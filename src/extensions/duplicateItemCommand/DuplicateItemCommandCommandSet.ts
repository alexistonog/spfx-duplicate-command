import { Log } from '@microsoft/sp-core-library';
import { SPHttpClient } from '@microsoft/sp-http';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IDuplicateItemCommandCommandSetProperties {
  sampleTextOne: string;
}

interface ISharePointFieldDefinition {
  InternalName: string;
  TypeAsString: string;
  Hidden: boolean;
  ReadOnlyField: boolean;
  Sealed: boolean;
}

interface IDuplicatePayloadResult {
  payload: Record<string, unknown>;
  skippedFields: string[];
}

const LOG_SOURCE: string = 'DuplicateItemCommandCommandSet';
const TARGET_LIST_ID: string = 'eba18070-be1a-4223-97e6-124c0dcec270';
const LIST_SPECIFIC_URL_FIELDS: Set<string> = new Set([
  'Event_x0020_Documents_x0020_Link'
]);
const SYSTEM_FIELD_NAMES: Set<string> = new Set([
  'Attachments',
  'AuthorId',
  'ComplianceAssetId',
  'CheckoutUserId',
  'ContentType',
  'ContentTypeId',
  'Created',
  'EditorId',
  'FileLeafRef',
  'FileRef',
  'FileSystemObjectType',
  'GUID',
  'ID',
  'Id',
  'Modified',
  'Order',
  'owshiddenversion',
  'WorkflowVersion'
]);

export default class DuplicateItemCommandCommandSet extends BaseListViewCommandSet<IDuplicateItemCommandCommandSetProperties> {

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized DuplicateItemCommandCommandSet');

    // initial state of the command's visibility
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      compareOneCommand.visible = false;
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public async onExecute(event: IListViewCommandSetExecuteEventParameters):Promise<void> {
    switch (event.itemId) {
      case 'COMMAND_1':
        if (!this._isTargetList()) {
          await this._showMessage('This command is only available on the configured list.');
          return;
        }

        await this._duplicateSelectedItem(event);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  private async _duplicateSelectedItem(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    try {
      const row = event.selectedRows[0];
      const itemId = row?.getValueByName('ID');
      const webUrl = this.context.pageContext.web.absoluteUrl;
      const listId = this.context.pageContext.list?.id.toString();

      if (!itemId || !listId) {
        throw new Error('Unable to resolve the selected item or current list.');
      }

      const getResponse = await this.context.spHttpClient.get(
        `${webUrl}/_api/web/lists(guid'${listId}')/items(${itemId})`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata'
          }
        }
      );

      if (!getResponse.ok) {
        throw new Error(await getResponse.text());
      }

      const sourceItem = await getResponse.json() as Record<string, unknown>;
      const fields = await this._getWritableFields(webUrl, listId);
      const { payload: newItem, skippedFields } = this._buildDuplicatePayload(sourceItem, fields);

      const createResponse = await this.context.spHttpClient.post(
        `${webUrl}/_api/web/lists(guid'${listId}')/items`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: 'application/json;odata=nometadata',
            'Content-Type': 'application/json;odata=nometadata'
          },
          body: JSON.stringify(newItem)
        }
      );

      if (!createResponse.ok) {
        throw new Error(await createResponse.text());
      }

      const successMessage = skippedFields.length > 0
        ? `Item duplicated. Unsupported fields were skipped: ${skippedFields.join(', ')}`
        : 'Item duplicated successfully.';

      await this._showMessage(successMessage);
      location.reload();
    } catch (error) {
      console.error(error);
      const message = error instanceof Error ? error.message : 'Error duplicating item.';
      await this._showMessage(message);
    }
  }

  private async _getWritableFields(webUrl: string, listId: string): Promise<ISharePointFieldDefinition[]> {
    const fieldsResponse = await this.context.spHttpClient.get(
      `${webUrl}/_api/web/lists(guid'${listId}')/fields?$select=InternalName,TypeAsString,Hidden,ReadOnlyField,Sealed`,
      SPHttpClient.configurations.v1,
      {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      }
    );

    if (!fieldsResponse.ok) {
      throw new Error(await fieldsResponse.text());
    }

    const fieldResult = await fieldsResponse.json() as { value: ISharePointFieldDefinition[] };

    return fieldResult.value.filter((field: ISharePointFieldDefinition) => {
      return !field.Hidden && !field.ReadOnlyField && !field.Sealed && !SYSTEM_FIELD_NAMES.has(field.InternalName);
    });
  }

  private _buildDuplicatePayload(
    sourceItem: Record<string, unknown>,
    fields: ISharePointFieldDefinition[]
  ): IDuplicatePayloadResult {
    const newItem: Record<string, unknown> = {};
    const skippedFields: string[] = [];

    fields.forEach((field: ISharePointFieldDefinition) => {
      const wasMapped = this._tryMapFieldValue(sourceItem, newItem, field);

      if (!wasMapped) {
        skippedFields.push(field.InternalName);
      }
    });

    if (typeof sourceItem.Title === 'string' && sourceItem.Title.trim()) {
      newItem.Title = `${sourceItem.Title} - Copy`;
    }

    return {
      payload: newItem,
      skippedFields
    };
  }

  private _tryMapFieldValue(
    sourceItem: Record<string, unknown>,
    payload: Record<string, unknown>,
    field: ISharePointFieldDefinition
  ): boolean {
    const fieldName = field.InternalName;
    const fieldValue = sourceItem[fieldName];
    const lookupFieldName = `${fieldName}Id`;
    const lookupValue = sourceItem[lookupFieldName];

    if (fieldName.startsWith('_') || fieldName.includes('@odata.')) {
      return false;
    }

    switch (field.TypeAsString) {
      case 'Text':
      case 'Note':
      case 'Choice':
      case 'Number':
      case 'Currency':
      case 'Integer':
      case 'Boolean':
      case 'DateTime':
        return this._assignPrimitiveValue(payload, fieldName, fieldValue);
      case 'MultiChoice':
        return this._assignMultiChoiceValue(payload, fieldName, fieldValue);
      case 'URL':
        return this._assignUrlValue(payload, fieldName, fieldValue);
      case 'User':
      case 'Lookup':
        return this._assignLookupValue(payload, lookupFieldName, lookupValue);
      case 'UserMulti':
      case 'LookupMulti':
        return this._assignLookupMultiValue(payload, lookupFieldName, lookupValue);
      default:
        return false;
    }
  }

  private _assignPrimitiveValue(
    payload: Record<string, unknown>,
    fieldName: string,
    fieldValue: unknown
  ): boolean {
    if (
      typeof fieldValue === 'string' ||
      typeof fieldValue === 'number' ||
      typeof fieldValue === 'boolean' ||
      fieldValue === null
    ) {
      payload[fieldName] = fieldValue;
      return true;
    }

    return false;
  }

  private _assignMultiChoiceValue(
    payload: Record<string, unknown>,
    fieldName: string,
    fieldValue: unknown
  ): boolean {
    if (Array.isArray(fieldValue)) {
      payload[fieldName] = fieldValue;
      return true;
    }

    return false;
  }

  private _assignUrlValue(
    payload: Record<string, unknown>,
    fieldName: string,
    fieldValue: unknown
  ): boolean {
    if (fieldValue === null) {
      payload[fieldName] = null;
      return true;
    }

    if (
      typeof fieldValue === 'object' &&
      fieldValue !== null &&
      ('Url' in fieldValue || 'url' in fieldValue)
    ) {
      const urlFieldValue = fieldValue as {
        Url?: unknown;
        url?: unknown;
        Description?: unknown;
        description?: unknown;
      };
      const url = typeof urlFieldValue.Url === 'string'
        ? urlFieldValue.Url
        : typeof urlFieldValue.url === 'string'
          ? urlFieldValue.url
          : undefined;
      const description = typeof urlFieldValue.Description === 'string'
        ? urlFieldValue.Description
        : typeof urlFieldValue.description === 'string'
          ? urlFieldValue.description
          : '';

      if (!url) {
        return false;
      }

      payload[fieldName] = {
        Url: url,
        Description: description
      };
      return true;
    }

    if (typeof fieldValue === 'string' && fieldValue.trim()) {
      payload[fieldName] = {
        Url: fieldValue,
        Description: LIST_SPECIFIC_URL_FIELDS.has(fieldName) ? fieldValue : ''
      };
      return true;
    }

    return false;
  }

  private _assignLookupValue(
    payload: Record<string, unknown>,
    fieldName: string,
    fieldValue: unknown
  ): boolean {
    if (typeof fieldValue === 'number') {
      payload[fieldName] = fieldValue;
      return true;
    }

    return false;
  }

  private _assignLookupMultiValue(
    payload: Record<string, unknown>,
    fieldName: string,
    fieldValue: unknown
  ): boolean {
    if (Array.isArray(fieldValue)) {
      payload[fieldName] = fieldValue;
      return true;
    }

    if (
      typeof fieldValue === 'object' &&
      fieldValue !== null &&
      'results' in fieldValue &&
      Array.isArray((fieldValue as { results: unknown[] }).results)
    ) {
      payload[fieldName] = (fieldValue as { results: unknown[] }).results;
      return true;
    }

    return false;
  }

  private _showMessage(message: string): Promise<void> {
    return Dialog.alert(message).catch(() => {
      alert(message);
    });
  }

  private _isTargetList(): boolean {
    const currentListId = this.context.pageContext.list?.id.toString().toLowerCase();

    return currentListId === TARGET_LIST_ID;
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = this._isTargetList() && this.context.listView.selectedRows?.length === 1;
    }

    // TODO: Add your logic here

    // You should call this.raiseOnChage() to update the command bar
    this.raiseOnChange();
  }
}
