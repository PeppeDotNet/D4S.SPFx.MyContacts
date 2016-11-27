import * as Model from '../model/Model';
import { IWebPartContext } from '@microsoft/sp-client-preview';

export interface IDataManager {
  SPContext: IWebPartContext;
  ListId: string;

  GetContacts(pageSize: number, pageNumber: number): Promise<Model.IGetContactsResult>;
  GetContactLists(): Promise<Model.IList[]>;
}
export { MockDataManager } from './MockDataManager';
export { SPDataManager } from './SPDataManager';
