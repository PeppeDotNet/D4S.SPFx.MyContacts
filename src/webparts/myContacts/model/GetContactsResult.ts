import * as Model from './Model';

export interface IGetContactsResult {
  Results: Model.IContact[];
  ItemCount: number;
}