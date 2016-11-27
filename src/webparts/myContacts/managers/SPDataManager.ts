import { IDataManager } from './Managers';
import * as Model from '../model/Model';
import { IWebPartContext } from '@microsoft/sp-client-preview';
import pnp from 'sp-pnp-js';

export class SPDataManager implements IDataManager {
  public SPContext: IWebPartContext;
  public ListId: string;

  private _itemCount: number = 0;

  public GetContacts(pageSize: number, pageNumber: number) : Promise<Model.IGetContactsResult>  {
    return new Promise<Model.IGetContactsResult>((resolve, reject) => {
      var result = { ItemCount: 0, Results: []};
      this.GetItemsCount().then((itemCount) => {
        result.ItemCount = itemCount;
        pnp.sp.web.lists.getById(this.ListId).items
                                            .select("Id", "FullName", "WebPage", "Email", "WorkPhone")
                                            .skip(pageNumber*pageSize)
                                            .top(pageSize)
                                            .get().then((items) => {
          items.forEach(item => {
            result.Results.push({
              Id: item.Id,
              DisplayName: item.FullName,
              Image: item.WebPage !== null ? item.WebPage.Url : '', //added as default contact list has no picture field
              Email: item.Email,
              Phone: item.WorkPhone
            });
          });
          resolve(result);
        });
      });
    });
  };

  private GetItemsCount(): Promise<number> {
      return new Promise<number>((resolve, reject) => {
        if(this._itemCount === 0) {
          pnp.sp.web.lists.getById(this.ListId).get().then((list) => {
            this._itemCount = list.ItemCount;
            resolve(this._itemCount);
          });
        }
        else {
          resolve(this._itemCount);
        }
      });
  };

  public GetContactLists(): Promise<Model.IList[]> {
    return new Promise<Model.IList[]>((resolve, reject) => {
      var results: Model.IList[] = [];
      pnp.sp.web.lists.filter("BaseTemplate eq 105").select("Id", "Title").get().then((lists) => {
        lists.forEach(list => {
          results.push({ Id: list.Id, Title: list.Title });
        });
        resolve(results);
      });
    });
  };
}