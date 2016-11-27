import * as Model from '../model/Model';
import { IDataManager } from './Managers';
import { IWebPartContext } from '@microsoft/sp-client-preview';

export class MockDataManager implements IDataManager {

  public SPContext: IWebPartContext;
  public ListId: string;

  private _itemCount: number = 16;
  private _serverLoadDelay: number = 1000;

  public GetContacts(pageSize: number, pageNumber: number): Promise<Model.IGetContactsResult> {
    return new Promise<Model.IGetContactsResult>((resolve, reject) => {
      var contacts = this.GetAllContacts();
      var pagedContacts: Model.IContact[] = [];

      //simulate server load...
      this.Sleep(this._serverLoadDelay).then(() => {
        for (var i = (pageSize * pageNumber); i < ((pageSize * pageNumber) + pageSize); i++) {
          if(i < contacts.length ) {
            pagedContacts.push(contacts[i]);
          }
        }
        var result: Model.IGetContactsResult = { ItemCount: this._itemCount, Results: pagedContacts };
        resolve(result);
      });
    });
  }

  public GetContactLists(): Promise<Model.IList[]> {
    return new Promise<Model.IList[]>((resolve, reject) => {
      //simulate server load...
      this.Sleep(this._serverLoadDelay).then(() => {
        var results = [
          { Title: "Contact list 1", Id: "f6785ba2-30a3-4b2d-a756-9371d416ae67" },
          { Title: "Contact list 2", Id: "cc475be1-b08b-49c4-9776-b26c31d18216" },
          { Title: "Contact list 3", Id: "3c15be57-41c8-45be-9456-abe04fbe00ed" }
        ];
        resolve(results);
      });
    });
  };

  private GetAllContacts(): Model.IContact[] {
    var mockData = [];
    for (var i = 0; i < this._itemCount; i++) {
      mockData.push({
        DisplayName: 'Peppe Marchi',
        Email: 'giuseppe.marchi@dev4side.com',
        Image: '../src/webparts/myContacts/content/peppe.jpg',
        Phone: '02-83439531',
        Id: (i + 1)
      });
    }
    return mockData;
  }

  private Sleep (delay) {
    return new Promise((resolve) => {
      setTimeout(resolve, delay);
    });
  }
}