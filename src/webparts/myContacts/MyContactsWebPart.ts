import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  IWebPartContext,
  PropertyPaneSlider,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption } from '@microsoft/sp-webpart-base';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';
import * as strings from 'myContactsStrings';
import MyContacts, { IMyContactsProps } from './components/MyContacts';
import { IMyContactsWebPartProps } from './IMyContactsWebPartProps';
import { Dictionary } from '../../utilities/Dictionary';
import * as Managers from './managers/Managers';

export default class MyContactsWebPart extends BaseClientSideWebPart<IMyContactsWebPartProps> {

  private _managers = new Dictionary([
    { key: EnvironmentType.Local.toString(), value: new Managers.MockDataManager() },
    { key: EnvironmentType.SharePoint.toString(), value: new Managers.SPDataManager() }
  ]);
  private _dataManger: Managers.IDataManager;
  private _contactLists: Array<IPropertyPaneDropdownOption>;
  private _pictureSizes: Array<IPropertyPaneDropdownOption>;

  public onInit(): Promise<void> {
    this._dataManger = this._managers[Environment.type.toString()];
    this._dataManger.SPContext = this.context;
    this._dataManger.GetContactLists().then((results) => {
      this._contactLists = new Array<IPropertyPaneDropdownOption>();
      results.forEach(element => {
        this._contactLists.push({ key: element.Id, text: element.Title });
      });
    });

    this._pictureSizes = new Array<IPropertyPaneDropdownOption>();
    this._pictureSizes.push({ key: 0, text: "tiny" });
    this._pictureSizes.push({ key: 1, text: "extraSmall" });
    this._pictureSizes.push({ key: 2, text: "small" });
    this._pictureSizes.push({ key: 3, text: "regular" });
    this._pictureSizes.push({ key: 4, text: "large" });
    this._pictureSizes.push({ key: 5, text: "extraLarge" });

    return super.onInit();
  };

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  public render(): void {
    this._dataManger.ListId = this.properties.listId;
    const element: React.ReactElement<IMyContactsProps> = React.createElement(MyContacts, {
      spContext: this.context,
      dataManager: this._dataManger,
      pageSize: this.properties.pageSize,
      listId: this.properties.listId,
      showPhone: this.properties.showPhone,
      pictureSize: this.properties.pictureSize
    });

    ReactDom.render(element, this.domElement);
  };

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if(propertyPath === "listId") {
      this._dataManger.ListId = newValue;
    }
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  };

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.VisualizationPage },
          groups: [{
              groupName: strings.PaginationGroup,
              groupFields: [
                PropertyPaneSlider("pageSize", { label: strings.PaginationGroupPageSize, min: 6, max: 18 }),
              ]
            }, {
              groupName: strings.VisualizationGroup,
              groupFields: [
                PropertyPaneDropdown('pictureSize', { label: strings.VisualizationGroupImageSize, disabled: false, options: this._pictureSizes }),
                PropertyPaneToggle("showPhone", { label: strings.VisualizationGroupShowPhone, disabled: false })
              ]
            }
          ]
        },
        {
          header: { description: strings.ConfigurationPage },
          groups: [{
            groupName: strings.ConnectionGroup,
            groupFields: [
              PropertyPaneDropdown('listId', { label: strings.ConnectionGroupListName, disabled: false, options: this._contactLists })
            ]
          }]
        }
      ]
    };
  }
}
