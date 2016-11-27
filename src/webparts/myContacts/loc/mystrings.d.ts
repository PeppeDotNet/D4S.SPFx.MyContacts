declare interface IMyContactsStrings {
  PropertyPaneDescription: string;
  ConfigurationPage: string;
  ConnectionGroup: string;
  ConnectionGroupListName: string;
  VisualizationPage: string;
  VisualizationGroup: string;
  VisualizationGroupShowPicture: string;
  VisualizationGroupShowPhone: string;
  VisualizationGroupImageSize: string;
  PaginationGroup: string;
  PaginationGroupPageSize: string;
  DescriptionFieldLabel: string;
}

declare module 'myContactsStrings' {
  const strings: IMyContactsStrings;
  export = strings;
}
