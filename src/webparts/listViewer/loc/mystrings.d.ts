declare interface IListViewerWebPartStrings {
  PropertyPaneGroupListAndView: string;
  PropertyPaneFieldListLabel: string;
  PropertyPaneFieldViewLabel: string;
  PropertyPaneButtonEditView: string;
  PropertyPaneFieldNoEntriesLabel: string;
  PropertyPaneFieldNoEntriesPlaceholder: string;
  PropertyPaneFieldNoEntriesDescription: string;
  PropertyPaneGroupWebPartTitle: string;
  PropertyPaneFieldWebPartTitleLabel: string;
  PropertyPaneFieldWebPartTitleDescription: string;
  PropertyPaneFieldWebPartTitlePlaceholder: string;
  ListViewerWelcomeTitle: string;
  ListViewerWelcomeSubtitle: string;
  ListViewerLearnMoreLabel: string;
  ListViewerLearnMoreUrl: string;
}

declare module 'ListViewerWebPartStrings' {
  const strings: IListViewerWebPartStrings;
  export = strings;
}
