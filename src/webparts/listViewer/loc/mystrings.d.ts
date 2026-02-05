declare interface IListViewerWebPartStrings {
  PropertyPaneGroupListAndView: string;
  PropertyPaneFieldListLabel: string;
  PropertyPaneFieldViewLabel: string;
  PropertyPaneDetailsViewLabel: string;
  PropertyPaneButtonEditView: string;
  PropertyPaneFieldNoEntriesLabel: string;
  PropertyPaneFieldNoEntriesPlaceholder: string;
  PropertyPaneFieldNoEntriesDescription: string;
  PropertyPaneGroupLabels: string;
  PropertyPaneFieldWebPartTitleLabel: string;
  PropertyPaneFieldWebPartTitleDescription: string;
  PropertyPaneFieldWebPartTitlePlaceholder: string;
  PropertyPaneFieldShowBodyCaptionLabel: string;
  SpinnerLoadingLabel: string;
  ConfigMissingMessage: string;
  DoubleClickForDetails: string;
  DetailsColumnHeader: string;
  DetailsDialogTitle: string;
  DetailsDialogClose: string;
  DetailsDialogCloseButtonAriaLabel: string;
  DetailsLoadingError: string;
  ViewHoverText: string;
}

declare module 'ListViewerWebPartStrings' {
  const strings: IListViewerWebPartStrings;
  export = strings;
}
