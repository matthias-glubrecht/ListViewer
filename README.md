## List Viewer

This web part allows you to display a SharePoint list in a custom view. It supports displaying list items in a table, filtering, and a details view for individual items.

## Compatibility

*   SharePoint 2019 and later
*   SharePoint Online

## Technology

*   SharePoint Framework (SPFx) 1.4.1
*   React
*   Office UI Fabric React

## Features

*   **Custom List View:** Displays SharePoint list items based on a selected list view.
*   **Details Modal:** Clicking on the "Use Details" icon or double-clicking a row opens a modal dialog with full item details.
*   **Rich Text Support:** Properly renders multi-line text fields with HTML content.
*   **Responsive Design:** The details dialog adapts to different screen sizes.
*   **Localization:** Supports English (en-us) and German (de-de).

### Building the code

```bash
git clone the repo
npm i
npm i -g gulp
gulp
```

This package produces the following:

* lib/* - intermediate-stage commonjs build artifacts
* dist/* - the bundled script, along with other resources
* deploy/* - all resources which should be uploaded to a CDN.

### Build options

gulp clean - TODO
gulp test - TODO
gulp serve - TODO
gulp bundle - TODO
gulp package-solution - TODO
