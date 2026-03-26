# Copilot Instructions — ListViewer SPFx Web Part

## Project Overview

This is a **SharePoint Framework (SPFx) 1.4** client-side web part built with **React 15**, **TypeScript**, and **Office UI Fabric React**. It displays SharePoint list data in a tabular view and provides a details dialog for individual items. The build toolchain uses **Gulp** and the SPFx build pipeline.

## Project Structure

Follow the existing folder conventions:

- `src/webparts/listViewer/` — web part entry point and property pane configuration.
- `src/webparts/listViewer/components/` — React components, each in its own subfolder containing the component `.tsx`, a props interface file, an optional state interface file, a `.module.scss` stylesheet, and an `index.ts` barrel export.
- `src/webparts/listViewer/service/` — service layer with an `IListViewerService` interface and the `ListViewerService` implementation that wraps `@pnp/sp` calls.
- `src/webparts/listViewer/utility/` — small, stateless helper classes (e.g. `Utility`).
- `src/webparts/listViewer/loc/` — localization files (`en-us.js`, `de-de.js`, etc.) and the `mystrings.d.ts` type declaration.
- `src/controls/` — reusable property pane controls (e.g. `PropertyPaneAsyncDropdown`).

## Coding Conventions

### TypeScript & Linting

- The project uses **TSLint** with the `@microsoft/sp-tslint-rules` base config. Respect the rule overrides defined in `tslint.json`.
- Where a TSLint rule must be suppressed locally (e.g. `no-any`, `max-line-length`), use inline `// tslint:disable` or `// tslint:disable-next-line` comments and keep the scope as narrow as possible.
- Prefer explicit type annotations on public members and function return types. The `typedef` rule is disabled, but explicit types improve readability.

### React Components

- **Class components** are used for stateful components (e.g. `ListViewer`, `DetailsView`). Use `React.Component<TProps, TState>` with dedicated `IXxxProps` and `IXxxState` interfaces.
- **Stateless functional components** are used for simple presentational components (e.g. `ListItem`). Define them as `React.StatelessComponent<TProps>`.
- Each component lives in its own subfolder under `components/`. Export the default component and its props interface via an `index.ts` barrel file:
  ```ts
  export { default as MyComponent } from './MyComponent';
  export * from './IMyComponentProps';
  ```
- Use **arrow-function class properties** for event handlers and callbacks to avoid manual `.bind(this)` in the constructor or render method.

### Styling with CSS Modules

- Each component has a co-located `*.module.scss` file that is imported as `styles`.
- Use `className={styles.ClassName}` in JSX — never use plain string class names.
- Import the SPFx Fabric Core when you need its SCSS mixins or variables:
  ```scss
  @import "~@microsoft/sp-office-ui-fabric-core/dist/sass/SPFabricCore.scss";
  ```

### Dialog & Modal Sizing

When using Fluent UI / Office UI Fabric React `Dialog` components, **never set the width directly on the `<Dialog>` element**. Instead, control sizing through CSS classes applied via the component's props objects:

- Use `modalProps.containerClassName` to target the outer modal container. This is the primary mechanism for controlling the overall dialog width. Set `min-width` and `max-width` with `!important` to override the Fabric defaults.
- Use `dialogContentProps.className` to style the inner dialog content area (e.g. minimum width constraints).
- Use a separate CSS class on the inner `<div>` for content-specific layout (padding, label widths, etc.).

Example from the existing `DetailsView` implementation:

```scss
// DetailsView.module.scss
.DetailsDialogModal {
    min-width: 600px !important;
    max-width: 1200px !important;
    background-color: #f0f0f0;
}

.DetailsDialogContent {
    min-width: 400px;
}

.DetailsDialogInner {
    min-width: 400px;
    padding: 8px;
}
```

```tsx
// DetailsView.tsx
const modalProps: IModalProps = {
    containerClassName: styles.DetailsDialogModal,
    // ...
};
const dialogContentProps: IDialogContentProps = {
    className: styles.DetailsDialogContent,
    // ...
};
```

This three-layer approach (modal container → dialog content → inner div) keeps sizing rules centralized in SCSS and avoids inline styles or ad-hoc width props.

### Localization

- All user-facing strings are defined in the `loc/` folder. Add new string keys to `mystrings.d.ts` and provide translations in each locale file (`en-us.js`, `de-de.js`, etc.).
- Import strings via the SPFx localization module:
  ```ts
  import * as strings from 'ListViewerWebPartStrings';
  ```
- Never hard-code user-visible text in components.

### Service Layer

- All SharePoint REST/CSOM calls go through the `IListViewerService` interface. Components receive the service via props — they never instantiate it directly.
- The `ListViewerService` implementation caches promises for repeated calls (views, fields, list title, etc.) to avoid redundant network requests. When a cached call fails, delete the cached promise so the next invocation retries.
- Use `@pnp/sp` for SharePoint data access. Build CAML queries for `getItemsByCAMLQuery` and use `select` / `expand` to minimize payload size.

### Property Pane

- Use `@pnp/spfx-property-controls` for list pickers and similar standard controls.
- For async dropdowns (e.g. view pickers), use the custom `PropertyPaneAsyncDropdown` control in `src/controls/`.
- The web part disables reactive property changes (`disableReactivePropertyChanges` returns `true`), so property pane changes are applied only when the pane is closed or the user clicks Apply.

## Build & Deployment

- `gulp bundle` — bundle the solution.
- `gulp package-solution` — produce the `.sppkg` package in `sharepoint/solution/`.
- `gulp serve` — start the local workbench for development.
- Configuration files for bundling, deployment, and serving live in the `config/` folder.
