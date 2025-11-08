NOTES:
https://stackblitz.com/edit/remote-data-fluintui-combbox?file=index.html,src%2Findex.tsx,src%2FApp.tsx,src%2Fmain.tsx
https://stackblitz.com/edit/fxzezjk9-eweftyxd?file=src%2Fexample.tsx

https://www.openplzapi.org/en/api/
https://stackblitz.com/

https://www.odata.org/odata-services/
https://medium.com/@manojtharindu11/mastering-typescripts-map-method-for-efficient-array-transformation-e9afb069ee46
https://carlrippon.com/fetch-with-async-await-and-typescript/
https://legacy.reactjs.org/docs/hooks-state.html

https://storybooks.fluentui.dev/react/?path=/docs/concepts-introduction--docs
https://fluent2.microsoft.design/get-started/whatisnew
https://storybooks.fluentui.dev/react/iframe.html?viewMode=docs&id=components-combobox--docs

Mailing:
https://learn.microsoft.com/en-us/graph/api/user-sendmail?view=graph-rest-1.0&tabs=javascript
https://learn.microsoft.com/en-us/graph/api/resources/itembody?view=graph-rest-1.0
https://learn.microsoft.com/en-us/graph/outlook-create-send-messages#reading-messages-with-control-over-the-body-format-returned


# Upgrade project dev-sky Dynamic Form App to v1.21.1

Date: 18.9.2025

## Findings

Following is the list of steps required to upgrade your project to SharePoint Framework version 1.21.1. [Summary](#Summary) of the modifications is included at the end of the report.

### FN001001 @microsoft/sp-core-library | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-core-library

Execute the following command:

```sh
npm i -SE @microsoft/sp-core-library@1.21.1
```

File: [./package.json:19:5](./package.json)

### FN001002 @microsoft/sp-lodash-subset | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-lodash-subset

Execute the following command:

```sh
npm i -SE @microsoft/sp-lodash-subset@1.21.1
```

File: [./package.json:20:5](./package.json)

### FN001003 @microsoft/sp-office-ui-fabric-core | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-office-ui-fabric-core

Execute the following command:

```sh
npm i -SE @microsoft/sp-office-ui-fabric-core@1.21.1
```

File: [./package.json:21:5](./package.json)

### FN001004 @microsoft/sp-webpart-base | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-webpart-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-webpart-base@1.21.1
```

File: [./package.json:23:5](./package.json)

### FN001021 @microsoft/sp-property-pane | Required

Upgrade SharePoint Framework dependency package @microsoft/sp-property-pane

Execute the following command:

```sh
npm i -SE @microsoft/sp-property-pane@1.21.1
```

File: [./package.json:22:5](./package.json)

### FN001034 @microsoft/sp-adaptive-card-extension-base | Optional

Upgrade SharePoint Framework dependency package @microsoft/sp-adaptive-card-extension-base

Execute the following command:

```sh
npm i -SE @microsoft/sp-adaptive-card-extension-base@1.21.1
```

File: [./package.json:18:5](./package.json)

### FN002001 @microsoft/sp-build-web | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-build-web

Execute the following command:

```sh
npm i -DE @microsoft/sp-build-web@1.21.1
```

File: [./package.json:33:5](./package.json)

### FN002002 @microsoft/sp-module-interfaces | Required

Upgrade SharePoint Framework dev dependency package @microsoft/sp-module-interfaces

Execute the following command:

```sh
npm i -DE @microsoft/sp-module-interfaces@1.21.1
```

File: [./package.json:34:5](./package.json)

### FN002022 @microsoft/eslint-plugin-spfx | Required

Upgrade SharePoint Framework dev dependency package @microsoft/eslint-plugin-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-plugin-spfx@1.21.1
```

File: [./package.json:30:5](./package.json)

### FN002023 @microsoft/eslint-config-spfx | Required

Upgrade SharePoint Framework dev dependency package @microsoft/eslint-config-spfx

Execute the following command:

```sh
npm i -DE @microsoft/eslint-config-spfx@1.21.1
```

File: [./package.json:29:5](./package.json)

### FN002024 eslint | Required

Upgrade SharePoint Framework dev dependency package eslint

Execute the following command:

```sh
npm i -DE eslint@8.57.1
```

File: [./package.json:40:5](./package.json)

### FN002026 typescript | Required

Upgrade SharePoint Framework dev dependency package typescript

Execute the following command:

```sh
npm i -DE typescript@5.3.3
```

File: [./package.json:43:5](./package.json)

### FN002029 @microsoft/rush-stack-compiler-5.3 | Required

Install SharePoint Framework dev dependency package @microsoft/rush-stack-compiler-5.3

Execute the following command:

```sh
npm i -DE @microsoft/rush-stack-compiler-5.3@0.1.0
```

File: [./package.json:28:3](./package.json)

### FN012017 tsconfig.json extends property | Required

Update tsconfig.json extends property

```json
{
  "extends": "./node_modules/@microsoft/rush-stack-compiler-5.3/includes/tsconfig-web.json"
}
```

File: [./tsconfig.json:2:3](./tsconfig.json)

### FN021003 package.json engines.node | Required

Update package.json engines.node property

```json
{
  "engines": {
    "node": ">=22.14.0 < 23.0.0"
  }
}
```

File: [./package.json:6:5](./package.json)

### FN002021 @rushstack/eslint-config | Required

Upgrade SharePoint Framework dev dependency package @rushstack/eslint-config

Execute the following command:

```sh
npm i -DE @rushstack/eslint-config@4.0.1
```

File: [./package.json:35:5](./package.json)

### FN017001 Run npm dedupe | Optional

If, after upgrading npm packages, when building the project you have errors similar to: "error TS2345: Argument of type 'SPHttpClientConfiguration' is not assignable to parameter of type 'SPHttpClientConfiguration'", try running 'npm dedupe' to cleanup npm packages.

Execute the following command:

```sh
npm dedupe
```

File: [./package.json](./package.json)

## Summary

### Execute script

```sh
npm i -SE @microsoft/sp-core-library@1.21.1 @microsoft/sp-lodash-subset@1.21.1 @microsoft/sp-office-ui-fabric-core@1.21.1 @microsoft/sp-webpart-base@1.21.1 @microsoft/sp-property-pane@1.21.1 @microsoft/sp-adaptive-card-extension-base@1.21.1
npm i -DE @microsoft/sp-build-web@1.21.1 @microsoft/sp-module-interfaces@1.21.1 @microsoft/eslint-plugin-spfx@1.21.1 @microsoft/eslint-config-spfx@1.21.1 eslint@8.57.1 typescript@5.3.3 @microsoft/rush-stack-compiler-5.3@0.1.0 @rushstack/eslint-config@4.0.1
npm dedupe
```

### Modify files

#### [./tsconfig.json](./tsconfig.json)

Update tsconfig.json extends property:

```json
{
  "extends": "./node_modules/@microsoft/rush-stack-compiler-5.3/includes/tsconfig-web.json"
}
```

#### [./package.json](./package.json)

Update package.json engines.node property:

```json
{
  "engines": {
    "node": ">=22.14.0 < 23.0.0"
  }
}
```
