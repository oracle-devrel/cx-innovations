ORACLE AND ITS AFFILIATES DO NOT PROVIDE ANY WARRANTY WHATSOEVER, EXPRESS OR IMPLIED, FOR ANY SOFTWARE, MATERIAL OR CONTENT OF ANY KIND CONTAINED OR PRODUCED WITHIN THIS REPOSITORY, AND IN PARTICULAR SPECIFICALLY DISCLAIM ANY AND ALL IMPLIED WARRANTIES OF TITLE, NON-INFRINGEMENT, MERCHANTABILITY, AND FITNESS FOR A PARTICULAR PURPOSE.  FURTHERMORE, ORACLE AND ITS AFFILIATES DO NOT REPRESENT THAT ANY CUSTOMARY SECURITY REVIEW HAS BEEN PERFORMED WITH RESPECT TO ANY SOFTWARE, MATERIAL OR CONTENT CONTAINED OR PRODUCED WITHIN THIS REPOSITORY.  IN ADDITION, AND WITHOUT LIMITING THE FOREGOING, THIRD PARTIES MAY HAVE POSTED SOFTWARE, MATERIAL OR CONTENT TO THIS REPOSITORY WITHOUT ANY REVIEW. USE AT YOUR OWN RISK. 

# CX Innovations

Developer samples for extending Oracle Customer Experience (CX). This repository
currently contains an Excel add-in that reads and updates Oracle CX Sales
opportunity revenue items.

## Samples

| Sample | Purpose | Documentation |
| --- | --- | --- |
| [CX Sales Excel Add-in](cx-excel-plugin/) | Load product-group revenue items into Excel and synchronize edited or new rows with Oracle CX Sales. | [Setup, configuration, and deployment](cx-excel-plugin/README.md) |

The add-in supports product-group revenue items. It does not delete revenue items
or detect concurrent edits. Reload an opportunity before editing it when other
users may have changed it. These samples demonstrate integration concepts and
require evaluation and adaptation before production use.

## Getting started

### Prerequisites and required services

- Node.js 22.15.0 or later and npm.
- Git and access to this repository.
- Microsoft Excel with support for ExcelApi 1.8 and permission to sideload an
  Office add-in. See Microsoft's [sideloading instructions](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).
- A nonproduction Oracle CX Sales environment or a compatible gateway. The Oracle
  account needs permission to read opportunities, product groups, and revenue
  items. Creating and updating items also requires write and REST batch access.
- Cross-Origin Resource Sharing (CORS) configured for the add-in origin, or a
  gateway providing the required REST contract. See the [sample prerequisites](cx-excel-plugin/README.md#getting-started).

### Install, test, and build

```sh
git clone https://github.com/oracle-devrel/cx-innovations.git
cd cx-innovations/cx-excel-plugin
npm install
npm test
npm run build
```

The build writes static files to `cx-excel-plugin/dist/`. Automated tests use
local fixtures and do not require an Oracle tenant or write to Oracle CX Sales.

### Run locally

From `cx-excel-plugin`, run:

```sh
npm start
```

The development server builds the add-in and serves it over HTTPS on port 8443.
The Office development certificate tooling may prompt you to trust a local
certificate. Sideload `cx-excel-plugin/manifest.xml` into Excel using the Microsoft
instructions linked above. Open **Home > CX Sales** in a blank workbook.

### Deploy

Run `npm run build`, host the generated `dist/` files at a stable HTTPS URL,
and update a deployment copy of `manifest.xml` with that host's asset URLs.
Validate the deployment manifest with `npx office-addin-manifest validate path/to/manifest.xml`.
See the [deployment instructions](cx-excel-plugin/README.md#verify-and-deploy)
for configuration, manifest handling, and tenant validation.

## Usage examples

For a read-only example, connect to a nonproduction Oracle CX Sales environment,
select an opportunity, and choose **Load into Excel**. Confirm that product-group
revenue items appear in the worksheet.

For a write example, edit a revenue amount or add a product-group row, choose
**Sync changes**, and reload the opportunity to verify the result. **Sync changes
creates or updates real Oracle CX Sales data.** Use a least-privilege test account.

Enter credentials only in the add-in's connection form. Public configuration files
must never contain passwords or access tokens. See the
[sample usage and configuration guide](cx-excel-plugin/README.md).

## Documentation

- [Sample setup, configuration, troubleshooting, and deployment](cx-excel-plugin/README.md)
- [Runtime configuration schema](cx-excel-plugin/config.schema.json)
- [Example public configuration](cx-excel-plugin/config.example.json)
- [Third-party notices](cx-excel-plugin/THIRD_PARTY_LICENSES.txt)

## Help

For reproducible bugs, documentation problems, or enhancement requests, open an
[issue](https://github.com/oracle-devrel/cx-innovations/issues). Include the sample
name, software versions, steps to reproduce the issue, and expected and actual
results. Remove credentials, tenant URLs, and customer data from examples and logs.
This repository provides sample code; it does not establish an Oracle product
support commitment.

## How to contribute

Contributions to code, tests, examples, and documentation are welcome. Before
submitting a pull request, read [CONTRIBUTING.md](CONTRIBUTING.md), sign the
Oracle Contributor Agreement, and follow the sign-off and review instructions.
Run the sample's tests and production build before submitting code changes.

## Security

Do not report security vulnerabilities in public GitHub issues or pull requests.
Follow [SECURITY.md](SECURITY.md) to report them privately to Oracle Security.

## License

Copyright (c) 2026 Oracle and/or its affiliates.

Released under the [Universal Permissive License, Version 1.0](LICENSE.txt).
See the [third-party notices](cx-excel-plugin/THIRD_PARTY_LICENSES.txt) for included
third-party material.
