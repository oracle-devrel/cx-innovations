<!--
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026 Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
-->

# CX Sales Excel Add-in

An independent open-source sample Excel add-in for reading and updating Oracle CX Sales opportunity revenue
items. It loads opportunities and product groups, writes revenue items to a worksheet, and sends
changed rows through an Oracle REST batch request.

The add-in supports product-group revenue items only. It does not delete items or check for
concurrent edits, so reload an opportunity before changing it if other users may have updated it.

This sample includes dependency declarations and npm commands for installing,
testing, building, and running the add-in. Repository-wide contribution and security policies
are linked below.

## Getting started

You need Node.js 22.15 or newer, npm, an Excel client that supports ExcelApi 1.8, and access to an
Oracle CX Sales environment or compatible gateway. Your Oracle user must be able to read
opportunities, product groups, and revenue items; creating or updating rows also requires write
and batch-request access.

From the repository root, install the dependencies and run the automated tests and build:

```sh
cd cx-excel-plugin
npm install
npm test
npm run build
npm start
```

`npm start` builds a development bundle and serves the add-in over HTTPS on port 8443.
The Office development certificate tooling may prompt you to trust a local certificate.
Sideload [manifest.xml](manifest.xml) into Excel using Microsoft's
[Office add-in sideloading instructions](https://learn.microsoft.com/office/dev/add-ins/testing/test-debug-office-add-ins).
The host and build examples are in [app.js](app.js), [webpack.config.js](webpack.config.js),
and [scripts/dev-server.js](scripts/dev-server.js).

## Usage examples

In a blank workbook, open **Home → CX Sales**, then:

1. Enter the environment origin, such as `https://your-pod.example.com` (no REST path).
2. Choose Basic authentication and enter a username/password, or choose Bearer and enter an access
   token.
3. Connect, select an opportunity, and choose **Load into Excel**.
4. Edit or add a product-group row and choose **Sync changes**. Reload to verify the result.

Use a nonproduction tenant and a least-privilege test account. **Sync changes makes real create
and update requests.** For a read-only smoke test, stop after loading the worksheet. Stop your
development server and remove the sideloaded add-in when finished.

Direct Oracle REST access requires CORS permission for the exact add-in origin (locally,
`https://localhost:8443`), methods `OPTIONS`, `GET`, and `POST`, and the `Authorization`,
`Content-Type`, and `REST-Framework-Version` headers. A gateway can provide the same REST contract
instead. Revenue-item responses must preserve each item's `self` or `canonical` link so updates
use the server-issued resource key.

## Libraries and tools

[package.json](package.json) declares the runtime and development dependencies.
Install them with `npm install`.

- `express` serves the add-in assets and public configuration in the hosting example.
- `webpack` and `webpack-cli` build the JavaScript bundles.
- `copy-webpack-plugin` and `html-webpack-plugin` copy assets and generate the HTML pages.
- `office-addin-dev-certs` manages local HTTPS certificates.
- `office-addin-manifest` validates the add-in manifest.

The Office.js library is loaded from Microsoft's CDN rather than installed through npm.

## Configure

The host serves public `config.json` to the task pane. Start with [config.example.json](config.example.json);
[config.schema.json](config.schema.json) lists the settings and accepted values. To change local
defaults without editing the tracked example:

```sh
cp config.example.json config.local.json
CX_CONFIG_FILE=config.local.json node scripts/dev-server.js
```

`config.local.json` is ignored by Git and is used only when selected with `CX_CONFIG_FILE`.
Restart the host after changing configuration. The host also accepts these public overrides:

| Variable                       | Setting                                |
| ------------------------------ | -------------------------------------- |
| `CX_API_ENVIRONMENT_URL`       | Default Oracle or gateway HTTPS origin |
| `CX_API_REST_PATH`             | REST root path                         |
| `CX_AUTH_MODE`                 | `basic` or `bearer`                    |
| `CX_AUTH_ALLOW_MODE_SELECTION` | `true` or `false`                      |

The example config includes the default resource paths, page limits, worksheet name, optional
product-group field mappings, branding, and extra revenue fields. `api.revenueItemLimit` cannot
exceed `workbook.maxRows`. Keep `ui.showDebug` disabled in production.

**Never put credentials or secrets in configuration or environment overrides.** `config.json` is
visible to anyone who can load the add-in. Enter a password or token in the task pane; it is kept
only for the current session and cleared from the input after connection. For unattended
authentication, use a customer-controlled gateway that keeps upstream credentials server-side.

## Verify and deploy

From the sample directory, run:

```sh
npm test
npm run build
npm run validate
```

The build writes static assets to `dist/`. Manifest validation calls Microsoft's service and
requires network access. An end-to-end test also needs a nonproduction Oracle tenant; exercise
load, create, update, authentication failure, and CORS failure before deployment.

For production, host `dist/` at a stable HTTPS URL. Copy `manifest.xml` for the deployment,
replace the `https://localhost:8443` asset URLs and sample metadata, and validate the copy with `npx office-addin-manifest validate path/to/manifest.xml`. Keep
the add-in ID stable across releases; leave the local manifest unchanged for local development.
Changing `CX_API_ENVIRONMENT_URL` does not change manifest asset URLs.

## Troubleshooting

- If **CX Sales** is missing from Excel, restart your development session, sideload the manifest
  again, and reopen a blank workbook.
- If the local certificate is not trusted, run `npx office-addin-dev-certs install` and accept
  the trust prompt.
- If the service is reachable but the authenticated request is blocked, check the tenant's CORS
  policy and authentication redirects. If reachability cannot be verified, check the URL, DNS,
  VPN, and TLS certificate.
- If sign-in succeeds but no data appears, check the user's Oracle roles and data-security grants.

## License and publication

This project is licensed under the Universal Permissive License, Version 1.0; see
[LICENSE.txt](LICENSE.txt). Attributions for the included third-party material are in
[THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt). Installed dependency directories and
generated build output are not included in the repository.

Before publishing a fork, confirm that you may distribute the bundled Oracle branding assets or
replace them. Review the Git history for old environment hostnames or deployment targets; removing
them from the current files does not remove them from prior commits.

## How to contribute

Read the repository [contribution guide](../CONTRIBUTING.md) for the Oracle Contributor
Agreement, commit sign-off requirements, and pull request process. Contributions to tests,
examples, accessibility, and documentation are welcome.

## Security

Do not report security vulnerabilities in public GitHub issues or pull requests. Follow
[SECURITY.md](../SECURITY.md) to report them privately to Oracle Security.

## Disclaimer

ORACLE AND ITS AFFILIATES DO NOT PROVIDE ANY WARRANTY WHATSOEVER, EXPRESS OR IMPLIED, FOR ANY SOFTWARE, MATERIAL OR CONTENT OF ANY KIND CONTAINED OR PRODUCED WITHIN THIS REPOSITORY, AND IN PARTICULAR SPECIFICALLY DISCLAIM ANY AND ALL IMPLIED WARRANTIES OF TITLE, NON-INFRINGEMENT, MERCHANTABILITY, AND FITNESS FOR A PARTICULAR PURPOSE.  FURTHERMORE, ORACLE AND ITS AFFILIATES DO NOT REPRESENT THAT ANY CUSTOMARY SECURITY REVIEW HAS BEEN PERFORMED WITH RESPECT TO ANY SOFTWARE, MATERIAL OR CONTENT CONTAINED OR PRODUCED WITHIN THIS REPOSITORY.  IN ADDITION, AND WITHOUT LIMITING THE FOREGOING, THIRD PARTIES MAY HAVE POSTED SOFTWARE, MATERIAL OR CONTENT TO THIS REPOSITORY WITHOUT ANY REVIEW. USE AT YOUR OWN RISK. 
