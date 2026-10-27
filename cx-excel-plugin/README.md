<!--
** CX Sales Excel Add-in version 1.0.
**
** Copyright (c) 2026, Oracle and/or its affiliates.
** Licensed under the Universal Permissive License v 1.0 as shown at https://oss.oracle.com/licenses/upl/
-->

# CX Sales Excel Add-in

An independent open-source sample Excel add-in for reading and updating Oracle CX Sales opportunity revenue
items. It loads opportunities and product groups, writes revenue items to a worksheet, and sends
changed rows through an Oracle REST batch request.

The add-in supports product-group revenue items only. It does not delete items or check for
concurrent edits, so reload an opportunity before changing it if other users may have updated it.

This repository provides reference code. `package.json` contains project metadata only, and no
npm dependency declarations or lockfile are distributed. To run or adapt the sample, supply the
libraries and development tools listed below in your own project.

## Use the sample

You need Node.js 22.15 or newer, npm, an Excel client that supports ExcelApi 1.8, and access to an
Oracle CX Sales environment or compatible gateway. Your Oracle user must be able to read
opportunities, product groups, and revenue items; creating or updating rows also requires write
and batch-request access.

Set up the referenced libraries in your own development environment. The examples in
[webpack.config.js](webpack.config.js), [app.js](app.js), and
[scripts/dev-server.js](scripts/dev-server.js) show how to build and host the add-in over HTTPS on
port 8443. Sideload [manifest.xml](manifest.xml) into Excel using your Office development tooling,
and trust the local development certificate if prompted.

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

## Referenced npm libraries and tools

The following packages are referenced for adapting and running the sample. Select compatible
versions and declare them in your own project's dependency manifest.

- `express` serves the add-in assets and public configuration in the hosting example.
- `webpack` and `webpack-cli` build the JavaScript bundles.
- `copy-webpack-plugin` and `html-webpack-plugin` copy assets and generate the HTML pages.
- `office-addin-dev-certs` manages local HTTPS certificates.
- `office-addin-debugging` is optional tooling for starting, stopping, and sideloading the add-in.
- `office-addin-manifest` is optional tooling for validating the add-in manifest.

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

After supplying the referenced libraries in your development environment, the example commands are:

```sh
node --test
npx webpack --mode production
npx office-addin-manifest validate manifest.xml
```

The build writes static assets to `dist/`. Manifest validation calls Microsoft's service and
requires network access. An end-to-end test also needs a nonproduction Oracle tenant; exercise
load, create, update, authentication failure, and CORS failure before deployment.

For production, host `dist/` at a stable HTTPS URL. Copy `manifest.xml` for the deployment,
replace the `https://localhost:8443` asset URLs and sample metadata, and validate the copy. Keep
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
[THIRD_PARTY_LICENSES.txt](THIRD_PARTY_LICENSES.txt). The npm library and tool references are listed
only in this README; their packages and generated build output are not included in the repository.

Before publishing a fork, confirm that you may distribute the bundled Oracle branding assets or
replace them. Review the Git history for old environment hostnames or deployment targets; removing
them from the current files does not remove them from prior commits.
