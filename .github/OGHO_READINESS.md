# OGHO publication readiness

Reviewed on October 27, 2026 against the Category 3 template at commit
`10f1ef9513802af1f4cb79014f87ffde2a024b90`:
<https://github.com/oracle-samples/template-repo-category3>.

## Repository preparation

- Root README includes the required disclaimer, purpose, prerequisites, required
  services, installation, test and build commands, deployment instructions, usage
  examples, documentation links, help, contribution guidance, security reporting,
  and license information.
- Root CONTRIBUTING.md and SECURITY.md match the Oracle template without changes.
- Root and sample LICENSE.txt files match the template with the copyright year
  changed to 2026. Both files use ASCII characters and LF line endings.
- Sample code remains under `cx-excel-plugin` with declared dependencies, npm
  commands, and an automated test and production build workflow.
- Relevant GitHub topics have been added while preserving the existing topics.
- [GitHub Actions run 40540307866](https://github.com/oracle-devrel/cx-innovations/actions/runs/40540307866)
  installed the declared dependencies, passed all 41 automated tests, and built
  production assets on Node.js 22.15.0. The workflow retains the resolved lockfile
  as a downloadable artifact for seven days. Direct dependency versions are pinned;
  the repository does not yet contain a lockfile because local registry and
  artifact download connections timed out.

## Required maintainer and governance follow-up

- **Repository name:** The maintainer requested keeping `cx-innovations` pending
  OGHO review. The current sample-only contents require a `-sample` or `-samples`
  suffix under the supplied naming rule. Ask the OGHO Admin Team to approve an
  appropriate name before publication; `cx-innovations-samples` is a candidate.
- **Repository description:** The current account has maintainer access but the
  GitHub repository update endpoint rejected description changes. An administrator
  must set the description. Suggested text: "Developer samples for extending
  Oracle Customer Experience, including an Excel add-in for Oracle CX Sales
  opportunity revenue items."
- **Documentation standards:** The supplied internal Oracle documentation,
  typography, and spelling guidance pages could not be accessed from this
  environment. An authorized reviewer must verify the documentation against them.
- **Tenant validation:** Automated tests do not replace a live Excel and Oracle CX
  Sales integration check. In a nonproduction tenant, verify load, create, update,
  authentication failure, and CORS failure. Validate the deployment manifest with
  Microsoft's validator. Confirm permission to distribute the included branding.
- **Publication approval:** The governance team must review readiness and make the
  repository public. The repository remains private.

## Timing

Prepare the repository within 30 days of the governance notice. If more time is
needed, notify the governance team. The notice states that repositories may be
deleted after 60 days without a response. The notice date was not supplied, so
this checklist does not assign a calendar deadline.
