# Foundry Quota Increase

An Agent Skill for preparing and submitting Microsoft Foundry and Azure OpenAI quota increase requests. It validates an Excel workbook or information collected interactively, then uses Playwright to fill the [official Microsoft quota request form](https://aka.ms/aoai/quotaincrease).

Prepare requests for multiple models and keep local screenshots and submission records for each request. The default mode is **fill and capture without submitting**. Actual submissions run one at a time, only after the user explicitly approves the exact information to be sent.

Skill name: `foundry-quota-increase`. Current version: `1.1.0`. See [SKILL.md](./SKILL.md) for the agent's detailed operating instructions.

> [!IMPORTANT]
> Setting an Excel row to `Y` selects a submission candidate; it is not submission approval. A successful form submission does not mean Microsoft has approved the quota or that the model can be deployed.

## Contents

- [Features](#features)
- [Supported scope](#supported-scope)
- [Getting started](#getting-started)
- [Usage examples](#usage-examples)
- [Input format](#input-format)
- [Workflow and submission approval](#workflow-and-submission-approval)
- [Local commands](#local-commands)
- [Repository layout](#repository-layout)
- [Development and testing](#development-and-testing)
- [Privacy and execution records](#privacy-and-execution-records)
- [License](#license)
- [References](#references)

## Features

- **Excel or interactive intake**: Read shared contact information and model-specific requests from a workbook, or ask for the required details in small groups.
- **Request validation**: Check required fields, email addresses, subscription GUIDs, duplicate request IDs, and TPM quantities. Match model and region selections against the live form.
- **Input verification and screenshots**: Read back actual form values, compare them with the source, and capture the screen. Block submission when values do not match.
- **Scoped approval**: Record approval only for the requests shown in the preview. Require fresh approval if the information to be sent changes.
- **Submission tracking and duplicate prevention**: Persist request states and submission attempts. Stop the batch when a result is uncertain instead of retrying automatically.

## Supported scope

| Item | Scope |
| --- | --- |
| Model types | Azure OpenAI chat and Azure Direct Model (ADM) |
| Deployment type | Global Standard |
| Requested unit | TPM, converted to the form's kTPM value |
| Input sources | Excel, or interactive intake saved as local JSON |
| Execution modes | Fill and capture; actual submission after user approval |

Audio/Realtime RPM, PTU/Provisioned, Batch, Fine-tuning, Managed Compute, Anthropic, and Fireworks requests are outside the automated scope. Unsupported requests are held or stopped, never silently converted to another request type.

Model names and available regions must be checked against the live form. When no region question appears, record it as "not asked." This is not evidence that the model is supported in a particular region.

## Getting started

### Prerequisites

- Python 3.10 or later
- The `openpyxl` Python package
- An agent with Playwright browser tools
- Node.js 22.20 or later and npm when installing with `npx skills`

The Python scripts validate input, manage execution records, and generate browser functions. **Running a script alone does not open a browser or submit a request.** The agent must execute generated functions with its Playwright tools.

### Installation options

Use `npx skills add <source>` to install a skill. `npm install skills` installs the CLI package itself; `npx skills add` runs that CLI to install skill files for an agent. This skill does not need to be published as an npm package.

| Method | Source | Purpose |
| --- | --- | --- |
| GitHub repository | `kyungtaak/skill-foundry-quota-increase` | Install from the repository's default branch |
| GitHub Release ZIP | URL of a ZIP attached to a release | Install a published versioned distribution |
| Local directory | Downloaded and extracted skill directory | Inspect files before installing, or register them manually |

The examples below use Windows PowerShell and this [GitHub repository](https://github.com/kyungtaak/skill-foundry-quota-increase). Release installation examples require the `v1.1.0` release and its ZIP asset to be published first. Change both the tag and asset name when installing another version.

#### Install from GitHub

Run this from the root of the project where you want to use the skill. This example targets GitHub Copilot:

```powershell
$repo = "kyungtaak/skill-foundry-quota-increase"
$env:DISABLE_TELEMETRY = "1"

npx skills add $repo --skill foundry-quota-increase --agent github-copilot --copy
```

`--skill` refers to the `name` in [SKILL.md](./SKILL.md), not the directory or repository name. The CLI can discover this skill because the instructions are at the repository root.

Installation is project-local by default. Add `-g` to make the skill available across projects. For Claude Code, use `--agent claude-code`; omit `--agent` to choose interactively.

`--copy` copies files instead of creating symlinks, avoiding dependence on Windows symlink permissions. `DISABLE_TELEMETRY` disables the CLI's usage telemetry and security-audit requests.

#### Install from a GitHub Release ZIP

A ZIP attached under a release's `Assets` can be used directly as the installation source. This downloads the specified release asset rather than the default branch:

```powershell
$env:DISABLE_TELEMETRY = "1"

npx skills add "https://github.com/kyungtaak/skill-foundry-quota-increase/releases/download/v1.1.0/foundry-quota-increase-1.1.0.zip" --agent github-copilot --copy
```

The tag and ZIP filename in the URL must match a published release asset.

Maintainers should attach the ZIP produced by the [release packaging command](#release-zip). It includes instructions, scripts, references, the blank template, README, and license under a `foundry-quota-increase` directory. It is separate from GitHub's automatically generated `Source code (zip)`.

If a direct download requires authentication, such as an asset in a private repository, download it with GitHub CLI first and install from the local directory.

#### Download and install locally

You can download and extract the ZIP from the GitHub Release page. With GitHub CLI, authenticate using `gh auth login` before running the following commands. Private repositories also require repository read access.

```powershell
$repo = "kyungtaak/skill-foundry-quota-increase"
$tag = "v1.1.0"
$asset = "foundry-quota-increase-1.1.0.zip"
$download = Join-Path $env:TEMP ("foundry-quota-" + [guid]::NewGuid().ToString("N"))

New-Item -ItemType Directory -Path $download | Out-Null
gh release download $tag --repo $repo --pattern $asset --dir $download
if ($LASTEXITCODE -ne 0) { throw "Release download failed." }

Expand-Archive -LiteralPath (Join-Path $download $asset) `
  -DestinationPath (Join-Path $download "extracted")

$source = Join-Path $download "extracted\foundry-quota-increase"
$env:DISABLE_TELEMETRY = "1"
npx skills add $source --skill foundry-quota-increase --agent github-copilot --copy
```

`$source` points to the skill directory containing the instructions, not its extracted parent directory. To inspect the files first, review that directory before running the final installation command. The temporary download directory can be removed after installation.

For agents not managed by the CLI, register the entire skill directory in the agent's Agent Skills path. Keep [assets](./assets), [scripts](./scripts), and [references](./references) alongside [SKILL.md](./SKILL.md). In Scout, configure the agent to read instructions and resources from the registered skill's `resourceDir`.

### Verify installation and install Python dependencies

The `skills` CLI installs skill files, not Python dependencies or Playwright browser tools. After a project-local installation for GitHub Copilot:

```powershell
npx skills list --agent github-copilot

Set-Location .\.agents\skills\foundry-quota-increase
python -m pip install -r .\requirements.txt
```

The current CLI uses `.agents\skills` for GitHub Copilot project installations and `$env:USERPROFILE\.copilot\skills` for global installations. If you installed with `-g`, also use `-g` when listing skills and change to the skill directory under the global path. See the [supported agent reference](https://github.com/vercel-labs/skills/blob/main/docs/agents.md) for other agents.

Verify that your agent can discover `foundry-quota-increase` and read its instructions. Connect Playwright tools separately in your agent. The template and local command examples below assume your working directory is the installed skill directory or downloaded repository root.

### Prepare a request workbook

Copy the [blank Excel template](./assets/quota_increase_request_template.xlsx) to a private directory outside the repository before entering request data:

```powershell
$private = Join-Path $env:USERPROFILE "foundry-quota-private"
New-Item -ItemType Directory -Path $private -Force | Out-Null
Copy-Item .\assets\quota_increase_request_template.xlsx `
  -Destination (Join-Path $private "quota-request.xlsx")
```

Fill in the `공통정보` (shared information) and `요청목록` (request list) worksheets in the copy. Do not overwrite the distributed template or change its headers and formulas.

## Usage examples

After registering the skill, ask your agent to perform one of the following tasks. Replace file paths and request details with your own data.

### Fill and capture from Excel

```text
Use the foundry-quota-increase skill to validate the attached workbook.
Fill the live form and capture screenshots for rows marked Y.
Do not submit. Report validation results and reasons for any held requests.
```

### Collect information without Excel

```text
I need a Foundry Global Standard TPM quota increase.
I do not have an Excel workbook. Ask me for contact details, subscription,
model, and target total TPM in small groups.
Only fill and capture the form for now.
```

### Prepare a batch for approval

```text
Prepare the Y rows in this workbook for batch submission.
Show all shared information and each request's model, region, and total TPM.
After I review the destination and private-data warning and explicitly approve,
submit the approved requests one at a time.
Report confirmed submissions separately from uncertain results.
```

These prompts express intent, not approval to send an unseen payload. The agent must show the actual submission preview and wait for the user's confirmation.

## Input format

### Excel

The template uses Korean worksheet names and column headers. Keep those labels unchanged.

In `공통정보`, enter the contact's first and last names, company email, company name and address, subscription ID, and justification. Preserve field names and descriptions; enter values in column B. Use only verified facts in the justification.

In `요청목록`, row 3 contains headers and data starts at row 4. One set of shared information plus one request row forms a single submission.

| Workbook field | Input rules |
| --- | --- |
| 실행(Y/N) | `Y` selects a submission candidate. `N` or blank excludes the row from submission. |
| 요청ID | Use a unique request ID. Do not renumber requests to fill gaps left by deleted rows. |
| Quota Type / Model Type / 배포 유형 | Match the live form selections. Question numbers in the template are reference-only. |
| 폼 모델명(정확한 표기) | Must exactly match a model option in the live form. |
| Region / Data Zone Pool | Select only when the live form displays a region question. |
| 신청 총 TPM(기존+추가) | Target total TPM, including existing quota, not just the increase. Must be a positive integer and a multiple of 1,000. |
| 폼 입력값 kTPM(자동) | The script divides total TPM by 1,000 rather than relying on cached formula results. |
| 구독 ID(행별 지정 시) | A row-specific subscription GUID overrides the shared subscription ID. |

For example, a target of `100,000,000 TPM` becomes `100000 kTPM` in the form. Do not replace the requested total with a published maximum tier or API `availableCapacity`.

### Interactive intake and JSON

Without a workbook, the agent collects intent, contact and company details, subscription, justification, model, total TPM, and any required region. Exact answers are saved as local JSON and processed with `prepare-json`.

See [interactive-intake.json](./references/interactive-intake.json) for the question order and `common` / `requests` structure. Missing contact details, model names, quantities, or justification facts must not be invented.

## Workflow and submission approval

1. **Prepare requests**: Read the latest workbook or JSON, create a new private run directory, and validate required values and supported scope.
2. **Inspect the live form**: Verify its official title and host. Select request type, deployment type, model, conditional region, and quantity in order.
3. **Fill and capture**: Read actual field values back and compare all shared and model-specific values. Save screenshots and validation evidence locally.
4. **Preview the submission**: Show all shared information, request quantities and units, whether a region is asked, held or excluded rows, and the actual destination URL.
5. **Obtain approval**: Show the private-data warning and wait for explicit confirmation. Batch approval covers only the displayed `Y` rows.
6. **Submit sequentially and record results**: Use a fresh form for each request and submit only once. Record success only after verifying an explicit completion message and a completion screenshot.

The default capture mode stops after step 3. Actual submissions require a separate `submit` mode run. Check existing execution records before starting another run to prevent duplicate submissions.

Changes to approved models, quantities, shared information, or whether a region question appears require fresh approval. If the approved batch scope is unchanged, do not repeat the same approval question for each request.

A timeout or disconnected session with an uncertain outcome is recorded as `unknown`, and the batch stops. Never automatically retry requests that are `submitting`, `unknown`, or already submitted. Reconciliation requires existing evidence and the user's confirmation of actual receipt.

## Local commands

The following validates input and checks execution status without running a browser. Run it from the skill or repository root after preparing your workbook. Use a new run directory for each execution.

```powershell
$private = Join-Path $env:USERPROFILE "foundry-quota-private"
$workbook = Join-Path $private "quota-request.xlsx"
$runDir = Join-Path $private "capture-001"

python .\scripts\quota_workbook.py prepare `
  --workbook $workbook `
  --run-dir $runDir `
  --mode capture

python .\scripts\quota_workbook.py status --run-dir $runDir
```

For JSON input, use `prepare-json --input` instead of `prepare`. In capture mode, add `--include-disabled` to include `N` rows in an input test. Use `--ids` to select specific requests.

| Command | Purpose |
| --- | --- |
| `prepare` / `prepare-json` | Validate input and create a new execution record |
| `render --action fill` | Generate a fill-and-capture function for Playwright |
| `record-ui` | Save input verification evidence returned by the browser |
| `preview` | Show submission scope, shared information, and the approval fingerprint |
| `approve` | Record the user's actual confirmation text and fingerprint |
| `begin-submit` | Create a durable submission attempt and single-request ticket |
| `render --action submit` | Generate a submission function that checks approval and the ticket |
| `finish-submit` | Validate completion evidence and record the outcome |
| `mark-unknown` | Record an uncertain submission outcome |
| `reconcile-submitted` | Reconcile an uncertain result using existing evidence and user confirmation |
| `status` | Show execution mode, request counts, and state totals |

Running `approve` is not itself user consent. Never fabricate confirmation text or browser evidence. See [SKILL.md](./SKILL.md) for full command syntax and submission rules.

## Repository layout

| Path | Description |
| --- | --- |
| [SKILL.md](./SKILL.md) | Skill metadata and agent operating instructions |
| [VERSION](./VERSION) | Distribution version |
| [LICENSE](./LICENSE) | MIT license |
| [requirements.txt](./requirements.txt) | Runtime dependencies |
| [requirements-dev.txt](./requirements-dev.txt) | Development and test dependencies |
| [assets/quota_increase_request_template.xlsx](./assets/quota_increase_request_template.xlsx) | Blank Excel request template |
| [references/form-map.json](./references/form-map.json) | Reference form fields and selections |
| [references/interactive-intake.json](./references/interactive-intake.json) | Intake questions and JSON structure |
| [scripts/quota_workbook.py](./scripts/quota_workbook.py) | Validation, approval and submission records, and browser function generation |
| [scripts/browser_flow.js](./scripts/browser_flow.js) | Playwright form filling, verification, and submission |
| [scripts/make_blank_template.py](./scripts/make_blank_template.py) | Template creation with customer data and private metadata removed |
| [scripts/build_release.py](./scripts/build_release.py) | Allowlist-based release packaging |
| [tests](./tests) | Regression tests using synthetic inputs and forms |
| [.github/workflows/tests.yml](./.github/workflows/tests.yml) | GitHub Actions test configuration |

## Development and testing

### Python tests

```powershell
python -m pip install -r .\requirements-dev.txt
python -m pytest tests -q
```

GitHub Actions runs synthetic pytest tests with Python 3.12 and Node.js 22. It does not access the live application form or submit requests.

### Browser regression checks

```powershell
$private = Join-Path $env:USERPROFILE "foundry-quota-private"
python .\tests\make_browser_smoke.py `
  --output (Join-Path $private "browser-smoke.js") `
  --captures (Join-Path $private "smoke-captures")
```

Execute the generated function with your agent's Playwright tools. It uses a separate browser context, intercepts network traffic, and serves synthetic forms only. No requests are submitted to the actual Microsoft form.

Passing synthetic tests does not guarantee current model or region options in the live form, or a successful real submission. Check the live form for changes during each execution.

### Release ZIP

```powershell
python .\scripts\build_release.py --output .\foundry-quota-increase-1.1.0.zip
```

The packager includes only allowlisted reusable files under `foundry-quota-increase`, including the README and MIT license. It refuses to overwrite an existing ZIP. Private request files and screenshots are not included. When changing versions, update both [VERSION](./VERSION) and the frontmatter in [SKILL.md](./SKILL.md).

## Privacy and execution records

Contact details, company addresses, subscription IDs, and justifications are private data. Keep original request files, execution manifests, confirmation text, tickets, screenshots, HAR files, and receipt evidence in a private local directory outside the repository. Generated browser functions can contain request values and must not be uploaded elsewhere.

[.gitignore](./.gitignore) helps prevent accidental inclusion of request data and browser output, but it is not a substitute for review. Check actual file contents and Excel metadata before committing or distributing files.

During an approved submission, the previewed information is sent to Microsoft's official form. The agent must show the destination URL and private-data warning before asking for approval. It must not modify or close the user's unrelated browser tabs.

## License

Licensed under the [MIT License](./LICENSE). Copyright (c) 2026 kyungtaak.

## References

- [Official Microsoft quota increase request form](https://aka.ms/aoai/quotaincrease)
- [Agent Skills specification](https://agentskills.io/)
- [skills CLI and requirements](https://github.com/vercel-labs/skills)
- [skills CLI installation sources and options](https://github.com/vercel-labs/skills/blob/main/docs/commands.md#skills-add)
- [GitHub CLI release download reference](https://cli.github.com/manual/gh_release_download)

The README structure follows the introduction, installation, repository overview, and development patterns of these public projects. Feature descriptions and commands are based on this repository's implementation.

- [anthropics/skills](https://github.com/anthropics/skills)
- [github/awesome-copilot](https://github.com/github/awesome-copilot)
