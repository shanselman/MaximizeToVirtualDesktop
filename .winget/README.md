# Winget Submission Files

This directory contains reference files for submitting MaximizeToVirtualDesktop to the Windows Package Manager.

## Files

- **manifest-template.yaml** — Sample manifest structure showing how the winget package will be configured
  - This is a *reference only*
  - The actual manifest should be created using `wingetcreate` tool
  - SHA256 hashes must be calculated by wingetcreate or manually
  
## Usage

**Do not submit this template directly.** Instead:

1. Follow the instructions in [WINGET.md](../WINGET.md)
2. Use `wingetcreate` to generate the actual manifest with correct hashes
3. Submit via `wingetcreate submit` or manual PR to [winget-pkgs](https://github.com/microsoft/winget-pkgs)

## Automated Publishing

The GitHub Actions workflow at `.github/workflows/winget-publish.yml` can automate submissions on new releases.

To enable it:
1. Create a GitHub Personal Access Token with `public_repo` scope
2. Add it as a repository secret named `WINGET_TOKEN`
3. The workflow will run automatically on new releases
