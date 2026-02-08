# Winget Integration

This app publishes to [Windows Package Manager](https://github.com/microsoft/winget-pkgs) (winget) automatically.

## How It Works

The release pipeline is fully automated:

```
Push tag v0.0.3
  → build.yml: build, sign, create GitHub release
  → build.yml (publish-winget job): submit manifest to winget-pkgs
  → Microsoft validates & merges
  → Users can run: winget install ScottHanselman.MaximizeToVirtualDesktop
```

Every tag push triggers the full pipeline. No manual steps after initial setup.

## One-Time Setup

Two things are needed before the first automated publish:

### 1. Create the `WINGET_TOKEN` secret

The winget-releaser action needs a GitHub Personal Access Token (PAT) to create pull requests against the [winget-pkgs](https://github.com/microsoft/winget-pkgs) repository.

1. Go to [GitHub Settings → Developer settings → Personal access tokens → Tokens (classic)](https://github.com/settings/tokens)
2. Generate new token with **`public_repo`** scope
3. Go to this repository's **Settings → Secrets and variables → Actions**
4. Create a new secret named **`WINGET_TOKEN`** and paste the token

### 2. Submit the initial package version

The first version must be submitted manually (winget-releaser handles updates, not new packages):

```powershell
# Install the tool
winget install Microsoft.WingetCreate

# Create manifest from the latest release
wingetcreate new https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.2/MaximizeToVirtualDesktop-v0.0.2-win-x64.zip
```

When prompted, use these values:

| Field | Value |
|-------|-------|
| Package Identifier | `ScottHanselman.MaximizeToVirtualDesktop` |
| Publisher | `Scott Hanselman` |
| Package Name | `MaximizeToVirtualDesktop` |
| License | `MIT` |
| Short Description | `Maximize windows to virtual desktops like macOS` |

Then submit:

```powershell
wingetcreate submit <path-to-generated-manifest>
```

After Microsoft approves the initial submission (typically 1-3 days), all subsequent versions are handled automatically by the pipeline.

## Manual Re-submission

If a winget submission fails or needs to be re-triggered, use the manual workflow:

**Actions → Publish to Winget (Manual) → Run workflow** → enter the version number (e.g., `0.0.3`).

Or via CLI:

```powershell
wingetcreate update ScottHanselman.MaximizeToVirtualDesktop `
  --version 0.0.3 `
  --urls https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.3/MaximizeToVirtualDesktop-v0.0.3-win-x64.zip `
         https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.3/MaximizeToVirtualDesktop-v0.0.3-win-arm64.zip `
  --submit
```

## Resources

- [winget-pkgs Repository](https://github.com/microsoft/winget-pkgs)
- [WingetCreate Tool](https://github.com/microsoft/winget-create)
- [winget-releaser Action](https://github.com/vedantmgoyal2009/winget-releaser)
- [Manifest Schema](https://learn.microsoft.com/en-us/windows/package-manager/package/manifest)
