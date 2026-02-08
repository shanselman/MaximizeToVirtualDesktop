# Adding MaximizeToVirtualDesktop to Winget

This guide explains how to submit MaximizeToVirtualDesktop to the [Windows Package Manager](https://github.com/microsoft/winget-pkgs) (winget).

## Prerequisites

Before submitting to winget, ensure:

1. **A stable release is published** on GitHub with downloadable installers/executables
2. **Installers are code-signed** (already done via Azure Trusted Signing in our build workflow)
3. **Release artifacts are accessible via direct URLs** (GitHub releases provide this)

## Step 1: Install WingetCreate

The easiest way to create a winget manifest is using Microsoft's official tool:

```powershell
winget install Microsoft.WingetCreate
```

Or download from: https://github.com/microsoft/winget-create/releases

## Step 2: Create the Manifest

Run the following command to generate a manifest for a new package:

```powershell
wingetcreate new https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.2/MaximizeToVirtualDesktop-v0.0.2-win-x64.zip
```

**Important**: Replace `v0.0.2` with the actual version number you're submitting.

The tool will prompt you for information. Here's what to provide:

| Prompt | Answer |
|--------|--------|
| **Package Identifier** | `ScottHanselman.MaximizeToVirtualDesktop` |
| **Package Version** | (match the tag, e.g., `0.0.2`) |
| **Publisher** | `Scott Hanselman` |
| **Package Name** | `MaximizeToVirtualDesktop` |
| **License** | `MIT` |
| **Short Description** | `Maximize windows to virtual desktops like macOS` |
| **Installer Type** | `zip` |
| **Architecture** | `x64` (then add ARM64 separately) |

### Adding Multiple Architectures

After creating the initial manifest, add the ARM64 installer:

```powershell
wingetcreate update ScottHanselman.MaximizeToVirtualDesktop --version 0.0.2 --urls https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.2/MaximizeToVirtualDesktop-v0.0.2-win-arm64.zip --architecture arm64
```

## Step 3: Review the Manifest

The manifest will be created in YAML format. Here's a sample structure:

```yaml
PackageIdentifier: ScottHanselman.MaximizeToVirtualDesktop
PackageVersion: 0.0.2
PackageName: MaximizeToVirtualDesktop
Publisher: Scott Hanselman
License: MIT
ShortDescription: Maximize windows to virtual desktops like macOS
Installers:
  - Architecture: x64
    InstallerType: zip
    InstallerUrl: https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.2/MaximizeToVirtualDesktop-v0.0.2-win-x64.zip
    InstallerSha256: [auto-generated]
  - Architecture: arm64
    InstallerType: zip
    InstallerUrl: https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.2/MaximizeToVirtualDesktop-v0.0.2-win-arm64.zip
    InstallerSha256: [auto-generated]
ManifestType: singleton
ManifestVersion: 1.0.0
```

**Note**: WingetCreate automatically calculates SHA256 hashes for you.

## Step 4: Validate the Manifest

Before submitting, validate the manifest:

```powershell
winget validate <path-to-manifest-folder>
```

## Step 5: Test in Sandbox (Recommended)

Test the installation in a clean environment:

```powershell
# Clone the winget-pkgs repo
git clone https://github.com/microsoft/winget-pkgs

# Test with Windows Sandbox
.\winget-pkgs\Tools\SandboxTest.ps1 <path-to-your-manifest-folder>
```

This launches a Windows Sandbox instance and tests the installation.

## Step 6: Submit to winget-pkgs

### Option A: Use WingetCreate to Submit (Easiest)

```powershell
wingetcreate submit <path-to-manifest-folder>
```

This will:
1. Fork the winget-pkgs repository (if you haven't already)
2. Create a new branch
3. Commit your manifest
4. Open a pull request

**Note**: You'll need a GitHub token with `public_repo` scope. The tool will prompt for it.

### Option B: Manual Submission

1. **Fork** the [winget-pkgs repository](https://github.com/microsoft/winget-pkgs)
2. **Clone** your fork
3. **Copy** your manifest to the correct path:
   ```
   manifests/s/ScottHanselman/MaximizeToVirtualDesktop/<version>/
   ```
4. **Commit** and push to a new branch
5. **Open a pull request** against the main repository

## Step 7: Wait for Validation

The winget-pkgs repository has automated validation checks:

- Schema validation
- Hash verification
- Installer scanning (SmartScreen, VirusTotal)
- Manual review by Microsoft

This typically takes 1-3 days. Once approved, the package will be available via:

```powershell
winget install ScottHanselman.MaximizeToVirtualDesktop
```

## Updating Future Versions

For each new release, update the manifest:

```powershell
wingetcreate update ScottHanselman.MaximizeToVirtualDesktop \
  --version 0.0.3 \
  --urls https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.3/MaximizeToVirtualDesktop-v0.0.3-win-x64.zip \
          https://github.com/shanselman/MaximizeToVirtualDesktop/releases/download/v0.0.3/MaximizeToVirtualDesktop-v0.0.3-win-arm64.zip \
  --submit
```

## Automating Winget Submissions

To automate manifest updates on release, consider using:

- **[winget-releaser GitHub Action](https://github.com/vedantmgoyal2009/winget-releaser)** — automatically creates/updates manifests on new GitHub releases

Example workflow (`.github/workflows/winget-publish.yml`):

```yaml
name: Publish to Winget

on:
  release:
    types: [published]

jobs:
  publish:
    runs-on: windows-latest
    steps:
      - uses: vedantmgoyal2009/winget-releaser@v2
        with:
          identifier: ScottHanselman.MaximizeToVirtualDesktop
          token: ${{ secrets.GITHUB_TOKEN }}
```

## Important Notes

### Portable vs Installer

Currently, this app is distributed as a ZIP file (portable). Winget prefers installers (MSI, EXE). Options:

1. **Keep as ZIP** — works fine, but users must manually extract
2. **Create an installer** — better UX. Consider using:
   - [Inno Setup](https://jrsoftware.org/isinfo.php) (free, scriptable)
   - [WiX Toolset](https://wixtoolset.org/) (MSI builder)
   - [Advanced Installer](https://www.advancedinstaller.com/) (commercial, has free tier)

### Code Signing

Our releases are already code-signed via Azure Trusted Signing. This helps with:
- SmartScreen reputation
- Winget validation (unsigned installers may be rejected)

## Troubleshooting

### "Package already exists"

If someone else has already submitted the package, you can:
- Contact them to transfer ownership
- Use a different Package Identifier (e.g., `Hanselman.MaximizeToVirtualDesktop`)

### Validation Failures

Common issues:
- **Incorrect SHA256** — re-run wingetcreate to recalculate hashes
- **Invalid installer URL** — ensure the GitHub release is public and URLs are correct
- **Installer fails SmartScreen** — code-signing helps; may need to build reputation

### GitHub Token Permissions

For automated submission, create a Personal Access Token with:
- `public_repo` scope (for public repositories)
- `workflow` scope (if using GitHub Actions)

## Resources

- [Winget Documentation](https://learn.microsoft.com/en-us/windows/package-manager/)
- [winget-pkgs Repository](https://github.com/microsoft/winget-pkgs)
- [WingetCreate Tool](https://github.com/microsoft/winget-create)
- [Manifest Schema Reference](https://learn.microsoft.com/en-us/windows/package-manager/package/manifest)
- [Package Submission Guide](https://learn.microsoft.com/en-us/windows/package-manager/package/repository)

## Quick Reference

```powershell
# Install WingetCreate
winget install Microsoft.WingetCreate

# Create manifest for new package
wingetcreate new <installer-url>

# Update existing package
wingetcreate update <PackageIdentifier> --version <new-version> --urls <url1> <url2> --submit

# Validate manifest
winget validate <manifest-path>

# Test installation locally
winget install --manifest <manifest-path>
```
