# Project: OutlookOkan
# Tech Stack: C# .NET, Outlook Add-in (VSTO)

## Architecture
- Visual Studio Solution: `OutlookOkan.sln`
- Main project: `OutlookOkan/`
- Tests: `OutlookOkanTest/`
- Setup/Installer: `SetupCustomAction/`
- Build scripts: `build.ps1` (PowerShell), `build.sh` (Bash)

## Coding Conventions
- C# .NET Framework / VSTO conventions
- Follow existing code style in `OutlookOkan/` project
- NuGet packages managed via `packages/` directory
- Version tracked in `version` file

## Building
- Windows: `.\build.ps1`
- MSBuild path: see `msbuild_path.txt`

## Testing
- Test project: `OutlookOkanTest/`

## Important Files
- `OutlookOkan.sln` — solution file
- `build.ps1` — build script
- `version` — version number
- `docs/` — documentation
