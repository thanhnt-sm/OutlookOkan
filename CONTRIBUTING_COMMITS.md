# Conventional Commits Guide for OutlookOkan

This project uses **Conventional Commits** for automatic versioning and changelog generation.

## Commit Message Format

```
<type>[optional scope]: <description>

[optional body]

[optional footer(s)]
```

## Types and Version Impact

| Type | Description | Version Bump |
|------|-------------|--------------|
| `feat:` | New feature | **Minor** (0.X.0) |
| `fix:` | Bug fix | **Patch** (0.0.X) |
| `docs:` | Documentation only | No bump |
| `style:` | Formatting, no code change | No bump |
| `refactor:` | Code refactoring | No bump |
| `perf:` | Performance improvement | Patch |
| `test:` | Adding tests | No bump |
| `build:` | Build system changes | No bump |
| `ci:` | CI configuration | No bump |
| `chore:` | Other changes | No bump |

## Breaking Changes (Major Version)

```bash
feat!: change API response format

# OR

feat: update login system

BREAKING CHANGE: password field is now encrypted by default
```

## Examples

```bash
# Feature (Minor bump: 1.0.0 → 1.1.0)
feat: add email confirmation dialog

# Bug fix (Patch bump: 1.1.0 → 1.1.1)
fix: resolve null reference in CheckList generation

# With scope
feat(ui): add dark mode support
fix(email): correct recipient validation logic

# Breaking change (Major bump: 1.1.1 → 2.0.0)
feat!: redesign settings storage format

# No version bump
docs: update installation guide
chore: update dependencies
refactor: simplify email parser logic
```

## Scope Examples

| Scope | Description |
|-------|-------------|
| `ui` | User interface changes |
| `email` | Email processing |
| `settings` | Configuration/settings |
| `ribbon` | Outlook ribbon |
| `checklist` | Check list functionality |
| `i18n` | Internationalization |

## Force Version Bump

Use `+semver:` in commit message:

```bash
chore: update config +semver:patch
chore: prepare release +semver:minor
refactor!: complete overhaul +semver:major
```

## Skip Release

```bash
docs: fix typo +semver:skip
```

## Quick Reference

```bash
# Start with the type
git commit -m "feat: your feature description"
git commit -m "fix: your bug fix description"

# Add scope for clarity
git commit -m "feat(email): add CC field validation"

# Breaking change
git commit -m "feat!: change configuration format"
```

---

**Learn more:** [Conventional Commits Specification](https://www.conventionalcommits.org/)
