# Plan: Fix Skipped Release Job in Build-Release Workflow

## Problem Analysis
The user reports that the `release` job in `.github/workflows/build-release.yml` is skipped when triggered via the GitHub Web UI (manual trigger), even though the build succeeds.

**Current Configuration:**
```yaml
release:
  if: startsWith(github.ref, 'refs/tags/v')
```

**Root Cause:**
*   When a workflow is triggered manually (via `workflow_dispatch`), the `github.ref` context variable typically points to the branch ref (e.g., `refs/heads/master`), NOT a tag ref.
*   Therefore, `startsWith('refs/heads/master', 'refs/tags/v')` evaluates to `false`, causing the `release` job to be skipped.

## Proposed Changes

### 1. Update Job Condition
Modify the `if` condition in the `release` job to allow execution when triggered manually.

**Change:**
```yaml
# From
if: startsWith(github.ref, 'refs/tags/v')

# To
if: startsWith(github.ref, 'refs/tags/v') || github.event_name == 'workflow_dispatch'
```

### 2. Handle Release Versioning for Manual Triggers
The `softprops/action-gh-release` action typically relies on the git tag to name the release. When running on a branch:
*   We need to ensure it doesn't fail or create a badly named release.
*   **Recommendation**: When manually triggered, we should probably set `name` and `tag_name` explicitly or accept that it might create a "draft" release or use the branch name.
*   However, for a quick fix that allows the job to *run*, simply allowing the condition is the first step.
*   Better approach: Add an input to `workflow_dispatch` for the tag/version, OR just let it run and see if the action handles it (it might default to standard behavior or fail if no tag is found).
*   *Refinement*: To be safe and "fix thoroughly", we will add a `tag_name` input to `workflow_dispatch` so the user can specify the tag manually if they want, OR we can default to `refs/tags/v...` only if it exists.
*   *Simpler approach for now*: Just enable the job. If it lacks a tag, `softprops/action-gh-release` might warn or use the commit SHA.

**Refined Strategy:**
Enable `workflow_dispatch` execution.
Note: If `softprops/action-gh-release` is run on a branch without `tag_name` specified, it might create a release named after the tag *if* the commit is tagged, or fail.
To be robust, we will treat manual runs as "Draft Releases" or ensure the user knows to tag first?
Actually, the user said "Run on web... doesn't create release". They probably want to *force* a release implementation.

**Decision**:
I will update the condition to allow `workflow_dispatch`.

## Verification Plan

### Automated Tests
*   I will use a script to validate the YAML syntax.

### Manual Verification
*   User will trigger the workflow manually on GitHub.
*   Check if the `release` job starts.

## User Review Required
> [!IMPORTANT]
> Enabling `release` on `workflow_dispatch` on a BRANCH (not a tag) might create a release pointing to the branch tip.
> Do you want to successfully create a release *artifact* even if not tagged? Or do you only want to release *tagged* commits?
> If you only want tagged commits, you should create the tag **before** running the workflow, or run the workflow **on the tag**.
> If you are running it on `master` branch, it is **correct** that it skips release if the intention is "Release only tags".
> **Hypothesis**: You want to test the release process OR you want to release the current 'master' build often.

**Assumption:** I will assume you want the release job to RUN when you click the button, even if it's just a draft.

