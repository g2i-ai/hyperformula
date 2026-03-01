# Design: Publish @g2i-ai/hyperformula to GitHub Packages

**Date:** 2026-03-01
**Status:** Implemented

## Context

HyperFormula is a fork of `handsontable/hyperformula` maintained at `g2i-ai/hyperformula`. We need to publish it as `@g2i-ai/hyperformula` via GitHub Packages so it can be installed via npm from consumer projects.

## Decisions

### Package scope: `@g2i-ai/hyperformula`
GitHub Packages requires the scope to match the GitHub org.

### Versioning: Independent semver starting at 4.0.0
Clearly distinguishes from upstream 3.x without implying a greenfield project.

### Release trigger: Git tags (`v*`)
- Development happens on feature branches, merged to `master`
- CI (build, test, lint) runs on push to master
- To release: bump version, commit, push tag `v4.0.0`
- GitHub Actions triggers on tag, validates version match, builds, publishes, creates GitHub Release
- No release branches needed — tags are immutable release pointers

### Documentation: Update all existing docs
All references to the `hyperformula` npm package updated to `@g2i-ai/hyperformula`. CDN sections note that jsdelivr/unpkg are not available for GitHub Packages.

## Changes made

### Build & config
- `package.json`: name → `@g2i-ai/hyperformula`, version → `4.0.0`, repository → `g2i-ai/hyperformula`, check:licenses updated
- `.config/webpack/languages.js`: externals updated to `@g2i-ai/hyperformula`
- `.github/workflows/npm-publish.yml`: replaced manual dispatch with tag-triggered release workflow

### Documentation
- `README.md`: badges, install commands, import examples, .npmrc instructions
- `docs/guide/client-side-installation.md`: .npmrc setup, install commands, CDN unavailability note
- `docs/guide/server-side-installation.md`: .npmrc setup, install commands
- `docs/guide/*.md`: all import/require paths updated
- `docs/examples/**/*.{js,ts}`: all import paths updated
- `docs/.vuepress/config.js`: CDN URL comments
- `docs/index.md`: badges, install commands
- `docs/guide/releasing.md`: new release process documentation
