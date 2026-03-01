# Releasing

This guide describes how to publish a new version of `@g2i-ai/hyperformula` to GitHub Packages.

## Prerequisites

- Push access to the `g2i-ai/hyperformula` repository
- The `master` branch must be in a releasable state (CI passing)

## Release process

### 1. Update the version

On the `master` branch (or a release prep branch), bump the version in `package.json`:

```bash
npm version <major|minor|patch> --no-git-tag-version
```

For example, to release `4.1.0`:

```bash
npm version minor --no-git-tag-version
```

### 2. Update the changelog

Add a new section to `CHANGELOG.md` under the `[Unreleased]` header:

```markdown
## [4.1.0] - YYYY-MM-DD

### Added
- ...

### Fixed
- ...
```

Move the relevant entries from `[Unreleased]` into the new version section.

### 3. Update the license checker

Update the `check:licenses` script in `package.json` to reference the new version:

```json
"check:licenses": "license-checker --production --excludePackages=\"@g2i-ai/hyperformula@4.1.0\" ..."
```

### 4. Commit and tag

```bash
git add package.json CHANGELOG.md
git commit -m "release: v4.1.0"
git tag v4.1.0
```

### 5. Push

```bash
git push origin master --tags
```

This triggers the GitHub Actions workflow which will:

1. Validate the tag version matches `package.json`
2. Install dependencies and build all bundles
3. Publish the package to GitHub Packages
4. Create a GitHub Release with auto-generated notes

### 6. Verify

- Check the [Actions tab](https://github.com/g2i-ai/hyperformula/actions) for workflow status
- Verify the package appears in [GitHub Packages](https://github.com/orgs/g2i-ai/packages)

## Versioning

This package uses [Semantic Versioning](https://semver.org/):

- **Major** (`X.0.0`): Breaking API changes
- **Minor** (`x.Y.0`): New features, backward-compatible
- **Patch** (`x.y.Z`): Bug fixes, backward-compatible

The package started at version `4.0.0` to distinguish from the upstream `hyperformula` package (which is at `3.x`).

## Troubleshooting

### Tag version mismatch

If the workflow fails with "Tag version does not match package.json version", ensure the git tag exactly matches the version in `package.json`. For example, tag `v4.1.0` requires `"version": "4.1.0"` in `package.json`.

### Republishing a version

GitHub Packages does not allow overwriting a published version. If you need to fix a release, bump to the next patch version and release again.
