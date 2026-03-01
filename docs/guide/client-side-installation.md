# Client-side installation

## Configure .npmrc for GitHub Packages

This package is published to GitHub Packages. Before installing, configure your
`.npmrc` file (in your project root or home directory) with:

```
@g2i-ai:registry=https://npm.pkg.github.com
//npm.pkg.github.com/:_authToken=${GITHUB_TOKEN}
```

## Install with npm or Yarn

You can install the latest version of HyperFormula with popular
packaging managers. Navigate to your project folder and run the
following command:

**npm:**

```bash
$ npm install @g2i-ai/hyperformula
```

**Yarn:**

```bash
$ yarn add @g2i-ai/hyperformula
```

The package will be added to your `package.json` file and installed to
the `./node_modules` directory.

Then you can import it into your file like this:

```javascript
import { HyperFormula } from '@g2i-ai/hyperformula';

// your code
```

## Use CDN

::: warning
CDN delivery (jsDelivr, unpkg) is **not available** for packages published to
GitHub Packages. You must install the package via npm/Yarn as described above,
or self-host the UMD bundle files from the `dist/` directory after building.
:::

<!--
The following CDN URLs were used with the original npm-published package and
will NOT work with the @g2i-ai/hyperformula GitHub Packages release:

<script src="https://cdn.jsdelivr.net/npm/hyperformula/dist/hyperformula.full.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/hyperformula/dist/hyperformula.min.js"></script>
-->

You can read more about the dependencies of HyperFormula on a dedicated [Dependencies](/guide/dependencies.md) page.

## Clone from GitHub

If you choose to clone the project or download it from GitHub you
will need to build it prior to usage. Check the
[building section](building.md) for a full list of commands and their
descriptions.

### Clone with HTTPS

```bash
git clone https://github.com/g2i-ai/hyperformula.git
```

### Clone with SSH

```bash
git clone git@github.com:g2i-ai/hyperformula.git
```

## Download from GitHub

You can download all resources as a ZIP archive directly from the
[GitHub repository](https://github.com/g2i-ai/hyperformula).
Then, you can use one of the above-mentioned methods to install the
library.
