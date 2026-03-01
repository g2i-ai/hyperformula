# Server-side installation

::: tip
For full compatibility, the minimum required version of **Node is 13**.
It is related to the support for ICU. There is a possibility to use
lower versions of Node but you need to install an additional package
as the dependency: [`full-icu`](https://github.com/unicode-org/full-icu-npm)
:::

The basic steps are very similar to the ones in the
[client-side installation](client-side-installation.md) process.

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

The package will be added to your `package.json` file and installed in
the `./node_modules` directory.

Then you can just `require` it:

```javascript
const { HyperFormula } = require('@g2i-ai/hyperformula');

// your code
```
