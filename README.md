<br>
<p align="center">
  <img src="https://raw.githubusercontent.com/handsontable/hyperformula/master/github-hf-logo-blue.svg" width="350" height="71" alt="HyperFormula - A headless spreadsheet, a parser and evaluator of Excel formulas"/>
</p>

<p align="center">
  <strong>An open-source headless spreadsheet for business web apps</strong>
</p>

<p align="center">
  <a href="https://github.com/g2i-ai/hyperformula/packages"><img src="https://img.shields.io/github/v/release/g2i-ai/hyperformula" alt="GitHub release"></a>
  <a href="https://github.com/g2i-ai/hyperformula/actions?query=workflow%3Abuild+branch%3Amaster"><img src="https://img.shields.io/github/actions/workflow/status/g2i-ai/hyperformula/build.yml?branch=master" alt="GitHub Workflow Status"></a>
</p>

---

> **Note:** This is a fork of [HyperFormula](https://github.com/handsontable/hyperformula) published to GitHub Packages as `@g2i-ai/hyperformula`. To install from GitHub Packages, configure your `.npmrc`:
> ```
> @g2i-ai:registry=https://npm.pkg.github.com
> //npm.pkg.github.com/:_authToken=${GITHUB_TOKEN}
> ```

HyperFormula is a headless spreadsheet built in TypeScript, serving as both a parser and evaluator of spreadsheet formulas. It can be integrated into your browser or utilized as a service with Node.js as your back-end technology.

## What HyperFormula can be used for?

HyperFormula doesn't assume any existing user interface, making it a general-purpose library that can be used in various business applications. Here are some examples:

- Deterministic compute layer for AI & LLMs
- Calculated fields in CRM and ERP software
- Custom spreadsheet-like app
- Business logic builder
- Forms and form builder
- Educational app
- Online calculator

## Features

- [Function syntax compatible with Microsoft Excel](docs/guide/compatibility-with-microsoft-excel.md) and [Google Sheets](docs/guide/compatibility-with-google-sheets.md)
- High-speed parsing and evaluation of spreadsheet formulas
- [A library of ~400 built-in functions](docs/guide/built-in-functions.md)
- [Support for custom functions](docs/guide/custom-functions.md)
- [Support for Node.js](docs/guide/server-side-installation.md#install-with-npm-or-yarn)
- [Support for undo/redo](docs/guide/undo-redo.md)
- [Support for CRUD operations](docs/guide/basic-operations.md)
- [Support for clipboard](docs/guide/clipboard-operations.md)
- [Support for named expressions](docs/guide/named-expressions.md)
- [Support for data sorting](docs/guide/sorting-data.md)
- [Support for formula localization with 17 built-in languages](docs/guide/i18n-features.md)
- Easy integration with any front-end or back-end application
- GPLv3 licensed (see [Licensing](docs/guide/licensing.md))

## Documentation

- [Client-side installation](docs/guide/client-side-installation.md)
- [Server-side installation](docs/guide/server-side-installation.md)
- [Basic usage](docs/guide/basic-usage.md)
- [Configuration options](docs/guide/configuration-options.md)
- [List of built-in functions](docs/guide/built-in-functions.md)
- [API Reference](https://g2i-ai.github.io/hyperformula/api/)

## Integrations

- [Integration with React](docs/guide/integration-with-react.md#demo)
- [Integration with Angular](docs/guide/integration-with-angular.md#demo)
- [Integration with Vue](docs/guide/integration-with-vue.md#demo)
- [Integration with Svelte](docs/guide/integration-with-svelte.md#demo)

## Installation and usage

Install the library from [GitHub Packages](https://github.com/g2i-ai/hyperformula/packages) like so:

```bash
npm install @g2i-ai/hyperformula
```

Once installed, you can use it to develop applications tailored to your specific business needs. Here, we've used it to craft a form that calculates mortgage payments using the `PMT` formula.

```js
import { HyperFormula } from '@g2i-ai/hyperformula';

// Create a HyperFormula instance
const hf = HyperFormula.buildEmpty({ licenseKey: 'gpl-v3' });

// Add an empty sheet
const sheetName = hf.addSheet('Mortgage Calculator');
const sheetId = hf.getSheetId(sheetName);

// Enter the mortgage parameters
hf.addNamedExpression('AnnualInterestRate', '8%');
hf.addNamedExpression('NumberOfMonths', 360);
hf.addNamedExpression('LoanAmount', 800000);

// Use the PMT function to calculate the monthly payment
hf.setCellContents({ sheet: sheetId, row: 0, col: 0 }, [['Monthly Payment', '=PMT(AnnualInterestRate/12, NumberOfMonths, -LoanAmount)']]);

// Display the result
console.log(`${hf.getCellValue({ sheet: sheetId, row: 0, col: 0 })}: ${hf.getCellValue({ sheet: sheetId, row: 0, col: 1 })}`);
```

## Contributing

Contributions are welcome. Please read the [Contributing Guide](docs/guide/contributing.md) before submitting a pull request.

## License

This fork is available under the [GPLv3 license](LICENSE.txt). Originally developed by [Handsoncode](https://handsontable.com/).
