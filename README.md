# React-XLSX-wrapper

[![Npm-version][npm-shield]][npm-url]
[![Forks][forks-shield]][forks-url]
[![Stargazers][stars-shield]][stars-url]
[![Issues][issues-shield]][issues-url]

A data export library built with and for [React](https://react.dev/), [Next.js](https://nextjs.org/)

### Note
> Please don't use this package with [2.0.0 - 2.0.7] versions. It has major issues. Either Stick to 1.1.5 or move to 2.0.8 and so on.

## Installation
With [npm](https://www.npmjs.org/package/react-xlsx-wrapper):

```sh
npm install --save react-xlsx-wrapper@latest
```

```sh
pnpm add react-xlsx-wrapper@latest
```

```sh
yarn add react-xlsx-wrapper@latest
```

## Framework Support

### Next.js (App Router)

This library is a **client-only** component — it uses browser APIs (`Blob`, `URL.createObjectURL`) to trigger downloads. Wrap it with `'use client'` in your component:

```tsx
'use client';

import { ExcelFile, ExcelSheet, ExcelColumn } from 'react-xlsx-wrapper';

export function ExportButton({ data }: { data: any[] }) {
  return (
    <ExcelFile filename="report" element={<button>Download</button>}>
      <ExcelSheet name="Sheet1" data={data}>
        <ExcelColumn label="Name" value="name" />
        <ExcelColumn label="Email" value="email" />
      </ExcelSheet>
    </ExcelFile>
  );
}
```

If the parent is a Server Component, lazy-load to avoid SSR:

```tsx
import dynamic from 'next/dynamic';

const ExportButton = dynamic(() => import('./ExportButton'), { ssr: false });
```

> **Note:** Server-side Excel generation (e.g. from an API route) is not supported. The library is designed for browser downloads only.

### Vite

Works out of the box with Vite via the ESM entry point (`dist/index.mjs`). No additional config needed.

If you encounter CJS interop warnings with an older version of this package, add to `vite.config.ts`:

```ts
optimizeDeps: { include: ['react-xlsx-wrapper', 'xlsx-js-style'] }
```

### TanStack (Router / Query / Table)

Fully compatible. Pass filtered/sorted rows from TanStack Table directly as the `data` prop:

```tsx
'use client'; // if using Next.js App Router

import { useReactTable, getCoreRowModel } from '@tanstack/react-table';
import { ExcelFile, ExcelSheet, ExcelColumn } from 'react-xlsx-wrapper';

export function DataTable({ data }) {
  const table = useReactTable({ data, columns, getCoreRowModel: getCoreRowModel() });

  return (
    <>
      {/* your table UI */}
      <ExcelFile filename="export" element={<button>Export</button>}>
        <ExcelSheet name="Data" data={table.getFilteredRowModel().rows.map(r => r.original)}>
          <ExcelColumn label="Name" value="name" />
        </ExcelSheet>
      </ExcelFile>
    </>
  );
}
```

## Code Examples

- [Simple Excel Export](examples/simple_excel_export_01.md)
- [Excel Export with Dataset](examples/simple_excel_export_02.md)
- [Excel Export with custom cells style](examples/styled_excel_sheet.md)

## Excel Props

| Prop          | Type                | Default    | Required | Description                                       |
| :------------ | :------------------ | :--------- | :------- | :------------------------------------------------ |
| hideElement   | `bool`              | false      | `false`  | To hide the button & directly download excel file |
| filename      | `string`            | Download   | `false`  | Excel file name to be downloaded                  |
| fileExtension | `string`            | xlsx       | `false`  | Download file extension [xlsx]                    |
| element       | `HTMLElement`       | `<button>` | `false`  | Element to download excel file                    |
| children      | `React.ReactElement<ExcelSheetProps>` | `null`     | `true`   | ExcelSheet Represents data                        |

### ExcelSheet Props

| Prop     | Type                    | Default | Required | Description        |
| :------- | :---------------------- | :------ | :------- | :----------------- |
| name     | `string`                | `""`    | `true`   | Sheet name in file |
| bigHeading | ` ExcelSheetCol` |`undefined`|`false`| Big Merged Cell Heading
| autoFilterForAllColumn | `boolean` | `false` | `false` | Auto Filter Generated Based on Colums
| data     | `any[]`         | `null`  | `false`  | Excel Sheet data   |
| dataSet  | `ExcelSheetData[]` | `null`  | `false`  | Excel Sheet data   |
| children | `ExcelColumn`           | `null`  | `false`  | ExcelColumns       |

**Note:** In ExcelSheet props `dataSet` has `precedence` over `data` and `children` props.

For further types and definitions [Read More](types/types.md)

## Cell Style

Cell styles are specified by a style object that roughly parallels the OpenXML structure. The style object has five
top-level attributes: `fill`, `font`, `numFmt`, `alignment`, and `border`.

| Style Attribute | Sub Attributes | Values                                                                                        |
| :-------------- | :------------- | :-------------------------------------------------------------------------------------------- |
| fill            | patternType    | `"solid"` or `"none"`                                                                         |
|                 | fgColor        | `COLOR_SPEC`                                                                                  |
|                 | bgColor        | `COLOR_SPEC`                                                                                  |
| font            | name           | `"Calibri"` // default                                                                        |
|                 | sz             | `11` // font size in points                                                                 |
|                 | color          | `COLOR_SPEC`                                                                                  |
|                 | bold           | `true` or `false`                                                                             |
|                 | underline      | `true` or `false`                                                                             |
|                 | italic         | `true` or `false`                                                                             |
|                 | strike         | `true` or `false`                                                                             |
|                 | outline        | `true` or `false`                                                                             |
|                 | shadow         | `true` or `false`                                                                             |
|                 | vertAlign      | `true` or `false`                                                                             |
| numFmt          |                | `"0"` // integer index to built in formats, see StyleBuilder.SSF property                     |
|                 |                | `"0.00%"` // string matching a built-in format, see StyleBuilder.SSF                          |
|                 |                | `"0.0%"` // string specifying a custom format                                                 |
|                 |                | `"0.00%;\\(0.00%\\);\\-;@"` // string specifying a custom format, escaping special characters |
|                 |                | `"m/dd/yy"` // string a date format using Excel's format notation                             |
| alignment       | vertical       | `"bottom"` or `"center"` or `"top"`                                                           |
|                 | horizontal     | `"bottom"` or `"center"` or `"top"`                                                           |
|                 | wrapText       | `true ` or ` false`                                                                           |
|                 | readingOrder   | `2` // for right-to-left                                                                      |
|                 | textRotation   | Number from `0` to `180` or `255` (default is `0`)                                            |
|                 |                | `90` is rotated up 90 degrees                                                                 |
|                 |                | `45` is rotated up 45 degrees                                                                 |
|                 |                | `135` is rotated down 45 degrees                                                              |
|                 |                | `180` is rotated down 180 degrees                                                             |
|                 |                | `255` is special, aligned vertically                                                          |
| border          | top            | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | bottom         | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | left           | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | right          | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | diagonal       | `{ style: BORDER_STYLE, color: COLOR_SPEC }`                                                  |
|                 | diagonalUp     | `true` or `false`                                                                             |
|                 | diagonalDown   | `true` or `false`                                                                             |

**COLOR_SPEC**: Colors for `fill`, `font`, and `border` are specified as objects, either:

- `{ auto: 1}` specifying automatic values
- `{ rgb: "FFFFAA00" }` specifying a hex ARGB value
- `{ theme: "1", tint: "-0.25"}` specifying an integer index to a theme color and a tint value (default 0)
- `{ indexed: 64}` default value for `fill.bgColor`

**BORDER_STYLE**: Border style is a string value which may take on one of the following values:

- `thin`
- `medium`
- `thick`
- `dotted`
- `hair`
- `dashed`
- `mediumDashed`
- `dashDot`
- `mediumDashDot`
- `dashDotDot`
- `mediumDashDotDot`
- `slantDashDot`

Borders for merged areas are specified for each cell within the merged area. So to apply a box border to a merged area of 3x3 cells, border styles would need to be specified for eight different cells:

- left borders for the three cells on the left,
- right borders for the cells on the right
- top borders for the cells on the top
- bottom borders for the cells on the left

[npm-shield]: https://img.shields.io/npm/v/react-xlsx-wrapper?style=for-the-badge&labelColor=44567bf&color=%23007bff
[npm-url]: https://www.npmjs.com/package/react-xlsx-wrapper
[forks-shield]: https://img.shields.io/github/forks/AS-Devs/react-xlsx-wrapper?color=44567bf&style=for-the-badge
[forks-url]: https://github.com/AS-Devs/react-xlsx-wrapper/network/members
[stars-shield]: https://img.shields.io/github/stars/AS-Devs/react-xlsx-wrapper?color=%23007bff&style=for-the-badge
[stars-url]: https://github.com/AS-Devs/react-xlsx-wrapper/stargazers
[issues-shield]: https://img.shields.io/github/issues/AS-Devs/react-xlsx-wrapper?style=for-the-badge
[issues-url]: https://github.com/AS-Devs/react-xlsx-wrapper/issues
