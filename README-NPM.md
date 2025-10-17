# DOCX Templater

Advanced DOCX template engine with support for placeholders, loops, conditionals, tables, and images. Built with TypeScript and designed for modern Node.js applications.

[![npm version](https://badge.fury.io/js/%40abdelmonimsamadi%2Fdocx-templater.svg)](https://badge.fury.io/js/%40abdelmonimsamadi%2Fdocx-templater)
[![TypeScript](https://img.shields.io/badge/TypeScript-Ready-blue.svg)](https://www.typescriptlang.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

## Features

- ✅ **Placeholders**: Replace `{name}` with data values
- ✅ **Loops**: Repeat content with `{#array}...{/array}`
- ✅ **Conditionals**: Show/hide content with `{?condition}...{/condition}`
- ✅ **If-Else**: Conditional branching with `{?condition}...{:else}...{/condition}`
- ✅ **Tables**: Generate table rows with `{table:arrayName}`
- ✅ **Images**: Embed images with size control
- ✅ **Buffer-based**: No file system dependencies
- ✅ **TypeScript**: Full type safety and intellisense
- ✅ **Statistics**: Optional detailed processing stats

## Installation

```bash
npm install @abdelmonimsamadi/docx-templater
```

## Quick Start

```typescript
import { generateDocx, TemplateData } from "@abdelmonimsamadi/docx-templater";
import { readFileSync, writeFileSync } from "fs";

const templateBuffer = readFileSync("template.docx");
const imageBuffer = readFileSync("signature.png");

const data: TemplateData = {
  name: "John Doe",
  date: "2024-10-17",
  items: [
    { product: "Laptop", price: "$999" },
    { product: "Mouse", price: "$25" },
  ],
  signature: {
    type: "image",
    buffer: imageBuffer,
    extension: "png",
    widthInches: 3,
  },
};

const outputBuffer = await generateDocx(templateBuffer, data);
writeFileSync("output.docx", outputBuffer);
```

## Template Syntax

### 1. Basic Placeholders

Replace simple values:

```
Template: Hello {name}, today is {date}
Data: { name: "John", date: "2024-10-17" }
Result: Hello John, today is 2024-10-17
```

### 2. Loops

Repeat content for arrays:

```
Template:
{#items}
- {product}: {price}
{/items}

Data: {
  items: [
    { product: "Laptop", price: "$999" },
    { product: "Mouse", price: "$25" }
  ]
}

Result:
- Laptop: $999
- Mouse: $25
```

### 3. Conditionals

Simple conditions:

```
Template: {?isVip}🌟 VIP Customer{/isVip}
Data: { isVip: true }
Result: 🌟 VIP Customer
```

If-else conditions:

```
Template: {?hasDiscount}Sale Price: {salePrice}{:else}Regular Price: {price}{/hasDiscount}
Data: { hasDiscount: false, price: "$99" }
Result: Regular Price: $99
```

### 4. Tables

Generate table rows from arrays:

1. Create a table in your DOCX template
2. Place `{table:arrayName}` in the first column of the row to repeat
3. Add field placeholders `{fieldName}` in other columns

```
Template table row:
| {table:employees} | {name} | {department} | {salary} |

Data: {
  employees: [
    { name: "John", department: "IT", salary: "$75000" },
    { name: "Jane", department: "HR", salary: "$65000" }
  ]
}

Result:
| (empty) | John | IT | $75000 |
| (empty) | Jane | HR | $65000 |
```

### 5. Images

Embed images with optional sizing:

```typescript
const data = {
  logo: {
    type: "image",
    buffer: imageBuffer,
    extension: "png",
    widthInches: 3, // Optional: width in inches
    heightInches: 2, // Optional: height in inches
  },
};
```

Template: `{logo}`

**Size Options:**

- No size specified: Auto-scale to max 6 inches, maintain aspect ratio
- Only width: Calculate height maintaining aspect ratio
- Only height: Calculate width maintaining aspect ratio
- Both width and height: Exact dimensions (may distort image)

## API Reference

### `generateDocx(templateBuffer, data, options?)`

Main function to generate DOCX from template and data.

**Parameters:**

- `templateBuffer: Buffer` - DOCX template file as Buffer
- `data: TemplateData` - Template data object
- `options?: GenerateDocxOptions` - Optional configuration

**Returns:** `Promise<Buffer>` - Generated DOCX as Buffer

### `generateDocxDetailed(templateBuffer, data, options?)`

Enhanced version that returns detailed statistics.

**Returns:** `Promise<DocxGenerationResult>` - Result with buffer and stats

```typescript
const result = await generateDocxDetailed(templateBuffer, data);
console.log("Generated DOCX with stats:", result.stats);
// {
//   placeholdersReplaced: 15,
//   loopsProcessed: 2,
//   conditionalsProcessed: 3,
//   tablesGenerated: 1,
//   imagesEmbedded: 2
// }
```

## License

MIT License - see [LICENSE](LICENSE) file for details.
