# DOCX Templater - Feature Guide

## Overview

This DOCX templater supports advanced features including conditionals, loops, tables, and image embedding.

## Features

### 1. Basic Placeholders

```
Template: Hello {name}, today is {date}
Data: { name: "John", date: "2024-10-17" }
Result: Hello John, today is 2024-10-17
```

### 2. Loops

Repeat content for each item in an array:

```
Template: {#employees}Employee: {name} - {department}{/employees}
Data: {
  employees: [
    { name: "John", department: "IT" },
    { name: "Jane", department: "HR" }
  ]
}
Result: Employee: John - IT
        Employee: Jane - HR
```

### 3. Conditionals

#### Simple Conditional

```
Template: {?isApproved}✅ Document approved{/isApproved}
Data: { isApproved: true }
Result: ✅ Document approved
```

#### If-Else Conditional

```
Template: {?hasErrors}❌ Errors found{:else}✅ No errors{/hasErrors}
Data: { hasErrors: false }
Result: ✅ No errors
```

#### Condition Evaluation

Conditions are considered "truthy" if:

- Boolean `true`
- Non-empty strings (except "false", "0", "")
- Non-zero numbers
- Non-empty arrays
- Non-null objects

### 4. Tables

Generate table rows from array data:

1. Create a table in your DOCX template
2. In the row where you want data repeated, place `{table:arrayName}`
3. Add placeholders for the data fields: `{field1}`, `{field2}`, etc.

```
Template table row: | {table:products} | {name} | {price} | {category} |
Data: {
  products: [
    { name: "Laptop", price: "$999", category: "Electronics" },
    { name: "Book", price: "$25", category: "Education" }
  ]
}
Result: Two table rows with the data filled in
```

### 5. Images

Embed images with size control:

```javascript
{
  signature: {
    type: "image",
    buffer: imageBuffer,      // Buffer containing image data
    extension: "png",         // File extension
    widthInches: 3,          // Optional: width in inches
    heightInches: 2          // Optional: height in inches
  }
}
```

## Usage Example

```typescript
import { generateDocx } from "./utils";
import { readFileSync, writeFileSync } from "fs";

const templateBuffer = readFileSync("template.docx");
const imageBuffer = readFileSync("signature.png");

const data = {
  // Basic data
  companyName: "Acme Corp",
  date: "2024-10-17",

  // Conditional data
  isUrgent: true,
  hasAttachments: false,

  // Array for loops
  items: [
    { name: "Item 1", status: "Complete" },
    { name: "Item 2", status: "Pending" },
  ],

  // Array for tables
  employees: [
    { name: "John", dept: "IT", salary: "50000" },
    { name: "Jane", dept: "HR", salary: "45000" },
  ],

  // Image
  logo: {
    type: "image",
    buffer: imageBuffer,
    extension: "png",
    widthInches: 2,
  },
};

const outputBuffer = await generateDocx(templateBuffer, data);
writeFileSync("output.docx", outputBuffer);
```

## Template Examples

### Document with all features:

```
{companyName} - {date}

{?isUrgent}🚨 URGENT DOCUMENT{/isUrgent}

Items:
{#items}
- {name}: {status}
{/items}

Employee Table:
| Name | Department | Salary |
|------|------------|--------|
| {table:employees} | {name} | {dept} | {salary} |

{?hasAttachments}📎 Attachments included{:else}No attachments{/hasAttachments}

Signature: {logo}
```

## Tips

1. **Loops**: Use `{#arrayName}...{/arrayName}` for repeating content
2. **Conditionals**: Use `{?condition}...{/condition}` or `{?condition}...{:else}...{/condition}`
3. **Tables**: Place `{table:arrayName}` in the first column of the row you want to repeat
4. **Images**: Always specify the correct file extension
5. **Nested structures**: You can nest loops and conditionals
