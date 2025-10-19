import { exec } from "child_process";
import {
  generateDocx,
  generateDocxDetailed,
  TemplateData,
  ImageData,
} from "./index";
import { readFileSync, unlinkSync, writeFileSync } from "fs";
import path from "path";

const __filename = path.basename(import.meta.url);
const __dirname = path.dirname(__filename);

unlinkSync(path.join(__dirname, "examples/output.docx"));

async function test() {
  // Read template as buffer
  const templateBuffer = readFileSync(
    path.join(__dirname, "examples/template.docx")
  );

  // Read image as buffer
  const imageBuffer = readFileSync(path.join(__dirname, "examples/image.jpg"));

  // Example usage with new buffer-based API including conditionals and tables
  const templateData: TemplateData = {
    name: "John Doe",
    date: new Date().toLocaleDateString(),

    // Array for loops
    array: [
      {
        value1: "Item 1",
        value2: "Description for item 1",
      },
      {
        value1: "Item 2",
        value2: "Description for item 2",
      },
    ],

    // Data for table example
    employees: [
      { name: "John Doe", position: "Developer", salary: "50000" },
      { name: "Jane Smith", position: "Designer", salary: "45000" },
      { name: "Bob Wilson", position: "Manager", salary: "60000" },
      { name: "Alice Johnson", position: "Senior Developer", salary: "65000" },
      { name: "Mike Brown", position: "QA Engineer", salary: "48000" },
    ],

    // Conditional data
    condition1: true,
    condition2: false,

    // Image
    image: {
      type: "image",
      buffer: imageBuffer,
      //   url: "https://planetofnames.biz/cdn/shop/products/SimpleDesign.jpg?v=1666923538&width=1946",
      extension: "png",
      widthInches: 5,
    } as ImageData,
  };

  // Generate the document
  const outputBuffer = await generateDocx(templateBuffer, templateData);

  // Write output buffer to file
  writeFileSync(path.join(__dirname, "examples/output.docx"), outputBuffer);

  console.log("✅ Generated: output.docx");

  exec("open " + path.join(__dirname, "examples/output.docx"));
}

test();
