import { generateDocx } from "./utils";
import { readFileSync, writeFileSync } from "fs";

async function test() {
  // Read template as buffer
  const templateBuffer = readFileSync("fiche.docx");

  // Read image as buffer
  const imageBuffer = readFileSync("signature.png");

  // Example usage with new buffer-based API including conditionals and tables
  const outputBuffer = await generateDocx(templateBuffer, {
    date: "2024-06-10",
    objet: "Demande de demande",
    reference: "Message N° 12345",

    // Array for loops
    references: [
      {
        valeur: "123",
        entite: "TEST/TEST",
      },
      {
        valeur: "435",
        entite: "TEST/TEST13213",
      },
    ],

    // Data for table example
    employees: [
      { name: "John Doe", position: "Developer", salary: "50000" },
      { name: "Jane Smith", position: "Designer", salary: "45000" },
      { name: "Bob Wilson", position: "Manager", salary: "60000" },
    ],

    // Conditional data
    condition: true,

    // Large text content
    contenu:
      "Nostrud laborum sint ex est duis culpa ea proident ea magna sit officia aliqua proident aute. Exercitation exercitation anim eu. Excepteur dolor commodo nulla excepteur duis do labore eiusmod commodo cupidatat ea. Sit culpa incididunt commodo minim culpa enim adipisicing excepteur aliqua do. Fugiat officia incididunt nostrud magna nulla fugiat. Irure esse laboris adipisicing excepteur incididunt commodo irure cupidatat duis magna voluptate fugiat exercitation. Et esse est proident voluptate nisi nostrud sit ipsum aliquip sint laborum minim ipsum aliqua.",

    // Image
    signature: {
      type: "image",
      buffer: imageBuffer,
      extension: "png",
      widthInches: 3,
    },
  });

  // Write output buffer to file
  writeFileSync("output.docx", outputBuffer);
  console.log("✅ Generated: output.docx");
}

test();
