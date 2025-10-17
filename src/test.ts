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
      { name: "Alice Johnson", position: "Senior Developer", salary: "65000" },
      { name: "Mike Brown", position: "QA Engineer", salary: "48000" },
      { name: "Sarah Davis", position: "Product Manager", salary: "70000" },
      { name: "Tom Anderson", position: "DevOps Engineer", salary: "58000" },
      { name: "Lisa Garcia", position: "UI/UX Designer", salary: "52000" },
      { name: "David Miller", position: "Backend Developer", salary: "55000" },
      { name: "Emma Wilson", position: "Data Analyst", salary: "47000" },
      { name: "Chris Taylor", position: "Technical Lead", salary: "75000" },
      { name: "Rachel White", position: "Frontend Developer", salary: "53000" },
      {
        name: "Kevin Martinez",
        position: "System Administrator",
        salary: "51000",
      },
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
