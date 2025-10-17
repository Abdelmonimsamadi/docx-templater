import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "xmldom";
import sizeOf from "image-size";

/**
 * Normalize DOCX text by merging split text runs to handle placeholders properly
 */
function normalizeDocxText(xmlString: string) {
  // This function merges adjacent <w:t> elements that might split placeholders
  const xmlDoc = new DOMParser().parseFromString(xmlString, "text/xml");
  const textNodes = xmlDoc.getElementsByTagName("w:t");

  // Extract all text content and merge it
  let fullText = "";
  const textElements = [];

  for (let i = 0; i < textNodes.length; i++) {
    const textNode = textNodes[i];
    const text = textNode.textContent || "";
    fullText += text;
    textElements.push({
      element: textNode,
      text: text,
      startIndex: fullText.length - text.length,
    });
  }

  // Find placeholders in the full text
  const placeholderRegex = /{[^}]+}/g;
  let match;
  const replacements = [];

  while ((match = placeholderRegex.exec(fullText)) !== null) {
    replacements.push({
      placeholder: match[0],
      start: match.index,
      end: match.index + match[0].length,
    });
  }

  // For each placeholder, clear affected text nodes and put the full placeholder in the first one
  replacements.forEach((replacement) => {
    let placeholderSet = false;

    textElements.forEach((textEl) => {
      const elStart = textEl.startIndex;
      const elEnd = textEl.startIndex + textEl.text.length;

      // If this element overlaps with the placeholder
      if (elStart < replacement.end && elEnd > replacement.start) {
        if (!placeholderSet) {
          // Set the complete placeholder in the first overlapping element
          textEl.element.textContent = replacement.placeholder;
          placeholderSet = true;
        } else {
          // Clear other overlapping elements
          textEl.element.textContent = "";
        }
      }
    });
  });

  return new XMLSerializer().serializeToString(xmlDoc);
}

/**
 * Replace placeholders {tag} in a DOCX with data values, loop arrays, conditionals, tables, and embed images.
 *
 * Supported features:
 *
 * Basic placeholders: {name} - Replaced with data.name
 *
 * Loops: {#arrayName}...{/arrayName} - Repeats content for each item in array
 *
 * Conditionals:
 * - {?condition}content{/condition} - Shows content if condition is truthy
 * - {?condition}if content{:else}else content{/condition} - Shows if/else content based on condition
 *
 * Tables: {table:arrayName} - Place in a table row, generates rows for each array item
 *
 * Images:
 * - { type: "image", buffer: Buffer.from(...), extension: "jpg" } - Auto-scale to max 6 inches, maintain aspect ratio
 * - { type: "image", buffer: Buffer.from(...), extension: "png", widthInches: 3 } - Set width to 3", auto-calculate height
 * - { type: "image", buffer: Buffer.from(...), extension: "jpg", heightInches: 2 } - Set height to 2", auto-calculate width
 * - { type: "image", buffer: Buffer.from(...), extension: "png", widthInches: 4, heightInches: 3 } - Exact 4"x3" (may distort)
 */
export async function generateDocx(
  templateBuffer: Buffer,
  data: any
): Promise<Buffer> {
  const zip = new JSZip();
  const doc = await zip.loadAsync(templateBuffer);

  const documentXml = await doc.file("word/document.xml")!.async("text");
  const xmlDoc = new DOMParser().parseFromString(documentXml, "text/xml");
  const serializer = new XMLSerializer();

  // First, we need to normalize the text to handle split placeholders
  let xmlString = normalizeDocxText(documentXml);

  // Handle loops: {#array}...{/array}
  xmlString = xmlString.replace(
    /{#(\w+)}([\s\S]*?){\/\1}/g,
    (match, key, inner) => {
      const arr = data[key];
      if (!arr) {
        console.warn(`Warning: Array '${key}' not found in data`);
        return "";
      }
      if (!Array.isArray(arr)) {
        console.warn(`Warning: '${key}' is not an array, got:`, typeof arr);
        return "";
      }
      if (arr.length === 0) {
        console.warn(`Warning: Array '${key}' is empty`);
        return "";
      }

      return arr
        .filter((item) => item != null) // Filter out null/undefined items
        .map((item) =>
          inner.replace(/{(\w+)}/g, (_, k) => {
            const value = item[k];
            return value != null ? String(value) : "";
          })
        )
        .join("");
    }
  );

  // Handle conditional statements: {?condition}...{/condition} and {?condition}...{:else}...{/condition}
  xmlString = xmlString.replace(
    /{\?(\w+)}([\s\S]*?)(?:{:else}([\s\S]*?))?{\/\1}/g,
    (match, key, ifContent, elseContent = "") => {
      const condition = data[key];

      // Evaluate condition - truthy values render the if content, falsy values render else content
      if (
        condition &&
        condition !== "false" &&
        condition !== "0" &&
        condition !== "" &&
        !(Array.isArray(condition) && condition.length === 0)
      ) {
        return ifContent;
      } else {
        return elseContent;
      }
    }
  );

  // Handle table generation: {table:arrayName}
  xmlString = xmlString.replace(/{table:(\w+)}/g, (match, key) => {
    const arr = data[key];
    if (!Array.isArray(arr) || arr.length === 0) {
      console.warn(`Warning: Table data '${key}' not found or empty`);
      return "";
    }

    // Find the table row template (should be the current row containing the {table:arrayName} placeholder)
    // We'll look for the containing <w:tr> element and duplicate it for each array item
    const tableRowRegex =
      /(<w:tr[^>]*>)([\s\S]*?{table:\w+}[\s\S]*?)(<\/w:tr>)/;
    const xmlBeforeTable = xmlString.substring(0, xmlString.indexOf(match));
    const xmlAfterTable = xmlString.substring(
      xmlString.indexOf(match) + match.length
    );

    // Find the table row that contains this placeholder
    const beforeTableReversed = xmlBeforeTable.split("").reverse().join("");
    const trStartMatch = beforeTableReversed.match(/>rt:w<([^<]*)/);
    const trEndMatch = xmlAfterTable.match(/(<\/w:tr>)/);

    if (!trStartMatch || !trEndMatch) {
      console.warn(
        `Warning: Could not find table row structure for {table:${key}}`
      );
      return match; // Return original if we can't find the table structure
    }

    // Extract the full table row
    const trStartIndex = xmlString.lastIndexOf(
      "<w:tr",
      xmlString.indexOf(match)
    );
    const trEndIndex =
      xmlString.indexOf("</w:tr>", xmlString.indexOf(match)) + 7;
    const rowTemplate = xmlString.substring(trStartIndex, trEndIndex);

    // Generate rows for each array item
    const generatedRows = arr
      .filter((item) => item != null)
      .map((item) => {
        let row = rowTemplate.replace(/{table:\w+}/, ""); // Remove the table placeholder

        // Replace placeholders in this row with item data
        row = row.replace(/{(\w+)}/g, (_, k) => {
          const value = item[k];
          return value != null ? String(value) : "";
        });

        return row;
      })
      .join("");

    // Replace the original row with generated rows in the full XML
    xmlString =
      xmlString.substring(0, trStartIndex) +
      generatedRows +
      xmlString.substring(trEndIndex);

    return ""; // Return empty since we've already replaced in xmlString
  });

  // Handle simple replacements: {name}
  xmlString = xmlString.replace(/{(\w+)}/g, (_, key) => {
    if (typeof data[key] === "string") return data[key];
    if (data[key] && data[key].type === "image") return `__IMAGE__${key}__`;
    return "";
  });

  // Replace the XML content
  doc.file("word/document.xml", xmlString);

  // Embed images if any
  if (Object.values(data).some((v) => v?.type === "image")) {
    const relsPath = "word/_rels/document.xml.rels";
    let relsXml = await doc.file(relsPath)!.async("text");
    const relsDoc = new DOMParser().parseFromString(relsXml, "text/xml");

    let relIdCounter = 100;
    for (const [key, value] of Object.entries(data)) {
      if (value?.type !== "image") continue;

      const imgBuffer = value.buffer;
      const dims = sizeOf(imgBuffer);

      const imgFile = `word/media/${key}.${value.extension}`;
      doc.file(imgFile, imgBuffer);

      // Add to relationships
      const relElem = relsDoc.createElement("Relationship");
      const rId = `rId${relIdCounter++}`;
      relElem.setAttribute("Id", rId);
      relElem.setAttribute(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      );
      relElem.setAttribute("Target", `media/${key}.${value.extension}`);
      relsDoc.documentElement.appendChild(relElem);

      // Replace image marker
      let docXml = await doc.file("word/document.xml")!.async("text");
      docXml = docXml.replace(
        `__IMAGE__${key}__`,
        createImageXML(rId, dims, value)
      );
      doc.file("word/document.xml", docXml);
    }

    relsXml = serializer.serializeToString(relsDoc);
    doc.file(relsPath, relsXml);

    // Update [Content_Types].xml to include image types
    const contentTypesPath = "[Content_Types].xml";
    let contentTypesXml = await doc.file(contentTypesPath)!.async("text");
    const contentTypesDoc = new DOMParser().parseFromString(
      contentTypesXml,
      "text/xml"
    );

    // Add image extensions if not already present
    const imageExtensions = ["jpg", "jpeg", "png", "gif"];
    imageExtensions.forEach((ext) => {
      if (!contentTypesXml.includes(`Extension="${ext}"`)) {
        const defaultElem = contentTypesDoc.createElement("Default");
        defaultElem.setAttribute("Extension", ext);
        defaultElem.setAttribute(
          "ContentType",
          `image/${ext === "jpg" ? "jpeg" : ext}`
        );
        contentTypesDoc.documentElement.appendChild(defaultElem);
      }
    });

    contentTypesXml = serializer.serializeToString(contentTypesDoc);
    doc.file(contentTypesPath, contentTypesXml);
  }

  const outputBuffer = await doc.generateAsync({ type: "nodebuffer" });
  return outputBuffer as Buffer;
}

function createImageXML(
  rId: string,
  dims: any,
  imageData: { widthInches?: number; heightInches?: number } = {}
) {
  // Convert inches to EMUs (914400 EMUs per inch)
  const inchesToEMU = 914400;

  let cx, cy;

  // If custom width/height specified in inches, use those
  if (imageData.widthInches || imageData.heightInches) {
    if (imageData.widthInches && imageData.heightInches) {
      // Both specified - use exact dimensions
      cx = imageData.widthInches * inchesToEMU;
      cy = imageData.heightInches * inchesToEMU;
    } else if (imageData.widthInches) {
      // Only width specified - maintain aspect ratio
      const aspectRatio = dims.height / dims.width;
      cx = imageData.widthInches * inchesToEMU;
      cy = cx * aspectRatio;
    } else {
      // Only height specified - maintain aspect ratio
      const aspectRatio = dims.width / dims.height;
      cy = imageData.heightInches * inchesToEMU;
      cx = cy * aspectRatio;
    }
  } else {
    // No custom size - use original logic with scaling
    const maxWidthTwips = 6 * inchesToEMU; // 6 inches max
    const maxHeightTwips = 6 * inchesToEMU;

    cx = dims.width * 9525; // Convert pixels to EMUs
    cy = dims.height * 9525;

    // Scale down if too large
    if (cx > maxWidthTwips) {
      const ratio = maxWidthTwips / cx;
      cx = maxWidthTwips;
      cy = cy * ratio;
    }
    if (cy > maxHeightTwips) {
      const ratio = maxHeightTwips / cy;
      cy = maxHeightTwips;
      cx = cx * ratio;
    }
  }

  return `<w:r>
      <w:drawing>
        <wp:inline distT="0" distB="0" distL="0" distR="0">
          <wp:extent cx="${Math.round(cx)}" cy="${Math.round(cy)}"/>
          <wp:effectExtent l="0" t="0" r="0" b="0"/>
          <wp:docPr id="1" name="Picture"/>
          <wp:cNvGraphicFramePr>
            <a:graphicFrameLocks xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" noChangeAspect="1"/>
          </wp:cNvGraphicFramePr>
          <a:graphic xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
            <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">
              <pic:pic xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture">
                <pic:nvPicPr>
                  <pic:cNvPr id="1" name="Picture"/>
                  <pic:cNvPicPr/>
                </pic:nvPicPr>
                <pic:blipFill>
                  <a:blip r:embed="${rId}"/>
                  <a:stretch>
                    <a:fillRect/>
                  </a:stretch>
                </pic:blipFill>
                <pic:spPr>
                  <a:xfrm>
                    <a:off x="0" y="0"/>
                    <a:ext cx="${Math.round(cx)}" cy="${Math.round(cy)}"/>
                  </a:xfrm>
                  <a:prstGeom prst="rect">
                    <a:avLst/>
                  </a:prstGeom>
                </pic:spPr>
              </pic:pic>
            </a:graphicData>
          </a:graphic>
        </wp:inline>
      </w:drawing>
    </w:r>`;
}
