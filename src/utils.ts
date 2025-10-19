import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "xmldom";
import {
  TemplateData,
  ImageData,
  GenerateDocxOptions,
  DocxGenerationResult,
  DocxGenerationError,
} from "./types";

// Cross-platform image size detection
function getImageSize(buffer: Buffer | Uint8Array): {
  width: number;
  height: number;
} {
  // Convert to Uint8Array for consistent handling
  const bytes = buffer instanceof Buffer ? new Uint8Array(buffer) : buffer;

  // PNG signature: 89 50 4E 47 0D 0A 1A 0A
  if (
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47
  ) {
    // PNG format - width and height are at bytes 16-19 and 20-23 (big-endian)
    const width =
      (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | bytes[19];
    const height =
      (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | bytes[23];
    return { width, height };
  }

  // JPEG signature: FF D8
  if (bytes[0] === 0xff && bytes[1] === 0xd8) {
    // JPEG format - scan for SOF markers
    let i = 2;
    while (i < bytes.length - 8) {
      if (bytes[i] === 0xff) {
        const marker = bytes[i + 1];
        // SOF0, SOF1, SOF2 markers (Start of Frame)
        if (marker >= 0xc0 && marker <= 0xc3) {
          const height = (bytes[i + 5] << 8) | bytes[i + 6];
          const width = (bytes[i + 7] << 8) | bytes[i + 8];
          return { width, height };
        }
        // Skip to next marker
        const length = (bytes[i + 2] << 8) | bytes[i + 3];
        i += length + 2;
      } else {
        i++;
      }
    }
  }

  // GIF signature: 47 49 46 38 (GIF8)
  if (
    bytes[0] === 0x47 &&
    bytes[1] === 0x49 &&
    bytes[2] === 0x46 &&
    bytes[3] === 0x38
  ) {
    // GIF format - width at bytes 6-7, height at bytes 8-9 (little-endian)
    const width = bytes[6] | (bytes[7] << 8);
    const height = bytes[8] | (bytes[9] << 8);
    return { width, height };
  }

  // Default fallback if format not recognized
  return { width: 100, height: 100 };
}

// Cross-platform buffer handling
function ensureBuffer(data: Buffer | Uint8Array | ArrayBuffer): Buffer {
  if (typeof Buffer !== "undefined" && Buffer.isBuffer(data)) {
    return data;
  }
  if (data instanceof ArrayBuffer) {
    return typeof Buffer !== "undefined"
      ? Buffer.from(data)
      : (new Uint8Array(data) as any);
  }
  if (data instanceof Uint8Array) {
    return typeof Buffer !== "undefined" ? Buffer.from(data) : (data as any);
  }
  return data as Buffer;
}

/**
 * Normalize DOCX text by merging split text runs to handle placeholders properly
 */
function normalizeDocxText(xmlString: string) {
  // This function merges adjacent <w:t> elements that might split placeholders
  const xmlDoc = new DOMParser().parseFromString(xmlString, "text/xml");
  const textNodes = xmlDoc.getElementsByTagName("w:t");

  // Extract all text content and merge it
  let fullText = "";
  const textElements: Array<{
    element: Element;
    text: string;
    startIndex: number;
  }> = [];

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
  const replacements: Array<{
    placeholder: string;
    start: number;
    end: number;
  }> = [];

  while ((match = placeholderRegex.exec(fullText)) !== null) {
    replacements.push({
      placeholder: match[0],
      start: match.index,
      end: match.index + match[0].length,
    });
  }

  // Create a map of elements to their replacement content
  const elementReplacements = new Map<Element, string>();

  // Process each element to build its final content
  textElements.forEach((textEl) => {
    const elStart = textEl.startIndex;
    const elEnd = textEl.startIndex + textEl.text.length;

    // Find all placeholders that overlap with this element
    const overlappingPlaceholders = replacements.filter(
      (replacement) => elStart < replacement.end && elEnd > replacement.start
    );

    if (overlappingPlaceholders.length === 0) {
      // No placeholders affect this element, keep original content
      return;
    }

    // Build the content for this element
    let elementContent = "";
    let currentPos = elStart;

    // Sort overlapping placeholders by start position
    overlappingPlaceholders.sort((a, b) => a.start - b.start);

    for (const replacement of overlappingPlaceholders) {
      // Add any text before this placeholder (within this element's range)
      const beforeStart = Math.max(currentPos, elStart);
      const beforeEnd = Math.min(replacement.start, elEnd);
      if (beforeStart < beforeEnd) {
        elementContent += fullText.substring(beforeStart, beforeEnd);
      }

      // Add the placeholder (only if it starts within this element)
      if (replacement.start >= elStart && replacement.start < elEnd) {
        elementContent += replacement.placeholder;
      }

      // Move current position past this placeholder
      currentPos = Math.max(currentPos, replacement.end);
    }

    // Add any remaining text after the last placeholder
    if (currentPos < elEnd) {
      elementContent += fullText.substring(currentPos, elEnd);
    }

    // Store the replacement for this element
    elementReplacements.set(textEl.element, elementContent);
  });

  // Apply all replacements
  elementReplacements.forEach((content, element) => {
    element.textContent = content;
  });

  // Clear elements that had overlapping placeholders but aren't the primary element
  replacements.forEach((replacement) => {
    let primaryElementSet = false;

    textElements.forEach((textEl) => {
      const elStart = textEl.startIndex;
      const elEnd = textEl.startIndex + textEl.text.length;

      if (elStart < replacement.end && elEnd > replacement.start) {
        if (elementReplacements.has(textEl.element) && !primaryElementSet) {
          // This is the primary element for this placeholder range, keep its content
          primaryElementSet = true;
        } else if (!elementReplacements.has(textEl.element)) {
          // This element was affected but doesn't have replacement content, clear it
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
  templateBuffer: Buffer | Uint8Array | ArrayBuffer,
  data: TemplateData
): Promise<Buffer | Uint8Array> {
  const zip = new JSZip();
  const doc = await zip.loadAsync(ensureBuffer(templateBuffer));

  const documentXml = await doc.file("word/document.xml")!.async("text");
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
          inner.replace(/{(\w+)}/g, (_match: string, k: string) => {
            const value = (item as any)[k];
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
  // Process each table placeholder
  const tableMatches = Array.from(xmlString.matchAll(/{table:(\w+)}/g));

  for (const tableMatch of tableMatches) {
    const [fullMatch, key] = tableMatch;
    const arr = data[key];

    if (!Array.isArray(arr) || arr.length === 0) {
      console.warn(`Warning: Table data '${key}' not found or empty`);
      xmlString = xmlString.replace(fullMatch, "");
      continue;
    }

    const matchIndex = tableMatch.index!;

    // Find the table row containing this placeholder
    const trStartIndex = xmlString.lastIndexOf("<w:tr", matchIndex);
    const trEndIndex = xmlString.indexOf("</w:tr>", matchIndex) + 7;

    if (trStartIndex === -1 || trEndIndex === 6) {
      // 6 means indexOf returned -1, +7 = 6
      console.warn(
        `Warning: Could not find table row structure for {table:${key}}`
      );
      xmlString = xmlString.replace(fullMatch, "");
      continue;
    }

    // Extract the row template
    const rowTemplate = xmlString.substring(trStartIndex, trEndIndex);

    console.log(`Processing table ${key} with ${arr.length} items`);

    // Generate new rows
    const newRows = arr
      .filter((item) => item != null)
      .map((item) => {
        // Start with the template and remove the table placeholder
        let newRow = rowTemplate.replace(/{table:\w+}/, "");

        // Replace data placeholders
        newRow = newRow.replace(/{(\w+)}/g, (_, fieldName) => {
          const value = item[fieldName];
          return value != null ? String(value) : "";
        });

        return newRow;
      })
      .join("");

    // Replace the original row with the new rows
    xmlString =
      xmlString.substring(0, trStartIndex) +
      newRows +
      xmlString.substring(trEndIndex);
  }

  // Handle simple replacements: {name}
  xmlString = xmlString.replace(/{(\w+)}/g, (_, key) => {
    const value = data[key];
    if (typeof value === "string") return value;
    if (typeof value === "number" || typeof value === "boolean")
      return String(value);
    if (
      value &&
      typeof value === "object" &&
      "type" in value &&
      value.type === "image"
    ) {
      return `__IMAGE__${key}__`;
    }
    return "";
  });

  // Replace the XML content
  doc.file("word/document.xml", xmlString);

  // Embed images if any
  const hasImages = Object.values(data).some(
    (v): v is ImageData =>
      v !== null && typeof v === "object" && "type" in v && v.type === "image"
  );

  if (hasImages) {
    const relsPath = "word/_rels/document.xml.rels";
    let relsXml = await doc.file(relsPath)!.async("text");
    const relsDoc = new DOMParser().parseFromString(relsXml, "text/xml");

    let relIdCounter = 100;
    for (const [key, value] of Object.entries(data)) {
      // Type guard for ImageData
      if (
        !value ||
        typeof value !== "object" ||
        !("type" in value) ||
        value.type !== "image"
      ) {
        continue;
      }

      const imageData = value as ImageData;
      const imgBuffer = ensureBuffer(imageData.buffer);
      const dims = getImageSize(imgBuffer);

      const imgFile = `word/media/${key}.${imageData.extension}`;
      doc.file(imgFile, imgBuffer);

      // Add to relationships
      const relElem = relsDoc.createElement("Relationship");
      const rId = `rId${relIdCounter++}`;
      relElem.setAttribute("Id", rId);
      relElem.setAttribute(
        "Type",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      );
      relElem.setAttribute("Target", `media/${key}.${imageData.extension}`);
      relsDoc.documentElement.appendChild(relElem);

      // Replace image marker
      let docXml = await doc.file("word/document.xml")!.async("text");
      docXml = docXml.replace(
        `__IMAGE__${key}__`,
        createImageXML(rId, dims, imageData)
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

  // Generate appropriate buffer type based on environment
  if (typeof Buffer !== "undefined") {
    // Node.js environment
    const outputBuffer = await doc.generateAsync({ type: "nodebuffer" });
    return outputBuffer as Buffer;
  } else {
    // Browser environment
    const outputBuffer = await doc.generateAsync({ type: "uint8array" });
    return outputBuffer as Uint8Array;
  }
}

/**
 * Enhanced version of generateDocx that returns detailed statistics
 */
export async function generateDocxDetailed(
  templateBuffer: Buffer | Uint8Array | ArrayBuffer,
  data: TemplateData,
  options: GenerateDocxOptions = {}
): Promise<DocxGenerationResult> {
  // Track statistics
  let placeholdersReplaced = 0;
  let loopsProcessed = 0;
  let conditionalsProcessed = 0;
  let tablesGenerated = 0;
  let imagesEmbedded = 0;

  try {
    const zip = new JSZip();
    const doc = await zip.loadAsync(ensureBuffer(templateBuffer));

    const documentXml = await doc.file("word/document.xml")!.async("text");
    const serializer = new XMLSerializer();

    // First, we need to normalize the text to handle split placeholders
    let xmlString = normalizeDocxText(documentXml);

    // Count and handle loops
    const loopMatches = xmlString.match(/{#(\w+)}([\s\S]*?){\/\1}/g);
    if (loopMatches) {
      loopsProcessed = loopMatches.length;
      if (options.debug) console.log(`Processing ${loopsProcessed} loops`);
    }

    xmlString = xmlString.replace(
      /{#(\w+)}([\s\S]*?){\/\1}/g,
      (match, key, inner) => {
        const arr = data[key];
        if (!arr) {
          if (options.debug)
            console.warn(`Warning: Array '${key}' not found in data`);
          return "";
        }
        if (!Array.isArray(arr)) {
          if (options.debug)
            console.warn(`Warning: '${key}' is not an array, got:`, typeof arr);
          return "";
        }
        if (arr.length === 0) {
          if (options.debug) console.warn(`Warning: Array '${key}' is empty`);
          return "";
        }

        return arr
          .filter((item) => item != null)
          .map((item) =>
            inner.replace(/{(\w+)}/g, (_match: string, k: string) => {
              const value = (item as any)[k];
              if (value != null) placeholdersReplaced++;
              return value != null ? String(value) : "";
            })
          )
          .join("");
      }
    );

    // Count and handle conditionals
    const conditionalMatches = xmlString.match(
      /{\?(\w+)}([\s\S]*?)(?:{:else}([\s\S]*?))?{\/\1}/g
    );
    if (conditionalMatches) {
      conditionalsProcessed = conditionalMatches.length;
      if (options.debug)
        console.log(`Processing ${conditionalsProcessed} conditionals`);
    }

    xmlString = xmlString.replace(
      /{\?(\w+)}([\s\S]*?)(?:{:else}([\s\S]*?))?{\/\1}/g,
      (match, key, ifContent, elseContent = "") => {
        const condition = data[key];
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

    // Handle tables
    const tableMatches = Array.from(xmlString.matchAll(/{table:(\w+)}/g));
    tablesGenerated = tableMatches.length;
    if (options.debug && tablesGenerated > 0) {
      console.log(`Processing ${tablesGenerated} tables`);
    }

    for (const tableMatch of tableMatches) {
      const [fullMatch, key] = tableMatch;
      const arr = data[key];

      if (!Array.isArray(arr) || arr.length === 0) {
        if (options.debug)
          console.warn(`Warning: Table data '${key}' not found or empty`);
        xmlString = xmlString.replace(fullMatch, "");
        continue;
      }

      const matchIndex = tableMatch.index!;
      const trStartIndex = xmlString.lastIndexOf("<w:tr", matchIndex);
      const trEndIndex = xmlString.indexOf("</w:tr>", matchIndex) + 7;

      if (trStartIndex === -1 || trEndIndex === 6) {
        if (options.debug)
          console.warn(
            `Warning: Could not find table row structure for {table:${key}}`
          );
        xmlString = xmlString.replace(fullMatch, "");
        continue;
      }

      const rowTemplate = xmlString.substring(trStartIndex, trEndIndex);
      if (options.debug)
        console.log(`Processing table ${key} with ${arr.length} items`);

      const newRows = arr
        .filter((item) => item != null)
        .map((item) => {
          let newRow = rowTemplate.replace(/{table:\w+}/, "");
          newRow = newRow.replace(/{(\w+)}/g, (_, fieldName) => {
            const value = (item as any)[fieldName];
            if (value != null) placeholdersReplaced++;
            return value != null ? String(value) : "";
          });
          return newRow;
        })
        .join("");

      xmlString =
        xmlString.substring(0, trStartIndex) +
        newRows +
        xmlString.substring(trEndIndex);
    }

    // Count and handle simple replacements
    xmlString = xmlString.replace(/{(\w+)}/g, (_, key) => {
      const value = data[key];
      if (typeof value === "string") {
        placeholdersReplaced++;
        return value;
      }
      if (typeof value === "number" || typeof value === "boolean") {
        placeholdersReplaced++;
        return String(value);
      }
      if (
        value &&
        typeof value === "object" &&
        "type" in value &&
        value.type === "image"
      ) {
        return `__IMAGE__${key}__`;
      }
      return "";
    });

    doc.file("word/document.xml", xmlString);

    // Handle images
    const hasImages = Object.values(data).some(
      (v): v is ImageData =>
        v !== null && typeof v === "object" && "type" in v && v.type === "image"
    );

    if (hasImages) {
      const relsPath = "word/_rels/document.xml.rels";
      let relsXml = await doc.file(relsPath)!.async("text");
      const relsDoc = new DOMParser().parseFromString(relsXml, "text/xml");

      let relIdCounter = 100;
      for (const [key, value] of Object.entries(data)) {
        if (
          !value ||
          typeof value !== "object" ||
          !("type" in value) ||
          value.type !== "image"
        ) {
          continue;
        }

        imagesEmbedded++;
        const imageData = value as ImageData;
        const imgBuffer = ensureBuffer(imageData.buffer);
        const dims = getImageSize(imgBuffer);

        const imgFile = `word/media/${key}.${imageData.extension}`;
        doc.file(imgFile, imgBuffer);

        const relElem = relsDoc.createElement("Relationship");
        const rId = `rId${relIdCounter++}`;
        relElem.setAttribute("Id", rId);
        relElem.setAttribute(
          "Type",
          "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
        );
        relElem.setAttribute("Target", `media/${key}.${imageData.extension}`);
        relsDoc.documentElement.appendChild(relElem);

        let docXml = await doc.file("word/document.xml")!.async("text");
        docXml = docXml.replace(
          `__IMAGE__${key}__`,
          createImageXML(rId, dims, imageData)
        );
        doc.file("word/document.xml", docXml);
      }

      relsXml = serializer.serializeToString(relsDoc);
      doc.file(relsPath, relsXml);

      const contentTypesPath = "[Content_Types].xml";
      let contentTypesXml = await doc.file(contentTypesPath)!.async("text");
      const contentTypesDoc = new DOMParser().parseFromString(
        contentTypesXml,
        "text/xml"
      );

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

    // Generate appropriate buffer type based on environment
    let outputBuffer: Buffer | Uint8Array;
    if (typeof Buffer !== "undefined") {
      // Node.js environment
      outputBuffer = (await doc.generateAsync({
        type: "nodebuffer",
      })) as Buffer;
    } else {
      // Browser environment
      outputBuffer = (await doc.generateAsync({
        type: "uint8array",
      })) as Uint8Array;
    }

    return {
      buffer: outputBuffer,
      stats: {
        placeholdersReplaced,
        loopsProcessed,
        conditionalsProcessed,
        tablesGenerated,
        imagesEmbedded,
      },
    };
  } catch (error) {
    throw new DocxGenerationError(
      `Failed to generate DOCX: ${error instanceof Error ? error.message : String(error)}`,
      "GENERATION_FAILED",
      error
    );
  }
}

function createImageXML(
  rId: string,
  dims: any,
  imageData: Pick<ImageData, "widthInches" | "heightInches"> = {}
) {
  // Convert inches to EMUs (914400 EMUs per inch)
  const inchesToEMU = 914400;

  // Initialize with default values
  let cx: number = dims.width * 9525; // Convert pixels to EMUs
  let cy: number = dims.height * 9525;

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
    } else if (imageData.heightInches) {
      // Only height specified - maintain aspect ratio
      const aspectRatio = dims.width / dims.height;
      cy = imageData.heightInches * inchesToEMU;
      cx = cy * aspectRatio;
    }
  } else {
    // No custom size - use original logic with scaling
    const maxWidthTwips = 6 * inchesToEMU; // 6 inches max
    const maxHeightTwips = 6 * inchesToEMU;

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
