// Main exports
export { generateDocx, generateDocxDetailed } from "./utils";

// Type exports
export type {
  TemplateData,
  ImageData,
  GenerateDocxOptions,
  DocxGenerationResult,
  DocxGenerationError,
} from "./types";

// Re-export everything for convenience
export * from "./types";
export * from "./utils";
