/**
 * TypeScript type definitions for DOCX Templater
 */

/**
 * Image configuration for embedding images in DOCX
 */
export interface ImageData {
  /** Specifies this is an image object */
  type: "image";
  /** Buffer containing the image data */
  buffer: Buffer;
  /** File extension of the image (jpg, png, gif, etc.) */
  extension: string;
  /** Optional width in inches. If only width is specified, height is calculated to maintain aspect ratio */
  widthInches?: number;
  /** Optional height in inches. If only height is specified, width is calculated to maintain aspect ratio */
  heightInches?: number;
}

/**
 * Template data that can be used in DOCX templates
 * - Strings are replaced directly in placeholders
 * - Arrays are used for loops and tables
 * - Objects with ImageData type are embedded as images
 * - Other values are converted to strings
 */
export interface TemplateData {
  [key: string]:
    | string
    | number
    | boolean
    | null
    | undefined
    | ImageData
    | TemplateData[]
    | TemplateData;
}

/**
 * Options for the generateDocx function
 */
export interface GenerateDocxOptions {
  /** Enable debug logging for troubleshooting */
  debug?: boolean;
  /** Custom image size limits in inches */
  maxImageWidth?: number;
  maxImageHeight?: number;
}

/**
 * Result of DOCX generation
 */
export interface DocxGenerationResult {
  /** The generated DOCX file as a Buffer */
  buffer: Buffer;
  /** Statistics about the generation process */
  stats?: {
    /** Number of placeholders replaced */
    placeholdersReplaced: number;
    /** Number of loops processed */
    loopsProcessed: number;
    /** Number of conditionals evaluated */
    conditionalsProcessed: number;
    /** Number of tables generated */
    tablesGenerated: number;
    /** Number of images embedded */
    imagesEmbedded: number;
  };
}

/**
 * Error thrown when DOCX generation fails
 */
export class DocxGenerationError extends Error {
  constructor(
    message: string,
    public readonly code: string,
    public readonly details?: any
  ) {
    super(message);
    this.name = "DocxGenerationError";
  }
}

/**
 * Main function to generate DOCX from template and data
 */
export declare function generateDocx(
  templateBuffer: Buffer,
  data: TemplateData,
  options?: GenerateDocxOptions
): Promise<Buffer>;

/**
 * Alternative function that returns detailed result with statistics
 */
export declare function generateDocxDetailed(
  templateBuffer: Buffer,
  data: TemplateData,
  options?: GenerateDocxOptions
): Promise<DocxGenerationResult>;
