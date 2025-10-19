import { readFile } from "fs/promises";
import { existsSync } from "fs";

export const downloadFile = async (file: string | Buffer) => {
  if (typeof file !== "string") {
    return file;
  }

  // Check if it's a URL
  if (file.startsWith("http://") || file.startsWith("https://")) {
    const response = await fetch(file);
    if (!response.ok) {
      throw new Error(
        `Failed to download file from ${file}: ${response.statusText}`
      );
    }
    return await response.arrayBuffer();
  }

  // Handle as file path
  if (!existsSync(file)) {
    throw new Error(`File not found: ${file}`);
  }

  const buffer = await readFile(file);

  return buffer;
};
