import path from "node:path";

/**
 * Calculate QuickXorHash for a file using the quickxorhash-rust CLI tool
 * @param filePath Path to the file to calculate hash for
 * @returns The hash value as a string (base64 encoded)
 */
async function calculateQuickXorHash(filePath: string): Promise<string> {
  // Resolve the absolute path to the CLI tool
  const cliPath = path.join(import.meta.dir, "vendor", "quickxorhash-rust");

  // Execute the CLI tool with the file path
  const process = Bun.spawn([cliPath, filePath], {
    stdout: "pipe",
    stderr: "pipe",
  });

  // Wait for the process to complete
  const result = await process.exited;

  if (result !== 0) {
    // Read error message if process failed
    const errorText = await new Response(process.stderr).text();
    throw new Error(`Failed to calculate hash: ${errorText}`);
  }

  // Read the hash output from stdout
  const hash = await new Response(process.stdout).text();

  // Trim whitespace (newline, etc.)
  return hash.trim();
}

// Example usage
async function main() {
  const filePath = process.argv[2];

  if (!filePath) {
    console.error("Usage: bun test.ts <filePath>");
    process.exit(1);
  }

  try {
    const hash = await calculateQuickXorHash(filePath);
    console.log(`QuickXorHash: ${hash}`);
  } catch (error) {
    console.error("Error:", error);
    process.exit(1);
  }
}

// Run if executed directly
if (import.meta.main) {
  main();
}

// Export the function for use in other modules
export { calculateQuickXorHash };
