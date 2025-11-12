import type { User } from "@microsoft/microsoft-graph-types";
import { calculateQuickXorHash } from "./quickxorhash";
import { GraphClient } from "./graph";

const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const tenantId = process.env.TENANT_ID;

if (!clientId || !clientSecret || !tenantId) {
  throw new Error("CLIENT_ID, CLIENT_SECRET and TENANT_ID must be set");
}

const graphClient = new GraphClient(clientId, clientSecret, tenantId);

/**
 * Display users in a formatted table
 */
function displayUsers(users: User[]) {
  console.log(`\n${"=".repeat(100)}`);
  console.log(`Found ${users.length} users in total:\n`);
  console.log(
    `${"Display Name".padEnd(30)} | ${"Email".padEnd(35)} | ${"Object ID"}`
  );
  console.log(`${"-".repeat(100)}`);

  users.forEach((user, index) => {
    const displayName = (user.displayName || "N/A").padEnd(30);
    const email = (user.mail || user.userPrincipalName || "N/A").padEnd(35);
    const id = user.id;

    console.log(`${displayName} | ${email} | ${id}`);

    // Show additional details if available
    if (user.jobTitle || user.department || user.officeLocation) {
      const details: string[] = [];
      if (user.jobTitle) details.push(`Job: ${user.jobTitle}`);
      if (user.department) details.push(`Dept: ${user.department}`);
      if (user.officeLocation) details.push(`Office: ${user.officeLocation}`);
      console.log(`  └─ ${details.join(" | ")}`);
    }
  });

  console.log(`${"=".repeat(100)}\n`);
}

/**
 * Format bytes to human readable format
 */
function formatBytes(bytes: number, withWhitespace: boolean = false): string {
  if (bytes === 0) return `${withWhitespace ? "0 " : "0"}B`;
  const k = 1024;
  const sizes = ["B", "KB", "MB", "GB", "TB"];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return `${Math.round((bytes / Math.pow(k, i)) * 100) / 100}${withWhitespace ? " " : ""}${sizes[i]}`;
}

/**
 * Format seconds to human readable format
 * Examples: 30s, 1m40s, 3h23m, 1d23h
 * Shows at most two units, no decimals
 * Ignores small remainders (e.g., 86401s shows as 1d, not 1d1s)
 */
function formatSeconds(seconds: number): string {
  if (seconds === 0) return "0s";

  const units = [
    { value: 604800, suffix: "w" }, // week
    { value: 86400, suffix: "d" }, // day
    { value: 3600, suffix: "h" }, // hour
    { value: 60, suffix: "m" }, // minute
    { value: 1, suffix: "s" }, // second
  ];

  const parts: string[] = [];
  let remaining = Math.floor(seconds);

  // Find the first unit that fits
  for (let i = 0; i < units.length; i++) {
    const unit = units[i]!;
    if (remaining >= unit.value) {
      const count = Math.floor(remaining / unit.value);
      parts.push(`${count}${unit.suffix}`);
      remaining = remaining % unit.value;

      // Only show second unit if it's the immediate next unit (no skipping)
      // This allows 1m1s but prevents 1h1s or 1d1s (which skip units)
      if (remaining > 0 && parts.length < 2 && i < units.length - 1) {
        const nextUnit = units[i + 1]!;
        if (remaining >= nextUnit.value) {
          const nextCount = Math.floor(remaining / nextUnit.value);
          parts.push(`${nextCount}${nextUnit.suffix}`);
        }
      }
      break;
    }
  }

  // If no unit matched (seconds < 60), just return seconds
  if (parts.length === 0) {
    return `${remaining}s`;
  }

  return parts.join("");
}

async function promptUserSelection() {
  console.log("⚠️ USER_ID environment variable is not set.\n");
  console.log("👥 Fetching all users from Azure AD...\n");

  const users = await graphClient.listAllUsers();
  displayUsers(users);

  // Prompt user to select a user and set USER_ID
  console.log(
    "📝 To continue, please select a user and set the USER_ID environment variable:\n"
  );
  console.log("   Example:");
  console.log("   export USER_ID=<Object ID from the list above>");
  console.log("   or");
  console.log("   export USER_ID=<Email from the list above>");
  console.log("\n   Then run the script again: bun index.ts\n");
}

/**
 * Main function - checks if USER_ID is set and acts accordingly
 */
async function main() {
  const args = Bun.argv;

  if (args.length < 4) {
    console.error("Usage: bun index.ts <file> <directory>");
    process.exit(1);
  }

  const filePath = args[2]!;
  const directory = args[3]!;

  const fileExists = await Bun.file(filePath).exists();
  if (!fileExists) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
  }

  console.log(`File: ${filePath}`);
  console.log(`Directory: ${directory}`);

  const userId = process.env.USER_ID;

  // Check if USER_ID environment variable is set
  if (!userId) {
    // USER_ID not set - list all users and prompt user to select one
    await promptUserSelection();
    process.exit(0);
  }

  graphClient.setUserId(userId);

  console.log("🔐 Computing QuickXorHash...");
  const quickXorHash = await calculateQuickXorHash(filePath);
  console.log(`QuickXorHash: ${quickXorHash}`);

  const uploadSession = await graphClient.createUploadSession(
    filePath,
    directory
  );

  console.log("🚀 Starting upload...");

  const driveItem = await graphClient.uploadFileToSession(
    filePath,
    uploadSession,
    (progress: {
      uploaded: number;
      total: number;
      percentage: string;
      speed: number;
      eta: number;
    }) => {
      console.log(
        `⬆️ Uploaded: ${formatBytes(progress.uploaded)} / ${formatBytes(progress.total)} [${progress.percentage}%] [${formatBytes(progress.speed)}/s] [ETA ${formatSeconds(progress.eta)}]`
      );
    }
  );

  const { id, name, size, file, parentReference } = driveItem;
  console.log(`✅ Uploaded file:`);
  console.log({ id, name, size, directory: parentReference?.path });

  if (file?.hashes?.quickXorHash && quickXorHash !== file.hashes.quickXorHash) {
    console.warn("⚠️ QuickXorHash mismatch");
    console.warn(`Expected: ${quickXorHash}`);
    console.warn(`Actual: ${file.hashes.quickXorHash}`);
  }
}

main();
