import type { User } from "@microsoft/microsoft-graph-types";
import { calculateQuickXorHash } from "./quickxorhash";
import { GraphClient } from "./graph";
import { UploadQueue } from "./upload-queue";
import { readdir } from "node:fs/promises";
import path from "node:path";
import cliProgress from "cli-progress";

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
 * Handle upload command
 */
async function handleUpload(filePath: string, directory: string) {
  const fileExists = await Bun.file(filePath).exists();
  if (!fileExists) {
    console.error(`File not found: ${filePath}`);
    process.exit(1);
  }

  const file = Bun.file(filePath);
  const fileSize = file.size;
  const MAX_DIRECT_UPLOAD_SIZE = 5 * 1024 * 1024; // 5MB

  console.log(`File: ${filePath}`);
  console.log(`Directory: ${directory}`);
  console.log(`File size: ${formatBytes(fileSize)}`);

  const userId = process.env.USER_ID;

  if (!userId) {
    await promptUserSelection();
    process.exit(0);
  }

  graphClient.setUserId(userId);

  console.log("🔐 Computing QuickXorHash...");
  const quickXorHash = await calculateQuickXorHash(filePath);
  console.log(`QuickXorHash: ${quickXorHash}`);

  let driveItem;

  if (fileSize <= MAX_DIRECT_UPLOAD_SIZE) {
    console.log("🚀 Starting direct upload (file size <= 5MB)...");
    driveItem = await graphClient.uploadFile(filePath, directory);
  } else {
    console.log("🚀 Starting chunked upload (file size > 5MB)...");
    const uploadSession = await graphClient.createUploadSession(
      filePath,
      directory
    );

    const progressBar = new cliProgress.SingleBar(
      {
        format:
          "⬆️ Uploading |{bar}| {percentage}% | {value}/{total} | Speed: {speed} | ETA: {eta_formatted}",
        barCompleteChar: "\u2588",
        barIncompleteChar: "\u2591",
        hideCursor: true,
      },
      cliProgress.Presets.shades_classic
    );

    progressBar.start(fileSize, 0, {
      speed: "N/A",
      eta_formatted: "N/A",
    });

    driveItem = await graphClient.uploadFileToSession(
      filePath,
      uploadSession,
      (progress: {
        uploaded: number;
        total: number;
        percentage: string;
        speed: number;
        eta: number;
      }) => {
        progressBar.update(progress.uploaded, {
          speed: `${formatBytes(progress.speed)}/s`,
          eta_formatted: formatSeconds(progress.eta),
        });
        progressBar.setTotal(progress.total);
      }
    );

    progressBar.update(fileSize, {
      speed: "Complete",
      eta_formatted: "Done",
    });
    progressBar.stop();
  }

  const { id, name, size, file: fileInfo, parentReference } = driveItem;
  console.log(`✅ Uploaded file:`);
  console.log({ id, name, size, directory: parentReference?.path });

  if (
    fileInfo?.hashes?.quickXorHash &&
    quickXorHash !== fileInfo.hashes.quickXorHash
  ) {
    console.warn("⚠️ QuickXorHash mismatch");
    console.warn(`Expected: ${quickXorHash}`);
    console.warn(`Actual: ${fileInfo.hashes.quickXorHash}`);
  }
}

async function handleList(directory: string) {
  const userId = process.env.USER_ID;

  if (!userId) {
    await promptUserSelection();
    process.exit(0);
  }

  graphClient.setUserId(userId);

  const files = await graphClient.listDirectory(directory);
  console.log(`Found ${files.length} files in directory: ${directory}`);
  console.log(files);
}

/**
 * Recursively get all files in a directory using Bun API
 */
async function getAllFiles(
  dirPath: string,
  basePath: string = dirPath
): Promise<Array<{ filePath: string; relativePath: string; size: number }>> {
  const files: Array<{ filePath: string; relativePath: string; size: number }> =
    [];
  const entries = await readdir(dirPath, { withFileTypes: true });

  for (const entry of entries) {
    if (entry.name.startsWith(".") || entry.name.startsWith("_")) {
      continue;
    }

    const fullPath = path.join(dirPath, entry.name);
    const relativePath = path.relative(basePath, fullPath);

    if (entry.isDirectory()) {
      const subFiles = await getAllFiles(fullPath, basePath);
      files.push(...subFiles);
    } else if (entry.isFile()) {
      const file = Bun.file(fullPath);
      const size = file.size;
      files.push({ filePath: fullPath, relativePath, size });
    }
  }

  return files;
}

/**
 * Handle upload directory command with queue and concurrency control
 */
async function handleUploadDirectory(
  localDirPath: string,
  remoteDirectory: string
) {
  try {
    await readdir(localDirPath);
  } catch (error) {
    console.error(`Directory not found: ${localDirPath}`);
    process.exit(1);
  }

  console.log(`Local directory: ${localDirPath}`);
  console.log(`Remote directory: ${remoteDirectory}`);
  console.log("📂 Scanning directory...\n");

  const files = await getAllFiles(localDirPath);
  console.log(`Found ${files.length} files to upload\n`);

  if (files.length === 0) {
    console.log("No files to upload.");
    return;
  }

  const userId = process.env.USER_ID;

  if (!userId) {
    await promptUserSelection();
    process.exit(0);
  }

  graphClient.setUserId(userId);

  const MAX_DIRECT_UPLOAD_SIZE = 5 * 1024 * 1024;
  const MAX_CONCURRENT_UPLOADS = 5;

  const sourceDirName = path.basename(path.resolve(localDirPath));
  const baseTargetDir = remoteDirectory.endsWith("/")
    ? `${remoteDirectory}${sourceDirName}`
    : `${remoteDirectory}/${sourceDirName}`;

  const queue = new UploadQueue(MAX_CONCURRENT_UPLOADS, files.length);
  let completedCount = 0;
  let failedCount = 0;

  console.log(
    `🚀 Starting upload with max ${MAX_CONCURRENT_UPLOADS} concurrent uploads...\n`
  );

  const multibar = new cliProgress.MultiBar(
    {
      clearOnComplete: true,
      hideCursor: true,
      stopOnComplete: false,
      synchronousUpdate: false,
      format:
        " {bar} | {percentage}% | {value}/{total} | Speed: {speed} | {filename}",
      barCompleteChar: "\u2588",
      barIncompleteChar: "\u2591",
    },
    cliProgress.Presets.shades_grey
  );

  const overallBar = multibar.create(files.length, 0, {
    filename: "📦 Overall Progress",
    speed: "N/A",
  });

  const activeBars = new Map<string, cliProgress.SingleBar>();
  const lastUpdateTime = new Map<string, number>();
  const UPDATE_THROTTLE_MS = 100;

  for (let i = 0; i < files.length; i++) {
    const { filePath, relativePath, size } = files[i]!;
    const fileDir = path.dirname(relativePath);
    const targetDir =
      fileDir === "." ? baseTargetDir : `${baseTargetDir}/${fileDir}`;

    queue.add(async () => {
      const currentIndex = ++completedCount;
      const fileBar = multibar.create(size, 0, {
        filename: `[${currentIndex}/${files.length}] ${relativePath}`,
        speed: "N/A",
      });
      activeBars.set(relativePath, fileBar);
      lastUpdateTime.set(relativePath, Date.now());

      try {
        if (size <= MAX_DIRECT_UPLOAD_SIZE) {
          await graphClient.uploadFile(filePath, targetDir);
          fileBar.update(size, { speed: "Complete" });
        } else {
          const uploadSession = await graphClient.createUploadSession(
            filePath,
            targetDir
          );

          await graphClient.uploadFileToSession(
            filePath,
            uploadSession,
            (progress: {
              uploaded: number;
              total: number;
              percentage: string;
              speed: number;
              eta: number;
            }) => {
              const now = Date.now();
              const lastUpdate = lastUpdateTime.get(relativePath) || 0;
              if (now - lastUpdate >= UPDATE_THROTTLE_MS) {
                fileBar.update(progress.uploaded, {
                  speed: `${formatBytes(progress.speed)}/s`,
                });
                fileBar.setTotal(progress.total);
                lastUpdateTime.set(relativePath, now);
              }
            }
          );
          fileBar.update(size, { speed: "Complete" });
        }

        fileBar.update(size, { speed: "Complete" });
        await new Promise((resolve) => setTimeout(resolve, 50));
        fileBar.stop();
        activeBars.delete(relativePath);
        lastUpdateTime.delete(relativePath);
        overallBar.increment();
      } catch (error) {
        failedCount++;
        fileBar.stop();
        activeBars.delete(relativePath);
        lastUpdateTime.delete(relativePath);
        overallBar.increment();
        console.error(
          `\n❌ Failed to upload ${relativePath}:`,
          error instanceof Error ? error.message : error
        );
        throw error;
      }
    });
  }

  await queue.waitForCompletion();
  multibar.stop();

  const stats = queue.getStats();
  console.log(`\n${"=".repeat(100)}`);
  console.log(`Upload summary:`);
  console.log(`  ✅ Success: ${stats.success}`);
  console.log(`  ❌ Failed: ${stats.failed}`);
  console.log(`  📊 Total: ${stats.total}`);
  console.log(`${"=".repeat(100)}\n`);
}

/**
 * Main function - parses command line arguments and routes to appropriate handler
 */
async function main() {
  const args = Bun.argv;

  if (args.length < 3) {
    console.error("Usage: bun index.ts <command> [options]");
    console.error("\nCommands:");
    console.error("  upload <file> <directory>     Upload a file to OneDrive");
    console.error(
      "  upload-dir <dir> <directory>   Upload a directory to OneDrive"
    );
    console.error("  list <directory>               List files in a directory");
    process.exit(1);
  }

  const command = args[2];

  switch (command) {
    case "upload":
      if (args.length < 5) {
        console.error("Usage: bun index.ts upload <file> <directory>");
        process.exit(1);
      }
      await handleUpload(args[3]!, args[4]!);
      break;

    case "upload-dir":
      if (args.length < 5) {
        console.error(
          "Usage: bun index.ts upload-dir <local_dir> <remote_directory>"
        );
        process.exit(1);
      }
      await handleUploadDirectory(args[3]!, args[4]!);
      break;

    case "list":
      if (args.length < 4) {
        console.error("Usage: bun index.ts list <directory>");
        process.exit(1);
      }
      await handleList(args[3]!);
      break;

    default:
      console.error(`Unknown command: ${command}`);
      console.error("\nAvailable commands:");
      console.error(
        "  upload <file> <directory>     Upload a file to OneDrive"
      );
      console.error(
        "  upload-dir <dir> <directory>   Upload a directory to OneDrive"
      );
      console.error(
        "  list <directory>               List files in a directory"
      );
      process.exit(1);
  }
}

main();
