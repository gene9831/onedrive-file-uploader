import * as msal from "@azure/msal-node";
import type { UploadSession, User } from "@microsoft/microsoft-graph-types";
import path from "node:path";
import { calculateQuickXorHash } from "./quickxorhash";

interface UserCollection {
  value: User[];
  "@odata.nextLink"?: string;
}

const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const tenantId = process.env.TENANT_ID;

if (!clientId || !clientSecret || !tenantId) {
  throw new Error("CLIENT_ID, CLIENT_SECRET and TENANT_ID must be set");
}

const clientConfig: msal.Configuration = {
  auth: {
    clientId,
    authority: `https://login.microsoftonline.com/${tenantId}`,
    clientSecret,
  },
  cache: {
    cachePlugin: {
      beforeCacheAccess: async (tokenCacheContext) => {
        const cacheFile = "./.cache/token.json";
        const cacheFileExists = await Bun.file(cacheFile).exists();
        if (!cacheFileExists) {
          await Bun.file(cacheFile).write(JSON.stringify({}));
        }

        const token = await Bun.file(cacheFile).text();
        tokenCacheContext.cache.deserialize(token);
      },
      afterCacheAccess: async (tokenCacheContext) => {
        if (tokenCacheContext.hasChanged) {
          console.log("📢 Access token changed, updating cache");
          await Bun.file("./.cache/token.json").write(
            tokenCacheContext.cache.serialize()
          );
        }
      },
    },
  },
};

const cca = new msal.ConfidentialClientApplication(clientConfig);

const baseUrl = "https://graph.microsoft.com/v1.0";
/**
 * Get access token using client credentials flow
 * This requires the application to have User.Read.All or User.ReadWrite.All permissions
 */
async function getAccessToken(): Promise<string> {
  try {
    const tokenRequest = {
      scopes: ["https://graph.microsoft.com/.default"],
    };

    const response = await cca.acquireTokenByClientCredential(tokenRequest);

    if (!response || !response.accessToken) {
      throw new Error("Failed to acquire access token");
    }

    return response.accessToken;
  } catch (error) {
    console.error("Error acquiring token:", error);
    throw error;
  }
}

/**
 * List all users in the Azure AD tenant
 * Requires User.Read.All or User.ReadWrite.All permission
 * @param maxUsers Optional limit on the number of users to fetch (default: 999)
 */
async function listAllUsers(maxUsers: number = 999): Promise<User[]> {
  try {
    const accessToken = await getAccessToken();
    const allUsers: User[] = [];
    let nextLink: string | undefined = `${baseUrl}/users?$top=999`;

    // Fetch users with pagination support
    while (nextLink && allUsers.length < maxUsers) {
      const response = await fetch(nextLink, {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      });

      if (!response.ok) {
        const errorText = await response.text();
        throw new Error(
          `Failed to list users: ${response.status} ${response.statusText}\n${errorText}`
        );
      }

      const data = (await response.json()) as UserCollection;
      allUsers.push(...data.value);

      // Check if there are more pages
      nextLink = data["@odata.nextLink"];

      console.log(`Fetched ${allUsers.length} users so far...`);
    }

    return allUsers.slice(0, maxUsers);
  } catch (error) {
    console.error("Error listing users:", error);
    throw error;
  }
}

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

export function normalizeDirectory(
  directory: string,
  options?: { encode?: boolean }
): string {
  const { encode = true } = options || {};

  if (directory.trim() === "") return "/";

  // 1️⃣ 替换反斜杠 → 正斜杠
  directory = directory.trim().replace(/\\/g, "/");

  // 2️⃣ 标准化路径
  let normalized = path.normalize(directory).replace(/\\/g, "/");

  // 3️⃣ 去掉首尾 /
  normalized = normalized.replace(/^\/+|\/+$/g, "");

  if (!normalized) return "/";

  // 4️⃣ 拆分 + 编码（条件提前判断）
  let parts = normalized.split("/");

  if (encode) {
    parts = parts.map((p) => encodeURIComponent(p));
  }

  // 5️⃣ 重新拼接
  return "/" + parts.join("/") + "/";
}

async function uploadFile(filePath: string, userId: string) {
  const filename = path.basename(filePath);
  const file = Bun.file(filePath);
  const accessToken = await getAccessToken();

  const response = await fetch(
    `${baseUrl}/users/${userId}/drive/items/root:/Backups/${filename}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": file.type,
      },
      body: Bun.file(filePath).stream(),
    }
  );

  if (!response.ok) {
    throw new Error(
      `Failed to upload file: ${response.status} ${response.statusText}`
    );
  }

  const data = await response.json();
  console.log(response.status, data);
}

async function createUploadSession(
  filePath: string,
  directory: string,
  userId: string
) {
  const filename = path.basename(filePath);
  const accessToken = await getAccessToken();

  const normalizedDirectory = normalizeDirectory(directory);

  const response = await fetch(
    `${baseUrl}/users/${userId}/drive/items/root:${normalizedDirectory}${filename}:/createUploadSession`,
    {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        item: {
          "@microsoft.graph.conflictBehavior": "replace",
          name: filename,
        },
      }),
    }
  );

  if (!response.ok) {
    throw new Error(
      `Failed to create upload session: ${response.status} ${response.statusText}\n${await response.text()}`
    );
  }

  return (await response.json()) as UploadSession;
}

const MIN_CHUNK_SIZE = 327680; // 320 KB
const MAX_CHUNK_SIZE = 1024 * 1024 * 50; // 50MB

function calculateBytesRange(
  fileSize: number,
  chunkSize: number,
  session: UploadSession
) {
  if (chunkSize < MIN_CHUNK_SIZE || chunkSize > MAX_CHUNK_SIZE) {
    throw new Error(
      `Default chunk size is out of range. Must be between 320 KB and 50 MB. chunkSize: ${chunkSize}`
    );
  }

  if (chunkSize % MIN_CHUNK_SIZE !== 0) {
    throw new Error(
      `Default chunk size is not a multiple of minimum chunk size. Must be a multiple of 320 KB. chunkSize: ${chunkSize}`
    );
  }

  // 确保 fileSize 大于等于 chunkSize
  if (fileSize < chunkSize) {
    throw new Error(
      `File size is smaller than chunk size. fileSize: ${fileSize}, chunkSize: ${chunkSize}`
    );
  }

  // 如果 nextExpectedRanges 为空，则返回 0 到 chunkSize - 1
  if (
    !Array.isArray(session.nextExpectedRanges) ||
    session.nextExpectedRanges.length === 0
  ) {
    return { start: 0, end: chunkSize - 1, chunkSize };
  }

  const nextExpectedRange = session.nextExpectedRanges[0]!;
  const [startStr, endStr] = nextExpectedRange.split("-");

  let start = startStr ? Number(startStr) : 0;
  let end =
    endStr && endStr.trim() !== "" ? Number(endStr) : start + chunkSize - 1;

  // 上传片段大小不超过 chunkSize
  end = Math.min(end, start + chunkSize - 1);
  // end 不超过文件大小
  end = Math.min(end, fileSize - 1);

  return { start, end, chunkSize: end - start + 1 };
}

/**
 * Upload file to OneDrive using upload session (for large files)
 * Supports chunked upload with progress tracking
 * @param filePath Local file path
 * @param session Upload session from createUploadSession
 */
async function uploadFileToSession(
  filePath: string,
  session: UploadSession
): Promise<void> {
  const uploadUrl = session.uploadUrl;

  if (!uploadUrl) {
    throw new Error("Upload URL is not set");
  }

  const file = Bun.file(filePath);
  const fileSize = file.size;
  const chunkSize = 1024 * 1024 * 5; // 5MB

  let currentSession = session;

  while (true) {
    const startTime = Date.now();
    const {
      start,
      end,
      chunkSize: calculatedChunkSize,
    } = calculateBytesRange(fileSize, chunkSize, currentSession);

    // Read specific chunk of the file using slice()
    const chunk = file.slice(start, end + 1);

    // Upload this chunk
    const response = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": `${calculatedChunkSize}`,
        "Content-Range": `bytes ${start}-${end}/${fileSize}`,
      },
      body: chunk,
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(
        `Failed to upload chunk: ${response.status} ${response.statusText}\n${errorText}`
      );
    }

    // Calculate and display progress
    const progress = (((end + 1) / fileSize) * 100).toFixed(2);
    const speed = (calculatedChunkSize / (Date.now() - startTime)) * 1000;
    const eta = (fileSize - (end + 1)) / speed;
    console.log(
      `⬆️ Uploaded: ${formatBytes(end + 1)} / ${formatBytes(fileSize)} [${progress}%] [${formatBytes(speed)}/s] [ETA ${formatSeconds(eta)}]`
    );

    const data = await response.json();

    if (response.status === 200 || response.status === 201) {
      console.log(data);
      break;
    } else {
      currentSession = data as UploadSession;
    }
  }
}

async function promptUserSelection() {
  console.log("⚠️ USER_ID environment variable is not set.\n");
  console.log("👥 Fetching all users from Azure AD...\n");

  const users = await listAllUsers();
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

  console.log("🔐 Computing QuickXorHash...\n");
  const quickXorHash = await calculateQuickXorHash(filePath);
  console.log(`QuickXorHash: ${quickXorHash}`);

  const uploadSession = await createUploadSession(filePath, directory, userId);
  console.log(JSON.stringify(uploadSession.uploadUrl));
  await uploadFileToSession(filePath, uploadSession);
}

main();
