import * as msal from "@azure/msal-node";
import type {
  DriveItem,
  UploadSession,
  User,
} from "@microsoft/microsoft-graph-types";
import path from "node:path";

interface UserCollection {
  value: User[];
  "@odata.nextLink"?: string;
}

const MIN_CHUNK_SIZE = 327680; // 320 KB
const MAX_CHUNK_SIZE = 1024 * 1024 * 50; // 50MB

/**
 * Microsoft Graph API client for uploading files and managing upload sessions
 */
export class GraphClient {
  private cca: msal.ConfidentialClientApplication;
  private baseUrl = "https://graph.microsoft.com/v1.0";
  private userId?: string;

  constructor(
    private clientId: string,
    private clientSecret: string,
    private tenantId: string,
    private cacheFile: string = "./.cache/token.json"
  ) {
    const clientConfig: msal.Configuration = {
      auth: {
        clientId,
        authority: `https://login.microsoftonline.com/${tenantId}`,
        clientSecret,
      },
      cache: {
        cachePlugin: {
          beforeCacheAccess: async (tokenCacheContext) => {
            const cacheFileExists = await Bun.file(this.cacheFile).exists();
            if (!cacheFileExists) {
              await Bun.file(this.cacheFile).write(JSON.stringify({}));
            }

            const token = await Bun.file(this.cacheFile).text();
            tokenCacheContext.cache.deserialize(token);
          },
          afterCacheAccess: async (tokenCacheContext) => {
            if (tokenCacheContext.hasChanged) {
              console.log("📢 Access token changed, updating cache");
              await Bun.file(this.cacheFile).write(
                tokenCacheContext.cache.serialize()
              );
            }
          },
        },
      },
    };

    this.cca = new msal.ConfidentialClientApplication(clientConfig);
  }

  /**
   * Set the target user ID for subsequent Graph API operations
   */
  setUserId(userId: string) {
    this.userId = userId;
  }

  /**
   * Get access token using client credentials flow
   * This requires the application to have User.Read.All or User.ReadWrite.All permissions
   */
  async getAccessToken(): Promise<string> {
    try {
      const tokenRequest = {
        scopes: ["https://graph.microsoft.com/.default"],
      };

      const response =
        await this.cca.acquireTokenByClientCredential(tokenRequest);

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
  async listAllUsers(maxUsers: number = 999): Promise<User[]> {
    try {
      const accessToken = await this.getAccessToken();
      const allUsers: User[] = [];
      let nextLink: string | undefined = `${this.baseUrl}/users?$top=999`;

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
   * Normalize directory path for Graph API
   */
  static normalizeDirectory(
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

  /**
   * Upload file directly to OneDrive (for small files)
   */
  async uploadFile(
    filePath: string,
    directory: string,
    userId?: string
  ): Promise<DriveItem> {
    const filename = path.basename(filePath);
    const file = Bun.file(filePath);
    const accessToken = await this.getAccessToken();

    const normalizedDirectory = GraphClient.normalizeDirectory(directory);
    const targetUserId = userId ?? this.userId;

    if (!targetUserId) {
      throw new Error("USER_ID must be provided or set via setUserId()");
    }

    const response = await fetch(
      `${this.baseUrl}/users/${targetUserId}/drive/items/root:${normalizedDirectory}${filename}:/content`,
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

    return (await response.json()) as DriveItem;
  }

  /**
   * Create an upload session for large file uploads
   */
  async createUploadSession(
    filePath: string,
    directory: string,
    userId?: string
  ): Promise<UploadSession> {
    const filename = path.basename(filePath);
    const accessToken = await this.getAccessToken();

    const normalizedDirectory = GraphClient.normalizeDirectory(directory);
    const targetUserId = userId ?? this.userId;

    if (!targetUserId) {
      throw new Error("USER_ID must be provided or set via setUserId()");
    }

    const response = await fetch(
      `${this.baseUrl}/users/${targetUserId}/drive/items/root:${normalizedDirectory}${filename}:/createUploadSession`,
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

  /**
   * Calculate bytes range for chunked upload
   */
  private calculateBytesRange(
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
   * @param onProgress Optional progress callback
   */
  async uploadFileToSession(
    filePath: string,
    session: UploadSession,
    onProgress?: (progress: {
      uploaded: number;
      total: number;
      percentage: string;
      speed: number;
      eta: number;
    }) => void
  ): Promise<DriveItem> {
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
      } = this.calculateBytesRange(fileSize, chunkSize, currentSession);

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

      const data = await response.json();

      // Calculate and report progress
      const uploaded = end + 1;
      const progress = ((uploaded / fileSize) * 100).toFixed(2);
      const speed = (calculatedChunkSize / (Date.now() - startTime)) * 1000;
      const eta = (fileSize - uploaded) / speed;

      if (onProgress) {
        onProgress({
          uploaded,
          total: fileSize,
          percentage: progress,
          speed,
          eta,
        });
      }

      if (response.status === 200 || response.status === 201) {
        return data as DriveItem;
      } else {
        currentSession = data as UploadSession;
      }
    }
  }
}

