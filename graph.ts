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
  // Promise to share token acquisition across concurrent calls
  private tokenPromise: Promise<string> | null = null;

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
   * Ensures only one token acquisition happens at a time, even with concurrent calls
   */
  async getAccessToken(): Promise<string> {
    // If there's already a token acquisition in progress, wait for it
    if (this.tokenPromise) {
      return this.tokenPromise;
    }

    // Create a new token acquisition promise
    this.tokenPromise = (async () => {
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
      } finally {
        // Clear the promise after completion (success or failure)
        // This allows retry on failure and new token acquisition when needed
        this.tokenPromise = null;
      }
    })();

    return this.tokenPromise;
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
  /**
   * Upload file directly to OneDrive (for small files)
   * Supports automatic retry on network errors and retryable HTTP status codes
   * @param filePath Local file path
   * @param directory Target directory in OneDrive
   * @param userId Optional user ID (uses setUserId if not provided)
   * @param maxRetries Maximum number of retries (default: 3)
   */
  async uploadFile(
    filePath: string,
    directory: string,
    userId?: string,
    maxRetries: number = 3
  ): Promise<DriveItem> {
    const filename = path.basename(filePath);
    const file = Bun.file(filePath);
    const normalizedDirectory = GraphClient.normalizeDirectory(directory);
    const targetUserId = userId ?? this.userId;

    if (!targetUserId) {
      throw new Error("USER_ID must be provided or set via setUserId()");
    }

    return await this.retryWithBackoff(
      async () => {
        const accessToken = await this.getAccessToken();

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
          const error = new Error(
            `Failed to upload file: ${response.status} ${response.statusText}`
          );
          (error as any).status = response.status;
          throw error;
        }

        return (await response.json()) as DriveItem;
      },
      maxRetries,
      1000,
      30000
    );
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
   * Check if an HTTP status code is retryable
   */
  private isRetryableStatus(status: number): boolean {
    // Retry on rate limiting (429), server errors (5xx), and some client errors
    return (
      status === 429 || // Too Many Requests
      status === 500 || // Internal Server Error
      status === 502 || // Bad Gateway
      status === 503 || // Service Unavailable
      status === 504 // Gateway Timeout
    );
  }

  /**
   * Sleep for a given number of milliseconds
   */
  private sleep(ms: number): Promise<void> {
    return new Promise((resolve) => setTimeout(resolve, ms));
  }

  /**
   * Retry a function with exponential backoff
   * @param fn Function to retry
   * @param maxRetries Maximum number of retries (default: 3)
   * @param initialDelay Initial delay in milliseconds (default: 1000)
   * @param maxDelay Maximum delay in milliseconds (default: 30000)
   */
  private async retryWithBackoff<T>(
    fn: () => Promise<T>,
    maxRetries: number = 3,
    initialDelay: number = 1000,
    maxDelay: number = 30000
  ): Promise<T> {
    let lastError: Error | unknown;
    let delay = initialDelay;

    for (let attempt = 0; attempt <= maxRetries; attempt++) {
      try {
        return await fn();
      } catch (error) {
        lastError = error;

        // Don't retry on the last attempt
        if (attempt === maxRetries) {
          break;
        }

        // Check if error is retryable
        let isRetryable = false;
        if (error instanceof Error) {
          const errorWithFlags = error as any;

          // Check if it's explicitly marked as a network error
          if (errorWithFlags.isNetworkError) {
            isRetryable = true;
          }
          // Network errors are retryable (check error message)
          else if (
            error.message.includes("fetch") ||
            error.message.includes("network") ||
            error.message.includes("ECONNRESET") ||
            error.message.includes("ETIMEDOUT")
          ) {
            isRetryable = true;
          }
          // Check if error has status code attached (from HTTP errors)
          else if (
            errorWithFlags.status &&
            typeof errorWithFlags.status === "number"
          ) {
            isRetryable = this.isRetryableStatus(errorWithFlags.status);
          } else {
            // Fallback: try to extract status from error message
            const statusMatch = error.message.match(/status[:\s]+(\d+)/i);
            if (statusMatch) {
              const status = parseInt(statusMatch[1]!, 10);
              isRetryable = this.isRetryableStatus(status);
            }
          }
        }

        // If not retryable, throw immediately
        if (!isRetryable) {
          throw error;
        }

        // Calculate exponential backoff with jitter
        const jitter = Math.random() * 0.3 * delay; // Add up to 30% jitter
        const backoffDelay = Math.min(delay + jitter, maxDelay);

        console.warn(
          `⚠️ Upload chunk failed (attempt ${attempt + 1}/${maxRetries + 1}), retrying in ${Math.round(backoffDelay)}ms...`
        );

        await this.sleep(backoffDelay);
        delay = Math.min(delay * 2, maxDelay); // Exponential backoff
      }
    }

    // If we get here, all retries failed
    throw lastError;
  }

  /**
   * Upload file to OneDrive using upload session (for large files)
   * Supports chunked upload with progress tracking and automatic retry
   * @param filePath Local file path
   * @param session Upload session from createUploadSession
   * @param onProgress Optional progress callback
   * @param maxRetries Maximum number of retries per chunk (default: 3)
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
    }) => void,
    maxRetries: number = 3
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

      // Upload this chunk with retry mechanism
      // Note: We re-read the chunk on each retry attempt since the stream may be consumed
      const response = await this.retryWithBackoff(
        async () => {
          try {
            // Re-read the chunk from file for each retry attempt
            const chunk = file.slice(start, end + 1);

            const res = await fetch(uploadUrl, {
              method: "PUT",
              headers: {
                "Content-Length": `${calculatedChunkSize}`,
                "Content-Range": `bytes ${start}-${end}/${fileSize}`,
              },
              body: chunk,
            });

            if (!res.ok) {
              const errorText = await res.text();
              const error = new Error(
                `Failed to upload chunk: ${res.status} ${res.statusText}\n${errorText}`
              );
              // Attach status code to error for retry logic
              (error as any).status = res.status;
              throw error;
            }

            return res;
          } catch (error) {
            // Re-throw HTTP errors (already handled above)
            if (error instanceof Error && (error as any).status) {
              throw error;
            }
            // Wrap network/fetch errors for retry logic
            const networkError = new Error(
              `Network error during chunk upload: ${error instanceof Error ? error.message : String(error)}`
            );
            // Mark as network error for retry detection
            (networkError as any).isNetworkError = true;
            throw networkError;
          }
        },
        maxRetries,
        1000, // Initial delay: 1 second
        30000 // Max delay: 30 seconds
      );

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

  async listDirectory(
    directory: string,
    userId?: string
  ): Promise<DriveItem[]> {
    const accessToken = await this.getAccessToken();
    const normalizedDirectory = GraphClient.normalizeDirectory(directory);
    const targetUserId = userId ?? this.userId;

    if (!targetUserId) {
      throw new Error("USER_ID must be provided or set via setUserId()");
    }

    const driveItemResp = await fetch(
      `${this.baseUrl}/users/${targetUserId}/drive/items/root:${normalizedDirectory}`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    if (!driveItemResp.ok) {
      throw new Error(
        `Failed to list directory: ${driveItemResp.status} ${driveItemResp.statusText}\n${await driveItemResp.text()}`
      );
    }

    const fileId = ((await driveItemResp.json()) as DriveItem).id;

    const childrenResp = await fetch(
      `${this.baseUrl}/users/${targetUserId}/drive/items/${fileId}/children`,
      {
        headers: {
          Authorization: `Bearer ${accessToken}`,
        },
      }
    );

    if (!childrenResp.ok) {
      throw new Error(
        `Failed to list children: ${childrenResp.status} ${childrenResp.statusText}\n${await childrenResp.text()}`
      );
    }

    const data = (await childrenResp.json()) as { value: DriveItem[] };

    return data.value;
  }
}
