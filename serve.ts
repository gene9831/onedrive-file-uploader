import { GraphClient } from "./graph";

const clientId = process.env.CLIENT_ID;
const clientSecret = process.env.CLIENT_SECRET;
const tenantId = process.env.TENANT_ID;

if (!clientId || !clientSecret || !tenantId) {
  throw new Error("CLIENT_ID, CLIENT_SECRET and TENANT_ID must be set");
}

const graphClient = new GraphClient(clientId, clientSecret, tenantId);

const port = parseInt(process.env.PORT || "8001", 10);

// Initialize GraphClient with USER_ID
const userId = process.env.USER_ID;
if (!userId) {
  throw new Error("USER_ID must be set");
}
graphClient.setUserId(userId);

// CORS headers for allowing any origin
const corsHeaders = {
  "Access-Control-Allow-Origin": "*",
  "Access-Control-Allow-Methods": "GET, POST, PUT, DELETE, OPTIONS",
  "Access-Control-Allow-Headers": "Content-Type, Authorization",
  "Access-Control-Max-Age": "86400", // 24 hours
};

// Create HTTP server
const server = Bun.serve({
  hostname: "0.0.0.0",
  port,
  async fetch(req) {
    const url = new URL(req.url);
    const pathname = url.pathname;

    // Handle OPTIONS preflight request
    if (req.method === "OPTIONS") {
      return new Response(null, {
        status: 204,
        headers: corsHeaders,
      });
    }

    // Handle /files/* path pattern (list directory)
    if (pathname.startsWith("/files/")) {
      try {
        // Extract directory path from URL (remove /files/ prefix)
        const directoryPath = pathname.slice(7); // Remove "/files/" (7 characters)

        // Decode URL-encoded path
        const decodedPath = decodeURIComponent(directoryPath);

        // List directory contents
        const files = await graphClient.listDirectory(decodedPath);

        if (!Array.isArray(files)) {
          return new Response("Not Found", {
            status: 404,
            headers: corsHeaders,
          });
        }

        // Return JSON response
        return new Response(JSON.stringify(files, null, 2), {
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        // Return error response
        return new Response(
          JSON.stringify({
            error: error instanceof Error ? error.message : String(error),
          }),
          {
            status: 500,
            headers: {
              ...corsHeaders,
              "Content-Type": "application/json",
            },
          }
        );
      }
    }

    // Handle /file/* path pattern (get single item)
    if (pathname.startsWith("/file/")) {
      try {
        // Extract item path from URL (remove /file/ prefix)
        const itemPath = pathname.slice(6); // Remove "/file/" (6 characters)

        // Decode URL-encoded path
        const decodedPath = decodeURIComponent(itemPath);

        // Get single item
        const item = await graphClient.getItem(decodedPath);

        // Return JSON response
        return new Response(JSON.stringify(item, null, 2), {
          headers: {
            ...corsHeaders,
            "Content-Type": "application/json",
          },
        });
      } catch (error) {
        // Return error response
        return new Response(
          JSON.stringify({
            error: error instanceof Error ? error.message : String(error),
          }),
          {
            status: 500,
            headers: {
              ...corsHeaders,
              "Content-Type": "application/json",
            },
          }
        );
      }
    }

    // Handle /content/* path pattern (redirect to download URL)
    if (pathname.startsWith("/content/")) {
      try {
        // Extract item path from URL (remove /content/ prefix)
        const itemPath = pathname.slice(9); // Remove "/content/" (9 characters)

        // Decode URL-encoded path
        const decodedPath = decodeURIComponent(itemPath);

        // Get single item
        const item = await graphClient.getItem(decodedPath);

        // Get download URL from item
        const downloadUrl = (item as any)["@microsoft.graph.downloadUrl"];

        if (!downloadUrl) {
          return new Response(
            JSON.stringify({
              error: "Item does not have a download URL (may be a folder)",
            }),
            {
              status: 400,
              headers: {
                ...corsHeaders,
                "Content-Type": "application/json",
              },
            }
          );
        }

        // Redirect to download URL (with CORS headers)
        const redirectResponse = Response.redirect(downloadUrl, 302);
        Object.entries(corsHeaders).forEach(([key, value]) => {
          redirectResponse.headers.set(key, value);
        });
        return redirectResponse;
      } catch (error) {
        // Return error response
        return new Response(
          JSON.stringify({
            error: error instanceof Error ? error.message : String(error),
          }),
          {
            status: 500,
            headers: {
              ...corsHeaders,
              "Content-Type": "application/json",
            },
          }
        );
      }
    }

    // Return 404 for other paths
    return new Response("Not Found", {
      status: 404,
      headers: corsHeaders,
    });
  },
});

console.log(`Server is running on http://localhost:${port}`);
console.log(`Access /files/<directory-path> to list directory contents`);
console.log(`Access /file/<item-path> to get single file/folder info`);
console.log(`Access /content/<file-path> to redirect to file download URL`);
