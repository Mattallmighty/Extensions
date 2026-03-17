import { createKnowledgeConnector } from "@cognigy/extension-tools";
import axios from "axios";
import {
    deleteKnowledgeSourceById,
    listKnowledgeSources,
    ManagedSource,
    patchKnowledgeSourceDescription,
} from "./cognigyManagementApi";

// ─── Constants ────────────────────────────────────────────────────────────────

const SUPPORTED_EXTENSIONS = [
    ".pdf", ".txt", ".docx", ".pptx",
    ".jpeg", ".jpg", ".png", ".bmp", ".heif", ".tiff",
];

const MIME_TYPES: Record<string, string> = {
    ".pdf":  "application/pdf",
    ".txt":  "text/plain",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".jpeg": "image/jpeg",
    ".jpg":  "image/jpeg",
    ".png":  "image/png",
    ".bmp":  "image/bmp",
    ".heif": "image/heif",
    ".tiff": "image/tiff",
};

function getMimeType(fileName: string): string {
    const ext = fileName.substring(fileName.lastIndexOf(".")).toLowerCase();
    return MIME_TYPES[ext] ?? "application/octet-stream";
}

function sleep(ms: number): Promise<void> {
    return new Promise(resolve => setTimeout(resolve, ms));
}

// ─── URL parsing ──────────────────────────────────────────────────────────────

interface IParsedSharePointUrl {
    siteUrl: string;
    libraryName: string;
    folderPath: string;
}

/**
 * Accept any SharePoint URL format and extract the site URL, library name,
 * and optional subfolder path.
 *
 * Examples:
 *   https://contoso.sharepoint.com/sites/hr
 *     → siteUrl=…/sites/hr, library="Shared Documents", folder=""
 *   https://contoso.sharepoint.com/sites/hr/Shared%20Documents/Forms/AllItems.aspx
 *     → siteUrl=…/sites/hr, library="Shared Documents", folder=""
 *   https://contoso.sharepoint.com/sites/hr/Policies/2025
 *     → siteUrl=…/sites/hr, library="Policies", folder="2025"
 */
function parseSharePointUrl(rawUrl: string): IParsedSharePointUrl {
    try {
        const url = new URL(rawUrl);
        const parts = url.pathname.split("/").filter(Boolean);

        // Find "sites" or "teams" segment
        const siteKeywordIdx = parts.findIndex(
            p => p.toLowerCase() === "sites" || p.toLowerCase() === "teams"
        );

        if (siteKeywordIdx >= 0 && parts.length > siteKeywordIdx + 1) {
            const siteUrl = `${url.origin}/${parts.slice(0, siteKeywordIdx + 2).join("/")}`;
            const afterSite = parts.slice(siteKeywordIdx + 2);

            if (afterSite.length === 0) {
                return { siteUrl, libraryName: "Shared Documents", folderPath: "" };
            }

            const libraryName = decodeURIComponent(afterSite[0]);

            // Strip "Forms" segment and any .aspx file at the end
            const folderParts = afterSite
                .slice(1)
                .filter(p => p.toLowerCase() !== "forms" && !p.toLowerCase().endsWith(".aspx"))
                .map(decodeURIComponent);

            return { siteUrl, libraryName, folderPath: folderParts.join("/") };
        }
    } catch {
        // fall through
    }
    return { siteUrl: rawUrl, libraryName: "Shared Documents", folderPath: "" };
}

// ─── Auth ────────────────────────────────────────────────────────────────────

async function getAccessToken(
    authMethod: string,
    tenantId: string,
    clientId: string,
    clientSecret: string,
    bearerToken: string
): Promise<string> {
    if (authMethod === "bearer") {
        return bearerToken.trim();
    }
    const url = `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`;
    const params = new URLSearchParams();
    params.append("grant_type", "client_credentials");
    params.append("client_id", clientId);
    params.append("client_secret", clientSecret);
    params.append("scope", "https://graph.microsoft.com/.default");

    const response = await axios.post(url, params.toString(), {
        headers: { "Content-Type": "application/x-www-form-urlencoded" }
    });
    return response.data.access_token as string;
}

// ─── SharePoint site / drive helpers ─────────────────────────────────────────

async function getSiteId(siteUrl: string, token: string): Promise<string> {
    const parsed = new URL(siteUrl);
    const graphUrl = `https://graph.microsoft.com/v1.0/sites/${parsed.hostname}:${parsed.pathname}`;
    const resp = await axios.get(graphUrl, {
        headers: { Authorization: `Bearer ${token}` }
    });
    return resp.data.id as string;
}

async function getDriveId(siteId: string, libraryName: string, token: string): Promise<string> {
    const resp = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, {
        headers: { Authorization: `Bearer ${token}` }
    });
    const drives: any[] = resp.data.value ?? [];
    const match = drives.find((d: any) =>
        d.name?.toLowerCase() === libraryName.toLowerCase() ||
        d.webUrl?.toLowerCase().endsWith(`/${libraryName.toLowerCase()}`)
    );
    if (match) return match.id as string;

    // Fall back to the site's default drive
    const defResp = await axios.get(`https://graph.microsoft.com/v1.0/sites/${siteId}/drive`, {
        headers: { Authorization: `Bearer ${token}` }
    });
    return defResp.data.id as string;
}

// ─── File listing ─────────────────────────────────────────────────────────────

interface IFileItem {
    id: string;
    name: string;
    webUrl: string;
    path: string;
    /** quickXorHash from Graph API (base64) — used for change detection. */
    quickXorHash: string;
    lastModified: string;
}

async function listFiles(
    driveId: string,
    folderPath: string,
    recursive: boolean,
    token: string
): Promise<IFileItem[]> {
    const headers = { Authorization: `Bearer ${token}` };
    const files: IFileItem[] = [];

    async function scanFolder(path: string): Promise<void> {
        const encodedPath = path
            .split("/")
            .map(s => encodeURIComponent(s))
            .join("/");

        let url: string = path === ""
            ? `https://graph.microsoft.com/v1.0/drives/${driveId}/root/children`
            : `https://graph.microsoft.com/v1.0/drives/${driveId}/root:/${encodedPath}:/children`;

        while (url) {
            const resp: any = await axios.get(url, { headers });
            const items: any[] = resp.data.value ?? [];

            for (const item of items) {
                if (item.folder) {
                    if (recursive) {
                        const childPath = path ? `${path}/${item.name}` : item.name;
                        await scanFolder(childPath);
                    }
                } else if (item.file) {
                    const ext = item.name.substring(item.name.lastIndexOf(".")).toLowerCase();
                    if (SUPPORTED_EXTENSIONS.includes(ext)) {
                        files.push({
                            id: item.id,
                            name: item.name,
                            webUrl: item.webUrl ?? "",
                            path: path ? `${path}/${item.name}` : item.name,
                            quickXorHash: item.file?.hashes?.quickXorHash ?? "",
                            lastModified: item.lastModifiedDateTime ?? "",
                        });
                    }
                }
            }

            url = resp.data["@odata.nextLink"] ?? null;
        }
    }

    const cleanPath = folderPath.replace(/^\/+/, "").replace(/\/+$/, "");
    await scanFolder(cleanPath);
    return files;
}

// ─── File download ────────────────────────────────────────────────────────────

async function downloadFile(driveId: string, itemId: string, token: string): Promise<Buffer> {
    const resp = await axios.get(
        `https://graph.microsoft.com/v1.0/drives/${driveId}/items/${itemId}/content`,
        {
            headers: { Authorization: `Bearer ${token}` },
            responseType: "arraybuffer",
            maxContentLength: 100 * 1024 * 1024,
        }
    );
    return Buffer.from(resp.data);
}

// ─── Cognigy file upload ──────────────────────────────────────────────────────

interface ISyncMetadata {
    spItemId?: string;
    spHash?: string;
    uploaded?: string;
}

/**
 * Upload a file buffer to Cognigy Knowledge AI via the Management REST API.
 * POST {apiUrl}/v2.0/knowledgestores/{storeId}/sources/upload
 */
async function uploadFileToCognigy(
    apiUrl: string,
    apiKey: string,
    storeId: string,
    fileName: string,
    fileBuffer: Buffer,
    tags: string[],
): Promise<void> {
    // Use form-data package (available as transitive dep via request, or direct dep)
    const FormData = require("form-data");
    const form = new FormData();

    if (tags && tags.length > 0) {
        form.append("tags", tags.join(","));
    }
    form.append("file", fileBuffer, {
        filename: fileName,
        contentType: getMimeType(fileName),
    });

    await axios.post(
        `${apiUrl}/v2.0/knowledgestores/${storeId}/sources/upload`,
        form,
        {
            headers: {
                "X-API-Key": apiKey,
                ...form.getHeaders(),
            },
            timeout: 120_000,
            maxBodyLength: 100 * 1024 * 1024,
            validateStatus: (s) => s >= 200 && s < 300,
        },
    );
}

// ─── Knowledge Connector ──────────────────────────────────────────────────────

export const sharepointKnowledgeConnector = createKnowledgeConnector({
    type: "sharepointKnowledgeConnector",
    label: "SharePoint Knowledge",
    summary: "Imports files from SharePoint into Cognigy Knowledge AI using Microsoft Graph API.",
    fields: [
        {
            key: "authMethod",
            label: "Authentication Method",
            type: "select",
            defaultValue: "appRegistration",
            params: {
                required: true,
                options: [
                    {
                        label: "App Registration (Tenant ID / Client ID / Secret) — Recommended",
                        value: "appRegistration"
                    },
                    {
                        label: "Bearer Token — For testing or delegated access",
                        value: "bearer"
                    }
                ]
            }
        },
        {
            key: "graphAuth",
            label: "Microsoft Graph Connection",
            type: "connection",
            params: {
                connectionType: "graphCloud",
                required: true
            },
            condition: {
                key: "authMethod",
                value: "appRegistration"
            }
        },
        {
            key: "bearerToken",
            label: "Bearer Token",
            type: "text",
            description: "Paste your Azure AD or SSO access token. Tokens typically expire after 1 hour.",
            params: { required: true },
            condition: {
                key: "authMethod",
                value: "bearer"
            }
        },
        {
            key: "siteUrl",
            label: "SharePoint Site or Folder URL",
            type: "text",
            description: "SharePoint site, library, or folder URL. Library and path are auto-detected from the URL.",
            params: { required: true }
        },
        {
            key: "libraryName",
            label: "Document Library Override",
            type: "text",
            defaultValue: "",
            description: "Overrides the library name auto-detected from the URL, e.g. Shared Documents.",
            params: { required: false }
        },
        {
            key: "folderPath",
            label: "Subfolder Path Override",
            type: "text",
            defaultValue: "",
            description: "Subfolder path within the library, e.g. HR Policies/2025. Leave empty to scan the library root.",
            params: { required: false }
        },
        {
            key: "recursive",
            label: "Scan Subfolders",
            type: "toggle",
            defaultValue: false,
            description: "When enabled, all subfolders are scanned recursively."
        },
        {
            key: "sourceTags",
            label: "Knowledge Tags",
            type: "chipInput",
            defaultValue: ["sharepoint"],
            description: "Tags applied to all uploaded knowledge sources. Press ENTER to add a tag."
        },
        {
            key: "cognigyApiUrl",
            label: "Cognigy API URL",
            type: "text",
            description: "Base URL of your Cognigy.AI API, e.g. https://api-jetstar-dev.cognigy.cloud",
            params: { required: true }
        },
        {
            key: "cognigyApiKey",
            label: "Cognigy API Key",
            type: "text",
            description: "API key from Profile > API Keys in Cognigy.AI",
            params: { required: true }
        },
        {
            key: "knowledgeStoreId",
            label: "Knowledge Store ID",
            type: "text",
            description: "The 24-character ID of the target knowledge store (visible in the store URL in Cognigy.AI)",
            params: { required: true }
        },
    ] as const,
    sections: [
        {
            key: "authSection",
            label: "Authentication",
            defaultCollapsed: false,
            fields: ["authMethod", "graphAuth", "bearerToken"]
        },
        {
            key: "contentSection",
            label: "SharePoint Content",
            defaultCollapsed: false,
            fields: ["siteUrl", "libraryName", "folderPath", "recursive"]
        },
        {
            key: "cognigySection",
            label: "Cognigy Upload Settings",
            defaultCollapsed: false,
            fields: ["cognigyApiUrl", "cognigyApiKey", "knowledgeStoreId"]
        }
    ],
    form: [
        { type: "section", key: "authSection" },
        { type: "section", key: "contentSection" },
        { type: "section", key: "cognigySection" },
        { type: "field", key: "sourceTags" },
    ],
    function: async ({ config }) => {
        const {
            authMethod,
            graphAuth,
            bearerToken,
            siteUrl,
            libraryName,
            folderPath,
            recursive,
            sourceTags,
            cognigyApiUrl,
            cognigyApiKey,
            knowledgeStoreId,
        } = config;

        const apiUrl   = (cognigyApiUrl as string).trim().replace(/\/+$/, "");
        const apiKey   = (cognigyApiKey as string).trim();
        const storeId  = (knowledgeStoreId as string).trim();
        const tags     = sourceTags as string[];

        // ── 1. Parse the SharePoint URL ──────────────────────────────────────
        const parsed = parseSharePointUrl(siteUrl as string);
        const effectiveSiteUrl    = parsed.siteUrl;
        const effectiveLibrary    = (libraryName as string)?.trim() || parsed.libraryName;
        const effectiveFolderPath = (folderPath as string)?.trim() || parsed.folderPath;

        console.log(`[SharePoint KC] Site: ${effectiveSiteUrl}`);
        console.log(`[SharePoint KC] Library: "${effectiveLibrary}"${effectiveFolderPath ? `, Folder: "${effectiveFolderPath}"` : ""}`);

        // ── 2. Authenticate with Microsoft Graph ─────────────────────────────
        const conn = graphAuth as any;
        const token = await getAccessToken(
            authMethod as string,
            conn?.tenantId ?? "",
            conn?.clientId ?? "",
            conn?.clientSecret ?? "",
            (bearerToken as string) ?? ""
        );

        // ── 3. Resolve site and drive IDs ────────────────────────────────────
        const siteId  = await getSiteId(effectiveSiteUrl, token);
        const driveId = await getDriveId(siteId, effectiveLibrary, token);

        // ── 4. List supported files in SharePoint ────────────────────────────
        const spFiles = await listFiles(driveId, effectiveFolderPath, recursive as boolean, token);
        console.log(`[SharePoint KC] Found ${spFiles.length} supported file(s) in SharePoint`);

        if (spFiles.length === 0) {
            console.log("[SharePoint KC] Nothing to do.");
            return;
        }

        // ── 5. List existing Cognigy sources for incremental sync ─────────────
        let existingSources: ManagedSource[] = [];
        try {
            existingSources = await listKnowledgeSources(apiUrl, apiKey, storeId);
            console.log(`[SharePoint KC] ${existingSources.length} existing source(s) in knowledge store`);
        } catch (err) {
            console.warn(`[SharePoint KC] Could not list existing sources: ${(err as Error).message}. All files will be uploaded.`);
        }

        // Build map: source name → { sourceId, metadata }
        const existingMap = new Map<string, { sourceId: string; meta: ISyncMetadata }>();
        const originalSourceIds = new Set<string>();
        for (const src of existingSources) {
            let meta: ISyncMetadata = {};
            try { meta = JSON.parse(src.description || "{}"); } catch { /* ignore */ }
            existingMap.set(src.name, { sourceId: src._id, meta });
            originalSourceIds.add(src._id);
        }

        // ── 6. Process each SharePoint file ──────────────────────────────────
        const currentFileNames = new Set<string>();

        for (const file of spFiles) {
            currentFileNames.add(file.name);

            // Effective hash: prefer quickXorHash, fall back to lastModified
            const effectiveHash = file.quickXorHash || file.lastModified;
            const existing = existingMap.get(file.name);

            if (existing) {
                if (existing.meta.spHash && existing.meta.spHash === effectiveHash) {
                    console.log(`[SharePoint KC] Unchanged — skipping: ${file.name}`);
                    continue;
                }
                if (!existing.meta.spHash) {
                    // No hash stored from previous run (PATCH may have failed) — skip conservatively
                    console.log(`[SharePoint KC] No sync metadata found — skipping: ${file.name} (delete and re-run to force refresh)`);
                    continue;
                }
                // Hash mismatch — file was updated in SharePoint
                console.log(`[SharePoint KC] Content changed — replacing: ${file.name}`);
                try {
                    await deleteKnowledgeSourceById(apiUrl, apiKey, storeId, existing.sourceId);
                } catch (delErr) {
                    console.warn(`[SharePoint KC] Could not delete old source for "${file.name}": ${(delErr as Error).message}`);
                }
            }

            // Download from SharePoint and upload to Cognigy
            try {
                const buffer = await downloadFile(driveId, file.id, token);
                console.log(`[SharePoint KC] Uploading "${file.name}" (${(buffer.length / 1024).toFixed(1)} KB)`);
                await uploadFileToCognigy(apiUrl, apiKey, storeId, file.name, buffer, tags);

                // Give Cognigy a moment to register the new source, then store sync metadata
                await sleep(2000);
                try {
                    const currentSources = await listKnowledgeSources(apiUrl, apiKey, storeId);
                    const newSource = currentSources.find(
                        s => s.name === file.name && !originalSourceIds.has(s._id)
                    );
                    if (newSource) {
                        const meta: ISyncMetadata = {
                            spItemId: file.id,
                            spHash: effectiveHash,
                            uploaded: new Date().toISOString(),
                        };
                        await patchKnowledgeSourceDescription(apiUrl, apiKey, storeId, newSource._id, JSON.stringify(meta));
                        originalSourceIds.add(newSource._id); // prevent re-patching on next file
                        console.log(`[SharePoint KC] Sync metadata stored for "${file.name}"`);
                    } else {
                        console.warn(`[SharePoint KC] Source for "${file.name}" not yet visible in API — sync metadata will be stored on next run`);
                    }
                } catch (patchErr) {
                    console.warn(`[SharePoint KC] Could not store sync metadata for "${file.name}": ${(patchErr as Error).message}`);
                }
            } catch (uploadErr) {
                console.error(`[SharePoint KC] Failed to upload "${file.name}": ${(uploadErr as Error).message}`);
            }
        }

        // ── 7. Remove sources for files no longer in SharePoint ──────────────
        for (const [name, { sourceId }] of existingMap) {
            if (!currentFileNames.has(name)) {
                console.log(`[SharePoint KC] Removing stale source: "${name}"`);
                try {
                    await deleteKnowledgeSourceById(apiUrl, apiKey, storeId, sourceId);
                } catch (err) {
                    console.warn(`[SharePoint KC] Could not remove stale source "${name}": ${(err as Error).message}`);
                }
            }
        }

        console.log("[SharePoint KC] Done.");
    },
});
