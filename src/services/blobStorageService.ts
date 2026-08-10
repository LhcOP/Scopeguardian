import { BlobServiceClient, ContainerClient } from "@azure/storage-blob";
import { ProjectScope } from "../models/ProjectScope";
import { compareVersions } from "../utils/versionCompare";

const SCOPE_CONTAINER = "master-scopes";
const SUMMARY_CONTAINER = "scope-summaries";

function getBlobServiceClient(): BlobServiceClient {
  const connStr = process.env.AZURE_STORAGE_CONNECTION_STRING;
  if (!connStr) throw new Error("AZURE_STORAGE_CONNECTION_STRING is not set");
  return BlobServiceClient.fromConnectionString(connStr);
}

async function ensureContainer(client: BlobServiceClient, name: string): Promise<ContainerClient> {
  const container = client.getContainerClient(name);
  await container.createIfNotExists();
  return container;
}

export async function uploadMasterScope(scope: ProjectScope): Promise<void> {
  const client = getBlobServiceClient();
  const container = await ensureContainer(client, SCOPE_CONTAINER);
  const blobName = `${scope.projectId}/scope-v${scope.version}.json`;
  const blob = container.getBlockBlobClient(blobName);
  const content = JSON.stringify(scope, null, 2);
  await blob.upload(content, Buffer.byteLength(content), {
    blobHTTPHeaders: { blobContentType: "application/json" },
  });
}

export async function downloadMasterScope(projectId: string): Promise<ProjectScope | null> {
  const client = getBlobServiceClient();
  const container = await ensureContainer(client, SCOPE_CONTAINER);

  // List blobs for this project and pick the latest version
  const blobs: { name: string; version: string }[] = [];
  for await (const blob of container.listBlobsFlat({ prefix: `${projectId}/scope-v` })) {
    const match = blob.name.match(/scope-v(\d+(?:\.\d+)*)\.json$/);
    if (match) {
      blobs.push({ name: blob.name, version: match[1] });
    }
  }

  if (blobs.length === 0) return null;

  blobs.sort((a, b) => compareVersions(b.version, a.version));
  const latest = container.getBlockBlobClient(blobs[0].name);
  const download = await latest.download();
  const text = await streamToString(download.readableStreamBody!);
  return JSON.parse(text) as ProjectScope;
}

export async function uploadScopeMarkdown(projectId: string, markdown: string): Promise<void> {
  const client = getBlobServiceClient();
  const container = await ensureContainer(client, SUMMARY_CONTAINER);
  const blobName = `${projectId}/scope-summary.md`;
  const blob = container.getBlockBlobClient(blobName);
  await blob.upload(markdown, Buffer.byteLength(markdown), {
    blobHTTPHeaders: { blobContentType: "text/markdown" },
  });
}

export async function downloadScopeMarkdown(projectId: string): Promise<string | null> {
  const client = getBlobServiceClient();
  const container = client.getContainerClient(SUMMARY_CONTAINER);
  const blob = container.getBlockBlobClient(`${projectId}/scope-summary.md`);
  try {
    const download = await blob.download();
    return await streamToString(download.readableStreamBody!);
  } catch {
    return null;
  }
}

async function streamToString(stream: NodeJS.ReadableStream): Promise<string> {
  const chunks: Uint8Array[] = [];
  for await (const chunk of stream) {
    chunks.push(typeof chunk === "string" ? Buffer.from(chunk) : (chunk as Uint8Array));
  }
  return Buffer.concat(chunks).toString("utf-8");
}
