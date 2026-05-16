import { SearchClient, AzureKeyCredential, SearchIndexClient } from "@azure/search-documents";
import { ScopeItem } from "../models/ProjectScope";

const INDEX_NAME = "scope-items";

interface ScopeItemDocument {
  id: string;
  projectId: string;
  title: string;
  description: string;
  deliverables: string;
  acceptanceCriteria: string;
  tags: string;
  contentVector: number[];
}

function getSearchClient(): SearchClient<ScopeItemDocument> {
  const endpoint = process.env.AI_SEARCH_ENDPOINT;
  const key = process.env.AI_SEARCH_KEY;
  if (!endpoint || !key) throw new Error("AI_SEARCH_ENDPOINT or AI_SEARCH_KEY is not set");
  return new SearchClient<ScopeItemDocument>(endpoint, INDEX_NAME, new AzureKeyCredential(key));
}

function getIndexClient(): SearchIndexClient {
  const endpoint = process.env.AI_SEARCH_ENDPOINT;
  const key = process.env.AI_SEARCH_KEY;
  if (!endpoint || !key) throw new Error("AI_SEARCH_ENDPOINT or AI_SEARCH_KEY is not set");
  return new SearchIndexClient(endpoint, new AzureKeyCredential(key));
}

export async function ensureIndexExists(): Promise<void> {
  const indexClient = getIndexClient();
  try {
    await indexClient.getIndex(INDEX_NAME);
  } catch {
    await indexClient.createIndex({
      name: INDEX_NAME,
      fields: [
        { name: "id", type: "Edm.String", key: true, filterable: true },
        { name: "projectId", type: "Edm.String", filterable: true, facetable: true },
        { name: "title", type: "Edm.String", searchable: true },
        { name: "description", type: "Edm.String", searchable: true },
        { name: "deliverables", type: "Edm.String", searchable: true },
        { name: "acceptanceCriteria", type: "Edm.String", searchable: true },
        { name: "tags", type: "Edm.String", searchable: true, filterable: true },
        {
          name: "contentVector",
          type: "Collection(Edm.Single)",
          searchable: true,
          vectorSearchDimensions: 1536,
          vectorSearchProfileName: "scope-vector-profile",
        },
      ],
      vectorSearch: {
        algorithms: [{ name: "hnsw-config", kind: "hnsw", parameters: { m: 4, efConstruction: 400, efSearch: 500, metric: "cosine" } }],
        profiles: [{ name: "scope-vector-profile", algorithmConfigurationName: "hnsw-config" }],
      },
    });
  }
}

export async function upsertScopeItems(projectId: string, items: ScopeItem[], embeddings: number[][]): Promise<void> {
  const client = getSearchClient();
  const documents: ScopeItemDocument[] = items.map((item, i) => ({
    id: `${projectId}-${item.id}`,
    projectId,
    title: item.title,
    description: item.description,
    deliverables: item.deliverables.join(" | "),
    acceptanceCriteria: item.acceptanceCriteria.join(" | "),
    tags: item.tags.join(", "),
    contentVector: embeddings[i],
  }));
  await client.mergeOrUploadDocuments(documents);
}

export async function vectorSearch(
  projectId: string,
  queryVector: number[],
  topK = 5
): Promise<{ item: Partial<ScopeItemDocument>; score: number }[]> {
  const client = getSearchClient();
  const results = await client.search("*", {
    filter: `projectId eq '${projectId}'`,
    vectorSearchOptions: {
      queries: [{ kind: "vector", vector: queryVector, kNearestNeighborsCount: topK, fields: ["contentVector"] }],
    },
    select: ["id", "projectId", "title", "description", "deliverables", "acceptanceCriteria", "tags"],
    top: topK,
  });

  const hits: { item: Partial<ScopeItemDocument>; score: number }[] = [];
  for await (const result of results.results) {
    hits.push({ item: result.document, score: result.score ?? 0 });
  }
  return hits;
}

export async function deleteScopeItemsByProject(projectId: string): Promise<void> {
  const client = getSearchClient();
  const results = await client.search("*", {
    filter: `projectId eq '${projectId}'`,
    select: ["id"],
  });
  const toDelete: ScopeItemDocument[] = [];
  for await (const r of results.results) {
    toDelete.push({ id: r.document.id } as ScopeItemDocument);
  }
  if (toDelete.length > 0) {
    await client.deleteDocuments("id", toDelete.map((d) => d.id));
  }
}
