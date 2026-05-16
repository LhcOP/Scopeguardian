/**
 * Lightweight token estimator (4 chars ≈ 1 token for English text).
 * Avoids a full tiktoken dependency for simple chunking decisions.
 */
export function countTokens(text: string): number {
  return Math.ceil(text.length / 4);
}

/**
 * Truncate text to a maximum token budget, appending an ellipsis marker.
 */
export function truncateToTokens(text: string, maxTokens: number): string {
  const maxChars = maxTokens * 4;
  if (text.length <= maxChars) return text;
  return text.slice(0, maxChars - 20) + "\n...[truncated]";
}

/**
 * Select the most relevant scope items to include in a prompt given a
 * token budget. Items are assumed to be pre-ranked by relevance.
 */
export function selectItemsWithinBudget<T>(
  items: T[],
  serializer: (item: T) => string,
  tokenBudget: number
): T[] {
  const selected: T[] = [];
  let used = 0;
  for (const item of items) {
    const cost = countTokens(serializer(item));
    if (used + cost > tokenBudget) break;
    selected.push(item);
    used += cost;
  }
  return selected;
}

/**
 * Compress a list of strings by removing duplicates and near-duplicates
 * using simple Jaccard similarity (word-level bigrams).
 */
export function deduplicateTexts(texts: string[], threshold = 0.8): string[] {
  const unique: string[] = [];
  const bigramSets: Set<string>[] = [];

  for (const text of texts) {
    const bg = bigrams(text.toLowerCase());
    const isDuplicate = bigramSets.some((existing) => jaccardSimilarity(existing, bg) >= threshold);
    if (!isDuplicate) {
      unique.push(text);
      bigramSets.push(bg);
    }
  }
  return unique;
}

function bigrams(text: string): Set<string> {
  const words = text.split(/\s+/);
  const result = new Set<string>();
  for (let i = 0; i < words.length - 1; i++) {
    result.add(`${words[i]} ${words[i + 1]}`);
  }
  return result;
}

function jaccardSimilarity(a: Set<string>, b: Set<string>): number {
  if (a.size === 0 && b.size === 0) return 1;
  const intersection = [...a].filter((x) => b.has(x)).length;
  const union = new Set([...a, ...b]).size;
  return union === 0 ? 0 : intersection / union;
}
