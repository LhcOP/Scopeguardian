import {
  countTokens,
  truncateToTokens,
  selectItemsWithinBudget,
  deduplicateTexts,
} from "../src/utils/tokenOptimizer";

describe("countTokens", () => {
  it("estimates token count as ceil(length/4)", () => {
    expect(countTokens("hello")).toBe(2); // 5 chars → ceil(5/4) = 2
    expect(countTokens("abcd")).toBe(1);  // 4 chars → 1
    expect(countTokens("")).toBe(0);
  });
});

describe("truncateToTokens", () => {
  it("returns original text when within budget", () => {
    const text = "short";
    expect(truncateToTokens(text, 100)).toBe(text);
  });

  it("truncates and appends ellipsis marker when over budget", () => {
    const text = "a".repeat(1000);
    const result = truncateToTokens(text, 10);
    expect(result.endsWith("\n...[truncated]")).toBe(true);
    expect(result.length).toBeLessThan(text.length);
  });
});

describe("selectItemsWithinBudget", () => {
  const items = ["short", "medium length text here", "this is quite a long string that uses more tokens than the others"];

  it("selects items fitting within token budget", () => {
    const selected = selectItemsWithinBudget(items, (s) => s, 5);
    expect(selected.length).toBeGreaterThan(0);
    expect(selected.length).toBeLessThanOrEqual(items.length);
  });

  it("returns empty array when budget is zero", () => {
    const selected = selectItemsWithinBudget(items, (s) => s, 0);
    expect(selected).toHaveLength(0);
  });
});

describe("deduplicateTexts", () => {
  it("removes near-duplicate strings", () => {
    const texts = [
      "the quick brown fox jumps over the lazy dog",
      "the quick brown fox jumps over the lazy dog", // exact duplicate
      "completely different text about something else",
    ];
    const result = deduplicateTexts(texts);
    expect(result.length).toBe(2);
  });

  it("keeps distinct texts", () => {
    const texts = ["apple orange mango", "car truck bus", "cat dog fish"];
    expect(deduplicateTexts(texts)).toHaveLength(3);
  });

  it("handles empty input", () => {
    expect(deduplicateTexts([])).toHaveLength(0);
  });
});
