import * as mammoth from "mammoth";
import * as XLSX from "xlsx";
import { PDFParse } from "pdf-parse";
import { simpleParser } from "mailparser";
import MsgReader from "@kenjiuno/msgreader";

export const SUPPORTED_DOC_EXTENSIONS = [".docx", ".pdf", ".xlsx", ".txt", ".md", ".eml", ".msg"];

export function isSupportedDocument(fileName: string): boolean {
  const lower = fileName.toLowerCase();
  return SUPPORTED_DOC_EXTENSIONS.some((ext) => lower.endsWith(ext));
}

/**
 * Extracts plain text from an estimation-basis document.
 * Returns null when the format is unsupported or extraction fails.
 */
export async function extractTextFromFile(fileName: string, buffer: Buffer): Promise<string | null> {
  const lower = fileName.toLowerCase();
  try {
    if (lower.endsWith(".docx")) {
      const result = await mammoth.extractRawText({ buffer });
      return result.value;
    }
    if (lower.endsWith(".xlsx")) {
      const workbook = XLSX.read(buffer, { type: "buffer" });
      return workbook.SheetNames.map((name) => {
        const csv = XLSX.utils.sheet_to_csv(workbook.Sheets[name]);
        return `--- Sheet: ${name} ---\n${csv}`;
      }).join("\n\n");
    }
    if (lower.endsWith(".pdf")) {
      const parser = new PDFParse({ data: buffer });
      try {
        const result = await parser.getText();
        return result.text;
      } finally {
        await parser.destroy();
      }
    }
    if (lower.endsWith(".txt") || lower.endsWith(".md")) {
      return buffer.toString("utf8");
    }
    if (lower.endsWith(".eml")) {
      const mail = await simpleParser(buffer);
      const parts = [
        mail.from?.text ? `From: ${mail.from.text}` : "",
        mail.to && !Array.isArray(mail.to) ? `To: ${mail.to.text}` : "",
        mail.subject ? `Subject: ${mail.subject}` : "",
        mail.date ? `Date: ${mail.date.toISOString()}` : "",
        "",
        mail.text ?? "",
      ];
      return parts.filter((p, i) => p || i === 4).join("\n");
    }
    if (lower.endsWith(".msg")) {
      const arrayBuffer = buffer.buffer.slice(buffer.byteOffset, buffer.byteOffset + buffer.byteLength);
      const reader = new MsgReader(arrayBuffer as ArrayBuffer);
      const data = reader.getFileData();
      const parts = [
        data.senderName ? `From: ${data.senderName} <${data.senderEmail ?? ""}>` : "",
        data.subject ? `Subject: ${data.subject}` : "",
        data.messageDeliveryTime ? `Date: ${data.messageDeliveryTime}` : "",
        "",
        data.body ?? "",
      ];
      return parts.filter((p, i) => p || i === 3).join("\n");
    }
    return null;
  } catch {
    return null;
  }
}
