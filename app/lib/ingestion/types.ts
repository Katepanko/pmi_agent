export type ExtractionStatus = "extracted" | "partial" | "failed";

export type ExtractedSegment = {
  ordinal: number;
  location: string;
  kind: "text" | "table" | "record" | "image" | "metadata";
  content: string;
  structured?: Record<string, unknown>;
  confidence: "high" | "medium" | "low";
};

export type ExtractedDocument = {
  fileId: string;
  fileName: string;
  fileType: string;
  mediaType: string;
  sizeBytes: number;
  status: ExtractionStatus;
  metadata: Record<string, unknown>;
  rawText: string;
  structuredElements: ExtractedSegment[];
  tables: Array<Record<string, unknown>>;
  sheets: Array<Record<string, unknown>>;
  slides: Array<Record<string, unknown>>;
  pages: Array<Record<string, unknown>>;
  images: Array<Record<string, unknown>>;
  extractionWarnings: string[];
};
