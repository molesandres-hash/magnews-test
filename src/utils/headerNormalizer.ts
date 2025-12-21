import type { NormalizedHeader } from '@/types/survey';

/**
 * Normalize a header string by cleaning up special characters,
 * smart quotes, and extracting value source information.
 */
export function normalizeHeader(rawHeader: string): NormalizedHeader {
  let cleaned = rawHeader;

  // Convert smart quotes to ASCII
  cleaned = cleaned.replace(/['']/g, "'");
  cleaned = cleaned.replace(/[""]/g, '"');

  // Replace carriage returns and line breaks
  cleaned = cleaned.replace(/_x000D_/g, '\n');
  cleaned = cleaned.replace(/\r\n/g, '\n');
  cleaned = cleaned.replace(/\r/g, '\n');

  // Collapse multiple spaces and trim
  cleaned = cleaned.replace(/\s+/g, ' ').trim();

  // Determine value source and remove suffix
  let valueSource: 'values' | 'labels' | 'unknown' = 'unknown';
  
  if (/\(values\)\s*$/i.test(cleaned)) {
    valueSource = 'values';
    cleaned = cleaned.replace(/\s*\(values\)\s*$/i, '').trim();
  } else if (/\(labels\)\s*$/i.test(cleaned)) {
    valueSource = 'labels';
    cleaned = cleaned.replace(/\s*\(labels\)\s*$/i, '').trim();
  }

  // Remove extra punctuation around numbering
  cleaned = cleaned.replace(/\t/g, ' ');

  return {
    rawHeader,
    cleanedHeader: cleaned,
    valueSource,
  };
}

/**
 * Extract question key (like "4.1") and question text from a cleaned header.
 */
export function extractKeyAndText(cleanedHeader: string): {
  questionKey: string | null;
  questionText: string;
  blockId: number | null;
  subId: number;
} {
  // Try to match patterns like "4.", "4.1", "10.2" at the start
  const regex = /^\s*(\d+(?:\.\d+)?)\s*[.):\-–]?\s*(.*)$/;
  const match = cleanedHeader.match(regex);

  if (match) {
    const questionKey = match[1];
    const questionText = match[2].trim() || cleanedHeader;
    
    // Parse block and sub IDs
    const parts = questionKey.split('.');
    const blockId = parseInt(parts[0], 10);
    const subId = parts[1] ? parseInt(parts[1], 10) : 0;

    return {
      questionKey,
      questionText,
      blockId,
      subId,
    };
  }

  // Also try pattern where question number is embedded after a prefix
  // e.g., "Comitato X - 4.1 Question text"
  const embeddedRegex = /[-–]\s*(\d+(?:\.\d+)?)\s*[.):\-–]?\s*(.*)$/;
  const embeddedMatch = cleanedHeader.match(embeddedRegex);

  if (embeddedMatch) {
    const questionKey = embeddedMatch[1];
    const questionText = embeddedMatch[2].trim() || cleanedHeader;
    
    const parts = questionKey.split('.');
    const blockId = parseInt(parts[0], 10);
    const subId = parts[1] ? parseInt(parts[1], 10) : 0;

    return {
      questionKey,
      questionText,
      blockId,
      subId,
    };
  }

  return {
    questionKey: null,
    questionText: cleanedHeader,
    blockId: null,
    subId: 0,
  };
}

/**
 * Get a short display text for a question (truncated if needed)
 */
export function getShortQuestionText(text: string, maxLength: number = 80): string {
  if (text.length <= maxLength) return text;
  return text.slice(0, maxLength - 3) + '...';
}
