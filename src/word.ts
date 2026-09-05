import {
  AlignmentType,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  TextRun,
} from 'docx';
import JSZip from 'jszip';
import { TextDecoder } from 'node:util';
import { type AuthConfig } from './config.js';
import { MAX_SAFE_UPLOAD_BYTES, safeCreateDriveFile, type VerifiedDriveItem } from './drive.js';
import { readSecureLocalFile } from './local-file.js';
import { validateOoxmlBytes, type OoxmlValidation } from './ooxml.js';

export const WORD_DOCUMENT_LIMITS = {
  titleCharacters: 200,
  metadataEntries: 50,
  metadataLabelCharacters: 100,
  metadataValueCharacters: 2000,
  sections: 100,
  sectionHeadingCharacters: 300,
  blocksPerSection: 100,
  totalBlocks: 2000,
  blockCharacters: 5000,
  totalCharacters: 250_000,
  inputJsonBytes: 1024 * 1024,
  docxBytes: MAX_SAFE_UPLOAD_BYTES,
} as const;

const REQUIRED_DOCX_ENTRIES = new Set(['[Content_Types].xml', '_rels/.rels', 'docProps/core.xml', 'word/document.xml']);
const FIXED_DOCUMENT_TIMESTAMP = '2000-01-01T00:00:00.000Z';
const FIXED_ZIP_DATE = new Date(FIXED_DOCUMENT_TIMESTAMP);
const ZIP_COMPRESSION_LEVEL = 9;

export interface WordMetadataEntry {
  readonly label: string;
  readonly value: string;
}

export interface WordDocumentSection {
  readonly heading: string;
  readonly paragraphs?: readonly string[];
  readonly bullets?: readonly string[];
}

export interface WordDocumentModel {
  readonly schemaVersion: 1;
  readonly title: string;
  readonly metadata?: readonly WordMetadataEntry[];
  readonly sections: readonly WordDocumentSection[];
}

export interface GeneratedWordDocument {
  bytes: Buffer;
  size: number;
  sha256: string;
}

function isRecord(value: unknown): value is Record<string, unknown> {
  if (value === null || typeof value !== 'object' || Array.isArray(value)) return false;
  const prototype = Object.getPrototypeOf(value);
  return prototype === Object.prototype || prototype === null;
}

function assertExactKeys(value: Record<string, unknown>, allowed: readonly string[], context: string): void {
  const allowedKeys = new Set(allowed);
  const unexpected = Reflect.ownKeys(value).find((key) => typeof key !== 'string' || !allowedKeys.has(key));
  if (unexpected !== undefined) throw new Error(`${context} contains an unexpected key`);
}

function codePointLength(value: string): number {
  return Array.from(value).length;
}

function hasInvalidXmlCharacter(value: string): boolean {
  return Array.from(value).some((character) => {
    const codePoint = character.codePointAt(0) ?? 0;
    return (codePoint < 0x20 && codePoint !== 0x09 && codePoint !== 0x0a && codePoint !== 0x0d)
      || codePoint === 0xfffe
      || codePoint === 0xffff;
  });
}

function parseText(value: unknown, context: string, maximum: number): string {
  if (typeof value !== 'string') throw new Error(`${context} must be a string`);
  const length = codePointLength(value);
  if (length === 0 || value.trim().length === 0 || length > maximum) {
    throw new Error(`${context} must contain 1 through ${maximum} characters`);
  }
  if (!value.isWellFormed() || hasInvalidXmlCharacter(value)) {
    throw new Error(`${context} contains a character that cannot be represented safely in XML`);
  }
  return value;
}

function assertDenseArray(value: unknown[], context: string): void {
  for (let index = 0; index < value.length; index += 1) {
    if (!Object.hasOwn(value, index)) throw new Error(`${context} must not contain missing entries`);
  }
  if (Reflect.ownKeys(value).some((key) => (
    key !== 'length'
    && (typeof key !== 'string' || !/^(?:0|[1-9]\d*)$/u.test(key) || Number(key) >= value.length)
  ))) {
    throw new Error(`${context} contains an unexpected key`);
  }
}

function parseTextList(value: unknown, context: string): string[] {
  if (!Array.isArray(value) || value.length === 0 || value.length > WORD_DOCUMENT_LIMITS.blocksPerSection) {
    throw new Error(`${context} must contain 1 through ${WORD_DOCUMENT_LIMITS.blocksPerSection} strings`);
  }
  assertDenseArray(value, context);
  return value.map((entry, index) => parseText(entry, `${context}[${index}]`, WORD_DOCUMENT_LIMITS.blockCharacters));
}

export function parseWordDocumentModel(value: unknown): WordDocumentModel {
  if (!isRecord(value)) throw new Error('Word document model must be a plain object');
  assertExactKeys(value, ['schemaVersion', 'title', 'metadata', 'sections'], 'Word document model');
  if (value.schemaVersion !== 1) throw new Error('Word document schemaVersion must be 1');
  const title = parseText(value.title, 'Word document title', WORD_DOCUMENT_LIMITS.titleCharacters);

  let totalCharacters = codePointLength(title);
  let metadata: WordMetadataEntry[] | undefined;
  if (value.metadata !== undefined) {
    if (!Array.isArray(value.metadata) || value.metadata.length > WORD_DOCUMENT_LIMITS.metadataEntries) {
      throw new Error(`Word document metadata must contain at most ${WORD_DOCUMENT_LIMITS.metadataEntries} entries`);
    }
    assertDenseArray(value.metadata, 'Word document metadata');
    metadata = value.metadata.map((entry, index) => {
      if (!isRecord(entry)) throw new Error(`Word document metadata[${index}] must be a plain object`);
      assertExactKeys(entry, ['label', 'value'], `Word document metadata[${index}]`);
      const label = parseText(entry.label, `Word document metadata[${index}].label`, WORD_DOCUMENT_LIMITS.metadataLabelCharacters);
      const entryValue = parseText(entry.value, `Word document metadata[${index}].value`, WORD_DOCUMENT_LIMITS.metadataValueCharacters);
      totalCharacters += codePointLength(label) + codePointLength(entryValue);
      if (totalCharacters > WORD_DOCUMENT_LIMITS.totalCharacters) {
        throw new Error(`Word document exceeds ${WORD_DOCUMENT_LIMITS.totalCharacters} total characters`);
      }
      return { label, value: entryValue };
    });
  }

  if (!Array.isArray(value.sections) || value.sections.length === 0 || value.sections.length > WORD_DOCUMENT_LIMITS.sections) {
    throw new Error(`Word document sections must contain 1 through ${WORD_DOCUMENT_LIMITS.sections} entries`);
  }
  assertDenseArray(value.sections, 'Word document sections');
  let totalBlocks = 0;
  const sections = value.sections.map((section, index): WordDocumentSection => {
    if (!isRecord(section)) throw new Error(`Word document sections[${index}] must be a plain object`);
    assertExactKeys(section, ['heading', 'paragraphs', 'bullets'], `Word document sections[${index}]`);
    const heading = parseText(section.heading, `Word document sections[${index}].heading`, WORD_DOCUMENT_LIMITS.sectionHeadingCharacters);
    const paragraphs = section.paragraphs === undefined
      ? undefined
      : parseTextList(section.paragraphs, `Word document sections[${index}].paragraphs`);
    const bullets = section.bullets === undefined
      ? undefined
      : parseTextList(section.bullets, `Word document sections[${index}].bullets`);
    if (!paragraphs && !bullets) throw new Error(`Word document sections[${index}] must include paragraphs and/or bullets`);
    const blockCount = (paragraphs?.length ?? 0) + (bullets?.length ?? 0);
    if (blockCount > WORD_DOCUMENT_LIMITS.blocksPerSection) {
      throw new Error(`Word document sections[${index}] exceeds ${WORD_DOCUMENT_LIMITS.blocksPerSection} total blocks`);
    }
    totalBlocks += blockCount;
    if (totalBlocks > WORD_DOCUMENT_LIMITS.totalBlocks) {
      throw new Error(`Word document exceeds ${WORD_DOCUMENT_LIMITS.totalBlocks} total blocks`);
    }
    totalCharacters += codePointLength(heading);
    for (const text of [...(paragraphs ?? []), ...(bullets ?? [])]) {
      totalCharacters += codePointLength(text);
      if (totalCharacters > WORD_DOCUMENT_LIMITS.totalCharacters) {
        throw new Error(`Word document exceeds ${WORD_DOCUMENT_LIMITS.totalCharacters} total characters`);
      }
    }
    return {
      heading,
      ...(paragraphs ? { paragraphs } : {}),
      ...(bullets ? { bullets } : {}),
    };
  });

  if (totalBlocks > WORD_DOCUMENT_LIMITS.totalBlocks) {
    throw new Error(`Word document exceeds ${WORD_DOCUMENT_LIMITS.totalBlocks} total blocks`);
  }
  if (totalCharacters > WORD_DOCUMENT_LIMITS.totalCharacters) {
    throw new Error(`Word document exceeds ${WORD_DOCUMENT_LIMITS.totalCharacters} total characters`);
  }
  return {
    schemaVersion: 1,
    title,
    ...(metadata ? { metadata } : {}),
    sections,
  };
}

export async function validateWordDocumentBytes(value: Uint8Array): Promise<OoxmlValidation> {
  if (value.byteLength === 0 || value.byteLength > WORD_DOCUMENT_LIMITS.docxBytes) {
    throw new Error(`DOCX must contain 1 through ${WORD_DOCUMENT_LIMITS.docxBytes} bytes`);
  }
  return validateOoxmlBytes(value, REQUIRED_DOCX_ENTRIES, 'DOCX');
}

function escapeXmlText(value: string): string {
  const entities: Record<string, string> = {
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&apos;',
  };
  return value.replace(/[&<>"']/gu, (character) => entities[character] ?? character);
}

function deterministicCoreProperties(title: string): Buffer {
  return Buffer.from(
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>${escapeXmlText(title)}</dc:title><dc:creator>Un-named</dc:creator><cp:lastModifiedBy>Un-named</cp:lastModifiedBy><cp:revision>1</cp:revision><dcterms:created xsi:type="dcterms:W3CDTF">${FIXED_DOCUMENT_TIMESTAMP}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${FIXED_DOCUMENT_TIMESTAMP}</dcterms:modified></cp:coreProperties>`,
    'utf8',
  );
}

async function rebuildDeterministicDocx(value: Uint8Array, title: string): Promise<Buffer> {
  const source = await JSZip.loadAsync(value, { checkCRC32: true, createFolders: false });
  const rebuilt = new JSZip();
  const entryNames = Object.keys(source.files).sort((left, right) => left < right ? -1 : left > right ? 1 : 0);
  for (const entryName of entryNames) {
    const entry = source.files[entryName];
    if (!entry) throw new Error('Generated DOCX contained an inaccessible ZIP entry');
    const date = new Date(FIXED_ZIP_DATE.getTime());
    if (entry.dir) {
      rebuilt.file(entryName, null, {
        createFolders: false,
        date,
        dir: true,
        compression: 'STORE',
        unixPermissions: 0o755,
      });
      continue;
    }
    const content = entryName === 'docProps/core.xml'
      ? deterministicCoreProperties(title)
      : await entry.async('nodebuffer');
    rebuilt.file(entryName, content, {
      createFolders: false,
      date,
      compression: 'DEFLATE',
      compressionOptions: { level: ZIP_COMPRESSION_LEVEL },
      unixPermissions: 0o644,
    });
  }
  return rebuilt.generateAsync({
    type: 'nodebuffer',
    platform: 'UNIX',
    compression: 'DEFLATE',
    compressionOptions: { level: ZIP_COMPRESSION_LEVEL },
    streamFiles: false,
    comment: '',
  });
}

export async function createWordDocument(value: unknown): Promise<GeneratedWordDocument> {
  const model = parseWordDocumentModel(value);
  const children: Paragraph[] = [
    new Paragraph({
      text: model.title,
      heading: HeadingLevel.TITLE,
      alignment: AlignmentType.CENTER,
      spacing: { after: 360 },
    }),
  ];
  for (const entry of model.metadata ?? []) {
    children.push(new Paragraph({
      children: [
        new TextRun({ text: `${entry.label}: `, bold: true, color: '34495E' }),
        new TextRun(entry.value),
      ],
      spacing: { after: 100, line: 300 },
    }));
  }
  if (model.metadata?.length) children.push(new Paragraph({ text: '', spacing: { after: 120 } }));
  for (const section of model.sections) {
    children.push(new Paragraph({
      text: section.heading,
      heading: HeadingLevel.HEADING_1,
      spacing: { before: 280, after: 140 },
      keepNext: true,
    }));
    for (const paragraph of section.paragraphs ?? []) {
      children.push(new Paragraph({ text: paragraph, spacing: { after: 160, line: 320 } }));
    }
    for (const bullet of section.bullets ?? []) {
      children.push(new Paragraph({ text: bullet, bullet: { level: 0 }, spacing: { after: 100, line: 300 } }));
    }
  }

  const document = new Document({
    title: model.title,
    styles: {
      default: {
        document: {
          run: { font: 'Aptos', size: 22, color: '1F2937' },
          paragraph: { spacing: { line: 320 } },
        },
        title: {
          run: { font: 'Aptos Display', size: 36, bold: true, color: '1F4E78' },
        },
        heading1: {
          run: { font: 'Aptos Display', size: 28, bold: true, color: '1F4E78' },
        },
      },
    },
    sections: [{
      properties: {
        page: { margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 } },
      },
      children,
    }],
  });
  const bytes = await rebuildDeterministicDocx(await Packer.toBuffer(document), model.title);
  const validation = await validateWordDocumentBytes(bytes);
  return { bytes, size: validation.bytes, sha256: validation.sha256 };
}

export async function readSecureWordDocumentModel(path: string): Promise<WordDocumentModel> {
  const bytes = await readSecureLocalFile(path, WORD_DOCUMENT_LIMITS.inputJsonBytes);
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    throw new Error('Input JSON file must contain valid UTF-8');
  }
  let value: unknown;
  try {
    value = JSON.parse(text) as unknown;
  } catch {
    throw new Error('Input JSON file must contain valid JSON');
  }
  return parseWordDocumentModel(value);
}

export async function createAndUploadWordDocument(
  config: AuthConfig,
  input: { model: unknown; path: string; driveId?: string },
): Promise<VerifiedDriveItem> {
  const document = await createWordDocument(input.model);
  return safeCreateDriveFile(config, {
    path: input.path,
    bytes: document.bytes,
    driveId: input.driveId,
    requiredExtension: '.docx',
  });
}
