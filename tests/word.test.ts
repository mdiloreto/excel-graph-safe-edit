import { chmod, mkdtemp, rm, symlink, writeFile } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { setTimeout as delay } from 'node:timers/promises';
import { inflateRawSync } from 'node:zlib';
import { afterEach, describe, expect, it } from 'vitest';
import {
  createWordDocument,
  parseWordDocumentModel,
  readSecureWordDocumentModel,
  validateWordDocumentBytes,
  WORD_DOCUMENT_LIMITS,
  type WordDocumentModel,
} from '../src/word.js';

const temporaryDirectories: string[] = [];

function readZipEntry(zip: Buffer, expectedName: string): Buffer {
  let cursor = zip.indexOf(Buffer.from([0x50, 0x4b, 0x01, 0x02]));
  while (cursor !== -1 && zip.readUInt32LE(cursor) === 0x0201_4b50) {
    const compressionMethod = zip.readUInt16LE(cursor + 10);
    const compressedSize = zip.readUInt32LE(cursor + 20);
    const nameLength = zip.readUInt16LE(cursor + 28);
    const extraLength = zip.readUInt16LE(cursor + 30);
    const commentLength = zip.readUInt16LE(cursor + 32);
    const localOffset = zip.readUInt32LE(cursor + 42);
    const name = zip.subarray(cursor + 46, cursor + 46 + nameLength).toString('utf8');
    if (name === expectedName) {
      const localNameLength = zip.readUInt16LE(localOffset + 26);
      const localExtraLength = zip.readUInt16LE(localOffset + 28);
      const dataStart = localOffset + 30 + localNameLength + localExtraLength;
      const compressed = zip.subarray(dataStart, dataStart + compressedSize);
      if (compressionMethod === 0) return compressed;
      if (compressionMethod === 8) return inflateRawSync(compressed);
      throw new Error(`Unsupported ZIP compression method ${compressionMethod}`);
    }
    cursor += 46 + nameLength + extraLength + commentLength;
  }
  throw new Error(`ZIP entry not found: ${expectedName}`);
}

const validModel: WordDocumentModel = {
  schemaVersion: 1,
  title: 'Class transcript',
  metadata: [
    { label: 'Date', value: '2026-09-04' },
    { label: 'Instructor', value: 'Ada Lovelace' },
  ],
  sections: [
    { heading: 'Overview', paragraphs: ['The class opened with a review.'] },
    { heading: 'Key points', paragraphs: ['The discussion continued.'], bullets: ['First point', 'Second point'] },
  ],
};

afterEach(async () => {
  await Promise.all(temporaryDirectories.splice(0).map((path) => rm(path, { recursive: true, force: true })));
});

describe('Word document model validation', () => {
  it('accepts a bounded generic transcript model and returns a defensive copy', () => {
    const parsed = parseWordDocumentModel(validModel);
    expect(parsed).toEqual(validModel);
    expect(parsed).not.toBe(validModel);
    expect(parsed.sections).not.toBe(validModel.sections);
  });

  it('rejects unknown keys at every modeled level', () => {
    expect(() => parseWordDocumentModel({ ...validModel, courseId: 'moodle-1' })).toThrow(/unexpected key/);
    expect(() => parseWordDocumentModel({
      ...validModel,
      metadata: [{ label: 'Date', value: 'today', extra: true }],
    })).toThrow(/unexpected key/);
    expect(() => parseWordDocumentModel({
      ...validModel,
      sections: [{ heading: 'Section', paragraphs: ['Text'], level: 1 }],
    })).toThrow(/unexpected key/);
  });

  it('rejects malformed versions, empty content, unsafe XML, and sections without blocks', () => {
    expect(() => parseWordDocumentModel({ ...validModel, schemaVersion: 2 })).toThrow(/schemaVersion/);
    expect(() => parseWordDocumentModel({ ...validModel, title: '   ' })).toThrow(/title/);
    expect(() => parseWordDocumentModel({ ...validModel, title: 'bad\u0000title' })).toThrow(/XML/);
    expect(() => parseWordDocumentModel({ ...validModel, sections: [] })).toThrow(/sections/);
    expect(() => parseWordDocumentModel({ ...validModel, sections: [{ heading: 'Empty' }] })).toThrow(/paragraphs and\/or bullets/);
    expect(() => parseWordDocumentModel({ ...validModel, sections: [{ heading: 'Empty', paragraphs: [] }] })).toThrow(/1 through/);
    const sparseSections = new Array<unknown>(1);
    expect(() => parseWordDocumentModel({ ...validModel, sections: sparseSections })).toThrow(/missing entries/);
  });

  it('enforces individual, count, total-block, and total-character bounds', () => {
    expect(() => parseWordDocumentModel({
      ...validModel,
      title: 'x'.repeat(WORD_DOCUMENT_LIMITS.titleCharacters + 1),
    })).toThrow(/title/);
    expect(() => parseWordDocumentModel({
      ...validModel,
      metadata: Array.from({ length: WORD_DOCUMENT_LIMITS.metadataEntries + 1 }, () => ({ label: 'x', value: 'y' })),
    })).toThrow(/metadata/);
    expect(() => parseWordDocumentModel({
      ...validModel,
      sections: [{
        heading: 'Too many blocks',
        paragraphs: Array.from({ length: WORD_DOCUMENT_LIMITS.blocksPerSection }, () => 'p'),
        bullets: ['extra'],
      }],
    })).toThrow(/total blocks/);
    expect(() => parseWordDocumentModel({
      schemaVersion: 1,
      title: 'Large transcript',
      sections: [{
        heading: 'Long section',
        paragraphs: Array.from({ length: 51 }, () => 'x'.repeat(WORD_DOCUMENT_LIMITS.blockCharacters)),
      }],
    })).toThrow(/total characters/);
  });

  it('supports a two-hour transcript aggregated into at most 100 timestamped paragraphs', () => {
    const paragraphs = Array.from({ length: WORD_DOCUMENT_LIMITS.blocksPerSection }, (_, index) => {
      const startSeconds = index * 72;
      const endSeconds = startSeconds + 71;
      const timestamp = (seconds: number) => `${String(Math.floor(seconds / 3600)).padStart(2, '0')}:${String(Math.floor((seconds % 3600) / 60)).padStart(2, '0')}:${String(seconds % 60).padStart(2, '0')}`;
      return `[${timestamp(startSeconds)}–${timestamp(endSeconds)}] ${'spoken transcript text '.repeat(100).trim()}`;
    });
    expect(parseWordDocumentModel({
      schemaVersion: 1,
      title: 'Two-hour class transcript',
      sections: [{ heading: 'Timestamped transcript', paragraphs }],
    }).sections[0]?.paragraphs).toHaveLength(100);
    expect(() => parseWordDocumentModel({
      schemaVersion: 1,
      title: 'Two-hour class transcript',
      sections: [{ heading: 'Timestamped transcript', paragraphs: [...paragraphs, '[02:00:00] Extra cue'] }],
    })).toThrow(/1 through 100 strings/);
  });
});

describe('DOCX generation and local input safety', () => {
  it('generates a bounded DOCX with required OOXML entries and a local SHA-256', async () => {
    const generated = await createWordDocument(validModel);
    expect(generated.bytes.length).toBe(generated.size);
    expect(generated.size).toBeGreaterThan(0);
    expect(generated.size).toBeLessThanOrEqual(WORD_DOCUMENT_LIMITS.docxBytes);
    expect(generated.sha256).toMatch(/^[a-f0-9]{64}$/u);
    await expect(validateWordDocumentBytes(generated.bytes)).resolves.toEqual({
      bytes: generated.size,
      sha256: generated.sha256,
    });
    const documentXml = readZipEntry(generated.bytes, 'word/document.xml').toString('utf8');
    const stylesXml = readZipEntry(generated.bytes, 'word/styles.xml').toString('utf8');
    const coreXml = readZipEntry(generated.bytes, 'docProps/core.xml').toString('utf8');
    for (const text of ['Class transcript', 'Date', '2026-09-04', 'Overview', 'The class opened with a review.', 'First point']) {
      expect(documentXml).toContain(text);
    }
    expect(documentXml).toContain('w:pStyle w:val="Title"');
    expect(documentXml).toContain('w:pStyle w:val="Heading1"');
    expect(documentXml).toContain('<w:numPr>');
    expect(stylesXml).toContain('Aptos');
    expect(coreXml).toContain('<dc:title>Class transcript</dc:title>');
    expect(coreXml.match(/2000-01-01T00:00:00\.000Z/gu)).toHaveLength(2);
  });

  it('produces identical bytes for the same canonical model after wall-clock time changes', async () => {
    const first = await createWordDocument(validModel);
    await delay(2100);
    const second = await createWordDocument(validModel);
    const changed = await createWordDocument({
      ...validModel,
      title: 'Changed & <safe>',
      sections: [{ heading: 'Overview', paragraphs: ['Changed content.'] }],
    });

    expect(second.bytes).toEqual(first.bytes);
    expect(second.sha256).toBe(first.sha256);
    expect(second.size).toBe(first.size);
    expect(changed.sha256).not.toBe(first.sha256);
    expect(changed.bytes.equals(first.bytes)).toBe(false);
    expect(readZipEntry(second.bytes, 'word/document.xml')).toEqual(readZipEntry(first.bytes, 'word/document.xml'));
    expect(readZipEntry(changed.bytes, 'docProps/core.xml').toString('utf8')).toContain(
      '<dc:title>Changed &amp; &lt;safe&gt;</dc:title>',
    );
  });

  it('rejects a DOCX package missing a required relationship entry', async () => {
    const generated = await createWordDocument(validModel);
    const corrupted = Buffer.from(generated.bytes);
    const original = Buffer.from('_rels/.rels');
    const replacement = Buffer.from('_rels/.relx');
    let offset = corrupted.indexOf(original);
    while (offset !== -1) {
      replacement.copy(corrupted, offset);
      offset = corrupted.indexOf(original, offset + replacement.length);
    }
    await expect(validateWordDocumentBytes(corrupted)).rejects.toThrow(/_rels\/\.rels/);
  });

  it('reads only private, regular local JSON files and validates their contents', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'word-model-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'model.json');
    await writeFile(path, JSON.stringify(validModel), { mode: 0o600 });
    await chmod(path, 0o600);
    await expect(readSecureWordDocumentModel(path)).resolves.toEqual(validModel);

    await chmod(path, 0o644);
    await expect(readSecureWordDocumentModel(path)).rejects.toThrow(/group or other/);
    await chmod(path, 0o600);
    const link = join(directory, 'model-link.json');
    await symlink(path, link);
    await expect(readSecureWordDocumentModel(link)).rejects.toThrow();
  });

  it('rejects malformed UTF-8 bytes before parsing input JSON', async () => {
    const directory = await mkdtemp(join(tmpdir(), 'word-model-'));
    temporaryDirectories.push(directory);
    const path = join(directory, 'malformed.json');
    await writeFile(path, Buffer.from([0x7b, 0x22, 0x78, 0x22, 0x3a, 0x22, 0xc3, 0x28, 0x22, 0x7d]), { mode: 0o600 });

    await expect(readSecureWordDocumentModel(path)).rejects.toThrow(/valid UTF-8/);
  });
});
