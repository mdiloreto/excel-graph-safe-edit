export { buildAuthConfig, type AuthConfig } from './config.js';
export {
  driveChildrenPath,
  driveItemPathById,
  driveItemPathByPath,
  driveRootPath,
  driveUploadSessionPath,
  DriveItemConflictError,
  encodeGraphPathSegment,
  encodeOneDrivePath,
  GRAPH_UPLOAD_CHUNK_BYTES,
  GRAPH_UPLOAD_GRANULARITY,
  MAX_SAFE_UPLOAD_BYTES,
  parseDriveFilePath,
  safeCreateDriveFile,
  type DrivePathTarget,
  type DriveUploadDependencies,
  type VerifiedDriveItem,
} from './drive.js';
export { type OoxmlValidation } from './ooxml.js';
export {
  createAndUploadLocalFile,
  readSecureLocalFile,
  type LocalFileUploadDependencies,
  type SecureLocalReadDependencies,
} from './local-file.js';
export {
  createAndUploadWordDocument,
  createWordDocument,
  parseWordDocumentModel,
  readSecureWordDocumentModel,
  validateWordDocumentBytes,
  WORD_DOCUMENT_LIMITS,
  type GeneratedWordDocument,
  type WordDocumentModel,
  type WordDocumentSection,
  type WordMetadataEntry,
} from './word.js';
