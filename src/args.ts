export interface CliArgs {
  _: string[];
  client_id?: string;
  authority?: string;
  scope?: string[];
  port?: string;
  no_persist?: boolean;
  json?: boolean;
  help?: boolean;
  item_id?: string;
  drive_id?: string;
  path?: string;
  sheet?: string;
  address?: string;
  dir?: string;
  backup_dir?: string;
  values_json?: string;
  formulas_json?: string;
  input_json?: string;
  input_file?: string;
  cache_key_file?: string;
  cache_path?: string;
}

const BOOLEAN_OPTIONS = new Set<keyof CliArgs>(['no_persist', 'json', 'help']);
const VALUE_OPTIONS = new Set<keyof CliArgs>([
  'client_id',
  'authority',
  'scope',
  'port',
  'item_id',
  'drive_id',
  'path',
  'sheet',
  'address',
  'dir',
  'backup_dir',
  'values_json',
  'formulas_json',
  'input_json',
  'input_file',
  'cache_key_file',
  'cache_path',
]);

export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = { _: [] };
  const occurrences = new Map<keyof CliArgs, number>();
  let positionalOnly = false;
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token) continue;
    if (positionalOnly) {
      args._.push(token);
      continue;
    }
    if (token === '--') {
      positionalOnly = true;
      continue;
    }
    if (token === '-h') {
      occurrences.set('help', (occurrences.get('help') ?? 0) + 1);
      args.help = true;
      continue;
    }
    if (!token.startsWith('--')) {
      args._.push(token);
      continue;
    }
    const option = token.slice(2);
    const equalsIndex = option.indexOf('=');
    const rawKey = equalsIndex === -1 ? option : option.slice(0, equalsIndex);
    const inlineValue = equalsIndex === -1 ? undefined : option.slice(equalsIndex + 1);
    const key = rawKey?.replaceAll('-', '_') as keyof CliArgs | undefined;
    if (!key) throw new Error(`Invalid option: ${token}`);
    if (BOOLEAN_OPTIONS.has(key)) {
      if (inlineValue !== undefined) throw new Error(`Option --${rawKey} does not accept a value`);
      occurrences.set(key, (occurrences.get(key) ?? 0) + 1);
      if (key === 'no_persist') args.no_persist = true;
      else if (key === 'json') args.json = true;
      else if (key === 'help') args.help = true;
      continue;
    }
    if (!VALUE_OPTIONS.has(key)) throw new Error(`Unknown option: --${rawKey}`);
    const value = inlineValue ?? argv[index + 1];
    if (value === undefined || value.startsWith('--')) throw new Error(`Missing value for --${rawKey}`);
    if (value.trim().length === 0) throw new Error(`Empty value for --${rawKey}`);
    if (inlineValue === undefined) index += 1;
    if (key === 'scope') args.scope = [...(args.scope ?? []), value];
    else if (key !== '_') {
      occurrences.set(key, (occurrences.get(key) ?? 0) + 1);
      args[key] = value as never;
    }
  }
  const uniqueOptions = args._[0] === 'upload-docx'
    ? ['client_id', 'authority', 'port', 'no_persist', 'json', 'help', 'input_json', 'path', 'drive_id', 'cache_key_file', 'cache_path'] as const
    : args._[0] === 'upload-file'
      ? ['client_id', 'authority', 'port', 'no_persist', 'json', 'help', 'input_file', 'path', 'drive_id', 'cache_key_file', 'cache_path'] as const
      : [];
  for (const key of uniqueOptions) {
      if ((occurrences.get(key) ?? 0) > 1) throw new Error(`Option --${key.replaceAll('_', '-')} may be specified only once`);
  }
  return args;
}
