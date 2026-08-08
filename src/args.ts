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
]);

export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = { _: [] };
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
    else if (key !== '_') args[key] = value as never;
  }
  return args;
}
