export interface CliArgs {
  _: string[];
  client_id?: string;
  authority?: string;
  scope?: string[];
  port?: string;
  no_persist?: boolean;
  json?: boolean;
  item_id?: string;
  path?: string;
  sheet?: string;
  address?: string;
  dir?: string;
  backup_dir?: string;
  values_json?: string;
  formulas_json?: string;
}

export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = { _: [] };
  for (let index = 0; index < argv.length; index += 1) {
    const token = argv[index];
    if (!token) continue;
    if (!token.startsWith('--')) {
      args._.push(token);
      continue;
    }
    const [rawKey, inlineValue] = token.slice(2).split('=', 2);
    const key = rawKey?.replaceAll('-', '_') as keyof CliArgs | undefined;
    if (!key) throw new Error(`Invalid option: ${token}`);
    if (key === 'no_persist' || key === 'json') {
      args[key] = true;
      continue;
    }
    const value = inlineValue ?? argv[index + 1];
    if (value === undefined || value.startsWith('--')) throw new Error(`Missing value for --${rawKey}`);
    if (inlineValue === undefined) index += 1;
    if (key === 'scope') args.scope = [...(args.scope ?? []), value];
    else if (key !== '_') args[key] = value as never;
  }
  return args;
}
