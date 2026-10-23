import { runCli, UsageError } from './cli.js';

// Entry point of the bundled skill script (skills/docentric/scripts/docentric.mjs).
try {
  process.exitCode = await runCli(process.argv.slice(2));
} catch (e) {
  console.error(e instanceof UsageError ? e.message : `Error: ${(e as Error).message}`);
  process.exitCode = 2;
}
