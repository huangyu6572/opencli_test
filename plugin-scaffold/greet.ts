/**
 * Sample TypeScript command for opencli-plugin-devcloud.
 * Demonstrates the programmatic cli() registration API.
 */

import { cli, Strategy } from '@jackwener/opencli/registry';

cli({
  site: 'opencli-plugin-devcloud',
  name: 'greet',
  description: 'Greet someone by name',
  strategy: Strategy.PUBLIC,
  browser: false,
  args: [
    { name: 'name', positional: true, required: true, help: 'Name to greet' },
  ],
  columns: ['greeting'],
  func: async (kwargs) => [{ greeting: `Hello, ${String(kwargs.name ?? 'World')}!` }],
});
