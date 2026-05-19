/**
 * Sample pipeline command for opencli-plugin-devcloud.
 * Demonstrates the declarative pipeline API.
 */

import { cli, Strategy } from '@jackwener/opencli/registry';

cli({
  site: 'opencli-plugin-devcloud',
  name: 'hello',
  description: 'A sample pipeline command',
  strategy: Strategy.PUBLIC,
  browser: false,
  columns: ['greeting'],
  pipeline: [
    { fetch: { url: 'https://httpbin.org/get?greeting=hello' } },
    { select: 'args' },
  ],
});
