CRUSH handbook for this repo

Build/lint/test
- Install: npm ci (Node >= 20.15)
- Build: npm run build (rimraf dist + tsc + gulp build:icons)
- Dev compile: npm run dev (tsc --watch)
- Lint: npm run lint (eslint nodes credentials package.json)
- Lint fix: npm run lintfix
- Format: npm run format (prettier nodes credentials --write)
- Prepublish check: npm run prepublishOnly (build + lint with prepublish config)
- Tests: No test framework configured; add one if needed. To run a single test later, prefer: npm test -- <pattern> (e.g., vitest/jest/mocha) once added.

Code style
- Language: TypeScript (strict, es2019 target, commonjs). JSON files under nodes are included by tsc.
- Imports: use ES import syntax; esModuleInterop enabled. Resolve JSON via resolveJsonModule. Prefer relative paths within nodes/ and credentials/.
- Formatting: Prettier config enforced: tabs, tabWidth 2, semi true, singleQuote true, trailingComma all, arrowParens always, printWidth 100, LF EOL, bracketSpacing true, quoteProps as-needed.
- Linting: ESLint with plugin:n8n-nodes-base rules. TS files only; JS files and dist are ignored. Use .eslintrc.prepublish.js to enforce package.json rules for publish.
- Naming: Follow n8n conventions:
  - Credentials classes end with Api and filename ends with .credentials.ts; names suffixed as required by rules.
  - Node classes/descriptions follow n8n node conventions (displayName casing, suffixes, directory/filename conventions). See rules in .eslintrc.js.
- Types: strict true, noImplicitAny, strictNullChecks, noUnusedLocals, noImplicitReturns. Prefer explicit types and const assertions where applicable. useUnknownInCatchVariables false (can type catch variable as any), but prefer narrowing.
- Error handling: For node execute blocks, throw NodeOperationError/NodeApiError as required by n8n rules; avoid generic errors. Ensure correct error messages and do not leak secrets.
- Files/structure: Keep code under nodes/ and credentials/. Icons (.png/.svg) are copied to dist by gulp task. Do not commit dist/.
- Peer deps: Uses n8n-workflow; ensure versions are compatible.

Tips for agents
- Before building, ensure devDependencies installed; do not change engines unless necessary.
- When adding tests, record the chosen runner in this file and add scripts (e.g., "test", "test:watch", and single-test invocation guidance).
- Keep secrets out of code and logs; never commit real credentials.
