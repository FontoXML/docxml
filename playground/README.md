# Docxml Playground

Minimal browser playground for `docxml`.

- Monaco editor with TypeScript.
- Runtime import uses the latest release from JSR via `https://esm.sh/jsr/@fontoxml/docxml?bundle`.
- Type information is loaded dynamically from `jsr.io` (latest package metadata + source files).
- One action: compile and generate a `.docx` file.

## Run locally with Deno

From the repository root:

```sh
deno task playground:dev
```

Open `http://localhost:5173`.

## Build static output with Deno

From the repository root:

```sh
deno task playground:build
```

The static output is generated in `playground/dist`.

## Deploy to GitHub Pages

`/.github/workflows/playground-pages.yml` publishes `playground/dist` to GitHub Pages when a new release is published.
