# Docxml Playground

Vite + Preact playground to try `docxml` in the browser.

The app bundles Monaco statically at build time (including workers), and still loads the latest published `docxml` release from JSR at runtime.
Type information for Monaco is loaded from the latest JSR source files in-memory and is not written to disk.

## Run locally with Deno

From the repository root:

```sh
deno task playground:dev
```

Then open `http://localhost:5173`.

## Build static output

From the repository root:

```sh
deno task playground:build
```

The generated static site is written to `playground/dist`.

## Deploy

The `.github/workflows/playground-pages.yml` workflow builds the playground and publishes `playground/dist` to GitHub Pages whenever a new release is published.
