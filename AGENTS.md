# AGENTS.md

## Cursor Cloud specific instructions

This is a static Jekyll-based GitHub Pages site that serves JSON files as API endpoints. There is no application code, no automated tests, and no linter configuration in the repository.

### Project overview

- **Technology**: Jekyll (Ruby) with the `jekyll-theme-slate` theme
- **Purpose**: Serves static JSON files (groceries, nutrition, swim results, content/articles) as public API endpoints
- **Production URL pattern**: `https://satsin06.github.io/customApis/<filename>.json`

### Development server

```bash
bundle exec jekyll serve --host 0.0.0.0 --port 4000
```

JSON endpoints are then available at `http://localhost:4000/<filename>.json`.

### Build

```bash
bundle exec jekyll build
```

Output goes to `_site/`.

### Testing / Linting

There are no automated tests or linting tools configured. JSON validity can be checked with:

```bash
python3 -c "import json; json.load(open('<file>.json'))"
```

### Gotchas

- The Gemfile and `.gitignore` were added for local development; the original repo did not have them. The update script creates the Gemfile if missing.
- `vendor/bundle/`, `_site/`, and `.bundle/` are gitignored and should not be committed.
- Sass `@import` deprecation warnings from the theme are expected and harmless.
- `bundle install` must use `--path vendor/bundle` (or `bundle config set --local path vendor/bundle`) since the system gem directory may not be writable.
