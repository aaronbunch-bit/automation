# Netlify deploy fix for `a-team-pace-dashboard`

This agent run is attached to **`aaronbunch-bit/automation`**, so it cannot push
directly to **`aaronbunch-bit/a-team-pace-dashboard`**. Apply these files there
(or re-run the cloud agent against that repo with write access).

## Root cause

1. `package.json` had `"@netlify/blobs": "^10.7.9"` with **no lockfile**.
2. npm resolved that to `10.7.11`, which depends on `@netlify/otel@^6.0.5`.
3. `@netlify/otel@6.0.5` was published the same day as the failed build and
   Netlify's install could not resolve it (`ETARGET`).
4. Separately, `netlify.toml` pointed at `netlify/functions`, but the function
   files live in `functions/` (the old `netlify/functions` directory was deleted).

## Fix (already prepared in this folder)

| File | Change |
|------|--------|
| `package.json` | Pin `@netlify/blobs` to `10.7.10`, add `@netlify/functions`, override `@netlify/otel` to `6.0.4` |
| `package-lock.json` | New lockfile so Netlify installs the known-good tree |
| `netlify.toml` | Point functions directory at `functions/` |
| `.gitignore` | Ignore `node_modules/` |

## Apply to `a-team-pace-dashboard`

From a clone of `a-team-pace-dashboard`:

```bash
# Option A — copy these files onto the repo root
cp package.json package-lock.json netlify.toml .gitignore /path/to/a-team-pace-dashboard/

# Option B — apply the patch
cd /path/to/a-team-pace-dashboard
git apply /path/to/0001-fix-netlify-otel-and-functions-path.patch

git add package.json package-lock.json netlify.toml .gitignore
git commit -m "Fix Netlify deploy: pin @netlify/otel and functions path"
git push
```

Then trigger a new Netlify deploy (push to the connected branch is enough).
