# bg-organic-registry

A fast, searchable mirror of the Bulgarian **organic-production operators registry**
(производители, преработватели и търговци на земеделски продукти и храни, произведени по
биологичен начин), published by the Ministry of Agriculture at
<https://bioreg.mzh.government.bg/Home/DataBaseList>.

The official site is a slow, server-paged jQuery DataTables UI with awkward search and no
usable export. This project provides instant client-side search, sorting, per-column filters,
and one-click **Excel/CSV export** — hosted entirely on **GitHub Pages**.

## How it works

The source API (`POST /Home/DataBaseListEffective` and `/Home/DataBaseListNonEffective`) requires
a session cookie and sends no CORS headers, so it cannot be called from a static site's browser.
Instead a Python script fetches the full dataset out-of-band and commits static artifacts that the
site loads directly:

- `public/data/effective.json` — active operators (~6,100)
- `public/data/noneffective.json` — inactive / expired (~11,400)
- `public/data/bioreg-*.xlsx` — prebuilt full-dataset Excel downloads
- `public/data/meta.json` — record counts + last-updated timestamp

A daily GitHub Action (`.github/workflows/deploy.yml`) refreshes the data, builds the Vite site,
and deploys to GitHub Pages.

### Why a Hetzner runner?

GitHub-hosted runners use US/Azure IP ranges that the `.government.bg` API may geo-block or
rate-limit. The workflow therefore provisions an **ephemeral self-hosted runner on Hetzner Cloud**
(Nuremberg, EU — via [`Cyclenerd/hcloud-github-runner`](https://github.com/Cyclenerd/hcloud-github-runner))
for the fetch + build step, then destroys it. The Pages deploy runs on a normal GitHub-hosted runner.

Two repository secrets are required:

- `HCLOUD_TOKEN` — a Hetzner Cloud API token (Read & Write).
- `PERSONAL_ACCESS_TOKEN` — a fine-grained GitHub PAT with **Administration: Read & write** on this
  repo (to register/unregister the self-hosted runner).

Enable Pages under **Settings → Pages → Build and deployment → GitHub Actions**. The site is served
at the default `https://<user>.github.io/<repo>/` URL (the build uses a relative base, so no custom
domain or path config is needed).

## Tech

- **Data:** Python (`requests`, `openpyxl`) — `scripts/fetch_data.py`
- **Frontend:** Vanilla JS + [Vite](https://vitejs.dev), [Tabulator](https://tabulator.info) table,
  [SheetJS](https://sheetjs.com) for in-browser export
- **Hosting:** GitHub Pages at the default `github.io` URL (relative base, no custom domain)
- **CI:** Ephemeral Hetzner Cloud self-hosted runner for the EU-side data fetch

## Develop locally

```bash
# 1. Fetch the data (needs Python 3 + requests + openpyxl)
pip install -r requirements.txt
python scripts/fetch_data.py        # writes public/data/*

# 2. Run the site
npm install
npm run dev                         # http://localhost:5173
npm run build && npm run preview    # production build
```

## Notes

This is an **unofficial** mirror for convenience. The authoritative source remains
<https://bioreg.mzh.government.bg>. Per-certificate detail pages are linked back to the official
site rather than copied.
