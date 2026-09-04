# GSO / GCTS Tools — web app

A small Streamlit site with three tabs, one per script:

1. **GCTS Report generator** — upload the Excel workbook, download a formatted
   Word report. Fast, no network calls, works perfectly hosted.
2. **GSO Model Extractor** — upload a workbook of product URLs, get back
   models + spec fields scraped from the public GSO site. Uses a real
   headless browser per row, so it's slow — good for testing a batch of
   rows through the browser, not for a full run of thousands.
3. **GSO Product-ID Lookup** — needs a session file you generate once on
   your own computer (the interactive login step can't happen inside a
   hosted website). Once you upload that session file, lookups run as
   plain fast HTTP requests.

## Why tools 2 & 3 have limits online

Free hosts (and honestly most paid ones too) kill requests or whole
containers after a fixed amount of time — usually seconds to a couple of
minutes for the free tiers below. Script 2 can take hours for a big sheet,
and script 3's login step requires a real interactive browser window,
which a server has no way to show you. So:

- Use the site for **testing / small batches** of tools 2 and 3.
- Do the **full multi-thousand-row runs locally** on your own machine,
  exactly as you were already doing — nothing about that changes.

## Files

```
app.py                     Streamlit front end (the website)
gcts_report.py              tool 1, used as-is
gso_model_extractor.py      tool 2, lightly patched to accept env-var overrides
gso_product_id_lookup.py    tool 3, lightly patched to accept env-var overrides
requirements.txt            Python deps
packages.txt                system libs Playwright's Chromium needs
```

## Hosting it for free — Streamlit Community Cloud

This is the easiest free option that supports Playwright out of the box
via `packages.txt`, and it doesn't sleep after every request the way some
serverless free tiers do.

1. Create a free GitHub account if you don't have one, and a new
   **public** repo (e.g. `gso-tools`).
2. Upload all the files in this folder to that repo (drag-and-drop on
   github.com works, or `git push` if you're comfortable with git).
3. Go to **share.streamlit.io**, sign in with GitHub, click **New app**.
4. Pick your repo, branch `main`, main file `app.py`. Click **Deploy**.
5. First load will take a minute or two while it installs Chromium
   (that's the `ensure_playwright_browser()` step in `app.py`) — after
   that it's cached for the life of the container.

You get a permanent URL like `yourname-gso-tools.streamlit.app`, free,
with no request limit (Community Cloud's only real constraints are that
the app can go to sleep after a period of *zero* visits — anyone opening
the link wakes it back up in a few seconds — and it runs on modest shared
hardware, so keep batch sizes in tool 2/3 reasonable).

### Alternative: Render.com free web service

Also free, also fine for tool 1. For tools 2/3 add `packages.txt`'s
contents to a Dockerfile's `apt-get install` instead (Render's native
Python runtime doesn't read `packages.txt` — that's a Streamlit Cloud
convention). Render's free tier spins the service down after 15 minutes
of inactivity and takes ~30–60s to wake back up, otherwise behaves the
same as Streamlit Cloud for this purpose.

## Running it locally first (recommended before deploying)

```bash
pip install -r requirements.txt
python -m playwright install chromium
streamlit run app.py
```

Then open the local URL it prints, and try each tab once so you know
what to expect before putting it online.
