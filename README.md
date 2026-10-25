# GharLeads

Windows desktop app that finds real estate agents, brokers and channel partners across the Mumbai region, station by station, and exports them to Excel.

## Features
- Pick stations on the Western, Central, Harbour, Trans-Harbour and Uran lines
- Categories: agents, brokers, channel partners, regional and institutional channel partners
- Sources: Google Maps and OpenStreetMap
- Phone, address, WhatsApp link, listing link, optional email and website
- Deep mode, relevance filter, "only new leads" memory, saved presets
- Excel export with one sheet per station

## Run from source
1. Install Python 3.12 (tick "Add python.exe to PATH")
2. Double-click `run_dev.bat`

## Build the Windows program and installer
1. Double-click `build.bat` (creates `dist\GharLeads\GharLeads.exe`)
2. Double-click `test_exe.bat` to check the bundled browser
3. Open `installer.iss` in Inno Setup and press Ctrl+F9, or run `make_installer.bat`

## Notes
- Leads are collected from publicly listed business information. Users are responsible for lawful use (India's DPDP Act, TRAI rules).
- Licensing and subscriptions are not built yet.

## 💡 What I Learned

### Scraping technique
- **`requests` can't handle JS-rendered sites.** Sulekha and MahaRERA returned nothing in v5 because the content only appears after JavaScript runs. Playwright fixed this in v6.
- **Stealth patches help, but they have limits.** Patching `navigator.webdriver`, plugins, languages and `window.chrome` got past basic bot checks. It did not beat Cloudflare Turnstile (99acres) or a login wall (JustDial).
- **Try the API first, then fall back to the browser.** This is faster and more stable (MahaRERA approach). The browser is the backup, not the default.
- **Warm up the session.** Visiting the homepage first to pick up cookies makes later requests look more human.
- **Class names change often.** Using multiple regex selectors with fallbacks (as in the 99acres and Sulekha parsers) keeps scrapers from breaking on small layout changes.
- **Try several ways to get a phone number.** `data-phone` attributes, `tel:` links, then visible text. No single method works on every site.

### Speed and efficiency
- **Dedupe URLs before visiting them.** Collecting all place links first and visiting each only once on Google Maps saved a lot of time and avoided duplicate records.
- **Parallelism has a ceiling.** Visible browser windows are heavy, so 2 parallel stations is the practical limit on an average PC.
- **Run the most reliable source first.** Google Maps runs first and MahaRERA runs once at the end, so a failure in a flaky source doesn't block the good data.

### Data quality
- **Dedupe by "richest record", not first seen.** Scoring each record by phone + email + address and keeping the best one gives a much more useful final list.
- **Validate phone numbers.** Strip `+91` and leading `0`, and accept only 10-digit numbers starting with 6-9. This removes landlines and junk.
- **Harvest emails from business websites.** `mailto:` links are the most reliable source; a regex scan of the page is the fallback.
- **Match MahaRERA records to stations by address.** Anything that doesn't match goes into a General (NM) bucket instead of being lost.

### Environment and process
- **Use home Wi-Fi.** VPNs and mobile hotspots are IP-blocked by these sites.
- **Keep settings at the top of the script.** `SHOW_BROWSER`, `MAX_PER_QUERY`, `PARALLEL_WORKERS` and `SCROLL_ROUNDS` can be tuned without touching the logic.
- **Zero results is a bug, not a result.** Silent failures (JustDial, 99acres) went unnoticed until the output was checked. Each source needs clear logging of its count.
- **Fix one source at a time and keep version history.** Tracking v1 to v6 made it easy to see what each version actually fixed.

### Bigger lessons
- **A login wall means the data is gated.** Don't fight it. Look for a public API or an alternative source (Housing.com, IndiaProperty) instead.
- **No single source is complete.** Combining Maps, directories and the government registry (MahaRERA) gives far better coverage than any one alone.
- **Scraping is maintenance, not a one-time build.** Sites change and add protection, so the scraper needs regular updates.

(c) 2026 GharLeads. All rights reserved.
