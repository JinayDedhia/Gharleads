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

(c) 2026 GharLeads. All rights reserved.