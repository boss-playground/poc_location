# Plum & Gold Calendar

Open the app: https://boss-playground.github.io/poc_location/calendar/

Modern purple/gold calendar with browser-only events and optional Outlook sync. This update removes the sidebar mini calendar, reduces bottom whitespace, and adds safe popup cancellation recovery. No bank affiliation or logo is used.

This directory contains the production build. Editable code, dependency lockfile, tests, and design assets are in `source/`. The existing root location-map page is unchanged.

## Develop and republish

From `calendar/source/`, run `npm ci`, then `npm run dev`. Run `npm test` and `npm run build` to verify changes. Refresh the parent production files with:

```sh
npm run build -- --base=/poc_location/calendar/ --outDir=../
```

Do not add `--emptyOutDir`: the parent also contains source files. Upload production and source changes together. GitHub Pages publishes from the `main` branch root.

## Outlook setup

Public Application (client) ID: `942aaf74-fa78-45d7-a764-a1a693ca4070`.

Register these **Single-page application** redirect URIs in Microsoft Entra:

- `https://boss-playground.github.io/poc_location/calendar/redirect.html`
- `https://boss-playground.github.io/poc_location/calendar/logout.html`

No client secret is included or needed. Real sign-in and calendar access require correct registration and your consent. Closed login popups release their own request state; old unowned locks require opening the calendar URL in a fresh tab. A blank popup stalled during discovery can wait up to about 20 seconds for bounded cleanup.

Browser-only events stay in localStorage for this origin; localhost events do not automatically transfer here. No personal events or authentication tokens are included in the repository.
