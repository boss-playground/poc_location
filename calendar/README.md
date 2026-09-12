# Aurora Night Calendar

Open the app: https://boss-playground.github.io/poc_location/calendar/

This directory contains the production build. The editable app, dependency lockfile, tests, and design assets are in `source/`. The repository's existing root location-map page is unchanged.

## Develop and republish

From `calendar/source/`, run `npm ci`, then `npm run dev`. Run `npm test` and `npm run build` to verify changes. To refresh the production files in the parent directory, run:

```sh
npm run build -- --base=/poc_location/calendar/ --outDir=../
```

Do not add `--emptyOutDir`: the parent also contains source files. Upload the changed production files alongside the source changes. GitHub Pages publishes from the existing `main` branch root.

## Outlook setup

The default public Application (client) ID is `942aaf74-fa78-45d7-a764-a1a693ca4070`. Register these **Single-page application** redirect URIs in Microsoft Entra:

- `https://boss-playground.github.io/poc_location/calendar/redirect.html`
- `https://boss-playground.github.io/poc_location/calendar/logout.html`

No client secret is needed or included. Actual Microsoft sign-in and calendar access require your account's consent and correct registration.

Browser-only events are saved in localStorage for this website origin. Events saved on localhost do not automatically transfer to the published site. No personal calendar data or authentication tokens are included in this repository.
