# Aurora Night design system

User approved the immersive night direction and requested a spectacular finish. This is a bounded visual restyle of the existing calendar, not a replacement app or new OAuth implementation.

## Reference and asset

- Concept: `aurora-night-concept.png`, 1536×1024, created with the built-in Image Gen tool.
- Production landscape: `../public/assets/aurora-night.png`, 1536×1024. Generated separately using the concept as a reference; no UI is embedded in the asset.
- Concept prompt: Full-screen personal month calendar “My calendar”; cinematic midnight aurora above mountains and a lake, two navy glass surfaces, mint primary action, serif month heading, Monday-first month grid, purple local events and cyan Outlook events; preserve existing information architecture and code-native controls, no marketing sections or fake metrics.
- Asset edit prompt: Extract/recreate only the complete aurora mountain/lake landscape from the concept, preserving composition and colors; remove all UI, text, panels, icons, numbers and borders; dark left/center negative space, aurora upper right, mountains in lower quarter.

## Tokens and components

- True midnight/navy backdrop #091426; translucent navy panels with subtle #a7d3df66 hairlines. No tint added to the landscape.
- Main text #eef5ff, captions #b4c4d9; mint CTA #7aefc7, lilac local #cc9afa, cyan Outlook #56e4f2.
- Georgia editorial brand/month/dialog headings; system sans-serif controls, field labels and event text.
- Desktop header 84px; workspace 36px top gap, roughly44px outer gutters, 20px column gap, 236–278px sidebar.
- Single calendar grid with thin cell divisions; dates aligned top-left, current day mint outline. No marketing hero or new navigation.
- All four dialogs inherit the more opaque dark editor treatment. All existing DOM ids, event operations, validation and storage/auth modules remain intact.
- Lightweight36-second background transform; pause/resume control and automatic system reduced-motion support. No WebGL, external font or additional runtime dependency.

## Intentional functional adaptations from concept

- Date/summary/status/account text and event content stay live rather than hardcoded to the concept's synthetic appointments. Source totals remain functional.
- Preserve existing Monday-first Mon–Sun labels and previous/next/Today controls. “Month view” is a label, not a fake dropdown; miniature calendar keeps its existing date-click behavior.
- Added a labeled pause/resume icon to meet continuous-motion accessibility needs; the concept omitted this necessary control.
- Mobile uses stacked controls and a readable scrollable month page; source checkboxes and connect/disconnect remain directly available. No forced fixed-height scroll trap.
- Microsoft configuration, permissions and live-login status are outside this restyle; no changes to auth, Graph or stored event records.
