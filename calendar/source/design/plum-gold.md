# Plum & Gold

Approved direction: modern SCB-inspired purple and restrained gold, no bank logo or affiliation; remove the sidebar mini calendar and reduce bottom whitespace. Existing functional calendar controls and copy stay native HTML.

## Visual specification

- Concept: `plum-gold-concept.png`, generated at 1536 × 1024.
- Background: `../public/assets/plum-gold.png`, generated as a standalone edit extracting the concept's violet glass/satin ribbons, dark plum center and sparse gold highlights. No text, UI, mountains or green aurora.
- Palette: deep plum `#20102f`, light violet text, lilac local events, warm gold `#e8c46c` actions and Outlook events.
- Type: system sans-serif; desktop month heading up to 42px, body and controls retain legible existing sizing.
- Layout: 232px natural-height sidebar; 16px gap; flexible month panel; 24px bottom gutter. Mobile stacks sidebar and month panel and retains day-detail access for truncated events.
- Icons: retain the existing 1.6px outline calendar, plus, chevrons, sync, external-link, close and motion controls. Do not add decorative dropdown affordances with no behavior.

## Fidelity ledger

Compared the concept with the current 1536 × 1024 browser render, mobile 390 × 844 and the event editor.

1. Background: violet ribbons and gold edge lighting match the concept's materials; standalone generated artwork replaces the former aurora landscape.
2. Palette: plum glass panels, gold primary action/Outlook marks and lilac local marks match; editor controls and focus states use the same palette.
3. Typography: serif month heading replaced with modern sans-serif at the concept's approximate scale; real labels remain readable rather than rasterized.
4. Sidebar: duplicate calendar removed. Intentional deviation: natural-height sidebar instead of stretching a mostly empty panel to the floor, addressing the user's whitespace concern.
5. Composition: main month grid expands to the desktop viewport with a 24px outer bottom gutter. Intentional deviation: retain month-before-navigation order to preserve the existing interaction hierarchy.
6. Functional fidelity: omit the concept's invented account/month dropdown arrows; Month view is an honest static label. Connected account state and real event titles take precedence over illustrative concept copy.
7. Responsive continuation: stacked compact controls, no horizontal document overflow; day detail/editor remain usable. Existing mobile month cells truncate long titles with full event information available on interaction.

These deviations preserve the approved visual direction and real functionality. No extra decorative feature or navigation was introduced.
