# The Earth, Written by Us - Design Handoff

## Design approach

This local design package treats the project as an interactive storybook publication. The Earth is the illustrated protagonist on the left/top, while the storybook panel is the readable publishing surface where the reader enters ecological words.

Chosen direction: storybook plus illustrated planet. The interface keeps paper, margin notes, field-note texture, visible damage, repair marks, and warm ecological color without becoming a generic green sustainability site.

## Delivered files

- `index.html`: static design preview with all six required screens.
- `styles.css`: design tokens, responsive layout, component states, and screen styling.
- `assets/earth/earth-fragile.svg`
- `assets/earth/earth-warming.svg`
- `assets/earth/earth-repairing.svg`
- `assets/earth/earth-blooming.svg`
- `assets/earth/earth-collective.svg`
- `assets/earth/earth-tipping.svg`
- `assets/earth/earth-collapsing.svg`
- `assets/earth/earth-fragile.png`
- `assets/earth/earth-warming.png`
- `assets/earth/earth-repairing.png`
- `assets/earth/earth-blooming.png`
- `assets/earth/earth-collective.png`
- `assets/earth/earth-tipping.png`
- `assets/earth/earth-collapsing.png`
- `assets/effects/effect-crack-overlay.svg`
- `assets/effects/effect-heat-haze.svg`
- `assets/effects/effect-repair-glow.svg`
- `assets/effects/effect-sprout.svg`
- `assets/effects/effect-water-ripple.svg`
- `assets/effects/effect-network-lines.svg`
- `assets/effects/effect-tipping-halo.svg`
- `assets/effects/texture-paper-grain.svg`
- `assets/effects/icon-repair.svg`
- `assets/effects/icon-temperature.svg`
- `assets/effects/icon-community.svg`

## Screens covered

- DES-SCREEN-01: Landing / Prologue
- DES-SCREEN-02: Main Story Interface
- DES-SCREEN-03: Earth State Change
- DES-SCREEN-04: Branching Choice Moment
- DES-SCREEN-05: Ending Card
- DES-SCREEN-06: Archive / Story Map

## Visual system

Primary tokens are defined in `:root` inside `styles.css`.

- `--paper-bg`: warm off-white page background.
- `--paper-panel`: readable storybook page surface.
- `--ink`: primary charcoal text.
- `--ink-muted`: secondary body and annotation text.
- `--moss`: repair, care, and positive action.
- `--sage`: soft ecological secondary color.
- `--clay`: warning, Earth material, and margin labels.
- `--ember`: collapse and negative action accents.
- `--sky-faint`: fragile atmosphere.
- `--gold-soft`: tipping glow and publication highlight.

Typography:

- Titles and chapter names use Georgia as an editorial storybook face.
- Interface labels, buttons, inputs, and navigation use the system sans-serif stack.

Shape and texture:

- Cards and panels use an 8px radius.
- Storybook panels use warm paper, side spine detail, and subtle grain/grid texture.
- Earth panels use atmospheric glows and field-note annotation blocks.

## Component notes

Storybook panel:

- Normal chapter state is shown in the main interface.
- Response state is represented by the blockquote Earth voice.
- Branch state is shown in the branching choice screen.
- Ending state is shown as a printable field-note card.

Keyword input:

- Empty/focus-ready input uses placeholder `Type a small action...`.
- Submit control is compact and non-technical.
- Invalid/negative actions should use the ember/dashed treatment shown on warning chips.

Action chips:

- Positive chips: muted moss background and moss text.
- Negative chips: ember text, pale ember background, dashed border.
- Negative actions should remain visible but less inviting than repair actions.

Progress:

- Current design uses `Chapter 2 / 5` plus five chapter dots.
- A future implementation can map current chapter to `.is-current` and previous chapters to `.is-complete`.

Archive:

- The archive is intentionally a publication ledger and keyword map, not a data dashboard.
- Distribution bars are included only as a quiet collective record.

## Earth state direction

- Fragile: pale blue-grey, weak glow, hairline cracks.
- Warming: orange haze, dry regions, red stress marks.
- Repairing: stitched cracks, green repair lines, incomplete healing.
- Blooming: moss green, river marks, sprouts, living surface.
- Collective: connected lights and care network lines.
- Tipping: strong warm halo, transformed cracks, irreversible positive shift.
- Collapsing: charcoal surface, broken fragments, ember cracks, polluted atmosphere.

The PNG set is generated from a single image-model reference sheet and split into 512x512 state assets. The SVG set remains available as scalable fallback artwork.

## Responsive behavior

- Desktop: main interface uses Earth left and storybook right.
- Tablet/mobile: layouts stack with Earth first and storybook below.
- Mobile navigation hides the section links to avoid cramped header text.
- Inputs and choice buttons become full-width so labels do not overflow.
- Earth cards collapse from four columns to two, then one column on small screens.

## Transition notes for development

- On keyword submit, fade/scale the current Earth asset out over 180ms, then crossfade the next state in over 260ms.
- Positive transitions can add `effect-repair-glow.svg`, `effect-sprout.svg`, or `effect-tipping-halo.svg`.
- Warming/collapse transitions can briefly overlay `effect-heat-haze.svg` or `effect-crack-overlay.svg`.
- Status label should update at the same time as the Earth asset.
- Story text should appear after the Earth transition starts, not before, so the visual change feels causal.

## Acceptance checklist

- [x] Main interface clearly shows Earth left and storybook right.
- [x] At least 5 Earth visual states are designed.
- [x] Storybook panel is readable and visually distinctive.
- [x] Keyword input and action chips are designed with clear states.
- [x] Ending card design exists.
- [x] Mobile layout exists through responsive CSS.
- [x] Asset files are named and exported with lowercase kebab-case.
- [x] Design avoids generic eco-website aesthetics.
- [x] Design supports both positive tipping and collapse states.

## Implementation boundary

This package is static design output only. It does not include a state machine, keyword parsing, story data wiring, persistence, or real archive storage. Placeholder copy is intentionally labelled where it stands in for future Story agent content.
