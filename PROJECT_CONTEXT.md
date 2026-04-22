# Endangered Species Rescue Mission - Project Context

## What this project is

This is a Google Apps Script web app for elementary or middle-school style classroom use. A student picks an endangered species, researches it, fills out a guided sequence of prompts, and submits a final "rescue report" that generates classroom-ready visual outputs.

The app is not a game in the traditional sense. It is a structured, mission-based learning experience with light game framing:

- students "start a mission"
- choose one species to protect
- gather ecological evidence
- identify threats
- propose conservation actions
- make a final pitch
- generate a finished poster and wall-tile PDF

Primary goal: make research-and-writing feel engaging, guided, and rewarding without making the workflow confusing.

## Current product outputs

When a student finishes, the app generates:

- a Google Slides poster
- a poster PDF
- a separate wall-tiles PDF

If sharing is enabled, the student can open those files from the submission-complete screen.

## Platform / deployment facts

- Platform: Google Apps Script web app
- Frontend: `Index.html`, `ClientJavaScript.html`, `Styles.html`
- Backend: `Code.gs`
- Runtime: V8
- Time zone: `America/Chicago`
- Web app access: `ANYONE`
- Executes as: `USER_DEPLOYING`
- Linked Apps Script project is configured through `.clasp.json`

## Core UX flow

The app currently uses a linear stage flow:

1. `login`
2. `species`
3. `briefing`
4. `ecology`
5. `threats`
6. `actions`
7. `final`
8. `submitted`

Important UX characteristics:

- progress bar across the whole mission
- back navigation between stages
- draft persistence in local storage
- emergency submit option if class ends early
- one selected species per mission
- exactly 2 threats and exactly 2 conservation actions
- final stage includes a 3-image chooser plus a short "why it matters" explanation

## Current visual state

The current UI is functional, readable, and classroom-safe, but visually conservative. It uses an earthy museum/workbook look:

- parchment / paper backgrounds
- dark green hero header
- rounded cards and soft shadows
- simple grid layouts
- practical form styling

This style works, but it is not yet distinctive or especially memorable. The next design pass should improve visual appeal, atmosphere, and delight without making the app harder for students to use.

## Current architecture

### Frontend responsibilities

`Index.html`

- stage markup
- modal markup
- overall shell and layout structure

`ClientJavaScript.html`

- stage transitions and progress
- form validation
- mission start / save / finish calls to Apps Script
- local draft persistence and restore
- image loading logic for species gallery and final image chooser
- graceful fallback when browser image embedding fails

`Styles.html`

- current theme and layout system
- species grid, panels, buttons, modal, image choice styling

### Backend responsibilities

`Code.gs`

- Apps Script entry points
- spreadsheet setup and recovery
- species catalog defaults
- submission saving and stage updates
- poster / PDF / wall-tile generation
- output sharing
- image normalization and image blob retrieval
- research links and species data loading

## Data model

There are 3 main sheets:

### `SPECIES_MASTER`

Stores available organisms and teacher-editable species data, including:

- ids and names
- biome / status / region
- briefing text
- hints
- 3 image URLs
- 2 research links

### `STUDENT_SUBMISSIONS`

Stores student progress and final outputs, including:

- student identity fields
- selected species
- mission stage
- ecology answers
- 2 threats + explanations
- 2 actions + explanation
- why-it-matters response
- selected image URL
- scoring / submission status
- generated file ids and file URLs

### `SETTINGS`

Stores configuration such as:

- output folder
- teacher email
- project title
- sharing behavior
- optional poster template file id

## Built-in species catalog

The default catalog currently includes 9 species:

- Hawksbill Turtle
- Whale Shark
- Blue Whale
- Bornean Orangutan
- Red Panda
- African Forest Elephant
- African Wild Dog
- Black-footed Ferret
- Addax

The app can also be teacher-customized through the species sheet.

## Important current implementation details

### Image handling

This is the most recently stabilized area of the project.

The app now supports a more defensive image pipeline because some URLs worked for final Slides generation but did not render reliably in browser `<img>` tags.

Current image strategy:

- normalize common image URL formats before use
- support Google Drive-style links more safely
- support Wikimedia direct file paths
- support Dropbox raw image conversion
- if the browser cannot embed an image directly, request a server-fetched preview data URI
- disable broken final image options instead of letting students click empty boxes
- preserve the student's manual final image choice across stage revisits when possible

This matters because image reliability was a major practical bug in earlier testing.

### Poster generation

The app supports 2 poster-generation paths:

- template-based poster generation if a valid Slides template id is configured
- fallback fully programmatic poster generation if not

The wall-tile PDF is also generated programmatically.

### State / recovery

The frontend persists drafts in local storage so a refresh does not immediately destroy work. The student can also emergency-submit incomplete work.

## Iteration history and why the current version looks like this

Based on git history plus the current working tree, the project appears to have evolved in this order:

### 1. Initial basic Apps Script app

Early commits were mostly file-by-file scaffolding. The app likely started as a straightforward classroom form with minimal polish and minimal structure.

### 2. Simplification for classroom practicality

Commit history shows the project reduced threats and actions down to 2 each and removed a previous image chooser at one point.

Reasoning:

- fewer choices reduces cognitive overload
- shorter inputs fit classroom time limits better
- less UI complexity improves reliability and load performance

### 3. Better setup and content guidance

Later commits added:

- automatic spreadsheet setup
- improved research links
- cleanup of dead code

Reasoning:

- reduce teacher setup friction
- make the tool more usable out of the box
- keep the student flow more guided

### 4. Richer final deliverables and stronger navigation

A later major feature commit added:

- full poster generation
- wall-tile PDF
- back navigation
- scientific name field
- image chooser

Reasoning:

- make the final artifact feel rewarding and classroom-display ready
- give students more control over their final product
- allow revision instead of forcing linear one-way completion

### 5. Recent stabilization pass for image reliability

The current working version includes a non-trivial image reliability pass beyond the last commit history snapshot. It fixed the bug where:

- species cards could show blank images during species selection
- the final 3-image chooser could show empty boxes that were still selectable
- the final generated poster could still contain the image, proving the problem was browser rendering rather than total image failure

Reasoning behind the fix:

- separate browser-display concerns from backend-fetch concerns
- normalize image URLs before use
- provide fallback previews through Apps Script
- make failure states explicit and non-confusing in the UI

### 6. Small polish pass after stabilization

Recent cleanup also improved:

- preserving selected final images more predictably
- restoring draft state more cleanly
- using species-specific briefing text
- showing both research links
- cleaning some encoded back-button text issues

## What should not be broken in the next design pass

Any redesign should preserve these behavioral truths:

- stage order stays clear and easy to follow
- progress bar remains understandable
- forms stay easy for students to complete quickly
- image selection remains stable and visible
- emergency submit remains obvious
- generated outputs and Apps Script backend behavior remain unchanged unless explicitly requested
- accessibility and readability matter more than flashy visuals

## What is open for redesign

These areas are good candidates for layout and visual exploration:

- hero/header treatment
- stage framing and transitions
- species selection card design
- progress visualization
- reference panel styling
- final pitch / image selection presentation
- overall typography, color system, and atmosphere

Good directions to explore:

- more immersive "mission control" / conservation HQ framing
- cleaner editorial / museum-exhibit framing
- brighter kid-friendly expedition aesthetic
- stronger visual hierarchy and more memorable cards

## Design objective for the next LLM

The next LLM should propose visual directions and layout variants, not rewrite the whole product concept.

It should assume:

- the educational flow is already working
- the main need is a more beautiful and distinctive UI
- the redesign must still be practical for young students in a classroom
- implementation should remain realistic for Apps Script HTML/CSS/JS, not a heavy framework rebuild

## Source of truth files

- [Code.gs](./Code.gs)
- [Index.html](./Index.html)
- [ClientJavaScript.html](./ClientJavaScript.html)
- [Styles.html](./Styles.html)
- [appsscript.json](./appsscript.json)
