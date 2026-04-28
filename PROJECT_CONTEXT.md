# Endangered Species Rescue Mission - Project Context

## What This Project Is

This is a Google Apps Script classroom web app where students complete a guided endangered-species research mission. The app frames the work as a conservation HQ workflow: students choose an organism, gather basic ecological evidence, identify two threats, propose two conservation actions, choose a final image, and submit a polished rescue-report PDF.

The goal is not a traditional game with scoring loops. It is a structured, mission-themed learning experience that makes research and writing feel more focused, visual, and rewarding for science students.

## Current Student Flow

The live flow is:

1. `login`
2. `species`
3. `briefing`
4. `ecology`
5. `threats`
6. `actions`
7. `final`
8. `submitted`

Important behavior to preserve:

- Students complete one species per mission.
- The mission uses exactly two threats and two conservation actions.
- The final stage includes a three-image chooser and a short "why this species matters" response.
- Draft work is saved in browser local storage.
- Emergency submit remains available for classroom timing problems.
- The final student-facing output is the PDF only.

## Current Output Model

The app still creates a Google Slides presentation internally because Apps Script exports Slides to PDF reliably. That Slide should be treated as an internal render artifact, not the student-facing product.

The intended final output is:

- one rescue-report PDF saved to the configured Google Drive folder
- no wall-tile output in the active student experience
- no student-facing Slide link

Some legacy spreadsheet columns still exist for older slide and wall-tile outputs so existing data is not broken, but new UX should stay PDF-first.

## Platform Facts

- Platform: Google Apps Script web app
- Frontend files: `Index.html`, `ClientJavaScript.html`, `Styles.html`
- Backend file: `Code.gs`
- Runtime: V8
- Time zone: `America/Chicago`
- Web app access target: `ANYONE`
- Executes as: `USER_DEPLOYING`
- Apps Script link is stored in `.clasp.json`

## Current Visual State

The main web app has been redesigned into a mission-control dashboard style:

- left mission rail
- top progress/status bar
- staged dashboard cards
- final image selection
- emergency submit affordance

The final PDF is a single-page endangered-species rescue report inspired by the reference mockup in:

`C:\Users\Keyur\Downloads\ChatGPT Image Apr 23, 2026, 07_35_52 PM.png`

The target PDF style is:

- dark green conservation-report frame
- large species title and scientific name
- large image panel
- quick facts card
- ecology, threats, conservation, and "why this species matters" cards
- strong icon/crest system
- footer call-to-action

The current programmatic renderer is good enough to ship, but it is not yet as polished as the target image. The biggest visual gap is the icon/crest system and the fact that the poster is hand-drawn through Apps Script coordinates instead of designed in a real Slides template.

## Important Architecture Notes

`Code.gs` currently owns most backend behavior:

- Apps Script entry points
- spreadsheet setup and data access
- species defaults
- submission saving
- image normalization and preview fetching
- poster/PDF generation
- test batch generation

`ClientJavaScript.html` owns:

- stage transitions
- form validation
- draft persistence
- image preview fallbacks
- Apps Script calls
- final PDF link display

`Index.html` owns stage markup and the mission shell.

`Styles.html` owns the mission-control web UI styling.

## Poster Renderer Notes

The active fallback poster renderer is programmatic. It builds a Google Slide using native shapes/text and exports that Slide as a PDF.

Known renderer constraints:

- Apps Script Slides support is more limited than browser HTML/CSS.
- Runtime SVG icon insertion has failed in prior tests, causing missing icons or failed sections.
- Safer font normalization was added because some web fonts disappeared during PDF export.
- Text is currently constrained by display helpers and fixed text boxes; this is better than raw character cuts, but still less robust than a real template.

Recommended future visual improvement:

- Build a real Google Slides template matching the target mockup.
- Configure `poster_render_mode=template` and `poster_template_file_id` when ready.
- Use the programmatic renderer as fallback only.

## Data Model

Main sheets:

- `SPECIES_MASTER`
- `STUDENT_SUBMISSIONS`
- `SETTINGS`

Default species currently include:

- Hawksbill Turtle
- Whale Shark
- Blue Whale
- Bornean Orangutan
- Red Panda
- African Forest Elephant
- African Wild Dog
- Black-footed Ferret
- Addax

## Iteration History

The project evolved roughly like this:

1. Basic Apps Script classroom form.
2. Simplified classroom flow with two threats and two actions.
3. Added spreadsheet setup, species defaults, and stronger research links.
4. Added poster generation, back navigation, scientific names, and image selection.
5. Fixed image reliability so species cards and final image choices no longer showed selectable blank boxes.
6. Redesigned the web app into a mission-control dashboard.
7. Removed wall tiles from the intended live UX.
8. Iterated heavily on the final PDF renderer after issues with missing text, fragile SVG icons, cramped conservation cards, and image stretching.
9. Switched the student-facing final product to PDF-only because PDF quality is better than the intermediate Slide.
10. Created the `5.5-attempt` branch to clean the shipped implementation and move the PDF closer to the ideal mockup.

## Current 5.5 Branch Goals

This branch is intended to:

- preserve the working student flow
- keep the final student product PDF-only
- reduce legacy wall-tile and duplicate renderer residue
- keep E2E testing available but less hard-coded
- improve poster text fitting and renderer maintainability
- push the branch to Apps Script for visual review with three generated PDFs

## Things Future LLMs Should Avoid

- Do not reintroduce wall tiles as a student-facing output.
- Do not expose the internal Google Slide as the final student artifact.
- Do not make broad app redesigns while debugging poster output.
- Do not rely on runtime SVG icons unless verified through exported PDF.
- Do not assume character count equals visual fit in Slides text boxes.

## Source Files

- `Code.gs`
- `Index.html`
- `ClientJavaScript.html`
- `Styles.html`
- `appsscript.json`
