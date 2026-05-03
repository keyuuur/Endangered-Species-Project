# Endangered Species Rescue Mission - Project Context

## Purpose Of This File

This file is a handoff brief for another LLM, specifically GPT-5.5 Pro, to review the current codebase and suggest the next set of improvements. It should give enough project history and current-state context to avoid repeating failed iterations.

The teacher/user is a non-traditional programmer learning as they go. Explanations should be plain English, concrete, and classroom-oriented.

## Project Summary

This is a Google Apps Script classroom web app for a 9th grade endangered species research activity. Students complete a guided "conservation mission" instead of a plain worksheet. The app helps them choose an organism, collect ecology information, identify two threats, choose two conservation actions, select a final image, and submit a final rescue-report PDF.

The project is not a score-heavy game. It is a structured research and writing workflow with a mission-control visual theme.

Core goals:

- Make endangered species research feel focused, visual, and student-friendly.
- Keep the workflow simple enough for iPad classroom use.
- Preserve the teacher's Google Apps Script workflow with no build tools or external frameworks.
- Save final student artifacts into the configured Google Drive destination folder.
- Produce a polished single-page PDF that science department colleagues can review.

## Current Repo And Deployment State

- Local repo path: `C:\Users\Keyur\Desktop\Claude Code YEET\Biology Games\Conservation LOCAL\Endangered Species Project LOCAL`
- Current branch at last update: `5.5-attempt`
- Current branch status at last update: clean and synced with `origin/5.5-attempt`
- Latest relevant commit: `9ddce6b Add template-first poster renderer`
- Apps Script project ID in `.clasp.json`: `1QSBVW3I45aY5cF4sZveGwueDdQx3ogny3Uh8ZIg9IWgboTEDXyKZP3tj`
- Live deployment used for testing: `https://script.google.com/macros/s/AKfycby--cP2u7g7ZSi4MHRxhO3ejQ7589TG3YyDdmd_SIrpaav4QwIYBC1_B--e3PmgW7Xn/exec`
- Most recent Apps Script deployment version from this branch: `53`
- Manifest target access: `ANYONE`
- Manifest execute-as: `USER_DEPLOYING`
- Runtime: Google Apps Script V8
- Time zone: `America/Chicago`

Important deployment note:

Apps Script deployment access sometimes flips or behaves as if it requires Google sign-in after redeployment. When that happens, the E2E endpoint returns Google sign-in HTML instead of JSON. The teacher has been manually changing the deployment access back to `Anyone`.

## Source Files

- `Code.gs`: backend, spreadsheet setup, species data, validation, submission saving, image helpers, poster/PDF generation, E2E test helper.
- `Index.html`: raw Apps Script HTML shell and student stage markup.
- `ClientJavaScript.html`: raw frontend JavaScript for stage transitions, local draft saving, validation, image preview fallbacks, Apps Script calls, and output-link rendering.
- `Styles.html`: mission-control UI CSS.
- `appsscript.json`: Apps Script manifest.
- `.clasp.json`: Apps Script push target.
- `PROJECT_CONTEXT.md`: this handoff file.

Project constraints:

- Keep everything compatible with Google Apps Script.
- Do not add npm packages, bundlers, transpilers, or external frameworks.
- Keep HTML/CSS/JS files as clean raw text that can be copied into Apps Script.
- Optimize the student UI for iPad classroom use.
- Do not change where student files are saved without teacher confirmation.

## Current Student Flow

The student flow is:

1. `login`
2. `species`
3. `briefing`
4. `ecology`
5. `threats`
6. `actions`
7. `final`
8. `submitted`

Important behavior to preserve:

- Students complete one organism per mission.
- Students choose exactly two threats.
- Students choose exactly two conservation actions.
- Students choose one final image from three options.
- Students write a "why this species matters" final response.
- Draft work is saved in browser local storage.
- Emergency submit remains available for classroom timing issues.
- The final student-facing artifact is a PDF only.

## Current Student-Facing UI

The web app has already been redesigned into a "Mission Control / Conservation HQ" dashboard style.

Current visual language:

- Left mission rail with stages.
- Top progress/status bar.
- Dark teal/green conservation dashboard.
- Card-based stage screens.
- Final image selection UI.
- Emergency submit affordance.
- Final submitted screen shows only the poster PDF link.

Known UI state:

- The main web app functionality is considered solid.
- Earlier image-preview bugs were fixed so species images and final image choices are visible instead of blank selectable boxes.
- Wall tiles were removed from the student-facing experience because they were not useful.
- Slide links were removed from the student-facing final output because PDFs looked better and were the only artifact the teacher wanted.

## Current Output Model

The final output should be:

- One rescue-report PDF in the configured Drive folder.
- No wall-tile PDF.
- No student-facing Google Slides link.

Implementation detail:

The app still creates a Google Slides presentation internally because Apps Script can export Slides to PDF reliably. That internal Slide is treated as a render artifact. After the PDF is exported, the internal Slide is trashed by `archiveInternalPosterSlide_`.

Important consequence:

- `poster_file_id` may point to a trashed internal Slide. This is intentional in the current PDF-only workflow.
- `poster_pdf_url` is the student-facing artifact.

## Data And Settings

Main sheets:

- `SPECIES_MASTER`
- `STUDENT_SUBMISSIONS`
- `SETTINGS`

Configured output folder:

- `APP.outputFolderId = 1BESzdmXkmswTBApTqgiAeBZxLHgQo2mY`

Default settings currently include:

- `poster_template_file_id = PASTE_TEMPLATE_FILE_ID_HERE`
- `output_folder_id = APP.outputFolderId`
- `teacher_email = PASTE_TEACHER_EMAIL_HERE`
- `project_title = Endangered Species Rescue Mission`
- `share_output_with_link = TRUE`
- `show_student_output_links = TRUE`
- `e2e_test_key = ''`
- `poster_template_version = template_first_v1`
- `use_template_poster = TRUE`
- `poster_render_mode = template`

Legacy columns still exist in `STUDENT_SUBMISSIONS`:

- `poster_slide_url`
- `tile_pdf_file_id`
- `tile_pdf_url`

These are retained for schema compatibility. Do not reintroduce wall tiles or student-facing Slides unless the teacher explicitly asks.

## Species Catalog

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

The E2E helper uses 9th-grade-style sample writing rather than "GPT scientist" writing. This was intentional because the teacher wants examples that resemble plausible student responses.

## Poster/PDF Target

The desired final PDF is a polished one-page endangered species rescue report inspired by this local reference image:

`C:\Users\Keyur\Downloads\ChatGPT Image Apr 23, 2026, 07_35_52 PM.png`

The target visual style includes:

- Dark green conservation-report frame.
- Large species title and scientific name.
- Strong crest/logo in the upper left.
- Large organism image panel.
- Quick facts card.
- Ecology card.
- Threats card.
- Conservation actions card.
- "Why this species matters" card.
- Footer call-to-action.
- Decorative ocean/leaf/coral background elements.
- Clean icon system.
- No text overlap.
- Stronger typography than default Arial-heavy output.

## Current Poster Renderer State

The current production path is "template-first" in name, but this needs an important warning.

Current behavior:

- `generatePosterForSubmission_` calls `ensurePolishedPosterTemplate_`.
- `ensurePolishedPosterTemplate_` creates or reuses a Slides template named `_Template - Endangered Species Rescue Report`.
- `buildPosterFromTemplate_` copies that template, replaces text placeholders, inserts the selected image into `POSTER_IMAGE`, exports PDF, and trashes the internal Slide.
- If the template path fails, `buildFallbackPoster_` uses the older coordinate-drawn renderer.

Critical lesson from the latest iteration:

The generated template was created by Apps Script drawing primitives, so visually it still looked too much like the old code-drawn poster. The teacher reviewed the Addax output from the template-first attempt and said it was not the desired end product. The architecture improved, but the visual design did not improve enough.

What this means:

- The next serious improvement should not be another small coordinate tweak.
- The next serious improvement should use a truly designed Google Slides template or a high-quality static designed background with Apps Script placing text/image into it.
- The current auto-created template can remain as a fallback, but should not be considered the final design solution.

## Important Poster Functions

Key functions in `Code.gs`:

- `runThreeEndToEndTests()`: creates 3 randomized test submissions and returns PDF URLs.
- `doGet(e)`: serves the app and exposes the E2E test endpoint when authorized.
- `generatePosterForSubmission_(submission)`: main poster/PDF generation flow.
- `ensurePolishedPosterTemplate_(settings, outputFolder)`: creates/reuses the current generated template.
- `createPolishedPosterTemplate_(outputFolder)`: creates the current generated template through Apps Script shapes.
- `drawPolishedPosterTemplate_(slide)`: draws the current generated template.
- `buildPosterFromTemplate_(templateId, outputFolder, posterName, submission)`: copies template, replaces placeholders, inserts image, saves.
- `replaceTemplateImagePlaceholder_(slide, imageUrl)`: finds the `POSTER_IMAGE` element by title and inserts the selected image.
- `buildFallbackPoster_(outputFolder, posterName, submission)`: older code-drawn poster fallback.
- `exportSlidesFileAsPdfWithRetry_(presentationId, posterName, outputFolder)`: exports the Slide artifact as PDF.
- `archiveInternalPosterSlide_(posterFile)`: trashes the internal Slide after PDF export.

## E2E Testing

The E2E endpoint is:

`https://script.google.com/macros/s/AKfycby--cP2u7g7ZSi4MHRxhO3ejQ7589TG3YyDdmd_SIrpaav4QwIYBC1_B--e3PmgW7Xn/exec?action=runE2ETests&key=codex-e2e-2026`

Notes:

- The fallback key `codex-e2e-2026` exists for testing convenience and is not a strong production secret.
- If Apps Script access is set correctly, the endpoint returns JSON with three PDF URLs.
- If access is broken, it returns Google sign-in HTML.
- Current expected successful output is 3 PDFs, not 6 files, because Slides are internal and trashed.

Recent successful test after version 53:

- Addax PDF generated successfully.
- Blue Whale PDF generated successfully.
- Bornean Orangutan PDF generated successfully.

But the Addax visual result was judged not close enough to the ideal target.

## Iteration History

High-level path to current state:

1. Started as a basic Apps Script classroom form.
2. Simplified into a guided research workflow with one species, two threats, and two conservation actions.
3. Added spreadsheet setup, species defaults, student submissions, settings, and research links.
4. Added poster generation and image selection.
5. Fixed image display reliability so organism-choice images and final image choices no longer showed blank boxes.
6. Redesigned the web app into a mission-control dashboard style.
7. Added wall tiles as a possible final output, then removed them because they were not useful.
8. Iterated heavily on the final PDF renderer to fix missing text, fragile icons, cramped conservation actions, image stretching, and overlapping text.
9. Removed student-facing Slides and kept PDF-only output.
10. Merged/pushed a stable version to main, then created the `5.5-attempt` branch for a stronger model review and improvement attempt.
11. Cleaned PDF-only workflow, removed duplicated renderer residue, and kept old wall-tile columns only for schema compatibility.
12. Added a "template-first" renderer path, but the template was generated by code and therefore did not achieve the desired visual leap.

## Current Problem To Solve

The core app works. The main remaining problem is final PDF visual quality.

The current final product is functional but still not close enough to the ideal polished rescue-report mockup. The teacher wants something visually impressive enough to showcase to the science department.

Most likely next direction:

- Keep the app flow and data model stable.
- Keep PDF-only output.
- Replace the generated-template approach with a real design template.
- Use Apps Script only to fill content into that design.

Potential implementation paths:

- Create a real Google Slides template manually, set `poster_template_file_id`, and update `buildPosterFromTemplate_` to support all needed placeholders and image frames.
- Use a high-quality static background image matching the ideal poster and overlay dynamic text/image boxes in Apps Script.
- If using a static background, ensure text areas are large enough for 9th-grade responses and that exports stay crisp.
- Keep `buildFallbackPoster_` as a safety fallback.

## Review Priorities For GPT-5.5 Pro

Please review with these priorities:

1. Identify the best architecture for making the final PDF match the target mockup more closely inside Apps Script constraints.
2. Decide whether the current generated-template code should be removed, kept as fallback, or converted to consume a real hand-designed template.
3. Check for bugs or fragility in `generatePosterForSubmission_`, template creation, image insertion, PDF export, and Slide trashing.
4. Check text fitting strategy. Character truncation is not the same as visual fit in Slides.
5. Check whether icon/SVG insertion is safe enough or should be replaced with static template assets.
6. Check whether E2E testing should be made more secure without making teacher testing painful.
7. Suggest cleanup only if it improves reliability or maintainability without disrupting the working classroom flow.

## Things To Avoid

- Do not reintroduce wall tiles as a student output.
- Do not expose the internal Google Slide to students.
- Do not change the Drive output folder without teacher confirmation.
- Do not add npm packages, frameworks, or build steps.
- Do not redesign the already-working web app while trying to fix the PDF.
- Do not assume the auto-generated template solved the visual problem.
- Do not rely on runtime SVG icons unless exported PDFs prove they render correctly.
- Do not optimize for desktop over iPad classroom use.

## Teacher Preference Notes

- The teacher wants plain English explanations and step-by-step test checklists.
- The teacher is comfortable with Git as a safety net and wants meaningful improvements, not overly timid edits.
- The teacher prefers functional classroom reliability over clever architecture.
- The teacher wants final PDF examples to look like plausible student work, not overly polished scientific writing.
- The teacher is open to a stronger model proposing a better approach if the current Apps Script drawing method is the limiting factor.
