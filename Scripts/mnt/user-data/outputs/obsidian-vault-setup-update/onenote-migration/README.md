# OneNote Migration

Scripts for pulling a OneNote notebook into this Obsidian vault, built
against a specific hard case: a deactivated school account whose notebook
went read-only mid-migration, which broke the normal image-export path for
anything inserted as a "printout" (a professor's slide deck, a scanned
worksheet) rather than typed text. Getting full fidelity back out of that
took several iterations — this README is the map so future-you doesn't have
to rediscover the same dead ends.

## The actual workflow (do this)

1. **OneNote Markdown Exporter** (`segunak/one-note-to-markdown`, external
   tool, not in this repo) does the first pass: converts every page in a
   notebook to `.md`, preserving the section/subpage folder structure. Run
   it once per notebook into a staging folder — this becomes the source of
   truth for folder placement in every later step.

2. **`current/Rebuild-CourseAttachments.ps1`** — the main script, run once
   per course:
   - Wipes and recopies that course's folder in the vault from the clean
     Markdown export (so re-runs are always safe and idempotent)
   - Cleans up duplicate titles the exporter creates when a page name has a
     Windows-illegal character (e.g. "Chapter 1: Intro" becomes two H1s:
     one from the sanitized filename, one from the real title — this keeps
     the real one)
   - Talks to OneNote directly via COM to find every page's real file
     attachments (not just the broken auto-generated preview images) and
     recovers them by reading OneNote's local `pathCache`, since that
     works even with a read-only/deactivated account
   - Renames each recovered attachment to match its page name, places it in
     a nested `Attachments/` folder, and embeds PDFs inline in the note
     (plain link for non-PDF originals, since Obsidian can't preview those)
   - Exports the whole course as one backup PDF too, as a last-resort
     safety net for pages with genuinely no recoverable file
   - Backs up every recovered file to a separate staging location
     (`C:\Users\<you>\Backups\Purdue-Attachments\` by default) before
     anything touches the vault
   - Supports `-SkipPages` to protect specific pages you've already fixed
     by hand from being overwritten on a re-run

3. **`current/Convert-AttachmentsToPDF.ps1`** — run after the above (once
   per course, or across the whole vault with no `-CourseName`). Finds any
   recovered `.pptx`/`.doc`/`.docx` that doesn't have a PDF twin, converts
   it via PowerPoint/Word automation, and embeds the new PDF in the note.
   Runs each conversion in an isolated background job with a 60-second
   timeout — an earlier version of this hung for 8 hours on an invisible
   Word dialog, so don't remove that safety net.

4. **`current/Rebuild-AllCourses.ps1`** — batch driver for step 2. Lists
   courses **explicitly by name**, not "all except the done ones" — safer
   given how many rounds of bugfixes this took. Update the list by hand
   before running.

5. **`current/Debug-OneNotePage.ps1`** and **`current/Export-SectionsToPDF.ps1`**
   are utilities, not part of the main flow: the debugger dumps a single
   page's raw XML when something's not matching expectations, and the
   section-exporter is the standalone version of what step 2 already does
   internally per course (useful if you just want to refresh one course's
   backup PDF without a full rebuild).

## Known quirks worth remembering

- **Character sanitization varies by symbol.** The exporter replaces `:`
  with `_` but `/` with `-` in filenames — normalize BOTH sides (page name
  and filename) by stripping all of `\/:*?"<>|-_` before comparing, or
  matches silently fail.
- **UTF-8 BOM breaks `^#` regex anchors.** Strip a leading BOM
  (`[char]0xFEFF`) before pattern-matching a file's first line, or
  title-cleanup logic will silently do nothing.
- **The broken-image placeholder text is wrapped in single asterisks**
  (`*[Image - no embedded data...]*`) — strip the asterisks too, or you'll
  leave orphaned `**` litter behind.
- **`callbackID="none"` in a page's raw XML means the image was never
  generated** (usually because the account is read-only) — there's nothing
  to fetch, it's not a bug to chase.
- **Force-killing OneNote repeatedly can corrupt its local state** — this
  happened once and required a full reboot to fix. The restart-cap in
  `Rebuild-CourseAttachments.ps1` (2 max per run) exists because of that;
  don't remove it.
- **A page's real ancestor is its Section, not a parent page** — OneNote
  represents subpages as flat siblings with a `pageLevel` attribute, not as
  real XML nesting. Don't try to compute folder placement independently;
  match by searching for where the already-exported note actually landed.
- **Windows PowerShell 5.1 (not PowerShell 7) misreads non-ASCII
  characters** (em dashes, emoji, curly quotes) in `.ps1` files, causing
  baffling "missing terminator" parse errors on lines that look fine.
  Keep script text plain ASCII.

## `legacy/` — earlier iterations, kept for reference

These were superseded by the workflow above but aren't deleted, since they
document approaches that were tried and why they didn't stick:

- `Extract-OneNoteAttachments.ps1` — the first standalone attachment
  puller, before it got folded into the per-course rebuild script
- `Merge-PurdueIntoVault.ps1` — an earlier three-stage process (export,
  extract, merge as separate manual steps) before it became one script
- `Link-RecoveredAttachments.ps1` — the first version of note-linking,
  before it moved inline into the rebuild script
- `Fix-BrokenPages.ps1` — a page-level PDF export approach (one small PDF
  per broken page) that was tried and dropped in favor of recovering real
  attachment files directly, which give better fidelity than any rendered
  page image
- `Find-ThinExports.ps1` — an early word-count heuristic for spotting
  broken pages, superseded by precise detection via OneNote's raw XML
