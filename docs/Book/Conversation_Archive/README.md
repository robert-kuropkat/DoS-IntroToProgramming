# Conversation Archive

Raw material for the eventual **annotated-conversation case study** (the third
document in this set, alongside `AI_Development_Case_Study.docx` and
`AI_Development_Production_Log.docx` in the `Book` folder). This folder holds
the real conversation data, extracted and restructured into a format meant to
be processed by Claude in small segments over time — not meant to be read
directly by a person. Robert has no format preference for this layer; the
only requirement is that it stays processable later.

## Where this came from

Robert exported local session data from the Claude desktop app's own internal
cache (via `Export_20261019-2217`, NOT the claude.ai account export, which
does not contain Cowork session data at all — see the Case Study's note on
that dead end). That export contained three local Cowork sessions as
`local_<id>.zip` files, each wrapping a `.claude/projects/.../<uuid>.jsonl`
raw transcript log. This folder is the result of running those raw logs
through `extract_transcript.py` (kept here for re-use/re-processing) to pull
out just the real conversational turns, stripped of internal tool-call noise,
thinking blocks, and queue/mode/cost bookkeeping entries.

## Files

- **`main_session_conversation.jsonl`** — the primary artifact. The full,
  real conversation for this entire project, from Robert's very first message
  ("I want to start building out my class...", 2026-07-09) through the
  present day. Turns 1-391 came from the raw log export described below.
  From turn 392 onward, turns are appended live (see Live-Append Policy
  below) rather than re-extracted from a fresh export each time. As of this
  update, live-appended content runs through turn 399, with one
  `compaction_suspected` marker (near turn 398) flagging a compaction Robert
  noticed between the archive-policy turn (397) and the Lesson 12 fix work
  that followed it. One JSON object per line; see Schema below.
- **`main_session_manifest.jsonl`** — a lightweight index over the same
  turns (turn_id, timestamp, a ~100-character preview, tool-call count,
  character counts, source). Meant to be read in full in one shot to decide
  which turn-ranges to pull from the main file next, without loading the
  full conversation into context.
- **`dispatch_a546e064_lesson6_sfx.jsonl`** and
  **`dispatch_b3364c44_lesson6_sfx_fixes.jsonl`** — two short, separate
  Claude Dispatch runs (the kind Robert can kick off remotely, e.g. from his
  phone), both titled "Lesson 6 Retro Sound Effects," dated 2026-09-14. The
  first generated real `sfx_start.wav` / `sfx_win.wav` assets for Lesson 6;
  the second fixed two Standards-Rule violations once those files existed
  (removing "reuse or create a sound" language that had been asking students
  to produce an asset, a Rule 3 self-containment break). Genuinely interesting
  material for the Case Study's section on Dispatch as a usage pattern, even
  though Robert says he hasn't yet taken full advantage of the feature.
- **`extract_transcript.py`** — the extraction script itself, kept here so
  the same process can be re-run if more raw session exports turn up later
  (e.g. if Robert finds additional local session zips), or re-run with a
  different filtering/summarization strategy without re-deriving the parsing
  logic from scratch.

## Schema (`record_type` discriminates two kinds of line)

**`"record_type": "turn"`** — one real back-and-forth exchange:
```json
{
  "record_type": "turn",
  "turn_id": 1,
  "timestamp_start": "2026-07-09T15:00:54.527Z",
  "timestamp_end": "2026-07-09T15:12:03.112Z",
  "human_text": "<Robert's actual message, verbatim>",
  "assistant_text": "<Claude's visible reply text, concatenated across all text blocks in the turn, paragraphs joined by blank lines>",
  "tool_calls": [ { "name": "Read", "input_brief": "file_path=..." }, ... ],
  "tool_call_count": 57,
  "commentary": "",
  "tags": [],
  "source": "raw_log"
}
```
`source` is `"raw_log"` for a turn extracted directly from an exported raw
transcript (ground truth, exact text) and `"live_append"` for a turn written
in by Claude from its own live context during an ongoing conversation (see
Live-Append Policy below) — these are reconstructed/condensed rather than
pulled verbatim from a log file, so they're marked distinctly. Turns 1-391
predate this field and have no `source` key at all; treat a missing
`source` as equivalent to `"raw_log"`.

`commentary` and `tags` are intentionally empty placeholders — that's where
the annotated-walkthrough pass fills in observations ("this is where the
GameManager pattern got generalized," "this request was underspecified and
should have been clarified first," etc.) turn by turn, without needing to
touch the original transcript data.

`tool_calls` is a *brief* record of what tools ran during that turn (tool
name plus a short, truncated summary of the call's main input field) — not
the tool outputs. Tool results were deliberately dropped during extraction;
they're often huge (full file contents, long command output) and rarely
needed to understand the narrative arc of a turn. If a specific tool result
genuinely matters for commentary on a specific turn, re-derive it from the
original session zip (paths below) rather than bloating this archive.

Internal "thinking" blocks (Claude's private reasoning before responding)
were also dropped — they're not part of what Robert actually saw, and
including them would roughly double the file size for content that isn't
part of the actual collaboration record.

**`"record_type": "compaction_boundary"`** — a marker showing where the
live conversation's context window was compacted (summarized) during the
session, interleaved in the correct position in the turn sequence:
```json
{ "record_type": "compaction_boundary", "index": 7, "timestamp": "..." }
```
The main session file has 39 of these (all from the original raw-log
extraction, turns 1-391). They're genuinely useful content for the
Production Log — they mark exactly where detail was lost from Claude's
working context during the live project (which is part of why this
third-document reconstruction effort came up in the first place), and they
roughly bound natural "chapters" for an annotated walkthrough.

**`"record_type": "compaction_suspected"`** — a lighter-weight cousin of
`compaction_boundary`, used from turn 392 onward. Claude doesn't have a
tool that directly detects a compaction event the way the raw log does; it
can only infer one happened (a sudden loss of earlier detail, a system
reminder, a conversation that was clearly continued from a summary). This
marker exists so that inference gets written down in the moment rather than
lost, flagging a stretch for priority re-checking at the next raw-log
resync:
```json
{ "record_type": "compaction_suspected", "near_turn_id": 420, "timestamp": "...", "note": "..." }
```

## Live-Append Policy (adopted 2026-10-19)

Starting at turn 392, Claude appends new real turns to this archive
opportunistically — at the end of a working session, or whenever it's a
natural pause point — using what's currently in its own context. This
keeps the archive close to current without requiring Robert to re-export
the local cache every time. The tradeoff, stated plainly: a live-appended
turn is only as good as what's still in Claude's context when it's written,
so anything lost to a compaction *before* being appended is gone the same
way the original pre-July-9-export conversation was. A `compaction_suspected`
marker (above) is the cheap mitigation — it can't recover the lost detail,
but it marks exactly where to look during the next resync.

Agreed resync cadence (ground-truth re-export from the local cache,
re-run through `extract_transcript.py`, then reconciled against whatever
was live-appended in the meantime): roughly monthly, at major project
milestones (Lesson 15 completion / end of Semester 1 being the next one
flagged), or ahead of any dedicated annotation-work session, so detail
doesn't fade before it's captured. When a resync happens, diff the two
manifests by turn content/timestamp rather than assuming turn_ids line up
exactly, since live-appended turn_ids are assigned sequentially by Claude
and may not match the numbering a fresh raw-log extraction would produce.

## Intended workflow (not started yet)

The plan Robert described: walk the conversation in manageable segments
(e.g. by compaction-boundary chapter, or by a fixed turn-count window),
reading each segment's turns from `main_session_conversation.jsonl`,
discussing what happened, and filling in each turn's `commentary` field
in place — where something worked well, where a request was underspecified,
where a bug was found, where a Standards Rule got written, where an approach
got refactored after a dead end. The `main_session_manifest.jsonl` index
exists specifically so a future session can orient quickly (find the turn
range for "the GameManager retrofit" or "the Unreal rebuild") without
re-reading hundreds of turns first.

Once a meaningful stretch is annotated, that's the point to decide on final
output format(s) — options include weaving selected annotated turns directly
into the Case Study as sidebars/pull-quotes, building a standalone annotated
transcript document, or keeping the annotation layer purely as source
material that informs prose elsewhere. Not decided; deliberately deferred,
per Robert's direction, until there's actual annotated content to look at.

## Where the originals still are

The raw, unprocessed export (three `local_<id>.zip` files plus matching
`.json` session-metadata files) is still sitting in
`Export_20261019-2217/` in the project root, untouched. This archive is a
derived, cleaned copy — if the extraction logic needs to change, re-run
`extract_transcript.py` against the originals there rather than editing
these files by hand.
