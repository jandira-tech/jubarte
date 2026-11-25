<!--
SPDX-FileCopyrightText: 2026 Jandira Technologies, LLC

SPDX-License-Identifier: Apache-2.0
-->

# Comparison settings

`WmlComparerSettings` configures a comparison. Build it from a preset with
`WmlComparerSettings::new(CompareMode::Word)` (the same as `default()`),
`CompareMode::PowerTools` or `CompareMode::WordWindows`, then change the
scalars that are safe to tune with the `with_*` methods:

| Method | Sets | Default |
|---|---|---|
| `with_author` | `author_for_revisions`, the `w:author` on every revision | `Open-Xml-PowerTools` |
| `with_date` | `date_time_for_revisions`, the `w:date` (ISO 8601) | `1970-01-01T00:00:00Z` |
| `with_detail_threshold` | `detail_threshold`: how little a paragraph may share before it is replaced whole instead of marked word by word | `0.02` (Word), `0.15` (PowerTools) |
| `with_input_limits` | `input_limits`, the budget both inputs are admitted under | `InputLimits::compare()` |

The presets:

| | `CompareMode::Word` (default) | `CompareMode::PowerTools` | `CompareMode::WordWindows` |
|---|---|---|---|
| Reproduces | Word Compare on Mac Word 16: word-level detail inside paragraphs, replaced paragraphs merged as Word merges them, Word's alignment passes | Open-Xml-PowerTools and Docxodus | Word Compare on Windows Word 2024, where its redline differs from Mac Word 16's |
| `word_platform` | Mac | Mac | Windows |
| `detail_threshold` | 0.02 | 0.15 | 0.02 |
| `detect_moves` | on | off | on |
| Word-visual passes (`merge_replaced_paragraphs`) | on | off | on |

`WordWindows` is the `Word` preset with `word_platform` set to
`WordPlatform::Windows`, and nothing else differs. The places where the two
Words' Compare writes different markup read that field: a paired table row
whose two sides have different cell counts gets `cellDel` on the original's
extra cells and `cellIns` on the revision's, and every replacement is
written deletion first, then insertion, inside a paragraph and between the
paragraphs of a replaced block (#2550), where Mac Word 16 writes the
insertion first.

The other fields are public for existing struct-update callers:
`detect_moves`, `move_similarity_threshold` (0.9) and
`move_minimum_word_count` (6) for moves; `detect_format_changes` (on) for
`w:rPrChange` and `w:pPrChange`; `simplify_move_markup` to write moves as
plain insertions and deletions; `case_insensitive`,
`conflate_breaking_and_nonbreaking_spaces` (on) and `word_separators`.
Setting the alignment switches by hand gives a configuration no oracle
covers; prefer a preset.

```rust,ignore
{{#include ../../../../crates/jubarte-redlines/examples/compare_settings.rs:example}}
```

Run it with `cargo run -p jubarte-redlines --example compare_settings`.

`get_revisions` reads a redline back as `WmlComparerRevision`s: the
type (inserted, deleted, moved, format change), text, author, date, part,
and the move group of a move. `revisions_to_json` writes the same list in
the JSON shape the Python and WebAssembly bindings return.

On the command line, `--mode word` (the default), `--mode powertools` and
`--mode word-windows` pick the preset.
