---
name: word-enhanced
description: Use when modifying, editing, formatting, adjusting, or repairing an existing Microsoft Word DOCX while preserving unspecified content. Includes filling an existing DOCX template. Do not use for creating a new DOCX from scratch.
---

# Word Enhanced

Use this workflow for an existing `.docx` that the user wants changed. An existing template filled with new content is an edit. If the request is to create a document from nothing, or only use a reference as inspiration for a new document, leave this workflow and use the normal document creation capability.

## Before editing

1. Keep the source intact. Default to a new output path; overwrite only when the user explicitly asks.
2. Restate the target, original value or state, and what must stay unchanged. Use `inspect_docx` and focused read-only inspection to locate the target and its surrounding paragraphs, table, section, field, bookmark, relationship, or drawing. Check the match count and original state before writing. If several candidates remain, resolve the ambiguity; do not guess.
3. Choose **one primary editing route** based on the actual structure and request. Do not chain editors on the same output. A failed route can be retried only from a fresh copy of the original.

## Choose the route

| Route | Choose when | Boundary |
| --- | --- | --- |
| OOXML Surgical | The target is unique, its existing OOXML location and expected value are known, and a small node or attribute edit expresses the change without Word recalculation. This can include run or paragraph format, cell shading, hyperlink or bookmark vicinity, a known field vicinity, or an existing drawing anchor value. | Verify context, relationships and preservation conditions first. Prefer one attribute over a rebuilt run, one node over a rebuilt paragraph, and one part over other parts. Use `patch_docx_part_exact` only with an exact fragment and expected count. Unknown object structure or mismatched original value ends this route. The four experiments support this route only for the four tested object types, not arbitrary DOCX objects. |
| OpenAI Documents | Natural-language target finding or more complex structure understanding helps, while programmatic editing remains suitable. | Follow the installed official `documents:documents` Skill and its own editing rules; this Skill does not implement Documents. Then apply Word Enhanced comparison and acceptance. Availability of the Skill does not imply every requested edit has a ready-made helper. |
| Microsoft Word native | The result depends on Word's own field, Track Changes, section or other proprietary behavior. On macOS the Word object model via AppleScript/JXA can be used. | Edit a copy, save once as needed, then independently inspect the saved file. Word saving may alter non-target structures; the observed side effects are not universal laws. |
| Computer Use + Word | Correctness is mainly visual, such as moving an object until it no longer overlaps text or aligning a region by eye. | Confirm the selection and control state, exclude paragraph marks unless intended, and check the saved value and layout. Do not choose mouse interaction merely for convenience when an exact value or XML property can be safely addressed. GUI values may be quantized by Word. |

An image with an exact known `posOffset` change may suit Surgical; “move it a little so it no longer blocks the heading” may suit Computer Use. Reconsider the route when the document structure does not support the initial choice.

## Validate before delivery

1. Run `compare_docx` on original and output. Distinguish raw package bytes from normalized XML structure. Check part additions/deletions, all changed parts, and relevant bookmarks, fields, relationships, styles, numbering, headers/footers, drawings and media. For a Surgical edit, confirm the exact intended replacement, target result, and unchanged non-target uncompressed parts.
2. Run `validate_docx_edit` with a task-specific allowlist and target checks. For Surgical, use `exact_replacements`; for other routes, use `allowed_xml_paths` and `expected_xml_values` only after inspecting actual paths. Do not approve an entire document part merely because it contains the target. Allow raw reserialization only when normalized structure and non-XML parts have been checked. The tool's PASS covers its declared structural checks, not visual or Word-native acceptance.
3. If layout could change—font size, table, image, section, header/footer, field, or floating object—render before and after with the same read-only renderer, fonts, settings and resolution; inspect every page and the target area. Prefer the official Documents render tools if appropriate. LibreOffice may export temporary PDF/PNG but must never resave the DOCX. Do not compare page counts from different renderers as a defect signal. A visually matching page does not prove structural identity.
4. If Word itself must verify behavior, open a verification copy without saving it. If an edit used Word, inspect the saved output again; a pre-save readback is insufficient. State separately what was structurally checked, visually checked, and actually reopened in Word.

If the target is ambiguous, the original value or relationship differs, the OOXML object is not understood, the semantic diff exceeds the allowed change, Word saving changes a protected structure, or an unrelated page changes, stop. Preserve the source; do not deliver the suspect copy as a successful result. Explain the observed difference and restart from the original if a different route is justified.

The tested examples do not prove that every Word save has the same effects, that visual equality means structural equality, or that a single changed part alone caused a layout difference.
