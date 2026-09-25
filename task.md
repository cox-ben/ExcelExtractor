Current Issues:

- [x] Grabbed Quote Item # from Tool # instead of Quote number.
  - Prioritized `Quote Number` / `MPC RFQ #` over generic `Item #` and strictly excluded columns with `tool`, `tooling`, `mold #`, `die #`.
  - Added regex extraction to isolate canonical quote numbers (e.g. `Q4378-4` from `Q4378-4 (2140)`).

- [x] Grabbed part numbers 90% correct except for two. It's because this sheet got very complicated and many sheets said MPC or QInfo in it. From now on grab information from the sheet that matches MPC the closest, use the sheet that matches QInfo the closest as validation, disregard all other sheets.
  - Implemented closest MPC sheet ranking (`score_sheet_title_mpc`) to extract primarily from the closest MPC sheet and disregard auxiliary sheets.
  - Implemented closest QInfo sheet validation (`score_sheet_title_qinfo`) to cross-reference and heal part numbers without pulling in external ghost items.

- [x] Description was called Desc so it skipped over it and grabbed resin information because the cell said "Generic Resin Description" which contains "Description" in it. Make it so if it finds a cell with discription it ensure it doesn't have any key words that would make it something else eg. resin discription. We want the part discription, in this case it was called Desc. 
  - Added `Desc` / `desc` to description keywords.
  - Added strict exclusion filter (`resin`, `material`, `generic`, `tool`, `machine`, `supplier`, `grade`) to ensure resin or generic descriptions are never picked as part descriptions.

- [x] It grabbed the resin name from resin supplier instead of "Resin Grade". We want Resin Grade not Generic Resin Discription or Resin Supplier
  - Prioritized `Resin Grade` / `Material Grade` (Priority 1) over generic material names.
  - Strictly banned `Resin Supplier`, `Supplier`, `Vendor`, and `Generic Resin Description`.

- [x] He couldn't find labour rate because that is only found in MPC and not in QInfo. 
  - Because extraction now operates primarily on the closest MPC sheet, `find_labour_rate` locates `Rate $CAD/Hr` directly on MPC.
  - Added cross-sheet fallback to scan visible MPC sheets if missing on the active tab, and improved spatial detection directly below `Rate $CAD/Hr`.

- [x] Failed on sheet with no QInfo sheet and MPC column headers starting at line 31, followed by a blank line, a column index row (1, 2, 3...), and data rows.
  - Increased header scan depth to scan top 60 rows (`max(header_max_row, 60)`).
  - Implemented multi-criteria candidate header scoring with lookahead data verification so top metadata rows (rows 1–25) are not falsely selected as table headers.
  - Updated `is_column_index_row` to recognize sequential integer series (`1, 2, 3...` or `2, 3, 4, 5...`), floats (`1.0, 2.0`), string representations, and optional leading labels.
  - Handled spacer rows between header and data rows without premature table termination.
  - Updated `clean_part_number_string` to reject single-digit integers (`1` through `9`).
