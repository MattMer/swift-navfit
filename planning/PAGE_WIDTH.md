# Wider Front and Back page viewing area

The fixed 12000-twip host clipped the rightmost report fields and left unused window space. The host now expands to the available width, with a minimum 12600-twip viewing area and a 12200-twip paper extent. Horizontal scrolling remains available when a window is too narrow. Explicit resize widths now reach the fixed-page layout instead of being ignored.

Official field positions, sizes, typography and Front/Back grouping are unchanged. This changes the viewing area, not the report layout or print geometry. FITREP Front and Back were visually checked in full Access; all rightmost fields are visible at the current display configuration.

204 bounds/wide/narrow/fixed-geometry assertions, 264 fixed-page assertions and 29 pending-edit/page-switch assertions pass. Both installed modules align with source. The demo retains the two validation practice groups. Separate blank/demo manifests and exact hashes are in PAGE_WIDTH_VERIFICATION.json.

Updated demo: build/wide-pages-demo.accdb
Blank candidate: build/wide-pages.accdb

Earlier artifacts and populated workspaces are not overwritten. Other physical scaling configurations and printed fit are not claimed.
