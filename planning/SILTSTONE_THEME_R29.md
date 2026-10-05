# Siltstone theme - r29

The native Access forms use the supplied palette:

| Role | Color |
|---|---|
| Field borders | Khaki #CBBD93 |
| Section/navigation buttons | Pale yellow #FFF5B8 |
| General action buttons | Peach #FFB16E |
| Selected section and primary actions | Ochre #CCA25A |
| Canvas | Warm light #F8F0EA |
| Field surfaces | White #FFFFFF |
| Text | Dark #262C26 |

All 51 forms receive the final shared presentation pass, including supporting
import, review, recovery and settings screens. Existing control names, event
handlers, layout and data services are retained. Section changes retain the
palette at runtime. Validation errors retain the existing dark red treatment
and text messages. Native checkbox glyphs, focus outlines, disabled controls,
Access menus and selected list rows follow Windows/Access rendering.

The shared pass is scripts/apply-ux-theme.ps1, run last by scripts/build.ps1.
The upgrade helper scripts/build-ux-theme.ps1 copies an existing workspace to
a new output and imports the two updated runtime modules before applying colors.

Demo: output/ux/Swift_NAVFIT_UI_Demo_2026-10-05-r29.accdb
Blank: build/ux-r29-final-blank.accdb

Verification: both builds compile. All 57 installed demo modules match source;
native editor/review/list events remain intact. This change does not add a schema
migration. Native visual and screen-reader review of the new colors is pending.
