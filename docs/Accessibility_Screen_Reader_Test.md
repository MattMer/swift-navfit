# Screen-reader acceptance test (NVDA and JAWS)

This is the attended test for Section 508 / WCAG 2.0 AA criteria that automated checks can't settle:
- 302.1 (without vision)
- 2.1.1 and 2.1.2 (keyboard, no trap)
- 2.4.3 and 2.4.7 (focus order and visibility)
- 3.2.1 (on focus)
- 4.1.2 (name, role, value)

`scripts/test-accessibility.ps1` already checks contrast, label links and button names.

Run it on a copy of a **Demo** build. Never use real personnel data. Record the results in the tables and keep the completed copy with the release evidence.

## 1. Set up

1. **Get a screen reader.**
   - **NVDA** is free from nvaccess.org.
   - **JAWS** runs in 40-minute demo sessions without a license. DoD users can request a licensed copy through the Computer/Electronic Accommodations Program (CAP).
   - On a Navy-managed workstation, use the installed copy or ask the ISSM. Don't install software yourself.
2. **Make a test copy.** Copy the Demo `.accdb` to a scratch folder and open the copy in desktop Access. Choose **Enable Content**.
3. **Turn on speech logging,** so you can copy exactly what was announced:
   - **NVDA:** press **Insert+N**, then choose **Tools > Speech viewer**.
   - **JAWS:** press **Insert+Space**, then **H**, for speech history.
4. **Use the keyboard only.** Put the mouse aside. Every step below uses the keyboard.

### Keys used in this test

| Purpose | NVDA | JAWS |
| --- | --- | --- |
| Start | Ctrl+Alt+N | Desktop shortcut |
| Say the focused control again | Insert+Tab | Insert+Tab |
| Say the window title | Insert+T | Insert+T |
| Read the whole window | Insert+B | Insert+B |
| Stop speaking | Ctrl | Ctrl |
| Quit | Insert+Q | Insert+F4 |

In NVDA's laptop layout, Caps Lock replaces Insert.

## 2. What to listen for on every focused control

- **Name:** the visible label or caption. Where there is no visible label, a short name such as "Reports", "Trait 33" or "E5 RSCA".
- **Role:** for example "edit", "combo box", "list", "button" or "check box".
- **Value or state:** the text or selection, or checked/unchecked. Also listen for "unavailable" on disabled buttons.

On every screen, also check the following:
- **Focus visible:** you can see where focus is.
- **No trap:** Tab and Shift+Tab never get stuck.
- **On focus:** tabbing into a field never opens a dialog or changes the screen by itself.

## 3. Walk-through

Run each step with NVDA, then repeat with JAWS. Under **Heard**, write what was actually announced; a short version is fine. Mark **P** (pass) or **F** (fail).

| # | Screen and action | Expected announcement | NVDA heard | P/F | JAWS heard | P/F |
| --- | --- | --- | --- | --- | --- | --- |
| 1 | Startup: tab to the contributor choice | The field's caption, its role, then each contributor as you arrow | | | | |
| 2 | Overview: Tab through the left navigation | "Overview", "Reports", "Summary groups" and the other destinations, each a button | | | | |
| 3 | Reports: Tab to the search box | "Search reports", edit | | | | |
| 4 | Reports: Tab through the filters | "Report type", "Local review status", "Rank / rate", "Archive status" and "Column list", each a combo box | | | | |
| 5 | Reports: Tab to the list and arrow through the rows | "Reports", list, then each row's text | | | | |
| 6 | Open a report with its open button | The window title names the report editor | | | | |
| 7 | Front Page: Tab from block 1 to block 20 | Each official block caption, such as "1. Name (Last, First MI Suffix)" and "2. Grade/Rate", in NAVPERS order | | | | |
| 8 | Front Page: Tab through the traits | "33. Professional Knowledge" through "39", each a combo box, with the selected mark | | | | |
| 9 | Change one trait and press **Save draft** | The status text changes to a saved message (Insert+B reads it) | | | | |
| 10 | Report tools > Validation checks | "Validation findings", list; then "Finding details" | | | | |
| 11 | Report tools > Notes: type a message, then Tab to **Post** | "Write a new report note", edit; then "Post", button. The log shows the message | | | | |
| 12 | Back Page: Tab through blocks 40 to 47 | Each block caption, in order | | | | |
| 13 | **Mark signed**: open it and Tab through the dialog | Each field's caption, and the buttons by caption | | | | |
| 14 | On the signed report, type into a field | A prompt asking whether to reopen the report, read in full | | | | |
| 15 | Summary groups: Tab to the list | "Summary groups", list, with each group's text | | | | |
| 16 | Open a group's trait table and Tab across one row | "Member", "Trait 33" through "Trait 39", "ITA", "Promotion recommendation", and so on | | | | |
| 17 | Open the RSCA table and Tab through one row | "E5 RSCA", "E5 previous graded traits", "E5 as of date" | | | | |
| 18 | My Work: switch views and Tab to each list | "My work", "Midterm counseling", "Contributors", each a list | | | | |
| 19 | Members: record counseling and Tab through the dialog | "Outcome", "Counseling date", "Reason not performed (block 31)" and the other fields | | | | |
| 20 | Settings & Help: open help with the **?** button | Help topics, then the topic text | | | | |
| 21 | Close each dialog with its Close or Cancel button | Focus returns to the screen that opened it | | | | |

## 4. Contrast theme and magnifier

1. Open **Windows Settings > Accessibility > Contrast themes**, pick a theme and apply it. Reopen Swift's screens and check the following:

   | Check | P/F |
   | --- | --- |
   | Text, borders and buttons use the theme's colors | |
   | Nothing is unreadable | |
   | The current navigation item is still identifiable (bold) | |

2. Turn the theme off, then zoom with Magnifier (**Windows key + Plus**) to 200%. Check that you can read the report pages and the trait table: P/F ___

## 5. Record failures

For each **F**, record:
- the step number,
- the screen and field (or the closest visible caption),
- what was announced,
- what was expected,
- the screen reader and its version, and the Access version.

Swift's maintainers turn each failure into a fix and an automated check where one is possible.

| Tester | Date | NVDA version | JAWS version | Access version | Build |
| --- | --- | --- | --- | --- | --- |
| | | | | | |
