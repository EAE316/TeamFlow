# TeamFlow 1.4.2 — preview

A modern look, new features, and SharePoint as the main way of working.
Test with a copy of your workbook first.

## New look
- **Soft app shell:** a tinted frame, a raised work area and cards with soft shadows
  instead of grey outlines, with one indigo colour for actions and selection. Light is the
  default; dark mode, chosen in the sidebar, has its own matching colours.
- **A colour for each person**, the same on the board, timeline, calendar, table and
  dashboard. The order of the colours is checked to stay distinguishable for colour-blind
  people, and the name is always shown beside them.
- **Calmer top bar:** search with a Ctrl+K quick-commands hint, icon buttons for the
  report, sharing and refresh, and one main "New task" button. The data source moved to
  the bottom of the sidebar.
- **Friendlier cards:** dates as words ("מחר", "באיחור 3 ימים"), a progress ring, counts
  of checklist items, comments and links, and the owner's avatar. A card lifts on hover;
  the whole card opens the task, and "סיום" appears on hover.
- **Task panel as a task page:** a large title you edit in place, properties as compact
  rows, and Details / Checklist / Activity tabs. The notes box shows only the notes;
  checklist items and links have their own places.
- **Richer dashboard:** a completion ring, tasks completed per day over the last two
  weeks, and workload as one bar per person (click a person to see their tasks).
- **Motion:** smooth screen changes, loading placeholders, a moved card briefly
  highlighted, and finished tasks fading out.
- **Hebrew labels throughout** (for example "לוח משימות" instead of "לוח Kanban") and a
  friendlier My day screen when there is nothing to do.
- **Dates as dd/mm/yyyy** everywhere: the task panel (type a date, or pick it from the
  calendar button), table, cards, history, comments and both reports. The Excel report
  holds real dates shown day first.
- **Printing the weekly report works again:** the printed page was blank in 1.3.

Nothing in the data or the SharePoint code changed for this redesign.

## SharePoint mode
- **Create and delete tasks.** New tasks get a random ID such as T-7KQ2M, so two
  PCs can never pick the same one, and go into the first empty row. If another PC
  creates a task into the same row at the same moment, TeamFlow adds yours again.
- **Deleted tasks leave their row in place.** The row is emptied and its ID marked
  deleted in column T, so no other task's row moves on anyone's PC.
- **Bulk actions and drag-to-person** work, one checked update per task. Tasks that
  changed in the meantime are skipped and reported.
- **Everything is done in the app**, with no need to edit the workbook by hand:
  - **Team lists** (people, categories, statuses, priorities) are edited in Settings.
    Only the cells that changed are written, and if someone else changed the lists
    first, TeamFlow asks you to reload instead of overwriting their change.
  - **Manual card order:** drag a card to a new place. Moving one card writes one
    order number (column U), and no rows move.
- **Personal OneDrive for testing:** an address starting with https://d.docs.live.net
  is accepted, so co-editing can be tried without a company account.

## Also new since 1.2
- **Dark mode:** choose it with the moon button at the bottom of the sidebar. Light is the default.
- **My day** (start page): your overdue tasks and those due today, this week and later.
- **Reminders:** a Windows notification for your overdue and due tasks, once per task
  per day, while TeamFlow is open. Turn them on from My day.
- **Calendar:** tasks on their due day; drag a task to another day to move its date.
- **Comments** on tasks, kept in the task's own history cell.
- **Links** on tasks (SharePoint documents, websites, email), kept as notes lines
  starting with 🔗 so they need no extra column.
- **Weekly report:** print or save as PDF, or download as Excel.

## Testing the SharePoint code on a PC without SharePoint
Close TeamFlow, then start `TeamTaskManagerWeb.exe --excel-test` from the install folder
(for example from a shortcut with that option). Choosing a local file then opens it through desktop Excel, using
exactly the code SharePoint mode uses. Sign-in and two-person editing are not covered.

## What was tested
- 72 automated backend tests, 12 of them driving real desktop Excel on local copies.
- 23 browser checks against a disposable workbook, in the built app's frontend.
- Not tested: real SharePoint, real OneDrive, and two PCs editing at once.

## Going back
Earlier versions stay on the releases page. In the source, git tag v1.3.1 is the
previous look with the same features, and v1.2.1 the version before 1.3.
