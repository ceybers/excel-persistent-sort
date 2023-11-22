# TODO
## Target ListObject
- [x] Show name of Target ListObject.
- [x] Show Current Sort Order, if any.
- [x] Allow saving Current Sort Order, if not already saved.
- [x] Check object references work OK with multiple workbooks.

## TreeView
- [x] Target ListObject should always be first on the tree.
- [x] Orphaned ListObject should always be last on the tree.
- [x] Initially select first Sort Order State that can be applied.
- [x] Highlight Sort Order State if it is current already applied.
- [x] List Current Sort State in the Tree View so we can see the list of fields.
- [x] Indicate when no Sort Order States were found to load (i.e., "No items found").
- [x] FIX Remove button not being disabled when selecting non-Sort Order State in tree view (e.g., ListObjects or Workbook).
- [ ] ~~Apply by double clicking TreeView Node.~~
- [ ] ~~FIX Logic when Apply via Tree View double click when mouse is over empty space.~~ 
  - Removed double click functionality from TreeView control.

## ListView
- [x] Show list of fields in a Sort Order State.
- [x] Indicate columns not present in target ListObject.
- [x] Indicate direction of sort order.

## Options
- [x] Re-associate Sort Order State when applying an orphaned state to a ListObject.
- [x] Option to show/hide partial matches.
- [x] Option to allow/prohibit partial apply.
- [x] Option to automatically close after applying.
- [x] Default values if no CustomXML found.

## Features
- [x] Apply selected Sort Order State.
- [x] Prune orphaned Sort Order States.
- [x] Remove selected Sort Order State.
- [x] Remove all Sort Order States.
- [x] Allow viewing Sort Order State in Base64 format.
- [x] Allow manually adding Sort Order State via textbox in Base64 format.
- [x] Sort on Cell Color, Font Color, and Icon.
- [x] Indicate Sort On type with text and icon.
- [x] Sort Order using Custom List.
- [x] Remap Columns from Sort Order State to existing Columns in target. 
- [ ] View and edit items in Custom List (multiline text box).
- [ ] Gracefully handle malformed/deprecated persistent states (e.g., change in schema)
- [ ] Ribbon button to store current state without prompts.
- [ ] Ribbon button to attempt to restore saved state without prompts, if only one good match.
- [ ] FIX Clicking on column to remap conflicts with having to select a column to Export. 
