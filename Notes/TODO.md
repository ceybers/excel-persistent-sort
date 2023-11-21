# TODO
## Target ListObject
- [x] Show name of Target ListObject.
- [x] Show Current Sort Order, if any.
- [x] Allow saving Current Sort Order, if not already saved.
- [x] Check object references work OK with multiple workbooks.
## TreeView
- [x] Target ListObject should always be first on the tree.
- [x] Orphaned ListObject should always be last on the tree.
- [ ] Initially select first Sort Order State that can be applied.
- [x] Highlight Sort Order State if it is current already applied.
- [x] List Current Sort State in the Tree View so we can see the list of fields.
- [x] Apply by double clicking TreeView Node. 
- [x] Indicate when no Sort Order States were found to load (i.e., "No items found").
- [x] FIX Remove button not being disabled when selecting non-Sort Order State in tree view (e.g., ListObjects or Workbook).
- [ ] ~~FIX Logic when Apply via Tree View double click when mouse is over empty space.~~ Removed double click functionality from TreeView control.
## ListView
- [x] Show list of fields in a Sort Order State.
- [x] Indicate columns not present in target ListObject.
- [x] Indicate direction of sort order.
## Features
- [x] Apply selected Sort Order State.
- [ ] Prune orphaned Sort Order States.
- [x] Remove selected Sort Order State.
- [x] Remove all Sort Order States.
- [ ] Allow viewing Sort Order State in Base64 format.
- [ ] Allow adding Sort Order State via textbox in Base64 format.
## Options
- [x] Re-associate Sort Order State when applying an orphaned state to a ListObject.
- [x] Option to show/hide partial matches.
- [x] Option to allow/prohibit partial apply.
- [x] Option to automatically close after applying.
- [x] Default values if no CustomXML found.