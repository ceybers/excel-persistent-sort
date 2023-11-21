# TODO
## Target ListObject
- [x] Show name of Target ListObject.
- [x] Show Current Sort Order, if any.
- [x] Allow saving Current Sort Order, if not already saved.
- [ ] Check object references work OK with multiple workbooks.
## TreeView
- [ ] Target ListObject should always be first on the tree.
- [ ] Orphaned ListObject should always be last on the tree.
- [ ] Initially select first Sort Order State that can be applied.
- [x] Highlight Sort Order State if it is current already applied.
- [ ] List Current Sort State in the Tree View so we can see the list of fields.
- [x] Apply by double clicking TreeView Node. 
- [x] Indicate when no Sort Order States were found to load (i.e., "No items found").
- [ ] FIX Remove button not being disabled when selecting non-Sort Order State in tree view (e.g., ListObjects or Workbook).
- [ ] FIX Logic when Apply via Tree View double click when mouse is over empty space.
## ListView
- [x] Show list of fields in a Sort Order State.
- [x] Indicate columns not present in target ListObject .
- [ ] Indicate direction of sort order.
## Features
- [x] Apply selected Sort Order State.
- [ ] Prune orphaned Sort Order States.
- [ ] Remove selected Sort Order State.
- [x] Remove all Sort Order States.
## Options
- [ ] Re-associate Sort Order State when applying it to a ListObject.
- [ ] Option to show/hide partial matches.
- [ ] Option to allow/prohibit partial apply.
- [x] Option to automatically close after applying.