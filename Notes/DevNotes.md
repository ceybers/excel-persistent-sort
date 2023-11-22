# Development Notes for Excel Persistent Sort Order Tool
## Notes
- The state of the Sort Order for a ListObject is stored by storing each SortField as a comma separated string, and the list of SortField strings as a semicolon separated string. Properties of type String are encoded as Base64 to avoid having to deal with escaping characters.
- Colons are used to separate the Worksheet name, ListObject name, and Sort Order State, as colons are not present in any of the above.

## TreeView and Double Click
- `DblClick` event for TreeView fires when clicking anywhere on the control. 
- Therefore it does not return the clicked Node, like the `NodeClick` event does.
- This can cause non-intuitive behaviour if the user clicks once on an item, then tries to double click on a different icon but misclicks next to it.
- We could hijack the `Expand` event, but that feels hacky.
- Alternatively, the `MouseDown` event has x/y coordinates, but then we have to mess around with Hit Testing.
- Last solution (which I saw somewhere on the internet) was to record a timestamp when the `NodeClick` event fires, then compare it when the `DblClick` event fires.

## Sort On
### XlSortOn enumeration (Excel)
| Name            | Value | Description |
| --------------- | ----- | ----------- |
| SortOnCellColor | 1     | Cell color. |
| SortOnFontColor | 2     | Font color. |
| SortOnIcon      | 3     | Icon.       |
| SortOnValues    | 0     | Values.     |

### Sort by Icon
- Test: `SortField.SortOn = xlSortOnIcon`
- Order direction: `SortField.Order = xlAscending` is "On Top" in the UI
- Order value: `SortField.SortOnValue` property is `Object/Icon`. 
  - We need to deserialize it using `ActiveWorkbook.IconSets(4).Item(3)`. The former is from `.SortOnValue.Parent.ID`, the latter from `.SortOnValue.Index`.
```vb
.Sortfields.Add(Range("Table1[Gamma]"), _
    xlSortOnIcon, xlDescending, , _
    xlSortNormal).SetIcon Icon:=ActiveWorkbook.IconSets(4).Item(3)
```

### Sort by Font Color (foreground)
- Test: `SortField.SortOn = xlSortOnFontColor`
- Order as above
- Order value: `SortField.SortOnValue` property is `Object/Font`. 
```vb
.Sortfields.Add(Range("Table1[Gamma]"), 
    xlSortOnFontColor, xlAscending, , _
    xlSortNormal).SortOnValue.Color = RGB(156, 0, 6)
```

### Sort by Cell Color (background)
- Test: `SortField.SortOn = xlSortOnCellColor`
- Order as above
- Order value: `SortField.SortOnValue` property is `Object/Font`. 
```vb
.Sortfields.Add(Range("Table1[Gamma]"), 
    xlSortOnCellColor, xlAscending, , _
    xlSortNormal).SortOnValue.Color = RGB(255, 199, 206)
```

### Sort by Value (Custom Order)
```vb
.Sortfields.Add2 Key:=Range("Table1[ColB]"), _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    CustomOrder:="xray,yacht,zebra", _
    DataOption:=xlSortNormal
```

## 📖 API References
- [Sort.SortFields property (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.sort.sortfields)
- [XlSortOrder enumeration (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsortorder)
- [XlSortOn enumeration (Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsorton)