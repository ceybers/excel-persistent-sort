# Sort by Icon
- Test: `SortField.SortOn = xlSortOnIcon`
- Order direction: `SortField.Order = xlAscending` is "On Top" in the UI
- Order value: `SortField.SortOnValue` property is `Object/Icon`. 
  - We need to deserialize it using `ActiveWorkbook.IconSets(4).Item(3)`. The former is from `.SortOnValue.Parent.ID`, the latter from `.SortOnValue.Index`.
```vb
.Sortfields.Add(Range("Table1[Gamma]"), _
    xlSortOnIcon, xlDescending, , _
    xlSortNormal).SetIcon Icon:=ActiveWorkbook.IconSets(4).Item(3)
```

# Sort by Font Color (foreground)
- Test: `SortField.SortOn = xlSortOnFontColor`
- Order as above
- Order value: `SortField.SortOnValue` property is `Object/Font`. 
```vb
.Sortfields.Add(Range("Table1[Gamma]"), 
    xlSortOnFontColor, xlAscending, , _
    xlSortNormal).SortOnValue.Color = RGB(156, 0, 6)
```

# Sort by Cell Color (background)
- Test: `SortField.SortOn = xlSortOnCellColor`
- Order as above
- Order value: `SortField.SortOnValue` property is `Object/Font`. 
```vb
.Sortfields.Add(Range("Table1[Gamma]"), 
    xlSortOnCellColor, xlAscending, , _
    xlSortNormal).SortOnValue.Color = RGB(255, 199, 206)
```

# Sort by Value (Custom Order)
```vb
.Sortfields.Add2 Key:=Range("Table1[ColB]"), _
    SortOn:=xlSortOnValues, _
    Order:=xlAscending, _
    CustomOrder:="xray,yacht,zebra", _
    DataOption:=xlSortNormal
```

# XlSortOn enumeration (Excel)
| Name            | Value | Description |
| --------------- | ----- | ----------- |
| SortOnCellColor | 1     | Cell color. |
| SortOnFontColor | 2     | Font color. |
| SortOnIcon      | 3     | Icon.       |
| SortOnValues    | 0     | Values.     |
