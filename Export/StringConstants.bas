Attribute VB_Name = "StringConstants"
'@Folder "MVVM.Resources.Constants"
Option Explicit

Public Const APP_TITLE As String = "Excel Persistent Sort Order Tool"
Public Const APP_VERSION As String = "Version 1.2"
Public Const APP_COPYRIGHT As String = "2024 Craig Eybers" & vbCrLf & "All rights reserved."

Public Const ERR_UNKNOWN_ERROR As String = "Something went wrong with Persistent Sort Order Tool."

Public Const MSG_REMOVE_STATE As String = "Remove this Sort Order state?"
Public Const MSG_REMOVE_ALL_STATES As String = "Remove ALL Sort Order states?"
Public Const MSG_EXPORT_SORTORDER As String = "Sort Order State in Base64 format:"
Public Const MSG_SELECT_TABLE_FIRST As String = "Select a table before running Persistent Sort Order Tool."

Public Const SETTING_NAME As String = "PersistentSortOrder"
Public Const SETTING_COLLECTION_NAME As String = "SortOrderStates"
Public Const SETTING_LAST_UPDATED As String = "LastUpdated"
Public Const SETTING_ASSOCIATE_ON_APPLY As String = "DO_ASSOCIATE_ON_APPLY"
Public Const SETTING_PARTIAL_MATCH As String = "DO_PARTIAL_MATCH"
Public Const SETTING_PARTIAL_APPLY As String = "DO_PARTIAL_APPLY"
Public Const SETTING_CLOSE_ON_APPLY As String = "DO_CLOSE_ON_APPLY"

Public Const CAPTION_NO_SORT_ORDER As String = "(no sort order)"
Public Const CAPTION_NO_STATES_FOUND As String = "No saved Sort Order States found."
Public Const CAPTION_ORPHAN As String = "(Orphaned)"
Public Const CAPTION_UNSAVED_SORTORDER As String = "(current sort order)"
Public Const CAPTION_DO_APPLY As String = "Apply"
Public Const CAPTION_ALREADY_APPLIED As String = "Applied"
Public Const CAPTION_DO_SAVE As String = "Save"
Public Const CAPTION_ALREADY_SAVED As String = "Saved"

Public Const SUFFIX_ACTIVE  As String = " (active)"
Public Const SUFFIX_SELECTED As String = " (selected)"

Public Const COLUMN_INDEX As String = "#"
Public Const COLUMN_NAME As String = "Column Name"
Public Const COLUMN_DIRECTION As String = "Direction"

Public Const ITEM_DIRECTION_DESC As String = "Desc"
Public Const ITEM_DIRECTION_ASC As String = "Asc"
Public Const ITEM_CELL_COLOR As String = "Cell Color"
Public Const ITEM_FONT_COLOR As String = "Font Color"
Public Const ITEM_ICON As String = "Icon"
Public Const ITEM_VALUES As String = "Values"
Public Const ITEM_VALUES_CUSTOM As String = "Values (Custom)"

Public Const COLUMN_SORT_ON As String = "Sort On"

Public Const KEY_UNSAVED As String = "UNSAVED"
Public Const KEY_ROOT As String = "ROOT"

Public Const GREY_TEXT_COLOR As Long = 12632256 'RGB(192,192,192)
