Attribute VB_Name = "Module1"
Sub testing1()
Attribute testing1.VB_Description = "testing the format to report code"
Attribute testing1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' testing1 Macro
' testing the format to report code
'

'
    Cells.Select
    Sheets.Add
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Crew List (5)!R1C1:R1048576C5", Version:=7).CreatePivotTable _
        TableDestination:="Sheet2!R3C1", TableName:="PivotTable1", DefaultVersion _
        :=7
    Sheets("Sheet2").Select
    Cells(3, 1).Select
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = True
        .HasAutoFormat = True
        .DisplayErrorString = False
        .DisplayNullString = True
        .EnableDrilldown = True
        .ErrorString = ""
        .MergeLabels = False
        .NullString = ""
        .PageFieldOrder = 2
        .PageFieldWrapCount = 0
        .PreserveFormatting = True
        .RowGrand = True
        .SaveData = True
        .PrintTitles = False
        .RepeatItemsOnEachPrintedPage = True
        .TotalsAnnotation = False
        .CompactRowIndent = 1
        .InGridDropZones = False
        .DisplayFieldCaptions = True
        .DisplayMemberPropertyTooltips = False
        .DisplayContextTooltips = True
        .ShowDrillIndicators = True
        .PrintDrillIndicators = False
        .AllowMultipleFilters = False
        .SortUsingCustomLists = True
        .FieldListSortAscending = False
        .ShowValuesRow = False
        .CalculatedMembersInFilters = False
        .RowAxisLayout xlCompactRow
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotCache
        .RefreshOnFileOpen = False
        .MissingItemsLimit = xlMissingItemsDefault
    End With
    ActiveSheet.PivotTables("PivotTable1").RepeatAllLabels xlRepeatLabels
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Project")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Asset")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1")
        .InGridDropZones = True
        .DisplayFieldCaptions = False
        .DisplayContextTooltips = False
        .ShowDrillIndicators = False
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Project").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Project").RepeatLabels = _
        True
End Sub
Sub tableSelect()
Attribute tableSelect.VB_Description = "start at a1 of tablet"
Attribute tableSelect.VB_ProcData.VB_Invoke_Func = " \n14"
'
' tableSelect Macro
' start at a1 of tablet
'

'
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
End Sub
