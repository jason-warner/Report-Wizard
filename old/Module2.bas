Attribute VB_Name = "Module2"
Sub setPivot()
Attribute setPivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' setPivot Macro
'

'
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Project")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Asset")
        .Orientation = xlRowField
        .Position = 2
    End With
    With ActiveSheet.PivotTables("PivotTable1")
        .ColumnGrand = False
        .InGridDropZones = True
        .DisplayFieldCaptions = False
        .RowAxisLayout xlTabularRow
    End With
    With ActiveSheet.PivotTables("PivotTable1")
        .DisplayContextTooltips = False
        .ShowDrillIndicators = False
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Project").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Project").RepeatLabels = _
        True
End Sub
Sub copyPTableData()
Attribute copyPTableData.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copyPTableData Macro
'

'
    Range("A3").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
End Sub
