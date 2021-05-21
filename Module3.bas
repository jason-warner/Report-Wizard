Attribute VB_Name = "Module3"
Sub configPvt2()
Attribute configPvt2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' configPvt2 Macro
'

'
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Project")
        .Orientation = xlRowField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Customer Vehicle ID")
        .Orientation = xlRowField
        .Position = 2
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Units"), "Sum of Units", xlSum
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Gross Cost"), "Sum of Gross Cost", xlSum
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Posted Date")
        .Orientation = xlPageField
        .Position = 1
    End With
    ActiveWindow.SmallScroll Down:=-18
    Application.CutCopyMode = False
    With ActiveSheet.PivotTables("PivotTable2")
        .ColumnGrand = False
        .RowGrand = False
        .InGridDropZones = True
        .DisplayFieldCaptions = False
        .DisplayContextTooltips = False
        .ShowDrillIndicators = False
        .RowAxisLayout xlTabularRow
    End With
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Project").Subtotals = Array _
        (False, False, False, False, False, False, False, False, False, False, False, False)
    ActiveSheet.PivotTables("PivotTable2").PivotFields("Project").RepeatLabels = _
        True
End Sub
