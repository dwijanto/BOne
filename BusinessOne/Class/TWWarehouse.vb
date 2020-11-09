﻿Imports Microsoft.Office.Interop
Public Class TWWarehouse
    Dim myadapter As New TWWarehouseAdapter
    Dim parent As Object
    Public ReadOnly Property GetDataSet As DataSet
        Get
            Return myadapter.DS
        End Get
    End Property
    Public Property errorMsg As String

    Public Sub New()

    End Sub
    Public Sub New(parent As Object)
        Me.parent = parent
    End Sub

    Public Sub LoadData(ByVal selecteddate As Date)
        myadapter.LoadData(selecteddate)
    End Sub
    Public Sub ExportTextFile()
        myadapter.ExportTextFile()
    End Sub

    Public Function DoExcelWork() As Boolean
        'ProgressReport(6, "Marquee")
        'ProgressReport(1, "Loading Data.")

        Dim myret As Boolean
        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("{0}\Warehouse\TW\Excel\{1}_{2:yyyyMMdd}.xlsx", myadapter.OutputFolder, "TWWarehouse", myadapter.SelectedDate)

        'If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
        Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
        Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

        Dim datasheet As Integer = 3

        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
        Dim sqlstr As String

        sqlstr = myadapter.GetSQLReport
        Dim myreport As New ExcelExtract(Me.parent, mysaveform.FileName, datasheet, "\templates\warehousetw.xltx", mycallback, PivotCallback)
        myret = myreport.GenerateReport(mysaveform.FileName, sqlstr, errorMsg)
        ' End If

        'ProgressReport(1, "Loading Data.Done!")
        'ProgressReport(5, "Continuous")
        Return myret
    End Function

    Private Sub FormattingReport(ByRef osheet As Excel.Worksheet, ByRef e As EventArgs)
        'Create DBRange

    End Sub

    Private Sub PivotTable(ByRef sender As Object, ByRef e As EventArgs)
        'Change PivotTable source
        Dim oXl As Excel.Application = Nothing
        Dim owb As Excel.Workbook = CType(sender, Excel.Workbook)
        oXl = owb.Parent
        owb.Worksheets(3).select()
        Dim osheet = owb.Worksheets(3)
        Dim orange = osheet.Range("A2")

        If osheet.cells(3, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If

        osheet.name = "RawData"


        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")
        Threading.Thread.Sleep(500)
        owb.Worksheets(2).select()
        osheet = owb.Worksheets(2)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")

        '
        Threading.Thread.Sleep(500)
        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(500)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        Threading.Thread.Sleep(500)
        'MessageBox.Show("refreshAll")
        owb.RefreshAll()
        Threading.Thread.Sleep(500)
        osheet.Cells.EntireColumn.AutoFit()
        'Threading.Thread.Sleep(100)
        Threading.Thread.Sleep(500)
    End Sub
End Class
