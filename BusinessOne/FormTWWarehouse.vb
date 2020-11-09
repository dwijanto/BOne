Imports System.Threading
Imports Microsoft.Office.Interop
Public Class FormTWWarehouse
    Dim selectedDate As Date
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Private myAdapter As New TWWarehouseAdapter
    Dim BS As BindingSource

    Public Property errMsg As String
        Get
            Return myAdapter.errorMsg
        End Get
        Set(value As String)
            myAdapter.errorMsg = value
        End Set
    End Property

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim mydialog As New DialogDate
        If mydialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            selectedDate = mydialog.mydate
            loaddata()
        End If
    End Sub

    Private Sub loaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoWork()
        ProgressReport(1, "Loading data. Please wait...")
        ProgressReport(6, "Marquee")

        BS = New BindingSource
        myAdapter = New TWWarehouseAdapter
        If myAdapter.LoadData(selectedDate) Then
            ProgressReport(4, "Fill Data..")
            ProgressReport(1, "Done.")
        Else
            ProgressReport(1, String.Format("Has error::{0}", myAdapter.errorMsg))
        End If
        ProgressReport(5, "Continuous")
    End Sub
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    ToolStripStatusLabel1.Text = message
                Case 4
                    DataGridView1.AutoGenerateColumns = True

                    BS = New BindingSource
                    BS.DataSource = myAdapter.DS.Tables(0)
                    DataGridView1.DataSource = BS
                    ToolStripStatusLabel2.Text = String.Format("Record(s) count: {0:#,##0}", BS.Count)
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 7

            End Select
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter.DS) Then
            myAdapter.SelectedDate = selectedDate
            If myAdapter.ExportTextFile() Then
                CreateExcel()
            Else
                ProgressReport(1, String.Format("{0}", myAdapter.errorMsg))
            End If
        Else
            MessageBox.Show("Refresh Data first.")
        End If
    End Sub

    Private Sub CreateExcel()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoExcelWork)
            myThread.TrySetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoExcelWork()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Loading Data.")


        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("{0}\Warehouse\TW\Excel\{1}_{2:yyyyMMdd}.xlsx", myAdapter.OutputFolder, "TWWarehouse", myAdapter.SelectedDate)

        'If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
        Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
        Dim reportname = IO.Path.GetFileName(mysaveform.FileName)
        Dim errormessage As String = String.Empty

        Dim datasheet As Integer = 3

        Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
        Dim sqlstr As String

        sqlstr = myAdapter.GetSQLReport
        Dim myreport As New ExcelExtract(Me, mysaveform.FileName, datasheet, "\templates\warehouseTW.xltx", mycallback, PivotCallback)
        myreport.GenerateReport(mysaveform.FileName, sqlstr, errormessage)
        ' End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

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
        Threading.Thread.Sleep(1000)
        If osheet.cells(3, 2).text.ToString = "" Then
            Err.Raise(100, Description:="Data not available.")
        End If

        osheet.name = "RawData"


        owb.Names.Add("db", RefersToR1C1:="=OFFSET('RawData'!R1C1,0,0,COUNTA('RawData'!C1),COUNTA('RawData'!R1))")
        Threading.Thread.Sleep(100)

        owb.Worksheets(2).select()
        osheet = owb.Worksheets(2)
        'MessageBox.Show("change pivot cache")
        osheet.PivotTables("PivotTable1").ChangePivotCache(owb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, SourceData:="db"))
        'oXl.Run("ShowFG")
        Thread.Sleep(100)
        '

        'MessageBox.Show("refresh1")
        osheet.PivotTables("PivotTable1").PivotCache.Refresh()
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refresh2")
        osheet.pivottables("PivotTable1").SaveData = True
        Threading.Thread.Sleep(100)
        'MessageBox.Show("refreshAll")
        owb.RefreshAll()
        Thread.Sleep(100)
        osheet.Cells.EntireColumn.AutoFit()
        Thread.Sleep(100)
        'Threading.Thread.Sleep(100)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        Dim mydialog As New DialogDate
        If mydialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            myAdapter.SelectedDate = mydialog.mydate
            GenerateExcel()
        End If
    End Sub

    Private Sub GenerateExcel()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoGenerate)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoGenerate()

        If myAdapter.GenerateFromTextFile() Then
            DoExcelWork()
        Else
            ProgressReport(1, String.Format("{0}", myAdapter.errorMsg))
        End If

    End Sub
End Class