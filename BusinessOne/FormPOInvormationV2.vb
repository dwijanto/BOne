Imports System.Threading
Imports Microsoft.Office.Interop
Public Class FormPOInvormationV2
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim MyLocation As LocationEnum
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim BS As BindingSource
    Dim myAdapter As Object

    Private startdate As Date
    Private enddate As Date

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub FormOrderItem_Load(sender As Object, e As EventArgs) Handles Me.Load
        'loaddata()
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


    Sub DoWork()
        ProgressReport(1, "Loading data. Please wait...")
        ProgressReport(6, "Marquee")
        BS = New BindingSource

        If MyLocation = LocationEnum.Hong_Kong Then
            myAdapter = New POAdapter
        Else
            myAdapter = New POAdapterTW
        End If

        If myAdapter.loadDataV2(startdate, enddate) Then
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
                    BS.DataSource = myAdapter.DS.Tables(0)
                    DataGridView1.DataSource = BS
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim myform = New DialogDateRange
        myform.GroupBox1.Visible = True
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            startdate = myform.startdate
            enddate = myform.enddate
            MyLocation = myform.MyLocation
            loaddata()
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter) Then
            CreateExcel()
        Else
            MessageBox.Show("Please refresh the data first.")
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
        mysaveform.FileName = String.Format("{0}Report{1:yyyyMMdd}.xlsx", "POInformation", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim myTable = myAdapter.DS.Tables(0).Copy
            Dim myreport As New ExcelExtract(Me, mysaveform.FileName, "\templates\ExcelTemplate01.xltx", myTable, mycallback, PivotCallback)
            myreport.ExtractFromDataTableUnsyncDT(Me, New System.EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub FormattingReport(ByRef osheet As Excel.Worksheet, ByRef e As EventArgs)
        'osheet.Columns("AD:AF").delete()
    End Sub

    Private Sub PivotTable()

    End Sub
End Class