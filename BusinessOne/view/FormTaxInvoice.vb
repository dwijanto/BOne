Imports System.Threading
Imports System.Text
Imports System.IO

Public Class FormTaxInvoice
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim MyLocation As LocationEnum
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim BS As BindingSource
    Dim BS2 As BindingSource
    Dim myAdapter As POAdapterTW

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
        BS2 = New BindingSource

        myAdapter = New POAdapterTW

        If myAdapter.loadDataTaxInvoice(startdate, enddate) Then
            ProgressReport(4, "Fill Data..")
            ProgressReport(1, String.Format("Done. Record(s) Count : {0}", BS.Count))
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
                    BS2.DataSource = myAdapter.DS.Tables(1)
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
        myform.GroupBox1.Visible = False
        myform.Label1.Text = "Invoice Start Date"
        myform.Label2.Text = "Invoice End Date"
        If myform.ShowDialog = Windows.Forms.DialogResult.OK Then
            startdate = myform.startdate
            enddate = myform.enddate
            MyLocation = myform.MyLocation
            loaddata()
        End If
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If Not IsNothing(myAdapter) Then
            CreateText()
        Else
            MessageBox.Show("Please refresh the data first.")
        End If
    End Sub

    Private Sub CreateText()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoTXT)
            myThread.TrySetApartmentState(ApartmentState.STA)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub DoTXT()
        ProgressReport(6, "Marquee")
        ProgressReport(1, "Preparing Text File.")

        Dim sb As New StringBuilder

        Dim mysaveform As New SaveFileDialog
        mysaveform.FileName = String.Format("{0}Report{1:yyyyMMdd}.txt", "TaxInvoice", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)
            Dim drv = BS2.Current
            'Create H Line
            sb.Append(String.Format("H,80030293,{0},{1},{2}{3}", drv.row.item("CompnyName"), drv.row.item("CompnyAddr"), drv.row.item("Phone1"), vbCrLf))

            Dim Check As String = String.Empty
            Dim CustRef As Boolean
            For Each drv2 As DataRowView In BS.List
                If Check <> drv2.Row.Item("invoice no") Then
                    'Print M
                    sb.Append(String.Format("M,{0},{1:yyyy/MM/dd},07,{2},{3},{4},1,5,{5:0},{6:0},{7:0},,,{8}",
                                            drv2.Row.Item("invoice no"),
                                            drv2.Row.Item("invoice date"),
                                            drv2.Row.Item("invoice ban"),
                                            drv2.Row.Item("invoice name"),
                                            drv2.Row.Item("address"),
                                            drv2.Row.Item("sales amount"),
                                            drv2.Row.Item("vatamount"),
                                            drv2.Row.Item("total"),
                                            vbCrLf))

                    Check = drv2.Row.Item("invoice no")
                    CustRef = True
                End If
                'Print Detail
                If CustRef Then
                    'Print with custref
                    sb.Append(String.Format("D,{0} {1},{2:0},{3:0},{4:0},{5}{6}", drv2.Row.Item("articlecode"), drv2.Row.Item("description"), drv2.Row.Item("quantity"), drv2.Row.Item("price"), drv2.Row.Item("amount"), drv2.Row.Item("custref"), vbCrLf))
                    CustRef = False
                Else
                    'Print without custref
                    sb.Append(String.Format("D,{0} {1},{2:0},{3:0},{4:0},{5}", drv2.Row.Item("articlecode"), drv2.Row.Item("description"), drv2.Row.Item("quantity"), drv2.Row.Item("price"), drv2.Row.Item("amount"), vbCrLf))
                End If
            Next
            Using mystream As New StreamWriter(mysaveform.FileName)
                mystream.WriteLine(sb.ToString)
            End Using
            Process.Start(mysaveform.FileName)
        End If

        ProgressReport(1, "Done!")
        ProgressReport(5, "Continuous")
    End Sub



End Class