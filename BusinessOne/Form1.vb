'Imports System.Security
Imports System.Text
Imports System.Threading

Public Class Form1
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim BS As BindingSource
    Dim myAdapter As OITMAdapter

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        loaddata()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()       
        ' Add any initialization after the InitializeComponent() call.

    End Sub

    Sub DoWork()
        ProgressReport(1, "Loading data. Please wait...")
        ProgressReport(6, "Marquee")
        Dim sqlstr = "select itemcode,itemname from OITM;"        
        bs = New BindingSource
        myAdapter = New OITMAdapter()
        'myAdapter = New OITMAdapter
        If myAdapter.load Then
            ProgressReport(4, "Fill Data..")
            ProgressReport(1, "Done.")
        Else
            ProgressReport(1, String.Format("Has error::{0}", myAdapter.errormsg))
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
                    DataGridView1.AutoGenerateColumns = False
                    bs.DataSource = myAdapter.DS.Tables(0)
                    DataGridView1.DataSource = BS
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select
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


    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub StatusStrip1_ItemClicked(sender As Object, e As ToolStripItemClickedEventArgs) Handles StatusStrip1.ItemClicked

    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub
End Class
