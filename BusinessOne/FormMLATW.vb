Imports System.Threading
Public Class FormMLATW
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim myAdapter As MLAAdapter

    Private Sub LoadData()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub

    Private Sub FormUser_Load(sender As Object, e As EventArgs) Handles Me.Load
        LoadData()
    End Sub

    Sub DoWork()
        myAdapter = New MLAAdapter
        Try
            ProgressReport(1, "Loading..")
            If myAdapter.LoadData() Then
                ProgressReport(4, "Init Data")
            End If
            ProgressReport(1, "Done.")
        Catch ex As Exception

            ProgressReport(1, ex.Message)
        End Try
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
                    DataGridView1.DataSource = myAdapter.BS
            End Select
        End If
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        LoadData()
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Try
            If User.can("createMLATW") Then
                ShowTx(TxRecord.AddRecord)
            Else

                MessageBox.Show("Sorry,You cannot create a new record.")
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
    End Sub

    Private Sub ShowTx(ByVal StatusTx As TxRecord)
        Dim drv As DataRowView = Nothing
        Select Case StatusTx
            Case TxRecord.AddRecord
                drv = myAdapter.BS.AddNew
            Case TxRecord.UpdateRecord
                drv = myAdapter.BS.Current
        End Select
        Dim myform As New DialogMLATWInput(drv)
        myform.ShowDialog()
    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        ShowTx(TxRecord.UpdateRecord)
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        If User.can("createMLATW") Then
            myAdapter.Save()
        Else
            MessageBox.Show("Sorry,You cannot save record.")
        End If

    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        If User.can("createMLATW") Then
            If Not IsNothing(myAdapter.BS.Current) Then
                If MessageBox.Show("Delete this record?", "Delete Record", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = DialogResult.OK Then
                    For Each drv As DataGridViewRow In DataGridView1.SelectedRows
                        myAdapter.BS.RemoveAt(drv.Index)
                    Next
                End If
            End If
        Else
            MessageBox.Show("Sorry, you cannot delete record.")
        End If
        
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        RemoveHandler DialogMLATWInput.FinishUpdate, AddressOf RefreshDataGrid
        AddHandler DialogMLATWInput.FinishUpdate, AddressOf RefreshDataGrid
    End Sub

    Private Sub RefreshDataGrid()
        DataGridView1.Invalidate()
    End Sub


   
End Class