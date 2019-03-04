Imports System.Threading
Imports System.Text

Public Class FormItemPrice
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim BS As BindingSource
    Dim myAdapter As New ItemPriceTMPController
    Dim BOneAdapter As ItemPriceAdapter

    Private docnum As Long
    Private ItemPriceDS As DataSet

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Private Sub FormOrderItem_Load(sender As Object, e As EventArgs) Handles Me.Load
        loaddata()
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

    Private Sub GetItemPriceDS()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork01)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Private Sub uploaddata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork02)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub
    Sub DoWork()
        ProgressReport(1, "Loading data. Please wait...")
        ProgressReport(6, "Marquee")
        BS = New BindingSource
        myAdapter = New ItemPriceTMPController
        If myAdapter.loaddata Then
            ProgressReport(4, "Fill Data..")
            ProgressReport(1, String.Format("Done. Records Count({0})", BS.Count))
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
                    DataGridView1.AutoGenerateColumns = False

                    Dim producttypeBS = myAdapter.Model.getProductTypeName


                    BS.DataSource = myAdapter.DS.Tables(0)
                    DataGridView1.DataSource = BS
                    Dim myCB As DataGridViewComboBoxColumn = DataGridView1.Columns("ProductTypeNameColumn")
                    myCB.DataSource = producttypeBS
                    myCB.DisplayMember = "producttypename"
                    myCB.ValueMember = "producttypeid"
                    myCB.DataPropertyName = "producttypeid"
                    ToolStripTextBox1.Clear()
                    ToolStripButton5.Enabled = myAdapter.Model.CanRollback

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 8
                    '***Refresh Table ItemPriceTmp
                    ToolStripStatusLabel1.Text = "Refresh data..."
                    'Set Added
                    BOneAdapter.DS.Tables(0).TableName = "ItemPriceBone"
                    For i = 0 To BOneAdapter.DS.Tables(0).Rows.Count - 1
                        BOneAdapter.DS.Tables(0).Rows(i).SetAdded()
                    Next
                    'Delete ItemPriceTmp
                    myAdapter.DeleteAll()
                    myAdapter.save(BOneAdapter.DS)
                    myAdapter.loaddata()
                    DataGridView1.AutoGenerateColumns = False
                    BS.DataSource = myAdapter.DS.Tables(0)
                    DataGridView1.DataSource = BS
                    ToolStripStatusLabel1.Text = "Done..."
                Case 9
                    '****Upload ItemPrice to E-Staff
                    ToolStripStatusLabel1.Text = "Upload data..."
                    'Delete ItemPrice
                    myAdapter.DeleteItemPrice()
                    'Set RowState Added for Record with instore

                    'Remove filter first
                    BS.Filter = ""
                    For Each drv As DataRowView In BS.List
                        If drv.Row.Item("instore") Then
                            If drv.Row.RowState = DataRowState.Unchanged Then
                                drv.Row.SetModified()
                            End If
                        End If
                    Next
                    'Save to E-Staff
                    myAdapter.saveItemPrice(myAdapter.DS)
                    DoWork()
                    ToolStripStatusLabel1.Text = "Upload data done"
                Case 10
                    ToolStripStatusLabel1.Text = "Rollback data..."
                    myAdapter.Rollback()
                    DoWork()
                    ToolStripStatusLabel1.Text = "Rollback data Done"
            End Select
        End If
    End Sub

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        If MessageBox.Show("Do you want to retrieve data from business one Server?", "Retrieve Data", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
            GetBoneData()
        End If

    End Sub

    Private Sub GetBoneData()
        GetItemPriceDS()
    End Sub

    Private Sub DoWork01()
        ProgressReport(1, "Loading data. Please wait...")
        ProgressReport(6, "Marquee")
        BS = New BindingSource
        BOneAdapter = New ItemPriceAdapter
        ItemPriceDS = New DataSet
        If BOneAdapter.loadData Then
            ProgressReport(8, "Fill Data..")
            ProgressReport(1, "Done.")
        Else
            ProgressReport(1, String.Format("Has error::{0}", BOneAdapter.errorMsg))
        End If
        ProgressReport(5, "Continuous")
    End Sub


    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        myAdapter.ApplyFilter = ToolStripTextBox1.Text
        ProgressReport(1, String.Format("Records found: ({0})", BS.Count))
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        loaddata()
    End Sub

    Private Sub DataGridView1_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles DataGridView1.DataError

    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        If MessageBox.Show("Do you want to upload to E-Staff Database?", "Upload Data", MessageBoxButtons.OKCancel, MessageBoxIcon.Question) = Windows.Forms.DialogResult.OK Then
            If Me.validate Then
                uploaddata()
            Else
                ProgressReport(1, "Upload failed. Please fix the error(s).")
            End If
        End If
    End Sub

    Private Sub DoWork02()

        'If Me.validate() Then
        ProgressReport(9, "Upload in process..")
        'Else
        'ProgressReport(1, "Upload failed. Please fix the error(s).")
        'End If
    End Sub
    Private Sub DoWork03()
        ProgressReport(10, "Rollback...")
    End Sub
    Public Overloads Function validate() As Boolean
        'Filter BindingSource only Instore
        DataGridView1.EndEdit()
        BS.EndEdit()

        'Clear Error First
        For Each drv As DataRowView In BS.List
            Dim errorsb As New StringBuilder
            addError(drv.Row, errorsb, "")
        Next


        'BS.Filter = "instore = true"


        For Each drv As DataRowView In BS.List
            If drv.Row.Item("instore") = True Then
                Dim errorsb As New StringBuilder
                'check BrandId

                If IsDBNull(drv.Row.Item("u_sebbran3")) Then
                    addError(drv.Row, errorsb, "Business One Brand is missing.")
                ElseIf drv.Row.Item("u_sebbran3") = "" Then
                    addError(drv.Row, errorsb, "Business One Brand is missing.")
                Else
                    If IsDBNull(drv.Row.Item("brandname")) Then
                        addError(drv.Row, errorsb, String.Format("Brand Name for Brandid {0} is missing.", drv.Row.Item("u_sebbran3")))
                    ElseIf drv.Row.Item("brandname") = "" Then
                        addError(drv.Row, errorsb, String.Format("Brand Name for Brandid {0} is missing.", drv.Row.Item("u_sebbran3")))
                    End If
                End If

                


                If IsDBNull(drv.Row.Item("productname")) Then
                    addError(drv.Row, errorsb, "Product Name is missing.")
                ElseIf drv.Row.Item("productname") = "" Then
                    addError(drv.Row, errorsb, "Product Name is missing.")
                End If

                If IsDBNull(drv.Row.Item("producttypeid")) Then
                    addError(drv.Row, errorsb, "Product Type Name is missing.")
                    'ElseIf drv.Row.Item("producttypeid") = "" Then
                    'addError(drv.Row, errorsb, "Product Type Name is missing.")
                End If

                If IsDBNull(drv.Row.Item("familyname")) Then
                    addError(drv.Row, errorsb, String.Format("Family Name for familyid {0} is missing.", drv.Row.Item("u_sebfamlev1cury")))
                ElseIf drv.Row.Item("familyname") = "" Then
                    addError(drv.Row, errorsb, String.Format("Family Name for familyid {0} is missing.", drv.Row.Item("u_sebfamlev1cury")))
                End If
            End If
            
        Next
        If myAdapter.DS.Tables(0).HasErrors Then
            Return False
        End If

        Return True
    End Function

    Private Sub addError(dr As DataRow, errorSB As StringBuilder, errormessage As String)
        If errorSB.Length > 0 Then
            errorSB.Append(",")
        End If
        errorSB.Append(errormessage)
        dr.RowError = errorSB.ToString
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        If MessageBox.Show("Do you want to rollback price?") = Windows.Forms.DialogResult.OK Then
            rollbackdata()
        End If
    End Sub

    Private Sub rollbackdata()
        If Not myThread.IsAlive Then
            ToolStripStatusLabel1.Text = ""
            myThread = New Thread(AddressOf DoWork03)
            myThread.Start()
        Else
            MessageBox.Show("Please wait until the current process is finished.")
        End If
    End Sub



End Class