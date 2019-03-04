Imports Microsoft.Office.Interop
Imports System.Threading
Public Class FormSalesReportTW
    Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
    Dim myThread As New System.Threading.Thread(AddressOf DoWork)
    Dim BS As BindingSource
    Private startdate As Date
    Private enddate As Date
    Private myAdapter As SalesOrderTWAdapter

    Private mySBUAdapter As SBUAdapter
    Private myBrandAdapter As BrandAdapter
    Private myFamilyAdapter As FamilyTWAdapter
    Private myFamilyLv2Adapter As subfamilyTWAdapter
    Private myMLAAdapter As MLAAdapter

    Private Sub ToolStripButton1_Click(sender As Object, e As EventArgs) Handles ToolStripButton1.Click
        Dim mydialog As New DialogDateRange
        If mydialog.ShowDialog = Windows.Forms.DialogResult.OK Then
            startdate = mydialog.startdate
            enddate = mydialog.enddate
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

    Sub DoWork()
        ProgressReport(1, "Loading data. Please wait...")
        ProgressReport(6, "Marquee")


        mySBUAdapter = New SBUAdapter
        If mySBUAdapter.getDataSet() Then

        Else
            ProgressReport(1, String.Format("Has error::{0}", mySBUAdapter.errorMsg))
            Exit Sub
        End If

        myBrandAdapter = New BrandAdapter
        If myBrandAdapter.getDataSet() Then

        Else
            ProgressReport(1, String.Format("Has error::{0}", myBrandAdapter.errorMsg))
            Exit Sub
        End If

        myFamilyAdapter = New FamilyTWAdapter
        If myFamilyAdapter.getDataSet() Then

        Else
            ProgressReport(1, String.Format("Has error::{0}", myFamilyAdapter.errorMsg))
            Exit Sub
        End If

        myFamilyLv2Adapter = New SubFamilyTWAdapter
        If myFamilyLv2Adapter.getDataSet() Then

        Else
            ProgressReport(1, String.Format("Has error::{0}", myFamilyLv2Adapter.errorMsg))
            Exit Sub
        End If

        myMLAAdapter = New MLAAdapter
        If myMLAAdapter.getDataSet() Then

        Else
            ProgressReport(1, String.Format("Has error::{0}", myMLAAdapter.errorMsg))
            Exit Sub
        End If
        ProgressReport(7, "Fill Data..")

        BS = New BindingSource
        myAdapter = New SalesOrderTWAdapter
        If myAdapter.LoadData(startdate, enddate) Then
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
                    DataGridView1.AutoGenerateColumns = False
                    'DataGridView1.AutoGenerateColumns = True
                    'Fill missing data
                    fillMissingData()
                    BS.DataSource = myAdapter.DS.Tables(0)
                    DataGridView1.DataSource = BS
                    ToolStripStatusLabel2.Text = String.Format("Record(s) count: {0:#,##0}", BS.Count)
                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee

                Case 7
                    'Fill in SBU
                    Dim pk(0) As DataColumn
                    pk(0) = mySBUAdapter.DS.Tables(0).Columns("sbuid")
                    mySBUAdapter.DS.Tables(0).PrimaryKey = pk
                    mySBUAdapter.DS.Tables(0).TableName = "SBU"


                    pk(0) = myBrandAdapter.DS.Tables(0).Columns("brandid")
                    myBrandAdapter.DS.Tables(0).PrimaryKey = pk
                    myBrandAdapter.DS.Tables(0).TableName = "BRAND"


                    pk(0) = myFamilyAdapter.DS.Tables(0).Columns("familyid")
                    myFamilyAdapter.DS.Tables(0).PrimaryKey = pk
                    myFamilyAdapter.DS.Tables(0).TableName = "FAMILY"


                    pk(0) = myFamilyLv2Adapter.DS.Tables(0).Columns("familylv2id")
                    myFamilyLv2Adapter.DS.Tables(0).PrimaryKey = pk
                    myFamilyLv2Adapter.DS.Tables(0).TableName = "FAMILYLV2"

                    pk(0) = myMLAAdapter.DS.Tables(0).Columns("id")
                    myMLAAdapter.DS.Tables(0).PrimaryKey = pk
                    myMLAAdapter.DS.Tables(0).TableName = "MLA"


            End Select
        End If
    End Sub

    Private Sub fillMissingData()
        For Each dr As DataRow In myAdapter.DS.Tables(0).Rows
            Dim sbusapkey(0) As Object
            sbusapkey(0) = dr.Item("Product Line")
            Dim result = mySBUAdapter.DS.Tables(0).Rows.Find(sbusapkey)
            If Not IsNothing(result) Then
                dr.Item("sbu") = result.Item("sbuname2")
                dr.EndEdit()
            End If

            Dim brandkey(0) As Object
            brandkey(0) = dr.Item("Brand Code")
            result = myBrandAdapter.DS.Tables(0).Rows.Find(brandkey)
            If Not IsNothing(result) Then
                dr.Item("brand") = result.Item("brandname")
                dr.EndEdit()
            End If

            Dim familykey(0) As Object
            familykey(0) = dr.Item("FamLv 1")
            result = myFamilyAdapter.DS.Tables(0).Rows.Find(familykey)
            If Not IsNothing(result) Then
                dr.Item("prodfamily") = result.Item("familyname").ToString.Trim
                dr.Item("E/C") = result.Item("type")
                dr.EndEdit()
            End If

            Dim familylv2key(0) As Object
            familylv2key(0) = dr.Item("Family lv2")
            result = myFamilyLv2Adapter.DS.Tables(0).Rows.Find(familylv2key)
            If Not IsNothing(result) Then
                dr.Item("subfamily") = result.Item("familylv2name")
                dr.EndEdit()
            End If

            Dim mlakey(0) As Object
            mlakey(0) = dr.Item("MLA Code")
            result = myMLAAdapter.DS.Tables(0).Rows.Find(mlakey)
            If Not IsNothing(result) Then
                dr.Item("MLA Name") = result.Item("mlaname")
                dr.EndEdit()
            End If

            'change Saleman
            If Not IsDBNull(dr.Item("saleman")) Then
                Select Case dr.Item("saleman")
                    Case "Jerry Yu"
                        dr.Item("saleman") = "Jerry"
                    Case "Marco Tang"
                        dr.Item("saleman") = "Marco"

                End Select
            End If

            'change Saleman
            If Not IsDBNull(dr.Item("Customer Name")) Then
                Select Case dr.Item("Customer Name")
                    Case "CARREFOUR"
                        dr.Item("Customer Name") = "Carrefour"
                    Case "CHUNG YO"
                        dr.Item("Customer Name") = "CHUNG YO DEPARTMENT"
                    Case "CUSTOMER SVC"
                        dr.Item("Customer Name") = "Customer Service"                  
                    Case "EMPLOYEE PURCH"
                        dr.Item("Customer Name") = "Employee Purchase"
                    Case "FENG PING SOGO"
                        dr.Item("Customer Name") = "FENG PING HSING YEH"
                    Case "KS SOGO"
                        dr.Item("Customer Name") = "Kuang San SOGO"
                    Case "MITSUKOSHI"
                        dr.Item("Customer Name") = "Mitsukoshi"
                    Case "PACIFIC SOGO"
                        dr.Item("Customer Name") = "Pacific SOGO"
                    Case "TAISUGAR"
                        dr.Item("Customer Name") = "TaiSuger"
                    Case "OTHERS"
                        If dr.Item("customer_i") = "C028A" Then
                            dr.Item("Customer Name") = "CHIN TAI"
                        ElseIf dr.Item("customer_i") = "C021A" Then
                            dr.Item("Customer Name") = "CLEARANCE SALES"
                        End If

                End Select
            End If

        Next
        

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
        mysaveform.FileName = String.Format("{0}Report{1:yyyyMMdd}.xlsx", "SalesExtract", Date.Today)

        If (mysaveform.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            Dim filename = IO.Path.GetDirectoryName(mysaveform.FileName)
            Dim reportname = IO.Path.GetFileName(mysaveform.FileName)

            Dim datasheet As Integer = 1

            Dim mycallback As FormatReportDelegate = AddressOf FormattingReport
            Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable
            Dim myTable = myAdapter.DS.Tables(0).Copy

            myTable.Columns.Remove("Local Description")
            myTable.Columns.Remove("Family lv2")
            myTable.Columns.Remove("Brand Code")
            myTable.Columns.Remove("Product Line")
            myTable.Columns.Remove("Price")
            myTable.Columns.Remove("Credit Note Reason")
            myTable.Columns.Remove("CN Number (by user)")
            myTable.Columns.Remove("subfamily")

            Dim myreport As New ExcelExtract(Me, mysaveform.FileName, "\templates\ExcelTemplate.xltx", myTable, mycallback, PivotCallback)
            myreport.ExtractFromDataTableUnsyncDT(Me, New System.EventArgs)
        End If

        ProgressReport(1, "Loading Data.Done!")
        ProgressReport(5, "Continuous")
    End Sub

    Private Sub FormattingReport(ByRef osheet As Excel.Worksheet, ByRef e As EventArgs)
        osheet.Columns("AD:AF").delete()
    End Sub

    Private Sub PivotTable()

    End Sub
End Class