Imports Microsoft.Office.Interop
Imports System.Text
Imports BusinessOne.ValidateClass

Public Class HKReport
    Inherits ModelAdapter

    Public ErrorMessage As String
    Public startdate As Date

    Dim myadapter As SalesOrderAdapter
    Dim mySBUAdapter As SBUAdapter
    Dim myCompanyAdapter As CompanyAdapter 'location Class is in SBUAdapterClass
    Dim myBrandAdapter As BrandAdapter
    Dim myFamilyAdapter As FamilyAdapter
    Dim myFamilyLv2Adapter As subfamilyAdapter


    Dim Parent As Object
    Public Sub New(ByVal Parent As Object)
        Me.Parent = Parent
    End Sub

    Public Function SalesExtract() As Boolean
        Dim myret As Boolean = False

        Try
            mySBUAdapter = New SBUAdapter
            If Not mySBUAdapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myCompanyAdapter = New CompanyAdapter
            If Not myCompanyAdapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myBrandAdapter = New BrandAdapter
            If Not myBrandAdapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myFamilyAdapter = New FamilyAdapter
            If Not myFamilyAdapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myFamilyLv2Adapter = New subfamilyAdapter
            If Not myFamilyLv2Adapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            Dim BS = New BindingSource
            myadapter = New SalesOrderAdapter
            'If myadapter.LoadData(Today.Date.AddDays(-1), Today.Date.AddDays(-1)) Then
            'If myadapter.LoadData(startdate, Today.Date.AddDays(-1)) Then
            If myadapter.LoadData(startdate, Today.Date) Then
                fillMissingData()
                GenerateSalesExtract()
                InsertIntoTable()
                myret = True
            Else
                ErrorMessage = myadapter.errorMsg
            End If
        Catch ex As Exception
            ErrorMessage = ex.Message
        End Try
        Return myret
    End Function
    Private Sub fillMissingData()
        'Dim pk(0) As DataColumn
        'pk(0) = mySBUAdapter.DS.Tables(0).Columns("sbuid")
        'mySBUAdapter.DS.Tables(0).PrimaryKey = pk
        'mySBUAdapter.DS.Tables(0).TableName = "SBU"


        'pk(0) = myBrandAdapter.DS.Tables(0).Columns("brandid")
        'myBrandAdapter.DS.Tables(0).PrimaryKey = pk
        'myBrandAdapter.DS.Tables(0).TableName = "BRAND"


        'pk(0) = myFamilyAdapter.DS.Tables(0).Columns("familyid")
        'myFamilyAdapter.DS.Tables(0).PrimaryKey = pk
        'myFamilyAdapter.DS.Tables(0).TableName = "FAMILY"


        'pk(0) = myFamilyLv2Adapter.DS.Tables(0).Columns("familylv2id")
        'myFamilyLv2Adapter.DS.Tables(0).PrimaryKey = pk
        'myFamilyLv2Adapter.DS.Tables(0).TableName = "FAMILYLV2"

        Dim pk(0) As DataColumn
        pk(0) = mySBUAdapter.DS.Tables(0).Columns("sbuid")
        mySBUAdapter.DS.Tables(0).PrimaryKey = pk
        mySBUAdapter.DS.Tables(0).TableName = "SBU"

        pk(0) = myCompanyAdapter.DS.Tables(0).Columns("customerid")
        myCompanyAdapter.DS.Tables(0).PrimaryKey = pk
        myCompanyAdapter.DS.Tables(0).TableName = "Customer"
        pk(0) = myCompanyAdapter.DS.Tables(1).Columns("customerid")
        myCompanyAdapter.DS.Tables(1).PrimaryKey = pk
        myCompanyAdapter.DS.Tables(1).TableName = "CustProdKAM"

        pk(0) = myBrandAdapter.DS.Tables(0).Columns("brandid")
        myBrandAdapter.DS.Tables(0).PrimaryKey = pk
        myBrandAdapter.DS.Tables(0).TableName = "BRAND"


        pk(0) = myFamilyAdapter.DS.Tables(0).Columns("familyid")
        myFamilyAdapter.DS.Tables(0).PrimaryKey = pk
        myFamilyAdapter.DS.Tables(0).TableName = "FAMILY"


        pk(0) = myFamilyLv2Adapter.DS.Tables(0).Columns("subfamilyid")
        myFamilyLv2Adapter.DS.Tables(0).PrimaryKey = pk
        myFamilyLv2Adapter.DS.Tables(0).TableName = "FAMILYLV2"

        For Each dr As DataRow In myadapter.DS.Tables(0).Rows
            If Not IsNumeric(dr.Item("cmmf")) Then
                dr.Item("cmmf") = 1
                dr.EndEdit()
            End If

            Dim sbusapkey(0) As Object
            sbusapkey(0) = dr.Item("U_SEBProdLinePi2")
            Dim result = mySBUAdapter.DS.Tables(0).Rows.Find(sbusapkey)
            If Not IsNothing(result) Then
                dr.Item("sbu") = result.Item("sbuname2")
                dr.EndEdit()
            End If

            Dim brandkey(0) As Object
            brandkey(0) = dr.Item("SEBbran2")
            result = myBrandAdapter.DS.Tables(0).Rows.Find(brandkey)
            If Not IsNothing(result) Then
                dr.Item("brand") = result.Item("brandname")
                dr.EndEdit()
            End If

            Dim familykey(0) As Object
            familykey(0) = dr.Item("familycode")
            result = myFamilyAdapter.DS.Tables(0).Rows.Find(familykey)
            If Not IsNothing(result) Then
                dr.Item("prodfamily") = result.Item("familyname").ToString.Trim
                dr.EndEdit()
            End If

            Dim familylv2key(0) As Object
            'familylv2key(0) = dr.Item("U_SEBFamLev2CurY")
            familylv2key(0) = dr.Item("subfamcode")
            result = myFamilyLv2Adapter.DS.Tables(0).Rows.Find(familylv2key)
            If Not IsNothing(result) Then
                'dr.Item("subfamily") = result.Item("familylv2name")
                dr.Item("subfamily") = result.Item("subfamilyname")
                dr.Item("subfamcode") = result.Item("subfamcode")
                dr.EndEdit()
            End If
            'change Saleman
            Select Case dr.Item("saleman")
                Case "Antonio SIN"
                    dr.Item("saleman") = "Antonio"
                Case "Catherine CHAN"
                    dr.Item("saleman") = "Catherine"
                Case "Jack PAU"
                    dr.Item("saleman") = "Jack"
                Case "Joe LO"
                    dr.Item("saleman") = "Joe"
                Case "Wai Kit CHAN"
                    dr.Item("saleman") = "Wai Kit"
                Case "Boris TAM"
                    dr.Item("saleman") = "Boris"
            End Select
            dr.EndEdit()

            'Check New Customer
            Dim customerkey(0) As Object
            customerkey(0) = dr.Item("customer_i")
            result = myCompanyAdapter.DS.Tables(0).Rows.Find(customerkey)
            If IsNothing(result) Then
                Dim mydr As DataRow = myCompanyAdapter.DS.Tables(0).NewRow
                mydr.Item("customerid") = dr.Item("customer_i")
                mydr.Item("customername") = dr.Item("customer_n")
                myCompanyAdapter.DS.Tables(0).Rows.Add(mydr)
                Dim kam As String = dr.Item("saleman")
                If kam.Contains("Kily") Then
                    kam = "Kily"
                End If

                Dim mydr1 As DataRow = myCompanyAdapter.DS.Tables(1).NewRow
                mydr1.Item("customerid") = dr.Item("customer_i")
                mydr1.Item("sda") = kam
                mydr1.Item("ckw") = kam
                myCompanyAdapter.DS.Tables(1).Rows.Add(mydr1)
            End If
        Next
        'Update Company
        Dim ds2 As DataSet = myCompanyAdapter.DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New ContentBaseEventArgs(ds2, True, ErrorMessage, ra, True)
            If Not myCompanyAdapter.Save(Me, mye) Then
                Logger.log(mye.message)
            Else
                myCompanyAdapter.DS.AcceptChanges()
            End If
        End If
    End Sub
    Private Sub fillMissingData1()
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

        For Each dr As DataRow In myadapter.DS.Tables(0).Rows
            Dim sbusapkey(0) As Object
            sbusapkey(0) = dr.Item("U_SEBProdLinePi2")
            Dim result = mySBUAdapter.DS.Tables(0).Rows.Find(sbusapkey)
            If Not IsNothing(result) Then
                dr.Item("sbu") = result.Item("sbuname2")
                dr.EndEdit()
            End If

            Dim brandkey(0) As Object
            brandkey(0) = dr.Item("SEBbran2")
            result = myBrandAdapter.DS.Tables(0).Rows.Find(brandkey)
            If Not IsNothing(result) Then
                dr.Item("brand") = result.Item("brandname")
                dr.EndEdit()
            End If

            Dim familykey(0) As Object
            familykey(0) = dr.Item("FamilyLv1")
            result = myFamilyAdapter.DS.Tables(0).Rows.Find(familykey)
            If Not IsNothing(result) Then
                dr.Item("prodfamily") = result.Item("familyname").ToString.Trim
                dr.EndEdit()
            End If

            Dim familylv2key(0) As Object
            familylv2key(0) = dr.Item("U_SEBFamLev2CurY")
            result = myFamilyLv2Adapter.DS.Tables(0).Rows.Find(familylv2key)
            If Not IsNothing(result) Then
                dr.Item("subfamily") = result.Item("familylv2name")
                dr.EndEdit()
            End If
        Next

    End Sub

    Private Sub GenerateSalesExtract()

        'FileName
        'Dim FileName = String.Format("{0}HKReport{1:yyyyMMdd}.xlsx", "c:\junk\SalesExtract", Date.Today)
        Dim FileName = String.Format("{0}HKReport{1:yyyyMMdd}.xlsx", My.Settings.HKAutoReport, Date.Today.AddDays(-1))
        Dim myfilename = IO.Path.GetDirectoryName(FileName)
        Dim reportname = IO.Path.GetFileName(FileName)

        Dim datasheet As Integer = 1
        Dim mycallback As FormatReportDelegate = AddressOf SalesExtractFormattingReport
        Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

        Dim myreport As New ExcelExtract(Parent, FileName, "\templates\ExcelTemplate.xltx", myadapter.DS.Tables(0), mycallback, PivotCallback, False)
        'myreport.ExtractFromDataTableUnsync(Parent, New System.EventArgs)
        myreport.ExtractFromDataTableUnsyncDT(Parent, New System.EventArgs)



    End Sub


    Private Sub PivotTable()

    End Sub

    Private Sub SalesExtractFormattingReport(ByRef osheet As Excel.Worksheet, ByRef e As EventArgs)
        osheet.Columns("AD:AF").delete()
    End Sub

    Private Sub InsertIntoTable()
        Dim mySB As New StringBuilder

        'Delete previous data
        'mySB.Append(String.Format("delete from sales.tx where invdate >= '{0:yyyy-MM-dd}' and invdate <='{0:yyyy-MM-dd}';", Date.Today.Date.AddDays(-1)))
        'Dim enddate As Date = Date.Today.Date.AddDays(-1)
        Dim enddate As Date = Date.Today.Date
        mySB.Append(String.Format("delete from sales.tx where invdate >= '{0:yyyy-MM-dd}' and invdate <='{1:yyyy-MM-dd}';", startdate, enddate))
        'Create Insert Data
        mySB.Append("insert into sales.tx(invid, invdate, orderno, customerid, customername, reportcode, saleforce, country, custtype, salesman, shipto, productid, cmmf, sbu, productfamily, brand, materialdesc, supplierid, qty, totalsales, totalcost, Region,location) values ")
        Dim i As Integer = 0
        For Each dr As DataRow In myadapter.DS.Tables(0).Rows
            If i > 0 Then
                mySB.Append(",")
            End If
            i = 1
            mySB.Append(String.Format("({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22})",
                                      ValidStr(dr.Item(0)),
                                      ValidDate(dr.Item(1)),
                                      ValidStr(dr.Item(2)),
                                      ValidStr(dr.Item(3)),
                                      ValidStr(dr.Item(4)),
                                      ValidStr(dr.Item(5)),
                                      ValidStr(dr.Item(6)),
                                      ValidStr(dr.Item(7)),
                                      ValidStr(dr.Item(8)),
                                      ValidStr(dr.Item(9)),
                                      ValidStr(dr.Item(10)),
                                      ValidStr(dr.Item(11)),
                                      ValidNumeric(dr.Item(12)),
                                      ValidStr(dr.Item(13)),
                                      ValidStr(dr.Item(14)),
                                      ValidStr(dr.Item(15)),
                                      ValidStr(dr.Item(16)),
                                      ValidStr(dr.Item(17)),
                                      ValidNumeric(dr.Item(18)),
                                      ValidNumeric(dr.Item(19)),
                                      ValidNumeric(dr.Item(20)),
                                      ValidStr(dr.Item(21)),
                                      ValidStr(dr.Item(22))
                                      ))

        Next
        mySB.Append(";")
        'Execute sql
        mySBUAdapter.dbAdapter1.ExecuteNonQuery(mySB.ToString)
    End Sub


End Class
Public Class ContentBaseEventArgs
    Inherits EventArgs
    Public Property dataset As DataSet
    Public Property message As String
    Public Property hasChanges As Boolean
    Public Property ra As Integer
    Public Property continueonerror As Boolean

    Public Sub New(ByVal dataset As DataSet, ByRef haschanges As Boolean, ByRef message As String, ByRef recordaffected As Integer, ByVal continueonerror As Boolean)
        Me.dataset = dataset
        Me.message = message
        Me.ra = ra
        Me.continueonerror = continueonerror
    End Sub
End Class