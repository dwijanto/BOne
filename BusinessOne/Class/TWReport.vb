Imports Microsoft.Office.Interop
Imports System.Text
Imports BusinessOne.ValidateClass
Public Class TWReport
    Inherits TaiwanModelAdapter

    Public ErrorMessage As String
    Public startdate As Date
    Dim myadapter As SalesOrderTWAdapter
    Dim mySBUAdapter As SBUAdapter
    Dim myBrandAdapter As BrandAdapter
    Dim myFamilyAdapter As FamilyTWAdapter
    Dim myFamilyLv2Adapter As SubFamilyTWAdapter
    Private myMLAAdapter As MLAAdapter
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

            myBrandAdapter = New BrandAdapter
            If Not myBrandAdapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myFamilyAdapter = New FamilyTWAdapter
            If Not myFamilyAdapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myFamilyLv2Adapter = New SubFamilyTWAdapter
            If Not myFamilyLv2Adapter.getDataSet() Then
                ErrorMessage = mySBUAdapter.errorMsg
                Return myret
            End If

            myMLAAdapter = New MLAAdapter
            If myMLAAdapter.getDataSet() Then
            Else
                ErrorMessage = myMLAAdapter.errorMsg
                Return myret
            End If

            Dim BS = New BindingSource
            myadapter = New SalesOrderTWAdapter
            'If myadapter.LoadData(Today.Date.AddDays(-1), Today.Date.AddDays(-1)) Then
            If myadapter.LoadData(startdate, Today.Date.AddDays(-1)) Then
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
        For Each dr As DataRow In myadapter.DS.Tables(0).Rows
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
    Private Sub fillMissingData1()
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

        For Each dr As DataRow In myadapter.DS.Tables(0).Rows
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
            familykey(0) = dr.Item("Family lv1")
            result = myFamilyAdapter.DS.Tables(0).Rows.Find(familykey)
            If Not IsNothing(result) Then
                dr.Item("prodfamily") = result.Item("familyname").ToString.Trim
                dr.EndEdit()
            End If

            Dim familylv2key(0) As Object
            familylv2key(0) = dr.Item("Family lv2")
            result = myFamilyLv2Adapter.DS.Tables(0).Rows.Find(familylv2key)
            If Not IsNothing(result) Then
                dr.Item("subfamily") = result.Item("familylv2name")
                dr.EndEdit()
            End If
        Next

    End Sub

    Private Sub GenerateSalesExtract()

        'FileName
        'Dim FileName = String.Format("{0}TWReport{1:yyyyMMdd}.xlsx", "c:\junk\SalesExtract", Date.Today)
        Dim FileName = String.Format("{0}TWReport{1:yyyyMMdd}.xlsx", My.Settings.TWAutoReport, Date.Today.AddDays(-1))
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
        'osheet.Columns("AD:AF").delete()
    End Sub

    Private Sub InsertIntoTable()
        Dim mySB As New StringBuilder

        'Delete previous data
        'mySB.Append(String.Format("delete from sales.txtw where invdate >= '{0:yyyy-MM-dd}' and invdate <='{0:yyyy-MM-dd}';", Date.Today.Date.AddDays(-1)))
        mySB.Append(String.Format("delete from sales.txtw where invdate >= '{0:yyyy-MM-dd}' and invdate <='{1:yyyy-MM-dd}';", startdate, Date.Today.Date.AddDays(-1)))
        'Create Insert Data

        mySB.Append("insert into sales.txtw(invid, invdate,  orderno ,  customerid ,  customername,  reportcode,  saleforce ,  country ,  custtype ,  salesman ,  shipto ,  productid ,  cmmf ,  sbu ,  productfamily ,  brand ,  materialdesc ,  supplierid ,  qty ,  totalsales ,  totalcost ,  retur,creditnote,  custname ,  merch ,  storename ,  mlacode ,  posid ,  od ) values ")
        Dim i As Integer = 0
        For Each dr As DataRow In myadapter.DS.Tables(0).Rows
            If i > 0 Then
                mySB.Append(",")
            End If
            i = 1
            mySB.Append(String.Format("({0},{1},{2},{3},{4},{5},{6},{7},{8},{9},{10},{11},{12},{13},{14},{15},{16},{17},{18},{19},{20},{21},{22},{23},{24},{25},{26},{27},{28})",
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
                                      ValidStr(dr.Item(22)),
                                      ValidStr(dr.Item(23)),
                                      ValidStr(dr.Item(25)),
                                      ValidStr(dr.Item(28)),
                                      ValidStr(dr.Item(29)),
                                      ValidStr(dr.Item(30)),
                                      ValidStr(dr.Item(32)),
                                      ValidStr(dr.Item(33))
                                      ))

        Next
        mySB.Append(";")
        'Execute sql
        mySBUAdapter.dbAdapter1.ExecuteNonQuery(mySB.ToString)
    End Sub

End Class
