Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop
Public Class ReportTWSalesReport
    Inherits PostgreSQLModelAdapter
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function

    Public Property errmsg As String
    Dim startdate As Date
    Dim enddate As Date
    Dim mypath As String = My.Settings.TWAutoReport
    Dim filename As String = "TWSalesReport.xlsx"

    Public Function GenerateReport() As Boolean
        'Dim myCriteria As String = String.Empty
        startdate = CDate(Today.Date.Year & "-1-1")
        enddate = CDate(Today.Date.AddDays(0))

        Dim result As Boolean = False
        Dim hwnd As System.IntPtr
        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()
        'Cursor.Current = Cursors.WaitCursor


        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty

        Try
            'Create Object Excel 
            'ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False
            'ProgressReport(2, "Opening Template...")
            'ProgressReport(2, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")

            Dim counter As Integer = 0
            'ProgressReport(2, "Creating Worksheet...")
            'backOrder
            For i = 0 To 2
                oWb.Worksheets.Add(After:=(oWb.Worksheets(3 + i)))
            Next i

            Dim sqlstr As String = String.Empty
            Dim obj As New ThreadPoolObj

            'Get Filter

            obj.osheet = oWb.Worksheets(6)
            Dim myfilter As New System.Text.StringBuilder

            obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                             " date_part('month',invdate) as month,retur as salesreturn,sales.get_sales(retur,totalsales) as sales,sales.get_return(retur,totalsales) as return,sales.get_salesreversal(retur,totalsales) as salesreversal," &
                             " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                             " from sales.txtw tx " &
                             " left join sales.customer c on c.customerid = tx.customerid " &
                             " left join sales.tw_family f on f.familyname = tx.productfamily" &
                             " left join sales.mla m on m.id = tx.mlacode" &
                             " where invdate >= '" & String.Format("{0:yyyy-MM-dd}", startdate) & "' and invdate <= '" & String.Format("{0:yyyy-MM-dd}", enddate) & "' order by invdate)"

            obj.strsql = "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,salesman,shipto,productid,cmmf,sbu,productfamily,brand,materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,totalcost,qty as totalqty ,totalsales as totalsales,totalcost as totalcost,region,location," &
                             " date_part('month',invdate) as month,retur as salesreturn,sales.get_sales(retur,totalsales) as sales,sales.get_return(retur,totalsales) as return,sales.get_salesreversal(retur,totalsales) as salesreversal," &
                             " custtype as channel,custname,f.type as ec,f.id as famlv1,merch,storename,mlacode,m.mlaname ,posid,od,invdate as filterdate1,invdate as filterdate2" &
                             " from sales.txtw tx " &
                             " left join sales.customer c on c.customerid = tx.customerid " &
                             " left join sales.tw_family f on f.familyname = tx.productfamily" &
                             " left join sales.mla m on m.id = tx.mlacode" &
                             " where invdate >= '" & String.Format("{0:yyyy-MM-dd}", startdate) & "' and invdate <= '" & String.Format("{0:yyyy-MM-dd}", enddate) & "' order by invdate)"



            obj.osheet.Name = "DATA"

            ExcellStuff.FillWorksheet(obj.osheet, obj.strsql, dbAdapter1)
            Dim lastrow = obj.osheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                'ProgressReport(2, "Generating Pivot Tables..")
                'oXl.Visible = True
                CreatePivotTable(oXl, oWb, 1, enddate)
                'createchart(oWb, 1, errmsg)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()
            'Filename = ValidateFileName(Filename, Filename & "\" & String.Format("Sales-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day)))

            'Filename = ValidateFileName(SelectedPath, Filename)
            filename = String.Format("{0}TWSalesReportYTD.xlsx", mypath)


            'ProgressReport(2, "Done ")
            'ProgressReport(5, "Saving File ...")
            oWb.SaveAs(filename)
            'ProgressReport(5, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            'ProgressReport(2, "")
            errmsg = ex.Message
        Finally
            'clear excel from memory
            oXl.Quit()
            ExcellStuff.releaseComObject(oSheet)
            ExcellStuff.releaseComObject(oWb)
            ExcellStuff.releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try
            Cursor.Current = Cursors.Default
        End Try
        Return result
    End Function

    Private Sub CreatePivotTable(ByVal oxl As Excel.Application, ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal mydate As Date)
        Dim osheet As Excel.Worksheet
        'oWb.Names.Add("dbRange", RefersToR1C1:="=OFFSET('data'!R1C1,0,0,COUNTA('data'!C1),COUNTA('data'!R1))")
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        'oWb.Worksheets("Sheet1").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R6C15", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.Cells(1, 1) = "SEB ASIA LTD - TAIWAN BRANCH"
        osheet.Cells(2, 1) = "GROSS SALES (including sales return)"
        osheet.Cells(3, 1) = "Currency: NTD"
        With osheet.Range("A1:A2")
            .Font.Size = 20
        End With

        With osheet.Range("A3").Font
            .Size = 10
            .FontStyle = "Bold"
            .Color = -16776961
            .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add(" GrandTotal", "=sales + return + salesreversal", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add(" Proportion", "=return /(sales + salesreversal)", True)

        osheet.PivotTables("PivotTable1").Pivotfields("mlaname").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("mlacode").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesreturn").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("famlv1").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("storename").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("channel").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("merch").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField



        osheet.PivotTables("PivotTable1").Pivotfields("custname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales"), " Totals Sales ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").Pivotfields("month").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        osheet.Range("A18").Select()
        oxl.ActiveWindow.FreezePanes = True

        osheet.Name = "Gross Sales"

        osheet.Cells.EntireColumn.AutoFit()

        ''Second PivotTable
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        'oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        oWb.Worksheets("Gross Sales").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.Cells(1, 1) = "SEB ASIA LTD - TAIWAN BRANCH"
        osheet.Cells(2, 1) = "Quantity (including sales return)"

        With osheet.Range("A1:A2")
            .Font.Size = 20
        End With


        osheet.PivotTables("PivotTable1").Pivotfields("mlaname").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("mlacode").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesreturn").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("famlv1").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("storename").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("channel").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("merch").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField



        osheet.PivotTables("PivotTable1").Pivotfields("custname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalqty"), " Totals QTY ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Totals QTY ").numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").Pivotfields("month").orientation = Excel.XlPivotFieldOrientation.xlColumnField

        osheet.Range("A18").Select()
        oxl.ActiveWindow.FreezePanes = True
        osheet.Name = "QTY"


        isheet = isheet + 1

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        'oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R16C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        oWb.Worksheets("Gross Sales").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R8C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.Cells(1, 1) = "SEB ASIA LTD - TAIWAN BRANCH"
        osheet.Cells(2, 1) = "Sales return on gross sales in % by customer"
        osheet.Cells(3, 1) = "Currency: NTD"
        With osheet.Range("A1:A2")
            .Font.Size = 20
        End With
        With osheet.Range("A3").Font
            .Size = 10
            .FontStyle = "Bold"
            .Color = -16776961
            .Underline = Excel.XlUnderlineStyle.xlUnderlineStyleSingle
        End With
        osheet.PivotTables("PivotTable1").Pivotfields("month").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("custname").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sales"), " Sales ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("salesreversal"), " Sales Reversal ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("return"), " Return ", Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields(" Sales ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").PivotFields(" Sales Reversal ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").PivotFields(" Return ").numberformat = "#,##0_);(#,##0)"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields(" GrandTotal"), " GrandTotal ", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields(" Proportion"), " Proportion on sales (in %)", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" GrandTotal ").numberformat = "#,##0_);(#,##0)"
        osheet.PivotTables("PivotTable1").PivotFields(" Proportion on sales (in %)").numberformat = "0%"


        osheet.Name = "Sales Return"


        osheet.Cells.EntireColumn.AutoFit()
        isheet = 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

    End Sub
End Class
Class ThreadPoolObj
    Public ObjectID As Integer
    Public signal As System.Threading.ManualResetEvent
    Public osheet As Excel.Worksheet
    Public ds As DataSet
    Public sb As System.Text.StringBuilder
    Public strsql As String
    Public Name As String
End Class

Class ThreadPoolManualResetEvent
    Public ObjectID As Integer
    Public signal As System.Threading.ManualResetEvent
End Class