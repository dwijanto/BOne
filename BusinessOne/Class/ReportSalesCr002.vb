Imports Microsoft.Office.Interop
Imports System.Runtime.InteropServices
Public Class ReportSalesCr002
    Inherits PostgreSQLModelAdapter
    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function

    Public Property errmsg As String
    Private criteria As String
    Dim mypath As String = My.Settings.HKAutoReport
    Dim filename As String = "SalesReportHK.xlsx"

    Public Sub New(ByVal criteria As String, ByVal myPath As String, ByVal FileName As String)
        Me.criteria = criteria
        Me.mypath = myPath
        Me.filename = FileName
    End Sub

    Public Function GenerateReport() As Boolean
        Dim myret As Boolean = False

        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        'Dim enddate As Date = Today.Date.AddDays(-1)
        Dim enddate As Date = Today.Date
        Dim hwnd As System.IntPtr
        Dim result As Boolean
        Try
            'Create Object Excel 

            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd

            oXl.Visible = False
            oXl.DisplayAlerts = False

            oWb = oXl.Workbooks.Open(Application.StartupPath & "\templates\ExcelTemplate.xltx")

            Dim counter As Integer = 0
            'ProgressReport(2, "Creating Worksheet...")
            'backOrder
            For i = 0 To 2
                oWb.Worksheets.Add(After:=(oWb.Worksheets(3 + i)))
            Next i

            Dim sqlstr As String = String.Empty
            '
            'Get Filter

            oSheet = oWb.Worksheets(6)
            Dim myfilter As New System.Text.StringBuilder
            sqlstr = String.Format("with cmmf as (" &
                    " select distinct cmmf,first_value(materialdesc) over (partition by cmmf order by invdate desc,cmmf,materialdesc  )as materialdesc  from sales.tx) " &
                    " (select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then cp.ckw else cp.sda end as salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,qty as qty" & enddate.Year - 1 & ",totalsales as totalsales" & enddate.Year - 1 & ",null::numeric as totalcost" & enddate.Year - 1 & ",null::integer as qty" & enddate.Year & ",null::numeric(13,2) as totalsales" & enddate.Year & ",null::numeric(15,5) as totalcost" & enddate.Year & ",qty as totalqty ,totalsales as totalsales,null::numeric as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                    ",rp.retailprice,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then  null else (1 - (totalsales/qty) / rp.retailprice)  end as sda,case brand when 'LAGOSTINA' then (1 - (totalsales/qty) / rp.retailprice)  when ('LAGOSTINA CASA')  then (1 - (totalsales/qty) / rp.retailprice)  end as lagocookware ,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then( case brand when 'TEFAL' then (1 - ((totalsales/qty) / (rp.retailprice * 0.7))) end )end as tefalcookwaretradediscount,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then( case brand when 'TEFAL' then (1 - ((totalsales/qty) / (rp.retailprice ))) end )end as tefalcookwaredirectdiscount" &
                    " ,cm.series,cm.range,cm.inductionproperty,cm.type,cm.size,cm.extmaterial,cm.intmaterial" &
                    " from sales.tx " &
                    " left join sales.customer c on c.customerid = tx.customerid " &
                    " left join sales.custprodkam cp on cp.customerid = tx.customerid" &
                    " left join cmmf on cmmf.cmmf = tx.cmmf " &
                    " left join sales.cmmfinfo cm on cm.cmmf = tx.cmmf" &
                    " left join sales.hkretailprice rp on rp.cmmf = tx.cmmf " &
                    " where invdate >= '{0:yyyy-MM-dd}' and invdate <= '{1:yyyy-MM-dd}' and tx.qty <> 0 {4} order by invdate) union all " &
                    "(select invid,invdate,orderno,tx.customerid,c.customername,reportcode,saleforce,country,custtype,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then cp.ckw else cp.sda end as salesman,shipto,productid,tx.cmmf,sbu,productfamily,brand,cmmf.materialdesc,supplierid,null::integer,null::numeric(13,2),null::numeric(15,5),qty,totalsales ,null::numeric as totalcost,qty as totalqty ,totalsales as totalsales,null::numeric as totalcost,region,location,invdate as filterdate1,invdate as filterdate2" &
                    ",rp.retailprice,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then  null else (1 - (totalsales/qty) / rp.retailprice)  end as sda,case brand when 'LAGOSTINA' then (1 - (totalsales/qty) / rp.retailprice)  when ('LAGOSTINA CASA')  then (1 - (totalsales/qty) / rp.retailprice)  end as lagocookware ,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then( case brand when 'TEFAL' then (1 - ((totalsales/qty) / (rp.retailprice * 0.7))) end )end as tefalcookwaretradediscount,case when  (sbu = 'COOKWARE & BAKEWARE' or sbu = 'KITCHENWARE & DINNER' or sbu = 'COOKWARE' ) then( case brand when 'TEFAL' then (1 - ((totalsales/qty) / (rp.retailprice ))) end )end as tefalcookwaredirectdiscount" &
                    " ,cm.series,cm.range,cm.inductionproperty,cm.type,cm.size,cm.extmaterial,cm.intmaterial" &
                    " from sales.tx " &
                     " left join sales.customer c on c.customerid = tx.customerid " &
                     " left join sales.custprodkam cp on cp.customerid = tx.customerid" &
                     " left join cmmf on cmmf.cmmf = tx.cmmf " &
                     " left join sales.cmmfinfo cm on cm.cmmf = tx.cmmf" &
                      " left join sales.hkretailprice rp on rp.cmmf = tx.cmmf  " &
                    " where invdate >= '{2:yyyy-MM-dd}' and invdate <= '{3:yyyy-MM-dd}' and tx.qty <> 0 {4} order by invdate)", CDate(enddate.Year - 1 & "-01-01"), CDate(enddate.Year - 1 & "-12-31"), CDate(enddate.Year & "-1-1"), enddate, criteria)


            oSheet.Name = "DATA"

            FillWorksheet(oSheet, sqlstr, dbAdapter1)
            Dim lastrow = oSheet.Cells.Find(What:="*", SearchDirection:=Excel.XlSearchDirection.xlPrevious, SearchOrder:=Excel.XlSearchOrder.xlByRows).Row

            If lastrow > 1 Then
                '
                ApplyFormat(oSheet)
                'oXl.Visible = True
                CreatePivotTable(oXl, oWb, 1, enddate)
                'createchart(oWb, 1, errmsg)
            End If

            'remove connection
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            'Stopwatch.Stop()
            'Filename = ValidateFileName(Filename, Filename & "\" & String.Format("Sales-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day)))
            filename = String.Format("{0}{1}", mypath, filename)


            oWb.SaveAs(filename)
            result = True

        Catch ex As Exception
            errmsg = ex.Message
        Finally
            'clear excel from memory
            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
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

        Return myret
    End Function

    Public Shared Sub FillWorksheet(ByVal osheet As Excel.Worksheet, ByVal sqlstr As String, ByVal dbAdapter As Object, Optional ByVal Location As String = "A1")
        'Dim oRange As Excel.Range
        Dim oExCon As String = My.Settings.oExcon '"ODBC;DSN=PostgreSQLhon03nt;"
        oExCon = oExCon.Insert(oExCon.Length, "UID=admin;Pwd=admin")
        Dim oRange As Excel.Range
        oRange = osheet.Range(Location)
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)
            'With osheet.QueryTables.Add(oExCon, osheet.Range("A1"))
            .CommandText = sqlstr
            .FieldNames = True
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = Excel.XlCellInsertionMode.xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh(BackgroundQuery:=False)
            Application.DoEvents()
        End With

        oRange = Nothing
    End Sub

    Private Sub CreatePivotTable(ByVal oxl As Excel.Application, ByVal oWb As Excel.Workbook, ByVal isheet As Integer, ByVal mydate As Date)
        Dim osheet As Excel.Worksheet

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.PivotCaches.Create(Excel.XlPivotTableSourceType.xlDatabase, "DATA!ExternalData_1").CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").calculatedfields.add("qtydif", "=qty" & mydate.Year & " - qty" & mydate.Year - 1, True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("qtydifpct", "=qty" & mydate.Year & " / qty" & mydate.Year - 1 & " - 1", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("amountdif", "=totalsales" & mydate.Year & " - totalsales" & mydate.Year - 1, True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("amountdifpct", "=totalsales" & mydate.Year & " / totalsales" & mydate.Year - 1 & " - 1", True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("margin" & mydate.Year - 1 & "pct", "=(totalsales" & mydate.Year - 1 & " - totalcost" & mydate.Year - 1 & ")/ totalsales" & mydate.Year - 1, True)
        osheet.PivotTables("PivotTable1").calculatedfields.add("margin" & mydate.Year & "pct", "=(totalsales" & mydate.Year & " - totalcost" & mydate.Year & ")/ totalsales" & mydate.Year, True)

        osheet.PivotTables("PivotTable1").Pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("A8").Group(True, True, Periods:={False, False, False, False, True, False, True})
        osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlHidden

        osheet.PivotTables("PivotTable1").Pivotfields("filterdate1").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.Range("A8").Group(True, True, Periods:={False, False, False, False, True, True, True})
        osheet.PivotTables("PivotTable1").pivotfields("Years2").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").pivotfields("Quarters").orientation = Excel.XlPivotFieldOrientation.xlHidden
        osheet.PivotTables("PivotTable1").pivotfields("filterdate1").orientation = Excel.XlPivotFieldOrientation.xlHidden

        osheet.PivotTables("PivotTable1").Pivotfields("location").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlPageField

        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField


        'osheet.PivotTables("PivotTable1").Pivotfields("sda").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("sda").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").Pivotfields("lagocookware").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("lagocookware").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").Pivotfields("tefalcookwaretradediscount").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaretradediscount").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").Pivotfields("tefalcookwaredirectdiscount").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaredirectdiscount").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("retailprice"), " RSP", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" RSP").NumberFormat = "#,##0"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sda"), " SDA", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" SDA").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("lagocookware"), " D. Disc (Lago)", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" D. Disc (Lago)").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaretradediscount"), " T Ckw T.Disc", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" T Ckw T.Disc").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaredirectdiscount"), " T Ckw D. Disc", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" T Ckw D. Disc").NumberFormat = "0.0%"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), " Total Sales " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty " & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), " Totals Sales " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydif"), " Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydifpct"), " %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdif"), " Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdifpct"), "%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year - 1 & "pct"), " %Margin " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year & "pct"), " %Margin " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Total Sales " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year).numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales " & mydate.Year).numberformat = "#,##0.00"


        'osheet.PivotTables("PivotTable1").PivotFields(" Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).NumberFormat = "#,##0"
        'osheet.PivotTables("PivotTable1").PivotFields(" %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        'osheet.PivotTables("PivotTable1").PivotFields(" Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields("%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        'osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year - 1).numberformat = "0.00%"
        'osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year).numberformat = "0.00%"
        'osheet.Columns("C:F").NumberFormat = "0.0%"
        osheet.Columns("C:G").HorizontalAlignment = Excel.Constants.xlRight

        osheet.Name = "YTD"



        osheet.Cells.EntireColumn.AutoFit()

        'Second PivotTable
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("location").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").currentpage = Format(mydate, "MMM")
        'For Each item As Object In osheet.PivotTables("PivotTable2").pivotfields("Years").pivotitems
        '    Dim obj = DirectCast(item, Excel.PivotItem)
        '    If obj.Value.ToString <> mydate.Year.ToString Then
        '        obj.Visible = False
        '    End If
        'Next



        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField

        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("retailprice"), " RSP", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" RSP").NumberFormat = "#,##0"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("sda"), " SDA", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" SDA").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("lagocookware"), " D. Disc (Lago)", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" D. Disc (Lago)").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaretradediscount"), " T Ckw T.Disc", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" T Ckw T.Disc").NumberFormat = "0.0%"
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaredirectdiscount"), " T Ckw D. Disc", Excel.XlConsolidationFunction.xlAverage)
        'osheet.PivotTables("PivotTable1").PivotFields(" T Ckw D. Disc").NumberFormat = "0.0%"

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), " Total Sales " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty " & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), " Totals Sales " & mydate.Year, Excel.XlConsolidationFunction.xlSum)


        'osheet.PivotTables("PivotTable1").Pivotfields("sda").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("sda").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").Pivotfields("lagocookware").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("lagocookware").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").Pivotfields("tefalcookwaretradediscount").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaretradediscount").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        'osheet.PivotTables("PivotTable1").Pivotfields("tefalcookwaredirectdiscount").orientation = Excel.XlPivotFieldOrientation.xlRowField
        'osheet.PivotTables("PivotTable1").PivotFields("tefalcookwaredirectdiscount").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}



        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydif"), " Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydifpct"), " %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdif"), " Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdifpct"), "%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year - 1 & "pct"), " %Margin " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year & "pct"), " %Margin " & mydate.Year, Excel.XlConsolidationFunction.xlSum)


        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Total Sales " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year).numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales " & mydate.Year).numberformat = "#,##0.00"

        'osheet.PivotTables("PivotTable1").PivotFields(" Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).NumberFormat = "#,##0.00"
        'osheet.PivotTables("PivotTable1").PivotFields(" %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        'osheet.PivotTables("PivotTable1").PivotFields(" Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields("%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        'osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year - 1).numberformat = "0.00%"
        'osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year).numberformat = "0.00%"
        osheet.Columns("C:G").HorizontalAlignment = Excel.Constants.xlRight

        osheet.Name = "MTD"

        osheet.Cells.EntireColumn.AutoFit()

        'Third PivotTable
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("salesman").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlPageField


        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("productid").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField



        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), " Total Sales " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty " & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), " Totals Sales " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydif"), " Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qtydifpct"), " %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdif"), " Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("amountdifpct"), "%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year - 1 & "pct"), " %Margin " & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("margin" & mydate.Year & "pct"), " %Margin " & mydate.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Total Sales " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty " & mydate.Year).numberformat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Totals Sales " & mydate.Year).numberformat = "#,##0.00"


        osheet.PivotTables("PivotTable1").PivotFields(" Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).NumberFormat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields(" %Qty Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "#,##0.00"
        osheet.PivotTables("PivotTable1").PivotFields("%Amt Diff " & mydate.Year & " VS " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year - 1).numberformat = "0.00%"
        osheet.PivotTables("PivotTable1").PivotFields(" %Margin " & mydate.Year).numberformat = "0.00%"

        osheet.Name = "Details"

        osheet.Cells.EntireColumn.AutoFit()

        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

        'Qty
        'Fourth Pivot Table
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").Caption = "Filter Years"

        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").caption = "Months"
        'osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("Months").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}

        'Dim mypivot As Excel.PivotItem
        'For Each mypivot In osheet.PivotTables("PivotTable1").pivotfields("Years").PivotItems
        '    mypivot.Value = "Qty " + mypivot.Value
        'Next

        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("productid").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField



        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalqty"), "Total Quantity", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year - 1), " Qty" & mydate.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("qty" & mydate.Year), " Qty" & mydate.Year, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").PivotFields(" Qty" & mydate.Year - 1).NumberFormat = "#,##0"
        osheet.PivotTables("PivotTable1").PivotFields(" Qty" & mydate.Year).NumberFormat = "#,##0"

        osheet.PivotTables("PivotTable1").ShowTableStyleColumnStripes = True
        osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16"
        oWb.TableStyles("PivotStyleLight16").Duplicate("PivotStyleLight16 2" _
    )
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeTop)
            .Weight = 2
            .LineStyle = 1
        End With
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeBottom)
            .Weight = 2
            .LineStyle = 1
        End With
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeLeft)
            .Weight = 2
            .LineStyle = 1
        End With
        With oxl.ActiveWorkbook.TableStyles("PivotStyleLight16 2").TableStyleElements(Excel.XlTableStyleElementType.xlColumnSubheading1).Borders(Excel.XlBordersIndex.xlEdgeRight)
            .Weight = 2
            .LineStyle = 1
        End With
        osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16 2"


        osheet.Name = "Quantity"

        osheet.Cells.EntireColumn.AutoFit()

        'Sales
        'Fifth Pivot Table
        isheet = isheet + 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)
        oWb.Worksheets("YTD").PivotTables("PivotTable1").PivotCache.CreatePivotTable(osheet.Name & "!R7C1", "PivotTable1", Excel.XlPivotTableVersionList.xlPivotTableVersionCurrent)
        With osheet.PivotTables("PivotTable1")
            .ingriddropzones = True
            .RowAxisLayout(Excel.XlLayoutRowType.xlTabularRow)
            .DisplayErrorString = True
        End With

        osheet.PivotTables("PivotTable1").Pivotfields("sbu").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").Pivotfields("customername").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").orientation = Excel.XlPivotFieldOrientation.xlPageField
        osheet.PivotTables("PivotTable1").pivotfields("Years2").Caption = "Filter Years"

        osheet.PivotTables("PivotTable1").pivotfields("invdate").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").pivotfields("invdate").caption = "Months"
        'osheet.PivotTables("PivotTable1").pivotfields("Years").orientation = Excel.XlPivotFieldOrientation.xlColumnField
        osheet.PivotTables("PivotTable1").PivotFields("Months").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}


        'For Each mypivot In osheet.PivotTables("PivotTable1").pivotfields("Years").PivotItems
        '    mypivot.Value = "Sales Amt " + mypivot.Value
        'Next

        osheet.PivotTables("PivotTable1").Pivotfields("productfamily").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("brand").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").Pivotfields("productid").orientation = Excel.XlPivotFieldOrientation.xlRowField
        osheet.PivotTables("PivotTable1").PivotFields("productid").Subtotals = {False, False, False, False, False, False, False, False, False, False, False, False}
        osheet.PivotTables("PivotTable1").Pivotfields("materialdesc").orientation = Excel.XlPivotFieldOrientation.xlRowField



        'osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales"), "Sales Amount", Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year - 1), "Sales Amount" & Today.Date.Year - 1, Excel.XlConsolidationFunction.xlSum)
        osheet.PivotTables("PivotTable1").AddDataField(osheet.PivotTables("PivotTable1").PivotFields("totalsales" & mydate.Year), "Sales Amount" & Today.Date.Year, Excel.XlConsolidationFunction.xlSum)

        osheet.PivotTables("PivotTable1").PivotFields("Sales Amount" & mydate.Year - 1).NumberFormat = "#,##0,00"
        osheet.PivotTables("PivotTable1").PivotFields("Sales Amount" & mydate.Year).NumberFormat = "#,##0,00"

        osheet.PivotTables("PivotTable1").ShowTableStyleColumnStripes = True
        osheet.PivotTables("PivotTable1").TableStyle2 = "PivotStyleLight16 2"
        osheet.Name = "Sales Amt"

        osheet.Cells.EntireColumn.AutoFit()
        isheet = 1
        oWb.Worksheets(isheet).select()
        osheet = oWb.Worksheets(isheet)

    End Sub

    Public Shared Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub

    Private Sub ApplyFormat(oSheet As Excel.Worksheet)
        oSheet.Columns("AG:AJ").NumberFormat = "0.0%"
    End Sub

End Class
