Imports Microsoft.Office.Interop
Imports System.Threading
Imports System.Text
Imports System.Runtime.InteropServices

Public Delegate Sub ProgressReportDelegate(ByVal id As Integer, ByVal message As String)
Public Delegate Sub FormatReportDelegate(ByRef sender As Object, ByRef e As EventArgs)
Public Class ExcelExtract
    Private DT As DataTable
    Dim myThread As New Threading.Thread(AddressOf DoWork)
    Private status As Boolean
    Public Property Directory As String
    Private Dataset1 As New DataSet
    Private Parent As Object
    Public Property Datasheet As Integer = 1
    Public Property ReportName As String
    Public Property mytemplate As String = "\templates\ExcelTemplate.xltx"
    Public Property FormatReportCallback As FormatReportDelegate
    Dim PivotCallback As FormatReportDelegate = AddressOf PivotTable

    Public Property DataTableList As List(Of DataTableWorksheet)
    Private OpenFile As Boolean

    <DllImport("user32.dll")> _
    Public Shared Function EndTask(ByVal hWnd As IntPtr, ByVal fShutDown As Boolean, ByVal fForce As Boolean) As Boolean
    End Function

    Public Sub New(ByRef Parent As Object, ByVal ReportName As String, ByVal datasheet As Integer, ByRef myTemplate As String, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, Optional openfile As Boolean = True)
        Me.Parent = Parent
        Me.ReportName = ReportName
        Me.mytemplate = myTemplate
        Me.Datasheet = Datasheet
        Me.FormatReportCallback = FormatReportCallBack
        Me.PivotCallback = PivotCallback
    End Sub

    Public Sub New(ByRef Parent As Object, ByVal ReportName As String, ByRef myTemplate As String, ByVal dt As DataTable, ByVal FormatReportCallBack As FormatReportDelegate, ByVal PivotCallback As FormatReportDelegate, Optional openfile As Boolean = True)
        Me.Parent = Parent
        Me.ReportName = ReportName
        Me.mytemplate = myTemplate
        Me.DT = dt
        Me.FormatReportCallback = FormatReportCallBack
        Me.OpenFile = openfile
    End Sub

    Public Sub ExtractFromDataTable(ByRef sender As Object, ByVal e As System.EventArgs)
        If Not myThread.IsAlive Then
            Try
                myThread = New System.Threading.Thread(New ThreadStart(AddressOf DoExtractFromDataTable))
                myThread.TrySetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        Else
            MsgBox("Please wait until the current process is finished")
        End If

    End Sub

    Public Sub ExtractFromDataTableUnsync(ByRef sender As Object, ByVal e As System.EventArgs)
        Try
             DoExtractFromDataTable
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try
      

    End Sub
    Public Sub ExtractFromDataTableUnsyncDT(ByRef sender As Object, ByVal e As System.EventArgs)
        Try
            DoExtractFromDataTableDT()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub

    Sub DoWork()
        Throw New NotImplementedException
    End Sub
    Private Sub DoExtractFromDataTableDT()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReportDT(Directory, errMsg)
        ProgressReport(5, "Continues..")
        If status Then


            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")
            If OpenFile Then
                If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(Directory)
                End If
            End If
            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()
    End Sub
    Private Sub DoExtractFromDataTable()
        Dim errMsg As String = String.Empty
        Dim i As Integer = 0
        Dim errSB As New StringBuilder
        Dim sw As New Stopwatch
        sw.Start()
        ProgressReport(2, "Export To Excel..")
        ProgressReport(6, "Marques..")
        status = GenerateReport(Directory, errMsg)
        ProgressReport(5, "Continues..")
        If status Then


            sw.Stop()
            ProgressReport(2, String.Format("Elapsed Time: {0}:{1}.{2} Done.", Format(sw.Elapsed.Minutes, "00"), Format(sw.Elapsed.Seconds, "00"), sw.Elapsed.Milliseconds.ToString))
            ProgressReport(3, "")
            If OpenFile Then
                If MsgBox("File name: " & Directory & vbCr & vbCr & "Open the file?", vbYesNo, "Export To Excel") = DialogResult.Yes Then
                    Process.Start(Directory)
                End If
            End If
            ProgressReport(3, "")
            'ProgressReport(4, errSB.ToString)
        Else
            errSB.Append(errMsg) '& vbCrLf)
            ProgressReport(3, errSB.ToString)
        End If
        sw.Stop()
    End Sub
    Private Function GenerateReportDT(ByRef FileName As String, ByRef errorMsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            oSheet = oWb.Worksheets(1)
            oSheet.Name = "RAWDATA"
            oWb.SaveAs(ReportName)
            'For i = 0 To oWb.Worksheets.Count - 1
            '    oWb.Sheets(1).Delete()
            'Next
            oXl.Visible = False

            oXl.Quit()
            releaseComObject(oSheet)
            releaseComObject(oWb)
            releaseComObject(oXl)
            GC.Collect()
            GC.WaitForPendingFinalizers()
            EndTask(hwnd, True, True)



            'ProgressReport(2, "Creating Worksheet...")
            'DATA

            If IsNothing(DataTableList) Then
                'Print One Worksheet 
                'oWb.Worksheets(Datasheet).select()
                'oSheet = oWb.Worksheets(Datasheet)

                If Me.DT.Rows.Count > 0 Then
                    ProgressReport(2, "Get records..")
                    'FillWorksheet(oSheet, DT)
                    FillWorksheetDT(ReportName, DT)
                    'Dim orange = oSheet.Range("A1")
                    'Dim lastrow = GetLastRow(oXl, oSheet, orange)
                    'If lastrow > 1 Then
                    ' FormatReportCallback.Invoke(oSheet, New EventArgs)
                    'End If
                Else
                    'FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Print Multiple Worksheet
                'Looping from here
                For i = 0 To DataTableList.Count - 1
                    Dim MyDataTable = CType(DataTableList(i), DataTableWorksheet)
                    oWb.Worksheets(MyDataTable.DataSheet).select()
                    oSheet = oWb.Worksheets(MyDataTable.DataSheet)
                    oSheet.Name = MyDataTable.SheetName
                    ProgressReport(2, "Get records..")
                    'FillWorksheet(oSheet, MyDataTable.DataTable)
                    FillWorksheetDT(ReportName, MyDataTable.DataTable)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Next



                'End Looping
            End If


            PivotCallback.Invoke(oWb, New EventArgs)
            'For i = 0 To oWb.Connections.Count - 1
            '    oWb.Connections(1).Delete()
            'Next
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            'If FileName.Contains("xlsm") Then
            '    oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            'Else
            '    oWb.SaveAs(FileName)
            'End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'clear excel from memory
            Try
                'oXl.Quit()
                'releaseComObject(oSheet)
                'releaseComObject(oWb)
                'releaseComObject(oXl)
                'GC.Collect()
                'GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                'EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function
    Private Function GenerateReport(ByRef FileName As String, ByRef errorMsg As String) As Boolean
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)            
            oXl.Visible = False
            ProgressReport(2, "Creating Worksheet...")
            'DATA

            If IsNothing(DataTableList) Then
                'Print One Worksheet 
                oWb.Worksheets(Datasheet).select()
                oSheet = oWb.Worksheets(Datasheet)

                If Me.DT.Rows.Count > 0 Then
                    ProgressReport(2, "Get records..")
                    FillWorksheet(oSheet, DT)
                    'FillWorksheetDT(ReportName, DT)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)
                    If lastrow > 1 Then
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Else
                    FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Print Multiple Worksheet
                'Looping from here
                For i = 0 To DataTableList.Count - 1
                    Dim MyDataTable = CType(DataTableList(i), DataTableWorksheet)
                    oWb.Worksheets(MyDataTable.DataSheet).select()
                    oSheet = oWb.Worksheets(MyDataTable.DataSheet)
                    oSheet.Name = MyDataTable.SheetName
                    ProgressReport(2, "Get records..")
                    FillWorksheet(oSheet, MyDataTable.DataTable)
                    'FillWorksheetDT(oSheet, MyDataTable.DataTable)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Next



                'End Looping
            End If


            PivotCallback.Invoke(oWb, New EventArgs)
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
            Next
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            If FileName.Contains("xlsm") Then
                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
        Finally
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Return result
    End Function
    Public Function GenerateReport(ByRef FileName As String, ByVal sqlstr As String, ByRef errorMsg As String) As Boolean
        Logger.log("Generate Report.")
        Dim myCriteria As String = String.Empty
        Dim result As Boolean = False

        Dim StopWatch As New Stopwatch
        StopWatch.Start()
        'Open Excel
        Application.DoEvents()

        'Excel Variable
        Dim oXl As Excel.Application = Nothing
        Dim oWb As Excel.Workbook = Nothing
        Dim oSheet As Excel.Worksheet = Nothing
        Dim SheetName As String = vbEmpty
        Dim hwnd As System.IntPtr
        Try
            'Create Object Excel 
            ProgressReport(2, "CreateObject..")
            oXl = CType(CreateObject("Excel.Application"), Excel.Application)
            hwnd = oXl.Hwnd
            'oXl.ScreenUpdating = False
            'oXl.Visible = False
            oXl.DisplayAlerts = False
            ProgressReport(2, "Opening Template...")
            ProgressReport(2, "Generating records..")
            Logger.log(String.Format("Open Template : {0}", Application.StartupPath & mytemplate))
            oWb = oXl.Workbooks.Open(Application.StartupPath & mytemplate)
            oXl.Visible = False
            ProgressReport(2, "Creating Worksheet...")
            'DATA

            If IsNothing(DataTableList) Then
                'Print One Worksheet 
                oWb.Worksheets(Datasheet).select()
                oSheet = oWb.Worksheets(Datasheet)

                If sqlstr <> "" Then
                    ProgressReport(2, "Get records..")
                    FillWorksheet(oSheet, sqlstr)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)
                    If lastrow > 1 Then
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Else
                    FormatReportCallback.Invoke(oSheet, New EventArgs)
                End If
            Else
                'Print Multiple Worksheet
                'Looping from here
                For i = 0 To DataTableList.Count - 1
                    Dim MyDataTable = CType(DataTableList(i), DataTableWorksheet)
                    oWb.Worksheets(MyDataTable.DataSheet).select()
                    oSheet = oWb.Worksheets(MyDataTable.DataSheet)
                    oSheet.Name = MyDataTable.SheetName
                    ProgressReport(2, "Get records..")
                    FillWorksheet(oSheet, sqlstr)
                    Dim orange = oSheet.Range("A1")
                    Dim lastrow = GetLastRow(oXl, oSheet, orange)


                    If lastrow > 1 Then
                        'Delegate for modification
                        'oSheet.Columns("A:A").numberformat = "dd-MMM-yyyy"
                        FormatReportCallback.Invoke(oSheet, New EventArgs)
                    End If
                Next



                'End Looping
            End If


            PivotCallback.Invoke(oWb, New EventArgs)
            For i = 0 To oWb.Connections.Count - 1
                oWb.Connections(1).Delete()
                Thread.Sleep(200)
            Next
            StopWatch.Stop()

            'FileName = FileName & "\" & String.Format("Report" & ReportName & "-{0}-{1}-{2}.xlsx", Today.Year, Format("00", Today.Month), Format("00", Today.Day))
            FileName = String.Format(ReportName)
            ProgressReport(3, "")
            ProgressReport(2, "Saving File ..." & FileName)
            'oSheet.Name = ReportName
            Logger.log(String.Format("Saving File...{0}", FileName))
            If FileName.Contains("xlsm") Then
                oWb.SaveAs(FileName, FileFormat:=Excel.XlFileFormat.xlOpenXMLWorkbookMacroEnabled)
            Else
                oWb.SaveAs(FileName)
            End If

            ProgressReport(2, "Elapsed Time: " & Format(StopWatch.Elapsed.Minutes, "00") & ":" & Format(StopWatch.Elapsed.Seconds, "00") & "." & StopWatch.Elapsed.Milliseconds.ToString)
            result = True
        Catch ex As Exception
            ProgressReport(3, ex.Message & FileName)
            errorMsg = ex.Message
            Logger.log(String.Format("Error : {0}", errorMsg))
        Finally
            'clear excel from memory
            Try
                oXl.Quit()
                releaseComObject(oSheet)
                releaseComObject(oWb)
                releaseComObject(oXl)
                GC.Collect()
                GC.WaitForPendingFinalizers()
            Catch ex As Exception

            End Try

            Try
                'to make sure excel is no longer in memory
                EndTask(hwnd, True, True)
            Catch ex As Exception
            End Try

        End Try
        Logger.log(String.Format("End {0}", result))
        Return result
    End Function
    Private Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Parent.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Try
                Parent.Invoke(d, New Object() {id, message})
            Catch ex As Exception

            End Try

        Else
            Select Case id
                Case 2
                    Parent.ToolStripStatusLabel1.Text = message
                Case 3
                    Parent.ToolStripStatusLabel2.Text = Trim(message)
                Case 4
                    Parent.close()
                Case 5
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    Parent.ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
            End Select

        End If

    End Sub

    Public Sub releaseComObject(ByRef o As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(o)
        Catch ex As Exception
        Finally
            o = Nothing
        End Try
    End Sub

    Public Shared Sub FillWorksheet(ByVal osheet As Excel.Worksheet, ByVal sqlstr As String, Optional ByVal Location As String = "A1")
        Dim oExCon As String = My.Settings.oExcon ' My.Settings.oExCon.ToString '"ODBC;DSN=PostgreSQL30;"
        'oExCon = oExCon.Insert(oExCon.Length, "UID=" & dbadapter1.userid & ";PWD=" & dbadapter1.password)
        'Dim dbAdapter1 = PostgreSQLDBAdapter.getInstance
        oExCon = oExCon.Insert(oExCon.Length, "UID=admin;PWD=admin;")
        Dim oRange As Excel.Range
        oRange = osheet.Range(Location)
        With osheet.QueryTables.Add(oExCon.Replace("Host=", "Server="), oRange)          
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

        oRange = osheet.Range("1:1")
        oRange = osheet.Range(Location)
        oRange.Select()
        osheet.Application.Selection.autofilter()
        osheet.Cells.EntireColumn.AutoFit()
    End Sub

    Private Sub FillWorksheet(ByVal osheet As Excel.Worksheet, DT As DataTable)
        Dim dc As System.Data.DataColumn
        Dim dr As System.Data.DataRow
        Dim colIndex As Integer = 0
        Dim rowIndex As Integer = 0

        For Each dc In DT.Columns
            colIndex = colIndex + 1
            Select Case dc.DataType.Name
                Case "String"
                    osheet.Columns(colIndex).numberformat = "@"
                Case "DateTime"
                    osheet.Columns(colIndex).numberformat = "dd-MMM-yyyy"
            End Select
           
            osheet.Cells(1, colIndex) = dc.ColumnName
        Next
        For Each dr In DT.Rows
            rowIndex = rowIndex + 1
            colIndex = 0
            For Each dc In DT.Columns
                colIndex = colIndex + 1
                osheet.Cells(rowIndex + 1, colIndex) = dr(dc.ColumnName)
            Next
        Next
        osheet.Cells.EntireColumn.AutoFit()
    End Sub


    Public Shared Function GetLastRow(ByVal oxl As Excel.Application, ByVal osheet As Excel.Worksheet, ByVal range As Excel.Range) As Long
        Dim lastrow As Long = 1
        oxl.ScreenUpdating = False
        Try
            lastrow = osheet.Cells.Find("*", range, , , Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlPrevious).Row
        Catch ex As Exception
        End Try
        Return lastrow
        oxl.ScreenUpdating = True
    End Function

    Sub PivotTable()

    End Sub

    'Private Sub FillWorksheetDT(oSheet As Excel.Worksheet, dataTable As DataTable)
    Private Sub FillWorksheetDT(Filename As String, dataTable As DataTable)
        Dim driver As String = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
        Dim ConStr As String = String.Format("Driver={0};Dbq={1};ReadOnly=0;", driver, Filename)

        Using conn As New Odbc.OdbcConnection(ConStr)
            conn.Open()
            'Create Table Query
            Dim strTableQ As String
            Dim colSB As New System.Text.StringBuilder
            'Dim j As Integer = 0
            For j = 0 To dataTable.Columns.Count - 1
                Dim dCol As DataColumn
                dCol = dataTable.Columns(j)
                Dim coltype As String = String.Empty
                Select Case dCol.DataType.Name
                    Case "String"
                        coltype = "TEXT"
                    Case "DateTime"
                        coltype = "DATE"
                    Case "Int32"
                        coltype = "int"
                    Case "Decimal"
                        coltype = "real"
                    Case Else
                End Select
                If colSB.Length > 0 Then
                    colSB.Append(",")
                End If
                colSB.Append(String.Format("[{0}] {1}", dCol.ColumnName, coltype))
            Next
            Dim mysheetname = "RAWDATA"
            'strTableQ = String.Format("CREATE TABLE [Sheet1$]({0})", colSB.ToString)
            strTableQ = String.Format("CREATE TABLE [{0}$]({1})", mysheetname, colSB.ToString)

            Using cmd As New Odbc.OdbcCommand(strTableQ, conn)
                cmd.ExecuteNonQuery()
            End Using

            'Insert Query
            Dim sbInsert As New StringBuilder
            Dim strInsertQ As String
            For k As Integer = 0 To dataTable.Columns.Count - 1
                If sbInsert.Length > 0 Then
                    sbInsert.Append(",")
                End If
                sbInsert.Append("?")
            Next
            'strInsertQ = String.Format("Insert Into [Sheet1$] Values({0})", sbInsert.ToString)
            strInsertQ = String.Format("Insert Into [{0}$] Values({1})", mysheetname, sbInsert.ToString)

            'Parameters Query
            For j = 0 To dataTable.Rows.Count - 1
                Using cmd = New Odbc.OdbcCommand(strInsertQ, conn)
                    For k As Integer = 0 To dataTable.Columns.Count - 1
                        cmd.Parameters.AddWithValue("?", dataTable.Rows(j)(k))
                        Select Case dataTable.Columns(k).DataType.Name
                            Case "String"
                                cmd.Parameters(k).DbType = DbType.String
                            Case "DateTime"
                                cmd.Parameters(k).DbType = DbType.Date
                            Case "Int32"
                                cmd.Parameters(k).DbType = DbType.Int32
                            Case "Decimal"
                                cmd.Parameters(k).DbType = DbType.Double
                            Case Else
                        End Select
                    Next
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End Using
            Next
        End Using
    End Sub
    Private Sub FillWorksheetDT1(Filename As String, dataTable As DataTable)
        'Using conn As New Odbc.OdbcConnection(String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source='{0}';Extended Properties=""Excel 12.0 Xml;HDR=YES;IMEX=1"";", "C:\Work\vb2013\BOne\BusinessOne\BusinessOne\bin\Debug\Templates\ExcelExport.xlsx"))
        Dim driver As String = "Microsoft Excel Driver (*.xls, *.xlsx, *.xlsm, *.xlsb)"
        Dim ConStr As String = String.Format("Driver={0};Dbq={1};ReadOnly=0;", driver, Filename)

        Dim StrSB As New StringBuilder

        Using conn As New Odbc.OdbcConnection(ConStr)
            conn.Open()
            'making table query

            'Dim strTableQ As String = "CREATE TABLE [" & dataTable.TableName & "]("

            Dim strTableQ As String = "CREATE TABLE [Sheet1$]("

            Dim j As Integer = 0
            For j = 0 To dataTable.Columns.Count - 1
                Dim dCol As DataColumn
                dCol = dataTable.Columns(j)
                Dim coltype As String = String.Empty
                Select Case dCol.DataType.Name
                    Case "String"
                        coltype = "] TEXT , "
                    Case "DateTime"
                        coltype = "] DATE , "

                    Case "Int32"
                        coltype = "] int , "

                    Case "Decimal"
                        coltype = "] real , "

                    Case Else

                End Select

                strTableQ &= " [" & dCol.ColumnName & coltype
            Next
            strTableQ = strTableQ.Substring(0, strTableQ.Length - 2)
            strTableQ &= ")"
            Using cmd As New Odbc.OdbcCommand(strTableQ, conn)
                cmd.ExecuteNonQuery()
            End Using

            ''making insert query
            Dim sbInsert As New StringBuilder
            Dim strInsertQ As String

            'strInsertQ = "Insert Into [" & dataTable.TableName & "] Values ("
            strInsertQ = "Insert Into [Sheet1$] Values ("
            For k As Integer = 0 To dataTable.Columns.Count - 1
                'strInsertQ &= "@" & dataTable.Columns(k).ColumnName & " , "
                If sbInsert.Length > 0 Then
                    sbInsert.Append(",")
                End If
                sbInsert.Append("?")
                'strInsertQ &= "?, "
            Next
            sbInsert.Append(")")
            'strInsertQ = strInsertQ.Substring(0, strInsertQ.Length - 2)
            'strInsertQ &= ")"
            strInsertQ &= sbInsert.ToString


            'Now inserting data
            'For i = 0 To ds.Tables.Count - 1
            For j = 0 To dataTable.Rows.Count - 1
                Using cmd = New Odbc.OdbcCommand(strInsertQ, conn)
                    For k As Integer = 0 To dataTable.Columns.Count - 1
                        'cmd.Parameters.AddWithValue("@" & dataTable.Columns(k).ColumnName.ToString(), dataTable.Rows(j)(k).ToString())

                        cmd.Parameters.AddWithValue("?", dataTable.Rows(j)(k))



                        Select Case dataTable.Columns(k).DataType.Name
                            Case "String"
                                cmd.Parameters(k).DbType = DbType.String
                            Case "DateTime"
                                cmd.Parameters(k).DbType = DbType.Date
                            Case "Int32"
                                cmd.Parameters(k).DbType = DbType.Int32
                            Case "Decimal"
                                cmd.Parameters(k).DbType = DbType.Double
                            Case Else

                        End Select
                    Next
                    cmd.ExecuteNonQuery()
                    cmd.Parameters.Clear()
                End Using

            Next
            'Next
        End Using

    End Sub
End Class

Public Class DataTableWorksheet
    Public Property DataSheet As Integer
    Public Property DataTable As DataTable
    Public Property SheetName As String
End Class