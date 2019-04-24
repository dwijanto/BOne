Imports System.Threading
Public Enum AUTOGENERATE
    AUTO = 0
    NON_AUTO = 1
End Enum

Public Class FormAutoGenerate
    Public dbAdapter1 As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private AutoGenerate As AUTOGENERATE
    Dim myThread As New System.Threading.Thread(AddressOf doWork)
    Private Sub FormAutoGenerate_Load(sender As Object, e As EventArgs) Handles Me.Load
        If AutoGenerate = BusinessOne.AUTOGENERATE.AUTO Then
            Me.WindowState = FormWindowState.Minimized
            LoadMe()
        End If
    End Sub

    Private Sub LoadMe()
        Label1.Text = String.Format("Server = {0};", My.Settings.host)
        If Not myThread.IsAlive Then
            Try
                ToolStripStatusLabel1.Text = ""
                myThread = New System.Threading.Thread(AddressOf doWork)
                myThread.TrySetApartmentState(ApartmentState.MTA)
                myThread.Start()
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
    End Sub

    Sub doWork()
        Logger.log("--------Start----------")

        ProgressReport(6, "Start")
        'Get Last Running Task
        Dim lastUpdate As Date = dbAdapter1.getLastImportDate

        'HONG KONG Report
        Logger.log("Generate HK Report.")


        Dim myHKReport = New HKReport(Me)
        'myHKReport.startdate = lastUpdate.AddDays(1) 'start date after last import date
        myHKReport.startdate = lastUpdate 'start date after last import date
        ProgressReport(6, "Start")
        ProgressReport(2, "1/15")
        Logger.log("SalesExtract")
        If Not myHKReport.SalesExtract Then
            Logger.log(myHKReport.ErrorMessage)
        End If

        ProgressReport(6, "Start")
        ProgressReport(2, "2A/15")
        'Dim mySalesReportHKCorrected = New ReportSales2016Corrected
        'If Not mySalesReportHKCorrected.GenerateReport Then
        '    Logger.log(mySalesReportHKCorrected.errmsg)
        'End If

        Logger.log("SalesReportHK001")
        Dim mySalesReportHK = New ReportSales("SalesReportHK001.xlsx", True)
        If Not mySalesReportHK.GenerateReport Then
            Logger.log(mySalesReportHK.errmsg)
        End If


        ProgressReport(6, "Start")
        ProgressReport(2, "2B/15")
        Logger.log("SalesReportHK")
        mySalesReportHK = New ReportSales
        If Not mySalesReportHK.GenerateReport Then
            Logger.log(mySalesReportHK.errmsg)
        End If

        ProgressReport(6, "Start")
        ProgressReport(2, "3/15")
        Logger.log("CookwareSalesReportHK.")
        Dim mypath As String = My.Settings.HKAutoReport
        Dim filename As String = "CookwareSalesReportHK.xlsx"
        Dim criteria As String = " and sbu in('COOKWARE & BAKEWARE','KITCHENWARE & DINNER','COOKWARE')"
        Dim mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        ProgressReport(6, "Start")
        ProgressReport(2, "4/15")
        Logger.log("SDASalesReportHK.")
        mypath = My.Settings.HKAutoReport
        filename = "SDASalesReportHK.xlsx"
        criteria = " and not sbu in ('COOKWARE & BAKEWARE','KITCHENWARE & DINNER','COOKWARE')"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        ProgressReport(6, "Start")
        ProgressReport(2, "5/15")
        Logger.log("AntonioSalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\asin\Report\" 'My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\asin\"
        filename = "AntonioSalesReportHK.xlsx"
        criteria = " and sbu in ('COOKWARE & BAKEWARE','KITCHENWARE & DINNER','COOKWARE') and salesman not in ('Singapore','Philippines')" ' " and salesman = 'Antonio'"
        Dim mySBUCookware0 = New ReportSalesCr002(criteria, mypath, filename)
        'If (Weekday(Today.Date, FirstDayOfWeek.Monday) = 1) Or Today.Date.Day = 1 Then
        If Not mySBUCookware0.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        'End If
        ProgressReport(6, "Start")
        ProgressReport(2, "6/15")
        Logger.log("CatherineSalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\cchan\Report\" 'My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\cchan\"
        filename = "CatherineSalesReportHK.xlsx"
        criteria = " and salesman not in ('Singapore','Philippines')" '" and salesman = 'Catherine'"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        'If (Weekday(Today.Date, FirstDayOfWeek.Monday) = 1) Or Today.Date.Day = 1 Then
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        'End If
        ProgressReport(6, "Start")
        ProgressReport(2, "7/15")
        Logger.log("FelixSalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\fewong\Report\" ' My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\fewong\"
        filename = "FelixSalesReportHK.xlsx"
        criteria = " and sbu in('COOKWARE & BAKEWARE','KITCHENWARE & DINNER','COOKWARE') and salesman not in ('Singapore','Philippines')" '" and salesman = 'Jack'"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        'If (Weekday(Today.Date, FirstDayOfWeek.Monday) = 1) Or Today.Date.Day = 1 Then
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        'End If
        ProgressReport(6, "Start")
        ProgressReport(2, "8/15")
        Logger.log("JoeSalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\jlo\Report\" 'My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\jlo\"
        filename = "JoeSalesReportHK.xlsx"
        criteria = " and salesman not in ('Singapore','Philippines')" '" and salesman = 'Joe'"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        'If (Weekday(Today.Date, FirstDayOfWeek.Monday) = 1) Or Today.Date.Day = 1 Then
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        'End If
        ProgressReport(6, "Start")
        ProgressReport(2, "9/15")
        Logger.log("SauYingSalesReportHK.")
        ' mypath = "\\172.22.10.34\RicohScanner1\sytsui\Report\" ' My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\sytsui\"
        filename = "SauYingSalesReportHK.xlsx"
        criteria = " and salesman not in ('Singapore','Philippines')" '" and salesman = 'Sau Ying'"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        'If (Weekday(Today.Date, FirstDayOfWeek.Monday) = 1) Or Today.Date.Day = 1 Then
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        'End If
        ProgressReport(6, "Start")
        ProgressReport(2, "10/15")
        Logger.log("WaiKitSalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\wkchan\Report\" 'My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\wkchan\"
        filename = "WaiKitSalesReportHK.xlsx"
        criteria = " and salesman not in ('Singapore','Philippines')" '" and salesman = 'Wai Kit'"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        'If (Weekday(Today.Date, FirstDayOfWeek.Monday) = 1) Or Today.Date.Day = 1 Then
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        'End If
        ProgressReport(6, "Start")
        ProgressReport(2, "11/15")
        Logger.log("KilySalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\kilai\Report\" ' My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\kilai\"
        filename = "KilySalesReportHK.xlsx"
        criteria = " and salesman not in ('Singapore','Philippines')"
        mySBUCookware = New ReportSalesCr(criteria, mypath, filename)
        If Not mySBUCookware.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        ProgressReport(6, "Start")
        ProgressReport(2, "12/15")
        Logger.log("MarketingSalesReportHK.")
        'mypath = "\\172.22.10.34\RicohScanner1\elam\Report\" 'My.Settings.HKAutoReport
        mypath = "\\SW07E601\report\elam\"
        filename = "MarketingSalesReportHK.xlsx"
        criteria = " and salesman not in ('Singapore','Philippines')"
        Dim mySBUCookware1 = New ReportSalesMarketing(mypath, filename)
        If Not mySBUCookware1.GenerateReport Then
            Logger.log(mySBUCookware.errmsg)
        End If
        ProgressReport(6, "Start")
        ProgressReport(2, "13/15")
        'TAIWAN REPORT
        Logger.log("Generate Taiwan Report.")
        'Email For CC Status Completed
        Dim myTWReport = New TWReport(Me)
        myTWReport.startdate = lastUpdate.AddDays(1) 'start date after last import date
        If Not myTWReport.SalesExtract Then
            Logger.log(myTWReport.ErrorMessage)
        End If

        ProgressReport(6, "Start")
        ProgressReport(2, "14/15")
        Dim mySalesReportTW = New ReportSalesTW
        If Not mySalesReportTW.GenerateReport Then
            Logger.log(mySalesReportTW.errmsg)
        End If

        ProgressReport(6, "Start")
        ProgressReport(2, "15/15")

        Dim myReportTWSalesReport = New ReportTWSalesReport
        If Not myReportTWSalesReport.GenerateReport Then
            Logger.log(myReportTWSalesReport.errmsg)
        End If

        'Update LastRunning Task

        dbAdapter1.setLastImportDate(Today.Date.AddDays(-1))
        ProgressReport(2, "Done.")
        Logger.log("--------End------------")

        ProgressReport(5, "End")
        ProgressReport(1, "Close Apps")
    End Sub

    Public Sub ProgressReport(ByVal id As Integer, ByVal message As String)
        If Me.InvokeRequired Then
            Dim d As New ProgressReportDelegate(AddressOf ProgressReport)
            Me.Invoke(d, New Object() {id, message})
        Else
            Select Case id
                Case 1
                    Me.Close()
                Case 2
                    ToolStripStatusLabel1.Text = message
                Case 3

                Case 4

                Case 5
                    ToolStripProgressBar1.Style = ProgressBarStyle.Continuous
                Case 6
                    ToolStripProgressBar1.Style = ProgressBarStyle.Marquee
                Case 7

            End Select
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        LoadMe()
    End Sub

    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        AutoGenerate = BusinessOne.AUTOGENERATE.AUTO
        ' Add any initialization after the InitializeComponent() call.

    End Sub
    Public Sub New(status As AUTOGENERATE)

        ' This call is required by the designer.
        InitializeComponent()
        AutoGenerate = status
        ' Add any initialization after the InitializeComponent() call.

    End Sub
End Class