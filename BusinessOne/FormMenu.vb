Public Class FormMenu
    Dim myuser As UserAdapter = New UserAdapter

    Private Sub Button2_Click(sender As Object, e As EventArgs)
        Dim myform = New FormSalesOrder
        myform.Show()
    End Sub


    Private Sub HKSalesExtractionToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HKSalesExtractionToolStripMenuItem.Click
        Dim myform = New FormSalesReport
        myform.Show()

    End Sub



    Private Sub FormMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Label2.Text = username
        Dim mydata = myuser.findByUserName(username.ToLower)
        If mydata.Tables(0).rows.count > 0 Then
            Dim identity = myuser.findIdentity(mydata.Tables(0).rows(0).item("id"))
            User.setIdentity(identity)
            User.login(identity)
            User.IdentityClass = myuser
            Try
                loglogin(username)
            Catch ex As Exception

            End Try
            displayMenuBar()
        Else
            disableMenuBar()
        End If
    End Sub

    Private Sub displayMenuBar()
        Dim identity As UserAdapter = User.getIdentity
        AdminToolStripMenuItem.Visible = User.can("View-Admin") 'identity.isAdmin
        UserToolStripMenuItem.Visible = User.can("createUser")
        QueryToolStripMenuItem.Visible = User.can("View-HKReport")
        QueryTaiwanToolStripMenuItem.Visible = User.can("View-TWReport")
        MLATWToolStripMenuItem.Visible = User.can("View-TWReport")
        MasterToolStripMenuItem.Visible = User.can("View-TWReport") Or User.can("View-HKReport")
        FamilyHKToolStripMenuItem.Visible = User.can("View-HKReport")
        AutoReportToolStripMenuItem.Visible = User.can("Run-AutoReport")
        ImportPOSDataToolStripMenuItem.Visible = User.can("importPostData")
        StaffPurchaseToolStripMenuItem.Visible = User.can("View-StaffPurchase")
        LogisticsToolStripMenuItem.Visible = User.can("View-Logistics")
    End Sub

    Private Sub disableMenuBar()
        AdminToolStripMenuItem.Visible = False
        QueryToolStripMenuItem.Visible = False
        MessageBox.Show(String.Format("You're not authorized to use this function. If not please contact Admin.User id: {0}.", Environment.UserDomainName & "\" & Environment.UserName))
        Me.Close()
    End Sub

    Private Sub UserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UserToolStripMenuItem.Click
        Dim myform = New FormUser
        myform.Show()
    End Sub

    Private Sub TWSalesExtractToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TWSalesExtractToolStripMenuItem.Click
        Dim myform = New FormSalesReportTW
        myform.Show()
    End Sub


    Private Sub MLATWToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MLATWToolStripMenuItem.Click
        Dim myform = New FormMLATW
        myform.Show()
    End Sub

    Private Sub FamilyHKToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FamilyHKToolStripMenuItem.Click
        Dim myform = New FormFamilyHK
        myform.Show()
    End Sub

    Private Sub loglogin(ByVal userid As String)
        Dim applicationname As String = "BOne"
        Dim username As String = Environment.UserDomainName & "\" & Environment.UserName
        Dim computername As String = My.Computer.Name
        Dim time_stamp As DateTime = Now
        myuser.loglogin(applicationname, userid, username, computername, time_stamp)
    End Sub

    Private Sub ImportPOSDataToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ImportPOSDataToolStripMenuItem.Click
        Dim myform As New FormImportPOS
        myform.ShowDialog()
    End Sub



    Private Sub AutoReportToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AutoReportToolStripMenuItem.Click
        Dim myform As New FormAutoGenerate(BusinessOne.AUTOGENERATE.NON_AUTO)
        myform.Show()
    End Sub

    Private Sub ItemPriceMasterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ItemPriceMasterToolStripMenuItem.Click
        Dim myform As New FormItemPrice
        myform.Show()
    End Sub

    Private Sub StaffPurchaseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles StaffPurchaseToolStripMenuItem.Click

    End Sub

    Private Sub POInformationToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles POInformationToolStripMenuItem.Click
        Dim myform = New FormPOInformation
        myform.ShowDialog()
    End Sub

    Private Sub POInformationV2ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles POInformationV2ToolStripMenuItem.Click
        Dim myform = New FormPOInvormationV2
        myform.ShowDialog()
    End Sub

    Private Sub TWInvoiceToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TWInvoiceToolStripMenuItem.Click
        Dim myform = FormTaxInvoice
        myform.ShowDialog()
    End Sub
End Class