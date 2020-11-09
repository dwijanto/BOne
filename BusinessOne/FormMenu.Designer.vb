<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FormMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FormMenu))
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.QueryToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HKSalesExtractionToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AutoReportToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HKWarehouseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QueryTaiwanToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TWSalesExtractToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TWInvoiceToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.TWWarehouseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AdminToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.UserToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ImportPOSDataToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MLATWToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.FamilyHKToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.StaffPurchaseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ItemPriceMasterToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.LogisticsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.POInformationToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.POInformationV2ToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.AutoReportWarehouseToolStripMenuItem1 = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 84)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(482, 22)
        Me.StatusStrip1.TabIndex = 3
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.QueryToolStripMenuItem, Me.QueryTaiwanToolStripMenuItem, Me.AdminToolStripMenuItem, Me.MasterToolStripMenuItem, Me.StaffPurchaseToolStripMenuItem, Me.LogisticsToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(482, 24)
        Me.MenuStrip1.TabIndex = 4
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'QueryToolStripMenuItem
        '
        Me.QueryToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HKSalesExtractionToolStripMenuItem, Me.AutoReportToolStripMenuItem, Me.HKWarehouseToolStripMenuItem})
        Me.QueryToolStripMenuItem.Name = "QueryToolStripMenuItem"
        Me.QueryToolStripMenuItem.Size = New System.Drawing.Size(70, 20)
        Me.QueryToolStripMenuItem.Text = "Query HK"
        '
        'HKSalesExtractionToolStripMenuItem
        '
        Me.HKSalesExtractionToolStripMenuItem.Name = "HKSalesExtractionToolStripMenuItem"
        Me.HKSalesExtractionToolStripMenuItem.Size = New System.Drawing.Size(162, 22)
        Me.HKSalesExtractionToolStripMenuItem.Text = "HK - SalesExtract"
        '
        'AutoReportToolStripMenuItem
        '
        Me.AutoReportToolStripMenuItem.Name = "AutoReportToolStripMenuItem"
        Me.AutoReportToolStripMenuItem.Size = New System.Drawing.Size(162, 22)
        Me.AutoReportToolStripMenuItem.Text = "Auto Report"
        '
        'HKWarehouseToolStripMenuItem
        '
        Me.HKWarehouseToolStripMenuItem.Name = "HKWarehouseToolStripMenuItem"
        Me.HKWarehouseToolStripMenuItem.Size = New System.Drawing.Size(162, 22)
        Me.HKWarehouseToolStripMenuItem.Text = "HK - Warehouse"
        '
        'QueryTaiwanToolStripMenuItem
        '
        Me.QueryTaiwanToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.TWSalesExtractToolStripMenuItem, Me.TWInvoiceToolStripMenuItem, Me.TWWarehouseToolStripMenuItem})
        Me.QueryTaiwanToolStripMenuItem.Name = "QueryTaiwanToolStripMenuItem"
        Me.QueryTaiwanToolStripMenuItem.Size = New System.Drawing.Size(91, 20)
        Me.QueryTaiwanToolStripMenuItem.Text = "Query Taiwan"
        '
        'TWSalesExtractToolStripMenuItem
        '
        Me.TWSalesExtractToolStripMenuItem.Name = "TWSalesExtractToolStripMenuItem"
        Me.TWSalesExtractToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.TWSalesExtractToolStripMenuItem.Text = "TW - SalesExtract"
        '
        'TWInvoiceToolStripMenuItem
        '
        Me.TWInvoiceToolStripMenuItem.Name = "TWInvoiceToolStripMenuItem"
        Me.TWInvoiceToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.TWInvoiceToolStripMenuItem.Text = "TW - Tax Invoice"
        '
        'TWWarehouseToolStripMenuItem
        '
        Me.TWWarehouseToolStripMenuItem.Name = "TWWarehouseToolStripMenuItem"
        Me.TWWarehouseToolStripMenuItem.Size = New System.Drawing.Size(164, 22)
        Me.TWWarehouseToolStripMenuItem.Text = "TW - Warehouse"
        '
        'AdminToolStripMenuItem
        '
        Me.AdminToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UserToolStripMenuItem, Me.ImportPOSDataToolStripMenuItem, Me.AutoReportWarehouseToolStripMenuItem1})
        Me.AdminToolStripMenuItem.Name = "AdminToolStripMenuItem"
        Me.AdminToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.AdminToolStripMenuItem.Text = "Admin"
        '
        'UserToolStripMenuItem
        '
        Me.UserToolStripMenuItem.Name = "UserToolStripMenuItem"
        Me.UserToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.UserToolStripMenuItem.Text = "User"
        '
        'ImportPOSDataToolStripMenuItem
        '
        Me.ImportPOSDataToolStripMenuItem.Name = "ImportPOSDataToolStripMenuItem"
        Me.ImportPOSDataToolStripMenuItem.Size = New System.Drawing.Size(200, 22)
        Me.ImportPOSDataToolStripMenuItem.Text = "Import POS Data"
        '
        'MasterToolStripMenuItem
        '
        Me.MasterToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MLATWToolStripMenuItem, Me.FamilyHKToolStripMenuItem})
        Me.MasterToolStripMenuItem.Name = "MasterToolStripMenuItem"
        Me.MasterToolStripMenuItem.Size = New System.Drawing.Size(55, 20)
        Me.MasterToolStripMenuItem.Text = "Master"
        '
        'MLATWToolStripMenuItem
        '
        Me.MLATWToolStripMenuItem.Name = "MLATWToolStripMenuItem"
        Me.MLATWToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.MLATWToolStripMenuItem.Text = "MLA-TW"
        '
        'FamilyHKToolStripMenuItem
        '
        Me.FamilyHKToolStripMenuItem.Name = "FamilyHKToolStripMenuItem"
        Me.FamilyHKToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.FamilyHKToolStripMenuItem.Text = "Family HK"
        '
        'StaffPurchaseToolStripMenuItem
        '
        Me.StaffPurchaseToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ItemPriceMasterToolStripMenuItem})
        Me.StaffPurchaseToolStripMenuItem.Name = "StaffPurchaseToolStripMenuItem"
        Me.StaffPurchaseToolStripMenuItem.Size = New System.Drawing.Size(94, 20)
        Me.StaffPurchaseToolStripMenuItem.Text = "Staff Purchase"
        '
        'ItemPriceMasterToolStripMenuItem
        '
        Me.ItemPriceMasterToolStripMenuItem.Name = "ItemPriceMasterToolStripMenuItem"
        Me.ItemPriceMasterToolStripMenuItem.Size = New System.Drawing.Size(166, 22)
        Me.ItemPriceMasterToolStripMenuItem.Text = "Item Price Master"
        '
        'LogisticsToolStripMenuItem
        '
        Me.LogisticsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.POInformationToolStripMenuItem, Me.POInformationV2ToolStripMenuItem})
        Me.LogisticsToolStripMenuItem.Name = "LogisticsToolStripMenuItem"
        Me.LogisticsToolStripMenuItem.Size = New System.Drawing.Size(65, 20)
        Me.LogisticsToolStripMenuItem.Text = "Logistics"
        '
        'POInformationToolStripMenuItem
        '
        Me.POInformationToolStripMenuItem.Name = "POInformationToolStripMenuItem"
        Me.POInformationToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.POInformationToolStripMenuItem.Text = "PO Information"
        '
        'POInformationV2ToolStripMenuItem
        '
        Me.POInformationV2ToolStripMenuItem.Name = "POInformationV2ToolStripMenuItem"
        Me.POInformationV2ToolStripMenuItem.Size = New System.Drawing.Size(172, 22)
        Me.POInformationV2ToolStripMenuItem.Text = "PO Information V2"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(312, 33)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(60, 13)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "User Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(378, 33)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(0, 13)
        Me.Label2.TabIndex = 6
        '
        'AutoReportWarehouseToolStripMenuItem1
        '
        Me.AutoReportWarehouseToolStripMenuItem1.Name = "AutoReportWarehouseToolStripMenuItem1"
        Me.AutoReportWarehouseToolStripMenuItem1.Size = New System.Drawing.Size(200, 22)
        Me.AutoReportWarehouseToolStripMenuItem1.Text = "Auto Report Warehouse"
        '
        'FormMenu
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(482, 106)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.MenuStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MainMenuStrip = Me.MenuStrip1
        Me.Name = "FormMenu"
        Me.Text = "FormMenu"
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents MenuStrip1 As System.Windows.Forms.MenuStrip
    Friend WithEvents QueryToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HKSalesExtractionToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AdminToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents UserToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents QueryTaiwanToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TWSalesExtractToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents MLATWToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents FamilyHKToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ImportPOSDataToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AutoReportToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents StaffPurchaseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ItemPriceMasterToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents LogisticsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POInformationToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents POInformationV2ToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TWInvoiceToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HKWarehouseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents TWWarehouseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AutoReportWarehouseToolStripMenuItem1 As System.Windows.Forms.ToolStripMenuItem
End Class
