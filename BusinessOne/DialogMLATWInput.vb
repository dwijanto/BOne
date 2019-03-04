Imports System.Windows.Forms

Public Class DialogMLATWInput
    Private DRV As DataRowView
    Public Shared Event FinishUpdate()
    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        RaiseEvent FinishUpdate()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        RaiseEvent FinishUpdate()
        Me.Close()
    End Sub

    Private Sub InitDataDRV()

        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        TextBox6.DataBindings.Clear()
        TextBox7.DataBindings.Clear()
        TextBox8.DataBindings.Clear()

        DateTimePicker1.DataBindings.Clear()
        DateTimePicker2.DataBindings.Clear()
        

        TextBox1.DataBindings.Add(New Binding("text", DRV, "id", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox2.DataBindings.Add(New Binding("text", DRV, "mlaname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox3.DataBindings.Add(New Binding("text", DRV, "countryid", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox4.DataBindings.Add(New Binding("text", DRV, "countryname", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox5.DataBindings.Add(New Binding("text", DRV, "distchannel", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox6.DataBindings.Add(New Binding("text", DRV, "distchanneldesc", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox7.DataBindings.Add(New Binding("text", DRV, "mlatype", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        TextBox8.DataBindings.Add(New Binding("text", DRV, "status", True, DataSourceUpdateMode.OnPropertyChanged, ""))

        DateTimePicker1.DataBindings.Add(New Binding("text", DRV, "validfrom", True, DataSourceUpdateMode.OnPropertyChanged, ""))
        DateTimePicker2.DataBindings.Add(New Binding("text", DRV, "validto", True, DataSourceUpdateMode.OnPropertyChanged, ""))

    End Sub

    Public Sub New(ByVal drv As DataRowView)

        ' This call is required by the designer.
        InitializeComponent()
        Me.DRV = drv
        ' Add any initialization after the InitializeComponent() call.
        InitDataDRV()
    End Sub

End Class
