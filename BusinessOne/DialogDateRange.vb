Imports System.Windows.Forms

Public Class DialogDateRange
    Public startdate As Date
    Public enddate As Date
    Public MyLocation As LocationEnum = LocationEnum.Hong_Kong

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        startdate = DateTimePicker1.Value.Date
        enddate = DateTimePicker2.Value.Date
        If RadioButton1.Checked Then
            MyLocation = LocationEnum.Hong_Kong
        Else
            MyLocation = LocationEnum.Taiwan
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()

    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

End Class
