<Serializable>
Public Class CreateUserRule
    Inherits Rule

    Public Overloads Property name As String = "CreateUserRule"

    Public Sub New()
        MyBase.New()
        MyBase.name = "CreateUserRule"
    End Sub

    Public Overrides Function executeRule(userid As Object, Optional item As Item = Nothing, Optional params As Hashtable = Nothing) As Boolean
        If Not IsNothing(params) Then
            Return params("role").ToString = "admin"
        End If
        Return True
    End Function
End Class