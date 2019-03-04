Public Class UserAssignmentAdapter
    Inherits PostgreSQLModelAdapter

    Public Function getRoles() As List(Of BusinessOne.Item)
        Dim RBAC = New DbManager
        Return RBAC.getRoles()
    End Function
End Class
