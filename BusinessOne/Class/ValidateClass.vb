Public Class ValidateClass
    Public Shared Function ValidStr(ByVal value As Object)
        If value.ToString.Length = 0 Then
            value = "Null"
        Else
            value = String.Format("'{0}'", value.ToString.Replace("'", "''"))
        End If
        Return value
    End Function
    Public Shared Function ValidDate(ByVal value As Object)
        If value.ToString.Length = 0 Then
            value = "Null"
        Else
            value = String.Format("'{0:yyyy-MM-dd}'", value)
        End If
        Return value
    End Function
    Public Shared Function ValidNumeric(ByVal value As Object)
        If value.ToString.Length = 0 Then
            value = "Null"
        End If
        Return value
    End Function
End Class
