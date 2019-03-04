
Public Class PostgreSQLModelAdapter
    Inherits ActiveRecord
    Protected sqlstr As String = String.Empty
    Public DS As DataSet
    'Private DBAdapter1 = PostgreSQLDBAdapter.getInstance
    Public errorMsg As String = String.Empty

    Public Sub New()

    End Sub

    Public Function load() As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        Try
            If dbAdapter1.GetDataset(sqlstr, DS) Then
                myret = True
            End If
        Catch ex As Exception
            errorMsg = ex.Message
        End Try
        Return myret
    End Function

    Public Function CompanyTx(ByVal company As Object, mye As ContentBaseEventArgs)
        Return dbAdapter1.CompanyTx(company, mye)
    End Function
    
End Class
