﻿Public Class ModelAdapter
    'Protected sqlstr As String = String.Empty
    Public sqlstr As String = String.Empty
    Public DS As DataSet
    Private DBAdapter1 = DbAdapter.getInstance
    Public errorMsg As String = String.Empty

    Public Sub New()

    End Sub

    Public Function load() As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        Try
            If DBAdapter1.getDataSet(sqlstr, DS, errorMsg) Then
                myret = True
            End If
        Catch ex As Exception
            errorMsg = ex.Message
        End Try        
        Return myret
    End Function

    Public Function Copy() As Boolean

        Return True
    End Function
End Class
