Imports System.Text

Public Class Param
    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Public FolderPath As String
    Public Shared myInstance As Param

    Public Sub New()

    End Sub

    Public Shared Function getInstance() As Param
        If myInstance Is Nothing Then
            myInstance = New Param
        End If
        Return myInstance
    End Function

    Public Function getOutputFolder() As String
        Dim sqlstr As String
        Dim myresult As String = String.Empty
        sqlstr = "select dt.cvalue from bone.paramdt dt left join bone.paramhd ph on ph.paramhdid = dt.paramhdid where ph.paramname = 'outputfolder';"
        myAdapter.ExecuteScalar(sqlstr, Nothing, recordAffected:=myresult)
        Return myresult
    End Function
    Public Function getOutputFolderHK() As String
        Dim sqlstr As String
        Dim myresult As String = String.Empty
        sqlstr = "select dt.cvalue from bone.paramdt dt left join bone.paramhd ph on ph.paramhdid = dt.paramhdid where dt.paramname = 'OutputFolderHK';"
        myAdapter.ExecuteScalar(sqlstr, Nothing, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function getOutputEmailWarehouseHK() As String
        Dim sqlstr As String
        Dim myresult As String = String.Empty
        sqlstr = "select dt.cvalue from bone.paramdt dt left join bone.paramhd ph on ph.paramhdid = dt.paramhdid where dt.paramname = 'emailwarehousehk';"
        myAdapter.ExecuteScalar(sqlstr, Nothing, recordAffected:=myresult)
        Return myresult
    End Function

    Public Function getOutputFolderTW() As String
        Dim sqlstr As String
        Dim myresult As String = String.Empty
        sqlstr = "select dt.cvalue from bone.paramdt dt left join bone.paramhd ph on ph.paramhdid = dt.paramhdid where dt.paramname = 'OutputFolderTW';"
        myAdapter.ExecuteScalar(sqlstr, Nothing, recordAffected:=myresult)
        Return myresult
    End Function
    Public Function GetHKFields() As String
        Dim sqlstr As String
        Dim myresult As String = String.Empty
        sqlstr = "select dt.cvalue from bone.paramdt dt  where dt.paramname = 'HKField' order by ivalue;"
        Dim DS As New DataSet
        If myAdapter.GetDataset(sqlstr, DS, Nothing) Then
            Dim sb As New StringBuilder
            For Each dr As DataRow In DS.Tables(0).Rows
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("""{0}"" int", dr.Item("cvalue")))
            Next
            myresult = sb.ToString
        End If
        Return myresult
    End Function

    Public Function GetTWFields() As String
        Dim sqlstr As String
        Dim myresult As String = String.Empty
        sqlstr = "select dt.cvalue from bone.paramdt dt  where dt.paramname = 'TWField' order by ivalue;"
        Dim DS As New DataSet
        If myAdapter.GetDataset(sqlstr, DS, Nothing) Then
            Dim sb As New StringBuilder
            For Each dr As DataRow In DS.Tables(0).Rows
                If sb.Length > 0 Then
                    sb.Append(",")
                End If
                sb.Append(String.Format("""{0}"" int", dr.Item("cvalue")))
            Next
            myresult = sb.ToString
        End If
        Return myresult
    End Function
End Class
