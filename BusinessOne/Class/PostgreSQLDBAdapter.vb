Imports Npgsql
Imports System.IO

Public Class PostgreSQLDBAdapter
    Implements IDisposable

    Dim mytransaction As NpgsqlTransaction

    Public Shared myInstance As PostgreSQLDBAdapter
    Private _connectionstring As String
    Private CopyIn1 As NpgsqlCopyIn
    Private Sub New()

        DbAdapterInitialize()

    End Sub


    Private Sub DbAdapterInitialize()
        '_connectionstring = "host=hon14nt;port=5432;database=LogisticDb;commandTimeout=1000;Timeout=1000;"
        '_connectionstring = "host=localhost;port=5432;database=LogisticDb20150120;CommandTimeout=1000;TimeOut=1000;User=admin;Password=admin"
        _connectionstring = String.Format(My.Settings.PostgreSQLCon & "{0}", "User=admin;Password=admin")
    End Sub
    Public Property ConnectionString As String
        Get
            Return _connectionstring
        End Get
        Set(value As String)
            _connectionstring = value
        End Set
    End Property

    Public Shared Function getInstance() As PostgreSQLDBAdapter
        If myInstance Is Nothing Then
            myInstance = New PostgreSQLDBAdapter
        End If
        Return myInstance
    End Function

    Public Overloads Function GetDataset(sqlstr As String, ds As DataSet, Optional params As List(Of IDataParameter) = Nothing) As Boolean
        Dim DataAdapter As IDbDataAdapter = New NpgsqlDataAdapter
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                DataAdapter.SelectCommand = cmd
                If Not IsNothing(params) Then
                    For Each param As IDataParameter In params
                        cmd.Parameters.Add(param)
                    Next
                End If
                DataAdapter.Fill(ds)
                myret = True
            End Using
        End Using
        Return myret
    End Function

    'Public Overloads Function getDataSet(ByVal sqlstr As String, ByVal DataSet As DataSet, Optional ByRef message As String = "") As Boolean
    '    Dim DataAdapter As New NpgsqlDataAdapter

    '    Dim myret As Boolean = False
    '    Try
    '        Using conn As New NpgsqlConnection(_connectionstring)
    '            conn.Open()
    '            DataAdapter.SelectCommand = New NpgsqlCommand(sqlstr, conn)
    '            DataAdapter.Fill(DataSet)
    '        End Using
    '        myret = True
    '    Catch ex As NpgsqlException
    '        Dim obj = TryCast(ex.Errors(0), NpgsqlError)
    '        Dim myerror As String = String.Empty
    '        If Not IsNothing(obj) Then
    '            myerror = obj.Message
    '        End If
    '        message = ex.Message & " " & myerror
    '    End Try
    '    Return myret
    'End Function

    Public Function ExecuteScalar(ByVal sqlstr As String, Optional ByVal params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Object = Nothing, Optional ByRef message As String = "") As Boolean
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                Try
                    If Not IsNothing(params) Then
                        For Each param As NpgsqlParameter In params
                            cmd.Parameters.Add(param)
                        Next
                    End If
                    recordAffected = cmd.ExecuteScalar
                    myret = True
                Catch ex As Exception
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myret
    End Function

    Public Function ExecuteNonQuery(ByVal sqlstr As String, Optional ByVal params As List(Of IDataParameter) = Nothing, Optional ByRef recordAffected As Int64 = 0, Optional ByRef message As String = "") As Boolean
        Dim myret As Boolean = False
        Using conn As New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using cmd As NpgsqlCommand = New NpgsqlCommand()
                cmd.CommandText = sqlstr
                cmd.Connection = conn
                Try
                    If Not IsNothing(params) Then
                        For Each param As IDataParameter In params
                            cmd.Parameters.Add(param)
                        Next
                    End If
                    recordAffected = cmd.ExecuteNonQuery
                    myret = True
                Catch ex As Exception
                    message = ex.Message
                End Try
            End Using
        End Using
        Return myret
    End Function

    Public Function getParam(ByVal ParameterName As String,
                             Optional ByVal value As Object = Nothing,
                         Optional ByVal dbType As DbType = Nothing,
                         Optional ByVal direction As ParameterDirection = ParameterDirection.Input,
                         Optional isNullable As Boolean = False,
                         Optional Precision As Byte = 0,
                         Optional scale As Byte = 0,
                         Optional size As Integer = Integer.MaxValue,
                         Optional SourceColumn As String = "",
                         Optional sourceversion As DataRowVersion = DataRowVersion.Current) As NpgsqlParameter
        Dim myparam = New NpgsqlParameter
        With myparam
            .ParameterName = ParameterName
            .Value = value
            .DbType = dbType
            .Direction = direction
            .IsNullable = isNullable
            .Precision = Precision
            .Scale = scale
            .Size = size
            .SourceColumn = SourceColumn
            .SourceVersion = sourceversion
        End With
        Return myparam
    End Function

    Public Function isAdmin(ByVal userid As String) As Boolean
        Dim sqlstr = String.Format("select * from _user where userid = :userid")
        Dim dbparams As New List(Of IDataParameter)
        dbparams.Add(getParam("userid", userid, DbType.String))
        Dim ra As Object = Nothing
        Dim message As String = String.Empty
        If ExecuteScalar(sqlstr, dbparams, ra, message) Then
            Return Not IsNothing(ra)
        End If
        Return False
    End Function

    Public Function copy(ByVal sqlstr As String, ByVal InputString As String, Optional ByRef result As Boolean = False) As String
        result = False
        Dim myReturn As String = ""
        'Convert string to MemoryStream
        Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.ASCII.GetBytes(InputString.Replace("\", "\\")))
        'Dim MemoryStream1 As New IO.MemoryStream(System.Text.Encoding.Default.GetBytes(InputString))
        Dim buf(9) As Byte
        Dim CopyInStream As Stream = Nothing
        Dim i As Long
        Using conn = New NpgsqlConnection(_connectionstring)
            conn.Open()
            Using command = New NpgsqlCommand(sqlstr, conn)
                CopyIn1 = New NpgsqlCopyIn(command, conn)
                Try
                    CopyIn1.Start()
                    CopyInStream = CopyIn1.CopyStream
                    i = MemoryStream1.Read(buf, 0, buf.Length)
                    While i > 0
                        CopyInStream.Write(buf, 0, i)
                        i = MemoryStream1.Read(buf, 0, buf.Length)
                        Application.DoEvents()
                    End While
                    CopyInStream.Close()
                    result = True
                Catch ex As NpgsqlException
                    Try
                        CopyIn1.Cancel("Undo Copy")
                        myReturn = ex.Message & vbCrLf & ex.Detail & vbCrLf & ex.Where
                    Catch ex2 As NpgsqlException
                        If ex2.Message.Contains("Undo Copy") Then
                            myReturn = ex2.Message & ex.Where
                        End If
                    End Try
                End Try

            End Using
        End Using

        Return myReturn
    End Function

    Public Function getLastImportDate() As Date
        Dim sqlstr = "select dvalue from bone.paramhd where paramname = 'lastImport';"
        Dim result As Date
        Me.ExecuteScalar(sqlstr, recordAffected:=result)
        Return result
    End Function
    Public Sub setLastImportDate(mydate As Date)
        Dim sqlstr = String.Format("update bone.paramhd set dvalue = '{0:yyyy-MM-dd}' where paramname = 'lastImport';", mydate)
        Me.ExecuteNonQuery(sqlstr)

    End Sub
#Region "IDisposable Support"
    Private disposedValue As Boolean ' To detect redundant calls

    ' IDisposable
    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not Me.disposedValue Then
            If disposing Then
                ' TODO: dispose managed state (managed objects).
            End If

            ' TODO: free unmanaged resources (unmanaged objects) and override Finalize() below.
            ' TODO: set large fields to null.
        End If
        Me.disposedValue = True
    End Sub

    ' TODO: override Finalize() only if Dispose(ByVal disposing As Boolean) above has code to free unmanaged resources.
    'Protected Overrides Sub Finalize()
    '    ' Do not change this code.  Put cleanup code in Dispose(ByVal disposing As Boolean) above.
    '    Dispose(False)
    '    MyBase.Finalize()
    'End Sub

    ' This code added by Visual Basic to correctly implement the disposable pattern.
    Public Sub Dispose() Implements IDisposable.Dispose
        ' Do not change this code.  Put cleanup code in Dispose(disposing As Boolean) above.
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
#End Region

    Sub onRowInsertUpdate(sender As Object, e As NpgsqlRowUpdatedEventArgs)
        'Table with autoincrement
        If e.StatementType = StatementType.Insert Or e.StatementType = StatementType.Update Then
            If e.Status <> UpdateStatus.ErrorsOccurred Then
                e.Status = UpdateStatus.SkipCurrentRow
            End If
        End If
    End Sub

    Public Function CompanyTx(ByVal Company As Object, ByVal mye As ContentBaseEventArgs) As Boolean
        Dim sqlstr As String = String.Empty
        Dim DataAdapter As New NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler DataAdapter.RowUpdated, New NpgsqlRowUpdatedEventHandler(AddressOf onRowInsertUpdate)
        Try
            Using conn As New NpgsqlConnection(ConnectionString)
                conn.Open()
                mytransaction = conn.BeginTransaction
                'Update
                'sqlstr = "doc.sp_insertupdatefamilygroupsbu"
                'DataAdapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "groupingcode").SourceVersion = DataRowVersion.Current
                'DataAdapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sbusapid").SourceVersion = DataRowVersion.Current
                'DataAdapter.UpdateCommand.CommandType = CommandType.StoredProcedure

                sqlstr = "sales.sp_insertcustomer"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "customerid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "customername").SourceVersion = DataRowVersion.Current                
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                'sqlstr = "doc.sp_deletefamilygroupsbu"
                'DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                'DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                'DataAdapter.UpdateCommand.Transaction = mytransaction
                'DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(0))

                sqlstr = "sales.sp_insertcustomerprodkam"
                DataAdapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "customerid").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "sda").SourceVersion = DataRowVersion.Current
                DataAdapter.InsertCommand.CommandType = CommandType.StoredProcedure

                'sqlstr = "doc.sp_deletefamilygroupsbu"
                'DataAdapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
                'DataAdapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "familyid").SourceVersion = DataRowVersion.Original
                'DataAdapter.DeleteCommand.CommandType = CommandType.StoredProcedure

                DataAdapter.InsertCommand.Transaction = mytransaction
                'DataAdapter.UpdateCommand.Transaction = mytransaction
                'DataAdapter.DeleteCommand.Transaction = mytransaction

                mye.ra = DataAdapter.Update(mye.dataset.Tables(1))

                mytransaction.Commit()
                myret = True

            End Using

        Catch ex As NpgsqlException
            Dim errordetail As String = String.Empty
            errordetail = "" & ex.Detail
            mye.message = ex.Message & ". " & errordetail
            Return False
        End Try
        Return myret
    End Function

    Public Function getConnection() As NpgsqlConnection       
        Return New NpgsqlConnection(_connectionString)
    End Function

    Public Function getDbDataAdapter() As NpgsqlDataAdapter
        Return New NpgsqlDataAdapter
    End Function

    Public Function getCommandObject() As NpgsqlCommand
        Return New NpgsqlCommand
    End Function

    Public Function getCommandObject(ByVal sqlstr As String, ByVal connection As Object) As NpgsqlCommand
        Return New NpgsqlCommand(sqlstr, connection)
    End Function
End Class
