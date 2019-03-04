Imports System.Data.SqlClient
Public Class DBAdapterTW
    Implements IDisposable
    Public Shared myInstance As DBAdapterTW
    Private _connectionstring As String
    Private _connectionstringTW As String

    Private Sub New()
        DbAdapterInitialize()
    End Sub

    Public Shared Function getInstance() As DBAdapterTW
        If myInstance Is Nothing Then
            myInstance = New DBAdapterTW
        End If
        Return myInstance
    End Function


    Private Sub DbAdapterInitialize()
        '_connectionstring = "Database=GSATaiwan_test;server=LIL161NT;user id =sa;Password=SB1Admin;"
        '_connectionstring = "Database=GSAHongkong;server=HON19NT;user id =sa;Password=SB19eVer;"
        '_connectionstring = "Database=GSAHongkong_Test;server=HON19NT;user id =sa;Password=SB19eVer;"
        '_connectionstring = "Database=GSATaiwan;server=HON19NT;user id =sa;Password=SB19eVer;"
        '_connectionstring = "Database=GSATaiwan_Test;server=HON19NT;user id =sa;Password=SB19eVer;"
        '_connectionstring = "Database=GSAHongkong_OB;server=HON19NT;user id =sa;Password=SB19eVer;"
        '_connectionstring = "Database=GSATaiwan_OB;server=HON19NT;user id =sa;Password=SB19eVer;"
        _connectionstring = "Database=GSATaiwan;server=HON19NT;user id =sa;Password=SB19eVer;"
    End Sub
    Public Property ConnectionString As String
        Get
            Return _connectionstring
        End Get
        Set(value As String)
            _connectionstring = value
        End Set
    End Property



    Public Overloads Function getDataSet(ByVal sqlstr As String, ByVal DataSet As DataSet, Optional ByRef message As String = "") As Boolean
        Dim DataAdapter As New SqlDataAdapter

        Dim myret As Boolean = False
        Try
            Using conn As New SqlConnection(_connectionstring)
                conn.Open()
                DataAdapter.SelectCommand = New SqlCommand(sqlstr, conn)
                DataAdapter.Fill(DataSet)
            End Using
            myret = True
        Catch ex As SqlException
            Dim obj = TryCast(ex.Errors(0), SqlError)
            Dim myerror As String = String.Empty
            If Not IsNothing(obj) Then
                myerror = obj.Source
            End If
            message = ex.Message & " " & myerror
        End Try
        Return myret
    End Function




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



End Class
