Imports System.Text

Public Class UserAdapter
    Inherits PostgreSQLModelAdapter
    Implements IIdentity
    Implements IdbAdapter

    Public Property userid As String
    Private Shared myInstance As UserAdapter
    Public Shared Property IdentityClass As Object
    Public Property id As Object
    Public Property username As String
    Public Property password_hash As String
    Public Property isAdmin As Boolean
    Public Property eamil As String
    Public Property isActive As Boolean

    Public Property BS As BindingSource Implements IdbAdapter.BS

    Public Shared Function getInstance() As UserAdapter
        If myInstance Is Nothing Then
            myInstance = New UserAdapter
        End If
        Return myInstance
    End Function

    Public Sub New()
        MyBase.New()
        tableName = "bone._user" 'tablename
        primarykey = "id" 'primarykey
    End Sub

    Public Function CanLogin(ByVal userid, ByVal password) As Boolean
        Return True
    End Function

    Public Function findByUserName(ByVal username As String)
        Dim myCondition As New Hashtable
        myCondition.Add("lower(username)", username.ToLower)
        Return findOne(myCondition)
    End Function

    Private Function populatedata(dr As DataRow) As IIdentity
        Dim Identity As New UserAdapter With {.id = dr.Item("id"),
                                            .username = dr.Item("username"),
                                            .isAdmin = dr.Item("isadmin"),
                                            .isActive = dr.Item("isactive"),
                                            .userid = "",
                                            .password_hash = ""}
        Return Identity
    End Function

    Public Function findIdentity(id As Object) As Object Implements IIdentity.findIdentity
        Dim ds As DataSet = Me.findOne(id)
        If Not IsNothing(ds) Then
            Return populatedata(ds.Tables(0).Rows(0))
        Else
            Return Nothing
        End If
    End Function

    Public Function findIdentityByAccessToken(token As Object, Optional type As Object = Nothing) As Object Implements IIdentity.findIdentityByAccessToken
        Dim myCondition As New Hashtable
        myCondition.Add("accestoken", token)
        Return findOne(myCondition)
    End Function

    Public Function getAuthKey() As String Implements IIdentity.getAuthKey
        Throw New NotImplementedException
    End Function

    Public Function getId() As Object Implements IIdentity.getId
        Return _id
    End Function

    Public Function validateAuthKey(authkey As String) As Boolean Implements IIdentity.validateAuthKey
        Throw New NotImplementedException
    End Function


    Public Function LoadData() As Boolean Implements IdbAdapter.LoadData
        DS = New DataSet
        BS = New BindingSource

        Dim sb As New StringBuilder
        sb.Append("select u.* from bone._user u order by u.username;")
        dbAdapter1.GetDataset(sb.ToString, DS)
        'Set Primary Key
        Dim pk(0) As DataColumn
        pk(0) = DS.Tables(0).Columns("id")
        DS.Tables(0).PrimaryKey = pk
        DS.Tables(0).Columns("id").AutoIncrement = True
        DS.Tables(0).Columns("id").AutoIncrementSeed = -1
        DS.Tables(0).Columns("id").AutoIncrementStep = -1

        'Unique Constrain
        Dim U0(0) As DataColumn
        U0(0) = DS.Tables(0).Columns("username")
        Dim UUser As UniqueConstraint = New UniqueConstraint(U0)
        DS.Tables(0).Constraints.Add(UUser)
        BS.DataSource = DS.Tables(0)
        Return True
    End Function

    Public Function Save() As Boolean Implements IdbAdapter.Save
        Dim myret As Boolean = False
        BS.EndEdit()

        Dim ds2 As DataSet = DS.GetChanges
        If Not IsNothing(ds2) Then
            Dim mymessage As String = String.Empty
            Dim ra As Integer
            Dim mye As New TxBaseEventArgs(ds2, True, mymessage, ra, True)
            Try
                If Save(mye) Then
                    DS.Merge(ds2)
                    DS.AcceptChanges()
                    MessageBox.Show("Saved.")
                    myret = True
                End If
            Catch ex As Exception
                MessageBox.Show(ex.Message)
                DS.Merge(ds2)
            End Try
        End If
        Return myret
    End Function

    Public Function Save(ByVal mye As TxBaseEventArgs) As Boolean Implements IdbAdapter.Save
        Dim dataadapter As Npgsql.NpgsqlDataAdapter = New Npgsql.NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf dbAdapter1.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Npgsql.NpgsqlConnection = New Npgsql.NpgsqlConnection(dbAdapter1.ConnectionString)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "bone.sp_updateuser"
            dataadapter.UpdateCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "bone.sp_insertuser"
            dataadapter.InsertCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)

            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "username").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "email").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "bone.sp_deleteuser"
            dataadapter.DeleteCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)

            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.RA = dataadapter.Update(mye.DataSet.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function loglogin(ByVal applicationname As String, ByVal userid As String, ByVal username As String, ByVal computername As String, ByVal time_stamp As Date)
        Dim result As Object
        Using conn As New Npgsql.NpgsqlConnection(dbAdapter1.ConnectionString)
            conn.Open()
            Dim cmd As Npgsql.NpgsqlCommand = New Npgsql.NpgsqlCommand("sp_insertlogonhistory", conn)
            cmd.CommandType = CommandType.StoredProcedure
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = applicationname
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = userid
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = username
            cmd.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0).Value = computername
            result = cmd.ExecuteNonQuery
        End Using
        Return result
    End Function
End Class

Public Class TxBaseEventArgs
    Inherits EventArgs
    Public Property DataSet As DataSet
    Public Property Message As String
    Public Property hasChanges As Boolean
    Public Property RA As Object
    Public Property ContinueOnError As Boolean

    Public Sub New(ByVal ds As DataSet, ByRef hasChanges As Boolean, ByRef message As String, ByRef recordAffected As Object, ByVal continueOnError As Boolean)
        With Me
            .DataSet = ds
            .hasChanges = hasChanges
            .Message = message
            .RA = recordAffected
            .ContinueOnError = continueOnError
        End With
    End Sub
End Class