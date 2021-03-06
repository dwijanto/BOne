﻿Imports System.Text
Public Class FamilyAdapter
    Inherits PostgreSQLModelAdapter
    Public BS As BindingSource
    Public Function getDataSet()
        'sqlstr = "select familyid,familyname::character varying from bone.family;"
        sqlstr = "select familyid,familyname::character(20) from sales.tb_family;"
        Return Me.load
    End Function

    Public Function LoadData() As Boolean
        Dim myret As Boolean = False
        DS = New DataSet
        BS = New BindingSource

        Dim sb As New StringBuilder
        sqlstr = "select  * from sales.tb_family;"
        If Me.load() Then
            'Set Primary Key
            Dim pk(0) As DataColumn
            pk(0) = DS.Tables(0).Columns("familyid")
            DS.Tables(0).PrimaryKey = pk
            'DS.Tables(0).Columns("id").AutoIncrement = True
            'DS.Tables(0).Columns("id").AutoIncrementSeed = -1
            'DS.Tables(0).Columns("id").AutoIncrementStep = -1

            BS.DataSource = DS.Tables(0)
            myret = True
        Else
            MessageBox.Show(Me.errorMsg)
        End If
        Return myret
    End Function

    Public Function Save() As Boolean
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

    Public Function Save(ByVal mye As TxBaseEventArgs) As Boolean
        Dim dataadapter As Npgsql.NpgsqlDataAdapter = New Npgsql.NpgsqlDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf dbAdapter1.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Npgsql.NpgsqlConnection = New Npgsql.NpgsqlConnection(dbAdapter1.ConnectionString)
            conn.Open()
            mytransaction = conn.BeginTransaction
            'Update
            Dim sqlstr = "sales.sp_updatetbfamily"
            dataadapter.UpdateCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)

            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyid").SourceVersion = DataRowVersion.Original
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_inserttbfamily"
            dataadapter.InsertCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)

            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "sales.sp_deletetbfamily"
            dataadapter.DeleteCommand = New Npgsql.NpgsqlCommand(sqlstr, conn)

            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "familyid").SourceVersion = DataRowVersion.Original
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
End Class
