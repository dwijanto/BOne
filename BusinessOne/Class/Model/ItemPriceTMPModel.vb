Imports Npgsql
Imports System.Text
Public Class ItemPriceTMPModel
    Implements IModel

    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance
    Private sqlstr As String = String.Empty

    Public ReadOnly Property sortField As String Implements IModel.sortField
        Get
            Return "instore desc,producttypename,U_sebcocod"
        End Get
    End Property

    Public ReadOnly Property tablename As String Implements IModel.tablename
        Get
            Return "shop.itempricetmp"
        End Get
    End Property

    Public ReadOnly Property FilterField()
        Get
            Return "[itemcode] like '*{0}*' or [itemname] like '*{0}*' or [frgnname] like '*{0}*' or [suppcatnum] like '*{0}*' or [u_sebcocod] like '*{0}*' or [u_sebacode] like '*{0}*' or [u_sebbran3] like '*{0}*' or [u_sebctype] like '*{0}*'  or [u_sebcwfamtype] like '*{0}*'  or [u_sebfami1] like '*{0}*'  or [u_sebfamilytype] like '*{0}*' or [u_sebfamlev1cury] like '*{0}*' or [familyname] like '*{0}*' or [brandname] like '*{0}*' or [producttypename] like '*{0}*' or [productname] like '*{0}*'"
        End Get
    End Property

    Public Function LoadData(DS As DataSet) As Boolean Implements IModel.LoadData
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        Using conn As Object = myAdapter.getConnection
            conn.Open()

            sqlstr = String.Format("select u.*,b.brandname,f.familyname,i.itemid,u.price - ip.retailprice as difference,ip.retailprice, case when not ip.retailprice isnull then true else false end as instore,fpt.producttypeid,pt.producttypename,i.productid,p.productname, case when u_sebcwfamtype = 'Y' then (price * 0.5)::integer when u_sebcwfamtype = 'N' then (price * 0.6)::integer end as staffprice,ip.promotionprice,ip.promotionstartdate,ip.promotionenddate from {0} u " &
                                   " left join shop.brand b on b.brandid = (case when u.u_sebbran3 = '' then '0' else u.u_sebbran3 end)::bigint" &
                                   " left join shop.family f on f.familyid = u.u_sebfamlev1cury::integer" &
                                   " left join shop.familyproducttype fpt on fpt.familyid = f.familyid" &
                                   " left join shop.producttype pt on pt.producttypeid = fpt.producttypeid" &
                                   " left join shop.item i on i.refno = u.u_sebcocod" &
                                   " left join shop.itemprice ip on ip.itempriceid = i.itemid " &
                                   " left join shop.product p on p.productid = i.productid" &
                                   " order by {1}", tablename, sortField)
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret = True
        End Using
        Return myret
    End Function

    Public Function CanRollback() As Boolean
        Dim sqlstr = String.Format("select count(itempriceid) from {0};", "shop.itempricebck")
        Dim ra As Integer
        myAdapter.ExecuteScalar(sqlstr, recordAffected:=ra)
        Return ra > 0
    End Function

    Public Function getProductTypeName() As BindingSource
        Dim DS As New DataSet
        Dim myret As New BindingSource
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Using conn As Object = myAdapter.getConnection
            conn.Open()

            sqlstr = String.Format("select 0::integer as producttypeid, ''::character varying as producttypename union all" &
                                   " (select * from shop.producttype order by producttypeid)")
            dataadapter.SelectCommand = myAdapter.getCommandObject(sqlstr, conn)
            dataadapter.SelectCommand.CommandType = CommandType.Text
            dataadapter.Fill(DS, tablename)
            myret.DataSource = DS.Tables(0)
        End Using
        Return myret
    End Function

    Public Function DeleteAll() As Boolean
        Dim sqlstr = String.Format("delete from {0};", tablename)
        Return myAdapter.ExecuteNonQuery(sqlstr)
    End Function

    Public Function DeleteItemPrice() As Boolean
        Dim sqlstr = String.Format("delete from shop.itempricebck;insert into shop.itempricebck select * from {0};  delete from {0};", "shop.itemprice")
        Return myAdapter.ExecuteNonQuery(sqlstr)
    End Function

    Public Function save(obj As Object, mye As ContentBaseEventArgs) As Boolean Implements IModel.save
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myAdapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            'ActionPlan
            Dim sqlstr As String
            sqlstr = "shop.sp_deleteitempricetmp"
            dataadapter.DeleteCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.DeleteCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "id").SourceVersion = DataRowVersion.Original
            dataadapter.DeleteCommand.CommandType = CommandType.StoredProcedure

            sqlstr = "shop.sp_insertitempricetmp"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "itemname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "frgnname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "SuppCatNum").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_Sebcocod").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBacode").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBbran3").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBctype").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBCWFamType").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBfami1").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBFamilyType").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBFamLev1CurY").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "Price").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "itemcode").Direction = ParameterDirection.InputOutput
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure


            sqlstr = "shop.sp_updateitempricetmp"
            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)

            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "itemname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "frgnname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "SuppCatNum").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_Sebcocod").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBacode").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBbran3").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "U_SEBctype").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "U_SEBCWFamType").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBfami1").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "U_SEBFamilyType").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_SEBFamLev1CurY").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "Price").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "currency").SourceVersion = DataRowVersion.Current

            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Bigint, 0, "itemcode").Direction = ParameterDirection.InputOutput
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction
            dataadapter.DeleteCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function


    Public Function saveStaffItemPrice(obj As Object, mye As ContentBaseEventArgs) As Boolean
        Dim dataadapter As NpgsqlDataAdapter = myAdapter.getDbDataAdapter
        Dim myret As Boolean = False
        AddHandler dataadapter.RowUpdated, AddressOf myAdapter.onRowInsertUpdate
        Dim mytransaction As Npgsql.NpgsqlTransaction
        Using conn As Object = myAdapter.getConnection
            conn.Open()
            mytransaction = conn.BeginTransaction
            Dim sqlstr As String

            sqlstr = "shop.sp_addupditemprice"
            dataadapter.InsertCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_Sebcocod").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "U_SEBbran3").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "U_SEBFamLev1CurY").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "producttypeid").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "itemname").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "Price").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "staffprice").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "promotionprice").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "promotionstartdate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "promotionenddate").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "instore").SourceVersion = DataRowVersion.Current
            dataadapter.InsertCommand.CommandType = CommandType.StoredProcedure

            dataadapter.UpdateCommand = New NpgsqlCommand(sqlstr, conn)
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "U_Sebcocod").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "U_SEBbran3").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "U_SEBFamLev1CurY").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "productname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Integer, 0, "producttypeid").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Varchar, 0, "itemname").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "Price").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "staffprice").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Numeric, 0, "promotionprice").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "promotionstartdate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Date, 0, "promotionenddate").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.Parameters.Add("", NpgsqlTypes.NpgsqlDbType.Boolean, 0, "instore").SourceVersion = DataRowVersion.Current
            dataadapter.UpdateCommand.CommandType = CommandType.StoredProcedure

            dataadapter.InsertCommand.Transaction = mytransaction
            dataadapter.UpdateCommand.Transaction = mytransaction

            mye.ra = dataadapter.Update(mye.dataset.Tables(0))

            mytransaction.Commit()
            myret = True
        End Using
        Return myret
    End Function

    Function RollBack() As Boolean
        Dim sqlstr = String.Format("delete from {0};insert into {0} select * from shop.itempricebck;", "shop.itemprice")
        Return myAdapter.ExecuteNonQuery(sqlstr)
    End Function


End Class
