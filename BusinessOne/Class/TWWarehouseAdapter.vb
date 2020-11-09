Imports System.IO
Imports System.Text
Public Class TWWarehouseAdapter
    Implements IAdapter
    Dim Model As New TWWarehouseModel
    Dim myParam As Param = Param.getInstance
    Public OutputFolder As String = myParam.getOutputFolderTW
    Public Property SelectedDate As Date

    Dim SqlStr As String
    Dim SB As StringBuilder

    Dim myAdapter As PostgreSQLDBAdapter = PostgreSQLDBAdapter.getInstance

    Public Property errorMsg As String
        Get
            Return Model.errorMsg
        End Get
        Set(value As String)
            Model.errorMsg = value
        End Set
    End Property

    Public ReadOnly Property DS As DataSet
        Get
            Return Model.DS
        End Get
    End Property

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        Return False
    End Function

    Public Function LoadData(ByVal selecteddate As Date) As Boolean
        Dim myret As Boolean = False
        Me.SelectedDate = selecteddate
        Model.sqlstr = String.Format("select 'DSV' as source, U_S037WhsCode as warehousecode,U_S037ItemCode as itemcode,sum(U_S037Quantity) as quantity,U_S037Date as date from [@S037ISR] " &
                               " where U_S037Date = '{0:yyyy/MM/dd}' and U_S037ItemCode > '100000000' and U_S037ItemCode <= '999999999' " &
                               " group by U_S037ItemCode, U_S037Date, U_S037WhsCode" &
                               " UNION ALL" &
                               " SELECT 'SEB' as source,T0.WhsCode,T0.ItemCode,T0.OnHand,'{0:yyyy/MM/dd}' as date " &
                               " FROM OITW T0 " &
                               " WHERE T0.[OnHand] > 0 ", selecteddate)
        myret = Model.load()
        Return myret
    End Function

    Public Function GetDataSB() As StringBuilder
        SB = New StringBuilder
        For Each dr As DataRow In Model.DS.Tables(0).Rows
            Dim mydata As TWWarehouseModel = New TWWarehouseModel With {.source = dr.Item("source"),
                                                                        .warehousecode = dr.Item("warehousecode"),
                                                                        .itemcode = dr.Item("itemcode"),
                                                                        .quantity = dr.Item("quantity"),
                                                                        .txdate = dr.Item("date")
                                                                        }
            SB.Append(mydata.source & vbTab &
                      mydata.warehousecode & vbTab &
                      mydata.itemcode & vbTab &
                      CInt(mydata.quantity) & vbTab &
                      mydata.txdate & vbCrLf)
        Next
        Return sb
    End Function

    Public Function ExportTextFile() As Boolean
        Dim myret As Boolean = True
        Dim OutputFile As String
        OutputFile = String.Format("{0}\Warehouse\TW\RawData\{1:yyyyMMdd}_TWStock.TXT", OutputFolder, SelectedDate)
        SB = GetDataSB()
        If SB.Length > 0 Then
            Using mystream As New StreamWriter(OutputFile)
                mystream.Write(SB.ToString)
            End Using
        End If

        If SB.Length > 0 Then
            SqlStr = "delete from bone.warehousetw;select setval('bone.warehousetw_id_seq',1,false);begin;set statement_timeout to 0;end;copy bone.warehousetw(source,location,itemcode,quantity,txdate ) from stdin with null as 'Null';"
            errorMsg = myAdapter.copy(SqlStr, SB.ToString, myret)
        End If
        Return myret
    End Function

    Public Function GenerateFromTextFile() As Boolean
        Dim myret As Boolean = True
        Dim InputFile As String
        InputFile = String.Format("{0}\Warehouse\TW\RawData\{1:yyyyMMdd}_TWStock.TXT", OutputFolder, SelectedDate)
        SB = New StringBuilder
        Try
            Using objTFParser = New FileIO.TextFieldParser(InputFile)
                With objTFParser
                    .TextFieldType = FileIO.FieldType.Delimited
                    .SetDelimiters(Chr(9))
                    .HasFieldsEnclosedInQuotes = True
                    Dim count As Long = 0

                    Do Until .EndOfData
                        Dim myrecord = .ReadFields

                        Dim mydata As TWWarehouseModel = New TWWarehouseModel With {.source = myrecord(0),
                                                                                    .warehousecode = myrecord(1),
                                                                                    .itemcode = myrecord(2),
                                                                                    .quantity = myrecord(3),
                                                                                    .txdate = myrecord(4)}
                        SB.Append(mydata.source & vbTab &
                                  mydata.warehousecode & vbTab &
                                  mydata.itemcode & vbTab &
                                  CInt(mydata.quantity) & vbTab &
                                  mydata.txdate & vbCrLf)


                    Loop
                End With
            End Using

            If SB.Length > 0 Then
                SqlStr = "delete from bone.warehousetw;select setval('bone.warehousetw_id_seq',1,false);begin;set statement_timeout to 0;end;copy bone.warehousetw(source,location,itemcode,quantity,txdate ) from stdin with null as 'Null';"
                errorMsg = myAdapter.copy(SqlStr, SB.ToString, myret)
            End If

        Catch ex As Exception
            errorMsg = ex.Message
            myret = False
        End Try
        Return myret

    End Function

    Public Function GetSQLReport() As String

        'SqlStr = String.Format("with ct as (select * from crosstab('with src as (select id,source,itemcode,case location when ''17'' then ''10'' else location end as location,quantity * case source when ''DSV'' then 1 else -1 end as qty from bone.warehousehk" &
        ' " where location in(select dt.cvalue from bone.paramdt dt where paramhdid = 2 and paramname = ''HKCriteria'')" &
        ' " order by source,itemcode,location)," &
        ' " dt as (select first_value(id) over (partition by source,itemcode order by source,itemcode) as myid, location,sum(qty) over (partition by source,itemcode,location) as qty from src)" &
        ' " select * from dt','select dt.cvalue from bone.paramdt dt where paramhdid = 2 and paramname = ''HKField'' order by ivalue') as " &
        ' " (myid bigint,{0}))" &
        ' " select 'HK' as country,wh.source,wh.itemcode,ct.* from ct" &
        ' " left join bone.warehousehk wh on wh.id = ct.myid", myParam.GetHKFields)

        SqlStr = String.Format("with ct as (select * from crosstab('with src as (select id,source,itemcode,case location when ''VAS'' then ''602'' else location end as location,quantity * case source when ''DSV'' then 1 else -1 end as qty from bone.warehousetw" &
                               " where location in(select dt.cvalue from bone.paramdt dt where paramhdid = 2 and paramname = ''TWCriteria'')" &
                               " order by source,itemcode,location)," &
                               " dt as (select first_value(id) over (partition by source,itemcode order by source,itemcode) as myid, location,sum(qty) over (partition by source,itemcode,location) as qty from src)" &
                               " select * from dt order by myid,location','select dt.cvalue from bone.paramdt dt where paramhdid = 2 and paramname = ''TWField''  order by ivalue') as " &
                               " (myid bigint,{0}))," &
                               " src as (select 'TW' as country,id,source,itemcode from bone.warehousetw where location in(select dt.cvalue from bone.paramdt dt where paramhdid = 2 and paramname = 'TWCriteria')" &
                               " order by source,itemcode,location) select wh.country,wh.source,wh.itemcode,ct.* from ct left join src wh on wh.id = ct.myid", myParam.GetTWFields)
        Return SqlStr
    End Function

End Class
