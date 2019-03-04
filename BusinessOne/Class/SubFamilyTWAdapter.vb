Public Class SubFamilyTWAdapter

    Inherits PostgreSQLModelAdapter

    Public Function getDataSet()
        sqlstr = "select familylv2id::character varying,familylv2name::character varying from bone.familylv2;"
        'sqlstr = "select replace(subfamilyid,'-','')as subfamilyid,subfamilyid as subfamcode,subfamilyname from sales.tb_subfamily;"
        Return Me.load
    End Function

End Class
