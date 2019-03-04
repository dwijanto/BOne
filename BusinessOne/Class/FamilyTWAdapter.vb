Public Class FamilyTWAdapter
    Inherits PostgreSQLModelAdapter

    Public Function getDataSet()
        'sqlstr = "select familyid,familyname::character varying from bone.family;"
        sqlstr = "select  id as familyid,familyname::character(20),type from sales.tw_family;"
        Return Me.load
    End Function
End Class
