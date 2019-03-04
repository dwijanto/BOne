Public Class BrandAdapter
    Inherits PostgreSQLModelAdapter

    Public Function getDataSet()
        sqlstr = "select to_char(brandid,'FM00') as brandid,brandname::character varying from bone.brand;"
        Return Me.load
    End Function
End Class
