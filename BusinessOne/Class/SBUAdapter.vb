Public Class SBUAdapter
    Inherits PostgreSQLModelAdapter

    Public Function getDataSet()
        sqlstr = "select * from bone.sbusap;"
        Return Me.load
    End Function
End Class

Public Class CompanyAdapter
    Inherits PostgreSQLModelAdapter
    Public Function getDataSet()
        sqlstr = "select * from sales.customer;select * from sales.custprodkam;"
        Return Me.load
    End Function

    Public Function Save(ByVal company As Object, ByVal mye As ContentBaseEventArgs)
        Return Me.CompanyTx(company, mye)
    End Function


End Class
