
Public Class OITMAdapter
    Inherits ModelAdapter

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function loadData()
        sqlstr = "select * from OITM;"
        Return MyBase.load
    End Function

End Class

