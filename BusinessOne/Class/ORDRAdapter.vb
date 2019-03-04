Public Class ORDRAdapter
    Inherits ModelAdapter

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function LoadData() As Boolean
        sqlstr = "select docnum,cardcode,cardname,docdate from ORDR;"
        Return Me.load
    End Function

End Class
