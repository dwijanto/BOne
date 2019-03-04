Public Class RDR1Adapter
    Inherits ModelAdapter

    Public Sub New()
        MyBase.New()
    End Sub

    Public Function getOrder(DocNum) As Boolean
        sqlstr = String.Format("select i.itemcode,i.dscription,d.frgnname,i.quantity,i.price,i.currency,i.rate,i.discprcnt,(i.price * i.quantity) as linetotal" &
                               " from rdr1 i" &
                               " left join oitm d on d.itemcode = i.itemcode where i.docentry = {0}", DocNum)
        Return Me.load()
    End Function
End Class
