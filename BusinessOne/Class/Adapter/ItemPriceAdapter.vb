Public Class ItemPriceAdapter
    Inherits ModelAdapter
    Public Sub New()
        MyBase.New()
    End Sub

    Public Function loadData()
        sqlstr = "SELECT T0.ItemCode, T1.ItemName, T1.FrgnName, T1.SuppCatNum, T1.U_Sebcocod, T1.U_SEBacode," &
            " T1.U_SEBbran3, T1.U_SEBctype, T1.U_SEBCWFamType, T1.U_SEBfami1, T1.U_SEBFamilyType, T1.U_SEBFamLev1CurY," &
            " T0.Price, T0.Currency, T0.PriceList, OP.ListName" &
            " FROM ITM1 T0" &
            " LEFT JOIN OITM T1 on T1.ItemCode = T0.ItemCode" &
            " LEFT JOIN OPLN OP on OP.ListNum = T0.PriceList" &
            " Where OP.ListNum = 2 and T0.currency = 'HKD'"
        Return MyBase.load
    End Function



End Class
