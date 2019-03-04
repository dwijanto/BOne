Public Class SalesOrderAdapter
    Inherits ModelAdapter
    Implements IAdapter

    Public Sub New()
        MyBase.New()
    End Sub
    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        Return False
    End Function

    Public Function LoadData(ByVal startdate As Date, ByVal enddate As Date) As Boolean
        'sqlstr = String.Format("SELECT T0.[DocNum] , T0.[DocDate] , T0.Numatcard , T0.CardCode, T0.CardName, '' as 'Report Code',  T2.U_SEBSalesForce, T2.Country," &
        '        "  '' as 'Cust type',T3.SlpName, T4.Country as 'Ship-to Country', T5.U_SEBcocod as 'Pi2 Commercial Code',T1.ItemCode, " &
        '        " '' as 'SBU', '' as prodfamily, '' as brand, T5.FrgnName as 'Pi2 Description', '' as 'Manufacturer', T1.Quantity, T1.LineTotal, " &
        '        " T1.StockPrice*T1.Quantity as 'Total Item Cost',T4.City ,T7.Name, '' as 'Retail', T1.Price, T8.GroupName, '' as subfamily," &
        '        " cast(T5.U_SEBFamLev1CurY as varchar(3)) as 'FamilyLv1', U_SEBfami2 as subfamcode, T5.U_SEBFamLev2CurY," &
        '        " cast(T5.U_SEBbran2 as varchar(2)) as 'SEBbran2',T5.U_SEBProdLinePi2" &
        '        " FROM OINV T0 " &
        '        " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
        '        " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
        '        " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
        '        " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.cardcode = T2.Cardcode and T4.adrestype = 'S'" &
        '        " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
        '        " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
        '        " LEFT JOIN OCST T7 on T7.Code = T4.State  and T7.country = T4.country" &
        '        " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
        '        " where  T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'" &
        '        " UNION ALL" &
        '        " SELECT T0.[DocNum], T0.[DocDate], T0.Numatcard, T0.CardCode, T0.CardName, '' as 'Report Code',  T2.U_SEBSalesForce, T2.Country," &
        '        "  '' as 'Cust type',T3.SlpName, T4.Country as 'Ship-to Country',T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode, '' as 'SBU', " &
        '        " '' as prodfamily,'' as brand,T5.FrgnName as 'Pi2 Description', '' as 'Manufacturer', -T1.Quantity, -T1.LineTotal, -T1.StockPrice*T1.Quantity as 'Total Item Cost'," &
        '        " T4.City ,T7.Name, '' as 'Retail', T1.Price, T8.GroupName,  '' as subfamily, cast(T5.U_SEBFamLev1CurY as varchar(3)) as 'FamilyLv1', " &
        '        " U_SEBfami2 as subfamcode,T5.U_SEBFamLev2CurY,cast(T5.U_SEBbran2 as varchar(2)) as 'SEBbran2',T5.U_SEBProdLinePi2" &
        '        " FROM ORIN T0 " &
        '        " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
        '        " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
        '        " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
        '        " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.cardcode = T2.cardcode and T4.adrestype = 'S'" &
        '        " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
        '        " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
        '        " LEFT JOIN OCST T7 on T7.Code = T4.State and T7.country = T4.country" &
        '        " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
        '        " where  T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'", startdate, enddate)
        sqlstr = String.Format("SELECT T0.[DocNum] as inv_id, T0.[DocDate] as inv_date , T0.Numatcard as order_no , T0.CardCode as customer_i, T0.CardName as customer_n, '' as 'reportcode',  371 as saleforce, T2.Country as country," &
                "  '' as custtype,T3.SlpName as saleman, T4.Country as shipto, T5.U_SEBcocod as 'product_id',T1.ItemCode as cmmf, " &
                " '' as 'sbu', '' as prodfamily, '' as brand, T5.FrgnName as 'cdesc', '' as 'supplier_i', T1.Quantity as quantity, T1.LineTotal as totalsales, " &
                " T1.StockPrice*T1.Quantity as 'totalcost',T4.City as region ,T7.Name as location, T9.price as 'retail', T1.Price as unit_price, T8.GroupName as business, '' as subfamily," &
                " T5.U_SEBFamLev1CurY as 'familycode', U_SEBfami2 as subfamcode, T5.U_SEBFamLev2CurY," &
                 " cast(T5.U_SEBbran2 as varchar(2)) as 'SEBbran2',T5.U_SEBProdLinePi2" &
                " FROM OINV T0 " &
                " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
                " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.cardcode = T2.Cardcode and T4.adrestype = 'S'" &
                " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                " LEFT JOIN OCST T7 on T7.Code = T4.State  and T7.country = T4.country" &
                " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                " LEFT JOIN ITM1 T9 on T9.ItemCode = T1.ItemCode and T9.pricelist = 2" &
                " where  T0.Doctype = 'I' and T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'" &
                " UNION ALL" &
                " SELECT T0.[DocNum], T0.[DocDate], T0.Numatcard, T0.CardCode, T0.CardName, '' as 'Report Code',  371 as U_SEBSalesForce, T2.Country," &
                "  '' as 'Cust type',T3.SlpName, T4.Country as 'Ship-to Country',T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode, '' as 'SBU', " &
                " '' as prodfamily,'' as brand,T5.FrgnName as 'Pi2 Description', '' as 'Manufacturer', -T1.Quantity, -T1.LineTotal, -T1.StockPrice*T1.Quantity as 'Total Item Cost'," &
                " T4.City ,T7.Name, T9.price as 'Retail', T1.Price, T8.GroupName,  '' as subfamily, cast(T5.U_SEBFamLev1CurY as varchar(3)) as 'FamilyLv1', " &
                " U_SEBfami2 as subfamcode,T5.U_SEBFamLev2CurY,cast(T5.U_SEBbran2 as varchar(2)) as 'SEBbran2',T5.U_SEBProdLinePi2" &
                " FROM ORIN T0 " &
                " LEFT JOIN RIN1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
                " LEFT JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.cardcode = T2.cardcode and T4.adrestype = 'S'" &
                " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                " LEFT JOIN OCST T7 on T7.Code = T4.State and T7.country = T4.country" &
                " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                " LEFT JOIN ITM1 T9 on T9.ItemCode = T1.ItemCode and T9.pricelist = 2" &
                " where  T0.Doctype = 'I' and T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'", startdate, enddate)

        '" INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
        '       " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
        '       " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
        '       " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.cardcode = T2.cardcode and T4.adrestype = 'S'" &
        '       " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
        '       " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
        '       " LEFT JOIN OCST T7 on T7.Code = T4.State and T7.country = T4.country" &
        '       " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
        '       " LEFT JOIN ITM1 T9 on T9.ItemCode = T1.ItemCode and T9.pricelist = 2" &
        '       " where  T0.Doctype = 'I' and T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'", startdate, enddate)

        Return Me.load
    End Function
End Class
