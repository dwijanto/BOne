Public Class SalesOrderTWAdapter
    Inherits TaiwanModelAdapter
    Implements IAdapter

    Public Sub New()

    End Sub

    Public Function LoadData() As Boolean Implements IAdapter.LoadData
        Return False
    End Function
    Public Function LoadData(startdate As Date, enddate As Date) As Boolean
        'sqlstr = " SELECT 'AR Invoice' AS 'Type'," &
        '    " T0.[DocDate],T0.[DocNum],T0.Numatcard,T0.U_SEBGUI," &
        '    " T0.CardCode, T0.CardName, T2.CardFName AS 'Store', " &
        '    " T9.AgentName as 'KAM', T3.SlpName as 'KAR', " &
        '    " T2.Country as 'Customer Country', T8.GroupName as 'Customer Group', " &
        '    " T2.U_SEBCurMLACode AS 'MLA Code', " &
        '    " T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode as 'CMMF Code', " &
        '    " T5.FrgnName as 'Pi2 Description', T5.ItemName as 'Local Description'," &
        '    " T5.U_SEBFamLev1CurY as 'Family lv1',T5.U_SEBFamLev2CurY as 'Family lv2'," &
        '    " T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
        '    " T5.CardCode as 'Supplier Code', " &
        '    " T1.Quantity, T1.Price, T1.LineTotal as 'Total Sales', T1.StockPrice*T1.Quantity as 'Total Cost'," &
        '    " T0.U_SEBCNReason as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)'  " &
        '    ",'' as sbu, '' as brand,'' as prodfamily,'' as subfamily" &
        '    " FROM OINV T0 " &
        '    " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
        '    " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
        '    " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
        '    " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
        '    " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
        '    " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
        '    " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
        '    " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
        '    " LEFT JOIN OAGP T9 on T9.AgentCode = T2.AgentCode" &
        '    " UNION ALL" &
        '    " SELECT 'Credit Note' AS 'Type'," &
        '    " T0.[DocDate],T0.[DocNum],T0.Numatcard,T0.U_SEBGUI," &
        '    " T0.CardCode, T0.CardName, T2.CardFName AS 'Store', " &
        '    " T9.AgentName as 'KAM', T3.SlpName as 'KAR', " &
        '    " T2.Country as 'Customer Country', T8.GroupName as 'Customer Group', " &
        '    " T2.U_SEBCurMLACode AS 'MLA Code', " &
        '    " T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode as 'CMMF Code', " &
        '    " T5.FrgnName as 'Pi2 Description', T5.ItemName as 'Local Description'," &
        '    " T5.U_SEBFamLev1CurY as 'Family lv1',T5.U_SEBFamLev2CurY as 'Family lv2'," &
        '    " T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
        '    " T5.CardCode as 'Supplier Code', " &
        '    " -T1.Quantity, T1.Price, -T1.LineTotal as 'Total Sales', -T1.StockPrice*T1.Quantity as 'Total Cost'," &
        '    " T0.U_SEBCNReason as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)'" &
        '     ",'' as sbu, '' as brand,'' as prodfamily,'' as subfamily" &
        '    " FROM ORIN T0 " &
        '    " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
        '    " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
        '    " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
        '    " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
        '    " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
        '    " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
        '    " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
        '    " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
        '    " LEFT JOIN OAGP T9 on T9.AgentCode = T2.AgentCode"
        sqlstr = String.Format(" SELECT 'AR Invoice' AS 'Type'," &
            " T0.[DocDate],T0.[DocNum],T0.Numatcard,T0.U_SEBGUI," &
            " T0.CardCode, T0.CardName, T2.CardFName AS 'Store', " &
            " T9.AgentName as 'KAM', T3.SlpName as 'KAR', " &
            " T2.Country as 'Customer Country', T8.GroupName as 'Customer Group', " &
            " T2.U_SEBCurMLACode AS 'MLA Code', " &
            " T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode as 'CMMF Code', " &
            " T5.FrgnName as 'Pi2 Description', T5.ItemName as 'Local Description'," &
            " T5.U_SEBFamLev1CurY as 'Family lv1',T5.U_SEBFamLev2CurY as 'Family lv2'," &
            " T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
            " T5.CardCode as 'Supplier Code', " &
            " T1.Quantity, T1.Price, T1.LineTotal as 'Total Sales', T1.StockPrice*T1.Quantity as 'Total Cost'," &
            " T0.U_SEBCNReason as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)'  " &
            ",'' as sbu, '' as brand,'' as prodfamily,'' as subfamily" &
            " FROM OINV T0 " &
            " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
            " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
            " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
            " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
            " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
            " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
            " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
            " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
            " LEFT JOIN OAGP T9 on T9.AgentCode = T2.AgentCode" &
             " where  T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'" &
            " UNION ALL" &
            " SELECT 'Credit Note' AS 'Type'," &
            " T0.[DocDate],T0.[DocNum],T0.Numatcard,T0.U_SEBGUI," &
            " T0.CardCode, T0.CardName, T2.CardFName AS 'Store', " &
            " T9.AgentName as 'KAM', T3.SlpName as 'KAR', " &
            " T2.Country as 'Customer Country', T8.GroupName as 'Customer Group', " &
            " T2.U_SEBCurMLACode AS 'MLA Code', " &
            " T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode as 'CMMF Code', " &
            " T5.FrgnName as 'Pi2 Description', T5.ItemName as 'Local Description'," &
            " T5.U_SEBFamLev1CurY as 'Family lv1',T5.U_SEBFamLev2CurY as 'Family lv2'," &
            " T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
            " T5.CardCode as 'Supplier Code', " &
            " -T1.Quantity, T1.Price, -T1.LineTotal as 'Total Sales', -T1.StockPrice*T1.Quantity as 'Total Cost'," &
            " T0.U_SEBCNReason as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)'" &
             ",'' as sbu, '' as brand,'' as prodfamily,'' as subfamily" &
            " FROM ORIN T0 " &
            " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
            " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
            " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
            " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
            " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
            " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
            " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
            " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
            " LEFT JOIN OAGP T9 on T9.AgentCode = T2.AgentCode" &
            " where  T0.[DocDate] >= '{0:yyyy-MM-dd}' and  T0.[DocDate] <= '{1:yyyy-MM-dd}'", startdate, enddate)

        sqlstr = String.Format("SELECT 'AR Invoice' AS 'Type',T0.[DocDate],T0.[DocNum],T0.Numatcard,T0.U_SEBGUI,  T0.CardCode, T0.CardName, T2.AliasName AS 'Store', " &
                 " T2.U_SEBKAM as 'KAM', T3.SlpName as 'KAR', T2.Country as 'Customer Country', T8.GroupName as 'Customer Group', T2.U_SEBCurMLACode AS 'MLA Code', " &
                 " T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode as 'CMMF Code', T5.FrgnName as 'Pi2 Description', T5.ItemName as 'Local Description'," &
                 " T5.U_SEBFamLev1CurY as 'Family lv1',T5.U_SEBFamLev2CurY as 'Family lv2',T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line'," &
                 " T5.CardCode as 'supplier_i', T1.Quantity as quantity, T1.Price, T1.LineTotal as 'Total Sales', T1.StockPrice*T1.Quantity as 'Total Cost'," &
                 " T9.Descr as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)'  " &
                  ",'' as sbu, '' as brand,'' as prodfamily,'' as subfamily,T1.U_SEBposnum" &
                 " FROM OINV T0 " &
                 " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] " &
                 " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                 " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                 " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
                 " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                 " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                 " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
                 " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                 " LEFT JOIN UFD1 T9 on T9.TableID = 'ORIN' AND T9.FieldID = '1' AND T9.FldValue = T0.U_SEBCNReason" &
                 " WHERE T0.Series <> '4' AND T5.ItmsGrpCod <> 103" &
                 " AND (T0.DocDate >= '{0:yyyy-MM-dd}' ) AND (T0.DocDate <= '{1:yyyy-MM-dd}' )" &
                 " UNION ALL" &
                 " SELECT 'Credit Note' AS 'Type',T0.[DocDate],T0.[DocNum],T0.Numatcard,T0.U_SEBGUI,  T0.CardCode, T0.CardName, T2.AliasName AS 'Store', " &
                 " T2.U_SEBKAM as 'KAM', T3.SlpName as 'KAR', T2.Country as 'Customer Country', T8.GroupName as 'Customer Group', T2.U_SEBCurMLACode AS 'MLA Code', " &
                 " T5.U_SEBcocod as 'Pi2 Commercial Code', T1.ItemCode as 'CMMF Code', T5.FrgnName as 'Pi2 Description', T5.ItemName as 'Local Description'," &
                 " T5.U_SEBFamLev1CurY as 'Family lv1',T5.U_SEBFamLev2CurY as 'Family lv2',T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
                 " T5.CardCode as 'supplier_i', -T1.Quantity as quantity, T1.Price, -T1.LineTotal as 'Total Sales', -T1.StockPrice*T1.Quantity as 'Total Cost'," &
                 " T9.Descr as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)'" &
                  ",'' as sbu, '' as brand,'' as prodfamily,'' as subfamily,'' as posnum" &
                 " FROM ORIN T0 " &
                 " INNER JOIN RIN1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
                 " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                 " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                 " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
                 " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                 " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                 " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
                 " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                 " LEFT JOIN UFD1 T9 on T9.TableID = 'ORIN' AND T9.FieldID = '1' AND T9.FldValue = T0.U_SEBCNReason" &
                 " WHERE T0.Series <> '5' AND T5.ItmsGrpCod <> 103" &
                 " AND (T0.DocDate >= '{0:yyyy-MM-dd}' ) AND (T0.DocDate <= '{1:yyyy-MM-dd}' )" &
                 " and T5.qryGroup2 = 'N'", startdate, enddate)

        '" AND (T0.CardCode = '[%2]' OR [%2] = '') and T5.qryGroup2 = 'N'" &
        ' " AND (T0.DocDate >= [%0] OR [%0] = '') AND (T0.DocDate <= '[%1]' OR [%1] = '')" &
        '" AND (T0.CardCode = '[%2]' OR [%2] = '') and T5.qryGroup2 = 'N'", startdate, enddate)

        sqlstr = String.Format("SELECT T0.[DocNum] as inv_id,T0.[DocDate] as inv_date,T0.Numatcard as 'order number',T0.CardCode as customer_i, T0.CardName as customer_n,  T2.Country as 'reportcode',373 as saleforce,'' as country,T8.GroupName as 'custtype', " &
                 " T2.U_SEBKAM as 'saleman', 'TAIWAN' as shipto, T5.U_SEBcocod as 'product_id',T1.ItemCode as 'cmmf','' as sbu,'' as prodfamily,'' as brand,T5.FrgnName as 'cdesc',T5.CardCode as 'supplier_i',T1.Quantity as quantity, T1.LineTotal as 'totalsales',T1.StockPrice*T1.Quantity as 'totalcost',datepart(mm,T0.[DocDate]) as 'Month'," &
                 "'Sales / Return' = CASE WHEN T1.Quantity < 0 THEN 'return' ELSE '' END " &
                 ",'Credit notes' = CASE WHEN T1.Quantity < 0 THEN T1.LineTotal ELSE Null END,T8.GroupName as 'Channel'," &
                 " T10.IndName as 'Customer Name','' as 'E/C',T5.U_SEBFamLev1CurY as 'FamLv 1'," &
                 " T3.SlpName as 'Merchandiser', T2.AliasName AS 'Store Name',  T2.U_SEBCurMLACode AS 'MLA code','' as 'MLA name',T1.U_SEBposnum as posid,'' as od, " &
                 " T5.ItemName as 'Local Description',T5.U_SEBFamLev2CurY as 'Family lv2',T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line'," &
                 "   T1.Price, T9.Descr as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)' , '' as subfamily" &
                 " FROM OINV T0 " &
                 " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] " &
                 " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                 " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                 " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
                 " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                 " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                 " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
                 " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                 " LEFT JOIN OOND T10 on T10.IndCode = T2.IndustryC" &
                 " LEFT JOIN UFD1 T9 on T9.TableID = 'ORIN' AND T9.FieldID = '1' AND T9.FldValue = T0.U_SEBCNReason" &
                 " WHERE T0.Series <> '4' AND T5.ItmsGrpCod <> 103" &
                 " AND (T0.DocDate >= '{0:yyyy-MM-dd}' ) AND (T0.DocDate <= '{1:yyyy-MM-dd}' ) and T5.qryGroup2 = 'N'" &
                 " UNION ALL" &
                 " SELECT T0.[DocNum] as inv_id,T0.[DocDate] as inv_date,T0.Numatcard as 'order number',T0.CardCode as customer_i, T0.CardName as customer_n,  T2.Country as 'reportcode',373 as saleforce,'' as country,T8.GroupName as 'custtype', " &
                 " T2.U_SEBKAM as 'saleman', 'TAIWAN' as shipto, T5.U_SEBcocod as 'product_id',T1.ItemCode as 'cmmf','' as sbu,'' as prodfamily,'' as brand,T5.FrgnName as 'cdesc',T5.CardCode as 'supplier_i',-T1.Quantity as quantity, -T1.LineTotal as 'totalsales',-T1.StockPrice*T1.Quantity as 'totalcost',datepart(mm,T0.[DocDate]) as 'Month'," &
                 " 'Sales / Return' = CASE WHEN T0.U_SEBCNReason IN ('1','2','7','8') THEN 'reversal' " &
                 " WHEN T0.U_SEBCNReason IN ('3','4','5','6') THEN 'return' END" &
                   " ,-T1.LineTotal as 'Credit notes',T8.GroupName as 'Channel'," &
                 " T10.IndName as 'Customer Name','' as 'E/C',T5.U_SEBFamLev1CurY as 'FamLv 1'," &
                 " T3.SlpName as 'Merchandiser', T2.AliasName AS 'storename',  T2.U_SEBCurMLACode AS 'MLA code','' as 'MLA name','' as posid,'' as od, " &
                 " T5.ItemName as 'Local Description',T5.U_SEBFamLev2CurY as 'Family lv2',T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
                 "  T1.Price, T9.Descr as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)', '' as subfamily" &
                 " FROM ORIN T0 " &
                 " INNER JOIN RIN1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
                 " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                 " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                 " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
                 " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                 " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                 " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
                 " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                 " LEFT JOIN OOND T10 on T10.IndCode = T2.IndustryC" &
                 " LEFT JOIN UFD1 T9 on T9.TableID = 'ORIN' AND T9.FieldID = '1' AND T9.FldValue = T0.U_SEBCNReason" &
                 " WHERE T0.Series <> '5' AND T5.ItmsGrpCod <> 103" &
                 " AND (T0.DocDate >= '{0:yyyy-MM-dd}' ) AND (T0.DocDate <= '{1:yyyy-MM-dd}' )" &
                 " and T5.qryGroup2 = 'N'", startdate, enddate)

        sqlstr = String.Format("SELECT T0.[DocNum] as inv_id,T0.[DocDate] as inv_date,T0.Numatcard as 'order number',T0.CardCode as customer_i, T0.CardName as customer_n,  T2.Country as 'reportcode',373 as saleforce,'' as country,T8.GroupName as 'custtype', " &
                 " T2.U_SEBKAM as 'saleman', 'TAIWAN' as shipto, T5.U_SEBcocod as 'product_id',T1.ItemCode as 'cmmf','' as sbu,'' as prodfamily,'' as brand,T5.FrgnName as 'cdesc',T5.CardCode as 'supplier_i',T1.Quantity as quantity, T1.LineTotal as 'totalsales',T1.StockPrice*T1.Quantity as 'totalcost',datepart(mm,T0.[DocDate]) as 'Month'," &
                 " 'Sales' as [Sales / Return]" &
                 ",null as 'Credit notes',T8.GroupName as 'Channel'," &
                 " T10.IndName as 'Customer Name','' as 'E/C',T5.U_SEBFamLev1CurY as 'FamLv 1'," &
                 " T3.SlpName as 'Merchandiser', T2.AliasName AS 'Store Name',  T2.U_SEBCurMLACode AS 'MLA code','' as 'MLA name',T1.U_SEBposnum as posid,'' as od, " &
                 " T5.ItemName as 'Local Description',T5.U_SEBFamLev2CurY as 'Family lv2',T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line'," &
                 "   T1.Price, T9.Descr as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)' , '' as subfamily" &
                 " FROM OINV T0 " &
                 " INNER JOIN INV1 T1 ON T0.[DocEntry] = T1.[DocEntry] " &
                 " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                 " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                 " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
                 " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                 " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                 " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
                 " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                 " LEFT JOIN OOND T10 on T10.IndCode = T2.IndustryC" &
                 " LEFT JOIN UFD1 T9 on T9.TableID = 'ORIN' AND T9.FieldID = '1' AND T9.FldValue = T0.U_SEBCNReason" &
                 " WHERE T0.Series <> '4' AND T5.ItmsGrpCod <> 103" &
                 " AND (T0.DocDate >= '{0:yyyy-MM-dd}' ) AND (T0.DocDate <= '{1:yyyy-MM-dd}' ) and T5.qryGroup2 = 'N'" &
                 " UNION ALL" &
                 " SELECT T0.[DocNum] as inv_id,T0.[DocDate] as inv_date,T0.Numatcard as 'order number',T0.CardCode as customer_i, T0.CardName as customer_n,  T2.Country as 'reportcode',373 as saleforce,'' as country,T8.GroupName as 'custtype', " &
                 " T2.U_SEBKAM as 'saleman', 'TAIWAN' as shipto, T5.U_SEBcocod as 'product_id',T1.ItemCode as 'cmmf','' as sbu,'' as prodfamily,'' as brand,T5.FrgnName as 'cdesc',T5.CardCode as 'supplier_i',-T1.Quantity as quantity, -T1.LineTotal as 'totalsales',-T1.StockPrice*T1.Quantity as 'totalcost',datepart(mm,T0.[DocDate]) as 'Month'," &
                 " 'Sales / Return' = CASE WHEN T0.U_SEBCNReason IN ('1','2','7','8') THEN 'Sales Reversal' " &
                 " WHEN T0.U_SEBCNReason IN ('3','4','5','6') THEN 'Sales Return' END" &
                   " ,-T1.LineTotal as 'Credit notes',T8.GroupName as 'Channel'," &
                 " T10.IndName as 'Customer Name','' as 'E/C',T5.U_SEBFamLev1CurY as 'FamLv 1'," &
                 " T3.SlpName as 'Merchandiser', T2.AliasName AS 'storename',  T2.U_SEBCurMLACode AS 'MLA code','' as 'MLA name','' as posid,'' as od, " &
                 " T5.ItemName as 'Local Description',T5.U_SEBFamLev2CurY as 'Family lv2',T5.U_SEBbran2 as 'Brand Code', T5.U_SEBProdLinePi2 as 'Product Line', " &
                 "  T1.Price, T9.Descr as 'Credit Note Reason', T0.U_SEBCNNumber as 'CN Number (by user)', '' as subfamily" &
                 " FROM ORIN T0 " &
                 " INNER JOIN RIN1 T1 ON T0.[DocEntry] = T1.[DocEntry]" &
                 " INNER JOIN OCRD T2 on T2.CardCode = T0.CardCode" &
                 " LEFT JOIN OSLP T3 ON T3.SlpCode = T0.SlpCode" &
                 " LEFT JOIN CRD1 T4 on T4.Address = T0.ShipToCode and T4.AdresType = 'S' AND T4.CardCode = T0.CardCode" &
                 " LEFT JOIN OITM T5 on T5.ItemCode = T1.ItemCode" &
                 " LEFT JOIN OITB T6 on T6.ItmsGrpCod = T5.ItmsGrpCod" &
                 " LEFT JOIN OCST T7 on T7.Code = T4.State AND T7.Country = T4.Country" &
                 " LEFT JOIN OCRG T8 on T8.GroupCode = T2.GroupCode" &
                 " LEFT JOIN OOND T10 on T10.IndCode = T2.IndustryC" &
                 " LEFT JOIN UFD1 T9 on T9.TableID = 'ORIN' AND T9.FieldID = '1' AND T9.FldValue = T0.U_SEBCNReason" &
                 " WHERE T0.Series <> '5' AND T5.ItmsGrpCod <> 103" &
                 " AND (T0.DocDate >= '{0:yyyy-MM-dd}' ) AND (T0.DocDate <= '{1:yyyy-MM-dd}' )" &
                 " and T5.qryGroup2 = 'N'", startdate, enddate)
        Return Me.load
        Return Me.load
    End Function
End Class
