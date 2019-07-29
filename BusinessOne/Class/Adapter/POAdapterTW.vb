Public Class POAdapterTW
    Inherits ModelAdapterTW
    Public Sub New()
        MyBase.New()
    End Sub

    Public Function loadData(ByVal startdate As Date, ByVal enddate As Date)
        sqlstr = String.Format("Select t3.docentry as docentry ,t3.docdate as postingdate,t3.cardcode as vendorcode,t3.cardname as vendorname,t3.numatcard as refno,t3.doccur as crcy,t3.docrate as exrate,t3.taxdate as documentdate,t0.linenum+1 as rownumber,t0.trgetentry as targetdocgit,t0.linestatus as rowstatus,t0.itemcode as itemno,t0.dscription as description,t0.quantity as qty,t0.shipdate as deliverydate ,t0.openqty as remainingqty,t0.price as price,t0.currency as crcy,t0.rate as exrate,t0.discprcnt as discountpct,t0.linetotal as rowtotallc,t0.totalfrgn as rowtotalfc,t0.opensum as openamountlc,t0.opensumfc as openamountfc,t0.vendornum as vendorcatalogno,t0.linmanclsd as closedmanually,t0.u_sebfqrsq as firstqtyreq,t0.u_sebfdrsq as firstdeldatereq, t4.docstatus as docstatus,t4.invntsttus as whstatus,t4.docdate as gitpostingdate,t4.numatcard as custinvoice,t4.createdate as creationdate,t4.u_sebapcontainertype as containertype,t4.u_sebapcontainer as containernum,t4.u_sebapbl as blnumber,t4.u_sebapdischargedate as dischargedate, t4.u_sebapeta as eta, t4.u_sebapseal as sealnumber,t4.u_sebapvessel as vessel,t1.docentry as gitdocnumber,t1.linenum+1 as gitrownumber,t1.trgetentry as targetporeceipt,t1.baseentry as gitbasedoc,t1.baseline+1 as gitbaserow,t1.linestatus as gitrowstatus,t1.invqty as qtyinventoryuom,t1.openinvqty as openqtyinventoryuom,t5.docstatus as docstatus,t5.invntsttus as whstatus,t5.docdate as porpostingdate,t2.docentry as pordocnum,t2.linenum+1 as porrownumr,t2.baseentry as basedoc,t2.baseline+1 as baserow,t2.linestatus as porrowstatus,t2.quantity as porqty,t2.shipdate as pordeliverydate,t2.openqty as poropenqty,t2.price as porprice,t2.currency as porcrcy,t2.rate as porexrate,t2.discprcnt as pordiscount,t2.linetotal as porrowtotal,t2.whscode as porwhscode,t2.opencreqty as creditmemoamount,t2.vatprcnt as taxrate,t2.vatgroup,t1.basetype" &
                " from por1 t0 " &
                " left join OPOR t3 on t3.docentry = t0.docentry" &
                " left join pch1 t1 on t1.docentry = t0.trgetentry and t1.linenum = t0.linenum" &
                " left join opch t4 on t4.docentry = t1.docentry " &
                " left join pdn1 t2 on t2.docentry = t1.trgetentry and t2.linenum = t1.linenum " &
                " left join opdn t5 on t5.docentry = t2.docentry" &
                " where t3.docdate >= '{0:yyyy-MM-dd}' and t3.docdate <= '{1:yyyy-MM-dd}' order by t3.docentry", startdate, enddate)
        Return MyBase.load
    End Function

    Public Function loadDataV2(ByVal startdate As Date, ByVal enddate As Date)
        sqlstr = String.Format("SELECT T0.CardCode AS 'Supplier Code', T0.CardName AS 'Supplier Name',T0.NumAtCard AS 'Vendor Sales Order', T0.DocDate AS 'Posting Date', T0.DocNum AS 'DocNumber', T1.docEntry AS 'PONumber',T1.LineNum+1 AS 'LineNumber',T1.ItemCode AS 'CMMF Code', T1.Dscription AS 'Desc', T12.U_SEBcocod AS 'Commercial Code', T1.OpenQty AS 'Open Quantity', T1.Price, T1.Currency, T1.Rate, T1.OpenQty*T1.Price AS 'Open Purchase Value', T1.Currency, T1.U_SEBfdrsq AS '1st delivery Date request', T1.U_SEBfqrsq AS '1st quantity request', T1.U_SEBrscod AS 'Last Date confirm',T1.U_SEBrsdl1 AS 'Delivery Date confirm 1', T1.U_SEBrstrq AS 'Last qty Transmit',T0.Comments, T1.trgetEntry AS 'Target Document Internal ID',T1.DocEntry AS 'Document Internal ID','AP R Invoice' AS 'Type',T2.BaseType, T2.BaseEntry,T2.BaseLine+1 AS 'LineNumber', T2.BaseAtCard AS 'Vendor Invoice', '' AS 'Bill of Lading',T2.U_SEBshct1 AS 'Container Number',T2.U_SEBshvn1 AS 'Vessel Name',T2.U_SEBshtrd AS 'Last Shipping Notif Date',T2.Quantity AS 'AP reserve Qty', T2.Quantity*T1.Price AS 'Last Shipping Purchase Value', T1.Currency, T2.U_SEBshdc1 AS 'Shipping Notif at Customer',T2.U_SEBshdd1 AS 'Departure Date',T2.U_SEBshdl1 AS 'ETA',T2.DocEntry AS 'AP Document Internal ID',T2.LineNum+1 AS 'AP Line Number', 'Good Receipt' AS 'Type_GR', T3.BaseEntry, T3.BaseLine+1 AS 'LineNumber', T3.U_SEBshtrd, T3.U_SEBshdc1 AS 'Shipping Notification at Vendor', T3.Quantity AS '1st Qty Ship', T3.U_SEBshdd1 AS '1st Depature Date',T3.U_SEBshdl1 AS '1st Depature ETA', T5.Comments FROM OPOR T0 INNER JOIN POR1 T1 ON T0.DocEntry = T1.docEntry INNER JOIN OITM T12 ON T1.ItemCode = T12.ItemCode LEFT JOIN PCH1 T2 ON T1.DocEntry = T2.BaseEntry AND T1.LineNum = T2.BaseLine LEFT JOIN  PDN1 T3 on T2.DocEntry  = T3.BaseEntry AND T2.LineNum = T3.BaseLine LEFT JOIN OPCH T4 ON T2.DocEntry = T4.DocEntry LEFT JOIN OPDN T5 ON T3.DocEntry = T5.DocEntry" &
                               " WHERE (T0.DocDate BETWEEN '{0:yyyy-MM-dd}' AND '{1:yyyy-MM-dd}') AND T0.CardCode > 99000000 ORDER BY T0.DocEntry, T1.LineNum", startdate, enddate)
        Return MyBase.load
    End Function

    Public Function loadDataTaxInvoice(ByVal startdate As Date, ByVal enddate As Date)
        'sqlstr = String.Format("SELECT T1.cardcode, 'M' as header,T0.U_SEBGUI as [Invoice No],T0.TAXDATE as [Invoice Date],'07' as [Invoice Types],T0.LicTradNum as [Invoice BAN],T1.cardname as [Invoice Name],T2.StreetNo as [Address],'1' as [VAT Type],'5' as [VAT], T0.doctotalsy - T0.vatsumsy as [Sales Amount] ,T0.vatsumsy as [VAT], T0.doctotalsy as [Total],T3.quantity,T3.price,T3.quantity * T3.price as [Amount],T4.suppcatnum as [articlecode],T3.dscription as [description],T0.numatcard as [custref]" &
        '                       " FROM OINV T0 " &
        '                       " INNER JOIN OCRD T1 ON T0.[CardCode] = T1.[CardCode] " &
        '                       " INNER JOIN CRD1 T2 ON T1.[CardCode] = T2.[CardCode] and T2.AdresType='B' and T2.address = '1'" &
        '                       " INNER JOIN INV1 T3	ON T0.[DocEntry] = T3.[DocEntry]" &
        '                       " INNER JOIN OITM T4 ON T3.[ItemCode] = T4.[ItemCode]" &
        '                       " WHERE T0.TAXDATE >= '{0:yyyy-MM-dd}' and T0.TAXDATE <= '{1:yyyy-MM-dd}'  " &
        '                       " and T0.U_SEBGUI  <> '' " &
        '                       " order by T0.U_SEBGUI", startdate, enddate)
        sqlstr = String.Format("SELECT t1.cardcode as [Customer Code], 'M' as [Header],T0.U_SEBGUI as [Invoice No],T0.TAXDATE as [Invoice Date],'07' as [Invoice Types],T0.LicTradNum as [Invoice BAN],T1.cardname as [Invoice Name],T2.StreetNo as [Address],'1' as [VAT Type],'5' as [VAT], T0.doctotalsy - T0.vatsumsy as [Sales Amount] ,T0.vatsumsy as [VATamount], T0.doctotalsy as [Total],T3.quantity,T3.price,T3.quantity * T3.price as [Amount],T4.suppcatnum as [articlecode],T3.dscription as [description],T0.numatcard as [custref]" &
                               " FROM OINV T0 " &
                               " INNER JOIN OCRD T1 ON T0.[CardCode] = T1.[CardCode] " &
                               " INNER JOIN CRD1 T2 ON T1.[CardCode] = T2.[CardCode] and T2.AdresType='B' and T2.address = T1.billtodef" &
                               " INNER JOIN INV1 T3	ON T0.[DocEntry] = T3.[DocEntry]" &
                               " INNER JOIN OITM T4 ON T3.[ItemCode] = T4.[ItemCode]" &
                               " WHERE T0.TAXDATE >= '{0:yyyy-MM-dd}' and T0.TAXDATE <= '{1:yyyy-MM-dd}'  " &
                               " and T0.U_SEBGUI  <> ''" &
                               " and not(T1.cardcode like ('C002%') or T1.cardcode like ('C003%') or T1.cardcode like ('C004%') or T1.cardcode like ('C005%') or T1.cardcode like ('C006%'))" &
                               " order by T0.U_SEBGUI;SELECT T0.[CompnyName], T0.[CompnyAddr], T0.[Phone1] FROM OADM T0;", startdate, enddate)
        Return MyBase.load
    End Function
End Class
