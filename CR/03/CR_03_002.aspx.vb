Public Class CR_03_002
    Inherits AuthBasePage 'System.Web.UI.Page

    '114年確定性需求5：<系統> 產投兩計畫_報表2：政策性產業核定課程統計

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        ddlAPPSTAGE_SCH = TIMS.GET_APPSTAGE2_N34(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

    End Sub

    '匯出 - 匯出統計表
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Call EXPORT_1()
    End Sub

    '匯出 - 匯出明細
    Protected Sub BtnExport2_Click(sender As Object, e As EventArgs) Handles BtnExport2.Click
        Call EXPORT_2()
    End Sub

    Public Function SEARCH_DATA_dt(iType As Integer) As DataTable
        'iType: 1:統計表/2:明細表/21:114統計表/22:114明細表
        Dim dt As DataTable = Nothing
        '訓練機構 '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        If v_rblOrgKind2 = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇計畫")
            Return dt
        End If
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
            Return dt
        End If
        'sSql &= " DECLARE @TPLANID VARCHAR(4) ='28' DECLARE @YEARS VARCHAR(4) ='2024' DECLARE @APPSTAGE INT =1; DECLARE @ORGKIND2 VARCHAR(4) ='W';" & vbCrLf
        'DECLARE @TPLANID VARCHAR(4) ='28' DECLARE @YEARS VARCHAR(4) ='2024'DECLARE @APPSTAGE INT =1; DECLARE @ORGKIND2 VARCHAR(4) ='G';
        'DECLARE @TPLANID VARCHAR(4) ='28' ; DECLARE @YEARS VARCHAR(4) ='2025'; DECLARE @APPSTAGE INT =1; DECLARE @ORGKIND2 VARCHAR(4) ='G';
        ' CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}, {"ORGKIND2", v_rblOrgKind2}}
        Dim sSql As String = ""

        If iType = 1 Then
            sSql = "
WITH WC1 AS ( SELECT pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO
,pp.DISTID,pp.ORGKIND2 
,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS,pp.TNUM 
,PP.TOTALCOST,PP.DEFGOVCOST 
,pf.CURESULT
,D.KID20_1, D.D20KNAME1
FROM VIEW2B pp 
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
JOIN V_PLAN_DEPOT d on d.PLANID=pp.PLANID AND d.COMIDNO=pp.COMIDNO AND d.SEQNO=pp.SEQNO
WHERE pf.CURESULT='Y' AND (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL /*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND pp.APPSTAGE=@APPSTAGE AND PP.ORGKIND2=@ORGKIND2
AND D.KID20_1 IS NOT NULL )
,WA1 AS ( SELECT CONCAT(C.KID20_1,'01') LB,C.KID20_1,C.D20KNAME1,C.DISTID
/*,(SELECT DISTNAME3 FROM V_DISTRICT where DISTID=C.DISTID) DISTNAME*/
,COUNT(1) CLSCNT,SUM(C.TNUM) STDCNT,SUM(TOTALCOST) TOTALCOST,SUM(DEFGOVCOST) DEFGOVCOST
,ROW_NUMBER() OVER (PARTITION BY C.KID20_1 ORDER BY C.DISTID ASC) AS ROWSEQNO
FROM WC1 C 
GROUP BY C.KID20_1,C.D20KNAME1,C.DISTID )
,WA2 AS ( SELECT CONCAT(C.KID20_1,'02') LB,C.KID20_1,C.D20KNAME1,NULL DISTID
,SUM(CLSCNT) CLSCNT,SUM(STDCNT) STDCNT,SUM(TOTALCOST) TOTALCOST,SUM(DEFGOVCOST) DEFGOVCOST
,0 ROWSEQNO
FROM WA1 C 
GROUP BY C.KID20_1,C.D20KNAME1 )
,WA3 AS ( SELECT LB,KID20_1,D20KNAME1,DISTID,CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST,ROWSEQNO FROM WA1 
UNION  
SELECT LB,KID20_1,D20KNAME1,DISTID,CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST,ROWSEQNO FROM WA2 )

SELECT C.LB,C.KID20_1,C.D20KNAME1,C.DISTID,C.CLSCNT,C.STDCNT,C.TOTALCOST,C.DEFGOVCOST ,C.ROWSEQNO
,(SELECT DISTNAME3 FROM V_DISTRICT where DISTID=C.DISTID) DISTNAME, R.ROWCNT
FROM WA3 C
JOIN (SELECT KID20_1,COUNT(1) ROWCNT FROM WA3 GROUP BY KID20_1) R ON R.KID20_1=C.KID20_1
ORDER BY C.LB,C.DISTID
"
        ElseIf iType = 2 Then
            sSql = "
SELECT ROW_NUMBER() OVER (ORDER BY D.KID20_1,pp.DISTID,pp.ORGNAME,pp.CLASSCNAME2) AS ROWSEQNO
,pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.DISTID,pp.ORGKIND2 
,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS
,pp.TNUM STDCNT ,PP.TOTALCOST,PP.DEFGOVCOST 
,pf.CURESULT
,D.KID20_1, D.D20KNAME1
,(SELECT DISTNAME3 FROM V_DISTRICT where DISTID=pp.DISTID) DISTNAME
,pp.ORGNAME,pp.CLASSCNAME2
FROM VIEW2B pp 
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
JOIN V_PLAN_DEPOT d on d.PLANID=pp.PLANID AND d.COMIDNO=pp.COMIDNO AND d.SEQNO=pp.SEQNO
WHERE pf.CURESULT='Y' AND (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL /*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND pp.APPSTAGE=@APPSTAGE AND PP.ORGKIND2=@ORGKIND2
AND D.KID20_1 IS NOT NULL
ORDER BY D.KID20_1,pp.DISTID,pp.ORGNAME,pp.CLASSCNAME2
"
        ElseIf iType = 21 Then
            sSql = "
WITH WC1 AS ( SELECT pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO
,pp.DISTID,pp.ORGKIND2 
,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS,pp.TNUM 
,PP.TOTALCOST,PP.DEFGOVCOST 
,pf.CURESULT
,D.KID25C, D.D25NAME2C
FROM VIEW2B pp 
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
JOIN V_PLAN_DEPOT d on d.PLANID=pp.PLANID AND d.COMIDNO=pp.COMIDNO AND d.SEQNO=pp.SEQNO
WHERE pf.CURESULT='Y' AND (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL /*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND pp.APPSTAGE=@APPSTAGE AND PP.ORGKIND2=@ORGKIND2
AND D.KID25C IS NOT NULL )
,WA1 AS ( SELECT CONCAT(C.KID25C,'01') LB,C.KID25C,C.D25NAME2C,C.DISTID
/*,(SELECT DISTNAME3 FROM V_DISTRICT where DISTID=C.DISTID) DISTNAME*/
,COUNT(1) CLSCNT,SUM(C.TNUM) STDCNT,SUM(TOTALCOST) TOTALCOST,SUM(DEFGOVCOST) DEFGOVCOST
,ROW_NUMBER() OVER (PARTITION BY C.KID25C ORDER BY C.DISTID ASC) AS ROWSEQNO
FROM WC1 C 
GROUP BY C.KID25C,C.D25NAME2C,C.DISTID )
,WA2 AS ( SELECT CONCAT(C.KID25C,'02') LB,C.KID25C,C.D25NAME2C,NULL DISTID
,SUM(CLSCNT) CLSCNT,SUM(STDCNT) STDCNT,SUM(TOTALCOST) TOTALCOST,SUM(DEFGOVCOST) DEFGOVCOST
,0 ROWSEQNO
FROM WA1 C 
GROUP BY C.KID25C,C.D25NAME2C )
,WA3 AS ( SELECT LB,KID25C,D25NAME2C,DISTID,CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST,ROWSEQNO FROM WA1 
UNION  
SELECT LB,KID25C,D25NAME2C,DISTID,CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST,ROWSEQNO FROM WA2 )

SELECT C.LB,C.KID25C,C.D25NAME2C,C.DISTID,C.CLSCNT,C.STDCNT,C.TOTALCOST,C.DEFGOVCOST ,C.ROWSEQNO
,(SELECT DISTNAME3 FROM V_DISTRICT where DISTID=C.DISTID) DISTNAME, R.ROWCNT
FROM WA3 C
JOIN (SELECT KID25C,COUNT(1) ROWCNT FROM WA3 GROUP BY KID25C) R ON R.KID25C=C.KID25C
ORDER BY C.LB,C.DISTID
"
        ElseIf iType = 22 Then
            sSql = "
SELECT ROW_NUMBER() OVER (ORDER BY D.KID25C,pp.DISTID,pp.ORGNAME,pp.CLASSCNAME2) AS ROWSEQNO
,pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.DISTID,pp.ORGKIND2 
,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS
,pp.TNUM STDCNT ,PP.TOTALCOST,PP.DEFGOVCOST 
,pf.CURESULT
,D.KID25C, D.D25NAME2C
,(SELECT DISTNAME3 FROM V_DISTRICT where DISTID=pp.DISTID) DISTNAME
,pp.ORGNAME,pp.CLASSCNAME2
FROM VIEW2B pp 
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
JOIN V_PLAN_DEPOT d on d.PLANID=pp.PLANID AND d.COMIDNO=pp.COMIDNO AND d.SEQNO=pp.SEQNO
WHERE pf.CURESULT='Y' AND (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL /*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND pp.APPSTAGE=@APPSTAGE AND PP.ORGKIND2=@ORGKIND2
AND D.KID25C IS NOT NULL
ORDER BY D.KID25C,pp.DISTID,pp.ORGNAME,pp.CLASSCNAME2
"
        End If

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms)))
            TIMS.WriteLog(Me, String.Concat("--##CR_03_002 sSql:", vbCrLf, sSql))
        End If
        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        Return dt
    End Function

    Private Sub EXPORT_1()
        Dim iType1 As Integer = If(sm.UserInfo.Years >= 2025, 21, 1)
        Dim dtXls As DataTable = SEARCH_DATA_dt(iType1)
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料。")
            Exit Sub
        End If

        Dim iSUMCLSCNT As Integer = 0
        Dim iSUMSTDCNT As Integer = 0
        Dim iSUMTOTALCOST As Integer = 0
        Dim iSUMDEFGOVCOST As Integer = 0
        For Each dr1 As DataRow In dtXls.Rows
            If Convert.ToString(dr1("DISTID")) <> "" Then
                iSUMCLSCNT += TIMS.VAL1(dr1("CLSCNT"))
                iSUMSTDCNT += TIMS.VAL1(dr1("STDCNT"))
                iSUMTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))
                iSUMDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
            End If
        Next

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TitleName2 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "-「政策性產業別大項」課程核定情形")

        '匯出excel /ods
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "-「政策性產業別大項」課程核定情形")
        Dim sPattern As String = "產業別,分署別,核定班次,核定訓練人次,核定訓練費用,核定補助費(以80%預估)"
        Dim sPatternA As String() = sPattern.Split(",")
        Dim iTitleColSpanCnt As Integer = sPatternA.Length
        Dim sColumn As String = "D20KNAME1,DISTNAME,CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST"
        If iType1 = 21 Then sColumn = "D25NAME2C,DISTNAME,CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST"
        Dim sColumnA As String() = sColumn.Split(",")
        Dim iColSpanCount As Integer = sColumnA.Length
        Dim sTDSUBTOTAL As String = "<td colspan=1 align='center'>小計</td>"
        Dim sCHKCOLNM As String = "DISTID"
        Dim sColB As String() = "CLSCNT,STDCNT,TOTALCOST,DEFGOVCOST".Split(",")

        Dim s_FootHtml2 As String = ""
        s_FootHtml2 &= "<tr>"
        s_FootHtml2 &= String.Format("<td colspan=2>{0}</td>", "總計") '合計
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMCLSCNT) '核定班次
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMSTDCNT) '核定訓練人次
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMTOTALCOST) '核定訓練費用
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMDEFGOVCOST) '核定補助費(以80%預估)
        s_FootHtml2 &= "</tr>"

        Dim parms As New Hashtable
        If iType1 = 21 Then
            parms.Add("CR03002_RowSpan2", "Y")
        Else
            parms.Add("CR03002_RowSpan1", "Y")
        End If
        parms.Add("TD_SUBTOTAL", sTDSUBTOTAL)
        parms.Add("CHKCOLNM", sCHKCOLNM)
        parms.Add("sColB", sColB)

        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", s_FILENAME1)
        parms.Add("TitleName", s_TitleName2)
        parms.Add("FootHtml2", s_FootHtml2)
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        Utl_ExportV2(Me, dtXls, parms)
        'TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Private Sub EXPORT_2()
        Dim iType2 As Integer = If(sm.UserInfo.Years >= 2025, 22, 2)
        Dim dtXls As DataTable = SEARCH_DATA_dt(iType2)
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料。")
            Exit Sub
        End If

        'Dim iSUMCLSCNT As Integer = 0
        Dim iSUMSTDCNT As Integer = 0
        Dim iSUMTOTALCOST As Integer = 0
        Dim iSUMDEFGOVCOST As Integer = 0
        For Each dr1 As DataRow In dtXls.Rows
            'iSUMCLSCNT += TIMS.VAL1(dr1("CLSCNT"))
            iSUMSTDCNT += TIMS.VAL1(dr1("STDCNT"))
            iSUMTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))
            iSUMDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
        Next

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TitleName2 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "-「5+2產業」課程核定情形明細")

        '匯出excel /ods
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "-「5+2產業」課程核定情形明細")
        Dim sPattern As String = "序號,產業別,分署別,訓練單位名稱,課程名稱,核定訓練人次,核定訓練費用,核定補助費(以80%預估)"
        Dim sPatternA As String() = sPattern.Split(",")
        Dim iTitleColSpanCnt As Integer = sPatternA.Length
        Dim sColumn As String = "ROWSEQNO,D20KNAME1,DISTNAME,ORGNAME,CLASSCNAME2,STDCNT,TOTALCOST,DEFGOVCOST"
        If iType2 = 22 Then sColumn = "ROWSEQNO,D25NAME2C,DISTNAME,ORGNAME,CLASSCNAME2,STDCNT,TOTALCOST,DEFGOVCOST"
        Dim sColumnA As String() = sColumn.Split(",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim s_FootHtml2 As String = ""
        s_FootHtml2 &= "<tr>"
        s_FootHtml2 &= String.Format("<td colspan=5>{0}</td>", "總計") '合計
        's_FootHtml2 &= String.Format("<td>{0}</td>", iSUMCLSCNT) '核定班次
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMSTDCNT) '核定訓練人次
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMTOTALCOST) '核定訓練費用
        s_FootHtml2 &= String.Format("<td>{0}</td>", iSUMDEFGOVCOST) '核定補助費(以80%預估)
        s_FootHtml2 &= "</tr>"

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", s_FILENAME1)
        parms.Add("TitleName", s_TitleName2)
        parms.Add("FootHtml2", s_FootHtml2)
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        'TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    ''' <summary> 匯出使用-dt </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="dt"></param>
    ''' <param name="htSS"></param>
    Public Shared Sub Utl_ExportV2(ByRef MyPage As Page, ByRef dt As DataTable, ByRef htSS As Hashtable)
        If (MyPage Is Nothing) Then Return
        Dim s_ExpType As String = TIMS.GetMyValue2(htSS, "ExpType") 'EXCEL/PDF/ODS
        'Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        Dim s_FileName As String = TIMS.GetMyValue2(htSS, "FileName")
        'Dim strHTML As String = TIMS.GetMyValue2(htSS, "strHTML")
        'Dim s_ResponseNoEnd As String = TIMS.GetMyValue2(htSS, "ResponseNoEnd")
        'TitleHtml
        Dim s_TitleHtml2 As String = TIMS.GetMyValue2(htSS, "TitleHtml2")
        'FootHtml2
        Dim s_FootHtml2 As String = TIMS.GetMyValue2(htSS, "FootHtml2")
        'Title parms
        Dim s_TitleName As String = TIMS.GetMyValue2(htSS, "TitleName")
        Dim s_TitleColSpanCnt As String = TIMS.GetMyValue2(htSS, "TitleColSpanCnt")
        'Columns
        Dim sPatternA As String() = TIMS.GetMyValue2(htSS, "sPatternA", TIMS.cst_oType_obj)
        Dim sColumnA As String() = TIMS.GetMyValue2(htSS, "sColumnA", TIMS.cst_oType_obj)

        Dim sColB As String() = TIMS.GetMyValue2(htSS, "sColB", TIMS.cst_oType_obj)
        Dim s_CHKCOLNM As String = TIMS.GetMyValue2(htSS, "CHKCOLNM")
        Dim s_TD_SUBTOTAL As String = TIMS.GetMyValue2(htSS, "TD_SUBTOTAL")
        Dim s_CR03002_RowSpan1 As String = TIMS.GetMyValue2(htSS, "CR03002_RowSpan1")
        Dim fg_use_CR03002_RowSpan1 As Boolean = (s_CR03002_RowSpan1 = "Y")
        Dim s_CR03002_RowSpan2 As String = TIMS.GetMyValue2(htSS, "CR03002_RowSpan2")
        Dim fg_use_CR03002_RowSpan2 As Boolean = (s_CR03002_RowSpan2 = "Y")

        Dim fg_use_TitleHtml2 As Boolean = ((s_TitleHtml2 IsNot Nothing) AndAlso (s_TitleHtml2.Length > 0))

        Dim fg_use_FootHtml2 As Boolean = ((s_FootHtml2 IsNot Nothing) AndAlso (s_FootHtml2.Length > 0))

        Dim fg_use_Pattern As Boolean = ((sPatternA IsNot Nothing) AndAlso (sPatternA.Length > 0))
        Dim fg_use_Column As Boolean = ((sColumnA IsNot Nothing) AndAlso (sColumnA.Length > 0))

        Dim fg_use_ColB As Boolean = ((sColB IsNot Nothing) AndAlso (sColB.Length > 0) AndAlso s_CHKCOLNM <> "" AndAlso s_TD_SUBTOTAL <> "")
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Common.MessageBox(MyPage, TIMS.cst_NODATAMsg1)
            Return 'Exit Sub
        End If

        Dim YMDSTR1x As String = DateTime.Now.ToString("ssHHddMMyyyymmss")
        If (s_FileName = "") Then s_FileName = YMDSTR1x
        Dim formatStr1 As String = "{0}.{1}"
        Dim strFileName As String = ""
        Dim s_log1 As String = ""
        Dim select_PrintType As String = s_ExpType.ToUpper()
        Select Case select_PrintType
            Case "EXCEL"
                strFileName = String.Format(formatStr1, s_FileName, "xls")
            Case "PDF"
                strFileName = String.Format(formatStr1, s_FileName, "pdf")
            Case "ODS"
                strFileName = String.Format(formatStr1, s_FileName, "ods")
            Case Else
                s_log1 = String.Format("ExpType(參數有誤)!!{0}", select_PrintType)
                Common.MessageBox(MyPage, s_log1)
                Exit Sub
        End Select
        '轉換UTF8
        strFileName = HttpUtility.UrlEncode(strFileName, System.Text.Encoding.UTF8)

        Dim strMETAXLS As String = ""
        Dim strSTYLE As String = ""
        Dim sbHTML As New StringBuilder
        Select Case select_PrintType
            Case "EXCEL"
                strMETAXLS = "<meta http-equiv='Content-Type' content='application/vnd.ms-excel; charset=UTF-8'>" & vbCrLf
                strSTYLE = "<style> .text { mso-number-format:\@; text-align:center;} td { mso-number-format:\@;} </style>" & vbCrLf
            Case "PDF"
            Case "ODS"
            Case Else
        End Select

        '建置表格內容
        'strHTML &= "<div style='font-family:標楷體;'>" & vbCrLf
        'strHTML &= "<table border='1' cellspacing='0' style='font-family:標楷體;border-collapse:collapse;border:solid thin #000000;'>" & vbCrLf
        sbHTML.Append("<div>" & vbCrLf)
        sbHTML.Append("<table border='1' cellspacing='0' style='border-collapse:collapse;border:solid thin #000000;'>" & vbCrLf)

        '表頭及查詢條件列
        If (s_TitleName <> "" AndAlso s_TitleColSpanCnt <> "") Then
            sbHTML.Append("  <tr><td align='center' style='font-weight:bold' colspan='" & s_TitleColSpanCnt & "'>" & s_TitleName & "</td></tr>" & vbCrLf)
        ElseIf (s_TitleName <> "") Then
            sbHTML.Append(" <tr><td align='center' colspan='" & dt.Columns.Count & "'>" & s_TitleName & "</td></tr>" & vbCrLf)
        ElseIf (s_TitleName = "") AndAlso sColumnA IsNot Nothing AndAlso sColumnA.Length > 0 Then
            sbHTML.Append(" <tr><td align='center' colspan='" & sColumnA.Length & "'>" & YMDSTR1x & "</td></tr>" & vbCrLf)
        Else
            sbHTML.Append(" <tr><td align='center' colspan='" & dt.Columns.Count & "'>" & YMDSTR1x & "</td></tr>" & vbCrLf)
        End If

        '建立表頭
        If fg_use_TitleHtml2 Then
            sbHTML.Append(s_TitleHtml2) '& vbTab
        Else
            sbHTML.Append("<tr align='center' style='font-weight:bold'>" & vbCrLf)
            If (fg_use_Pattern) Then
                For i As Integer = 0 To sPatternA.Length - 1
                    sbHTML.Append(String.Format(" <td>{0}</td>", sPatternA(i))) '& vbTab
                Next
            Else
                'Dim coli As Integer = 0
                For Each col As Data.DataColumn In dt.Columns
                    sbHTML.Append(String.Format(" <td>{0}</td>", col.ColumnName)) '& vbCrLf
                Next
            End If
            sbHTML.Append("</tr>" & vbCrLf)
        End If

        '建立資料面
        If (fg_use_Column) Then
            '建立資料面
            For Each dr As DataRow In dt.DefaultView.Table.Rows
                If fg_use_ColB AndAlso Convert.ToString(dr(s_CHKCOLNM)) = "" Then
                    sbHTML.Append("<tr>")
                    sbHTML.Append(s_TD_SUBTOTAL) '& vbTab
                    For i As Integer = 0 To sColB.Length - 1
                        sbHTML.Append(String.Format("<td>{0}</td>", Convert.ToString(dr(sColB(i))))) '& vbTab
                    Next
                    sbHTML.Append("</tr>")
                ElseIf fg_use_CR03002_RowSpan1 Then
                    sbHTML.Append("<tr>")
                    For i As Integer = 0 To sColumnA.Length - 1
                        If sColumnA(i) = "D20KNAME1" AndAlso Convert.ToString(dr("ROWSEQNO")) = "1" Then
                            Dim s_ROWSPAN As String = Convert.ToString(dr("ROWCNT"))
                            Dim s_D20KNAME1 As String = Convert.ToString(dr("D20KNAME1"))
                            sbHTML.Append($"<td rowspan={s_ROWSPAN}>{s_D20KNAME1}</td>") '& vbTab
                        ElseIf sColumnA(i) = "D20KNAME1" AndAlso Convert.ToString(dr("ROWSEQNO")) <> "1" Then
                        Else
                            sbHTML.Append(String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))) '& vbTab
                        End If
                    Next
                    sbHTML.Append("</tr>")
                ElseIf fg_use_CR03002_RowSpan2 Then
                    sbHTML.Append("<tr>")
                    For i As Integer = 0 To sColumnA.Length - 1
                        If sColumnA(i) = "D25NAME2C" AndAlso Convert.ToString(dr("ROWSEQNO")) = "1" Then
                            Dim s_ROWSPAN As String = Convert.ToString(dr("ROWCNT"))
                            Dim s_D20KNAME1 As String = Convert.ToString(dr("D25NAME2C"))
                            sbHTML.Append($"<td rowspan={s_ROWSPAN}>{s_D20KNAME1}</td>") '& vbTab
                        ElseIf sColumnA(i) = "D25NAME2C" AndAlso Convert.ToString(dr("ROWSEQNO")) <> "1" Then
                        Else
                            sbHTML.Append(String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))) '& vbTab
                        End If
                    Next
                    sbHTML.Append("</tr>")
                Else
                    sbHTML.Append("<tr>")
                    For i As Integer = 0 To sColumnA.Length - 1
                        sbHTML.Append(String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))) '& vbTab
                    Next
                    sbHTML.Append("</tr>")
                End If
            Next
        Else
            '建立資料面
            For Each dr As DataRow In dt.Rows
                sbHTML.Append("<tr>")
                For i As Integer = 0 To dt.Columns.Count - 1
                    Dim s_OneRow As String = dr(i).ToString.Trim '清空白
                    If s_OneRow.IndexOf(vbCrLf) > 0 Then s_OneRow = Replace(s_OneRow, vbCrLf, "") '換行清除
                    sbHTML.Append(String.Format("<td align='center'>{0}</td>", s_OneRow))
                Next
                sbHTML.Append("</tr>")
            Next
        End If

        If fg_use_FootHtml2 Then
            sbHTML.Append(s_FootHtml2) '& vbTab
        End If

        sbHTML.Append("</table>" & vbCrLf)
        sbHTML.Append("</div>" & vbCrLf)

        htSS.Add("strMETAXLS", strMETAXLS)
        htSS.Add("strSTYLE", strSTYLE)
        htSS.Add("strHTML", sbHTML.ToString())
        TIMS.Utl_ExportRp1(MyPage, htSS)
    End Sub

End Class
