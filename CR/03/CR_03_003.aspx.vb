Public Class CR_03_003
    Inherits AuthBasePage 'System.Web.UI.Page

    '114年確定性需求5：<系統> 產投兩計畫_報表3：19大類課程核定統計
    '    <匯出>
    '1.匯出報表請依照附件格式產出。
    '2.報表要區分產投、自主
    '3.報表名稱：年度 + 申請階段 + 計畫 + "-19大類課程核定情形"
    '如：113年度下半年提升勞工自主學習計畫-19大類課程核定情形
    '4.匯出欄位：序號、課程分類、核定班次、核班比率、核定訓練人次、核定補助費(以80%預估)、核定經費比率
    '5.課程分類：19大類，但注意：【美容、推拿整復類】要再額外區分為美容、推拿整復
    '6.核定比率：【核定班次】/ 【核定班次總計】% 四捨五入到小數第二位
    '7.核定經費比率：【核定補助費】/ 【核定補助費總計】% 四捨五入到小數第二位


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

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.GET_APPSTAGE2_N34(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

    End Sub


    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Public Function SEARCH_DATA1_dt() As DataTable
        'sm As SessionModel, hPMS As Hashtable
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
        'DECLARE @TPLANID NVARCHAR(2)='28';/*2*/ DECLARE @YEARS SMALLINT=CONVERT(SMALLINT,'2024');/*3*/ DECLARE @APPSTAGE NVARCHAR(1)='1';/*4*/ DECLARE @ORGKIND2 NVARCHAR(1)='G';/*1*/
        ' CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}, {"ORGKIND2", v_rblOrgKind2}}
        Dim sSql As String = ""
        sSql = "WITH WK1 AS (
SELECT GCODE GCODE33,CNAME PNAME2 FROM dbo.ID_GOVCLASSCAST3 C WITH(NOLOCK) where PARENTS is null AND GCODE!='05'
UNION SELECT '051' GCODE33,concat(CNAME,'-美容') PNAME2 FROM dbo.ID_GOVCLASSCAST3 C WITH(NOLOCK) where PARENTS is null AND GCODE='05'
UNION SELECT '052' GCODE33,concat(CNAME,'-推拿整復') PNAME2 FROM dbo.ID_GOVCLASSCAST3 C WITH(NOLOCK) where PARENTS is null AND GCODE='05'
)
,WC1 AS (
SELECT pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.DISTID,pp.ORGKIND2 
,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS,pp.TNUM 
,PP.TOTALCOST,PP.DEFGOVCOST 
,pf.CURESULT 
,IG3.GCODE33,IG3.PNAME,IG3.PNAME2,IG3.PTMID
FROM VIEW2B pp 
JOIN V_GOVCLASSCAST3 IG3 ON IG3.GCID3=PP.GCID3
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL /*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND pp.APPSTAGE=@APPSTAGE AND PP.ORGKIND2=@ORGKIND2
)
,WC2 AS (
SELECT CC.GCODE33,CC.PNAME,CC.PNAME2
,COUNT(CASE WHEN CC.CURESULT='Y' THEN 1 END) CLSCNTY
,SUM(CASE WHEN CC.CURESULT='Y' THEN CC.TNUM END) STDCNTY
,SUM(CASE WHEN CC.CURESULT='Y' THEN CC.DEFGOVCOST END) DEFGOVCOSTY
FROM WC1 CC
GROUP BY CC.GCODE33,CC.PNAME,CC.PNAME2
)
,WC2A AS ( SELECT SUM(CLSCNTY) SUM_CLSCNTY,SUM(STDCNTY) SUM_STDCNTY,SUM(DEFGOVCOSTY) SUM_DEFGOVCOSTY FROM WC2  )

SELECT ROW_NUMBER() OVER (ORDER BY K.GCODE33 ASC) AS ROWSEQNO
,K.GCODE33 ,K.PNAME2
/*,case when C.PNAME2 is not null then concat(C.PNAME,'-',C.PNAME2) collate Chinese_Taiwan_Stroke_CI_AS else C.PNAME end PNAME2,C.PNAME,C.PNAME2*/
,ISNULL(C.CLSCNTY,0) CLSCNTY,ISNULL(C.STDCNTY,0) STDCNTY,ISNULL(C.DEFGOVCOSTY,0) DEFGOVCOSTY
,CASE WHEN A.SUM_CLSCNTY>0 THEN CASE WHEN C.CLSCNTY>0 THEN concat(FORMAT(100.0*C.CLSCNTY/A.SUM_CLSCNTY,'0.00'),'%') ELSE '0%' END END CLSRATE1
,CASE WHEN A.SUM_STDCNTY>0 THEN CASE WHEN C.STDCNTY>0 THEN concat(FORMAT(100.0*C.STDCNTY/A.SUM_STDCNTY,'0.00'),'%') ELSE '0%' END END STDRATE1
,CASE WHEN A.SUM_DEFGOVCOSTY>0 THEN CASE WHEN C.DEFGOVCOSTY>0 THEN concat(FORMAT(100.0*C.DEFGOVCOSTY/A.SUM_DEFGOVCOSTY,'0.00'),'%') ELSE '0%' END END DEFGOVCOSTRATE1
,A.SUM_CLSCNTY,A.SUM_STDCNTY,A.SUM_DEFGOVCOSTY
FROM WK1 K
LEFT JOIN WC2 C ON C.GCODE33=K.GCODE33
CROSS JOIN WC2A A
ORDER BY K.GCODE33
"

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, $"--{vbCrLf}{TIMS.GetMyValue5(parms)} --#CR_03_003:{vbCrLf}{sSql}")
        End If
        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        Return dt
    End Function

    ''' <summary> 匯出 </summary>
    Sub EXPORT_4()
        'Dim dtXls As DataTable = Nothing
        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料。")
            Exit Sub
        End If

        Dim dr1 As DataRow = dtXls.Rows(0)

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        '年度 + 申請階段 + 計畫 + "-19大類課程核定情形
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_TitleName2 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "-19大類課程核定情形")

        '匯出excel /ods
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "-19大類課程核定情形")
        Dim sPattern As String = "序號,課程分類,核定班次,核班比率,核定訓練人次,核定補助費(以80%預估),核定經費比率"
        Dim sPatternA As String() = sPattern.Split(",")
        Dim iTitleColSpanCnt As Integer = sPatternA.Length
        Dim sColumn As String = "ROWSEQNO,PNAME2,CLSCNTY,CLSRATE1,STDCNTY,DEFGOVCOSTY,DEFGOVCOSTRATE1"
        Dim sColumnA As String() = sColumn.Split(",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim s_FootHtml2 As String = ""
        s_FootHtml2 &= "<tr>"
        s_FootHtml2 &= String.Format("<td colspan=2>{0}</td>", "總計") '合計
        s_FootHtml2 &= String.Format("<td>{0}</td>", Convert.ToString(dr1("SUM_CLSCNTY"))) '核定班次
        s_FootHtml2 &= String.Format("<td>{0}</td>", "100.00%") '核班比率
        s_FootHtml2 &= String.Format("<td>{0}</td>", Convert.ToString(dr1("SUM_STDCNTY"))) '核定訓練人次
        s_FootHtml2 &= String.Format("<td>{0}</td>", Convert.ToString(dr1("SUM_DEFGOVCOSTY"))) '核定補助費(以80%預估)
        s_FootHtml2 &= String.Format("<td>{0}</td>", "100.00%") '核定經費比率
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

    '匯出 
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Call EXPORT_4()
    End Sub

End Class
