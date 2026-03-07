Public Class CR_02_002
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'iType 1:查詢用 11:匯出(1) 12:匯出(2)
    'Const cst_iType_查詢1 As Integer = 1
    'Const cst_iType_匯出1 As Integer = 11
    'Const cst_iType_匯出2 As Integer = 12
    'Const cst_SCORELEVEL_A As String = "A"
    'Const cst_SCORELEVEL_B As String = "B"
    'Const cst_SCORELEVEL_C As String = "C"
    'Const cst_SCORELEVEL_D As String = "D"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            CCreate1()
        End If

        '委訓
        Select Case sm.UserInfo.LID
            Case 2
                Button2.Visible = False
            Case Else
                'Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
                'If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
                '    '署(局) 或 分署(中心)
                '    TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                '    If HistoryRID.Rows.Count <> 0 Then
                '        center.Attributes("onclick") = "showObj('HistoryList2');"
                '        center.Style("CURSOR") = "hand"
                '    End If
                '    Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
                'End If
                TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                If HistoryRID.Rows.Count <> 0 Then
                    center.Attributes("onclick") = "showObj('HistoryList2');"
                    center.Style("CURSOR") = "hand"
                End If
                Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
                Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End Select
    End Sub

    Sub CCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '訓練機構
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ 

        '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位 '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位

        '初審建議結論 --Y 通過、N 不通過、P 調整後通過
        'ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        'Result 初審建議結論 / 審查結果 - -Y 通過、N 不通過、P 調整後通過
        'ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)
    End Sub

    'Function GET_ORG_SQL1() As String
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
    '    sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
    '    sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
    '    Return sql
    'End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt() As DataTable
        Dim dt As DataTable = Nothing

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        '篩選範圍 1: 不區分 2: 轄區單位 3:  19大類主責課程 SYS_GCODEREVIE
        'Dim v_RBL_RANGE1_SCH As String = TIMS.GetListValue(RBL_RANGE1_SCH)
        '跨區/ 轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        'Dim v_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        '開訓日期
        APPLIEDDATE1.Text = TIMS.Cdate3(APPLIEDDATE1.Text)
        APPLIEDDATE2.Text = TIMS.Cdate3(APPLIEDDATE2.Text)
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)

        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        ElseIf v_APPSTAGE_SCH = "3" AndAlso v_rblOrgKind2 <> "" Then
            v_rblOrgKind2 = ""
        End If

        '--DECLARE @YEARS VARCHAR(4)='2021';DECLARE @TPLANID VARCHAR(3)='28';DECLARE @APPSTAGE NUMERIC(10,0)=2; 
        '--DECLARE @ORGKIND2 VARCHAR(1)='G'; DECLARE @COMIDNO VARCHAR(11)='06313774'; DECLARE @DISTID VARCHAR(4)='006';

        Dim parms As New Hashtable From {
            {"YEARS", sm.UserInfo.Years},
            {"TPLANID", sm.UserInfo.TPlanID},
            {"DISTID", s_DISTID},
            {"APPSTAGE", v_APPSTAGE_SCH}
        }
        If (v_rblOrgKind2 <> "") Then parms.Add("ORGKIND2", v_rblOrgKind2)

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.RID,cc.PSNO28,cc.DISTID" & vbCrLf
        sql &= " ,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= " ,cc.TOTALCOST,cc.DEFGOVCOST,cc.DEFSTDCOST" & vbCrLf
        sql &= " ,cc.TPLANID,cc.YEARS,cc.APPSTAGE" & vbCrLf
        sql &= " ,dbo.FN_GET_DISTNAME(cc.DISTID,3) DISTNAME3,cc.ORGPLANNAME2" & vbCrLf
        sql &= " ,oo.RSID, OO.ORGLEVEL,oo.PLANID,cc.ORGKIND2,OO.ORGID" & vbCrLf
        sql &= " ,OO.COMIDNO" & vbCrLf '統一編號	統一編號" & vbCrLf
        sql &= " ,OO.ORGTYPENAME" & vbCrLf '單位屬性	EX：A－全國性總工會" & vbCrLf
        sql &= " ,OO.ORGNAME" & vbCrLf '訓練單位	單位名稱" & vbCrLf
        sql &= " ,OO.MASTERNAME" & vbCrLf '(二計畫欄位名稱不同) 產投：負責人 提升勞工自主：理事長" & vbCrLf
        'sql &= " ,dbo.FN_GET_CSTUDCNT14(cc.OCID) CSTUDCNT14" & vbCrLf
        sql &= " ,cc.TNUM CSTUDTNUM" & vbCrLf '訓練人次
        '複審通過 'sql &= " ,o2.IMPLEVEL_1 LEVEL1" & vbCrLf(避免重複統編)
        sql &= " ,o2.SECONDCHK" & vbCrLf
        If v_APPSTAGE_SCH = "3" Then
            sql &= " ,dbo.FN_SCORING2_TOTALSCORE(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE) TOTALSCORE" & vbCrLf
            sql &= " ,dbo.FN_SCORING2_RLEVEL_2(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE) RLEVEL2" & vbCrLf
        Else
            '總分-複審等級-通過才顯示
            '計算優先順序：'【分署小計】+署/部加分項目/【匯入成績】'【匯入成績】優先 【分署小計】
            '(1)	當【匯入成績】有資料時 '總分 =【匯入成績】+【署/部加分項目】
            '(2)	當【匯入成績】為空、【分署小計】有資料時 '總分 =【分署小計】+【署/部加分項目】
            '(3)	當【匯入成績】、【分署小計】皆為空時     '總分 = 顯示 "-"
            sql &= " ,CASE WHEN o2.SECONDCHK='Y' THEN case when o2.IMPSCORE_1>0 then convert(varchar,o2.IMPSCORE_1+isnull(o2.BRANCHPNT,0))" & vbCrLf
            sql &= " when o2.SUBTOTAL>0 then convert(varchar,o2.SUBTOTAL+isnull(o2.BRANCHPNT,0)) else '-' end END TOTALSCORE" & vbCrLf
            '複審等級-通過才顯示
            sql &= " ,CASE WHEN o2.SECONDCHK='Y' THEN o2.RLEVEL_2 END RLEVEL2" & vbCrLf
        End If

        sql &= " FROM dbo.VIEW2B cc" & vbCrLf
        sql &= " JOIN dbo.VIEW_ORGPLANINFO oo on oo.RID=cc.RID" & vbCrLf
        If v_APPSTAGE_SCH = "3" Then
            sql &= " LEFT JOIN dbo.ORG_SCORING2 o2 on o2.OSID2=cc.OSID2" & vbCrLf
        Else
            sql &= " LEFT JOIN dbo.ORG_SCORING2 o2 on o2.OSID2=cc.OSID2" & vbCrLf
        End If
        sql &= " WHERE (cc.RESULTBUTTON IS NULL OR cc.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND cc.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)

        sql &= " AND OO.ORGLEVEL=2" & vbCrLf
        'sql &= " AND cc.TPLANID='28'  AND cc.YEARS='2022' AND cc.ORGKIND2='G'" & vbCrLf
        sql &= " AND cc.YEARS=@YEARS AND cc.TPLANID=@TPLANID AND cc.DISTID=@DISTID AND cc.APPSTAGE=@APPSTAGE" & vbCrLf
        If (v_rblOrgKind2 <> "") Then sql &= " AND cc.ORGKIND2=@ORGKIND2" & vbCrLf

        If APPLIEDDATE1.Text <> "" Then
            sql &= " AND cc.APPLIEDDATE >=@APPLIEDDATE1 "
            parms.Add("APPLIEDDATE1", TIMS.Cdate2(APPLIEDDATE1.Text))
        End If
        If APPLIEDDATE2.Text <> "" Then
            sql &= " AND cc.APPLIEDDATE <=@APPLIEDDATE2 "
            parms.Add("APPLIEDDATE2", TIMS.Cdate2(APPLIEDDATE2.Text))
        End If
        If STDate1.Text <> "" Then
            sql &= " AND cc.STDATE >=@STDATE1"
            parms.Add("STDATE1", TIMS.Cdate2(STDate1.Text))
        End If
        If STDate2.Text <> "" Then
            sql &= " AND cc.STDATE <=@STDATE2"
            parms.Add("STDATE2", TIMS.Cdate2(STDate2.Text))
        End If
        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
            sql &= " AND cc.RID =@RID" & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If
        sql &= " )" & vbCrLf

        'COMIDNO (過濾1),cc.LEVEL1(避免重複統編)
        sql &= " ,WO1 AS (SELECT DISTINCT cc.ORGPLANNAME2,cc.COMIDNO,cc.TPLANID,cc.YEARS,cc.APPSTAGE,cc.ORGTYPENAME,cc.ORGNAME,cc.MASTERNAME,cc.RLEVEL2 FROM WC1 cc)" & vbCrLf

        'COMIDNO/APPSTAGE  (過濾2),oo.LEVEL1(避免重複統編)
        sql &= " ,WO2 AS (SELECT oo.ORGPLANNAME2,oo.COMIDNO,oo.APPSTAGE,oo.RLEVEL2 FROM WO1 oo)" & vbCrLf

        'COMIDNO/APPSTAGE  (過濾3)
        sql &= " ,WO3 AS (SELECT cc.COMIDNO,cc.APPSTAGE" & vbCrLf '統一編號	統一編號" & vbCrLf
        sql &= " ,max(cc.ORGTYPENAME) ORGTYPENAME" & vbCrLf '單位屬性	EX：A－全國性總工會" & vbCrLf
        sql &= " ,max(cc.ORGNAME) ORGNAME" & vbCrLf '訓練單位	單位名稱" & vbCrLf
        sql &= " ,max(cc.MASTERNAME) MASTERNAME" & vbCrLf '理事長	(二計畫欄位名稱不同) 產投：負責人 提升勞工自主：理事長" & vbCrLf
        sql &= " FROM WO1 cc" & vbCrLf
        sql &= " GROUP BY cc.COMIDNO,cc.APPSTAGE )" & vbCrLf

        'COMIDNO/APPSTAGE  (過濾4)
        sql &= " ,WC2 AS( SELECT cc.COMIDNO,cc.APPSTAGE" & vbCrLf
        sql &= " ,max(cc.RID) RID ,COUNT(1) CLASSCNT1" & vbCrLf
        sql &= " ,SUM(cc.DEFGOVCOST) DEFGOVCOST ,SUM(cc.CSTUDTNUM) CSTUDTNUM" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " GROUP BY cc.COMIDNO,cc.APPSTAGE )" & vbCrLf

        'COMIDNO/APPSTAGE  (過濾5)
        sql &= " ,WC3 AS ( SELECT cc.COMIDNO,cc.APPSTAGE" & vbCrLf
        sql &= " ,max(cc.RID) RID" & vbCrLf
        sql &= " ,COUNT(CASE WHEN cp.CURESULT ='Y' THEN 1 END) CLASSCNT1B" & vbCrLf '改 核班結果=通過來篩選
        sql &= " ,SUM(CASE WHEN cp.CURESULT ='Y' THEN cc.DEFGOVCOST END) DEFGOVCOSTB" & vbCrLf '改 核班結果=通過來篩選
        sql &= " ,SUM(CASE WHEN cp.CURESULT ='Y' THEN cc.CSTUDTNUM END) CSTUDTNUMB" & vbCrLf '改 核班結果=通過來篩選
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " JOIN PLAN_STAFFOPIN cp on cp.PSNO28=cc.PSNO28" & vbCrLf
        sql &= " GROUP BY cc.COMIDNO,cc.APPSTAGE )" & vbCrLf
        'MAIN
        sql &= " SELECT o2.ORGPLANNAME2,oo.ORGNAME,oo.COMIDNO" & vbCrLf
        sql &= " ,oo.APPSTAGE ,oo.ORGTYPENAME ,oo.MASTERNAME" & vbCrLf
        ',o2.LEVEL1 (避免重複統編)
        sql &= " ,o2.RLEVEL2" & vbCrLf
        sql &= " ,ISNULL(c2.CLASSCNT1,0) CLASSCNT1" & vbCrLf
        sql &= " ,ISNULL(c2.DEFGOVCOST,0) DEFGOVCOST" & vbCrLf
        sql &= " ,ISNULL(c2.CSTUDTNUM,0) CSTUDTNUM" & vbCrLf '訓練人次'CSTUDTNUM-申請

        sql &= " ,ISNULL(c3.CLASSCNT1B,0) CLASSCNT1B" & vbCrLf
        sql &= " ,ISNULL(c3.DEFGOVCOSTB,0) DEFGOVCOSTB" & vbCrLf
        sql &= " ,ISNULL(c3.CSTUDTNUMB,0) CSTUDTNUMB" & vbCrLf '訓練人次'CSTUDTNUMB-核定

        sql &= " ,convert(varchar(30),null) MEMO1" & vbCrLf
        sql &= " FROM WO3 oo" & vbCrLf
        sql &= " JOIN WO2 o2 on o2.COMIDNO=oo.COMIDNO and o2.APPSTAGE=oo.APPSTAGE" & vbCrLf
        sql &= " LEFT JOIN WC2 c2 on c2.COMIDNO=oo.COMIDNO and c2.APPSTAGE=oo.APPSTAGE" & vbCrLf
        sql &= " LEFT JOIN WC3 c3 on c3.COMIDNO=oo.COMIDNO and c3.APPSTAGE=oo.APPSTAGE" & vbCrLf

        'ROW_NUMBER() OVER(ORDER BY cc.ORGNAME,cc.FIRSTSORT,cc.STDATE) SEQNUM
        sql &= " ORDER BY o2.ORGPLANNAME2 desc,oo.ORGNAME,oo.COMIDNO,oo.APPSTAGE" & vbCrLf

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, $"{vbCrLf}--parms{TIMS.GetMyValue5(parms)}{vbCrLf}--##CR_02_002 sql:{vbCrLf}{sql}{vbCrLf}")
        End If
        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        Return dt
    End Function

    ''' <summary> 匯出-申請核定差異統計表 </summary>
    Sub EXPORT_4()
        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        If s_DISTID = "000" Then
            Common.MessageBox(Me, "請選擇有效的轄區分署／機構單位!")
            Exit Sub
        End If
        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If dtXls Is Nothing Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        If dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS

        '年度 + 申請階段 + 計畫 + 訓練課程申請/核定差異統計表
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim sTitle1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "訓練課程申請/核定差異統計表")
        'Dim iColSpanCount As Integer = 12
        ' 產投：負責人 提升勞工自主：理事長
        Dim S_MASTERNM As String = If(v_ddlAPPSTAGE_SCH = "3", "負責人／理事長", If(v_rblOrgKind2 = "G", "負責人", If(v_rblOrgKind2 = "W", "理事長", "理事長")))

        '匯出excel /ods
        Dim s_FILENAME1 As String = String.Concat("申請核定差異統計表_", TIMS.GetDateNo2(3))

        'in_parms.Clear(){"EXP", "Y"}, '匯出查詢條件{"ExpType", v_ExpType} '匯出查詢條件
        Dim in_parms As New Hashtable From {{"EXP", "Y"}, {"ExpType", v_ExpType}}

        Dim iCLASSCNT1 As Integer = 0 '訓練班次'CLASSCNT1-申請
        Dim iCLASSCNT1B As Integer = 0  '訓練班次'CLASSCNT1B-核定

        Dim iCSTUDTNUM As Integer = 0  '訓練人次'CSTUDTNUM-申請
        Dim iCSTUDTNUMB As Integer = 0  '訓練人次'CSTUDTNUMB-核定

        Dim iDEFGOVCOST As Integer = 0  '訓練補助費'DEFGOVCOST-申請
        Dim iDEFGOVCOSTB As Integer = 0  '訓練補助費'DEFGOVCOSTB-核定

        For Each dr1 As DataRow In dtXls.Rows
            iCLASSCNT1 += If(Val(dr1("CLASSCNT1")) > 0, Val(dr1("CLASSCNT1")), 0)
            iCLASSCNT1B += If(Val(dr1("CLASSCNT1B")) > 0, Val(dr1("CLASSCNT1B")), 0)
            iCSTUDTNUM += If(Val(dr1("CSTUDTNUM")) > 0, Val(dr1("CSTUDTNUM")), 0)
            iCSTUDTNUMB += If(Val(dr1("CSTUDTNUMB")) > 0, Val(dr1("CSTUDTNUMB")), 0) '訓練人次'CSTUDTNUMB-核定
            iDEFGOVCOST += If(Val(dr1("DEFGOVCOST")) > 0, Val(dr1("DEFGOVCOST")), 0)
            iDEFGOVCOSTB += If(Val(dr1("DEFGOVCOSTB")) > 0, Val(dr1("DEFGOVCOSTB")), 0)
        Next

        Dim s_TitleHtml2 As String = ""
        s_TitleHtml2 &= "<tr>"
        If v_ddlAPPSTAGE_SCH = "3" Then s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "計畫別") 'COMIDNO 
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "統一編號") 'COMIDNO
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "單位屬性") 'ORGTYPENAME
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "訓練單位") 'ORGNAME
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", S_MASTERNM) 'MASTERNAME  ' 產投：負責人 提升勞工自主：理事長
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "審查計分表等級") '審查計分表等級: (初審)'LEVEL1 / (複審)RLEVEL2
        s_TitleHtml2 &= String.Format("<td colspan=2>{0}</td>", "訓練班次")
        s_TitleHtml2 &= String.Format("<td colspan=2>{0}</td>", "訓練人次")
        s_TitleHtml2 &= String.Format("<td colspan=2>{0}</td>", "訓練補助費")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "備註") 'MEMO1
        s_TitleHtml2 &= "</tr>"
        s_TitleHtml2 &= "<tr>"
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "申請") '訓練班次'CLASSCNT1
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "核定") '訓練班次'CLASSCNTB
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "申請") '訓練人次'CSTUDTNUM
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "核定") '訓練人次'CSTUDTNUMB
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "申請") '訓練補助費'DEFGOVCOST
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "核定") '訓練補助費'DEFGOVCOSTB
        s_TitleHtml2 &= "</tr>"

        Dim s_FootHtml2 As String = ""
        s_FootHtml2 &= "<tr>"
        Dim i_COLSPAN_f1 As Integer = If(v_ddlAPPSTAGE_SCH = "3", 6, 5)
        s_FootHtml2 &= String.Format("<td colspan={0}>{1}</td>", i_COLSPAN_f1, " 合計 ") '訓練班次'CLASSCNT1
        s_FootHtml2 &= String.Format("<td>{0}</td>", iCLASSCNT1) '訓練班次'CLASSCNT1
        s_FootHtml2 &= String.Format("<td>{0}</td>", iCLASSCNT1B) '訓練班次'CLASSCNT1B
        s_FootHtml2 &= String.Format("<td>{0}</td>", iCSTUDTNUM) '訓練人次 'CSTUDTNUM
        s_FootHtml2 &= String.Format("<td>{0}</td>", iCSTUDTNUMB) '訓練人次 'CSTUDTNUMB
        s_FootHtml2 &= String.Format("<td>{0}</td>", iDEFGOVCOST) '訓練補助費'DEFGOVCOST
        s_FootHtml2 &= String.Format("<td>{0}</td>", iDEFGOVCOSTB) '訓練補助費'DEFGOVCOSTB
        s_FootHtml2 &= String.Format("<td>{0}</td>", " ") '"備註") 'MEMO1
        s_FootHtml2 &= "</tr>"

        '審查計分表等級: (初審)'LEVEL1 / (複審)RLEVEL2
        Dim sColumn As String = "COMIDNO,ORGTYPENAME,ORGNAME,MASTERNAME,RLEVEL2,CLASSCNT1,CLASSCNT1B,CSTUDTNUM,CSTUDTNUMB,DEFGOVCOST,DEFGOVCOSTB,MEMO1"
        If v_ddlAPPSTAGE_SCH = "3" Then sColumn = "ORGPLANNAME2,COMIDNO,ORGTYPENAME,ORGNAME,MASTERNAME,RLEVEL2,CLASSCNT1,CLASSCNT1B,CSTUDTNUM,CSTUDTNUMB,DEFGOVCOST,DEFGOVCOSTB,MEMO1"
        'Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        'parms.Add("sPatternA", sPatternA)
        Dim parms As New Hashtable From {
            {"ExpType", v_ExpType}, 'EXCEL/PDF/ODS
            {"FileName", s_FILENAME1},
            {"TitleHtml2", s_TitleHtml2},
            {"FootHtml2", s_FootHtml2},
            {"TitleName", TIMS.ClearSQM(sTitle1)},
            {"TitleColSpanCnt", iColSpanCount},
            {"sColumnA", sColumnA}
        }
        TIMS.Utl_Export(Me, dtXls, parms)
    End Sub

    '匯出  '表單04_申請核定差異統計表.xls
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Call EXPORT_4()
    End Sub

End Class
