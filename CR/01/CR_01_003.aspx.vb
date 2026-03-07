Public Class CR_01_003
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'PLAN_STAFFOPIN / PSOID / PSNO28
    Const Cst_PLAN_STAFFOPIN_pkName As String = "PSOID"
    Const Cst_PLAN_STAFFOPIN_PSNO28 As String = "PSNO28"

    '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
    Dim fg_Work2026x02 As Boolean = False 'TIMS.SHOW_W2026x02(sm)
    'iType 1:查詢用 11:匯出(1) 12:匯出(2)
    Const cst_iType_查詢1 As Integer = 1
    Const cst_iType_匯出1 As Integer = 11
    'Const cst_iType_匯出2 As Integer = 12
    '課程申請流水號	其他應修正內容	送請委員確認內容	初審建議結論
    'PSNO28,OTHFIXCONT,CONFIRMCONT,ST1RESULT
    Dim s_COL_PSNO28 As String = "" '課程申請流水號
    Dim g_IMP_ERR1 As Boolean = False

    '課程審查/一階審查/審查幕僚意見開關機制,產業人才投資方案／PLAN_STAFFOPINSWITCH
    Dim fg_can_staffopinswitch As Boolean = True

    '課程申請流水號	分署幕僚意見 其他應修正內容 初審建議結論 分署確認課程分類
    Const cst_col_PSNO28 As Integer = 0 '課程申請流水號
    Const cst_col_ST1SUGGEST As Integer = 1 '分署幕僚意見
    Const cst_col_OTHFIXCONT As Integer = 2 '其他應修正內容 'Const cst_col_CONFIRMCONT As Integer = 3 '送請委員確認內容
    Const cst_col_ST1RESULT As Integer = 3 '初審建議結論
    Const cst_col_GCODE As Integer = 4 '分署確認課程分類
    Const cst_col_iMaxLen As Integer = 5 '(欄位總數)

    Dim ff3 As String = "" '搜尋暫存使用

    Const cst_DG1CMD_ADD1 As String = "ADD1" '新增/修改
    Const cst_DG1CMD_EDT1 As String = "EDT1" '新增/修改
    Const cst_DG1CMD_VIE1 As String = "VIE1" '查看
    Const cst_DG1CMD_DEL1 As String = "DEL1" '刪除

    'UPDATE PLAN_STAFFOPIN
    'Set GCODE=ig3.gcode31 FROM PLAN_STAFFOPIN pf 
    'join PLAN_PLANINFO pp On pp.psno28=pf.psno28
    'join V_GOVCLASSCAST3 ig3 On ig3.GCID3=pp.GCID3 WHERE pf.GCODE Is null 

    '分署確認課程分類 / 職類課程/ 訓練業別
    Dim dtGCODE3 As DataTable = Nothing

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1
        Call TIMS.OpenDbConn(objconn)
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        fg_Work2026x02 = TIMS.SHOW_W2026x02(sm)

        '審查職類代碼
        dtGCODE3 = TIMS.Get_GOVCODE3dt(objconn)

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
                '    TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                '    If HistoryRID.Rows.Count <> 0 Then
                '        center.Attributes("onclick") = "showObj('HistoryList2');"
                '        center.Style("CURSOR") = "hand"
                '    End If
                '    Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
                'End If
                '署(局) 或 分署(中心)
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
        PanelEdit1.Visible = False

        msg1.Text = ""
        tbDataGrid1.Visible = False

        BtnDELChkBxDG1.Attributes("onclick") = "return confirm('確定要刪除勾選的所有資料?');"

        Const cst_Title_msg1 As String = "當有勾選，於資料匯入時，系統不檢核允許直接覆蓋匯入資料。"
        TIMS.Tooltip(ChkBxCover1, cst_Title_msg1, True)

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim v_APPSTAGE_SCH_DEF As String = "1"
        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, v_APPSTAGE_SCH_DEF)
        'Dim tit_PLANYEARS As String.Concat("計畫年度：", sm.UserInfo.Years)
        TIMS.Tooltip(ddlAPPSTAGE_SCH, String.Concat("計畫年度：", sm.UserInfo.Years))

        '課程審查/一階審查/審查幕僚意見開關機制,產業人才投資方案／PLAN_STAFFOPINSWITCH
        Hid_YEARS_SCH.Value = Convert.ToString(sm.UserInfo.Years)
        Hid_APPSTAGE_SCH.Value = v_APPSTAGE_SCH_DEF
        '課程審查/一階審查/審查幕僚意見開關機制,產業人才投資方案／PLAN_STAFFOPINSWITCH
        Dim aParms As New Hashtable From {{"YEARS", Hid_YEARS_SCH.Value}, {"APPSTAGE", Hid_APPSTAGE_SCH.Value}}
        fg_can_staffopinswitch = TIMS.CAN_STAFFOPINSWITCH_1(objconn, aParms)

        Hid_YN_STAFFOPINSWITCH.Value = If(fg_can_staffopinswitch, TIMS.cst_YES, TIMS.cst_NO)
        Hid_t_STAFFOPINSWITCH.Value = If(Not fg_can_staffopinswitch, "非「審查幕僚意見」開放增修時間，僅可查詢及匯出", String.Concat("「審查幕僚意見」開放增修時間：", TIMS.GetMyValue2(aParms, "SFOPENDATE_ROC")))
        '(匯入是否啟用)'配合增修需求OJT-23092202：<審查幕僚意見開關機制>  功能設定
        '僅在該年度/ 申請階段， 所設定「審查幕僚意見」開放增修時間內， 才可開放分署增修，
        '若不在時間內， 包含： 新增、匯入、修改及刪除按鈕均須反灰不可操作，僅可進行查詢及匯出。
        fg_can_staffopinswitch = If(Hid_YN_STAFFOPINSWITCH.Value = TIMS.cst_YES, True, False)
        BtnIMPORT1.Enabled = fg_can_staffopinswitch
        TIMS.Tooltip(BtnIMPORT1, Hid_t_STAFFOPINSWITCH.Value, True)
        BtnDELChkBxDG1.Enabled = fg_can_staffopinswitch
        TIMS.Tooltip(BtnDELChkBxDG1, Hid_t_STAFFOPINSWITCH.Value, True)

        'Dim dtD As DataTable = TIMS.Get_DISTIDT2(objconn)
        'cblDistid = TIMS.Get_DistID(cblDistid, dtD)
        'cblDistid.Items.Insert(0, New ListItem("全部", 0))
        'Common.SetListItem(cblDistid, sm.UserInfo.DistID)
        ''cblDistid.Enabled = If(sm.UserInfo.DistID <> "000", False, True)
        ''選擇全部轄區
        'cblDistid.Attributes("onclick") = "SelectAll('cblDistid','DistHidden');"

        '訓練機構
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位

        '初審建議結論 / Y 通過、N 不通過、P 調整後通過
        ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)

        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        ddlGCODE = TIMS.Get_GOVCODE3(dtGCODE3, ddlGCODE, False)
    End Sub

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Call sSearch1()
    End Sub

    Function GET_ORG_SQL1() As String
        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @APPSTAGE NUMERIC(10,0)=2; 'DECLARE @TPLANID VARCHAR(4)='28';DECLARE @DISTID VARCHAR(4)='003';
        Dim sql As String = "
SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST,oo.COMIDNO,oo.ORGID FROM ORG_ORGINFO oo WITH(NOLOCK)
"
        Return sql
    End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <param name="iType">iType 1:查詢用 11:匯出(1) 12:匯出(2)</param>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt(ByVal iType As Integer) As DataTable
        'iType 1:查詢用 11:匯出(1) 12:匯出(2) 'Const cst_iType_查詢1 As Integer = 1 'Const cst_iType_匯出1 As Integer = 11 'Const cst_iType_匯出2 As Integer = 12
        Dim dt As DataTable = Nothing

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過
        Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

        'Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        'Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        'Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Hid_YEARS_SCH.Value = Convert.ToString(sm.UserInfo.Years)
        Hid_APPSTAGE_SCH.Value = v_APPSTAGE_SCH
        '課程審查/一階審查/審查幕僚意見開關機制,產業人才投資方案／PLAN_STAFFOPINSWITCH
        Dim aParms As New Hashtable From {{"YEARS", Hid_YEARS_SCH.Value}, {"APPSTAGE", Hid_APPSTAGE_SCH.Value}}
        fg_can_staffopinswitch = TIMS.CAN_STAFFOPINSWITCH_1(objconn, aParms)

        Hid_YN_STAFFOPINSWITCH.Value = If(fg_can_staffopinswitch, TIMS.cst_YES, TIMS.cst_NO)
        Hid_t_STAFFOPINSWITCH.Value = If(Not fg_can_staffopinswitch, "非「審查幕僚意見」開放增修時間，僅可查詢及匯出", String.Concat("「審查幕僚意見」開放增修時間：", TIMS.GetMyValue2(aParms, "SFOPENDATE_ROC")))
        '(匯入是否啟用)'配合增修需求OJT-23092202：<審查幕僚意見開關機制>  功能設定
        '僅在該年度/ 申請階段， 所設定「審查幕僚意見」開放增修時間內， 才可開放分署增修，
        '若不在時間內， 包含： 新增、匯入、修改及刪除按鈕均須反灰不可操作，僅可進行查詢及匯出。
        fg_can_staffopinswitch = If(Hid_YN_STAFFOPINSWITCH.Value = TIMS.cst_YES, True, False)
        BtnIMPORT1.Enabled = fg_can_staffopinswitch
        TIMS.Tooltip(BtnIMPORT1, Hid_t_STAFFOPINSWITCH.Value, True)
        BtnDELChkBxDG1.Enabled = fg_can_staffopinswitch
        TIMS.Tooltip(BtnDELChkBxDG1, Hid_t_STAFFOPINSWITCH.Value, True)

        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        If s_DISTID = "" AndAlso v_RBL_CrossDist_SCH <> "C" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If
        If v_RBL_CrossDist_SCH = "C" Then s_DISTID = ""
        If v_RBL_CrossDist_SCH = "C" Then RIDValue.Value = ""

        '篩選範圍 1:不區分 2:轄區單位 3:19大類主責課程 SYS_GCODEREVIE
        Dim v_RBL_RANGE1_SCH As String = TIMS.GetListValue(RBL_RANGE1_SCH)

        '計畫'TRPlanPoint28
        'Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        '開訓日期
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        '課程申請流水號
        schPSNO28.Text = TIMS.ClearSQM(schPSNO28.Text)

        Dim dtGQ As DataTable = GET_GRADEQUOTA_TB(sm.UserInfo.Years.ToString(), v_APPSTAGE_SCH)

        Dim sql_WORG1 As String = String.Format("WITH WORG1 AS ({0})", GET_ORG_SQL1())

        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @TPLANID VARCHAR(3)='28';DECLARE @APPSTAGE NUMERIC(10,0)=2;
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim parms As New Hashtable From {{"YEARS", sm.UserInfo.Years.ToString()}, {"TPLANID", sm.UserInfo.TPlanID}, {"APPSTAGE", v_APPSTAGE_SCH}}

        Dim sql As String = ""
        sql &= sql_WORG1
        sql &= " SELECT pp.YEARS ,dbo.FN_CYEAR2(pp.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,pp.APPSTAGE ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" & vbCrLf
        sql &= " ,pp.PLANNAME ,pp.PSNO28 ,pp.RID" & vbCrLf
        sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        sql &= " ,pp.OCID" & vbCrLf
        sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf
        sql &= " ,pp.CLASSCNAME" & vbCrLf
        sql &= " ,FORMAT(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,pp.GCID3" & vbCrLf
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf

        sql &= " ,pf.PSOID" & vbCrLf '審查幕僚意見 SEQNO
        sql &= " ,pf.ST1SUGGEST" & vbCrLf '初審幕僚建議/分署幕僚意見
        sql &= " ,pf.OTHFIXCONT" & vbCrLf '其他應修正內容
        sql &= " ,pf.CONFIRMCONT" & vbCrLf '送請委員確認內容
        '初審建議結論	(轄區分署)顯示通過、不通過、調整後通過
        '(19大類主責分署)下拉選單，選項包括：==請選擇==、通過、不通過、調整後通過
        sql &= " ,pf.ST1RESULT" & vbCrLf '初審建議結論
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
        '初審建議結論 Y 通過、N 不通過、P 調整後通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END ST1RESULT_N" & vbCrLf
        sql &= " ,pf.RESULT" & vbCrLf '審查結果
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf
        sql &= " ,gr1.DISTID GRDISTID" & vbCrLf '19大類主責課程 SYS_GCODEREVIE
        sql &= " ,pp.iCAPNUM " & vbCrLf '/iCAP標章證號" & vbCrLf

        '匯出1/'匯出2
        Dim flag_EXP_12 As Boolean = If(iType = cst_iType_匯出1, True, False)
        If flag_EXP_12 Then
            'Dim sql As String = ""
            sql = ""
            sql &= sql_WORG1
            sql &= " SELECT rr.ORGPLANNAME" & vbCrLf '計畫別、" & vbCrLf
            sql &= " ,pp.ORGID ,pp.OCID,pp.COMIDNO" & vbCrLf
            sql &= " ,pp.YEARS" & vbCrLf
            sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
            'sql &= " ,pp.DISTNAME" & vbCrLf '分署別" & vbCrLf
            'sql &= " ,pp.ORGNAME" & vbCrLf '訓練單位名稱" & vbCrLf
            sql &= " ,pp.FIRSTSORT" & vbCrLf ',cc.FIRSTSORT
            sql &= " ,pp.PSNO28" & vbCrLf '課程申請流水號" & vbCrLf
            sql &= " ,pp.CLASSCNAME" & vbCrLf '班級名稱" & vbCrLf
            sql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期" & vbCrLf
            sql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期" & vbCrLf
            '分署確認課程分類 / 職類課程 / 訓練業別
            sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
            '課程分類
            sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
            '訓練業別編碼
            sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf
            '訓練業別名稱
            sql &= " ,ig3.CNAME GCNAME" & vbCrLf
            '分署確認課程分類
            sql &= " ,gc.PFCNAME" & vbCrLf

            sql &= " ,kc.CCNAME" & vbCrLf '訓練職能" & vbCrLf
            sql &= " ,tt.JOBNAME" & vbCrLf '課程分類 JOBNAME/PKNAME12
            sql &= " ,pp.THOURS" & vbCrLf '訓練時數" & vbCrLf
            sql &= " ,pp.TNUM" & vbCrLf '訓練人次" & vbCrLf
            sql &= " ,pp.ACTHUMCOST" & vbCrLf '實際人時成本" & vbCrLf
            sql &= " ,pp.METSUMCOST" & vbCrLf '實際材料費" & vbCrLf
            '/實際材料費比率,pp.METCOSTPER
            sql &= " ,CASE WHEN pp.METCOSTPER>=0 THEN concat(convert(float, pp.METCOSTPER),'%') END METCOSTPER" & vbCrLf '/實際材料費比率" & vbCrLf
            'sql &= " ,dbo.FN_GET_KID20NAME(pp.PLANID,pp.COMIDNO,pp.SEQNO) D20KNAME" & vbCrLf '政府政策性產業
            sql &= " ,dd.D20KNAME,dd.D25KNAME,dd.D26KNAME" & vbCrLf '政府政策性產業
            '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
            If fg_Work2026x02 Then
                '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
                '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
                sql &= " ,dd.D26KNAME1,dd.D26KNAME2,dd.D26KNAME3,dd.D26KNAME4,dd.D26KNAME5,dd.D26KNAME6,dd.D26KNAME7,dd.D26KNAME8,dd.D26KNAME9"
            Else
                '5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策
                sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf '"5+2產業創新計畫"" & vbCrLf
                ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
                sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
            End If
            'sql &= " 5+2產業/新南向政策/台灣AI行動計畫/數位國家創新經濟發展方案/國家資通安全發展方案/前瞻基礎建設計畫" & vbCrLf
            'sql &= " ,dbo.FN_GET_CROSSDIST(cc.YEARS,cc.COMIDNO,cc.APPSTAGE) CROSSDIST" & vbCrLf '/是否跨區提案" & vbCrLf
            'I_CROSSDIST 是否跨區提案
            sql &= " ,CASE wo.I_CROSSDIST WHEN -1 THEN '否' ELSE '是' END CROSSDIST" & vbCrLf '/是否跨區提案" & vbCrLf
            sql &= " ,pp.iCAPNUM " & vbCrLf '/iCAP標章證號" & vbCrLf
            sql &= " ,pf.PSOID" & vbCrLf
            sql &= " ,pf.ST1SUGGEST" & vbCrLf '/、初審幕僚建議" & vbCrLf
            sql &= " ,pf.OTHFIXCONT" & vbCrLf '/其他應修正內容"  /初審綜合意見-其他應修正內容述明
            sql &= " ,pf.CONFIRMCONT" & vbCrLf '/送請委員確認內容 /初審綜合意見-送請委員確認內容述明
            sql &= " ,pf.ST1RESULT" & vbCrLf '/初審建議結論" /初審綜合意見
            '1:通過/2:調整後通過/3:不通過
            sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
            '初審建議結論" & vbCrLf'初審建議結論 Y 通過、N 不通過、P 調整後通過
            sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END ST1RESULT_N" & vbCrLf
            sql &= " ,pf.RESULT" & vbCrLf '審查結果
            sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
            '分署確認課程分類 / 職類課程 / 訓練業別
            sql &= " ,pf.GCODE PFGCODE" & vbCrLf

            sql &= " ,NULL SAMEOCFIXREC" & vbCrLf '同單位同類課程建議修正意見 SAME UNIT	SIMILAR COURSES	SUGGEST	FIX	OPINION	
            sql &= " ,NULL SAMEOCNT1" & vbCrLf '同單位同類課程(班數)
            sql &= " ,NULL REMARK1" & vbCrLf '備註
            '19大類主責課程 SYS_GCODEREVIE
            sql &= " ,gr1.DISTID GRDISTID" & vbCrLf

            sql &= " ,pp.CYCLTYPE" & vbCrLf '/期別" & vbCrLf
            sql &= " ,pp.TOTALCOST" & vbCrLf '/每班總訓練費(元)" & vbCrLf
            '訓練時數 48小時以上(含)：160  <48(不含)：184  (+15%，160*1.15=184)
            sql &= " ,CASE WHEN pp.THOURS>=48 THEN 160 ELSE 184 END UPMANCOST" & vbCrLf '/人時成本上限
            sql &= " ,ig3.UPPERATE" & vbCrLf ' /材料費上限" & vbCrLf
            'sql &= " ,pp.METCOSTPER" & vbCrLf '/實際材料費比率" & vbCrLf
            sql &= " ,pp.TOTAL" & vbCrLf '/每人訓練費" & vbCrLf
            sql &= " ,pp.POINTYN" & vbCrLf '/是否為學分班(Y/N)" & vbCrLf
            sql &= " ,pp.ISiCAPCOUR" & vbCrLf '/是否為iCAP課程" & vbCrLf
            'sql &= " ,dd.KID12,dd.KNAME12" & vbCrLf '/課程分類" & vbCrLf
            'sql &= " ,o2.RLEVEL_2" & vbCrLf '審查計分等級" & vbCrLf
            'sql &= " ,dbo.FN_GET_CLASSQUOTA(pp.ORGKIND2,pp.YEARS,pp.APPSTAGE,pp.OSID2) CLASSQUOTA" & vbCrLf '/等級額度核配上限" & vbCrLf
            sql &= " ,dbo.FN_SCORING2_RLEVEL_2(pp.COMIDNO,pp.TPLANID,pp.DISTID,pp.YEARS,pp.APPSTAGE) RLEVEL_2" & vbCrLf '審查計分等級" & vbCrLf
            'CLASSQUOTA
            sql &= " ,convert(int,NULL) CLASSQUOTA" & vbCrLf '可核配上限
            'sql &= " ,dbo.FN_SCORING2_UPLIMIT(pp.COMIDNO,pp.TPLANID,pp.YEARS,pp.APPSTAGE,pp.ORGKIND2) UPLIMIT" & vbCrLf '可核配上限,等級額度核配上限" & vbCrLf
            'sql &= " /,dbo.FN_SCORING2_GRADE(pp.COMIDNO,pp.TPLANID,pp.YEARS,pp.APPSTAGE) GRADE" & vbCrLf
            '匯出之【單位屬性】欄位：目前抓錯欄位，應該要抓「訓練機構設定」下方訓練機構屬性設定的機構別
            sql &= " ,pp.ORGTYPENAME" & vbCrLf '機構別/單位屬性。" & vbCrLf
            sql &= " ,pp.ORGKIND2,pp.ORGTYPENAME2" & vbCrLf '機構別/單位屬性。" & vbCrLf
        End If

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=pp.TMID" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " JOIN WORG1 wo on wo.ORGID=pp.ORGID" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.ORG_SCORING2 o2 on o2.OSID2=pp.OSID2" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        '審查計分等級'19大類主責課程 SYS_GCODEREVIE
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf

        sql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sql &= " AND pp.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND pp.YEARS=@YEARS" & vbCrLf
        sql &= " AND pp.APPSTAGE=@APPSTAGE" & vbCrLf

        '僅限有iCAP標章課程
        If CB1_ICAPSTAMP.Checked Then sql &= " AND pp.iCAPNUM IS NOT NULL AND LEN(pp.iCAPNUM)>1" & vbCrLf

        '課程申請流水號
        If schPSNO28.Text <> "" Then
            sql &= " AND pp.PSNO28=@PSNO28" & vbCrLf
            parms.Add("PSNO28", schPSNO28.Text)
        End If

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過
        'Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        Select Case v_RBL_ST1RESULT_SCH
            Case "2"
                sql &= " AND pf.ST1RESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.ST1RESULT IS NULL" & vbCrLf
            Case "Y", "N", "P"
                sql &= " AND pf.ST1RESULT=@ST1RESULT" & vbCrLf
                parms.Add("ST1RESULT", v_RBL_ST1RESULT_SCH)
        End Select

        If sm.UserInfo.LID = 0 AndAlso s_DISTID = "000" Then s_DISTID = ""
        If s_DISTID <> "" Then
            '篩選範圍 1:不區分 2:轄區單位 3:19大類主責課程 SYS_GCODEREVIE
            Select Case Val(v_RBL_RANGE1_SCH)
                Case 1
                    sql &= " AND (pp.DISTID=@DISTID OR gr1.DISTID=@DISTID)" & vbCrLf
                    parms.Add("DISTID", s_DISTID)
                Case 2
                    sql &= " AND pp.DISTID=@DISTID" & vbCrLf
                    parms.Add("DISTID", s_DISTID)
                Case 3
                    sql &= " AND gr1.DISTID=@DISTID" & vbCrLf
                    parms.Add("DISTID", s_DISTID)
            End Select
        End If

        ''跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        'Dim v_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH) FN_GET_CROSSDIST
        'Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                RIDValue.Value = ""
                sql &= " and wo.I_CROSSDIST !=-1" & vbCrLf
            Case "J" 'J:轄區提案單位
                sql &= " and wo.I_CROSSDIST =-1" & vbCrLf
        End Select

        '計畫'TRPlanPoint28
        If TRPlanPoint28.Visible Then
            Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
            Select Case v_rblOrgKind2'rblOrgKind2.SelectedValue
                Case "G", "W"
                    sql &= " and pp.ORGKIND2=@ORGKIND2" & vbCrLf
                    parms.Add("ORGKIND2", v_rblOrgKind2)
            End Select
        End If

        If STDate1.Text <> "" Then
            sql &= " and pp.STDATE >=@STDATE1 "
            parms.Add("STDATE1", TIMS.Cdate2(STDate1.Text))
        End If
        If STDate2.Text <> "" Then
            sql &= " and pp.STDATE <=@STDATE2 "
            parms.Add("STDATE2", TIMS.Cdate2(STDate2.Text))
        End If

        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
            sql &= " AND pp.RID =@RID" & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If

        'ROW_NUMBER()  OVER(ORDER BY cc.ORGNAME,cc.FIRSTSORT,cc.STDate) SEQNUM
        sql &= " ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE" & vbCrLf

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms), vbCrLf, "--##CR_01_003:", vbCrLf, sql))
        End If

        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If flag_EXP_12 AndAlso dtGQ IsNot Nothing AndAlso dtGQ.Rows.Count > 0 Then
            For Each dr1 As DataRow In dt.Rows
                Dim iCLASSQUOTA As Integer = GET_CLASSQUOTA(dtGQ, dr1)
                If (iCLASSQUOTA > 0) Then dr1("CLASSQUOTA") = iCLASSQUOTA
            Next
        End If

        Return dt
    End Function

    Function GET_CLASSQUOTA(dtGQ As DataTable, dr1 As DataRow) As Integer
        Dim rst As Integer = 0
        If dr1 Is Nothing Then Return rst
        Dim ORGKIND2 As String = Convert.ToString(dr1("ORGKIND2"))
        Dim RLEVEL_2 As String = Convert.ToString(dr1("RLEVEL_2"))
        If RLEVEL_2 = "" OrElse ORGKIND2 = "" Then Return rst
        Dim fff As String = String.Concat("SCORELEVEL='", RLEVEL_2, "'")
        Select Case Convert.ToString(dr1("ORGKIND2"))
            Case "G"
                fff &= " AND CLASSQUOTAG>0"
                If dtGQ.Select(fff).Length > 0 Then
                    rst = Val(dtGQ.Select(fff)(0)("CLASSQUOTAG"))
                    Return rst
                End If

            Case "W"
                fff &= " AND CLASSQUOTAW>0"
                If dtGQ.Select(fff).Length > 0 Then
                    rst = Val(dtGQ.Select(fff)(0)("CLASSQUOTAW"))
                    Return rst
                End If

        End Select

        Return rst
    End Function
    Function GET_GRADEQUOTA_TB(YEARS As String, s_APPSTAGE As String) As DataTable
        Dim hPMS As New Hashtable From {{"YEARS", YEARS}, {"APPSTAGE", Val(s_APPSTAGE)}}
        Dim sSql As String = " SELECT SGQID,SCORELEVEL,CLASSQUOTAG,CLASSQUOTAW FROM SYS_GRADEQUOTA WHERE YEARS=@YEARS AND APPSTAGE=@APPSTAGE" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, hPMS)
        Return dt1
    End Function

    ''' <summary>查詢</summary>
    Sub sSearch1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False

        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return
        End If

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        tbDataGrid1.Visible = False
        msg1.Text = TIMS.cst_NODATAMsg1

        Dim dt As DataTable = SEARCH_DATA1_dt(cst_iType_查詢1)

        'msg1.Text = TIMS.cst_NODATAMsg1
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        msg1.Text = ""
        tbDataGrid1.Visible = True
        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>提案彙整表</summary>
    Sub EXPORT_1()
        Dim dtX1 As DataTable = SEARCH_DATA1_dt(cst_iType_匯出1)

        Dim sb_Pattern As New StringBuilder '序號,
        sb_Pattern.Append("計畫別,分署別,訓練單位名稱,課程申請流水號,班級名稱,期別,開訓日期,結訓日期")
        sb_Pattern.Append(",訓練業別編碼,訓練業別名稱,訓練職能,訓練時數,訓練人次,每班總訓練費(元),人時成本上限,實際人時成本,材料費上限,實際材料費比率,實際材料費,每人訓練費")
        'sb_Pattern.Append(",5+2產業,新南向政策,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,是否為學分班(Y/N),是否跨區提案")
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            '2026 政府政策性產業 (產投)
            sb_Pattern.Append(",五大信賴產業推動方案,六大區域產業及生活圈,台灣AI行動計畫2.0,智慧國家2.0鋼領,國家人才競爭力躍升方案,新南向政策推動計畫,AI新十大建設推動方案,智慧機器人產業推動方案,臺灣2050淨零轉型")
        Else
            sb_Pattern.Append(",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航")
        End If
        sb_Pattern.Append(",是否為學分班(Y/N),是否跨區提案")
        sb_Pattern.Append(",iCAP標章證號,是否為iCAP課程,初審幕僚建議-分署幕僚意見,初審幕僚建議-其他應修正內容,初審建議結論")
        sb_Pattern.Append(",課程分類,分署確認課程分類,審查計分等級,等級額度核配上限,單位屬性,統一編號")

        'CONFIRMCONT 送請委員確認內容 送請委員確認內容
        Dim sb_Column As New StringBuilder
        sb_Column.Append("ORGPLANNAME,DISTNAME,ORGNAME,PSNO28,CLASSCNAME,CYCLTYPE,STDATE,FTDATE")
        sb_Column.Append(",GCODENAME,GCNAME,CCNAME,THOURS,TNUM,TOTALCOST,UPMANCOST,ACTHUMCOST,UPPERATE,METCOSTPER,METSUMCOST,TOTAL")
        'sb_Column.Append(",D20KNAME1,D20KNAME6,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5,POINTYN,CROSSDIST")
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            '2026 政府政策性產業 (產投)
            sb_Column.Append(",D26KNAME1,D26KNAME2,D26KNAME7,D26KNAME3,D26KNAME5,D26KNAME4,D26KNAME6,D26KNAME8,D26KNAME9")
        Else
            sb_Column.Append(",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,D25KNAME7,D25KNAME8")
        End If
        sb_Column.Append(",POINTYN,CROSSDIST")
        sb_Column.Append(",iCAPNUM,ISiCAPCOUR,ST1SUGGEST,OTHFIXCONT,ST1RESULT_N")
        sb_Column.Append(",GCODEPNAME,PFCNAME,RLEVEL_2,CLASSQUOTA,ORGTYPENAME2,COMIDNO")

        Dim sPatternA() As String = Split(sb_Pattern.ToString(), ",")
        Dim sColumnA() As String = Split(sb_Column.ToString(), ",")

        'Dim s_FILENAME1 As String = String.Concat("跨區提案彙整表_", TIMS.GetDateNo2(3))
        Dim s_FILENAME1 As String = String.Concat("提案彙整表_", TIMS.GetDateNo2(3))

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim sbHTML As New StringBuilder ' = ""
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr &= "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "序號") '& vbTab
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        sbHTML.Append(ExportStr)

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dtX1.Rows
            iNum += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", iNum) '& vbTab
            For i As Integer = 0 To sColumnA.Length - 1
                ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr(sColumnA(i))))
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)}, 'EXCEL/PDF/ODS
            {"FileName", s_FILENAME1},
            {"strSTYLE", strSTYLE},
            {"strHTML", sbHTML.ToString()},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'TIMS.CloseDbConn(objconn) 'Response.End()
    End Sub

    Private Sub DISABLE_SHOW1(ByRef dr1 As DataRow, ByRef sCMDNM As String)
        If dr1 Is Nothing Then Return

        Dim flag_OTHFIXCONT_OK As Boolean = If(sm.UserInfo.DistID = Convert.ToString(dr1("DISTID")), True, False)
        'Dim flag_CONFIRMCONT_OK As Boolean = If(sm.UserInfo.DistID = Convert.ToString(dr1("GRDISTID")), True, False)
        'Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'If Not flag_OTHFIXCONT_OK AndAlso flagS1 Then flag_OTHFIXCONT_OK = True

        ST1SUGGEST.Enabled = flag_OTHFIXCONT_OK '分署幕僚意見
        OTHFIXCONT.Enabled = flag_OTHFIXCONT_OK '其他應修正內容
        'CONFIRMCONT.Enabled = flag_CONFIRMCONT_OK '送請委員確認內容
        ddlST1RESULT.Enabled = flag_OTHFIXCONT_OK '初審建議結論 /通過、不通過、調整後通過
        'ddlGCODE.Enabled = flag_OTHFIXCONT_OK '分署確認課程分類

        '(1)各分署可針對自己所屬轄區之訓練單位，填寫「其他應修正內容」。
        '(2)19大類的主責分署(設定於功能：首頁>> 課程審查 >> 一階審查 >> 19大類審查分署設定)，可填寫「送請委員確認內容」、「初審建議結論」
        If (Not ST1SUGGEST.Enabled) Then TIMS.Tooltip(ST1SUGGEST, "所屬轄區之訓練單位，可填寫「分署幕僚意見」", True)
        If (Not OTHFIXCONT.Enabled) Then TIMS.Tooltip(OTHFIXCONT, "所屬轄區之訓練單位，可填寫「其他應修正內容」", True)
        'If (Not CONFIRMCONT.Enabled) Then TIMS.Tooltip(CONFIRMCONT, "19大類的主責分署，可填寫「送請委員確認內容」", True)
        If (Not ddlST1RESULT.Enabled) Then TIMS.Tooltip(ddlST1RESULT, "所屬轄區之訓練單位，可填寫「初審建議結論」", True)
        'If (Not ddlGCODE.Enabled) Then TIMS.Tooltip(ddlST1RESULT, "所屬轄區之訓練單位，可填寫「分署確認課程分類」", True)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing Then Return
        Dim sCmdArg As String = e.CommandArgument
        Dim sCMDNM As String = e.CommandName
        If sCmdArg = "" OrElse sCMDNM = "" Then Return

        Call CLEAR_DATA1()

        Hid_PSOID.Value = TIMS.GetMyValue(sCmdArg, "PSOID")
        Hid_PSNO28.Value = TIMS.GetMyValue(sCmdArg, "PSNO28")
        Hid_GCODE.Value = TIMS.GetMyValue(sCmdArg, "GCODE") 'vGCODE
        Hid_PFGCODE.Value = TIMS.GetMyValue(sCmdArg, "PFGCODE") 'vPFGCODE
        Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))
        If Hid_PSNO28.Value = "" Then Return

        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"PSNO28", Hid_PSNO28.Value}}
        If Hid_PSOID.Value <> "" Then parms.Add("PSOID", Hid_PSOID.Value)
        Dim dr1 As DataRow = GET_DATA1(objconn, parms)

        Select Case sCMDNM'e.CommandName
            Case cst_DG1CMD_ADD1, cst_DG1CMD_EDT1 ' "ADD1", "EDT1" '新增/修改
                btnSAVE1.Visible = True
                Call SHOW_DATA1(dr1)
                Call DISABLE_SHOW1(dr1, sCMDNM)

            Case cst_DG1CMD_VIE1 '"VIE1" '查看
                btnSAVE1.Visible = False
                Call SHOW_DATA1(dr1)
                Call DISABLE_SHOW1(dr1, sCMDNM)

            Case cst_DG1CMD_DEL1 ' "DEL1"
                Call DEL_DATA1(dr1)

            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg9)

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim labserial_number As Label = e.Item.FindControl("labserial_number")
                Dim ChkBxDG1_DEL As CheckBox = e.Item.FindControl("ChkBxDG1_DEL")
                Dim ChkBxDG1_ICAP As CheckBox = e.Item.FindControl("ChkBxDG1_ICAP")
                Dim Hid_DataKey As HiddenField = e.Item.FindControl("Hid_DataKey")
                Dim Hid_DataKey2 As HiddenField = e.Item.FindControl("Hid_DataKey2")

                'e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                labserial_number.Text = TIMS.Get_DGSeqNo(sender, e) '序號

                ChkBxDG1_DEL.Visible = If(Convert.ToString(drv("PSOID")) <> "", True, False)
                ChkBxDG1_ICAP.Visible = If(Convert.ToString(drv("iCAPNUM")) <> "", True, False)

                Hid_DataKey.Value = TIMS.EncryptAes(drv(Cst_PLAN_STAFFOPIN_pkName)) 'PSOID
                Hid_DataKey2.Value = TIMS.EncryptAes(Convert.ToString(drv("PSNO28"))) 'PSNO28

                '初審建議結論 Y 通過、N 不通過、P 調整後通過
                Dim labST1RESULT_N As Label = e.Item.FindControl("labST1RESULT_N")
                labST1RESULT_N.Text = Convert.ToString(drv("ST1RESULT_N")) 'TIMS.Get_ST1RESULT_N(Convert.ToString(drv("ST1RESULT_N")))

                '新增、修改、刪除、查看
                Dim BtnADD1 As Button = e.Item.FindControl("BtnADD1")    '新增
                Dim BtnEDT1 As Button = e.Item.FindControl("BtnEDT1")    '修改
                Dim BtnDEL1 As Button = e.Item.FindControl("BtnDEL1")    '刪除
                Dim BtnVIE1 As Button = e.Item.FindControl("BtnVIE1")    '查看
                BtnADD1.Visible = If(Convert.ToString(drv("PSOID")) = "", True, False)
                BtnEDT1.Visible = If(Convert.ToString(drv("PSOID")) <> "", True, False)
                BtnDEL1.Visible = If(Convert.ToString(drv("PSOID")) <> "", True, False)
                BtnVIE1.Visible = If(Convert.ToString(drv("PSOID")) <> "", True, False)

                '課程審查/一階審查/審查幕僚意見開關機制,產業人才投資方案／PLAN_STAFFOPINSWITCH
                '(匯入是否啟用)'配合增修需求OJT-23092202：<審查幕僚意見開關機制>  功能設定
                '僅在該年度/ 申請階段， 所設定「審查幕僚意見」開放增修時間內， 才可開放分署增修，
                '若不在時間內， 包含： 新增、匯入、修改及刪除按鈕均須反灰不可操作，僅可進行查詢及匯出。
                fg_can_staffopinswitch = If(Hid_YN_STAFFOPINSWITCH.Value = TIMS.cst_YES, True, False)
                '僅限有iCAP標章課程
                Dim fg_can_icapsendswitch As Boolean = If(Convert.ToString(drv("iCAPNUM")) <> "", True, False)
                ChkBxDG1_ICAP.Enabled = fg_can_icapsendswitch
                TIMS.Tooltip(ChkBxDG1_ICAP, "僅限有iCAP標章課程", True)

                ChkBxDG1_DEL.Enabled = fg_can_staffopinswitch
                TIMS.Tooltip(ChkBxDG1_DEL, Hid_t_STAFFOPINSWITCH.Value, True)
                BtnIMPORT1.Enabled = fg_can_staffopinswitch
                TIMS.Tooltip(BtnIMPORT1, Hid_t_STAFFOPINSWITCH.Value, True)
                BtnADD1.Enabled = fg_can_staffopinswitch
                TIMS.Tooltip(BtnADD1, Hid_t_STAFFOPINSWITCH.Value, True)
                BtnEDT1.Enabled = fg_can_staffopinswitch
                TIMS.Tooltip(BtnEDT1, Hid_t_STAFFOPINSWITCH.Value, True)
                BtnDEL1.Enabled = fg_can_staffopinswitch
                TIMS.Tooltip(BtnDEL1, Hid_t_STAFFOPINSWITCH.Value, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PSOID", Convert.ToString(drv("PSOID")))
                TIMS.SetMyValue(sCmdArg, "PSNO28", Convert.ToString(drv("PSNO28")))
                TIMS.SetMyValue(sCmdArg, "GCODE", Convert.ToString(drv("GCODE")))
                TIMS.SetMyValue(sCmdArg, "PFGCODE", Convert.ToString(drv("PFGCODE")))
                BtnADD1.CommandArgument = sCmdArg
                BtnEDT1.CommandArgument = sCmdArg
                BtnDEL1.CommandArgument = sCmdArg
                BtnVIE1.CommandArgument = sCmdArg
                BtnDEL1.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End Select
    End Sub

    Sub CLEAR_DATA1()
        lbYEARS_ROC.Text = ""
        lbDistName.Text = ""
        lbOrgName.Text = ""
        lbPSNO28.Text = ""
        lbClassName.Text = ""
        lbSFTDate.Text = ""
        '分署確認課程分類 / 職類課程 / 訓練業別
        lbGCODEPNAME.Text = ""
        'lbGCNAME.Text = ""
        lbCCNAME.Text = ""
        lbTNum.Text = ""
        lbTHours.Text = ""
        lbACTHUMCOST.Text = ""
        lbMETSUMCOST.Text = ""
        lbIsCROSSDIST.Text = ""
        lbiCAPNUM.Text = "" 'iCAP標章課程
        lbD20KNAME.Text = "" '$"{dr("D20KNAME")}" '政府政策性產業
        lbD25KNAME.Text = "" '$"{dr("D25KNAME")}" '政府政策性產業
        lbD26KNAME.Text = "" '$"{dr("D26KNAME")}" '政府政策性產業

        ST1SUGGEST.Text = "" '分署幕僚意見
        OTHFIXCONT.Text = ""
        'CONFIRMCONT.Text = ""
        ddlST1RESULT.SelectedIndex = -1
        Common.SetListItem(ddlST1RESULT, "")
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        ddlGCODE.SelectedIndex = -1
        Common.SetListItem(ddlGCODE, "")
    End Sub

    Function GET_DATA1(ByRef oConn As SqlConnection, ByRef parms As Hashtable) As DataRow
        Dim dr1 As DataRow = Nothing
        Dim v_TPLANID As String = TIMS.GetMyValue2(parms, "TPLANID") '(必填)
        Dim v_PSNO28 As String = TIMS.GetMyValue2(parms, "PSNO28") '(必填)
        Dim v_PSOID As String = TIMS.GetMyValue2(parms, "PSOID") '(單1資料必填)
        If v_TPLANID = "" OrElse v_PSNO28 = "" Then Return dr1

        Dim sql As String = ""
        sql &= " SELECT rr.ORGPLANNAME" & vbCrLf '/ 計畫別、" & vbCrLf
        sql &= " ,pp.PSNO28,pp.OCID" & vbCrLf
        sql &= " ,pp.YEARS" & vbCrLf
        sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        'sql &= " ,pp.DISTID,pp.DISTNAME" & vbCrLf '/分署別" & vbCrLf
        'sql &= " ,pp.ORGNAME" & vbCrLf '/訓練單位名稱" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf 'FIRSTSORT
        sql &= " ,pp.PSNO28" & vbCrLf '/課程申請流水號" & vbCrLf
        sql &= " ,pp.CLASSCNAME" & vbCrLf '/班級名稱" & vbCrLf
        sql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '/開訓日期" & vbCrLf
        sql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '/結訓日期" & vbCrLf
        sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '/: 訓練業別編碼" & vbCrLf
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sql &= " ,ig3.CNAME GCNAME" & vbCrLf '/訓練業別名稱" & vbCrLf
        sql &= " ,kc.CCNAME" & vbCrLf '/訓練職能" & vbCrLf

        sql &= " ,pp.TNUM" & vbCrLf '/訓練人次" & vbCrLf
        sql &= " ,pp.THOURS" & vbCrLf '/訓練時數" & vbCrLf
        sql &= " ,pp.ACTHUMCOST" & vbCrLf '/實際人時成本" & vbCrLf
        sql &= " ,pp.METSUMCOST" & vbCrLf '/實際材料費" & vbCrLf
        sql &= " ,dd.D20KNAME,dd.D25KNAME,dd.D26KNAME" & vbCrLf '政府政策性產業
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
            '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
            sql &= " ,dd.D26KNAME1,dd.D26KNAME2,dd.D26KNAME3,dd.D26KNAME4,dd.D26KNAME5,dd.D26KNAME6,dd.D26KNAME7,dd.D26KNAME8,dd.D26KNAME9"
        Else
            '5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策
            sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf '"5+2產業創新計畫"" & vbCrLf
            ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
            sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
        End If

        'sql &= " /5+2產業/新南向政策/台灣AI行動計畫/數位國家創新經濟發展方案/國家資通安全發展方案/前瞻基礎建設計畫" & vbCrLf
        sql &= " ,dbo.FN_GET_CROSSDIST(pp.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST" & vbCrLf '/是否跨區提案" & vbCrLf
        sql &= " ,pp.iCAPNUM " & vbCrLf '/iCAP標章證號" & vbCrLf

        sql &= " ,pf.PSOID" & vbCrLf
        sql &= " ,pf.ST1SUGGEST" & vbCrLf '/、初審幕僚建議/分署幕僚意見
        sql &= " ,pf.OTHFIXCONT" & vbCrLf '/其他應修正內容" & vbCrLf
        sql &= " ,pf.CONFIRMCONT" & vbCrLf '/送請委員確認內容" & vbCrLf
        sql &= " ,pf.ST1RESULT" & vbCrLf '/初審建議結論" & vbCrLf'初審建議結論 Y 通過、N 不通過、P 調整後通過
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
        '初審建議結論" & vbCrLf'初審建議結論 Y 通過、N 不通過、P 調整後通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END ST1RESULT_N" & vbCrLf
        sql &= " ,pf.RESULT" & vbCrLf '審查結果
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf
        sql &= " ,gc.PFCNAME" & vbCrLf
        '19大類主責課程 SYS_GCODEREVIE
        sql &= " ,gr1.DISTID GRDISTID" & vbCrLf

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        'PLAN_STAFFOPIN
        sql &= String.Concat(If(v_PSOID <> "", "", " LEFT"), " JOIN dbo.PLAN_STAFFOPIN pf On pf.PSNO28=pp.PSNO28")

        '審查計分等級'19大類主責課程 SYS_GCODEREVIE
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 On gr1.YEARS=pp.YEARS And gr1.APPSTAGE=pp.APPSTAGE And gr1.GCODE=pf.GCODE" & vbCrLf
        'sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 On gr1.YEARS=pp.YEARS And gr1.APPSTAGE=pp.APPSTAGE And gr1.GCODE=ig3.GCODE31" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc On gc.GCODE=pf.GCODE" & vbCrLf

        'sql &= " And CC.YEARS='2022'" & vbCrLf
        sql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)

        sql &= " AND pp.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND pp.PSNO28=@PSNO28" & vbCrLf
        If v_PSOID <> "" Then sql &= " AND pf.PSOID=@PSOID" & vbCrLf 'PLAN_STAFFOPIN
        dr1 = DbAccess.GetOneRow(sql, oConn, parms)
        Return dr1
    End Function

    Sub SHOW_DATA1(ByRef dr As DataRow)
        PanelSch1.Visible = False
        PanelEdit1.Visible = True

        If dr Is Nothing Then Return
        Hid_PSOID.Value = Convert.ToString(dr("PSOID"))
        Hid_PSNO28.Value = Convert.ToString(dr("PSNO28"))

        lbYEARS_ROC.Text = TIMS.GET_YEARS_ROC(dr("YEARS"))
        lbDistName.Text = Convert.ToString(dr("DISTNAME"))
        lbOrgName.Text = Convert.ToString(dr("ORGNAME"))
        lbPSNO28.Text = Convert.ToString(dr("PSNO28"))
        lbClassName.Text = Convert.ToString(dr("CLASSCNAME"))
        lbSFTDate.Text = String.Format("{0}~{1}", dr("STDATE"), dr("FTDATE"))

        '分署確認課程分類 / 職類課程 / 訓練業別
        lbGCODEPNAME.Text = Convert.ToString(dr("GCODEPNAME"))
        'lbGCNAME.Text = Convert.ToString(dr("GCNAME")) '訓練業別
        lbCCNAME.Text = Convert.ToString(dr("CCNAME")) '訓練職能
        lbTNum.Text = Convert.ToString(dr("TNUM"))
        lbTHours.Text = Convert.ToString(dr("THOURS"))
        lbACTHUMCOST.Text = Convert.ToString(dr("ACTHUMCOST"))
        lbMETSUMCOST.Text = Convert.ToString(dr("METSUMCOST"))
        '是否跨區提案
        Dim s_CROSSDIST As String = If(Convert.ToString(dr("CROSSDIST")) <> "", If(Val(dr("CROSSDIST")) > -1, "是", "否"), "")
        lbIsCROSSDIST.Text = s_CROSSDIST 'Convert.ToString(dr("CROSSDIST")) '是否跨區提案
        lbiCAPNUM.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(dr("iCAPNUM"))) 'iCAP標章課程
        'lbD20KNAME.Text = TIMS.NullToStr(dr("D20KNAME"), "無") '政府政策性產業
        lbD20KNAME.Text = $"{dr("D20KNAME")}" '政府政策性產業
        lbD25KNAME.Text = $"{dr("D25KNAME")}" '政府政策性產業
        lbD26KNAME.Text = $"{dr("D26KNAME")}" '政府政策性產業
        If $"{dr("D20KNAME")}{dr("D25KNAME")}{dr("D26KNAME")}" = "" Then lbD20KNAME.Text = "無"

        '初審幕僚建議
        ST1SUGGEST.Text = Convert.ToString(dr("ST1SUGGEST")) '分署幕僚意見
        OTHFIXCONT.Text = Convert.ToString(dr("OTHFIXCONT")) '其他應修正內容
        'CONFIRMCONT.Text = Convert.ToString(dr("CONFIRMCONT")) '送請委員確認內容
        Common.SetListItem(ddlST1RESULT, dr("ST1RESULT")) '初審建議結論 /通過、不通過、調整後通過
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        Hid_GCODE.Value = Convert.ToString(dr("GCODE"))
        Hid_PFGCODE.Value = Convert.ToString(dr("PFGCODE"))
        Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))
    End Sub

    Protected Sub BtnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        Call CLEAR_DATA1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False
    End Sub

    Protected Sub BtnSAVE1_Click(sender As Object, e As EventArgs) Handles btnSAVE1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return 'Exit Sub
        End If

        Call SAVE_DATA1()
    End Sub

    Sub DEL_DATA1(ByRef dr1 As DataRow)
        If Hid_PSOID.Value = "" OrElse Hid_PSNO28.Value = "" Then Return
        If (Convert.ToString(dr1("PSOID")) <> Hid_PSOID.Value) Then Return '(檢核)
        If (Convert.ToString(dr1("PSNO28")) <> Hid_PSNO28.Value) Then Return '(檢核)

        Dim d_parms As New Hashtable From {{"PSOID", Hid_PSOID.Value}, {"PSNO28", Hid_PSNO28.Value}}
        Dim d_sql As String = " DELETE PLAN_STAFFOPIN WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)

        Call sSearch1()
        Common.MessageBox(Me, TIMS.cst_DELETEOKMsg1)
    End Sub

    ''' <summary>儲存檢核</summary>
    ''' <param name="s_ERRMSG"></param>
    Sub CheckData1(ByRef s_ERRMSG As String)
        'Dim v_ddlST1RESULT As String = TIMS.GetListValue(ddlST1RESULT)
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        Dim v_ddlGCODE As String = TIMS.GetListValue(ddlGCODE)
        '分署確認課程分類
        If v_ddlGCODE = "" Then
            s_ERRMSG &= "請選擇 分署確認課程分類" & vbCrLf
        ElseIf v_ddlGCODE <> "" Then
            Dim ff3 As String = String.Format("GCODE='{0}'", v_ddlGCODE)
            If dtGCODE3.Select(ff3).Length = 0 Then s_ERRMSG &= "分署確認課程分類 代碼有誤" & vbCrLf
        End If
    End Sub

    Sub SAVE_DATA1()
        Hid_PSOID.Value = TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSNO28.Value = "" Then Return
        Dim v_ddlST1RESULT As String = TIMS.GetListValue(ddlST1RESULT)
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        Dim v_ddlGCODE As String = TIMS.GetListValue(ddlGCODE)

        'PLAN_STAFFOPIN / PSOID / PSNO28
        Dim iPSOID As Integer = If(Hid_PSOID.Value <> "", Val(Hid_PSOID.Value), 0)
        Dim iRst As Integer = 0
        If Hid_PSOID.Value = "" Then
            iPSOID = DbAccess.GetNewId(objconn, "PLAN_STAFFOPIN_PSOID_SEQ,PLAN_STAFFOPIN,PSOID")
            'i_parms.Add("CONFIRMCONT", CONFIRMCONT.Text)
            'i_parms.Add("RESULT", RESULT) 'i_parms.Add("COMMENTS", COMMENTS)
            Dim i_parms As New Hashtable From {
                {"PSOID", iPSOID},
                {"PSNO28", Hid_PSNO28.Value},
                {"ST1SUGGEST", ST1SUGGEST.Text},
                {"OTHFIXCONT", OTHFIXCONT.Text},
                {"ST1RESULT", v_ddlST1RESULT},
                {"GCODE", v_ddlGCODE},
                {"ST1ACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim i_sql As String = ""
            i_sql &= " INSERT INTO PLAN_STAFFOPIN (PSOID,PSNO28,ST1SUGGEST,OTHFIXCONT,ST1RESULT,GCODE,ST1ACCT,ST1DATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
            i_sql &= " VALUES (@PSOID,@PSNO28,@ST1SUGGEST,@OTHFIXCONT,@ST1RESULT,@GCODE,@ST1ACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf
            iRst += DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
        Else
            Dim parms As New Hashtable From {{"PSOID", Val(Hid_PSOID.Value)}, {"PSNO28", Hid_PSNO28.Value}}
            Dim s_sql As String = "SELECT PSOID FROM PLAN_STAFFOPIN WHERE PSOID=@PSOID AND PSNO28=@PSNO28"
            Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, parms)
            If TIMS.dtNODATA(dt) Then Return

            'PLAN_STAFFOPIN / PSOID / PSNO28
            'u_parms.Add("CONFIRMCONT", CONFIRMCONT.Text)
            Dim u_parms As New Hashtable From {
                {"ST1SUGGEST", ST1SUGGEST.Text},
                {"OTHFIXCONT", OTHFIXCONT.Text},
                {"ST1RESULT", v_ddlST1RESULT},
                {"GCODE", v_ddlGCODE},
                {"ST1ACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"PSOID", iPSOID},
                {"PSNO28", Hid_PSNO28.Value}
            }

            Dim u_sql As String = ""
            u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
            u_sql &= " SET ST1SUGGEST=@ST1SUGGEST" & vbCrLf
            u_sql &= " ,OTHFIXCONT=@OTHFIXCONT" & vbCrLf
            'u_sql &= " ,CONFIRMCONT=@CONFIRMCONT" & vbCrLf
            u_sql &= " ,ST1RESULT=@ST1RESULT" & vbCrLf
            u_sql &= " ,GCODE=@GCODE" & vbCrLf
            'u_sql &= " ,RESULT=@RESULT ,COMMENTS=@COMMENTS" & vbCrLf
            u_sql &= " ,ST1ACCT=@ST1ACCT ,ST1DATE=GETDATE()" & vbCrLf
            u_sql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
            u_sql &= " WHERE PSOID=@PSOID" & vbCrLf
            u_sql &= " AND PSNO28=@PSNO28" & vbCrLf
            iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        End If
        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Call sSearch1()
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

    '匯出  '表單02_跨區提案彙整表(線上填表).xlsx /提案彙整表
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Call EXPORT_1()
    End Sub

    '匯出審查意見綜整表 表單03_跨區/審查意見綜整表.xlsx/審查意見綜整表
    'Protected Sub BtnExport2_Click(sender As Object, e As EventArgs) Handles BtnExport2.Click
    '    Call EXPORT_2()
    'End Sub

    Protected Sub BtnIMPORT1_Click(sender As Object, e As EventArgs) Handles BtnIMPORT1.Click
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Call ImportXLSX_1()
        Call CCreate1()
    End Sub

    ''' <summary>匯入審查幕僚意見 匯入等級/分數</summary>
    Private Sub ImportXLSX_1()
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        ElseIf Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Const cst_Upload_Path As String = "~/CR/01/Temp/" '暫存路徑
        Call TIMS.MyCreateDir(Me, cst_Upload_Path)
        Const Cst_Filetype As String = "xlsx" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        '檢查檔案格式與大小 Start
        If File1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        ElseIf File1.PostedFile.ContentLength = 0 Then
            '檔案位置錯誤
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型 'FileOCIDValue = Split(Split(MyFileName, "-")(1), ".")(0)
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> LCase(Cst_Filetype) Then
            Common.MessageBox(Me, String.Concat("檔案類型錯誤，必須為", UCase(Cst_Filetype), "檔!"))
            Exit Sub
        End If
        '檢查檔案格式與大小 End

        Dim Errmag As String = ""
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{cst_Upload_Path}{MyFileName}")
        File1.PostedFile.SaveAs(filePath1) '上傳檔案
        '(讀取XLSX檔案轉為dt_xls)
        Dim dt_xls As DataTable = TIMS.ReadXLSX(New IO.FileInfo(filePath1), Errmag)
        '刪除檔案 IO.File.Delete(filePath1)
        TIMS.MyFileDelete(filePath1)

        If Errmag <> "" Then
            Errmag &= "資料有誤，故無法匯入，請修正Excel檔案!"
            Common.MessageBox(Me, Errmag)
            Exit Sub
        End If
        If dt_xls Is Nothing Then '有資料
            Common.MessageBox(Me, "資料為空，故無法匯入，請修正Excel檔案!")
            Exit Sub
        End If
        If dt_xls.Rows.Count = 0 Then '有資料
            Common.MessageBox(Me, "查無資料，故無法匯入，請修正Excel檔案!")
            Exit Sub
        End If

        '取出資料庫的所有欄位    Start
        'Dim sql As String = ""
        'Dim da As SqlDataAdapter = Nothing
        '建立錯誤資料格式Table Start
        'Dim Reason As String  '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("PSNO28"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        Dim sHtb As New Hashtable
        Dim iRowIndex As Integer = 0 '讀取行累計數
        Dim Reason As String = "" '做一次驗証的即可
        If Reason = "" Then
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                Reason = ""
                Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                Try
                    Reason = SAVE_PLAN_STAFFOPIN(colArray, sHtb)  '驗証(單筆) 並 儲存
                Catch ex As Exception
                    Reason = ex.Message
                    'TIMS.LOG.Error(ex.Message, ex)
                    Call TIMS.WriteTraceLog(ex.Message, ex) 'Throw ex
                End Try
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = String.Concat("第", CStr(iRowIndex + 2), "列")
                    drWrong("PSNO28") = s_COL_PSNO28 '統一編號
                    drWrong("Reason") = If(Reason <> "", Reason, "(錯誤)") 'Reason
                End If

                iRowIndex += 1 '讀取行累計數
                If g_IMP_ERR1 Then Exit For
            Next
            'Loop
        End If

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""
        '開始判別欄位存入 End
        If dtWrong.Rows.Count = 0 Then
            explain = ""
            explain = String.Concat(explain, "匯入資料共", iRowIndex, "筆", vbCrLf)
            explain = String.Concat(explain, "成功：", (iRowIndex - dtWrong.Rows.Count), "筆", vbCrLf)
            explain = String.Concat(explain, "失敗：", dtWrong.Rows.Count, "筆", vbCrLf)
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Reason = String.Concat("錯誤訊息如下:", vbCrLf, Reason)
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            explain2 = String.Concat(explain2, "匯入資料共", iRowIndex, "筆\n")
            explain2 = String.Concat(explain2, "成功：", (iRowIndex - dtWrong.Rows.Count), "筆\n")
            explain2 = String.Concat(explain2, "失敗：", dtWrong.Rows.Count, "筆\n")
            Session("MyWrongTable") = dtWrong
            Const CST_WRONG_ASPX_1 As String = "CR_01_003_Wrong.aspx"
            Dim s_FMT1 As String = String.Format("window.open('{0}','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');", CST_WRONG_ASPX_1)
            Dim s_DYW2CRES As String = String.Concat(explain2, "是否要檢視原因?")
            Dim s_JS1 As String = String.Concat("<script>if(confirm('", s_DYW2CRES, "')){", s_FMT1, "}</script>")
            Page.RegisterStartupScript("", s_JS1)
        End If

    End Sub

    Function CheckImp1(ByRef ErrMsg1 As String) As Boolean
        Dim rst As Boolean = False '正常:true '異常:false 
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        If v_ddlAPPSTAGE_SCH = "" Then
            ErrMsg1 &= "申請階段，未選擇無法匯入，請先選擇申請階段!" & vbCrLf
            Return rst
        End If
        'Common.MessageBox(Me, "分署未選擇，無法匯入，請先選擇分署!")
        'Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        'If v_ddlSCORING = "" Then
        '    ErrMsg1 &= "審查計分區間未選擇，無法匯入，請先選擇審查計分區間!" & vbCrLf
        '    Return rst
        'End If
        rst = True '正常:true '異常:false 
        Return rst
    End Function

    ''' <summary>匯入驗証</summary>
    ''' <param name="colArray">比對資料</param>
    ''' <param name="Htb">輸入查詢</param>
    ''' <param name="o_parms">取得有效值</param>
    ''' <returns></returns>
    Function CheckImportData(ByRef colArray As Array, ByRef Htb As Hashtable, ByRef o_parms As Hashtable) As String
        Dim Reason As String = ""
        s_COL_PSNO28 = ""

        If colArray.Length < cst_col_iMaxLen Then
            g_IMP_ERR1 = True
            Reason &= "欄位對應有誤<BR>,請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        'Dim s_col_PSNO28 As String = "" '課程申請流水號
        If colArray.Length > cst_col_PSNO28 Then s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col_PSNO28)) '課程申請流水號

        Dim s_COL_PCS As String = "" '(計畫)課程流水號
        Dim s_COL_ST1SUGGEST As String = "" '分署幕僚意見
        Dim s_COL_OTHFIXCONT As String = "" '其他應修正內容
        'Dim s_COL_CONFIRMCONT As String = "" '送請委員確認內容
        Dim s_COL_ST1RESULT As String = "" '初審建議結論
        Dim s_COL_ST1RESULT_YNP As String = "" '初審建議結論(YNP)
        Dim s_COL_GCODE As String = "" '分署確認課程分類

        s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col_PSNO28)) '課程申請流水號
        s_COL_ST1SUGGEST = TIMS.NullToStr(colArray(cst_col_ST1SUGGEST)) '其他應修正內容
        s_COL_OTHFIXCONT = TIMS.NullToStr(colArray(cst_col_OTHFIXCONT)) '其他應修正內容
        's_COL_CONFIRMCONT = TIMS.NullToStr(colArray(cst_col_CONFIRMCONT)) '送請委員確認內容
        s_COL_ST1RESULT = TIMS.ClearSQM(colArray(cst_col_ST1RESULT)) '初審建議結論
        '分署確認課程分類
        s_COL_GCODE = TIMS.ClearSQM(colArray(cst_col_GCODE))
        If s_COL_GCODE <> "" Then s_COL_GCODE = TIMS.AddZero(s_COL_GCODE, 2)

        Dim flag_TXT_NG As Boolean = (s_COL_ST1SUGGEST = "" AndAlso s_COL_OTHFIXCONT = "") 'AndAlso s_COL_CONFIRMCONT = ""

        '先確認資料不為空
        If s_COL_PSNO28 = "" Then Reason &= "課程申請流水號 不可為空<br>"
        'If s_col_OTHFIXCONT = "" Then Reason &= "其他應修正內容不可為空<br>"
        'If s_col_CONFIRMCONT = "" Then Reason &= "送請委員確認內容不可為空<br>"
        If flag_TXT_NG AndAlso s_COL_ST1RESULT = "" Then Reason &= "分署幕僚意見/其他應修正內容/初審建議結論 不可皆為空<br>"
        If Reason <> "" Then Return Reason

        s_COL_PCS = TIMS.Get_PCSforPSNO28(sm, s_COL_PSNO28, objconn)
        If s_COL_PCS = "" Then Reason &= String.Format("課程申請流水號 有誤，查無計畫資料({0})或確認登入年度<br>", s_COL_PSNO28)
        If Reason <> "" Then Return Reason

        Dim drPP As DataRow = TIMS.GetPCSDate(s_COL_PCS, objconn)
        If drPP Is Nothing Then Reason &= String.Format("課程申請流水號 有誤，查無計畫資料({0}-{1})或確認登入年度<br>", s_COL_PSNO28, s_COL_PCS)
        If Reason <> "" Then Return Reason

        'Y/N/P'用文字方式輸入
        s_COL_ST1RESULT_YNP = If(s_COL_ST1RESULT = "通過", "Y", If(s_COL_ST1RESULT = "不通過", "N", If(s_COL_ST1RESULT = "調整後通過", "P", "")))
        s_COL_ST1RESULT_YNP = If(s_COL_ST1RESULT_YNP <> "", s_COL_ST1RESULT_YNP, If(s_COL_ST1RESULT = "1", "Y", If(s_COL_ST1RESULT = "3", "N", If(s_COL_ST1RESULT = "2", "P", ""))))
        If s_COL_ST1RESULT_YNP = "" AndAlso s_COL_ST1RESULT <> "" Then
            Select Case s_COL_ST1RESULT '用代碼方式輸入
                Case "Y", "P", "N"
                    s_COL_ST1RESULT_YNP = s_COL_ST1RESULT
            End Select
        End If
        '/送請委員確認內容
        If flag_TXT_NG AndAlso s_COL_ST1RESULT_YNP = "" Then Reason &= "分署幕僚意見/其他應修正內容/初審建議結論 資料皆為空<br>"
        'If flag_TXT_NG AndAlso s_col_ST1RESULT_YNP = "" Then Reason &= String.Format("初審建議結論 有誤，(通過/調整後通過/不通過):{0}<br>", s_col_ST1RESULT)

        '分署確認課程分類
        If s_COL_GCODE <> "" Then
            Dim ff3 As String = String.Format("GCODE='{0}'", s_COL_GCODE)
            If dtGCODE3.Select(ff3).Length = 0 Then Reason &= String.Format("分署確認課程分類 代碼輸入有誤({0})<br>", s_COL_GCODE)
        ElseIf drPP IsNot Nothing Then
            s_COL_GCODE = Convert.ToString(drPP("GCODE"))
        End If
        'Dim s_COL_PPGCODE As String = If(drPP IsNot Nothing, Convert.ToString(drPP("GCODE")), "")

        'Select Case s_col_ST1RESULT_YNP
        '    Case "Y"
        '    Case Else
        '        If s_col_OTHFIXCONT = "" Then Reason &= "其他應修正內容不可為空<br>"
        '        If s_col_CONFIRMCONT = "" Then Reason &= "送請委員確認內容不可為空<br>"
        'End Select

        If Not ChkBxCover1.Checked Then
            Dim flag_EXISTS_1 As Boolean = CHK_PLAN_STAFFOPIN_EXISTS(s_COL_PSNO28)
            If flag_EXISTS_1 Then Reason &= String.Format("課程申請流水號 已有資料，不再匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If
        If Reason <> "" Then Return Reason

        If o_parms Is Nothing Then o_parms = New Hashtable
        o_parms.Add("PSNO28", s_COL_PSNO28)

        o_parms.Add("ST1SUGGEST", s_COL_ST1SUGGEST) '分署幕僚意見
        o_parms.Add("OTHFIXCONT", s_COL_OTHFIXCONT)
        'o_parms.Add("CONFIRMCONT", s_COL_CONFIRMCONT)
        o_parms.Add("ST1RESULT", s_COL_ST1RESULT_YNP)
        o_parms.Add("GCODE", s_COL_GCODE)
        'o_parms.Add("PPGCODE", s_COL_PPGCODE)
        Return Reason
    End Function

    ''' <summary>檢核資料是否存在</summary>
    ''' <param name="s_PSNO28"></param>
    ''' <returns></returns>
    Function CHK_PLAN_STAFFOPIN_EXISTS(ByVal s_PSNO28 As String) As Boolean
        Dim rst As Boolean = False
        If s_PSNO28 = "" Then Return rst

        Dim dt1 As New DataTable
        Dim s_sql As String = " SELECT PSOID FROM PLAN_STAFFOPIN WHERE PSNO28=@PSNO28" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            dt1.Load(.ExecuteReader)
        End With
        rst = (dt1.Rows.Count > 0)
        Return rst
    End Function

    Function SAVE_PLAN_STAFFOPIN(ByRef colArray As Array, ByRef Htb As Hashtable) As String
        Dim o_parms As New Hashtable
        Dim rst As String = CheckImportData(colArray, Htb, o_parms)
        If rst <> "" Then Return rst

        Dim s_PSNO28 As String = TIMS.GetMyValue2(o_parms, "PSNO28")
        Dim vST1SUGGEST As String = TIMS.GetMyValue2(o_parms, "ST1SUGGEST")
        Dim vOTHFIXCONT As String = TIMS.GetMyValue2(o_parms, "OTHFIXCONT")
        'Dim vCONFIRMCONT As String = TIMS.GetMyValue2(o_parms, "CONFIRMCONT")
        Dim vST1RESULT As String = TIMS.GetMyValue2(o_parms, "ST1RESULT")
        Dim vGCODE As String = TIMS.GetMyValue2(o_parms, "GCODE")
        'Dim vPPGCODE As String = TIMS.GetMyValue2(o_parms, "PPGCODE")

        '手動勾選「確定覆蓋」匯入。
        Dim fg_CanCover1 As Boolean = ChkBxCover1.Checked
        If Not fg_CanCover1 Then
            Dim flag_EXISTS_1 As Boolean = CHK_PLAN_STAFFOPIN_EXISTS(s_PSNO28)
            If flag_EXISTS_1 Then Return rst
        End If

        Dim dt1 As New DataTable
        Dim s_sql As String = " SELECT PSOID,ST1SUGGEST,OTHFIXCONT,ST1RESULT,GCODE FROM PLAN_STAFFOPIN WHERE PSNO28=@PSNO28" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            dt1.Load(.ExecuteReader)
        End With
        If dt1.Rows.Count = 0 Then
            Dim iPSOID As Integer = DbAccess.GetNewId(objconn, "PLAN_STAFFOPIN_PSOID_SEQ,PLAN_STAFFOPIN,PSOID")
            'i_parms.Add("CONFIRMCONT", If(vCONFIRMCONT <> "", vCONFIRMCONT, Convert.DBNull)) 'i_parms.Add("RESULTACCT", sm.UserInfo.UserID)
            Dim i_parms As New Hashtable From {
                {"PSOID", iPSOID}, {"PSNO28", s_PSNO28},
                {"ST1SUGGEST", If(vST1SUGGEST <> "", vST1SUGGEST, Convert.DBNull)},
                {"OTHFIXCONT", If(vOTHFIXCONT <> "", vOTHFIXCONT, Convert.DBNull)},
                {"ST1RESULT", If(vST1RESULT <> "", vST1RESULT, Convert.DBNull)},
                {"GCODE", If(vGCODE <> "", vGCODE, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID}, {"ST1ACCT", sm.UserInfo.UserID}
            }
            Dim i_sql As String = ""
            i_sql &= " INSERT INTO PLAN_STAFFOPIN ( PSOID,PSNO28,ST1SUGGEST,OTHFIXCONT,ST1RESULT,GCODE" & vbCrLf
            i_sql &= " ,MODIFYACCT ,MODIFYDATE, ST1ACCT,ST1DATE)" & vbCrLf
            i_sql &= " VALUES (@PSOID,@PSNO28,@ST1SUGGEST,@OTHFIXCONT,@ST1RESULT,@GCODE" & vbCrLf
            i_sql &= " ,@MODIFYACCT,GETDATE(), @ST1ACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(i_sql, objconn, i_parms)
        ElseIf ChkBxCover1.Checked Then
            '手動勾選「確定覆蓋」匯入。
            Dim dro1 As DataRow = dt1.Rows(0)
            Dim iPSOID As Integer = dro1("PSOID")
            'ST1SUGGEST,OTHFIXCONT,ST1RESULT,GCODE
            Dim oST1SUGGEST As String = If(Convert.ToString(dro1("ST1SUGGEST")) <> "", dro1("ST1SUGGEST"), "")
            Dim oOTHFIXCONT As String = If(Convert.ToString(dro1("OTHFIXCONT")) <> "", dro1("OTHFIXCONT"), "")
            Dim oST1RESULT As String = If(Convert.ToString(dro1("ST1RESULT")) <> "", dro1("ST1RESULT"), "")
            Dim oGCODE As String = If(Convert.ToString(dro1("GCODE")) <> "", dro1("GCODE"), "")
            '若輸入值為空，保留原本資料
            If vST1SUGGEST = "" AndAlso oST1SUGGEST <> "" Then vST1SUGGEST = oST1SUGGEST
            If vOTHFIXCONT = "" AndAlso oOTHFIXCONT <> "" Then vOTHFIXCONT = oOTHFIXCONT
            If vST1RESULT = "" AndAlso oST1RESULT <> "" Then vST1RESULT = oST1RESULT
            If vGCODE = "" AndAlso oGCODE <> "" Then vGCODE = oGCODE
            Dim u_parms As New Hashtable From {
                {"ST1SUGGEST", If(vST1SUGGEST <> "", vST1SUGGEST, Convert.DBNull)},
                {"OTHFIXCONT", If(vOTHFIXCONT <> "", vOTHFIXCONT, Convert.DBNull)},
                {"ST1RESULT", If(vST1RESULT <> "", vST1RESULT, Convert.DBNull)},
                {"GCODE", If(vGCODE <> "", vGCODE, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID}, {"ST1ACCT", sm.UserInfo.UserID},
                {"PSOID", iPSOID}, {"PSNO28", s_PSNO28}
            }
            Dim u_sql As String = ""
            u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
            u_sql &= " SET ST1SUGGEST=@ST1SUGGEST,OTHFIXCONT=@OTHFIXCONT,ST1RESULT=@ST1RESULT,GCODE=@GCODE" & vbCrLf
            u_sql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE(),ST1ACCT=@ST1ACCT,ST1DATE=GETDATE()" & vbCrLf
            u_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        End If

        Return rst
    End Function

    Private Sub DelBatchPLAN_STAFFOPIN(ByRef dt As DataTable, ByRef DGobj As DataGrid)
        Const cst_errmsg16 As String = "傳入表格資訊有誤，刪除失敗!"
        Const cst_errmsg10 As String = "審查幕僚意見明細，儲存資料有誤!"
        If DGobj Is Nothing Then
            sm.LastErrorMessage = cst_errmsg16
            Return ' Exit Sub
        End If
        Dim iCNT As Integer = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim ChkBxDG1_DEL As CheckBox = eItem.FindControl("ChkBxDG1_DEL")
            Dim Hid_DataKey As HiddenField = eItem.FindControl("Hid_DataKey") 'PSOID
            Dim Hid_DataKey2 As HiddenField = eItem.FindControl("Hid_DataKey2") 'PSNO28

            Dim fg_continue As Boolean = (ChkBxDG1_DEL Is Nothing OrElse Hid_DataKey Is Nothing OrElse Hid_DataKey2 Is Nothing)
            If fg_continue Then Continue For

            Hid_PSOID.Value = If(ChkBxDG1_DEL.Checked, TIMS.DecryptAes(Hid_DataKey.Value), "")
            Hid_PSNO28.Value = If(ChkBxDG1_DEL.Checked, TIMS.DecryptAes(Hid_DataKey2.Value), "")
            If ChkBxDG1_DEL.Checked AndAlso Hid_PSOID.Value <> "" AndAlso Hid_PSNO28.Value <> "" Then
                Dim sfilter As String = String.Concat(Cst_PLAN_STAFFOPIN_pkName, "=", Hid_PSOID.Value)
                '搜尋刪除資料刪除
                If dt.Select(sfilter).Length <> 0 Then
                    'iCNT += 1
                    'For Each dr As DataRow In dt.Select(sfilter)
                    '    If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    'Next
                    Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"PSOID", Hid_PSOID.Value}, {"PSNO28", Hid_PSNO28.Value}}
                    Dim dr1 As DataRow = GET_DATA1(objconn, parms)
                    If dr1 IsNot Nothing Then
                        Dim d_parms As New Hashtable From {{"PSOID", Hid_PSOID.Value}, {"PSNO28", Hid_PSNO28.Value}}
                        Dim d_sql As String = " DELETE PLAN_STAFFOPIN WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
                        iCNT += DbAccess.ExecuteNonQuery(d_sql, objconn, d_parms)
                    End If

                End If
            End If
        Next
        If iCNT = 0 Then sm.LastErrorMessage = cst_errmsg10
    End Sub

    Protected Sub BtnDELChkBxDG1_Click(sender As Object, e As EventArgs) Handles BtnDELChkBxDG1.Click
        Dim dt As DataTable = SEARCH_DATA1_dt(cst_iType_查詢1)

        Call DelBatchPLAN_STAFFOPIN(dt, DataGrid1)

        dt = SEARCH_DATA1_dt(cst_iType_查詢1)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            msg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        msg1.Text = ""
        tbDataGrid1.Visible = True
        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>寄送iCAP課程比對結果</summary>
    ''' <param name="dt"></param>
    ''' <param name="DGobj"></param>
    Private Sub SENDiCAP_RESULTS(ByRef dt As DataTable, ByRef DGobj As DataGrid)
        Const cst_errmsg1a As String = "傳入表格資訊有誤，寄送iCAP課程比對結果失敗!"
        Const cst_errmsg1b As String = "iCAP比對沒有勾選，寄送iCAP課程比對結果失敗!!"
        Const cst_errmsg1c As String = "iCAP比對資料有誤，寄送iCAP課程比對結果失敗!"
        If DGobj Is Nothing Then
            sm.LastErrorMessage = cst_errmsg1a
            Return ' Exit Sub
        End If
        Dim iCNT As Integer = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim ChkBxDG1_ICAP As CheckBox = eItem.FindControl("ChkBxDG1_ICAP")
            If ChkBxDG1_ICAP IsNot Nothing AndAlso ChkBxDG1_ICAP.Checked Then iCNT += 1
        Next
        If iCNT = 0 Then
            sm.LastErrorMessage = cst_errmsg1b
            Return ' Exit Sub
        End If

        iCNT = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim ChkBxDG1_ICAP As CheckBox = eItem.FindControl("ChkBxDG1_ICAP")
            'Dim Hid_DataKey As HiddenField = eItem.FindControl("Hid_DataKey") 'PSOID
            Dim Hid_DataKey2 As HiddenField = eItem.FindControl("Hid_DataKey2") 'PSNO28

            Dim fg_continue As Boolean = (ChkBxDG1_ICAP Is Nothing OrElse Hid_DataKey2 Is Nothing)
            If fg_continue Then Continue For

            'Hid_PSOID.Value = If(ChkBxDG1_ICAP.Checked, TIMS.DecryptAes(Hid_DataKey.Value), "")
            Hid_PSNO28.Value = If(ChkBxDG1_ICAP.Checked, TIMS.DecryptAes(Hid_DataKey2.Value), "")
            If ChkBxDG1_ICAP.Checked AndAlso Hid_PSNO28.Value <> "" Then
                Dim sfilter As String = String.Format("{0}='{1}'", Cst_PLAN_STAFFOPIN_PSNO28, Hid_PSNO28.Value)
                '搜尋資料
                If dt.Select(sfilter).Length <> 0 Then
                    Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"PSNO28", Hid_PSNO28.Value}}
                    Dim dr1 As DataRow = GET_DATA1(objconn, parms)
                    If dr1 IsNot Nothing Then
                        Dim vPSNO28 As String = Convert.ToString(dr1("PSNO28"))
                        Dim viCAPNUM As String = TIMS.ChangeIDNO(TIMS.ClearSQM(dr1("iCAPNUM")))
                        Dim vDISTID As String = Convert.ToString(dr1("DISTID"))
                        Dim vEMAILSENDiCAP As String = TIMS.GET_EMAIL_SENDICAP(vDISTID, objconn)
                        If vEMAILSENDiCAP = "" Then
                            sm.LastErrorMessage = "查無分署信箱EMAIL通知對象為空，請設定要接收iCAP課程比對通知的對象!" 'Common.MessageBox(Me, vERRMSG1)
                            Return
                        End If
                        Dim vOCID As String = Convert.ToString(dr1("OCID"))
                        '課程申請流水號 PSNO28、iCAP標章證號 iCAPNUM、該分署信箱、OCID
                        Dim pms_r1 As New Hashtable From {{"PSNO28", vPSNO28}, {"ICAPNUM", viCAPNUM}, {"EMAILSENDICAP", vEMAILSENDiCAP}, {"OCID", vOCID}}
                        Dim fg_send1 As Boolean = TIMS.SENDICAP_RESULTS_DIFF1(pms_r1)
                        Dim vERRMSG1 As String = TIMS.GetMyValue2(pms_r1, "ERRMSG1")
                        If vERRMSG1 <> "" Then
                            sm.LastErrorMessage = vERRMSG1 'Common.MessageBox(Me, vERRMSG1)
                            Return
                        ElseIf Not fg_send1 Then
                            sm.LastErrorMessage = cst_errmsg1c 'Common.MessageBox(Me, vERRMSG1)
                            Return
                        End If
                        iCNT += 1
                    End If
                End If
            End If
        Next
        If iCNT > 0 Then
            sm.LastResultMessage = "已寄出，分署接收iCAP課程比對通知對象。"
            Return
        ElseIf iCNT = 0 Then
            sm.LastErrorMessage = String.Concat(cst_errmsg1c, " : 0") 'Common.MessageBox(Me, vERRMSG1)
            Return
        End If
    End Sub

    ''' <summary>勾選寄送iCAP課程比對結果</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnCBL_SENDICAP_RESULTS_Click(sender As Object, e As EventArgs) Handles BtnCBL_SENDICAP_RESULTS.Click
        Dim dt As DataTable = SEARCH_DATA1_dt(cst_iType_查詢1)

        '寄送iCAP課程比對結果
        Call SENDiCAP_RESULTS(dt, DataGrid1)

        dt = SEARCH_DATA1_dt(cst_iType_查詢1)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            msg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        msg1.Text = ""
        tbDataGrid1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
