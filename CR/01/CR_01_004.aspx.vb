Public Class CR_01_004
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'PLAN_STAFFOPIN / PSOID / PSNO28
    Dim s_COL_PSNO28 As String = "" '課程申請流水號
    Dim g_IMP_ERR1 As Boolean = False

    Const cst_APPSTAGE_政策性產業_3 As String = "3"

    Const cst_iType_查詢1 As Integer = 1
    'Const cst_iType_匯出1 As Integer = 11
    Const cst_iType_匯出12 As Integer = 12
    Const cst_iType_匯出13 As Integer = 13

    Const cst_col1_PSNO28 As Integer = 0 '課程申請流水號
    Const cst_col1_CONFIRMCONT As Integer = 1 '送請委員確認內容
    Const cst_col1_iMaxLen As Integer = 2

    Const cst_col2_PSNO28 As Integer = 0 '課程申請流水號
    Const cst_col2_COMMENTS As Integer = 1 '委員審查意見與建議 
    Const cst_col2_RESULT As Integer = 2 '審核結果
    Const cst_col2_iMaxLen As Integer = 3

    'directcover
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
                '署(局) 或 分署(中心)
                'If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
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
        PanelEdit1.Visible = False

        Const cst_Title_msg1 As String = "當有勾選，於資料匯入時，系統不檢核允許直接覆蓋匯入資料。"
        TIMS.Tooltip(ChkBxCover1, cst_Title_msg1, True)
        TIMS.Tooltip(ChkBxCover2, cst_Title_msg1, True)

        msg1.Text = ""
        tbDataGrid1.Visible = False

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
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ 

        '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位 '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位

        '初審建議結論 --Y 通過、N 不通過、P 調整後通過
        ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        ddlGCODE = TIMS.Get_GOVCODE3(dtGCODE3, ddlGCODE, False)
        '一階審查結果 RESULT 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
        ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)
    End Sub

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Call sSearch1()
    End Sub

    Function GET_ORG_SQL1() As String
        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @APPSTAGE NUMERIC(10,0)=2;
        Dim sql As String = ""
        sql &= " SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
        sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
        sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
        Return sql
    End Function

    ''' <summary>檢核查詢輸入值</summary>
    ''' <returns></returns>
    Function CHECK_SCH_DATA1() As Boolean
        'Dim dt As DataTable = Nothing

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過
        Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        '審查結果'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        Dim v_RBL_RESULT_SCH As String = TIMS.GetListValue(RBL_RESULT_SCH)

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        If s_DISTID = "" AndAlso v_RBL_CrossDist_SCH <> "C" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return False 'dt
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

        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return False ' dt
        End If

        Return True
    End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt(ByRef iType As Integer) As DataTable
        Dim dt As DataTable = Nothing

        '檢核查詢輸入值
        Dim flag_OK As Boolean = CHECK_SCH_DATA1()
        If Not flag_OK Then Return dt

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過
        Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        '審查結果'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        Dim v_RBL_RESULT_SCH As String = TIMS.GetListValue(RBL_RESULT_SCH)

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

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

        Dim sql_WORG1 As String = String.Format("WITH WORG1 AS ({0})", GET_ORG_SQL1())

        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @TPLANID VARCHAR(3)='28';DECLARE @APPSTAGE NUMERIC(10,0)=2;
        Dim parms As New Hashtable From {
            {"YEARS", sm.UserInfo.Years},
            {"TPLANID", sm.UserInfo.TPlanID},
            {"APPSTAGE", v_APPSTAGE_SCH}
        }
        Dim sql As String = ""
        sql = sql_WORG1
        sql &= " SELECT pp.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(pp.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,pp.APPSTAGE" & vbCrLf
        sql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" & vbCrLf
        sql &= " ,pp.PLANNAME" & vbCrLf
        sql &= " ,pp.PSNO28 ,pp.RID" & vbCrLf
        sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        sql &= " ,pp.OCID" & vbCrLf
        sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf
        sql &= " ,pp.CLASSCNAME" & vbCrLf
        sql &= " ,FORMAT(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,pp.GCID3" & vbCrLf
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '訓練業別編碼" & vbCrLf
        sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sql &= " ,ig3.CNAME GCNAME" & vbCrLf '/訓練業別名稱" & vbCrLf
        'sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        'sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sql &= " ,pf.PSOID" & vbCrLf '審查幕僚意見 --SEQNO
        sql &= " ,pf.ST1SUGGEST" & vbCrLf '初審幕僚建議/分署幕僚建議
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
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END RESULT_C" & vbCrLf

        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END RESULT_N" & vbCrLf
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf
        '分署確認課程分類
        sql &= " ,gc.PFCNAME" & vbCrLf
        sql &= " ,gr1.DISTID GRDISTID" & vbCrLf '19大類主責課程 SYS_GCODEREVIE

        '匯出1/'匯出2
        Dim flag_EXP_12 As Boolean = If(iType = cst_iType_匯出12, True, False)
        Dim flag_EXP_13 As Boolean = If(iType = cst_iType_匯出13, True, False)
        If flag_EXP_12 Then
            'Dim sql As String = ""
            sql = sql_WORG1
            sql &= " SELECT rr.ORGPLANNAME" & vbCrLf '計畫別、" & vbCrLf
            sql &= " ,pp.ORGID" & vbCrLf
            sql &= " ,pp.OCID" & vbCrLf
            sql &= " ,pp.YEARS" & vbCrLf
            sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
            'sql &= " ,pp.DISTNAME" & vbCrLf '分署別" & vbCrLf
            'sql &= " ,pp.ORGNAME" & vbCrLf '訓練單位名稱" & vbCrLf
            sql &= " ,pp.FIRSTSORT" & vbCrLf ',cc.FIRSTSORT
            sql &= " ,pp.PSNO28" & vbCrLf '課程申請流水號" & vbCrLf
            sql &= " ,pp.CLASSCNAME" & vbCrLf '班級名稱" & vbCrLf
            sql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期" & vbCrLf
            sql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期" & vbCrLf

            sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '訓練業別編碼" & vbCrLf
            '分署確認課程分類 / 職類課程 / 訓練業別
            sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
            sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
            sql &= " ,ig3.CNAME GCNAME" & vbCrLf '/訓練業別名稱" & vbCrLf
            'sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf ': 訓練業別編碼" & vbCrLf
            'sql &= " ,ig3.CNAME GCNAME" & vbCrLf '訓練業別名稱" & vbCrLf 
            sql &= " ,kc.CCNAME" & vbCrLf '訓練職能" & vbCrLf
            sql &= " ,tt.JOBNAME" & vbCrLf '課程分類 JOBNAME/PKNAME12
            sql &= " ,ig3.PNAME PKNAME12" & vbCrLf '課程分類 JOBNAME/PKNAME12

            sql &= " ,pp.THOURS" & vbCrLf '訓練時數" & vbCrLf
            sql &= " ,pp.TNUM" & vbCrLf '訓練人次" & vbCrLf
            sql &= " ,pp.ACTHUMCOST" & vbCrLf '實際人時成本" & vbCrLf
            sql &= " ,pp.METSUMCOST" & vbCrLf '實際材料費" & vbCrLf
            '/實際材料費比率,pp.METCOSTPER
            sql &= " ,CASE WHEN pp.METCOSTPER>=0 THEN concat(convert(float, pp.METCOSTPER),'%') END METCOSTPER" & vbCrLf '/實際材料費比率" & vbCrLf
            'sql &= " ,dbo.FN_GET_KID20NAME(pp.PLANID,pp.COMIDNO,pp.SEQNO) D20KNAME" & vbCrLf '政府政策性產業
            sql &= " ,dd.D20KNAME,dd.D25KNAME,dd.D26KNAME" & vbCrLf '政府政策性產業
            '5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策
            sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf '"5+2產業創新計畫"" & vbCrLf
            ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
            sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
            'sql &= " 5+2產業/新南向政策/台灣AI行動計畫/數位國家創新經濟發展方案/國家資通安全發展方案/前瞻基礎建設計畫" & vbCrLf
            'sql &= " ,dbo.FN_GET_CROSSDIST(cc.YEARS,cc.COMIDNO,cc.APPSTAGE) CROSSDIST" & vbCrLf '/是否跨區提案" & vbCrLf
            'I_CROSSDIST 是否跨區提案
            sql &= " ,CASE wo.I_CROSSDIST WHEN -1 THEN '否' ELSE '是' END CROSSDIST" & vbCrLf '/是否跨區提案" & vbCrLf
            sql &= " ,pp.ICAPNUM" & vbCrLf '/iCAP標章證號" & vbCrLf

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
            '分署確認課程分類
            sql &= " ,gc.PFCNAME" & vbCrLf

            sql &= " ,NULL SAMEOCFIXREC" & vbCrLf '同單位同類課程建議修正意見 SAME UNIT	SIMILAR COURSES	SUGGEST	FIX	OPINION	
            sql &= " ,NULL SAMEOCNT1" & vbCrLf '同單位同類課程(班數)
            sql &= " ,NULL REMARK1" & vbCrLf '備註
            sql &= " ,gr1.DISTID GRDISTID" & vbCrLf '19大類主責課程 SYS_GCODEREVIE

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
            'sql &= " ,o2.RLEVEL_2" & vbCrLf '--審查計分等級" & vbCrLf
            'sql &= " ,dbo.FN_GET_CLASSQUOTA(pp.ORGKIND2,pp.YEARS,pp.APPSTAGE,pp.OSID2) CLASSQUOTA" & vbCrLf '/等級額度核配上限" & vbCrLf
            'sql &= " ,dbo.FN_SCORING2_UPLIMIT(pp.COMIDNO,pp.TPLANID,pp.YEARS,pp.APPSTAGE,pp.ORGKIND2) UPLIMIT" & vbCrLf '/等級額度核配上限" & vbCrLf
            'sql &= " /,dbo.FN_SCORING2_GRADE(pp.COMIDNO,pp.TPLANID,pp.YEARS,pp.APPSTAGE) GRADE" & vbCrLf
            sql &= " ,pp.ORGTYPENAME" & vbCrLf '/單位屬性。" & vbCrLf

        ElseIf flag_EXP_13 Then
            'Dim sSql As String = ""
            sql = sql_WORG1
            sql &= " SELECT pp.PLANID,pp.SEQNO" & vbCrLf
            sql &= " ,pp.ORGPLANNAME2" '計畫別" & vbCrLf
            sql &= " ,pp.ORGNAME" '訓練單位名稱" & vbCrLf
            sql &= " ,pp.APPSTAGE,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" '申請階段" & vbCrLf
            sql &= " ,pp.COMIDNO" '統一編號" & vbCrLf
            sql &= " ,pp.ORGTYPENAME" '單位屬性" & vbCrLf
            sql &= " ,pp.DISTNAME" '分署別" & vbCrLf
            sql &= " ,pp.CLASSCNAME" '課程名稱" & vbCrLf
            sql &= " ,pp.FIRSTSORT" '提案意願順序" & vbCrLf
            sql &= " ,pp.PSNO28" '課程申請流水號" & vbCrLf
            sql &= " ,ig3.GCODE31 GCODE" '課程分類編碼" & vbCrLf
            sql &= " ,ig3.PNAME GCODEPNAME" '課程分類" & vbCrLf
            sql &= " ,pp.THOURS" '訓練時數" & vbCrLf
            sql &= " ,pp.TNUM" '訓練人次" & vbCrLf
            sql &= " ,pp.DEFGOVCOST1" '每人訓練費用(元)" & vbCrLf
            sql &= " ,pp.TOTALCOST" '訓練單位可向學員收取之訓練費用(元)" & vbCrLf
            sql &= " ,pp.DEFGOVCOST" '總補助費(元)(以訓練費用之80%估算)" & vbCrLf
            sql &= " ,dbo.FN_GET_PLAN_TRAINDESC(pp.PLANID,pp.COMIDNO,pp.SEQNO,2) PROTECHHOURS" '術科時數" & vbCrLf
            sql &= " ,pp.FIXSUMCOST" '固定費用總計" & vbCrLf
            sql &= " ,pp.ACTHUMCOST" '實際人時成本" & vbCrLf
            'sql &= " "‘METDET,METSUMCOST,METCOSTPER,ALLSUMCOST"‘材料明細" & vbCrLf
            sql &= " ,dbo.FN_GET_PLANCNAME2(pp.PLANID,pp.COMIDNO,pp.SEQNO) METDET" '材料明細" & vbCrLf
            sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'P') METDETP" & vbCrLf
            sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'C') METDETC" & vbCrLf
            sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'S') METDETS" & vbCrLf
            sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'O') METDETO" & vbCrLf
            sql &= " ,pp.METSUMCOST" '材料費總計" & vbCrLf
            sql &= " ,pp.METCOSTPER" '--材料費占比" & vbCrLf
            sql &= " ,ISNULL(pp.FIXSUMCOST,0)+ISNULL(pp.METSUMCOST,0) ALLSUMCOST" '費用總計" & vbCrLf
            sql &= " ,pp.STDATE" '開訓日期" & vbCrLf
            sql &= " ,pp.FTDATE" '結訓日期" & vbCrLf
            sql &= " ,ig3.GCODE2 GCODENAME" '訓練業別編碼" & vbCrLf
            sql &= " ,ig3.CNAME GCNAME" '訓練業別" & vbCrLf
            sql &= " ,kc.CODEID" '訓練職能編碼" & vbCrLf
            sql &= " ,kc.CCNAME" '訓練職能" & vbCrLf
            sql &= " ,dd.D20KNAME,dd.D25KNAME,dd.D26KNAME" & vbCrLf '政府政策性產業
            '5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策
            sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf '"5+2產業創新計畫"" & vbCrLf
            ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
            sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf

            sql &= " ,dd.KNAME22" '進階政策性產業類別" & vbCrLf
            sql &= " ,dd.KNAME15" '轄區重點產業" & vbCrLf
            sql &= " ,pp.POINTYN" '是否為學分班(Y/N)" & vbCrLf
            sql &= " ,pp.CTNAME" '辦訓縣市別" & vbCrLf
            sql &= " ,pp.CONTACTNAME" '聯絡人" & vbCrLf
            sql &= " ,pp.CONTACTPHONE" '聯絡電話" & vbCrLf
            sql &= " ,pp.ISiCAPCOUR" '是否為iCAP課程" & vbCrLf
            sql &= " ,pp.ICAPNUM" 'iCAP標章證號" & vbCrLf
            sql &= " ,(SELECT iz2.CTNAME FROM dbo.VIEW_ZIPNAME iz2 WHERE iz2.ZipCode=pp.ORGZIPCODE) CTNAME2" '立案縣市" & vbCrLf
            sql &= " ,pp.OUTDOOR" '室外教學課程" & vbCrLf
            sql &= " ,pp.REPORTE" '報請主管機關核備" & vbCrLf
            'sql &= " ,dbo.FN_GET_CROSSDIST(pp.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST" ‘跨區/轄區提案" & vbCrLf
            sql &= " ,CASE wo.I_CROSSDIST WHEN -1 THEN '否' ELSE '是' END CROSSDIST" & vbCrLf '/是否跨區提案" & vbCrLf
            'sql &= " "‘CASE wo.I_CROSSDIST WHEN -1 THEN '否' ELSE '是' END CROSSDIST"‘跨區/轄區提案" & vbCrLf
            sql &= " ,pp.TMIDCORRECT " '訓練業別同意協助重新歸類" & vbCrLf
            sql &= " ,pf.OTHFIXCONT " '初審綜合意見-其他應修正內容" & vbCrLf
            sql &= " ,pf.CONFIRMCONT " '初審綜合意見-送請委員確認內容" & vbCrLf

        End If

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=pp.TMID" & vbCrLf
        sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc WITH(NOLOCK) on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " JOIN WORG1 wo on wo.ORGID=pp.ORGID" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.ORG_SCORING2 o2 on o2.OSID2=pp.OSID2" & vbCrLf
        '審查計分等級'19大類主責課程 SYS_GCODEREVIE
        'sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf

        sql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)

        sql &= " AND pp.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND pp.YEARS=@YEARS" & vbCrLf
        sql &= " AND pp.APPSTAGE=@APPSTAGE" & vbCrLf

        If STDate1.Text <> "" Then
            sql &= " AND pp.STDATE >=@STDATE1 "
            parms.Add("STDATE1", TIMS.Cdate2(STDate1.Text))
        End If
        If STDate2.Text <> "" Then
            sql &= " AND pp.STDATE <=@STDATE2 "
            parms.Add("STDATE2", TIMS.Cdate2(STDate2.Text))
        End If
        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
            sql &= " AND pp.RID =@RID" & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If

        '課程申請流水號
        If schPSNO28.Text <> "" Then
            sql &= " AND pp.PSNO28=@PSNO28" & vbCrLf
            parms.Add("PSNO28", schPSNO28.Text)
        End If

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 / ST1RESULT
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
        '審查結果'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        Select Case v_RBL_RESULT_SCH
            Case "2"
                sql &= " AND pf.RESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.RESULT IS NULL" & vbCrLf
            Case "Y", "N", "P"
                sql &= " AND pf.RESULT=@RESULT" & vbCrLf
                parms.Add("RESULT", v_RBL_RESULT_SCH)
        End Select

        '篩選範圍 1:不區分 2:轄區單位 3:19大類主責課程 SYS_GCODEREVIE
        If s_DISTID <> "" Then
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

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                sql &= " and wo.I_CROSSDIST !=-1" & vbCrLf
            Case "J" 'J:轄區提案單位
                sql &= " and wo.I_CROSSDIST =-1" & vbCrLf
        End Select

        '計畫'TRPlanPoint28
        If TRPlanPoint28.Visible Then
            Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
            Select Case v_rblOrgKind2'rblOrgKind2.SelectedValue
                Case "G", "W"
                    sql &= " AND pp.ORGKIND2=@ORGKIND2" & vbCrLf
                    parms.Add("ORGKIND2", v_rblOrgKind2)
            End Select
        End If

        'ROW_NUMBER() OVER(ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE) SEQNUM
        sql &= " ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE" & vbCrLf

        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    Sub sSearch1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        msg1.Text = TIMS.cst_NODATAMsg1
        tbDataGrid1.Visible = False

        Dim dt As DataTable = SEARCH_DATA1_dt(cst_iType_查詢1)
        If TIMS.dtNODATA(dt) Then
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

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Return
        Dim sCMDNM As String = e.CommandName
        If sCMDNM = "" Then Return

        Call CLEAR_DATA1()

        Hid_PSOID.Value = TIMS.GetMyValue(sCmdArg, "PSOID")
        Hid_PSNO28.Value = TIMS.GetMyValue(sCmdArg, "PSNO28")
        Hid_GCODE.Value = TIMS.GetMyValue(sCmdArg, "GCODE")
        Hid_PFGCODE.Value = TIMS.GetMyValue(sCmdArg, "PFGCODE")
        Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))
        If Hid_PSNO28.Value = "" Then Return

        Dim dr1 As DataRow = GET_DATA1()
        Select Case sCMDNM'e.CommandName
            Case "EDT1"
                btnSAVE1.Visible = True
                Call SHOW_DATA1(dr1)
                Call DISABLE_SHOW1(dr1) 'EDT1
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg9)
        End Select
    End Sub

    Private Sub DISABLE_SHOW1(ByRef dr1 As DataRow)
        ', ByRef sCMDNM As String
        If dr1 Is Nothing Then Return

        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        'Dim flag_OTHFIXCONT_OK As Boolean = If(sm.UserInfo.DistID = Convert.ToString(dr1("DISTID")), True, False)
        Dim flag_CONFIRMCONT_OK As Boolean = False
        If v_APPSTAGE_SCH = cst_APPSTAGE_政策性產業_3 Then
            flag_CONFIRMCONT_OK = If(sm.UserInfo.DistID = Convert.ToString(dr1("DISTID")), True, False)
        Else
            flag_CONFIRMCONT_OK = If(sm.UserInfo.DistID = Convert.ToString(dr1("GRDISTID")), True, False)
        End If

        'ST1SUGGEST.Enabled = False ' flag_OTHFIXCONT_OK '分署幕僚意見
        'OTHFIXCONT.Enabled = False 'flag_OTHFIXCONT_OK '其他應修正內容
        'ddlST1RESULT.Enabled = False 'flag_OTHFIXCONT_OK '初審建議結論 --通過、不通過、調整後通過
        CONFIRMCONT.Enabled = flag_CONFIRMCONT_OK '送請委員確認內容
        COMMENTS.Enabled = flag_CONFIRMCONT_OK '委員審查意見與建議
        ddlRESULT.Enabled = flag_CONFIRMCONT_OK  '一階審查結果

        '(1)各分署可針對自己所屬轄區之訓練單位，填寫「其他應修正內容」。
        '(2)19大類的主責分署(設定於功能：首頁>> 課程審查 >> 一階審查 >> 19大類審查分署設定)，可填寫「送請委員確認內容」、「初審建議結論」
        'If (Not ST1SUGGEST.Enabled) Then TIMS.Tooltip(ST1SUGGEST, "所屬轄區之訓練單位，可填寫「分署幕僚意見」", True)
        'If (Not OTHFIXCONT.Enabled) Then TIMS.Tooltip(OTHFIXCONT, "所屬轄區之訓練單位，可填寫「其他應修正內容」", True)
        'If (Not ddlST1RESULT.Enabled) Then TIMS.Tooltip(ddlST1RESULT, "所屬轄區之訓練單位，可填寫「初審建議結論」", True)
        If v_APPSTAGE_SCH = cst_APPSTAGE_政策性產業_3 Then
            If (Not CONFIRMCONT.Enabled) Then TIMS.Tooltip(CONFIRMCONT, "所屬轄區之訓練單位，可填寫「送請委員確認內容」", True)
            If (Not COMMENTS.Enabled) Then TIMS.Tooltip(COMMENTS, "所屬轄區之訓練單位，可填寫「委員審查意見與建議」", True)
            If (Not ddlRESULT.Enabled) Then TIMS.Tooltip(ddlRESULT, "所屬轄區之訓練單位，可填寫「一階審查結果」", True)
        Else
            If (Not CONFIRMCONT.Enabled) Then TIMS.Tooltip(CONFIRMCONT, "19大類的主責分署，可填寫「送請委員確認內容」", True)
            If (Not COMMENTS.Enabled) Then TIMS.Tooltip(COMMENTS, "19大類的主責分署，可填寫「委員審查意見與建議」", True)
            If (Not ddlRESULT.Enabled) Then TIMS.Tooltip(ddlRESULT, "19大類的主責分署，可填寫「一階審查結果」", True)
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                '初審建議結論 Y 通過、N 不通過、P 調整後通過
                Dim labST1RESULT_N As Label = e.Item.FindControl("labST1RESULT_N")
                labST1RESULT_N.Text = Convert.ToString(drv("ST1RESULT_N")) 'TIMS.Get_ST1RESULT_N(Convert.ToString(drv("ST1RESULT_N")))

                'dg_RESULT 審查結果 RESULT 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
                'Dim dg_RESULT As DropDownList = e.Item.FindControl("dg_RESULT")
                'dg_RESULT = TIMS.Get_ST1RESULT(dg_RESULT)
                'Common.SetListItem(dg_RESULT, Convert.ToString(drv("RESULT")))
                '審查結果 RESULT 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
                Dim labRESULT_N As Label = e.Item.FindControl("labRESULT_N")
                labRESULT_N.Text = Convert.ToString(drv("RESULT_N")) 'TIMS.Get_ST1RESULT_N(Convert.ToString(drv("ST1RESULT_N")))

                Dim BtnEDT1 As Button = e.Item.FindControl("BtnEDT1")    '編輯
                BtnEDT1.Visible = If(Convert.ToString(drv("PSOID")) <> "", True, False)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PSOID", Convert.ToString(drv("PSOID")))
                TIMS.SetMyValue(sCmdArg, "PSNO28", Convert.ToString(drv("PSNO28")))
                TIMS.SetMyValue(sCmdArg, "GCODE", Convert.ToString(drv("GCODE")))
                TIMS.SetMyValue(sCmdArg, "PFGCODE", Convert.ToString(drv("PFGCODE")))
                BtnEDT1.CommandArgument = sCmdArg

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
        lbiCAPNUM.Text = ""
        lbD20KNAME.Text = ""
        lbD25KNAME.Text = ""
        lbD26KNAME.Text = ""

        ST1SUGGEST.Text = ""
        OTHFIXCONT.Text = ""
        ddlST1RESULT.SelectedIndex = -1
        Common.SetListItem(ddlST1RESULT, "")

        ddlGCODE.SelectedIndex = -1
        Common.SetListItem(ddlGCODE, "")

        CONFIRMCONT.Text = ""
        COMMENTS.Text = "" 'Convert.ToString(dr("COMMENTS")) '委員審查意見與建議
        ddlRESULT.SelectedIndex = -1
        Common.SetListItem(ddlRESULT, "")
    End Sub

    Function GET_DATA1() As DataRow
        Dim dr1 As DataRow = Nothing
        If Hid_PSNO28.Value = "" Then Return dr1
        Dim parms As New Hashtable From {{"PSNO28", Hid_PSNO28.Value}}

        Dim sql As String = ""
        sql &= " SELECT rr.ORGPLANNAME" & vbCrLf '-- 計畫別、" & vbCrLf
        sql &= " ,pp.PSNO28" & vbCrLf
        sql &= " ,pp.YEARS" & vbCrLf
        sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        'sql &= " ,pp.DISTID,pp.DISTNAME" & vbCrLf '分署別" & vbCrLf
        'sql &= " ,pp.ORGNAME" & vbCrLf '訓練單位名稱" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf 'FIRSTSORT
        sql &= " ,pp.PSNO28" & vbCrLf '課程申請流水號" & vbCrLf
        sql &= " ,pp.CLASSCNAME" & vbCrLf '班級名稱" & vbCrLf
        sql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期" & vbCrLf
        sql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期" & vbCrLf
        sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '/: 訓練業別編碼" & vbCrLf
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sql &= " ,ig3.CNAME GCNAME" & vbCrLf '/訓練業別名稱" & vbCrLf
        'sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '訓練業別編碼" & vbCrLf
        'sql &= " ,ig3.CNAME GCNAME" & vbCrLf '訓練業別名稱" & vbCrLf
        sql &= " ,kc.CCNAME" & vbCrLf '訓練職能" & vbCrLf

        sql &= " ,pp.TNUM" & vbCrLf '訓練人次" & vbCrLf
        sql &= " ,pp.THOURS" & vbCrLf '訓練時數" & vbCrLf
        sql &= " ,pp.ACTHUMCOST" & vbCrLf '實際人時成本" & vbCrLf
        sql &= " ,pp.METSUMCOST" & vbCrLf '實際材料費" & vbCrLf
        sql &= " ,dd.D20KNAME,dd.D25KNAME,dd.D26KNAME" & vbCrLf '政府政策性產業
        '5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策
        sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf '"5+2產業創新計畫"" & vbCrLf
        ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
        sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
        ' 5+2產業--新南向政策--台灣AI行動計畫--數位國家創新經濟發展方案--國家資通安全發展方案--前瞻基礎建設計畫" & vbCrLf
        sql &= " ,dbo.FN_GET_CROSSDIST(pp.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST" & vbCrLf '是否跨區提案" & vbCrLf
        sql &= " ,pp.ICAPNUM" & vbCrLf 'iCAP標章證號" & vbCrLf

        sql &= " ,pf.PSOID" & vbCrLf
        sql &= " ,pf.ST1SUGGEST" & vbCrLf '初審幕僚建議/分署幕僚意見
        sql &= " ,pf.OTHFIXCONT" & vbCrLf '其他應修正內容" & vbCrLf
        sql &= " ,pf.CONFIRMCONT" & vbCrLf '送請委員確認內容" & vbCrLf
        sql &= " ,pf.ST1RESULT" & vbCrLf '初審建議結論" & vbCrLf'初審建議結論 Y 通過、N 不通過、P 調整後通過
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
        '初審建議結論 Y 通過、N 不通過、P 調整後通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END ST1RESULT_N" & vbCrLf
        sql &= " ,pf.RESULT" & vbCrLf '審查結果
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END RESULT_C" & vbCrLf
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END RESULT_N" & vbCrLf
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf
        '分署確認課程分類
        sql &= " ,gc.PFCNAME" & vbCrLf
        '19大類主責課程 SYS_GCODEREVIE
        sql &= " ,gr1.DISTID GRDISTID" & vbCrLf

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf

        '審查計分等級'19大類主責課程 SYS_GCODEREVIE
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        'sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf

        If Hid_PSOID.Value <> "" Then
            sql &= " AND pf.PSOID =@PSOID" & vbCrLf
            parms.Add("PSOID", Hid_PSOID.Value)
        End If
        'sql &= " AND CC.YEARS='2022'" & vbCrLf
        sql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)

        sql &= " AND pp.TPLANID ='28'" & vbCrLf
        sql &= " AND pp.PSNO28 =@PSNO28" & vbCrLf
        dr1 = DbAccess.GetOneRow(sql, objconn, parms)

        Return dr1
    End Function

    Sub SHOW_DATA1(ByRef dr As DataRow)
        PanelSch1.Visible = False
        PanelEdit1.Visible = True
        '初審幕僚建議
        ST1SUGGEST.ReadOnly = True '分署幕僚意見
        OTHFIXCONT.ReadOnly = True  '其他應修正內容

        'CONFIRMCONT.ReadOnly = True  '送請委員確認內容
        'CONFIRMCONT.ApplyStyle(TIMS.GET_RO_STYLE())
        'TIMS.Tooltip(CONFIRMCONT, "僅提供顯示")
        ST1SUGGEST.ApplyStyle(TIMS.GET_RO_STYLE())
        OTHFIXCONT.ApplyStyle(TIMS.GET_RO_STYLE())
        TIMS.Tooltip(ST1SUGGEST, "僅提供顯示", True)
        TIMS.Tooltip(OTHFIXCONT, "僅提供顯示", True)

        ddlST1RESULT.Enabled = False '初審建議結論 --通過、不通過、調整後通過
        TIMS.Tooltip(ddlST1RESULT, "僅提供顯示", True)
        ddlGCODE.Enabled = False
        TIMS.Tooltip(ddlGCODE, "僅提供顯示", True)

        If dr Is Nothing Then Return

        Hid_PSOID.Value = Convert.ToString(dr("PSOID"))
        Hid_PSNO28.Value = Convert.ToString(dr("PSNO28"))

        lbYEARS_ROC.Text = TIMS.GET_YEARS_ROC(dr("YEARS"))
        lbDistName.Text = Convert.ToString(dr("DISTNAME"))
        lbOrgName.Text = Convert.ToString(dr("ORGNAME"))
        lbPSNO28.Text = Convert.ToString(dr("PSNO28"))
        lbClassName.Text = Convert.ToString(dr("CLASSCNAME"))
        lbSFTDate.Text = String.Format("{0}~{1}", dr("STDATE"), dr("STDATE"))

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
        lbiCAPNUM.Text = Convert.ToString(dr("ICAPNUM"))

        lbD20KNAME.Text = $"{dr("D20KNAME")}" '政府政策性產業
        lbD25KNAME.Text = $"{dr("D25KNAME")}" '政府政策性產業
        lbD26KNAME.Text = $"{dr("D26KNAME")}" '政府政策性產業
        If $"{dr("D20KNAME")}{dr("D25KNAME")}{dr("D26KNAME")}" = "" Then lbD20KNAME.Text = "無"

        '初審幕僚建議
        ST1SUGGEST.Text = Convert.ToString(dr("ST1SUGGEST")) '分署幕僚意見
        OTHFIXCONT.Text = Convert.ToString(dr("OTHFIXCONT")) '其他應修正內容
        Common.SetListItem(ddlST1RESULT, dr("ST1RESULT")) '初審建議結論 --通過、不通過、調整後通過
        '分署確認課程分類
        Common.SetListItem(ddlST1RESULT, dr("ST1RESULT"))
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        Hid_GCODE.Value = Convert.ToString(dr("GCODE"))
        Hid_PFGCODE.Value = Convert.ToString(dr("PFGCODE")) '分署確認課程分類
        Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))

        CONFIRMCONT.Text = Convert.ToString(dr("CONFIRMCONT")) '送請委員確認內容
        '審查結果
        COMMENTS.Text = Convert.ToString(dr("COMMENTS")) '委員審查意見與建議
        Common.SetListItem(ddlRESULT, dr("RESULT")) '審查結果
    End Sub

    Protected Sub BtnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        Call CLEAR_DATA1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False
    End Sub

    Protected Sub BtnSAVE1_Click(sender As Object, e As EventArgs) Handles btnSAVE1.Click
        Call SAVE_DATA1()
    End Sub

    Sub SAVE_DATA1()
        Hid_PSOID.Value = TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSOID.Value = "" Then Return
        If Hid_PSNO28.Value = "" Then Return
        'Dim v_ddlST1RESULT As String = TIMS.GetListValue(ddlST1RESULT)
        Dim v_ddlRESULT As String = TIMS.GetListValue(ddlRESULT)

        Dim iRst As Integer = 0
        Dim iPSOID As Integer = Val(Hid_PSOID.Value)
        Dim parms As New Hashtable From {{"PSOID", Val(Hid_PSOID.Value)}, {"PSNO28", Hid_PSNO28.Value}}
        'parms.Add("YEARS", Hid_YEARS.Value) 'parms.Add("APPSTAGE", Hid_APPSTAGE.Value)
        Dim s_sql As String = ""
        s_sql &= " SELECT PSOID FROM PLAN_STAFFOPIN" & vbCrLf
        s_sql &= " WHERE PSOID=@PSOID" & vbCrLf
        s_sql &= " AND PSNO28=@PSNO28" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, parms)
        If dt.Rows.Count = 0 Then Return
        'Dim dr1 As DataRow = dt.Rows(0)

        Dim u_parms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"CONFIRMCONT", CONFIRMCONT.Text},
            {"COMMENTS", COMMENTS.Text},
            {"RESULT", v_ddlRESULT},
            {"RESULTACCT", sm.UserInfo.UserID},
            {"PSOID", iPSOID},
            {"PSNO28", Hid_PSNO28.Value}
        }
        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        u_sql &= " SET MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf

        u_sql &= " ,CONFIRMCONT=@CONFIRMCONT" & vbCrLf
        u_sql &= " ,COMMENTS=@COMMENTS" & vbCrLf
        u_sql &= " ,RESULT=@RESULT" & vbCrLf
        u_sql &= " ,RESULTACCT=@RESULTACCT" & vbCrLf
        u_sql &= " ,RESULTDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PSOID=@PSOID" & vbCrLf
        u_sql &= " AND PSNO28=@PSNO28" & vbCrLf
        iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Call sSearch1()
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

#Region "匯入送請委員確認內容"
    ''' <summary>匯入送請委員確認內容</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnIMPORT1_Click(sender As Object, e As EventArgs) Handles BtnIMPORT1.Click
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入送請委員確認內容檢核有誤!請再確認匯入參數!")
            Return
        End If

        Call ImportXLSX_1(File1)
        Call CCreate1()
    End Sub

    ''' <summary>匯入送請委員確認內容</summary>
    ''' <param name="oFile1"></param>
    Private Sub ImportXLSX_1(ByRef oFile1 As HtmlInputFile)
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
        If Not TIMS.HttpCHKFile(Me, oFile1, MyPostedFile, Cst_Filetype) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        '檢查檔案格式與大小 Start
        If oFile1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        ElseIf oFile1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(oFile1.PostedFile.FileName, "\")((Split(oFile1.PostedFile.FileName, "\")).Length - 1)
        'FileOCIDValue = Split(Split(MyFileName, "-")(1), ".")(0)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> LCase(Cst_Filetype) Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小 End

        Dim Errmag As String = ""
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{cst_Upload_Path}{MyFileName}")
        oFile1.PostedFile.SaveAs(filePath1) '上傳檔案
        '(讀取XLSX檔案轉為dt_xls)
        Dim dt_xls As DataTable = TIMS.ReadXLSX(New IO.FileInfo(filePath1), Errmag)
        '刪除檔案'If IO.File.Exists(filePath1) Then IO.File.Delete(filePath1)
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

                Reason = SAVE_PLAN_STAFFOPIN1(colArray, sHtb)  '驗証(單筆) 並 儲存

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
        Session("MyWrongTable") = Nothing
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

    ''' <summary> 檢核-匯入送請委員確認內容 </summary>
    ''' <param name="ErrMsg1"></param>
    ''' <returns></returns>
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

    ''' <summary>匯入驗証-匯入送請委員確認內容</summary>
    ''' <param name="colArray">比對資料</param>
    ''' <param name="Htb">輸入查詢</param>
    ''' <param name="o_parms">取得有效值</param>
    ''' <returns></returns>
    Function CheckImportData1(ByRef colArray As Array, ByRef Htb As Hashtable, ByRef o_parms As Hashtable) As String
        Dim Reason As String = ""
        s_COL_PSNO28 = ""

        If colArray.Length < cst_col1_iMaxLen Then
            g_IMP_ERR1 = True
            Reason += "欄位對應有誤<BR>,請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        'Dim s_col_PSNO28 As String = "" '課程申請流水號
        If colArray.Length > cst_col1_PSNO28 Then s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col1_PSNO28)) '課程申請流水號

        Dim s_COL_CONFIRMCONT As String = "" '送請委員確認內容

        s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col1_PSNO28)) '課程申請流水號
        s_COL_CONFIRMCONT = TIMS.NullToStr(colArray(cst_col1_CONFIRMCONT)) '送請委員確認內容
        Dim flag_TXT_NG As Boolean = (s_COL_CONFIRMCONT = "")

        '先確認資料不為空
        If s_COL_PSNO28 = "" Then Reason += "課程申請流水號 不可為空<br>"
        'If s_col_OTHFIXCONT = "" Then Reason += "其他應修正內容不可為空<br>"
        'If s_col_CONFIRMCONT = "" Then Reason += "送請委員確認內容不可為空<br>"
        If flag_TXT_NG Then Reason += "送請委員確認內容 不可為空<br>"
        If Reason <> "" Then Return Reason

        Dim s_COL_PCS As String = "" '課程流水號
        s_COL_PCS = TIMS.Get_PCSforPSNO28(sm, s_COL_PSNO28, objconn)
        If s_COL_PCS = "" Then Reason += String.Format("課程申請流水號 有誤，查無班級資料({0})<br>", s_COL_PSNO28)
        If Reason <> "" Then Return Reason

        If Not ChkBxCover1.Checked Then
            Dim flag_EXISTS_1 As Boolean = TIMS.CHK_STAFFOPIN_RESULT_EXISTS(objconn, s_COL_PSNO28)
            If flag_EXISTS_1 Then Reason += String.Format("課程申請流水號 已有審核結果，不再匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
            Dim flag_EXISTS_2 As Boolean = TIMS.CHK_STAFFOPIN_CONFIRMCONT_EXISTS(objconn, s_COL_PSNO28)
            If flag_EXISTS_2 Then Reason += String.Format("課程申請流水號 已有送請委員確認內容，不再匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If

        If sm.UserInfo.LID <> 0 Then
            '19大類主責分署有誤
            Dim flag_DISTID_NG As Boolean = CHK_STAFFOPIN_19_GRDISTID_NG(s_COL_PSNO28, sm.UserInfo.DistID)
            If flag_DISTID_NG Then Reason += String.Format("課程申請流水號 19大類主責分署有誤，不可匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If

        If o_parms Is Nothing Then o_parms = New Hashtable
        o_parms.Add("PSNO28", s_COL_PSNO28)
        o_parms.Add("CONFIRMCONT", s_COL_CONFIRMCONT)
        Return Reason
    End Function
#End Region

#Region "匯入審查結果"
    ''' <summary>匯入審查結果</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnIMPORT2_Click(sender As Object, e As EventArgs) Handles BtnIMPORT2.Click
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp2(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入審查結果檢核有誤!請再確認匯入參數!")
            Return
        End If

        Call ImportXLSX_2(File2)
        Call CCreate1()
    End Sub

    ''' <summary>匯入審查結果</summary>
    ''' <param name="oFile1"></param>
    Private Sub ImportXLSX_2(ByRef oFile1 As HtmlInputFile)
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp2(ErrMsg1)
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
        If Not TIMS.HttpCHKFile(Me, oFile1, MyPostedFile, Cst_Filetype, 1) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        '檢查檔案格式與大小 Start
        If oFile1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        ElseIf oFile1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(oFile1.PostedFile.FileName, "\")((Split(oFile1.PostedFile.FileName, "\")).Length - 1)
        'FileOCIDValue = Split(Split(MyFileName, "-")(1), ".")(0)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> LCase(Cst_Filetype) Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小 End

        'Dim dt_xls As DataTable
        'Dim Errmag As String = ""
        'oFile1.PostedFile.SaveAs(Server.MapPath(cst_Upload_Path & MyFileName)) '上傳檔案
        'dt_xls = TIMS.GetDataTable_XlsFile(Server.MapPath(cst_Upload_Path & MyFileName).ToString, "", Errmag, "課程申請流水號") '取得內容
        'IO.File.Delete(Server.MapPath(cst_Upload_Path & MyFileName)) '刪除檔案
        Dim Errmag As String = ""

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath(cst_Upload_Path & MyFileName)
        oFile1.PostedFile.SaveAs(filePath1) '上傳檔案
        '(讀取XLSX檔案轉為dt_xls)
        Dim dt_xls As DataTable = TIMS.ReadXLSX(New IO.FileInfo(filePath1), Errmag)
        '刪除檔案 'If IO.File.Exists(FullFileName1) Then IO.File.Delete(FullFileName1)
        Call TIMS.MyFileDelete(filePath1)

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

                Reason = SAVE_PLAN_STAFFOPIN2(colArray, sHtb)  '驗証(單筆) 並 儲存

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
        Session("MyWrongTable") = Nothing
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

    ''' <summary> 檢核-匯入審查結果 </summary>
    ''' <param name="ErrMsg1"></param>
    ''' <returns></returns>
    Function CheckImp2(ByRef ErrMsg1 As String) As Boolean
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

    ''' <summary>匯入驗証-匯入審查結果</summary>
    ''' <param name="colArray">比對資料</param>
    ''' <param name="Htb">輸入查詢</param>
    ''' <param name="o_parms">取得有效值</param>
    ''' <returns></returns>
    Function CheckImportData2(ByRef colArray As Array, ByRef Htb As Hashtable, ByRef o_parms As Hashtable) As String
        Dim Reason As String = ""
        s_COL_PSNO28 = ""

        If colArray.Length < cst_col2_iMaxLen Then
            g_IMP_ERR1 = True
            Reason += "欄位對應有誤<BR>,請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        'Dim s_col_PSNO28 As String = "" '課程申請流水號
        If colArray.Length > cst_col2_PSNO28 Then s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col2_PSNO28)) '課程申請流水號

        'Dim s_COL_COMMENTS As String = "" '委員審查意見與建議
        'Dim s_COL_RESULT As String = "" '審核結果
        'Dim s_COL_RESULT_YNP As String = "" '審核結果(YNP)

        s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col2_PSNO28)) '課程申請流水號
        Dim s_COL_COMMENTS As String = TIMS.NullToStr(colArray(cst_col2_COMMENTS)) '委員審查意見與建議
        Dim s_COL_RESULT As String = TIMS.ClearSQM(colArray(cst_col2_RESULT)) '初審建議結論
        'Dim flag_TXT_NG As Boolean = (s_COL_COMMENTS = "")

        '先確認資料不為空
        If s_COL_PSNO28 = "" Then Reason += "課程申請流水號 不可為空<br>"
        'If s_col_OTHFIXCONT = "" Then Reason += "其他應修正內容不可為空<br>"
        'If s_col_CONFIRMCONT = "" Then Reason += "送請委員確認內容不可為空<br>"
        If (s_COL_COMMENTS = "") AndAlso s_COL_RESULT = "" Then Reason += "委員審查意見與建議/審核結果 不可皆為空<br>"
        If Reason <> "" Then Return Reason

        'Dim s_COL_PCS As String = "" '課程流水號
        Dim s_COL_PCS As String = TIMS.Get_PCSforPSNO28(sm, s_COL_PSNO28, objconn)
        If s_COL_PCS = "" Then Reason += String.Format("課程申請流水號 有誤，查無班級資料({0})<br>", s_COL_PSNO28)
        If Reason <> "" Then Return Reason

        'Y/N/P'用文字方式輸入 '審核結果(YNP)
        Dim s_COL_RESULT_YNP As String = If(s_COL_RESULT = "通過", "Y", If(s_COL_RESULT = "不通過", "N", If(s_COL_RESULT = "調整後通過", "P", "")))
        s_COL_RESULT_YNP = If(s_COL_RESULT_YNP <> "", s_COL_RESULT_YNP, If(s_COL_RESULT = "1", "Y", If(s_COL_RESULT = "3", "N", If(s_COL_RESULT = "2", "P", ""))))
        If s_COL_RESULT_YNP = "" AndAlso s_COL_RESULT <> "" Then
            Select Case s_COL_RESULT '用代碼方式輸入
                Case "Y", "P", "N"
                    s_COL_RESULT_YNP = s_COL_RESULT
            End Select
        End If
        'If flag_TXT_NG AndAlso s_COL_RESULT = "" Then Reason += "委員審查意見與建議/審核結果 資料皆為空<br>"
        'If flag_TXT_NG AndAlso s_col_ST1RESULT_YNP = "" Then Reason += String.Format("初審建議結論 有誤，(通過/調整後通過/不通過):{0}<br>", s_col_ST1RESULT)
        'Select Case s_col_ST1RESULT_YNP
        '    Case "Y"
        '    Case Else
        '        If s_col_OTHFIXCONT = "" Then Reason += "其他應修正內容不可為空<br>"
        '        If s_col_CONFIRMCONT = "" Then Reason += "送請委員確認內容不可為空<br>"
        'End Select

        If Not ChkBxCover2.Checked Then
            Dim flag_EXISTS_1 As Boolean = TIMS.CHK_STAFFOPIN_RESULT_EXISTS(objconn, s_COL_PSNO28)
            If flag_EXISTS_1 Then Reason += String.Format("課程申請流水號 已有審核結果，不再匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason

            Dim flag_EXISTS_3 As Boolean = TIMS.CHK_STAFFOPIN_COMMENTS_EXISTS(objconn, s_COL_PSNO28)
            If flag_EXISTS_3 Then Reason += String.Format("課程申請流水號 已有委員審查意見與建議，不再匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If

        If sm.UserInfo.LID <> 0 Then
            '19大類主責分署有誤
            Dim flag_DISTID_NG As Boolean = CHK_STAFFOPIN_19_GRDISTID_NG(s_COL_PSNO28, sm.UserInfo.DistID)
            If flag_DISTID_NG Then Reason += String.Format("課程申請流水號 19大類主責分署有誤，不可匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If

        If o_parms Is Nothing Then o_parms = New Hashtable
        o_parms.Add("PSNO28", s_COL_PSNO28)
        o_parms.Add("COMMENTS", s_COL_COMMENTS)
        o_parms.Add("RESULT", s_COL_RESULT_YNP)
        Return Reason
    End Function
#End Region

    ''' <summary>課程申請流水號 19大類主責分署有誤，不可匯入</summary>
    ''' <param name="s_PSNO28"></param>
    ''' <param name="s_GRDISTID"></param>
    ''' <returns></returns>
    Function CHK_STAFFOPIN_19_GRDISTID_NG(ByVal s_PSNO28 As String, ByVal s_GRDISTID As String) As Boolean
        Dim rst As Boolean = False
        If s_PSNO28 = "" Then Return rst
        Dim dt1 As New DataTable
        Dim s_sql As String = ""
        s_sql &= " SELECT pf.PSOID,pp.DISTID"
        s_sql &= " ,gr1.DISTID GRDISTID" & vbCrLf '19大類主責課程 SYS_GCODEREVIE
        s_sql &= " FROM dbo.PLAN_STAFFOPIN pf WITH(NOLOCK)" & vbCrLf
        s_sql &= " JOIN dbo.VIEW2B pp on pp.PSNO28=pf.PSNO28" & vbCrLf
        s_sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        's_sql &= " JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        s_sql &= " JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        s_sql &= " WHERE pf.PSNO28=@PSNO28 AND gr1.DISTID!=@GRDISTID" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            .Parameters.Add("GRDISTID", SqlDbType.VarChar).Value = s_GRDISTID
            dt1.Load(.ExecuteReader)
        End With
        rst = (dt1.Rows.Count > 0)
        Return rst
    End Function

    ''' <summary>取得流水號</summary>
    ''' <param name="s_PSNO28"></param>
    ''' <returns></returns>
    Function GET_PLAN_STAFFOPIN_PSOID(ByVal s_PSNO28 As String) As Integer
        Dim rst As Integer = -1
        If s_PSNO28 = "" Then Return rst
        Dim dt1 As New DataTable
        Dim s_sql As String = " SELECT PSOID FROM PLAN_STAFFOPIN WITH(NOLOCK) WHERE PSNO28=@PSNO28" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            dt1.Load(.ExecuteReader)
        End With
        If (dt1.Rows.Count = 1) Then rst = dt1.Rows(0)("PSOID")
        Return rst
    End Function

    ''' <summary>匯入檔[儲存]-匯入送請委員確認內容</summary>
    ''' <param name="colArray"></param>
    ''' <param name="Htb"></param>
    ''' <returns></returns>
    Function SAVE_PLAN_STAFFOPIN1(ByRef colArray As Array, ByRef Htb As Hashtable) As String
        Dim o_parms As New Hashtable
        Dim rst As String = CheckImportData1(colArray, Htb, o_parms)
        If rst <> "" Then Return rst

        Dim s_PSNO28 As String = TIMS.GetMyValue2(o_parms, "PSNO28")
        Dim vCONFIRMCONT As String = TIMS.GetMyValue2(o_parms, "CONFIRMCONT")
        'Dim vCOMMENTS As String = TIMS.GetMyValue2(o_parms, "COMMENTS")
        'Dim vRESULT As String = TIMS.GetMyValue2(o_parms, "RESULT")

        'Dim flag_EXISTS_1 As Boolean = TIMS.CHK_STAFFOPIN_RESULT_EXISTS(objconn, s_PSNO28)
        'If flag_EXISTS_1 Then Return rst
        Dim iPSOID As Integer = GET_PLAN_STAFFOPIN_PSOID(s_PSNO28)
        If iPSOID = -1 Then Return rst

        'u_parms.Add("COMMENTS", If(vCOMMENTS <> "", vCOMMENTS, Convert.DBNull))
        'u_parms.Add("RESULT", If(vRESULT <> "", vRESULT, Convert.DBNull))
        Dim u_parms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"CONFIRMCONT", If(vCONFIRMCONT <> "", vCONFIRMCONT, Convert.DBNull)},
            {"RESULTACCT", sm.UserInfo.UserID},
            {"PSOID", iPSOID},
            {"PSNO28", s_PSNO28}
        }
        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        u_sql &= " SET MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " ,CONFIRMCONT=@CONFIRMCONT" & vbCrLf
        'u_sql &= " ,COMMENTS=@COMMENTS,RESULT=@RESULT" & vbCrLf
        'u_sql &= " ,RESULTACCT=@RESULTACCT ,RESULTDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        Return rst
    End Function

    ''' <summary>匯入檔[儲存]-匯入審查結果</summary>
    ''' <param name="colArray"></param>
    ''' <param name="Htb"></param>
    ''' <returns></returns>
    Function SAVE_PLAN_STAFFOPIN2(ByRef colArray As Array, ByRef Htb As Hashtable) As String
        Dim o_parms As New Hashtable
        Dim rst As String = CheckImportData2(colArray, Htb, o_parms)
        If rst <> "" Then Return rst

        Dim s_PSNO28 As String = TIMS.GetMyValue2(o_parms, "PSNO28")
        'Dim vCONFIRMCONT As String = TIMS.GetMyValue2(o_parms, "CONFIRMCONT")
        Dim vCOMMENTS As String = TIMS.GetMyValue2(o_parms, "COMMENTS")
        Dim vRESULT As String = TIMS.GetMyValue2(o_parms, "RESULT")

        'Dim flag_EXISTS_1 As Boolean = TIMS.CHK_STAFFOPIN_RESULT_EXISTS(objconn, s_PSNO28)
        'If flag_EXISTS_1 Then Return rst
        Dim iPSOID As Integer = GET_PLAN_STAFFOPIN_PSOID(s_PSNO28)
        If iPSOID = -1 Then Return rst

        'u_parms.Add("CONFIRMCONT", If(vCONFIRMCONT <> "", vCONFIRMCONT, Convert.DBNull))
        Dim u_parms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"COMMENTS", If(vCOMMENTS <> "", vCOMMENTS, Convert.DBNull)},
            {"RESULT", If(vRESULT <> "", vRESULT, Convert.DBNull)},
            {"RESULTACCT", sm.UserInfo.UserID},
            {"PSOID", iPSOID},
            {"PSNO28", s_PSNO28}
        }
        'i_parms.Add("ST1ACCT", sm.UserInfo.UserID)
        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        u_sql &= " SET MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        ',CONFIRMCONT=@CONFIRMCONT
        u_sql &= " ,COMMENTS=@COMMENTS,RESULT=@RESULT" & vbCrLf
        u_sql &= " ,RESULTACCT=@RESULTACCT ,RESULTDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        Return rst
    End Function

    Protected Sub BtnExport2_Click(sender As Object, e As EventArgs) Handles BtnExport2.Click
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = cst_APPSTAGE_政策性產業_3 Then
            Call EXPORT_3()
            Return
        End If
        Call EXPORT_2()
    End Sub

    ''' <summary>跨區/審查意見綜整表</summary>
    Sub EXPORT_2()
        '序號, 單位代碼, 課程申請流水號, 分署別	計畫, 訓練單位名稱, 課程名稱, iCAP標章證號, 初審綜合意見 - 其他應修正內容述明, 初審綜合意見 - 送請委員確認內容述明, 初審綜合意見, 同單位同類課程建議修正意見, 同單位同類課程(班數), 課程分類, 訓練業別編碼, 訓練業別, 備註
        Dim dtX1 As DataTable = SEARCH_DATA1_dt(cst_iType_匯出12)

        Dim sPattern As String = "" '序號,
        Dim sColumn As String = ""
        'CONFIRMCONT 送請委員確認內容 送請委員確認內容
        sPattern = "單位代碼,課程申請流水號,分署別,計畫別,訓練單位名稱,課程名稱,期別,iCAP標章證號,是否為iCAP課程,初審綜合意見-其他應修正內容述明,初審綜合意見-送請委員確認內容述明,初審綜合意見,同單位同類課程建議修正意見,同單位同類課程(班數)"
        sPattern &= ",課程分類,分署確認課程分類,訓練業別編碼,訓練業別,備註"
        sColumn = "ORGID,PSNO28,DISTNAME,ORGPLANNAME,ORGNAME,CLASSCNAME,CYCLTYPE,ICAPNUM,ISiCAPCOUR,OTHFIXCONT,CONFIRMCONT,ST1RESULT_N,SAMEOCFIXREC,SAMEOCNT1"
        sColumn &= ",GCODEPNAME,PFCNAME,GCODENAME,GCNAME,REMARK1"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        '跨區/審查意見綜整表_ 
        Dim s_FILENAME1 As String = String.Concat("審查意見綜整表_", TIMS.GetDateNo2(3))
        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr &= "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "序號") '& vbTab
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

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
            strHTML &= ExportStr
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)}, 'EXCEL/PDF/ODS
            {"FileName", s_FILENAME1},
            {"strSTYLE", strSTYLE},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'TIMS.CloseDbConn(objconn)
        'Response.End()
    End Sub

    ''' <summary>「政策性產業」之「匯出審查意見綜整表」</summary>
    Sub EXPORT_3()
        '序號,計畫別,訓練單位名稱,申請階段,統一編號,單位屬性,分署別,課程名稱,提案意願順序,課程申請流水號,課程分類編碼,課程分類,訓練時數,訓練人次
        ',每人訓練費用(元),訓練單位可向學員收取之訓練費用(元),總補助費(元)(以訓練費用之80%估算),術科時數
        ',固定費用總計,實際人時成本,材料明細,材料費總計,材料費占比,費用總計,開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能
        ',5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策,進階政策性產業類別,轄區重點產業
        ',是否為學分班(Y/N),辦訓縣市別,聯絡人,聯絡電話,是否為iCAP課程,iCAP標章證號,立案縣市,室外教學課程,報請主管機關核備,跨區/轄區提案
        ',訓練業別同意協助重新歸類,初審綜合意見-其他應修正內容,初審綜合意見-送請委員確認內容
        Dim dtX1 As DataTable = SEARCH_DATA1_dt(cst_iType_匯出13)

        Dim sPattern As String = "" '序號,
        sPattern &= "計畫別,訓練單位名稱,申請階段,統一編號,單位屬性,分署別,課程名稱,提案意願順序,課程申請流水號,課程分類編碼,課程分類,訓練時數,訓練人次"
        sPattern &= ",每人訓練費用(元),訓練單位可向學員收取之訓練費用(元),總補助費(元)(以訓練費用之80%估算),術科時數"
        sPattern &= ",固定費用總計,實際人時成本,材料明細,材料費總計,材料費占比,費用總計,開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能"
        'sPattern &= ",5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策,進階政策性產業類別,轄區重點產業"
        sPattern &= ",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航,進階政策性產業類別,轄區重點產業"
        sPattern &= ",是否為學分班(Y/N),辦訓縣市別,聯絡人,聯絡電話,是否為iCAP課程,iCAP標章證號,立案縣市,室外教學課程,報請主管機關核備,跨區/轄區提案"
        sPattern &= ",訓練業別同意協助重新歸類,初審綜合意見-其他應修正內容,初審綜合意見-送請委員確認內容"

        Dim sColumn As String = "" '序號,
        sColumn &= "ORGPLANNAME2,ORGNAME,APPSTAGE_N,COMIDNO,ORGTYPENAME,DISTNAME,CLASSCNAME,FIRSTSORT,PSNO28,GCODE,GCODEPNAME,THOURS,TNUM"
        sColumn &= ",DEFGOVCOST1,TOTALCOST,DEFGOVCOST,PROTECHHOURS"
        sColumn &= ",FIXSUMCOST,ACTHUMCOST,METDET,METSUMCOST,METCOSTPER,ALLSUMCOST,STDATE,FTDATE,GCODENAME,GCNAME,CODEID,CCNAME"
        'sColumn &= ",D20KNAME1,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5,D20KNAME6,KNAME22,KNAME15"
        sColumn &= ",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,D25KNAME7,D25KNAME8,KNAME22,KNAME15"
        sColumn &= ",POINTYN,CTNAME,CONTACTNAME,CONTACTPHONE,ISiCAPCOUR,ICAPNUM,CTNAME2,OUTDOOR,REPORTE,CROSSDIST"
        sColumn &= ",TMIDCORRECT,OTHFIXCONT,CONFIRMCONT"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        '跨區/審查意見綜整表_ 
        Dim s_FILENAME1 As String = String.Concat("審查意見綜整表_", TIMS.GetDateNo2(3))
        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr &= "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "序號") '& vbTab
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

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
            strHTML &= ExportStr
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)}, 'EXCEL/PDF/ODS
            {"FileName", s_FILENAME1},
            {"strSTYLE", strSTYLE},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'TIMS.CloseDbConn(objconn) 'Response.End()
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    '匯出  '表單02_跨區/提案彙整表(線上填表).xlsx
    'Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
    '    Call EXPORT_1()
    'End Sub

    '匯出審查意見綜整表 表單03_跨區/審查意見綜整表.xlsx
    'Protected Sub BtnExport2_Click(sender As Object, e As EventArgs) Handles BtnExport2.Click
    '    Call EXPORT_2()
    'End Sub
End Class

