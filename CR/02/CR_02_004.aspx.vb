Public Class CR_02_004
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
    '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
    Dim fg_Work2026x02 As Boolean = False 'TIMS.SHOW_W2026x02(sm)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        fg_Work2026x02 = TIMS.SHOW_W2026x02(sm)

        If Not IsPostBack Then
            CCreate1()
        End If

        '委訓
        'Select Case sm.UserInfo.LID
        '    Case 2
        '        Button2.Visible = False
        '    Case Else
        '        'Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        '        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
        '            '署(局) 或 分署(中心)
        '            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        '            If HistoryRID.Rows.Count <> 0 Then
        '                center.Attributes("onclick") = "showObj('HistoryList2');"
        '                center.Style("CURSOR") = "hand"
        '            End If
        '            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        '        End If
        'End Select
    End Sub

    Public Sub CCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)
        Select Case sm.UserInfo.LID
            Case 0
                ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, TIMS.Get_DISTIDdt(objconn))
            Case Else
                ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, TIMS.Get_DISTIDT2(objconn))
        End Select
        Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '訓練機構
        'center.Text = sm.UserInfo.OrgName
        'RIDValue.Value = sm.UserInfo.RID

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ 

        '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位 '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位

        '初審建議結論 Y 通過、N 不通過、P 調整後通過
        'ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        'Result 初審建議結論 / 審查結果 - -Y 通過、N 不通過、P 調整後通過
        'ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)
    End Sub

#Region "NO USE"
    'Function GET_ORG_SQL1() As String
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
    '    sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
    '    sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
    '    Return sql
    'End Function

    'Function GET_CLASS_SQL1(ByRef parms As Hashtable) As String
    '    'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
    '    Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
    '    '訓練機構
    '    'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
    '    'If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
    '    'Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
    '    Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH)

    '    '計畫'TRPlanPoint28
    '    Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
    '    '開訓日期
    '    'STDate1.Text = TIMS.cdate3(STDate1.Text)
    '    'STDate2.Text = TIMS.cdate3(STDate2.Text)

    '    parms.Add("YEARS", sm.UserInfo.Years)
    '    parms.Add("TPLANID", sm.UserInfo.TPlanID)
    '    parms.Add("DISTID", v_ddlDISTID_SCH)
    '    parms.Add("APPSTAGE", v_APPSTAGE_SCH)
    '    '計畫'TRPlanPoint28
    '    parms.Add("ORGKIND2", v_rblOrgKind2)

    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT cc.RID,cc.PSNO28,cc.OCID,cc.DISTID,cc.DISTNAME" & vbCrLf
    '    sql &= " ,cc.APPSTAGE,cc.STDATE,cc.FTDATE" & vbCrLf
    '    sql &= " ,cc.TOTALCOST,cc.DEFGOVCOST,cc.DEFSTDCOST" & vbCrLf
    '    sql &= " ,cc.TPLANID,cc.YEARS" & vbCrLf ',cc.MONTHS
    '    sql &= " ,oo.RSID, OO.ORGLEVEL,oo.PLANID,cc.ORGKIND2,OO.ORGID" & vbCrLf
    '    sql &= " ,OO.COMIDNO" & vbCrLf
    '    sql &= " ,OO.ORGTYPENAME" & vbCrLf
    '    sql &= " ,OO.ORGNAME" & vbCrLf
    '    sql &= " ,OO.MASTERNAME" & vbCrLf
    '    sql &= " ,dbo.FN_GET_CSTUDCNT14(cc.OCID) CSTUDCNT14" & vbCrLf
    '    'sql &= " /*建議結論 Y 通過、N 不通過、P 調整後通過*/" & vbCrLf
    '    'sql &= " AND cp.RESULT IN ('Y','P')" & vbCrLf
    '    sql &= " ,cp.RESULT" & vbCrLf
    '    sql &= " FROM dbo.VIEW2 cc" & vbCrLf
    '    sql &= " JOIN dbo.VIEW_ORGPLANINFO oo on oo.RID=cc.RID" & vbCrLf
    '    sql &= " JOIN dbo.PLAN_STAFFOPIN cp on cp.PSNO28=cc.PSNO28" & vbCrLf
    '    sql &= " WHERE 1=1" & vbCrLf
    '    sql &= " AND OO.ORGLEVEL=2" & vbCrLf
    '    'sql &= " AND cc.TPLANID='28' AND cc.YEARS='2022' AND cc.DISTID='001' AND cc.APPSTAGE=1 AND cc.ORGKIND2='G'" & vbCrLf
    '    'sql &= " /*建議結論 Y 通過、N 不通過、P 調整後通過*/" & vbCrLf
    '    'sql &= " AND cp.RESULT IN ('Y','P')" & vbCrLf
    '    sql &= " AND cc.YEARS=@YEARS" & vbCrLf
    '    sql &= " AND cc.TPLANID=@TPLANID" & vbCrLf
    '    sql &= " AND cc.DISTID=@DISTID" & vbCrLf
    '    sql &= " AND cc.APPSTAGE=@APPSTAGE" & vbCrLf
    '    '計畫'TRPlanPoint28
    '    sql &= " AND cc.ORGKIND2=@ORGKIND2" & vbCrLf

    '    'If STDate1.Text <> "" Then
    '    '    sql &= " and cc.STDATE >=@STDATE1 "
    '    '    parms.Add("STDATE1", TIMS.cdate2(STDate1.Text))
    '    'End If
    '    'If STDate2.Text <> "" Then
    '    '    sql &= " and cc.STDATE <=@STDATE2 "
    '    '    parms.Add("STDATE2", TIMS.cdate2(STDate2.Text))
    '    'End If

    '    'If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
    '    '    sql &= " AND cc.RID =@RID  " & vbCrLf
    '    '    parms.Add("RID", RIDValue.Value)
    '    'End If
    '    Return sql
    'End Function
#End Region

    Function GET_CLASS_SQL2(ByRef parms As Hashtable, ByVal s_CURESULT As String) As String
        's_CURESULT 匯出格式	Y:通過明細表 N:未通過明細表
        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH)

        '訓練機構
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        'Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        If v_APPSTAGE_SCH = "3" Then v_rblOrgKind2 = ""
        '開訓日期 'STDate1.Text = TIMS.cdate3(STDate1.Text) 'STDate2.Text = TIMS.cdate3(STDate2.Text)
        parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("APPSTAGE", v_APPSTAGE_SCH)
        '計畫'TRPlanPoint28
        If (v_rblOrgKind2 <> "") Then parms.Add("ORGKIND2", v_rblOrgKind2)

        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT cc.PSNO28" & vbCrLf
        sql &= " ,dbo.FN_GET_DISTNAME(cc.DISTID,3) DISTNAME3,cc.ORGKIND2,cc.ORGPLANNAME2" & vbCrLf
        sql &= " ,cc.ORGNAME,cc.COMIDNO,cc.ORGTYPENAME2,cc.TYPEID2NAME,cc.DISTID,cc.DISTNAME,cc.CLASSCNAME2,cc.FIRSTSORT" & vbCrLf
        sql &= " ,cc.PLANID,cc.SEQNO,cc.YEARS,cc.APPSTAGE" & vbCrLf
        sql &= " ,ig3.GCODE31,ig3.PNAME" & vbCrLf '【提案】課程分類/課程分類,ig3.CNAME ,課程分類
        sql &= " ,CC.THOURS,CC.TNUM" & vbCrLf
        sql &= " ,CC.TOTAL,CC.TOTALCOST" & vbCrLf
        sql &= " ,CC.DEFSTDCOST,CC.DEFGOVCOST,CC.DEFSTDCOST1,CC.DEFGOVCOST1" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC(cc.PLANID,cc.COMIDNO,cc.SEQNO,2) PROTECHHOURS" & vbCrLf
        sql &= " ,cc.FIXSUMCOST,cc.ACTHUMCOST" & vbCrLf
        sql &= " ,cc.METSUMCOST,cc.METCOSTPER" & vbCrLf
        sql &= " ,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf
        sql &= " ,ig3.CNAME GCNAME" & vbCrLf
        sql &= " ,kc.CODEID,kc.CCNAME" & vbCrLf
        sql &= " ,dd.KNAME15" & vbCrLf
        sql &= " ,dd.D20KNAME1,dd.D20KNAME6,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5" & vbCrLf
        ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
        sql &= " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8" & vbCrLf
        sql &= " ,dd.D26KNAME1,dd.D26KNAME2,dd.D26KNAME3,dd.D26KNAME4,dd.D26KNAME5,dd.D26KNAME6,dd.D26KNAME7,dd.D26KNAME8,dd.D26KNAME9"
        sql &= " ,cc.POINTYN,cc.CTNAME,cc.CONTACTNAME,cc.CONTACTPHONE,cc.CONTACTMOBILE" & vbCrLf
        sql &= " ,iz2.CTNAME CTNAME2" & vbCrLf
        sql &= " ,os.IMPLEVEL_1 IMPLEVEL1" & vbCrLf
        sql &= " ,os.IMPSCORE_1 IMPSCORE1" & vbCrLf
        '複審通過
        sql &= " ,os.SECONDCHK" & vbCrLf
        If v_APPSTAGE_SCH = "3" Then
            sql &= " ,dbo.FN_SCORING2_TOTALSCORE(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE) TOTALSCORE" & vbCrLf
            sql &= " ,dbo.FN_SCORING2_RLEVEL_2(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE) RLEVEL_2" & vbCrLf
        Else
            '總分-複審等級-通過才顯示
            '計算優先順序：'【分署小計】+署/部加分項目/【匯入成績】'【匯入成績】優先 【分署小計】
            '(1)	當【匯入成績】有資料時 '總分 =【匯入成績】+【署/部加分項目】
            '(2)	當【匯入成績】為空、【分署小計】有資料時 '總分 =【分署小計】+【署/部加分項目】
            '(3)	當【匯入成績】、【分署小計】皆為空時     '總分 = 顯示 "-"
            sql &= " ,CASE WHEN os.SECONDCHK='Y' THEN case when os.IMPSCORE_1>0 then convert(varchar,os.IMPSCORE_1+isnull(os.BRANCHPNT,0))" & vbCrLf
            sql &= " when os.SUBTOTAL>0 then convert(varchar,os.SUBTOTAL+isnull(os.BRANCHPNT,0)) else '-' end END TOTALSCORE" & vbCrLf
            '複審等級-通過才顯示
            sql &= " ,case when os.SECONDCHK='Y' THEN os.RLEVEL_2 END RLEVEL_2" & vbCrLf
        End If

        sql &= " ,pf.ST1SUGGEST" & vbCrLf '初審幕僚建議/分署幕僚建議/分署幕僚意見
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
        sql &= " ,pf.ST1RESULT" & vbCrLf '初審建議結論
        sql &= " ,pf.OTHFIXCONT" & vbCrLf '其他應修正內容
        sql &= " ,pf.CONFIRMCONT" & vbCrLf '送請委員確認內容
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf 'PFGCODE 1.「課程分類」欄位：撈取審查幕僚意見 -「分署確認課程分類」資料
        sql &= " ,(SELECT x.PFCNAME FROM V_GOVCLASS x WHERE x.GCODE=pf.GCODE) PFCNAME" & vbCrLf 'PFCNAME/PFGCODE 1.「課程分類」欄位：撈取審查幕僚意見 -「分署確認課程分類」資料
        sql &= " ,cc.ICAPNUM" & vbCrLf '備註(如有iCap課程請註明，並註明為額度內或額度外）
        '是否為跨區單位課程 (Y/N)"
        sql &= " ,dbo.FN_GET_CROSSDIST_YN(cc.YEARS,cc.COMIDNO,cc.APPSTAGE) CROSSDIST_YN" & vbCrLf
        sql &= " ,pf.RESULT" & vbCrLf '一階審查結果
        sql &= " ,pf.CURESULT" & vbCrLf '核班結果
        sql &= " ,pf.NGREASON" & vbCrLf '不通過明細表增加「未核班原因」，說明如下

        sql &= " ,op.MASTERNAME" & vbCrLf '負責人(產投計畫)
        sql &= " FROM dbo.VIEW2B cc" & vbCrLf
        sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=cc.GCID3" & vbCrLf
        sql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.planid=cc.planid and dd.comidno=cc.comidno and dd.seqno=cc.seqno" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc WITH(NOLOCK) on kc.CCID=cc.CLASSCATE" & vbCrLf
        sql &= " JOIN dbo.VIEW_ORGPLANINFO op on op.RID=cc.RID" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz2 on iz2.ZipCode=cc.ORGZIPCODE" & vbCrLf
        sql &= " LEFT JOIN dbo.ORG_SCORING2 os on os.OSID2=cc.OSID2" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=cc.PSNO28" & vbCrLf
        sql &= " WHERE (cc.RESULTBUTTON IS NULL OR cc.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND cc.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        'sql &= " AND cc.TPLANID='28' AND cc.YEARS='2022' AND cc.DISTID='001' AND cc.APPSTAGE=1 AND cc.ORGKIND2='G'" & vbCrLf
        sql &= " AND cc.TPLANID=@TPLANID AND cc.YEARS=@YEARS" & vbCrLf
        Dim fg_NO_DISTID As Boolean = (v_ddlDISTID_SCH = "000" AndAlso sm.UserInfo.LID = 0)
        If Not fg_NO_DISTID Then
            parms.Add("DISTID", v_ddlDISTID_SCH)
            sql &= " AND cc.DISTID=@DISTID" & vbCrLf
        End If

        'sql &= If(s_RESULT = "Y", " AND pf.RESULT IN ('Y','P')", If(s_RESULT = "N", " AND pf.RESULT ='N'", ""))
        '改 核班結果：通過/不通過 判斷  '只能是Y/N
        sql &= If(s_CURESULT = "Y", " AND pf.CURESULT='Y'", If(s_CURESULT = "N", " AND pf.CURESULT ='N'", " AND 1!=1"))
        sql &= " AND cc.APPSTAGE=@APPSTAGE"
        If (v_rblOrgKind2 <> "") Then sql &= " AND cc.ORGKIND2=@ORGKIND2"
        sql &= " )" & vbCrLf

        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY cc.ORGKIND2,cc.ORGNAME,cc.FIRSTSORT,cc.STDATE) SEQNUM" & vbCrLf '序號
        sql &= " ,cc.ORGNAME" & vbCrLf '訓練單位名稱
        sql &= " ,cc.COMIDNO" & vbCrLf '統一編號
        sql &= " ,cc.ORGTYPENAME2" & vbCrLf '單位屬性／(產投)單位屬性
        sql &= " ,cc.TYPEID2NAME" & vbCrLf '機構別／(產投)單位屬性
        sql &= " ,cc.ORGKIND2,cc.ORGPLANNAME2" & vbCrLf '計畫別
        sql &= " ,cc.DISTNAME,cc.DISTNAME3" & vbCrLf '分署別
        sql &= " ,cc.CLASSCNAME2" & vbCrLf '課程名稱(含期別)
        sql &= " ,cc.PSNO28" & vbCrLf '課程申請流水號
        sql &= " ,cc.FIRSTSORT" & vbCrLf '提案意願順序
        sql &= " ,cc.GCODE31" & vbCrLf '課程分類編碼
        sql &= " ,cc.PNAME" & vbCrLf '【提案】課程分類/課程分類,ig3.CNAME ,課程分類 2.「【提案】課程分類」欄位：撈取訓練單位研提課程時的「職類課程」資料
        sql &= " ,cc.PFCNAME" & vbCrLf 'PFCNAME/PFGCODE 1.「課程分類」欄位：撈取審查幕僚意見 -「分署確認課程分類」資料

        sql &= " ,CC.THOURS" & vbCrLf '訓練時數
        sql &= " ,CC.TNUM" & vbCrLf '訓練人次
        sql &= " ,CC.TOTAL" & vbCrLf '每人訓練費用(元)
        sql &= " ,CC.TOTALCOST" & vbCrLf '訓練單位可向學員收取之訓練費用(元)
        'sql &= " ,CC.DEFSTDCOST" & vbCrLf
        sql &= " ,CC.DEFGOVCOST" & vbCrLf '總補助費(元)(以訓練費用之80%估算)
        'sql &= " ,CC.DEFSTDCOST1,CC.DEFGOVCOST1" & vbCrLf
        sql &= " ,cc.PROTECHHOURS" & vbCrLf '術科時數

        sql &= " ,cc.FIXSUMCOST" & vbCrLf '固定費用總計
        sql &= " ,cc.ACTHUMCOST" & vbCrLf '固定費用人時成本
        sql &= " ,cc.METSUMCOST" & vbCrLf '材料費總計
        'sql &= " ,cc.METCOSTPER" & vbCrLf '材料費占比
        sql &= " ,CASE WHEN cc.METCOSTPER>=0 THEN concat(convert(float, cc.METCOSTPER),'%') END METCOSTPER" & vbCrLf '/實際材料費比率" & vbCrLf

        sql &= " ,FORMAT(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期
        sql &= " ,FORMAT(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期

        sql &= " ,cc.GCODENAME" & vbCrLf '訓練業別編碼
        sql &= " ,cc.GCNAME" & vbCrLf '訓練業別
        sql &= " ,cc.CODEID" & vbCrLf '訓練職能編碼
        sql &= " ,cc.CCNAME" & vbCrLf '訓練職能
        ''5+2產業,'新南向政策,'台灣AI行動計畫,'數位國家創新經濟發展方案,'國家資通安全發展方案,'前瞻基礎建設計畫,
        sql &= " ,cc.D20KNAME1,cc.D20KNAME6,cc.D20KNAME2,cc.D20KNAME3,cc.D20KNAME4,cc.D20KNAME5" & vbCrLf '5+2產業
        ',亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
        sql &= " ,cc.D25KNAME1,cc.D25KNAME2,cc.D25KNAME3,cc.D25KNAME4,cc.D25KNAME5,cc.D25KNAME6,cc.D25KNAME7,cc.D25KNAME8" & vbCrLf '5+2產業
        sql &= " ,cc.D26KNAME1,cc.D26KNAME2,cc.D26KNAME3,cc.D26KNAME4,cc.D26KNAME5,cc.D26KNAME6,cc.D26KNAME7,cc.D26KNAME8,cc.D26KNAME9"
        sql &= " ,cc.KNAME15" & vbCrLf '轄區重點產業
        sql &= " ,cc.POINTYN" & vbCrLf '是否為學分班(Y/N)
        sql &= " ,cc.CTNAME" & vbCrLf '辦訓縣市別
        sql &= " ,cc.CONTACTNAME" & vbCrLf '聯絡人
        'sql &= " ,cc.CONTACTPHONE" & vbCrLf '聯絡電話
        'sql &= " ,cc.CONTACTMOBILE" & vbCrLf '聯絡電話
        sql &= " ,dbo.FN_PHONEMOBILE(cc.CONTACTPHONE,cc.CONTACTMOBILE) PHONEMOBILE" & vbCrLf '聯絡電話
        sql &= " ,cc.CTNAME2" & vbCrLf '立案縣市
        '審查計分表等級(初)'審查計分表總分(初)
        sql &= " ,cc.IMPLEVEL1 ,cc.IMPSCORE1" & vbCrLf
        '複審等級-通過才顯示(複審)
        sql &= " ,cc.RLEVEL_2" & vbCrLf
        '總分-複審等級-通過才顯示(複審)
        sql &= " ,cc.TOTALSCORE" & vbCrLf

        '匯出欄位「分署幕僚意見」(欄位名稱不變)，內容改抓「送請委員確認內容
        sql &= " ,cc.ST1SUGGEST" & vbCrLf '分署幕僚意見
        sql &= " ,cc.COMMENTS" & vbCrLf '委員審查意見
        sql &= " ,cc.CONFIRMCONT" & vbCrLf '送請委員確認內容
        sql &= " ,cc.ST1RESULT" & vbCrLf '初審建議結論 (1:通過/ 2: 調整後通過/ 3: 不通過) 請寫數字
        sql &= " ,CASE ISNULL(cc.ST1RESULT,cc.RESULT) WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
        sql &= " ,cc.ICAPNUM" & vbCrLf '備註(如有iCap課程請註明，並註明為額度內或額度外）
        sql &= " ,cc.CROSSDIST_YN" & vbCrLf '是否為跨區單位課程 (Y/N)"
        '一階審查結果 Y 通過、N 不通過、P 調整後通過 'sql &= " AND pf.RESULT IN ('Y','P')" & vbCrLf
        sql &= " ,cc.RESULT ,CASE cc.RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END RESULT_C" & vbCrLf
        '核班結果,核班結果
        sql &= " ,cc.CURESULT ,CASE cc.CURESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END CURESULT_C" & vbCrLf
        sql &= " ,cc.NGREASON" & vbCrLf '不通過明細表增加「未核班原因」，說明如下

        sql &= " ,cc.MASTERNAME" & vbCrLf '負責人(產投計畫)
        sql &= " FROM WC1 cc" & vbCrLf
        'ORDER BY 
        sql &= " ORDER BY cc.ORGKIND2,cc.ORGNAME,cc.FIRSTSORT,cc.STDATE" & vbCrLf

        Return sql
    End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt(ByVal iType2 As Integer) As DataTable
        Dim dt As DataTable = Nothing
        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

        Dim parms As New Hashtable

        Dim sql As String = ""

        'iType2'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        Select Case iType2
            Case 1
                'Dim sql_WC1 As String = String.Format("WITH WC1 AS ({0})", GET_CLASS_SQL1(parms))
                'sql = "" & vbCrLf
                'sql &= sql_WC1 'WITH WC1 
                'sql &= " SELECT cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf
                'sql &= " ,COUNT(1) CLASSCNT1" & vbCrLf
                'sql &= " ,SUM(cc.CSTUDCNT14) CSTUDCNT14" & vbCrLf
                'sql &= " ,SUM(cc.DEFGOVCOST) DEFGOVCOST" & vbCrLf
                'sql &= " FROM WC1 cc" & vbCrLf
                ''/*建議結論 Y 通過、N 不通過、P 調整後通過*/
                'sql &= " WHERE cc.RESULT IN ('Y','P')" & vbCrLf
                'sql &= " GROUP BY cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf
            Case 2
                sql = GET_CLASS_SQL2(parms, "Y")
            Case 3
                'Dim sql_WC1 As String = String.Format("WITH WC1 AS ({0})", GET_CLASS_SQL1(parms))
                'sql = "" & vbCrLf
                'sql &= sql_WC1 'WITH WC1 
                'sql &= " SELECT cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf
                'sql &= " ,COUNT(1) CLASSCNT1" & vbCrLf
                'sql &= " ,SUM(cc.CSTUDCNT14) CSTUDCNT14" & vbCrLf
                'sql &= " ,SUM(cc.DEFGOVCOST) DEFGOVCOST" & vbCrLf
                'sql &= " FROM WC1 cc" & vbCrLf
                ''/*建議結論 Y 通過、N 不通過、P 調整後通過*/
                'sql &= " WHERE cc.RESULT IN ('N')" & vbCrLf
                'sql &= " GROUP BY cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf
            Case 4
                sql = GET_CLASS_SQL2(parms, "N")
        End Select

        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function


    '匯出  (請參考附件：表單05_通過課程審查結果彙整總表+明細表.xls、表單05_未通過課程審查結果彙整總表+明細表.xls)
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        '2:通過明細表/4:未通過明細表
        Dim v_RBLExpType2 As String = TIMS.GetListValue(RBLExpType2)
        If v_RBLExpType2 = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return
        End If

        'iType2 'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        '2:通過明細表/4:未通過明細表
        Dim iType2 As Integer = Val(v_RBLExpType2)
        Call EXPORT_5(iType2)
    End Sub

    '(請參考附件：表單05_通過課程審查結果彙整總表+明細表.xls、表單05_未通過課程審查結果彙整總表+明細表.xls)
    Sub EXPORT_5(ByVal iType2 As Integer)
        'Dim dtXls As DataTable = Nothing
        Dim dtXls As DataTable = SEARCH_DATA1_dt(iType2)
        If dtXls Is Nothing Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        If dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        '年度 + 申請階段 + 計畫 + 通過明細表 / 未通過明細表
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_APPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G:產投計畫/W:提升勞工自主學習計畫 「負責人(產投計畫)」名稱修正為「理事長(自主計畫)」
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        If v_APPSTAGE_SCH = "3" Then
            v_rblOrgKind2 = ""
            s_PLANNAME = ""
        End If

        '111年度上半年提升勞工自主學習計畫通過課程彙整表

        'iType2 'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        'Dim sTYPE2_NM As String = If(iType2 = 1, "通過彙整總表", If(iType2 = 2, "通過明細表", If(iType2 = 3, "未通過彙整總表", If(iType2 = 4, "未通過明細表", "未通過明細表-"))))
        Dim sTYPE2_NM As String = If(iType2 = 2, "通過明細表", If(iType2 = 4, "未通過明細表", "未通過明細表-"))
        Dim sTitle1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, sTYPE2_NM)
        Dim v_OUTTYPE2_YN As String = If(iType2 = 2, "Y", If(iType2 = 4, "N", "N"))

        Dim s_DISTID_NM As String = TIMS.GetListText(ddlDISTID_SCH)

        '匯出excel /ods
        Dim s_FILENAME1 As String = String.Concat(sTYPE2_NM, "-", s_DISTID_NM, "_", TIMS.GetDateNo2(3))

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        'Dim in_parms As New Hashtable 'in_parms.Clear()
        'in_parms.Add("EXP", "Y") '匯出查詢條件
        'in_parms.Add("ExpType", v_ExpType) '匯出查詢條件

        Dim parms As New Hashtable From {
            {"ExpType", v_ExpType}, 'EXCEL/PDF/ODS
            {"FileName", s_FILENAME1},
            {"TitleName", TIMS.ClearSQM(sTitle1)},
            {"ORGKIND2", v_rblOrgKind2},
            {"OUTTYPE2_YN", v_OUTTYPE2_YN},
            {"APPSTAGE_SCH", v_APPSTAGE_SCH}
        } 'parms.Clear()

        Call EXPORT_5_24(dtXls, parms)
    End Sub

    '2:通過明細表 /4:未通過明細表
    Sub EXPORT_5_24(ByRef dtXls As DataTable, ByRef parms As Hashtable)
        'Dim iCLASSCNT1 As Integer = 0 '訓練班次'CLASSCNT1
        'Dim iCSTUDCNT14 As Integer = 0  '訓練人次'CSTUDCNT14
        'Dim iDEFGOVCOST As Integer = 0  '訓練補助費'DEFGOVCOST

        'For Each dr1 As DataRow In dtXls.Rows
        '    iCLASSCNT1 += If(Val(dr1("CLASSCNT1")) > 0, Val(dr1("CLASSCNT1")), 0)
        '    iCSTUDCNT14 += If(Val(dr1("CSTUDCNT14")) > 0, Val(dr1("CSTUDCNT14")), 0)
        '    iDEFGOVCOST += If(Val(dr1("DEFGOVCOST")) > 0, Val(dr1("DEFGOVCOST")), 0)
        'Next
        Dim v_APPSTAGE_SCH As String = TIMS.GetMyValue2(parms, "APPSTAGE_SCH")
        Dim v_OUTTYPE2_YN As String = TIMS.GetMyValue2(parms, "OUTTYPE2_YN")
        'G:產投計畫/W:提升勞工自主學習計畫 「負責人(產投計畫)」名稱修正為「理事長(自主計畫)」
        Dim v_ORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")
        Dim s_MASTERNAME_t As String = String.Concat(",", If(v_ORGKIND2 = "", "負責人/理事長", If(v_ORGKIND2 = "G", "負責人(產投計畫)", If(v_ORGKIND2 = "W", "理事長(自主計畫)", v_ORGKIND2))))

        Dim sb_Pattern As New StringBuilder
        'sb_Pattern.Append(If(v_APPSTAGE_SCH = "3", "序號,計畫別,分署別", "序號"))
        sb_Pattern.Append("序號,計畫別,分署別")
        sb_Pattern.Append(",訓練單位名稱,統一編號,單位屬性,課程名稱(含期別),課程申請流水號,提案意願順序,課程分類編碼,課程分類,訓練時數,訓練人次")
        sb_Pattern.Append(",每人訓練費用(元),訓練單位可向學員收取之訓練費用(元),總補助費(元)(以訓練費用之80%估算),術科時數,固定費用總計,固定費用人時成本,材料費總計,材料費占比")
        sb_Pattern.Append(",開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能")
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            '2026 政府政策性產業 (產投)
            sb_Pattern.Append(",五大信賴產業推動方案,六大區域產業及生活圈,台灣AI行動計畫2.0,智慧國家2.0鋼領,國家人才競爭力躍升方案,新南向政策推動計畫,AI新十大建設推動方案,智慧機器人產業推動方案,臺灣2050淨零轉型")
        Else
            'sb_Column.Append(",D20KNAME1,D20KNAME6,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5")
            'sb_Pattern.Append(",5+2產業,新南向政策,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫")
            sb_Pattern.Append(",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航")
        End If
        sb_Pattern.Append(",轄區重點產業,是否為學分班(Y/N),辦訓縣市別,聯絡人,聯絡電話,立案縣市")
        sb_Pattern.Append(",審查計分表等級,審查計分表總分,分署幕僚意見,委員審查意見")
        sb_Pattern.Append(",第1階段實質審查結果(1:通過/2:調整後通過/3:不通過)請寫數字")
        sb_Pattern.Append(If(v_OUTTYPE2_YN = "N", ",未核班原因", ""))
        sb_Pattern.Append(",備註(如有iCap課程請註明，並註明為額度內或額度外）,是否為跨區單位課程(Y/N)")
        sb_Pattern.Append(String.Concat(s_MASTERNAME_t, ",【提案】課程分類"))
        'sb_Pattern.Append(If(v_APPSTAGE_SCH = "3", "審查計分表等級,審查計分表總分", ""))

        '機構別／,(單位屬性) ORGTYPENAME2/,(產投)單位屬性 TYPEID2NAME
        Dim sb_Column As New StringBuilder
        'sb_Column.Append(If(v_APPSTAGE_SCH = "3", "SEQNUM,ORGPLANNAME2,DISTNAME3", "SEQNUM"))
        sb_Column.Append("SEQNUM,ORGPLANNAME2,DISTNAME3")
        sb_Column.Append(",ORGNAME,COMIDNO,ORGTYPENAME2,CLASSCNAME2,PSNO28,FIRSTSORT,GCODE31,PFCNAME,THOURS,TNUM")
        sb_Column.Append(",TOTAL,TOTALCOST,DEFGOVCOST,PROTECHHOURS,FIXSUMCOST,ACTHUMCOST,METSUMCOST,METCOSTPER")
        sb_Column.Append(",STDATE,FTDATE,GCODENAME,GCNAME,CODEID,CCNAME")
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            '2026 政府政策性產業 (產投)
            sb_Column.Append(",D26KNAME1,D26KNAME2,D26KNAME7,D26KNAME3,D26KNAME5,D26KNAME4,D26KNAME6,D26KNAME8,D26KNAME9")
        Else
            'sb_Column.Append(",D20KNAME1,D20KNAME6,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5")
            sb_Column.Append(",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,D25KNAME7,D25KNAME8")
        End If
        sb_Column.Append(",KNAME15,POINTYN,CTNAME,CONTACTNAME,PHONEMOBILE,CTNAME2")
        'sb_Column.Append(",IMPLEVEL1,IMPSCORE1,RLEVEL_2,TOTALSCORE,ST1SUGGEST,COMMENTS")
        sb_Column.Append(",RLEVEL_2,TOTALSCORE,ST1SUGGEST,COMMENTS")

        sb_Column.Append(",RESULT_C")  '第1階段實質審查結果(1:通過/2:調整後通過/3:不通過)請寫數字/RESULT_C 'sb_Column.Append(",ST1RESULT_C") '初審建議結論 (1:通過/ 2: 調整後通過/ 3: 不通過) 請寫數字
        sb_Column.Append(If(v_OUTTYPE2_YN = "N", ",NGREASON", ""))
        sb_Column.Append(",ICAPNUM,CROSSDIST_YN")
        sb_Column.Append(",MASTERNAME,PNAME")
        'sb_Pattern.Append(If(v_APPSTAGE_SCH = "3", "RLEVEL_2,TOTALSCORE", ""))

        Dim sPatternA() As String = Split(sb_Pattern.ToString(), ",")
        Dim sColumnA() As String = Split(sb_Column.ToString(), ",")
        'Dim iColSpanCount As Integer = 46
        Dim iColSpanCount As Integer = sColumnA.Length 'If(sColumnA.Length > sPatternA.Length, sColumnA.Length, sPatternA.Length)

        'Dim s_FootHtml2 As String = ""
        's_FootHtml2 &= "<tr>"
        's_FootHtml2 &= String.Format("<td colspan=2>{0}</td>", "合計") '合計
        's_FootHtml2 &= String.Format("<td>{0}</td>", iCLASSCNT1) '訓練班次
        's_FootHtml2 &= String.Format("<td>{0}</td>", iCSTUDCNT14) '訓練人次
        's_FootHtml2 &= String.Format("<td>{0}</td>", iDEFGOVCOST) '總補助費(元)〈以訓練費用之80%估算〉
        's_FootHtml2 &= "</tr>"

        'parms.Add("TitleHtml2", s_TitleHtml2) 'parms.Add("FootHtml2", s_FootHtml2)
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

End Class

