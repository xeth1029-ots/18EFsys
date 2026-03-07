Partial Class SD_15_022
    Inherits AuthBasePage

    'https://jira.turbotech.com.tw/browse/TIMSC-96
    'https://jira.turbotech.com.tw/browse/TIMSC-296
    'Const cst_warning1 As String = "該計畫不提供此功能!"
    '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
    Dim fg_Work2026x02 As Boolean = False 'TIMS.SHOW_W2026x02(sm)

    'Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線

        If TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg17)
            Return 'Exit Sub
        End If

        If Not IsPostBack Then
            CCreate1()
        End If

        Select Case sm.UserInfo.LID
            Case 2 '委訓
                Button2.Visible = False
            Case Else
                If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
                    '署(局) 或 分署(中心)
                    TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                    If HistoryRID.Rows.Count <> 0 Then
                        center.Attributes("onclick") = "showObj('HistoryList2');"
                        center.Style("CURSOR") = "hand"
                    End If
                End If
                Button2.Attributes("onclick") = TIMS.Get_javascript_openOrg_js(sm)
        End Select

    End Sub

    Sub CCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If tr_AppStage_TP28.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
        End If
        tr_CrossDist_TP28.Visible = If(sm.UserInfo.TPlanID = TIMS.Cst_TPlanID28, True, False)
        'tr_CrossDist_TP28.Visible = tr_AppStage_TP28.Visible
        If tr_CrossDist_TP28.Visible Then
            '跨區/轄區提案 //D:不區分/C:跨區提案單位/J:轄區提案單位 (選擇 跨區提案單位，排除【訓練機構】條件)
            RBL_CrossDist_SCH.Attributes("onclick") = "javascript:return CHK_RBL_CROSSDIST_SCH();"
        End If

        '(V_DEPOT12)
        '課程分類 'KID12 'SELECT * FROM KEY_BUSINESS A WHERE 1=1 AND A.DEPID='12'
        HidcblDepot12.Value = "0"
        cblDepot12 = TIMS.Get_KeyBusiness(cblDepot12, "12", objconn) '課程分類
        cblDepot12.Attributes("onclick") = "SelectAll('cblDepot12','HidcblDepot12');"
    End Sub

    Function checkData1(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True 'False
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)

        If STDate1.Text <> "" AndAlso Not TIMS.IsDate1(STDate1.Text) Then errMsg &= "開訓期間 起始日期有誤!" & vbCrLf
        If STDate2.Text <> "" AndAlso Not TIMS.IsDate1(STDate2.Text) Then errMsg &= "開訓期間 結束日期有誤!" & vbCrLf
        If tr_AppStage_TP28.Visible Then
            Dim v_AppStage As String = TIMS.GetListValue(AppStage)
            If v_AppStage = "" Then errMsg &= "請選擇 申請階段!" & vbCrLf
        End If

        If errMsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>查詢 (SQL) 語法匯出 tr_CrossDist_TP28-區分(產投／充飛)</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function Search_SQL_2854(ByRef parms As Hashtable) As String
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = Convert.ToString(sm.UserInfo.RID)
        'Dim s_RID1 As String = Mid(RIDValue.Value, 1, 1)
        '(選擇 跨區提案單位，排除【訓練機構】條件)
        Dim v_RBL_CrossDist_SCH As String = ""
        Dim s_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        If tr_CrossDist_TP28.Visible Then
            'RBL_CrossDist_SCH 跨區/轄區提案 D:不區分 C:跨區提案單位 J:轄區提案單位
            ',dbo.FN_GET_CROSSDIST(ip.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST
            '跨區/轄區提案 //D:不區分/C:跨區提案單位/J:轄區提案單位
            v_RBL_CrossDist_SCH = TIMS.GetListValue(RBL_CrossDist_SCH)

            If s_DistID = "" AndAlso v_RBL_CrossDist_SCH <> "C" Then Return ""
            If v_RBL_CrossDist_SCH = "C" Then s_DistID = ""
            If v_RBL_CrossDist_SCH = "C" Then RIDValue.Value = ""
        End If

        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        fg_Work2026x02 = TIMS.SHOW_W2026x02(sm)
        Dim v_IR_AppStage As Integer = TIMS.CINT1(TIMS.GetListValue(AppStage))
        If (v_IR_AppStage = 0) Then v_IR_AppStage = 3 'NULL(強制轉為政策性)
        Dim fg_USE_trKID25 As Boolean = $"{sm.UserInfo.Years}.{v_IR_AppStage}" <= "2026.1" '(2026上半年)強制使用trKID25 或有值
        'If TIMS.sUtl_ChkTest() Then fg_USE_trKID25 = $"{sm.UserInfo.Years}.{v_IR_AppStage}" <= "2025.1" '(TEST)

        Dim sql_WS1 As String = " ,dd.D25KNAME1,dd.D25KNAME2,dd.D25KNAME3,dd.D25KNAME4,dd.D25KNAME5,dd.D25KNAME6,dd.D25KNAME7,dd.D25KNAME8,dd.KID22,dd.KNAME22"
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            If Not fg_USE_trKID25 Then
                '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
                '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
                sql_WS1 = " ,dd.D26KNAME1,dd.D26KNAME2,dd.D26KNAME3,dd.D26KNAME4,dd.D26KNAME5,dd.D26KNAME6,dd.D26KNAME7,dd.D26KNAME8,dd.D26KNAME9"
            End If
        End If

        Dim sql As String = ""
        '跨區/轄區提案
        If tr_CrossDist_TP28.Visible Then
            sql &= " WITH WR1 AS ( SELECT rr.RID,rr.PLANID,rr.ORGLEVEL,rr.OrgKind2,rr.ORGPLANNAME2" & vbCrLf
            sql &= " ,dbo.FN_GET_CROSSDIST(rr.YEARS,rr.COMIDNO,@AppStage) i_CROSSDIST" & vbCrLf
            sql &= " FROM dbo.VIEW_RIDNAME rr" & vbCrLf
            sql &= " WHERE rr.ORGLEVEL=2 AND rr.YEARS=@Years )" & vbCrLf
        End If

        sql &= " SELECT convert(int,'0') SEQNUM" & vbCrLf
        sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        '訓練單位名稱" & vbCrLf
        sql &= " ,oo.ORGNAME" & vbCrLf
        'sql &= " ,oo.COMIDNO" & vbCrLf '   --'統一編號" & vbCrLf
        '單位屬性" & vbCrLf
        sql &= " ,concat(o1.typeid2,'-',o1.typeid2name) OrgTypeName" & vbCrLf
        '分署別" & vbCrLf
        sql &= " ,ip.DISTNAME" & vbCrLf
        '課程名稱
        sql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) CLASSNAME" & vbCrLf
        '提案意願順序" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf
        '課程申請流水號" & vbCrLf
        sql &= " ,pp.PSNO28" & vbCrLf
        'sql &= " ,cc.OCID" & vbCrLf '課程申請流水號" & vbCrLf
        '2019年啟用'申請階段
        sql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.APPRESULT,'Y',dd.KID12,ISNULL(dd.KID12,vd12.KID)) KID12" & vbCrLf ' '課程分類編碼" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.APPRESULT,'Y',dd.KNAME12,ISNULL(dd.KNAME12,vd12.kname)) KNAME12" & vbCrLf '課程分類" & vbCrLf
        '訓練時數" & vbCrLf
        sql &= " ,pp.THOURS" & vbCrLf
        '訓練人次" & vbCrLf
        sql &= " ,pp.TNUM" & vbCrLf
        '每人訓練費用(元)" & vbCrLf
        sql &= " ,format(CASE WHEN pp.TNum IS NULL THEN '0' ELSE (ISNULL(pp.Defstdcost,0) + ISNULL(pp.DefGovCost,0))/pp.TNum END,'N') Total" & vbCrLf
        '每班總訓練費(元)" & vbCrLf
        sql &= " ,format(ISNULL(pp.TotalCost,0),'N') TotalCost" & vbCrLf
        '每班總補助費(元)" & vbCrLf
        sql &= " ,format(ISNULL(pp.DefGovCost,0),'N') DefGovCost" & vbCrLf
        'ProTechHours 
        'sql &= " ,ISNULL(PP.ProTechHours,0) ProTechHours" & vbCrLf '術科時數" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC(pp.PlanID,pp.COMIDNO,pp.SEQNO,2) PROTECHHOURS" & vbCrLf '術科時數" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_SHEETCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,1) SHEETCOST" & vbCrLf '教材費明細" & vbCrLf
        'sql &= " ,dbo.fn_GET_PLAN_SHEETCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,2) SHEETCOST2" & vbCrLf '教材費明細" & vbCrLf

        sql &= " ,dbo.FN_GET_PLAN_PERSONCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,1) PERSONCOST" & vbCrLf '教材費明細" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_COMMONCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,1) COMMONCOST" & vbCrLf '教材費明細" & vbCrLf
        'sql &= " ,dbo.fn_GET_PLAN_PERSONCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,2) PERSONCOST2" & vbCrLf
        'sql &= " ,dbo.fn_GET_PLAN_COMMONCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,2) COMMONCOST2" & vbCrLf ' --材料費明細" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_OTHERCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,1) OTHERCOST" & vbCrLf '其他費用明細" & vbCrLf
        'sql &= " ,dbo.fn_GET_PLAN_OTHERCOST(pp.PlanID,pp.COMIDNO,pp.SEQNO,2) OTHERCOST2" & vbCrLf ' --其他費用明細" & vbCrLf
        '03,人 教材費-(人)04,人	材料費-(人)11,班	其他費用
        'sql &= " ,ISNULL(cc2.PT03,0) PT03" & vbCrLf
        'sql &= " ,ISNULL(cc2.PT04,0) PT04" & vbCrLf
        'sql &= " ,ISNULL(cc2.PT11,0) PT11" & vbCrLf
        '人時成本上限" & vbCrLf
        sql &= " ,ig2.MAXUP" & vbCrLf
        '實際人時成本" & vbCrLf
        sql &= " ,format(CASE WHEN pp.TNum IS NULL THEN 0 else CASE WHEN pp.Thours IS NULL THEN 0 else isnull(pp.TotalCost,0)/PP.TNum/pp.Thours end end,'N') TIMECOST" & vbCrLf
        '開訓日期" & vbCrL
        sql &= " ,CONVERT(varchar, pp.STDATE, 111) STDATE" & vbCrLf
        '結訓日期" & vbCrLf
        sql &= " ,CONVERT(varchar, pp.FDDATE, 111) FDDATE" & vbCrLf
        'sql &= " ,CONVERT(varchar, cc.STDate, 111) STDATE" & vbCrLf '開訓日期" & vbCrLf
        'sql &= " ,CONVERT(varchar, cc.FTDATE, 111) FTDATE" & vbCrLf '結訓日期" & vbCrLf

        'sql &= " ,tt.JOBID" & vbCrLf '職訓業別編碼" & vbCrLf
        'sql &= " ,'['+tt.JOBID+']'+tt.JOBNAME TJOBNAME" & vbCrLf '職訓業別" & vbCrLf

        'GCODENAME:「業別分類代碼」: 提案彙總表 -【訓練業別編碼】欄位 
        If sm.UserInfo.Years >= 2018 Then
            sql &= " ,ISNULL(ig3.GCODE2,ig2.GCODE2) GCodeName" & vbCrLf '訓練業別編碼" & vbCrLf
            sql &= " ,ISNULL(ig3.CNAME,ig2.CNAME) GCNAME" & vbCrLf '訓練業別編碼" & vbCrLf
        Else
            sql &= " ,ISNULL(ig.GOVCLASSN,ig2.GCODE2) GCodeName" & vbCrLf '訓練業別編碼" & vbCrLf
            sql &= " ,ISNULL(ig.CNAME,ig2.CNAME) GCNAME" & vbCrLf '訓練業別編碼" & vbCrLf
        End If
        '訓練職能編碼" & vbCrLf
        sql &= " ,kc.CODEID" & vbCrLf
        '訓練職能" & vbCrLf
        sql &= " ,kc.CCName" & vbCrLf
        'KID17
        'sql &= " ,dd.KNAME17 kname52" & vbCrLf '----dd.kname13-'5+2產業" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME18 ,NULL) KNAME18" & vbCrLf '新南向政策" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME06 ,ISNULL(dd.KNAME06 ,vd06.kname)) kname1" & vbCrLf '新興產業" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.Kname10 ,d10.Kname) Kname10" & vbCrLf '重點服務業" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',K10.KNAME ,NULL) D10KNAME " & vbCrLf '10大重點服務業(9項)
        'sql &= " ,dd.kname13" & vbCrLf '轄區重點產業

        'sql &= " ,dbo.DECODE(pp.PointYN,'Y','Y',NULL) PointYN '是否為學分班(Y/N)" & vbCrLf
        sql &= " ,pp.PointYN" & vbCrLf
        sql &= " ,case pp.PointYN when 'Y' then '是' else '否' end PointYN_N" & vbCrLf
        '辦訓縣市別
        sql &= " ,COALESCE(iz.CTName,vtp.s1CTName,vtp.t1CTName) CTName" & vbCrLf
        'TC_01_017_add.aspx?orgid=1523
        '立案縣市別
        sql &= " ,iz3.CTName OrgCTName" & vbCrLf
        '「聯絡電話」辦公室電話及行動電話
        sql &= " ,pp.CONTACTNAME,pp.CONTACTPHONE,pp.CONTACTMOBILE" & vbCrLf
        sql &= " ,CASE WHEN LEN(pp.CONTACTPHONE)>1 AND LEN(pp.CONTACTMOBILE)>1 THEN CONCAT(pp.CONTACTPHONE,'、',pp.CONTACTMOBILE)" & vbCrLf
        sql &= " WHEN LEN(pp.CONTACTPHONE)>1 THEN pp.CONTACTPHONE" & vbCrLf
        sql &= " WHEN LEN(pp.CONTACTMOBILE)>1 THEN pp.CONTACTMOBILE END CONTACTPHONEMOB" & vbCrLf
        '辦理方式
        sql &= " ,pp.DISTANCE,dbo.FN_DISTANCE_N(pp.DISTANCE) DISTANCE_N" & vbCrLf
        '課程分類編碼
        sql &= " ,ISNULL(dd.KID12,ig3.GCODE31) KID12" & vbCrLf
        '轄區重點產業
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.SEQNOD15 ,NULL) SEQNOD15" & vbCrLf
        '六大新興產業
        'sql &= " ,dd.KID06" & vbCrLf
        '重點服務業
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID10 ,NULL) KID10" & vbCrLf
        '政府政策性產業(5+2)
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID17 ,NULL) KID17" & vbCrLf
        '政府政策性產業(5+2)-2018
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID19 ,NULL) KID19" & vbCrLf
        '新南向政策
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID18 ,NULL) KID18" & vbCrLf
        '課程分類 (view_depot12 vd12)
        sql &= " ,ISNULL(dd.KNAME12,ig3.PNAME) D12KNAME" & vbCrLf
        'ISNULL(dbo.DECODE(dd.APPRESULT,'Y',dd.kname12,vd12.kname),ig3.PNAME) KNAME12
        '轄區重點產業
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME15 ,NULL) D15KNAME" & vbCrLf
        '6大新興產業
        'sql &= " ,dd.KNAME06 D06KNAME" & vbCrLf
        '10大重點服務業(9項)
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME10 ,NULL) D10KNAME" & vbCrLf
        '政府政策性產業(5+2)
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME17 ,NULL) D17KNAME" & vbCrLf
        '新南向政策
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME18 ,NULL) D18KNAME" & vbCrLf
        'V_PLAN_DEPOT
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KNAME19 ,NULL) D19KNAME" & vbCrLf
        '2019年啟用 work2019x01:2019 政府政策性產業
        sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID20 ,NULL) KID20" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.D20KNAME1 ,'無') D20KNAME1" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.D20KNAME2 ,'無') D20KNAME2" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.D20KNAME3 ,'無') D20KNAME3" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.D20KNAME4 ,'無') D20KNAME4" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.D20KNAME5 ,'無') D20KNAME5" & vbCrLf
        'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.D20KNAME6 ,'無') D20KNAME6" & vbCrLf

        '亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航
        '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
        '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
        sql &= sql_WS1

        sql &= " ,pp.FIXSUMCOST ,pp.ACTHUMCOST" & vbCrLf
        'dbo.fn_GET_PLANCNAME(PLANID,COMIDNO,SEQNO)
        sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'P') METDETP" & vbCrLf
        sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'C') METDETC" & vbCrLf
        sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'S') METDETS" & vbCrLf
        sql &= " ,dbo.FN_GET_PLANCNAME(pp.PLANID,pp.COMIDNO,pp.SEQNO,'O') METDETO" & vbCrLf
        sql &= " ,convert(varchar(max),'') METDET" & vbCrLf
        sql &= " ,pp.METSUMCOST" & vbCrLf
        sql &= " ,CASE WHEN pp.METCOSTPER IS NOT NULL THEN CONVERT(varchar, pp.METCOSTPER)+'%' END METCOSTPER" & vbCrLf
        sql &= " ,ISNULL(pp.FIXSUMCOST,0)+ISNULL(pp.METSUMCOST,0) ALLSUMCOST" & vbCrLf
        'sql &= " ,dbo.FN_GET_ROC_YEAR(CONVERT(int,vp.YEARS)+0) YR1" & vbCrLf
        'sql &= " ,dbo.FN_GET_ROC_YEAR(CONVERT(int,vp.YEARS)+1) YR2" & vbCrLf
        'sql &= " ,dbo.FN_GET_ROC_YEAR(CONVERT(int,vp.YEARS)+2) YR3" & vbCrLf
        'sql &= " ,dbo.FN_GET_PRECLASS_PCNT1(pp.PLANID,pp.COMIDNO,pp.SEQNO,ip.YEARS,1) PCNT11" & vbCrLf
        'sql &= " ,dbo.FN_GET_PRECLASS_PCNT1(pp.PLANID,pp.COMIDNO,pp.SEQNO,ip.YEARS,2) PCNT12" & vbCrLf
        'sql &= " ,dbo.FN_GET_PRECLASS_PCNT1(pp.PLANID,pp.COMIDNO,pp.SEQNO,ip.YEARS,3) PCNT13" & vbCrLf
        '跨區/轄區提案
        If tr_CrossDist_TP28.Visible Then
            sql &= " ,wr.ORGPLANNAME2" & vbCrLf
            '跨區/轄區提案 CROSSDIST_N
            sql &= " ,wr.i_CROSSDIST,case wr.i_CROSSDIST when -1 then '轄區提案' else '跨區提案' end CROSSDIST_N" & vbCrLf
        Else
            '計畫別
            sql &= " ,vr.ORGPLANNAME2" & vbCrLf
        End If
        'ICAPNUM-iCAP標章證號
        sql &= " ,pp.ICAPNUM" & vbCrLf
        '是否為iCAP課程
        sql &= " ,CASE pvr.ISiCAPCOUR WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END ISICAPCOUR_N" & vbCrLf
        'iCAP有效期限 班級申請 iCAP有效期限
        sql &= " ,FORMAT(pp.iCAPMARKDATE,'yyyy/MM/dd') iCAPMARKDATE" & vbCrLf
        '實體課程時數 PHOUR 訓練總時數減掉有勾選遠距教學之時數 分署要檢核該課程實體課程是否有佔總訓練課程的1/3以上
        sql &= " ,dbo.FN_GET_PLAN_TRAINDESC(pp.PLANID,pp.COMIDNO,pp.SEQNO,6) ENTPHOUR" & vbCrLf
        '課程內容有室外教學 室外教學課程
        sql &= " ,pp.OUTDOOR" & vbCrLf
        sql &= " ,case pp.OUTDOOR when 'Y' then '是' else '否' end OUTDOOR_N" & vbCrLf
        '報請主管機關核備 是否報請主管機關核備
        sql &= " ,pvr.REPORTE" & vbCrLf
        sql &= " ,case pvr.REPORTE when 'Y' then '是' else '否' end REPORTE_N" & vbCrLf
        '未檢送資料
        sql &= " ,pp.DataNotSent" & vbCrLf
        sql &= " ,case pp.DataNotSent when 'Y' then '是' else '否' end DataNotSent_N" & vbCrLf
        'TMIDCORRECT:若訓練業別有誤是否同意協助重新歸類
        sql &= " ,pp.TMIDCORRECT" & vbCrLf
        '線上送件。倘該班級係透過線上申辦送件並有送出至分署(【申辦狀態】非暫存)，則於此欄位顯示Y
        sql &= " ,dbo.FN_GET_BIDCASEPI(pp.PlanID,pp.ComIDNO,pp.SeqNo,'Y') BIDCASEPI" & vbCrLf

        sql &= " FROM dbo.PLAN_PLANINFO pp WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc WITH(NOLOCK) on pp.ClassCate=kc.CCID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.planid=pp.planid" & vbCrLf
        '跨區/轄區提案
        If tr_CrossDist_TP28.Visible Then
            sql &= " JOIN WR1 wr on wr.RID=pp.RID" & vbCrLf
        Else
            sql &= " JOIN dbo.VIEW_RIDNAME vr on vr.RID=pp.RID" & vbCrLf
        End If
        sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) on oo.ComIDNO=pp.ComIDNO" & vbCrLf
        '機構別鍵詞檔
        'sql &= " JOIN KEY_ORGTYPE ko WITH(NOLOCK) on ko.ORGTYPEID=oo.ORGKIND" & vbCrLf
        '若為學分班(沒有 VIEW_COSTITEM2 訓練費用明細)
        'sql &= " LEFT JOIN VIEW_COSTITEM2 cc2 ON cc2.PlanID=pp.PlanID AND cc2.ComIDNO=pp.ComIDNO AND cc2.SeqNO=pp.SeqNO" & vbCrLf
        '機構別鍵詞檔1
        sql &= " LEFT JOIN dbo.KEY_ORGTYPE1 o1 on o1.OrgTypeID1=oo.OrgKind1" & vbCrLf
        '開班計畫
        sql &= " LEFT JOIN dbo.PLAN_VERREPORT pvr WITH(NOLOCK) ON pp.PlanID=pvr.PlanID AND pp.ComIDNO=pvr.ComIDNO AND pp.SeqNo=pvr.SeqNo" & vbCrLf
        '訓練業別
        sql &= " LEFT JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=pp.TMID" & vbCrLf

        sql &= " LEFT JOIN dbo.VIEW_GOVCLASSCAST ig on pp.GCID=ig.GCID" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST2 ig2 on pp.GCID2=ig2.GCID2" & vbCrLf '依GCID2
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 on pp.GCID3=ig3.GCID3" & vbCrLf '依GCID3
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz ON iz.ZipCode = pp.TaddressZip" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME iz3 on iz3.zipCode=oo.orgzipcode" & vbCrLf
        '產投上課地址學科場地代碼 ／產投上課地址術科場地代碼
        sql &= " LEFT JOIN dbo.VIEW_TRAINPLACE vtp on vtp.PlanID=pp.PlanID and vtp.ComIDNO=pp.ComIDNO and vtp.SeqNo=pp.SeqNo" & vbCrLf
        '課程分類 '訓練課程分類 KID12
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PlanID=pp.PlanID and dd.ComIDNO=pp.ComIDNO and dd.SeqNo=pp.SeqNo" & vbCrLf
        sql &= " WHERE pp.ISAPPRPAPER ='Y' AND pp.RESULTBUTTON IS NULL" & vbCrLf '審核送出(已送審)
        sql &= " AND pvr.ISAPPRPAPER='Y'" & vbCrLf '正式 'sql &= " AND (cc.IsSuccess = 'Y') AND (cc.NotOpen = 'N')" & vbCrLf
        sql &= " AND ip.TPlanID=@TPlanID AND ip.Years=@Years " & vbCrLf

        '含未檢送研提資料
        '【提案匯總表】匯出時請排除，不匯出有勾選【未檢送資料】之班級。(預設)
        'If Not CB_DataNotSent_SCH.Checked Then sql &= " and pp.DataNotSent IS NULL" & vbCrLf

        'RBL_DataNotSent_SCH 研提資料 Y:有檢送 / O:含未檢送研提資料
        Dim v_RBL_DataNotSent_SCH As String = TIMS.GetListValue(RBL_DataNotSent_SCH)
        Select Case v_RBL_DataNotSent_SCH
            Case "Y" 'Y:有檢送
                sql &= " and pp.DataNotSent IS NULL" & vbCrLf
        End Select

        '跨區/轄區提案
        If tr_CrossDist_TP28.Visible Then
            'RBL_CrossDist_SCH 跨區/轄區提案 D:不區分 C:跨區提案單位 J:轄區提案單位
            ',dbo.FN_GET_CROSSDIST(ip.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST
            '跨區/轄區提案 //D:不區分/C:跨區提案單位/J:轄區提案單位
            'Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
            'Select Case v_RBL_CrossDist_SCH
            '    Case "C" 'C:跨區提案單位
            '        sql &= " and dbo.FN_GET_CROSSDIST(ip.YEARS,pp.COMIDNO,pp.APPSTAGE)!=-1" & vbCrLf
            '    Case "J" 'J:轄區提案單位
            '        sql &= " and dbo.FN_GET_CROSSDIST(ip.YEARS,pp.COMIDNO,pp.APPSTAGE)=-1 " & vbCrLf
            'End Select

            'wr.i_CROSSDIST 
            Select Case v_RBL_CrossDist_SCH
                Case "C" 'C:跨區提案單位
                    sql &= " and wr.i_CROSSDIST !=-1" & vbCrLf
                Case "J" 'J:轄區提案單位
                    sql &= " and wr.i_CROSSDIST=-1 " & vbCrLf
            End Select
        End If

        '課程分類 '訓練課程分類 KID12
        Dim cblDepot12_ValIn As String = TIMS.GetCblValueIn(cblDepot12)
        If cblDepot12_ValIn <> "" Then sql &= String.Concat(" AND ISNULL(dd.KID12,ig3.GCODE31) IN (", cblDepot12_ValIn, ")", vbCrLf)

        'test 測試環境測試
        'Dim flag_chktest As Boolean = If(TIMS.sUtl_ChkTest(), True, False) '(測試環境中)
        'If (flag_chktest) Then TIMS.writeLog(Me, String.Concat("##SD_15_022.aspx,", vbCrLf, ",Search_SQL_2854 sql:", vbCrLf, sql))

        parms.Clear()
        parms.Add("TPlanID", sm.UserInfo.TPlanID) '登入計畫
        parms.Add("Years", CStr(sm.UserInfo.Years)) '登入年度
        Select Case sm.UserInfo.LID
            Case 0
                If s_DistID <> "" Then
                    sql &= " and ip.DistID=@DistID " & vbCrLf
                    parms.Add("DistID", s_DistID) '依業務選擇轄區
                End If
            Case 1
                If s_DistID <> "" Then
                    sql &= " and ip.DistID=@DistID " & vbCrLf
                    parms.Add("DistID", s_DistID) '依業務選擇轄區
                End If
            Case Else
                sql &= " and ip.DistID=@DistID " & vbCrLf
                parms.Add("DistID", sm.UserInfo.DistID) '登入轄區
        End Select
        If TRPlanPoint28.Visible Then
            Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
            Select Case v_rblOrgKind2'rblOrgKind2.SelectedValue
                Case "G", "W"
                    'sql &= " and vr.OrgKind2=@OrgKind2 " & vbCrLf
                    sql &= " and wr.OrgKind2=@OrgKind2 " & vbCrLf
                    parms.Add("OrgKind2", rblOrgKind2.SelectedValue)
                Case "A"
                Case Else
                    sql &= " and 1<>1 " & vbCrLf
            End Select
        End If

        'Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        'RID LEN !=1 有選取某機構
        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length <> 1 Then
            sql &= " and pp.RID = @RID " & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If

        If STDate1.Text <> "" Then
            sql &= " and pp.STDATE >= @STDATE1 " & vbCrLf
            parms.Add("STDATE1", STDate1.Text)
        End If

        If STDate2.Text <> "" Then
            sql &= " and pp.STDATE <= @STDATE2 " & vbCrLf
            parms.Add("STDATE2", STDate2.Text)
        End If

        Dim v_AppStage As String = "" 'TIMS.GetListValue(AppStage)
        If tr_AppStage_TP28.Visible Then
            v_AppStage = TIMS.GetListValue(AppStage)
            'Dim v_AppStage As String = TIMS.ClearSQM(AppStage.SelectedValue)
            'If v_AppStage <> "" Then End If
            sql &= " AND pp.AppStage=@AppStage " & vbCrLf '依申請階段
            parms.Add("AppStage", v_AppStage)
        End If

        sql &= " ORDER BY oo.ORGNAME,ip.DISTID,pp.COMIDNO,pp.FIRSTSORT,pp.STDATE" & vbCrLf
        'Select Case sm.UserInfo.Years
        '    Case Is >= 2019
        '        sql &= " ORDER BY pp.COMIDNO,pp.FIRSTSORT" & vbCrLf
        '    Case Else
        '        sql &= " ORDER BY pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        'End Select

        Return sql
    End Function

    ''' <summary> 匯出欄位設定 區分(產投／充飛) </summary>
    ''' <param name="sPattern"></param>
    ''' <param name="sColumn"></param>
    Sub UTL_PCSTRING(ByRef sPattern As String, ByRef sColumn As String)

        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        fg_Work2026x02 = TIMS.SHOW_W2026x02(sm)
        Dim v_IR_AppStage As Integer = TIMS.CINT1(TIMS.GetListValue(AppStage))
        If (v_IR_AppStage = 0) Then v_IR_AppStage = 3 'NULL(強制轉為政策性)
        Dim fg_USE_trKID25 As Boolean = $"{sm.UserInfo.Years}.{v_IR_AppStage}" <= "2026.1" '(2026上半年)強制使用trKID25 或有值
        'If TIMS.sUtl_ChkTest() Then fg_USE_trKID25 = $"{sm.UserInfo.Years}.{v_IR_AppStage}" <= "2025.1" '(TEST)

        Dim sPatternS1 As String = ",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航,進階政策性產業類別"
        Dim sColumnS1 As String = ",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,D25KNAME7,D25KNAME8,KNAME22"
        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
        If fg_Work2026x02 Then
            If Not fg_USE_trKID25 Then
                '1.五大信賴產業推動方案,'2.六大區域產業及生活圈,'3.智慧國家2.0綱領,'4.新南向政策推動計畫,
                '5.國家人才競爭力躍升方案,'6.AI新十大建設推動方案,'7.台灣AI行動計畫2.0,'8.智慧機器人產業推動方案,'9.臺灣2050淨零轉型
                sPatternS1 = ",五大信賴產業推動方案,六大區域產業及生活圈,台灣AI行動計畫2.0,智慧國家2.0綱領,國家人才競爭力躍升方案,新南向政策推動計畫,AI新十大建設推動方案,智慧機器人產業推動方案,臺灣2050淨零轉型"
                sColumnS1 = ",D26KNAME1,D26KNAME2,D26KNAME7,D26KNAME3,D26KNAME5,D26KNAME4,D26KNAME6,D26KNAME8,D26KNAME9"
            End If
        End If

        sPattern = ""
        sColumn = ""
        '跨區/轄區提案
        If tr_CrossDist_TP28.Visible Then
            sPattern &= "序號,計畫別,訓練單位名稱,申請階段,統一編號,單位屬性,分署別,課程名稱,提案意願順序,課程申請流水號,課程分類編碼,課程分類,訓練時數,訓練人次"
            sPattern &= ",每人訓練費用(元),訓練單位可向學員收取之訓練費用(元),總補助費(元)(以訓練費用之80%估算),辦理方式,實體課程時數"
            sPattern &= ",術科時數,固定費用總計,實際人時成本,材料明細,材料費總計,材料費占比,費用總計"
            ',教材費明細,教材費總計,材料費明細,材料費總計,其他費用明細,其他費用總計,人時成本上限,實際人時成本"
            'sPattern &= ",開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能,5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策"
            sPattern &= ",開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能"
            sPattern &= sPatternS1 '",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航,進階政策性產業類別"
            sPattern &= ",轄區重點產業,是否為學分班(Y/N),辦訓縣市別,聯絡人,聯絡電話,是否為iCAP課程,iCAP標章證號,iCAP有效期限,立案縣市"
            sPattern &= ",室外教學課程,報請主管機關核備,跨區/轄區提案,訓練業別同意協助重新歸類,線上送件"

            sColumn &= "SEQNUM,ORGPLANNAME2,ORGNAME,APPSTAGE,COMIDNO,ORGTYPENAME,DISTNAME,CLASSNAME,FIRSTSORT,PSNO28,KID12,D12KNAME,THOURS,TNUM"
            sColumn &= ",TOTAL,TOTALCOST,DEFGOVCOST,DISTANCE_N,ENTPHOUR"
            sColumn &= ",PROTECHHOURS,FIXSUMCOST,ACTHUMCOST,METDET,METSUMCOST,METCOSTPER,ALLSUMCOST"
            ',SHEETCOST,PT03,COMMONCOST,PT04,OTHERCOST,PT11,MAXUP,TIMECOST"
            'sColumn &= ",STDATE,FDDATE,GCODENAME,GCNAME,CODEID,CCNAME,D20KNAME1,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5,D20KNAME6"
            sColumn &= ",STDATE,FDDATE,GCODENAME,GCNAME,CODEID,CCNAME"
            sColumn &= sColumnS1 '",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,D25KNAME7,D25KNAME8,KNAME22"
            sColumn &= ",D15KNAME,PointYN_N,CTNAME,CONTACTNAME,CONTACTPHONEMOB,ISICAPCOUR_N,ICAPNUM,iCAPMARKDATE,ORGCTNAME"
            sColumn &= ",OUTDOOR_N,REPORTE_N,CROSSDIST_N,TMIDCORRECT,BIDCASEPI"
            Return
        End If

        sPattern &= "序號,計畫別,訓練單位名稱,統一編號,單位屬性,分署別,課程名稱,提案意願順序,課程申請流水號,課程分類編碼,課程分類,訓練時數,訓練人次"
        sPattern &= ",每人訓練費用(元),訓練單位可向學員收取之訓練費用(元),總補助費(元)(以訓練費用之80%估算),辦理方式,實體課程時數"
        sPattern &= ",術科時數,固定費用總計,實際人時成本,材料明細,材料費總計,材料費占比,費用總計"
        ',教材費明細,教材費總計,材料費明細,材料費總計,其他費用明細,其他費用總計,人時成本上限,實際人時成本"
        'sPattern &= ",開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能,5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策"
        sPattern &= ",開訓日期,結訓日期,訓練業別編碼,訓練業別,訓練職能編碼,訓練職能"
        sPattern &= sPatternS1 '",亞洲矽谷,重點產業,台灣AI行動計畫,智慧國家方案,國家人才競爭力躍升方案,新南向政策,AI加值應用,職場續航,進階政策性產業類別"
        sPattern &= ",轄區重點產業,是否為學分班(Y/N),辦訓縣市別,聯絡人,聯絡電話,是否為iCAP課程,iCAP標章證號,iCAP有效期限,立案縣市"
        sPattern &= ",室外教學課程,報請主管機關核備"

        sColumn &= "SEQNUM,ORGPLANNAME2,ORGNAME,COMIDNO,ORGTYPENAME,DISTNAME,CLASSNAME,FIRSTSORT,PSNO28,KID12,D12KNAME,THOURS,TNUM"
        sColumn &= ",TOTAL,TOTALCOST,DEFGOVCOST,DISTANCE_N,ENTPHOUR"
        sColumn &= ",PROTECHHOURS,FIXSUMCOST,ACTHUMCOST,METDET,METSUMCOST,METCOSTPER,ALLSUMCOST"
        ',SHEETCOST,PT03,COMMONCOST,PT04,OTHERCOST,PT11,MAXUP,TIMECOST"
        'sColumn &= ",STDATE,FDDATE,GCODENAME,GCNAME,CODEID,CCNAME,D20KNAME1,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5,D20KNAME6"
        sColumn &= ",STDATE,FDDATE,GCODENAME,GCNAME,CODEID,CCNAME"
        sColumn &= sColumnS1 '",D25KNAME1,D25KNAME2,D25KNAME3,D25KNAME4,D25KNAME5,D25KNAME6,D25KNAME7,D25KNAME8,KNAME22"
        sColumn &= ",D15KNAME,PointYN_N,CTNAME,CONTACTNAME,CONTACTPHONEMOB,ISICAPCOUR_N,ICAPNUM,iCAPMARKDATE,ORGCTNAME"
        sColumn &= ",OUTDOOR_N,REPORTE_N"
        Return

    End Sub

    ''' <summary>匯出鈕</summary>
    Sub Export1_2854()
        If TIMS.Cst_TPlanID2854.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg17)
            Return 'Exit Sub
        End If
        'If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then,'    Common.MessageBox(Me, "該計畫不提供此功能!!"),'    Exit Sub,'End If,
        Dim sErrMsg As String = ""
        Call checkData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        Dim parms As New Hashtable()
        '查詢 (SQL) 語法匯出 (產投／充飛)
        Dim sql As String = Search_SQL_2854(parms)
        If sql = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        TIMS.LOG.Debug(String.Format("##SD_15_022 dt.Rows.Count: {0}", dt.Rows.Count))

        Dim iRows1 As Integer = 0
        For Each dr As DataRow In dt.Rows
            iRows1 += 1
            dr("SEQNUM") = iRows1
            If Not Convert.IsDBNull(dr("PERSONCOST")) Then
                Dim s_COMMONCOST As String = If(Not Convert.IsDBNull(dr("COMMONCOST")), String.Format("{0};{1})", dr("PERSONCOST"), dr("COMMONCOST")), Convert.ToString(dr("PERSONCOST")))
                dr("COMMONCOST") = s_COMMONCOST
            End If
            Dim s_METDET As String = ""
            If Not Convert.IsDBNull(dr("METDETP")) Then
                s_METDET &= String.Concat(If(s_METDET <> "", ",", ""), dr("METDETP"))
            End If
            If Not Convert.IsDBNull(dr("METDETC")) Then
                s_METDET &= String.Concat(If(s_METDET <> "", ",", ""), dr("METDETC"))
            End If
            If Not Convert.IsDBNull(dr("METDETS")) Then
                s_METDET &= String.Concat(If(s_METDET <> "", ",", ""), dr("METDETS"))
            End If
            If Not Convert.IsDBNull(dr("METDETO")) Then
                s_METDET &= String.Concat(If(s_METDET <> "", ",", ""), dr("METDETO"))
            End If
            dr("METDET") = s_METDET
        Next

        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_FileName1 As String = String.Format("PVS-{0}-{1}-{2}", v_rblOrgKind2, TIMS.GetToday(objconn), TIMS.GetGUID())

        Dim sTitle1 As String = ""
        Select Case v_rblOrgKind2 'rblOrgKind2.SelectedValue
            Case "G", "W", "A"
                '產業人才投資計畫 'sTitle1 = CStr(sm.UserInfo.Years - 1911) & "年度產業人才投資計畫－課程提案彙總表"
                '提升勞工自主學習計畫 'sTitle1 = CStr(sm.UserInfo.Years - 1911) & "年度提升勞工自主學習計畫－課程提案彙總表"
                sTitle1 = $"{(sm.UserInfo.Years - 1911)}年度－課程提案彙總表"
            Case Else
                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Common.MessageBox(Me, "請選擇計畫!!")
                    Return
                End If
        End Select

        '匯出欄位設定
        Dim sPattern As String = ""
        Dim sColumn As String = ""
        'UTL_PCSTRING(v_rblOrgKind2, sPattern, sColumn)
        Call UTL_PCSTRING(sPattern, sColumn)

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", s_FileName1)
        parmsExp.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parmsExp.Add("TitleColSpanCnt", iColSpanCount)
        parmsExp.Add("sPatternA", sPatternA)
        parmsExp.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dt, parmsExp)
    End Sub

    '匯出
    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles BtnExp1.Click
        Call Export1_2854()
    End Sub

End Class
