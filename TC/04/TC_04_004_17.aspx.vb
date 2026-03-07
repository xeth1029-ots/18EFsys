Partial Class TC_04_004_17
    Inherits AuthBasePage

    '產業人才投資方案，審核計畫專用
    'Dim Auth_Relship As DataTable
    'select * from view_function where name ='各類課程明細表'
    '477
    '綜合查詢統計表	SD/15/SD_15_012.aspx	
    '各類課程明細表 SD/15/SD_15_009.aspx	
    'v_Depot04'view_Depot06
    'UPDATE PLAN_DEPOT
    'SELECT * FROM KEY_DEPOT
    'SELECT * FROM KEY_BUSINESS WHERE DEPID = '12' 
    'SELECT KID,KNAME FROM V_DEPOT12 ORDER BY KID
    'SELECT * FROM KEY_DEPOT
    'SELECT * FROM KEY_BUSINESS WHERE ROWNUM <=10
    'SELECT MAX (SEQNO) SEQNO FROM KEY_BUSINESS WHERE 1=1--ROWNUM <=10
    'INSERT INTO KEY_DEPOT (DEPID,DNAME,YEARS)  SELECT N'15',N'轄區重點產業',N'2017' ;
    'PLAN_DEPOT
    'Dim strYears As String = "" '2014 / 2015'(經費分類代碼。)
    'Const cst_y2017 As String = "2017"
    'Const cst_y2018 As String = "2018"
    'Const cst_y2014 As String = "2014" '(此功能for 2017使用,其他年度停用)
    'Const cst_y2015 As String = "2015" '(此功能for 2017使用,其他年度停用)
    'Const cst_y2017 As String = "2017"
    'Const cst_DG1_新南向政策 As Integer = 12
    Dim fg_File1_xls As Boolean = False
    Dim fg_File1_ods As Boolean = False

    '取得SQL iType 1:1筆資料 2:多筆資料 3:匯出欄位修改
    Const cst_Exp_iType_1筆資料 As Integer = 1
    Const cst_Exp_iType_2多筆資料 As Integer = 2
    Const cst_Exp_iType_3匯出欄位 As Integer = 3

    '(署)啟動產業別匯入及其他功能 署:True/非署:False
    Dim fg_CAN_USE_1_LID_0 As Boolean = True ' False

    Dim aOCID As String = ""
    Dim aKID60 As String = ""
    Dim aPLANID As String = ""
    Dim aCOMIDNO As String = ""
    Dim aSEQNO As String = ""

    Const cst_aOCID As Integer = 0 '班別代碼／課程申請流水號
    Const cst_aKID60 As Integer = 1 '產業別代號
    Const cst_iFiledColumnNum As Integer = 2 '欄位對應至少欄位數 

    'https://jira.turbotech.com.tw/browse/TIMSC-276
    '2018年度開始適用，政府政策性產業欄位選項修改，由舊有7項改為17項，選項修改對照表詳如附件。
    Dim flag_Years2017 As Boolean = False '2017轉程式執行
    Dim flag_Years2018 As Boolean = False '2018啟用
    '<summary>trKID20.Visible '2018/2019 (政府政策性產業) </summary> 'Dim flag_Years2019 As Boolean = False '2019啟用
    Dim gflag_test As Boolean = False '測試環境為true
    ''' <summary>
    ''' 繼續搜尋動作
    ''' </summary>
    Dim gflag_can_continue_sch As Boolean = True
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        fg_CAN_USE_1_LID_0 = (sm.UserInfo.LID = 0)
        tr_txtCLASSNAMEKeyWORDS_SCH.Visible = fg_CAN_USE_1_LID_0
        tr_Btn_XlsImport1.Visible = fg_CAN_USE_1_LID_0
        tr_RBListExpType.Visible = fg_CAN_USE_1_LID_0
        btnExp1.Visible = fg_CAN_USE_1_LID_0
        tr_CBLKID60.Visible = fg_CAN_USE_1_LID_0

        Call TIMS.OpenDbConn(objconn) '開啟連線

        iPYNum = TIMS.sUtl_GetPYNum(Me)
        'Dim flag2017 As Boolean = False '2017轉程式執行
        If sm.UserInfo.Years >= 2017 Then flag_Years2017 = True
        If sm.UserInfo.Years >= 2018 Then flag_Years2018 = True

        'Dim gflag_test As Boolean = TIMS.sUtl_ChkTest() '測試
        gflag_test = TIMS.sUtl_ChkTest() '測試
        If gflag_test Then
            flag_Years2017 = True '測試啟用
            flag_Years2018 = True '測試啟用
            'flag_Years2019 = True '測試啟用
        End If
        If flag_Years2018 Then Hid_Years2018.Value = TIMS.cst_YES

        'trKID20.Visible '2018/2019 (政府政策性產業)
        'flag_Years2019 = TIMS.SHOW_2019_1(sm)
        'If flag_Years2019 Then Hid_Years2019.Value = TIMS.cst_YES

        If Not flag_Years2017 Then
            '(非測試環境 不可使用) 2017->2014轉程式執行
            Dim url1 As String = "TC_04_004.aspx?ID=" & Request("ID")
            TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        '(經費分類代碼。)
        'strYears = cst_y2014 '2014年  顯示層級。
        'If sm.UserInfo.Years >= "2015" Then strYears = cst_y2015 '2015年 不顯示層級。
        'If flag2017 Then strYears = cst_y2017 '測試啟用
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then LabTMID.Text = "訓練業別"

        HyperLink1.NavigateUrl = "../../Doc/PlanDEPOT60v23.zip"

        PageControler1.PageDataGrid = Me.DataGrid1

        'btnQuery.Enabled = True
        'If Not au.blnCanSech Then btnQuery.Enabled = False
        'If Not au.blnCanSech Then TIMS.Tooltip(btnQuery, "沒有搜尋權限")

        If Not IsPostBack Then
            Call Create1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        'Button1.Attributes("onclick") = "return SavaData();"
    End Sub

    '初始載入
    Sub Create1()
        msg.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        DataGridTable1.Visible = False '預設搜尋資料不顯示
        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉

        '有不區分
        OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
        Common.SetListItem(OrgKind2, "A")

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)
        If tr_AppStage_TP28.Visible Then
            AppStage2 = TIMS.Get_AppStage2_NotCase(AppStage2)
            Common.SetListItem(AppStage2, "")
        End If

        '課程分類 ddlDEPOT12
        Dim sql As String = ""
        ddlDEPOT12.Items.Clear()
        sql = "SELECT KID,KNAME FROM V_DEPOT12 ORDER BY KID" '19:其他類
        DbAccess.MakeListItem(ddlDEPOT12, sql, objconn)
        ddlDEPOT12.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
        ddlDEPOT12.Enabled = False
        TIMS.Tooltip(ddlDEPOT12, "不能修改")

        '轄區重點產業(排除停用)(轄區)
        Call TIMS.CHG_ddlDEPOT15(ddlDEPOT15, sm.UserInfo.DistID, sm.UserInfo.TPlanID, objconn)

        '產業別(管考)
        CBLKID60.Items.Clear()
        sql = "SELECT KID,KNAME FROM VIEW_DEPOT60 ORDER BY KID"
        DbAccess.MakeListItem(CBLKID60, sql, objconn)
        'ddlKID60.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        'ddlKID06.Items.Clear() '2015 六大新興產業 (DEPID='10')
        'sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='10' ORDER BY KID"
        'DbAccess.MakeListItem(ddlKID06, sql, objconn)
        'ddlKID06.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        ''2017 (重點服務業) SELECT * FROM KEY_DEPOT WHERE DEPID='16' 
        'ddlKID10.Items.Clear()
        'sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='16' ORDER BY KID"
        'DbAccess.MakeListItem(ddlKID10, sql, objconn)
        'ddlKID10.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        'https://jira.turbotech.com.tw/browse/TIMSC-276
        'ddlKID17.Visible = True
        'ddlKID19.Visible = False
        'If flag_Years2018 Then
        '    ddlKID17.Visible = False
        '    ddlKID19.Visible = True
        'End If
        'trKID20.Visible = False
        ''trKID20.Visible '2018/2019 (政府政策性產業)
        'If flag_Years2019 Then
        '    trKID19.Visible = False
        '    trKID18.Visible = False
        '    trKID20.Visible = True
        'End If

        'If ddlKID17.Visible Then
        '    '2017 (政府政策性產業) SELECT * FROM KEY_DEPOT WHERE DEPID = '17' 
        '    ddlKID17.Items.Clear()
        '    sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='17' ORDER BY KID"
        '    DbAccess.MakeListItem(ddlKID17, sql, objconn)
        '    ddlKID17.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
        'End If
        '2018 (政府政策性產業)
        'If ddlKID19.Visible Then
        '    ddlKID19.Items.Clear()
        '    sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='19' ORDER BY KID"
        '    DbAccess.MakeListItem(ddlKID19, sql, objconn)
        '    ddlKID19.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
        'End If
        '2017 (新南向政策) SELECT * FROM KEY_DEPOT WHERE DEPID = '18' 
        'ddlKID18.Items.Clear()
        'sql = "SELECT KID,KNAME FROM KEY_BUSINESS WHERE DEPID='18' ORDER BY KID"
        'DbAccess.MakeListItem(ddlKID18, sql, objconn)
        'ddlKID18.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        'trKID20.Visible '2018 (政府政策性產業)
        'sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID = '20' ORDER BY KID"
        'Dim dtKID_N20 As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim dtKID_N20 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "20")
        Call TIMS.GET_CBL_KID20(CBLKID20_1, dtKID_N20, 1)
        Call TIMS.GET_CBL_KID20(CBLKID20_2, dtKID_N20, 2)
        Call TIMS.GET_CBL_KID20(CBLKID20_3, dtKID_N20, 3)
        Call TIMS.GET_CBL_KID20(CBLKID20_4, dtKID_N20, 4)
        Call TIMS.GET_CBL_KID20(CBLKID20_5, dtKID_N20, 5)
        Call TIMS.GET_CBL_KID20(CBLKID20_6, dtKID_N20, 6)

        Dim dtKID_N22 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "22")
        Call TIMS.GET_CBL_KID22(CBLKID22, dtKID_N22)
        Call TIMS.GET_CBL_KID22(CBLKID22B, dtKID_N22)

        Dim dtKID_N25 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "25")
        Call TIMS.GET_CBL_KID25(CBLKID25_1, dtKID_N25, 1)
        Call TIMS.GET_CBL_KID25(CBLKID25_2, dtKID_N25, 2)
        Call TIMS.GET_CBL_KID25(CBLKID25_3, dtKID_N25, 3)
        Call TIMS.GET_CBL_KID25(CBLKID25_4, dtKID_N25, 4)
        Call TIMS.GET_CBL_KID25(CBLKID25_5, dtKID_N25, 5)
        Call TIMS.GET_CBL_KID25(CBLKID25_6, dtKID_N25, 6)
        Call TIMS.GET_CBL_KID25(CBLKID25_7, dtKID_N25, 7)
        Call TIMS.GET_CBL_KID25(CBLKID25_8, dtKID_N25, 8)
    End Sub

    Function str_WCS1() As String
        Dim sSql As String = ""
        sSql &= " SELECT cc.PLANID,cc.COMIDNO,cc.SEQNO" & vbCrLf
        sSql &= " ,min(cc.OCID) OCID" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='03' then 1 else 0 end ) openstudcount1" ' --,實際就保開訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='02' then 1 else 0 end ) openstudcount2" ' --,實際就安開訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID in ('02','03') then 1 else 0 end ) openstudcount12" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 and cs.BudgetID='97' then 1 else 0 end ) openstudcount97" ' --,實際公務(ECFA)開訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE)=1 then 1 else 0 end) openstudcountall" ' --,實際合計開訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='03' then 1 else 0 end ) closestudcout03" ' --,就保結訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='02' then 1 else 0 end ) closestudcout02" ' --,就安結訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='97' then 1 else 0 end ) closestudcout97" ' --,公務(ECFA)結訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID IN ('03','02','97') then 1 else 0 end ) closestudcoutall" ' --,合計結訓人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and ss.AppliedStatus = '1' and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='03' then 1 else 0 end ) budcountall3" ' --,就保撥款人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and ss.AppliedStatus = '1' and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='02' then 1 else 0 end ) budcountall2" ' --,就安撥款人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and ss.AppliedStatus = '1' and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='97' then 1 else 0 end ) budcountall97" ' --,公務(ECFA)撥款人數" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and ss.AppliedStatus = '1' and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='03' then ss.SumOfMoney else 0 end ) budmoneyall3" ' --,就保撥款補助費" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and ss.AppliedStatus = '1' and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='02' then ss.SumOfMoney else 0 end ) budmoneyall2" ' --,就安撥款補助費" & vbCrLf
        sSql &= " ,sum(case when cc.Notopen='N' and cs.IsApprPaper='Y' and cs.CreditPoints is not NULL and ss.AppliedStatus = '1' and cs.StudStatus Not IN (2,3) and cc.FTDate < GETDATE() and cs.BudgetID='97' then ss.SumOfMoney else 0 end ) budmoneyall97" ' --,公務(ECFA)撥款補助費" & vbCrLf
        sSql &= " ,min(cc.NOTOPEN) NOTOPEN" & vbCrLf
        sSql &= " FROM dbo.VIEW_PLAN ip" & vbCrLf
        sSql &= " JOIN dbo.CLASS_CLASSINFO cc WITH(NOLOCK) on cc.PLANID=ip.PLANID" & vbCrLf
        sSql &= " JOIN dbo.CLASS_STUDENTSOFCLASS cs WITH(NOLOCK) on cs.OCID=cc.OCID" & vbCrLf
        sSql &= " JOIN dbo.STUD_STUDENTINFO ssi WITH(NOLOCK) on ssi.SID = cs.SID" & vbCrLf
        sSql &= " LEFT JOIN dbo.STUD_SUBSIDYCOST ss WITH(NOLOCK) on ss.SOCID=cs.SOCID" & vbCrLf
        sSql &= " WHERE ip.PlanKind=2" & vbCrLf '計畫種類:1.自辦／2.委外
        sSql &= " AND ip.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf '登入計畫
        sSql &= " AND ip.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf '登入年度
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0" '署(局)
            Case Else '非署(局)
                sSql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf '登入委訓計畫
                sSql &= " AND ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf '登入轄區
        End Select
        sSql &= " GROUP BY cc.PLANID,cc.COMIDNO,cc.SEQNO" & vbCrLf
        Return sSql
    End Function

    '取得SQL iType 1:1筆資料 2:多筆資料 3:匯出欄位修改
    Function sUtl_GetSQL2(ByVal iType As Integer, ByVal ss As String) As String
        Dim sql As String = ""
        '(匯出欄位重新定義)
        If iType = cst_Exp_iType_3匯出欄位 Then
            Dim sSql As String = ""
            sSql = String.Concat("WITH WCS1 AS (", str_WCS1(), ")", vbCrLf)
            sSql &= " SELECT ip.YEARS 年度" & vbCrLf
            sSql &= " ,ar.ORGPLANNAME 計畫別" & vbCrLf
            sSql &= " ,ip.DISTNAME 轄區" & vbCrLf
            sSql &= " ,ar.ORGNAME 訓練機構" & vbCrLf
            sSql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) 班級名稱" & vbCrLf
            sSql &= " ,pp.PSNO28 課程申請流水號" & vbCrLf
            sSql &= " ,cs.OCID 班別代碼" & vbCrLf
            sSql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) 申請階段" & vbCrLf
            sSql &= " ,format(pp.STDATE,'yyyy/MM/dd') 開訓日期" & vbCrLf
            sSql &= " ,format(pp.FDDATE,'yyyy/MM/dd') 結訓日期" & vbCrLf
            sSql &= " ,dd.KID60 產業別代號" & vbCrLf
            sSql &= " ,dbo.FN_GET_KID60NAME(pp.PLANID,pp.COMIDNO,pp.SEQNO) 產業別名稱" & vbCrLf
            sSql &= " ,pp.TNUM 核定人數" & vbCrLf
            sSql &= " ,pp.DefGovCost 核定補助費" & vbCrLf
            sSql &= " ,cs.openstudcount1 實際就保開訓人數" & vbCrLf
            sSql &= " ,cs.openstudcount2 實際就安開訓人數" & vbCrLf
            sSql &= " ,cs.openstudcount12 ""就安就保開訓人數""" & vbCrLf
            sSql &= " ,cs.openstudcount97 實際公務ECFA開訓人數" & vbCrLf
            sSql &= " ,cs.openstudcountall 實際合計開訓人數" & vbCrLf
            sSql &= " ,cs.closestudcout03 就保結訓人數" & vbCrLf
            sSql &= " ,cs.closestudcout02 就安結訓人數" & vbCrLf
            sSql &= " ,cs.closestudcout97 公務ECFA結訓人數" & vbCrLf
            sSql &= " ,cs.closestudcoutall 合計結訓人數" & vbCrLf
            sSql &= " ,cs.budcountall3 就保撥款人數" & vbCrLf
            sSql &= " ,cs.budcountall2 就安撥款人數" & vbCrLf
            sSql &= " ,cs.budcountall97 公務ECFA撥款人數" & vbCrLf
            sSql &= " ,cs.budmoneyall3 就保撥款補助費" & vbCrLf
            sSql &= " ,cs.budmoneyall2 就安撥款補助費" & vbCrLf
            sSql &= " ,cs.budmoneyall97 公務ECFA撥款補助費" & vbCrLf
            sSql &= " ,cs.NOTOPEN 是否停辦" & vbCrLf
            sql = sSql

        Else
            sql = ""
            sql &= " SELECT ip.YEARS,ip.DISTNAME,ip.DISTID,ar.ORGNAME" & vbCrLf
            sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
            sql &= " ,pp.THOURS,vt.JOBNAME,vt.TRAINNAME" & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) CLASSNAME" & vbCrLf
            sql &= " ,FORMAT(pp.STDate,'yyyy/MM/dd') STDate" & vbCrLf
            sql &= " ,FORMAT(pp.FDDate,'yyyy/MM/dd') FDDate" & vbCrLf
            sql &= " ,ig3.GCODE2 GOVCLASSN" & vbCrLf
            '職類課程分類名稱 PNAME,JOBNAME /訓練業別 CNAME '(課程分類／職類課程)
            sql &= " ,ig3.GCODE31,ig3.PNAME,ig3.CNAME" & vbCrLf
            sql &= " ,pp.GCID,pp.GCID2,pp.GCID3" & vbCrLf
            sql &= " ,dd.AppResult" & vbCrLf
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID12 ,ISNULL(dd.KID12,d12.KID)) KID12" & vbCrLf
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.SEQNOD15 ,NULL) SEQNOD15" & vbCrLf
            '課程分類 
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',K12.KNAME ,ISNULL(K12.KNAME ,d12.KNAME)) D12KNAME" & vbCrLf
            '轄區重點產業
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',K15.KNAME ,NULL) D15KNAME" & vbCrLf
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',K06.KNAME ,ISNULL(K06.KNAME ,d06.KNAME)) D06KNAME" & vbCrLf '6大新興產業
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',K10.KNAME ,NULL) D10KNAME" & vbCrLf '10大重點服務業(9項)
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID17 ,NULL) KID17" & vbCrLf
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',K17.KNAME ,NULL) D17KNAME" & vbCrLf
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID19 ,NULL) KID19" & vbCrLf
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',K19.KNAME ,NULL) D19KNAME" & vbCrLf
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID18 ,NULL) KID18" & vbCrLf
            'sql &= " ,dbo.DECODE(dd.AppResult,'Y',K18.KNAME ,NULL) D18KNAME" & vbCrLf
            '2019年啟用 work2019x01:2019 政府政策性產業
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID20 ,NULL) KID20" & vbCrLf
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',dd.KID25 ,NULL) KID25" & vbCrLf
            sql &= " ,dbo.DECODE(dd.AppResult,'Y',dbo.FN_GET_KID20NAME(pp.PLANID,pp.COMIDNO,pp.SEQNO) ,NULL) D20KNAME" & vbCrLf
            sql &= " ,dd.KID22" & vbCrLf
            sql &= " ,dd.KID60" & vbCrLf '產業別代號
            sql &= " ,pp.PSNO28" & vbCrLf '課程申請流水號
            sql &= " ,dbo.FN_OCID(pp.PLANID,pp.COMIDNO,pp.SEQNO) OCID" & vbCrLf

        End If
        sql &= " FROM dbo.PLAN_PLANINFO pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.PLANID =pp.PLANID" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME ar ON ar.RID=pp.RID" & vbCrLf '(業務確認)
        'sql &= " JOIN dbo.ORG_ORGINFO oo ON oo.ORGID=ar.ORGID" & vbCrLf '(機構確認)
        'sql &= " JOIN dbo.ORG_ORGPLANINFO oop ON oop.RSID=ar.RSID" & vbCrLf '(機構資料確認)
        sql &= " LEFT JOIN dbo.VIEW_TRAINTYPE vt on vt.TMID=pp.TMID" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERREPORT pvr ON pp.PLANID=pvr.PLANID AND pp.COMIDNO=pvr.COMIDNO AND pp.SEQNO=pvr.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERRECORD pvrd ON pp.PLANID=pvrd.PLANID AND pp.COMIDNO=pvrd.COMIDNO AND pp.SEQNO=pvrd.SEQNO" & vbCrLf
        '(2015-2017)'依GCID
        sql &= " LEFT JOIN dbo.VIEW_GOVCLASSCAST ig on pp.GCID=ig.GCID" & vbCrLf
        '依GCID2
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST2 ig2 on pp.GCID2=ig2.GCID2" & vbCrLf
        '依GCID3
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 on pp.GCID3=ig3.GCID3" & vbCrLf
        '依GCID2
        sql &= " LEFT JOIN dbo.VIEW_DEPOT12 d12 ON d12.GCID2 =pp.GCID2" & vbCrLf
        '依GCID2
        sql &= " LEFT JOIN dbo.VIEW_DEPOT10 d06 ON d06.GCID2 =pp.GCID2" & vbCrLf
        '2017-
        sql &= " LEFT JOIN dbo.PLAN_DEPOT dd ON pp.PLANID=dd.PLANID AND pp.COMIDNO=dd.COMIDNO AND pp.SEQNO=dd.SEQNO" & vbCrLf
        sql &= " LEFT JOIN (SELECT KID,KNAME FROM dbo.KEY_BUSINESS WHERE DEPID='12') K12 ON K12.KID =dd.KID12" & vbCrLf
        sql &= " LEFT JOIN (SELECT SEQNO SEQNOD15,KID,KNAME FROM dbo.KEY_BUSINESS WHERE DEPID IN ('15','21')) K15 ON K15.SEQNOD15 =dd.SEQNOD15" & vbCrLf
        If iType = cst_Exp_iType_3匯出欄位 Then
            sql &= " LEFT JOIN WCS1 cs on pp.PLANID=cs.PLANID AND pp.COMIDNO=cs.COMIDNO AND pp.SEQNO=cs.SEQNO" & vbCrLf
        End If
        sql &= " WHERE ip.PlanKind=2" & vbCrLf '計畫種類:1.自辦／2.委外
        'sql &= " AND pvr.SecResult='Y' --複審結果通過" & vbCrLf
        sql &= " AND ip.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf '登入計畫
        sql &= " AND ip.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf '登入年度
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0" '署(局)
            Case Else '非署(局)
                sql &= " AND ip.PlanID ='" & sm.UserInfo.PlanID & "'" & vbCrLf '登入委訓計畫
                sql &= " AND ip.DistID ='" & sm.UserInfo.DistID & "'" & vbCrLf '登入轄區
        End Select
        '業務權限(中心)
        If ViewState("relship") <> "" Then
            'sql &= " AND EXISTS ( SELECT 'x' FROM AUTH_RELSHIP x WHERE x.RELSHIP LIKE '" & ViewState("relship") & "%' and x.RID =pp.RID )" & vbCrLf
            sql &= " AND ar.RELSHIP LIKE '" & ViewState("relship") & "%'" & vbCrLf
        End If
        'If ViewState("jobValue") <> "" Then sql &= " and pp.TMID = " & ViewState("jobValue") & vbCrLf
        'If ViewState("trainValue") <> "" Then sql &= " and pp.TMID = " & ViewState("trainValue") & vbCrLf
        If Hid_TMIDVALUE.Value <> "" Then sql &= " and pp.TMID = " & Hid_TMIDVALUE.Value & vbCrLf
        '通俗職類
        If ViewState("cjobValue") <> "" Then
            sql &= " and pp.CJOB_UNKEY = " & ViewState("cjobValue") & "" & vbCrLf
        End If
        If ViewState("ClassName") <> "" Then
            sql &= " and pp.ClassName like '%" & ViewState("ClassName") & "%'" & vbCrLf
        End If
        If fg_CAN_USE_1_LID_0 Then
            '1.在查詢條件，增加一個【課程關鍵字】搜尋(圖一紅框1) 可輸入關鍵字， 針對【課程大綱/內容】搜尋
            txtCLASSNAMEKeyWORDS_SCH.Text = TIMS.ClearSQM(txtCLASSNAMEKeyWORDS_SCH.Text)
            If txtCLASSNAMEKeyWORDS_SCH.Text <> "" Then
                sql &= " and dbo.FN_GET_TRAINDESC(pp.PLANID,pp.COMIDNO,pp.SEQNO,'PCONT') LIKE '%'+'" & txtCLASSNAMEKeyWORDS_SCH.Text & "'+'%'" & vbCrLf
            End If
        End If

        Select Case iType
            Case cst_Exp_iType_1筆資料 '1 '單筆
                sql &= " AND ip.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf '登入計畫
                sql &= " AND ip.Years ='" & sm.UserInfo.Years & "'" & vbCrLf '登入年度
                Dim PlanID As String = TIMS.GetMyValue(ss, "PlanID")
                Dim ComIDNO As String = TIMS.GetMyValue(ss, "ComIDNO")
                Dim SeqNo As String = TIMS.GetMyValue(ss, "SeqNo")
                '單筆'(正式須檢核)
                If Not gflag_test Then sql &= " AND pvr.IsApprPaper='Y'" & vbCrLf '正式
                'sql &= " AND ip.PlanKind=2" & vbCrLf '計畫種類:1.自辦／2.委外
                sql &= " AND pp.PlanID ='" & PlanID & "'" & vbCrLf
                sql &= " AND pp.ComIDNO ='" & ComIDNO & "'" & vbCrLf
                sql &= " AND pp.SeqNo ='" & SeqNo & "'" & vbCrLf
                Return sql
            Case Else
                '多筆
        End Select
        '課程申請流水號
        'txtPSNO28_SCH.Text = TIMS.ClearSQM(txtPSNO28_SCH.Text)
        'If txtPSNO28_SCH.Text <> "" AndAlso txtPSNO28_SCH.Text.Length >= 8 Then
        '    sql &= " and pp.PSNO28='" & txtPSNO28_SCH.Text & "'" & vbCrLf
        'End If
        If ViewState("CyclType") <> "" Then
            sql &= " AND pp.CyclType='" & ViewState("CyclType") & "'" & vbCrLf
        End If
        If ViewState("STDate1") <> "" Then
            sql &= " AND pp.STDate >=" & TIMS.To_date(ViewState("STDate1")) & vbCrLf
        End If
        If ViewState("STDate2") <> "" Then
            sql &= " AND pp.STDate <=" & TIMS.To_date(ViewState("STDate2")) & vbCrLf
        End If
        If ViewState("FDDate1") <> "" Then
            sql &= " AND pp.FDDate >=" & TIMS.To_date(ViewState("FDDate1")) & vbCrLf
        End If
        If ViewState("FDDate2") <> "" Then
            sql &= " AND pp.FDDate <=" & TIMS.To_date(ViewState("FDDate2")) & vbCrLf
        End If
        If ViewState("OrgKind2") <> "" Then
            sql &= " AND ar.OrgKind2 ='" & ViewState("OrgKind2") & "'" & vbCrLf
        End If
        '依申請階段
        If ViewState("AppStage2") <> "" Then sql &= " AND pp.AppStage= '" & ViewState("AppStage2") & "'" & vbCrLf

        'If ViewState("sqlSecResult") <> "" Then
        '    sql &= ViewState("sqlSecResult") & vbCrLf
        'End If
        'ViewState("sqlSecResult") = ""
        '(正式須檢核) '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
        If Not gflag_test Then
            sql &= " AND pvr.IsApprPaper='Y'" & vbCrLf '正式
        End If

        Dim v_PlanMode As String = TIMS.GetListValue(PlanMode)
        Select Case v_PlanMode 'PlanMode.SelectedValue
            Case "S" '審核中的
                'sql &= " AND pvr.IsApprPaper='Y'" & vbCrLf '正式
                sql &= " AND pvr.SecResult IS NULL "
            Case "Y" '已通過
                sql &= " AND pvr.SecResult='Y'"
            Case "R" '退件修正
                sql &= " AND pvr.SecResult in ('R','N')"
        End Select

        Return sql
    End Function

    'SQL PageDataTable LIST
    Sub Search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉

        Dim sql As String = sUtl_GetSQL2(cst_Exp_iType_2多筆資料, "")

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        'trKID20.Visible '2018/2019 (政府政策性產業)
        'DataGrid1.Columns(cst_DG1_新南向政策).Visible = True
        'If flag_Years2019 Then DataGrid1.Columns(cst_DG1_新南向政策).Visible = False

        DataGridTable1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    'SQL return datarow
    Function search2(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String) As DataRow
        Dim rst As DataRow = Nothing
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Return rst

        Dim ss As String = ""
        TIMS.SetMyValue(ss, "PlanID", PlanID)
        TIMS.SetMyValue(ss, "ComIDNO", ComIDNO)
        TIMS.SetMyValue(ss, "SeqNo", SeqNo)
        Dim sql As String = sUtl_GetSQL2(cst_Exp_iType_1筆資料, ss)

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then rst = dt.Rows(0)

        Return rst
    End Function

    '設定 ViewState Value
    Sub sUtl_SetViewStateValue()
        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        Hid_TMIDVALUE.Value = trainValue.Value 'If(trainValue.Value <> "",trainValue.Value, "")
        'If iPYNum >= 3 Then
        '    '訓練職類
        '    Hid_TMIDVALUE.Value = trainValue.Value 'If(trainValue.Value <> "",trainValue.Value, "")
        'Else
        '    '訓練業別
        '    Hid_TMIDVALUE.Value = jobValue.Value 'If(jobValue.Value <> "",jobValue.Value, "")
        'End If

        '通俗職類/'班別名稱
        ViewState("cjobValue") = cjobValue.Value 'If(cjobValue.Value <> "",cjobValue.Value, "")
        ViewState("ClassName") = ClassName.Text 'If(Trim(ClassName.Text) <> "", Trim(ClassName.Text).Replace("'", "''"), "")
        'ViewState("jobValue") = TIMS.ClearSQM(ViewState("jobValue"))
        'ViewState("trainValue") = TIMS.ClearSQM(ViewState("trainValue"))
        'Hid_TMIDVALUE.Value = TIMS.ClearSQM(Hid_TMIDVALUE.Value)
        ViewState("cjobValue") = TIMS.ClearSQM(ViewState("cjobValue"))
        ViewState("ClassName") = TIMS.ClearSQM(ViewState("ClassName"))
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
        ViewState("CyclType") = CyclType.Text 'TIMS.ClearSQM(ViewState("CyclType"))
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FDDate1.Text = TIMS.ClearSQM(FDDate1.Text)
        FDDate2.Text = TIMS.ClearSQM(FDDate2.Text)
        ViewState("STDate1") = If(STDate1.Text <> "", STDate1.Text, "")
        ViewState("STDate2") = If(STDate2.Text <> "", STDate2.Text, "")
        ViewState("FDDate1") = If(FDDate1.Text <> "", FDDate1.Text, "")
        ViewState("FDDate2") = If(FDDate2.Text <> "", FDDate2.Text, "")
        ViewState("OrgKind2") = ""
        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Select Case v_OrgKind2
            Case "G", "W"
                ViewState("OrgKind2") = v_OrgKind2 'OrgKind2.SelectedValue
        End Select
        ViewState("OrgKind2") = TIMS.ClearSQM(ViewState("OrgKind2"))
        '依申請階段
        Dim v_AppStage2 As String = "" 'TIMS.GetListValue(AppStage2)
        If tr_AppStage_TP28.Visible Then v_AppStage2 = TIMS.GetListValue(AppStage2)
        ViewState("AppStage2") = v_AppStage2

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Select Case sm.UserInfo.LID
            Case 1, 2
                If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        End Select
        'ViewState("relship") = "" '業務權限(中心)
        Dim Relship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)
        ViewState("relship") = TIMS.ClearSQM(Relship)
    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        Call sUtl_SetViewStateValue()  '設定 ViewState Value
        Call Search1()
    End Sub

    ''' <summary>儲存 INSERT/UPDATE Plan_Depot (SAVE) </summary>
    ''' <param name="sSearchW"></param>
    Sub sUtl_SAVEDATA1_ChkOk(ByVal sSearchW As String)
        gflag_can_continue_sch = True

        '確認 (若課程分類選擇其他)
        'Const cst_KID12_其他 As String = "19" 'KID12 課程分類 '19:其他類 ddlDepot12

        'sSearchW = e.CommandArgument
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")
        '確認 (若課程分類選擇其他)
        Dim GCODE31 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "GCODE31"))
        Dim KID12 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "KID12")) '19:其他類
        If KID12 = "" Then KID12 = GCODE31
        Dim SEQNOD15 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "SEQNOD15"))
        Dim cvKID20 As String = TIMS.CombiSQM2IN(TIMS.GetMyValue(sSearchW, "KID20"), True)
        Dim cvKID25 As String = TIMS.CombiSQM2IN(TIMS.GetMyValue(sSearchW, "KID25"), True)
        Dim KID22 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "KID22"))
        Dim cvKID60 As String = TIMS.CombiSQM2IN(TIMS.GetMyValue(sSearchW, "KID60"), True)
        'Dim KID06 As String = TIMS.GetMyValue(sSearchW, "KID06")
        'Dim KID10 As String = TIMS.GetMyValue(sSearchW, "KID10")
        'Dim KID17 As String = TIMS.GetMyValue(sSearchW, "KID17")
        'Dim KID18 As String = TIMS.GetMyValue(sSearchW, "KID18")
        'Dim KID19 As String = TIMS.GetMyValue(sSearchW, "KID19")
        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Exit Sub
        If cvKID25 <> "" Then cvKID20 = ""

        'trKID20.Visible '2018/2019 (政府政策性產業) If flag_Years2019 Then End If
        Dim sErrMsg1 As String = CHK_KID20_VAL()
        If sErrMsg1 <> "" Then
            gflag_can_continue_sch = False
            Common.MessageBox(Me, sErrMsg1)
            Return 'Exit Sub
        End If
        Dim sErrMsg2 As String = TIMS.CHK_KID60_VAL(CBLKID60)
        If sErrMsg2 <> "" Then
            gflag_can_continue_sch = False
            Common.MessageBox(Me, sErrMsg2)
            Return
        End If
        If Not TIMS.OpenDbConn(objconn) Then
            gflag_can_continue_sch = False
            sErrMsg1 = TIMS.cst_ErrorMsg13
            Common.MessageBox(Me, sErrMsg1)
            Return
        End If
        Dim sErrMsg3 As String = CHK_KID25_VAL_OTH() '2025 政府政策性產業
        If trKID25.Visible AndAlso sErrMsg3 <> "" Then
            gflag_can_continue_sch = False
            Common.MessageBox(Me, sErrMsg3)
            Return  'Exit Sub
        End If

        Dim s_parms As New Hashtable From {{"PLANID", PlanID}, {"COMIDNO", ComIDNO}, {"SEQNO", SeqNo}}
        Dim sql As String = ""
        sql = "SELECT 'X' FROM PLAN_DEPOT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, s_parms)

        'If fg_CAN_USE_1_LID_0 Then '(增加:KID60=@KID60)
        Dim i_sql As String = ""
        i_sql &= " INSERT INTO PLAN_DEPOT(PLANID,COMIDNO,SEQNO,KID12,SEQNOD15,KID20,KID22,KID25" & vbCrLf
        If fg_CAN_USE_1_LID_0 Then i_sql &= " ,KID60" & vbCrLf
        i_sql &= " ,APPRESULT,MODIFYACCT ,MODIFYDATE)" & vbCrLf
        i_sql &= " VALUES (@PLANID,@COMIDNO,@SEQNO,@KID12,@SEQNOD15,@KID20,@KID22,@KID25" & vbCrLf
        If fg_CAN_USE_1_LID_0 Then i_sql &= " ,@KID60" & vbCrLf
        i_sql &= " ,'Y',@MODIFYACCT ,GETDATE())" & vbCrLf

        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_DEPOT" & vbCrLf
        u_sql &= " SET KID12=@KID12,SEQNOD15=@SEQNOD15,KID20=@KID20,KID22=@KID22,KID25=@KID25" & vbCrLf
        If fg_CAN_USE_1_LID_0 Then u_sql &= " ,KID60=@KID60" & vbCrLf
        u_sql &= " ,APPRESULT='Y',MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf

        If dt1.Rows.Count > 0 Then
            '確認 (若課程分類選擇其他)
            'sql &= " ,KID06=@KID06" & vbCrLf 'sql &= " ,KID10=@KID10" & vbCrLf 'sql &= " ,KID17=@KID17" & vbCrLf
            'u_sqll &= " ,KID18=@KID18" & vbCrLf 'sql &= " ,KID19=@KID19" & vbCrLf
            Dim uCmd As New SqlCommand(u_sql, objconn)
            With uCmd
                .Parameters.Clear()
                .Parameters.Add("@PLANID", SqlDbType.VarChar).Value = PlanID
                .Parameters.Add("@COMIDNO", SqlDbType.VarChar).Value = ComIDNO
                .Parameters.Add("@SEQNO", SqlDbType.VarChar).Value = SeqNo

                .Parameters.Add("@KID12", SqlDbType.VarChar).Value = If(KID12 <> "", KID12, Convert.DBNull) 'KID12
                .Parameters.Add("@SEQNOD15", SqlDbType.VarChar).Value = If(SEQNOD15 <> "", SEQNOD15, Convert.DBNull) 'SEQNOD15
                .Parameters.Add("@KID20", SqlDbType.VarChar).Value = If(cvKID20 <> "", cvKID20, Convert.DBNull) 'KID20
                .Parameters.Add("@KID22", SqlDbType.VarChar).Value = If(KID22 <> "", KID22, Convert.DBNull) 'KID22
                .Parameters.Add("@KID25", SqlDbType.VarChar).Value = If(cvKID25 <> "", cvKID25, Convert.DBNull) 'KID25
                '(增加:KID60=@KID60)
                If fg_CAN_USE_1_LID_0 Then .Parameters.Add("@KID60", SqlDbType.VarChar).Value = If(cvKID60 <> "", cvKID60, Convert.DBNull) 'KID60
                .Parameters.Add("@ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
            End With
        Else
            'sql &= " ,KID17,KID19,KID18" & vbCrLf 'sql &= " ,@KID17,@KID19,@KID18" & vbCrLf
            Dim iCmd As New SqlCommand(i_sql, objconn)
            With iCmd
                .Parameters.Clear()
                .Parameters.Add("@PLANID", SqlDbType.VarChar).Value = PlanID
                .Parameters.Add("@COMIDNO", SqlDbType.VarChar).Value = ComIDNO
                .Parameters.Add("@SEQNO", SqlDbType.VarChar).Value = SeqNo
                .Parameters.Add("@KID12", SqlDbType.VarChar).Value = If(KID12 <> "", KID12, Convert.DBNull) 'KID12
                .Parameters.Add("@SEQNOD15", SqlDbType.VarChar).Value = If(SEQNOD15 <> "", SEQNOD15, Convert.DBNull) 'SEQNOD15

                .Parameters.Add("@KID20", SqlDbType.VarChar).Value = If(cvKID20 <> "", cvKID20, Convert.DBNull) 'KID20
                .Parameters.Add("@KID22", SqlDbType.VarChar).Value = If(KID22 <> "", KID22, Convert.DBNull) 'KID22
                .Parameters.Add("@KID25", SqlDbType.VarChar).Value = If(cvKID25 <> "", cvKID25, Convert.DBNull) 'KID25
                '(增加:KID60=@KID60)
                If fg_CAN_USE_1_LID_0 Then .Parameters.Add("@KID60", SqlDbType.VarChar).Value = If(cvKID60 <> "", cvKID60, Convert.DBNull)
                .Parameters.Add("@MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
            End With
        End If
        '若選其他 異動 PLAN_PLANINFO
        'If KID12 = cst_KID12_其他 Then Call sUtl_UpdatePlaninfo(sSearchW)
    End Sub

    'update plan_planinfo class_classinfo 若選其他 異動 PLAN_PLANINFO
    'Sub sUtl_UpdatePlaninfo(ByVal sSearchW As String)
    '    Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
    '    Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
    '    Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")
    '    If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Exit Sub
    '    'select * from Key_TrainType where TMID=554
    '    '(select * from view_traintype where tmid =554)
    '    'select * from ID_GOVCLASSCAST2 where GCID2=1160
    '    '(1160 1 99 其他(其他) NULL 其他(其他) )

    '    If Not TIMS.OpenDbConn(objconn) Then Return

    '    Dim u_sql As String = ""
    '    u_sql &= " UPDATE CLASS_CLASSINFO "
    '    u_sql &= If(iPYNum >= 3, " SET TMID=754", " SET TMID=554")
    '    u_sql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
    '    Dim uCmd1 As New SqlCommand(u_sql, objconn)
    '    With uCmd1
    '        .Parameters.Clear()
    '        .Parameters.Add("PLANID", SqlDbType.VarChar).Value = PlanID
    '        .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = ComIDNO
    '        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = SeqNo
    '        DbAccess.ExecuteNonQuery(uCmd1.CommandText, objconn, uCmd1.Parameters)
    '    End With

    '    'select * from ID_GOVCLASSCAST3 where GCID3=2154
    '    'select * from V_GOVCLASSCAST2 where GCID2 =1160
    '    'select * from view_traintype where tmid =554
    '    Dim u_sql2 As String = ""
    '    u_sql2 &= " UPDATE PLAN_PLANINFO "
    '    u_sql2 &= If(iPYNum >= 3, " SET TMID=754,GCID2=NULL,GCID3=2154", " SET TMID=554,GCID2=1160,GCID3=NULL")
    '    u_sql2 &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
    '    Dim uCmd2 As New SqlCommand(u_sql2, objconn)
    '    With uCmd2
    '        .Parameters.Clear()
    '        .Parameters.Add("PLANID", SqlDbType.VarChar).Value = PlanID
    '        .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = ComIDNO
    '        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = SeqNo
    '        DbAccess.ExecuteNonQuery(uCmd2.CommandText, objconn, uCmd2.Parameters)
    '    End With
    'End Sub

    '清除
    Sub sUtl_ClearPanelEdit1()
        Call ClearTrainDesc()

        hidPlanID.Value = ""
        hidComIDNO.Value = ""
        hidSeqNO.Value = ""

        'ddlKID04.SelectedIndex = -1
        'ddlKID10.SelectedIndex = -1
        ddlDEPOT12.SelectedIndex = -1 '12 課程分類 '19:其他類
        ddlDEPOT15.SelectedIndex = -1 '15 轄區重點產業
        'ddlKID06.SelectedIndex = -1 '06	四大新興智慧型產業
        'ddlKID10.SelectedIndex = -1 '10/16	重點服務業
        'ddlKID17.SelectedIndex = -1 '17	政府政策性產業
        'ddlKID18.SelectedIndex = -1 '18	新南向政策
        'ddlKID14.SelectedIndex = -1

        lbYears.Text = ""
        lbDistName.Text = ""
        lbOrgName.Text = ""
        lbClassName.Text = ""
        'lbPSNO28.Text = ""
        lbOCID.Text = ""
        lbSFTDate.Text = ""
        lbTHours.Text = ""
        lbJobName.Text = ""
        lbGovClassN.Text = ""
    End Sub

    '顯示
    Sub sUtl_ShowPanelEdit1(ByVal sSearchW As String)
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        '修改 'sSearchW = e.CommandArgument
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")
        Dim sDistID As String = TIMS.GetMyValue(sSearchW, "DistID")

        Dim GCODE31 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "GCODE31"))
        Dim KID12 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "KID12")) '19:其他類
        If KID12 = "" Then KID12 = GCODE31
        Dim SEQNOD15 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "SEQNOD15"))
        Dim cvKID20 As String = TIMS.CombiSQM2IN(TIMS.GetMyValue(sSearchW, "KID20"), True)
        Dim cvKID25 As String = TIMS.CombiSQM2IN(TIMS.GetMyValue(sSearchW, "KID25"), True)
        Dim KID22 As String = TIMS.ClearSQM(TIMS.GetMyValue(sSearchW, "KID22"))
        Dim cvKID60 As String = TIMS.CombiSQM2IN(TIMS.GetMyValue(sSearchW, "KID60"), True)

        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Exit Sub

        Dim drS2 As DataRow = search2(PlanID, ComIDNO, SeqNo)
        If drS2 Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        lbYears.Text = drS2("Years").ToString
        lbDistName.Text = drS2("DistName").ToString
        lbOrgName.Text = drS2("OrgName").ToString
        lbClassName.Text = drS2("ClassName").ToString
        'lbPSNO28.Text = Convert.ToString(dr("PSNO28"))
        lbOCID.Text = Convert.ToString(drS2("OCID"))
        lbSFTDate.Text = String.Concat(drS2("STDate"), "~", drS2("FDDate")) 'Convert.ToString(dr("STDate")) & "~" & Convert.ToString(dr("FDDate"))
        lbTHours.Text = drS2("THours").ToString
        lbJobName.Text = drS2("JobName").ToString
        lbGovClassN.Text = drS2("GovClassN").ToString

        hidPlanID.Value = PlanID
        hidComIDNO.Value = ComIDNO
        hidSeqNO.Value = SeqNo

        Call ShowTrainDesc(PlanID, ComIDNO, SeqNo)

        Dim dr2 As DataRow = TIMS.GET_PLANDEPOT(PlanID, ComIDNO, SeqNo, objconn)
        If dr2 Is Nothing Then Return

        '轄區重點產業(排除停用)(轄區)
        Call TIMS.CHG_ddlDEPOT15(ddlDEPOT15, sDistID, sm.UserInfo.TPlanID, objconn)
        '轄區重點產業(排除停用)(轄區) 15/21 搜尋舊值顯示
        If SEQNOD15 <> "" Then
            TIMS.CHG_ddlDEPOT15_21(ddlDEPOT15, sm.UserInfo.DistID, objconn, SEQNOD15)
            Common.SetListItem(ddlDEPOT15, SEQNOD15)
        End If

        If KID12 <> "" Then Common.SetListItem(ddlDEPOT12, KID12)
        'If KID06 <> "" Then Common.SetListItem(ddlKID06, KID06)
        'If KID10 <> "" Then Common.SetListItem(ddlKID10, KID10)
        'If KID17 <> "" Then Common.SetListItem(ddlKID17, KID17)
        'If KID19 <> "" Then Common.SetListItem(ddlKID19, KID19)
        'If KID18 <> "" Then Common.SetListItem(ddlKID18, KID18)
        '2019年啟用 work2019x01:2019 政府政策性產業
        Call TIMS.SetCblValue(CBLKID20_1, cvKID20)
        Call TIMS.SetCblValue(CBLKID20_2, cvKID20)
        Call TIMS.SetCblValue(CBLKID20_3, cvKID20)
        Call TIMS.SetCblValue(CBLKID20_4, cvKID20)
        Call TIMS.SetCblValue(CBLKID20_5, cvKID20)
        Call TIMS.SetCblValue(CBLKID20_6, cvKID20)
        '進階政策性產業類別
        'Dim KID22 As String = Convert.ToString(dr2("KID22"))
        'Call TIMS.SetCblValue(CBLKID22, KID22) 'Call TIMS.SetCblValue(CBLKID22B, KID22)
        '進階政策性產業類別
        If trKID25.Visible Then
            Call TIMS.SetCblValue(CBLKID22B, KID22)
            'Dim v_KID22 As String = Convert.ToString(dr2("KID22")) 'Call TIMS.SetCblValue(CBLKID22B, v_KID22)
        Else
            Call TIMS.SetCblValue(CBLKID22, KID22)
            'Dim v_KID22 As String = Convert.ToString(dr2("KID22")) 'Call TIMS.SetCblValue(CBLKID22, v_KID22)
        End If

        Call TIMS.SetCblValue(CBLKID60, cvKID60)

        Call TIMS.SetCblValue(CBLKID25_1, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_2, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_3, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_4, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_5, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_6, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_7, cvKID25)
        Call TIMS.SetCblValue(CBLKID25_8, cvKID25)

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub

        Dim fg_SHOW_2025_1 As Boolean = TIMS.SHOW_2025_1(sm)
        trKID20.Visible = If(fg_SHOW_2025_1, False, True)
        trKID25.Visible = If(fg_SHOW_2025_1, True, False)

        Select Case e.CommandName
            Case "CHKOK"
                '確認'SAVE
                Dim sSearchW As String = e.CommandArgument
                Call sUtl_SAVEDATA1_ChkOk(sSearchW)
                If gflag_can_continue_sch Then Call Search1()

            Case "Edit"
                '修改
                Dim sSearchW As String = e.CommandArgument
                Call sUtl_ClearPanelEdit1()
                Call sUtl_ShowPanelEdit1(sSearchW)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim lbD12KNAME As Label = e.Item.FindControl("lbD12KNAME")
                Dim lbD15KNAME As Label = e.Item.FindControl("lbD15KNAME")
                'Dim lbD06KNAME As Label = e.Item.FindControl("lbD06KNAME")
                'Dim lbD10KNAME As Label = e.Item.FindControl("lbD10KNAME")
                'Dim lbD17KNAME As Label = e.Item.FindControl("lbD17KNAME")
                'Dim lbD19KNAME As Label = e.Item.FindControl("lbD19KNAME")
                'Dim lbD18KNAME As Label = e.Item.FindControl("lbD18KNAME")
                Dim lbD20KNAME As Label = e.Item.FindControl("lbD20KNAME")

                Dim BtnCHKOK As Button = e.Item.FindControl("BtnCHKOK")
                Dim BtnEdit As Button = e.Item.FindControl("BtnEdit")

                '序號
                'e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Const Cst_emptyTxt As String = "無"

                'lbD12KNAME.Text = If(Convert.ToString(drv("D12KNAME")) <> "", Convert.ToString(drv("D12KNAME")), Cst_emptyTxt)
                lbD15KNAME.Text = If(Convert.ToString(drv("D15KNAME")) <> "", Convert.ToString(drv("D15KNAME")), Cst_emptyTxt)
                'lbD06KNAME.Text = If(Convert.ToString(drv("D06KNAME")) <> "", Convert.ToString(drv("D06KNAME")), Cst_emptyTxt)
                'lbD10KNAME.Text = If(Convert.ToString(drv("D10KNAME")) <> "", Convert.ToString(drv("D10KNAME")), Cst_emptyTxt)

                'lbD17KNAME.Visible = False
                'lbD19KNAME.Visible = False
                'lbD17KNAME.Text = ""
                'lbD19KNAME.Text = ""

                'lbD20KNAME.Visible = False
                lbD20KNAME.Text = If(Convert.ToString(drv("D20KNAME")) <> "", Convert.ToString(drv("D20KNAME")), Cst_emptyTxt)
                '2019年啟用 work2019x01:2019 政府政策性產業
                'If Hid_Years2019.Value = TIMS.cst_YES Then
                '    lbD17KNAME.Visible = False
                '    lbD19KNAME.Visible = False
                '    lbD20KNAME.Visible = True
                '    'Dim vKID20_NAME As String = TIMS.Get_DEPOTNAME("20", Convert.ToString(drv("KID20")), objconn)
                '    lbD20KNAME.Text = If(Convert.ToString(drv("D20KNAME")) <> "", Convert.ToString(drv("D20KNAME")), Cst_emptyTxt)
                'Else
                '    If Hid_Years2018.Value = TIMS.cst_YES Then
                '        lbD19KNAME.Visible = True
                '        lbD19KNAME.Text = If(Convert.ToString(drv("D19KNAME")) <> "", Convert.ToString(drv("D19KNAME")), Cst_emptyTxt)
                '    Else
                '        lbD17KNAME.Visible = True
                '        lbD17KNAME.Text = If(Convert.ToString(drv("D17KNAME")) <> "", Convert.ToString(drv("D17KNAME")), Cst_emptyTxt)
                '    End If
                '    lbD18KNAME.Text = If(Convert.ToString(drv("D18KNAME")) <> "", Convert.ToString(drv("D18KNAME")), Cst_emptyTxt)
                'End If

                Dim cmdArg As String = ""
                cmdArg &= "&PlanID=" & Convert.ToString(drv("PlanID"))
                cmdArg &= "&ComIDNO=" & Convert.ToString(drv("ComIDNO"))
                cmdArg &= "&SeqNo=" & Convert.ToString(drv("SeqNo"))
                cmdArg &= "&DistID=" & Convert.ToString(drv("DistID"))
                'cmdArg &= "&KID04=" & Convert.ToString(drv("KID04"))
                cmdArg &= "&GCODE31=" & Convert.ToString(drv("GCODE31"))
                cmdArg &= "&KID12=" & Convert.ToString(drv("KID12"))
                cmdArg &= "&SEQNOD15=" & Convert.ToString(drv("SEQNOD15"))
                '2019年啟用 work2019x01:2019 政府政策性產業
                cmdArg &= "&KID20=" & Convert.ToString(drv("KID20"))
                cmdArg &= "&KID22=" & Convert.ToString(drv("KID22"))
                cmdArg &= "&KID60=" & Convert.ToString(drv("KID60"))
                cmdArg &= "&KID25=" & Convert.ToString(drv("KID25"))

                'cmdArg &= "&KID06=" & Convert.ToString(drv("KID06"))
                'cmdArg &= "&KID10=" & Convert.ToString(drv("KID10"))
                'cmdArg &= "&KID17=" & Convert.ToString(drv("KID17"))
                'cmdArg &= "&KID19=" & Convert.ToString(drv("KID19"))
                'cmdArg &= "&KID18=" & Convert.ToString(drv("KID18"))

                BtnCHKOK.CommandArgument = cmdArg
                BtnEdit.CommandArgument = cmdArg

                BtnEdit.Enabled = True
                If Convert.ToString(drv("AppResult")) <> "Y" Then '未確認
                    BtnCHKOK.Enabled = True '使用確認鈕
                Else
                    BtnCHKOK.Enabled = False '停止確認鈕
                    'BtnCHKOK.CssClass = ""
                    TIMS.Tooltip(BtnCHKOK, "已確認")
                End If

        End Select
    End Sub

    '儲存
    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Dim v_ddlDEPOT12 As String = TIMS.GetListValue(ddlDEPOT12)
        Dim v_ddlDEPOT15 As String = TIMS.GetListValue(ddlDEPOT15)

        Dim cv_KID20 As String = GET_KID20_VAL() '2019 政府政策性產業
        Dim cv_KID25 As String = GET_KID25_VAL() '2025 政府政策性產業

        'Dim v_CBLKID22 As String = TIMS.GetCblValue(CBLKID22)
        'Dim v_CBLKID22B As String = TIMS.GetCblValue(CBLKID22B)
        Dim v_CBLKID60 As String = TIMS.GetCblValue(CBLKID60)

        Dim v_CBLKID22 As String = If(trKID25.Visible, TIMS.GetCblValue(CBLKID22B), TIMS.GetCblValue(CBLKID22)) 'v_CBLKID22

        Dim cmdArg As String = ""
        cmdArg &= "&PlanID=" & hidPlanID.Value
        cmdArg &= "&ComIDNO=" & hidComIDNO.Value
        cmdArg &= "&SeqNo=" & hidSeqNO.Value
        'cmdArg &= "&KID04=" & ddlKID04.SelectedValue
        cmdArg &= "&KID12=" & v_ddlDEPOT12 'ddlDEPOT12.SelectedValue '課程分類
        cmdArg &= "&SEQNOD15=" & v_ddlDEPOT15 'ddlDEPOT15.SelectedValue 'SEQNOD15
        cmdArg &= "&KID20=" & cv_KID20 'TIMS.EncryptAes(vKID20) 
        cmdArg &= "&KID25=" & cv_KID25
        cmdArg &= "&KID22=" & v_CBLKID22
        cmdArg &= "&KID60=" & v_CBLKID60

        'cmdArg &= "&KID06=" & v_ddlKID06 'ddlKID06.SelectedValue
        'cmdArg &= "&KID10=" & v_ddlKID10 'ddlKID10.SelectedValue
        'Dim vKID17 As String = ""
        'Dim vKID19 As String = ""
        'If ddlKID17.Visible Then vKID17 = v_ddlKID17 'ddlKID17.SelectedValue
        'If ddlKID19.Visible Then vKID19 = v_ddlKID19 'ddlKID19.SelectedValue
        'cmdArg &= "&KID17=" & vKID17 'ddlKID17.SelectedValue
        'cmdArg &= "&KID19=" & vKID19 'ddlKID19.SelectedValue
        'cmdArg &= "&KID18=" & v_ddlKID18 'ddlKID18.SelectedValue
        '2019年啟用 work2019x01:2019 政府政策性產業

        Call sUtl_SAVEDATA1_ChkOk(cmdArg)
        If gflag_can_continue_sch Then Call Search1()
    End Sub

    '回上一頁
    Private Sub btnBack1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉
    End Sub

    ''' <summary>'2019年啟用 work2019x01:2019 政府政策性產業 (取得checkboxValue)</summary>
    ''' <returns></returns>
    Function GET_KID20_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function

    ''' <summary>取值-啟用-2025 政府政策性產業</summary>
    ''' <returns></returns>
    Function GET_KID25_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID25_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_7)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_8)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function

    '#Region "SSS"

    '課程大綱 清理
    Sub ClearTrainDesc()
        Dim sql As String = " SELECT * FROM dbo.PLAN_TRAINDESC WHERE 1<>1"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        dt.DefaultView.Sort = "STrainDate asc,PName asc,PTDID asc"
        dt = TIMS.dv2dt(dt.DefaultView)
        With Datagrid3
            .DataSource = dt
            .DataKeyField = "PTDID"
            .DataBind()
        End With
    End Sub

    '課程大綱
    Sub ShowTrainDesc(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String)
        Dim parms As New Hashtable From {{"PLANID", PlanID}, {"COMIDNO", ComIDNO}, {"SEQNO", SeqNo}}

        Dim sSql As String = ""
        sSql &= " SELECT PTDID ,PLANID ,COMIDNO ,SEQNO" & vbCrLf
        sSql &= " ,PNAME ,PHOUR,PCONT" & vbCrLf
        sSql &= " ,STRAINDATE ,ETRAINDATE ,TRAINDEP" & vbCrLf
        sSql &= " ,CLASSIFICATION1 ,CLASSIFICATION2 ,PTID ,TECHID" & vbCrLf
        sSql &= " FROM dbo.PLAN_TRAINDESC "
        sSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, parms)
        dt.DefaultView.Sort = "STrainDate asc,PName asc,PTDID asc"
        dt = TIMS.dv2dt(dt.DefaultView)
        Datagrid3Table.Visible = (dt IsNot Nothing AndAlso dt.Rows.Count > 0)
        With Datagrid3
            .DataSource = dt
            .DataKeyField = "PTDID"
            .DataBind()
        End With
    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'SHOW
                Dim drv As DataRowView = e.Item.DataItem
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel")
                Dim lbContText As Label = e.Item.FindControl("lbContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1")

                PHourLabel.Text = Convert.ToString(drv("PHOUR")) '時數
                lbContText.Text = Convert.ToString(drv("PCONT")) '內容
                If Convert.ToString(drv("CLASSIFICATION1")) <> "" Then
                    Common.SetListItem(drpClassification1, drv("CLASSIFICATION1").ToString)
                End If
        End Select
    End Sub

    '2019年啟用 work2019x01:2019 政府政策性產業
    Function CHK_KID20_VAL() As String
        Dim Errmsg As String = ""
        '「5+2」產業創新計畫 5+2產業'【台灣AI行動計畫】 KID='08''【數位國家創新經濟發展方案】KID='09''【國家資通安全發展方案】KID='10''【前瞻基礎建設計畫】'【新南向政策】KID='19'
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "「5+2」產業創新計畫，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【台灣AI行動計畫】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【數位國家創新經濟發展方案】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【國家資通安全發展方案】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【前瞻基礎建設計畫】，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【新南向政策】，不可複選(僅可單一勾選)" & vbCrLf
        Return Errmsg
    End Function

    ''' <summary> 檢核-2025年啟用-2025 政府政策性產業</summary>
    ''' <returns></returns>
    Function CHK_KID25_VAL_OTH() As String
        Dim Errmsg As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID25_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "亞洲矽谷，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "重點產業，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "台灣AI行動計畫，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "智慧國家方案，不可複選(僅可單一勾選)" & vbCrLf

        tmp01 = TIMS.GetCblValue(CBLKID25_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "國家人才競爭力躍升方案，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "新南向政策，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_7)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "AI加值應用，不可複選(僅可單一勾選)" & vbCrLf
        tmp01 = TIMS.GetCblValue(CBLKID25_8)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "職場續航，不可複選(僅可單一勾選)" & vbCrLf

        tmp01 = TIMS.GetCblValue(CBLKID22B)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= "【進階政策性產業類別】，不可複選(僅可單一勾選)" & vbCrLf
        Return Errmsg
    End Function

    ''' <summary>匯出產業別分類</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExp1_Click(sender As Object, e As EventArgs) Handles btnExp1.Click
        Call sExprot2()
    End Sub

    ''' <summary>匯出產業別分類</summary>
    Sub sExprot2()
        '匯出excel 
        Dim sFileName1 As String = String.Concat("產業別分類", TIMS.GetDateNo2(4))

        '設定 ViewState Value
        Call sUtl_SetViewStateValue()
        Dim sql23 As String = sUtl_GetSQL2(cst_Exp_iType_3匯出欄位, "")

        Dim dtXls As DataTable = DbAccess.GetDataTable(sql23, objconn)

        If dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        '匯出excel 
        'Dim filename As String = "ExpFile" & TIMS.GetRnd6Eng
        'Dim strfileext As String = ".xls"
        'HttpContext.Current.Response.ContentType = "application/vnd.ms-excel"
        'HttpContext.Current.Response.HeaderEncoding = System.Text.Encoding.GetEncoding("big5")
        'HttpContext.Current.Response.AppendHeader("Content-Disposition", "attachment; filename=" & filename & strfileext)
        'HttpContext.Current.Response.Write("<meta http-equiv=Content-Type content=text/html;charset=utf-8>")

        '先把分頁關掉
        Dim GridView1 As New GridView
        GridView1.AllowPaging = False
        GridView1.DataSource = dtXls
        GridView1.DataBind()

        'Get the HTML for the control.
        Dim tw As IO.StringWriter = New IO.StringWriter()
        Dim hw As HtmlTextWriter = New HtmlTextWriter(tw)
        GridView1.RenderControl(hw)
        Dim strHTML As String = tw.ToString()

        'Dim v_ExpType As String = TIMS.GetListValue(RBListExpType)
        'parmsExp.Add("strSTYLE", strSTYLE)
        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)}, 'EXCEL/PDF/ODS
            {"FileName", sFileName1},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub Btn_XlsImport1_Click(sender As Object, e As EventArgs) Handles Btn_XlsImport1.Click
        Dim sErrMsg2 As String = ""
        Call ChkImpData1(sErrMsg2)
        If sErrMsg2 <> "" Then
            Common.MessageBox(Me, sErrMsg2)
            Exit Sub
        End If

        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, fg_File1_xls, fg_File1_ods)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If fg_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf fg_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If

        Const Cst_FileSavePath As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        '匯入xls or ods
        Call Utl_ImpData1(FullFileName1)
    End Sub

    Private Sub ChkImpData1(ByRef sErrMsg2 As String)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                If RIDValue.Value = "" Then
                    sErrMsg2 &= "匯入請先選擇訓練機構!" & vbCrLf
                    Return
                ElseIf drR Is Nothing Then
                    sErrMsg2 &= "匯入訓練機構選擇有誤，請重新選擇!" & vbCrLf
                    Return
                ElseIf drR IsNot Nothing AndAlso Convert.ToString(drR("ORGLEVEL")) <> "2" Then
                    sErrMsg2 &= "匯入訓練機構選擇有誤，請重新選擇!!" & vbCrLf
                    Return
                End If
                Select Case sm.UserInfo.LID
                    Case 1
                        If drR IsNot Nothing AndAlso (Convert.ToString(drR("DISTID")) <> sm.UserInfo.DistID) Then
                            sErrMsg2 &= "匯入訓練機構轄區有誤，請重新選擇!!!" & vbCrLf
                            Return
                        End If
                    Case 2
                        If drR IsNot Nothing AndAlso (Convert.ToString(drR("RID")) <> sm.UserInfo.RID) Then
                            sErrMsg2 &= "匯入訓練機構選擇有誤，請重新選擇!!!" & vbCrLf
                            Return
                        End If
                End Select
        End Select

    End Sub

    ''' <summary>'匯入xls or ods</summary>
    ''' <param name="fullFileName1"></param>
    Private Sub Utl_ImpData1(fullFileName1 As String)
        '上傳檔案
        File1.PostedFile.SaveAs(fullFileName1)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (fg_File1_xls) Then
            'Const cst_FirstCol1 As String = "課程申請流水號"
            Const cst_FirstCol1 As String = "班別代碼"
            dt_xls = TIMS.GetDataTable_XlsFile(fullFileName1, "", Reason, cst_FirstCol1)
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        ElseIf (fg_File1_ods) Then
            dt_xls = TIMS.GetDataTable_ODSFile(fullFileName1)
        End If
        'TIMS.LOG.DebugFormat("#fg_File1_xls: {0}", fg_File1_xls)
        'TIMS.LOG.DebugFormat("#fg_File1_ods: {0}", fg_File1_ods)

        '刪除檔案 'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(fullFileName1)
        Reason = TIMS.Chk_DTXLS1(dt_xls, fg_File1_xls, fg_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '儲存錯誤資料的DataTable
        Dim dtWrong As New DataTable
        Dim drWrong As DataRow = Nothing
        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        Dim i_sSql As String = ""
        i_sSql &= " INSERT INTO PLAN_DEPOT(PLANID,COMIDNO,SEQNO ,APPRESULT,MODIFYACCT,MODIFYDATE,KID60)" & vbCrLf
        i_sSql &= " VALUES (@PLANID,@COMIDNO,@SEQNO ,'Y',@MODIFYACCT,GETDATE(),@KID60)" & vbCrLf

        Dim u_sSql As String = ""
        u_sSql &= " UPDATE PLAN_DEPOT" & vbCrLf
        u_sSql &= " SET APPRESULT='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE(),KID60=@KID60" & vbCrLf
        u_sSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf

        Dim s_sSql As String = ""
        s_sSql &= " SELECT 1 FROM PLAN_DEPOT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf

        Dim lstImp1s As List(Of String) = New List(Of String)()
        Dim iRowIndex As Integer = 1
        For Each dr1 As DataRow In dt_xls.Rows
            Reason = ""
            Dim colArray As Array = dr1.ItemArray
            '轉換正確欄位值/'檢查正確欄位值
            Reason += CheckImportData1(colArray, lstImp1s)

            '通過檢查，開始輸入資料---------------------Start
            If Reason = "" Then
                Dim sParms As New Hashtable From {{"PLANID", aPLANID}, {"COMIDNO", aCOMIDNO}, {"SEQNO", aSEQNO}}
                Dim dtMain As DataTable = DbAccess.GetDataTable(s_sSql, objconn, sParms)
                If dtMain.Rows.Count = 0 Then
                    'iParms.Add("APPRESULT", APPRESULT)
                    'iParms.Add("MODIFYDATE", MODIFYDATE)
                    Dim iParms As New Hashtable From {{"PLANID", aPLANID}, {"COMIDNO", aCOMIDNO}, {"SEQNO", aSEQNO},
                        {"MODIFYACCT", sm.UserInfo.UserID}, {"KID60", If(aKID60 <> "", aKID60, Convert.DBNull)}}
                    DbAccess.ExecuteNonQuery(i_sSql, objconn, iParms)
                Else
                    'uParms.Add("APPRESULT", APPRESULT)
                    'uParms.Add("MODIFYDATE", MODIFYDATE)
                    Dim uParms As New Hashtable From {{"PLANID", aPLANID}, {"COMIDNO", aCOMIDNO}, {"SEQNO", aSEQNO},
                        {"MODIFYACCT", sm.UserInfo.UserID}, {"KID60", If(aKID60 <> "", aKID60, Convert.DBNull)}}
                    DbAccess.ExecuteNonQuery(u_sSql, objconn, uParms)
                End If
            Else
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                drWrong("Index") = iRowIndex
                drWrong("Reason") = Reason
            End If

            iRowIndex += 1
        Next

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""
        explain = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf

        explain2 = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        '開始判別欄位存入------------   End
        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, explain)
            Exit Sub
        End If

        Session("MyWrongTable") = dtWrong
        Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視失敗原因?')){window.open('TC_04_004_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
    End Sub

    Private Function CheckImportData1(colArray As Array, lstImp1s As List(Of String)) As String
        Const cst_必須填寫 As String = "必須填寫"
        Dim Reason As String = ""
        If colArray.Length < cst_iFiledColumnNum Then
            'Reason += "欄位數量不正確(應該為" & cst_filedNum & "個欄位)<BR>"
            Reason &= "欄位對應有誤<BR>"
            Reason &= "請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        aOCID = TIMS.ClearSQM(colArray(cst_aOCID)) '班別代碼／課程申請流水號
        aKID60 = TIMS.ClearSQM(colArray(cst_aKID60)) '產業別代號

        If aOCID = "" Then Reason += cst_必須填寫 & "班別代碼<Br>"
        If aKID60 = "" Then Reason += cst_必須填寫 & "產業別代號<Br>"
        If Not TIMS.IsNumeric2(aOCID) Then Reason += String.Concat("班別代碼 必須為正確的格式! ", aOCID, "<BR>")
        'If aOCID <> "" AndAlso (aOCID.Length <= 8 OrElse aOCID.Length >= 12) Then
        '    Reason += String.Concat("課程申請流水號 必須為正確的格式 ", aOCID, "<BR>")
        'End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(aOCID, objconn)
        If drCC Is Nothing Then
            Reason += String.Concat("班別代碼 課程查無資料! ", aOCID, "<BR>")
        ElseIf drCC IsNot Nothing AndAlso Convert.ToString(drCC("TPLANID")) <> TIMS.Cst_TPlanID28 Then
            Reason += String.Concat("班別代碼 查詢資料有誤! ", aOCID, "<BR>")
        End If
        'Dim drPP As DataRow = TIMS.Get_PSNO28_row(aPSNO28, objconn)
        'If drPP Is Nothing OrElse Convert.ToString(drPP("TPLANID")) <> TIMS.Cst_TPlanID28 Then
        '    Reason += String.Concat("依課程申請流水號 課程查無資料! ", aOCID, "<BR>")
        'ElseIf drPP IsNot Nothing AndAlso Convert.ToString(drPP("TPLANID")) <> TIMS.Cst_TPlanID28 Then
        '    Reason += String.Concat("依課程申請流水號 查詢資料有誤! ", aOCID, "<BR>")
        'End If
        If Reason <> "" Then Return Reason

        aPLANID = Convert.ToString(drCC("PLANID"))
        aCOMIDNO = Convert.ToString(drCC("COMIDNO"))
        aSEQNO = Convert.ToString(drCC("SEQNO"))

        aKID60 = TIMS.CombiSQM4IN(aKID60, "00")
        TIMS.SetCblValue(CBLKID60, aKID60)
        aKID60 = TIMS.GetCblValue(CBLKID60)
        Dim sErrMsg2 As String = TIMS.CHK_KID60_VAL(CBLKID60)
        If sErrMsg2 <> "" Then Reason = sErrMsg2
        If Reason <> "" Then Return Reason

        'If lstImp1s.Contains(aOCID) Then Reason += String.Concat("(匯入重複)課程申請流水號 已重複!", aOCID, "<BR>")
        If lstImp1s.Contains(aOCID) Then Reason += String.Concat("(匯入重複)班別代碼 已重複!", aOCID, "<BR>")
        If Reason <> "" Then Return Reason
        lstImp1s.Add(aOCID)

        Return Reason
    End Function

    Protected Sub Datagrid3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Datagrid3.SelectedIndexChanged

    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class

