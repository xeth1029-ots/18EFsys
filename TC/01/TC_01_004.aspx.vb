Partial Class TC_01_004
    Inherits AuthBasePage

    '開班資料查詢
    '一般 TIMS 'TC_01_004_add
    Const cst_TIMS_EDITASPX1 As String = "TC_01_004_add.aspx?"
    '產投、充飛 'TC_01_004_BusAdd
    Const cst_TIMS28_EDITASPX1 As String = "TC_01_004_BusAdd.aspx?"
    '刪除
    Const cst_TIMS_DEL_ASPX1 As String = "TC_01_004_del.aspx?"

    'sUrl = "TC_01_004_add.aspx?ProcessType=Update&" & e.CommandArgument' '一般 TIMSTC_01_004_add
    'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then sUrl = "TC_01_004_BusAdd.aspx?" & e.CommandArgument '產投、充飛'TC_01_004_BusAdd
    'Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
    'Dim s_TargetTable As String = "CLASS_CLASSINFO"
    'Dim s_FuncPath As String = "/TC/06/TC_06_001_chk"
    'Dim s_WHERE As String = "" 'insert省略/update必要'String.Format(cst_fWHERE, pkVALUE)

    Dim vsMsg2 As String = "" '確認機構是否為黑名單
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection = Nothing

    Const cst_ClassSearchStr As String = "ClassSearchStr" ' Session(cst_ClassSearchStr) = Nothing

    Const cst_dg_SelectAll As Integer = 0
    Const cst_dg_序號 As Integer = 1
    Const cst_dg_訓練機構 As Integer = 2
    Const cst_dg_班別代碼 As Integer = 3 '(自辦)
    Const cst_dg_開結訓日 As Integer = 4

    Const cst_dg_班別名稱 As Integer = 5
    Const cst_dg_訓練職類 As Integer = 6 '(自辦)
    Const cst_dg_通俗職類 As Integer = 7 '(自辦)

    Const cst_dg_上架日期 As Integer = 8 '(產投)
    Const cst_dg_報名開始日期 As Integer = 9 '(產投)
    Const cst_dg_報名結束日期 As Integer = 10 '(產投)
    'Const cst_dg_功能 As Integer = 11

    'Const cst_dg_序號 As Integer = 1
    'Const cst_dg_管控單位 As Integer = 2 '管控單位
    'Const cst_dg_ClassID4 As Integer = 4 '班別代碼
    'Const cst_dg_ClassNum5 As Integer = 5 '班數
    'Const cst_dg_SFTdate6 As Integer = 6 '開結訓日
    'Const cst_dg_ClassName7 As Integer = 7 '班別名稱
    'Const cst_dg_TPropertyID10 As Integer = 10 '訓練性質
    'Const cst_dg_OnShellDate As Integer = 12 '上架日期

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        iPYNum = TIMS.sUtl_GetPYNum(Me)

        'DG_ClassInfo.Columns(cst_dg_ClassNum5).Visible = False '班數不SHOW (20181002 依照承辦人要求,在職計畫別裡,此欄位用不到)

        'bt_EXPORT.Visible = False '匯出鈕
        LabTPeriod.Text = "訓練時段"
        AppStage.Visible = False '申請階段[停用]
        TPeriod_List.Visible = True '訓練時段[顯示]
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            btnSave.Attributes("onclick") = "return doBatchSet();"
            'LabTPeriod.Visible = False
            'bt_EXPORT.Visible = True '匯出鈕
            LabTPeriod.Text = "申請階段" '將 [訓練時段] 改為 [申請階段]
            AppStage.Visible = True '申請階段[顯示]
            TPeriod_List.Visible = False '訓練時段[停用]

            LabTMID.Text = "訓練業別"
            'LabTPropertyID.Visible = False
            'RB_TPropertyID.Visible = False
            'If Not LabTPropertyID.Visible And Not RB_TPropertyID.Visible And Not LabTPeriod.Visible And Not TPeriod_List.Visible Then trTrainX.Visible = False  'edit，by:20181030
            'trTrainX.Visible = False

            'XXX.Visible = False
            'DG_ClassInfo.Columns(cst_dg_ClassID4).Visible = False '班別代碼不SHOWcst_dg_SelectAll
            '2018-10-11 add 訓練單位人員沒有權限設定上架日期
            DG_ClassInfo.Columns(cst_dg_班別代碼).Visible = False
            Select Case sm.UserInfo.LID
                Case 2
                    DG_ClassInfo.Columns(cst_dg_SelectAll).Visible = False '選取欄SHOW
                    DG_ClassInfo.Columns(cst_dg_上架日期).Visible = False '上架日SHOW
                    DG_ClassInfo.Columns(cst_dg_報名開始日期).Visible = False
                    DG_ClassInfo.Columns(cst_dg_報名結束日期).Visible = False
                Case Else
                    DG_ClassInfo.Columns(cst_dg_訓練職類).Visible = False
                    DG_ClassInfo.Columns(cst_dg_通俗職類).Visible = False
                    DG_ClassInfo.Columns(cst_dg_SelectAll).Visible = True '選取欄SHOW
                    DG_ClassInfo.Columns(cst_dg_上架日期).Visible = True '上架日SHOW
                    DG_ClassInfo.Columns(cst_dg_報名開始日期).Visible = True
                    DG_ClassInfo.Columns(cst_dg_報名結束日期).Visible = True
            End Select
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            DG_ClassInfo.Columns(cst_dg_SelectAll).Visible = False '選取欄不SHOW
            DG_ClassInfo.Columns(cst_dg_上架日期).Visible = False '上架日不SHOW
            DG_ClassInfo.Columns(cst_dg_報名開始日期).Visible = False
            DG_ClassInfo.Columns(cst_dg_報名結束日期).Visible = False
        End If

        '分頁設定 Start
        PageControler1.PageDataGrid = DG_ClassInfo
        '分頁設定 End
        'Dim ProcessType As String = Request("ProcessType")

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not Page.IsPostBack Then
            Call cCreate1()
            'Session(cst_ClassSearchStr) 依查詢值 查詢1次
            Use_SearchStr1()
        End If

        '確認機構是否為黑名單
        'Dim vsMsg2 As String = "" '確認機構是否為黑名單
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            'Button2.Enabled = False
            'TIMS.Tooltip(Button2, vsMsg2)
            Page.RegisterStartupScript("", String.Concat("<script>alert('", vsMsg2, "');</script>"))
        End If
    End Sub

    Sub cCreate1()
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        'End If
        tr_SetOnShellDate1.Visible = False '批次設定上架日
        tr_setSEnterDate.Visible = False
        tr_setFEnterDate.Visible = False
        btnSave.Visible = False
        bt_EXPORT.Visible = False '匯出鈕

        '取得查詢條件
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso AppStage.Visible Then
            AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_AppStage2(AppStage), TIMS.Get_AppStage(AppStage))
        End If

        '訓練時段
        TPeriod_List = TIMS.GET_HOURRAN(TPeriod_List, objconn, sm)
        'RB_TPropertyID = TIMS.Get_TPropertyID(sm, RB_TPropertyID)

        Call TIMS.SUB_SET_HR_MI(OnShellDate_HR, OnShellDate_MI)
        Call TIMS.SUB_SET_HR_MI(SEnterDate_HR, SEnterDate_MI)
        Call TIMS.SUB_SET_HR_MI(FEnterDate_HR, FEnterDate_MI)
    End Sub

    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            isBlack.Value = "Y"
            orgname.Value = sm.UserInfo.OrgName
            'btnAdd.Visible = False
            'Button8.Visible = False
        End If
        Return rst
    End Function

    ''' <summary>'查詢 [SQL]</summary>
    ''' <returns></returns>
    Function Get_Search1dt() As DataTable
        'ByRef rst As DataTable
        Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objconn)

        Dim parms As Hashtable = New Hashtable()
        parms.Clear()

        Dim sqlstr As String = ""
        sqlstr &= " SELECT a.Years" & vbCrLf
        'sqlstr &= " ,a.ClassNum" & vbCrLf'班數
        sqlstr &= " ,b.ClassID ,a.OCID" & vbCrLf
        sqlstr &= " ,a.PlanID ,a.ComIDNO ,a.SeqNO" & vbCrLf
        '訓練機構
        sqlstr &= " ,d.ORGNAME ,pp.CYCLTYPE" & vbCrLf
        sqlstr &= " ,dbo.FN_GET_CLASSCNAME(pp.CLASSNAME,pp.CYCLTYPE) CLASSCNAME" & vbCrLf
        'sqlstr &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sqlstr &= " ,CASE WHEN f.JobID IS NULL THEN f.TrainName ELSE f.JobName END TrainName" & vbCrLf
        sqlstr &= " ,s.CJOB_NAME" & vbCrLf
        '申請階段 
        sqlstr &= " ,pp.AppStage ,dbo.FN_GET_APPSTAGE(pp.AppStage) MyAppStage" & vbCrLf
        '0:職前/ 1:在職/ 2:接受委託
        sqlstr &= " ,a.TPropertyID ,case a.TPropertyID when 0 then '職前' when 1 then '在職' when 2 then '接受委託' end TPropertyID_N" & vbCrLf
        sqlstr &= " ,e.HourRanName" & vbCrLf
        sqlstr &= " ,a.TNUM ,a.THOURS" & vbCrLf
        sqlstr &= " ,FORMAT(a.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sqlstr &= " ,FORMAT(a.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sqlstr &= " ,a.RID" & vbCrLf
        'sqlstr &= " ,c2.OrgName2" & vbCrLf
        sqlstr &= " ,dbo.FN_CYEAR2(ip.YEARS) PLANYEARS" & vbCrLf
        sqlstr &= " ,ip.DISTNAME" & vbCrLf
        sqlstr &= " ,CONVERT(VARCHAR(30) ,a.ONSHELLDATE,120) ONSHELLDATE" & vbCrLf '上架日期
        sqlstr &= " ,CONVERT(VARCHAR(30) ,a.SENTERDATE ,120) SENTERDATE" & vbCrLf '報名開始日期
        sqlstr &= " ,CONVERT(VARCHAR(30) ,a.FENTERDATE ,120) FENTERDATE" & vbCrLf '報名結束日期

        sqlstr &= " FROM dbo.CLASS_CLASSINFO a WITH(NOLOCK)" & vbCrLf
        sqlstr &= " JOIN dbo.ID_CLASS b WITH(NOLOCK) ON a.CLSID = b.CLSID" & vbCrLf
        sqlstr &= " JOIN dbo.VIEW_PLAN ip ON ip.planid = a.planid" & vbCrLf
        sqlstr &= " JOIN dbo.AUTH_RELSHIP c ON c.RID = a.RID" & vbCrLf
        sqlstr &= " JOIN dbo.ORG_ORGINFO d ON c.OrgID = d.OrgID" & vbCrLf
        sqlstr &= " JOIN dbo.PLAN_PLANINFO pp WITH(NOLOCK) ON pp.PlanID=a.PlanID AND pp.ComIDNO=a.ComIDNO AND pp.SeqNO=a.SeqNO" & vbCrLf
        'sqlstr &= " LEFT JOIN MVIEW_RELSHIP23 c2 ON c2.RID3 = a.RID" & vbCrLf
        sqlstr &= " LEFT JOIN dbo.KEY_HOURRAN e ON a.TPeriod = e.HRID" & vbCrLf
        sqlstr &= " LEFT JOIN dbo.KEY_TRAINTYPE f ON a.TMID = f.TMID" & vbCrLf
        sqlstr &= " LEFT JOIN dbo.SHARE_CJOB s ON s.CJOB_UNKEY = a.CJOB_UNKEY" & vbCrLf
        sqlstr &= " WHERE 1=1" & vbCrLf
        sqlstr &= " AND ip.TPlanID = @TPlanID AND ip.Years = @Years" & vbCrLf
        parms.Add("TPlanID", sm.UserInfo.TPlanID)
        parms.Add("Years", sm.UserInfo.Years)
        If sm.UserInfo.RID <> "A" Then
            sqlstr &= " AND a.PlanID = @PlanID" & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value <> "" Then
            sqlstr &= " AND a.RID LIKE @RID" & vbCrLf
            parms.Add("RID", RIDValue.Value & "%")
        Else
            sqlstr &= " AND a.RID LIKE @RID" & vbCrLf
            parms.Add("RID", sm.UserInfo.RID & "%")
        End If
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            sqlstr &= " AND a.CJOB_UNKEY = @CJOB_UNKEY" & vbCrLf
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If trainValue.Value <> "" Then
            sqlstr &= " AND a.TMID = @TMID" & vbCrLf
            parms.Add("TMID", trainValue.Value)
        End If

        '申請階段 
        Dim v_AppStage As String = TIMS.GetListValue(AppStage) 'SelectedValue
        If v_AppStage <> "" Then
            sqlstr &= " AND pp.AppStage = @AppStage" & vbCrLf
            parms.Add("AppStage", v_AppStage)
        End If
        Dim v_TPeriod_List As String = TIMS.GetListValue(TPeriod_List) 'SelectedValue
        If v_TPeriod_List <> "" Then
            sqlstr &= " AND a.TPeriod = @TPeriod" & vbCrLf
            parms.Add("TPeriod", v_TPeriod_List)
        End If
        TB_ClassName.Text = TIMS.ClearSQM(TB_ClassName.Text)
        If TB_ClassName.Text <> "" Then
            sqlstr &= " AND a.ClassCName LIKE @ClassCName" & vbCrLf
            parms.Add("ClassCName", "%" & Me.TB_ClassName.Text & "%")
        End If
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        If start_date.Text <> "" Then
            sqlstr &= " AND a.STDate >= @STDate1" & vbCrLf
            parms.Add("STDate1", start_date.Text)
        End If
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        If end_date.Text <> "" Then
            sqlstr &= " AND a.STDate <= @STDate2" & vbCrLf
            parms.Add("STDate2", end_date.Text)
        End If

        TB_cycltype.Text = TIMS.FmtCyclType(TB_cycltype.Text)
        If TB_cycltype.Text <> "" Then
            sqlstr &= " AND pp.CYCLTYPE= @CYCLTYPE" & vbCrLf
            parms.Add("CYCLTYPE", TB_cycltype.Text)
        End If

        '開訓狀態--'不區分/已開訓/未開訓
        Select Case ClassState.SelectedIndex
            Case 1 '已開訓
                sqlstr &= " AND a.STDate <= dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
            Case 2 '未開訓
                sqlstr &= " AND a.STDate > dbo.TRUNC_DATETIME(GETDATE())" & vbCrLf
        End Select

        Dim v_NotOpen As String = TIMS.GetListValue(NotOpen)
        Select Case v_NotOpen 'NotOpen.SelectedValue
            Case "Y", "N"
                sqlstr &= " AND a.NotOpen=@NotOpen" & vbCrLf
                parms.Add("NotOpen", v_NotOpen)
        End Select
        If sm.UserInfo.RID <> "A" Then '非署(局)才有此限制
            If sPlanKind = "1" Then '自辦者只能列出賦予給此帳號的班級
                sqlstr &= " AND a.OCID IN (SELECT OCID FROM dbo.AUTH_ACCRWCLASS WITH(NOLOCK) WHERE Account=@Account)" & vbCrLf
                parms.Add("Account", sm.UserInfo.UserID)
            End If
        End If

        ClassID.Text = TIMS.ClearSQM(ClassID.Text)
        If ClassID.Text <> "" Then
            sqlstr &= " AND (b.ClassID LIKE @ClassID OR ISNULL(b.ClassID2,'') LIKE @ClassID2)" & vbCrLf
            parms.Add("ClassID", ClassID.Text & "%")
            parms.Add("ClassID2", ClassID.Text & "%")
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn, parms)
        Return dt
    End Function

    Sub sSearch1()
        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DG_ClassInfo)

        Dim dt As DataTable = Get_Search1dt() '查詢

        Panel_ClassInfo.Visible = False
        DG_ClassInfo.Visible = False
        msg.Text = "查無資料!!"

        bt_EXPORT.Visible = False '匯出鈕
        tr_SetOnShellDate1.Visible = False '批次設定上架日
        tr_setSEnterDate.Visible = False
        tr_setFEnterDate.Visible = False
        btnSave.Visible = False

        If dt Is Nothing Then Return
        If dt.Rows.Count = 0 Then Return

        Panel_ClassInfo.Visible = True
        DG_ClassInfo.Visible = True
        msg.Text = ""

        '2018 add 產投:顯示批次設定上架日期功能
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            bt_EXPORT.Visible = True '匯出鈕
            '2018-10-11 add訓練單位無批次設定權限
            Select Case sm.UserInfo.LID'階層代碼
                Case 2 '委訓單位
                    'divSetOnShellDate1.Visible = False
                    'btnSave.Visible = False
                Case Else
                    tr_SetOnShellDate1.Visible = True '批次設定上架日
                    tr_setSEnterDate.Visible = True
                    tr_setFEnterDate.Visible = True
                    btnSave.Visible = True
            End Select
            OnShellDate.Text = ""
            TIMS.SET_DateHMC00(OnShellDate_HR, OnShellDate_MI)
            SEnterDate.Text = ""
            TIMS.SET_DateHMC00(SEnterDate_HR, SEnterDate_MI)
            FEnterDate.Text = ""
            TIMS.SET_DateHMC00(FEnterDate_HR, FEnterDate_MI)
        End If

        'PageControler1.SqlString = sqlstr
        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "OCID"
        PageControler1.Sort = "ClassID,CYCLTYPE"
        PageControler1.ControlerLoad()

    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call GetSearchStr()
        Call sSearch1()
    End Sub


    ''' <summary>
    ''' Session(cst_ClassSearchStr) 依查詢值 查詢1次
    ''' </summary>
    Sub Use_SearchStr1()
        If Session(cst_ClassSearchStr) Is Nothing Then Return

        Dim myValue1 As String = Session(cst_ClassSearchStr)
        Session(cst_ClassSearchStr) = Nothing

        center.Text = TIMS.GetMyValue(myValue1, "center")
        RIDValue.Value = TIMS.GetMyValue(myValue1, "RIDValue")
        TB_career_id.Text = TIMS.GetMyValue(myValue1, "TB_career_id")
        trainValue.Value = TIMS.GetMyValue(myValue1, "trainValue")
        jobValue.Value = TIMS.GetMyValue(myValue1, "jobValue")
        txtCJOB_NAME.Text = TIMS.GetMyValue(myValue1, "txtCJOB_NAME")
        cjobValue.Value = TIMS.GetMyValue(myValue1, "cjobValue")
        'Common.SetListItem(RB_TPropertyID, TIMS.GetMyValue(myValue1, "RB_TPropertyID"))
        Common.SetListItem(AppStage, TIMS.GetMyValue(myValue1, "AppStage")) '申請階段 
        Common.SetListItem(TPeriod_List, TIMS.GetMyValue(myValue1, "TPeriod_List")) '訓練時段
        TB_ClassName.Text = TIMS.GetMyValue(myValue1, "TB_ClassName")
        TB_cycltype.Text = TIMS.GetMyValue(myValue1, "TB_cycltype")
        start_date.Text = TIMS.GetMyValue(myValue1, "start_date")
        end_date.Text = TIMS.GetMyValue(myValue1, "end_date")

        Common.SetListItem(ClassState, TIMS.GetMyValue(myValue1, "ClassState"))
        Common.SetListItem(NotOpen, TIMS.GetMyValue(myValue1, "NotOpen"))

        PageControler1.PageIndex = 0
        'PageControler1.PageIndex = TIMS.GetMyValue(myValue1, "PageIndex")
        If TIMS.GetMyValue(myValue1, "Button1") = TIMS.cst_True Then
            Dim MyValue As String = TIMS.GetMyValue(myValue1, "PageIndex")
            If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                MyValue = CInt(MyValue)
                PageControler1.PageIndex = MyValue
            End If
            sSearch1()
        End If
        'If TIMS.GetMyValue(myValue1, "Button1") = "True" Then bt_search_Click(sender, e)
        'Session(cst_ClassSearchStr) = Nothing
    End Sub

    ''' <summary>
    '''  Session(cst_ClassSearchStr) 儲存查詢值
    ''' </summary>
    Sub GetSearchStr()
        'Const cst_ClassSearchStr As String = "ClassSearchStr"
        Session(cst_ClassSearchStr) = Nothing
        Dim myValue As String = ""
        myValue = "center=" & TIMS.ClearSQM(center.Text)
        myValue &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        myValue &= "&TB_career_id=" & TIMS.ClearSQM(TB_career_id.Text)
        myValue &= "&trainValue=" & TIMS.ClearSQM(trainValue.Value)
        myValue &= "&jobValue=" & TIMS.ClearSQM(jobValue.Value)
        myValue &= "&cjobValue=" & TIMS.ClearSQM(cjobValue.Value)
        myValue &= "&txtCJOB_NAME=" & TIMS.ClearSQM(txtCJOB_NAME.Text)
        'myValue &= "&RB_TPropertyID=" & RB_TPropertyID.SelectedValue
        myValue &= "&AppStage=" & TIMS.GetListValue(AppStage) '.SelectedValue '申請階段 
        myValue &= "&TPeriod_List=" & TIMS.GetListValue(TPeriod_List) '.SelectedValue '訓練時段
        myValue &= "&TB_ClassName=" & TIMS.ClearSQM(TB_ClassName.Text)
        myValue &= "&TB_cycltype=" & TIMS.ClearSQM(TB_cycltype.Text)
        myValue &= "&start_date=" & TIMS.ClearSQM(start_date.Text)
        myValue &= "&end_date=" & TIMS.ClearSQM(end_date.Text)

        myValue &= "&ClassState=" & TIMS.GetListValue(ClassState) '.SelectedValue
        myValue &= "&NotOpen=" & TIMS.GetListValue(NotOpen) '.SelectedValue
        If DG_ClassInfo.Visible Then
            myValue &= "&PageIndex=" & DG_ClassInfo.CurrentPageIndex + 1
            myValue &= "&Button1=" & TIMS.cst_True
        Else
            myValue &= "&PageIndex=1"
            myValue &= "&Button1=" & TIMS.cst_False
        End If
        Session(cst_ClassSearchStr) = myValue
    End Sub

    Sub CHK_SAVEDATA1(ByRef errMsg As String)
        Dim str_errMSG_FM1 As String = "第{0}筆班級資料 報名結束日期 最晚可為開訓日後第14天，若為短期班，開訓後，14天內就結訓的班級，報名結束日期 最晚為結訓日期前一天。" & vbCrLf

        Dim blFlag As Boolean = False ' 檢核欄位值
        For i As Int64 = 0 To DG_ClassInfo.Items.Count - 1
            Dim eItem As DataGridItem = DG_ClassInfo.Items(i)
            Dim chkItem As HtmlInputCheckBox = eItem.FindControl("chkItem") '勾選
            Dim hidOCID As HiddenField = eItem.FindControl("hidOCID")
            Dim hidSTDATE As HiddenField = eItem.FindControl("hidSTDATE")
            Dim hidFTDATE As HiddenField = eItem.FindControl("hidFTDATE")

            '批次設定上架日期
            Dim OnShellDate_i As TextBox = eItem.FindControl("OnShellDate_i")
            'Dim imgOnShellDate_i As HtmlImage = eItem.FindControl("imgOnShellDate_i")
            'imgOnShellDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & OnShellDate_i.ClientID & "','','','CY/MM/DD');")
            Dim OnShellDate_HR_i As DropDownList = eItem.FindControl("OnShellDate_HR_i")
            Dim OnShellDate_MI_i As DropDownList = eItem.FindControl("OnShellDate_MI_i")
            Dim hidOnShellDate As HiddenField = eItem.FindControl("hidOnShellDate")
            Dim hidOnShellDate_HR As HiddenField = eItem.FindControl("hidOnShellDate_HR")
            Dim hidOnShellDate_MI As HiddenField = eItem.FindControl("hidOnShellDate_MI")

            Dim SEnterDate_i As TextBox = eItem.FindControl("SEnterDate_i")
            'Dim imgSEnterDate_i As HtmlImage = eItem.FindControl("imgSEnterDate_i")
            'imgSEnterDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & SEnterDate_i.ClientID & "','','','CY/MM/DD');")
            Dim SEnterDate_HR_i As DropDownList = eItem.FindControl("SEnterDate_HR_i")
            Dim SEnterDate_MI_i As DropDownList = eItem.FindControl("SEnterDate_MI_i")
            Dim hidSEnterDate As HiddenField = eItem.FindControl("hidSEnterDate")
            Dim hidSEnterDate_HR As HiddenField = eItem.FindControl("hidSEnterDate_HR")
            Dim hidSEnterDate_MI As HiddenField = eItem.FindControl("hidSEnterDate_MI")

            Dim FEnterDate_i As TextBox = eItem.FindControl("FEnterDate_i")
            'Dim imgFEnterDate_i As HtmlImage = eItem.FindControl("imgFEnterDate_i")
            'imgFEnterDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & FEnterDate_i.ClientID & "','','','CY/MM/DD');")
            Dim FEnterDate_HR_i As DropDownList = eItem.FindControl("FEnterDate_HR_i")
            Dim FEnterDate_MI_i As DropDownList = eItem.FindControl("FEnterDate_MI_i")
            Dim hidFEnterDate As HiddenField = eItem.FindControl("hidFEnterDate")
            Dim hidFEnterDate_HR As HiddenField = eItem.FindControl("hidFEnterDate_HR")
            Dim hidFEnterDate_MI As HiddenField = eItem.FindControl("hidFEnterDate_MI")

            '上架日期 '(OLD)'(NEW)'(NEW/OLD) '取得有效值(若有新值依新VAL，無則用舊VAL)
            Dim oldOnShellDate As String = TIMS.GET_YMDHM1(hidOnShellDate.Value, hidOnShellDate_HR.Value, hidOnShellDate_MI.Value)
            Dim vsOnShellDate As String = TIMS.GET_YMDHM1(OnShellDate_i.Text, TIMS.GetListValue(OnShellDate_HR_i), TIMS.GetListValue(OnShellDate_MI_i))
            Dim v_OnShellDate_i As String = If(vsOnShellDate <> "", vsOnShellDate, oldOnShellDate)

            '報名開始日期'(OLD)'(NEW)'(NEW/OLD) '取得有效值(若有新值依新VAL，無則用舊VAL)
            Dim oldSEnterDate As String = TIMS.GET_YMDHM1(hidSEnterDate.Value, hidSEnterDate_HR.Value, hidSEnterDate_MI.Value)
            Dim vsSEnterDate As String = TIMS.GET_YMDHM1(SEnterDate_i.Text, TIMS.GetListValue(SEnterDate_HR_i), TIMS.GetListValue(SEnterDate_MI_i))
            Dim v_SEnterDate_i As String = If(vsSEnterDate <> "", vsSEnterDate, oldSEnterDate)

            '報名結束日期'(OLD)'(NEW)'(NEW/OLD) '取得有效值(若有新值依新VAL，無則用舊VAL)
            Dim oldFEnterDate As String = TIMS.GET_YMDHM1(hidFEnterDate.Value, hidFEnterDate_HR.Value, hidFEnterDate_MI.Value)
            Dim vsFEnterDate As String = TIMS.GET_YMDHM1(FEnterDate_i.Text, TIMS.GetListValue(FEnterDate_HR_i), TIMS.GetListValue(FEnterDate_MI_i))
            Dim v_FEnterDate_i As String = If(vsFEnterDate <> "", vsFEnterDate, oldFEnterDate)

            '勾選
            If chkItem.Checked Then
                blFlag = True '"尚未勾選班級!!"(有勾選)
                '(_i 正確值比對)
                If v_SEnterDate_i <> "" AndAlso v_OnShellDate_i <> "" Then
                    '有勾選且有填上架日期的資料，再進一步檢核設定結果不得超過報名日期
                    If DateDiff(DateInterval.Minute, CDate(v_SEnterDate_i), CDate(v_OnShellDate_i)) > 0 Then
                        errMsg &= If(errMsg <> "", "<br>", "")
                        errMsg &= String.Format("第{0}筆班級資料 [上架日期]不能晚於[報名開始日期]!", (i + 1).ToString())
                    End If
                End If
                '(_i 正確值比對)
                If v_SEnterDate_i <> "" AndAlso v_FEnterDate_i <> "" Then
                    '有勾選且有填上架日期的資料，再進一步檢核設定結果不得超過報名日期
                    If DateDiff(DateInterval.Day, CDate(v_SEnterDate_i), CDate(v_FEnterDate_i)) <= 0 Then
                        errMsg &= If(errMsg <> "", "<br>", "")
                        errMsg &= String.Format("第{0}筆班級資料 [報名開始日期]不能晚於等於[報名結束日期]!", (i + 1).ToString())
                    End If
                End If
                '(_i 正確值比對)
                If v_FEnterDate_i <> "" AndAlso hidFTDATE.Value <> "" Then
                    If DateDiff(DateInterval.Day, CDate(v_FEnterDate_i), CDate(hidFTDATE.Value)) < 0 Then
                        errMsg &= If(errMsg <> "", "<br>", "")
                        errMsg &= String.Format("第{0}筆班級資料 [報名結束日期]不能晚於[結訓日期]!", (i + 1).ToString())
                    End If
                End If
                '上架日期
                If v_OnShellDate_i = "" OrElse vsOnShellDate = "" Then
                    errMsg &= If(errMsg <> "", "<br>", "")
                    errMsg &= String.Format("第{0}筆班級資料 [上架日期] 日期格式有誤!!", (i + 1).ToString()) 'Return False
                End If
                If hidSTDATE.Value = "" OrElse hidFTDATE.Value = "" Then
                    errMsg &= If(errMsg <> "", "<br>", "")
                    errMsg &= String.Format("第{0}筆班級資料 [開結訓日期] 日期格式有誤!!", (i + 1).ToString()) 'Return False
                End If
                If v_FEnterDate_i <> "" AndAlso hidSTDATE.Value <> "" AndAlso hidFTDATE.Value <> "" Then
                    If (DateDiff(DateInterval.Day, CDate(DateAdd(DateInterval.Day, 15, CDate(hidSTDATE.Value))), CDate(v_FEnterDate_i)) >= 0) Then
                        errMsg &= If(errMsg <> "", "<br>", "")
                        errMsg &= String.Format(str_errMSG_FM1, (i + 1).ToString()) 'Return False
                    Else
                        If (DateDiff(DateInterval.Day, CDate(hidSTDATE.Value), CDate(hidFTDATE.Value)) <= 14) Then
                            If (DateDiff(DateInterval.Day, CDate(hidFTDATE.Value), CDate(v_FEnterDate_i)) >= 0) Then
                                errMsg &= If(errMsg <> "", "<br>", "")
                                errMsg &= String.Format(str_errMSG_FM1, (i + 1).ToString()) 'Return False
                            End If
                        End If
                    End If
                End If

            Else
                '"尚未勾選班級!!"
                '沒勾選的資料則做新舊資料的比對，若資料已有更動時則提示告知已異動請勾選
                If oldOnShellDate <> vsOnShellDate Then
                    errMsg &= String.Concat(If(errMsg <> "", "<br>", ""), String.Format("第{0}筆班級資料 [上架日期]已異動 請勾選!", (i + 1).ToString()))
                End If
                If oldSEnterDate <> vsSEnterDate Then
                    errMsg &= String.Concat(If(errMsg <> "", "<br>", ""), String.Format("第{0}筆班級資料 [報名開始日期]已異動 請勾選!", (i + 1).ToString()))
                End If
                If oldFEnterDate <> vsFEnterDate Then
                    errMsg &= String.Concat(If(errMsg <> "", "<br>", ""), String.Format("第{0}筆班級資料 [報名結束日期]已異動 請勾選!", (i + 1).ToString()))
                End If
            End If
        Next

        If Not blFlag Then
            errMsg &= "尚未勾選班級!!"
        End If
    End Sub

    ''' <summary>'儲存</summary>
    ''' <param name="sm"></param>
    ''' <param name="MyPage"></param>
    ''' <param name="s_SAVEMSG"></param>
    ''' <param name="oDG1"></param>
    ''' <returns></returns>
    Public Shared Function UTL_SAVEDATA1(ByRef sm As SessionModel, ByRef MyPage As Page, ByRef s_SAVEMSG As String, ByRef oDG1 As DataGrid) As Boolean
        Dim rst As Boolean = False 'false:異常 /true:正常
        'oDG1:DG_ClassInfo
        '批次設定上架日
        Dim uSql As String = ""
        uSql &= " UPDATE CLASS_CLASSINFO" & vbCrLf
        uSql &= " SET ONSHELLDATE = @ONSHELLDATE,MODIFYDATE = GETDATE(),MODIFYACCT = @MODIFYACCT"
        uSql &= " WHERE OCID = @OCID"

        Dim uSql2 As String = ""
        uSql2 &= " UPDATE CLASS_CLASSINFO" & vbCrLf
        uSql2 &= " SET SENTERDATE = @SENTERDATE,MODIFYDATE = GETDATE(),MODIFYACCT = @MODIFYACCT"
        uSql2 &= " WHERE OCID = @OCID"

        Dim uSql3 As String = ""
        uSql3 &= " UPDATE CLASS_CLASSINFO" & vbCrLf
        uSql3 &= " SET FENTERDATE = @FENTERDATE,MODIFYDATE = GETDATE(),MODIFYACCT = @MODIFYACCT"
        uSql3 &= " WHERE OCID = @OCID"

        '儲存設定結果
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            For Each eItem As DataGridItem In oDG1.Items
                Dim chkItem As HtmlInputCheckBox = eItem.FindControl("chkItem")
                Dim hidOCID As HiddenField = eItem.FindControl("hidOCID")
                '批次設定上架日期
                Dim OnShellDate_i As TextBox = eItem.FindControl("OnShellDate_i")
                'Dim imgOnShellDate_i As HtmlImage = eItem.FindControl("imgOnShellDate_i")
                'imgOnShellDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & OnShellDate_i.ClientID & "','','','CY/MM/DD');")
                Dim OnShellDate_HR_i As DropDownList = eItem.FindControl("OnShellDate_HR_i")
                Dim OnShellDate_MI_i As DropDownList = eItem.FindControl("OnShellDate_MI_i")
                Dim hidOnShellDate As HiddenField = eItem.FindControl("hidOnShellDate")
                Dim hidOnShellDate_HR As HiddenField = eItem.FindControl("hidOnShellDate_HR")
                Dim hidOnShellDate_MI As HiddenField = eItem.FindControl("hidOnShellDate_MI")

                Dim SEnterDate_i As TextBox = eItem.FindControl("SEnterDate_i")
                'Dim imgSEnterDate_i As HtmlImage = eItem.FindControl("imgSEnterDate_i")
                'imgSEnterDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & SEnterDate_i.ClientID & "','','','CY/MM/DD');")
                Dim SEnterDate_HR_i As DropDownList = eItem.FindControl("SEnterDate_HR_i")
                Dim SEnterDate_MI_i As DropDownList = eItem.FindControl("SEnterDate_MI_i")
                Dim hidSEnterDate As HiddenField = eItem.FindControl("hidSEnterDate")
                Dim hidSEnterDate_HR As HiddenField = eItem.FindControl("hidSEnterDate_HR")
                Dim hidSEnterDate_MI As HiddenField = eItem.FindControl("hidSEnterDate_MI")

                Dim FEnterDate_i As TextBox = eItem.FindControl("FEnterDate_i")
                'Dim imgFEnterDate_i As HtmlImage = eItem.FindControl("imgFEnterDate_i")
                'imgFEnterDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & FEnterDate_i.ClientID & "','','','CY/MM/DD');")
                Dim FEnterDate_HR_i As DropDownList = eItem.FindControl("FEnterDate_HR_i")
                Dim FEnterDate_MI_i As DropDownList = eItem.FindControl("FEnterDate_MI_i")
                Dim hidFEnterDate As HiddenField = eItem.FindControl("hidFEnterDate")
                Dim hidFEnterDate_HR As HiddenField = eItem.FindControl("hidFEnterDate_HR")
                Dim hidFEnterDate_MI As HiddenField = eItem.FindControl("hidFEnterDate_MI")

                '上架日期 '(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
                Dim vsOnShellDate As String = TIMS.GET_YMDHM1(OnShellDate_i.Text, TIMS.GetListValue(OnShellDate_HR_i), TIMS.GetListValue(OnShellDate_MI_i))
                Dim oldOnShellDate As String = TIMS.GET_YMDHM1(hidOnShellDate.Value, hidOnShellDate_HR.Value, hidOnShellDate_MI.Value)
                Dim v_OnShellDate_i As String = If(vsOnShellDate <> "", vsOnShellDate, oldOnShellDate)

                '報名開始日期'(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
                Dim vsSEnterDate As String = TIMS.GET_YMDHM1(SEnterDate_i.Text, TIMS.GetListValue(SEnterDate_HR_i), TIMS.GetListValue(SEnterDate_MI_i))
                Dim oldSEnterDate As String = TIMS.GET_YMDHM1(hidSEnterDate.Value, hidSEnterDate_HR.Value, hidSEnterDate_MI.Value)
                Dim v_SEnterDate_i As String = If(vsSEnterDate <> "", vsSEnterDate, oldSEnterDate)

                '報名結束日期'(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
                Dim vsFEnterDate As String = TIMS.GET_YMDHM1(FEnterDate_i.Text, TIMS.GetListValue(FEnterDate_HR_i), TIMS.GetListValue(FEnterDate_MI_i))
                Dim oldFEnterDate As String = TIMS.GET_YMDHM1(hidFEnterDate.Value, hidFEnterDate_HR.Value, hidFEnterDate_MI.Value)
                Dim v_FEnterDate_i As String = If(vsFEnterDate <> "", vsFEnterDate, oldFEnterDate)

                's_TransType = TIMS.cst_TRANS_LOG_Update 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
                's_TargetTable = "CLASS_CLASSINFO"
                's_FuncPath = "/TC/01/TC_01_004"
                's_WHERE = String.Format("OCID={0}", hidOCID.Value)

                If chkItem.Checked Then hidOCID.Value = TIMS.ClearSQM(hidOCID.Value)
                '與舊值不同直接修改
                If hidOCID.Value <> "" AndAlso chkItem.Checked Then
                    If vsOnShellDate <> "" AndAlso vsOnShellDate <> oldFEnterDate Then 'CLASS_CLASSINFO
                        Dim mParms As New Hashtable
                        mParms.Add("ONSHELLDATE", If(vsOnShellDate <> "", vsOnShellDate, Convert.DBNull))
                        mParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                        mParms.Add("OCID", Val(hidOCID.Value))
                        DbAccess.ExecuteNonQuery(uSql, trans, mParms)
                    End If
                    If vsSEnterDate <> "" AndAlso vsSEnterDate <> oldSEnterDate Then
                        Dim mParms2 As New Hashtable
                        mParms2.Add("SENTERDATE", If(vsSEnterDate <> "", vsSEnterDate, Convert.DBNull))
                        mParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
                        mParms2.Add("OCID", Val(hidOCID.Value))
                        DbAccess.ExecuteNonQuery(uSql2, trans, mParms2)
                    End If
                    If vsFEnterDate <> "" AndAlso vsFEnterDate <> oldFEnterDate Then
                        Dim mParms3 As New Hashtable
                        mParms3.Add("FENTERDATE", If(vsFEnterDate <> "", vsFEnterDate, Convert.DBNull))
                        mParms3.Add("MODIFYACCT", sm.UserInfo.UserID)
                        mParms3.Add("OCID", Val(hidOCID.Value))
                        DbAccess.ExecuteNonQuery(uSql3, trans, mParms3)
                    End If
                End If
            Next
            Call DbAccess.CommitTrans(trans)
        Catch ex As Exception
            TIMS.WriteTraceLog(MyPage, ex, ex.ToString)
            Call DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(conn)
            s_SAVEMSG = "儲存失敗!!"
            Return rst
            'Common.MessageBox(Me, "儲存失敗")
            'Exit Sub
        End Try
        Call TIMS.CloseDbConn(conn)
        'bt_search_Click(sender, e)
        'Common.MessageBox(Me, "儲存成功")
        s_SAVEMSG = "儲存成功"
        rst = True
        Return rst
    End Function

    ''' <summary> 批次設定(檢核／儲存) </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        Dim errMsg As String = ""
        '批次設定-檢核
        Call CHK_SAVEDATA1(errMsg)

        If errMsg <> "" Then
            Common.MessageBox(Me, errMsg)
            Return
        End If

        Dim saveMsg As String = ""
        Dim flag_SAVE_OK As Boolean = False

        '批次設定-儲存
        flag_SAVE_OK = UTL_SAVEDATA1(sm, Me, saveMsg, DG_ClassInfo)

        If saveMsg <> "" Then Common.MessageBox(Me, saveMsg)

        If Not flag_SAVE_OK Then Return

        Call GetSearchStr()
        Call sSearch1()
    End Sub

    ''' <summary> 匯出-開班資料檔 </summary>
    Sub Utl_EXPORT1(ByRef objtable As DataTable)
        msg.Text = ""
        If objtable Is Nothing Then
            msg.Text = "查無資料!!"
            Return
        End If
        If objtable.Rows.Count = 0 Then
            msg.Text = "查無資料!!"
            Return
        End If

        'Dim objtable As DataTable = Session("TC_table")
        Dim s_fileName1 As String = HttpUtility.UrlEncode(String.Format("開班資料檔-{0}.xls", TIMS.GetDateNo2()), System.Text.Encoding.UTF8)

        '計畫年度,分署,申請階段,訓練機構,班別名稱,訓練人數,訓練時數,開訓日期,結訓日期,報名開始日期(含時間),報名結束日期(含時間),上架日期(含時間)
        Const s_title1 As String = "計畫年度,分署,申請階段,訓練機構,班別名稱,課程代碼,訓練人數,訓練時數,開訓日期,結訓日期,報名開始日期(含時間),報名結束日期(含時間),上架日期(含時間)"
        Const s_data1 As String = "PLANYEARS,DISTNAME,MYAPPSTAGE,ORGNAME,CLASSCNAME,OCID,TNUM,THOURS,STDATE,FTDATE,SENTERDATE,FENTERDATE,ONSHELLDATE"
        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")

        Dim sFileName1 As String = String.Format("開班資料檔-{0}", TIMS.GetDateNo2())

        Dim strSTYLE As String = ""
        '套CSS值
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String '建立輸出文字
        ExportStr = "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "編號") '"<td>編號</td>"
        For Each s_T1 As String In As_title1
            ExportStr &= String.Format("<td>{0}</td>", s_T1) ' "<td>" & s_T1 & "</td>"   '& vbTab
        Next
        ExportStr &= "</tr>"
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim i_num As Integer = 0
        For Each oDr1 As DataRow In objtable.Rows
            i_num += 1
            ExportStr = "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", CStr(i_num))
            For Each s_D1 As String In As_data1
                Select Case s_D1
                    Case "STDATE", "FTDATE"
                        ExportStr &= String.Format("<td>{0}</td>", TIMS.Cdate3(oDr1(s_D1))) '"<td>" & TIMS.cdate3(oDr1(s_D1)) & "</td>"
                    Case Else
                        ExportStr &= String.Format("<td>{0}</td>", TIMS.ClearSQM(oDr1(s_D1))) '"<td>" & TIMS.ClearSQM(oDr1(s_D1)) & "</td>"
                End Select
            Next
            ExportStr &= "</tr>"
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")
        objtable = Nothing

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    ''' <summary>匯出鈕</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub bt_EXPORT_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_EXPORT.Click
        Call GetSearchStr()
        Dim dt As DataTable = Get_Search1dt() '查詢 'sUtl_Search1(dt)
        Utl_EXPORT1(dt)
    End Sub

    '刪除。
    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand
        If Me.isBlack.Value = "Y" Then
            Common.MessageBox(Me, vsMsg2)
            Exit Sub
        End If
        Select Case e.CommandName
            Case "add"
                'SmartQuery匯出
                'but.Attributes("onclick") =     ReportQuery.ReportScript(Me, "MultiBlock", "TENS", "OCID=" & dr("OCID"))
                'Dim cGuid As String =   ReportQuery.GetGuid(Page)
                'Dim Url As String =   ReportQuery.GetUrl(Page)
                'Page.RegisterStartupScript("0000", "<script>window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&filename=TENS&path=TIMS&OCID=" & e.CommandArgument & "');</script>")
            Case "del"
                Call GetSearchStr()
                'Response.Redirect("TC_01_004_del.aspx?" & e.CommandArgument & "")
                Dim sUrl1 As String = String.Format("{0}{1}", cst_TIMS_DEL_ASPX1, e.CommandArgument)
                TIMS.Utl_Redirect(Me, objconn, sUrl1)
            Case "edit"
                Call GetSearchStr()
                Dim sUrl As String = ""
                '一般 TIMS 'TC_01_004_add
                sUrl = String.Format("{0}ProcessType=Update&{1}", cst_TIMS_EDITASPX1, e.CommandArgument)
                '產投、充飛 'TC_01_004_BusAdd
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then sUrl = String.Format("{0}{1}", cst_TIMS28_EDITASPX1, e.CommandArgument)
                'Response.Redirect(sUrl)
                'Dim url1 As String = "TC_01_004_del.aspx?" & e.CommandArgument & ""
                TIMS.Utl_Redirect(Me, objconn, sUrl)
        End Select
    End Sub

    'list
    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim img_sort As New UI.WebControls.Image
                Dim i_sortimg As Integer = 0
                If Me.ViewState("sort") IsNot Nothing Then
                    Select Case Me.ViewState("sort")
                        Case "OrgName", "OrgName desc"
                            i_sortimg = cst_dg_訓練機構 '2
                    End Select
                    img_sort.ImageUrl = If(Me.ViewState("sort").ToString.IndexOf("desc") = -1, "../../images/SortUp.gif", "../../images/SortDown.gif")
                    e.Item.Cells(i_sortimg).Controls.Add(img_sort)
                End If
                '2018-09-20 管控單位是職前（補助地方政府計畫用的資料）才有的資訊，產投沒有用到，所以就先拿掉不顯示
                'e.Item.Cells(cst_dg_OrgName2).Visible = False
            Case ListItemType.Item, ListItemType.AlternatingItem ',ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidOCID As HiddenField = e.Item.FindControl("hidOCID")
                Dim hidSTDATE As HiddenField = e.Item.FindControl("hidSTDATE")
                Dim hidFTDATE As HiddenField = e.Item.FindControl("hidFTDATE")
                hidOCID.Value = Convert.ToString(drv("OCID"))
                hidSTDATE.Value = Convert.ToString(drv("STDATE"))
                hidFTDATE.Value = Convert.ToString(drv("FTDATE"))

                '批次設定上架日期
                Dim OnShellDate_i As TextBox = e.Item.FindControl("OnShellDate_i")
                Dim imgOnShellDate_i As HtmlImage = e.Item.FindControl("imgOnShellDate_i")
                imgOnShellDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & OnShellDate_i.ClientID & "','','','CY/MM/DD');")
                Dim OnShellDate_HR_i As DropDownList = e.Item.FindControl("OnShellDate_HR_i")
                Dim OnShellDate_MI_i As DropDownList = e.Item.FindControl("OnShellDate_MI_i")
                Dim hidOnShellDate As HiddenField = e.Item.FindControl("hidOnShellDate")
                Dim hidOnShellDate_HR As HiddenField = e.Item.FindControl("hidOnShellDate_HR")
                Dim hidOnShellDate_MI As HiddenField = e.Item.FindControl("hidOnShellDate_MI")

                Dim SEnterDate_i As TextBox = e.Item.FindControl("SEnterDate_i")
                Dim imgSEnterDate_i As HtmlImage = e.Item.FindControl("imgSEnterDate_i")
                imgSEnterDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & SEnterDate_i.ClientID & "','','','CY/MM/DD');")
                Dim SEnterDate_HR_i As DropDownList = e.Item.FindControl("SEnterDate_HR_i")
                Dim SEnterDate_MI_i As DropDownList = e.Item.FindControl("SEnterDate_MI_i")
                Dim hidSEnterDate As HiddenField = e.Item.FindControl("hidSEnterDate")
                Dim hidSEnterDate_HR As HiddenField = e.Item.FindControl("hidSEnterDate_HR")
                Dim hidSEnterDate_MI As HiddenField = e.Item.FindControl("hidSEnterDate_MI")

                Dim FEnterDate_i As TextBox = e.Item.FindControl("FEnterDate_i")
                Dim imgFEnterDate_i As HtmlImage = e.Item.FindControl("imgFEnterDate_i")
                imgFEnterDate_i.Attributes.Add("onclick", "javascript:show_calendar('" & FEnterDate_i.ClientID & "','','','CY/MM/DD');")
                Dim FEnterDate_HR_i As DropDownList = e.Item.FindControl("FEnterDate_HR_i")
                Dim FEnterDate_MI_i As DropDownList = e.Item.FindControl("FEnterDate_MI_i")
                Dim hidFEnterDate As HiddenField = e.Item.FindControl("hidFEnterDate")
                Dim hidFEnterDate_HR As HiddenField = e.Item.FindControl("hidFEnterDate_HR")
                Dim hidFEnterDate_MI As HiddenField = e.Item.FindControl("hidFEnterDate_MI")

                e.Item.Cells(cst_dg_序號).Text = e.Item.ItemIndex + 1 + DG_ClassInfo.PageSize * DG_ClassInfo.CurrentPageIndex

                Call TIMS.SUB_SET_HR_MI(OnShellDate_HR_i, OnShellDate_MI_i)
                If Not IsDBNull(drv("OnShellDate")) Then
                    hidOnShellDate.Value = TIMS.Cdate3(drv("OnShellDate"))
                    OnShellDate_i.Text = hidOnShellDate.Value 'TIMS.cdate3(drv("OnShellDate"))
                    TIMS.SET_DateHM(CDate(drv("OnShellDate")), OnShellDate_HR_i, OnShellDate_MI_i)

                    Dim v_HR As String = CDate(drv("OnShellDate")).Hour.ToString().PadLeft(2, "0")
                    Dim v_MI As String = CDate(drv("OnShellDate")).Minute.ToString().PadLeft(2, "0")
                    hidOnShellDate_HR.Value = v_HR 'CDate(drv("OnShellDate")).Hour
                    hidOnShellDate_MI.Value = v_MI 'CDate(drv("OnShellDate")).Minute
                End If
                Call TIMS.SUB_SET_HR_MI(SEnterDate_HR_i, SEnterDate_MI_i)
                If Not IsDBNull(drv("SEnterDate")) Then
                    hidSEnterDate.Value = TIMS.Cdate3(drv("SEnterDate"))
                    SEnterDate_i.Text = hidSEnterDate.Value 'TIMS.cdate3(drv("OnShellDate"))
                    TIMS.SET_DateHM(CDate(drv("SEnterDate")), SEnterDate_HR_i, SEnterDate_MI_i)

                    Dim v_HR As String = CDate(drv("SEnterDate")).Hour.ToString().PadLeft(2, "0")
                    Dim v_MI As String = CDate(drv("SEnterDate")).Minute.ToString().PadLeft(2, "0")
                    hidSEnterDate_HR.Value = v_HR 'CDate(drv("OnShellDate")).Hour
                    hidSEnterDate_MI.Value = v_MI 'CDate(drv("OnShellDate")).Minute
                End If
                Call TIMS.SUB_SET_HR_MI(FEnterDate_HR_i, FEnterDate_MI_i)
                If Not IsDBNull(drv("FEnterDate")) Then
                    hidFEnterDate.Value = TIMS.Cdate3(drv("FEnterDate"))
                    FEnterDate_i.Text = hidFEnterDate.Value 'TIMS.cdate3(drv("OnShellDate"))
                    TIMS.SET_DateHM(CDate(drv("FEnterDate")), FEnterDate_HR_i, FEnterDate_MI_i)
                    Dim v_HR As String = CDate(drv("FEnterDate")).Hour.ToString().PadLeft(2, "0")
                    Dim v_MI As String = CDate(drv("FEnterDate")).Minute.ToString().PadLeft(2, "0")
                    hidFEnterDate_HR.Value = v_HR 'CDate(drv("OnShellDate")).Hour
                    hidFEnterDate_MI.Value = v_MI 'CDate(drv("OnShellDate")).Minute
                End If

                'TPropertyID : 0:職前/ 1:在職/ 2:接受委託
                'Dim TPropertyID_N As String = ""
                'Select Case Convert.ToString(drv("TPropertyID"))
                '    Case "0"
                '        TPropertyID_N = "職前"
                '    Case "1"
                '        TPropertyID_N = "在職"
                '    Case "2"
                '        TPropertyID_N = "接受委託"
                'End Select
                'If TPropertyID_N <> "" Then e.Item.Cells(cst_dg_TPropertyID10).Text = TPropertyID_N

                'If Convert.ToString(drv("OCID")) <> "" Then TIMS.Tooltip(e.Item.Cells(cst_dg_TPropertyID10), Convert.ToString(drv("OCID")))

                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit") '修改
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel") '刪除
                Dim lbtExport As LinkButton = e.Item.FindControl("lbtExport") '匯出

                lbtEdit.CommandArgument = "ocid=" & drv("OCID") & "&ID=" & Request("ID") & "" '修改

                lbtDel.CommandArgument = "ocid=" & drv("OCID") & "&PlanID=" & drv("PlanID") & "&ComIDNO=" & drv("ComIDNO") & "&SeqNO=" & drv("SeqNO") & "&Years=" & drv("Years") & "&ID=" & Request("ID") & "" '刪除

                lbtExport.CommandArgument = Convert.ToString(drv("OCID")) '匯出
                lbtExport.Attributes("onclick") = ReportQuery.ReportScript(Me, "MultiBlock", "TENS", "OCID=" & drv("OCID"))

                Dim parms_A As New Hashtable
                parms_A.Add("OCID", Val(drv("OCID")))

                Dim sqlstr_A As String = ""
                Dim i_NUMa As Integer = 0
                'Try 'Catch ex As Exception '有錯就算了 'End Try
                sqlstr_A = "SELECT COUNT(1) x FROM dbo.STUD_ENTERTYPE WITH(NOLOCK) WHERE ocid1 = @OCID"
                i_NUMa += Val(DbAccess.ExecuteScalar(sqlstr_A, objconn, parms_A))
                If i_NUMa = 0 Then
                    sqlstr_A = " SELECT COUNT(1) x FROM dbo.STUD_ENTERTYPE2 WITH(NOLOCK) WHERE OCID1 = @OCID"
                    i_NUMa += Val(DbAccess.ExecuteScalar(sqlstr_A, objconn, parms_A))
                End If
                Dim sqlstr_B As String = " SELECT COUNT(1) FROM dbo.CLASS_STUDENTSOFCLASS WITH(NOLOCK) WHERE OCID =@OCID"
                Dim i_NUMb As Integer = Val(DbAccess.ExecuteScalar(sqlstr_B, objconn, parms_A))
                Dim sqlstr_C As String = " SELECT COUNT(1) FROM dbo.CLASS_SCHEDULE WITH(NOLOCK) WHERE OCID =@OCID"
                Dim i_NUMc As Integer = Val(DbAccess.ExecuteScalar(sqlstr_C, objconn, parms_A))

                Dim is_parent As String = TIMS.c_false
                If (i_NUMa + i_NUMb + i_NUMc) >= 1 Then is_parent = TIMS.c_true
                If is_parent = TIMS.c_true Then
                    Dim strMsg As String = ""
                    If i_NUMa > 0 Then strMsg &= "此班級檔 尚有報名資料(Stud_EnterType:" & drv("OCID") & "),已有資料參照,不可刪除!!!\n"
                    If i_NUMb > 0 Then strMsg &= "此班級檔 尚有班級學員(Class_StudentsOfClass:" & drv("OCID") & "),已有資料參照,不可刪除!!!\n"
                    If i_NUMc > 0 Then strMsg &= "此班級檔 尚有排課(Class_Schedule:" & drv("OCID") & "),已有資料參照,不可刪除!!!\n"
                    If strMsg <> "" Then
                        lbtDel.Attributes("onclick") = "javascript:alert('" & strMsg & "');return false;"
                    Else
                        lbtDel.Attributes("onclick") = "return confirm('此動作會刪除班別資料，您確定要刪除這一筆紀錄?');"
                    End If
                Else
                    lbtDel.Attributes("onclick") = "return confirm('此動作會刪除班別資料，您確定要刪除這一筆紀錄?');"
                    'lbtDel.Attributes("onclick") = "javascript:return confirm('此動作會刪除班別資料，是否確定刪除?');"
                End If
                'but_del.Attributes.Add("onclick", "but_del(" & dr("OCID") & "," & dr("PlanID") & "," & dr("ComIDNO") & "," & dr("SeqNO") & ",'" & dr("Years") & "'," & is_parent & "," & Request("ID") & ");return false;")
                'lbtEdit.Enabled = False
                'If check_mod.Value = "1" Then lbtEdit.Enabled = True

                '2005/5/30系統管理者,才可以刪除班級-Melody
                Dim flagAdmin As Boolean = False
                Select Case sm.UserInfo.RoleID
                    Case "1", "0"
                        flagAdmin = True
                End Select
                lbtDel.Visible = If(flagAdmin, True, False)
                lbtDel.Enabled = If(flagAdmin, True, False)

                '班別名稱
                TIMS.Tooltip(e.Item.Cells(cst_dg_班別名稱), Convert.ToString(drv("OCID")))

                Dim v_ClassID As String = Convert.ToString(drv("ClassID"))
                Dim s_span1 As String = String.Concat(drv("Years"), "0")
                Dim s_span2 As String = TIMS.GET_SPAN_RED1(v_ClassID)
                Dim s_span3 As String = If(Convert.ToString(drv("CYCLTYPE")) <> "", Convert.ToString(drv("CYCLTYPE")), "01")
                e.Item.Cells(cst_dg_班別代碼).Text = If(v_ClassID.Length > 4, s_span2, String.Concat(s_span1, s_span2, s_span3))
                'Result = courName 'myTableCell.Text = courName 'Result
                TIMS.Tooltip(e.Item.Cells(cst_dg_班別代碼), "前2碼為年度，中間4碼為班別代碼，後2碼為期別")

                Dim myTableCell1 As TableCell = e.Item.Cells(cst_dg_開結訓日)
                myTableCell1.Text = String.Format("{0}<br>|<br>{1}", TIMS.Cdate3(drv("STDate")), TIMS.Cdate3(drv("FTDate"))) 'Result1

        End Select
    End Sub

    Private Sub DG_ClassInfo_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DG_ClassInfo.SortCommand
        Me.ViewState("sort") = If(e.SortExpression = Me.ViewState("sort"), String.Concat(e.SortExpression, " desc"), e.SortExpression)

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

End Class