Partial Class TC_01_004_BusAdd
    Inherits AuthBasePage

    '(產投)專用。
    'CLASS_CLASSINFO

    Sub sUtl_PageInit1()
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        'InitializeComponent()
        Dim strTables As String = "'ORG_ORGINFO','ID_CLASS','CLASS_CLASSINFO','KEY_TRAINTYPE','PLAN_PLANINFO','PLAN_ONCLASS'"
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS(strTables, objconn)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "ORGNAME", OrgName) '訓練機構
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSID", TBclass_id) '班別代碼
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSCNAME", ClassCName) '班級中文名稱
        Call TIMS.sUtl_SetMaxLen(dt, "CYCLTYPE", CyclType) '期別
        Call TIMS.sUtl_SetMaxLen(dt, "CLASSENGNAME", ClassEngName) '班級英文名稱
        'Call TIMS.sUtl_SetMaxLen(dt, "TRAINID", OrgName)
        Call TIMS.sUtl_SetMaxLen(dt, "TRAINNAME", TB_career_id) '訓練職類
        Call TIMS.sUtl_SetMaxLen(dt, "CTNAME", TechName) '師資
        Call TIMS.sUtl_SetMaxLen(dt, "OTHERREASON", OtherReason) '不開班原因其他原因說明
        Call TIMS.sUtl_SetMaxLen(dt, "ROOMNAME", RoomName) '上課教室名稱
        Call TIMS.sUtl_SetMaxLen(dt, "FACTMODEOTHER", FactModeOther) '場地類型其他說明
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTNAME", ContactName) '聯絡人
        'Call TIMS.sUtl_SetMaxLen(dt, "CONTACTPHONE", ContactPhone) '電話
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTEMAIL", ContactEmail) '電子郵件
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTFAX", ContactFax) '傳真
        Call TIMS.sUtl_SetMaxLen(dt, "TIMES", Times) '時間
    End Sub

    ''sm.UserInfo.TPlanID = "28" 產業人才投資計劃用
    'Dim connStr As String = System.Configuration.ConfigurationSettings.AppSettings("ConnectionString")
    ''Dim ProcessType As String = ""
    Const cst_errmsg1 As String = "班級基本資料有誤，請確認資料正確性!"
    Const cst_errmsg2 As String = "班級基本資料有誤，請重新確認資料正確性!"
    'Const cst_errmsg3 As String = "班級已有學員報名資料，不可再修改學員報名時間!"
    Const cst_errmsg4 As String = "使用者登入計畫有誤，不提供儲存!!"

    Const cst_temp_classinfo As String = "temp_classinfo" 'Session(cst_temp_classinfo)

    Dim vsSEnterDate As String = "" '報名開始日期
    Dim vsFEnterDate As String = "" '報名結束日期
    Dim vsOnShellDate As String = "" '上架日期(產投) CLASS_CLASSINFO

    Dim rq_OCID As String = "" ' TIMS.ClearSQM(Request("OCID"))
    Dim dtTeacherInfo As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call sUtl_PageInit1()
        '檢查Session是否存在 End

        rq_OCID = TIMS.ClearSQM(Request("OCID"))
        If Convert.ToString(Request("STDate")) <> "" Then STDate.Text = TIMS.ClearSQM(Request("STDate")) '若有輸入值，以輸入值為準

        '各物件 鎖定
        Call Create00()

        If Not IsPostBack Then
            Call CreateItem()
            ProecessType.Text = If(rq_OCID = "", "新增", "修改")

            If Session("ClassSearchStr") IsNot Nothing Then
                ViewState("ClassSearchStr") = Session("ClassSearchStr")
                Session("ClassSearchStr") = Nothing
            End If

            Call SHOW_CLASSINFO() '帶出相關資料
        End If

        IsBusiness.Enabled = False
        EnterpriseName.Enabled = False
        IsBusiness.ToolTip = "本年度暫不開放此功能"
        EnterpriseName.ToolTip = "本年度暫不開放此功能"

        Button1.Attributes("onclick") = "return CheckData();"
        Button4.Attributes("onclick") = "wopen('../../Common/TechID.aspx?RID=" & RIDValue.Value & "&TextField=TechName&ValueField=CTName&CTName='+document.getElementById('CTName').value,'Tech',350,400,1);"
        Button5.Attributes("onclick") = "return CheckAddClass();"

        'trainValue.Value '取後 jobValue 的 TMID 97年產業人才投資方案
        Choice_Button2.Attributes("onclick") = "wopen('TC_01_004_Class.aspx?TMID='+'" & jobValue.Value & "' ,'班別代碼',300,300,1);"
        Img4.Style("display") = "none"
        Img5.Style("display") = "none"
    End Sub

    '檢核登入者的計畫 異常為False
    Function ChkTPlanID(ByRef MyPage As Page) As Boolean
        Dim rst As Boolean = False
        If Convert.ToString(sm.UserInfo.TPlanID) = "" Then Return rst
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then rst = True '可使用的計畫 (true:正常)
        Return rst
    End Function

    '某些日期的檢核條件。
    Function ChkSTDate(ByRef ErrMsg As String) As Boolean
        Dim blnRst As Boolean = True
        ErrMsg = ""
        'tmpSTDate15 = DateAdd(DateInterval.Day, 15, CDate(STDate.Text))
        STDate.Text = TIMS.Cdate3(STDate.Text)
        FTDate.Text = TIMS.Cdate3(FTDate.Text)
        SEnterDate.Text = TIMS.Cdate3(SEnterDate.Text)
        FEnterDate.Text = TIMS.Cdate3(FEnterDate.Text)
        'ExamDate.Text = TIMS.cdate3(ExamDate.Text)
        OnShellDate.Text = TIMS.Cdate3(OnShellDate.Text) '上架日期

        'vsSEnterDate = ""
        'vsFEnterDate = ""
        'vsSEnterDate = SEnterDate.Text & " " & Format(CInt(HR1.SelectedValue), "00") & ":" & Format(CInt(MM1.SelectedValue), "00")
        vsSEnterDate = TIMS.GET_YMDHM1(SEnterDate.Text, TIMS.GetListValue(HR1), TIMS.GetListValue(MM1))
        If Not TIMS.IsDate1(vsSEnterDate) Then vsSEnterDate = ""
        If vsSEnterDate = "" Then
            'Common.MessageBox(Me, "報名開始日期 不可為空白!!") 'Exit Function
            ErrMsg = "報名開始日期 不可為空白!!"
            Return False
        End If
        'vsFEnterDate = FEnterDate.Text & " " & Format(CInt(HR2.SelectedValue), "00") & ":" & Format(CInt(MM2.SelectedValue), "00")
        vsFEnterDate = TIMS.GET_YMDHM1(FEnterDate.Text, TIMS.GetListValue(HR2), TIMS.GetListValue(MM2))
        If Not TIMS.IsDate1(vsFEnterDate) Then vsFEnterDate = ""
        If vsFEnterDate = "" Then
            'Common.MessageBox(Me, "報名結束日期 不可為空白!!") 'Exit Function
            ErrMsg = "報名結束日期 不可為空白!!"
            Return False
        End If
        If Not IsDate(vsSEnterDate) Then
            'Common.MessageBox(Me, "報名開始日期有誤!!") 'Exit Function
            ErrMsg = "報名開始日期有誤!!"
            Return False
        End If
        If Not IsDate(vsFEnterDate) Then
            'Common.MessageBox(Me, "報名結束日期有誤!!") 'Exit Function
            ErrMsg = "報名結束日期有誤!!"
            Return False
        End If
        If DateDiff(DateInterval.Day, CDate(vsSEnterDate), CDate(vsFEnterDate)) < 0 Then
            'Common.MessageBox(Me, "報名開始~結束日期 起迄有誤請確認!!") 'Exit Function
            ErrMsg = "報名開始~結束日期 起迄有誤請確認!!"
            Return False
        End If
        If DateDiff(DateInterval.Day, CDate(vsFEnterDate), CDate(FTDate.Text)) < 0 Then
            'Common.MessageBox(Me, "報名結束日期 不可大於 結訓日期 請確認!!") 'Exit Function
            ErrMsg = "報名結束日期 不可大於 結訓日期 請確認!!"
            Return False
        End If

        'Dim bflagTrans1 As Boolean = False '無異動報名日期
        'If Not bflagTrans1 AndAlso hid_SEnterDate_old.Value <> "" Then
        '    If DateDiff(DateInterval.Day, CDate(vsSEnterDate), CDate(hid_SEnterDate_old.Value)) <> 0 Then bflagTrans1 = True '有異動
        'End If
        'If Not bflagTrans1 AndAlso hid_FEnterDate_old.Value <> "" Then
        '    If DateDiff(DateInterval.Day, CDate(vsFEnterDate), CDate(hid_FEnterDate_old.Value)) <> 0 Then bflagTrans1 = True '有異動
        'End If
        'If rq_OCID <> "" AndAlso bflagTrans1 Then
        '    Dim i2 As Integer = TIMS.Get_EnterCount(rq_OCID, objconn)
        '    If i2 > 0 Then
        '        Common.MessageBox(Me, cst_errmsg3)
        '        Exit Sub
        '    End If
        'End If

        'https://jira.turbotech.com.tw/browse/TIMSC-43
        '2:.檢核():上架日期不能晚於報名起日, 預設日期值為當日
        '上架日期
        vsOnShellDate = ""
        If OnShellDate.Text <> "" Then
            vsOnShellDate = TIMS.GET_YMDHM1(OnShellDate.Text, TIMS.GetListValue(OnShellDate_HR), TIMS.GetListValue(OnShellDate_MI))
            If Not TIMS.IsDate1(vsOnShellDate) OrElse vsOnShellDate = "" Then vsOnShellDate = ""
            If vsOnShellDate = "" Then
                ErrMsg = "上架日期日期有誤!"
                Return False
            End If
        End If

        If (DateDiff(DateInterval.Day, CDate(DateAdd(DateInterval.Day, 15, CDate(STDate.Text))), CDate(FEnterDate.Text)) >= 0) Then
            ErrMsg = ""
            ErrMsg += "報名結束日期 最晚可為開訓日後第14天，" & vbCrLf
            ErrMsg += "若為短期班，開訓後，14天內就結訓的班級，" & vbCrLf
            ErrMsg += "報名結束日期最晚為結訓日期前一天。" & vbCrLf
            'Common.MessageBox(Me.Page, msg) 'blnRst = False
        Else
            If (DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FTDate.Text)) <= 14) Then
                If (DateDiff(DateInterval.Day, CDate(FTDate.Text), CDate(FEnterDate.Text)) >= 0) Then
                    ErrMsg = ""
                    ErrMsg += "報名結束日期 最晚可為開訓日後第14天，" & vbCrLf
                    ErrMsg += "若為短期班，開訓後，14天內就結訓的班級，" & vbCrLf
                    ErrMsg += "報名結束日期 最晚為結訓日期前一天。" & vbCrLf
                    'Common.MessageBox(Me.Page, msg) 'blnRst = False
                End If
            End If
        End If
        If ErrMsg <> "" Then Return False

        '090608 andy edit
        'If ExamDate.Text <> "" Then
        '    If ExamPeriod.SelectedIndex = 0 Then '20100329 add 甄試時段
        '        ErrMsg = "「甄試日期」全天、上午、下午 時段請擇一選擇!"
        '        Return False
        '        'ErrMsg = "<script language=""javascript"">" + vbCrLf
        '        'ErrMsg += "alert('「甄試日期」全天、上午、下午 時段請擇一選擇!\n');" + vbCrLf
        '        'ErrMsg += "</script>"
        '        'Page.RegisterStartupScript("", ErrMsg)
        '        'Exit Function
        '    End If
        '    If (CDate(ExamDate.Text) <= CDate(FEnterDate.Text)) Then
        '        ErrMsg = "「甄試日期」必須大於「報名結束日期」!"
        '        Return False
        '        'ErrMsg = "<script language=""javascript"">" + vbCrLf
        '        'ErrMsg += "alert('「甄試日期」必須大於「報名結束日期」\n');" + vbCrLf
        '        'ErrMsg += "</script>"
        '        'Page.RegisterStartupScript("errMsg", ErrMsg)
        '        'Exit Function
        '    End If
        '    If (CDate(ExamDate.Text) > CDate(STDate.Text)) Then
        '        ErrMsg = "[甄試日期]必須小於或等於[開訓日期]!"
        '        Return False
        '        'ErrMsg = "<script language=""javascript"">" + vbCrLf
        '        'ErrMsg += "alert('[甄試日期]必須小於或等於[開訓日期]\n');" + vbCrLf
        '        'ErrMsg += "</script>"
        '        'Page.RegisterStartupScript("errMsg", ErrMsg)
        '        'Exit Function
        '    End If
        'End If

        'https://jira.turbotech.com.tw/browse/TIMSC-43
        '2:.檢核():上架日期不能晚於報名起日, 預設日期值為當日
        'If OnShellDate.Text <> "" Then '上架日期
        '    If DateDiff(DateInterval.Day, CDate(SEnterDate.Text), CDate(OnShellDate.Text)) > 0 Then
        '        ErrMsg = "[上架日期]不能晚於[報名起日]!"
        '        Return False
        '    End If
        'End If
        If vsOnShellDate <> "" Then '上架日期
            If DateDiff(DateInterval.Minute, CDate(vsSEnterDate), CDate(vsOnShellDate)) > 0 Then
                ErrMsg = "[上架日期]不能晚於[報名起日]!"
                Return False
            End If
        End If

        If ErrMsg <> "" Then blnRst = False
        Return blnRst
    End Function

    ''' <summary>異動報名時間。(計算)(預設值) ChangeSEnterDate</summary>
    ''' <param name="rqSDate"></param>
    ''' <returns></returns>
    Function ChgSEnterDate(ByVal rqSDate As String) As Boolean
        Dim blnRst As Boolean = True    '可報名(回傳值)
        'Dim Period As Integer = 0       '可報名期間 (依天)
        'Dim Period2 As Integer = 0      '可報名期間 (依月)
        If TIMS.ClearSQM(rqSDate) = "" Then Return blnRst

        '檢核報名日期 (若OK 轉出OUT SEnterDate/FEnterDate)
        rqSDate = TIMS.Cdate3(rqSDate)
        Dim vSEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "SEnterDate")
        Dim vFEnterDate As String = "" 'TIMS.GetMyValue2(htCC, "FEnterDate")
        '異動報名時間。(計算)(預設值)-班級轉入 'Dim flag_chkSEnDate As Boolean = False 'false:異常 
        Call TIMS.ChangeSEnterDate(rqSDate, vSEnterDate, vFEnterDate)
        SEnterDate.Text = TIMS.Cdate3(vSEnterDate)
        FEnterDate.Text = TIMS.Cdate3(vFEnterDate)
        TIMS.SET_DateHM(HR1, MM1, "1200") 'Common.SetListItem(HR1, 12) 'Common.SetListItem(MM1, 0)
        TIMS.SET_DateHM(HR2, MM2, "1800") 'Common.SetListItem(HR2, 18) Common.SetListItem(MM2, 0)

        '上架日期(預設值)
        'https://jira.turbotech.com.tw/browse/TIMSC-43
        '2:.檢核():上架日期不能晚於報名起日, 預設日期值為當日
        If OnShellDate.Text = "" Then '上架日期
            OnShellDate.Text = Common.FormatDate(DateAdd(DateInterval.Day, 0, CDate(Now)))
            TIMS.SET_DateHM(CDate(OnShellDate.Text), OnShellDate_HR, OnShellDate_MI)
        End If

        '報名日期順序有誤
        'If DateDiff(DateInterval.Day, CDate(FEnterDate.Text), CDate(SEnterDate.Text)) > 0 Then
        '    Dim tmpSEnterDate As String = SEnterDate.Text
        '    SEnterDate.Text = FEnterDate.Text
        '    FEnterDate.Text = tmpSEnterDate
        'End If

        '可報名期間
        'Dim Period As Integer = 0 '可報名期間 (依天)
        'Dim flag_chkSEnDate3 As Boolean = TIMS.ChkEnterDayS3(rqSDate)
        Dim Period As Integer = DateDiff(DateInterval.Day, CDate(Now), CDate(rqSDate))
        blnRst = If(Period <= 3, False, True) '小於、等於 開訓前三天 '不可報名／可報名
        Return blnRst
    End Function

    ''' <summary>
    ''' 物件 鎖定 (沒有開放修改儲存)
    ''' </summary>
    Sub Create00()
        'ProcessType = Request("ProcessType")  '20100110 andy 
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        'End If

        Me.LabTMID.Text = "訓練業別"
        'NotOpen_TD.InnerHtml = ""
        'NotOpen.Enabled = True
        IsFullDate_TR.Style("display") = "none" ''2
        IsFullDate.Visible = False
        '申請階段
        AppStage.Enabled = False
        'TPropertyID_TD.InnerHtml = "" ''3
        'TPropertyID.Visible = False
        IsApplic_TD.InnerHtml = "" ''4
        IsApplic.Visible = False

        '(選擇)班別代碼'5 本來就有了
        Choice_Button2.Visible = False
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "1"
                Choice_Button2.Visible = True
        End Select

        OrgName.Enabled = False
        TBclass_id.Enabled = False
        ClassCName.Enabled = False
        CyclType.Enabled = False
        ClassEngName.Enabled = False
        career.Visible = False
        TB_career_id.Enabled = False
        TNum.Enabled = False
        'Content.Enabled = False
        Purpose.Enabled = False
        THours.Enabled = False
        'TPeriod.Enabled = False
        STDate.Enabled = False
        FTDate.Enabled = False
        ConNum.Enabled = False

        RoomNameTD.InnerHtml = ""
        RoomName.Visible = False
        oldPlace.Style("display") = "none"

        'ExamDate.ReadOnly = True
        'CheckInDate.ReadOnly = True
        NotOpenTR.Visible = False
        NORIDTR.Visible = False
        NotOpenTR.Style("display") = "none"
        NORIDTR.Style("display") = "none"

        'TechName.ReadOnly = True
        'Button4.Visible = False -- 此狀態沒有作用
        Button4.Disabled = True
        Weeks.Enabled = False
        'Times.ReadOnly = True
        Button5.Enabled = False
        DataGrid1.Enabled = False
        'CredPoint.ReadOnly = True

        'txtCJOB_NAME.ReadOnly = True  '通俗職類 'onfocus="this.blur()"
        btu_sel2.Disabled = True

        'SEnterDate.ReadOnly = True 'onfocus="this.blur()"
        'FEnterDate.ReadOnly = True 'onfocus="this.blur()"
        Img1.Disabled = False '顯示
        Img2.Disabled = False '顯示
        SEnterDate.Enabled = True '啟用
        FEnterDate.Enabled = True '啟用
        HR1.Enabled = True
        MM1.Enabled = True
        HR2.Enabled = True
        MM2.Enabled = True

        '上架日期
        Img_OnShellDate.Disabled = False '顯示
        OnShellDate.Enabled = True '啟用
        OnShellDate_HR.Enabled = True '啟用
        OnShellDate_MI.Enabled = True '啟用

        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0", "1"
            Case Else
                '委訓單位登入時
                Img1.Disabled = True
                Img2.Disabled = True
                SEnterDate.Enabled = False
                FEnterDate.Enabled = False
                HR1.Enabled = False
                MM1.Enabled = False
                HR2.Enabled = False
                MM2.Enabled = False

                '上架日期
                Img_OnShellDate.Disabled = True '(Disabled)顯示
                OnShellDate.Enabled = False '(False)啟用
                OnShellDate_HR.Enabled = False '(False)啟用
                OnShellDate_MI.Enabled = False '(False)啟用

                Dim vMsg As String = "委訓單位登入"
                TIMS.Tooltip(SEnterDate, vMsg)
                TIMS.Tooltip(HR1, vMsg)
                TIMS.Tooltip(MM1, vMsg)
                TIMS.Tooltip(FEnterDate, vMsg)
                TIMS.Tooltip(HR2, vMsg)
                TIMS.Tooltip(MM2, vMsg)

                '上架日期
                TIMS.Tooltip(OnShellDate, vMsg)
                TIMS.Tooltip(OnShellDate_HR, vMsg)
                TIMS.Tooltip(OnShellDate_MI, vMsg)
        End Select

    End Sub

    ''' <summary>
    ''' 1.下拉物件設定/2.預設參數判斷
    ''' </summary>
    Sub CreateItem()

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))

        HidSYSDATE.Value = TIMS.GetSysDate(objconn)
        'With TPropertyID
        '    .Items.Clear()
        '    .Items.Add(New ListItem("職前", "0"))
        '    .Items.Add(New ListItem("在職", "1"))
        'End With
        '/職前(0)/在職(1)/
        Call TIMS.GET_TRAINEXP(TDeadline, objconn, sm)
        'TPeriod = TIMS.Get_HourRan(TPeriod)
        Call TIMS.Get_ClassCatelog(ClassCate, objconn)

        With Weeks
            .Items.Clear()
            .Items.Add(New ListItem("==請選擇==", ""))
            .Items.Add(New ListItem("星期一", "星期一"))
            .Items.Add(New ListItem("星期二", "星期二"))
            .Items.Add(New ListItem("星期三", "星期三"))
            .Items.Add(New ListItem("星期四", "星期四"))
            .Items.Add(New ListItem("星期五", "星期五"))
            .Items.Add(New ListItem("星期六", "星期六"))
            .Items.Add(New ListItem("星期日", "星期日"))
        End With
        Call TIMS.Get_NotOpenReason(NORID, objconn)

        '20100329 add 甄試時段
        'ExamPeriod = TIMS.GET_ExamPeriod(ExamPeriod, objconn)

        'HR1.Items.Clear()
        'HR2.Items.Clear()
        'OnShellDate_HR.Items.Clear()
        'MM1.Items.Clear()
        'MM2.Items.Clear()
        'OnShellDate_MI.Items.Clear()
        '初始化物件
        TIMS.SUB_SET_HR_MI(HR1, MM1)
        TIMS.SUB_SET_HR_MI(HR2, MM2)
        TIMS.SUB_SET_HR_MI(OnShellDate_HR, OnShellDate_MI)

        '設定初始值
        TIMS.SET_DateHM(HR1, MM1, "1200")
        TIMS.SET_DateHM(HR2, MM2, "1800")
        '上架日期
        TIMS.SET_DateHMC00(OnShellDate_HR, OnShellDate_MI)

        'TIMS.ClearSQM(Request("STDate"))
        If rq_OCID = "" AndAlso STDate.Text = "" Then
            Button1.Enabled = False '儲存鈕鎖定
            TIMS.Tooltip(Button1, "開訓日期為空 系統異常 停止轉班(開班)功能")
            Exit Sub
        End If

        If STDate.Text <> "" Then
            Select Case CInt(sm.UserInfo.LID)
                Case "0", "1"
                    '署(局)、分署(中心)帳號登入
                    'If ProcessType = "PlanUpdate" Then  '開班轉入時
                    'End If
                    If SEnterDate.Text = "" OrElse FEnterDate.Text = "" Then Call ChgSEnterDate(STDate.Text) '改變報名日期區間
                Case Else
                    '委訓單位登入
                    'If ProcessType = "PlanUpdate" Then  '開班轉入時 'End If
                    '改變報名日期區間
                    If Not ChgSEnterDate(STDate.Text) Then
                        Button1.Enabled = False '儲存鈕鎖定
                        TIMS.Tooltip(Button1, "小於、等於 開訓前三天停止轉班(開班)功能")
                    End If
            End Select
        End If
    End Sub

    ''' <summary>
    ''' 帶出相關資料-課程-CLASS_CLASSINFO
    ''' </summary>
    Sub SHOW_CLASSINFO()
        '檢核登入者的計畫 異常為False
        If Not ChkTPlanID(Me) Then
            Common.MessageBox(Me, cst_errmsg4)
            Exit Sub
        End If

        '產投專用。
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        If rq_OCID <> "" Then
            Dim pms_cc As New Hashtable From {{"OCID", rq_OCID}}
            Dim sql_cc As String = "SELECT * FROM dbo.CLASS_CLASSINFO WHERE OCID=@OCID "
            dt = DbAccess.GetDataTable(sql_cc, objconn, pms_cc)
            If dt.Rows.Count > 0 Then dr = dt.Rows(0) 'DbAccess.GetOneRow(sql)
        Else
            If Not Session(cst_temp_classinfo) Is Nothing Then
                dt = Session(cst_temp_classinfo)
                If dt.Rows.Count > 0 Then dr = dt.Rows(0)
                Session(cst_temp_classinfo) = Nothing
            End If
        End If

        Button1.Visible = True '儲存鈕
        If dt Is Nothing Then
            'Dim errMsg As String = ""
            'errMsg = "<script language=""javascript"">" + vbCrLf
            'errMsg += "alert('班級基本資料有誤，請確認資料正確性!\n');" + vbCrLf
            'errMsg += "</script>"
            'Page.RegisterStartupScript("", errMsg)
            Button1.Visible = False '停用儲存鈕
            TIMS.Tooltip(Button1, cst_errmsg1)
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If
        If dr Is Nothing Then
            Button1.Visible = False '停用儲存鈕
            TIMS.Tooltip(Button1, cst_errmsg1)
            Common.MessageBox(Me, cst_errmsg1)
            Exit Sub
        End If

        If dt.Rows.Count > 0 Then
            'If IsDBNull(dr("TPeriod")) Then TPeriod.Enabled = True
            '學科場地1
            SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, dr("ComIDNO").ToString, 1, "", objconn)
            '術科場地1
            TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, dr("ComIDNO").ToString, 1, "", objconn)
            '學科場地2
            SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, dr("ComIDNO").ToString, 1, "", objconn)
            '術科場地2
            TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, dr("ComIDNO").ToString, 1, "", objconn)

            AddressSciPTID = TIMS.Get_SciPTID(AddressSciPTID, dr("ComIDNO").ToString, 3, objconn)
            AddressTechPTID = TIMS.Get_TechPTID(AddressTechPTID, dr("ComIDNO").ToString, 3, objconn)

            Me.PlanID.Value = dr("PlanID").ToString
            Me.ComIDNO.Value = dr("ComIDNO").ToString
            Me.SeqNO.Value = dr("SeqNO").ToString

            Me.Years.Value = dr("Years").ToString
            RIDValue.Value = dr("RID").ToString
            Me.clsid.Value = dr("CLSID").ToString
            ClassCName.Text = dr("ClassCName").ToString
            CyclType.Text = TIMS.FmtCyclType(dr("CyclType"))
            ClassEngName.Text = dr("ClassEngName").ToString
            trainValue.Value = dr("TMID").ToString
            jobValue.Value = dr("TMID").ToString
            '/職前(0)/在職(1)/
            'Common.SetListItem(TPropertyID, dr("TPropertyID"))
            TNum.Text = dr("TNum").ToString

            '上架日期
            If Not IsDBNull(dr("OnShellDate")) Then
                OnShellDate.Text = TIMS.Cdate3(dr("OnShellDate"))
                TIMS.SET_DateHM(CDate(dr("OnShellDate")), OnShellDate_HR, OnShellDate_MI)
            End If

            '起始日
            If Not IsDBNull(dr("SEnterDate")) Then
                SEnterDate.Text = FormatDateTime(dr("SEnterDate"), 2)
                SEnterDate.Text = TIMS.Cdate3(SEnterDate.Text)
                TIMS.SET_DateHM(CDate(dr("SEnterDate")), HR1, MM1)
            End If

            hid_SEnterDate_old.Value = ""
            If Convert.ToString(dr("SEnterDate")) <> "" Then
                '針對已開放報名的班級，報名開始日期欄無法更改。
                SEnterDate.Enabled = True
                HR1.Enabled = True
                MM1.Enabled = True
                'Dim dSEtrDate As String = TIMS.cdate3(dr("SEnterDate"))
                If HidSYSDATE.Value = "" Then HidSYSDATE.Value = TIMS.GetSysDate(objconn)
                hid_SEnterDate_old.Value = TIMS.Cdate3(dr("SEnterDate"), "yyyy/MM/dd HH:mm")

                'hid_SEnterDate_old.Value = CDate(dr("SEnterDate")).ToString("yyyy/MM/dd HH:mm")
                'If DateDiff(DateInterval.Minute, CDate(dSEtrDate), CDate(HidSYSDATE.Value)) >= 0 Then
                '    SEnterDate.Enabled = False
                '    HR1.Enabled = False
                '    MM1.Enabled = False
                '    Dim vMsg As String = "報名時間已開始無法變更。"
                '    TIMS.Tooltip(SEnterDate, vMsg)
                '    TIMS.Tooltip(HR1, vMsg)
                '    TIMS.Tooltip(MM1, vMsg)
                'End If
            End If

            '結束日
            hid_FEnterDate_old.Value = ""
            If Not IsDBNull(dr("FEnterDate")) Then
                FEnterDate.Text = FormatDateTime(dr("FEnterDate"), 2)
                FEnterDate.Text = TIMS.Cdate3(FEnterDate.Text)
                TIMS.SET_DateHM(CDate(dr("FEnterDate")), HR2, MM2)
                hid_FEnterDate_old.Value = TIMS.Cdate3(dr("FEnterDate"), "yyyy/MM/dd HH:mm")
            End If

            'Content.Text = dr("Content").ToString
            Purpose.Text = dr("Purpose").ToString
            'If Not IsDBNull(dr("ExamDate")) Then
            '    ExamDate.Text = FormatDateTime(dr("ExamDate"), 2)
            '    ExamDate.Text = TIMS.cdate3(ExamDate.Text)
            'End If
            '20100329 andy add  甄試時段
            'If Not IsDBNull(dr("ExamPeriod")) Then
            '    Common.SetListItem(ExamPeriod, dr("ExamPeriod"))
            '    'ExamPeriod.SelectedValue = Convert.ToString(dr("ExamPeriod"))
            'End If
            Common.SetListItem(TDeadline, dr("TDeadline"))
            'If Not IsDBNull(dr("TaddressZip")) Then
            '    TBCity.Text = "(" & dr("TaddressZip") & ")" & TIMS.Get_ZipName(dr("TaddressZip"))
            '    city_code.Value = dr("TaddressZip").ToString
            'End If
            'TAddress.Text = dr("TAddress").ToString
            THours.Text = dr("THours").ToString
            'If dr("TPeriod").ToString <> "" Then
            '    Common.SetListItem(TPeriod, dr("TPeriod"))
            'End If
            '
            Common.SetListItem(IsFullDate, dr("IsFullDate"))
            If Not IsDBNull(dr("STDate")) Then STDate.Text = FormatDateTime(dr("STDate"), 2)
            If Not IsDBNull(dr("FTDate")) Then FTDate.Text = FormatDateTime(dr("FTDate"), 2)
            If Not IsDBNull(dr("CheckInDate")) Then CheckInDate.Text = FormatDateTime(dr("CheckInDate"), 2)
            STDate.Text = TIMS.Cdate3(STDate.Text)
            FTDate.Text = TIMS.Cdate3(FTDate.Text)
            CheckInDate.Text = TIMS.Cdate3(CheckInDate.Text)

            IsApplic.Checked = If(dr("IsApplic").ToString = "Y", True, False)
            'add by nick
            NotOpen.Checked = If(dr("NotOpen").ToString = "Y", True, False)
            'end mark

            '不開班原因
            TIMS.SetCblValue(NORID, Convert.ToString(dr("NORID")))

            OtherReason.Text = dr("OtherReason").ToString
            Relship.Value = dr("Relship").ToString
        End If

        Dim PlanID As Integer = dr("PlanID") 'Integer
        Dim ComIDNO As String = dr("ComIDNO")
        Dim SeqNo As Integer = dr("SeqNo") 'Integer
        Dim TMID As Integer = dr("TMID") 'Integer
        Dim CLSID As Integer = dr("CLSID") 'Integer
        Dim RID As String = dr("RID")
        'Dim TechID As String = dr("CTName").ToString'''''已經停用了

        Dim pms_o As New Hashtable From {{"RID", RID}}
        Dim sql_o As String = " SELECT b.OrgName FROM (SELECT * FROM Auth_Relship WHERE RID =@RID ) a JOIN Org_OrgInfo b ON a.OrgID = b.OrgID "
        OrgName.Text = DbAccess.ExecuteScalar(sql_o, objconn, pms_o)

        Dim pms_t As New Hashtable From {{"TMID", TMID}}
        Dim sql_t As String = " SELECT '[' + TrainID + ']' + TrainName FROM Key_TrainType WHERE TMID =@TMID"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql_t = " SELECT CASE WHEN JobID IS NULL THEN '[' + trainID + ']' + trainName ELSE '[' + JobID + ']' + JobName END FROM Key_TrainType WHERE TMID =@TMID"
        End If
        TB_career_id.Text = DbAccess.ExecuteScalar(sql_t, objconn, pms_t)

        'sql = "SELECT TeachCName FROM Teach_TeacherInfo WHERE TechID='" & TechID & "'"
        'TechName.Text = DbAccess.ExecuteScalar(sql)

        Dim pms_ic As New Hashtable From {{"CLSID", CLSID}}
        Dim sql_ic As String = " SELECT ClassID FROM dbo.ID_Class WHERE CLSID =@CLSID"
        TBclass_id.Text = DbAccess.ExecuteScalar(sql_ic, objconn, pms_ic)
        OldClassID.Value = TBclass_id.Text

        Dim drPP As DataRow = Nothing
        Call SHOW_PLANINFO(drPP, PlanID, ComIDNO, SeqNo)
        If drPP Is Nothing Then Return


        Dim dtTecher As DataTable = Nothing
        If rq_OCID <> "" Then
            '轉班後資料
            Dim pms_ct As New Hashtable From {{"OCID", rq_OCID}}
            Dim sql_ct As String = ""
            sql_ct += " SELECT b.TechID,b.TeachCName "
            sql_ct += " FROM CLASS_TEACHER a "
            sql_ct += " JOIN Teach_TeacherInfo b ON a.TechID = b.TechID "
            sql_ct += " WHERE a.OCID = @OCID "
            dtTecher = DbAccess.GetDataTable(sql_ct, objconn, pms_ct)
        Else
            If drPP IsNot Nothing Then 'dr (plan_planinfo)
                Dim pms_tt As New Hashtable From {{"PLANID", PlanID}, {"COMIDNO", ComIDNO}, {"SEQNO", SeqNo}}
                Dim sql_tt As String = ""
                sql_tt += " SELECT DISTINCT b.TechID ,b.TeachCName"
                sql_tt += " FROM PLAN_TEACHER a"
                sql_tt += " JOIN TEACH_TEACHERINFO b ON a.TechID = b.TechID " & vbCrLf
                sql_tt += " WHERE a.PLANID = @PLANID AND a.COMIDNO = @COMIDNO AND a.SEQNO = @SEQNO " & vbCrLf
                dtTecher = DbAccess.GetDataTable(sql_tt, objconn, pms_tt)
            End If
        End If
        'Dim dtTecher As DataTable = DbAccess.GetDataTable(sql, objconn)
        If Not dtTecher Is Nothing Then
            TechName.Text = ""
            CTName.Value = ""
            For Each dr3 As DataRow In dtTecher.Rows
                If TechName.Text <> "" Then TechName.Text &= ","
                TechName.Text &= dr3("TeachCName").ToString

                If CTName.Value <> "" Then CTName.Value &= ","
                CTName.Value &= dr3("TechID").ToString
            Next
            If dtTecher.Rows.Count = 0 Then Button4.Visible = True
        End If

        Call CreateClassTime(drPP)
        Call CreateTrainDesc(drPP)
        'If TPeriod.SelectedIndex = 0 Then Common.SetListItem(TPeriod, TIMS.Get_Plan_VerReport(PlanID, ComIDNO, SeqNo, "TPeriod"))

        '若是班級尚未有資料，可使用計劃時的老師
        If Me.TechName.Text = "" Then
            Me.CTName.Value = TIMS.Get_Plan_VerReport(PlanID, ComIDNO, SeqNo, "TecherID", objconn)
            Me.TechName.Text = TIMS.Get_TeachCName(Me.CTName.Value, objconn) 'TIMS.Get_TeacherName(Me.CTName.Value)
        End If
    End Sub

    ''' <summary>訓練計畫資料顯示--PLAN_PLANINFO</summary>
    ''' <param name="dr"></param>
    ''' <param name="PlanID"></param>
    ''' <param name="ComIDNO"></param>
    ''' <param name="SeqNO"></param>
    Sub SHOW_PLANINFO(ByRef dr As DataRow, ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNO As String)
        'Dim dr As DataRow = Nothing
        Dim sParms As New Hashtable From {{"PlanID", PlanID}, {"ComIDNO", ComIDNO}, {"SeqNO", SeqNO}}
        Dim sql As String = ""
        sql = " SELECT * FROM dbo.PLAN_PLANINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        dr = DbAccess.GetOneRow(sql, objconn, sParms)
        If dr Is Nothing Then Return

        '申請階段 
        Dim v_AppStage As String = Convert.ToString(dr("AppStage"))
        Common.SetListItem(AppStage, v_AppStage)

        CredPoint.Text = If(IsDBNull(dr("CredPoint")), " 0", Convert.ToString(dr("CredPoint")))

        '通俗職類 SHARE_CJOB
        Dim dr99 As DataRow = TIMS.Get_SHARECJOB(Convert.ToString(dr("CJOB_UNKEY")), objconn)
        txtCJOB_NAME.Text = If(dr99 IsNot Nothing, dr99("CJOB_NAME"), "")
        cjobValue.Value = If(dr99 IsNot Nothing, dr99("CJOB_UNKEY"), "")

        RoomName.Text = dr("RoomName").ToString
        Common.SetListItem(FactMode, dr("FactMode").ToString)
        FactModeOther.Text = dr("FactModeOther").ToString
        ConNum.Text = Convert.ToString(dr("ConNum"))
        ContactName.Text = dr("ContactName").ToString
        'ContactPhone.Text = dr("ContactPhone").ToString
        Dim hCtPhone As New Hashtable
        TIMS.CHK_ContactPhoneFMT(Convert.ToString(dr("ContactPhone")), hCtPhone)
        ContactPhone_1.Text = hCtPhone("ContactPhone_1")
        ContactPhone_2.Text = hCtPhone("ContactPhone_2")
        ContactPhone_3.Text = hCtPhone("ContactPhone_3")
        Dim hCtMobile As New Hashtable
        TIMS.CHK_ContactMobileFMT(Convert.ToString(dr("ContactMobile")), hCtMobile)
        ContactMobile_1.Text = hCtMobile("ContactMobile_1")
        ContactMobile_2.Text = hCtMobile("ContactMobile_2")
        ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
        ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
        ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
        ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
        ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)

        ContactEmail.Text = dr("ContactEmail").ToString
        ContactFax.Text = dr("ContactFax").ToString
        Common.SetListItem(ClassCate, dr("ClassCate"))
        If dr("SciPlaceID").ToString <> "" Then Common.SetListItem(SciPlaceID, dr("SciPlaceID").ToString)
        If dr("TechPlaceID").ToString <> "" Then Common.SetListItem(TechPlaceID, dr("TechPlaceID").ToString)
        If dr("SciPlaceID2").ToString <> "" Then Common.SetListItem(SciPlaceID2, dr("SciPlaceID2").ToString)
        If dr("TechPlaceID2").ToString <> "" Then Common.SetListItem(TechPlaceID2, dr("TechPlaceID2").ToString)
        If dr("AddressSciPTID").ToString <> "" Then Common.SetListItem(AddressSciPTID, dr("AddressSciPTID").ToString)
        If dr("AddressTechPTID").ToString <> "" Then Common.SetListItem(AddressTechPTID, dr("AddressTechPTID").ToString)

        IsBusiness.Checked = If(IsDBNull(dr("IsBusiness")), False, If(Convert.ToString(dr("IsBusiness")) = "N", False, True))

        EnterpriseName.Text = dr("EnterpriseName").ToString

        ClassCount.Value = If(Convert.ToString(dr("ClassCount")) <> "", Convert.ToString(dr("ClassCount")), "1")
    End Sub

    Sub CreateTrainDesc(drPP As DataRow)
        If drPP Is Nothing Then Return
        Dim v_RID As String = Convert.ToString(drPP("RID"))
        Dim v_PlanID As String = Convert.ToString(drPP("PLANID"))
        Dim v_ComIDNO As String = Convert.ToString(drPP("COMIDNO"))
        Dim v_SeqNO As String = Convert.ToString(drPP("SEQNO"))

        Dim parms_t As New Hashtable From {{"RID", v_RID}}
        Dim sql_t As String = " SELECT TECHID ,TEACHCNAME FROM TEACH_TEACHERINFO WHERE RID=@RID"
        dtTeacherInfo = DbAccess.GetDataTable(sql_t, objconn, parms_t)

        '20090313 andy edit 與 首頁>>訓練機構管理>>班級查詢作業 、首頁>>訓練機構管理>>計畫變更申請  一致 'sql += " ORDER  BY  PTDID"
        Dim parms As New Hashtable From {{"PLANID", v_PlanID}, {"COMIDNO", v_ComIDNO}, {"SEQNO", v_SeqNO}}
        Dim sql As String = ""
        sql &= " SELECT * FROM PLAN_TRAINDESC WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        sql &= " ORDER BY STrainDate,PName ,PTDID"   '20090401 andy edit 變更排序欄位為 STrainDate,PName
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Datagrid3Table.Style.Item("display") = "none"
            Return
        End If

        'Datagrid3Table.Style.Item("display") = "inline"
        '上面是原寫法
        Datagrid3Table.Style.Item("display") = ""
        With Datagrid3
            .DataSource = dt
            .DataKeyField = "PTDID"
            .DataBind()
        End With
    End Sub

    Sub CreateClassTime(ByRef drPP As DataRow)
        'If drPP Is Nothing Then Return
        'Dim v_RID As String = Convert.ToString(drPP("RID"))
        Dim dt As DataTable = Nothing
        If drPP Is Nothing Then
            dt = Session("Plan_OnClass")
        Else
            Dim v_PlanID As String = Convert.ToString(drPP("PLANID"))
            Dim v_ComIDNO As String = Convert.ToString(drPP("COMIDNO"))
            Dim v_SeqNO As String = Convert.ToString(drPP("SEQNO"))
            Dim parms As New Hashtable From {{"PLANID", v_PlanID}, {"COMIDNO", v_ComIDNO}, {"SEQNO", v_SeqNO}}
            Dim sql As String = "SELECT * FROM PLAN_ONCLASS WHERE PLANID=@PLANID and COMIDNO=@COMIDNO and SEQNO=@SEQNO "
            dt = DbAccess.GetDataTable(sql, objconn, parms)
            dt.Columns("POCID").AutoIncrement = True
            dt.Columns("POCID").AutoIncrementSeed = -1
            dt.Columns("POCID").AutoIncrementStep = -1
            Session("Plan_OnClass") = dt
        End If
        If dt Is Nothing Then
            DataGrid1.Visible = False
            Return
        ElseIf dt.Select(Nothing, Nothing, DataViewRowState.CurrentRows).Length = 0 Then
            DataGrid1.Visible = False
            Return
        End If
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Weeks As Label = e.Item.FindControl("Weeks1")
                Dim Times As Label = e.Item.FindControl("Times1")
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button6")
                Dim btn2 As Button = e.Item.FindControl("Button7")
                btn1.Enabled = Button5.Enabled
                btn2.Enabled = Button5.Enabled
                Weeks.Text = drv("Weeks").ToString
                Times.Text = drv("Times").ToString
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn2.CommandArgument = drv("POCID")
            Case ListItemType.EditItem
                Dim Weeks As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times As TextBox = e.Item.FindControl("Times2")
                Dim btn1 As Button = e.Item.FindControl("Button8")
                Dim btn2 As Button = e.Item.FindControl("Button9")
                Dim drv As DataRowView = e.Item.DataItem
                With Weeks
                    .Items.Add(New ListItem("==請選擇==", ""))
                    .Items.Add(New ListItem("星期一", "星期一"))
                    .Items.Add(New ListItem("星期二", "星期二"))
                    .Items.Add(New ListItem("星期三", "星期三"))
                    .Items.Add(New ListItem("星期四", "星期四"))
                    .Items.Add(New ListItem("星期五", "星期五"))
                    .Items.Add(New ListItem("星期六", "星期六"))
                    .Items.Add(New ListItem("星期日", "星期日"))
                End With
                Common.SetListItem(Weeks, drv("Weeks").ToString)
                Times.Text = drv("Times").ToString
                btn1.CommandArgument = drv("POCID")
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                DataGrid1.EditItemIndex = e.Item.ItemIndex
            Case "del"
                Dim dt As DataTable
                dt = Session("Plan_OnClass")
                If dt.Select("POCID='" & e.CommandArgument & "'").Length <> 0 Then dt.Select("POCID='" & e.CommandArgument & "'")(0).Delete()
                Session("Plan_OnClass") = dt
                DataGrid1.EditItemIndex = -1
            Case "save"
                Dim dt As DataTable
                Dim dr As DataRow
                Dim Weeks As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times As TextBox = e.Item.FindControl("Times2")
                dt = Session("Plan_OnClass")
                If dt.Select("POCID='" & e.CommandArgument & "'").Length <> 0 Then
                    dr = dt.Select("POCID='" & e.CommandArgument & "'")(0)
                    dr("Weeks") = Weeks.SelectedValue
                    dr("Times") = Times.Text
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                Session("Plan_OnClass") = dt
                DataGrid1.EditItemIndex = -1
            Case "cancel"
                DataGrid1.EditItemIndex = -1
        End Select
        CreateClassTime(Nothing)
        Page.RegisterStartupScript("1111", "<script>window.scroll(0,document.body.scrollHeight)</script>")
    End Sub

    '新增(上課時間!!)
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim dt As DataTable = Nothing

        If Session("Plan_OnClass") Is Nothing Then
            Dim rq_PlanID As String = TIMS.ClearSQM(Request("PlanID"))
            Dim rq_ComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
            Dim rq_SeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
            Dim parms As New Hashtable From {{"PLANID", rq_PlanID}, {"COMIDNO", rq_ComIDNO}, {"SEQNO", rq_SeqNo}}
            Dim sql As String = " SELECT * FROM PLAN_ONCLASS WHERE PLANID = @PLANID AND COMIDNO = @COMIDNO AND SEQNO = @SEQNO "
            dt = DbAccess.GetDataTable(sql, objconn, parms)
            dt.Columns("POCID").AutoIncrement = True
            dt.Columns("POCID").AutoIncrementSeed = -1
            dt.Columns("POCID").AutoIncrementStep = -1
        Else
            dt = Session("Plan_OnClass")
        End If

        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        If Request("PlanID") <> "" Then
            dr("PlanID") = Request("PlanID")
            dr("ComIDNO") = Request("ComIDNO")
            dr("SeqNo") = Request("SeqNo")
        End If
        dr("Weeks") = Weeks.SelectedValue
        dr("Times") = Times.Text
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session("Plan_OnClass") = dt
        Call CreateClassTime(Nothing)
        Page.RegisterStartupScript("1111", "<script>window.scroll(0,document.body.scrollHeight)</script>")
    End Sub


    Sub CheckDataPhone(ByRef ErrMsg As String)
        '【辦公室電話】、【行動電話】至少須擇一填寫
        '2023/2024/ContactPhone
        ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
        ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
        ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
        ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
        ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
        Dim s_ContactPhone As String = TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text)
        Dim s_ContactMobile As String = TIMS.ChangContactMobile(ContactMobile_1.Text, ContactMobile_2.Text)
        If s_ContactPhone = "" AndAlso s_ContactMobile = "" Then ErrMsg &= "請輸入 班別資料-【辦公室電話】、【行動電話】至少須擇一填寫" & vbCrLf

        If s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_1.Text) Then
            ErrMsg &= "班別資料-【辦公室電話】區碼，僅能為數字" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso Not ContactPhone_1.Text.StartsWith("0") Then
            ErrMsg &= "班別資料-【辦公室電話】區碼，第1碼應該為0" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso ContactPhone_1.Text.Length < 2 Then
            ErrMsg &= "班別資料-【辦公室電話】區碼，長度須大於1" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_2.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_2.Text) Then
            ErrMsg &= "班別資料-【辦公室電話】電話(8碼)，僅能為數字" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso ContactPhone_3.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_3.Text) Then
            ErrMsg &= "班別資料-【辦公室電話】分機，僅能為數字" & vbCrLf
        ElseIf s_ContactPhone <> "" AndAlso (ContactPhone_1.Text = "" OrElse ContactPhone_2.Text = "") Then
            ErrMsg &= "班別資料-【辦公室電話】不為空，請填寫完整(區碼與電話)為必填" & vbCrLf
        End If

        If s_ContactMobile <> "" AndAlso ContactMobile_1.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactMobile_1.Text) Then
            ErrMsg &= "班別資料-【行動電話】手機前4碼，僅能為數字" & vbCrLf
        ElseIf s_ContactMobile <> "" AndAlso ContactMobile_1.Text <> "" AndAlso Not ContactMobile_1.Text.StartsWith("0") Then
            ErrMsg &= "班別資料-【行動電話】手機前4碼，第1碼應該為0" & vbCrLf
        ElseIf s_ContactMobile <> "" AndAlso ContactMobile_2.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactMobile_2.Text) Then
            ErrMsg &= "班別資料-【行動電話】手機後6碼，僅能為數字" & vbCrLf
        ElseIf s_ContactMobile <> "" AndAlso (ContactMobile_1.Text = "" OrElse ContactMobile_2.Text = "") Then
            ErrMsg &= "班別資料-【行動電話】不為空，請填寫完整(前4碼與後6碼)為必填" & vbCrLf
        End If
    End Sub

    '儲存按鈕 (開班轉入)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        '檢核登入者的計畫 異常為False
        If Not ChkTPlanID(Me) Then
            Common.MessageBox(Me, cst_errmsg4)
            Return 'Exit Sub
        End If
        Dim sErrorMsg As String = ""
        Dim rst As Boolean = ChkSTDate(sErrorMsg)
        If sErrorMsg <> "" Then
            Common.MessageBox(Me, sErrorMsg)
            Return 'Exit Sub
        End If
        'Dim sErrMsg As String = ""
        CheckDataPhone(sErrorMsg)
        If sErrorMsg <> "" Then
            '有錯誤訊息
            'sm.LastErrorMessage = sErrMsg
            Common.MessageBox(Me, sErrorMsg)
            Return 'Exit Sub 'Return False '不可儲存
        End If

        clsid.Value = TIMS.ClearSQM(clsid.Value) '班別代碼
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text) '期別
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        rq_OCID = TIMS.ClearSQM(rq_OCID)

        '2006/03/ add conn by matt
        '檢查班級代碼是否被使用過
        Dim sql_cr As String = " SELECT 'X' FROM CLASS_CLASSINFO WHERE CLSID = @CLSID AND PLANID = @PLANID AND CYCLTYPE = @CYCLTYPE AND RID = @RID "
        Dim parms_cr As New Hashtable From {{"CLSID", Me.clsid.Value}, {"PLANID", PlanID.Value}, {"CYCLTYPE", CyclType.Text}, {"RID", RIDValue.Value}}
        If rq_OCID <> "" Then
            parms_cr.Add("OCID", rq_OCID)
            sql_cr &= " AND OCID != @OCID "
        End If
        Dim dtX As DataTable = DbAccess.GetDataTable(sql_cr, objconn, parms_cr)
        If dtX.Rows.Count > 0 Then
            Common.MessageBox(Me, "新增開班資料重複(該機構在當年度計畫有相同的班別代碼與期別!!)")
            Return ' Exit Sub
        End If

        'Dim sql As String = ""
        If rq_OCID.Trim = "" Then
            Dim pms_cc As New Hashtable From {{"PlanID", PlanID.Value}, {"ComIDNO", ComIDNO.Value}, {"SeqNO", SeqNO.Value}}
            Dim sql_cc As String = " SELECT 1 FROM CLASS_CLASSINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
            Dim dtCC As DataTable = DbAccess.GetDataTable(sql_cc, objconn, pms_cc)
            If dtCC.Rows.Count > 0 Then
                Common.MessageBox(Me, "新增開班資料重複(已有轉班資料!!!)")
                Return '  Exit Sub
            End If
        End If

        Call SaveData1()  '儲存(CLASS_CLASSINFO)
    End Sub

    '儲存(CLASS_CLASSINFO)
    Sub SaveData1()
        'Sub SaveData1(ByVal ss As String)
        'Dim vsSEnterDate As String = TIMS.GetMyValue(ss, "vsSEnterDate")
        'Dim vsFEnterDate As String = TIMS.GetMyValue(ss, "vsFEnterDate")
        'Dim vsOnShellDate As String = TIMS.GetMyValue(ss, "vsOnShellDate")
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        ComIDNO.Value = TIMS.ClearSQM(ComIDNO.Value)
        SeqNO.Value = TIMS.ClearSQM(SeqNO.Value)

        Dim parms As New Hashtable From {{"PLANID", PlanID.Value}, {"COMIDNO", ComIDNO.Value}, {"SEQNO", SeqNO.Value}}
        Dim sql As String = "" 'sql = ""
        sql &= " SELECT * FROM PLAN_TEACHER WHERE TechTYPE = 'A' " 'TechTYPE: A:師資/B:助教
        sql &= " AND PLANID = @PLANID AND COMIDNO = @COMIDNO AND SEQNO = @SEQNO "
        'parms.Clear()
        Dim dtPt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        'Call TIMS.OpenDbConn(tConn)
        'Dim sql As String = ""
        Dim tConn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)

        Try
            '2006/03/ add conn by matt
            Dim dr As DataRow = Nothing
            Dim dt As DataTable = Nothing
            Dim da As SqlDataAdapter = Nothing
            If rq_OCID.Trim <> "" Then
                sql = " SELECT * FROM CLASS_CLASSINFO WHERE OCID = '" & rq_OCID & "' "
            Else
                sql = " SELECT * FROM CLASS_CLASSINFO WHERE 1<>1 "
            End If
            dt = DbAccess.GetDataTable(sql, da, trans)
            Dim iOCID_New As Integer = 0

            If ClassEng.Value.ToString <> "" Then
                If ClassEng.Value.ToString <> ClassEngName.Text Then ClassEngName.Text = ClassEng.Value.ToString
            End If

            TechName.Text = TIMS.ClearSQM(TechName.Text)

            If dt.Rows.Count = 0 Then
                dr = dt.NewRow
                dt.Rows.Add(dr)
                iOCID_New = DbAccess.GetNewId(trans, "CLASS_CLASSINFO_OCID_SEQ,CLASS_CLASSINFO,OCID") 'fix ora-00001 違反必須唯一的限制條件
                dr("OCID") = iOCID_New

                '2019-02-19 add 記操作歷程（sys_trans_log）'新增
                Dim BeforeValues As String = Get_BeforeValuesStr1(iOCID_New) '修改後資料-新增
                Dim t_iSql As String = ""
                't_iSql = ""
                t_iSql += " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                t_iSql += " VALUES(@SessionID, FORMAT(GETDATE(),'yyyy-MM-dd HH:mm:ss.fff'), '/TC/01/TC_01_004_BusAdd.aspx', @UserID, 'Insert', 'CLASS_CLASSINFO', '',@BeforeValues, '') "
                Dim myParam As New Hashtable
                myParam.Clear()
                myParam.Add("SessionID", sm.SessionID.ToString)
                myParam.Add("UserID", sm.UserInfo.UserID)
                myParam.Add("BeforeValues", BeforeValues)
                DbAccess.ExecuteNonQuery(t_iSql, objconn, myParam)
            Else
                dr = dt.Rows(0)
                iOCID_New = dr("OCID")
                If iOCID_New <> Val(rq_OCID) Then
                    '異常發生問題
                    DbAccess.RollbackTrans(trans)
                    TIMS.CloseDbConn(tConn)
                    Common.MessageBox(Me, "儲存失敗!!")
                    Exit Sub
                End If

                '2019-02-19 add 記操作歷程（sys_trans_log）'修改
                Dim Conditions As String = ""
                Conditions &= "OCID=" & iOCID_New
                Dim BeforeValues As String = Get_BeforeValuesStr2(iOCID_New) '操作前資料
                Dim AfterValues As String = Get_BeforeValuesStr1(iOCID_New) '修改後資料
                Dim t_iSql As String = ""
                't_iSql = ""
                t_iSql += " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
                t_iSql += " VALUES(@SessionID, FORMAT(GETDATE(),'yyyy-MM-dd HH:mm:ss.fff'), '/TC/01/TC_01_004_BusAdd.aspx', @UserID, 'Update', 'CLASS_CLASSINFO',@Conditions ,@BeforeValues,@AfterValues) "
                Dim myParam As New Hashtable
                myParam.Clear()
                myParam.Add("SessionID", sm.SessionID.ToString)
                myParam.Add("UserID", sm.UserInfo.UserID)
                myParam.Add("Conditions", Conditions)
                myParam.Add("BeforeValues", BeforeValues)
                myParam.Add("AfterValues", AfterValues)
                DbAccess.ExecuteNonQuery(t_iSql, objconn, myParam)
            End If

            dr("ClassNum") = ClassCount.Value '班數
            dr("CLSID") = clsid.Value
            dr("PlanID") = PlanID.Value '(PCS)
            dr("Years") = Years.Value
            dr("RID") = RIDValue.Value

            Dim vCyclType As String = TIMS.ClearSQM(CyclType.Text)
            If vCyclType = "" Then vCyclType = TIMS.cst_Default_CyclType
            vCyclType = TIMS.FmtCyclType(vCyclType)
            dr("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull) 'CyclType.Text

            dr("ClassCName") = ClassCName.Text
            If ClassEng.Value.ToString <> "" Then
                If ClassEng.Value.ToString <> ClassEngName.Text Then ClassEngName.Text = ClassEng.Value.ToString
            End If
            dr("ClassEngName") = If(ClassEngName.Text = "", Convert.DBNull, ClassEngName.Text)
            'dr("Content") = Content.Text
            dr("Purpose") = Purpose.Text
            '在職VALUE '/職前(0)/在職(1)/
            dr("TPropertyID") = If(Convert.ToString(dr("TPropertyID")) <> "", Convert.ToString(dr("TPropertyID")), "1")
            '訓練業別
            dr("TMID") = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, jobValue.Value, trainValue.Value)
            '通俗職類
            dr("CJOB_UNKEY") = If(cjobValue.Value <> "", CInt(cjobValue.Value), Convert.DBNull)
            dr("SEnterDate") = If(vsSEnterDate <> "", vsSEnterDate, dr("SEnterDate"))
            dr("FEnterDate") = If(vsFEnterDate <> "", vsFEnterDate, dr("FEnterDate"))
            dr("OnShellDate") = If(vsOnShellDate <> "", vsOnShellDate, dr("OnShellDate"))
            dr("CheckInDate") = If(CheckInDate.Text <> "", CheckInDate.Text, STDate.Text)
            'If ExamDate.Text = "" Then   '20100329 add 甄試時段
            '    dr("ExamDate") = Convert.DBNull
            '    dr("ExamPeriod") = Convert.DBNull
            'Else
            '    dr("ExamDate") = ExamDate.Text
            '    dr("ExamPeriod") = ExamPeriod.SelectedValue
            'End If
            dr("TDeadline") = TDeadline.SelectedValue
            dr("STDate") = STDate.Text
            dr("FTDate") = FTDate.Text
            'dr("TaddressZip") = city_code.Value
            'dr("TAddress") = TAddress.Text
            dr("THours") = THours.Text
            dr("TNum") = TNum.Text
            'If TPeriod.SelectedValue <> "" Then
            '    dr("TPeriod") = TPeriod.SelectedValue
            'Else
            '    dr("TPeriod") = TPeriod.SelectedValue
            'End If
            ' open mark by nick 060308
            Dim sNORID As String = ""
            If (NotOpen.Checked) Then
                For i As Integer = 0 To NORID.Items.Count - 1
                    If NORID.Items(i).Selected AndAlso NORID.Items(i).Value <> "" Then
                        sNORID &= String.Concat(If(sNORID <> "", ",", ""), NORID.Items(i).Value)
                    End If
                Next
            End If

            dr("NotOpen") = If(NotOpen.Checked, "Y", "N")
            dr("NORID") = If(NotOpen.Checked, sNORID, Convert.DBNull)
            dr("OtherReason") = If(NotOpen.Checked, If(OtherReason.Text = "", Convert.DBNull, OtherReason.Text), Convert.DBNull)
            dr("IsApplic") = If(IsApplic.Checked, "Y", "N")
            dr("Relship") = Relship.Value
            dr("ComIDNO") = ComIDNO.Value '(PCS)
            dr("SeqNO") = SeqNO.Value '(PCS)
            dr("IsCalculate") = "Y"
            dr("IsSuccess") = "Y"
            TechName.Text = TIMS.ClearSQM(TechName.Text)
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                dr("IsFullDate") = "N" '產學訓預設值為否
                'dr("CTName") = Left("X:" & Me.TechName.Text, 40) 
                'dr("CTName") = Left(TechName.Text, 40)
            Else
                dr("IsFullDate") = IsFullDate.SelectedValue
                'dr("CTName") = Left(TechName.Text, 40)
            End If
            dr("CTName") = TIMS.Get_CTNAME1(TechName.Text)
            dr("IsBusiness") = If(IsBusiness.Checked = True, "Y", "N")
            'dr("EnterpriseName") = If(EnterpriseName.Text <> "", EnterpriseName.Text, Convert.DBNull)

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)
            'If rq_OCID <> "" Then iOCID_New = Val(rq_OCID)

            'ContactPhone.Text = TIMS.ClearSQM(ContactPhone.Text)
            ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
            ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
            ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
            ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
            ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
            'Dim s_ContactPhone As String = If(fg_phone_2024, TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text), ContactPhone.Text)
            Dim s_ContactPhone As String = TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text)
            'dr("ContactPhone") = If(s_ContactPhone <> "", s_ContactPhone, Convert.DBNull)
            Dim s_ContactMobile As String = TIMS.ChangContactMobile(ContactMobile_1.Text, ContactMobile_2.Text)
            'dr("ContactMobile") = If(s_ContactMobile <> "", s_ContactMobile, Convert.DBNull)

            Dim v_FactMode As String = TIMS.GetListValue(FactMode)
            sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID = '" & PlanID.Value & "' AND ComIDNO = '" & ComIDNO.Value & "' AND SeqNo = '" & SeqNO.Value & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            dr = dt.Rows(0)
            dr("CredPoint") = CredPoint.Text
            dr("RoomName") = RoomName.Text
            dr("FactMode") = v_FactMode 'FactMode.SelectedValue
            dr("FactModeOther") = If(v_FactMode = "99", FactModeOther.Text, Convert.DBNull)
            dr("ConNum") = If(ConNum.Text <> "", Val(ConNum.Text), Convert.DBNull)
            dr("ContactName") = ContactName.Text

            dr("ContactPhone") = If(s_ContactPhone <> "", s_ContactPhone, Convert.DBNull)
            dr("ContactMobile") = If(s_ContactMobile <> "", s_ContactMobile, Convert.DBNull)
            dr("ContactEmail") = ContactEmail.Text
            dr("ContactFax") = ContactFax.Text
            dr("ClassCate") = ClassCate.SelectedValue
            dr("TransFlag") = "Y"
            '正常後關閉 
            dr("IsBusiness") = If(IsBusiness.Checked = True, "Y", "N")
            dr("EnterpriseName") = If(EnterpriseName.Text <> "", EnterpriseName.Text, Convert.DBNull)
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)

            sql = " SELECT * FROM PLAN_ONCLASS WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, da, trans)
            Dim TempDataTable As DataTable = Session("Plan_OnClass")
            dt = TempDataTable.Copy
            DbAccess.UpdateDataTable(dt, da, trans)

            '儲存 班級申請老師(CLASS_TEACHER)
            'Call SAVE_CLASS_TEACHER(iOCID_New, dtPt, trans)

            '假如有學員的話,要更新學員學號--------------------Strat
            If TBclass.Value <> "" Then
                If OldClassID.Value <> TBclass.Value Then  'TBclass_id.Text Then
                    TBclass_id.Text = TBclass.Value.ToString
                    sql = " SELECT SOCID,StudentID FROM CLASS_STUDENTSOFCLASS WHERE OCID='" & iOCID_New & "'"
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count <> 0 Then
                        For Each dr In dt.Rows
                            dr("StudentID") = Replace(dr("StudentID"), OldClassID.Value, TBclass_id.Text)
                        Next
                        DbAccess.UpdateDataTable(dt, da, trans)
                    End If
                End If
            End If
            '假如有學員的話,要更新學員學號--------------------End
            'DbAccess.RollbackTrans(trans)
            DbAccess.CommitTrans(trans)
            'Call TIMS.CloseDbConn(tConn)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Common.MessageBox(Me, "儲存失敗!!")
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
            'Call TIMS.CloseDbConn(tConn)
            'DbAccess.RollbackTrans(trans)
            'Throw ex
        End Try

        '重複 為 true
        Dim Double_flag As Boolean = False 'false 沒有重複。
        Dim iDouble As Integer = 0
        Do
            iDouble += 1
            Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
            '刪除重複轉班資料。
            Try
                '至少判斷1次是否有重複轉班
                Double_flag = TIMS.sUtl_DeleteDoubleClassInfo(PlanID.Value, ComIDNO.Value, SeqNO.Value, objconn)
            Catch ex As Exception
                Double_flag = False '只要有1次失敗就算了吧。
            End Try
            If iDouble >= 5 Then
                '判斷5次也太 ...
                Double_flag = False
            End If
        Loop Until Not Double_flag '直到沒有重複。

        If Not Me.ViewState("ClassSearchStr") Is Nothing Then Session("ClassSearchStr") = Me.ViewState("ClassSearchStr")
        Common.RespWrite(Me, "<script>alert('儲存成功!');location.href='TC_01_004.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    '回查詢頁面
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not Me.ViewState("ClassSearchStr") Is Nothing Then Session("ClassSearchStr") = Me.ViewState("ClassSearchStr")
        'Response.Redirect("TC_01_004.aspx?ID=" & Request("ID") & "")
        Dim url1 As String = "TC_01_004.aspx?ID=" & Request("ID") & ""
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '上課時間DG3
    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim STrainDateLabel As Label = e.Item.FindControl("STrainDateLabel")
                Dim PNameLabel As Label = e.Item.FindControl("PNameLabel")
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel")
                Dim PContText As TextBox = e.Item.FindControl("PContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1")
                Dim drpPTID As DropDownList = e.Item.FindControl("drpPTID")
                Dim Tech1Value As HtmlInputHidden = e.Item.FindControl("Tech1Value")
                Dim Tech1Text As TextBox = e.Item.FindControl("Tech1Text")
                Dim Tech2Value As HtmlInputHidden = e.Item.FindControl("Tech2Value")
                Dim Tech2Text As TextBox = e.Item.FindControl("Tech2Text")
                If drv("STrainDate").ToString <> "" Then STrainDateLabel.Text = Common.FormatDate(drv("STrainDate").ToString)
                PNameLabel.Text = drv("PName").ToString
                PHourLabel.Text = drv("PHour").ToString
                PContText.Text = drv("PCont").ToString
                If drv("Classification1").ToString <> "" Then
                    Common.SetListItem(drpClassification1, drv("Classification1").ToString)
                    Select Case drpClassification1.SelectedValue
                        Case "1" '學科
                            If Request("ComIDNO") Is Nothing Then
                                If Me.ComIDNO.Value = "" Then Me.ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                                drpPTID = TIMS.Get_SciPTID(drpPTID, Me.ComIDNO.Value, 1, objconn)
                            Else
                                drpPTID = TIMS.Get_SciPTID(drpPTID, Request("ComIDNO"), 1, objconn)
                            End If
                        Case "2" '術科
                            If Request("ComIDNO") Is Nothing Then
                                If Me.ComIDNO.Value = "" Then Me.ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                                drpPTID = TIMS.Get_TechPTID(drpPTID, Me.ComIDNO.Value, 1, objconn)
                            Else
                                drpPTID = TIMS.Get_TechPTID(drpPTID, Request("ComIDNO"), 1, objconn)
                            End If
                    End Select
                    If Convert.ToString(drv("PTID")) <> "" Then Common.SetListItem(drpPTID, drv("PTID"))
                End If
                If Convert.ToString(drv("TechID")) <> "" Then
                    Tech1Value.Value = drv("TechID").ToString
                    Tech1Text.Text = TIMS.Get_TeacherName(drv("TechID").ToString, dtTeacherInfo)
                End If
                If Convert.ToString(drv("TechID2")) <> "" Then
                    Tech2Value.Value = drv("TechID2").ToString
                    Tech2Text.Text = TIMS.Get_TeachCName(Tech2Value.Value, objconn) '
                End If
        End Select
    End Sub

    Function Get_BeforeValuesStr1(ByVal iOCID As Integer) As String
        Dim rst As String = ""
        'rst = ""
        rst &= "OCID=" & iOCID
        rst &= ",CLSID=" & clsid.Value
        rst &= ",PLANID=" & PlanID.Value '(PCS)
        rst &= ",YEARS=" & Years.Value
        rst &= ",CYCLTYPE=" & CyclType.Text 'If(String.IsNullOrEmpty(CyclType.Text), "01", CyclType.Text)
        rst &= ",RID=" & RIDValue.Value
        rst &= ",CLASSCNAME=" & ClassCName.Text
        rst &= ",CLASSENGNAME=" & ClassEngName.Text
        rst &= ",PURPOSE=" & Purpose.Text
        'rst &= ",TPROPERTYID=" & If(TPropertyID.SelectedValue = "", "1", TPropertyID.SelectedValue)
        rst &= ",TPROPERTYID=1" '/職前(0)/在職(1)/
        rst &= ",TMID=" & If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, jobValue.Value, trainValue.Value)
        rst &= ",SENTERDATE=" & vsSEnterDate
        rst &= ",FENTERDATE=" & vsFEnterDate
        rst &= ",ONSHELLDATE=" & vsOnShellDate
        rst &= ",CHECKINDATE=" & If(CheckInDate.Text = "", STDate.Text, CheckInDate.Text)
        'rst &= ",EXAMDATE=" & ExamDate.Text
        'rst &= ",EXAMPERIOD=" & If(ExamDate.Text = "", "", ExamPeriod.SelectedValue)
        rst &= ",STDATE=" & STDate.Text
        rst &= ",FTDATE=" & FTDate.Text
        rst &= ",THOURS=" & THours.Text
        rst &= ",TNUM=" & TNum.Text
        rst &= ",TDEADLINE=" & TDeadline.SelectedValue

        If NotOpen.Checked Then
            Dim strNORID As String = ""
            For i As Integer = 0 To NORID.Items.Count - 1
                If NORID.Items(i).Selected = True AndAlso NORID.Items(i).Value <> "" Then
                    If strNORID <> "" Then strNORID &= ","
                    strNORID &= NORID.Items(i).Value
                End If
            Next
            rst &= ",NOTOPEN=" & "Y"
            rst &= ",NORID=" & strNORID
            rst &= ",OTHERREASON=" & OtherReason.Text
        Else
            rst &= ",NOTOPEN=N"
            rst &= ",NORID="
            rst &= ",OTHERREASON="
        End If
        rst &= ",ISAPPLIC=" & If(IsApplic.Checked, "Y", "N")
        rst &= ",RELSHIP=" & Relship.Value
        rst &= ",COMIDNO=" & ComIDNO.Value '(PCS)
        rst &= ",SEQNO=" & SeqNO.Value '(PCS)
        rst &= ",ISCALCULATE=Y"
        rst &= ",ISSUCCESS=Y"
        rst &= ",ISFULLDATE=" & If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, "N", IsFullDate.SelectedValue)
        'rst &= ",CTNAME=" & Left(TechName.Text, 40)
        rst &= ",CTNAME=" & TIMS.Get_CTNAME1(TechName.Text)
        rst &= ",MODIFYACCT=" & sm.UserInfo.UserID
        'rst &= ",MODIFYDATE=FORMAT(GETDATE(),'yyyy/MM/dd HH:mm:ss.fff')" '& DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss.fff")
        rst &= ",MODIFYDATE=" & TIMS.Cdate3(Now, "yyyy/MM/dd HH:mm:ss.fff")
        rst &= ",CLASSNUM=" & ClassCount.Value '班數
        rst &= ",ISBUSINESS=" & If(IsBusiness.Checked = True, "Y", "N")
        rst &= ",CJOB_UNKEY=" & cjobValue.Value
        Return rst
    End Function

    Function Get_BeforeValuesStr2(ByVal iOCID As Integer) As String
        'Dim rst As String = ""
        '2019-02-19 add 記操作歷程（sys_trans_log）
        Dim t_BeforeSql As String = ""
        't_BeforeSql = ""
        t_BeforeSql &= " SELECT OCID,CLSID,PLANID,YEARS,CYCLTYPE,RID,CLASSCNAME,CLASSENGNAME,PURPOSE,TPROPERTYID,TMID"
        t_BeforeSql &= " ,SENTERDATE,FENTERDATE,ONSHELLDATE,CHECKINDATE,STDATE,FTDATE,THOURS,TNUM,TDEADLINE,NOTOPEN"
        t_BeforeSql &= " ,ISAPPLIC, RELSHIP, COMIDNO, SEQNO, ISCALCULATE, ISSUCCESS, CTNAME, MODIFYACCT, MODIFYDATE, CLASSNUM"
        t_BeforeSql &= " ,ISFULLDATE, NORID, OTHERREASON, ISBUSINESS, CJOB_UNKEY"
        t_BeforeSql &= " FROM CLASS_CLASSINFO Where OCID = " & iOCID
        Dim t_BeforeRow As DataRow = DbAccess.GetOneRow(t_BeforeSql, objconn)
        'Dim t_BeforeRow As DataRow = t_BeforeTB.Rows(0)
        'Dim t_iSql As String = ""
        'Dim myParam As Hashtable = New Hashtable
        'Dim BeforeValues As String = ""
        'Dim AfterValues As String = ""
        'Dim strNORID As String = ""
        Dim rst As String = ""
        rst &= "OCID=" & Convert.ToString(t_BeforeRow("OCID"))
        rst &= ",CLSID=" & Convert.ToString(t_BeforeRow("CLSID"))
        rst &= ",PLANID=" & Convert.ToString(t_BeforeRow("PLANID"))
        rst &= ",YEARS=" & Convert.ToString(t_BeforeRow("YEARS"))
        rst &= ",CYCLTYPE=" & Convert.ToString(t_BeforeRow("CYCLTYPE"))
        rst &= ",RID=" & Convert.ToString(t_BeforeRow("RID"))
        rst &= ",CLASSCNAME=" & Convert.ToString(t_BeforeRow("CLASSCNAME"))
        rst &= ",CLASSENGNAME=" & Convert.ToString(t_BeforeRow("CLASSENGNAME"))
        rst &= ",PURPOSE=" & Convert.ToString(t_BeforeRow("PURPOSE"))
        rst &= ",TPROPERTYID=" & Convert.ToString(t_BeforeRow("TPROPERTYID")) '/職前(0)/在職(1)/
        rst &= ",TMID=" & Convert.ToString(t_BeforeRow("TMID"))

        rst &= ",SENTERDATE=" & TIMS.Cdate3(t_BeforeRow("SENTERDATE"), "yyyy/MM/dd HH:mm")
        rst &= ",FENTERDATE=" & TIMS.Cdate3(t_BeforeRow("FENTERDATE"), "yyyy/MM/dd HH:mm")
        rst &= ",ONSHELLDATE=" & TIMS.Cdate3(t_BeforeRow("ONSHELLDATE"), "yyyy/MM/dd HH:mm")
        rst &= ",CHECKINDATE=" & TIMS.Cdate3(t_BeforeRow("CHECKINDATE"))

        'rst &= ",EXAMDATE=" & TIMS.cdate3(t_BeforeRow("EXAMDATE"), "yyyy/MM/dd HH:mm:ss.fff")
        'rst &= ",EXAMPERIOD=" & Convert.ToString(t_BeforeRow("EXAMPERIOD"))
        rst &= ",STDATE=" & TIMS.Cdate3(t_BeforeRow("STDATE"))
        rst &= ",FTDATE=" & TIMS.Cdate3(t_BeforeRow("FTDATE"))
        rst &= ",THOURS=" & Convert.ToString(t_BeforeRow("THOURS"))
        rst &= ",TNUM=" & Convert.ToString(t_BeforeRow("TNUM"))
        rst &= ",TDEADLINE=" & Convert.ToString(t_BeforeRow("TDEADLINE"))
        rst &= ",NOTOPEN=" & Convert.ToString(t_BeforeRow("NOTOPEN"))
        rst &= ",NORID=" & Convert.ToString(t_BeforeRow("NORID"))
        rst &= ",OTHERREASON=" & Convert.ToString(t_BeforeRow("OTHERREASON"))
        rst &= ",ISAPPLIC=" & Convert.ToString(t_BeforeRow("ISAPPLIC"))
        rst &= ",RELSHIP=" & Convert.ToString(t_BeforeRow("RELSHIP"))
        rst &= ",COMIDNO=" & Convert.ToString(t_BeforeRow("COMIDNO"))
        rst &= ",SEQNO=" & Convert.ToString(t_BeforeRow("SEQNO"))
        rst &= ",ISCALCULATE=" & Convert.ToString(t_BeforeRow("ISCALCULATE"))
        rst &= ",ISSUCCESS=" & Convert.ToString(t_BeforeRow("ISSUCCESS"))
        rst &= ",ISFULLDATE=" & Convert.ToString(t_BeforeRow("ISFULLDATE"))
        rst &= ",CTNAME=" & Convert.ToString(t_BeforeRow("CTNAME"))
        rst &= ",MODIFYACCT=" & Convert.ToString(t_BeforeRow("MODIFYACCT"))
        rst &= ",MODIFYDATE=" & TIMS.Cdate3(t_BeforeRow("MODIFYDATE"), "yyyy/MM/dd HH:mm:ss.fff")

        rst &= ",CLASSNUM=" & Convert.ToString(t_BeforeRow("CLASSNUM")) '班數
        rst &= ",ISBUSINESS=" & Convert.ToString(t_BeforeRow("ISBUSINESS"))
        rst &= ",CJOB_UNKEY=" & Convert.ToString(t_BeforeRow("CJOB_UNKEY"))
        Return rst
    End Function

End Class