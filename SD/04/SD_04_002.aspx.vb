Imports System.Threading

Partial Class SD_04_002
    Inherits AuthBasePage

    'TEACH_TEACHERINFO,COURSE_COURSEINFO,SYS_HOLIDAY,STUD_TRAININGRESULTS,CLASS_SCHVERIFY
    'CLASS_SCHEDULE,PLAN_SCHEDULE 
    'V_SCHEDULETYPE
    'select * from Auth_AccRWClass  where 1=1 and ocid =99721
    'DataGrid4
    'Call TIMS.OpenDbConn(objconn)
    Const cst_max_Timeout As Integer = 500
    Dim ff3 As String = ""
    Dim g_ErrSql As String = ""
    Private dtCourse As DataTable '可用課程嗎？Course_CourseInfo (CourID,CourseName)
    Private HolidayTable As DataTable
    Const cst_vsDetailTable As String = "DetailTable" 'ViewState(cst_vsDetailTable)
    Dim dtTeacher As DataTable

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Dim blnCanDeleteAuth As Boolean = False
    'Dim PageControler1 As New PageControler
    Dim Key_Degree As DataTable
    Dim Key_Military As DataTable
    Dim Key_Identity As DataTable
    'Dim Gary_day7() As String '每星期的上課時間 星期日~星期六 0~6
    Const Cst_OCIDNULL As String = "此訓練機構，暫無班別供選擇"
    Const cst_holiday As String = "holiday" '假日
    Dim ff As String = ""
    Dim vsErrMsg As String = ""

    Const cst_alertMsg1 As String = "您已經使用過全期排課，所以無法使用本功能!"
    Const cst_alertMsg2 As String = "排課資料不存在，請重新查詢建立排課資料!!"
    Const cst_alertMsg3 As String = "排課日期區間未落在開、結訓日期中 無法修改排課!!"
    Const cst_alertMsg4 As String = "排課時數異常，查無排課資料!!"
    Const cst_alertMsg5 As String = "排課時數已經用完!"
    Const cst_alertMsg6 As String = "班級資訊異常，請重新查詢班級資料!!"
    Const cst_alertMsg7 As String = "排課區間迄止日期有誤，請重新設定 排課區間迄止日期!!"
    Const cst_alertMsg8 As String = "指定日期有誤，請重新選擇指定日期!!"
    Const cst_alertMsg9 As String = "排課日期區間未落在開、結訓日期中 無法新增排課!!"
    Const cst_alertMsg10 As String = "此班級已經輸入成績，不能刪除資料"
    Const cst_alertMsg11 As String = "指定日期有誤，請再確認!!"
    'Const cst_alertMsg12 As String = "該計畫學科 教師 僅限1位!!"
    Const cst_alertMsg13 As String = "系統後端正在排課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!"
    Const cst_alertMsg14 As String = "班級已結訓，不可再修改!!"
    Const cst_alertMsg15 As String = "已審核確認，不可再修改!!"
    Const cst_alertMsg16 As String = "該課程代碼有誤!"
    Const cst_alertMsg17 As String = "查無課程代碼!"

    Const cst_flag As String = ","
    Const cst_DG2_CourseName As Integer = 1
    Const cst_DG2_ClassRoom As Integer = 2
    Const cst_DG2_Teacher1 As Integer = 3
    Const cst_DG2_Teacher2 As Integer = 4
    Const cst_DG2_Teacher3 As Integer = 5
    Const cst_DG2_Teacher4 As Integer = 6

    'Dim blnRC As Boolean = True
    Public Shared blnRC As Boolean = True

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        'blnCanDeleteAuth = TIMS.CheckAuthUse(sm.UserInfo.UserID, objconn, 1)
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'PageControler1 = Me.FindControl("PageControler1")
        'Me.HyperLink1.NavigateUrl = "../../Doc/Class_Schedule_frm.zip"
        'Me.HyperLink1.NavigateUrl = "../../Doc/ClassScheduletForrmat_v12.zip"
        Me.HyperLink1.NavigateUrl = "../../Doc/ClassScheduletForrmat_v13.zip"

        'Dim blnRC As Boolean = True
        If TIMS.sUtl_ChkTest() Then blnRC = False '正式環境為TRUE'測試環境為FALSE

        '檢查帳號的功能權限 Start
        'Button2.Enabled = True
        'If Not au.blnCanSech Then Button2.Enabled = False
        'If Not au.blnCanSech Then TIMS.Tooltip(Button2, "您無權限使用該功能", True)
        '檢查帳號的功能權限 End

        msg.Text = ""

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?btnName=Button1');"
        Button11.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        'Button1.Style("display") = "none"
        '依sm.UserInfo.PlanID取得PlanKind  '1:自辦(內訓) 2:委外
        'sql = "SELECT PlanKind FROM ID_Plan WHERE PlanID='" & sm.UserInfo.PlanID & "'"
        'dr = DbAccess.GetOneRow(sql, objConn)

        '取得該業務單位的所有老師
        'Dim dtTeacher As DataTable = Nothing
        dtTeacher = New DataTable
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID

        If RIDValue.Value <> "" Then RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim TeacherSql As String = ""
        TeacherSql &= " SELECT TeacherID,TechID,TeachCName "
        TeacherSql &= " FROM Teach_TeacherInfo "
        TeacherSql &= " WHERE WorkStatus = '1' AND RID = '" & RIDValue.Value & "' "
        dtTeacher.Load(DbAccess.GetReader(TeacherSql, objconn))

        Button2.Attributes("onclick") = "return chkdata();"
        Button15.Attributes("onclick") = "return chkdata();"
        'Button3.Attributes("onclick") = "return Locker_bt3();"
        Button9.Attributes("onclick") = "return CheckNewCourse();"  '新增排課
        Button10.Attributes("onclick") = "return DelClass();"
        Button12.Attributes("onclick") = "return DelALLClass();"

        'CourseID.Attributes("onKeypress") = "GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','Room');"
        'CourseID.Attributes("onclick") = "Course('Add','CourseID','CourseIDValue');"
        'OLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');" '''Class TIMS.CreateTeacherScript
        'OLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','OLessonTeah2');"
        'OLessonTeah1.Attributes("ondblclick") = "Get_Teah('OLessonTeah1','OLessonTeah1Value');"
        'OLessonTeah2.Attributes("ondblclick") = "Get_Teah('OLessonTeah2','OLessonTeah2Value');"
        'CourseID.Attributes("onDblClick") = "Course_search();"
        CourseID.Attributes("onDblClick") = "Course_search('Edit','CourseIDValue','CourseID');"

        Dim sjj As String = "GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','OLessonTeah3Value','OLessonTeah3','OLessonTeah4Value','OLessonTeah4','Room');"
        CourseID.Attributes("onClick") = sjj '"GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','Room');"
        'CourseID.Attributes("onChange") = sjj '"GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','Room');"
        CourseID.Style.Item("CURSOR") = "hand"

        Call TIMS.CreateTeacherScript(Me, RIDValue.Value, dtTeacher)

        'If RIDValue.Value <> "" Then
        '    TIMS.CreateTeacherScript(Me, RIDValue.Value, dtTeacher)
        'Else
        '    TIMS.CreateTeacherScript(Me, sm.UserInfo.RID, dtTeacher)
        'End If

        OLessonTeah1.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','1','','');")
        OLessonTeah1.Style.Item("CURSOR") = "hand"
        OLessonTeah2.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','2','OLessonTeah2','OLessonTeah2Value');")
        OLessonTeah2.Style.Item("CURSOR") = "hand"
        OLessonTeah3.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','3','OLessonTeah3','OLessonTeah3Value');")
        OLessonTeah3.Style.Item("CURSOR") = "hand"
        OLessonTeah4.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','4','OLessonTeah4','OLessonTeah4Value');")
        OLessonTeah4.Style.Item("CURSOR") = "hand"

        ClassSort1.Attributes("onclick") = "CheckClassTime();"
        ClassSort2.Attributes("onclick") = "GetClassTime(1);"
        ClassSort3.Attributes("onclick") = "GetClassTime(2);"
        ClassSort4.Attributes("onclick") = "GetClassTime(3);"
        ClassSort5.Attributes("onclick") = "GetClassTime(4);"
        LinkButton2.Attributes("onclick") = "ShowCourseList(this);return false;"
        'LinkButton4.Attributes("onclick") = "ShowCourseList4(this);return false;"
        'OpShowCourse
        btnShowCourse.Attributes("onclick") = "OpShowCourse();return false;"
        ClearButton.Attributes("onclick") = "document.form1.OCID2.value='';document.form1.OCIDValue2.value='';"
        LoadIntoClass.Attributes("onclick") = "return CheckImportData();"
        OCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        Button13.Style("display") = "none"

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button1")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Style("CURSOR") = "hand"
        End If

        '開放正式機使用 by AMU 201012
        '正式機判斷(新增排課資料)暫不顯示 '測試
        'If ReportQuery.GetSmartQueryPath = "TIMS" Then '測試
        '    Button15.Visible = False '正式機不顯示 測試機顯示
        'Else
        '    Button15.Visible = True '正式機不顯示 測試機顯示
        'End If '測試

        If Not IsPostBack Then
            Call Utl_ShowX1()

            'ViewState("GUID1") = TIMS.GetGUID
            'Session("GUID1") = ViewState("GUID1")
            SearchTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
            CourseTable.Style.Item("display") = "none"
            DetailTable.Style.Item("display") = "none"

            If Request("k") = "SD_04_002_ADD" Then
                center.Text = TIMS.GetMyValue(Session("SearchStr"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("SearchStr"), "RIDValue")

                Button12.Enabled = False
                TIMS.Tooltip(Button12, "新增排課，不提供刪除排課資料", True)

                Call sUtl_Search1() 'Button1_Click(sender, e)'sender, e
                TypeRadio.Items(0).Selected = True

                Common.SetListItem(Me.OCID, TIMS.GetMyValue(Session("SearchStr"), "OCIDValue1"))
                Me.OCID_SelectedIndexChanged(sender, e)

                CheckBox1.Checked = True
                CSDate.Text = STDate.Value
                CFDate.Text = FTDate.Value
            Else
                RIDValue.Value = sm.UserInfo.RID
                center.Text = sm.UserInfo.OrgName

                Button12.Enabled = False
                TIMS.Tooltip(Button12, "新增排課，不提供刪除排課資料", True)

                Call sUtl_Search1() 'Button1_Click(sender, e)sender, e
                TypeRadio.Items(0).Selected = True
            End If

            Message.Text = ""
        End If

        If dtCourse Is Nothing OrElse HolidayTable Is Nothing Then
            Dim v_OCID As String = TIMS.GetListValue(OCID) 'OCID.SelectedValue 
            If v_OCID <> "" Then
                Dim ssRID As String = sm.UserInfo.RID
                If RIDValue.Value <> "" Then ssRID = RIDValue.Value
                dtCourse = TIMS.Get_COURSEINFOdt(ssRID, objconn)

                Dim sql As String = ""
                sql = " SELECT * FROM SYS_HOLIDAY WHERE 1=1 "
                sql &= If(RIDValue.Value <> "", " AND RID = '" & RIDValue.Value & "' ", " AND RID = '" & sm.UserInfo.RID & "' ")

                HolidayTable = New DataTable
                HolidayTable.Load(DbAccess.GetReader(sql, objconn))
                'HolidayTable = DbAccess.GetDataTable(sql, objConn)

                '刪除排課確認(Type: 1:批次；2:單月)
                sql = " SELECT 'x' FROM CLASS_SCHEDULE WHERE OCID = '" & v_OCID & "' AND Type = '2' "
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)

                Button12.Enabled = False '不可刪除排課
                TIMS.Tooltip(Button12, "無權限使用該功能 或 已審核確認,不可刪除排課資料", True)
                If dr IsNot Nothing Then
                    If Button12.Enabled = False AndAlso blnCanDeleteAuth = True AndAlso ViewState("IsVerify") <> "Y" Then '審核確認
                        Button12.Enabled = True  '可 刪除排課
                        TIMS.Tooltip(Button12, "有權限使用刪除功能，請小心使用。", True)
                    End If
                    If Button12.Enabled = False AndAlso ViewState("IsVerify") <> "Y" Then '審核確認
                        Button12.Enabled = True  '可 刪除排課
                        TIMS.Tooltip(Button12, "已提供刪除權限 且 未審核確認", True)
                    End If
                End If
            End If
        End If

        'Me.TypeRadio.Attributes("onclick") = "ChangeClassMode();"
        LoadIntoClass.Attributes("onclick") = "return chkdata();"
    End Sub

    Sub Utl_ShowX1()
        '68:照顧服務員自訓自用訓練計畫  
        Const cst_textname1 As String = "教師1"
        Const cst_textname2 As String = "教師2"
        'https://jira.turbotech.com.tw/browse/TIMSC-207
        Const cst_textname1b As String = "老師1"
        Const cst_textname2b As String = "老師2"
        Const cst_textname3b As String = "老師3"
        Const cst_textname4b As String = "助教1"

        trlabTechN4.Visible = False '不顯示助教1<tr>
        labTechN4.Visible = False '不顯示助教1
        OLessonTeah4.Visible = False '不顯示助教1
        OLessonTeah4Value.Visible = False '不顯示助教1
        DataGrid2.Columns(cst_DG2_Teacher4).Visible = False '不顯示助教1(4)

        '68:照顧服務員自訓自用訓練計畫  
        If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            labTechN1.Text = cst_textname1 '"教師1"
            labTechN2.Text = cst_textname2 '"教師2"
            labTechN3.Visible = False '不顯示助教2
            OLessonTeah3.Visible = False '不顯示助教2
            OLessonTeah3Value.Visible = False '不顯示助教2

            DataGrid2.Columns(cst_DG2_Teacher1).HeaderText = cst_textname1
            DataGrid2.Columns(cst_DG2_Teacher2).HeaderText = cst_textname2
            DataGrid2.Columns(cst_DG2_Teacher3).Visible = False '不顯示助教2
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-207
        '47:補助辦理照顧服務員職業訓練 / 58:補助辦理托育人員職業訓練
        '單月排課作業，修改為可設定3位老師(老師1、老師2、老師3)與1位助教(助教1)
        If TIMS.Cst_TPlanID47AppPlan8.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            labTechN1.Text = cst_textname1b '"教師1"
            labTechN2.Text = cst_textname2b '"教師2"
            labTechN3.Text = cst_textname3b '"教師1"
            labTechN4.Text = cst_textname4b '"教師2"

            trlabTechN4.Visible = True '顯示助教1<tr>(4)
            labTechN4.Visible = True '顯示助教1
            OLessonTeah4.Visible = True '顯示助教1
            OLessonTeah4Value.Visible = True '顯示助教1

            DataGrid2.Columns(cst_DG2_Teacher1).HeaderText = cst_textname1b
            DataGrid2.Columns(cst_DG2_Teacher2).HeaderText = cst_textname2b
            DataGrid2.Columns(cst_DG2_Teacher3).HeaderText = cst_textname3b
            DataGrid2.Columns(cst_DG2_Teacher4).HeaderText = cst_textname4b
            DataGrid2.Columns(cst_DG2_Teacher3).Visible = True '顯示
            DataGrid2.Columns(cst_DG2_Teacher4).Visible = True '顯示助教1
        End If
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If CSDate.Text <> "" Then
            If Not TIMS.IsDate1(CSDate.Text) Then Errmsg += "排課區間 起始日期格式有誤" & vbCrLf
            If Errmsg = "" Then CSDate.Text = CDate(CSDate.Text).ToString("yyyy/MM/dd")
        Else
            Errmsg += "排課區間 起始日期 為必填" & vbCrLf
        End If

        If CFDate.Text <> "" Then
            If Not TIMS.IsDate1(CFDate.Text) Then Errmsg += "排課區間 結束日期格式有誤" & vbCrLf
            If Errmsg = "" Then CFDate.Text = CDate(CFDate.Text).ToString("yyyy/MM/dd")
        Else
            Errmsg += "排課區間 結束日期 為必填" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call sSearch2()
    End Sub

    Sub sSearch2()
        'Call TIMS.Utl_Sleep(1000)
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)

        '檢查是否為全日制,若是全日制 檢查是否符合規則
        TIMS.IsAllDateCheck(Me, OCIDValue1.Value, "ShowMsg", objconn)  '檢查是否為全日制,若是全日制檢查是否符合規則

        Call GetData1()

        'ViewState("IsVerify") = ""
        '假如審核確認，則不可以修改資料
        If ViewState("IsVerify") = "Y" Then
            'TIMS.Chk_ClassSchVerify(OCID.SelectedValue)
            '已審核確認
            'If SysInfo.Text <> "" Then SysInfo.Text += "，已審核確認" Else SysInfo.Text += "此班級已審核確認"
            'ViewState("IsVerify") = "Y"
            Button12.Enabled = False '刪除排課資料
            TIMS.Tooltip(Button12, "此班級已審核確認", True)
            Button14.Enabled = False '匯入單月排課作業
            LoadIntoClass.Enabled = False '載入排課班級
            File1.Disabled = True '匯入單月排課作業_檔案選擇
        End If
    End Sub

    ''' <summary>
    ''' 取得資料
    ''' </summary>
    Sub GetData1()
        'Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim DGTable As DataTable = Nothing

        Dim v_OCID As String = TIMS.GetListValue(OCID) 'OCID.SelectedValue
        If CSDate.Text <> "" AndAlso CFDate.Text <> "" Then
            Dim StartDate As Date = CSDate.Text
            Dim EndDate As Date = CFDate.Text

            Dim sql As String = ""
            sql &= " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & v_OCID & "' "
            sql &= " AND SchoolDate >= " & TIMS.To_date(StartDate)
            sql &= " AND SchoolDate <= " & TIMS.To_date(EndDate)
            sql &= " ORDER BY SchoolDate "
            dt = DbAccess.GetDataTable(sql, objconn)

            While (StartDate <= EndDate)
                If dt.Select("SchoolDate='" & StartDate & "'").Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("OCID") = v_OCID 'OCID.SelectedValue
                    dr("SchoolDate") = StartDate
                End If
                StartDate = StartDate.AddDays(1)
            End While

            If dt.Rows.Count = 0 Then
                msg.Text = "查無資料!"
                CourseTable.Style.Item("display") = "none"
                Exit Sub
            End If
            If dt.Select("Type=1").Length <> 0 Then
                msg.Text = cst_alertMsg1 '"您已經使用過全期排課，所以無法使用本功能!"
                CourseTable.Style.Item("display") = "none"
                Exit Sub
            End If

            CourseTable.Style.Item("display") = TIMS.cst_inline1 '"inline"

            PageControler1.PageDataTable = dt
            If Not ViewState("PageIndex") Is Nothing Then PageControler1.PageIndex = ViewState("PageIndex")
            PageControler1.Sort = "SchoolDate"
            PageControler1.ControlerLoad()

            Select Case ShowClassNum.SelectedIndex
                Case 0
                    For i As Integer = 1 To 12
                        DataGrid1.Columns(i).Visible = True
                    Next
                Case 1
                    For i As Integer = 1 To 8
                        DataGrid1.Columns(i).Visible = True
                    Next
                    For i As Integer = 9 To 12
                        DataGrid1.Columns(i).Visible = False
                    Next
                Case 2
                    For i As Integer = 1 To 8
                        DataGrid1.Columns(i).Visible = False
                    Next
                    For i As Integer = 9 To 12
                        DataGrid1.Columns(i).Visible = True
                    Next
            End Select

            '將目前所有的使用課程列出 Start
            Call sHide_TableShow1()
            '將目前所有的使用課程列出 End
        End If
    End Sub

    ''' <summary>
    ''' 將目前所有的使用課程列出 'Public Shared Sub sHide_TableShow1()
    ''' </summary>
    Sub sHide_TableShow1()
        'Optional ByVal sType As Integer = 1
        'sType 1@List 班級課表 2@List 某天的班級課表
        Dim dt As DataTable = Nothing
        'Dim sql As String = ""
        Dim CourseList As String = ""
        Dim CourseArray As Array = Nothing
        'Const Cst_d1 As String = "convert(varchar(7),Schooldate,111)"
        'Const Cst_ym1 As String = " ym1"
        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If v_OCID = "" Then Exit Sub

        Call TIMS.OpenDbConn(objconn)
        'sql = " SELECT *,convert(varchar(7),Schooldate,111) ym1 FROM CLASS_SCHEDULE   WHERE OCID='" & OCID.SelectedValue & "' "
        'sql = " SELECT *,replace(convert(varchar(7),Schooldate,111),'/','_') ym3 FROM CLASS_SCHEDULE   WHERE OCID='" & OCID.SelectedValue & "' " '★
        Dim sql As String = " SELECT a.*, FORMAT(a.Schooldate, 'yyyy_MM') ym3 FROM CLASS_SCHEDULE a WHERE OCID = '" & v_OCID & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Exit Sub

        For Each dr As DataRow In dt.Rows
            For i As Integer = 1 To 12
                If Not IsDBNull(dr("Class" & i)) Then
                    Dim Flag As Boolean = False
                    CourseArray = Split(CourseList, ",")
                    For j As Integer = 0 To CourseArray.Length - 1
                        If CourseArray(j) = dr("Class" & i) Then Flag = True
                    Next
                    If Flag = False Then '表示不存在,增加課程代碼
                        If CourseList <> "" Then CourseList &= ","
                        CourseList &= dr("Class" & i)
                    End If
                End If
            Next
        Next

        DataGrid3.Visible = False
        'DataGrid4.Visible = False '檢視每月已排課時數
        If CourseList <> "" Then '有科目範圍
            Dim Sql_1 As String = ""
            Sql_1 &= " SELECT a.CourID, a.CourseName"
            Sql_1 &= " ,b.CourseName MCourseName"
            Sql_1 &= " ,0 TotalHours " & vbCrLf
            Sql_1 &= " FROM Course_CourseInfo a " & vbCrLf
            Sql_1 &= " LEFT JOIN Course_CourseInfo b ON a.MainCourID = b.CourID " & vbCrLf
            Sql_1 &= " WHERE 1=1 AND a.CourID IN (" & CourseList & ") " & vbCrLf
            Sql_1 &= " ORDER BY a.CourseName " & vbCrLf
            Dim dt1 As DataTable = DbAccess.GetDataTable(Sql_1, objconn) '預備塞入課表 (DataGrid3)

#Region "(No Use)"

            'sql = "" & vbCrLf
            ''ym1 設計比對 與dt table之中要建立喔
            'sql += " select 0 as TotalHours ,convert(varchar(7),Schooldate,111) ym1" & vbCrLf
            'sql += " FROM CLASS_SCHEDULE   " & vbCrLf
            'sql += " where ocid ='" & OCID.SelectedValue & "'" & vbCrLf
            'sql += " group by convert(varchar(7),Schooldate,111) " & vbCrLf
            'sql += " order by convert(varchar(7),Schooldate,111) " & vbCrLf
            ''1.取出上課年月
            'sql = "" & vbCrLf
            'sql += " select * FROM (" & vbCrLf
            ''ym2 設計比對 與dt table之中要建立喔
            'sql += " select distinct replace(convert(varchar(7),Schooldate,111),'/','年')+'月' ym2" & vbCrLf
            'sql += " ,convert(varchar(7),Schooldate,111) ym1 " & vbCrLf
            'sql += " ,replace(convert(varchar(7),Schooldate,111),'/','_') ym3 " & vbCrLf
            'sql += " FROM CLASS_SCHEDULE   " & vbCrLf
            'sql += " where ocid ='" & OCID.SelectedValue & "'" & vbCrLf
            'Sql += " ) g" & vbCrLf
            'sql += " order by ym2" & vbCrLf
            'Dim dt2 As DataTable '預備塞入課表 (DataGrid4)
            'dt2 = DbAccess.GetDataTable(sql, objConn)
            ''2.取出上課科目，並組合年月
            'sql = "SELECT a.CourID,a.CourseName,b.CourseName as MCourseName,0 as TotalHours " & vbCrLf
            'For i As Integer = 0 To dt2.Rows.Count - 1
            '    sql += ",0 as '" & dt2.Rows(i)("ym3") & "'" & vbCrLf
            '    'sql += ",0 as '" & dt2.Rows(i)("ym2") & "'" & vbCrLf
            'Next
            'sql += "FROM (SELECT * FROM Course_CourseInfo   WHERE 1=1 AND CourID IN (" & CourseList & ")) a " & vbCrLf
            'sql += "LEFT JOIN Course_CourseInfo b   ON a.MainCourID=b.CourID " & vbCrLf
            'sql += "Order By a.CourseName" & vbCrLf
            'Dim dt3 As DataTable '預備塞入課表 (DataGrid4)
            'dt3 = DbAccess.GetDataTable(sql, objConn)
            'Const cst_科目 As String = "科目"
            'Const cst_累計時數 As String = "累計時數"
            ''Empty 中文表格 'dt3b 沒有資料
            'sql = "SELECT "
            'sql += " '' as '" & cst_科目 & "'" & vbCrLf
            'For i As Integer = 0 To dt2.Rows.Count - 1
            '    sql += ",0 as '" & dt2.Rows(i)("ym2") & "'" & vbCrLf '中文年月
            'Next
            'sql += ",0 as '" & cst_累計時數 & "' " & vbCrLf
            'sql += "FROM Course_CourseInfo   WHERE 1<>1" & vbCrLf
            'Dim dt3b As DataTable '預備塞入課表 (DataGrid4)
            'dt3b = DbAccess.GetDataTable(sql, objConn)

#End Region

            'dt3 沒有資料
            For Each dr1 As DataRow In dt1.Rows '預備塞入課表
                For Each dr As DataRow In dt.Rows '目前排入課程
                    For i As Integer = 1 To 12 '每天12節
                        If Not IsDBNull(dr("Class" & i)) Then '有課程
                            If dr("Class" & i) = dr1("CourID") Then
                                dr1("TotalHours") += 1 '時數+1
                                'If dt3.Select("CourID='" & dr("Class" & i) & "'").Length > 0 Then
                                '    Dim dr3 As DataRow
                                '    dr3 = dt3.Select("CourID='" & dr("Class" & i) & "'")(0)
                                '    dr3(dr("ym3")) += 1 '某年月 時數+1
                                '    dr3("TotalHours") += 1 '時數+1
                                'End If
                            End If
                        End If
                    Next
                Next
            Next

#Region "(No Use)"

            ''dt3 已有資料
            'For i As Integer = 0 To dt3.Rows.Count - 1
            '    Dim dr3 As DataRow
            '    dr3 = dt3.Rows(i) '取得
            '    Dim dr3b As DataRow
            '    If dr3("TotalHours") > 0 Then '有時數
            '        If dt3b.Select(cst_科目 & "='" & dr3("CourseName") & "'").Length > 0 Then
            '            '取得資料
            '            dr3b = dt3b.Select(cst_科目 & "='" & dr3("CourseName") & "'")(0)
            '        Else
            '            '建立欄位初始值
            '            dr3b = dt3b.NewRow
            '            dt3b.Rows.Add(dr3b)
            '            dr3b(cst_科目) = dr3("CourseName")
            '            For j As Integer = 0 To dt2.Rows.Count - 1
            '                dr3b(dt2.Rows(j)("ym2")) = 0
            '            Next
            '            dr3b(cst_累計時數) = 0
            '        End If
            '        '補資料
            '        For j As Integer = 0 To dt2.Rows.Count - 1
            '            dr3b(dt2.Rows(j)("ym2")) += CInt(dr3(dt2.Rows(j)("ym3")))
            '        Next
            '        dr3b(cst_累計時數) += CInt(dr3("TotalHours"))
            '    End If
            'Next

#End Region

            If dt1.Rows.Count > 0 Then
                'dt1.AcceptChanges()
                DataGrid3.Visible = True
                DataGrid3.DataSource = dt1
                DataGrid3.DataBind()
            End If

            'If dt2.Rows.Count > 0 Then
            '    'dt2.AcceptChanges()
            '    DataGrid4.Visible = True '檢視每月已排課時數
            '    DataGrid4.DataSource = dt3b 'dt3 'dt2
            '    DataGrid4.DataBind()
            'End If
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If Not TIMS.IsDate1(e.CommandArgument) Then
            Common.MessageBox(Me, "日期格式有誤，請重新查詢操作..")
            Exit Sub
        End If

        ViewState("PageIndex") = DataGrid1.CurrentPageIndex + 1
        SearchTable.Style.Item("display") = "none"
        DetailTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        vsErrMsg = ""
        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If Not CheckGetVal1(e.CommandArgument, v_OCID, vsErrMsg) Then
            Common.MessageBox(Me, vsErrMsg)
            Exit Sub
        End If

        CreateDetailCourse(e.CommandArgument, dtTeacher) '檢視詳細課程

        ''將目前所有的使用課程列出 Start
        'Call sHide_TableShow1(2)
        ''將目前所有的使用課程列出 End
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyLink As LinkButton = e.Item.FindControl("LinkButton1")
                Dim vMsg As String = ""

                '無新增的權限
                MyLink.Enabled = True
                Dim flag_NO_USE As Boolean = If(HolidayTable IsNot Nothing AndAlso HolidayTable.Select("HolDate='" & drv("SchoolDate") & "'").Length <> 0, True, False)
                '假日不排課。'MyLink.Enabled = False
                If flag_NO_USE Then vMsg = "假日：" & HolidayTable.Select("HolDate='" & drv("SchoolDate") & "'")(0)("Reason")
                If vMsg <> "" Then TIMS.Tooltip(MyLink, vMsg)

                MyLink.Text = Common.FormatDate(drv("SchoolDate")) & "(" & TIMS.GetWeekDay(CDate(drv("SchoolDate")).DayOfWeek) & ")"
                MyLink.CommandArgument = Common.FormatDate(drv("SchoolDate"))
                MyLink.ForeColor = Color.Blue

                '班級已結訓
                If ViewState("IsClosed") = "Y" Then
                    vMsg = "班級已結訓，不可再修改!!"
                    'MyLink.Attributes("onclick") = "return false;"
                    'MyLink.Attributes("onclick") = "alert('" & vMsg & "');"
                    MyLink.ForeColor = Color.Black
                    TIMS.Tooltip(MyLink, vMsg)
                End If

                '已審核確認
                If ViewState("IsVerify") = "Y" Then
                    vMsg = "已審核確認，不可再修改!!"
                    'MyLink.Attributes("onclick") = "return false;"
                    'MyLink.Attributes("onclick") = "alert('" & vMsg & "');"
                    MyLink.ForeColor = Color.Black
                    TIMS.Tooltip(MyLink, vMsg)
                End If

                Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
                Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"

                If dtCourse Is Nothing Then
                    Dim ssRID As String = sm.UserInfo.RID
                    If RIDValue.Value <> "" Then ssRID = RIDValue.Value
                    dtCourse = TIMS.Get_COURSEINFOdt(ssRID, objconn)
                End If

                For i As Integer = 1 To e.Item.Cells.Count - 2
                    Dim s_COURID As String = e.Item.Cells(i).Text
                    If s_COURID <> "&nbsp;" AndAlso s_COURID <> "" Then
                        ff3 = String.Concat("CourID='", s_COURID, "'")
                        Dim v_COURSENAME As String = ""
                        If dtCourse.Select(ff3).Length <> 0 Then
                            Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
                            v_COURSENAME = If(Convert.ToString(drv("Vacation")) = "Y", String.Format("{0}(假日)", dr1("CourseName")), dr1("CourseName"))
                        ElseIf Convert.ToString(drv("Vacation")) = "Y" Then
                            v_COURSENAME = s_HoliDayCourseName '"假日"
                        Else
                            e.Item.Cells(i).ForeColor = Color.Red
                            TIMS.Tooltip(e.Item.Cells(i), cst_alertMsg16)
                        End If
                        If v_COURSENAME <> "" Then e.Item.Cells(i).Text = v_COURSENAME

                        'ff3 = "CourID='" & e.Item.Cells(i).Text & "'"
                        'Dim v_CourseName As String = ""
                        'If dtCourse.Select(ff3).Length <> 0 Then
                        '    Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
                        '    v_CourseName = If(Convert.ToString(drv("Vacation")) = "Y", String.Format("{0}(假日)", dr1("CourseName")), dr1("CourseName"))
                        '    e.Item.Cells(i).Text = v_CourseName 'dr1("CourseName")
                        'ElseIf Convert.ToString(drv("Vacation")) = "Y" Then
                        '    v_CourseName = s_HoliDayCourseName '"假日"
                        '    e.Item.Cells(i).Text = v_CourseName ' dr("CourseID") = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
                        'Else
                        '    e.Item.Cells(i).ForeColor = Color.Red
                        '    TIMS.Tooltip(e.Item.Cells(i), cst_alertMsg16)
                        'End If
                    End If
                Next

        End Select
    End Sub

    '68:照顧服務員自訓自用訓練計畫  
    Sub Chk_TPlanID68(ByRef Reason As String, ByVal CourID As String, ByVal TechID1 As String, ByVal TechID2 As String, ByVal TechID3 As String, ByVal TechID4 As String)
        If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) = -1 Then Exit Sub '限定計畫檢核
        If dtCourse Is Nothing Then
            Dim ssRID As String = sm.UserInfo.RID
            If RIDValue.Value <> "" Then ssRID = RIDValue.Value
            dtCourse = TIMS.Get_COURSEINFOdt(ssRID, objconn)
        End If

#Region "(No Use)"

        'If dtCourse.Select("CourID='" & CourID & "'").Length = 0 Then
        '    errmsg1 &= cst_alertMsg2
        '    Exit Sub
        'End If

        'Dim dr As DataRow = dtCourse.Select("CourID='" & CourID & "'")(0)
        ''CLASSIFICATION1 1:學科/2術科
        'Select Case Convert.ToString(dr("CLASSIFICATION1"))
        '    Case "1"
        '        If TechID2 <> "" Then
        '            errmsg1 &= cst_alertMsg12
        '            Exit Sub
        '        End If
        'End Select

#End Region

        ff3 = "COURID='" & CourID & "'"
        If dtCourse.Select(ff3).Length = 0 Then
            Reason &= cst_alertMsg2
            Exit Sub
        End If
        Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
        Dim ss3 As String = ""
        TIMS.SetMyValue(ss3, "Classification1", Convert.ToString(dr1("CLASSIFICATION1")))
        TIMS.SetMyValue(ss3, "MaxTNum", hid_TNum.Value)
        TIMS.SetMyValue(ss3, "OLessonTeah1_Value", TechID1)
        TIMS.SetMyValue(ss3, "OLessonTeah2_Value", TechID2)
        TIMS.SetMyValue(ss3, "OLessonTeah3_Value", TechID3)
        TIMS.SetMyValue(ss3, "OLessonTeah4_Value", TechID4)

        'out@Reason
        Call TIMS.Chk_TIMSB1251(Me, Reason, ss3, objconn)
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")
        Dim dt As DataTable = ViewState(cst_vsDetailTable)
        ' IsNothingdtFlag 
        If TIMS.dtNODATA(dt) Then
            '"排課資料不存在，請重新查詢建立排課資料!!"
            Common.MessageBox(Me, cst_alertMsg2)
            Exit Sub
        End If

        Select Case e.CommandName
            Case "edit"
                DataGrid2.EditItemIndex = e.Item.ItemIndex
            Case "del"
                '刪除
                ViewState("LeftHour") = Nothing
                If dt.Select("ClassNum='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select("ClassNum='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'")(0)
                    dr("CourseName") = ""
                    dr("CourseID") = ""
                    dr("ClassRoom") = ""
                    dr("Teacher1") = ""
                    dr("Teacher1ID") = ""
                    dr("Teacher2") = ""
                    dr("Teacher2ID") = ""
                    dr("Teacher3") = ""
                    dr("Teacher3ID") = ""
                    dr("Teacher4") = ""
                    dr("Teacher4ID") = ""
                    dr("VACATION") = ""
                    ViewState("LeftHour") = Me.LeftHour.Text
                End If
                'dt.AcceptChanges()
                ViewState(cst_vsDetailTable) = dt
            Case "cancel"
                DataGrid2.EditItemIndex = -1
            Case "update"
                '修改
                Dim CourseName As TextBox = e.Item.FindControl("CourseName")
                Dim CourseValue As HtmlInputHidden = e.Item.FindControl("CourseValue")
                Dim ClassRoom As TextBox = e.Item.FindControl("ClassRoom")
                Dim Teacher1 As TextBox = e.Item.FindControl("Teacher1")
                Dim Teacher1Value As HtmlInputHidden = e.Item.FindControl("Teacher1Value")
                Dim Teacher2 As TextBox = e.Item.FindControl("Teacher2")
                Dim Teacher2Value As HtmlInputHidden = e.Item.FindControl("Teacher2Value")
                Dim Teacher3 As TextBox = e.Item.FindControl("Teacher3")
                Dim Teacher3Value As HtmlInputHidden = e.Item.FindControl("Teacher3Value")
                Dim Teacher4 As TextBox = e.Item.FindControl("Teacher4")
                Dim Teacher4Value As HtmlInputHidden = e.Item.FindControl("Teacher4Value")

                If dt.Select("ClassNum='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select("ClassNum='" & DataGrid2.DataKeys(e.Item.ItemIndex) & "'")(0)

                    Teacher1.Text = TIMS.ClearSQM(Teacher1.Text)
                    Teacher1Value.Value = TIMS.ClearSQM(Teacher1Value.Value)
                    If Teacher1.Text = "" Then Teacher1Value.Value = ""

                    Teacher2.Text = TIMS.ClearSQM(Teacher2.Text)
                    Teacher2Value.Value = TIMS.ClearSQM(Teacher2Value.Value)
                    If Teacher2.Text = "" Then Teacher2Value.Value = ""

                    Teacher3.Text = TIMS.ClearSQM(Teacher3.Text)
                    Teacher3Value.Value = TIMS.ClearSQM(Teacher3Value.Value)
                    If Teacher3.Text = "" Then Teacher3Value.Value = ""

                    Teacher4.Text = TIMS.ClearSQM(Teacher4.Text)
                    Teacher4Value.Value = TIMS.ClearSQM(Teacher4Value.Value)
                    If Teacher4.Text = "" Then Teacher4Value.Value = ""

                    '68:照顧服務員自訓自用訓練計畫  
                    Dim errmsg1 As String = ""
                    Call Chk_TPlanID68(errmsg1, CourseValue.Value, Teacher1Value.Value, Teacher2Value.Value, Teacher3Value.Value, Teacher4Value.Value)

                    If errmsg1 <> "" Then
                        DataGrid2.EditItemIndex = -1
                        Common.MessageBox(Me, errmsg1)
                        Exit Select '只離開此判斷
                    End If

                    dr("CourseName") = CourseName.Text 'TIMS.Get_CourseName(CourseValue.Value)
                    'CourseName.Text = TIMS.Get_CourseName(CourseValue.Value, dtCourse, objConn)
                    'dr("CourseName") = CourseName.Text 'TIMS.Get_CourseName(CourseValue.Value)
                    dr("CourseID") = CourseValue.Value
                    dr("ClassRoom") = TIMS.Get_Substr1(ClassRoom.Text, 30)
                    dr("Teacher1") = Teacher1.Text 'TIMS.Get_TeacherName(Teacher1Value.Value)
                    dr("Teacher1ID") = Teacher1Value.Value
                    dr("Teacher2") = Teacher2.Text 'TIMS.Get_TeacherName(Teacher2Value.Value)
                    dr("Teacher2ID") = Teacher2Value.Value
                    dr("Teacher3") = Teacher3.Text 'TIMS.Get_TeacherName(Teacher2Value.Value)
                    dr("Teacher3ID") = Teacher3Value.Value
                    dr("Teacher4") = Teacher4.Text 'TIMS.Get_TeacherName(Teacher2Value.Value)
                    dr("Teacher4ID") = Teacher4Value.Value
                    dr("VACATION") = If(v_Vacation <> "", v_Vacation, "") '假日排課
                End If
                'dt.AcceptChanges()
                ViewState(cst_vsDetailTable) = dt
                DataGrid2.EditItemIndex = -1
        End Select
        'Try
        'Catch ex As Exception
        'End Try
        'IFRAME1.Attributes("src") = "SD_04_002_Course.aspx?RID=" & RIDValue.Value
        DataGrid2.DataSource = ViewState(cst_vsDetailTable)
        DataGrid2.DataBind()

        'Call GetUsedClass()
        Call GetLeftCourseHour(Int(THours.Text))
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Dim drv As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim btn1 As Button = e.Item.FindControl("Button5") '修改
                Dim btn2 As Button = e.Item.FindControl("Button6") '刪除
                btn2.Attributes("onclick") = "return confirm('確定要刪除本節課程?\n\n(離開時請記得按儲存鈕)\n\n');"

                If drv("CourseName").ToString = "" Then
                    btn1.Enabled = False
                    btn2.Enabled = False
                    TIMS.Tooltip(btn1, "尚未選擇課程")
                    TIMS.Tooltip(btn2, "尚未選擇課程")
                End If
                If ViewState("IsVerify") = "Y" Then
                    btn1.Enabled = False
                    btn2.Enabled = False
                    TIMS.Tooltip(btn1, "此班級已審核確認")
                    TIMS.Tooltip(btn2, "此班級已審核確認")
                End If

                If btn1.Enabled Then
                    '若日期區間落在開、結訓日期中，為True，其餘為False
                    If Not CheckDate_NoUpdata(CDate(MyDate.Text), CDate(STDate.Value), CDate(FTDate.Value)) Then
                        '修改鈕停用
                        btn1.Enabled = False
                        ' "排課日期區間未落在開、結訓日期中 無法修改排課!!"
                        TIMS.Tooltip(btn1, cst_alertMsg3)
                    End If
                End If

                e.Item.Cells(cst_DG2_CourseName).Text = Convert.ToString(drv("CourseName"))
                e.Item.Cells(cst_DG2_ClassRoom).Text = Convert.ToString(drv("ClassRoom"))
                e.Item.Cells(cst_DG2_Teacher1).Text = Convert.ToString(drv("Teacher1"))
                e.Item.Cells(cst_DG2_Teacher2).Text = Convert.ToString(drv("Teacher2"))
                e.Item.Cells(cst_DG2_Teacher3).Text = Convert.ToString(drv("Teacher3"))
                e.Item.Cells(cst_DG2_Teacher4).Text = Convert.ToString(drv("Teacher4"))

            Case ListItemType.EditItem
                Dim btn As Button = e.Item.FindControl("Button7")
                Dim CourseName As TextBox = e.Item.FindControl("CourseName")
                Dim CourseValue As HtmlInputHidden = e.Item.FindControl("CourseValue")
                Dim ClassRoom As TextBox = e.Item.FindControl("ClassRoom")
                Dim Teacher1 As TextBox = e.Item.FindControl("Teacher1")
                Dim Teacher1Value As HtmlInputHidden = e.Item.FindControl("Teacher1Value")
                Dim Teacher2 As TextBox = e.Item.FindControl("Teacher2")
                Dim Teacher2Value As HtmlInputHidden = e.Item.FindControl("Teacher2Value")
                Dim Teacher3 As TextBox = e.Item.FindControl("Teacher3")
                Dim Teacher3Value As HtmlInputHidden = e.Item.FindControl("Teacher3Value")
                Dim Teacher4 As TextBox = e.Item.FindControl("Teacher4")
                Dim Teacher4Value As HtmlInputHidden = e.Item.FindControl("Teacher4Value")
                btn.Attributes("onclick") = "var msg='';if(document.form1." & CourseName.ClientID & ".value=='') msg+='請輸入課程代碼\n';if(document.form1." & ClassRoom.ClientID & ".value=='') msg+='請輸入教室\n';if(document.form1." & Teacher1.ClientID & ".value=='') msg+='請輸入教師1\n';if(msg!=''){alert(msg);return false;}"

                CourseName.Text = drv("CourseName").ToString
                CourseValue.Value = drv("CourseID").ToString
                ClassRoom.Text = TIMS.Get_Substr1(drv("ClassRoom").ToString, 30)
                Teacher1.Text = drv("Teacher1").ToString
                Teacher1Value.Value = drv("Teacher1ID").ToString
                Teacher2.Text = drv("Teacher2").ToString
                Teacher2Value.Value = drv("Teacher2ID").ToString
                Teacher3.Text = drv("Teacher3").ToString
                Teacher3Value.Value = drv("Teacher3ID").ToString
                Teacher4.Text = drv("Teacher4").ToString
                Teacher4Value.Value = drv("Teacher4ID").ToString

                Teacher1.Text = TIMS.ClearSQM(Teacher1.Text)
                Teacher1Value.Value = TIMS.ClearSQM(Teacher1Value.Value)
                If Teacher1.Text = "" Then Teacher1Value.Value = ""

                Teacher2.Text = TIMS.ClearSQM(Teacher2.Text)
                Teacher2Value.Value = TIMS.ClearSQM(Teacher2Value.Value)
                If Teacher2.Text = "" Then Teacher2Value.Value = ""

                Teacher3.Text = TIMS.ClearSQM(Teacher3.Text)
                Teacher3Value.Value = TIMS.ClearSQM(Teacher3Value.Value)
                If Teacher3.Text = "" Then Teacher3Value.Value = ""

                Teacher4.Text = TIMS.ClearSQM(Teacher4.Text)
                Teacher4Value.Value = TIMS.ClearSQM(Teacher4Value.Value)
                If Teacher4.Text = "" Then Teacher4Value.Value = ""

                'Dim srcstr As String = ""
                'srcstr += "&TextField=" & CourseName.ClientID
                'srcstr += "&HiddenField=" & CourseValue.ClientID
                'srcstr += "&Tech1Field=" & Teacher1Value.ClientID
                'srcstr += "&TechName1Field=" & Teacher1.ClientID
                'srcstr += "&Tech2Field=" & Teacher2Value.ClientID
                'srcstr += "&TechName2Field=" & Teacher2.ClientID
                'srcstr += "&Tech3Field=" & Teacher3Value.ClientID
                'srcstr += "&TechName3Field=" & Teacher3.ClientID
                'srcstr += "&RoomField=" & ClassRoom.ClientID

                'SD_04_002_Course.aspx
                'CourseName.Attributes("onclick") = "Course_search('Edit','" & CourseValue.ClientID & "','" & CourseName.ClientID & "');"
                CourseName.Attributes("onclick") = "Course_search('Edit','" & CourseValue.ClientID & "','" & CourseName.ClientID & "');"
                CourseName.Attributes.Add("onkeydown", "this.blur()")
                CourseName.Attributes.Add("oncontextmenu", "return false;")
                CourseName.Style.Item("CURSOR") = "hand"
                'IFRAME1.Attributes("src") = "SD_04_002_Course.aspx?RID=" & RIDValue.Value & srcstr
                'IFRAME1.Attributes("src") = "SD_04_002_Course.aspx?RID=" & RIDValue.Value & "&TextField=" & CourseName.ClientID & "&HiddenField=" & CourseValue.ClientID & "&Tech1Field=" & Teacher1Value.ClientID & "&TechName1Field=" & Teacher1.ClientID & "&Tech2Field=" & Teacher1Value.ClientID & "&TechName2Field=" & Teacher2.ClientID & "&RoomField=" & ClassRoom.ClientID
                'CourseName.Attributes("onclick") = "Course('Add','" & CourseName.ClientID & "','" & CourseValue.ClientID & " ');"
                'CourseName.Style.Item("CURSOR") = "hand"
                'CourseName.ReadOnly = True
                'Teacher1.Attributes("onclick") = "Get_Teah('" & Teacher1.ClientID & "','" & Teacher1Value.ClientID & "');"
                'Teacher1.ReadOnly = True
                'Teacher2.Attributes("onclick") = "Get_Teah('" & Teacher2.ClientID & "','" & Teacher2Value.ClientID & "');"
                'Teacher2.ReadOnly = True

                Teacher1.Style.Item("CURSOR") = "hand"
                Teacher1.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add2','1','" & Teacher1.ClientID & "','" & Teacher1Value.ClientID & "');")
                Teacher1.Attributes.Add("onkeydown", "this.blur()")
                Teacher1.Attributes.Add("oncontextmenu", "return false;")

                Teacher2.Style.Item("CURSOR") = "hand"
                Teacher2.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','2','" & Teacher2.ClientID & "','" & Teacher2Value.ClientID & "');")
                Teacher2.Attributes.Add("onkeydown", "this.blur()")
                Teacher2.Attributes.Add("oncontextmenu", "return false;")

                Teacher3.Style.Item("CURSOR") = "hand"
                Teacher3.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','3','" & Teacher3.ClientID & "','" & Teacher3Value.ClientID & "');")
                Teacher3.Attributes.Add("onkeydown", "this.blur()")
                Teacher3.Attributes.Add("oncontextmenu", "return false;")

                Teacher4.Style.Item("CURSOR") = "hand"
                Teacher4.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','4','" & Teacher4.ClientID & "','" & Teacher4Value.ClientID & "');")
                Teacher4.Attributes.Add("onkeydown", "this.blur()")
                Teacher4.Attributes.Add("oncontextmenu", "return false;")
        End Select
    End Sub

    '計算可用的時數(排課時數) 將可用值存入 LeftHour.Text(LeftHour1.Value) 
    '，若無法使用 則  Button9.Enabled 為 False
    Sub GetLeftCourseHour(ByVal Total As Integer)
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim holiday_CourID As String = ""
        holiday_CourID = TIMS.Get_CourID(cst_holiday, objconn)

        Dim dtDetailTable As DataTable = Nothing
        Try
            dtDetailTable = ViewState(cst_vsDetailTable)
        Catch ex As Exception
            '"排課時數異常，查無排課資料!!"
            Common.MessageBox(Me, cst_alertMsg4)
            Exit Sub
        End Try
        '刪除資料不存在
        If TIMS.dtNODATA(dtDetailTable) Then
            '排課資料不存在，請重新查詢建立排課資料!!
            Common.MessageBox(Me, cst_alertMsg2)
            Exit Sub
        End If

        '加上判斷排除不計算排課時數
        Dim isCountHours As String = ""
        Dim sql99 As String = " SELECT isCountHours FROM Course_CourseInfo WHERE CourID = @CourID "  '& dr("Class" & i).ToString
        'objConn = DbAccess.GetConnection()
        Dim cmd As New SqlCommand(sql99, objconn)
        Call TIMS.OpenDbConn(objconn)
        'If objConn.State = ConnectionState.Closed Then objConn.Open()

        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        sql = " SELECT * FROM CLASS_SCHEDULE WHERE SchoolDate <> " & TIMS.To_date(MyDate.Text) & " AND OCID = '" & v_OCID & "' AND Type = '2' "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            For Each dr In dt.Rows
                For i As Integer = 1 To 12
                    If Convert.ToString(dr("Class" & i)) <> "" AndAlso Convert.ToString(dr("Class" & i)) <> holiday_CourID Then
                        '在資料庫裡新增一筆假日的資料供排假日用
                        'insert INTO Course_CourseInfo(CourseID,CoursENAME,Classification1,Classification2,RID,ModifyAcct,ModifyDATE)
                        'VALUES('holiday','假日',1,0,'','sys',getdate())
                        cmd.Parameters.Clear()
                        cmd.Parameters.Add("CourID", SqlDbType.VarChar).Value = CStr(dr("Class" & i))
                        isCountHours = Convert.ToString(cmd.ExecuteScalar())
                        If isCountHours = "" Then Total -= 1
                    End If
                Next
            Next
        End If

        TodayUseHour.Value = "0"
        'dt = ViewState(cst_vsDetailTable)
        If TIMS.dtHaveDATA(dtDetailTable) Then
            For Each dr In dtDetailTable.Rows
                'If dr("CourseName").ToString <> "" Then
                If Convert.ToString(dr("CourseName")) <> "" AndAlso Convert.ToString(dr("CourseName")) <> "假日" Then
                    cmd.Parameters.Clear()
                    cmd.Parameters.Add("CourID", SqlDbType.VarChar).Value = dr("CourseID").ToString
                    isCountHours = Convert.ToString(cmd.ExecuteScalar())
                    If isCountHours = "" Then
                        Total -= 1
                        TodayUseHour.Value = Int(TodayUseHour.Value) + 1
                    End If
                End If
            Next
        End If

        LeftHour.Text = Total
        Button9.Enabled = True
        Button3.Enabled = True
        If Total = 0 Then
            Button9.Enabled = False
            '"排課時數已經用完!"
            TIMS.Tooltip(Button9, cst_alertMsg5, True)
            Button3.Enabled = True
            Common.MessageBox(Me, cst_alertMsg5)
        ElseIf Total < 0 Then
            Button9.Enabled = False
            Button3.Enabled = False
            TIMS.Tooltip(Button9, cst_alertMsg5, True)
            TIMS.Tooltip(Button3, cst_alertMsg5, True)
            If Not ViewState("LeftHour") Is Nothing Then
                ViewState("LeftHour") = Nothing
                '採取刪除動作打開儲存鈕
                Button3.Enabled = True
            End If
            Common.MessageBox(Me, cst_alertMsg5)
        End If
    End Sub

    '檢核資料
    Function CheckGetVal1(ByVal MyDateStr As Object, ByVal OCID As String, ByRef errMsg As String) As Boolean
        Dim Rst As Boolean = True
        errMsg = ""

        '檢核資料
        If OCID = "" OrElse Not IsNumeric(OCID) Then
            errMsg += "未選擇 職類/班別或有誤，請再確認!!" & vbCrLf
            'Common.MessageBox(Me, "未選擇 職類/班別，請再確認!!")
            'Exit Function
        End If

        Try
            If Convert.ToString(MyDateStr).Trim <> "" Then
                MyDateStr = Convert.ToString(MyDateStr).Trim
                MyDateStr = CDate(MyDateStr).ToString("yyyy/MM/dd")
            Else
                errMsg += "排課日期未輸入，請再確認!!" & vbCrLf
            End If
        Catch ex As Exception
            errMsg += "排課日期異常，請再確認!!" & vbCrLf
            'Common.MessageBox(Me, "該班日期異常，請再確認!!")
            'Exit Function
        End Try

        If errMsg <> "" Then Rst = False
        Return Rst
    End Function

    '建立課表資料 
    Sub CreateDetailCourse(ByVal MyDateStr As String, ByRef dtTeacher As DataTable)
        If Convert.ToString(MyDateStr).Trim <> "" Then
            MyDateStr = Convert.ToString(MyDateStr).Trim
            MyDateStr = CDate(MyDateStr).ToString("yyyy/MM/dd")
        End If

        '假如審核確認，則不可以修改資料
        ViewState("IsVerify") = ""
        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If TIMS.Chk_ClassSchVerify(v_OCID, objconn) Then ViewState("IsVerify") = "Y"

        MyDate.Text = MyDateStr
        MyWeek.Text = "(" & TIMS.GetWeekDay(CDate(MyDateStr).DayOfWeek) & ")"
        Dim sql As String = ""
        Dim dr As DataRow = Nothing

        '建立空白資料表 Start
        Dim drTemp As DataRow = Nothing
        Dim dtTemp As New DataTable
        dtTemp.Columns.Add(New DataColumn("ClassNum"))
        dtTemp.Columns.Add(New DataColumn("CourseName"))
        dtTemp.Columns.Add(New DataColumn("CourseID"))
        dtTemp.Columns.Add(New DataColumn("Teacher1"))
        dtTemp.Columns.Add(New DataColumn("Teacher1ID"))
        dtTemp.Columns.Add(New DataColumn("Teacher2"))
        dtTemp.Columns.Add(New DataColumn("Teacher2ID"))
        dtTemp.Columns.Add(New DataColumn("Teacher3"))
        dtTemp.Columns.Add(New DataColumn("Teacher3ID"))
        dtTemp.Columns.Add(New DataColumn("Teacher4"))
        dtTemp.Columns.Add(New DataColumn("Teacher4ID"))
        dtTemp.Columns.Add(New DataColumn("ClassRoom"))
        dtTemp.Columns.Add(New DataColumn("VACATION"))
        '建立空白資料表 End

        'sql = "SELECT TechID,TeachCName FROM Teach_TeacherInfo WHERE WorkStatus='1' "
        'dtTeacher = DbAccess.GetDataTable(sql)
        Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
        Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"

        Dim VACATION As String = ""
        Dim iUsedHour As Integer = 0
        'Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        sql = "SELECT * FROM CLASS_SCHEDULE WHERE CONVERT(varchar, SchoolDate, 111) = '" & MyDateStr & "' AND OCID = " & v_OCID 'OCID.SelectedValue
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            CSID.Value = ""
            For i As Integer = 1 To 12
                drTemp = dtTemp.NewRow
                dtTemp.Rows.Add(drTemp)
                drTemp("ClassNum") = i
                drTemp("CourseName") = ""
                drTemp("CourseID") = ""
                drTemp("Teacher1") = ""
                drTemp("Teacher1ID") = ""
                drTemp("Teacher2") = ""
                drTemp("Teacher2ID") = ""
                drTemp("Teacher3") = ""
                drTemp("Teacher3ID") = ""
                drTemp("Teacher4") = ""
                drTemp("Teacher4ID") = ""
                drTemp("ClassRoom") = ""
                drTemp("VACATION") = VACATION
            Next
        Else
            CSID.Value = dr("CSID")
            VACATION = Convert.ToString(dr("VACATION"))
            For i As Integer = 1 To 12
                drTemp = dtTemp.NewRow
                dtTemp.Rows.Add(drTemp)
                drTemp("ClassNum") = i
                drTemp("CourseName") = ""
                drTemp("CourseID") = ""
                drTemp("Teacher1") = ""
                drTemp("Teacher1ID") = ""
                drTemp("Teacher2") = ""
                drTemp("Teacher2ID") = ""
                drTemp("Teacher3") = ""
                drTemp("Teacher3ID") = ""
                drTemp("Teacher4") = ""
                drTemp("Teacher4ID") = ""
                drTemp("ClassRoom") = ""
                drTemp("VACATION") = VACATION
                If dr("Class" & i).ToString <> "" Then
                    Dim sTmp As String = ""
                    sTmp = TIMS.ClearSQM(Convert.ToString(dr("Class" & i)))
                    drTemp("CourseID") = If(sTmp = "", Convert.DBNull, sTmp)
                    If sTmp <> "" Then
                        ff3 = "CourID='" & sTmp & "'"
                        Dim v_COURSENAME As String = ""
                        If dtCourse.Select(ff3).Length <> 0 Then
                            Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
                            v_COURSENAME = If(VACATION = "Y", String.Format("{0}(假日)", dr1("CourseName")), dr1("CourseName"))
                        ElseIf VACATION = "Y" Then
                            v_COURSENAME = s_HoliDayCourseName '"假日"
                        End If
                        If v_COURSENAME <> "" Then drTemp("CourseName") = v_COURSENAME
                        'If dtCourse.Select(ff3).Length <> 0 Then drTemp("CourseName") = dtCourse.Select(ff3)(0)("CourseName")
                    End If
                    '1~12
                    sTmp = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i)))
                    If sTmp <> "" Then
                        drTemp("Teacher1ID") = sTmp
                        If dtTeacher.Select("TechID='" & sTmp & "'").Length <> 0 Then drTemp("Teacher1") = dtTeacher.Select("TechID='" & sTmp & "'")(0)("TeachCName")
                    End If
                    '13~24
                    sTmp = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 12)))
                    If sTmp <> "" Then
                        drTemp("Teacher2ID") = sTmp
                        If dtTeacher.Select("TechID='" & sTmp & "'").Length <> 0 Then drTemp("Teacher2") = dtTeacher.Select("TechID='" & sTmp & "'")(0)("TeachCName")
                    End If
                    '25~36
                    sTmp = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 24)))
                    If sTmp <> "" Then
                        drTemp("Teacher3ID") = sTmp
                        If dtTeacher.Select("TechID='" & sTmp & "'").Length <> 0 Then drTemp("Teacher3") = dtTeacher.Select("TechID='" & sTmp & "'")(0)("TeachCName")
                    End If
                    '37~48
                    sTmp = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 36)))
                    If sTmp <> "" Then
                        drTemp("Teacher4ID") = sTmp
                        If dtTeacher.Select("TechID='" & sTmp & "'").Length <> 0 Then drTemp("Teacher4") = dtTeacher.Select("TechID='" & sTmp & "'")(0)("TeachCName")
                    End If
                    sTmp = TIMS.Get_Substr1(TIMS.ClearSQM(Convert.ToString(dr("Room" & i))), 30)
                    drTemp("ClassRoom") = If(sTmp = "", Convert.DBNull, sTmp)
                    iUsedHour += 1
                End If
            Next
        End If
        TodayUseHour.Value = iUsedHour

        'dtTemp.AcceptChanges()
        ViewState(cst_vsDetailTable) = dtTemp '將ViewState(cst_vsDetailTable) 取得到最新的資料，若無資料塞入空白

        DataGrid2.DataSource = dtTemp
        DataGrid2.DataKeyField = "ClassNum"
        DataGrid2.DataBind()

        'Call GetUsedClass()  '顯示能夠編輯的課程節次(日班只能編輯1-8節...)
        Call GetLeftCourseHour(Int(THours.Text))  '計算可用的時數(排課時數) 將可用值存入 LeftHour.Text ，若無法使用 則  Button9.Enabled 為 False

        'Button12.Enabled = True '刪除
        'Button10.Enabled = True '刪除
        'Button9.Enabled = True '新增
        'Button3.Enabled = True '儲存

        Dim vMsg As String = ""
        vMsg = ""
        If ViewState("IsClosed") = "Y" Then vMsg += "此班級已結訓，不可再修改!!"  '已結訓
        If ViewState("IsVerify") = "Y" Then vMsg += "此班級已審核確認，不可再修改!!" '已審核確認
        If vMsg <> "" Then
            Button12.Enabled = False '刪除
            Button10.Enabled = False '刪除
            Button9.Enabled = False '新增
            Button3.Enabled = False '儲存
            TIMS.Tooltip(Button12, vMsg, True)
            TIMS.Tooltip(Button10, vMsg, True)
            TIMS.Tooltip(Button9, vMsg, True)
            TIMS.Tooltip(Button3, vMsg, True)
        Else
            Button12.Enabled = True '刪除
            Button10.Enabled = True '刪除
            Button9.Enabled = True '新增
            Button3.Enabled = True '儲存
            TIMS.Tooltip(Button12, vMsg)
            TIMS.Tooltip(Button10, vMsg)
            TIMS.Tooltip(Button9, vMsg)
            TIMS.Tooltip(Button3, vMsg)
        End If
    End Sub

#Region "(No Use)"

    '顯示能夠編輯的課程節次(日班只能編輯1-8節...)
    'Sub GetUsedClass()
    '    Dim oDG As DataGrid = DataGrid2
    '    For i As Integer = 0 To oDG.Items.Count - 1
    '        oDG.Items(i).Visible = True
    '    Next
    '    Select Case TPeriodValue.Value
    '        Case "01"
    '            For i As Integer = 8 To 11
    '                oDG.Items(i).Visible = False
    '            Next
    '        Case "02"
    '            For i As Integer = 0 To 7
    '                oDG.Items(i).Visible = False
    '            Next
    '        Case "03"
    '        Case "04"
    '    End Select
    'End Sub

#End Region

    '是否不開啟修改鈕: 若日期區間落在開、結訓日期中，為True，其餘為False
    Function CheckDate_NoUpdata(ByVal MyDate As Date, ByVal STDate As Date, ByVal FTDate As Date) As Boolean
        Dim rst As Boolean = False
        If DateDiff(DateInterval.Day, STDate, MyDate) >= 0 And DateDiff(DateInterval.Day, MyDate, FTDate) >= 0 Then rst = True
        Return rst
    End Function

    '儲存課表
    Function SAVE_CLASS_SCHEDULE(ByVal CSID As String, ByVal OCID As String, ByVal SchoolDate As DateTime) As String
        'ByRef csClass() As String, ByRef Room() As String, _
        'ByRef Teacher1() As String, ByRef Teacher2() As String)
        'Save_Class_Schedule = "" '返回的錯誤訊息
        Dim sErrmsg As String = "" '返回的錯誤訊息

        CSID = TIMS.ClearSQM(CSID)

        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")
        Dim Mydt As DataTable = ViewState(cst_vsDetailTable)

        '刪除資料不存在
        If TIMS.dtNODATA(Mydt) Then
            'Common.MessageBox(Me, "排課資料不存在，請重新查詢建立排課資料!!") 'Exit Function
            sErrmsg &= "排課資料不存在，請重新查詢建立排課資料!!"
            Return sErrmsg 'Exit Function
        End If

        OCID = TIMS.ClearSQM(OCID)
        If OCID = "" Then
            'Common.MessageBox(Me, "班級資訊異常，請重新查詢班級資料!!")
            'Exit Function
            sErrmsg &= cst_alertMsg6 '"班級資訊異常，請重新查詢班級資料!!"
            Return sErrmsg 'Exit Function
        End If

        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        'Dim sql As String = ""

        Dim StartDate As Date
        Dim EndDate As Date
        Dim i As Integer = 0
        Dim j As Integer = 0

        If Mydt Is Nothing Then
            sErrmsg &= cst_alertMsg2 '"排課資料不存在，請重新查詢建立排課資料!!"
            Return sErrmsg 'Exit Function
        End If

        Using tConn As SqlConnection = DbAccess.GetConnection()
            Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
            Try
                '先將課程本日課程資料填入Class_Schedule Start
                'trans = DbAccess.BeginTrans(tConn)
                Dim sql As String = ""
                If CSID = "" Then
                    'Call TIMS.OpenDbConn(tConn)
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE SchoolDate = " & TIMS.To_date(SchoolDate.ToString("yyyy/MM/dd")) & " AND OCID = '" & OCID.ToString & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count = 0 Then
                        Dim iCSID As Integer = DbAccess.GetNewId(trans, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("CSID") = iCSID
                        dr("OCID") = OCID.ToString
                        dr("SchoolDate") = Common.FormatDate(SchoolDate)
                    Else
                        dr = dt.Rows(0)
                    End If
                Else
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE CSID = '" & CSID & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    dr = dt.Rows(0)
                End If

                dr("Formal") = "Y"
                dr("Type") = 2 '1:全期排課/2:單月排課
                Dim Tmps As String = ""
                For i = 1 To 12
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("CourseID")))
                    dr("Class" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                    '1~12
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher1ID")))
                    dr("Teacher" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                    '13~24
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher2ID")))
                    dr("Teacher" & i + 12) = If(Tmps = "", Convert.DBNull, Tmps)
                    '25~36
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher3ID")))
                    dr("Teacher" & i + 24) = If(Tmps = "", Convert.DBNull, Tmps)
                    '37~48
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher4ID")))
                    dr("Teacher" & i + 36) = If(Tmps = "", Convert.DBNull, Tmps)

                    Tmps = TIMS.Get_Substr1(TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("ClassRoom"))), 30)
                    dr("Room" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                Next
                dr("VACATION") = If(v_Vacation <> "", v_Vacation, Convert.DBNull) '假日排課
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                DbAccess.UpdateDataTable(dt, da, trans)
                '先將課程本日課程資料填入Class_Schedule End

                'Plan_Schedule
                '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
                Call AddNew_Plan_Schedule(Me, STDate.Value, FTDate.Value, OCID.ToString, dt, da, trans, dtCourse)
                DbAccess.CommitTrans(trans)
            Catch ex As Exception
                Common.MessageBox(Me, ex.Message)
                Dim strErrmsg As String = String.Concat(" *ex.Message:", vbCrLf, ex.Message, vbCrLf, TIMS.GetErrorMsg(Me)) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg, ex)

                DbAccess.RollbackTrans(trans)
                Call TIMS.CloseDbConn(tConn)
                'DbAccess.CloseDbConn(objconn)
                Throw ex
            End Try
            Call TIMS.CloseDbConn(tConn)
        End Using

        '增加空白日期存入資料庫--   Start
        Dim flag As Boolean = False             '要是為True，必須回存日期
        StartDate = STDate.Value
        EndDate = FTDate.Value
        Using tConn As SqlConnection = DbAccess.GetConnection()
            '儲存課表
            Call TIMS.OpenDbConn(tConn)
            'sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & OCID.ToString & "' and SchoolDate>=convert(datetime, '" & StartDate & "', 111) and SchoolDate<=convert(datetime, '" & EndDate & "', 111) and Formal='Y'"
            Dim sql As String = ""
            sql &= " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & OCID & "' "
            sql &= " AND SchoolDate >= " & TIMS.To_date(StartDate)
            sql &= " AND SchoolDate <= " & TIMS.To_date(EndDate)
            sql &= " AND Formal = 'Y' "
            dt = DbAccess.GetDataTable(sql, da, tConn)
            While (StartDate <= EndDate)
                If dt.Select("SchoolDate='" & StartDate & "'").Length = 0 Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("CSID") = DbAccess.GetNewId(tConn, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                    dr("OCID") = OCID.ToString
                    dr("SchoolDate") = StartDate
                    dr("Formal") = "Y"
                    dr("Type") = 2 '1:全期排課/2:單月排課
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    flag = True
                End If
                StartDate = StartDate.AddDays(1)
            End While
            If flag Then
                DbAccess.UpdateDataTable(dt, da)
            End If
            '增加空白日期存入資料庫--   End
            Call TIMS.CloseDbConn(tConn)
        End Using

        Return sErrmsg
    End Function

    '儲存課表3
    Function Save_Class_Schedule3(ByVal CSID As String, ByVal OCID As String, ByVal SchoolDate As DateTime) As String
        'ByRef csClass() As String, ByRef Room() As String, _
        'ByRef Teacher1() As String, ByRef Teacher2() As String)
        Dim sErrmsg As String = "" '返回的錯誤訊息'Save_Class_Schedule3 

        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")
        Dim Mydt As DataTable = ViewState(cst_vsDetailTable)
        '刪除資料不存在
        If TIMS.dtNODATA(Mydt) Then
            'Common.MessageBox(Me, "排課資料不存在，請重新查詢建立排課資料!!") 'Exit Function
            sErrmsg &= "排課資料不存在，請重新查詢建立排課資料!!"
            Return sErrmsg 'Exit Function
        End If
        If Mydt Is Nothing Then
            sErrmsg &= cst_alertMsg2 '"排課資料不存在，請重新查詢排課資料!!"
            Return sErrmsg 'Exit Function
        End If

        Using tConn As SqlConnection = DbAccess.GetConnection()
            Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
            Try
                '先將課程本日課程資料填入Class_Schedule Start
                'trans = DbAccess.BeginTrans(tConn)
                Dim da As SqlDataAdapter = Nothing
                Dim dt As DataTable = Nothing
                Dim dr As DataRow = Nothing
                Dim sql As String = ""
                If CSID.ToString = "" Then
                    'Call TIMS.OpenDbConn(tConn)
                    sql = "SELECT * FROM CLASS_SCHEDULE WHERE SchoolDate = " & TIMS.To_date(Common.FormatDate(SchoolDate)) & " AND OCID = '" & OCID.ToString & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("CSID") = DbAccess.GetNewId(trans, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                        dr("OCID") = OCID.ToString
                        dr("SchoolDate") = Common.FormatDate(SchoolDate)
                    Else
                        dr = dt.Rows(0)
                    End If
                Else
                    sql = "SELECT * FROM CLASS_SCHEDULE WHERE CSID = '" & CSID.ToString & "' "
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    dr = dt.Rows(0)
                End If

                dr("Formal") = "Y"
                dr("Type") = 2 '1:全期排課/2:單月排課
                For i As Integer = 1 To 12
                    Dim Tmps As String = ""
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("CourseID")))
                    dr("Class" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                    '1~12
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher1ID")))
                    dr("Teacher" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                    '13~24
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher2ID")))
                    dr("Teacher" & i + 12) = If(Tmps = "", Convert.DBNull, Tmps)
                    '25~36
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher3ID")))
                    dr("Teacher" & i + 24) = If(Tmps = "", Convert.DBNull, Tmps)
                    '37~48
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher4ID")))
                    dr("Teacher" & i + 36) = If(Tmps = "", Convert.DBNull, Tmps)

                    Tmps = TIMS.Get_Substr1(TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("ClassRoom"))), 30)
                    dr("Room" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                Next
                dr("VACATION") = If(v_Vacation <> "", v_Vacation, Convert.DBNull) '假日排課
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                DbAccess.UpdateDataTable(dt, da, trans)
                '先將課程本日課程資料填入Class_Schedule End

                'Plan_Schedule
                '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
                Call AddNew_Plan_Schedule(Me, STDate.Value, FTDate.Value, OCID.ToString, dt, da, trans, dtCourse)
                DbAccess.CommitTrans(trans)
            Catch ex As Exception
                Common.MessageBox(Me, ex.Message)
                Dim strErrmsg As String = String.Concat(" *ex.Message:", vbCrLf, ex.Message, vbCrLf, TIMS.GetErrorMsg(Me)) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg, ex)

                DbAccess.RollbackTrans(trans)
                Call TIMS.CloseDbConn(tConn)
                Throw ex
            End Try
            Call TIMS.CloseDbConn(tConn)
        End Using
        Return ""
    End Function

    '儲存課表
    Sub SaveCommonClass()
        'ByVal sender As System.Object, ByVal e As System.EventArgs
        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")
        Dim Mydt As DataTable = ViewState(cst_vsDetailTable)
        'IsNothingdtFlag 
        If TIMS.dtNODATA(Mydt) Then
            '"排課資料不存在，請重新查詢建立排課資料!!"
            Common.MessageBox(Me, cst_alertMsg2)
            Exit Sub
        End If
        Try
            CFDate.Text = CDate(CFDate.Text).ToString("yyyy/MM/dd")
        Catch ex As Exception
            '"排課區間迄止日期有誤，請重新設定 排課區間迄止日期!!"
            Common.MessageBox(Me, cst_alertMsg7)
            Exit Sub
        End Try
        Try
            MyDate.Text = CDate(MyDate.Text).ToString("yyyy/MM/dd")
        Catch ex As Exception
            '"指定日期有誤，請重新選擇指定日期!!"
            Common.MessageBox(Me, cst_alertMsg8)
            Exit Sub
        End Try

        'CSID.Value = TIMS.ClearSQM(CSID.Value)
        'If CSID.Value = "" Then
        '    '"排課資料不存在，請重新查詢建立排課資料!!"
        '    Common.MessageBox(Me, cst_alertMsg2)
        '    Exit Sub
        'End If
        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If CSID.Value <> "" Then
            '若不為空就檢測
            Dim dt As DataTable
            Dim sql As String = ""
            sql &= " SELECT * FROM CLASS_SCHEDULE "
            sql &= " WHERE CSID = '" & CSID.Value & "' AND OCID = '" & v_OCID & "' AND SchoolDate = " & TIMS.To_date(MyDate.Text)
            dt = DbAccess.GetDataTable(sql, objconn)
            If dt.Rows.Count = 0 Then
                '"排課資料不存在，請重新查詢建立排課資料!!"
                Common.MessageBox(Me, cst_alertMsg2)
                Exit Sub
            End If
        End If

        'Dim i As Integer = 0
        'SearchTable.Style.Item("display") = "inline"
        'DetailTable.Style.Item("display") = "none"

        If Mydt IsNot Nothing Then
            Dim StartDate As Date
            Dim EndDate As Date
            Dim dr As DataRow
            Dim dt As DataTable
            'Dim sql As String = ""
            Dim da As SqlDataAdapter = Nothing

            'Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
            Using tConn As SqlConnection = DbAccess.GetConnection()
                Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
                Try
                    'trans = DbAccess.BeginTrans(tConn)
                    '先將課程本日課程資料填入Class_Schedule Start
                    Dim sql As String = ""
                    If CSID.Value = "" Then
                        sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & v_OCID & "' AND SchoolDate = " & TIMS.To_date(MyDate.Text)
                        dt = DbAccess.GetDataTable(sql, da, trans)
                        If dt.Rows.Count = 0 Then
                            Dim iCSID As Integer = DbAccess.GetNewId(trans, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("CSID") = iCSID
                            dr("OCID") = v_OCID 'OCID.SelectedValue
                            dr("SchoolDate") = CDate(MyDate.Text)
                        Else
                            dr = dt.Rows(0)
                        End If
                    Else
                        sql = " SELECT * FROM CLASS_SCHEDULE WHERE CSID = '" & CSID.Value & "' "
                        dt = DbAccess.GetDataTable(sql, da, trans)
                        If dt.Rows.Count = 0 Then
                            '"排課資料不存在，請重新查詢建立排課資料!!"
                            Common.MessageBox(Me, cst_alertMsg2)
                            Exit Sub
                        End If
                        dr = dt.Rows(0)
                    End If

                    dr("Formal") = "Y"
                    dr("Type") = 2 '1:全期排課/2:單月排課
                    Dim Tmps As String = ""
                    For i As Integer = 1 To 12
                        Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("CourseID")))
                        '查無此課程代碼，不進行課堂時數相加了
                        dr("Class" & i) = Convert.DBNull 'IIf(Tmps = "", Convert.DBNull, Tmps)
                        If Tmps <> "" Then
                            If dtCourse.Select("CourID='" & Tmps & "'").Length > 0 Then dr("Class" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                        End If
                        Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher1ID")))
                        dr("Teacher" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                        Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher2ID")))
                        dr("Teacher" & i + 12) = If(Tmps = "", Convert.DBNull, Tmps)
                        Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher3ID")))
                        dr("Teacher" & i + 24) = If(Tmps = "", Convert.DBNull, Tmps)
                        '37~48
                        Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher4ID")))
                        dr("Teacher" & i + 36) = If(Tmps = "", Convert.DBNull, Tmps)
                        Tmps = TIMS.Get_Substr1(TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("ClassRoom"))), 30)
                        dr("Room" & i) = If(Tmps = "", Convert.DBNull, Tmps)
                    Next
                    dr("VACATION") = If(v_Vacation <> "", v_Vacation, Convert.DBNull) '假日排課
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, trans)
                    '先將課程本日課程資料填入Class_Schedule End

                    'Plan_Schedule
                    '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
                    Call AddNew_Plan_Schedule(Me, STDate.Value, FTDate.Value, v_OCID, dt, da, trans, dtCourse)
                    DbAccess.CommitTrans(trans)
                Catch ex As Exception
                    Common.MessageBox(Me, ex.Message)
                    Dim strErrmsg As String = String.Concat(" *ex.Message:", vbCrLf, ex.Message, vbCrLf, TIMS.GetErrorMsg(Me)) '取得錯誤資訊寫入
                    Call TIMS.WriteTraceLog(strErrmsg, ex)

                    DbAccess.RollbackTrans(trans)
                    Call TIMS.CloseDbConn(tConn)
                    'Throw ex
                    Exit Sub 'Throw ex
                End Try
            End Using

            'Call TIMS.CloseDbConn(tConn)

            '增加空白日期存入資料庫--   Start
            Dim flag As Boolean = False '要是為True，必須回存日期
            StartDate = STDate.Value
            EndDate = FTDate.Value
            'tConn = DbAccess.GetConnection
            '儲存課表
            Using tConn As SqlConnection = DbAccess.GetConnection()
                Call TIMS.OpenDbConn(tConn)
                'v_OCID OCID.SelectedValue
                Dim sql As String = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & v_OCID & "' AND SchoolDate >= CONVERT(DATETIME, '" & StartDate & "', 111) AND SchoolDate <= CONVERT(DATETIME, '" & EndDate & "', 111) AND Formal = 'Y' "
                dt = DbAccess.GetDataTable(sql, da, tConn)
                While (StartDate <= EndDate)
                    If dt.Select("SchoolDate='" & StartDate & "'").Length = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("CSID") = DbAccess.GetNewId(tConn, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                        dr("OCID") = v_OCID 'OCID.SelectedValue
                        dr("SchoolDate") = StartDate
                        dr("Formal") = "Y"
                        dr("Type") = 2 '1:全期排課/2:單月排課
                        dr("Vacation") = Convert.DBNull
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        flag = True
                    End If
                    StartDate = StartDate.AddDays(1)
                End While

                If flag Then DbAccess.UpdateDataTable(dt, da)
                '增加空白日期存入資料庫--   End
                Call TIMS.CloseDbConn(tConn)
            End Using

            'ClearItemValue()
            'Button2_Click(sender, e)
            Call sSearch2()
            If DesDate.Checked = True Then
                Page.RegisterStartupScript("ChangeDate", "<script>ChangeDate('" & MyDate.Text & "','" & STDate.Value & "','" & FTDate.Value & "')</script>")
                'Common.MessageBox(Me, "儲存成功!")
            Else
                ViewState(cst_vsDetailTable) = Nothing
                'Common.MessageBox(Me, "儲存成功!")
                If CDate(MyDate.Text) >= CDate(CFDate.Text) Then
                    Common.MessageBox(Me, "已經到最後一天!")
                Else
                    Dim Temp As Date = CDate(MyDate.Text).AddDays(1)
                    While HolidayTable.Select("HolDate='" & Temp & "'").Length <> 0
                        If CDate(Temp) < CDate(CFDate.Text) Then
                            Temp = CDate(Temp).AddDays(1)
                        Else
                            Common.MessageBox(Me, "已經到最後一天!")
                            Exit Sub
                        End If
                    End While

                    vsErrMsg = ""
                    'v_OCID OCID.SelectedValue
                    If CheckGetVal1(Temp, v_OCID, vsErrMsg) Then
                        CreateDetailCourse(Temp, dtTeacher)
                    Else
                        Common.MessageBox(Me, vsErrMsg)
                    End If
                End If
            End If
        End If
    End Sub

    '依 儲存課表 新增 課表計畫2
    Public Shared Sub sUtl_PTI(ByVal MyPage As Page, ByVal STDate_Value As String, ByVal FTDate_Value As String, ByVal OCID_Value As String, ByRef dt As DataTable, ByRef da As SqlDataAdapter, ByRef trans As SqlTransaction, ByRef CourseDataTable As DataTable, ByVal ss3 As String)
        'Dim Ti1 As String = TIMS.GetMyValue(ss3, "Ti1")
        'Dim Yc1 As String = TIMS.GetMyValue(ss3, "Yc1")
        Dim yj As String = TIMS.GetMyValue(ss3, "yj")
        Dim sm As SessionModel = SessionModel.Instance()
        'fix error start

        Dim dtErr As DataTable = Nothing
        Do
            dtErr = New DataTable
            Dim sql As String = ""
            sql &= " SELECT yearcount, titleitem, COUNT(1) cnt, MAX(psid) PSID " & vbCrLf
            sql &= " FROM PLAN_SCHEDULE" & vbCrLf
            sql &= " WHERE OCID=@OCID AND titleitem <= 4 " & vbCrLf
            sql &= " GROUP BY yearcount, titleitem " & vbCrLf
            sql &= " HAVING COUNT(1) > 1 " & vbCrLf
            Using sCmd As New SqlCommand(sql, trans.Connection, trans)
                With sCmd
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID_Value
                    dtErr.Load(.ExecuteReader())
                End With
            End Using
            If TIMS.dtHaveDATA(dtErr) Then
                For Each drErr As DataRow In dtErr.Rows
                    Dim dtS As New DataTable
                    Dim sqlS As String = " SELECT 1 FROM PLAN_SCHEDULE WHERE PSID=@PSID"
                    Using dCmd As New SqlCommand(sqlS, trans.Connection, trans)
                        With dCmd
                            .Parameters.Add("PSID", SqlDbType.Int).Value = drErr("PSID")
                            dtS.Load(.ExecuteReader())
                        End With
                    End Using
                    If TIMS.dtHaveDATA(dtS) Then
                        Dim sqlD As String = " DELETE PLAN_SCHEDULE WHERE PSID=@PSID"
                        Using dCmd As New SqlCommand(sqlD, trans.Connection, trans)
                            With dCmd
                                .Parameters.Add("PSID", SqlDbType.Int).Value = drErr("PSID")
                                .ExecuteNonQuery()
                            End With
                        End Using
                    End If
                Next
            End If
        Loop Until TIMS.dtNODATA(dtErr) 'dtErr.Rows.Count = 0
        'fix error end

        For yi As Integer = 0 To Val(yj)
            'Ti1 = "1"
            Dim Ti1 As String = ""
            Dim intNextYearWeek As Integer = 0 '設定下一輪週數
            Dim i As Integer = 1
            Dim StartDate As Date ' = CDate(STDate_Value)
            Dim EndDate As Date ' = CDate(FTDate_Value)
            Dim NextStartDate As Date '下一年度起始日

            StartDate = CDate(STDate_Value)
            EndDate = CDate(FTDate_Value)
            StartDate = StartDate.AddYears(yi)
            NextStartDate = StartDate.AddYears(1) '下一年度起始日
            i = 1
            Ti1 = "1" '週數
            Dim Sql_1 As String = String.Concat(" SELECT * FROM PLAN_SCHEDULE WHERE OCID = '", OCID_Value, "' AND TitleItem = ", Ti1, " AND YearCount = ", yi)
            dt = DbAccess.GetDataTable(Sql_1, da, trans)
            '儲存週數 Start
            Dim iPSID As Integer = 0
            Dim dr As DataRow = Nothing
            If dt.Rows.Count = 0 Then
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                iPSID = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                'If Not blnRC Then iPSID = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("PSID") = iPSID
                dr("OCID") = OCID_Value
                dr("TitleItem") = Ti1 '週數
                dr("YearCount") = yi '年數
            Else
                'If dt.Rows.Count <> 1 Then sql = " DELETE PLAN_SCHEDULE WHERE "
                dr = dt.Rows(0)
                iPSID = dr("PSID")
            End If

            'column loop
            While (StartDate <= EndDate) And StartDate <= NextStartDate
                dr("W" & i) = intNextYearWeek + i
                StartDate = StartDate.AddDays(7)
                i += 1
            End While
            intNextYearWeek = i - 1 '設定下一輪週數
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)
            '儲存週數 End

            '儲存月數 Start
            StartDate = CDate(STDate_Value)
            EndDate = CDate(FTDate_Value)
            StartDate = StartDate.AddYears(yi)
            NextStartDate = StartDate.AddYears(1) '下一年度起始日
            i = 1
            Ti1 = "2" '月數
            Dim Sql_2 As String = String.Concat(" SELECT * FROM PLAN_SCHEDULE WHERE OCID = '" & OCID_Value & "' ", " AND TitleItem = " & Ti1, " AND YearCount = " & yi)
            dt = DbAccess.GetDataTable(Sql_2, da, trans)
            If dt.Rows.Count = 0 Then
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                iPSID = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                'If Not blnRC Then iPSID = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("PSID") = iPSID
                dr("OCID") = OCID_Value
                dr("TitleItem") = Ti1 '月數
                dr("YearCount") = yi '年數
            Else
                dr = dt.Rows(0)
                iPSID = dr("PSID")
            End If
            While (StartDate <= EndDate) And StartDate <= NextStartDate
                dr("W" & i) = Month(StartDate)
                StartDate = StartDate.AddDays(7)
                i += 1
            End While
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)
            '儲存月數 End

            '儲存起始日 Start
            StartDate = CDate(STDate_Value)
            EndDate = CDate(FTDate_Value)
            StartDate = StartDate.AddYears(yi)
            NextStartDate = StartDate.AddYears(1) '下一年度起始日
            i = 1
            Ti1 = "3" '起始日
            Dim Sql_3 As String = String.Concat(" SELECT * FROM PLAN_SCHEDULE WHERE OCID = '" & OCID_Value & "' ", " AND TitleItem = " & Ti1, " AND YearCount = " & yi)
            dt = DbAccess.GetDataTable(Sql_3, da, trans)
            If dt.Rows.Count = 0 Then
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                iPSID = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                'If Not blnRC Then iPSID = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("PSID") = iPSID
                dr("OCID") = OCID_Value
                dr("TitleItem") = Ti1 '起始日
                dr("YearCount") = yi '年數
            Else
                dr = dt.Rows(0)
                iPSID = dr("PSID")
            End If
            While (StartDate <= EndDate) And StartDate <= NextStartDate
                dr("W" & i) = Day(StartDate)
                StartDate = StartDate.AddDays(7)
                i += 1
            End While
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)
            '儲存起始日 End

            '儲存結束日 Start
            StartDate = CDate(STDate_Value)
            EndDate = CDate(FTDate_Value)
            StartDate = StartDate.AddYears(yi)
            NextStartDate = StartDate.AddYears(1) '下一年度起始日
            i = 1
            Ti1 = "4" '結束日
            Dim Sql_4 As String = String.Concat(" SELECT * FROM PLAN_SCHEDULE WHERE OCID = '" & OCID_Value & "' ", " AND TitleItem = " & Ti1, " AND YearCount = " & yi)
            dt = DbAccess.GetDataTable(Sql_4, da, trans)
            If dt.Rows.Count = 0 Then
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                'Dim iPSID As Integer = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                iPSID = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                'If Not blnRC Then iPSID = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("PSID") = iPSID
                dr("OCID") = OCID_Value
                dr("TitleItem") = Ti1 '結束日
                dr("YearCount") = yi '年數
            Else
                dr = dt.Rows(0)
                iPSID = dr("PSID")
            End If
            StartDate = StartDate.AddDays(6)
            While (StartDate < EndDate) And StartDate <= NextStartDate
                dr("W" & i) = Day(StartDate)
                StartDate = StartDate.AddDays(7)
                i += 1
            End While
            If StartDate >= EndDate Then
                dr("W" & i) = Day(EndDate)
            End If
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, trans)
            '儲存結束日 End
        Next
    End Sub

    '依 儲存課表 新增 課表計畫
    Public Shared Sub AddNew_Plan_Schedule(ByVal MyPage As Page, ByVal STDate_Value As String, ByVal FTDate_Value As String, ByVal OCID_Value As String, ByRef dt As DataTable, ByRef da As SqlDataAdapter, ByRef trans As SqlTransaction, ByRef CourseDataTable As DataTable)
        Dim i As Integer = 0
        Dim NextStartDate As Date '下一年度起始日
        Dim sm As SessionModel = SessionModel.Instance()

        '將Plan_Schedule中所有的排課資料清除 Start
        Dim dtS As New DataTable
        Dim sqlS As String = " SELECT 1 FROM PLAN_SCHEDULE WHERE OCID=@OCID AND TitleItem=5"
        Using SCmd As New SqlCommand(sqlS, trans.Connection, trans)
            With SCmd
                .Parameters.Add("OCID", SqlDbType.BigInt).Value = OCID_Value
                dtS.Load(.ExecuteReader())
            End With
        End Using
        If TIMS.dtHaveDATA(dtS) Then
            Dim pmsD As New Hashtable From {{"OCID", OCID_Value}}
            Dim sqlD As String = " DELETE PLAN_SCHEDULE WHERE OCID=@OCID AND TitleItem=5"
            DbAccess.ExecuteNonQuery(sqlD, trans, pmsD)
        End If
        '將Plan_Schedule中所有的排課資料清除 End

        Dim StartDate As Date = CDate(STDate_Value)
        Dim EndDate As Date = CDate(FTDate_Value)
        Dim yj As Integer = 0 '總年度數
        yj = TIMS.DateDiffYears(StartDate, EndDate) '判斷年度(起始日與結束日是否超過1年)

        Call TIMS.OpenDbConn(trans.Connection)
        '假如沒有TitleItem=1~4的屬性，先加入 Start
        Dim ss3 As String = ""
        Call TIMS.SetMyValue(ss3, "yj", yj)
        Call sUtl_PTI(MyPage, STDate_Value, FTDate_Value, OCID_Value, dt, da, trans, CourseDataTable, ss3)
        '假如沒有TitleItem=1~4的屬性，先加入 End

        '將Plan_Schedule的資料更新 Start
        'sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & OCID.SelectedValue & "' Order BY SchoolDate"
        'CourseDataTable TIMS共用課程資料
        Dim sqlS2 As String = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & OCID_Value & "' ORDER BY SCHOOLDATE "
        Dim da1 As SqlDataAdapter = Nothing
        Dim ClassSch As DataTable = DbAccess.GetDataTable(sqlS2, da1, trans)

        'StartDate = STDate.Value
        'EndDate = FTDate.Value
        'NextStartDate = StartDate.AddYears(1) '下個年度
        For yi As Integer = 0 To yj
            StartDate = CDate(STDate_Value)
            StartDate = StartDate.AddYears(yi)
            NextStartDate = StartDate.AddYears(1)
            EndDate = CDate(FTDate_Value)
            i = 1 '(每次迴圈從1計算)計算週數  i表示第幾週
            '假如尚未到最後一天
            While (StartDate <= EndDate) And StartDate <= NextStartDate
                For Each ClassSchRow As DataRow In ClassSch.Select("SchoolDate>='" & StartDate & "' and SchoolDate<'" & StartDate.AddDays(7) & "'", "SchoolDate")
                    '課堂數，一天12堂
                    For j As Integer = 1 To 12   '迴圈跑Class1  Class2  .....
                        If ClassSchRow("Class" & j).ToString <> "" Then
                            '查無此課程代碼，不進行課堂時數相加了
                            If CourseDataTable.Select("CourID='" & ClassSchRow("Class" & j) & "'").Length > 0 Then
                                Dim iPSID As Integer = 0
                                Dim dr As DataRow = Nothing
                                If CourseDataTable.Select("CourID='" & ClassSchRow("Class" & j) & "'")(0)("MainCourID").ToString = "" Then            '表示為主課程
                                    If dt.Select("CourID='" & ClassSchRow("Class" & j) & "'").Length = 0 Then           '表示Plan_Schedule無使此課程資料
                                        'Dim iPSID As Integer = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                                        'Dim iPSID As Integer = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                                        iPSID = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                                        'If Not blnRC Then iPSID = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)
                                        dr("PSID") = iPSID
                                        dr("OCID") = OCID_Value
                                        dr("CourID") = ClassSchRow("Class" & j)
                                        dr("TitleItem") = 5
                                    Else
                                        dr = dt.Select("CourID='" & ClassSchRow("Class" & j) & "'")(0)
                                    End If
                                Else
                                    '副課程，所以要將時數存入主課程之時數
                                    If dt.Select("CourID='" & CourseDataTable.Select("CourID='" & ClassSchRow("Class" & j) & "'")(0)("MainCourID").ToString & "'").Length = 0 Then
                                        '表示Plan_Schedule無主課程資料
                                        'Dim iPSID As Integer = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                                        'Dim iPSID As Integer = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                                        iPSID = DbAccess.GetNewId(trans, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID")
                                        'If Not blnRC Then iPSID = DbAccess.GetNewId(trans, trans.Connection, "PLAN_SCHEDULE_PSID_SEQ,PLAN_SCHEDULE,PSID", 10)
                                        dr = dt.NewRow
                                        dt.Rows.Add(dr)
                                        dr("PSID") = iPSID
                                        dr("OCID") = OCID_Value
                                        dr("CourID") = CourseDataTable.Select("CourID='" & ClassSchRow("Class" & j) & "'")(0)("MainCourID").ToString
                                        dr("TitleItem") = 5
                                    Else
                                        dr = dt.Select("CourID='" & CourseDataTable.Select("CourID='" & ClassSchRow("Class" & j) & "'")(0)("MainCourID").ToString & "'")(0)
                                    End If
                                End If

                                If IsDBNull(dr("W" & i)) Then
                                    dr("W" & i) = 1 '每週，課堂數加1
                                Else
                                    dr("W" & i) = dr("W" & i) + 1 '每週，課堂數加1
                                End If
                                dr("YearCount") = Convert.DBNull '課堂資訊不分年度,全部加起來
                                dr("ModifyAcct") = sm.UserInfo.UserID
                                dr("ModifyDate") = Now
                            End If
                        End If
                    Next
                Next
                StartDate = StartDate.AddDays(7)
                i += 1
            End While
        Next

        '將Plan_Schedule的資料更新 End
        DbAccess.UpdateDataTable(dt, da, trans)
    End Sub

    '(儲存) 課表
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'If Session("GUID1") = ViewState("GUID1") Then '必免重複存取
        '    Session("GUID1") = "" '必免重複存取
        'Button3.Enabled = False
        '一般排課
        Call SaveCommonClass()
        Exit Sub

    End Sub

    '(回排課列表)
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        SearchTable.Style.Item("display") = TIMS.cst_inline1 '"inline"
        DetailTable.Style.Item("display") = "none"
        ClearItemValue()
        'GetMySearch.Value = "0"
        'IFRAME1.Attributes("src") = "SD_04_002_Course.aspx?RID=" & RIDValue.Value
        ViewState(cst_vsDetailTable) = Nothing
    End Sub

    '檢查，新增排課 Class_Schedule (匯入用陣列)
    Function Insert_New_Class_Schedule(ByVal colArray As System.Array) As String
        Dim Reason As String = "" '錯誤訊息
        '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
        Dim dt As DataTable = ViewState(cst_vsDetailTable)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, "排課資料不存在，請重新查詢建立排課資料!!") 'Exit Function'刪除資料不存在
            Return Reason
        End If

        Dim UsedHour As Integer = 0
        Dim TodayLeftHour As Integer = 0
        Dim ItemValue As String = ""
        Dim ErrorMsg As String = ""
        'Dim i As Integer = 0
        'colArray
        '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
        Const cst_Classaddvalue As Integer = 0
        Const cst_Teacher1addvalue As Integer = 12
        Const cst_Teacher2addvalue As Integer = 24
        Const cst_ClassRoomaddvalue As Integer = 36
        Const cst_Teacher3addvalue As Integer = 48
        Const cst_Teacher4addvalue As Integer = 50

        TodayLeftHour = Int(LeftHour.Text) - Int(TodayUseHour.Value)
        For i As Integer = 1 To 12
            '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
            If colArray(i + cst_Classaddvalue).ToString <> "" Then
                If dt.Select("ClassNum='" & i & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select("ClassNum='" & i & "'")(0)
                    If Convert.ToString(dr("CourseName")) = "" Then LeftHour.Text = Int(Val(LeftHour.Text)) - 1 '可用時數減少
                    Dim Tmps As String = ""
                    Tmps = TIMS.Get_CourseName(colArray(i + cst_Classaddvalue).ToString, Nothing, objconn)
                    dr("CourseName") = If(Tmps = "", Convert.DBNull, Tmps) 'CourseID.Text
                    Tmps = TIMS.ClearSQM(colArray(i + cst_Classaddvalue).ToString)
                    dr("CourseID") = If(Tmps = "", Convert.DBNull, Tmps) 'CourseIDValue.Value
                    If Convert.ToString(dr("CourseName")) = "" Then Reason += "第" & CStr(i) & "堂 " & Convert.ToString(dr("CourseID")) & "課程名稱代碼輸入有誤<BR>"
                    Tmps = TIMS.Get_Substr1(TIMS.ClearSQM(colArray(i + cst_ClassRoomaddvalue).ToString), 30)
                    dr("ClassRoom") = If(Tmps = "", Convert.DBNull, Tmps) 'Room.Text
                    If Convert.ToString(dr("ClassRoom")) = "" Then Reason += "第" & CStr(i) & "堂 教室 不可為空<BR>"

                    Dim xi As Integer = 0
                    xi = i + cst_Teacher1addvalue
                    If colArray(xi).ToString <> "" Then
                        If Not TIMS.Get_TeacherDegree(colArray(xi), objconn) = Nothing Then
                            dr("Teacher1") = TIMS.Get_TeachCName(colArray(xi).ToString, objconn) 'TIMS.Get_TeacherName(colArray(xi).ToString)  'OLessonTeah1.Text
                            dr("Teacher1ID") = TIMS.ClearSQM(colArray(xi).ToString) 'OLessonTeah1Value.Value
                        Else
                            Reason += "第" & CStr(i) & "堂 老師一 師資流水ID有誤(" & colArray(xi).ToString & ")<BR>"
                        End If
                    Else
                        Reason += "第" & CStr(i) & "堂 老師一 不可為空<BR>"
                    End If

                    xi = i + cst_Teacher2addvalue
                    If colArray(xi).ToString <> "" Then
                        If Not TIMS.Get_TeacherDegree(colArray(xi), objconn) = Nothing Then
                            dr("Teacher2") = TIMS.Get_TeachCName(colArray(xi).ToString, objconn) 'TIMS.Get_TeacherName(colArray(xi).ToString) 'OLessonTeah2.Text
                            dr("Teacher2ID") = TIMS.ClearSQM(colArray(xi).ToString) 'OLessonTeah2Value.Value
                        Else
                            Reason += "第" & CStr(i) & "堂 助教一 師資流水ID有誤(" & colArray(xi).ToString & ")<BR>"
                        End If
                    End If

                    xi = i + cst_Teacher3addvalue
                    If colArray(xi).ToString <> "" Then
                        If Not TIMS.Get_TeacherDegree(colArray(xi), objconn) = Nothing Then
                            dr("Teacher3") = TIMS.Get_TeachCName(colArray(xi).ToString, objconn) 'TIMS.Get_TeacherName(colArray(xi).ToString) 'OLessonTeah3.Text
                            dr("Teacher3ID") = TIMS.ClearSQM(colArray(xi).ToString) 'OLessonTeah3Value.Value
                        Else
                            Reason += "第" & CStr(i) & "堂 助教二 師資流水ID有誤(" & colArray(xi).ToString & ")<BR>"
                        End If
                    End If

                    xi = i + cst_Teacher4addvalue
                    If colArray.Length > xi Then
                        If colArray(xi).ToString <> "" Then
                            If Not TIMS.Get_TeacherDegree(colArray(xi), objconn) = Nothing Then
                                dr("Teacher4") = TIMS.Get_TeachCName(colArray(xi).ToString, objconn) 'TIMS.Get_TeacherName(colArray(xi).ToString) 'OLessonTeah4.Text
                                dr("Teacher4ID") = TIMS.ClearSQM(colArray(xi).ToString) 'OLessonTeah4Value.Value
                            Else
                                Reason += "第" & CStr(i) & "堂 助教1 師資流水ID有誤(" & colArray(xi).ToString & ")<BR>"
                            End If
                        End If
                    End If
                End If
            End If
        Next
        For Each dr1 As DataRow In dt.Rows
            If dr1("CourseName").ToString <> "" Then UsedHour += 1
        Next

        Call GetLeftCourseHour(Int(THours.Text)) '更新今天使用的時數和剩餘時數

        ViewState(cst_vsDetailTable) = dt
        DataGrid2.DataSource = ViewState(cst_vsDetailTable)
        DataGrid2.DataBind()

        Return Reason
    End Function

    '新增排課
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        'ViewState("GUID1") = TIMS.GetGUID
        'Session("GUID1") = ViewState("GUID1")
        Dim Reason As String = ""
        'Reason = ""
        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")
        MyDate.Text = TIMS.ClearSQM(MyDate.Text)
        STDate.Value = TIMS.ClearSQM(STDate.Value)
        FTDate.Value = TIMS.ClearSQM(FTDate.Value)
        If MyDate.Text = "" Then
            Reason += cst_alertMsg9 & vbCrLf
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If
        If STDate.Value = "" Then
            Reason += cst_alertMsg9 & vbCrLf
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If
        If FTDate.Value = "" Then
            Reason += cst_alertMsg9 & vbCrLf
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '若日期區間落在開、結訓日期中，為True，其餘為False
        If Not CheckDate_NoUpdata(CDate(MyDate.Text), CDate(STDate.Value), CDate(FTDate.Value)) Then
            '新增排課停用
            'Reason += "排課日期區間未落在開、結訓日期中 無法新增排課!!" & vbCrLf
            Reason += cst_alertMsg9 & vbCrLf
        End If
        CourseIDValue.Value = TIMS.ClearSQM(CourseIDValue.Value)
        If CourseIDValue.Value = "" Then
            Reason += cst_alertMsg17 & vbCrLf '"查無課程代碼!!" & vbCrLf
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        If Reason = "" Then
            If dtCourse Is Nothing Then
                Dim ssRID As String = sm.UserInfo.RID
                If RIDValue.Value <> "" Then ssRID = RIDValue.Value
                dtCourse = TIMS.Get_COURSEINFOdt(ssRID, objconn)
            End If
            ff3 = "COURID='" & CourseIDValue.Value & "'"
            If dtCourse.Select(ff3).Length = 0 Then
                Reason += cst_alertMsg16 & vbCrLf
                Common.MessageBox(Me, Reason)
                Exit Sub
            End If

            Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
            Dim ss3 As String = ""
            TIMS.SetMyValue(ss3, "Classification1", Convert.ToString(dr1("CLASSIFICATION1")))
            TIMS.SetMyValue(ss3, "MaxTNum", hid_TNum.Value)
            TIMS.SetMyValue(ss3, "OLessonTeah1_Value", OLessonTeah1Value.Value)
            TIMS.SetMyValue(ss3, "OLessonTeah2_Value", OLessonTeah2Value.Value)
            TIMS.SetMyValue(ss3, "OLessonTeah3_Value", OLessonTeah3Value.Value)
            TIMS.SetMyValue(ss3, "OLessonTeah4_Value", OLessonTeah4Value.Value)
            'out@Reason
            Call TIMS.Chk_TIMSB1251(Me, Reason, ss3, objconn)
        End If
        If Reason <> "" Then
            Reason = "錯誤訊息如下:" & vbCrLf & Reason
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        Dim dt As DataTable = ViewState(cst_vsDetailTable)
        If TIMS.dtNODATA(dt) Then
            '"排課資料不存在，請重新查詢建立排課資料!!"
            Common.MessageBox(Me, cst_alertMsg2)
            Exit Sub
        End If

        Dim UsedHour As Integer = 0
        Dim TodayLeftHour As Integer = 0
        'Dim ItemValue, ErrorMsg As String

        TodayLeftHour = Int(LeftHour.Text) - Int(TodayUseHour.Value)
        'TodayLeftHour = Int(LeftHour.Text) - Int(TodayUseHour.Value)
        ' 判斷同一教師是否有出現在不同的教室  Start By Amu 2006/10/13
        'For Each Item As ListItem In ClassSort1.Items
        '    If Item.Selected = True And Int(LeftHour.Text) > 0 Then
        '        ItemValue += Item.Value & ","
        '    End If
        'Next

        'If Check_Class_Schedule_Duplicate(ItemValue, ErrorMsg) Then
        '    If ErrorMsg <> "" Then
        '        Common.RespWrite(Me, "<script>alert('" & ErrorMsg & "');</script>")
        '        'Response.End()
        '        Exit Sub
        '    End If
        'End If
        ' 判斷同一教師是否有出現在不同的教室 End

        For Each Item As ListItem In ClassSort1.Items
            If Item.Selected = True AndAlso Int(LeftHour.Text) > 0 Then
                If dt.Select("ClassNum='" & Item.Value & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select("ClassNum='" & Item.Value & "'")(0)
                    If dr("CourseName").ToString = "" Then LeftHour.Text = Int(LeftHour.Text) - 1 '可用時數減少
                    dr("CourseName") = CourseID.Text
                    dr("CourseID") = CourseIDValue.Value 'CourseIDValue.Value 流水id 用以顯現資料
                    dr("ClassRoom") = TIMS.Get_Substr1(Room.Text, 30)
                    dr("Teacher1") = OLessonTeah1.Text
                    dr("Teacher1ID") = OLessonTeah1Value.Value
                    dr("Teacher2") = OLessonTeah2.Text
                    dr("Teacher2ID") = OLessonTeah2Value.Value
                    dr("Teacher3") = OLessonTeah3.Text
                    dr("Teacher3ID") = OLessonTeah3Value.Value
                    dr("Teacher4") = OLessonTeah4.Text
                    dr("Teacher4ID") = OLessonTeah4Value.Value
                    dr("VACATION") = If(v_Vacation <> "", v_Vacation, "") '假日排課
                End If
            End If
        Next
        For Each dr1 As DataRow In dt.Rows
            If Convert.ToString(dr1("CourseName")) <> "" Then UsedHour += 1
        Next

        '更新今天使用的時數和剩餘時數
        'TodayUseHour.Value = UsedHour
        'LeftHour.Text = Int(LeftHour.Text) - Int(TodayUseHour.Value)
        Call GetLeftCourseHour(Int(THours.Text))

        ViewState(cst_vsDetailTable) = dt
        DataGrid2.DataSource = ViewState(cst_vsDetailTable)
        DataGrid2.DataBind()

        'ClearItemValue()
        'Call GetUsedClass()
        'IFRAME1.Attributes("src") = "SD_04_002_Course.aspx?RID=" & RIDValue.Value
    End Sub

    Sub ClearItemValue()
        CourseID.Text = ""
        CourseIDValue.Value = ""
        Room.Text = ""
        OLessonTeah1.Text = ""
        OLessonTeah1Value.Value = ""
        OLessonTeah2.Text = ""
        OLessonTeah2Value.Value = ""
        OLessonTeah3.Text = ""
        OLessonTeah3Value.Value = ""
        OLessonTeah4.Text = ""
        OLessonTeah4Value.Value = ""

        For i As Integer = 0 To ClassSort1.Items.Count - 1
            ClassSort1.Items(i).Selected = False
        Next

        ClassSort2.Checked = False
        ClassSort3.Checked = False
        ClassSort4.Checked = False
        ClassSort5.Checked = False
    End Sub

    '確認勾選的節次
    Function CheckClassSort1x() As String
        Dim rst As String = ""
        For Each Item As ListItem In ClassSort1.Items
            'If Item.Selected = True And Int(LeftHour.Text) > 0 Then
            '有勾選的節次
            If Item.Selected Then
                If rst <> "" Then rst &= ","
                rst &= Item.Value
            End If
        Next
        Return rst
    End Function

    '刪除排課-依節數
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        'ViewState("GUID1") = TIMS.GetGUID
        'Session("GUID1") = ViewState("GUID1")
        Dim dt As DataTable = ViewState(cst_vsDetailTable)
        If TIMS.dtNODATA(dt) Then
            '"排課資料不存在，請重新查詢建立排課資料!!"
            Common.MessageBox(Me, cst_alertMsg2)
            Exit Sub
        End If

        '確認勾選節次
        Dim vClassSort1x As String = CheckClassSort1x()
        If vClassSort1x = "" Then
            Dim msg As String = "節次不可為空，請勾選要刪除的排課節次!!!" & vbCrLf
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        ViewState("LeftHour") = Nothing
        For Each MyItem As ListItem In ClassSort1.Items
            If MyItem.Selected Then
                If dt.Select("ClassNum='" & MyItem.Value & "'").Length <> 0 Then
                    Dim dr As DataRow = dt.Select("ClassNum='" & MyItem.Value & "'")(0)
                    If dr("CourseName").ToString <> "" Then LeftHour.Text = Int(LeftHour.Text) + 1 '可用時數增加 
                    dr("CourseName") = ""
                    dr("CourseID") = ""
                    dr("ClassRoom") = ""
                    dr("Teacher1") = ""
                    dr("Teacher1ID") = ""
                    dr("Teacher2") = ""
                    dr("Teacher2ID") = ""
                    dr("Teacher3") = ""
                    dr("Teacher3ID") = ""
                    dr("Teacher4") = ""
                    dr("Teacher4ID") = ""
                    ViewState("LeftHour") = Me.LeftHour.Text
                End If
            End If
        Next

        ViewState(cst_vsDetailTable) = dt
        DataGrid2.DataSource = ViewState(cst_vsDetailTable)
        DataGrid2.DataBind()

        'ClearItemValue()
        If LeftHour.Text >= 0 Then
            '假如審核確認，則不可以修改資料
            If ViewState("IsVerify") = "Y" Then
                '已審核確認
                Button12.Enabled = False '刪除
                Button10.Enabled = False '刪除
                Button9.Enabled = False '新增
                Button3.Enabled = False '儲存
                TIMS.Tooltip(Button12, "此班級已審核確認", True)
                TIMS.Tooltip(Button10, "此班級已審核確認", True)
                TIMS.Tooltip(Button9, "此班級已審核確認", True)
                TIMS.Tooltip(Button3, "此班級已審核確認", True)
            Else
                Button9.Enabled = True
                Button3.Enabled = True
                TIMS.Tooltip(Button3, "")
                TIMS.Tooltip(Button9, "")
            End If
        Else
            '排課時數已經用完!
            If Not ViewState("LeftHour") Is Nothing Then
                ViewState("LeftHour") = Nothing
                '採取刪除動作打開儲存鈕
                Button3.Enabled = True
            End If
        End If

        'Call GetUsedClass()
        'IFRAME1.Attributes("src") = "SD_04_002_Course.aspx?RID=" & RIDValue.Value
    End Sub

    '刪除排課-依排課區間刪除
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        'CheckBox2'依排課區間刪除
        Dim CSDateText As String = TIMS.Cdate3(CSDate.Text)
        Dim CFDateText As String = TIMS.Cdate3(CFDate.Text)

        Dim v_OCID As String = TIMS.GetListValue(OCID)
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        'tConn = DbAccess.GetConnection
        sql = " SELECT 'x' x FROM Stud_TrainingResults WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID = '" & v_OCID & "') "
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            ' "此班級已經輸入成績，不能刪除資料"
            Common.MessageBox(Me, cst_alertMsg10)
            Exit Sub
        End If

        Using tConn As SqlConnection = DbAccess.GetConnection()
            Dim trans As SqlTransaction = DbAccess.BeginTrans(tConn)
            Try
                'trans = DbAccess.BeginTrans(tConn)
                Dim dtS As New DataTable
                Dim sqlS As String = " SELECT 1 FROM CLASS_SCHEDULE WHERE OCID=@OCID AND Type = '2'"
                Using SCmd As New SqlCommand(sqlS, trans.Connection, trans)
                    With SCmd
                        .Parameters.Add("OCID", SqlDbType.BigInt).Value = v_OCID
                        dtS.Load(.ExecuteReader())
                    End With
                End Using
                If TIMS.dtHaveDATA(dtS) Then
                    Dim pmsD As New Hashtable From {{"OCID", v_OCID}}
                    Dim sqlD As String = " DELETE CLASS_SCHEDULE WHERE OCID =@OCID AND Type = '2' "
                    If CheckBox2.Checked Then
                        If CSDateText <> "" Then sqlD &= " AND SchoolDate >= " & TIMS.To_date(CSDateText)
                        If CFDateText <> "" Then sqlD &= " AND SchoolDate <= " & TIMS.To_date(CFDateText)
                    End If
                    DbAccess.ExecuteNonQuery(sqlD, trans, pmsD)
                End If

                Dim dtS2 As New DataTable
                Dim sqlS2 As String = " SELECT 1 FROM PLAN_SCHEDULE WHERE OCID=@OCID"
                Using SCmd2 As New SqlCommand(sqlS2, trans.Connection, trans)
                    With SCmd2
                        .Parameters.Add("OCID", SqlDbType.BigInt).Value = v_OCID
                        dtS2.Load(.ExecuteReader())
                    End With
                End Using
                If TIMS.dtHaveDATA(dtS2) Then
                    '計算排課資料 刪除
                    Dim pmsD As New Hashtable From {{"OCID", v_OCID}}
                    Dim sqlD As String = " DELETE PLAN_SCHEDULE WHERE OCID=@OCID"
                    DbAccess.ExecuteNonQuery(sqlD, trans, pmsD)
                End If
                DbAccess.CommitTrans(trans)
            Catch ex As Exception
                Common.MessageBox(Me, ex.Message)
                Dim strErrmsg As String = String.Concat(" *ex.Message:", vbCrLf, ex.Message, vbCrLf, TIMS.GetErrorMsg(Me)) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg, ex)

                DbAccess.RollbackTrans(trans)
                Call TIMS.CloseDbConn(tConn)
                Throw ex
            End Try
            Call TIMS.CloseDbConn(tConn)
        End Using

        'Button2_Click(sender, e) '重算排課資料
        Call sSearch2()
        Common.MessageBox(Me, "刪除成功!")

        Button12.Enabled = False
        TIMS.Tooltip(Button12, "已刪除排課資料")
    End Sub

    '指定日期按鈕
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        'ViewState("GUID1") = TIMS.GetGUID
        'Session("GUID1") = ViewState("GUID1")
        DecideDate.Value = TIMS.ClearSQM(DecideDate.Value)
        'Dim sDecideDate As String = DecideDate.Value
        Try
            DecideDate.Value = CDate(DecideDate.Value).ToString("yyyy/MM/dd")
        Catch ex As Exception
            '"指定日期有誤，請再確認!!"
            Common.MessageBox(Me, cst_alertMsg11)
            Exit Sub
        End Try

        If Not IsDate(DecideDate.Value) Then
            '"指定日期有誤，請再確認!!"
            Common.MessageBox(Me, cst_alertMsg11)
            Exit Sub
        End If

        If HolidayTable.Select("HolDate='" & DecideDate.Value & "'").Length <> 0 Then
            Common.MessageBox(Me, "指定日期為假日!")
            Exit Sub
        End If

        ViewState(cst_vsDetailTable) = Nothing
        vsErrMsg = ""
        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If Not CheckGetVal1(DecideDate.Value, v_OCID, vsErrMsg) Then
            Common.MessageBox(Me, vsErrMsg)
            Exit Sub
        End If

        CreateDetailCourse(DecideDate.Value, dtTeacher)
        DecideDate.Value = ""
    End Sub

    '(第1次) 進入時查詢班級SQL
    Sub sUtl_Search1()
        'Dim sql As String = ""
        'Dim dt As DataTable
        'Dim dr As DataRow
        '依sm.UserInfo.PlanID取得PlanKind  '1:自辦(內訓) 2:委外
        Dim PlanKind As Integer = TIMS.Get_PlanKind(Me, objconn) 'dr("PlanKind")
        CourKeyWord.Text = TIMS.ClearSQM(CourKeyWord.Text)

        Dim sql As String = ""
        '依登入計畫年度
        sql = ""
        sql &= " SELECT p.OCID, p.ClassCName, p.CyclType"
        sql &= " ,dbo.FN_GET_CLASSCNAME(p.ClassCName,p.CyclType) ClassCName2"
        sql &= " FROM CLASS_CLASSINFO p " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        If sm.UserInfo.RID <> "A" Then
            sql &= " AND p.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        sql &= " AND p.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= " AND p.IsSuccess = 'Y' " & vbCrLf
        sql &= " AND p.NotOpen = 'N' " & vbCrLf
        If (CourKeyWord.Text <> "") Then sql &= String.Format(" AND dbo.FN_GET_CLASSCNAME(p.ClassCName,p.CyclType) LIKE '%{0}%'", CourKeyWord.Text) & vbCrLf
        If sm.UserInfo.RID = "A" Then
            '依登入者 計畫查詢
            sql &= " AND EXISTS ( " & vbCrLf
            sql &= "    SELECT 'x' FROM ID_PLAN x " & vbCrLf
            sql &= "    WHERE 1=1 " & vbCrLf
            sql &= "    AND x.Years = '" & sm.UserInfo.Years & "' " & vbCrLf
            sql &= "    AND x.TPlanID = '" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= "    AND p.PlanID = x.PlanID) " & vbCrLf
        Else
            '依登入者 可查詢權限
            If PlanKind = 1 Then '依sm.UserInfo.PlanID取得PlanKind  '1:自辦(內訓) 2:委外
                sql &= " AND EXISTS (SELECT 'x' FROM AUTH_ACCRWCLASS WHERE Account = '" & sm.UserInfo.UserID & "' AND p.OCID = OCID) " & vbCrLf
            End If
        End If
        sql &= " ORDER BY p.ClassCName ,p.CyclType " & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        OCID.Items.Clear()
        If dt.Rows.Count = 0 Then
            OCID.Items.Add(New ListItem(Cst_OCIDNULL, ""))
        Else
            Dim OCIDStr As String = ""
            OCIDStr = ""
            For Each dr As DataRow In dt.Rows
                If OCIDStr <> "" Then OCIDStr &= ","
                OCIDStr &= CStr(dr("OCID"))
            Next

            sql = ""
            sql &= " SELECT OCID, Type FROM V_SCHEDULETYPE WHERE 1=1 "
            If OCIDStr.IndexOf(",") > -1 Then
                sql &= " AND OCID IN (" & OCIDStr & ") " & vbCrLf
            Else
                sql &= " AND OCID = " & OCIDStr & " " & vbCrLf
            End If
            Dim dt1 As DataTable
            dt1 = DbAccess.GetDataTable(sql, objconn)

            '審核確認
            sql = ""
            sql &= " SELECT DISTINCT OCID, APPRESULT FROM CLASS_SCHVERIFY WHERE 1=1 "
            If OCIDStr.IndexOf(",") > -1 Then
                sql &= " AND OCID IN (" & OCIDStr & ") " & vbCrLf
            Else
                sql &= " AND OCID = " & OCIDStr & " " & vbCrLf
            End If
            Dim dt2 As DataTable
            dt2 = DbAccess.GetDataTable(sql, objconn)

            OCID.Items.Add(New ListItem("請選擇", ""))
            'Dim ClassName As String

            For Each dr As DataRow In dt.Rows
                Dim ClassName As String = Convert.ToString(dr("ClassCName2"))
                '審核確認
                If dt2.Select("OCID='" & dr("OCID") & "'").Length <> 0 Then ClassName &= "-(已審核)"

                If dt1.Select("OCID='" & dr("OCID") & "'").Length = 0 Then
                    ClassName &= "---(無紀錄)"
                Else
                    '1:全期排課/2:單月排課
                    If dt1.Select("OCID='" & dr("OCID") & "'")(0)("Type") = 1 Then
                        ClassName &= "---(全期)"
                    Else
                        ClassName &= "---(單月)"
                    End If
                End If
                OCID.Items.Add(New ListItem(ClassName, dr("OCID")))
            Next
        End If

        '980224 fix 只有一筆資料則將TIMS.cst_ddl_PleaseChoose3拿掉
        If dt.Rows.Count = 1 Then
            OCID.SelectedIndex = 1
            'OCID_SelectedIndexChanged(sender, e)
            Call Utl_OCIDSel()
        End If
    End Sub

    Sub Utl_OCIDSel()
        Session("Class_CourseName") = ""

        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If OCID.SelectedIndex <> 0 AndAlso v_OCID <> "" Then

            Dim tOCID As String = OCID.SelectedValue
            Dim dr As DataRow = TIMS.GetOCIDDate(tOCID, objconn)

            If Not dr Is Nothing Then
                hid_TNum.Value = Convert.ToString(dr("TNum"))
                Session("Class_CourseName") = TIMS.Get_PNamePlanTrainDesc(dr("planid"), dr("comidno"), dr("seqno"), objconn)

                CheckBox1.Checked = False
                CSDate.Text = ""
                CFDate.Text = ""
                Years.Text = Convert.ToString(dr("Years"))

                CyclType.Text = "無分期"
                If Convert.ToString(dr("CyclType")) <> "" AndAlso Convert.ToString(dr("CyclType")) <> "0" AndAlso Convert.ToString(dr("CyclType")) <> "00" Then CyclType.Text = "第" & Convert.ToString(dr("CyclType")) & "期"

                CyclTypeValue.Value = ""
                If Convert.ToString(dr("CyclType")) <> "" Then CyclTypeValue.Value = Val(dr("CyclType"))

                DayCount.Value = DateDiff(DateInterval.Day, CDate(dr("STDate")), CDate(dr("FTDate"))) + 1

                OCID2.Text = ""
                OCIDValue2.Value = ""
                OCIDValue1.Value = tOCID
                TPeriodValue.Value = Convert.ToString(dr("TPeriod"))
                TPeriod.Text = TIMS.Get_Value(TPeriodValue.Value, "Key_HourRan,HRID,HourRanName", objconn)
                If Convert.ToString(dr("STDate")) <> "" Then STDate.Value = Common.FormatDate(dr("STDate"))
                If Convert.ToString(dr("FTDate")) <> "" Then FTDate.Value = Common.FormatDate(dr("FTDate"))
                labTDate.Text = STDate.Value & "~" & FTDate.Value
                THours.Text = Convert.ToString(dr("THours"))

                CheckBox1.Attributes("onclick") = "if(this.checked){document.form1.CSDate.value='" & Common.FormatDate(dr("STDate")) & "';document.form1.CFDate.value='" & Common.FormatDate(dr("FTDate")) & "';}else{document.form1.CSDate.value='';document.form1.CFDate.value='';}"

                CourseTable.Style.Item("display") = "none"
                DataGrid3.Visible = False

                Dim sql As String = ""
                sql = "SELECT DISTINCT OCID, Type FROM CLASS_SCHEDULE WHERE OCID = " & tOCID
                Dim dt1 As New DataTable
                dt1.Load(DbAccess.GetReader(sql, objconn))
                'dt1 = DbAccess.GetDataTable(sql, objConn)

                SysInfo.Text = ""
                Button2.Enabled = True
                Button14.Enabled = True
                LoadIntoClass.Enabled = True
                File1.Disabled = False

                ff = "OCID='" & tOCID & "'"
                If dt1.Select(ff).Length > 0 Then
                    '1:全期排課/2:單月排課
                    If CStr(dt1.Select(ff)(0)("Type")) = "1" Then
                        SysInfo.Text = "此班級已經使用全期排課" & "<br>"
                        Button2.Enabled = False
                        Button14.Enabled = False
                        LoadIntoClass.Enabled = False
                        File1.Disabled = True
                    End If
                End If
                If dr("IsClosed").ToString = "Y" Then
                    SysInfo.Text += "此班級已經結訓"
                    Button14.Enabled = False
                    LoadIntoClass.Enabled = False
                    File1.Disabled = True
                End If

                ViewState("IsClosed") = Convert.ToString(dr("IsClosed"))
                ViewState("IsVerify") = ""
                '假如審核確認，則不可以修改資料
                Button15.Enabled = True
                If TIMS.Chk_ClassSchVerify(tOCID, objconn) Then
                    If SysInfo.Text <> "" Then SysInfo.Text += "，已審核確認" Else SysInfo.Text += "此班級已審核確認" '已審核確認
                    ViewState("IsVerify") = "Y"
                    Button12.Enabled = False
                    Button15.Enabled = False
                    TIMS.Tooltip(Button12, "")
                    TIMS.Tooltip(Button12, "此班級已審核確認")
                    Button14.Enabled = False
                    LoadIntoClass.Enabled = False
                    File1.Disabled = True
                End If

                '判斷訓練時段
                ClassSort2.Enabled = True
                ClassSort3.Enabled = True
                ClassSort4.Enabled = True
                ClassSort5.Enabled = True

                Select Case TPeriodValue.Value
                    Case "01"
                        If ShowClassNum.SelectedItem Is Nothing Then ShowClassNum.Items(1).Selected = True
                        ClassSort5.Enabled = False
                    Case "02"
                        If ShowClassNum.SelectedItem Is Nothing Then ShowClassNum.Items(2).Selected = True
                        ClassSort2.Enabled = False
                        ClassSort3.Enabled = False
                        ClassSort4.Enabled = False
                    Case Else
                        If ShowClassNum.SelectedItem Is Nothing Then ShowClassNum.Items(0).Selected = True
                End Select
            End If
        End If
    End Sub

    '班級選擇
    Private Sub OCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OCID.SelectedIndexChanged
        Call Utl_OCIDSel()
    End Sub

    '匯入單月排課作業
    Sub ImportExcels1()
        'Dim OCIDValue1 As String = v_OCID
        Dim STDateValue As String = STDate.Value
        Dim FTDateValue As String = FTDate.Value

        Dim CSDateText As String = CSDate.Text
        Dim CFDateText As String = CFDate.Text

        Dim Upload_Path As String = "~/SD/11/Temp/" '暫存路徑
        Call TIMS.MyCreateDir(Me, Upload_Path)

        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        '檢查檔案格式與大小 Start
        If File1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If

        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
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
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
        File1.PostedFile.SaveAs(filePath1) '上傳檔案
        Dim dt_xls As DataTable = TIMS.GetDataTable_XlsFile(filePath1, "", Errmag, "SchoolDate") '取得內容
        '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
        TIMS.MyFileDelete(filePath1)

        If Errmag <> "" Then
            Common.MessageBox(Me, Errmag)
            Common.MessageBox(Me, "資料有誤，故無法匯入，請修正Excel檔案，謝謝")
            Exit Sub
        End If
        If dt_xls Is Nothing Then '有資料
            Common.MessageBox(Me, "資料為空，故無法匯入，請修正Excel檔案，謝謝")
            Exit Sub
        End If
        If dt_xls.Rows.Count = 0 Then '有資料
            Common.MessageBox(Me, "查無資料，故無法匯入，請修正Excel檔案，謝謝")
            Exit Sub
        End If

        '將檔案讀出放入記憶體
        'Dim sr As System.IO.Stream
        'Dim srr As System.IO.StreamReader
        'sr = MyFile.OpenRead(Server.MapPath(Upload_Path & MyFileName))
        'srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)
        Dim RowIndex As Integer = 0 '讀取行累計數
        'Dim OneRow As String        'srr.ReadLine 一行一行的資料
        'Dim col As String           '欄位
        Dim colArray As Array       '陣列

        '取出資料庫的所有欄位 Start
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing

        '建立錯誤資料格式Table Start
        Dim Reason As String                    '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow

        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("SchoolDate"))
        dtWrong.Columns.Add(New DataColumn("OCID"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        '取出所有鍵值當判斷 Start
        'sql = "SELECT * FROM Key_Degree"
        'Key_Degree = DbAccess.GetDataTable(sql)
        'sql = "SELECT * FROM Key_Military"
        'Key_Military = DbAccess.GetDataTable(sql)
        'sql = "SELECT * FROM Key_Identity"
        'Key_Identity = DbAccess.GetDataTable(sql)
        '取出所有鍵值當判斷 End

        Reason = "" '做一次驗証的即可
        ' 驗証表格中的資料確認可以輸入使用 Start
        If OCIDValue1.Value = "" Then
            Reason += "未選擇 職類/班別(OCID) 無法匯入" & vbCrLf
        End If
        If Reason = "" Then
            sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ") AND TYPE=1 "
            Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
            If dt1.Rows.Count > 0 Then
                SysInfo.Text = "此班級已經使用全期排課" & "<br>"
                Reason += "此班級已經使用全期排課" & vbCrLf
            Else
                'sql = "SELECT * FROM CLASS_SCHEDULE   WHERE OCID IN (" & OCIDValue1.Value & ")"
                'sql += " AND SchoolDate >= '" & STDateValue & "' "
                'sql += " AND SchoolDate <= '" & FTDateValue & "' "
                'dt1 = DbAccess.GetDataTable(sql, objConn)
                'If dt1.Rows.Count > 0 Then Reason += "此班級已經建立單月排課" & vbCrLf
            End If
        End If
        If CSDateText = "" Or CFDateText = "" Then
            Reason += "排課區間為必填資料" & vbCrLf
        End If
        'If Reason = "" Then
        '    sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ")"
        '    sql &= " AND SchoolDate >= '" & CSDateText & "' "
        '    sql &= " AND SchoolDate <= '" & CFDateText & "' "
        '    Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objConn)
        '    If dt1.Rows.Count > 0 Then Reason += "此班級已經建立單月排課" & vbCrLf
        'End If
        ' 驗証表格中的資料確認可以輸入使用 End

        If Reason = "" Then
            Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                'Do While srr.Peek >= 0
                'OneRow = srr.ReadLine
                'If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈
                Reason = ""
                colArray = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                DetailTable.Style.Item("display") = TIMS.cst_inline1 '"inline" ' 每次驗証位置 Start
                'Dim j = i + 1
                Reason = CheckImportData(colArray, i + 1)  '驗証

                '取得最詳細的資料，寫入 ViewState(cst_vsDetailTable) = dtTemp
                If Reason = "" Then
                    vsErrMsg = ""
                    If CheckGetVal1(colArray(0), v_OCID, vsErrMsg) Then
                        CreateDetailCourse(colArray(0), dtTeacher)
                    Else
                        Reason = vsErrMsg
                        Common.MessageBox(Me, vsErrMsg)
                    End If
                End If

                '寫入暫存資料庫()
                If Reason = "" Then Reason = Insert_New_Class_Schedule(colArray) '改變  ViewState(cst_vsDetailTable) = dtTemp

                '存入資料庫()且重新計算
                If Reason = "" Then Reason = SAVE_CLASS_SCHEDULE("", OCIDValue1.Value, colArray(0))

                DetailTable.Style.Item("display") = "none" ' 每次驗証位置 End

                ' 寫入資料庫 Strat
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = "EXCEL 位置：第" & CStr(RowIndex + 2) & "列"
                    If colArray.Length > 5 Then
                        drWrong("SchoolDate") = colArray(0) '排課日期
                        drWrong("OCID") = OCIDValue1.Value '開班編號
                        drWrong("Reason") = Reason 'Reason
                    End If
                End If
                ' 寫入資料庫 End
                'If RowIndex <> 0 Then '第0行不進入
                'End If
                RowIndex += 1 '讀取行累計數
            Next
            'Loop
        End If

        '判斷匯出資料是否有誤
        Dim explain, explain2 As String
        'explain = ""
        'explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        'explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        'explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
        'explain2 = ""
        'explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        'explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        'explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"
        'RowIndex
        explain = ""
        explain += "匯入資料共" & RowIndex & "筆" & vbCrLf
        explain += "成功：" & (RowIndex - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
        explain2 = ""
        explain2 += "匯入資料共" & RowIndex & "筆\n"
        explain2 += "成功：" & (RowIndex - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"

        '開始判別欄位存入 End
        If dtWrong.Rows.Count = 0 Then
            If Reason = "" Then
                'Common.MessageBox(Me, "資料匯入成功")
                Common.MessageBox(Me, explain)
            Else
                Reason = "錯誤訊息如下:" & vbCrLf & Reason
                Common.MessageBox(Me, explain & Reason)
            End If

        Else
            Session("MyWrongTable") = dtWrong
            'Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_04_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('SD_04_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If

        'sr.Close()
        'srr.Close()
        'MyFile.Delete(Server.MapPath(Upload_Path & MyFileName)) '刪除暫存檔案
    End Sub

    '匯入單月排課作業
    Private Sub Button14_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button14.Click
        Call ImportExcels1()
    End Sub

    '匯入驗証 '課程代碼，只會驗證是否為數字
    Function CheckImportData(ByVal colArray As Array, Optional ByVal row As String = "") As String
        Dim Reason As String = ""
        Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        'Dim sql As String
        'Dim dr As DataRow
        'Const cst_Len as Integer = 49
        '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
        Const cst_Len As Integer = 61

        If colArray.Length < cst_Len Then
            'Reason += "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
        Else
            Try
                If IsDate(colArray(0)) Then
                    If DateDiff(DateInterval.Day, CDate(STDate.Value), CDate(colArray(0))) < 0 Then Reason += "日期輸入範圍有誤，尚未到達訓練期間<br>"
                    If DateDiff(DateInterval.Day, CDate(colArray(0)), CDate(FTDate.Value)) < 0 Then Reason += "日期輸入範圍有誤，已經超過訓練期間<br>"
                Else
                    Reason += "日期輸入格式有誤<br>"
                End If

                If IsDate(colArray(0)) Then
                    If DateDiff(DateInterval.Day, CDate(CSDate.Text), CDate(colArray(0))) < 0 Then Reason += "日期輸入範圍有誤，尚未到達排課期間<br>"
                    If DateDiff(DateInterval.Day, CDate(colArray(0)), CDate(CFDate.Text)) < 0 Then Reason += "日期輸入範圍有誤，已經超過排課期間<br>"
                End If

                '1~12(CLASS)
                For i As Integer = 1 To 12
                    If colArray(i).ToString <> "" Then
                        If Not IsNumeric(colArray(i).ToString) Then
                            Reason += "課堂 課程流水ID有誤，請使用 排課匯入用的課程代碼 欄位<br>"
                            Exit For
                        End If
                    End If
                Next

                Select Case TPeriodValue.Value
                    Case "01"
                        For i As Integer = 9 To 12
                            If colArray(i).ToString <> "" Then
                                Reason += "課堂 排課訓練時段有誤，已經超過可排課訓練時段<br>"
                                Exit Select
                            End If
                        Next
                    Case "02"
                        For i As Integer = 1 To 8
                            If colArray(i).ToString <> "" Then
                                Reason += "課堂 排課訓練時段有誤，已經超過可排課訓練時段<br>"
                                Exit Select
                            End If
                        Next
                    Case "03", "04"
                        'Case Else '"03","04"
                        '    Reason += "課堂 排課訓練時段有誤，只能為01~04<br>"
                        '    Exit Select
                End Select

                '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3) 61~72(Teacher4) 
                'Dim dt9 As DataTable = Nothing
                'Dim x As Integer = 0
                'Dim z As Integer = 0
                Dim sqlstr As String = ""
                sqlstr = " SELECT TECHID FROM TEACH_TEACHERINFO WHERE RID = '" & sm.UserInfo.RID & "' "
                Dim dt9 As New DataTable
                dt9.Load(DbAccess.GetReader(sqlstr, objconn))
                If dt9.Rows.Count = 0 Then Reason += "排課單位無輸入的師資代碼<br>"
                If dt9.Rows.Count > 0 Then
                    'dt9 = DbAccess.GetDataTable(sqlstr, objConn)
                    '13~24(Teacher1) 25~36(Teacher2) Teacher 1~24
                    For x As Integer = 13 To 36
                        Dim z As Integer = x - 12 '算出是欄位名稱 "Teacher 1~24" 字串的後面數字部分
                        If colArray(x).ToString <> "" Then
                            Try
                                If IsNumeric(colArray(x).ToString) = False Then
                                    Reason += "Teacher" & z & "的師資代碼不是合法的師資代碼,請輸入合法的師資代碼,須為數字!!<br>"
                                Else
                                    Dim iTechIDVal As Integer = 0
                                    iTechIDVal = CInt(colArray(x).ToString)
                                    If dt9.Select("TechID in ('" & Convert.ToString(iTechIDVal) & "')").Length = 0 Then Reason += "Teacher" & z & "的師資代碼不在本年度師資代碼設定中，請至師資設定功能做設定<br>"
                                End If
                            Catch ex As Exception
                                Reason += "Teacher" & z & "的師資代碼不是合法的師資代碼,請輸入合法的師資代碼,須為數字!!<br>"
                            End Try
                        End If
                    Next
                End If

                '37~48(Room)
                For x As Integer = 37 To 48
                    Dim z As Integer = x - 36 '算出是欄位名稱 "Room 1~12" 字串的後面數字部分
                    If colArray(x).ToString <> "" Then
                        If Len(colArray(x).ToString) > 30 Then Reason += "Room" & z & "的長度超過系統範圍(30),請重新輸入合法長度!!<br>"
                    End If
                Next

                If dt9.Rows.Count > 0 Then
                    '49~60(Teacher3) Teacher 25~36
                    For x As Integer = 49 To 60
                        Dim z As Integer = x - 24 '算出是欄位名稱 "Teacher 25~36" 字串的後面數字部分
                        If colArray(x).ToString <> "" Then
                            Try
                                If IsNumeric(colArray(x).ToString) = False Then
                                    Reason += "Teacher" & z & "的師資代碼不是合法的師資代碼,請輸入合法的師資代碼,須為數字!!<br>"
                                Else
                                    Dim iTechIDVal As Integer = 0
                                    iTechIDVal = CInt(colArray(x).ToString)
                                    If dt9.Select("TechID =" & Convert.ToString(iTechIDVal)).Length = 0 Then Reason += "Teacher" & z & "的師資代碼不在本年度師資代碼設定中，請至師資設定功能做設定<br>"
                                End If
                            Catch ex As Exception
                                Reason += "Teacher" & z & "的師資代碼不是合法的師資代碼,請輸入合法的師資代碼,須為數字!!<br>"
                            End Try
                        End If
                    Next
                End If
            Catch ex As Exception
                Reason += "欄位對應有誤<BR>"
                Reason += "請注意欄位中是否有半形逗點<BR>"
                'Exit For
            End Try
        End If

        Return Reason
    End Function

#Region "(No Use)"

    'Function Create_ColArray(ByRef colArray As System.Array, ByVal dr As DataRow, ByVal OCIDValue1 As String, ByVal diffday As Integer) As System.Array
    '    'diffday DateAdd(DateInterval.Day, diffday, CDate(dr("SchoolDate").ToString))
    '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
    '    'SELECT * FROM CLASS_SCHEDULE WHERE ROWNUM <=10
    '    dr("OCID") = OCIDValue1.ToString
    '    dr("SchoolDate") = DateAdd(DateInterval.Day, diffday, CDate(dr("SchoolDate").ToString))
    '    dr("ModifyAcct") = sm.UserInfo.UserID
    '    dr("ModifyDate") = Now
    '    Dim OneRow As String = Common.FormatDate(dr("SchoolDate")) '第1格為日期
    '    '0,1,2
    '    Dim sTeacher1 As String = ""
    '    Dim sTeacher2 As String = ""
    '    Dim sClass As String = ""
    '    Dim sRoom As String = ""
    '    Dim sTeacher3 As String = ""
    '    For i As Integer = 1 To 12
    '        Dim Tmps As String = ""
    '        Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i)))
    '        sTeacher1 &= cst_flag & Tmps
    '        Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 12)))
    '        sTeacher2 &= cst_flag & Tmps
    '        Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 24)))
    '        sTeacher3 &= cst_flag & Tmps
    '        Tmps = TIMS.ClearSQM(Convert.ToString(dr("Class" & i)))
    '        sClass &= cst_flag & Tmps
    '        Tmps = TIMS.ClearSQM(Convert.ToString(dr("Room" & i)))
    '        sRoom &= cst_flag & Tmps
    '    Next
    '    OneRow &= sTeacher1
    '    OneRow &= sTeacher2
    '    OneRow &= sClass
    '    OneRow &= sRoom
    '    OneRow &= sTeacher3
    '    'For i As Integer = 3 To CInt(3 + 48) - 1 'dr2.ItemArray.Length - 1
    '    '    OneRow += cst_flag & Convert.ToString(dr(i))
    '    'Next
    '    colArray = Split(OneRow, cst_flag)
    '    Return colArray
    'End Function

#End Region

    Function Create_ColArray(ByRef colArray As System.Array, ByVal dr As DataRow, ByVal OCIDValue1 As String, ByVal SchoolDate As Date) As System.Array
        '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
        'SELECT * FROM CLASS_SCHEDULE WHERE ROWNUM <=10
        'Dim OneRow As String = ""
        'Const cst_flag As String = ","
        dr("OCID") = OCIDValue1.ToString
        dr("SchoolDate") = Common.FormatDate(SchoolDate)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        Dim OneRow As String = Common.FormatDate(dr("SchoolDate"))
        Dim sTeacher1 As String = ""
        Dim sTeacher2 As String = ""
        Dim sClass As String = ""
        Dim sRoom As String = ""
        Dim sTeacher3 As String = ""
        Dim sTeacher4 As String = ""
        For i As Integer = 1 To 12
            '1~12(CLASS) 13~24(Teacher1) 25~36(Teacher2) 37~48(Room) 49~60(Teacher3)
            Dim Tmps As String = ""
            Tmps = TIMS.ClearSQM(Convert.ToString(dr("Class" & i)))
            sClass &= cst_flag & Tmps
            Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i)))
            sTeacher1 &= cst_flag & Tmps
            Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 12)))
            sTeacher2 &= cst_flag & Tmps
            Tmps = TIMS.ClearSQM(Convert.ToString(dr("Room" & i)))
            sRoom &= cst_flag & Tmps
            Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 24)))
            sTeacher3 &= cst_flag & Tmps
            Tmps = TIMS.ClearSQM(Convert.ToString(dr("Teacher" & i + 36)))
            sTeacher4 &= cst_flag & Tmps
        Next
        OneRow &= sClass
        OneRow &= sTeacher1
        OneRow &= sTeacher2
        OneRow &= sRoom
        OneRow &= sTeacher3
        OneRow &= sTeacher4
        'For i As Integer = 3 To CInt(3 + 48) - 1 'dr2.ItemArray.Length - 1
        '    OneRow += cst_flag & Convert.ToString(dr(i))
        'Next
        colArray = Split(OneRow, cst_flag)
        Return colArray
    End Function

#Region "(No Use)"

    'Public Sub LoadIntoClass2()
    '    Dim sql As String = ""
    '    Dim dr As DataRow = Nothing
    '    Dim dt As DataTable = Nothing
    '    'Dim sql, sql2 As String
    '    'Dim dr, dr2 As DataRow
    '    'Dim dt, dt2 As DataTable
    '    'Dim da As SqlDataAdapter = Nothing
    '    'Dim trans As SqlTransaction
    '    ''Dim conn As SqlConnection
    '    Dim colArray As Array       '陣列
    '    Dim i As Integer
    '    Dim diffday As Integer '差距天數

    '    '建立錯誤資料格式Table Start
    '    Dim Reason As String                    '儲存錯誤的原因
    '    Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
    '    Dim drWrong As DataRow

    '    dtWrong.Columns.Add(New DataColumn("Index"))
    '    dtWrong.Columns.Add(New DataColumn("SchoolDate"))
    '    dtWrong.Columns.Add(New DataColumn("OCID"))
    '    dtWrong.Columns.Add(New DataColumn("Reason"))
    '    '建立錯誤資料格式Table End

    '    ' 驗証表格中的資料確認可以輸入使用 Start
    '    If OCIDValue1.Value = "" Then
    '        Reason += "未選擇 職類/班別(OCID) 無法匯入" & vbCrLf
    '    End If
    '    If Reason = "" Then
    '        Dim dt1 As DataTable
    '        sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ") AND TYPE=1 "
    '        dt1 = DbAccess.GetDataTable(sql, objConn)
    '        If dt1.Rows.Count > 0 Then
    '            SysInfo.Text = "此班級已經使用全期排課" & "<br>"
    '            Reason += "此班級已經使用全期排課" & vbCrLf
    '        Else
    '            sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ")"
    '            sql &= " AND SchoolDate >= " & TIMS.to_date(STDate.Value)
    '            sql &= " AND SchoolDate <= " & TIMS.to_date(FTDate.Value)
    '            dt1 = DbAccess.GetDataTable(sql, objConn)
    '            If dt1.Rows.Count > 0 Then
    '                'Reason += "此班級已經建立單月排課" & vbCrLf
    '            End If
    '        End If
    '    End If
    '    If CSDate.Text = "" Or CFDate.Text = "" Then
    '        Reason += "排課區間為必填資料" & vbCrLf
    '    End If
    '    If Reason = "" Then
    '        sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ")"
    '        sql &= " AND SchoolDate >= " & TIMS.to_date(CSDate.Text)
    '        sql &= " AND SchoolDate <= " & TIMS.to_date(CFDate.Text)
    '        Dim dt1 As DataTable
    '        dt1 = DbAccess.GetDataTable(sql, objConn)
    '        If dt1.Rows.Count > 0 Then
    '            'Reason += "此班級已經建立單月排課" & vbCrLf
    '        End If
    '    End If
    '    ' 驗証表格中的資料確認可以輸入使用 End

    '    sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & OCIDValue2.Value & "' AND TYPE ='2' AND  Formal='Y' order by schoolDate "
    '    dt = DbAccess.GetDataTable(sql, objConn)
    '    dr = dt.Rows(0)
    '    diffday = CInt(DateDiff(DateInterval.Day, CDate(dr("schoolDate")), CDate(STDate.Value)))

    '    If Reason = "" Then
    '        DetailTable.Style.Item("display") = "inline" ' 每次驗証位置 Start
    '        For i = 0 To dt.Rows.Count - 1
    '            dr = dt.Rows(i)
    '            Reason = "" '做一次驗証的即可

    '            Call Create_ColArray(colArray, dr, OCIDValue1.Value, diffday)

    '            'DetailTable.Style.Item("display") = "inline" ' 每次驗証位置 Start
    '            Reason = CheckImportData(colArray) '驗証
    '            '取得最詳細的資料，寫入 ViewState(cst_vsDetailTable) = dtTemp

    '            If Reason = "" Then
    '                vsErrMsg = ""
    '                If CheckGetVal1(colArray(0), v_OCID, vsErrMsg) Then
    '                    CreateDetailCourse(colArray(0), dtTeacher)
    '                Else
    '                    Common.MessageBox(Me, vsErrMsg)
    '                End If
    '            End If

    '            '寫入暫存資料庫()
    '            If Reason = "" Then Reason = Insert_New_Class_Schedule(colArray) '改變  ViewState(cst_vsDetailTable) = dtTemp
    '            '存入資料庫()且重新計算
    '            If Reason = "" Then Reason = Save_Class_Schedule("", OCIDValue1.Value, colArray(0))
    '            'DetailTable.Style.Item("display") = "none" ' 每次驗証位置 End

    '            ' 寫入資料庫 Strat
    '            If Reason <> "" Then
    '                '錯誤資料，填入錯誤資料表
    '                drWrong = dtWrong.NewRow
    '                dtWrong.Rows.Add(drWrong)

    '                drWrong("Index") = CInt(i + 1)
    '                If colArray.Length > 5 Then
    '                    drWrong("SchoolDate") = colArray(0) '排課日期
    '                    drWrong("OCID") = OCIDValue1.Value '開班編號
    '                    drWrong("Reason") = Reason 'Reason
    '                End If
    '            End If
    '            ' 寫入資料庫 End
    '        Next
    '        DetailTable.Style.Item("display") = "none" ' 每次驗証位置 End
    '    End If

    '    '開始判別欄位存入 End
    '    If dtWrong.Rows.Count = 0 Then
    '        If Reason = "" Then
    '            Common.MessageBox(Me, "資料匯入成功")
    '        Else
    '            Common.MessageBox(Me, Reason)
    '        End If
    '    Else
    '        Session("MyWrongTable") = dtWrong
    '        Page.RegisterStartupScript("js2", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_04_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
    '    End If

    'End Sub

#End Region

    '判斷是否有上課資料
    Function Check_onClass(ByVal dr As DataRow) As Boolean
        Dim Flag As Boolean = False

        For i12 As Integer = 1 To 12
            If Not IsDBNull(dr("Class" & i12)) Then
                Flag = True
                Exit For
            End If
        Next

        Return Flag
    End Function

    '驗證當日是否為每星期的上課日
    Function Check_EffectiveClassDay(ByVal strDate As String, ByVal Gary_day7 As String()) As Boolean
        'Gary_day7 為已經取得原資料(每星期的上課日)系統共用變數
        Dim match_flag As Boolean = False
        Dim DateTemp As Date
        DateTemp = Common.FormatDate(strDate)

        If strDate <> "" And Not Gary_day7 Is Nothing Then
            match_flag = False
            For i7 As Integer = 0 To UBound(Gary_day7)
                If Gary_day7(i7) = DateTemp.DayOfWeek() Then
                    match_flag = True
                    Exit For
                End If
            Next
        End If

        Return match_flag
    End Function

    Function Get_WorkDay7(ByVal OCID As String) As String()
        'Dim Gary_day7() As String '第一個星期的上課時間 星期日~星期六 0,1,..,6 (系統共用變數)
        Dim Rst() As String = Nothing
        Dim sql_7 As String
        sql_7 = "SELECT TOP 7 * FROM CLASS_SCHEDULE WHERE 1=1 AND OCID=@OCID  AND TYPE ='2' AND FORMAL='Y' ORDER BY SCHOOLDATE" & vbCrLf
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("OCID", OCID)
        Dim dt As DataTable = DbAccess.GetDataTable(sql_7, objconn, parms)
        'dt.Load(DbAccess.GetReader(sql_7, objconn, parms))

        Dim str As String = ""
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim dr As DataRow = dt.Rows(i)
            Dim Flag As Boolean = Check_onClass(dr) '判斷是否有上課資料
            If Flag AndAlso dr("SchoolDate").ToString <> "" Then
                If str <> "" Then str &= ","
                str &= CInt(CDate(Common.FormatDate(dr("SchoolDate"))).DayOfWeek).ToString
            End If
        Next

        If str <> "" Then Rst = str.Split(",")
        Return Rst
    End Function

    '將資料轉換為Null空白行
    Function Get_DrNull(ByVal dtTempB As DataTable, ByVal dr As DataRow) As DataRow
        Dim drB As DataRow = dtTempB.NewRow
        dtTempB.Rows.Add(drB)

        For Each aDataColumn As DataColumn In dtTempB.Columns
            drB(aDataColumn.ColumnName) = dr(aDataColumn.ColumnName)
        Next

        drB("OCID") = OCIDValue1.Value ' OCIDValue1.ToString
        'dr("SchoolDate") = DateAdd(DateInterval.Day, diffday, CDate(dr("SchoolDate").ToString))
        drB("ModifyAcct") = sm.UserInfo.UserID
        drB("ModifyDate") = Now
        For i As Integer = 3 To CInt(3 + 48) - 1 'dr2.ItemArray.Length - 1
            drB(i) = Convert.DBNull
        Next

        Return drB
    End Function

    '匯入功能
    Public Sub LoadIntoClass3()
        Dim colArray As Array = Nothing '陣列
        'Dim i As Integer
        'Dim i2, i3 As Integer  '記錄資料使用筆數 i3:總筆數
        'Dim diffday As Integer '差距天數
        Dim sql As String = ""

        '建立錯誤資料格式Table Start
        Dim Reason As String = ""                '儲存錯誤的原因
        Dim dtWrong As New DataTable             '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("SchoolDate"))
        dtWrong.Columns.Add(New DataColumn("OCID"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        ' 驗証表格中的資料確認可以輸入使用 Start
        If OCIDValue1.Value = "" Then Reason += "未選擇 職類/班別(OCID) 無法匯入" & vbCrLf

        If Reason = "" Then
            sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ") AND TYPE=1 "
            Dim dt1 As DataTable
            dt1 = DbAccess.GetDataTable(sql, objconn)

            If dt1.Rows.Count > 0 Then
                SysInfo.Text = "此班級已經使用全期排課" & "<br>"
                Reason += "此班級已經使用全期排課" & vbCrLf
            Else
                sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ")"
                sql &= " AND SchoolDate >= " & TIMS.To_date(STDate.Value)
                sql &= " AND SchoolDate <= " & TIMS.To_date(FTDate.Value)
                dt1 = DbAccess.GetDataTable(sql, objconn)

                If dt1.Rows.Count > 0 Then
                    'Reason += "此班級已經建立單月排課" & vbCrLf
                End If
            End If
        End If

        If CSDate.Text = "" Or CFDate.Text = "" Then Reason += "排課區間為必填資料" & vbCrLf

        If Reason = "" Then
            sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID IN (" & OCIDValue1.Value & ")"
            sql &= " AND SchoolDate >= " & TIMS.To_date(CSDate.Text)
            sql &= " AND SchoolDate <= " & TIMS.To_date(CFDate.Text)
            Dim dt1 As DataTable
            dt1 = DbAccess.GetDataTable(sql, objconn)
            'If dt1.Rows.Count > 0 Then
            '    Reason += "此班級已經建立單月排課" & vbCrLf
            'End If
        End If
        'If Reason = "" Then
        'End If
        ' 驗証表格中的資料確認可以輸入使用 End

        sql = " SELECT * FROM CLASS_SCHEDULE WHERE  1<>1 "
        Dim dtTempB As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim drBNull As DataRow = Nothing

        sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & OCIDValue2.Value & "' AND TYPE = '2' AND Formal = 'Y' ORDER BY schoolDate "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        'dr = dt.Rows(0)
        'i2 = 0
        If dt.Rows.Count > 0 Then
            drBNull = Get_DrNull(dt, dt.Rows(0))  'COPY一筆含有資料的內容轉換成無內容的資料
            For ix As Integer = 0 To dt.Rows.Count - 1 '欲載入的班級迴圈
                Dim dr As DataRow = dt.Rows(ix)
                If Check_onClass(dr) Then '判斷是否有課程
                    'ViewState("dtTempB_" & i2) = dr
                    'i2 += 1
                    dtTempB.ImportRow(dr)
                    'drB = dtTempB.NewRow
                    'drB.ItemArray = dr.ItemArray
                    'dtTempB.Rows.Add(drB)
                    'For Each aDataColumn As DataColumn In dtTempB.Columns
                    '    drB(aDataColumn.ColumnName) = dr(aDataColumn.ColumnName)
                    'Next
                End If
            Next
            'i3 = i2
        End If

        Dim Gary_day7() As String '每星期的上課時間 星期日~星期六 0~6
        Gary_day7 = Get_WorkDay7(OCIDValue2.Value)

        'Dim CntTB As Integer = dtTempB.Rows.Count
        Dim indexDate As Date = CDate(STDate.Value)
        DetailTable.Style.Item("display") = TIMS.cst_inline1 '"inline"  ' 每次驗証位置 Start
        Dim i As Integer = 0
        Dim i2 As Integer = 0

        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        Do
            If Check_EffectiveClassDay(indexDate, Gary_day7) Then
                '有效上課日
                'dr = ViewState("dtTempB_" & i2)  'dtTempB.Rows(i2)
                Dim dr As DataRow = dtTempB.Rows(i2) 'CLASS_SCHEDULE

                Call Create_ColArray(colArray, dr, OCIDValue1.Value, indexDate)

                Reason = CheckImportData(colArray) '驗証
                '取得最詳細的資料，寫入 ViewState(cst_vsDetailTable) = dtTemp

                If Reason = "" Then
                    vsErrMsg = ""
                    If CheckGetVal1(colArray(0), v_OCID, vsErrMsg) Then
                        CreateDetailCourse(colArray(0), dtTeacher)
                    Else
                        Common.MessageBox(Me, vsErrMsg)
                    End If
                End If

                '寫入暫存資料庫()
                If Reason = "" Then Reason = Insert_New_Class_Schedule(colArray) '改變  ViewState(cst_vsDetailTable) = dtTemp

                '存入資料庫()且重新計算
                If Reason = "" Then Reason = Save_Class_Schedule3("", OCIDValue1.Value, colArray(0))

                ' 寫入資料庫 Strat
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = CInt(i + 1)
                    If colArray.Length > 5 Then
                        drWrong("SchoolDate") = colArray(0) '排課日期
                        drWrong("OCID") = OCIDValue1.Value '開班編號
                        drWrong("Reason") = Reason 'Reason
                    End If
                End If
                ' 寫入資料庫 End

                i2 += 1
            Else
                '非上課日
                Call Create_ColArray(colArray, drBNull, OCIDValue1.Value, indexDate)

                Reason = CheckImportData(colArray) '驗証
                '取得最詳細的資料，寫入 ViewState(cst_vsDetailTable) = dtTemp

                If Reason = "" Then
                    vsErrMsg = ""
                    If CheckGetVal1(colArray(0), v_OCID, vsErrMsg) Then
                        CreateDetailCourse(colArray(0), dtTeacher)
                    Else
                        Common.MessageBox(Me, vsErrMsg)
                    End If
                End If

                '寫入暫存資料庫()
                If Reason = "" Then Reason = Insert_New_Class_Schedule(colArray) '改變  ViewState(cst_vsDetailTable) = dtTemp

                '存入資料庫()且重新計算
                If Reason = "" Then Reason = Save_Class_Schedule3("", OCIDValue1.Value, colArray(0))

                ' 寫入資料庫 Strat
                'Try
                'Catch ex As Exception
                'End Try
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = CInt(i + 1)
                    If colArray.Length > 5 Then
                        drWrong("SchoolDate") = colArray(0) '排課日期
                        drWrong("OCID") = OCIDValue1.Value '開班編號
                        drWrong("Reason") = Reason 'Reason
                    End If

                End If
                ' 寫入資料庫 End
            End If

            indexDate = DateAdd(DateInterval.Day, 1, indexDate)
        Loop Until (DateDiff(DateInterval.Day, indexDate, CDate(FTDate.Value)) <= -1)
        DetailTable.Style.Item("display") = "none"  '每次驗証位置 End

        '開始判別欄位存入 End
        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, If(Reason <> "", Reason, "資料匯入成功"))
        Else
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("js2", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_04_002_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
    End Sub

    '載入排課班級
    Private Sub LoadIntoClass_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles LoadIntoClass.Click
        'Dim Sign As String = "Y"
        'Dim sql, ERRMsg, strScript As String
        'Dim dr As DataRow
        'Dim dt As DataTable
        OCIDValue2.Value = TIMS.ClearSQM(OCIDValue2.Value)
        Dim ERRMsg As String = ""
        Dim sql As String = ""
        Dim dt As New DataTable
        If OCIDValue2.Value <> "" Then
            sql = " SELECT OCID, COUNT(1) DayCount FROM CLASS_SCHEDULE WHERE OCID = '" & OCIDValue2.Value & "' AND TYPE = '2' AND Formal = 'Y' GROUP BY OCID "
            dt.Load(DbAccess.GetReader(sql, objconn))
        End If
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt.Rows(0)
            If CInt(dr1("DayCount").ToString) > 0 Then '先判斷是否有可用天數
                Dim Gary_day7() As String '每星期的上課時間 星期日~星期六 0~6
                Gary_day7 = Get_WorkDay7(OCIDValue2.Value)
                If Gary_day7 Is Nothing Then ERRMsg += "欲載入的班級第一週無排課資料，無法對應每星期上課時間" & vbCrLf
            End If
        Else
            ERRMsg += "欲載入的班級尚無單月排課資料" & vbCrLf
        End If

        If ERRMsg <> "" Then
            ERRMsg += "請確認" & vbCrLf
            Common.MessageBox(Me, ERRMsg)
            Exit Sub
        End If

        Dim dr As DataRow = dt.Rows(0)
        'dr = dt.Rows(0)
        If CInt(dr("DayCount").ToString) <= CInt(Me.DayCount.Value) Then
            '可執行匯入
        Else
            Message.Text = "載入排課天數 " & dr("DayCount").ToString & " 大於可排課天數 " & Me.DayCount.Value.ToString & " 系統將依本班天數載入班級資料"
            'ERRMsg += "欲載入的班級排課天數 " & dr("DayCount").ToString & "\n"
            'ERRMsg += "    大於本班可排天數 " & Me.DayCount.Value.ToString & "\n"
            'ERRMsg += "確定要繼續轉入作業 ?"
            'strScript = "<script language=""javascript"">" + vbCrLf
            'strScript += " if(window.confirm('" & ERRMsg & "')){" + vbCrLf
            'strScript += "alert('系統將依本班天數載入班級資料.');} else {return false;}"
            'strScript += "</script>"
            'Page.RegisterStartupScript("", strScript)
        End If
        'LoadIntoClass2() '執行匯入 依日
        'LoadIntoClass3() '執行匯入 依每星期

        'Dim myThreadDelegate As New ThreadStart(AddressOf LoadIntoClass2)
        Dim myThreadDelegate As New ThreadStart(AddressOf LoadIntoClass3)
        Dim myThread As New Thread(myThreadDelegate)
        myThread.Start()
        'Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在課程，請稍候10~20分，再回到系統來查詢!!');</script>")
        'Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在排課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & Me.OCIDValue1.Value & "&Single=" & Sign & "&Formal=Y'</script>")
        Page.RegisterStartupScript("JS1", "<script>alert('系統後端正在排課程，如果稍後看到的課程列表是空白的，請稍候10~20分，再回到系統來查詢!!');location.href='SD_04_003.aspx?ID=" & Request("ID") & "&ClassID=" & Me.OCIDValue1.Value & "&Single=Y&Formal=Y'</script>")
    End Sub

    Sub GetSearchStr()
        THours.Text = Trim(THours.Text)
        LeftHour.Text = Trim(LeftHour.Text)
        If THours.Text = "" Then THours.Text = "0"
        If LeftHour.Text = "" Then LeftHour.Text = "0"
        Dim sUseHour As String = Convert.ToString(CInt(THours.Text) - CInt(LeftHour.Text))
        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        Dim sSearchStr As String = ""
        Call TIMS.SetMyValue(sSearchStr, "k", "1")
        Call TIMS.SetMyValue(sSearchStr, "center", center.Text)
        Call TIMS.SetMyValue(sSearchStr, "RIDValue", RIDValue.Value)
        Call TIMS.SetMyValue(sSearchStr, "OCID1", OCID.Items.FindByValue(v_OCID).Text)
        Call TIMS.SetMyValue(sSearchStr, "OCIDValue1", OCIDValue1.Value)
        If CSDate.Text <> "" Then
            CSDate.Text = TIMS.Cdate3(CSDate.Text)
            If DateDiff(DateInterval.Day, CDate(STDate.Value), CDate(CSDate.Text)) >= 0 Then
                Call TIMS.SetMyValue(sSearchStr, "start_date", CSDate.Text) '使用者輸入起始值
            Else
                Call TIMS.SetMyValue(sSearchStr, "start_date", STDate.Value) '原班起始值
            End If
        Else
            STDate.Value = TIMS.Cdate3(STDate.Value)
            Call TIMS.SetMyValue(sSearchStr, "start_date", STDate.Value) '原班起始值
        End If
        If CFDate.Text <> "" Then
            CFDate.Text = TIMS.Cdate3(CFDate.Text)
            If DateDiff(DateInterval.Day, CDate(CFDate.Text), CDate(FTDate.Value)) >= 0 Then
                Call TIMS.SetMyValue(sSearchStr, "end_date", CFDate.Text) '使用者輸入結束值
            Else
                Call TIMS.SetMyValue(sSearchStr, "end_date", FTDate.Value) '原班結束值
            End If
        Else
            FTDate.Value = TIMS.Cdate3(FTDate.Value)
            Call TIMS.SetMyValue(sSearchStr, "end_date", FTDate.Value) '原班結束值
        End If
        Call TIMS.SetMyValue(sSearchStr, "TPeriod", TPeriod.Text)
        Call TIMS.SetMyValue(sSearchStr, "TPeriodValue", TPeriodValue.Value)
        Call TIMS.SetMyValue(sSearchStr, "THours", THours.Text)
        Call TIMS.SetMyValue(sSearchStr, "LeftHour", LeftHour.Text)
        Call TIMS.SetMyValue(sSearchStr, "UseHour", sUseHour)
        Session("SearchStr") = sSearchStr
    End Sub

    '新增排課資料
    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        Dim Reason As String = ""
        Reason = ""

        Dim v_OCID As String = TIMS.GetListValue(OCID) ' OCID.SelectedValue
        If v_OCID <> "" Then OCIDValue1.Value = v_OCID 'OCID.SelectedValue
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then Reason += "未選擇 職類/班別(OCID) 無法進入排課" & vbCrLf
        If Reason <> "" Then
            Reason = "錯誤訊息如下:" & vbCrLf & Reason
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        'Dim TEST As String = ""
        'TEST = Microsoft.JScript.GlobalObject.escape("測試打")
        Call GetSearchStr()
        Call TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "SD_04_002_add.aspx?ID=" & Request("ID"))
    End Sub

    '匯出
    Private Sub Button1_ServerClick(sender As Object, e As System.EventArgs) Handles Button1.ServerClick
        Call sUtl_Search1() 'sender, e
    End Sub

    ''' <summary> '匯出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExport_Click(sender As Object, e As EventArgs) Handles BtnExport.Click
        OCIDValue1.Value = TIMS.ClearSQM(TIMS.GetListValue(OCID))
        vsErrMsg = ""
        If OCIDValue1.Value = "" Then vsErrMsg &= "未選擇 職類/班別 無法匯出" & vbCrLf
        If vsErrMsg <> "" Then
            msg.Text = vsErrMsg
            Common.MessageBox(Me, vsErrMsg)
            Return
        End If
        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別！")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別！")
            Return
        End If

        '匯出 xlsx
        Call ExpRptXLSX()
    End Sub

    ''' <summary> '匯出 xlsx</summary>
    Private Sub ExpRptXLSX()
        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別！")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別！")
            Return
        End If

        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'If OCIDValue1.Value = "" Then Return
        Try
            'Call ExpRpt() '匯出SUB'SQL
            Call ExpXLSX1()
        Catch ex As System.Threading.ThreadAbortException
            TIMS.LOG.Warn(ex.Message, ex)
            Server.ClearError()
        Catch ex As Exception
            Dim sErrMsg1 As String = String.Concat("發生錯誤:", vbCrLf, ex.ToString, vbCrLf, "g_ErrSql : ", vbCrLf, g_ErrSql, vbCrLf, "OCIDValue1.Value : ", OCIDValue1.Value)
            Call TIMS.WriteTraceLog(Page, ex, sErrMsg1)
            Common.MessageBox(Me.Page, "發生錯誤:" & vbCrLf & ex.Message)
            Call TIMS.CloseDbConn(objconn)
            'If Response IsNot Nothing AndAlso (Response.IsClientConnected) Then Response.End()
            'TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Return
        End Try

        '結束狀況無誤
        If msg.Text = "" Then TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    ''' <summary>取得 DataTable 匯入xlsx </summary>
    Sub ExpXLSX1()
        Dim dt1 As New DataTable
        Call Utl_ExpXLSXdt1(dt1)
        dt1.TableName = "課表"

        vsErrMsg = ""
        Dim flag_DATAOK1 As Boolean = If(dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0, True, False)
        If (Not flag_DATAOK1) Then vsErrMsg &= "該班級資料筆數有誤 無法匯出" & vbCrLf
        If vsErrMsg <> "" Then
            msg.Text = vsErrMsg
            Common.MessageBox(Me, vsErrMsg)
            Return
        End If

        Dim dt2 As New DataTable
        Call Utl_ExpXLSXdt2(dt2)
        dt2.TableName = "常用老師以及舊資料對應"

        Dim dt3 As New DataTable
        Call Utl_ExpXLSXdt3(dt3)
        dt3.TableName = "常用教室"

        Dim dt4 As New DataTable
        Call Utl_ExpXLSXdt4(dt4)
        dt4.TableName = "常用課程以及單位課程"

        Dim dt5 As New DataTable
        Call Utl_ExpXLSXdt5(dt5)
        dt5.TableName = "作息時間"

        Dim ds1 As New DataSet
        ds1.Tables.Add(dt1)
        ds1.Tables.Add(dt2)
        ds1.Tables.Add(dt3)
        ds1.Tables.Add(dt4)
        ds1.Tables.Add(dt5)

        Call TIMS.CloseDbConn(objconn)
        Call TIMS.Get_XLSX_Response(Me, ds1) ', Cst_FileSavePath
        If Response IsNot Nothing AndAlso (Response.IsClientConnected) Then Response.End()
        Return
    End Sub

    ''' <summary> 作息時間-識別碼,內容</summary>
    ''' <param name="dt"></param>
    Private Sub Utl_ExpXLSXdt5(ByRef dt As DataTable)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS ( select cr.CRTID" & vbCrLf
        sql &= " ,CONCAT(cr.C1,dbo.FN_GETCRLF(1),cr.C2,dbo.FN_GETCRLF(1),cr.C3,dbo.FN_GETCRLF(1),cr.C4,dbo.FN_GETCRLF(1)" & vbCrLf
        sql &= " ,cr.C5,dbo.FN_GETCRLF(1),cr.C6,dbo.FN_GETCRLF(1),cr.C7,dbo.FN_GETCRLF(1),cr.C8,dbo.FN_GETCRLF(1)" & vbCrLf
        sql &= " ,cr.C9,dbo.FN_GETCRLF(1),cr.C10,dbo.FN_GETCRLF(1),cr.C11,dbo.FN_GETCRLF(1),cr.C12,dbo.FN_GETCRLF(1)) CTEXT" & vbCrLf
        sql &= " from V_CLASSRESTTIME cr" & vbCrLf
        sql &= " WHERE cr.RID =@RID )" & vbCrLf
        sql &= " SELECT DISTINCT CRTID 識別碼,CTEXT 內容 FROM WC1" & vbCrLf
        g_ErrSql = sql
        Dim sCmd As New SqlCommand(g_ErrSql, objconn)
        sCmd.CommandTimeout = cst_max_Timeout '500
        'Dim dt As New DataTable
        If dt Is Nothing Then dt = New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            dt.Load(.ExecuteReader())
        End With
    End Sub

    ''' <summary>常用課程以及單位課程-課程代碼</summary>
    ''' <param name="dt"></param>
    Private Sub Utl_ExpXLSXdt4(ByRef dt As DataTable)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT cr.COURID 代碼	" & vbCrLf
        sql &= " ,cr.COURSENAME 課程名稱" & vbCrLf
        sql &= " ,null 類別" & vbCrLf
        sql &= " FROM COURSE_COURSEINFO cr" & vbCrLf
        sql &= " WHERE cr.RID=@RID" & vbCrLf
        g_ErrSql = sql
        Dim sCmd As New SqlCommand(g_ErrSql, objconn)
        sCmd.CommandTimeout = cst_max_Timeout '500
        'Dim dt As New DataTable
        If dt Is Nothing Then dt = New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            dt.Load(.ExecuteReader())
        End With
    End Sub

    ''' <summary>常用教室</summary>
    ''' <param name="dt"></param>
    Private Sub Utl_ExpXLSXdt3(ByRef dt As DataTable)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select ch.ROOM" & vbCrLf
        sql &= " FROM dbo.VIEW_CLASS_SCHEDULE ch" & vbCrLf
        sql &= " WHERE ch.OCID=@OCIDV1" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT DISTINCT ROOM 教室名稱" & vbCrLf
        sql &= " FROM WC1" & vbCrLf
        g_ErrSql = sql
        Dim sCmd As New SqlCommand(g_ErrSql, objconn)
        sCmd.CommandTimeout = cst_max_Timeout '500
        'Dim dt As New DataTable
        If dt Is Nothing Then dt = New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCIDV1", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            dt.Load(.ExecuteReader())
        End With
    End Sub

    ''' <summary>常用老師以及舊資料對應</summary>
    ''' <param name="dt"></param>
    Sub Utl_ExpXLSXdt2(ByRef dt As DataTable)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT tt.TECHID 代碼" & vbCrLf
        sql &= " ,tt.TEACHCNAME 老師姓名" & vbCrLf
        sql &= " ,case when tt.TECHTYPE3='Y' THEN '教師' when tt.TECHTYPE4='Y' THEN '第二教師'" & vbCrLf
        sql &= " when tt.TECHTYPE1='Y' THEN '講師' when tt.TECHTYPE2='Y' THEN '助教' END 類別" & vbCrLf
        sql &= " FROM TEACH_TEACHERINFO tt" & vbCrLf
        sql &= " WHERE tt.RID=@RID" & vbCrLf
        g_ErrSql = sql
        Dim sCmd As New SqlCommand(g_ErrSql, objconn)
        sCmd.CommandTimeout = cst_max_Timeout '500
        'Dim dt As New DataTable
        If dt Is Nothing Then dt = New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = sm.UserInfo.RID
            dt.Load(.ExecuteReader())
        End With
    End Sub

    ''' <summary>課表-資料</summary>
    ''' <param name="dt"></param>
    Sub Utl_ExpXLSXdt1(ByRef dt As DataTable)
        ''Call TIMS.OpenDbConn(objconn)
        'Const cst_max_Timeout As Integer = 500
        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別！")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "班別代碼有誤，請確認點選職類/班別！")
            Return
        End If

        'Dim parmas As New Hashtable
        'parmas.Add("OCIDV1", OCIDValue1.Value)
        Dim sql As String = ""
        sql &= " SELECT dbo.FN_CDATE1B(ch.SCHOOLDATE) 日期" & vbCrLf
        sql &= " ,dbo.[FN_RESTTIMECRTID](ch.OCID) ""作息識別碼""" & vbCrLf
        sql &= " ,ch.LESSON 節次" & vbCrLf
        sql &= " ,dbo.FN_RESTTIME(ch.OCID,ch.LESSON) 時間" & vbCrLf
        'sql &= " ,ch.COURSEID 課程代碼" & vbCrLf
        sql &= " ,ch.CLASS 課程代碼" & vbCrLf
        sql &= " ,ch.ROOM 教室" & vbCrLf
        sql &= " ,[dbo].[FN_TECHID34](ch.OCID,CH.SCHOOLDATE,CH.LESSON,1) 教師1代碼" & vbCrLf
        sql &= " ,[dbo].[FN_TECHID34](ch.OCID,CH.SCHOOLDATE,CH.LESSON,2) 教師2代碼" & vbCrLf
        sql &= " ,[dbo].[FN_TECHID34](ch.OCID,CH.SCHOOLDATE,CH.LESSON,3) 教師3代碼" & vbCrLf
        sql &= " ,[dbo].[FN_TECHID34](ch.OCID,CH.SCHOOLDATE,CH.LESSON,4) 助教代碼" & vbCrLf
        sql &= " FROM dbo.VIEW_CLASS_SCHEDULE ch" & vbCrLf
        sql &= " WHERE ch.OCID=@OCIDV1" & vbCrLf
        sql &= " ORDER BY ch.OCID,ch.SCHOOLDATE,ch.LESSON" & vbCrLf
        g_ErrSql = sql
        Dim sCmd As New SqlCommand(g_ErrSql, objconn)
        sCmd.CommandTimeout = cst_max_Timeout '500
        'Dim dt As New DataTable
        If dt Is Nothing Then dt = New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCIDV1", SqlDbType.Int).Value = Val(OCIDValue1.Value)
            dt.Load(.ExecuteReader())
        End With
    End Sub

#Region "(No Use)"

    ''檢查 CLASS_SCHEDULE 是否重複
    'Function Check_Class_Schedule_Duplicate(ByVal ItemValue As String, ByRef ErrorMsg As String) As Boolean
    '    Dim rst As Boolean = True '重複
    '    ErrorMsg = ""
    '    '1.教師：同一日同個時段，同一教師不可上同一堂課(無法成立)
    '    '2.教室：同一日同個時段不可使用同一教室(無法成立)
    '    '同一日 同個時段 同一教師 不可在不同教室出現
    '    Dim dr As DataRow
    '    Dim sql As String
    '    Dim strarray As String()
    '    'Dim i As Int32
    '    strarray = Split(ItemValue, ",")
    '    sql = " select SUM(qty) cnt FROM (select 0 as qty "
    '    For i As Integer = 0 To UBound(strarray, 1) - 1
    '        If OLessonTeah1Value.Value <> "" Then
    '            sql &= " union SELECT count(1) qty FROM CLASS_SCHEDULE WHERE SCHOOLDATE=convert(datetime, '" & MyDate.Text & "', 111) and Room" & strarray(i) & "<>'" & Room.Text & "' and (Teacher" & strarray(i) & "='" & OLessonTeah1Value.Value & "' or Teacher" & strarray(i) + 12 & "='" & OLessonTeah1Value.Value & "') " & vbCrLf
    '        End If
    '        If OLessonTeah2Value.Value <> "" Then
    '            sql &= " union SELECT count(1) qty FROM CLASS_SCHEDULE WHERE SCHOOLDATE=convert(datetime, '" & MyDate.Text & "', 111) and Room" & strarray(i) & "<>'" & Room.Text & "' and (Teacher" & strarray(i) & "='" & OLessonTeah2Value.Value & "' or Teacher" & strarray(i) + 12 & "='" & OLessonTeah2Value.Value & "') " & vbCrLf
    '        End If
    '    Next
    '    sql &= " ) a "
    '    dr = DbAccess.GetOneRow(sql, objConn)

    '    If dr("cnt") > 0 Then
    '        ErrorMsg = MyDate.Text & " 當日 同個時段 同一教師 不可在不同教室出現!!"
    '        rst = True '重複
    '    Else
    '        rst = False '不重複
    '    End If
    '    Return rst
    'End Function

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    'End Sub

#End Region

End Class