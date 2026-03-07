'Imports Microsoft.Web.UI.WebControls

Partial Class SD_04_002_add
    Inherits AuthBasePage

    'TEACH_TEACHERINFO,COURSE_COURSEINFO,SYS_HOLIDAY,STUD_TRAININGRESULTS,CLASS_SCHVERIFY
    'CLASS_SCHEDULE,PLAN_SCHEDULE 
    'V_SCHEDULETYPE
    'SELECT * FROM CLASS_SCHEDULE WHERE ROWNUM <=10
    Dim ff3 As String = ""
    Private dtCourse As DataTable
    Private HolidayTable As DataTable
    Private dtTeacher As DataTable
    Const cst_holiday As String = "holiday"

    'Dim PageControler1 As New PageControler
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'Dim sql As String = ""
        'If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        'sql = "SELECT CourID,CourseName,MainCourID FROM Course_CourseInfo WHERE RID LIKE '" & Mid(RIDValue.Value, 1, 1) & "%'"
        'dtCourse = New DataTable
        'dtCourse.Load(DbAccess.GetReader(sql, objconn))
        Dim ssRID As String = sm.UserInfo.RID
        If RIDValue.Value <> "" Then ssRID = RIDValue.Value
        dtCourse = TIMS.Get_COURSEINFOdt(ssRID, objconn)

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim sql As String = ""
        sql = " SELECT * FROM SYS_HOLIDAY WHERE RID = '" & RIDValue.Value & "' "
        'HolidayTable = New DataTable
        'HolidayTable.Load(DbAccess.GetReader(sql, objconn))
        HolidayTable = DbAccess.GetDataTable(sql, objconn)

        If Not IsPostBack Then
            Call Utl_ShowX1()

            OCIDValue1.Value = Show_SearchStr_Session()
            If OCIDValue1.Value = "" Then
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
            End If

            Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drCC Is Nothing Then
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Exit Sub
            End If

            hid_TNum.Value = Convert.ToString(drCC("TNum"))

            TIMS.IsAllDateCheck(Me, OCIDValue1.Value, "ShowMsg", objconn) '檢查是否為全日制,若是全日制檢查是否符合規則

            Call GetData(Me.OCIDValue1.Value)

            ''審核確認，則不可以修改資料
            'CheckIsVerify(Me.OCIDValue1.Value)

            '執行查詢條件
            GetCourseData()
            ''清理查詢條件
            'If Not Session("Class_CourseName") Is Nothing Then Session("Class_CourseName") = Nothing

            'CourseName.Attributes("onclick") = "Course_search('Edit','" & CourseValue.ClientID & "','" & CourseName.ClientID & "');"
            'CourseName.Style.Item("CURSOR") = "hand"
            'CourseID.Attributes("onDblClick") = "Course_search();"
            'CourseID.Attributes("onchange") = "GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','Room');"
            ''CourseID.Attributes("onclick") = "Course('Add','CourseID','CourseIDValue');"
            'CourseID.Style.Item("CURSOR") = "hand"
            'OLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');" '''Class TIMS.CreateTeacherScript
            'OLessonTeah1.Attributes("ondblclick") = "Get_Teah('OLessonTeah1','OLessonTeah1Value');"
            'OLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','OLessonTeah2');"
            'OLessonTeah2.Attributes("ondblclick") = "Get_Teah('OLessonTeah2','OLessonTeah2Value');"
            Button1.Attributes("onclick") = "Course_search('notepad');"
            CourseID.Attributes("onDblClick") = "Course_search('Edit');"

            Dim sjj As String = "GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','OLessonTeah3Value','OLessonTeah3','OLessonTeah4Value','OLessonTeah4','Room');"
            'CourseID.Attributes("onClick") = sjj ' "GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','Room');"
            'CourseID.Attributes("onChange") = sjj '"GetCourseID(this.value,'CourseID','CourseIDValue','OLessonTeah1Value','OLessonTeah1','OLessonTeah2Value','OLessonTeah2','Room');"
            CourseID.Style.Item("CURSOR") = "hand"

            OLessonTeah1.Attributes.Add("onDblClick", "LessonTeah3('Add','1','','');")
            OLessonTeah1.Style.Item("CURSOR") = "hand"
            OLessonTeah2.Attributes.Add("onDblClick", "LessonTeah3('Add','2','OLessonTeah2','OLessonTeah2Value');")
            OLessonTeah2.Style.Item("CURSOR") = "hand"
            OLessonTeah3.Attributes.Add("onDblClick", "LessonTeah3('Add','3','OLessonTeah3','OLessonTeah3Value');")
            OLessonTeah3.Style.Item("CURSOR") = "hand"
            OLessonTeah4.Attributes.Add("onDblClick", "LessonTeah3('Add','4','OLessonTeah4','OLessonTeah4Value');")
            OLessonTeah4.Style.Item("CURSOR") = "hand"

            ClassSort1.Attributes("onclick") = "CheckClassTime();"
            ClassSort2.Attributes("onclick") = "GetClassTime(1);"
            ClassSort3.Attributes("onclick") = "GetClassTime(2);"
            ClassSort4.Attributes("onclick") = "GetClassTime(3);"
            ClassSort5.Attributes("onclick") = "GetClassTime(4);"

            Button9.Attributes("onclick") = "return CheckNewCourse();"
            Button10.Attributes("onclick") = "return DelChoicClass();"
            LinkButton2.Attributes("onclick") = "ShowCourseList(this);return false;"
        End If

        TIMS.CreateTeacherScript(Me, RIDValue.Value, objconn)
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

        '68:照顧服務員自訓自用訓練計畫  
        If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            labTechN1.Text = cst_textname1 '"教師1"
            labTechN2.Text = cst_textname2 '"教師2"
            labTechN3.Visible = False '不顯示助教2
            OLessonTeah3.Visible = False '不顯示助教2
            OLessonTeah3Value.Visible = False '不顯示助教2
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-207
        '47:補助辦理照顧服務員職業訓練 / 58:補助辦理托育人員職業訓練
        '單月排課作業，修改為可設定3位老師(老師1、老師2、老師3)與1位助教(助教1)
        If TIMS.Cst_TPlanID47AppPlan8.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            labTechN1.Text = cst_textname1b '老師1
            labTechN2.Text = cst_textname2b '老師2
            labTechN3.Text = cst_textname3b '老師3
            labTechN4.Text = cst_textname4b '助教1

            trlabTechN4.Visible = True '顯示助教1<tr>
            labTechN4.Visible = True '顯示助教1
            OLessonTeah4.Visible = True '顯示助教1
            OLessonTeah4Value.Visible = True '顯示助教1
        End If
    End Sub

    '重載課程清單選擇
    Sub GetCourseData2()
        TreeView1.Nodes.Clear() '清理

        Dim sql As String
        sql = "" & vbCrLf
        sql &= " SELECT a.* " & vbCrLf
        sql &= "  ,b.TeachCName TechName1, c.TeachCName TechName2, c3.TeachCName TechName3, c4.TeachCName TechName4 " & vbCrLf
        sql &= "  ,CASE WHEN a2.CourID IS NOT NULL THEN 1 ELSE 0 END Selected " & vbCrLf
        sql &= " FROM Course_CourseInfo a " & vbCrLf
        sql &= " JOIN Auth_CourseInfo a2 ON a.CourID = a2.CourID AND a2.Account = '" & sm.UserInfo.UserID & "' " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo b ON a.Tech1 = b.TechID " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo c ON a.Tech2 = c.TechID " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo c3 ON a.Tech3 = c3.TechID " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo c4 ON a.Tech4 = c4.TechID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= "  AND a.Valid = 'Y' " & vbCrLf
        sql &= "  AND a.RID = '" & RIDValue.Value & "' " & vbCrLf
        sql &= " ORDER BY a.MainCourID ,a.CourID " & vbCrLf
        'Dim dt As New DataTable
        'dt.Load(DbAccess.GetReader(sql, objconn))
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "(課程-依登入者)查無資料!"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            AddTreeView(dt)
        End If
    End Sub

    '課程清單選擇
    Sub GetCourseData()
        TreeView1.Nodes.Clear() '清理

        Dim Class_CourseName As String = ""
        '預設查詢條件
        If Convert.ToString(Session("Class_CourseName")) <> "" Then
            Class_CourseName = Convert.ToString(Session("Class_CourseName"))
            'Session("Class_CourseName") = Nothing
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.* " & vbCrLf
        sql &= " ,b.TeachCName TechName1, c.TeachCName TechName2, c3.TeachCName TechName3, c4.TeachCName TechName4 " & vbCrLf
        sql &= " FROM COURSE_COURSEINFO a" & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo b ON a.Tech1 = b.TechID " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo c ON a.Tech2 = c.TechID " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo c3 ON a.Tech3 = c3.TechID " & vbCrLf
        sql &= " LEFT JOIN Teach_TeacherInfo c4 ON a.Tech4 = c4.TechID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.Valid = 'Y' " & vbCrLf
        sql &= " AND a.RID = '" & RIDValue.Value & "' " & vbCrLf

        '預設查詢條件有值時
        If Class_CourseName <> "" Then sql &= " AND a.CourseName IN (" & Class_CourseName & ") " & vbCrLf
        'Dim dt As New DataTable
        'dt.Load(DbAccess.GetReader(sql, objconn))
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "(課程)查無資料!"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            AddTreeView(dt)
        End If

        ''清理查詢條件
        'If Not Session("Class_CourseName") Is Nothing Then Session("Class_CourseName") = Nothing
    End Sub

    '加入節點
    Sub AddTreeView(ByVal dt As DataTable, Optional ByVal ParentsNode As TreeNode = Nothing, Optional ByVal MainCourID As String = "")
        'Dim dr As DataRow
        Dim RowFilterStr As String = ""

        If ParentsNode Is Nothing Then
            RowFilterStr = "MainCourID IS NULL"
        Else
            RowFilterStr = "MainCourID ='" & MainCourID & "'"
        End If

        For Each dr As DataRow In dt.Select(RowFilterStr)
            Dim sCourseName As String = TIMS.ClearSQM(dr("CourseName"))
            Dim sCourseID As String = TIMS.ClearSQM(dr("CourseID"))
            Dim sCourID As String = TIMS.ClearSQM(dr("CourID"))
            Dim sClassification1 As String = TIMS.ClearSQM(dr("Classification1"))
            Dim sClassification2 As String = TIMS.ClearSQM(dr("Classification2"))
            Dim sTech1 As String = TIMS.ClearSQM(dr("Tech1"))
            Dim sTechName1 As String = TIMS.ClearSQM(dr("TechName1"))
            Dim sTech2 As String = TIMS.ClearSQM(dr("Tech2"))
            Dim sTechName2 As String = TIMS.ClearSQM(dr("TechName2"))
            Dim sTech3 As String = TIMS.ClearSQM(dr("Tech3"))
            Dim sTechName3 As String = TIMS.ClearSQM(dr("TechName3"))
            Dim sTech4 As String = TIMS.ClearSQM(dr("Tech4"))
            Dim sTechName4 As String = TIMS.ClearSQM(dr("TechName4"))
            Dim sRoom As String = TIMS.ClearSQM(dr("Room"))

            Dim h_param As New Hashtable
            h_param.Clear()
            h_param.Add("CourseName", sCourseName)
            h_param.Add("CourseID", sCourseID)
            h_param.Add("Classification1", sClassification1)
            h_param.Add("Classification2", sClassification2)
            Dim NewNode As New TreeNode
            NewNode.Text = TIMS.Get_CourseName(h_param)

            Dim sNNNU2 As String = ""
            Dim sNNNU As String = ""
            sNNNU = ""
            sNNNU &= "'" & sCourseName & "'"
            sNNNU &= ",'" & sCourID & "'"
            sNNNU &= ",'" & sTech1 & "'"
            sNNNU &= ",'" & sTechName1 & "'"
            sNNNU &= ",'" & sTech2 & "'"
            sNNNU &= ",'" & sTechName2 & "'"
            sNNNU &= ",'" & sTech3 & "'"
            sNNNU &= ",'" & sTechName3 & "'"
            sNNNU &= ",'" & sRoom & "'"
            sNNNU2 = "javascript:returnValue(" & sNNNU & ");"

            '68:照顧服務員自訓自用訓練計畫 
            If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                sNNNU = ""
                sNNNU &= "'" & sCourseName & "'"
                sNNNU &= ",'" & sCourID & "'"
                sNNNU &= ",'" & sTech1 & "'"
                sNNNU &= ",'" & sTechName1 & "'"
                sNNNU &= ",'" & sTech2 & "'"
                sNNNU &= ",'" & sTechName2 & "'"
                sNNNU &= ",'" & sRoom & "'"
                sNNNU &= ",'" & sClassification1 & "'"
                sNNNU2 = "javascript:returnValue68(" & sNNNU & ");"
            End If

            If TIMS.Cst_TPlanID47AppPlan8.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                sNNNU = ""
                sNNNU &= "'" & sCourseName & "'"
                sNNNU &= ",'" & sCourID & "'"
                sNNNU &= ",'" & sTech1 & "'"
                sNNNU &= ",'" & sTechName1 & "'"
                sNNNU &= ",'" & sTech2 & "'"
                sNNNU &= ",'" & sTechName2 & "'"
                sNNNU &= ",'" & sTech3 & "'"
                sNNNU &= ",'" & sTechName3 & "'"
                sNNNU &= ",'" & sTech4 & "'"
                sNNNU &= ",'" & sTechName4 & "'"
                sNNNU &= ",'" & sRoom & "'"
                sNNNU2 = "javascript:returnValue47(" & sNNNU & ");"
            End If

            NewNode.NavigateUrl = sNNNU2 'sNNNU

            If ParentsNode Is Nothing Then
                TreeView1.Nodes.Add(NewNode)
                AddTreeView(dt, NewNode, sCourID)
            Else
                'ParentsNode.Nodes.Add(NewNode)
                ParentsNode.ChildNodes.Add(NewNode)
            End If
        Next
    End Sub

    ''' <summary>  取得到傳入參數</summary>
    ''' <returns></returns>
    Function Show_SearchStr_Session() As String
        Dim rst As String = ""

        If Not Session("SearchStr") Is Nothing Then
            center.Text = TIMS.GetMyValue(Session("SearchStr"), "center")
            RIDValue.Value = TIMS.GetMyValue(Session("SearchStr"), "RIDValue")
            OCID1.Text = TIMS.GetMyValue(Session("SearchStr"), "OCID1")
            OCIDValue1.Value = TIMS.GetMyValue(Session("SearchStr"), "OCIDValue1")
            rst = OCIDValue1.Value

            TPeriod.Text = TIMS.GetMyValue(Session("SearchStr"), "TPeriod")
            TPeriodValue.Value = TIMS.GetMyValue(Session("SearchStr"), "TPeriodValue")
            STDate.Value = TIMS.GetMyValue(Session("SearchStr"), "start_date")
            FTDate.Value = TIMS.GetMyValue(Session("SearchStr"), "end_date")

            labTDate.Text = ""
            labTDate.Text &= STDate.Value
            labTDate.Text &= "~"
            labTDate.Text &= FTDate.Value

            THours.Text = TIMS.GetMyValue(Session("SearchStr"), "THours")
            'UseHour.Text = TIMS.GetMyValue(Session("SearchStr"), "UseHour")
            'LeftHour.Text = TIMS.GetMyValue(Session("SearchStr"), "LeftHour")
            Call GetLeftCourseHour(CInt(THours.Text), OCIDValue1.Value)

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
                    If ShowClassNum.SelectedItem Is Nothing Then
                        ShowClassNum.Items(0).Selected = True
                    End If
            End Select

            ShowClassNum.Enabled = False
            Session("SearchStr") = Nothing
        End If

        Return rst
    End Function

    '您已經使用過全期排課，所以無法使用本功能!
    Function ChkClsScheduleType1(ByVal OCIDSelectedValue As String) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM CLASS_SCHEDULE "
        sql &= " WHERE OCID = '" & OCIDSelectedValue & "' "
        sql &= "    AND TYPE = 1 " '您已經使用過全期排課，所以無法使用本功能!
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable

        With sCmd
            '.Parameters.Clear()
            'dt.Load(.ExecuteReader())
            dt = DbAccess.GetDataTable(sql, objconn)
        End With

        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    Sub GetData(ByVal OCIDSelectedValue As String)
        '您已經使用過全期排課，所以無法使用本功能!
        If ChkClsScheduleType1(OCIDSelectedValue) Then
            msg.Text = "您已經使用過全期排課，所以無法使用本功能!"
            CourseTable.Style.Item("display") = "none"
            Exit Sub
        End If

        If STDate.Value <> "" AndAlso FTDate.Value <> "" Then
            Dim StartDate As Date = CDate(STDate.Value)
            Dim EndDate As Date = CDate(FTDate.Value)

            Dim sql As String = ""
            sql = ""
            sql &= " SELECT * FROM CLASS_SCHEDULE "
            sql &= " WHERE OCID = '" & OCIDSelectedValue & "' "
            sql &= " AND SCHOOLDATE >= " & TIMS.To_date(StartDate)
            sql &= " AND SCHOOLDATE <= " & TIMS.To_date(EndDate)
            sql &= " ORDER BY SCHOOLDATE "
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

            While (StartDate <= EndDate)
                If dt.Select("SchoolDate='" & StartDate & "'").Length = 0 Then
                    Dim dr As DataRow = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("OCID") = OCIDSelectedValue
                    dr("SchoolDate") = StartDate
                End If
                StartDate = StartDate.AddDays(1)
            End While

            msg.Text = "(排課)查無資料!"
            CourseTable.Style.Item("display") = "none"

            If dt.Rows.Count >= 0 Then
                msg.Text = ""
                CourseTable.Style.Item("display") = TIMS.cst_inline1 '"inline"

                If dt.Select("Type=1").Length <> 0 Then
                    msg.Text = "您已經使用過全期排課，所以無法使用本功能!"
                    CourseTable.Style.Item("display") = "none"
                Else
                    CourseTable.Style.Item("display") = TIMS.cst_inline1 '"inline"

                    PageControler1.PageDataTable = dt
                    PageControler1.Sort = "SchoolDate"
                    PageControler1.ControlerLoad()

                    Const Cst_位移量 As Integer = 1 '多了一格checkbox
                    Select Case ShowClassNum.SelectedIndex
                        Case 0
                            For i As Integer = 1 To 12
                                DataGrid1.Columns(i + Cst_位移量).Visible = True
                            Next
                        Case 1
                            For i As Integer = 1 To 8
                                DataGrid1.Columns(i + Cst_位移量).Visible = True
                            Next
                            For i As Integer = 9 To 12
                                DataGrid1.Columns(i + Cst_位移量).Visible = False
                            Next
                        Case 2
                            For i As Integer = 1 To 8
                                DataGrid1.Columns(i + Cst_位移量).Visible = False
                            Next
                            For i As Integer = 9 To 12
                                DataGrid1.Columns(i + Cst_位移量).Visible = True
                            Next
                    End Select

                    '將目前所有的使用課程列出----------------------------------------Start
                    Dim CourseList As String = ""
                    Dim CourseArray As Array '課程陣列(排除重複)
                    sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & OCIDSelectedValue & "' ORDER BY SCHOOLDATE "
                    dt = DbAccess.GetDataTable(sql, objconn)

                    For Each dr As DataRow In dt.Rows
                        For i As Integer = 1 To 12
                            If Not IsDBNull(dr("Class" & i)) Then
                                Dim Flag As Boolean = False
                                CourseArray = Split(CourseList, ",")

                                For j As Integer = 0 To CourseArray.Length - 1
                                    If CourseArray(j) = dr("Class" & i) Then Flag = True '課程比對存在
                                Next

                                If Flag = False Then
                                    '表示不存在,增加課程代碼
                                    If CourseList <> "" Then CourseList &= ","
                                    CourseList &= Convert.ToString(dr("Class" & i))
                                End If
                            End If
                        Next
                    Next

                    DataGrid3.Visible = False
                    If CourseList <> "" Then '取出所有相關課程
                        sql = ""
                        sql &= " SELECT a.CourID, a.CourseName, b.CourseName MCourseName, 0 TotalHours "
                        sql &= " FROM (SELECT * FROM Course_CourseInfo WHERE 1=1 AND CourID IN (" & CourseList & ")) a "
                        sql &= " LEFT JOIN Course_CourseInfo b ON a.MainCourID = b.CourID "
                        sql &= " ORDER BY a.CourseName "
                        Dim dt1 As DataTable
                        dt1 = DbAccess.GetDataTable(sql, objconn)
                        For Each dr1 As DataRow In dt1.Rows 'Course_CourseInfo
                            For Each dr As DataRow In dt.Rows 'CLASS_SCHEDULE
                                For i As Integer = 1 To 12
                                    If Not IsDBNull(dr("Class" & i)) Then
                                        If dr("Class" & i) = dr1("CourID") Then dr1("TotalHours") += 1 '上課時數+1(增加)
                                    End If
                                Next
                            Next
                        Next

                        If dt1.Rows.Count > 0 Then
                            DataGrid3.Visible = True
                            DataGrid3.DataSource = dt1
                            DataGrid3.DataBind()
                        End If
                    End If
                    '將目前所有的使用課程列出----------------------------------------End
                End If
            End If
        End If
    End Sub

    '儲存
    Sub Save_Class_Schedule(ByVal CSID As String, ByVal OCID As String, ByVal SchoolDate As Date, ByRef Mydt As DataTable) 'As String
        'CSID: Class_Schedule pk流水號
        Dim rst As String = "" '返回的錯誤訊息
        'Save_Class_Schedule = "" '返回的錯誤訊息
        'Dim Mydt As DataTable = Me.ViewState("DetailTable")
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing

        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")
        If Mydt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim str_SchoolDate As String = SchoolDate.ToString("yyyy/MM/dd")
        Using conn As SqlConnection = DbAccess.GetConnection()
            Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)
            Try
                Dim sql As String = ""
                '先將課程本日課程資料填入Class_Schedule------   Start'Call TIMS.OpenDbConn(conn)'trans = DbAccess.BeginTrans(conn)
                'If trans.Connection.State = ConnectionState.Closed Then trans.Connection.Open() 
                If $"{CSID}" = "" Then
                    sql = $" SELECT * FROM CLASS_SCHEDULE WHERE OCID={OCID} AND SCHOOLDATE ={TIMS.To_date(str_SchoolDate)}" 'sql &= " ORDER BY SCHOOLDATE"
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    If dt.Rows.Count = 0 Then
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("CSID") = DbAccess.GetNewId(trans, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                        dr("OCID") = TIMS.CINT1(OCID)
                        dr("SchoolDate") = TIMS.Cdate2(SchoolDate) 'FormatDateTime(SchoolDate, DateFormat.ShortDate)
                    Else
                        dr = dt.Rows(0)
                    End If
                Else
                    sql = $" SELECT * FROM CLASS_SCHEDULE WHERE CSID={CSID} AND OCID={OCID}"
                    dt = DbAccess.GetDataTable(sql, da, trans)
                    dr = dt.Rows(0)
                End If

                dr("Formal") = "Y"
                dr("Type") = 2
                Dim Tmps As String = ""
                For i As Integer = 1 To 12
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("CourseID")))
                    dr("Class" & i) = IIf(Tmps = "", Convert.DBNull, Tmps)
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher1ID")))
                    dr("Teacher" & i) = IIf(Tmps = "", Convert.DBNull, Tmps)
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher2ID")))
                    dr("Teacher" & i + 12) = IIf(Tmps = "", Convert.DBNull, Tmps)
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher3ID")))
                    dr("Teacher" & i + 24) = IIf(Tmps = "", Convert.DBNull, Tmps)
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("Teacher4ID")))
                    dr("Teacher" & i + 36) = IIf(Tmps = "", Convert.DBNull, Tmps) '37~48
                    Tmps = TIMS.ClearSQM(Convert.ToString(Mydt.Rows(i - 1)("ClassRoom")))
                    dr("Room" & i) = IIf(Tmps = "", Convert.DBNull, Tmps)
                Next
                dr("VACATION") = If(v_Vacation <> "", v_Vacation, Convert.DBNull) '假日排課
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                DbAccess.UpdateDataTable(dt, da, trans)
                '先將課程本日課程資料填入Class_Schedule------   End

                'Plan_Schedule '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
                SD_04_002.AddNew_Plan_Schedule(Me, STDate.Value, FTDate.Value, OCID, dt, da, trans, dtCourse)
                DbAccess.CommitTrans(trans)
            Catch ex As Exception
                DbAccess.RollbackTrans(trans)
                Throw ex
            End Try
        End Using

        'Return rst
    End Sub

    ''' <summary> 計算可用的時數(排課時數) 將可用值存入 LeftHour.Text ，若無法使用 則  Button9.Enabled 為 False </summary>
    ''' <param name="Total"></param>
    ''' <param name="OCID"></param>
    Sub GetLeftCourseHour(ByVal Total As Integer, ByVal OCID As String)
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
        Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"

        'Dim holiday_CourID As String
        'holiday_CourID = TIMS.Get_CourID(cst_holiday, objconn)
        UseHour.Text = "0"

        Call TIMS.OpenDbConn(objconn)
        sql = " SELECT * FROM CLASS_SCHEDULE WHERE OCID = '" & OCID & "' AND TYPE = '2' ORDER BY SCHOOLDATE "
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count <> 0 Then
            For Each dr In dt.Rows
                For i As Integer = 1 To 12
                    If dr("Class" & i).ToString <> "" AndAlso dr("Class" & i).ToString <> i_HoliDayCourID.ToString() Then
                        '在資料庫裡新增一筆假日的資料供排假日用
                        'insert INTO Course_CourseInfo(CourseID,CoursENAME,Classification1,Classification2,RID,ModifyAcct,ModifyDATE)
                        'VALUES('holiday','假日',1,0,'','sys',getdate())
                        Dim sql99 As String = ""       '加上判斷排除不計算排課時數
                        Dim countHours As String = ""
                        sql99 = "SELECT ISCOUNTHOURS FROM COURSE_COURSEINFO WHERE COURID=" & dr("Class" & i).ToString
                        Dim dt99 As DataTable = DbAccess.GetDataTable(sql99, objconn)
                        If dt99.Rows.Count > 0 Then countHours = Convert.ToString(dt99.Rows(0)("isCountHours"))
                        If countHours = "" Then
                            Total -= 1
                            UseHour.Text = Int(UseHour.Text) + 1
                        End If
                    End If
                Next
            Next
        End If

        LeftHour.Text = Total
        Button9.Enabled = True '新增排課

        If Total = 0 Then
            Button9.Enabled = False '新增排課
            TIMS.Tooltip(Button9, "排課時數已經用完!", True)
            Common.MessageBox(Me, "排課時數已經用完!")
        ElseIf Total < 0 Then
            Button9.Enabled = False '新增排課
            TIMS.Tooltip(Button9, "排課時數已經用完!", True)
            Common.MessageBox(Me, "排課時數已經用完!")
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Case ListItemType.Header, ListItemType.Footer
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim MyLink As LinkButton = e.Item.FindControl("LinkButton1")
                'Dim dt As DataTable = dtCourse
                MyLink.Enabled = False

                If HolidayTable.Select("HolDate='" & drv("SchoolDate") & "'").Length <> 0 Then
                    'MyLink.Enabled = False
                    Dim ttip As String = Convert.ToString(HolidayTable.Select("HolDate='" & drv("SchoolDate") & "'")(0)("Reason"))
                    TIMS.Tooltip(MyLink, ttip)
                    'MyLink.ToolTip = ttip 'HolidayTable.Select("HolDate='" & drv("SchoolDate") & "'")(0)("Reason")
                End If

                'MyLink.Text = FormatDateTime(drv("SchoolDate"), DateFormat.ShortDate) & "(" & TIMS.GetWeekDay(CDate(drv("SchoolDate")).DayOfWeek) & ")"
                'MyLink.CommandArgument = FormatDateTime(drv("SchoolDate"), DateFormat.ShortDate)
                MyLink.Text = TIMS.Cdate3(drv("SchoolDate")) & "(" & TIMS.GetWeekDay(CDate(drv("SchoolDate")).DayOfWeek) & ")"
                MyLink.CommandArgument = TIMS.Cdate3(drv("SchoolDate")) 'FormatDateTime(drv("SchoolDate"), DateFormat.ShortDate)
                MyLink.ForeColor = Color.Blue

                If Me.ViewState("IsClosed") = "Y" Then
                    TIMS.Tooltip(MyLink, "班級已結訓")
                    MyLink.Attributes("onclick") = "return false;"
                    MyLink.ForeColor = Color.Black
                End If

                '已審核確認
                If Me.ViewState("IsVerify") = "Y" Then
                    TIMS.Tooltip(MyLink, "課程已審核確認")
                    MyLink.Attributes("onclick") = "return false;"
                    MyLink.ForeColor = Color.Black
                End If

                Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
                Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"

                For i As Integer = 1 To e.Item.Cells.Count - 2
                    If e.Item.Cells(i).Text <> "&nbsp;" And e.Item.Cells(i).Text <> "" Then
                        ff3 = "CourID='" & e.Item.Cells(i).Text & "'"
                        Dim v_COURSENAME As String = ""
                        If dtCourse.Select(ff3).Length <> 0 Then
                            Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
                            v_COURSENAME = If(Convert.ToString(drv("Vacation")) = "Y", String.Format("{0}(假日)", dr1("CourseName")), dr1("CourseName"))
                        ElseIf Convert.ToString(drv("Vacation")) = "Y" Then
                            v_COURSENAME = s_HoliDayCourseName '"假日"
                        End If
                        If v_COURSENAME <> "" Then e.Item.Cells(i).Text = v_COURSENAME
                    End If
                Next

        End Select
    End Sub

    '取得或新增排課
    Function GetDetailCourse(ByVal MyDateStr As String, ByVal OCID As String) As DataTable
        'MyDate.Text = MyDateStr
        'MyWeek.Text = "(" & TIMS.GetWeekDay(CDate(MyDateStr).DayOfWeek) & ")"
        'Dim UsedHour As Integer
        '建立空白資料表----------------------------------------Start
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
        '建立空白資料表----------------------------------------End

        Dim CSIDValue As String = ""
        Dim VACATION As String = ""
        Dim sql As String = ""
        Dim dr As DataRow
        sql = ""
        sql &= " SELECT * FROM CLASS_SCHEDULE "
        sql &= " WHERE OCID = '" & OCID & "' "
        sql &= " AND SCHOOLDATE = " & TIMS.To_date(MyDateStr)
        'sql &= " Order By SchoolDate"
        dr = DbAccess.GetOneRow(sql, objconn)

        If dr Is Nothing Then
            CSIDValue = ""
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
                drTemp("VACATION") = ""
            Next
        Else
            Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
            Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"

            CSIDValue = dr("CSID")
            VACATION = Convert.ToString(dr("VACATION"))
            For i As Integer = 1 To 12
                drTemp = dtTemp.NewRow
                dtTemp.Rows.Add(drTemp)
                drTemp("ClassNum") = i
                If dr("Class" & i).ToString = "" Then
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
                    drTemp("VACATION") = ""
                Else
                    drTemp("CourseID") = dr("Class" & i)
                    ff3 = String.Concat("CourID='", dr("Class" & i), "'")
                    Dim v_COURSENAME As String = ""
                    If dtCourse.Select(ff3).Length <> 0 Then
                        Dim dr1 As DataRow = dtCourse.Select(ff3)(0)
                        v_COURSENAME = If(VACATION = "Y", String.Format("{0}(假日)", dr1("CourseName")), dr1("CourseName"))
                    ElseIf VACATION = "Y" Then
                        v_COURSENAME = s_HoliDayCourseName '"假日"
                    End If
                    drTemp("CourseName") = v_COURSENAME
                    Dim s1 As String = CStr(i)
                    If dr("Teacher" & s1).ToString <> "" Then
                        drTemp("Teacher1") = TIMS.Get_TeacherName(dr("Teacher" & s1), dtTeacher)
                        drTemp("Teacher1ID") = Val(dr("Teacher" & s1))
                    End If
                    Dim s12 As String = CStr(i + 12)
                    If dr("Teacher" & s12).ToString <> "" Then
                        drTemp("Teacher2") = TIMS.Get_TeacherName(dr("Teacher" & s12), dtTeacher)
                        drTemp("Teacher2ID") = Val(dr("Teacher" & s12))
                    End If
                    Dim s24 As String = CStr(i + 24)
                    If dr("Teacher" & s24).ToString <> "" Then
                        drTemp("Teacher3") = TIMS.Get_TeacherName(dr("Teacher" & s24), dtTeacher)
                        drTemp("Teacher3ID") = Val(dr("Teacher" & s24))
                    End If
                    Dim s36 As String = CStr(i + 36) '37~48
                    If dr("Teacher" & s36).ToString <> "" Then
                        drTemp("Teacher4") = TIMS.Get_TeacherName(dr("Teacher" & s36), dtTeacher)
                        drTemp("Teacher4ID") = Val(dr("Teacher" & s36))
                    End If
                    drTemp("ClassRoom") = dr("Room" & i)
                    drTemp("VACATION") = VACATION
                    'UsedHour += 1
                End If
            Next
        End If

        Return dtTemp

#Region "(No Use)"

        'TodayUseHour.Value = UsedHour
        'Me.ViewState("DetailTable") = dtTemp '將ViewState("DetailTable") 取得到最新的資料，若無資料塞入空白

        'DataGrid2.DataSource = dtTemp
        'DataGrid2.DataKeyField = "ClassNum"
        'DataGrid2.DataBind()

        'GetUsedClass()  '顯示能夠編輯的課程節次(日班只能編輯1-8節...)
        'GetLeftCourseHour(Int(THours.Text)) '計算可用的時數(排課時數) 將可用值存入 LeftHour.Text ，若無法使用 則  Button9.Enabled 為 False

        'If Me.ViewState("IsVerify") = "Y" Then
        '    '已審核確認
        '    Button12.Enabled = False '刪除
        '    Button10.Enabled = False '刪除
        '    Button9.Enabled = False '新增
        '    Button3.Enabled = False '儲存

        '    TIMS.Tooltip(Button12, "此班級已審核確認", True)
        '    TIMS.Tooltip(Button10, "此班級已審核確認", True)
        '    TIMS.Tooltip(Button9, "此班級已審核確認", True)
        '    TIMS.Tooltip(Button3, "此班級已審核確認", True)
        'End If

#End Region
    End Function

    '刪除 Class_Schedule
    Public Shared Sub DeleteDetailCourse_A(ByRef MyPage As Page, ByRef oDG1 As DataGrid, ByRef vClassSort1x As String,
                                           ByVal OCID As String, ByVal oConn As SqlConnection)
        Const cst_f1a As String = "CSID,OCID,SCHOOLDATE,CLASS1,CLASS2,CLASS3,CLASS4,CLASS5,CLASS6,CLASS7,CLASS8,CLASS9,CLASS10,CLASS11,CLASS12,TEACHER1,TEACHER2,TEACHER3,TEACHER4,TEACHER5,TEACHER6,TEACHER7,TEACHER8,TEACHER9,TEACHER10,TEACHER11,TEACHER12,TEACHER13,TEACHER14,TEACHER15,TEACHER16,TEACHER17,TEACHER18,TEACHER19,TEACHER20,TEACHER21,TEACHER22,TEACHER23,TEACHER24,ROOM1,ROOM2,ROOM3,ROOM4,ROOM5,ROOM6,ROOM7,ROOM8,ROOM9,ROOM10,ROOM11,ROOM12,TYPE,MODIFYACCT,MODIFYDATE,FORMAL,VACATION,TEACHER25,TEACHER26,TEACHER27,TEACHER28,TEACHER29,TEACHER30,TEACHER31,TEACHER32,TEACHER33,TEACHER34,TEACHER35,TEACHER36"
        'getdate()
        Const cst_f1b As String = "CSID,OCID,SCHOOLDATE,CLASS1,CLASS2,CLASS3,CLASS4,CLASS5,CLASS6,CLASS7,CLASS8,CLASS9,CLASS10,CLASS11,CLASS12,TEACHER1,TEACHER2,TEACHER3,TEACHER4,TEACHER5,TEACHER6,TEACHER7,TEACHER8,TEACHER9,TEACHER10,TEACHER11,TEACHER12,TEACHER13,TEACHER14,TEACHER15,TEACHER16,TEACHER17,TEACHER18,TEACHER19,TEACHER20,TEACHER21,TEACHER22,TEACHER23,TEACHER24,ROOM1,ROOM2,ROOM3,ROOM4,ROOM5,ROOM6,ROOM7,ROOM8,ROOM9,ROOM10,ROOM11,ROOM12,TYPE,MODIFYACCT,getdate(),FORMAL,VACATION,TEACHER25,TEACHER26,TEACHER27,TEACHER28,TEACHER29,TEACHER30,TEACHER31,TEACHER32,TEACHER33,TEACHER34,TEACHER35,TEACHER36"

        Dim sm As SessionModel = SessionModel.Instance()
        Dim sql As String = ""
        sql = ""
        sql &= " UPDATE Class_Schedule "
        sql &= " SET modifydate = GETDATE(), modifyacct = @modifyacct "
        sql &= " WHERE ocid = @ocid AND SchoolDate = CONVERT(DATE, @SchoolDate)"
        Dim uCmd As New SqlCommand(sql, oConn)
        'Dim uSql As String = sql

        sql = ""
        sql &= " INSERT INTO CLASS_SCHEDULELOG (" & cst_f1a & ") "
        sql &= " SELECT " & cst_f1b
        sql &= " FROM CLASS_SCHEDULE "
        sql &= " WHERE 1=1 "
        sql &= " AND modifyacct = @modifyacct AND ocid = @ocid AND SchoolDate = CONVERT(DATE, @SchoolDate)"
        Dim iCmd As New SqlCommand(sql, oConn)
        'Dim iSql As String = sql

        'sql = ""
        'sql &= " DELETE Class_Schedule WHERE 1=1 "
        'sql &= "    AND modifyacct = @modifyacct AND ocid = @ocid AND SchoolDate = CONVERT(DATETIME, @SchoolDate, 111) "
        'Dim dCmd As New SqlCommand(sql, oConn)

        Dim vClassSort1x2 As String = TIMS.CombiSQM2IN(vClassSort1x)
        Dim tmpV2 As String = "Vacation=NULL"
        For Xi As Integer = 1 To 12
            Dim tmpXo As String = "'" & CStr(Xi) & "'"
            If vClassSort1x2.IndexOf(tmpXo) > -1 Then
                If tmpV2 <> "" Then tmpV2 &= ","
                tmpV2 &= "CLASS" & CStr(Xi) & "=NULL"
                If tmpV2 <> "" Then tmpV2 &= ","
                tmpV2 &= "TEACHER" & CStr(Xi) & "=NULL"
                If tmpV2 <> "" Then tmpV2 &= ","
                tmpV2 &= "TEACHER" & CStr(Xi + 12) & "=NULL"
                If tmpV2 <> "" Then tmpV2 &= ","
                tmpV2 &= "TEACHER" & CStr(Xi + 24) & "=NULL"
                If tmpV2 <> "" Then tmpV2 &= ","
                tmpV2 &= "TEACHER" & CStr(Xi + 36) & "=NULL"
                If tmpV2 <> "" Then tmpV2 &= ","
                tmpV2 &= "ROOM" & CStr(Xi) & "=NULL"
            End If
        Next

        sql = ""
        sql &= " UPDATE CLASS_SCHEDULE "
        sql &= " SET " & tmpV2
        sql &= " WHERE 1=1 "
        sql &= " AND modifyacct = @modifyacct AND ocid = @ocid AND SchoolDate = CONVERT(DATE, @SchoolDate)"
        Dim uCmd2 As New SqlCommand(sql, oConn)
        'Dim uSql2 As String = sql

        '勾選的日期
        TIMS.OpenDbConn(oConn)
        For Each ItemA As DataGridItem In oDG1.Items
            Dim SelectClass1 As HtmlInputCheckBox = ItemA.FindControl("SelectClass1")
            '有勾選的日期
            If SelectClass1.Checked Then
                Dim LinkButton1 As LinkButton = ItemA.FindControl("LinkButton1")
                Dim SchoolData As String = LinkButton1.CommandArgument '(SchoolData)
                'SchoolData = MyLink.CommandArgument
                'Dim myParam As Hashtable = New Hashtable
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("ocid", SqlDbType.Int).Value = OCID
                    .Parameters.Add("SchoolDate", SqlDbType.VarChar).Value = TIMS.Cdate3(SchoolData)
                    .ExecuteNonQuery()
                    'myParam.Add("modifyacct", sm.UserInfo.UserID)
                    'myParam.Add("ocid", OCID)
                    'myParam.Add("SchoolDate", TIMS.cdate3(SchoolData))
                    'DbAccess.ExecuteNonQuery(uSql, oConn, myParam)
                End With
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("ocid", SqlDbType.Int).Value = OCID
                    .Parameters.Add("SchoolDate", SqlDbType.VarChar).Value = TIMS.Cdate3(SchoolData)
                    .ExecuteNonQuery()
                    'myParam.Add("modifyacct", sm.UserInfo.UserID)
                    'myParam.Add("ocid", OCID)
                    'myParam.Add("SchoolDate", TIMS.cdate3(SchoolData))
                    'DbAccess.ExecuteNonQuery(iSql, oConn, myParam)
                End With
                With uCmd2 'dCmd
                    .Parameters.Clear()
                    .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                    .Parameters.Add("ocid", SqlDbType.Int).Value = OCID
                    .Parameters.Add("SchoolDate", SqlDbType.VarChar).Value = TIMS.Cdate3(SchoolData)
                    .ExecuteNonQuery()
                    'myParam.Add("modifyacct", sm.UserInfo.UserID)
                    'myParam.Add("ocid", OCID)
                    'myParam.Add("SchoolDate", TIMS.cdate3(SchoolData))
                    'DbAccess.ExecuteNonQuery(uSql2, oConn, myParam)
                End With
                'sql = ""
                'sql &= " DELETE Class_Schedule"
                'sql &= " WHERE OCID='" & OCID & "'" 'and Type='2'"'1:批次；2:單月
                'sql &= " and SchoolDate= " & TIMS.to_date(SchoolData)
                'DbAccess.ExecuteNonQuery(sql, objconn)
            End If
        Next
    End Sub

    ''' <summary>新增排課_A (save) </summary>
    ''' <param name="OCID"></param>
    Sub CreateDetailCourse_A(ByVal OCID As String)
        ''假如審核確認，則不可以修改資料
        'If TIMS.Chk_ClassSchVerify(OCIDValue1.Value) Then Me.ViewState("IsVerify") = "Y"

        If RIDValue.Value = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim sql As String = ""
        sql = " SELECT TECHID, TEACHCNAME FROM TEACH_TEACHERINFO WHERE WORKSTATUS = '1' AND RID = @RID "
        Dim sCmd As New SqlCommand(sql, objconn)
        dtTeacher = New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
            dtTeacher.Load(.ExecuteReader())
            'Dim myParam As Hashtable = New Hashtable
            'myParam.Add("RID", RIDValue.Value)
            'dtTeacher = DbAccess.GetDataTable(sql, objconn, myParam)
        End With
        'dtTeacher.Load(DbAccess.GetReader(sql, objconn))
        'dtTeacher = DbAccess.GetDataTable(sql, objconn)

        OLessonTeah1.Text = TIMS.ClearSQM(OLessonTeah1.Text)
        OLessonTeah1Value.Value = TIMS.ClearSQM(OLessonTeah1Value.Value)
        If OLessonTeah1.Text = "" Then OLessonTeah1Value.Value = ""

        OLessonTeah2.Text = TIMS.ClearSQM(OLessonTeah2.Text)
        OLessonTeah2Value.Value = TIMS.ClearSQM(OLessonTeah2Value.Value)
        If OLessonTeah2.Text = "" Then OLessonTeah2Value.Value = ""

        OLessonTeah3.Text = TIMS.ClearSQM(OLessonTeah3.Text)
        OLessonTeah3Value.Value = TIMS.ClearSQM(OLessonTeah3Value.Value)
        If OLessonTeah3.Text = "" Then OLessonTeah3Value.Value = ""

        OLessonTeah4.Text = TIMS.ClearSQM(OLessonTeah4.Text)
        OLessonTeah4Value.Value = TIMS.ClearSQM(OLessonTeah4Value.Value)
        If OLessonTeah4.Text = "" Then OLessonTeah4Value.Value = ""

        'Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
        'Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"

        '排課選項 TypeRadio 0 '一般排課  1 '假日排課
        Dim v_TypeRadio As String = TIMS.GetListValue(TypeRadio)
        Dim v_Vacation As String = If(v_TypeRadio = "1", "Y", "")

        '勾選的日期
        For Each ItemA As DataGridItem In DataGrid1.Items
            Dim dr As DataRow
            Dim dt As DataTable
            Dim drv As DataRowView = ItemA.DataItem
            Dim SelectClass1 As HtmlInputCheckBox = ItemA.FindControl("SelectClass1")

            '有勾選的日期
            If SelectClass1.Checked Then
                Dim MyLink As LinkButton = ItemA.FindControl("LinkButton1")
                dt = GetDetailCourse(MyLink.CommandArgument, OCID)

                '勾選的節次
                For Each Item As ListItem In ClassSort1.Items
                    'If Item.Selected = True And Int(LeftHour.Text) > 0 Then
                    '有勾選的節次
                    If Item.Selected = True Then
                        If dt.Select("ClassNum='" & Item.Value & "'").Length <> 0 Then
                            dr = dt.Select("ClassNum='" & Item.Value & "'")(0)
                            If dr("CourseName").ToString = "" Then LeftHour.Text = Int(LeftHour.Text) - 1

                            dr("CourseName") = CourseID.Text
                            dr("CourseID") = CourseIDValue.Value 'CourseIDValue.Value 流水id 用以顯現資料

                            dr("ClassRoom") = Room.Text
                            dr("Teacher1") = OLessonTeah1.Text
                            dr("Teacher1ID") = OLessonTeah1Value.Value
                            dr("Teacher2") = OLessonTeah2.Text
                            dr("Teacher2ID") = OLessonTeah2Value.Value
                            dr("Teacher3") = OLessonTeah3.Text
                            dr("Teacher3ID") = OLessonTeah3Value.Value
                            dr("Teacher4") = OLessonTeah4.Text
                            dr("Teacher4ID") = OLessonTeah4Value.Value
                            dr("VACATION") = If(v_Vacation <> "", v_Vacation, "")
                        End If
                    End If
                Next

                'Select Case TypeRadio.SelectedIndex
                '    Case 0 '一般排課
                '        '勾選的節次
                '        For Each Item As ListItem In ClassSort1.Items
                '            'If Item.Selected = True And Int(LeftHour.Text) > 0 Then
                '            '有勾選的節次
                '            If Item.Selected = True Then
                '                If dt.Select("ClassNum='" & Item.Value & "'").Length <> 0 Then
                '                    dr = dt.Select("ClassNum='" & Item.Value & "'")(0)
                '                    If dr("CourseName").ToString = "" Then LeftHour.Text = Int(LeftHour.Text) - 1
                '                    dr("CourseName") = CourseID.Text
                '                    dr("CourseID") = CourseIDValue.Value 'CourseIDValue.Value 流水id 用以顯現資料
                '                    dr("ClassRoom") = Room.Text
                '                    dr("Teacher1") = OLessonTeah1.Text
                '                    dr("Teacher1ID") = OLessonTeah1Value.Value
                '                    dr("Teacher2") = OLessonTeah2.Text
                '                    dr("Teacher2ID") = OLessonTeah2Value.Value
                '                    dr("Teacher3") = OLessonTeah3.Text
                '                    dr("Teacher3ID") = OLessonTeah3Value.Value
                '                    dr("Teacher4") = OLessonTeah4.Text
                '                    dr("Teacher4ID") = OLessonTeah4Value.Value
                '                End If
                '            End If
                '        Next
                '        'For Each dr In dt.Rows
                '        '    If dr("CourseName").ToString <> "" Then UsedHour += 1
                '        'Next
                '    Case 1 '假日排課
                '        '勾選的節次
                '        Dim i_HoliDayCourID As Integer = TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
                '        Dim s_HoliDayCourseName As String = TIMS.Get_CourseName(i_HoliDayCourID, Nothing, objconn) '"假日"
                '        For Each Item As ListItem In ClassSort1.Items
                '            '有勾選的節次
                '            If Item.Selected = True Then
                '                If dt.Select("ClassNum='" & Item.Value & "'").Length <> 0 Then
                '                    dr = dt.Select("ClassNum='" & Item.Value & "'")(0)
                '                    dr("CourseName") = s_HoliDayCourseName 'TIMS.Get_CourseName(TIMS.Get_CourID(cst_holiday, objconn), Nothing, objconn) '"假日"
                '                    dr("CourseID") = i_HoliDayCourID 'TIMS.Get_CourID(cst_holiday, objconn) '"10000000"
                '                    'dr("CourseName") = CourseID.Text
                '                    'dr("CourseID") = CourseIDValue.Value 'CourseIDValue.Value 流水id 用以顯現資料
                '                    dr("ClassRoom") = Convert.DBNull
                '                    dr("Teacher1") = Convert.DBNull
                '                    dr("Teacher1ID") = Convert.DBNull
                '                    dr("Teacher2") = Convert.DBNull
                '                    dr("Teacher2ID") = Convert.DBNull
                '                    dr("Teacher3") = Convert.DBNull
                '                    dr("Teacher3ID") = Convert.DBNull
                '                    dr("Teacher4") = Convert.DBNull
                '                    dr("Teacher4ID") = Convert.DBNull
                '                End If
                '            End If
                '        Next
                'End Select
                'Me.ViewState("DetailTable") = dt
                Call Save_Class_Schedule("", OCID, MyLink.CommandArgument, dt)
            End If

            'Dim FunID As HtmlInputCheckBox = Item.FindControl("FunID")
            'Dim LID As HtmlInputHidden = Item.FindControl("LID")
            'Dim PlanID As HtmlInputHidden = Item.FindControl("PlanID")
            'Dim OrgID As HtmlInputHidden = Item.FindControl("OrgID")
            'If FunID.Checked Then
            '    If Session("PlanIDValue" & CStr(PlanID.Value)) Is Nothing Then
            '        Session("PlanIDValue" & CStr(PlanID.Value)) = 1
            '    Else
            '        flag = False
            '        errmsg = "同一計劃，不同機構，不可申請兩次"
            '        Session.Clear()
            '        Exit For
            '        'Session("PlanIDValue" & PlanID.Value) += 1
            '    End If
            '    i += 1
            'End If
        Next
    End Sub

    '確認勾選日期
    Function CheckDetailCourse() As Boolean
        Dim rst As Boolean = False
        '=== === === === check DetailCourse start === === === === 
        'Dim msg As String = ""

        For Each ItemA As DataGridItem In DataGrid1.Items
            'Dim drv As DataRowView = ItemA.DataItem
            'Dim MyLink As LinkButton = ItemA.FindControl("LinkButton1")
            Dim SelectClass1 As HtmlInputCheckBox = ItemA.FindControl("SelectClass1")

            If SelectClass1.Checked Then
                rst = True
                Exit For
            End If
        Next

#Region "(No Use)"

        'If Not rst Then msg += "請勾選排課日期!!!" & vbCrLf
        'If msg <> "" Then
        '    Common.MessageBox(Me, msg)
        '    Exit Function
        'End If
        '=== === === === check DetailCourse end === === === ===

#End Region
        Return rst
    End Function

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

    '新增排課
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '確認勾選日期
        Dim Cdc1 As Boolean = CheckDetailCourse()
        If Not Cdc1 Then
            Dim msg As String = "請勾選排課日期!!!" & vbCrLf
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        '68:照顧服務員自訓自用訓練計畫 
        If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'labTechN1.Text = "教師1"
            'labTechN2.Text = "教師2"
            'labTechN3.Visible = False '不顯示助教2
            'OLessonTeah3.Visible = False '不顯示助教2
            'OLessonTeah3Value.Visible = False '不顯示助教2
            OLessonTeah3.Text = "" '不顯示助教2
            OLessonTeah3Value.Value = "" '不顯示助教2
        End If

        Dim Reason As String = ""
        If Reason = "" Then
            CourseIDValue.Value = TIMS.ClearSQM(CourseIDValue.Value)
            If CourseIDValue.Value = "" Then
                Reason &= "該課程代碼有誤!!" & vbCrLf
                Common.MessageBox(Me, Reason)
                Exit Sub
            End If

            If Not TIMS.IsNumeric2(CourseIDValue.Value) Then
                Reason &= "該課程代碼有誤!!" & vbCrLf
                Common.MessageBox(Me, Reason)
                Exit Sub
            End If

            ff3 = "COURID='" & CourseIDValue.Value & "'"
            If dtCourse.Select(ff3).Length = 0 Then
                Reason &= "該課程代碼有誤!!" & vbCrLf
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

        'can save
        Call CreateDetailCourse_A(Me.OCIDValue1.Value)
        Call GetData(Me.OCIDValue1.Value)
        Call GetLeftCourseHour(CInt(THours.Text), Me.OCIDValue1.Value)
    End Sub

    '刪除排課-依節數
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '確認勾選日期
        Dim Cdc1 As Boolean = CheckDetailCourse()
        If Not Cdc1 Then
            Dim msg As String = "請勾選要刪除的排課日期!!!" & vbCrLf
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        '確認勾選節次
        Dim vClassSort1x As String = CheckClassSort1x()
        If vClassSort1x = "" Then
            Dim msg As String = "節次不可為空，請勾選要刪除的排課節次!!!" & vbCrLf
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        'can save delete
        Call DeleteDetailCourse_A(Me, DataGrid1, vClassSort1x, OCIDValue1.Value, objconn)
        Call GetData(Me.OCIDValue1.Value)
        Call GetLeftCourseHour(CInt(THours.Text), Me.OCIDValue1.Value)
    End Sub

    Sub GetSearchStr()
        THours.Text = Trim(THours.Text)
        LeftHour.Text = Trim(LeftHour.Text)
        If THours.Text = "" Then THours.Text = "0"
        If LeftHour.Text = "" Then LeftHour.Text = "0"
        Dim sUseHour As String = Convert.ToString(CInt(THours.Text) - CInt(LeftHour.Text))
        'Session("SearchStr") = "k=1"
        'Session("SearchStr") += "&center=" & center.Text
        'Session("SearchStr") += "&RIDValue=" & RIDValue.Value
        ''Session("SearchStr") += "&TMID1=" & TMID1.Text
        'Session("SearchStr") += "&OCID1=" & OCID1.Text
        ''Session("SearchStr") += "&TMIDValue1=" & TMIDValue1.Value
        'Session("SearchStr") += "&OCIDValue1=" & OCIDValue1.Value
        'Session("SearchStr") += "&start_date=" & STDate.Value
        'Session("SearchStr") += "&end_date=" & FTDate.Value
        'Session("SearchStr") += "&TPeriod=" & TPeriod.Text
        'Session("SearchStr") += "&TPeriodValue=" & TPeriodValue.Value
        'Session("SearchStr") += "&THours=" & THours.Text
        'Session("SearchStr") += "&LeftHour=" & LeftHour.Text
        'Session("SearchStr") += "&UseHour=" & sUseHour
        Dim sSearchStr As String = ""
        Call TIMS.SetMyValue(sSearchStr, "k", "1")
        Call TIMS.SetMyValue(sSearchStr, "center", center.Text)
        Call TIMS.SetMyValue(sSearchStr, "RIDValue", RIDValue.Value)
        Call TIMS.SetMyValue(sSearchStr, "OCID1", OCID1.Text)
        Call TIMS.SetMyValue(sSearchStr, "OCIDValue1", OCIDValue1.Value)
        Call TIMS.SetMyValue(sSearchStr, "start_date", STDate.Value) '原班起始值
        Call TIMS.SetMyValue(sSearchStr, "end_date", FTDate.Value) '原班結束值
        Call TIMS.SetMyValue(sSearchStr, "TPeriod", TPeriod.Text)
        Call TIMS.SetMyValue(sSearchStr, "TPeriodValue", TPeriodValue.Value)
        Call TIMS.SetMyValue(sSearchStr, "THours", THours.Text)
        Call TIMS.SetMyValue(sSearchStr, "LeftHour", LeftHour.Text)
        Call TIMS.SetMyValue(sSearchStr, "UseHour", sUseHour)
        Session("SearchStr") = sSearchStr
    End Sub

    '回排課列表
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        Call GetSearchStr()
        TIMS.Utl_Redirect1(Me, "SD_04_002.aspx?ID=" & Request("ID") & "&k=SD_04_002_ADD")
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call GetCourseData2()
        ''CreateDetailCourse_A(Me.OCIDValue1.Value)
        'GetData(Me.OCIDValue1.Value)
        'GetLeftCourseHour(CInt(THours.Text), Me.OCIDValue1.Value)
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class