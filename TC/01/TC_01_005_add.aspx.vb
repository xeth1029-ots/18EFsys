Partial Class TC_01_005_add
    Inherits AuthBasePage

    'COURSE_COURSEINFO
    'Public Overrides Sub Validate()
    '    Re_Classification2.Enabled = True
    '    Select Case Classification1_List.SelectedValue
    '        Case "2" '術科
    '            Re_Classification2.Enabled = False
    '    End Select
    '    'Select Case Classification1_List.SelectedValue
    '    '    Case "2" '術科
    '    '        Re_Classification2.Enabled = False
    '    '    Case "1" '學科
    '    'End Select
    '    MyBase.Validate()
    'End Sub

    '若計畫異常或機構名稱異常，則表示該機構該計畫無權設定
    Sub GetPlanName1(ByVal PlanID As String, ByVal RID As String)
        Dim PlanName As String = ""
        Dim OrgName As String = ""

        TBplan.Text = ""
        orgid_value.Value = ""

        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn)
        Dim dr As DataRow
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT a.Years + c.Name + b.PlanName + a.Seq PlanName "
        sql &= " FROM ID_Plan a   "
        sql &= " JOIN Key_Plan b ON a.TPlanID=b.TPlanID "
        sql &= " JOIN ID_District c ON a.DistID=c.DistID "
        sql &= " WHERE 1=1"
        sql &= " AND a.PlanID='" & PlanID & "'"
        dr = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then PlanName = dr("PlanName")

        sql = ""
        sql &= " SELECT b.OrgName,a.orgid,a.PlanID "
        sql &= " FROM Auth_Relship a   "
        sql &= " JOIN Org_OrgInfo b   ON a.OrgID=b.OrgID "
        sql &= " WHERE 1=1"
        sql &= " AND a.RID='" & RID & "'"
        If Len(RID) <> "1" Then
            '非委訓單位
            sql &= " AND a.PlanID='" & PlanID & "'"
        End If
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then
            OrgName = dr("OrgName").ToString()
            If PlanName <> "" AndAlso OrgName <> "" Then
                TBplan.Text = PlanName & "_" & OrgName
                orgid_value.Value = dr("orgid").ToString()
            End If
        End If
    End Sub



    Const cst_Techer1s As String = "教師,教師2,教師3,助教1,助教2,助教1"
    Const cst_Techer2s As String = "教師1,教師2,教師3,助教1,助教2,助教1" '限定 Cst_TPlanID47AppPlan7

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End
        'Dim FunDr As DataRow
        'Dim vsClassYears As String = ""

        Type_str.Value = TIMS.ClearSQM(Request("ProcessType"))
        Re_ID.Value = TIMS.ClearSQM(Request("courid"))
        '依 ProcessType 做判斷 順序要寫對
        bt_save.Enabled = True
        bt_save.Enabled = True
        'Select Case Type_str.Value
        '    Case "Update"
        '        If Not au.blnCanMod Then
        '            bt_save.Enabled = False
        '            TIMS.Tooltip(bt_save, "無修改權限", True)
        '        End If
        '    Case "Insert"
        '        If Not au.blnCanAdds Then
        '            bt_save.Enabled = False
        '            TIMS.Tooltip(bt_save, "無新增權限", True)
        '        End If
        'End Select

        If Not IsPostBack Then
            Call Create1()
        End If
    End Sub

    Sub Create1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        LabelTeah1_3.Visible = True
        OLessonTeah1_3.Visible = True
        '顯示名稱改變
        LabelTeah1.Text = cst_Techer1s.Split(",")(0)
        LabelTeah1_2.Text = cst_Techer1s.Split(",")(1)
        LabelTeah1_3.Text = cst_Techer1s.Split(",")(2)
        LabelTeah2.Text = cst_Techer1s.Split(",")(3)
        LabelTeah3.Text = cst_Techer1s.Split(",")(4)
        LabelTeah2b.Text = cst_Techer1s.Split(",")(5)
        If TIMS.Cst_TPlanID47AppPlan7.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            LabelTeah1.Text = cst_Techer2s.Split(",")(0)
            LabelTeah1_2.Text = cst_Techer2s.Split(",")(1)
            LabelTeah1_3.Text = cst_Techer2s.Split(",")(2)
            LabelTeah2.Text = cst_Techer2s.Split(",")(3)
            LabelTeah3.Text = cst_Techer2s.Split(",")(4)
            LabelTeah2b.Text = cst_Techer2s.Split(",")(5)
        End If

        trTeahList1.Visible = False '教師2.3
        trTeahList2.Visible = True '助教1.2
        trTeahList3.Visible = False '助教1
        '依計畫確認是否顯示(委訓單位)
        If TIMS.Cst_TPlanID47AppPlan7.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trTeahList1.Visible = True '教師2.3
            trTeahList2.Visible = False '助教1.2
            trTeahList3.Visible = True '助教1
        End If

        '68:照顧服務員自訓自用訓練計畫  
        If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trTeahList3.Visible = False '助教1
            LabelTeah1_3.Visible = False '不顯示教師3
            OLessonTeah1_3.Visible = False '不顯示教師3
        End If

        Dim vsClassYears As String = ""
        Hidsave1.Value = ""
        RIDValue1.Value = sm.UserInfo.RID
        PlanIDValue.Value = sm.UserInfo.PlanID
        TPlanIDValue.Value = sm.UserInfo.TPlanID

        '保留搜尋條件
        ViewState("MySreach") = Session("MySreach")
        'Session("MySreach") = Nothing

        '20100208 按新增時代查詢之 課程代碼 & 課程名稱
        TB_CourseID.Text = ""
        TB_CourseName.Text = ""
        If Convert.ToString(Request("ClassID")) <> "" Then
            TB_CourseID.Text = Convert.ToString(Request("ClassID"))
        End If
        If Convert.ToString(Request("ClassName")) <> "" Then
            TB_CourseName.Text = Convert.ToString(Request("ClassName"))
        End If

        OLessonTeah1.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','1','','');")
        OLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');"

        OLessonTeah1_2.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','12','OLessonTeah1_2','OLessonTeah1_2Value');")
        OLessonTeah1_2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1_2Value','OLessonTeah1_2');"
        OLessonTeah1_3.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','13','OLessonTeah1_3','OLessonTeah1_3Value');")
        OLessonTeah1_3.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1_3Value','OLessonTeah1_3');"

        OLessonTeah2.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','2','OLessonTeah2','OLessonTeah2Value');")
        OLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','OLessonTeah2');"
        OLessonTeah3.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','3','OLessonTeah3','OLessonTeah3Value');")
        OLessonTeah3.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah3Value','OLessonTeah3');"

        OLessonTeah2b.Attributes.Add("onDblClick", "javascript:LessonTeah3('Add','2','OLessonTeah2b','OLessonTeah2bValue');")
        OLessonTeah2b.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2bValue','OLessonTeah2b');"

        Select Case Type_str.Value
            Case "Insert" '新增
                lblProecessType.Text = "新增"
                '新增時添加預設值-----------------------------2005/4/11
                RIDValue.Value = sm.UserInfo.RID
                PlanIDValue.Value = sm.UserInfo.PlanID
                TPlanIDValue.Value = sm.UserInfo.TPlanID

                '若計畫異常或機構名稱異常，則表示該機構該計畫無權設定
                Call GetPlanName1(PlanIDValue.Value, RIDValue.Value)

                Common.SetListItem(Classification1_List, "1")
                Common.SetListItem(Classification2_List, "2")
                '新增時添加預設值-----------------------------2005/4/11

                '是否有效
                CB_Valid.Checked = True
            Case "Update" '修改
                lblProecessType.Text = "修改"
                Re_ID.Value = TIMS.ClearSQM(Re_ID.Value)
                If Re_ID.Value = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                Dim row_list As DataRow
                Dim sqlstr_list As String = ""
                sqlstr_list = "SELECT * FROM COURSE_COURSEINFO WHERE COURID=@COURID"
                Dim parms As New Hashtable
                parms.Add("COURID", Re_ID.Value)
                row_list = DbAccess.GetOneRow(sqlstr_list, objconn, parms)
                If row_list Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                'trainValue.Value = ""
                'vsTrainValue =trainValue.Value
                trainValue.Value = Convert.ToString(row_list("TMID"))
                If (trainValue.Value <> "") Then
                    Dim parms2 As New Hashtable
                    parms2.Add("TMID", trainValue.Value)
                    Dim train_sql As String = ""
                    train_sql = "SELECT TRAINNAME FROM KEY_TRAINTYPE WHERE TMID=@TMID "
                    Dim train_name As String = DbAccess.ExecuteScalar(train_sql, objconn, parms2)
                    TB_career_id.Text = train_name
                End If

                RIDValue.Value = Convert.ToString(row_list("RID"))
                TB_CourseID.Text = TIMS.ClearSQM(row_list("CourseID"))
                TB_CourseName.Text = TIMS.ClearSQM(row_list("CourseName"))
                TB_Hours.Text = Convert.ToString(row_list("Hours"))
                If Convert.ToString(row_list("Classification1")) <> "" Then
                    'Classification1_List.SelectedValue = row_list("Classification1")
                    Common.SetListItem(Classification1_List, row_list("Classification1"))
                End If

                Classification2_List.SelectedIndex = 0
                If Convert.ToString(row_list("Classification1")) = "2" Then
                    Classification2_List.Enabled = False
                End If
                Common.SetListItem(Classification2_List, row_list("Classification2"))

                courid.Value = Convert.ToString(row_list("MainCourID"))
                '修改 若MainCourID欄位為Null 
                Dim MainName As String = TIMS.Get_CourseName(courid.Value, objconn)
                MainName = TIMS.Get_CourseName(courid.Value, objconn)
                TB_CourName.Text = ""
                If MainName <> "" Then TB_CourName.Text = MainName

                If Convert.ToString(Request("ErrorData")) = "1" Then
                    TIMS.Tooltip(TB_CourName, "主課程資料異常，請重新設定!!!")
                    TB_CourName.Text += "(主課程資料異常，請重新設定!!!)"
                    TB_CourName.ForeColor = Color.Red 'TB_CourName.ForeColor.Red
                End If

                CB_Valid.Checked = False
                If Convert.ToString(row_list("Valid")) = "Y" Then
                    CB_Valid.Checked = True
                End If

                Dim sTmp As String = ""
                OLessonTeah1Value.Value = ""
                OLessonTeah1.Text = ""
                sTmp = Convert.ToString(row_list("Tech1"))
                OLessonTeah1Value.Value = sTmp
                OLessonTeah1.Text = TIMS.Get_TeachCName(sTmp, objconn) ' drA("TeachCName")

                OLessonTeah1_2Value.Value = ""
                OLessonTeah1_2.Text = ""
                sTmp = Convert.ToString(row_list("Tech1_2"))
                OLessonTeah1_2Value.Value = sTmp
                OLessonTeah1_2.Text = TIMS.Get_TeachCName(sTmp, objconn) ' drA("TeachCName")

                OLessonTeah1_3Value.Value = ""
                OLessonTeah1_3.Text = ""
                sTmp = Convert.ToString(row_list("Tech1_3"))
                OLessonTeah1_3Value.Value = sTmp
                OLessonTeah1_3.Text = TIMS.Get_TeachCName(sTmp, objconn) ' drA("TeachCName")

                OLessonTeah2Value.Value = "" '(助教1)
                OLessonTeah2.Text = ""
                sTmp = Convert.ToString(row_list("Tech2"))
                OLessonTeah2Value.Value = sTmp
                OLessonTeah2.Text = TIMS.Get_TeachCName(sTmp, objconn) ' drA("TeachCName")

                OLessonTeah3Value.Value = "" '(助教2)
                OLessonTeah3.Text = ""
                sTmp = Convert.ToString(row_list("Tech3"))
                OLessonTeah3Value.Value = sTmp
                OLessonTeah3.Text = TIMS.Get_TeachCName(sTmp, objconn) ' drA("TeachCName")

                OLessonTeah2bValue.Value = "" '(助教1)
                OLessonTeah2b.Text = ""
                sTmp = Convert.ToString(row_list("Tech2"))
                OLessonTeah2bValue.Value = sTmp
                OLessonTeah2b.Text = TIMS.Get_TeachCName(sTmp, objconn) ' drA("TeachCName")

                Room.Text = ""
                If Convert.ToString(row_list("Room")) <> "" Then
                    Room.Text = row_list("Room")
                End If

                'If Convert.IsDBNull(row_list("IsCountHours")) Then '是否計算排課時數
                '    Common.SetListItem(IsCountHours, "Y")
                'Else
                '    Common.SetListItem(IsCountHours, row_list("IsCountHours"))
                'End If
                'Dim planid As String
                'Dim sqlstr_Planid As String = "select OrgName,a.orgid from Auth_Relship a join org_orginfo b on a.orgid =b.orgid where a.rid='" & row_list("RID") & "'"
                'Dim drow As DataRow = DbAccess.GetOneRow(sqlstr_Planid)
                'If drow Is Nothing Then
                '   TBplan.Text = drow("OrgName")
                '    '2005/6/13-新增機構代碼欄位-Melody
                '   orgid_value.Value = drow("orgid")
                'End If

                '若計畫異常或機構名稱異常，則表示該機構該計畫無權設定
                Call GetPlanName1(PlanIDValue.Value, RIDValue.Value)

                '歸屬班別代碼
                Classid_Hid.Value = ""
                Classid.Text = ""
                If Convert.ToString(row_list("CLSID")) <> "" Then
                    Classid_Hid.Value = row_list("CLSID")

                    Dim class_name As DataRow
                    Dim sqlstr_className As String = "SELECT YEARS,TPLANID,CLASSID,CLASSNAME FROM ID_CLASS WHERE CLSID='" & row_list("CLSID") & "'"
                    class_name = DbAccess.GetOneRow(sqlstr_className, objconn)
                    If Not class_name Is Nothing Then
                        vsClassYears = Convert.ToString(class_name("Years"))
                        Classid.Text = Convert.ToString(class_name("ClassID")) & "(" & Convert.ToString(class_name("CLassName")) & ")"
                        'Dim TPlan As String = "select TPlanID from ID_Class where CLSID='" & row_list("CLSID") & "'"
                        'Dim Tplan_ID As String = Convert.ToString(DbAccess.ExecuteScalar(TPlan, objconn))
                        TPlanIDValue.Value = class_name("TPlanID") 'Tplan_ID
                    End If
                End If
                'Classification1_List.Attributes("onchange") = "ChangeCourseSort(this.value);"
                Dim CF1v1 As String = Classification1_List.SelectedValue
                If CF1v1 <> "" Then
                    Dim strScript As String = ""
                    strScript = ""
                    strScript &= "<script>"
                    strScript &= "ChangeCourseSort('" & CF1v1 & "');"
                    strScript &= "</script>"
                    TIMS.RegisterStartupScript(Me, TIMS.xBlockName, "")
                End If

        End Select

        TIMS.CreateTeacherScript(Me, RIDValue.Value, objconn)
        'bt_save.Attributes("onclick") = "check();"
        Classification1_List.Attributes("onchange") = "ChangeCourseSort(this.value);"

        '隸屬班級 選擇鈕
        If vsClassYears = "" Then
            Button2.Attributes("onclick") = "javascript:wopen('TC_01_005_Classid.aspx?ProcessType=" & Type_str.Value & "&amp;tplanid=" & TPlanIDValue.Value & "','班級代碼',1000,630,1)"
        Else
            Button2.Attributes("onclick") = "javascript:wopen('TC_01_005_Classid.aspx?ProcessType=" & Type_str.Value & "&amp;tplanid=" & TPlanIDValue.Value & "&amp;ID_ClassYears=" & Convert.ToString(vsClassYears) & "','班級代碼',1000,630,1)"
        End If

    End Sub

    '儲存'檢核1
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        orgid_value.Value = TIMS.ClearSQM(Me.orgid_value.Value)
        If orgid_value.Value = "" Then
            Errmsg += "計畫階段 機構有誤，請重新選擇!!" & vbCrLf
        End If

        '課程代碼 最多8碼 寬限10碼
        TB_CourseID.Text = TIMS.ClearSQM(TB_CourseID.Text)
        If TB_CourseID.Text <> "" Then
            'TB_CourseID.Text = Trim(TB_CourseID.Text)
            If Len(TB_CourseID.Text) > 10 Then
                Errmsg += "課程代碼 長度超過系統範圍(10)" & vbCrLf
            End If
        Else
            TB_CourseID.Text = ""
            Errmsg += "請輸入 課程代碼" & vbCrLf
        End If

        TB_CourseName.Text = TIMS.ClearSQM(TB_CourseName.Text)
        If TB_CourseName.Text <> "" Then
            'TB_CourseName.Text = Trim(TB_CourseName.Text)
            If Len(TB_CourseName.Text) > 50 Then
                Errmsg += "課程名稱 長度超過系統範圍(50)" & vbCrLf
            End If
        Else
            TB_CourseName.Text = ""
            Errmsg += "請輸入 課程名稱" & vbCrLf
        End If

        '學/術科
        Select Case Classification1_List.SelectedValue
            Case "1", "2" '學/術科
            Case Else
                Errmsg += "請選擇 學/術科" & vbCrLf
        End Select

        'Dim flagchkSub2 As Boolean = True
        Select Case Classification1_List.SelectedValue
            Case "1"
            Case "2" '術科 (目前只能選專業)
                Select Case Classification2_List.SelectedValue
                    Case "2"
                    Case Else
                        Errmsg &= "學/術科：若選擇為 「術科」，共同/一般/專業：只能選擇「專業」" & vbCrLf
                End Select
        End Select
        '術科 (目前只能選專業)
        '共同/一般/專業
        Select Case Classification2_List.SelectedValue
            Case "0", "1", "2"
            Case Else
                Errmsg += "請選擇 共同/一般/專業" & vbCrLf
        End Select

        '訓練職類
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        If trainValue.Value = "" Then
            Errmsg += "請選擇 訓練職類" & vbCrLf
        End If

        'If trTeahList1.Visible Then '教師2.3
        'If trTeahList2.Visible Then '助教1.2
        If trTeahList1.Visible Then '教師2.3
            If OLessonTeah1_3Value.Value <> "" AndAlso OLessonTeah1_2Value.Value = "" Then
                Errmsg += "教師3有填資料，教師2不可無資料!! " & vbCrLf
            End If
            If OLessonTeah1_2Value.Value <> "" AndAlso OLessonTeah1Value.Value = "" Then
                Errmsg += "教師2有填資料，教師1不可無資料!! " & vbCrLf
            End If
        End If
        If trTeahList2.Visible Then '助教1.2
            If OLessonTeah3Value.Value <> "" AndAlso OLessonTeah2Value.Value = "" Then
                Errmsg += "助教2有填資料，助教1不可無資料!! " & vbCrLf
            End If
        End If

        Room.Text = TIMS.ClearSQM(Room.Text)
        If Room.Text <> "" Then
            'Room.Text = Trim(Room.Text)
            If Len(Room.Text) > 30 Then
                Errmsg += "教室 內容長度超過系統範圍(30)" & vbCrLf
            End If
            'Else : Room.Text = ""
            'Errmsg += "請輸入 教室內容" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        Dim tchName As String = ""
        Dim sTmp As String = ""
        Dim sChkValue As String = ""
        If OLessonTeah1Value.Value <> "" Then
            '第1筆資料無須判斷
            sTmp = "'" & OLessonTeah1Value.Value & "'"
            If sChkValue <> "" Then sChkValue &= ","
            sChkValue &= sTmp
        End If
        If trTeahList1.Visible Then '教師2.3
            '因為前面有資料
            If OLessonTeah1_2Value.Value <> "" Then
                tchName = cst_Techer1s.Split(",")(1)
                sTmp = "'" & OLessonTeah1_2Value.Value & "'"
                If sChkValue.IndexOf(sTmp) = -1 Then
                    If sChkValue <> "" Then sChkValue &= ","
                    sChkValue &= sTmp
                Else
                    Errmsg += tchName & "資料已重複!! " & vbCrLf
                End If
            End If
            '因為前面有資料
            If OLessonTeah1_3Value.Value <> "" Then
                tchName = cst_Techer1s.Split(",")(2)
                sTmp = "'" & OLessonTeah1_3Value.Value & "'"
                If sChkValue.IndexOf(sTmp) = -1 Then
                    If sChkValue <> "" Then sChkValue &= ","
                    sChkValue &= sTmp
                Else
                    Errmsg += tchName & "資料已重複!! " & vbCrLf
                End If
            End If
        End If

        If trTeahList2.Visible Then '助教1.2
            '不一定有資料
            If OLessonTeah2Value.Value <> "" Then
                tchName = cst_Techer1s.Split(",")(3)
                sTmp = "'" & OLessonTeah2Value.Value & "'"
                If sChkValue <> "" Then
                    If sChkValue.IndexOf(sTmp) = -1 Then
                        If sChkValue <> "" Then sChkValue &= ","
                        sChkValue &= sTmp
                    Else
                        Errmsg += tchName & "資料已重複!! " & vbCrLf
                    End If
                Else
                    If sChkValue <> "" Then sChkValue &= ","
                    sChkValue &= sTmp
                End If
            End If

            '因為前面有資料
            If OLessonTeah3Value.Value <> "" Then
                tchName = cst_Techer1s.Split(",")(4)
                sTmp = "'" & OLessonTeah3Value.Value & "'"
                If sChkValue.IndexOf(sTmp) = -1 Then
                    If sChkValue <> "" Then sChkValue &= ","
                    sChkValue &= sTmp
                Else
                    Errmsg += tchName & "資料已重複!! " & vbCrLf
                End If
            End If
        End If

        If trTeahList3.Visible Then '助教1
            '不一定有資料
            If OLessonTeah2bValue.Value <> "" Then
                tchName = cst_Techer1s.Split(",")(5)
                sTmp = "'" & OLessonTeah2bValue.Value & "'"
                If sChkValue <> "" Then
                    If sChkValue.IndexOf(sTmp) = -1 Then
                        If sChkValue <> "" Then sChkValue &= ","
                        sChkValue &= sTmp
                    Else
                        Errmsg += tchName & "資料已重複!! " & vbCrLf
                    End If
                Else
                    If sChkValue <> "" Then sChkValue &= ","
                    sChkValue &= sTmp
                End If
            End If
        End If

        'If Errmsg <> "" Then Return False
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存'檢核2
    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT 'X' FROM COURSE_COURSEINFO WHERE COURID=@CourID"
        Dim sCmd As New SqlCommand(sql, objconn)

        'Select Case Type_str.Value
        '    Case "Insert"
        '    Case "Update"
        'End Select

        Select Case Type_str.Value
            Case "Insert"
                sql = ""
                sql &= " SELECT COURSEID FROM COURSE_COURSEINFO "
                sql &= " WHERE CourseID= '" & TB_CourseID.Text & "'"
                sql &= " AND RID='" & RIDValue.Value & "'"
                If DbAccess.GetCount(sql, objconn) > 0 Then
                    Errmsg &= "課程代碼重複!!!!"
                    Return False 'Exit Function
                End If

            Case "Update"
                If Re_ID.Value = "" Then
                    Errmsg &= "課程代碼有誤，請重新查詢操作!!!!"
                    Return False 'Exit Function
                End If

                sql = ""
                sql &= " SELECT COURSEID FROM COURSE_COURSEINFO"
                sql &= " where CourID<>'" & Re_ID.Value & "'"
                sql &= " and CourseID= '" & TB_CourseID.Text & "'"
                sql &= " and RID='" & RIDValue.Value & "'"
                If DbAccess.GetCount(sql, objconn) > 0 Then
                    Errmsg &= "課程代碼重複!!"
                    Return False 'Exit Function
                End If

                TIMS.OpenDbConn(objconn)
                Dim dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("CourID", SqlDbType.VarChar).Value = Re_ID.Value
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count = 0 Then
                    Errmsg &= "課程代碼有誤，請重新查詢操作!!!!"
                    Return False 'Exit Function
                End If

            Case Else
                Errmsg &= "資料有誤，請重新操作功能!!"
                Return False 'Exit Function
        End Select

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '儲存
    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        Dim strScript As String = ""
        If Hidsave1.Value <> "" Then
            'Common.MessageBox(Me, "儲存動作執行中!!")
            'Exit Sub '儲存動作已經執行
            strScript = ""
            strScript = "<script language=""javascript"">" + vbCrLf
            strScript &= "location.href='TC_01_005.aspx?ID=" & Request("ID") & "';" + vbCrLf
            strScript &= "</script>"
            TIMS.RegisterStartupScript(Page, TIMS.xBlockName(), strScript)
            Exit Sub
        End If

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        If Not Page.IsValid Then
            Common.MessageBox(Me, "資料有誤，請檢查輸入資料!!")
            Exit Sub
        End If

        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call SaveData1()
    End Sub

    '儲存
    Sub SaveData1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " INSERT INTO COURSE_COURSEINFO (" & vbCrLf
        sql &= " COURID" & vbCrLf '/*PK*/ 
        sql &= " ,COURSEID" & vbCrLf
        sql &= " ,COURSENAME" & vbCrLf
        sql &= " ,HOURS" & vbCrLf
        sql &= " ,CLASSIFICATION1" & vbCrLf
        sql &= " ,CLASSIFICATION2" & vbCrLf
        sql &= " ,VALID" & vbCrLf
        sql &= " ,MAINCOURID" & vbCrLf
        sql &= " ,RID" & vbCrLf '分署(無計畫差別)/委訓才有計畫差別
        sql &= " ,TMID" & vbCrLf
        sql &= " ,MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE" & vbCrLf
        sql &= " ,CLSID" & vbCrLf
        sql &= " ,ORGID" & vbCrLf
        sql &= " ,TECH1" & vbCrLf
        If trTeahList1.Visible Then '教師2.3
            sql &= " ,TECH1_2" & vbCrLf
            sql &= " ,TECH1_3" & vbCrLf
        End If
        If trTeahList2.Visible Then '助教1.2
            sql &= " ,TECH2" & vbCrLf
            sql &= " ,TECH3" & vbCrLf
        End If
        If trTeahList3.Visible Then '助教1
            sql &= " ,TECH2" & vbCrLf
        End If
        sql &= " ,ROOM" & vbCrLf
        sql &= " ,PLANID" & vbCrLf
        'sql += " ,ISCOUNTHOURS" & vbCrLf
        sql &= " ) VALUES ( " & vbCrLf
        sql &= " @COURID" & vbCrLf '/*PK*/ 
        sql &= " ,@COURSEID" & vbCrLf
        sql &= " ,@COURSENAME" & vbCrLf
        sql &= " ,@HOURS" & vbCrLf
        sql &= " ,@CLASSIFICATION1" & vbCrLf
        sql &= " ,@CLASSIFICATION2" & vbCrLf
        sql &= " ,@VALID" & vbCrLf
        sql &= " ,@MAINCOURID" & vbCrLf
        sql &= " ,@RID" & vbCrLf
        sql &= " ,@TMID" & vbCrLf
        sql &= " ,@MODIFYACCT" & vbCrLf
        sql &= " ,getdate()" & vbCrLf '@MODIFYDATE
        sql &= " ,@CLSID" & vbCrLf
        sql &= " ,@ORGID" & vbCrLf
        sql &= " ,@TECH1" & vbCrLf
        If trTeahList1.Visible Then '教師2.3
            sql &= " ,@TECH1_2" & vbCrLf
            sql &= " ,@TECH1_3" & vbCrLf
        End If
        If trTeahList2.Visible Then '助教1.2
            sql &= " ,@TECH2" & vbCrLf
            sql &= " ,@TECH3" & vbCrLf
        End If
        If trTeahList3.Visible Then '助教1
            sql &= " ,@TECH2" & vbCrLf
        End If
        sql &= " ,@ROOM" & vbCrLf
        sql &= " ,@PLANID" & vbCrLf
        'sql += " ,@ISCOUNTHOURS" & vbCrLf
        sql &= " )" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql &= " UPDATE COURSE_COURSEINFO" & vbCrLf
        sql &= " SET COURSEID=@COURSEID" & vbCrLf
        sql &= " ,COURSENAME=@COURSENAME" & vbCrLf
        sql &= " ,HOURS=@HOURS" & vbCrLf
        sql &= " ,CLASSIFICATION1=@CLASSIFICATION1" & vbCrLf
        sql &= " ,CLASSIFICATION2=@CLASSIFICATION2" & vbCrLf
        sql &= " ,VALID=@VALID" & vbCrLf
        sql &= " ,MAINCOURID=@MAINCOURID" & vbCrLf
        sql &= " ,RID=@RID" & vbCrLf
        sql &= " ,TMID=@TMID" & vbCrLf
        sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql &= " ,MODIFYDATE=getdate()" & vbCrLf '@MODIFYDATE
        sql &= " ,CLSID=@CLSID" & vbCrLf
        sql &= " ,ORGID=@ORGID" & vbCrLf
        sql &= " ,TECH1=@TECH1" & vbCrLf
        If trTeahList1.Visible Then '教師2.3
            sql &= " ,TECH1_2=@TECH1_2" & vbCrLf
            sql &= " ,TECH1_3=@TECH1_3" & vbCrLf
        End If
        If trTeahList2.Visible Then '助教1.2
            sql &= " ,TECH2=@TECH2" & vbCrLf
            sql &= " ,TECH3=@TECH3" & vbCrLf
        End If
        If trTeahList3.Visible Then '助教1
            sql &= " ,TECH2=@TECH2" & vbCrLf
        End If
        sql &= " ,ROOM=@ROOM" & vbCrLf
        sql &= " ,PLANID=@PLANID" & vbCrLf
        'sql += " ,ISCOUNTHOURS=@ISCOUNTHOURS" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND COURID=@COURID" & vbCrLf '/*PK*/ 
        Dim uCmd As New SqlCommand(sql, objconn)

        Dim iCOURID As Integer = 0
        If Re_ID.Value <> "" Then iCOURID = Re_ID.Value 'UPDATE
        Dim vCB_Valid As String = "N"
        If CB_Valid.Checked Then vCB_Valid = "Y"

        Select Case Type_str.Value
            Case "Insert"
                iCOURID = DbAccess.GetNewId(objconn, "COURSE_COURSEINFO_COURID_SEQ,COURSE_COURSEINFO,COURID")
                With iCmd
                    .Parameters.Clear()
                    .Parameters.Add("COURID", SqlDbType.Int).Value = iCOURID
                    .Parameters.Add("COURSEID", SqlDbType.VarChar).Value = TB_CourseID.Text
                    .Parameters.Add("COURSENAME", SqlDbType.NVarChar).Value = TB_CourseName.Text

                    .Parameters.Add("HOURS", SqlDbType.Int).Value = If(TB_Hours.Text = "", Convert.DBNull, Val(TB_Hours.Text))

                    .Parameters.Add("CLASSIFICATION1", SqlDbType.Int).Value = Classification1_List.SelectedValue
                    '術科 只能選專業
                    .Parameters.Add("CLASSIFICATION2", SqlDbType.Int).Value = Classification2_List.SelectedValue '"0", "1", "2"

                    .Parameters.Add("VALID", SqlDbType.VarChar).Value = vCB_Valid
                    .Parameters.Add("MAINCOURID", SqlDbType.Int).Value = If(courid.Value = "", Convert.DBNull, Val(courid.Value)) 'MAINCOURID.Text

                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    .Parameters.Add("TMID", SqlDbType.VarChar).Value = If(trainValue.Value = "", Convert.DBNull, Val(trainValue.Value)) 'TMID.Text
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID 'MODIFYACCT.Text
                    '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE.Text
                    .Parameters.Add("CLSID", SqlDbType.Int).Value = If(Classid_Hid.Value = "", Convert.DBNull, Val(Classid_Hid.Value)) 'CLSID.Text
                    '2005/6/13-新增機構代碼欄位-Melody
                    .Parameters.Add("ORGID", SqlDbType.Int).Value = Val(Me.orgid_value.Value)
                    '2005/08/15 新增授課老師一、授課老師二及教室
                    .Parameters.Add("TECH1", SqlDbType.Int).Value = If(OLessonTeah1Value.Value = "", Convert.DBNull, Val(OLessonTeah1Value.Value)) ' TECH1.Text
                    If trTeahList1.Visible Then '教師2.3
                        .Parameters.Add("TECH1_2", SqlDbType.Int).Value = If(OLessonTeah1_2Value.Value = "", Convert.DBNull, Val(OLessonTeah1_2Value.Value)) ' TECH1.Text
                        .Parameters.Add("TECH1_3", SqlDbType.Int).Value = If(OLessonTeah1_3Value.Value = "", Convert.DBNull, Val(OLessonTeah1_3Value.Value)) ' TECH1.Text
                    End If
                    If trTeahList2.Visible Then '助教1.2
                        .Parameters.Add("TECH2", SqlDbType.Int).Value = If(OLessonTeah2Value.Value = "", Convert.DBNull, Val(OLessonTeah2Value.Value)) 'TECH2.Text
                        .Parameters.Add("TECH3", SqlDbType.Int).Value = If(OLessonTeah3Value.Value = "", Convert.DBNull, Val(OLessonTeah3Value.Value)) 'TECH2.Text
                    End If
                    If trTeahList3.Visible Then '助教1
                        .Parameters.Add("TECH2", SqlDbType.Int).Value = If(OLessonTeah2bValue.Value = "", Convert.DBNull, Val(OLessonTeah2bValue.Value)) 'TECH2.Text
                    End If
                    .Parameters.Add("ROOM", SqlDbType.VarChar).Value = If(Room.Text = "", Convert.DBNull, Room.Text)
                    .Parameters.Add("PLANID", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
                    '.Parameters.Add("ISCOUNTHOURS", SqlDbType.VarChar).Value = IsCountHours.SelectedValue 'Y/N '是否計算排課時數
                    '.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
                End With
            Case "Update"
                If iCOURID = 0 Then Exit Sub '異常

                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("COURSEID", SqlDbType.VarChar).Value = TB_CourseID.Text
                    .Parameters.Add("COURSENAME", SqlDbType.NVarChar).Value = TB_CourseName.Text

                    .Parameters.Add("HOURS", SqlDbType.Int).Value = If(TB_Hours.Text = "", Convert.DBNull, Val(TB_Hours.Text))

                    .Parameters.Add("CLASSIFICATION1", SqlDbType.Int).Value = Classification1_List.SelectedValue
                    '術科 只能選專業
                    .Parameters.Add("CLASSIFICATION2", SqlDbType.Int).Value = Classification2_List.SelectedValue '"0", "1", "2"

                    .Parameters.Add("VALID", SqlDbType.VarChar).Value = vCB_Valid
                    .Parameters.Add("MAINCOURID", SqlDbType.Int).Value = If(courid.Value = "", Convert.DBNull, Val(courid.Value)) 'MAINCOURID.Text

                    .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                    .Parameters.Add("TMID", SqlDbType.VarChar).Value = If(trainValue.Value = "", Convert.DBNull, Val(trainValue.Value)) 'TMID.Text
                    .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID 'MODIFYACCT.Text
                    '.Parameters.Add("MODIFYDATE", SqlDbType.VarChar).Value = MODIFYDATE.Text
                    .Parameters.Add("CLSID", SqlDbType.Int).Value = If(Classid_Hid.Value = "", Convert.DBNull, Val(Classid_Hid.Value)) 'CLSID.Text

                    '2005/6/13-新增機構代碼欄位-Melody
                    .Parameters.Add("ORGID", SqlDbType.Int).Value = Val(Me.orgid_value.Value)
                    '2005/08/15 新增授課老師一、授課老師二及教室
                    .Parameters.Add("TECH1", SqlDbType.Int).Value = If(OLessonTeah1Value.Value = "", Convert.DBNull, Val(OLessonTeah1Value.Value)) ' TECH1.Text
                    If trTeahList1.Visible Then '教師2.3
                        .Parameters.Add("TECH1_2", SqlDbType.Int).Value = If(OLessonTeah1_2Value.Value = "", Convert.DBNull, Val(OLessonTeah1_2Value.Value)) ' TECH1_2
                        .Parameters.Add("TECH1_3", SqlDbType.Int).Value = If(OLessonTeah1_3Value.Value = "", Convert.DBNull, Val(OLessonTeah1_3Value.Value)) ' TECH1_3
                    End If
                    If trTeahList2.Visible Then '助教1.2
                        .Parameters.Add("TECH2", SqlDbType.Int).Value = If(OLessonTeah2Value.Value = "", Convert.DBNull, Val(OLessonTeah2Value.Value)) 'TECH2.Text
                        .Parameters.Add("TECH3", SqlDbType.Int).Value = If(OLessonTeah3Value.Value = "", Convert.DBNull, Val(OLessonTeah3Value.Value)) 'TECH2.Text
                    End If
                    If trTeahList3.Visible Then '助教1
                        .Parameters.Add("TECH2", SqlDbType.Int).Value = If(OLessonTeah2bValue.Value = "", Convert.DBNull, Val(OLessonTeah2bValue.Value)) 'TECH2.Text
                    End If
                    .Parameters.Add("ROOM", SqlDbType.VarChar).Value = If(Room.Text = "", Convert.DBNull, Room.Text)
                    .Parameters.Add("PLANID", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
                    '.Parameters.Add("ISCOUNTHOURS", SqlDbType.VarChar).Value = IsCountHours.SelectedValue 'Y/N '是否計算排課時數
                    .Parameters.Add("COURID", SqlDbType.Int).Value = iCOURID
                    '.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                End With
        End Select

        'Select Case Type_str.Value
        '    Case "Insert"
        '    Case "Update"
        'End Select

        If Session("MySreach") Is Nothing Then
            Session("MySreach") = ViewState("MySreach")
        End If

        Hidsave1.Value = "Y"

        Dim strScript As String = ""
        strScript = ""
        strScript &= "<script language=""javascript"">" + vbCrLf
        strScript &= "document.getElementById('Hidsave1').value=""Y"";" + vbCrLf
        Select Case Type_str.Value
            Case "Insert"
                strScript &= "alert('課程設定儲存成功!!');" + vbCrLf
            Case "Update"
                strScript &= "alert('課程設定修改成功!!');" + vbCrLf
        End Select
        strScript &= "location.href='TC_01_005.aspx?ID=" & Request("ID") & "';" + vbCrLf
        'strScript &= "location.href='TC_01_005.aspx?ID='+document.getElementById('Re_ID').value;" + vbCrLf
        strScript &= "</script>"
        TIMS.RegisterStartupScript(Page, TIMS.xBlockName(), strScript)
        'Page.RegisterStartupScript("", strScript)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        courid.Value = ""
        TB_CourName.Text = ""
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        TB_career_id.Text = ""
        trainValue.Value = ""
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Classid_Hid.Value = ""
        Classid.Text = ""
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        If Session("MySreach") Is Nothing Then
            Session("MySreach") = ViewState("MySreach")
        End If
        'Response.Redirect("TC_01_005.aspx?ID=" & Request("ID"))
        Dim url1 As String = "TC_01_005.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

End Class
