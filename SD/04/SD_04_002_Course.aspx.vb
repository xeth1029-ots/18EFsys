'Imports Microsoft.Web.UI.WebControls

Partial Class SD_04_002_Course
    Inherits AuthBasePage

    'select * from COURSE_COURSEINFO  where orgid =167 and modifydate>='2019-02-20'
    Const Cst_Edit As String = "Edit"
    Const Cst_notepad As String = "notepad"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '在這裡放置使用者程式碼以初始化網頁
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        Call TIMS.ChkSession(Me, 0, sm)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        HidrqType.Value = TIMS.ClearSQM(Request("Type"))
        HidReqRID.Value = TIMS.ClearSQM(Request("RID"))

        If Not IsPostBack Then
            msg.Text = ""
            Button1.Attributes("onclick") = "GetCLSID();"
            Button3.Attributes("onclick") = "ClearClassid();"

            Classid.Text = ""
            Classid_Hid.Value = ""

            Call GetBusID(GetItemsDataTable())

            btnSaveCheckBox.Visible = False
            hid_Type1.Value = ""

            Select Case HidrqType.Value 'Request("Type").ToString
                Case Cst_Edit
                    hid_Type1.Value = "1" '1@Request("Type")=Cst_Edit ""@Request("Type")=Cst_notepad 
                Case Cst_notepad
                    btnSaveCheckBox.Visible = True
                    'hid_Type1.Value = ""
            End Select

            Dim ItemValue As String = ""
            Dim dt As DataTable
            dt = TIMS.GetCookieTable(Me, objconn)
            ItemValue = Get_dtItemValue(dt, "SD04_Classification1")

            If ItemValue <> "" Then
                CourseID.Text = Get_dtItemValue(dt, "SD04_CourseID")
                'CLASSIFICATION1 1:學科/2術科
                ItemValue = Get_dtItemValue(dt, "SD04_Classification1") : Common.SetListItem(Classification1, ItemValue)
                ItemValue = Get_dtItemValue(dt, "SD04_Classification2") : Common.SetListItem(Classification2, ItemValue)
                Classid.Text = Get_dtItemValue(dt, "SD04_Classid")
                Classid_Hid.Value = Get_dtItemValue(dt, "SD04_Classid_Hid")
                ItemValue = Get_dtItemValue(dt, "SD04_BusID") : Common.SetListItem(BusID, ItemValue)
                ItemValue = Get_dtItemValue(dt, "SD04_JobID") : Common.SetListItem(JobID, ItemValue)
                ItemValue = Get_dtItemValue(dt, "SD04_TrainID") : Common.SetListItem(TrainID, ItemValue)
                Call GetCourseData()
            End If

            If Session("Class_CourseName") IsNot Nothing Then Session("Class_CourseName") = Nothing
            'If Request("Used") = "1" And Request("MySearch") <> "0" Then GetCourseData(Request("MySearch"))
        End If

    End Sub

    Function Get_dtItemValue(ByVal dt As DataTable, ByVal ItemName As String) As String
        Dim rst As String = ""
        Dim ff3 As String = "ItemName='" & TIMS.ClearSQM(ItemName) & "'"
        If dt.Select(ff3).Length > 0 Then rst = TIMS.ClearSQM(dt.Select(ff3)(0)("ItemValue"))
        Return rst
    End Function

    Function GetItemsDataTable() As DataTable
        Dim sql As String = " SELECT * FROM dbo.KEY_TRAINTYPE ORDER BY 1 "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        dt.TableName = "Key_TrainType"
        Return dt
    End Function

    Sub GetBusID(ByVal dt As DataTable)
        Dim dv As New DataView
        dv.Table = dt
        dv.RowFilter = "BusID Is Not NULL"
        With BusID
            .DataSource = dv
            .DataTextField = "BusName"
            .DataValueField = "TMID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        JobID.Items.Clear()
        TrainID.Items.Clear()

        JobID.Items.Insert(0, New ListItem("===請選擇行業別代碼===", ""))
        TrainID.Items.Insert(0, New ListItem("===請選擇行業別代碼===", ""))
    End Sub

    Sub GetJobID(ByVal dt As DataTable)
        Dim v_BusID As String = TIMS.GetListValue(BusID) 'BusID.SelectedValue 
        Dim dv As New DataView
        dv.Table = dt
        dv.RowFilter = "JobID IS NOT NULL AND [Parent] = '" & v_BusID & "' "
        JobID.Items.Clear()

        With JobID
            .DataSource = dv
            .DataTextField = "JobName"
            .DataValueField = "TMID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

        TrainID.Items.Clear()
        TrainID.Items.Insert(0, New ListItem("===請選擇職業分類===", ""))
    End Sub

    Sub GetTrainID(ByVal dt As DataTable)
        Dim v_JobID As String = TIMS.GetListValue(JobID) 'JobID.SelectedValue 
        Dim dv As New DataView
        dv.Table = dt
        dv.RowFilter = "TrainID IS NOT NULL AND [Parent] = '" & v_JobID & "' "
        TrainID.Items.Clear()

        With TrainID
            .DataSource = dv
            .DataTextField = "TrainName"
            .DataValueField = "TMID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
    End Sub

    Private Sub BusID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BusID.SelectedIndexChanged
        If BusID.SelectedIndex = 0 Then
            JobID.Items.Clear()
            TrainID.Items.Clear()
            JobID.Items.Insert(0, New ListItem("===請選擇行業別代碼===", ""))
            TrainID.Items.Insert(0, New ListItem("===請選擇行業別代碼===", ""))
        Else
            GetJobID(GetItemsDataTable())
        End If
    End Sub

    Private Sub JobID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles JobID.SelectedIndexChanged
        If JobID.SelectedIndex = 0 Then
            TrainID.Items.Clear()
            TrainID.Items.Insert(0, New ListItem("===請選擇職業分類===", ""))
        Else
            GetTrainID(GetItemsDataTable())
        End If
    End Sub

    '查詢
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call GetCourseData()
        If Session("Class_CourseName") IsNot Nothing Then Session("Class_CourseName") = Nothing  '清理查詢條件
    End Sub

    '查詢所有課程
    Sub GetCourseData()
        TreeView1.Nodes.Clear() '清理

        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        Dim str_Search1 As String = ""
        Dim str_CourseName As String = ""

        If HidReqRID.Value = "" Then Exit Sub

        Classid_Hid.Value = Trim(Classid_Hid.Value).Replace(vbCrLf, "")
        Classid.Text = Trim(Classid.Text).Replace(vbCrLf, "")
        Classid_Hid.Value = TIMS.ClearSQM(Classid_Hid.Value)
        Classid.Text = TIMS.ClearSQM(Classid.Text)

        Dim v_Classification1 As String = TIMS.GetListValue(Classification1)
        Dim v_Classification2 As String = TIMS.GetListValue(Classification2)
        'v_Classification1 = If(v_Classification1 = "X", "", v_Classification1)
        'v_Classification2 = If(v_Classification2 = "X", "", v_Classification2)
        Dim v_TrainID As String = TIMS.GetListValue(TrainID)

        Me.ViewState("Classid_Hid") = TIMS.ClearSQM(Classid_Hid.Value) '.Replace("'", "''")
        Me.ViewState("Classid") = TIMS.ClearSQM(Classid.Text) '.Replace("'", "''")
        Me.ViewState("CourseID") = TIMS.ClearSQM(CourseID.Text) '.Replace("'", "''")
        Me.ViewState("CourseName") = TIMS.ClearSQM(CourseName.Text) '.Replace("'", "''")
        Me.ViewState("Classification1") = v_Classification1 '"" 'CLASSIFICATION1 1:學科/2術科 'If Classification1.SelectedIndex <> 0 Then Me.ViewState("Classification1") = TIMS.ClearSQM(Classification1.SelectedValue) '.Replace("'", "''")
        Me.ViewState("Classification2") = v_Classification2 '"" 'If Classification2.SelectedIndex <> 0 Then Me.ViewState("Classification2") = TIMS.ClearSQM(Classification2.SelectedValue) '.Replace("'", "''")
        Me.ViewState("TrainID") = v_TrainID '""
        'If TrainID.SelectedIndex <> 0 Then Me.ViewState("TrainID") = TIMS.ClearSQM(TrainID.SelectedValue) '.Replace("'", "''")
        'Me.ViewState("Request_RID") = TIMS.ClearSQM(Request("RID")) '.Replace("'", "''")

        str_Search1 = ""
        If Me.ViewState("CourseID") <> "" Then str_Search1 &= " AND a.CourseID LIKE '%" & Me.ViewState("CourseID") & "%' " & vbCrLf

        '課程名稱
        Dim sMAINCOURID As String = ""
        sMAINCOURID = ""
        If Me.ViewState("CourseName") <> "" Then
            sql = ""
            sql &= " SELECT DISTINCT x.MAINCOURID "
            sql &= " FROM COURSE_COURSEINFO x "
            sql &= " WHERE 1=1 "
            sql &= " AND x.COURSENAME LIKE '%" & Me.ViewState("CourseName") & "%' "
            sql &= " AND x.RID = '" & HidReqRID.Value & "' "
            sql &= " AND x.MAINCOURID IS NOT NULL "
            Dim dtX As DataTable = DbAccess.GetDataTable(sql, objconn)
            For Each drX As DataRow In dtX.Rows
                If sMAINCOURID <> "" Then sMAINCOURID &= ","
                sMAINCOURID &= Convert.ToString(drX("MAINCOURID"))
            Next
        End If
        '課程名稱
        If Me.ViewState("CourseName") <> "" Then
            str_Search1 += " AND (1!=1" & vbCrLf
            str_Search1 += " OR a.CourseName LIKE '%" & Me.ViewState("CourseName") & "%' " & vbCrLf
            str_Search1 += " OR c2.CourseName LIKE '%" & Me.ViewState("CourseName") & "%' " & vbCrLf '搜尋 主課程名稱
            If sMAINCOURID <> "" Then str_Search1 += " OR a.COURID IN (" & sMAINCOURID & ") " & vbCrLf  '為主課程的子課程名稱
            str_Search1 += " )"
        End If
        If Me.ViewState("Classification1") <> "" Then str_Search1 += " AND a.Classification1 = '" & Me.ViewState("Classification1") & "' " & vbCrLf 'CLASSIFICATION1 1:學科/2術科
        If Me.ViewState("Classification2") <> "" Then str_Search1 += " AND a.Classification2 = '" & Me.ViewState("Classification2") & "' " & vbCrLf
        If Me.ViewState("Classid") <> "" AndAlso Me.ViewState("Classid_Hid") <> "" Then
            str_Search1 += " AND a.CLSID = '" & Me.ViewState("Classid_Hid") & "' " & vbCrLf
        Else
            '若有一值為空,清空
            Classid_Hid.Value = ""
        End If
        If Me.ViewState("TrainID") <> "" Then str_Search1 += " and a.TMID='" & Me.ViewState("TrainID") & "'" & vbCrLf

        SetCookie()

        '預設查詢條件
        str_CourseName = ""
        If Session("Class_CourseName") IsNot Nothing Then
            If Convert.ToString(Session("Class_CourseName")) <> "" Then
                str_CourseName = Convert.ToString(Session("Class_CourseName"))
                'Session("Class_CourseName") = Nothing
            End If
        End If

        Select Case HidrqType.Value
            Case Cst_notepad
                sql = "" & vbCrLf
                sql &= " SELECT a.* " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech1) TechName1 " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech2) TechName2 " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech3) TechName3 " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech4) TechName4 " & vbCrLf
                sql &= " ,CASE WHEN a2.CourID IS NOT NULL THEN 1 " & vbCrLf
                '預設查詢條件有值時(display)
                If str_CourseName <> "" Then
                    sql &= " WHEN a.CourseName IN (" & str_CourseName & ") THEN 1 " & vbCrLf
                End If
                sql &= " ELSE 0 END Selected " & vbCrLf
                sql &= " FROM COURSE_COURSEINFO a " & vbCrLf
                'HidReqRID.Value
                'sql &= " LEFT JOIN Course_CourseInfo c2 on c2.RID ='" & Me.ViewState("Request_RID") & "' and c2.MAINCOURID is null AND c2.COURID=a.MAINCOURID" & vbCrLf
                sql &= " LEFT JOIN COURSE_COURSEINFO c2 ON c2.RID = '" & HidReqRID.Value & "' AND c2.MAINCOURID IS NULL AND c2.COURID = a.MAINCOURID " & vbCrLf
                '編輯選擇常用清單
                sql &= " LEFT JOIN AUTH_COURSEINFO a2 ON a.CourID = a2.CourID AND a2.Account = '" & sm.UserInfo.UserID & "' " & vbCrLf
                'sql &= " LEFT JOIN Teach_TeacherInfo b ON a.Tech1=b.TechID " & vbCrLf
                'sql &= " LEFT JOIN Teach_TeacherInfo c ON a.Tech2=c.TechID " & vbCrLf
                'sql &= " LEFT JOIN Teach_TeacherInfo c3 ON a.Tech3=c3.TechID " & vbCrLf
                sql &= " WHERE 1=1 " & vbCrLf
                sql &= " AND a.Valid = 'Y' " & vbCrLf
                'sql &= " AND a.RID='" & Me.ViewState("Request_RID") & "'" & vbCrLf
                sql &= " AND a.RID = '" & HidReqRID.Value & "' " & vbCrLf
                sql &= str_Search1 & vbCrLf
            Case Else
                sql = "" & vbCrLf
                sql &= " SELECT a.* " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech1) TechName1 " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech2) TechName2 " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech3) TechName3 " & vbCrLf
                sql &= " ,dbo.FN_GET_TEACHCNAME(a.Tech4) TechName4 " & vbCrLf
                sql &= " FROM COURSE_COURSEINFO a " & vbCrLf
                sql &= " LEFT JOIN COURSE_COURSEINFO c2 ON c2.RID = '" & HidReqRID.Value & "' AND c2.MAINCOURID IS NULL AND c2.COURID = a.MAINCOURID " & vbCrLf
                'sql &= " LEFT JOIN Teach_TeacherInfo b ON a.Tech1=b.TechID " & vbCrLf
                'sql &= " LEFT JOIN Teach_TeacherInfo c ON a.Tech2=c.TechID " & vbCrLf
                'sql &= " LEFT JOIN Teach_TeacherInfo c3 ON a.Tech3=c3.TechID " & vbCrLf
                sql &= " WHERE 1=1 " & vbCrLf
                sql &= " AND a.Valid = 'Y' " & vbCrLf
                sql &= " AND a.RID = '" & HidReqRID.Value & "' " & vbCrLf
                '預設查詢條件有值時(where)
                If str_CourseName <> "" Then
                    sql &= " AND a.CourseName IN (" & str_CourseName & ") " & vbCrLf
                Else
                    sql &= str_Search1 & vbCrLf
                End If
        End Select

        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            ' ==== 追錯email  ==== 
            Dim strErrmsg As String = ""
            strErrmsg += "/*追錯Email,sql:*/" & vbCrLf
            strErrmsg += sql & vbCrLf
            'For i As Integer = 0 To daObj.SelectCommand.Parameters.Count - 1
            '    Dim xName As String = daObj.SelectCommand.Parameters(i).ParameterName
            '    Dim xValue As String = Convert.ToString(daObj.SelectCommand.Parameters(i).Value)
            '    strErrmsg += " AND " & xName & "='" & xValue & "'" & vbCrLf
            'Next
            strErrmsg &= "/*ex.ToString*/" & vbCrLf
            strErrmsg &= ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            ' ==== 追錯email  ==== 

            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Common.MessageBox2(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End Try

        trTreeView1.Visible = False
        trCheckBoxList1.Visible = False

        Select Case HidrqType.Value
            Case Cst_notepad
                trCheckBoxList1.Visible = True
                msg2.Text = "查無資料!"
                If dt.Rows.Count > 0 Then
                    msg2.Text = ""
                    Call AddCheckBoxList(CheckBoxList1, dt)
                    Call SaveCheckBoxList(CheckBoxList1)
                End If
            Case Else
                trTreeView1.Visible = True
                msg.Text = "查無資料!"
                If dt.Rows.Count > 0 Then
                    msg.Text = ""
                    Call AddTreeView(dt)
                End If
        End Select

        ''清理查詢條件
        'If Not Session("Class_CourseName") Is Nothing Then
        '    Session("Class_CourseName") = Nothing
        'End If
    End Sub

    Sub SetCookie()
        Dim v_Classification1 As String = TIMS.GetListValue(Classification1)
        Dim v_Classification2 As String = TIMS.GetListValue(Classification2)
        Dim v_BusID As String = TIMS.GetListValue(BusID)
        Dim v_JobID As String = TIMS.GetListValue(JobID)
        Dim v_TrainID As String = TIMS.GetListValue(TrainID)

        'v_Classification1 = If(v_Classification1 = "X", "", v_Classification1)
        'v_Classification2 = If(v_Classification2 = "X", "", v_Classification2)

        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        TIMS.InsertCookieTable(Me, dt, da, "SD04_CourseID", CourseID.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_Classification1", v_Classification1, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_Classification2", v_Classification2, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_Classid", Classid.Text, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_Classid_Hid", Classid_Hid.Value, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_BusID", v_BusID, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_JobID", v_JobID, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "SD04_TrainID", v_TrainID, True, objconn)
    End Sub

    Sub AddCheckBoxList(ByRef objCBL1 As CheckBoxList, ByVal dt3 As DataTable)
        Dim z As Integer = 0
        Dim strCourseName As String = ""

        objCBL1.Items.Clear()
        For Each dr As DataRow In dt3.Rows
            Dim h_param As New Hashtable
            h_param.Clear()
            h_param.Add("CourseName", Convert.ToString(dr("CourseName")))
            h_param.Add("CourseID", Convert.ToString(dr("CourseID")))
            h_param.Add("Classification1", Convert.ToString(dr("Classification1")))
            h_param.Add("Classification2", Convert.ToString(dr("Classification2")))

            strCourseName = TIMS.Get_CourseName(h_param)
            objCBL1.Items.Add(strCourseName)
            objCBL1.Items.Item(z).Value = Convert.ToString(dr("CourID"))
            objCBL1.Items.Item(z).Selected = If(Convert.ToString(dr("Selected")) = "1", True, False)

            z += 1
        Next
    End Sub

    Sub SaveCheckBoxList(ByRef objCBL1 As CheckBoxList)
        Dim dr As DataRow
        Dim dt As DataTable
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = ""

        sql = " SELECT * FROM AUTH_COURSEINFO WHERE Account = '" & sm.UserInfo.UserID & "' "
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For i As Integer = 0 To objCBL1.Items.Count - 1
            If objCBL1.Items.Item(i).Selected Then
                '無則增
                If dt.Select("CourID=" & objCBL1.Items.Item(i).Value).Length = 0 Then
                    dr = dt.NewRow()
                    dt.Rows.Add(dr)
                    dr("Account") = sm.UserInfo.UserID
                    dr("CourID") = objCBL1.Items.Item(i).Value
                End If
            Else
                '有則刪
                If dt.Select("CourID=" & objCBL1.Items.Item(i).Value).Length > 0 Then
                    dr = dt.Select("CourID=" & objCBL1.Items.Item(i).Value)(0)
                    dr.Delete()
                End If
            End If
        Next
        DbAccess.UpdateDataTable(dt, da)

        sql = "" & vbCrLf
        sql &= " SELECT DISTINCT a.MainCourID " & vbCrLf
        sql &= " FROM Course_CourseInfo a " & vbCrLf
        sql &= " JOIN Auth_CourseInfo a2 ON a.CourID = a2.CourID " & vbCrLf
        'sql += " LEFT JOIN Teach_TeacherInfo b ON a.Tech1 = b.TechID " & vbCrLf
        'sql += " LEFT JOIN Teach_TeacherInfo c ON a.Tech2 = c.TechID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.MainCourID IS NOT NULL " & vbCrLf
        sql &= " AND a2.Account = '" & sm.UserInfo.UserID & "' " & vbCrLf
        Dim dt2 As DataTable
        dt2 = DbAccess.GetDataTable(sql, objconn)

        sql = " SELECT * FROM Auth_CourseInfo WHERE Account = '" & sm.UserInfo.UserID & "' "
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For i As Integer = 0 To dt2.Rows.Count - 1
            Dim dr2 As DataRow = dt2.Rows(i)
            '無則增
            If dt.Select("CourID=" & dr2("MainCourID").ToString).Length = 0 Then
                dr = dt.NewRow()
                dt.Rows.Add(dr)
                dr("Account") = sm.UserInfo.UserID
                dr("CourID") = Convert.ToString(dr2("MainCourID"))
            End If
        Next

        DbAccess.UpdateDataTable(dt, da)
    End Sub

    Sub AddTreeView(ByVal dt As DataTable, Optional ByVal ParentsNode As TreeNode = Nothing, Optional ByVal MainCourID As String = "")
        'Dim dr As DataRow
        Dim RowFilterStr As String = ""
        If ParentsNode Is Nothing Then
            RowFilterStr = "MainCourID is NULL"
        Else
            RowFilterStr = "MainCourID ='" & MainCourID & "'"
        End If

        For Each dr As DataRow In dt.Select(RowFilterStr)
            Dim h_param As New Hashtable
            h_param.Clear()
            h_param.Add("CourseName", Convert.ToString(dr("CourseName")))
            h_param.Add("CourseID", Convert.ToString(dr("CourseID")))
            h_param.Add("Classification1", Convert.ToString(dr("Classification1")))
            h_param.Add("Classification2", Convert.ToString(dr("Classification2")))

            Dim NewNode As New TreeNode
            'NewNode.Text = dr("CourseName")
            NewNode.Text = TIMS.Get_CourseName(h_param)
            Dim sUrl As String = ""
            sUrl = "'" & dr("CourseName") & "(" & dr("CourseID") & ")" & "'"
            sUrl &= ",'" & dr("CourID") & "'"
            sUrl &= ",'" & Convert.ToString(dr("Tech1")) & "','" & Convert.ToString(dr("TechName1")) & "'"
            sUrl &= ",'" & Convert.ToString(dr("Tech2")) & "','" & Convert.ToString(dr("TechName2")) & "'"
            sUrl &= ",'" & Convert.ToString(dr("Tech3")) & "','" & Convert.ToString(dr("TechName3")) & "'"
            sUrl &= ",'" & Convert.ToString(dr("Room")) & "'"

            'NewNode.NavigateUrl = "javascript:returnValue('" & dr("CourseName") & "(" & dr("CourseID") & ")" & "','" & dr("CourID") & "','" & dr("Tech1").ToString & "','" & dr("TechName1").ToString & "','" & dr("Tech2").ToString & "','" & dr("TechName2").ToString & "','" & dr("Room").ToString & "');"
            Dim sUrl2 As String = ""
            sUrl2 = "javascript:returnValue(" & sUrl & ");"

            '68:照顧服務員自訓自用訓練計畫 
            If TIMS.Cst_TPlanID68.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                sUrl = ""
                sUrl &= "'" & dr("CourseName") & "(" & dr("CourseID") & ")" & "'"
                sUrl &= ",'" & dr("CourID") & "'"
                sUrl &= ",'" & Convert.ToString(dr("Tech1")) & "','" & Convert.ToString(dr("TechName1")) & "'"
                sUrl &= ",'" & Convert.ToString(dr("Tech2")) & "','" & Convert.ToString(dr("TechName2")) & "'"
                sUrl &= ",'" & Convert.ToString(dr("Room")) & "'"
                sUrl &= ",'" & Convert.ToString(dr("Classification1")) & "'"

                sUrl2 = "javascript:returnValue68(" & sUrl & ");"
            End If
            NewNode.NavigateUrl = sUrl2

            If ParentsNode Is Nothing Then
                TreeView1.Nodes.Add(NewNode)
                AddTreeView(dt, NewNode, dr("CourID").ToString)
            Else
                'ParentsNode.Nodes.Add(NewNode)
                ParentsNode.ChildNodes.Add(NewNode)
            End If
        Next
    End Sub

    Private Sub btnSaveCheckBox_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSaveCheckBox.Click
        Call SaveCheckBoxList(CheckBoxList1)
        Common.MessageBox(Me, "儲存完畢!!")
    End Sub
End Class