Partial Class Course
    Inherits AuthBasePage

    Const cst_Class_CourseName As String = "Class_CourseName"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Me.modifytype.Value = Me.Request("type")
        Me.fieldname.Value = Me.Request("fieldname")
        Me.modifytype.Value = TIMS.ClearSQM(Me.modifytype.Value)
        Me.fieldname.Value = TIMS.ClearSQM(Me.fieldname.Value)
        'If Me.modifytype.Value <> "" Then Me.modifytype.Value = Trim(Me.modifytype.Value)
        'If Me.fieldname.Value <> "" Then Me.fieldname.Value = Trim(Me.fieldname.Value)

        If Not Page.IsPostBack Then
            btnClear.Attributes("onclick") = "returnValue('','');return false;"
            Call Create1()
        End If

    End Sub

    Sub Create1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Call GetCookie1()

        Call Search1()
    End Sub

    Function Get_TrainName(ByVal obj As ListControl, ByVal iLEVELS As Integer, _
                           ByVal parentTMID As String) As ListControl
        'levels 0@BusName 1@JobName 2@TrainName
        'Dim dt As DataTable = Nothing
        parentTMID = TIMS.ClearSQM(parentTMID)
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM KEY_TRAINTYPE"
        sql &= " WHERE LEVELS=@LEVELSXX"
        If parentTMID <> "" Then sql &= " AND PARENT=@PARENTXX" 'NUMBER 
        sql &= " ORDER BY TMID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("LEVELSXX", SqlDbType.Int).Value = iLEVELS
            If parentTMID <> "" Then .Parameters.Add("PARENTXX", SqlDbType.VarChar).Value = parentTMID
            dt.Load(.ExecuteReader())
        End With
        'dt = DbAccess.GetDataTable(sql, objconn)
        obj.Items.Clear()
        If dt Is Nothing Then Return obj
        If dt.Rows.Count = 0 Then Return obj
        With obj
            .DataSource = dt
            Select Case iLEVELS
                Case 0
                    .DataTextField = "BusName"
                Case 1
                    .DataTextField = "JobName"
                Case 2
                    .DataTextField = "TrainName"
            End Select
            .DataValueField = "TMID"
            .DataBind()
            If TypeOf obj Is DropDownList Then
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End If
            'If TypeOf obj Is CheckBoxList Then
            '    .Items.Insert(0, New ListItem("全部", ""))
            'End If
        End With

        Return obj
    End Function

    Sub AddTreeNodes(ByVal dr As DataRow, ByVal objTable As DataTable, ByVal objTreeView As TreeView, ByVal ParentNode As TreeNode)
        Dim NewNode As New TreeNode
        Dim drChild As DataRow
        Dim strFilter As String

        Dim h_param As New Hashtable
        h_param.Clear()
        h_param.Add("CourseName", Convert.ToString(dr("CourseName")))
        h_param.Add("CourseID", Convert.ToString(dr("CourseID")))
        h_param.Add("Classification1", Convert.ToString(dr("Classification1")))
        h_param.Add("Classification2", Convert.ToString(dr("Classification2")))
        NewNode.Text = TIMS.Get_CourseName(h_param)
        NewNode.NavigateUrl = "javascript:returnValue('" & dr("CourID") & "','" & dr("CourseName") & "');"

        If ParentNode Is Nothing Then
            objTreeView.Nodes.Add(NewNode)
        Else
            'ParentNode.Nodes.Add(NewNode)
            ParentNode.ChildNodes.Add(NewNode)
        End If

        '加入子節點
        Dim strRid As String = dr("RID") & "/"
        strFilter = "MainCourID = '" & dr("courID") & "'"        '先找出符合父節點 xxx\ 開頭的關係
        For Each drChild In objTable.Select(strFilter)
            Call AddTreeNodes(drChild, objTable, objTreeView, NewNode)
        Next
    End Sub

    Sub bus_Selected()
        '' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim mydv2 As New DataView(objdataset.Tables(0))
        If Me.bus.SelectedValue <> "" Then
            'mydv2.RowFilter = "levels='1' and [parent]='" & Me.bus.SelectedValue & "'"
            'Me.job.DataSource = mydv2
            'Me.job.DataTextField = "JobName"
            'Me.job.DataValueField = "TMID"
            'Me.job.DataBind()
            'Me.job.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

            job = Get_TrainName(job, 1, Me.bus.SelectedValue)
            Me.train.Items.Clear()
            Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Else
            Me.job.Items.Clear()
            Me.job.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Me.train.Items.Clear()
            Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End If
    End Sub

    Private Sub bus_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bus.SelectedIndexChanged
        Call bus_Selected()
    End Sub

    Sub job_Selected()
        'Dim mydv3 As New DataView(objdataset.Tables(0))
        If Me.job.SelectedValue <> "" Then
            'mydv3.RowFilter = "levels='2' and [parent]='" & Me.job.SelectedValue & "'"
            'Me.train.DataSource = mydv3
            'Me.train.DataTextField = "TrainName"
            'Me.train.DataValueField = "TMID"
            'Me.train.DataBind()
            'Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

            train = Get_TrainName(train, 2, Me.job.SelectedValue)
        Else
            Me.train.Items.Clear()
            Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End If
    End Sub

    Private Sub job_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles job.SelectedIndexChanged
        Call job_Selected()
    End Sub

    '查詢sql
    Sub Search1()
        'treeview begin
        'Me.TreeView1.Nodes.Clear()
        'Dim strFilter, objstr, classidstr As String
        'Dim objtable As DataTable
        ''Dim objadapter As SqlDataAdapter
        'Dim dr As DataRow
        'Dim first, final, classsidint As Integer

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Call SetCookie()

        Dim first As Integer = 0
        Dim final As Integer = 0
        Dim classidstr As String = Me.Classid.Text
        first = classidstr.IndexOf("(")
        final = classidstr.IndexOf(")")
        If first <> -1 And final <> -1 Then
            classidstr = classidstr.Substring(0, first)
        End If

        'Dim objstr As String = ""
        Dim sql As String = ""
        Dim iClasssIDX As Integer = 0
        If classidstr <> "" Then
            sql = ""
            sql &= " select CLSID "
            sql &= " from ID_Class "
            sql &= " where DistID = '" & sm.UserInfo.DistID & "' "
            sql &= " and Years='" & sm.UserInfo.Years & "' "
            sql &= " and ClassID ='" & classidstr & "'"
            Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
            If dt1.Rows.Count > 0 Then
                iClasssIDX = dt1.Rows(0)("CLSID")
            End If
            'iClasssIDX = DbAccess.ExecuteScalar(objstr, objconn)
        End If

        '預設查詢條件
        Dim def_Class_CourseName As String = ""
        If Convert.ToString(Session(cst_Class_CourseName)) <> "" Then
            def_Class_CourseName = Convert.ToString(Session(cst_Class_CourseName))
            'Session(cst_Class_CourseName) = Nothing
        End If

        Dim rqRID As String = TIMS.ClearSQM(Request("RID"))
        CourseName.Text = TIMS.ClearSQM(CourseName.Text)

        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " SELECT MAINCOURID CourID FROM COURSE_COURSEINFO" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " AND Valid='Y'" & vbCrLf
        sql &= " AND MAINCOURID IS NOT NULL" & vbCrLf '必為子層
        If rqRID = "" Then
            sql &= " AND RID='" & sm.UserInfo.RID & "'" & vbCrLf
        Else
            sql &= " AND RID='" & rqRID & "'" & vbCrLf
        End If
        If def_Class_CourseName <> "" Then
            sql &= " AND COURSENAME IN (" & def_Class_CourseName & ")" & vbCrLf
        Else
            If Me.Classification1.SelectedValue <> "" Then
                sql &= " and Classification1 = " & Me.Classification1.SelectedValue & vbCrLf
            End If
            If Me.Classification2.SelectedValue <> "" Then
                sql &= " and Classification2 = " & Me.Classification2.SelectedValue & vbCrLf
            End If
            If CourseName.Text <> "" Then
                sql &= " AND COURSENAME LIKE '%" & CourseName.Text & "%'" & vbCrLf
            End If
        End If
        sql &= " )" & vbCrLf

        'If CB_Main.Checked Then
        '    sql &= " ,WC2 AS (" & vbCrLf
        '    sql &= " SELECT COURID FROM COURSE_COURSEINFO" & vbCrLf
        '    sql &= " where 1=1" & vbCrLf
        '    sql &= " AND Valid='Y'" & vbCrLf
        '    sql &= " AND MAINCOURID IS NULL" & vbCrLf '必為父層
        '    If rqRID = "" Then
        '        sql &= " AND RID='" & sm.UserInfo.RID & "'" & vbCrLf
        '    Else
        '        sql &= " AND RID='" & rqRID & "'" & vbCrLf
        '    End If
        '    If Me.Classification1.SelectedValue <> "" Then
        '        sql &= " and Classification1 = " & Me.Classification1.SelectedValue & vbCrLf
        '    End If
        '    If Me.Classification2.SelectedValue <> "" Then
        '        sql &= " and Classification2 = " & Me.Classification2.SelectedValue & vbCrLf
        '    End If
        '    If CourseName.Text <> "" Then
        '        sql &= " AND COURSENAME LIKE '%" & CourseName.Text & "%'" & vbCrLf
        '    End If
        '    sql &= " )" & vbCrLf
        'End If

        sql &= " SELECT *" & vbCrLf
        sql &= " FROM COURSE_COURSEINFO" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " AND Valid='Y'" & vbCrLf
        If rqRID = "" Then
            sql &= " AND RID='" & sm.UserInfo.RID & "'" & vbCrLf
        Else
            sql &= " AND RID='" & rqRID & "'" & vbCrLf
        End If
        '預設查詢條件有值時
        If def_Class_CourseName <> "" Then
            'sql &= " AND COURSENAME IN (" & def_Class_CourseName & ")" & vbCrLf
            sql &= " AND (1!=1" & vbCrLf
            sql &= " OR COURSENAME IN (" & def_Class_CourseName & ")" & vbCrLf
            sql &= " OR CourID IN (SELECT CourID FROM WC1)" & vbCrLf
            'If CB_Main.Checked Then
            '    sql &= " OR MAINCOURID IN (SELECT COURID FROM WC2)" & vbCrLf
            'End If
            sql &= " )" & vbCrLf
        Else
            If Me.Classification1.SelectedValue <> "" Then
                sql &= " and Classification1 = " & Me.Classification1.SelectedValue & vbCrLf
            End If
            If Me.Classification2.SelectedValue <> "" Then
                sql &= " and Classification2 = " & Me.Classification2.SelectedValue & vbCrLf
            End If
            If CourseName.Text <> "" Then
                sql &= " and (1!=1" & vbCrLf
                sql &= " OR COURSENAME like '%" & CourseName.Text & "%'" & vbCrLf
                sql &= " OR CourID IN (SELECT CourID FROM WC1)" & vbCrLf
                'If CB_Main.Checked Then
                '    sql &= " OR MAINCOURID IN (SELECT COURID FROM WC2)" & vbCrLf
                'End If
                sql &= " )" & vbCrLf
                'sql &= " and (1!=1" & vbCrLf
                'sql &= " OR COURSENAME like '%" & CourseName.Text & "%'" & vbCrLf
                'sql &= " OR COURID IN (SELECT MAINCOURID FROM COURSE_COURSEINFO WHERE COURSENAME LIKE '%" & CourseName.Text & "%' AND MAINCOURID IS NOT NULL)" & vbCrLf
                'sql &= " )" & vbCrLf
                'sql &= " and (CourseName like '%" & CourseName.Text & "%' OR COURID IN (SELECT MAINCOURID FROM COURSE_COURSEINFO WHERE COURSENAME LIKE '%" & CourseName.Text & "%') or MainCourID in (select CourID from Course_CourseInfo where CourseName like '%" & CourseName.Text & "%'and MainCourID is Null))"
            End If
            If Me.train.SelectedValue <> "" Then
                sql &= " and TMID = " & Me.train.SelectedValue
            End If
            If iClasssIDX <> 0 AndAlso Me.Classid_Hid.Value <> "" Then
                sql &= " and CLSID = " & iClasssIDX
            End If
        End If
        sql &= " ORDER BY COURSEID"

        Dim objtable As DataTable = DbAccess.GetDataTable(sql, objconn)
        'objAdapter = New SqlDataAdapter(objstr, objconn)
        'objAdapter.Fill(objtable)

        TreeView1.Nodes.Clear()
        Dim strFilter As String = "MAINCOURID IS NULL"
        For Each dr As DataRow In objtable.Select(strFilter)
            AddTreeNodes(dr, objtable, Me.TreeView1, Nothing)
        Next
        'treeview end

        '清理查詢條件
        If Not Session(cst_Class_CourseName) Is Nothing Then
            Session(cst_Class_CourseName) = Nothing
        End If
    End Sub

    Private Sub But_Sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Sub.Click
        Call Search1()
    End Sub

    '保留搜尋值
    Sub SetCookie()
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim MyValue As String = ""

        TIMS.InsertCookieTable(Me, dt, da, "Course_Classification1", Classification1.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "Course_Classification2", Classification2.SelectedValue, False, objconn)

        MyValue = TIMS.ClearSQM(CourseName.Text) : If MyValue <> "" Then MyValue = Trim(MyValue)
        TIMS.InsertCookieTable(Me, dt, da, "Course_name", MyValue, False, objconn)
        MyValue = TIMS.ClearSQM(Classid.Text) : If MyValue <> "" Then MyValue = Trim(MyValue)
        TIMS.InsertCookieTable(Me, dt, da, "Course_Classid", MyValue, False, objconn)
        MyValue = TIMS.ClearSQM(Classid_Hid.Value) : If MyValue <> "" Then MyValue = Trim(MyValue)
        TIMS.InsertCookieTable(Me, dt, da, "Course_Classid_Hid", MyValue, False, objconn)

        TIMS.InsertCookieTable(Me, dt, da, "Course_bus", bus.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "Course_job", job.SelectedValue, False, objconn)
        TIMS.InsertCookieTable(Me, dt, da, "Course_train", train.SelectedValue, True, objconn)
    End Sub

    '取得 保留(搜尋)值 
    Sub GetCookie1()
        'Dim rst As Boolean = False '異常
        bus = Get_TrainName(bus, 0, "")
        Me.job.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Me.train.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

        Dim MyValue As String = ""
        Dim ff3 As String = ""
        Dim ff4 As String = ""

        Dim dtCK As DataTable = TIMS.GetCookieTable(Me, objconn)
        ff3 = "ItemName='Course_Classification1'"
        ff4 = "ItemName='Course_Classification2'"
        If dtCK.Select(ff3).Length > 0 Then
            Common.SetListItem(Classification1, dtCK.Select(ff3)(0)("ItemValue"))
        End If
        If dtCK.Select(ff4).Length > 0 Then
            Common.SetListItem(Classification2, dtCK.Select(ff4)(0)("ItemValue"))
        End If

        ff3 = "ItemName='Course_name'"
        If dtCK.Select(ff3).Length > 0 Then
            CourseName.Text = dtCK.Select(ff3)(0)("ItemValue")
        End If

        ff3 = "ItemName='Course_Classid'"
        ff4 = "ItemName='Course_Classid_Hid'"
        'If dt.Select(ff3).Length = 0 Then Exit Sub
        'If dt.Select(ff4).Length = 0 Then Exit Sub
        If dtCK.Select(ff3).Length > 0 Then
            Classid.Text = dtCK.Select(ff3)(0)("ItemValue")
        End If
        If dtCK.Select(ff4).Length > 0 Then
            Classid_Hid.Value = dtCK.Select(ff4)(0)("ItemValue")
        End If
        Classid.Text = TIMS.ClearSQM(Classid.Text)
        Classid_Hid.Value = TIMS.ClearSQM(Classid_Hid.Value)

        ff3 = "ItemName='Course_bus'"
        If dtCK.Select(ff3).Length > 0 Then
            MyValue = dtCK.Select(ff3)(0)("ItemValue")
            MyValue = TIMS.ClearSQM(MyValue)
            Common.SetListItem(bus, MyValue)
        End If

        If bus.SelectedIndex = 0 Then Exit Sub
        Call bus_Selected()

        ff3 = "ItemName='Course_job'"
        If dtCK.Select(ff3).Length > 0 Then
            MyValue = dtCK.Select(ff3)(0)("ItemValue")
            If MyValue <> "" Then MyValue = TIMS.ClearSQM(MyValue)
            If MyValue <> "" Then Common.SetListItem(job, MyValue)
        End If

        If job.SelectedIndex = 0 Then Exit Sub
        Call job_Selected()

        ff3 = "ItemName='Course_train'"
        'If dtCK.Select(ff3).Length = 0 Then Exit Sub
        If dtCK.Select(ff3).Length > 0 Then
            MyValue = dtCK.Select(ff3)(0)("ItemValue")
            If MyValue <> "" Then MyValue = TIMS.ClearSQM(MyValue)
            If MyValue <> "" Then Common.SetListItem(train, MyValue)
        End If
    End Sub

End Class
