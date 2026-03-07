Partial Class TC_01_003
    Inherits AuthBasePage

    Const cst_sess_sch1_txt As String = "_searchTc0103"

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '依轄區、計畫、年度，顯示班別代碼 by AMU 2010
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DG_Class
        '分頁設定 End
        iPYNum = TIMS.sUtl_GetPYNum(Me)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Me.LabTMID.Text = "訓練業別"
        End If

        check_del.Value = "1"
        check_mod.Value = "1"

        If Not IsPostBack Then
            cCreate1()
        End If
        cCreate2()

    End Sub

    Sub cCreate1()
        ddlYears = TIMS.Get_Years(ddlYears, objconn)
        ddlYears.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(ddlYears, "") 'sm.UserInfo.Years)

        ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)

        'objstr = "select * from Key_Plan"
        Dim objstr As String = ""
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            objstr = "SELECT * FROM Key_Plan WHERE (Clsyear IS NULL OR Clsyear > '" & Year(Now) & "') AND TPlanID='" & Convert.ToString(sm.UserInfo.TPlanID) & "'"
        Else
            objstr = "SELECT * FROM Key_Plan WHERE (Clsyear IS NULL OR Clsyear > '" & Year(Now) & "') AND TPlanID NOT IN (" & TIMS.Cst_NotTPlanID2 & ") "
        End If

        Dim objtable As DataTable
        objtable = DbAccess.GetDataTable(objstr, objconn)
        With Plan_List
            .Items.Clear()
            .DataSource = objtable
            .DataValueField = "TPlanID"
            .DataTextField = "PlanName"
            .DataBind()
        End With
        'Me.Plan_List.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        If objtable.Rows.Count > 0 Then
            For i As Integer = 0 To objtable.Rows.Count - 1
                Dim objdr As DataRow = objtable.Rows(i)
                If objdr("TPlanID") = sm.UserInfo.TPlanID Then
                    Common.SetListItem(Plan_List, objdr("TPlanID"))
                    'Exit Sub '以下狀況不執行了
                    'Me.Plan_List.Items.Insert(0, New ListItem(objdr("PlanName"), objdr("TPlanID")))
                End If
            Next
        End If

    End Sub

    Sub cCreate2()
        '取得訓練計畫2005/3/21
        'Dim Sqlstr As String = "select TPlanID  from ID_Plan where PlanID=" & sm.UserInfo.PlanID & ""
        'TPlanID.Value = DbAccess.ExecuteScalar(Sqlstr, objconn)
        If Session(cst_sess_sch1_txt) Is Nothing Then Return
        Dim myValue1 As String = Session(cst_sess_sch1_txt)
        Session(cst_sess_sch1_txt) = Nothing

        Common.SetListItem(Plan_List, TIMS.GetMyValue(myValue1, "Plan_List"))
        Common.SetListItem(ddlYears, TIMS.GetMyValue(myValue1, "ddlYears"))
        Common.SetListItem(ddlDISTID, TIMS.GetMyValue(myValue1, "ddlDistID"))
        'Common.SetListItem(planlist, TIMS.GetMyValue(Session("_search"), "planlist"))
        'Plan_List.SelectedValue = TIMS.GetMyValue(Session("_search"), "Plan_List")
        TB_classid.Text = UCase(TIMS.GetMyValue(myValue1, "TB_classid"))
        TB_ClassName.Text = TIMS.GetMyValue(myValue1, "TB_ClassName")
        TB_career_id.Text = TIMS.GetMyValue(myValue1, "TB_career_id")
        trainValue.Value = TIMS.GetMyValue(myValue1, "trainValue")
        jobValue.Value = TIMS.GetMyValue(myValue1, "jobValue")
        txtCJOB_NAME.Text = TIMS.GetMyValue(myValue1, "txtCJOB_NAME")
        cjobValue.Value = TIMS.GetMyValue(myValue1, "cjobValue")
        PageControler1.PageIndex = 0
        'PageControler1.PageIndex = TIMS.GetMyValue(myValue1, "PageIndex")
        Dim MyValue As String = TIMS.GetMyValue(myValue1, "PageIndex")
        If MyValue <> "" AndAlso IsNumeric(MyValue) Then
            MyValue = CInt(MyValue)
            PageControler1.PageIndex = MyValue
        End If
        If TIMS.GetMyValue(myValue1, "submit") = "1" Then
            'bt_search_Click(sender, e)
            Call Search1()
        End If

        'Button1.Attributes("onclick") = "window.open('TC_01_003_import.aspx?bt_search=bt_search&TPlanid=" & Plan_List.SelectedValue & "','','width=750,height=550,location=0,status=0,menubar=0,scrollbars=0,resizable=0');"
        'Button1.Attributes("oNclick") = "window.open('tc_01_001_IMPORT.ASPX','','WIDTH = 250,HEIGHT = 150,LOCATION = 0,STATUS');"
    End Sub

    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DG_Class)
        Me.Panel.Visible = True

        'Dim sqlAdapter As SqlDataAdapter
        'Dim dtOrgInfo As DataTable

        TB_classid.Text = TIMS.ClearSQM(UCase(TB_classid.Text))
        TB_ClassName.Text = TIMS.ClearSQM(TB_ClassName.Text)
        txtCJOB_NAME.Text = TIMS.ClearSQM(txtCJOB_NAME.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List)
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)

        Dim parms As Hashtable = New Hashtable()
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " SELECT a.CLSID" & vbCrLf
        sqlstr &= " ,a.ClassID" & vbCrLf
        sqlstr &= " ,a.ClassName" & vbCrLf
        sqlstr &= " ,b.TPlanID" & vbCrLf
        sqlstr &= " ,b.PlanName " & vbCrLf
        sqlstr &= " ,a.TMID" & vbCrLf
        sqlstr &= " ,s.CJOB_NAME" & vbCrLf
        sqlstr &= " ,ISNULL(c.TrainName, c.jobName) TrainName " & vbCrLf
        sqlstr &= " ,'['+c3.GCODE2+']'+c3.CNAME CNAME3 " & vbCrLf
        sqlstr &= " ,a.YEARS" & vbCrLf
        sqlstr &= " ,a.DISTID" & vbCrLf
        sqlstr &= " FROM ID_Class a " & vbCrLf
        sqlstr &= " JOIN Key_Plan b ON a.TPlanID = b.TPlanID " & vbCrLf
        sqlstr &= " LEFT JOIN KEY_TRAINTYPE c ON a.TMID = c.TMID " & vbCrLf
        sqlstr &= " LEFT JOIN V_GOVCLASSCAST3 c3 ON c.TMID = c3.TMID " & vbCrLf
        sqlstr &= " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY = a.CJOB_UNKEY "
        sqlstr &= " WHERE 1=1 " & vbCrLf
        sqlstr &= " AND a.TPlanID=@TPlanID " & vbCrLf
        parms.Add("@TPlanID", v_Plan_List)
        If v_ddlYears <> "" Then
            sqlstr &= " AND a.YEARS=@YEARS " & vbCrLf
            parms.Add("@YEARS", v_ddlYears)
        End If
        If v_ddlDistID <> "" Then
            sqlstr &= " AND a.DISTID=@DISTID " & vbCrLf
            parms.Add("@DISTID", v_ddlDistID)
        End If
        If TB_classid.Text <> "" Then
            sqlstr &= " AND a.ClassID LIKE @ClassID " & vbCrLf
            parms.Add("ClassID", "%" & TB_classid.Text & "%")
        End If
        If TB_ClassName.Text <> "" Then
            sqlstr &= " AND a.ClassName LIKE @ClassName " & vbCrLf
            parms.Add("ClassName", "%" & TB_ClassName.Text & "%")
        End If
        If txtCJOB_NAME.Text <> "" AndAlso cjobValue.Value <> "" Then
            sqlstr &= " AND s.CJOB_UNKEY = @CJOB_UNKEY " & vbCrLf
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If iPYNum >= 3 Then
                If trainValue.Value <> "" Then
                    sqlstr &= " AND a.TMID = @TMID " & vbCrLf
                    parms.Add("TMID", trainValue.Value)
                End If
            Else
                'Me.LabTMID.Text = "訓練業別"
                If jobValue.Value <> "" Then
                    'sqlstr &= " AND (a.TMID = " & jobValue.Value & " " & vbCrLf
                    sqlstr &= " AND (a.TMID = @TMID " & vbCrLf
                    sqlstr &= " OR a.TMID IN ( " & vbCrLf
                    sqlstr &= " SELECT TMID FROM Key_TrainType WHERE parent IN ( " & vbCrLf '職類別
                    sqlstr &= " SELECT TMID FROM Key_TrainType WHERE parent IN ( " & vbCrLf '業別
                    sqlstr &= " SELECT TMID FROM Key_TrainType WHERE busid ='G') " & vbCrLf '產業人才投資方案類
                    'sqlstr &= " AND tmid = " & jobValue.Value & " " & vbCrLf
                    sqlstr &= " AND tmid = @TMID " & vbCrLf
                    sqlstr &= " )))" & vbCrLf

                    parms.Add("TMID", jobValue.Value)
                End If
            End If
        Else
            If trainValue.Value <> "" Then
                sqlstr &= " AND a.TMID = @TMID " & vbCrLf
                parms.Add("TMID", trainValue.Value)
            End If
        End If

        'sqlstr += " order by a.ClassID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sqlstr, objconn, parms)

        If dt.Rows.Count > 0 Then
            For Each dr1 As DataRow In dt.Rows
                If Convert.ToString(dr1("CNAME3")) <> "" Then
                    dr1("TrainName") = Convert.ToString(dr1("CNAME3"))
                End If
            Next
            'dt.AcceptChanges()
        End If

        msg.Text = "查無資料!!"
        Panel.Visible = False
        DG_Class.Visible = False

        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Panel.Visible = True
            DG_Class.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "CLSID"
            PageControler1.Sort = "ClassID"
            PageControler1.ControlerLoad()

        End If
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call Search1()
    End Sub

    Sub KeepSearch()
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List)
        Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)
        Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)

        Dim myValue1 As String = ""
        myValue1 = "Plan_List=" & v_Plan_List 'Plan_List.SelectedValue
        myValue1 += "&ddlYears=" & v_ddlYears
        myValue1 += "&ddlDistID=" & v_ddlDistID
        myValue1 += "&TB_classid=" & TB_classid.Text
        myValue1 += "&TB_ClassName=" & TB_ClassName.Text
        myValue1 += "&TB_career_id=" & TB_career_id.Text
        myValue1 += "&trainValue=" & trainValue.Value
        myValue1 += "&jobValue=" & jobValue.Value
        myValue1 += "&txtCJOB_NAME=" & txtCJOB_NAME.Text
        myValue1 += "&cjobValue=" & cjobValue.Value
        myValue1 += "&PageIndex=" & DG_Class.CurrentPageIndex + 1
        myValue1 += If(DG_Class.Visible, "&submit=1", "&submit=0")

        Session(cst_sess_sch1_txt) = myValue1
    End Sub

    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        KeepSearch()
        TB_classid.Text = UCase(TB_classid.Text)
        'Response.Redirect("TC_01_003_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "")
        '20100208 按新增時代查詢之 班別代碼 & 班別名稱
        Dim url1 As String = "TC_01_003_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&ClassID=" & TB_classid.Text & "&ClassName=" & TB_ClassName.Text & ""
        TIMS.Utl_Redirect(Me, objconn, url1)
        'Response.Redirect("TC_01_003_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&ClassID=" & TB_classid.Text & "&ClassName=" & TB_ClassName.Text & "")
    End Sub

    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click
        KeepSearch()
        Dim url1 As String = "TC_01_003_print.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
        'Response.Redirect("TC_01_003_print.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub DG_Class_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Class.ItemDataBound
        Dim dr As DataRowView
        'Dim is_parent As String
        dr = e.Item.DataItem
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DG_Class.PageSize * DG_Class.CurrentPageIndex

            'Dim but_edit, but_del, but_share, but_copy As Button
            Dim lbtEdit, lbtDel, lbtCopy As LinkButton

            lbtEdit = e.Item.Cells(5).FindControl("lbtEdit") '修改
            lbtEdit.CommandArgument = dr("CLSID")

            lbtDel = e.Item.Cells(5).FindControl("lbtDel") '刪除
            lbtDel.CommandArgument = dr("CLSID")

            lbtCopy = e.Item.Cells(5).FindControl("lbtCopy") '複製
            lbtCopy.CommandArgument = dr("CLSID")

            '非登入此計畫.不能做刪除、修改
            If sm.UserInfo.TPlanID <> dr("TPlanID") Then
                lbtEdit.Visible = False
                lbtDel.Visible = False
            Else
                lbtEdit.Visible = True
                lbtDel.Visible = True
            End If

            Dim flag_can_del As Boolean = False '可以刪除預設為false
            Dim sqlstr_A As String = "SELECT 'x' FROM Class_ClassInfo a JOIN ID_Class b ON a.CLSID = b.CLSID WHERE a.CLSID = '" & dr("CLSID") & "' "
            If DbAccess.GetCount(sqlstr_A, objconn) > 0 Then
                lbtDel.Attributes("onclick") = "alert('此班別代碼資料尚有開班資料檔參照,不可刪除!!');return false;"
            Else
                lbtDel.Attributes("onclick") = "return confirm('此動作會刪除班別代碼資料，是否確定刪除?');"
                flag_can_del = True
            End If
            If flag_can_del Then
                '若可以刪除，再檢查一下。
                sqlstr_A = "SELECT * FROM Course_CourseInfo a WHERE CLSID = '" & dr("CLSID") & "' "
                If DbAccess.GetCount(sqlstr_A, objconn) > 0 Then
                    lbtDel.Attributes("onclick") = "alert('此班級代碼有歸屬的課程代碼，不可刪除!');return false;"
                End If
            End If

            lbtDel.Enabled = If(check_del.Value = "1", True, False)
            lbtEdit.Enabled = If(check_mod.Value = "1", True, False)
        End If
    End Sub

    Private Sub DG_Class_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Class.ItemCommand
        Select Case e.CommandName
            Case "edit"
                KeepSearch()
                Dim url1 As String = "TC_01_003_add.aspx?ID=" & Request("ID") & "&clsid=" & e.CommandArgument & "&ProcessType=Update"
                TIMS.Utl_Redirect(Me, objconn, url1)
            Case "del"
                Dim parms As Hashtable = New Hashtable()
                Dim aCLSID As String = e.CommandArgument
                Dim sql As String = "SELECT 'x' FROM Class_ClassInfo a JOIN ID_Class b ON a.CLSID = b.CLSID WHERE a.CLSID = @CLSID "
                parms.Clear()
                parms.Add("CLSID", aCLSID)
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "此班別代碼資料尚有開班資料檔參照,不可刪除!!")
                    Exit Sub
                End If

                sql = "DELETE ID_CLASS WHERE CLSID = @CLSID "
                parms.Clear()
                parms.Add("CLSID", aCLSID)
                DbAccess.ExecuteNonQuery(sql, objconn, parms)

                Common.MessageBox(Me, "刪除成功")
                'bt_search_Click(bt_search, Nothing)
                Call Search1()

            Case "copy"
                KeepSearch()
                Dim url1 As String = "TC_01_003_add.aspx?ID=" & Request("ID") & "&clsid=" & e.CommandArgument & "&ProcessType=Copy"
                TIMS.Utl_Redirect(Me, objconn, url1)
                'Response.Redirect("TC_01_003_add.aspx?ID=" & Request("ID") & "&clsid=" & e.CommandArgument & "&ProcessType=Copy")
        End Select
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'Response.Redirect("TC_01_003_print.aspx?ID=" & Request("ID"))
        'Button1.Attributes("onclick") = "window.open('TC_01_003_import.aspx?bt_search=bt_search&TPlanid=" & Plan_List.SelectedValue & "','','width=750,height=550,location=0,status=0,menubar=0,scrollbars=0,resizable=0');"
        KeepSearch()

        'Dim v_ddlDistID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_Plan_List As String = TIMS.GetListValue(Plan_List)
        'Dim v_ddlYears As String = TIMS.GetListValue(ddlYears)

        Dim mValue As String = ""
        mValue &= "ID=" & TIMS.sUtl_GetRqValue(Me, "ID")
        mValue &= "&TPlanid=" & v_Plan_List 'Plan_List.SelectedValue

        Dim url1 As String = "TC_01_003_import.aspx?" & mValue
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub
End Class