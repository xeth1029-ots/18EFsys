Partial Class SYS_01_001
    Inherits AuthBasePage

    Const cst_search1 As String = "search1"
    Const cst_sys01001addaspx As String = "SYS_01_001_add.aspx?ID="

    'Dim au As New cAUTH
    Dim objconn As SqlConnection
    Dim flag_SAH3 As Boolean = False

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        '檢查帳號的功能權限-----------------------------------Start
        'but_add.Enabled = False
        'If au.blnCanAdds Then but_add.Enabled = True
        'but_search.Enabled = False
        'If au.blnCanSech Then but_search.Enabled = True

        but_search.ToolTip = "機構共用不鎖轄區代碼"
        '檢查帳號的功能權限-----------------------------------End        

        'choice_button
        '帳號設定使用-跨轄區與計畫
        'Dim flag_SAH3 As Boolean = TIMS.IsSuperUser(sm, 3)
        flag_SAH3 = TIMS.IsSuperUser(sm, 3)

        If Not IsPostBack Then
            Call Create1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "TBplan", "", False, "SYS_01_001")
        If HistoryRID.Rows.Count <> 0 Then
            TBplan.Attributes("onclick") = "showObj('HistoryList2');"
            TBplan.Style("CURSOR") = "hand"
        End If

        If flag_SAH3 Then
            '帳號設定使用-跨轄區與計畫
            'choice_button.Attributes("onclick") = "wopen('../../Common/LevPlan.aspx?SAH=Y','計畫階段',850,400,1);"
            Dim js_win1 As String = String.Format("javascript:openOrg('../../Common/LevOrg.aspx?OrgField={0}&selected_year={1}');", "TBplan", sm.UserInfo.Years)
            choice_button.Attributes("onclick") = js_win1
        Else
            If sm.UserInfo.RoleID = 0 Or sm.UserInfo.RoleID = 1 Then
                choice_button.Attributes("onclick") = "wopen('../../Common/LevPlan.aspx','計畫階段',850,400,1);"
            Else
                'getParamValue('IDSetUp')==1 'opener.document.form1.TBplan.value=orgname;
                choice_button.Attributes("onclick") = "wopen('../../Common/LevOrg1.aspx?IDSetUp=1','計畫階段',480,420,1);"
            End If
        End If
        'choice_button.Attributes("onclick") = "wopen('../../Common/LevOrg1.aspx?IDSetUp=1','計畫階段',450,400,1);"
    End Sub

    Sub Create1()
        msg.Text = ""
        Me.ViewState("sort") = "RoleID"
        DataGridTable.Visible = False

        'Me.isused.SelectedValue = "Y"
        Common.SetListItem(isused, "Y")

        '所有超過3個月未登入的帳號，設為不啟用
        Call TIMS.UPDATE_STOP_ACCOUNT(objconn)
        Call TIMS.SUB_SET_HR_MI(ddlLastDATE1_HH, ddlLastDATE1_MM)
        Call TIMS.SUB_SET_HR_MI(ddlLastDATE2_HH, ddlLastDATE2_MM)

        Dim i_type_role As Integer = 0 'i_type: 1:全部/0:排除 0的角色

        i_type_role = If(sm.UserInfo.RoleID = 0, 1, 0)
        Role = TIMS.Get_IDRole(Role, TIMS.dtNothing(), i_type_role, objconn)

        If sm.UserInfo.RoleID = 0 Or sm.UserInfo.RoleID = 1 Then
            TBplan.Text = String.Concat(TIMS.GetPlanName(sm.UserInfo.PlanID, objconn), "_", sm.UserInfo.OrgName)
            RIDValue.Value = sm.UserInfo.RID
        Else
            TBplan.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Call UseKeepSearch()
    End Sub

    '保留  Session(cst_search1) = search1
    Sub KeepSearch()
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        Dim search1 As String = ""
        'Session("search1") = Nothing
        search1 = ""
        search1 &= "&prg=sys01001"
        search1 &= "&nameid=" & TIMS.ClearSQM(nameid.Text)
        search1 &= "&namefield=" & TIMS.ClearSQM(namefield.Text)
        search1 &= "&TBplan=" & TIMS.ClearSQM(TBplan.Text)
        search1 &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        'search1 &= "&PlanIDValue=" & TIMS.ClearSQM(PlanIDValue.Value)
        search1 &= "&hidPlanID=" & TIMS.ClearSQM(hidPlanID.Value)
        search1 &= "&isused=" & TIMS.ClearSQM(isused.SelectedValue)
        search1 &= "&Role=" & TIMS.ClearSQM(Role.SelectedValue)
        Session(cst_search1) = search1
    End Sub

    Sub UseKeepSearch()
        If Session(cst_search1) IsNot Nothing Then
            Dim str_search1 As String = Session(cst_search1)
            Session(cst_search1) = Nothing

            Dim myValue As String = ""
            myValue = TIMS.GetMyValue(str_search1, "prg")
            If myValue = "sys01001" Then
                nameid.Text = TIMS.GetMyValue(str_search1, "nameid")
                nameid.Text = TIMS.ClearSQM(nameid.Text)
                namefield.Text = TIMS.GetMyValue(str_search1, "namefield")
                TBplan.Text = TIMS.GetMyValue(str_search1, "TBplan")
                RIDValue.Value = TIMS.GetMyValue(str_search1, "RIDValue")
                'PlanIDValue.Value = TIMS.GetMyValue(str_search1, "PlanIDValue")
                hidPlanID.Value = TIMS.GetMyValue(str_search1, "hidPlanID")

                myValue = TIMS.GetMyValue(str_search1, "isused")
                If myValue <> "" Then Common.SetListItem(isused, myValue)
                myValue = TIMS.GetMyValue(str_search1, "Role")
                If myValue <> "" Then Common.SetListItem(Role, myValue)
                If but_search.Enabled Then Call Search1()
            End If
        End If
    End Sub

    Public Shared Sub TEST_LOGIN_TT2(ByRef MyPage As Page, ByVal v_txt_account2 As String, ByVal oConn As SqlConnection)
        v_txt_account2 = TIMS.ClearSQM(v_txt_account2)
        'txt_account2.Text = TIMS.ClearSQM(txt_account2.Text)
        Dim v_test_user As String = v_txt_account2 'txt_account2.Text
        Dim sUrl1 As String = "https://wltest.wda.gov.tw/login.aspx?USER=:USER&TID=:TID"
        Dim dr1 As DataRow = TIMS.Get_AccountData(v_test_user, oConn)
        If dr1 Is Nothing Then
            Common.MessageBox(MyPage, "查無該帳號資料!!")
            Exit Sub
        End If
        Dim v_test_pwd As String = Convert.ToString(dr1("PASSWORD"))
        If v_test_pwd.Length >= 32 Then v_test_pwd = TIMS.DecryptAes(v_test_pwd)

        Dim v_test_TID As String = ""
        TIMS.SetMyValue(v_test_TID, "BACKUSE", "Y")
        TIMS.SetMyValue(v_test_TID, "USER", v_test_user)
        TIMS.SetMyValue(v_test_TID, "PPWD", v_test_pwd)
        TIMS.SetMyValue(v_test_TID, "GSTIME", TIMS.GetSysDateNow(oConn))
        'Session("test_usr") = v_test_user
        sUrl1 = Replace(sUrl1, ":USER", v_test_user)
        sUrl1 = Replace(sUrl1, ":TID", TIMS.EncryptAes(v_test_TID))

        Dim strScript As String = ""
        strScript = "" & vbCrLf
        strScript &= "<script language=""javascript"">" & vbCrLf
        strScript &= ReportQuery.strWOScript(sUrl1)
        strScript &= "</script>" & vbCrLf
        MyPage.RegisterStartupScript(TIMS.xBlockName(), strScript)
        'TIMS.Utl_Redirect(MyPage, oConn, sUrl1)
        'MyPage.Response.Redirect(sUrl1)
    End Sub

    Function get_testurl() As Boolean
        Dim rst As Boolean = False
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        Const cst_ppp As String = "!@#"
        Dim s_Right As String = ""
        If (nameid.Text = "") Then Return rst
        If Not (nameid.Text.Length > 3) Then Return rst

        If (nameid.Text.IndexOf(cst_ppp) > -1) Then s_Right = nameid.Text.Substring(nameid.Text.Length - 3, 3)
        If (s_Right = "") Then Return rst

        Dim flag_get_test_url As Boolean = False
        If s_Right.Equals(cst_ppp) Then flag_get_test_url = True
        If Not flag_get_test_url Then Return rst

        If flag_get_test_url Then nameid.Text = Replace(nameid.Text, cst_ppp, "")
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        If (nameid.Text = "") Then Return rst

        If TIMS.IsSuperUser(sm, 1) AndAlso nameid.Text <> "" AndAlso flag_get_test_url Then
            rst = True
            '帳號權限足夠 帳號不為空 取得測試url
            Call TEST_LOGIN_TT2(Me, nameid.Text, objconn)
            Return rst
        End If
        Return rst
    End Function

    '查詢 SQL
    Sub Search1()
        Dim flag_gettesturl As Boolean = get_testurl()
        If flag_gettesturl Then Return

        Dim myParam As Hashtable = New Hashtable
        myParam.Clear()

        '帳號未賦予計畫時，可賦予計畫 09/06/24 by waiming
        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr &= " SELECT DISTINCT a.Account, a.Name" & vbCrLf
        sqlstr &= " ,e.Name RoleName" & vbCrLf
        sqlstr &= " ,a.IsUsed" & vbCrLf
        sqlstr &= " ,a.OrgID" & vbCrLf
        sqlstr &= " ,d.OrgName" & vbCrLf
        sqlstr &= " ,a.RoleID " & vbCrLf
        sqlstr &= " ,format(a.Last_LoginDate,'yyyy-MM-dd HH:mm:ss') LastDate" & vbCrLf
        sqlstr &= " FROM AUTH_ACCOUNT a " & vbCrLf
        sqlstr &= " JOIN ID_ROLE e ON e.RoleID = a.RoleID " & vbCrLf
        sqlstr &= " JOIN ORG_ORGINFO d ON d.OrgID = a.OrgID " & vbCrLf '帳號跟隨機構
        sqlstr &= " JOIN AUTH_RELSHIP c ON c.OrgID = a.OrgID " & vbCrLf  '業務權限。
        sqlstr &= " WHERE 1=1"
        'If RIDValue.Value <> "" Then sqlstr += "join Auth_Relship c on c.OrgID=a.OrgID" & vbCrLf '有選擇
        sqlstr &= " AND a.RoleID >= '" & sm.UserInfo.RoleID & "' " & vbCrLf '角色層級不可查大過於自已的

        'If Me.nameid.Text <> "" Then Me.nameid.Text = Trim(Me.nameid.Text)
        If Me.nameid.Text <> "" Then
            '帳號
            '若是為身分證號格式，試著查身分證號
            If TIMS.CheckIDNO(UCase(Me.nameid.Text)) Then
                sqlstr &= " AND (1!=1" & vbCrLf
                sqlstr &= " OR UPPER(a.account) = UPPER(@nameid) " & vbCrLf
                sqlstr &= " OR UPPER(a.idno) = UPPER(@nameid2) " & vbCrLf
                sqlstr &= " )" & vbCrLf
                myParam.Add("nameid", UCase(Me.nameid.Text))
                myParam.Add("nameid2", TIMS.CheckIDNO(UCase(Me.nameid.Text)))
            Else
                sqlstr &= " AND UPPER(a.account) = UPPER(@nameid) " & vbCrLf
                myParam.Add("nameid", UCase(Me.nameid.Text))
            End If
        End If

        '姓名
        If Me.namefield.Text <> "" Then Me.namefield.Text = Trim(Me.namefield.Text)
        If Me.namefield.Text <> "" Then
            'sqlstr += " and upper(a.name) like '%" & UCase(Me.namefield.Text) & "%'" & vbCrLf
            sqlstr &= " AND UPPER(a.name) LIKE '%' + @namefield + '%' " & vbCrLf
            myParam.Add("namefield", Me.namefield.Text)
        End If

        '啟用
        Select Case Me.isused.SelectedValue
            Case "Y", "N"
                sqlstr &= " AND a.IsUsed = @IsUsed " & vbCrLf
                myParam.Add("IsUsed", Me.isused.SelectedValue)
        End Select

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If sm.UserInfo.RID = "A" AndAlso (sm.UserInfo.RoleID = 0 OrElse sm.UserInfo.RoleID = 1) Then
            '署(局)的管理者
            If flag_SAH3 Then
                If RIDValue.Value <> "" Then
                    '有選擇計畫與單位 依選擇單位做判斷
                    sqlstr &= " AND c.RID = @RIDValue "
                    myParam.Add("RIDValue", RIDValue.Value)
                End If
            Else
                '有選擇計畫與單位 依選擇單位做判斷
                sqlstr &= " AND c.RID = @RIDValue "
                myParam.Add("RIDValue", If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID))
            End If
        Else
            If RIDValue.Value <> "" Then
                '有選擇計畫與單位 依選擇單位做判斷
                sqlstr &= " AND EXISTS (SELECT 'x' FROM AUTH_ACCRWPLAN x WHERE UPPER(x.account) = UPPER(a.account) " & vbCrLf  '09/06/24 by waiming
                sqlstr &= " AND c.RID = x.RID "
                'sqlstr += " AND x.RID = '" & RIDValue.Value & "') "
                sqlstr &= " AND x.RID = @RIDValue) "
                myParam.Add("RIDValue", RIDValue.Value)
            Else
                '未選擇計畫與單位 依登入權限做判斷
                'sm.UserInfo.RoleID 
                '0:超級使用者
                '1:系統管理者()
                '2:一級以上()
                '3:一級()
                '4:二級()
                '5:承辦人()
                '99一般使用者()
                If sm.UserInfo.RoleID <= 5 Then
                    'sqlstr += " and c.RID='" & sm.UserInfo.RID & "'" & vbCrLf
                    sqlstr &= " AND EXISTS (SELECT 'x' FROM Auth_AccRWPlan x " & vbCrLf
                    'sqlstr += " JOIN auth_relship x2 ON x2.rid = x.rid "
                    sqlstr &= " WHERE UPPER(x.account) = UPPER(a.account) " & vbCrLf
                    sqlstr &= " AND c.RID = x.RID " & vbCrLf
                    sqlstr &= " AND c.Relship LIKE '%/" & sm.UserInfo.RID & "/%' " & vbCrLf 'RID依 Relship比對
                    sqlstr &= " AND c.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf '同登入轄區
                    'sqlstr += " AND x.PlanID = '" & sm.UserInfo.PlanID & "' "
                    sqlstr &= " ) " & vbCrLf
                    'sqlstr += " AND x2.Relship LIKE '%/" & sm.UserInfo.RID & "/%') "
                Else
                    '99一般使用者() '應該不會有此層級
                    sqlstr &= " AND c.RID = '" & sm.UserInfo.RID & "' " & vbCrLf
                    sqlstr &= " AND EXISTS (SELECT 'x' FROM Auth_AccRWPlan x WHERE UPPER(x.account) = UPPER(a.account) " & vbCrLf
                    sqlstr &= " AND c.RID = x.RID " & vbCrLf
                    sqlstr &= " AND c.Relship LIKE '%/" & sm.UserInfo.RID & "/%' " & vbCrLf
                    sqlstr &= " AND c.DistID = '" & sm.UserInfo.DistID & "' " & vbCrLf
                    sqlstr &= " AND c.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
                    sqlstr &= " AND x.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
                    sqlstr &= " AND x.RID = '" & sm.UserInfo.RID & "') " & vbCrLf 'RID限定
                End If
            End If

        End If

        '角色
        If Role.SelectedValue <> "" Then
            sqlstr &= " AND e.RoleID = @RoleID " & vbCrLf
            'sqlstr += " and e.RoleID='" & Role.SelectedValue & "'" & vbCrLf
            myParam.Add("RoleID", Role.SelectedValue)
        End If

        '單位機構名
        TBplan.Text = TIMS.ClearSQM(TBplan.Text)
        If RIDValue.Value = "" AndAlso TBplan.Text <> "" Then
            sqlstr &= " AND d.OrgName like '%'+@OrgName+'%'" & vbCrLf
            myParam.Add("OrgName", TBplan.Text)
        End If

        Dim s_LastDATE1 As String = TIMS.GET_DateHM(LastDATE1, ddlLastDATE1_HH, ddlLastDATE1_MM)
        Dim s_LastDATE2 As String = TIMS.GET_DateHM(LastDATE2, ddlLastDATE2_HH, ddlLastDATE2_MM)
        If s_LastDATE1 <> "" AndAlso Not TIMS.IsDate1(s_LastDATE1) Then
            '異常日期修正為今天 00:00
            s_LastDATE1 = Now.ToString("yyyy/MM/dd") & " 00:00"
            LastDATE1.Text = TIMS.Cdate3(s_LastDATE1)
            TIMS.SET_DateHM(CDate(s_LastDATE1), ddlLastDATE1_HH, ddlLastDATE1_MM)
        End If
        If s_LastDATE2 <> "" AndAlso Not TIMS.IsDate1(s_LastDATE2) Then
            '異常日期修正為今天 23:59
            s_LastDATE2 = Now.ToString("yyyy/MM/dd") & " 23:59"
            LastDATE2.Text = TIMS.Cdate3(s_LastDATE2)
            TIMS.SET_DateHM(CDate(s_LastDATE2), ddlLastDATE2_HH, ddlLastDATE2_MM)
        End If
        If (s_LastDATE1 <> "" AndAlso s_LastDATE2 <> "") Then
            If DateDiff(DateInterval.Minute, CDate(s_LastDATE1), CDate(s_LastDATE2)) < 0 Then
                '順序異常(查詢對調)
                Dim s_LastDATE_tmp As String = s_LastDATE1
                s_LastDATE1 = s_LastDATE2
                s_LastDATE2 = s_LastDATE_tmp

                LastDATE1.Text = TIMS.Cdate3(s_LastDATE1)
                TIMS.SET_DateHM(CDate(s_LastDATE1), ddlLastDATE1_HH, ddlLastDATE1_MM)
                LastDATE2.Text = TIMS.Cdate3(s_LastDATE2)
                TIMS.SET_DateHM(CDate(s_LastDATE2), ddlLastDATE2_HH, ddlLastDATE2_MM)
            End If
        End If
        If s_LastDATE1 <> "" Then
            sqlstr &= " AND a.Last_LoginDate >= @LastDATE1" & vbCrLf
            myParam.Add("LastDATE1", CDate(s_LastDATE1))
        End If
        If s_LastDATE2 <> "" Then
            sqlstr &= " AND a.Last_LoginDate <= @LastDATE2" & vbCrLf
            myParam.Add("LastDATE2", CDate(s_LastDATE2))
        End If

        'dt.Load(.ExecuteReader())
        Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        'https://dotblogs.com.tw/harry/2016/10/14/181017
        Dim slogMsg1 As String = ""
        slogMsg1 &= "##SYS_01_001, sqlstr: " & sqlstr & vbCrLf
        slogMsg1 &= "##SYS_01_001, myParam: " & TIMS.GetMyValue3(myParam) & vbCrLf
        If flag_chktest Then TIMS.WriteLog(Me, slogMsg1)

        Dim dt As New DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn, myParam)

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.PrimaryKey = "Account"
        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub but_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_search.Click
        Call Search1()
    End Sub

    '新增
    Private Sub but_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_add.Click
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        'If nameid.Text <> "" Then nameid.Text = Trim(nameid.Text)
        Call KeepSearch()

        Dim sUrl_1 As String = String.Concat(cst_sys01001addaspx, TIMS.Get_MRqID(Me), "&myid=", nameid.Text, "&act=add")
        TIMS.Utl_Redirect1(Me, sUrl_1)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "btnEdit"
                Dim s_Account As String = TIMS.GetMyValue(e.CommandArgument, "Account")
                If s_Account = "" Then Return

                Call KeepSearch()

                Dim sUrl_1 As String = String.Concat(cst_sys01001addaspx, TIMS.Get_MRqID(Me), "&Account=", s_Account, "&act=edit")
                TIMS.Utl_Redirect1(Me, sUrl_1)
        End Select

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim mylabel As String
                Dim mysort As New System.Web.UI.WebControls.Image
                Dim i As Integer
                Select Case Me.ViewState("sort")
                    Case "RoleID", "RoleID desc"
                        mylabel = "RoleID"
                        i = 2
                        If Me.ViewState("sort") = "RoleID" Then
                            mysort.ImageUrl = "../../images/SortUp.gif"
                        Else
                            mysort.ImageUrl = "../../images/SortDown.gif"
                        End If
                    Case "account", "account desc"
                        mylabel = "account"
                        i = 0
                        If Me.ViewState("sort") = "account" Then
                            mysort.ImageUrl = "../../images/SortUp.gif"
                        Else
                            mysort.ImageUrl = "../../images/SortDown.gif"
                        End If
                End Select
                e.Item.Cells(i).Controls.Add(mysort)
                'DataGrid1.Columns(i).SortExpression = mylabel
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnEdit As LinkButton = e.Item.FindControl("btnEdit")
                Dim cmdArg As String = ""
                cmdArg = ""
                cmdArg &= "&Account=" & Convert.ToString(drv("Account"))
                'cmdArg &= "&Account=" & drv("Account")
                btnEdit.CommandArgument = cmdArg

                'If sm.UserInfo.RoleID <> 0 Then
                '    btnEdit.Enabled = False
                '    If au.blnCanMod Then btnEdit.Enabled = True
                'End If
        End Select

    End Sub

    '排序
    Private Sub DataGrid1_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid1.SortCommand
        If Me.ViewState("sort") = e.SortExpression Then
            Me.ViewState("sort") = e.SortExpression & " desc"
        Else
            Me.ViewState("sort") = e.SortExpression
        End If

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    '清除
    Private Sub btnClear1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnClear1.Click
        TBplan.Text = ""
        RIDValue.Value = ""
        'PlanIDValue.Value = ""
        hidPlanID.Value = ""
    End Sub
End Class
