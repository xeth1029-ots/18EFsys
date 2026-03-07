Partial Class SYS_01_004
    Inherits AuthBasePage

    'Dim sql As String = ""

#Region "Sub"
    '代入清單資料
    Private Sub get_list()
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        Dim obj As Object = Nothing
        Dim sql As String = ""
        For i As Integer = 1 To 2
            If i = 1 Then
                sql = "select roleid id,name from ID_Role "
                If Not (sm.UserInfo.LID = 0 And sm.UserInfo.RoleID = 0) Then  '非署(局)帳號(super user)
                    sql += "where roleid <> 0"
                End If
                obj = rdo_role
            Else
                sql = ""
                sql += " select distinct a.years id,a.years name "
                sql += " from id_plan a "
                sql += " join key_plan b on a.tplanid=b.tplanid "
                sql += " where a.years is not null order by a.years "
                obj = ddl_years
            End If
            Call TIMS.OpenDbConn(objconn)
            Dim dt As New DataTable
            Dim oCmd As New SqlCommand(sql, objconn)
            With oCmd
                dt.Load(.ExecuteReader())
            End With
            If dt.Rows.Count > 0 Then
                With obj
                    .DataSource = dt
                    .DataTextField = "name"
                    .DataValueField = "id"
                    .DataBind()
                End With
            End If
            'With sda
            '    .SelectCommand = New SqlCommand(sql, objconn)
            '    .Fill(ds, i.ToString)
            'End With
            'If ds.Tables(i.ToString).Rows.Count > 0 Then
            '    With obj
            '        .DataSource = ds.Tables(i.ToString)
            '        .DataTextField = "name"
            '        .DataValueField = "id"
            '        .DataBind()
            '    End With
            'End If
        Next

        rdo_role.Items.Insert(0, New ListItem("不區分", ""))
        rdo_role.SelectedIndex = 0

        ddl_years.SelectedValue = Me.sm.UserInfo.Years
    End Sub

    '取得單位底下的使用者帳號(只取正常帳號使用者，排除已停用,自然人憑證已清除,三個月未登入系統之使用者)
    Private Sub get_Acct()
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        'Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql += " select  a.Name as accname" & vbCrLf
        sql += " ,a.Name+'('+b.Name+') ['+a.Account+']' Name" & vbCrLf
        sql += " ,dbo.TRUNC_DATETIME(getdate()-a.last_loginDate)  as  interval" & vbCrLf
        sql += " ,a.IsUsed,a.account,a.roleid,a.lid,a.note,a.Last_LoginDate,a.Serialno" & vbCrLf
        sql += " from auth_account a" & vbCrLf
        sql += " join ID_Role b on a.RoleID=b.RoleID" & vbCrLf
        sql += " join Auth_Relship c on c.OrgID=a.OrgID" & vbCrLf
        sql += " where c.RID= @RID " & vbCrLf
        sql += " and a.RoleID >=" & sm.UserInfo.RoleID & vbCrLf
        If sm.UserInfo.RID <> RIDValue.Value.Chars(0) Then
            If sm.UserInfo.RoleID = "0" Then '超級使用者 
                Select Case rdoIsUsed.SelectedValue
                    Case "A"
                    Case "Y"
                        sql += "and a.IsUsed='Y'" & vbCrLf
                        If rdo_role.SelectedValue <> "0" Then '超級使用者(無自然人憑證)不做驗証
                            sql += "  and a.Serialno is not null"
                        End If
                    Case Else
                        If rdo_role.SelectedValue <> "" Then
                            sql += " and a.IsUsed='N'"
                        Else
                            sql += " and (a.IsUsed='N' or a.Serialno is null )"
                        End If
                End Select
            Else
                Select Case rdoIsUsed.SelectedValue '非超級使用者(跨區了)
                    Case "A"
                        sql += "and a.RoleID not in (0,1)" & vbCrLf
                    Case "Y"
                        sql += "and a.RoleID not in (0,1) and a.IsUsed='Y'" & vbCrLf
                        If rdo_role.SelectedValue <> "0" Then '超級使用者(無自然人憑證)不做驗証
                            sql += "  and a.Serialno is not null"
                        End If
                    Case Else
                        If rdo_role.SelectedValue <> "" Then
                            sql += " and a.RoleID not in (0,1) and a.IsUsed='N'"
                        Else
                            sql += " and (and a.RoleID not in (0,1) and a.IsUsed='N' or a.Serialno is null )" '帳號角色不區分時
                        End If
                End Select
            End If
        Else
            '限定不可為超級使用者
            Select Case rdoIsUsed.SelectedValue
                Case "A"
                    sql += "and a.RoleID<>0" & vbCrLf
                Case "Y"
                    sql += "and a.RoleID<>0 and a.IsUsed='Y'" & vbCrLf
                    If rdo_role.SelectedValue <> "0" Then '超級使用者(無自然人憑證)不做驗証
                        sql += "  and a.Serialno is not null"
                    End If
                Case Else
                    If rdo_role.SelectedValue <> "" Then
                        sql += " and a.RoleID<>0 and a.IsUsed='N'"
                    Else
                        sql += " and a.RoleID<>0 and (a.IsUsed='N' or a.Serialno is null)" '帳號角色不區分時
                    End If
            End Select
        End If

        If rdo_role.SelectedValue <> "" Then
            sql += " and a.RoleID = @roleid "
        End If

        If (rdo_role.SelectedValue = "" Or rdo_role.SelectedValue = "0") And rdoIsUsed.SelectedValue = "Y" Then
            'super user snoopy 帳號清單中要顯示snoopy,除此之外,任何一個帳號登入至此功能,都不得顯示snoopy帳號
            If sm.UserInfo.UserID = str_superuser1 Then
                sql += " or a.account='" & str_superuser1 & "' "
            Else
                sql += " and a.account<>'" & str_superuser1 & "'"
            End If
        End If
        sql += " order by accname,a.account "

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = Me.RIDValue.Value
            If rdo_role.SelectedValue <> "" Then
                .Parameters.Add("roleid", SqlDbType.VarChar).Value = rdo_role.SelectedValue
            End If
            dt.Load(.ExecuteReader())
        End With
        msg.Text = "無可供賦予權限之帳號!"
        Datagrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Datagrid1.Visible = True
            Datagrid1.DataSource = dt
            Datagrid1.DataBind()
        End If

        'With sda
        '    .SelectCommand = New SqlCommand(sql, objconn)
        '    .SelectCommand.Parameters.Clear()
        '    .SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = Me.RIDValue.Value
        '    If rdo_role.SelectedValue <> "" Then
        '        .SelectCommand.Parameters.Add("roleid", SqlDbType.VarChar).Value = rdo_role.SelectedValue
        '    End If
        '    .Fill(ds)
        'End With
        'msg.Text = "無可供賦予權限之帳號!"
        'Datagrid1.Visible = False
        'If ds.Tables(0).Rows.Count > 0 Then
        '    msg.Text = ""
        '    Datagrid1.Visible = True
        '    Datagrid1.DataSource = ds.Tables(0)
        '    Datagrid1.DataBind()
        'End If

        chk_state()

    End Sub

    '取得RID之相關計畫(依單位及年度)
    Private Sub get_plan()
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        Dim filterStr As String = ""
        Dim sql2 As String = ""
        'Call TIMS.OpenDbConn(objconn)

        '先取出 orgid
        Dim sql As String = ""
        sql += " select a.orgid,b.orgname "
        sql += " from Auth_Relship a " & vbCrLf
        sql += " join org_orginfo b on a.orgid=b.orgid " & vbCrLf
        sql += " where a.rid= @rid "
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("rid", SqlDbType.VarChar).Value = RIDValue.Value
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            center.Text = Convert.ToString(dt.Rows(0)("orgname"))
        End If

        'With sda
        '    .SelectCommand = New SqlCommand(sql, objconn)
        '    .SelectCommand.Parameters.Clear()
        '    .SelectCommand.Parameters.Add("rid", SqlDbType.VarChar).Value = RIDValue.Value
        '    .Fill(ds, "org")
        'End With
        'If ds.Tables("org").Rows.Count > 0 Then
        '    center.Text = Convert.ToString(ds.Tables("org").Rows(0)("orgname"))
        'End If


        If sm.UserInfo.RoleID <= 1 Then  '假如登入者角色代碼為系統管理者(或超級使用者)
            filterStr = " aa.Years='" & ddl_years.SelectedValue & "'"
            If sm.UserInfo.LID <> "0" Then  '非署(局)帳號登入時
                filterStr += " and  aa.DistID='" & sm.UserInfo.DistID & "'"
            End If
        Else
            Select Case rdo_role.SelectedValue        '角色代碼 1：系統管理者
                Case 1    '選擇賦予的使用者 為系統管理者 ，依登入者轄區顯示計劃
                    filterStr = " aa.DistID='" & sm.UserInfo.DistID & "' and aa.Years='" & ddl_years.SelectedValue & "'"
                Case Else '選擇賦予的使用者 不為系統管理者，依登入者計畫顯示計畫
                    filterStr = " aa.PlanID='" & sm.UserInfo.PlanID & "' and aa.Years='" & ddl_years.SelectedValue & "'"
            End Select
        End If

        sql2 = "" & vbCrLf
        sql2 += " join" & vbCrLf
        sql2 += " ( select RID" & vbCrLf
        sql2 += " ,CASE WHEN len(Relship)-4 >0 THEN replace(replace(substring(Relship,5,len(Relship)-4),RID,''),'/','') end AS CRID" & vbCrLf
        'sql2 += " WHEN Len(@Relship)-4 >0" & vbCrLf
        'sql2 += " THEN replace(replace( dbo.SUBSTR(Relship,5,Len(@Relship)-4),RID,''),'/','')" & vbCrLf
        'sql2 += " END AS CRID from Auth_Relship) c  on  aa.RID=c.RID" & vbCrLf
        sql2 += " from Auth_Relship) c  on  aa.RID=c.RID" & vbCrLf
        sql2 += " left join Auth_Relship a2 on a2.RID=c.CRID" & vbCrLf
        sql2 += " left join Org_OrgInfo o2 on a2.OrgID=o2.OrgID" & vbCrLf

        If RIDValue.Value.Length <= 1 Then '分署(中心)、署(局)  
            sql = ""
            sql += " select '('+o2.OrgName+')' as SubOrgName" & vbCrLf
            sql += " ,aa.* " & vbCrLf
            sql += " from (" & vbCrLf
            sql += "   select a.planid,a.years,b.distid,a.planname,b.rid,b.orgid,o.orgname " & vbCrLf
            sql += "   from view_LoginPlan a" & vbCrLf
            sql += "   join auth_relship b on b.distid=a.distid" & vbCrLf
            sql += "   join org_orginfo o on b.orgid=o.orgid" & vbCrLf
            sql += "   where a.years = @years "
            sql += "   and b.rid= @rid " & vbCrLf

            '可將計畫附與不同轄區單位
            If Convert.ToString(sm.UserInfo.RID) <> RIDValue.Value Then
                sql += " union " & vbCrLf
                sql += "   select a.planid,a.years,b.distid,a.planname,b.rid,b.orgid,o.orgname " & vbCrLf
                sql += "   from view_LoginPlan a" & vbCrLf
                sql += "   join auth_relship b on b.distid=a.distid" & vbCrLf
                sql += "   join org_orginfo o on b.orgid=o.orgid" & vbCrLf
                sql += "   where a.years = @years "
                sql += "   and b.rid= '" & Convert.ToString(sm.UserInfo.RID) & "'" & vbCrLf
                sql += " ) aa"

                sql += sql2

                If filterStr <> "" Then
                    sql += " where  " & filterStr
                End If

                If RIDValue.Value > Convert.ToString(sm.UserInfo.RID) Then
                    sql += "order by aa.rid desc,aa.planid"
                Else
                    sql += "order by aa.rid,aa.planid"
                End If
            Else
                sql += " ) aa " & sql2
            End If

        Else  '非分署(中心)、署(局)  
            sql = ""
            sql += " select '('+o2.OrgName+')' as SubOrgName, aa.* from ( "
            sql += " select o.orgname,vp.planname,ip.years,ar.* " & vbCrLf
            sql += " from auth_relship ar " & vbCrLf
            sql += " left join id_plan ip on ip.planid=ar.planid " & vbCrLf
            sql += " join view_LoginPlan vp on ip.planid=vp.planid " & vbCrLf
            sql += " left join org_orginfo o on ar.orgid=o.orgid " & vbCrLf
            sql += " where ip.years= @years  " & vbCrLf
            sql += " and ar.orgid=(select orgid from Auth_Relship where rid= @rid ) " & vbCrLf
            sql += " ) aa  "

            sql += sql2

            If filterStr <> "" Then
                sql += " where  " & filterStr
            End If
        End If
        Call TIMS.OpenDbConn(objconn)
        dt = New DataTable
        oCmd = New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("rid", SqlDbType.VarChar).Value = RIDValue.Value
            .Parameters.Add("years", SqlDbType.VarChar).Value = ddl_years.SelectedValue
            dt.Load(.ExecuteReader())
        End With
        dt.Columns.Add(New DataColumn("allow"))   '是否有權限對此計畫項目進行賦予
        msg2.Text = "無可供賦予之計畫!"
        Datagrid2.Visible = False
        If dt.Rows.Count > 0 Then
            msg2.Text = ""
            Datagrid2.Visible = True

            Datagrid2.DataSource = dt
            Datagrid2.DataBind()
        End If


        'With sda
        '      .SelectCommand = New SqlCommand(sql, objconn)
        '      .SelectCommand.CommandText = sql
        '      .SelectCommand.Parameters.Clear()
        '      .SelectCommand.Parameters.Add("rid", SqlDbType.VarChar).Value = RIDValue.Value
        '      .SelectCommand.Parameters.Add("years", SqlDbType.VarChar).Value = ddl_years.SelectedValue
        '      .Fill(ds, "data")
        '  End With
        '  ds.Tables("data").Columns.Add(New DataColumn("allow"))   '是否有權限對此計畫項目進行賦予
        '  If ds.Tables("data").Rows.Count > 0 Then
        '      Datagrid2.DataSource = ds.Tables("data")
        '      Datagrid2.DataBind()
        '      Datagrid2.Visible = True
        '      msg2.Text = ""
        '  Else
        '      msg2.Text = "無可供賦予之計畫!"
        '      Datagrid2.Visible = False
        '  End If

        chk_state()

    End Sub

    Private Sub chk_state()
        If Datagrid1.Visible = True And Datagrid2.Visible = True Then
            bt_save.Enabled = True
        Else
            bt_save.Enabled = False
        End If
    End Sub

    '新增計畫權限
    Sub Add_AccRWPlan()
        'ByVal da As SqlDataAdapter
        'Dim dt As DataTable = Nothing

        For i As Integer = 0 To Datagrid2.Items.Count - 1
            Dim chk2 As CheckBox = Datagrid2.Items(i).Cells(1).FindControl("chk2")
            Dim RID As HtmlInputHidden = Datagrid2.Items(i).FindControl("hid_RID")
            Dim planid As HtmlInputHidden = Datagrid2.Items(i).FindControl("hid_planid")

            Dim acct As String = "" '賦予權限帳號
            Dim sql As String = ""
            If chk2.Checked = True Then
                Dim dt As DataTable
                dt = get_AccRWPlan(planid.Value, RID.Value)
                For m1 As Integer = 0 To Datagrid1.Items.Count - 1 '確認是該計劃是否已有賦予帳號計畫權限，若無才加入新增帳號
                    Dim chk1 As CheckBox = Datagrid1.Items(m1).Cells(1).FindControl("chk1")
                    Dim hid_account As HtmlInputHidden = Datagrid1.Items(m1).FindControl("hid_account")
                    If chk1.Checked = True Then
                        If dt.Rows.Count > 0 Then
                            If dt.Select(" account='" & hid_account.Value & "'").Length = 0 Then
                                If acct = "" Then
                                    acct = "'" & hid_account.Value & "'"
                                Else
                                    acct = acct & ",'" & hid_account.Value & "'"
                                End If
                            End If
                        Else
                            If acct = "" Then
                                acct = "'" & hid_account.Value & "'"
                            Else
                                acct = acct & ",'" & hid_account.Value & "'"
                            End If
                        End If
                    End If
                Next

                sql = ""
                sql &= " insert  into  Auth_AccRWPlan (Account,PlanID,RID,CreateByAcc,ModifyAcct,ModifyDate) " & vbCrLf
                sql &= " Select  account , @planid, @rid,'N',@ModifyAcct, getdate()" & vbCrLf
                sql &= " from  auth_account  where   account  in " & vbCrLf
                sql &= " ( " & acct & " ) "
                If acct <> "" And planid.Value <> "" And RID.Value <> "" Then
                    Dim iCmd As New SqlCommand(sql, objconn)
                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("planid", SqlDbType.VarChar).Value = planid.Value
                        .Parameters.Add("rid", SqlDbType.VarChar).Value = RID.Value
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .ExecuteNonQuery()
                    End With
                End If
            End If
        Next
        'If Not dt Is Nothing Then dt.Dispose()
    End Sub
#End Region

#Region "Function"

    Private Function get_AccRWPlan(ByVal planid As String, ByVal rid As String) As DataTable
        Dim dt As New DataTable
        If planid <> "" AndAlso rid <> "" Then
            Dim sql As String = ""
            sql = " select  * from  Auth_AccRWPlan  where planid= @planid  and rid= @rid "
            Call TIMS.OpenDbConn(objconn)
            Dim oCmd As New SqlCommand(sql, objconn)
            With oCmd
                .Parameters.Clear()
                .Parameters.Add("planid", SqlDbType.VarChar).Value = planid
                .Parameters.Add("rid", SqlDbType.VarChar).Value = rid
                dt.Load(.ExecuteReader())
            End With
        End If
        Return dt

        'With da
        '    .SelectCommand.CommandText = sql
        '    .SelectCommand.Parameters.Clear()
        '    .SelectCommand.Parameters.Add("planid", SqlDbType.VarChar).Value = planid
        '    .SelectCommand.Parameters.Add("rid", SqlDbType.VarChar).Value = rid
        '    If planid <> "" And rid <> "" Then
        '        .Fill(dt)
        '    End If
        'End With
        'Return dt
        'If Not dt Is Nothing Then dt.Dispose()        
    End Function
#End Region

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        chkall_1.Value = "0" '全選值清空
        chkall_2.Value = "0"

        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
            str_superuser1 = CStr(sm.UserInfo.UserID)
        End If

        If Chk_ROLE_NOUSE(sm) Then
            bt_save.Enabled = False
            btu_org.Disabled = True
            Common.RespWrite(Me, "<script>alert('您目前無此功能使用權限!');</script>")
            Exit Sub
        End If

        If Not IsPostBack Then
            bt_save.Enabled = True
            btu_org.Disabled = False

            get_list()

            RIDValue.Attributes.Add("onpropertychange", "document.getElementById('but_search').click();")
            but_search.Style.Add("visibility", "hidden")
            chkAll1.Attributes.Add("onclick", "switch_chk('Datagrid1','" & chkAll1.ID & "')")
            chkAll2.Attributes.Add("onclick", "switch_chk('Datagrid2','" & chkAll2.ID & "')")

            tbSch.Visible = False
            bt_save.Enabled = False
        End If

        If sm.UserInfo.RID <> "A" And sm.UserInfo.RoleID = "1" Then
            btu_org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?submit=true&GetOther=1'+'&OrgField=center')"
        ElseIf sm.UserInfo.RID = "A" Then
            btu_org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx?submit=true'+'&OrgField=center')"
        Else
            btu_org.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx?btnName=but_search'+'&OrgField=center&')"
        End If

    End Sub

    Private Sub but_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_search.Click
        If Chk_ROLE_NOUSE(sm) Then
            bt_save.Enabled = False
            btu_org.Disabled = True
            Common.RespWrite(Me, "<script>alert('您目前無此功能使用權限!');</script>")
            Exit Sub
        End If

        If RIDValue.Value <> "" Then
            tbSch.Visible = True
            get_Acct()
            get_plan()
        End If
    End Sub

    Private Sub ddl_years_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddl_years.SelectedIndexChanged
        get_plan()
    End Sub

    Private Sub rdoIsUsed_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdoIsUsed.SelectedIndexChanged
        tr_save.Visible = True
        If rdoIsUsed.SelectedValue = "N" Then
            tr_save.Visible = False
        End If
        but_search_Click(sender, e)
    End Sub

    Private Sub rdo_role_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles rdo_role.SelectedIndexChanged
        but_search_Click(sender, e)
    End Sub

    Private Sub lnkBtn_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles lnkBtn.Click
        Dim v_ddl_years As String = TIMS.GetListValue(ddl_years)
        Session("SYS_01_004_acc") = hid_acc.Value
        Session("SYS_01_004_years") = v_ddl_years 'ddl_years.SelectedValue
        Common.RespWrite(Me, "<script language=""javascript"">window.open(""SYS_01_004_view.aspx"",""openwind"",""height=300px,width=650px,scrollbars=yes,resizable=yes"");</script>")

        Me.RegisterStartupScript("clientScript", "<script language=""javascript"">showList('Datagrid1');showList('Datagrid2');</script>")
    End Sub

    Private Sub bt_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_save.Click
        'true:無權使用
        If Chk_ROLE_NOUSE(sm) Then
            bt_save.Enabled = False
            btu_org.Disabled = True
            Common.RespWrite(Me, "<script>alert('您目前無此功能使用權限!');</script>")
            Exit Sub
        End If

        Dim intCnt As Integer = 0
        Try
            Call Add_AccRWPlan() 'sda
            intCnt = 1
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= "/*  ex.ToString: */" & vbCrLf
            strErrmsg &= ex.ToString & vbCrLf
            'strErrmsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(Me, ex, strErrmsg)
        End Try

        'Dim conn As SqlConnection
        'Dim sda As New SqlDataAdapter
        'Dim trans As SqlTransaction = Nothing
        'Dim intCnt As Integer = 0
        'conn = DbAccess.GetConnection()
        'Try
        '    Call TIMS.OpenDbConn(conn)
        '    trans = conn.BeginTransaction()
        '    'sda.SelectCommand = New SqlCommand(sql, conn, trans)
        '    Call Add_AccRWPlan() 'sda
        '    intCnt = 1
        '    trans.Commit()
        'Catch ex As Exception
        '    trans.Rollback()
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        '    If Not trans Is Nothing Then trans.Dispose()
        '    If Not sda Is Nothing Then sda.Dispose()
        'End Try
        'Call TIMS.CloseDbConn(conn)

        If intCnt = 1 Then
            Common.MessageBox(Me.Page, "計畫賦予執行完成!")
            but_search_Click(sender, e)
        End If
    End Sub

    Private Sub Datagrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                For i As Integer = 0 To Datagrid1.Columns.Count - 1
                    e.Item.Cells(i).Visible = False
                Next

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim chk1 As CheckBox = e.Item.FindControl("chk1")
                Dim note1 As Label = e.Item.FindControl("note1")
                Dim note As String = ""
                Dim hid_account As HtmlInputHidden = e.Item.FindControl("hid_account")
                chk1.Enabled = True
                If Convert.ToString(drv("IsUsed")) = "N" Or Convert.ToString(drv("IsUsed")) = "" Then
                    chk1.Enabled = False
                    If note = "" Then
                        note = "[帳號未啟用]"
                    Else
                        note = note & "；" & "[帳號未啟用]"
                    End If
                End If
                If Convert.ToString(drv("Serialno")) = "" And Convert.ToString(drv("roleid")) <> "0" Then '超級使用者無自然人憑證不做驗証
                    chk1.Enabled = False
                    If note1.Text = "" Then
                        note = "[自然人憑證已清除]"
                    Else
                        note = note & "；" & "[自然人憑證已清除]"
                    End If
                End If

                If note <> "" Then
                    chk1.ToolTip = note
                End If

                If rdoIsUsed.SelectedValue <> "Y" Then
                    note1.Text = Convert.ToString(drv("note")) & note
                End If

                e.Item.Cells(1).ToolTip = "點擊查看此帳號 [" & Convert.ToString(drv("accname")) & "] 計畫賦予情形"
                e.Item.Cells(1).Attributes("onmouseover") = "this.style.cursor='hand';setAccValue('" & Convert.ToString(drv("account")) & "')"
                e.Item.Cells(1).Attributes.Add("onclick", "document.getElementById('lnkBtn').click();")
                hid_account.Value = Convert.ToString(drv("account"))

                chk1.Attributes.Add("onclick", "showList('Datagrid1');")
        End Select
    End Sub

    Private Sub Datagrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                For i As Integer = 0 To Datagrid2.Columns.Count - 1
                    e.Item.Cells(i).Visible = False
                Next

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hid_RID As HtmlInputHidden = e.Item.FindControl("hid_RID")
                Dim hid_planid As HtmlInputHidden = e.Item.FindControl("hid_planid")
                Dim lb_planname As Label = e.Item.FindControl("lb_planname")
                Dim chk2 As CheckBox = e.Item.FindControl("chk2")
                Dim note2 As Label = e.Item.FindControl("note2")
                Dim note As String = ""

                lb_planname.Text = Convert.ToString(drv("planname")) & Convert.ToString(drv("suborgname"))
                hid_RID.Value = Convert.ToString(drv("rid"))
                hid_planid.Value = Convert.ToString(drv("planid"))

                '署(局)帳號登入有最大權限
                If sm.UserInfo.RID <> "A" Then
                    '不同轄區 不可互相執行計畫賦予
                    If sm.UserInfo.RID <> Convert.ToString(drv("RID")).Chars(0) Then
                        chk2.Enabled = False
                        note = "登入者與賦予計劃的轄區不相同,無權限修改!"
                    Else
                        If drv("DistID") <> sm.UserInfo.DistID Then
                            chk2.Enabled = False
                            note = "登入者與賦予計劃的轄區不相同,無權限修改!"
                        End If
                    End If
                End If

                note2.Text = note
                chk2.Attributes.Add("onclick", "showList('Datagrid2');")
        End Select
    End Sub

    Public Shared Function Chk_ROLE_NOUSE(ByRef sm As SessionModel) As Boolean
        Dim flag_rst As Boolean = False 'false:可以使用  true:無權使用
        '角色階層 0：署(局) 1：分署(中心) 2：委訓 (縣市政府、一般培訓單位)
        If CInt(sm.UserInfo.LID) > 1 Then
            flag_rst = True '委訓單位不可使用 2
            Return flag_rst
        End If

        '角色 0：超級管理者 1：系統管理者 2：一級以上 3：一級 4：二級 5：承辦人 99：一般使用
        If sm.UserInfo.RoleID = 0 Or sm.UserInfo.RoleID = 1 Then '限定只有系統管理者可以使用
            flag_rst = False '可使用
        Else
            flag_rst = True '您目前無此功能使用權限
        End If
        Return flag_rst
    End Function

    Protected Sub Datagrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Datagrid1.SelectedIndexChanged

    End Sub
End Class
