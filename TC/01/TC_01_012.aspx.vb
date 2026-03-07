Partial Class TC_01_012
    Inherits System.Web.UI.Page

    Dim ProcessType As String
    'Dim FunDr As DataRow
    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End
        '分頁設定---------------Start
        PageControler1.PageDataGrid = DG_Org
        '分頁設定---------------End

        Button3.Visible = False '目前尚未開發此功能
        ProcessType = Request("ProcessType")
        bt_search.Attributes("onclick") = "return Search();"
        'check_bt_add()

        'If sm.UserInfo.RoleID <> 0 Then
        'End If
        If sm.UserInfo.FunDt Is Nothing Then
            Common.RespWrite(Me, "<script>alert('Session過期');</script>")
            Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        Else
            Dim FunDt As DataTable = sm.UserInfo.FunDt
            Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

            If FunDrArray.Length = 0 Then
                Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
                Common.RespWrite(Me, "<script>location.href='../../main.aspx';</script>")
            Else
                FunDr = FunDrArray(0)
                If FunDr("Adds") = "1" Then
                    check_add.Value = "1"
                Else
                    check_add.Value = "0"
                End If
                If FunDr("Sech") = "1" Then
                    bt_search.Enabled = True
                Else
                    bt_search.Enabled = False
                End If
                If FunDr("Del") = "1" Then
                    check_del.Value = "1"
                Else
                    check_del.Value = "0"
                End If
                If FunDr("Mod") = "1" Then
                    check_mod.Value = "1"
                Else
                    check_mod.Value = "0"
                End If
            End If
        End If

        If Not Me.IsPostBack Then
            yearlist = TIMS.GetSyear(yearlist)
            'Common.SetListItem(yearlist, Now.Year) 
            Common.SetListItem(yearlist, sm.UserInfo.Years) 'sm.UserInfo.Years 
            OrgKindList = TIMS.Get_OrgType(OrgKindList)

            '取得查詢條件
            If Not Session("_Search") Is Nothing Then
                TB_OrgName.Text = TIMS.GetMyValue(Session("_Search"), "TB_OrgName")
                TB_ComIDNO.Text = TIMS.GetMyValue(Session("_Search"), "TB_ComIDNO")
                TBCity.Text = TIMS.GetMyValue(Session("_Search"), "TBCity")
                city_code.Value = TIMS.GetMyValue(Session("_Search"), "city_code")
                Common.SetListItem(OrgKindList, TIMS.GetMyValue(Session("_Search"), "OrgKindList"))

                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("_Search"), "PageIndex")
                Dim MyValue As String = TIMS.GetMyValue(Session("_Search"), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If

                Common.SetListItem(yearlist, TIMS.GetMyValue(Session("_Search"), "yearlist"))
                If TIMS.GetMyValue(Session("_Search"), "Button1") = "True" Then
                    bt_search_Click(sender, e)
                End If

                Session("_Search") = Nothing
            End If
        End If
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
            Me.TxtPageSize.Text = 10
        End If
        Me.DG_Org.PageSize = Me.TxtPageSize.Text

        'Dim sqlAdapter As OracleDataAdapter
        'Dim dtOrgInfo As DataTable
        Dim relship As String
        relship = sm.UserInfo.Relship

        Dim dt As DataTable
        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr += " select distinct a.OrgID, d.Years PlanYear, c.Name, a.OrgName, g.CHARID, h.CHARNAME" & vbCrLf
        sqlstr += " from Org_orginfo a" & vbCrLf
        sqlstr += " join Auth_Relship b on a.OrgID  = b.OrgID" & vbCrLf
        sqlstr += " join ID_District c on b.DistID  = c.DistID" & vbCrLf
        sqlstr += " join view_LoginPlan d on b.PlanID = d.PlanID" & vbCrLf
        sqlstr += " join Org_OrgPlanInfo f on f.RSID = b.RSID" & vbCrLf
        sqlstr += " left join org_yearchar g  on a.OrgID = g.OrgID AND d.Years = g.PlanYear" & vbCrLf
        sqlstr += " left join key_ClassChar h on g.CHARID = h.CHARID" & vbCrLf
        sqlstr += " where 1=1" & vbCrLf

        sqlstr += " and b.OrgLevel >= '" & sm.UserInfo.OrgLevel & "'"
        If sm.UserInfo.RID <> "A" Then '局不受區域限制
            sqlstr += " and b.distid = '" & sm.UserInfo.DistID & "' "
        End If
        sqlstr += " and d.TPlanID = '" & sm.UserInfo.TPlanID & "' "

        If zip_code.Value <> "" Then
            sqlstr += " and f.ZipCode='" & Me.zip_code.Value & "' "
        ElseIf city_code.Value <> "" Then
            sqlstr += "and f.ZipCode in (select zipcode from ID_Zip where ctid='" & Me.city_code.Value & "') "
        End If

        If Me.TB_OrgName.Text <> "" Then
            sqlstr += " and a.OrgName like '%" & Me.TB_OrgName.Text & "%'"
        End If
        If Me.TB_ComIDNO.Text <> "" Then
            sqlstr += " and a.ComIDNO =  '" & Me.TB_ComIDNO.Text & "'"
        End If
        If Me.OrgKindList.SelectedValue <> "" Then
            sqlstr += " and a.OrgKind='" & Me.OrgKindList.SelectedValue & "'"
        End If
        If Me.yearlist.SelectedValue <> "" Then
            sqlstr += " and d.Years='" & Me.yearlist.SelectedValue & "'"
        End If
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        Panel.Visible = False
        DG_Org.Visible = False
        msg.Text = "查無資料!!"
        Me.bt_save.Visible = False
        If dt.Rows.Count > 0 Then
            Panel.Visible = True
            DG_Org.Visible = True
            msg.Text = ""
            'Me.bt_save.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OrgID"
            PageControler1.Sort = "OrgID"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        KeepSearch()
        TIMS.Utl_Redirect1(Me, "TC_01_012_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "") ''305
    End Sub

    Private Sub DG_Org_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_Org.ItemCommand
        'Dim plan_list, plan_sql As String
        'Dim sPLANYEAR As String = TIMS.GetMyValue(e.CommandArgument, "PLANYEAR")
        Dim sORGID As String = TIMS.GetMyValue(e.CommandArgument, "ORGID")
        'Dim sORGNAME As String = TIMS.GetMyValue(e.CommandArgument, "ORGNAME")
        'Dim sCHARID As String = TIMS.GetMyValue(e.CommandArgument, "CHARID")
        Select Case e.CommandName
            Case "edit"
                KeepSearch()
                Dim sRIDVALUE As String = Convert.ToString(TIMS.Get_RIDforOrgID(sORGID, "", objconn))
                TIMS.Utl_Redirect1(Me, "TC_01_012_add.aspx?ProcessType=Update&RIDVALUE=" & sRIDVALUE & "&" & e.CommandArgument & "")
            Case "del"
                '有計畫
                Dim dt As DataTable
                Dim sqlstr As String = ""
                sqlstr = "" & vbCrLf
                sqlstr += " SELECT 'x' x" & vbCrLf
                sqlstr += " FROM Org_YearChar a " & vbCrLf
                sqlstr += " join Org_OrgInfo b ON a.OrgID = b.OrgID " & vbCrLf
                sqlstr += " join Plan_PlanInfo c ON  c.ComIDNO = b.ComIDNO " & vbCrLf
                sqlstr += " join KEY_ClassProc d ON d.procid = c.procid and d.charid = a.charid" & vbCrLf
                sqlstr += " " & vbCrLf
                dt = DbAccess.GetDataTable(sqlstr, objconn)
                If dt.Rows.Count > 0 Then
                    Common.MessageBox(Me, "此機構已有計畫資料，不可以刪除!!")
                    Exit Sub
                End If
                KeepSearch()
                TIMS.Utl_Redirect1(Me, "TC_01_012_del.aspx?" & e.CommandArgument & "")
        End Select

    End Sub

    Private Sub DG_Org_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Org.ItemDataBound
        If Trim(Me.TxtPageSize.Text) <> "" And IsNumeric(Me.TxtPageSize.Text) Then
            If CInt(Me.TxtPageSize.Text) >= 1 Then
                Me.TxtPageSize.Text = Trim(Me.TxtPageSize.Text)
            Else
                Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
                Me.TxtPageSize.Text = 10
            End If
        Else
            Common.RespWrite(Me, "<script>alert('顯示列數不正確，以10 帶入');</script>")
            Me.TxtPageSize.Text = 10
        End If
        Me.DG_Org.PageSize = Me.TxtPageSize.Text

        Dim dr As DataRowView
        'Dim strScript1 As String
        Const cst_func As Integer = 5 '功能欄位在(0-5之間的5)
        dr = e.Item.DataItem
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DG_Org.PageSize * DG_Org.CurrentPageIndex

            'Dim but_edit, but_del As Button
            Dim but_add, but_edit, but_del As Button
            Dim MyValueAgm As String = "ORGID=" & dr("orgid") & "&ORGNAME=" & dr("OrgName") & "&PLANYEAR=" & dr("PlanYear") & "&CHARID=" & dr("charid") & "&ID=" & Request("ID") & ""

            but_add = e.Item.Cells(cst_func).FindControl("add_but") '新增
            but_add.CommandArgument = MyValueAgm
            If check_add.Value = "0" Then but_add.Enabled = False Else but_add.Enabled = True
            If IsDBNull(dr("CHARNAME")) = True Then but_add.Visible = True Else but_add.Visible = False

            but_edit = e.Item.Cells(cst_func).FindControl("edit_but") '修改
            but_edit.CommandArgument = MyValueAgm
            If check_mod.Value = "0" Then but_edit.Enabled = False Else but_edit.Enabled = True
            If IsDBNull(dr("CHARNAME")) = False Then but_edit.Visible = True Else but_edit.Visible = False

            but_del = e.Item.Cells(cst_func).FindControl("del_but") '刪除
            but_del.CommandArgument = "MyValueAgm"
            but_del.Attributes("onclick") = "return confirm('確定刪除此評鑑資料?');"
            If check_del.Value = "0" Then but_del.Enabled = False Else but_del.Enabled = True
            If IsDBNull(dr("CHARNAME")) = False Then but_del.Visible = True Else but_del.Visible = False
        End If
    End Sub

    Sub KeepSearch()
        Session("_Search") = "TB_OrgName=" & TB_OrgName.Text
        Session("_Search") += "&TB_ComIDNO=" & TB_ComIDNO.Text
        Session("_Search") += "&TBCity=" & TBCity.Text
        Session("_Search") += "&city_code=" & city_code.Value
        Session("_Search") += "&zip_code=" & zip_code.Value
        Session("_Search") += "&OrgKindList=" & OrgKindList.SelectedValue
        Session("_Search") += "&PageIndex=" & DG_Org.CurrentPageIndex + 1
        Session("_Search") += "&yearlist=" & yearlist.SelectedValue
        Session("_Search") += "&Button1=" & DG_Org.Visible
    End Sub

End Class
