Partial Class OB_01_002
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            '因有傳入值 yearlist.SelectedValue.ToString 故放此位置，才可讀到值
            DataGridTable.Visible = False
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        GetKeepSearch(sender, e)
        Page.RegisterStartupScript("ClearOrg1", "<script>ClearOrg();</script>")
    End Sub

    Sub KeepSearch()
        Session("_SearchStr") = "Class=OB_01_002"
        Session("_SearchStr") += "&center=" & center.Text.Trim
        Session("_SearchStr") += "&RIDValue=" & RIDValue.Value
        Session("_SearchStr") += "&orgid_value=" & orgid_value.Value
        Session("_SearchStr") += "&DeptName=" & DeptName.Text.Trim
        Session("_SearchStr") += "&memName=" & memName.Text.Trim
        Session("_SearchStr") += "&rblQualified=" & Me.rblQualified.SelectedValue

        'Session("_SearchStr") += "&radio1=" & radio1.Checked
        'Session("_SearchStr") += "&radio2=" & radio2.Checked
        'Session("_SearchStr") += "&radio3=" & radio3.Checked

        If DataGridTable.Visible Then
            Session("_SearchStr") += "&submit=1"
        Else
            Session("_SearchStr") += "&submit=0"
        End If
    End Sub

    Sub GetKeepSearch(ByVal sender As System.Object, ByVal e As System.EventArgs)
        '執行 Session("_SearchStr") 保留查詢值
        If Not Session("_SearchStr") Is Nothing Then
            ViewState("_SearchStr") = Session("_SearchStr")
            If TIMS.GetMyValue(ViewState("_SearchStr"), "Class") = "OB_01_002" Then
                center.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "center")
                RIDValue.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "RIDValue")
                orgid_value.Value = TIMS.GetMyValue(ViewState("_SearchStr"), "orgid_value")
                DeptName.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "DeptName")
                memName.Text = TIMS.GetMyValue(ViewState("_SearchStr"), "memName")

                Common.SetListItem(Me.rblQualified, TIMS.GetMyValue(ViewState("_SearchStr"), "rblQualified"))

                'If TIMS.GetMyValue(ViewState("_SearchStr"), "radio1") Then radio1.Checked = True
                'If TIMS.GetMyValue(ViewState("_SearchStr"), "radio2") Then radio2.Checked = True
                'If TIMS.GetMyValue(ViewState("_SearchStr"), "radio3") Then radio3.Checked = True

                If TIMS.GetMyValue(ViewState("_SearchStr"), "submit") = "1" Then
                    'btnQuery_Click(sender, e)
                    Call SchQuery()
                End If
            End If
            Session("_SearchStr") = Nothing
        End If
    End Sub

    Sub SchQuery(Optional ByVal sql As String = "")
        Dim dt As DataTable
        If sql <> "" Then
        Else
            sql = "" & vbCrLf
            sql += " select b.OrgName, a.* " & vbCrLf
            sql += " from OB_Member a " & vbCrLf
            sql += " JOIN OB_Org b on a.OrgSN=b.OrgSN " & vbCrLf
            sql += " WHERE 1=1 " & vbCrLf
            If sm.UserInfo.DistID <> "000" Then
                sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            End If

            If center.Text.Trim <> "" Then
                center.Text = center.Text.Trim
                sql += " AND b.OrgName like '%" & center.Text & "%'" & vbCrLf
            End If

            If orgid_value.Value <> "" Then
                sql += " AND a.OrgID='" & orgid_value.Value & "'" & vbCrLf
            End If

            If DeptName.Text.Trim <> "" Then
                DeptName.Text = DeptName.Text.Trim()
                sql += " AND b.DeptName like '%" & DeptName.Text & "%'" & vbCrLf
            End If

            Select Case Me.rblQualified.SelectedValue
                Case "N", "B", "A"
                    sql += " AND a.Qualified ='" & Me.rblQualified.SelectedValue & "'" & vbCrLf
            End Select

        End If

        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            DataGridTable.Visible = False
            msg.Text = "查無資料!!"
        End If

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        SchQuery()
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        'Response.Redirect("OB_01_002_add.aspx?ID=" & Request("ID") & "&Action=ADD")
        Dim url1 As String = "OB_01_002_add.aspx?ID=" & Request("ID") & "&Action=ADD"
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '檢查是否已被使用
    Function Chk_OB_TMember(ByVal msn As Integer) As Boolean
        Dim str_flag As String = Nothing
        Dim sql As String = "select msn from OB_TMember where msn=" & msn
        str_flag = DbAccess.ExecuteScalar(sql, objconn)
        If str_flag Is Nothing Then
            Chk_OB_TMember = False
        Else
            Chk_OB_TMember = True
        End If
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim drv As DataRowView = e.Item.DataItem
                'Dim btn_view As Button = e.Item.FindControl("btn_view")
                'btn_view.CommandArgument = drv("tsn")

                Dim btn_edit As Button = e.Item.FindControl("btn_edit")
                btn_edit.CommandArgument = drv("msn")

                Dim btn_del As Button = e.Item.FindControl("btn_del")
                btn_del.CommandArgument = drv("msn")
                btn_del.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                If Chk_OB_TMember(drv("msn")) Then
                    btn_del.Enabled = False
                    TIMS.Tooltip(btn_del, "")
                    TIMS.Tooltip(btn_del, "此工作小組已被選擇，無法刪除")
                End If

        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        KeepSearch()

        Dim sql As String = ""
        Select Case e.CommandName
            Case "del"
                sql = "delete OB_Member where msn='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功！")
                SchQuery()

            Case "edit"
                'Response.Redirect("OB_01_002_add.aspx?ID=" & Request("ID") & "&Action=EDIT&msn=" & e.CommandArgument)
                Dim url1 As String = "OB_01_002_add.aspx?ID=" & Request("ID") & "&Action=EDIT&msn=" & e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
            Case "view"
                'Response.Redirect("OB_01_002_add.aspx?ID=" & Request("ID") & "&Action=VIEW&msn=" & e.CommandArgument)
                Dim url1 As String = "OB_01_002_add.aspx?ID=" & Request("ID") & "&Action=VIEW&msn=" & e.CommandArgument
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

End Class
