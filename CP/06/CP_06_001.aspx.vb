Partial Class CP_06_001
    Inherits AuthBasePage

    Dim objconn As SqlConnection

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
        PageControler1.PageDataGrid = DataGrid_Main
        '分頁設定---------------End

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button5.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button5.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PageControler1.Visible = False

            If Not Session("_SearchStr") Is Nothing Then
                center.Text = TIMS.GetMyValue(Session("_SearchStr"), "center")
                QuestNum.Text = TIMS.GetMyValue(Session("_SearchStr"), "QuestNum")

                PageControler1.PageIndex = 0
                'PageControler1.PageIndex = TIMS.GetMyValue(Session("_SearchStr"), "PageIndex")
                Dim MyValue As String = TIMS.GetMyValue(Session("_SearchStr"), "PageIndex")
                If MyValue <> "" AndAlso IsNumeric(MyValue) Then
                    MyValue = CInt(MyValue)
                    PageControler1.PageIndex = MyValue
                End If

                If TIMS.GetMyValue(Session("_SearchStr"), "submit") = "1" Then
                    Query_Click(sender, e)
                End If

                Session("_SearchStr") = Nothing
            End If
        End If
    End Sub

    Private Sub Btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Btn_Add.Click
        GetSearchStr()
        'Response.Redirect("CP_06_001_add.aspx?ID=" & Request("ID") & "&RID=" & Me.RIDValue.Value & "&Process=add")
        Dim url1 As String = "CP_06_001_add.aspx?ID=" & Request("ID") & "&RID=" & Me.RIDValue.Value & "&Process=add"
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Query.Click
        Call create("")
    End Sub

    Sub create(ByVal sql As String)
        If sql = "" Then
            sql = " select * from Org_GradedQuest where RID='" & Me.RIDValue.Value & "' "
            If QuestNum.Text <> "" Then
                sql += " and QuestNum= '" & QuestNum.Text & "' "
            End If
        End If
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        DataGrid_Main.Visible = False
        PageControler1.Visible = False
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            DataGrid_Main.Visible = True
            PageControler1.Visible = True
            msg.Text = ""

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OGQID"
            PageControler1.Sort = "QuestNum"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Const cst_Useing As Integer = 2
    Const cst_Null As String = "未啟用"
    Const cst_Y As String = "啟用"
    Const cst_N As String = "不啟用"

    Private Sub DataGrid_Main_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid_Main.ItemDataBound
        Dim drv As DataRowView
        drv = e.Item.DataItem

        Dim btn_edit As Button = e.Item.FindControl("Edit")
        Dim btn_del As Button = e.Item.FindControl("Del")
        Dim btn_view As Button = e.Item.FindControl("View")

        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            '啟用狀態

            If e.Item.Cells(cst_Useing).Text = "" Or IsDBNull(e.Item.Cells(cst_Useing).Text) Then
                e.Item.Cells(cst_Useing).Text = cst_Null
            ElseIf e.Item.Cells(cst_Useing).Text = "Y" Then
                e.Item.Cells(cst_Useing).Text = cst_Y
            Else
                e.Item.Cells(cst_Useing).Text = cst_N
            End If

            btn_edit.CommandArgument = drv("OGQID")

            btn_del.CommandArgument = drv("OGQID")
            btn_del.Attributes("onclick") = "return confirm('您確定要刪除這一筆資料?');"

            btn_view.CommandArgument = drv("OGQID")

        End If

    End Sub

    Private Sub DataGrid_Main_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid_Main.ItemCommand
        Select Case e.CommandName
            Case "edit"
                GetSearchStr()
                'Response.Redirect("CP_06_001_add.aspx?Process=update&OGQID=" & e.CommandArgument & "&RID=" & RIDValue.Value & "&ID=" & Request("ID"))
                Dim url1 As String = "CP_06_001_add.aspx?Process=update&OGQID=" & e.CommandArgument & "&RID=" & RIDValue.Value & "&ID=" & Request("ID")
                Call TIMS.Utl_Redirect(Me, objconn, url1)

            Case "del"
                Dim sql As String
                sql = "delete Org_GradedQuest where OGQID='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功！")
                Call create("")
            Case "view"
                GetSearchStr()
                'Response.Redirect("CP_06_001_add.aspx?Process=view&OGQID=" & e.CommandArgument & "&RID=" & RIDValue.Value & "&ID=" & Request("ID"))
                Dim url1 As String = "CP_06_001_add.aspx?Process=view&OGQID=" & e.CommandArgument & "&RID=" & RIDValue.Value & "&ID=" & Request("ID")
                Call TIMS.Utl_Redirect(Me, objconn, url1)

        End Select
    End Sub

    Sub GetSearchStr()
        Session("_SearchStr") = "center=" & center.Text & "&"
        Session("_SearchStr") += "QuestNum=" & QuestNum.Text & "&"
        Session("_SearchStr") += "PageIndex=" & DataGrid_Main.CurrentPageIndex + 1 & "&"
        If DataGrid_Main.Visible Then
            Session("_SearchStr") += "submit=1"
        Else
            Session("_SearchStr") += "submit=0"
        End If
    End Sub
End Class
