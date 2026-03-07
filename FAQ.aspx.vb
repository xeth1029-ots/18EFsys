Partial Class FAQ
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(objconn) Then Exit Sub
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Me.msg.Text = ""
            If Not Session("MySearch") Is Nothing Then
                TextBox1.Text = TIMS.GetMyValue(Session("MySearch"), "TextBox1")
                Me.ViewState("PageIndex") = TIMS.GetMyValue(Session("MySearch"), "PageIndex")
                Session("MySearch") = Nothing

                If IsNumeric(Me.ViewState("PageIndex")) Then
                    PageControler1.PageIndex = Me.ViewState("PageIndex")
                    'PageControler1.CreateData()
                End If

                'create()
                ''If Me.PageControler1.Visible AndAlso IsNumeric(Me.ViewState("PageIndex")) Then
                ''    PageControler1.PageIndex = Me.ViewState("PageIndex")
                ''    PageControler1.CreateData()
                ''End If
                'Else
                'create()
            End If
            Call create()

        End If

    End Sub

    Sub create()
        Dim sqlstr As String
        Dim sFunctionName As String = ""
        sFunctionName = Convert.ToString(TextBox1.Text).Trim

        Dim rtnPath As String = ""
        rtnPath = Request.FilePath
        If TIMS.CheckInput(sFunctionName) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg2, rtnPath)
            Exit Sub
        End If
        sFunctionName = sFunctionName.Replace("'", "''").Trim

        sqlstr = ""
        sqlstr += " SELECT FunID"
        sqlstr += " ,FunctionName"
        sqlstr += " ,COUNT(1) FunctionCount "
        sqlstr += " FROM Q_A "
        sqlstr += " where FunctionName like '%" & sFunctionName & "%' "
        sqlstr += " GROUP BY FunID"
        sqlstr += " ,FunctionName "

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        Me.msg.Text = "查無資料"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.msg.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "FunID"
            PageControler1.Sort = "FunctionName"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim btn As Button = e.Item.FindControl("Detail")
                btn.CommandArgument = e.Item.Cells(1).Text
        End Select
    End Sub

    Protected Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        KeepSearch()
        Session("MyFunctionName") = e.CommandArgument

        '詳細 按鈕
        Dim url1 As String = "FAQ_Detail.aspx"
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        Call create()
    End Sub

    Private Sub reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles reset.Click
        Me.TextBox1.Text = ""
        Call create()
    End Sub

    Sub KeepSearch()
        Session("MySearch") = "TextBox1=" & Me.TextBox1.Text
        Session("MySearch") += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
