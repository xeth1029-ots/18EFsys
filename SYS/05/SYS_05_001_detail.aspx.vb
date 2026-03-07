Partial Class SYS_05_001_detail
    Inherits AuthBasePage

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
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Me.ViewState("ButtonFunctionID") = Session("ButtonFunctionID")
            Me.ViewState("FunctionID") = Session("FunctionID")
            Session("ButtonFunctionName") = Nothing
            Session("FunctionID") = Nothing
        End If

        Dim sqlstr As String = ""
        sqlstr = "SELECT * FROM Q_A "
        If ViewState("ButtonFunctionID") <> "" Then
            sqlstr += "WHERE FunID= '" & ViewState("ButtonFunctionID") & "' "
        End If

        '以序號做排序
        If Not IsPostBack Then
            Dim dt As DataTable
            dt = DbAccess.GetDataTable(sqlstr, objconn)
            Me.msg.Text = "查無資料"
            Me.DataGrid1.Visible = False
            PageControler1.Visible = False
            If dt.Rows.Count > 0 Then
                Me.msg.Text = ""
                Me.DataGrid1.Visible = True
                PageControler1.Visible = True

                'PageControler1.SqlDataCreate(sqlstr, "FunID")
                Me.DataGrid1.DataKeyField = "SeqNO"

                PageControler1.PageDataTable = dt
                PageControler1.PrimaryKey = "SeqNO"
                PageControler1.ControlerLoad()
            End If

            Me.Num.Text = CStr(dt.Rows.Count)
        End If
    End Sub

    Private Sub Button1_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
        Session("FunctionID") = Me.ViewState("FunctionID")

        '回上一頁
        TIMS.Utl_Redirect1(Me, "SYS_05_001.aspx")
    End Sub

    Protected Sub ECmd(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        '編輯
        Me.DataGrid1.EditItemIndex = e.Item.ItemIndex
        Me.PageControler1.CreateData()
    End Sub

    Protected Sub CCmd(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        '取消 
        Me.DataGrid1.EditItemIndex = -1
        Me.PageControler1.CreateData()
    End Sub

    Protected Sub DCmd(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        '刪除 
        Dim sqlstr As String

        sqlstr = "DELETE Q_A WHERE SeqNO=" & DataGrid1.DataKeys(e.Item.ItemIndex)
        DbAccess.ExecuteNonQuery(sqlstr, objconn)
        PageControler1.CreateData()
    End Sub

    Protected Sub UCmd(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        '更新
        Dim sqlstr As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim sqlAdapter As SqlDataAdapter = Nothing

        sqlstr = "SELECT * FROM Q_A WHERE SeqNO=" & DataGrid1.DataKeys(e.Item.ItemIndex)
        dt = DbAccess.GetDataTable(sqlstr, sqlAdapter, objconn)
        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0)
            Dim FunctionListTemp As DropDownList = e.Item.FindControl("functionlist")
            Dim QuestionTextBox As TextBox = e.Item.FindControl("QuestionTextBox")
            Dim PostUnitTextBox As TextBox = e.Item.FindControl("PostUnitTextBox")
            Dim AnswerTextBox As TextBox = e.Item.FindControl("AnswerTextBox")

            dr("FunID") = FunctionListTemp.SelectedValue
            dr("FunctionName") = FunctionListTemp.SelectedItem.Text
            dr("Question") = QuestionTextBox.Text
            dr("PostUnit") = PostUnitTextBox.Text
            dr("PostDate") = FormatDateTime(Now(), 2)
            dr("Deal") = AnswerTextBox.Text
            dr("ModifyAcct") = sm.UserInfo.UserID
            DbAccess.UpdateDataTable(dt, sqlAdapter)
            Me.DataGrid1.EditItemIndex = -1
        End If

        PageControler1.CreateData()

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim myTableCell As TableCell
        Dim myDeleteButton As Button
        Dim FunctionListTemp As DropDownList
        Dim dr As DataRow
        Dim dt As DataTable
        Dim sqlstr As String

        sqlstr = "SELECT * FROM ID_Function "
        sqlstr += "WHERE SPage is not null "
        sqlstr += "ORDER BY FunID "
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        If e.Item.ItemType = ListItemType.Item Or e.Item.ItemType = ListItemType.AlternatingItem Then
            myTableCell = e.Item.Cells(3)
            myDeleteButton = myTableCell.Controls(0)
            myDeleteButton.Attributes.Add("onclick", "return confirm('您確定要刪除嗎?');")
        End If

        '編輯模式時的功能項目
        If e.Item.ItemType = ListItemType.EditItem Then
            FunctionListTemp = e.Item.FindControl("functionlist")
            FunctionListTemp.DataSource = dt
            FunctionListTemp.DataTextField = "Name"
            FunctionListTemp.DataValueField = "FunID"
            FunctionListTemp.DataBind()
            FunctionListTemp.Items.Insert(0, New ListItem("不區分功能", "1"))
            FunctionListTemp.Items.Insert(1, New ListItem("外網相關問題", "2"))
            FunctionListTemp.Items.Insert(2, New ListItem("其他功能", "3"))

            Select Case Me.ViewState("ButtonFunctionID")
                Case 1
                    FunctionListTemp.SelectedValue = "1"
                Case 2
                    FunctionListTemp.SelectedValue = "2"
                Case 3
                    FunctionListTemp.SelectedValue = "3"
                Case Else
                    sqlstr = "SELECT * FROM ID_Function "
                    sqlstr += "WHERE SPage is not null "
                    sqlstr += "AND FunID=" & Me.ViewState("ButtonFunctionID") & " "
                    sqlstr += "ORDER BY FunID "
                    dt = DbAccess.GetDataTable(sqlstr, objconn)
                    dr = dt.Rows(0)
                    FunctionListTemp.SelectedValue = dr("FunID")
            End Select
        End If
    End Sub

End Class
