Partial Class SYS_05_001
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

        If Not Page.IsPostBack Then
            Dim dt As DataTable
            'Dim dr As DataRow
            Dim sqlstr As String

            sqlstr = "SELECT * FROM ID_Function WHERE SPage is not null ORDER BY FunID "
            dt = DbAccess.GetDataTable(sqlstr, objconn)

            Me.FunctionNameList.DataSource = dt
            Me.FunctionNameList.DataTextField = "Name"
            Me.FunctionNameList.DataValueField = "FunID"
            Me.FunctionNameList.DataBind()
            Me.FunctionNameList.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Me.FunctionNameList.Items.Insert(1, New ListItem("全部", "0"))
            Me.FunctionNameList.Items.Insert(2, New ListItem("不區分功能", "1"))
            Me.FunctionNameList.Items.Insert(3, New ListItem("外網相關問題", "2"))
            Me.FunctionNameList.Items.Insert(4, New ListItem("其他功能", "3"))

            If Not Session("_search") Is Nothing Then
                Dim MyArray As Array = Split(Session("_search"), "&")
                For i As Integer = 0 To MyArray.Length - 1
                    Dim MyItem As String = Split(MyArray(i), "=")(0)
                    Dim MyValue As String = Split(MyArray(i), "=")(1)

                    Select Case MyItem
                        Case "FunctionNameList"
                            FunctionNameList.SelectedIndex = MyValue
                        Case "PageIndex"
                            PageControler1.PageIndex = MyValue
                    End Select
                Next

                Session("_search") = Nothing
            End If

            create()
        End If

    End Sub

    Sub create()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        Dim sqlstr As String = ""
        sqlstr = ""
        sqlstr &= " SELECT a.FunID,a.FunctionName,COUNT(1) FunctionCount FROM Q_A a "
        sqlstr += " LEFT JOIN ID_Function b ON  a.FunID=b.FunID "
        sqlstr += " GROUP BY a.FunID ,a.FunctionName "
        'sqlstr += "ORDER BY a.FunID "
        If Me.FunctionNameList.SelectedIndex > 1 Then
            sqlstr = ""
            sqlstr &= " SELECT a.FunID,a.FunctionName,COUNT(1) FunctionCount FROM Q_A a "
            sqlstr += " LEFT JOIN ID_Function b ON  a.FunID=b.FunID "
            sqlstr += " WHERE a.FunID = '" & Me.FunctionNameList.SelectedValue & "' "
            sqlstr += " GROUP BY a.FunID ,a.FunctionName "
            'sqlstr += " ORDER BY a.FunID "
        End If
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        msg.Text = "查無資料"
        DataGrid1.Visible = False
        PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGrid1.Visible = True
            PageControler1.Visible = True

            'PageControler1.SqlDataCreate(sqlstr, "FunID")
            PageControler1.PageDataTable = dt
            PageControler1.Sort = "FunID"
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound1(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '序號
                e.Item.Cells(0).Text = (Me.DataGrid1.PageSize * Me.DataGrid1.CurrentPageIndex) + e.Item.ItemIndex + 1

                Dim lbtDetail As LinkButton = e.Item.FindControl("lbtDetail")
                lbtDetail.CommandArgument = e.Item.Cells(4).Text
        End Select
    End Sub

    Protected Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand

        '細節 按鈕對應的功能項目名稱
        Session("ButtonFunctionID") = e.CommandArgument

        '下拉選單的功能項目名稱
        KeepSearch()

        '細節 按鈕
        TIMS.Utl_Redirect1(Me, "SYS_05_001_detail.aspx")
    End Sub

    Private Sub reset_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        'RESET
        Me.FunctionNameList.SelectedIndex = 0
        Me.msg.Text = ""
        Me.ViewState("FunctionID") = ""
        create()
    End Sub

    Private Sub bt_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_add.Click
        '新增 按鈕
        KeepSearch()
        TIMS.Utl_Redirect1(Me, "SYS_05_001_add.aspx")
    End Sub

    Private Sub bt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_search.Click
        '查詢
        Me.msg.Text = ""
        create()
    End Sub

    '保留查詢值
    Sub KeepSearch()
        Session("_search") = "FunctionNameList=" & FunctionNameList.SelectedIndex
        Session("_search") += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
    End Sub
End Class
