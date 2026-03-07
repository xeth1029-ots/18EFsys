Partial Class SYS_05_001_add
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

        If Not Page.IsPostBack Then
            Dim dt As DataTable
            'Dim dr As DataRow
            Dim sqlstr As String

            '保留查詢字串
            If Not Session("_search") Is Nothing Then
                Me.AcceptSearch.Value = Session("_search")
                Session("_search") = Nothing
            End If

            sqlstr = "SELECT * FROM ID_Function WHERE SPage is not null ORDER BY FunID "
            dt = DbAccess.GetDataTable(sqlstr, objconn)

            Me.FunctionList.DataSource = dt
            Me.FunctionList.DataTextField = "Name"
            Me.FunctionList.DataValueField = "FunID"
            Me.FunctionList.DataBind()
            Me.FunctionList.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Me.FunctionList.Items.Insert(2, New ListItem("不區分功能", "1"))
            Me.FunctionList.Items.Insert(3, New ListItem("外網相關問題", "2"))
            Me.FunctionList.Items.Insert(4, New ListItem("其他功能", "3"))
        End If

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Session("_search") = Me.AcceptSearch.Value

        '回上一頁
        TIMS.Utl_Redirect1(Me, "SYS_05_001.aspx")
    End Sub

    Private Sub bt_addrow_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_addrow.Click
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim sqlstr As String = ""
        'Dim objconn As SqlConnection

        '新增實作
        sqlstr = " select * from Q_A "
        'dr = DbAccess.GetOneRow(sqlstr, objconn)
        dt = DbAccess.GetDataTable(sqlstr, sqlAdapter, objconn)
        dr = dt.NewRow
        dt.Rows.Add(dr)

        dr("FunID") = Me.FunctionList.SelectedValue
        dr("FunctionName") = Me.FunctionList.SelectedItem.Text
        dr("PostDate") = FormatDateTime(Now(), 2)
        dr("PostUnit") = Me.Textbox2.Text
        dr("Question") = Me.TextBox3.Text
        dr("Deal") = Me.TextBox4.Text
        dr("ModifyAcct") = sm.UserInfo.UserID

        DbAccess.UpdateDataTable(dt, sqlAdapter)

        Session("_search") = Me.AcceptSearch.Value
        TIMS.Utl_Redirect1(Me, "SYS_05_001.aspx")
    End Sub
End Class
