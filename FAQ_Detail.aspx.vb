Partial Class FAQ_Detail
    Inherits AuthBasePage

    Const cst_FunctionName As String = "FunctionName"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(objconn) Then Exit Sub
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Me.ViewState("Search") = Session("MySearch")
            Session("MySearch") = Nothing
            Me.ViewState(cst_FunctionName) = Session("MyFunctionName")
            Session("MyFunctionName") = Nothing

            If ViewState("Search") <> "" Then
                Me.AcceptSearch.Value = ViewState("Search")
            End If

            Me.FunctionName.Text = ViewState(cst_FunctionName)
            Call create()
        End If

    End Sub

    Sub create()
        Dim dt As DataTable
        'Dim dr As DataRow
        Dim sqlstr As String
        If ViewState(cst_FunctionName) <> "" Then
            ViewState(cst_FunctionName) = TIMS.ClearSQM(ViewState(cst_FunctionName))
        End If

        sqlstr = ""
        sqlstr &= " SELECT * FROM Q_A"
        sqlstr &= " WHERE 1=1"
        If ViewState(cst_FunctionName) <> "" Then
            sqlstr &= " AND FunctionName like '%" & ViewState(cst_FunctionName) & "%'"
        End If
        sqlstr &= " ORDER BY SEQNO"
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        Me.msg.Text = "查無資料"
        Me.DataGrid1.Visible = False
        Me.PageControler1.Visible = False
        If dt.Rows.Count > 0 Then
            Me.msg.Text = ""
            Me.DataGrid1.Visible = True
            Me.PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "SEQNO" '以序號做排序
            PageControler1.ControlerLoad()
        End If
        '筆數
        Me.Num.Text = CStr(dt.Rows.Count)

    End Sub

    Private Sub Button1_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
        Session("MySearch") = Me.AcceptSearch.Value

        '回上一頁
        Dim url1 As String = "FAQ.aspx"
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

End Class
