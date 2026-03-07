Partial Class TC_01_012_del
    Inherits System.Web.UI.Page

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
        Dim Re_orgid, Re_OrgName, Re_planyear, Re_charid As String
        Re_orgid = Request("orgid")
        Re_OrgName = Request("OrgName")
        Re_planyear = Request("planyear")
        Re_charid = Request("charid")

        If Not IsPostBack Then
            Me.ViewState("_Search") = Session("_Search")
            Session("_Search") = Nothing
        End If

        Dim objTrans As OracleTransaction = Nothing
        ''刪除[計畫名稱年度]-[機構名稱]
        Dim str As String = "刪除[" & Re_planyear & "]-[" & Re_OrgName & "] "
        Try
            objTrans = DbAccess.BeginTrans(objconn)
            Dim sqlstr_del As String = ""
            sqlstr_del = " DELETE Org_YearChar WHERE OrgID ='" & Re_orgid & "' and PlanYear ='" & Re_planyear & "'"
            DbAccess.ExecuteNonQuery(sqlstr_del, objTrans)
            DbAccess.CommitTrans(objTrans)
        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            Common.MessageBox(Page, "訓練機構評鑑 刪除失敗!!")
            Throw ex
        End Try

        Session("_Search") = Me.ViewState("_Search")
        TIMS.Utl_Redirect1(Me, "TC_01_012.aspx?ProcessType=del&ID=" & Request("ID") & "")

    End Sub

End Class
