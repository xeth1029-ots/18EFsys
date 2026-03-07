Partial Class TC_01_014_del
    Inherits System.Web.UI.Page

    'Dim Re_PlanID, Re_ComIDNO, Re_SeqNo As String
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

        Dim Re_PlanID As String = Request("PlanID")
        Dim Re_ComIDNO As String = Request("ComIDNO")
        Dim Re_SeqNo As String = Request("SeqNo")

        If Not IsPostBack Then
            If Not Session("_Search") Is Nothing Then Me.ViewState("_Search") = Session("_Search")
            'Session("_Search") = Nothing
        End If

        Dim sqlstr As String = ""
        sqlstr = " DELETE Plan_VerReport WHERE PlanID=" & Re_PlanID & " and ComIDNO='" & Re_ComIDNO & "' and SeqNo=" & Re_SeqNo
        DbAccess.ExecuteNonQuery(sqlstr, objconn)

        'Session("_Search") = Me.ViewState("_Search")
        If Not Me.ViewState("_Search") Is Nothing Then
            If Session("_Search") Is Nothing Then Session("_Search") = Me.ViewState("_Search")
        End If
        'Response.Redirect("TC_01_014.aspx?ProcessType=del&ID=" & Request("ID") & "")

        'Dim str, sqlstr_del As String
        'Dim objTrans As OracleTransaction
        'Dim dr As DataRow

        ''刪除[計畫名稱年度]-[機構名稱]
        'str = "刪除[" & Re_PlanID & "]-[" & Re_ComIDNO & "]-[" & Re_SeqNo
        'Try
        '    Call TIMS.OpenDbConn(objconn)
        '    objTrans = DbAccess.BeginTrans(objconn)
        '    sqlstr = " DELETE Plan_VerReport WHERE PlanID=" & Re_PlanID & " and ComIDNO='" & Re_ComIDNO & "' and SeqNo=" & Re_SeqNo
        '    DbAccess.ExecuteNonQuery(sqlstr_del, objTrans)
        '    DbAccess.CommitTrans(objTrans)
        '    Call TIMS.CloseDbConn(objconn)

        '    Session("_Search") = Me.ViewState("_Search")
        '    TIMS.Utl_Redirect1(Me, "TC_01_014.aspx?ProcessType=del&ID=" & Request("ID") & "")

        'Catch ex As Exception
        '    DbAccess.RollbackTrans(objTrans)
        '    Session("_Search") = Me.ViewState("_Search")
        '    Common.MessageBox(Page, "開班計畫表資料維護作業 刪除失敗!!")
        '    Throw ex
        'End Try

    End Sub

End Class
