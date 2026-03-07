Partial Class TC_10_001_Wrong
    Inherits AuthBasePage 'Inherits System.Web.UI.Page

    'Dim PageControler1 As New PageControler
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            Response.End()
        End If
        '檢查Session是否存在 Start
        'TIMS.CheckSession(Me)
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt)  '2011取得功能按鈕權限值
        '檢查Session是否存在 End

        PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid1
        If Not IsPostBack Then
            create()
            Session("MyWrongTable") = Nothing
        End If
    End Sub

    Sub create()
        If Session("MyWrongTable") Is Nothing Then Return
        Dim dt As DataTable = Session("MyWrongTable")
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

End Class
