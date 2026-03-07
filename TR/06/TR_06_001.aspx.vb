Partial Class TR_06_001
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    'http://disemp.evta.gov.tw/
    'http://www3.evta.gov.tw/disemp/index.aspx

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        Page.RegisterStartupScript("", "<script>window.open('TR_06_001_DOC.htm','','width=700,height=700,location=0,status=0,menubar=0,scrollbars=1,resizable=0');</script>")

    End Sub

End Class
