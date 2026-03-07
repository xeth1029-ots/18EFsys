Partial Class Calendar
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        If Not IsPostBack Then
            NowYear.Attributes("onchange") = "ChangeDate();"
            NowMonth.Attributes("onchange") = "ChangeDate();"
        End If
    End Sub
End Class