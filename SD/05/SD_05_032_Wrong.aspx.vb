Public Class SD_05_032_Wrong
    Inherits AuthBasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        'TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            Call create()
            Session("MyWrongTable") = Nothing
        End If
    End Sub

    Sub create()
        Dim dt As DataTable = Session("MyWrongTable")
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

End Class