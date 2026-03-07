Partial Class TR_05_002_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "TR_05_002_R"

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End
        If Not IsPostBack Then
            CreateItem()
            Button1.Attributes("onclick") = "return search();"
        End If

    End Sub

    Sub CreateItem()
        Syear1 = TIMS.GetSyear(Syear1)
        Common.SetListItem(Syear1, Now.Year)

        Syear2 = TIMS.GetSyear(Syear2)
        Common.SetListItem(Syear2, Now.Year)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim v_Syear1 As String = TIMS.GetListValue(Syear1)
        Dim v_Syear2 As String = TIMS.GetListValue(Syear2)
        Dim v_PlanType As String = TIMS.GetListValue(PlanType)
        Dim v_PlanTypeTxt As String = TIMS.GetListText(PlanType)

        Dim MyValue As String = ""
        MyValue &= "&Years1=" & v_Syear1
        MyValue &= "&Years2=" & v_Syear2
        MyValue &= "&PlanType=" & v_PlanType
        MyValue &= "&PlanTypeName=" & v_PlanTypeTxt
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub
End Class
