Partial Class TC_01_003_print
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        Dim onClickScript1 As String = ""
        onClickScript1 = "if (document.getElementById('TB_classid').value==''){"
        onClickScript1 += ReportQuery.ReportScript(Me, "MultiBlock", "Class_Rpt_1", "DistID=" & sm.UserInfo.DistID & "&Years=" & sm.UserInfo.Years & "")
        onClickScript1 += "}"
        onClickScript1 += "else{"
        onClickScript1 += ReportQuery.ReportScript(Me, "MultiBlock", "Class_Rpt", "ClassID='+document.getElementById('TB_classid').value+'&DistID=" & sm.UserInfo.DistID & "&Years=" & sm.UserInfo.Years & "")
        onClickScript1 += "}"
        print.Attributes("onclick") = onClickScript1

        If Not Session("_search") Is Nothing Then
            Me.ViewState("_search") = Session("_search")
            Session("_search") = Nothing
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Session("_search") = Me.ViewState("_search")
        TIMS.Utl_Redirect1(Me, "TC_01_003.aspx?ID=" & Request("ID"))
    End Sub
End Class
