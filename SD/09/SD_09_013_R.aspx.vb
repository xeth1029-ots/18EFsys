Partial Class SD_09_013_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        If Not IsPostBack Then
            DistID = TIMS.Get_DistID(DistID)
            DistID.SelectedValue = sm.UserInfo.DistID
            Call CreateItem()
        End If
        Button1.Attributes("onclick") = "return print();"
    End Sub

    Sub CreateItem()
        For i As Integer = Now.Year To 2005 Step -1
            SYear.Items.Add(i)
            EYear.Items.Add(i)
        Next
        For i As Integer = 1 To 12
            SMonth.Items.Add(i)
            EMonth.Items.Add(i)
        Next
        Common.SetListItem(SYear, Now.Year)
        Common.SetListItem(SMonth, Now.Month)
        Common.SetListItem(EYear, Now.Year)
        Common.SetListItem(EMonth, Now.Month)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim SDate As Date = CDate(SYear.SelectedValue & "/" & SMonth.SelectedValue & "/01")
        Dim EDate As Date
        If (SYear.SelectedValue = EYear.SelectedValue And EMonth.SelectedValue = "12") Then
            EDate = CDate((EYear.SelectedValue + 1) & "/01/01")
        Else
            Select Case EMonth.SelectedValue
                Case "12"
                    EDate = CDate((EYear.SelectedValue + 1) & "/01/01")
                Case Else
                    EDate = CDate(EYear.SelectedValue & "/" & (EMonth.SelectedValue + 1) & "/01")
            End Select
        End If
        ReportQuery.PrintReport(Me, "Member", "SD_09_013_R", "DistID=" & DistID.SelectedValue & "&TPlanID=" & sm.UserInfo.TPlanID & "&SDate=" & SDate & "&EDate=" & EDate & "")
    End Sub

End Class
