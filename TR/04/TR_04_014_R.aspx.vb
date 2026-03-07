Partial Class TR_04_014_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
        End If

    End Sub

    Sub CreateItem()
        For i As Integer = Now.Year To 2005 Step -1
            SYear.Items.Add(i)
            FYear.Items.Add(i)
        Next
        For i As Integer = 1 To 12
            SMonth.Items.Add(i)
            FMonth.Items.Add(i)
        Next
        Common.SetListItem(SMonth, Now.Month - 3)
        Common.SetListItem(FMonth, Now.Month)
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Years, stdate_start, stdate_end As String

        Years = Convert.ToString(SYear.SelectedValue - 1911)
        If SMonth.SelectedValue <= "9" Then
            stdate_start = Convert.ToString(SYear.SelectedValue) & "0" & Convert.ToString(SMonth.SelectedValue)
        Else
            stdate_start = Convert.ToString(SYear.SelectedValue) & Convert.ToString(SMonth.SelectedValue)
        End If

        If FMonth.SelectedValue <= "9" Then
            stdate_end = Convert.ToString(FYear.SelectedValue) & "0" & Convert.ToString(FMonth.SelectedValue)
        Else
            stdate_end = Convert.ToString(FYear.SelectedValue) & Convert.ToString(FMonth.SelectedValue)
        End If

        Dim TPlanIDStr As String = ""
        For Each item As ListItem In TPlanID.Items
            If item.Selected = True Then
                If TPlanIDStr <> "" Then TPlanIDStr &= ","
                TPlanIDStr &= "\'" & item.Value & "\'"
            End If
        Next

        ReportQuery.PrintReport(Me, "TR", "TR_04_014_R", "Years=" & Years & "&FTDate=" & stdate_start & "&FTDate2=" & stdate_end & "&TPlanID=" & TPlanIDStr)
    End Sub
End Class
