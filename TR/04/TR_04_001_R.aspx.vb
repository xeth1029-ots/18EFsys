Partial Class TR_04_001_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (眻諉婓 AuthBasePage ?燴, 祥蚚?脤 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()
        End If
        ' Button1.Attributes("onclick") = "javascript:return print();"
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

        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", "%"))
        DistID.Items(0).Selected = True
        If Not DistID.Items.FindByValue("000") Is Nothing Then
            DistID.Items.Remove(DistID.Items.FindByValue("000"))
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stdate_start, stdate_end, title As String
        stdate_start = Convert.ToString(SYear.SelectedValue) & "/" & Convert.ToString(SMonth.SelectedValue) & "/1"
        stdate_end = DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(FYear.SelectedValue & "/" & FMonth.SelectedValue & "/1")))
        title = Convert.ToString(SYear.SelectedValue) & "/" & Convert.ToString(SMonth.SelectedValue) & "~" & Convert.ToString(FYear.SelectedValue) & "/" & Convert.ToString(FMonth.SelectedValue) & "  輔助失業者參加提升數位能力研習計畫參加人數統計月報表"
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_001_R", "STDate=" & stdate_start & "&STDate2=" & stdate_end & "&DistID=" & DistID.SelectedValue & "&DistName=" & Server.UrlEncode(DistID.SelectedItem.Text) & "&Years=" & FYear.SelectedValue & "&title=" & Server.UrlEncode(title) & "")

        Dim MyValue As String = ""
        MyValue = "STDate=" & stdate_start & "&STDate2=" & stdate_end & "&DistID=" & DistID.SelectedValue & "&DistName=" & Server.UrlEncode(DistID.SelectedItem.Text) & "&Years=" & FYear.SelectedValue & "&title=" & Server.UrlEncode(title)
        ReportQuery.Redirect(Me, "TR_04_001_R_Rpt.aspx", MyValue)
    End Sub
End Class
