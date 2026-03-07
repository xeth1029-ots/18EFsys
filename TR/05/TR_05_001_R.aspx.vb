Partial Class TR_05_001_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
        End If

        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stitle As String = ""
        Dim etitle As String = ""
        If STDate1.Text <> "" Or STDate2.Text <> "" Then
            stitle = STDate1.Text + " ~ " + STDate2.Text
        End If
        If FTDate1.Text <> "" Or FTDate2.Text <> "" Then
            etitle = FTDate1.Text + " ~ " + FTDate2.Text
        End If

        Dim myvalue As String = ""
        TIMS.SetMyValue(myvalue, "Years1=", Syear.SelectedValue)
        TIMS.SetMyValue(myvalue, "&STDate1=", STDate1.Text)
        TIMS.SetMyValue(myvalue, "&STDate2=", STDate2.Text)
        TIMS.SetMyValue(myvalue, "&FTDate1=", FTDate1.Text)
        TIMS.SetMyValue(myvalue, "&FTDate2=", FTDate2.Text)
        TIMS.SetMyValue(myvalue, "&stitle=", stitle)
        TIMS.SetMyValue(myvalue, "&etitle=", etitle)
        ReportQuery.PrintReport(Me, "TR_05_001_R", myvalue)
    End Sub
End Class
