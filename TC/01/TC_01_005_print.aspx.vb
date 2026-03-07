Partial Class TC_01_005_print
    Inherits AuthBasePage

    'Cours_Rpt_1_R
    'Cours_Rpt
    Const cst_printFN1 As String = "Cours_Rpt_1_R"
    Const cst_printFN2 As String = "Cours_Rpt"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'print.Attributes("onclick") =     ReportQuery.ReportScript(Me, "MultiBlock", "Cours_Rpt", "TMID='+document.getElementById('trainValue').value+'&OrgID=" & sm.UserInfo.OrgID & "")
        If Not Session("MySreach") Is Nothing Then
            Me.ViewState("MySreach") = Session("MySreach")
        End If

        If Not IsPostBack Then
            yearlist = TIMS.Get_Years(yearlist, objconn)
            yearlist.Items.Insert(0, New ListItem("不區分", ""))
            If sm.UserInfo.LID < 2 Then
                yearlist.Enabled = False
            End If

        End If
    End Sub

    '列印。
    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        CourseName.Text = TIMS.ClearSQM(CourseName.Text)

        Dim sMyValue As String = ""
        sMyValue = ""
        sMyValue += "&TMID=" & trainValue.Value
        sMyValue += "&OrgID=" & sm.UserInfo.OrgID
        'If CourseName.Text <> "" Then CourseName.Text = Trim(CourseName.Text)
        If CourseName.Text <> "" Then
            sMyValue += "&CourseName=" & CourseName.Text
        End If
        If sm.UserInfo.LID = 2 Then '階層代碼【0:署(局) 1:分署(中心) 2:委訓】 NUMBER
            If yearlist.SelectedValue <> "" Then
                sMyValue += "&Years=" & yearlist.SelectedValue
            End If
        End If
        Dim sPrintFN As String = ""
        sPrintFN = cst_printFN2
        If RadioButtonList1.SelectedValue = 0 Then sPrintFN = cst_printFN1
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sPrintFN, sMyValue)

    End Sub

    '回上一頁。
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Session("MySreach") Is Nothing Then
            Session("MySreach") = Me.ViewState("MySreach")
        End If
        'Response.Redirect("TC_01_005.aspx?ID=" & Request("ID"))
        Dim url1 As String = "TC_01_005.aspx?ID=" & Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)

    End Sub
End Class
