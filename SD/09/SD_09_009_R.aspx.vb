Partial Class SD_09_009_R
    Inherits AuthBasePage

    'price_list

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            syears = TIMS.GetSyear(syears)
            Button1_Click(sender, e)
        End If
        print_submit.Attributes("onclick") = "return ReportPrint();"
        'print_submit.Attributes("onclick") = "if(ReportPrint()){"
        ''print_submit.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "list", "price_list", "OCID='+document.getElementById('OCIDValue1').value+'&Years='+document.getElementById('syears').value+'")
        'print_submit.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "list", "price_list", "OCID=" & Me.OCIDValue1.Value & "&CJOB_UNKEY=" & cjobValue.Value & "&Years=" & syears.SelectedValue & "")
        'print_submit.Attributes("onclick") += "}return false;"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub print_submit_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print_submit.Click
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=price_list&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&Years=" & syears.SelectedValue & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

        Dim strParam As String = ""
        strParam = "OCID=" & Me.OCIDValue1.Value
        strParam &= "&CJOB_UNKEY=" & cjobValue.Value
        strParam &= "&Years=" & syears.SelectedValue
        strParam &= "&DISTID=" & Convert.ToString(sm.UserInfo.DistID)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "price_list", "OCID=" & Me.OCIDValue1.Value & "&CJOB_UNKEY=" & cjobValue.Value & "&Years=" & syears.SelectedValue)
        ReportQuery.PrintReport(Me, "list", "price_list", strParam)
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class