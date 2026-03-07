Partial Class SD_03_003_R
    Inherits AuthBasePage

    'student_IdentityCard
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在---------------------------End

        Button1.Attributes("onclick") = "if(search()){"
        Button1.Attributes("onclick") += ReportQuery.ReportScript(Me, "list", "student_IdentityCard", "OCID='+document.getElementById('OCIDValue1').value+'")
        Button1.Attributes("onclick") += "}return false;"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
#Region "(No Use)"

        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=student_IdentityCard&path=TIMS&OCID=" & Me.OCIDValue1.Value & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        'Response.Redirect("" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=student_IdentityCard&path=TIMS&OCID=" & Me.OCIDValue1.Value & "")

#End Region
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class