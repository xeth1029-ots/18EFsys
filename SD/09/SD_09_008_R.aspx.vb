Partial Class SD_09_008_R
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            syears = TIMS.GetSyear(syears)
            Button2_Click(sender, e)
        End If
        'print_submit.Attributes("onclick") = "if(ReportPrint()){"
        'print_submit.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "Member", "behavior_list", "OCID='+document.getElementById('OCIDValue1').value+'&CJOB_UNKEY='+document.getElementById('cjobValue').value+'&Years='+document.getElementById('syears').value+'")
        'print_submit.Attributes("onclick") += "}return false;"

        print_submit.Attributes("onclick") = "javascript:return ReportPrint();"

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
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=behavior_list&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&Years=" & syears.SelectedValue & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "behavior_list", "syear=" & syears.SelectedValue & "&eyear=" & eyears.SelectedValue & "&RID=" & Me.RIDValue.Value & "&TPlanID=" & TPlan.SelectedValue & "&OCID=" & ClassName.SelectedValue & "")
        Dim MyValue As String = ""
        MyValue &= "&OCID=" & OCIDValue1.Value
        MyValue &= "&CJOB_UNKEY=" & cjobValue.Value
        MyValue &= "&Years=" & syears.SelectedValue
        'SD_09_008_R_Rpt.aspx
        Dim sUrl As String = ""
        sUrl = "SD_09_008_R_Rpt.aspx"
        ReportQuery.Redirect(Me, sUrl, MyValue)
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub
End Class
