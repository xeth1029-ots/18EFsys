Partial Class SD_09_005_R
    Inherits System.Web.UI.Page

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents Hidden1 As System.Web.UI.HtmlControls.HtmlInputHidden
    Protected WithEvents Hidden2 As System.Web.UI.HtmlControls.HtmlInputHidden

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
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            years = TIMS.GetSyear(years)
            months.Items.Add(New ListItem("===請選擇===", 0))
            For i As Integer = 1 To 12
                months.Items.Add(i)
            Next
        End If
        Button1.Attributes("onclick") = "javascript@return print();"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim strScript As String
        Dim cGuid As String =   ReportQuery.GetGuid(Page)
        Dim Url As String =   ReportQuery.GetUrl(Page)
        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=charge_list_2&path=TIMS&OCID=" & Me.OCIDValue1.Value & "&school_year=" & years.SelectedValue & "&school_month=" & months.SelectedValue & "');" + vbCrLf
        strScript += "</script>"
        Page.RegisterStartupScript("window_onload", strScript)
    End Sub
End Class
