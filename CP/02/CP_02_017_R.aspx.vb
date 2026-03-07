Partial Class CP_02_017_R
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
    Protected WithEvents SYear As System.Web.UI.WebControls.DropDownList
    Protected WithEvents SMonth As System.Web.UI.WebControls.DropDownList

    '注意: 下列預留位置宣告是 Web Form 設計工具需要的項目。
    '請勿刪除或移動它。
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        InitializeComponent()
    End Sub

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在--------------------------End

        Button1.Attributes("onclick") = "javascript:return search()"
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = Trim(start_date.Text)
        end_date.Text = Trim(end_date.Text)

        If start_date.Text <> "" Then
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "結訓日期 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            Errmsg += "結訓日期 起始日期 為必填" & vbCrLf
        End If

        If end_date.Text <> "" Then
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "結訓日期 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            Errmsg += "結訓日期 迄止日期 為必填" & vbCrLf
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then
                    Errmsg += "【結訓日期】的起日不得大於【結訓日期】的迄日!!" & vbCrLf
                End If
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim cGuid As String =   ReportQuery.GetGuid(Page)
        Dim Url As String =   ReportQuery.GetUrl(Page)
        Dim strScript As String
        Dim Y1, Y2, SYMD, EYMD As String
        Y1 = Convert.ToInt16(Convert.ToString(start_date.Text).Substring(0, 4)) - 1911
        SYMD = Y1 & "年" & Convert.ToString(start_date.Text).Substring(5, 2) & "月"
        Y2 = Convert.ToInt16(Convert.ToString(end_date.Text).Substring(0, 4)) - 1911
        EYMD = Y2 & "年" & Convert.ToString(end_date.Text).Substring(5, 2) & "月"

        strScript = "<script language=""javascript"">" + vbCrLf
        strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&filename=CP_02_017_R&path=TIMS&start_date=" & start_date.Text & "&end_date=" & end_date.Text & "&SYMD='+escape('" & SYMD & "')+'&EYMD='+escape('" & EYMD & "'));" + vbCrLf
        strScript += "</script>"

        Page.RegisterStartupScript("window_onload", strScript)

    End Sub
End Class
