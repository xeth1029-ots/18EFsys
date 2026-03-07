Partial Class CP_02_021_R
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

        Dim i As Integer
        If Not IsPostBack Then
            syear = TIMS.GetSyear(syear)
            smonth.Items.Add(New ListItem("===請選擇===", 0))
            For i = 1 To 12
                smonth.Items.Add(i)
            Next
            Common.SetListItem(syear, Now.Year)
            Common.SetListItem(smonth, Now.Month)
        End If

        Button1.Attributes("onclick") = "javascript:return print();"

    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Dim SDate, EDate As Date
        Dim YY, Y1, Y2, Y3 As String

        YY = Mid(syear.SelectedValue, 4, 1)
        Y1 = (YY - 1) + 2000
        Y2 = (syear.SelectedValue - 1911)
        Y3 = (Y2 - 1)

        SDate = CDate(syear.SelectedValue & "/" & smonth.SelectedValue & "/1")
        EDate = CDate(syear.SelectedValue & "/" & (smonth.SelectedValue + 1) & "/1")

        ReportQuery.PrintReport(Me, "Report", "CP_02_021_R", "Years=" & syear.SelectedValue & "&Y1=" & Y1 & "&Y2=" & Y2 & "&Y3=" & Y3 & "&smonth=" & smonth.SelectedValue & "&EDate=" & EDate & "")

    End Sub
End Class
