Partial Class SD_04_008_c
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
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在--------------------------End

        Dim sql As String
        Dim dr As DataRow

        sql = "SELECT ProPercent FROM Log_Thread WHERE ProcessID='" & Request("ProcessID") & "' Order By ProDate Desc"
        dr = DbAccess.GetOneRow(sql)
        Percent.Text = dr("ProPercent")

        Select Case dr("ProPercent")
            Case 100
                Common.RespWrite(Me, "<script>if(confirm('執行結束,您是否要導到「師資衝堂查詢」的功能頁面?')){opener.location.href='SD_04_008.aspx?ID=" & Request("ID") & "';}window.close();</script>")
            Case -1
                Common.RespWrite(Me, "<script>alert('執行失敗,請聯絡開發人員!');</script>")
        End Select

    End Sub

End Class
