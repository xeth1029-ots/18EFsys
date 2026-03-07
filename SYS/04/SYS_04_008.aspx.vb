Partial Class SYS_04_008
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
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call create1()
        End If
    End Sub

    Sub create1()
        Dim sql As String = "SELECT * FROM Sys_VisitAlert"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If Not dr Is Nothing Then
            GRate1.Text = dr("GRate1").ToString
            YRate1.Text = dr("YRate1").ToString
            RRate1.Text = dr("RRate1").ToString
            GRate2.Text = dr("GRate2").ToString
            YRate2.Text = dr("YRate2").ToString
            RRate2.Text = dr("RRate2").ToString
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        'Dim conn As SqlConnection
        '2006/03/28 add conn by matt
        'conn = DbAccess.GetConnection
        sql = "SELECT * FROM Sys_VisitAlert"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
        Else
            dr = dt.Rows(0)
        End If
        dr("GRate1") = GRate1.Text
        dr("YRate1") = YRate1.Text
        dr("RRate1") = RRate1.Text
        dr("GRate2") = GRate2.Text
        dr("YRate2") = YRate2.Text
        dr("RRate2") = RRate2.Text
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
        Common.MessageBox(Me, "儲存成功")

    End Sub
End Class
