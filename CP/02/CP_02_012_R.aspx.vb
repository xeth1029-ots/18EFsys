
Partial Class CP_02_012_R
    Inherits AuthBasePage

#Region " Web Form 設計工具產生的程式碼 "

    '此為 Web Form 設計工具所需的呼叫。
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub
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

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            syear = TIMS.GetSyear(syear)
            If sm.UserInfo.DistID = "000" Then
                'Tplan = TIMS.Get_TPlan(Tplan)
                Tplan = TIMS.Get_TPlan(Tplan, , 1)
            Else
                'Tplan = TIMS.Get_TPlan(Tplan)
                Tplan = TIMS.Get_TPlan(Tplan, , 1)
                Tplan.Enabled = False
            End If
            Tplan.SelectedValue = sm.UserInfo.TPlanID
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button1.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Button2.Attributes("onclick") = "javascript:return search()"

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim sql As String = "select Relship  from Auth_Relship where RID='" & RIDValue.Value & "'"
        'Dim dr As DataRow
        'dr = DbAccess.GetOneRow(sql)
        'Dim relship_str As String = dr("Relship")
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&filename=CP_02_012_Rpt&path=TIMS&Relship=" & relship_str & "&TPlanID=" & Tplan.SelectedValue & "&Years=" & syear.SelectedValue & "&start_date=" & start_date.Text & "&end_date=" & end_date.Text & "&start_date1=" & start_date1.Text & "&end_date1=" & end_date1.Text & "');" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&RID=" & RIDValue.Value
        MyValue &= "&Years=" & syear.SelectedValue
        MyValue &= "&TPlanID=" & Tplan.SelectedValue
        MyValue &= "&DistID=" & TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        MyValue &= "&start_date=" & start_date.Text
        MyValue &= "&end_date=" & end_date.Text
        MyValue &= "&start_date1=" & start_date1.Text
        MyValue &= "&end_date1=" & end_date1.Text
        'Relship=" & relship_str & "&TPlanID=" & Tplan.SelectedValue & "&Years=" & syear.SelectedValue & "&start_date=" & start_date.Text & "&end_date=" & end_date.Text & "&start_date1=" & start_date1.Text & "&end_date1=" & end_date1.Text & "
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "CP", "CP_02_012_Rpt", MyValue)
    End Sub
End Class
