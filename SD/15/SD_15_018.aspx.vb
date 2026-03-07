Public Class SD_15_018
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_15_018_R"

    Dim aNow As DateTime
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        aNow = TIMS.GetSysDateNow(objconn)
        '檢查Session是否存在 End

        '訓練機構
        If Not IsPostBack Then
            'msg.Text = ""
            'DataGridTable.Visible = False

            '年度
            ddlYears = TIMS.GetSyear(ddlYears)
            '月份
            ddlMonths = TIMS.Get_Month(ddlMonths, "")

            Common.SetListItem(ddlYears, sm.UserInfo.Years)
            Common.SetListItem(ddlMonths, Month(aNow))

            '轄區
            Dim sql As String
            sql = "SELECT DISTID,NAME FROM ID_DISTRICT ORDER BY DISTID"
            Dim dtD1 As DataTable = DbAccess.GetDataTable(sql, objconn)
            ddlDistID = TIMS.Get_DistID(ddlDistID, dtD1)
            Common.SetListItem(ddlDistID, sm.UserInfo.DistID)

            ddlDistID.Enabled = True
            If sm.UserInfo.DistID <> "000" Then '若登入者非署(局)署，鎖定轄區
                Common.SetListItem(ddlDistID, sm.UserInfo.DistID)
                ddlDistID.Enabled = False
            End If

            'btnSearch1.Attributes("onclick") = "javascript:return CheckSearch1();"
            'btnPrint1
            btnPrint1.Attributes("onclick") = "javascript:return CheckSearch1();"
        End If

    End Sub

    '列印
    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        'Call sUtl_ViewStateValue()  '設定 ViewState Value
        Dim sMonths As String = ddlMonths.SelectedValue
        If sMonths.Length < 2 Then sMonths = "0" & sMonths '補0

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&YYYYMM1=" & ddlYears.SelectedValue & "01"
        MyValue += "&YYYYMM2=" & ddlYears.SelectedValue & sMonths
        MyValue += "&YEARS=" & ddlYears.SelectedValue
        MyValue += "&DISTID=" & ddlDistID.SelectedValue
        MyValue += "&ORGKIND2=" & rblSearchPlan.SelectedValue

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub
End Class