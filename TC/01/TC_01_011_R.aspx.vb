Partial Class TC_01_011_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "TC_01_011_R"

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        'Dim strJsOnClick As String = ""
        'strJsOnClick = ""
        'strJsOnClick &= "if(ReportPrint()){"
        'strJsOnClick +=     ReportQuery.ReportScript(Me, "MultiBlock", "TC_01_011_R", "Years='+document.getElementById('yearlist').value+'&TPlanID='+document.getElementById('planlist').value+'&DistID='+document.getElementById('DistrictList').value+'&CTID='+document.getElementById('city_code').value+'&ComIDNO='+document.getElementById('ComIDNO').value+'&OrgKind='+document.getElementById('OrgTypeList').value+'&OrgName='+escape(document.getElementById('OrgName').value)+'")
        'strJsOnClick += "}return false;"
        'btnPrint.Attributes("onclick") = strJsOnClick

        If Not Me.IsPostBack Then
            Call Create1()
        End If

    End Sub

    Sub Create1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        yearlist = TIMS.GetSyear(yearlist)
        '2005/4/8--Ellen年度帶預設值Ellen 
        Common.SetListItem(yearlist, sm.UserInfo.Years)

        '取得訓練計畫
        If yearlist.SelectedValue <> "" Then
            planlist = TIMS.sUtl_ShowTPlanID(yearlist.SelectedValue, planlist, objconn)
        Else
            planlist = TIMS.sUtl_ShowTPlanID(sm.UserInfo.Years, planlist, objconn)
        End If
        '設定登入訓練計畫
        Common.SetListItem(planlist, sm.UserInfo.TPlanID)
        'Try
        '    Me.planlist.SelectedValue = sm.UserInfo.TPlanID
        'Catch ex As Exception
        'End Try

        '取得轄區
        DistrictList = TIMS.Get_DistID(DistrictList)

        'Dim Sqlstr As String = ""
        'Dim dr As DataRow
        If sm.UserInfo.LID <> "000" Then
            Me.planlist.Enabled = False
            Common.SetListItem(DistrictList, sm.UserInfo.DistID)
            'Try
            '    DistrictList.SelectedValue = sm.UserInfo.DistID
            'Catch ex As Exception
            'End Try
            DistrictList.Enabled = False
        End If

        '取得機構別
        Dim Sqlstr As String = ""
        Sqlstr = "SELECT ORGTYPEID,NAME FROM KEY_ORGTYPE ORDER BY ORGTYPEID"
        Me.OrgTypeList.Items.Clear()
        DbAccess.MakeListItem(Me.OrgTypeList, Sqlstr, objconn)
        Me.OrgTypeList.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        'Dim cGuid As String =   ReportQuery.GetGuid(Page)
        'Dim Url As String =   ReportQuery.GetUrl(Page)
        'Dim strScript As String
        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=MultiBlock&path=TIMS&filename=TC_01_011_R&Years=" & Me.yearlist.SelectedValue & "&TPlanID=" & Me.planlist.SelectedValue & "&DistID=" & Me.DistrictList.SelectedValue & "&CTID=" & Me.city_code.Value & "&ComIDNO=" & Me.ComIDNO.Text & "&OrgKind=" & Me.OrgTypeList.SelectedValue & "&OrgName='+escape('" & Me.OrgName.Text & "'));" + vbCrLf
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        TIMS.CloseDbConn(objconn)
        Dim sMyValue As String = ""
        TIMS.SetMyValue(sMyValue, "Years", Me.yearlist.SelectedValue)
        TIMS.SetMyValue(sMyValue, "TPlanID", Me.planlist.SelectedValue)
        TIMS.SetMyValue(sMyValue, "DistID", Me.DistrictList.SelectedValue)
        TIMS.SetMyValue(sMyValue, "CTID", Me.city_code.Value)
        TIMS.SetMyValue(sMyValue, "ComIDNO", Me.ComIDNO.Text)
        TIMS.SetMyValue(sMyValue, "OrgKind", Me.OrgTypeList.SelectedValue)
        TIMS.SetMyValue(sMyValue, "OrgName", OrgName.Text)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, sMyValue)
    End Sub

    Protected Sub yearlist_SelectedIndexChanged(sender As Object, e As EventArgs) Handles yearlist.SelectedIndexChanged
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        planlist = TIMS.sUtl_ShowTPlanID(yearlist.SelectedValue, planlist, objconn)
        Common.SetListItem(planlist, sm.UserInfo.TPlanID)
        'Try
        '    Me.planlist.SelectedValue = sm.UserInfo.TPlanID
        'Catch ex As Exception
        'End Try
    End Sub
End Class
