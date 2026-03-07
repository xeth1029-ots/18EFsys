Partial Class SD_15_011
    Inherits AuthBasePage

    'ReportQuery
    'SD_15_011_R '全部
    'SD_15_011_2R '含有轄區值
    ' (select NAME from v_OrgKind1 WHERE VALUE='$P!{Plankind}') Plankind,

    Const cst_printFN1 As String = "SD_15_011_R" '全部？
    Const cst_printFN2 As String = "SD_15_011_2R" '含有轄區值

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call cCreate1()
        End If
    End Sub

    Sub cCreate1()
        DistID = TIMS.Get_VDISTRICT(DistID, objconn)

        '沒有不區分
        OrgKind2 = TIMS.Get_RblOrgPlanKind(OrgKind2, objconn)
        Common.SetListItem(OrgKind2, "G")

        Years = TIMS.Get_Years(Years) '年度
        end_date.Text = Format(Now(), "yyyy/MM/dd") '開訓結束日

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)'AppStage = TIMS.Get_AppStage(AppStage)
        AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_AppStage2(AppStage), TIMS.Get_AppStage(AppStage))
        'AppStage = TIMS.Get_AppStage2(AppStage)

        trPlanKind.Style("display") = "none"
        trPackageType.Style("display") = "none"
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trPackageType.Style("display") = TIMS.cst_inline1 '"inline"
        Else
            '28:產業人才投資方案
            '計畫範圍 產投
            If sm.UserInfo.Years >= 2008 Then
                trPlanKind.Style("display") = TIMS.cst_inline1 '"inline"
            End If
        End If

    End Sub

    Function sUtl_CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        If start_date.Text <> "" AndAlso Not TIMS.IsDate1(start_date.Text) Then
            Errmsg += "開訓日期 起始日期格式有誤" & vbCrLf
        End If
        If end_date.Text <> "" AndAlso Not TIMS.IsDate1(end_date.Text) Then
            Errmsg += "開訓日期 迄止日期格式有誤" & vbCrLf
        End If
        If Errmsg <> "" Then Return False
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        Return rst
    End Function

    'Private Sub Print_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.ServerClick
    'End Sub

    Sub subPrint1()
        Dim SearchPlan1 As String = ""
        Dim sPackType As String = ""

        Dim MyValue As String = ""
        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sPackType = ""
            If OrgKind2.SelectedValue <> "" Then
                SearchPlan1 = OrgKind2.SelectedValue
            End If
        End If

        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            SearchPlan1 = ""
            If PackageType.SelectedValue <> "A" Then
                sPackType = PackageType.SelectedValue
            End If
        End If

        If DistID.SelectedValue = "0" Then
            '全部？
            'Plankind = PlanList.SelectedValue
            MyValue = ""
            MyValue += "&Years=" & Years.SelectedValue
            MyValue += "&TPlanID=" & sm.UserInfo.TPlanID
            MyValue += "&Plankind=" & SearchPlan1 '"",G,W
            MyValue += "&PackageType=" & sPackType '"",2,3

            MyValue += "&start_date=" & Me.start_date.Text
            MyValue += "&end_date=" & Me.end_date.Text
            '依申請階段 
            Dim v_AppStage As String = TIMS.GetListValue(AppStage)
            If v_AppStage <> "" AndAlso v_AppStage > "0" Then MyValue &= "&AppStage=" & v_AppStage
            MyValue += "&UserID=" & sm.UserInfo.UserID

            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

        Else
            'SD_15_011_2R
            'Me.DistID.SelectedValue含有轄區值
            MyValue = ""
            MyValue += "&Years=" & Years.SelectedValue
            MyValue += "&TPlanID=" & sm.UserInfo.TPlanID
            MyValue += "&PlanList=" & SearchPlan1 '"",G,W
            MyValue += "&PackageType=" & sPackType '"",2,3

            MyValue += "&DISTID=" & Me.DistID.SelectedValue
            MyValue += "&start_date=" & Me.start_date.Text
            MyValue += "&end_date=" & Me.end_date.Text
            '依申請階段 
            Dim v_AppStage As String = TIMS.GetListValue(AppStage)
            If v_AppStage <> "" AndAlso v_AppStage > "0" Then MyValue &= "&AppStage=" & v_AppStage
            MyValue += "&UserID=" & sm.UserInfo.UserID

            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
        End If
    End Sub

    '列印
    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""
        Call sUtl_CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        subPrint1()
    End Sub
End Class

