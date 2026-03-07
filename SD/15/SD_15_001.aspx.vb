Partial Class SD_15_001
    Inherits AuthBasePage

    'SQControl
    'SD_15_001_1~5
    'SD_15_012(綜合查詢統計表)
    'SD_15_001*.jrxml
    Const cst_prt_轄區 As String = "SD_15_001_1"
    'Const cst_prt_訓練類別 As String = "SD_15_001_2"
    Const cst_prt_職能別 As String = "SD_15_001_3"
    Const cst_prt_訓練單位類別 As String = "SD_15_001_4"
    Const cst_prt_訓練單位 As String = "SD_15_001_5"
    'SELECT * FROM Key_Depot WHERE DEPID='14'
    'SELECT * FROM Key_Business WHERE DEPID='14' AND status is null
    'SELECT KID14,KNAME14  FROM V_PLAN_DEPOT WHERE KID14 IS NOT NULL AND ROWNUM <=10
    Const cst_prt_課程分類 As String = "SD_15_001_6"
    Const cst_prt_生產力4 As String = "SD_15_001_7"
    Const cst_prt_新興產業 As String = "SD_15_001_8"
    Const cst_prt_重點服務業 As String = "SD_15_001_9"
    Const cst_prt_新興智慧型產業 As String = "SD_15_001_10" '新興智慧型產業
    'Const cst_prt_新南向政策 As String = "SD_15_001_11" '新南向政策
    Const cst_inline1 As String = "" '"inline"

    Sub print1()
        Dim MyValue As String = ""
        Dim SearchOrgKind As String = ""
        Dim SchOrgKind2 As String = ""

        '28:產業人才投資方案
        'SearchOrgKind = ""
        Dim v_SearchMode As String = TIMS.GetListValue(SearchMode)
        Dim v_PrintMode As String = TIMS.GetListValue(PrintMode)
        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim fileName1 As String = ""
        Select Case v_PrintMode'PrintMode.SelectedValue
            Case "1"
                fileName1 = cst_prt_轄區 '轄區
            Case "2"
                 'fileName1 = cst_prt_訓練類別 '訓練類別
            Case "3"
                fileName1 = cst_prt_職能別 '職能別
            Case "4"
                fileName1 = cst_prt_訓練單位類別 '訓練單位類別
            Case "5"
                fileName1 = cst_prt_訓練單位 '訓練單位
            Case "6"
                fileName1 = cst_prt_課程分類
            Case "7"
                fileName1 = cst_prt_生產力4
            Case "8"
                fileName1 = cst_prt_新興產業
            Case "9"
                fileName1 = cst_prt_重點服務業
            Case "10"
                fileName1 = cst_prt_新興智慧型產業
            Case "11"
                'fileName1 = cst_prt_新南向政策
        End Select

        '28:產業人才投資方案
        SearchOrgKind = ""
        SchOrgKind2 = "G,W" '""
        If v_OrgKind2 <> "A" AndAlso v_OrgKind2 <> "" Then
            SearchOrgKind = v_OrgKind2 'OrgKind2.SelectedValue
            SchOrgKind2 = v_OrgKind2 'OrgKind2.SelectedValue
        End If

        Select Case v_SearchMode'SearchMode.SelectedValue
            Case "1"
                '依照訓練期間查詢
                MyValue = ""
                MyValue &= "&STDate1=" & STDate1.Text
                MyValue &= "&STDate2=" & STDate2.Text
                MyValue &= "&FTDate1=" & FTDate1.Text
                MyValue &= "&FTDate2=" & FTDate2.Text
                MyValue &= "&FDDate1=" & FTDate1.Text
                MyValue &= "&FDDate2=" & FTDate2.Text
                If RIDValue.Value = "A" Then RIDValue.Value = ""
                MyValue &= "&RIDValue=" & RIDValue.Value

            Case "2"
                'SearchMode.SelectedValue:2 依照年度月份查詢
                MyValue = ""
                MyValue &= "&Years=" & Years.SelectedValue
                MyValue &= "&Months=" & Months.SelectedValue
                If RIDValue.Value.ToString = "A" Then RIDValue.Value = ""
                MyValue &= "&RIDValue=" & RIDValue.Value

        End Select

        '54:充電起飛計畫（在職）判斷方式
        Dim sPackType As String = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            SearchOrgKind = "" '(產投計畫分類)
            SchOrgKind2 = "" '(產投計畫分類)
            If v_PackageType <> "A" AndAlso v_PackageType <> "" Then
                sPackType = v_PackageType 'PackageType.SelectedValue '包班種類
            End If
        End If

        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID '計畫代碼 (產投或充飛)
        MyValue &= "&OrgKind=" & SearchOrgKind '"",G,W'(產投計畫分類)(單選)
        MyValue &= "&OrgKind2=" & SchOrgKind2 '"",G,W'(產投計畫分類)(複選)
        MyValue &= "&PackageType=" & sPackType '"",2,3'包班種類

        If fileName1 = "" Then
            Common.MessageBox(Me, "找不到對應的報表!!")
            Exit Sub
        End If
        '列印報表。
        ReportQuery.PrintReport(Me, fileName1, MyValue)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call create1()

            Button1.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"
            btnPrint1.Attributes("onclick") = "return CheckPrint();"

            SearchMode.Attributes("onclick") = "ChangeMode();"
            PrintMode.Attributes("onclick") = "ChangePrintMode();"
        End If
        'DistID_TR.Style("display") = "none"
        'Org_TR.Style("display") = "none"
        'If PrintMode.SelectedIndex = 4 Then  '訓練單位
        '    DistID_TR.Style("display") = "inline"
        '    Org_TR.Style("display") = "inline"
        'End If
    End Sub

    Sub create1()
        'Page.RegisterStartupScript("SD_15_001_load1", "<script>ChangeMode();</script>")
        TIMS.RegisterStartupScript(Me, "", "<script>ChangeMode();</script>")
        'For i As Integer = 7 To 10
        '    If Not PrintMode.Items.FindByValue(CStr(i)) Is Nothing Then
        '        PrintMode.Items.Remove(PrintMode.Items.FindByValue(CStr(i)))
        '    End If
        'Next
        OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
        Common.SetListItem(OrgKind2, "G")

        Years = TIMS.GetSyear(Years)
        Months.Items.Add(New ListItem("請選擇", ""))
        For i As Integer = 1 To 12
            Months.Items.Add(New ListItem(i & "月份", i))
        Next
        DistID = TIMS.Get_DistID(DistID)

        trPlanKind.Style("display") = "none"
        trPackageType.Style("display") = "none"
        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trPackageType.Style("display") = cst_inline1 '"inline"
        Else
            '28:產業人才投資方案
            '計畫範圍 產投
            If sm.UserInfo.Years >= 2008 Then
                trPlanKind.Style("display") = cst_inline1 '"inline"
            End If
        End If
    End Sub

    Private Sub PrintMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PrintMode.SelectedIndexChanged
        Dim v_PrintMode As String = TIMS.GetListValue(PrintMode)

        DistID.SelectedIndex = -1
        center.Text = ""
        RIDValue.Value = ""
        PlanID.Value = ""

        DistID_TR.Style("display") = "none"
        Org_TR.Style("display") = "none"
        If v_PrintMode = "5" Then '訓練單位
            DistID_TR.Style("display") = cst_inline1 '"inline"
            Org_TR.Style("display") = cst_inline1 '"inline"
        End If
    End Sub

    Private Sub btnPrint1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint1.Click
        Call print1()
    End Sub
End Class
