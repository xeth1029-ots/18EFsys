Partial Class SD_15_004
    Inherits AuthBasePage

    'SD_15_004*.jrxml (1~4) "SD_15_004_
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            SearchPlan = TIMS.Get_RblSearchPlan(Me, SearchPlan)
            Common.SetListItem(SearchPlan, "G")

            DistID = TIMS.Get_DistID(DistID)

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
        End If

        Button1.Attributes("onclick") = "OpenOrg('" & sm.UserInfo.TPlanID & "');"
        btnPrint1.Attributes("onclick") = "return CheckPrint();"
    End Sub

    Private Sub btnPrint1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint1.Click
        Dim MyValue As String = ""
        Dim fileName1 As String = ""
        Dim SearchPlan1 As String = ""

        If SearchPlan.SelectedValue <> "A" Then
            SearchPlan1 = SearchPlan.SelectedValue
        End If

        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            SearchPlan1 = ""
        End If
        Dim sPackType As String = ""
        If PackageType.SelectedValue <> "A" Then
            sPackType = PackageType.SelectedValue
        End If

        MyValue = ""
        MyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        MyValue &= "&FTDate1=" & FTDate1.Text
        MyValue &= "&FTDate2=" & FTDate2.Text
        MyValue &= "&DistID=" & DistID.SelectedValue
        MyValue &= "&PlanID=" & PlanID.Value
        MyValue &= "&RID=" & RIDValue.Value
        MyValue &= "&SearchPlan=" & SearchPlan1 '"",G,W
        MyValue &= "&PackageType=" & sPackType '"",2,3

        Select Case PrintMode.SelectedValue
            Case "1"
                fileName1 = "SD_15_004_1"
            Case "2"
                MyValue &= "&FTDate1_2=" & FTDate1.Text
                MyValue &= "&FTDate2_2=" & FTDate2.Text
                MyValue &= "&DistID2=" & DistID.SelectedValue
                MyValue &= "&PlanID2=" & PlanID.Value
                MyValue &= "&RID2=" & RIDValue.Value
                fileName1 = "SD_15_004_2"
            Case "3"
                fileName1 = "SD_15_004_3"
            Case "4"
                fileName1 = "SD_15_004_4"
        End Select
        ReportQuery.PrintReport(Me, "BussinessTrain", fileName1, MyValue)

    End Sub
End Class
