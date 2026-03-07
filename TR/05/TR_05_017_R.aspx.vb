Partial Class TR_05_017_R
    Inherits AuthBasePage

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
            Call CreateItem()

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

            '選擇訓練計畫
            chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
            chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
            chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"
            ''選擇全部訓練計畫
            'TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
            '選擇全部預算來源
            BudgetList.Attributes("onclick") = "SelectAll('BudgetList','hidBudgetList');"
            '列印檢查
            Print.Attributes("onclick") = "javascript:return CheckPrint();"
        End If

    End Sub

    Sub CreateItem()
        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, sm.UserInfo.Years) '預設值

        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))

        '計畫
        Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn, 1)
        'TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

        '預算來源
        BudgetList = TIMS.Get_Budget(BudgetList, 3)
        BudgetList.Items.Insert(0, New ListItem("全部", ""))

    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click

        '報表要用的轄區參數
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數
        'Dim TPlanID1 As String = ""
        'TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 3)
        'newTPlanID = TPlanID1

        '報表要用的預算來源參數
        Dim BudgetID As String = ""
        BudgetID = TIMS.GetCheckBoxListRptVal(BudgetList, 1)

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&Years=" & Syear.SelectedValue '年度
        MyValue += "&STDate1=" & Me.STDate1.Text
        MyValue += "&STDate2=" & Me.STDate2.Text
        MyValue += "&FTDate1=" & Me.FTDate1.Text
        MyValue += "&FTDate2=" & Me.FTDate2.Text

        MyValue += "&DistID=" & DistID1
        MyValue += "&TPlanID=" & TPlanID1
        MyValue += "&BudgetID=" & BudgetID '預算來源

        Dim sFileName As String = ""
        'sFileName = "TR_05_017_1" '依分署(中心)
        sFileName = "TR_05_017_2" '依計畫
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", sFileName, MyValue)
    End Sub
End Class
