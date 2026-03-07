Partial Class TR_05_011_R
    Inherits AuthBasePage

    'TR_05_011_R*.jrxml
    'TR_05_011_R (依訓練職類)
    'TR_05_011_R2 (依通俗職類)
    Const cst_printFN1 As String = "TR_05_011_R" '(依訓練職類)
    Const cst_printFN2 As String = "TR_05_011_R2" '(依通俗職類)

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'Call TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()
        End If

    End Sub

    Sub CreateItem()
        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部縣市
        CityList.Attributes("onclick") = "SelectAll('CityList','CityHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        '選擇全部預算
        BudgetList.Attributes("onclick") = "SelectAll('BudgetList','BudgetHidden');"

        Button1.Attributes("onclick") = "return search();"

        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        Syear = TIMS.GetSyear(Syear)
        'Common.SetListItem(Syear, Now.Year) '年度設定
        Common.SetListItem(Syear, sm.UserInfo.Years)  '年度設定

        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", "")) '新增全部
        Common.SetListItem(DistID, sm.UserInfo.DistID)

        '縣市別
        CityList = TIMS.Get_CityName(CityList, TIMS.dtNothing)

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")  '大計畫 "Y" 含全部

        BudgetList = TIMS.Get_Budget(BudgetList, 3) '含特別預算
        BudgetList.Items.Insert(0, New ListItem("全部", "")) '新增全部
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim msg As String = ""
        msg = ""
        If Me.STDate1.Text = "" AndAlso Me.STDate2.Text = "" Then
            If Me.FTDate1.Text = "" AndAlso Me.FTDate2.Text = "" Then
                If Syear.SelectedValue = "" Then
                    msg += "年度、開訓日期、結訓日期擇一為查詢條件!!" & vbCrLf
                End If
            End If
        End If
        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        '報表要用的轄區參數
        Dim DistID1 As String = ""
        'Dim DistName As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected AndAlso Me.DistID.Items(i).Value <> "" Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += Me.DistID.Items(i).Value

                'If DistName <> "" Then DistName += ","
                'DistName += Me.DistID.Items(i).Text
            End If
        Next

        '沒有選的動作
        'If DistID1 = "" Then
        '    Common.SetListItem(DistID, sm.UserInfo.DistID)
        '    For i As Integer = 1 To Me.DistID.Items.Count - 1
        '        If Me.DistID.Items(i).Selected Then
        '            If DistID1 <> "" Then DistID1 += ","
        '            DistID1 += Me.DistID.Items(i).Value

        '            If DistName <> "" Then DistName += ","
        '            DistName += Me.DistID.Items(i).Text
        '        End If
        '    Next
        'End If
        '選擇縣市別
        Dim itemcity As String = ""
        itemcity = ""
        For Each objitem As ListItem In Me.CityList.Items
            If objitem.Selected = True Then
                If itemcity <> "" Then itemcity += ","
                itemcity += objitem.Value
            End If
        Next

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Me.TPlanID.Items(i).Value
                'If TPlanName <> "" Then TPlanName += ","
                'TPlanName += Me.TPlanID.Items(i).Text
            End If
        Next
        'If TPlanID1 <> "" Then
        '    If TPlanID1.Split(",").Length = (Me.TPlanID.Items.Count - 1) Then
        '        TPlanName = "全部"
        '    End If
        'End If
        '沒有選的動作
        If TPlanID1 = "" Then
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            For i As Integer = 1 To Me.TPlanID.Items.Count - 1
                If Me.TPlanID.Items(i).Selected Then
                    If TPlanID1 <> "" Then TPlanID1 += ","
                    TPlanID1 += Me.TPlanID.Items(i).Value
                    'If TPlanName <> "" Then TPlanName += ","
                    'TPlanName += Me.TPlanID.Items(i).Text
                End If
            Next
        End If

        '報表要用的預算別參數
        Dim BudgetID1 As String = ""
        'Dim BudgetName As String = ""
        For i As Integer = 1 To Me.BudgetList.Items.Count - 1
            If Me.BudgetList.Items(i).Selected Then
                If BudgetID1 <> "" Then BudgetID1 += ","
                BudgetID1 += Me.BudgetList.Items(i).Value
                'If BudgetName <> "" Then BudgetName += ","
                'BudgetName += Me.BudgetList.Items(i).Text
            End If
        Next
        'If BudgetID1 <> "" Then
        '    If BudgetID1.Split(",").Length = (Me.BudgetList.Items.Count - 1) Then
        '        'BudgetID1 = ""
        '        BudgetName = "全部"
        '    End If
        'End If

        Dim MyValue As String = ""
        MyValue = "sType=Print"
        MyValue += "&Years=" & Syear.SelectedValue
        MyValue += "&SSTDate=" & Me.STDate1.Text
        MyValue += "&ESTDate=" & Me.STDate2.Text
        MyValue += "&SFTDate=" & Me.FTDate1.Text
        MyValue += "&EFTDate=" & Me.FTDate2.Text
        If DistID1 <> "" Then MyValue += "&DistID=" & DistID1
        If TPlanID1 <> "" Then MyValue += "&TPlanID=" & TPlanID1
        If BudgetID1 <> "" Then MyValue += "&BudgetID=" & BudgetID1
        If itemcity <> "" Then MyValue += "&CTID=" & itemcity
        'MyValue += "&DistName=" & Server.UrlEncode(DistName)
        'MyValue += "&PlanName=" & Server.UrlEncode(TPlanName)
        'MyValue += "&BudgetName=" & Server.UrlEncode(BudgetName)
        'Exit Sub

        Select Case Convert.ToString(PrintType.SelectedValue)
            Case "1" '依訓練職類
                ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
            Case "2" '依通俗職類
                ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
        End Select

    End Sub

End Class

