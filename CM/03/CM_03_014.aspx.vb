Partial Class CM_03_014
    Inherits AuthBasePage

    Dim vsDistID1 As String = ""
    Dim vsTPlanID1 As String = ""
    Dim vsBudgetID As String = ""
    Dim vsDistName As String = ""
    Dim vsTPlanName As String = ""
    Dim vsBudgetName As String = ""
    'Dim objconn As SqlConnection

    'Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call TIMS.CloseDbConn(objconn)
    'End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        If Not IsPostBack Then
            '建立初始資料
            Call CreateItem()
            '取得session 
            Call GetSessionSearch()
        End If
    End Sub

    '建立初始資料
    Sub CreateItem()

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

        '列印檢查
        Print.Attributes("onclick") = "javascript:return CheckPrint();"
        '查詢檢查
        btnSearch1.Attributes("onclick") = "javascript:return CheckPrint();"

        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year)

        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))

        '計畫
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

        '預算來源
        BudgetList = TIMS.Get_Budget(BudgetList, 3)

    End Sub


    '取得session 
    Sub GetSessionSearch()
        Const Cst_MySearch As String = "_MySearch"
        Dim MyValue As String = ""

        If Not Session(Cst_MySearch) Is Nothing Then
            MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "prgid")
            If MyValue = "CM_03_014" Then
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "Years")
                If MyValue <> "" Then Common.SetListItem(Me.Syear, MyValue)
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "STDate1")
                If MyValue <> "" Then STDate1.Text = MyValue
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "STDate2")
                If MyValue <> "" Then STDate2.Text = MyValue
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "FTDate1")
                If MyValue <> "" Then FTDate1.Text = MyValue
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "FTDate2")
                If MyValue <> "" Then FTDate2.Text = MyValue

                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "DistID1")
                If MyValue <> "" Then TIMS.SetCblValue(DistID, MyValue)
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "TPlanID1")
                If MyValue <> "" Then TIMS.SetCblValue(TPlanID, MyValue)
                MyValue = TIMS.GetMyValue(Session(Cst_MySearch), "BudgetID")
                If MyValue <> "" Then TIMS.SetCblValue(BudgetList, MyValue)
            End If
        End If

    End Sub


    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)

        If STDate1.Text <> "" Then
            'STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If STDate2.Text <> "" Then
            'STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If FTDate1.Text <> "" Then
            'FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If FTDate2.Text <> "" Then
            'FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        Dim v_Syear As String = TIMS.GetListValue(Syear)
        If v_Syear = "" _
            AndAlso STDate1.Text = "" AndAlso STDate2.Text = "" _
            AndAlso FTDate1.Text = "" AndAlso FTDate2.Text = "" Then
            Errmsg += "[年度]、[開訓區間]、[結訓區間],請擇一輸入查詢" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '保留查詢Session
    Function KeepSearch() As String
        Const Cst_MySearch As String = "_MySearch"
        Session(Cst_MySearch) = Nothing

        Dim MySearch As String = ""
        MySearch = "prgid=" & "CM_03_014"
        MySearch += "&Years=" & Me.Syear.SelectedValue
        MySearch += "&STDate1=" & Me.STDate1.Text
        MySearch += "&STDate2=" & Me.STDate2.Text
        MySearch += "&FTDate1=" & Me.FTDate1.Text
        MySearch += "&FTDate2=" & Me.FTDate2.Text
        MySearch += "&DistID1=" & vsDistID1
        MySearch += "&TPlanID1=" & vsTPlanID1
        MySearch += "&BudgetID=" & vsBudgetID

        Session(Cst_MySearch) = MySearch
        Return MySearch
    End Function

    '供smartQuery查詢使用
    Function KeepSearch2() As String
        Dim MySearch As String = ""

        MySearch = "prgid=" & "CM_03_014"
        MySearch += "&Years=" & Me.Syear.SelectedValue
        MySearch += "&STDate1=" & Me.STDate1.Text
        MySearch += "&STDate2=" & Me.STDate2.Text
        MySearch += "&FTDate1=" & Me.FTDate1.Text
        MySearch += "&FTDate2=" & Me.FTDate2.Text

        '請先執行 GetSearchValue2() 取得vs參數
        MySearch += "&DistID=" & vsDistID1
        MySearch += "&TPlanID=" & vsTPlanID1
        MySearch += "&BudgetID=" & vsBudgetID
        MySearch += "&BudgetID2=" & vsBudgetID
        MySearch += "&DistName=" & vsDistName
        MySearch += "&TPlanName=" & vsTPlanName
        MySearch += "&BudgetName=" & vsBudgetName

        Return MySearch
    End Function

    '組合查詢條件
    Sub GetSearchValue1()
        Dim Rst As String = ""

        'Dim DistID1 As String = ""
        'Dim TPlanID1 As String = ""
        'Dim BudgetID As String = ""

        'SQL要用的轄區參數
        Dim DistID1 As String = ""
        'DistID1 = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected AndAlso Me.DistID.Items(i).Value <> "" Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= "'" & Me.DistID.Items(i).Value & "'"
            End If
        Next

        'SQL要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'TPlanID1 = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= "'" & Me.TPlanID.Items(i).Value & "'"
            End If
        Next

        'SQL要用的預算來源參數
        Dim BudgetID As String = ""
        'BudgetID = ""
        For i As Integer = 0 To Me.BudgetList.Items.Count - 1
            If Me.BudgetList.Items(i).Selected AndAlso Me.BudgetList.Items(i).Value <> "" Then
                If BudgetID <> "" Then BudgetID &= ","
                BudgetID &= "'" & Me.BudgetList.Items(i).Value & "'"
            End If
        Next

        vsDistID1 = ""
        If DistID1 <> "" Then vsDistID1 = DistID1
        vsTPlanID1 = ""
        If TPlanID1 <> "" Then vsTPlanID1 = TPlanID1
        vsBudgetID = ""
        If BudgetID <> "" Then vsBudgetID = BudgetID

    End Sub

    '取得smartQuery 要傳入的參數
    Sub GetSearchValue2()
        Dim Rst As String = ""

        Dim DistID1 As String = ""
        Dim DistName As String = ""
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        Dim BudgetID As String = ""
        Dim BudgetName As String = ""

        '報表要用的轄區參數
        DistID1 = ""
        DistName = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected AndAlso Me.DistID.Items(i).Value <> "" Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= "\'" & Me.DistID.Items(i).Value & "\'"

                If DistName <> "" Then DistName &= ","
                DistName &= Me.DistID.Items(i).Text
            End If
        Next

        '報表要用的訓練計畫參數
        TPlanID1 = ""
        TPlanName = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= "\'" & Me.TPlanID.Items(i).Value & "\'"

                If TPlanName <> "" Then TPlanName &= ","
                TPlanName &= Me.TPlanID.Items(i).Text
            End If
        Next

        '報表要用的預算來源參數
        BudgetID = ""
        BudgetName = ""
        For i As Integer = 0 To Me.BudgetList.Items.Count - 1
            If Me.BudgetList.Items(i).Selected AndAlso Me.BudgetList.Items(i).Value <> "" Then
                If BudgetID <> "" Then BudgetID &= ","
                BudgetID &= "\'" & Me.BudgetList.Items(i).Value & "\'"

                If BudgetName <> "" Then BudgetName &= ","
                BudgetName &= Me.BudgetList.Items(i).Text
            End If
        Next

        vsDistID1 = ""
        If DistID1 <> "" Then vsDistID1 = DistID1
        vsTPlanID1 = ""
        If TPlanID1 <> "" Then vsTPlanID1 = TPlanID1
        vsBudgetID = ""
        If BudgetID <> "" Then vsBudgetID = BudgetID

        vsDistName = ""
        If DistName <> "" Then vsDistName = DistName
        vsTPlanName = ""
        If TPlanName <> "" Then vsTPlanName = TPlanName
        vsBudgetName = ""
        If BudgetName <> "" Then vsBudgetName = BudgetName

        ''報表要用的標題轄區參數
        'TDistName = ""
    End Sub

    '查詢鈕
    Private Sub btnSearch1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call GetSearchValue1()
        Call KeepSearch()
        TIMS.Utl_Redirect1(Me, "CM_03_014_A.aspx?ID=" & Request("ID"))
    End Sub

    '列印鈕
    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'Common.MessageBox(Me, "功能開發中，敬請期待，造成您的不便，請見諒")
        'Exit Sub

        Call GetSearchValue2()
        Dim MyValue As String = ""

        '取得smartQuery 要傳入的參數
        Call GetSearchValue2()
        'smartQuery查詢使用
        MyValue = KeepSearch2()
        ReportQuery.PrintReport(Me, "Report2011", "CM_03_014", MyValue)
    End Sub

End Class
