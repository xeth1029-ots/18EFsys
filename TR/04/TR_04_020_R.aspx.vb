Partial Class TR_04_020_R
    Inherits AuthBasePage

#Region "NO USE"
    ''測試function
    'Sub TestSub()
    '    '------------- ------ ---- ---test'------------- ------ ---- ---
    '    Me.Syear.SelectedValue = "2010"
    '    Me.DistID.SelectedValue = "001"
    '    'TPlanID
    '    FTDate1.Text = "2010/07/01"
    '    FTDate2.Text = "2010/11/01"
    '    Dim j As Integer = 0
    '    Dim tmpStr As String = ""
    '    Dim CBLobj As CheckBoxList
    '    j = 0
    '    tmpStr = ""
    '    CBLobj = DistID
    '    For i As Integer = 1 To CBLobj.Items.Count - 1
    '        Dim objitem As ListItem = CBLobj.Items(i)
    '        If "001,002,003".IndexOf(objitem.Value) > -1 Then
    '            objitem.Selected = True
    '        End If
    '    Next
    '    j = 0
    '    tmpStr = ""
    '    CBLobj = TPlanID
    '    For i As Integer = 1 To CBLobj.Items.Count - 1
    '        Dim objitem As ListItem = CBLobj.Items(i)
    '        If "01,02,03,04,05".IndexOf(objitem.Value) > -1 Then
    '            objitem.Selected = True
    '        End If
    '    Next
    '    'Dim itemPlan As String = tmpStr
    '    'If j = Me.TPlanID.Items.Count - 1 Then itemPlan = ""
    '    j = 0
    '    tmpStr = ""
    '    CBLobj = BudgetList
    '    For i As Integer = 1 To CBLobj.Items.Count - 1
    '        Dim objitem As ListItem = CBLobj.Items(i)
    '        If "01,02,03".IndexOf(objitem.Value) > -1 Then
    '            objitem.Selected = True
    '        End If
    '    Next
    '    'Dim itemBudget As String = tmpStr
    '    '------------- ------ ---- ---test'------------- ------ ---- ---
    'End Sub

#End Region

    Const cst_printFN1 As String = "TR_04_020_R"

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

        ''分頁設定 Start
        'PageControler1.PageDataGrid = DataGrid1
        ''分頁設定 End

        If Not IsPostBack Then
            'msg.Text = ""
            'PageControler1.Visible = False
            'DataGrid1.Visible = False

            Call CreateItem()

            ''選擇全部轄區
            'DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

            Syear.Enabled = True
            DistID.Enabled = True
            Common.SetListItem(Syear, sm.UserInfo.Years)
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            If sm.UserInfo.DistID <> "000" Then
                Syear.Enabled = False
                DistID.Enabled = False
            End If

            btnPrint.Attributes("onclick") = "return chkSearch();"
            'btnSearch.Attributes("onclick") = "return chkSearch();"
            'btnExport1.Attributes("onclick") = "return chkSearch();"
        End If

        'If Not IsPostBack Then
        '    If TIMS.sUtl_ChkTest() Then '測試用
        '        Call TestSub()
        '    End If
        'End If
    End Sub

    '關鍵字詞建立
    Sub CreateItem()
        '年度
        Syear = TIMS.GetSyear(Syear)
        '轄區別
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Remove(DistID.Items.FindByValue(""))
        DistID.Items.Insert(0, New ListItem("全部", ""))

        '計畫別
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        '預算來源
        BudgetList = TIMS.Get_Budget(BudgetList, 3)
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Me.Syear.SelectedValue = "" Then
            Errmsg += "請選擇年度" & vbCrLf
        End If

        'If Me.DistID.SelectedValue = "" Then
        '    Errmsg += "請選擇轄區中心" & vbCrLf
        'End If

        Dim j As Integer = 0
        Dim CBLobj As CheckBoxList
        'j = 0
        'CBLobj = DistID
        'For i As Integer = 1 To CBLobj.Items.Count - 1
        '    Dim objitem As ListItem = CBLobj.Items(i)
        '    If objitem.Selected = True Then
        '        j += 1
        '        Exit For
        '    End If
        'Next
        'If j = 0 Then Errmsg += "請選擇轄區中心" & vbCrLf


        If Trim(FTDate1.Text) <> "" Then
            FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓期間 的起始日不是正確的日期格式" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate1.Text = ""
        End If

        If Trim(FTDate2.Text) <> "" Then
            FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓期間 的迄止日不是正確的日期格式" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate2.Text = ""
        End If

        If Errmsg = "" Then
            If Me.FTDate1.Text <> "" AndAlso Me.FTDate2.Text <> "" Then
                If DateDiff(DateInterval.Day, CDate(FTDate1.Text), CDate(FTDate2.Text)) < 0 Then
                    Errmsg += "結訓期間 日期起迄有誤，迄日需大於起日" & vbCrLf
                End If
            End If
        End If

        j = 0
        CBLobj = TPlanID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇訓練計畫" & vbCrLf


        j = 0
        CBLobj = BudgetList
        For i As Integer = 0 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇預算來源" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '列印
    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim j As Integer = 0
        Dim tmpStr As String = ""
        Dim CBLobj As CheckBoxList

        'j = 0
        'tmpStr = ""
        'CBLobj = DistID
        'For i As Integer = 1 To CBLobj.Items.Count - 1
        '    Dim objitem As ListItem = CBLobj.Items(i)
        '    If objitem.Selected = True Then
        '        j += 1
        '        If tmpStr <> "" Then tmpStr += ","
        '        tmpStr += "\'" & objitem.Value & "\'"
        '    End If
        'Next
        'Dim itemDist As String = tmpStr

        Dim itemDist As String = ""
        If DistID.SelectedValue <> "" Then
            itemDist = "\'" & DistID.SelectedValue & "\'"
        Else
            tmpStr = ""
            itemDist = ""
            For Each objitem As ListItem In DistID.Items
                If objitem.Value <> "" Then
                    If tmpStr <> "" Then tmpStr += ","
                    tmpStr += "\'" & objitem.Value & "\'"
                End If
            Next
            itemDist = tmpStr
        End If

        j = 0
        tmpStr = ""
        CBLobj = TPlanID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "\'" & objitem.Value & "\'"
            End If
        Next
        Dim itemPlan As String = tmpStr

        j = 0
        tmpStr = ""
        CBLobj = BudgetList
        For i As Integer = 0 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "\'" & objitem.Value & "\'"
            End If
        Next
        Dim itemBudget As String = tmpStr

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&Years=" & Syear.SelectedValue
        'MyValue += "&DistID=" & DistID.SelectedValue
        'MyValue += "&DistID2=" & DistID.SelectedValue
        MyValue += "&DistID=" & itemDist
        MyValue += "&DistID2=" & itemDist

        MyValue += "&TPlanID=" & itemPlan
        MyValue += "&FTDate1=" & FTDate1.Text
        MyValue += "&FTDate2=" & FTDate2.Text
        MyValue += "&BudgetID=" & itemBudget

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

    End Sub

End Class