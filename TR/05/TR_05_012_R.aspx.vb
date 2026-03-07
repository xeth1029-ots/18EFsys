Partial Class TR_05_012_R
    Inherits AuthBasePage

    'TR_05_012_R
    Const cst_printFN1 As String = "TR_05_012_R"

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
        End If

        ''選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        '列印鈕
        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(start_date.Text) <> "" Then
            start_date.Text = Trim(start_date.Text)
            If Not TIMS.IsDate1(start_date.Text) Then
                Errmsg += "開訓期間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            start_date.Text = ""
            Errmsg += "開訓期間 起始日期為必填資料"
        End If

        If Trim(end_date.Text) <> "" Then
            end_date.Text = Trim(end_date.Text)
            If Not TIMS.IsDate1(end_date.Text) Then
                Errmsg += "開訓期間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If
        Else
            end_date.Text = ""
            Errmsg += "開訓期間 迄止日期為必填資料"
        End If

        If Errmsg = "" Then
            If Me.start_date.Text <> "" AndAlso Me.end_date.Text <> "" Then
                If DateDiff(DateInterval.Day, CDate(start_date.Text), CDate(end_date.Text)) < 0 Then
                    Errmsg += "開訓期間 日期起迄，迄日需大起日" & vbCrLf
                End If
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        'Dim j As Integer = 0 '數量平衡計算
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= "\'" & Me.TPlanID.Items(i).Value & "\'"
                'j = j + 1
            End If
        Next

        'If TPlanID1 <> "" Then
        '    ''If Me.TPlanID.Items(0).Selected Then
        '    ''    TPlanName = "全部"
        '    ''End If
        '    If j = (Me.TPlanID.Items.Count - 1) Then
        '        TPlanName = "全部"
        '    End If
        'End If

        Dim MyValue As String = ""
        MyValue = "prgid=" & Request("ID")
        MyValue += "&start_date=" & start_date.Text
        MyValue += "&end_date=" & end_date.Text
        MyValue += "&TPlanID=" & TPlanID1
        'MyValue += "&PlanName=" & TPlanName
        'MyValue += "&PlanName=" & Server.UrlEncode(TPlanName)
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub
End Class
