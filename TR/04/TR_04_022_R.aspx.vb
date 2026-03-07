Partial Class TR_04_022_R
    Inherits AuthBasePage

    'ReportQuery
    'TR_04_022_R @TR

    'Dim objconn As SqlConnection

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        'objconn = DbAccess.GetConnection()
        '檢查Session是否存在 End

        If Not IsPostBack Then
            '年度
            Syear = TIMS.GetSyear(Syear)
            Syear.Enabled = True
            Common.SetListItem(Syear, sm.UserInfo.Years)
            If sm.UserInfo.DistID <> "000" Then
                Syear.Enabled = False
            End If

            '選擇全部轄區
            DistID = TIMS.Get_DistID(DistID)
            DistID.Items.Remove(DistID.Items.FindByValue(""))
            DistID.Items.Insert(0, New ListItem("全部", ""))
            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

            '訓練計畫
            TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

            btnPrint1.Attributes("onclick") = "return chkSearch();"
        End If

        DistID.Enabled = False
        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            Common.SetListItem(DistID, sm.UserInfo.DistID)
        End If
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Dim DistID1 As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 &= Convert.ToString(",")
                DistID1 &= Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
            End If
        Next
        Dim TPlanID1 As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 &= Convert.ToString(",")
                TPlanID1 &= Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
            End If
        Next
        If Me.Syear.SelectedValue = "" Then
            Errmsg += "請選擇年度" & vbCrLf
        End If
        If DistID1 = "" Then
            Errmsg += "請選擇轄區" & vbCrLf
        End If
        If TPlanID1 = "" Then
            Errmsg += "請選擇訓練計畫" & vbCrLf
        End If

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

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '列印
    Private Sub btnPrint1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim etitle As String = ""
        If FTDate1.Text <> "" OrElse FTDate2.Text <> "" Then
            etitle = FTDate1.Text & " ~ " & FTDate2.Text
        End If

        '報表要用的轄區參數
        Dim DistID1 As String = ""
        'Dim DistName As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 &= Convert.ToString(",")
                DistID1 &= Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
                'If DistName <> "" Then DistName &= Convert.ToString(",")
                'DistName &= Convert.ToString(Me.DistID.Items(i).Text)
            End If
        Next

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 &= Convert.ToString(",")
                TPlanID1 &= Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
                'If TPlanName <> "" Then TPlanName &= Convert.ToString(",")
                'TPlanName &= Convert.ToString(Me.TPlanID.Items(i).Text)
            End If
        Next

        Dim myValue As String = ""
        myValue = "prg=TR_04_022_R"
        myValue += "&Years=" & Syear.SelectedValue 'sm.UserInfo.Years
        myValue += "&DistID1=" & DistID1
        'myValue += "&DistName=" & Server.UrlEncode(DistName)
        myValue += "&TPlanID=" & TPlanID1
        'myValue += "&PlanName=" & Server.UrlEncode(TPlanName)
        myValue += "&FTDate1=" & Me.FTDate1.Text
        myValue += "&FTDate2=" & Me.FTDate2.Text
        myValue += "&etitle=" & etitle
        'ReportQuery 'TR_04_022_R @TR
        ReportQuery.PrintReport(Me, "TR_04_022_R", myValue)
    End Sub
End Class
