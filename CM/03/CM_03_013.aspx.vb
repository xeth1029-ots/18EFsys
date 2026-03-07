Partial Class CM_03_013
    Inherits AuthBasePage

    'ReportQuery
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
            CreateItem()
        End If

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        ''選擇全部訓練計畫
        'TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        '列印檢查
        Print.Attributes("onclick") = "javascript:return CheckPrint();"
    End Sub

    Sub CreateItem()
        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year)

        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))

        ''計畫
        'TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

        ''Common.SetListItem(DistID, sm.UserInfo.DistID)
        'For i As Integer = 0 To DistID.Items.Count - 1
        '    DistID.Items(i).Selected = True
        'Next
        'DistID.Enabled = False
        'Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        'DistName = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
                'If DistName <> "" Then DistName += ","
                'DistName += Convert.ToString(Me.DistID.Items(i).Text)
            End If
        Next

        ''報表要用的訓練計畫參數
        'TPlanID1 = ""
        'TPlanName = ""
        'For i = 1 To Me.TPlanID.Items.Count - 1
        '    If Me.TPlanID.Items(i).Selected Then
        '        If TPlanID1 = "" Then
        '            TPlanID1 = Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
        '            TPlanName = Convert.ToString(Me.TPlanID.Items(i).Text)
        '        Else
        '            TPlanID1 += "," & Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
        '            TPlanName += "," & Convert.ToString(Me.TPlanID.Items(i).Text)
        '        End If
        '    End If
        'Next

        Dim myValue As String = ""
        myValue = "prg=CM_03_013"
        myValue += "&Years=" & Syear.SelectedValue
        myValue += "&SFTDate=" & Me.FTDate1.Text
        myValue += "&FFTDate=" & Me.FTDate2.Text
        myValue += "&DistID=" & DistID1
        'myValue += "&DistName=" & DistName
        'myValue += "&TPlanID=" & TPlanID1
        'myValue += "&TPlanName=" & TPlanName

        'ReportQuery
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report2011", "CM_03_013", myValue)
    End Sub

End Class
