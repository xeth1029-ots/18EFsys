Partial Class TR_04_017_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            DistID = TIMS.Get_DistID(DistID)
            DistID.Items.Insert(0, New ListItem("全部", ""))
        End If

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If

        Print.Attributes("onclick") = "return search();"
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '選擇轄區
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        Dim DistName As String = ""
        Dim MyValue As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += "" & Me.DistID.Items(i).Value & ""
                If DistName <> "" Then DistName += ","
                DistName += Convert.ToString(Me.DistID.Items(i).Text)
            End If
        Next

        Select Case PrintStyle.SelectedValue
            Case "1"
                '依性別、年齡、教育程度 
                If DistID1 <> "" Then
                    MyValue = "DistID=" & DistID1
                End If
                MyValue += "&DistName=" & Convert.ToString(DistName) & _
                                                 "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & _
                                                 "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text
                ReportQuery.PrintReport(Me, "TR", "TR_04_017_R", MyValue)
            Case "2"
                '依身分別
                If DistID1 <> "" Then
                    MyValue = "DistID=" & DistID1
                End If
                MyValue += "&DistName=" & Convert.ToString(DistName) & _
                                                 "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & _
                                                 "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text
                ReportQuery.PrintReport(Me, "TR", "TR_04_017_R_1", MyValue)
            Case "3"
                '依訓練職類
                If DistID1 <> "" Then
                    MyValue = "DistID=" & DistID1
                End If
                MyValue += "&DistName=" & Convert.ToString(DistName) & _
                                                 "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & _
                                                 "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text
                ReportQuery.PrintReport(Me, "TR", "TR_04_017_R_2", MyValue)
        End Select

    End Sub

End Class