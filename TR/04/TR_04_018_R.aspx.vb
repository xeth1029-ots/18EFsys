Partial Class TR_04_018
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            DistID = TIMS.Get_DistID(DistID)
            DistID.Items.Insert(0, New ListItem("全部", ""))
        End If

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        DistID.Enabled = False
        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            Common.SetListItem(DistID, sm.UserInfo.DistID)
        End If

        Print.Attributes("onclick") = "return rptsearch();"

    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '報表要用的轄區參數 '選擇轄區
        Dim DistID1 As String
        Dim DistName As String
        DistID1 = ""
        DistName = ""
        For Each objitem As ListItem In Me.DistID.Items '選擇轄區
            If objitem.Selected = True Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += "\'" & objitem.Value & "\'"
                If DistName <> "" Then DistName += ","
                DistName += objitem.Text
            End If
        Next

        If DistID1 <> "" Then
            If Me.DistID.Items(0).Selected Then
                DistName = "全部"
            End If
        End If

        Dim MyValue As String = ""
        MyValue = "actid=print"
        MyValue += "&DistID=" & DistID1
        MyValue += "&DistName=" & DistName
        MyValue += "&STDate1=" & STDate1.Text
        MyValue += "&STDate2=" & STDate2.Text
        MyValue += "&FTDate1=" & FTDate1.Text
        MyValue += "&FTDate2=" & FTDate2.Text

        Select Case PrintStyle.SelectedValue
            Case "1"
                '依性別、年齡、教育程度 
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_018_R", MyValue)
                ReportQuery.PrintReport(Me, "Report2011", "TR_04_018_R_2011", MyValue)
            Case "2"
                '依身分別
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_018_R_1", MyValue)
                ReportQuery.PrintReport(Me, "Report2011", "TR_04_018_R1_2011", MyValue)
            Case "3"
                '依訓練職類
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_018_R_2", MyValue)
                ReportQuery.PrintReport(Me, "Report2011", "TR_04_018_R2_2011", MyValue)
        End Select

    End Sub
End Class
