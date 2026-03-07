Partial Class TR_05_010_R
    Inherits AuthBasePage

    '報表格式: 2009:舊格式(2009年度含之前) / 2010:新格式(2010年度含之後)
    'TR_05_010_R3
    Const cst_printFN1 As String = "TR_05_010_R" '2009
    Const cst_printFN2 As String = "TR_05_010_R3" '2010

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If Not IsPostBack Then
            Call CreateItem()
        End If
    End Sub

    Sub CreateItem()
        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        '報表格式 PrintStyle
        Syear = TIMS.GetSyear(Syear) '年度
        Common.SetListItem(Syear, sm.UserInfo.Years)

        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

        Button1.Attributes("onclick") = "return search();"
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim msg As String = ""
        If Me.STDate1.Text = "" And Me.STDate2.Text = "" Then
            If Me.FTDate1.Text = "" And Me.FTDate2.Text = "" Then
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
        Dim DistName As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += "\'" & Me.DistID.Items(i).Value & "\'"

                If DistName <> "" Then DistName += ","
                DistName += Me.DistID.Items(i).Text
            End If
        Next

        '沒有選的動作
        If DistID1 = "" Then
            Common.SetListItem(DistID, sm.UserInfo.DistID)
            For i As Integer = 1 To Me.DistID.Items.Count - 1
                If Me.DistID.Items(i).Selected Then
                    If DistID1 <> "" Then DistID1 += ","
                    DistID1 += "\'" & Me.DistID.Items(i).Value & "\'"

                    If DistName <> "" Then DistName += ","
                    DistName += Me.DistID.Items(i).Text
                End If
            Next
        End If

        '報表要用的訓練計畫參數
        Dim j As Integer = 0
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += "\'" & Me.TPlanID.Items(i).Value & "\'"

                If TPlanName <> "" Then TPlanName += ","
                TPlanName += Me.TPlanID.Items(i).Text
                j = j + 1
            End If
        Next
        If TPlanID1 <> "" Then
            If j = (Me.TPlanID.Items.Count - 1) Then
                TPlanName = "全部"
            End If
        End If

        '沒有選的動作
        If TPlanID1 = "" Then
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
            For i As Integer = 1 To Me.TPlanID.Items.Count - 1
                If Me.TPlanID.Items(i).Selected Then
                    If TPlanID1 <> "" Then TPlanID1 += ","
                    TPlanID1 += "\'" & Me.TPlanID.Items(i).Value & "\'"

                    If TPlanName <> "" Then TPlanName += ","
                    TPlanName += Me.TPlanID.Items(i).Text
                End If
            Next
        End If

        '報表格式: 2009:舊格式(2009年度含之前) / 2010:新格式(2010年度含之後)
        Select Case PrintStyle.SelectedValue
            Case "2009"
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_05_010_R", "Years=" & Syear.SelectedValue & "&SSTDate=" & Me.STDate1.Text & "&ESTDate=" & Me.STDate2.Text & "&SFTDate=" & Me.FTDate1.Text & "&EFTDate=" & Me.FTDate2.Text & "&DistID=" & newDistID & "&DistName=" & Server.UrlEncode(newDistName) & "&TPlanID=" & newTPlanID & "&PlanName=" & Server.UrlEncode(newTPlanIDName))
                Dim MyValue As String = ""
                MyValue = "jkl=" & Convert.ToString(Request("ID"))
                MyValue += "&Years=" & Syear.SelectedValue
                MyValue += "&SSTDate=" & Me.STDate1.Text
                MyValue += "&ESTDate=" & Me.STDate2.Text
                MyValue += "&SFTDate=" & Me.FTDate1.Text
                MyValue += "&EFTDate=" & Me.FTDate2.Text
                MyValue += "&DistID=" & DistID1
                MyValue += "&TPlanID=" & TPlanID1
                MyValue += "&DistName=" & Server.UrlEncode(DistName)
                MyValue += "&PlanName=" & Server.UrlEncode(TPlanName)
                ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

            Case "2010"
                Dim MyValue As String = ""
                MyValue = "jkl=" & Convert.ToString(Request("ID"))
                MyValue += "&Years=" & Syear.SelectedValue
                MyValue += "&STDate1=" & Me.STDate1.Text
                MyValue += "&STDate2=" & Me.STDate2.Text
                MyValue += "&FTDate1=" & Me.FTDate1.Text
                MyValue += "&FTDate2=" & Me.FTDate2.Text
                MyValue += "&DistID=" & DistID1
                MyValue += "&TPlanID=" & TPlanID1
                MyValue += "&DistName=" & Server.UrlEncode(DistName)
                MyValue += "&PlanName=" & Server.UrlEncode(TPlanName)
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_05_010_R2", "Years=" & Syear.SelectedValue & "&SSTDate=" & Me.STDate1.Text & "&ESTDate=" & Me.STDate2.Text & "&SFTDate=" & Me.FTDate1.Text & "&EFTDate=" & Me.FTDate2.Text & "&DistID=" & newDistID & "&DistName=" & Server.UrlEncode(newDistName) & "&TPlanID=" & newTPlanID & "&PlanName=" & Server.UrlEncode(newTPlanIDName))
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_05_010_R2", MyValue)
                ReportQuery.PrintReport(Me, cst_printFN2, MyValue)

        End Select

    End Sub
End Class
