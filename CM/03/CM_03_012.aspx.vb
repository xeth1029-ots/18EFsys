Partial Class CM_03_012
    Inherits AuthBasePage

    Const cst_printFN1 As String = "CM_03_012"

    'CM_03_012
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        If Not IsPostBack Then
            Call CreateItem()

            DistID.Attributes("onclick") = "ClearData();"
            TPlanID.Attributes("onclick") = "ClearData();"
            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
            '列印檢查
            Print.Attributes("onclick") = "javascript:return CheckPrint();"
        End If

    End Sub

    Sub CreateItem()
        '年度
        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year)
        '轄區
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Insert(0, New ListItem("全部", ""))
        '計畫
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        'Dim dt As DataTable
        'Dim sqlstr As String
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        'Dim DistName As String = ""
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += "\'" & Me.DistID.Items(i).Value & "\'"
                'If DistName <> "" Then DistName += ","
                'DistName += Me.DistID.Items(i).Text
            End If
        Next
        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        'Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
                'If TPlanName <> "" Then TPlanName += ","
                'TPlanName += Convert.ToString(Me.TPlanID.Items(i).Text)
            End If
        Next

        Dim MyValue As String = ""
        MyValue = "jkl=jkl"
        MyValue += "&start_date=" & Me.STDate1.Text
        MyValue += "&end_date=" & Me.STDate2.Text
        MyValue += "&start_date2=" & Me.FTDate1.Text
        MyValue += "&end_date2=" & Me.FTDate2.Text
        MyValue += "&DistID=" & DistID1
        MyValue += "&TPlanID=" & TPlanID1
        MyValue += "&Years=" & Syear.SelectedValue
        'MyValue += "&DistName=" & Server.UrlEncode(DistName)
        'MyValue += "&TPlanName=" & Server.UrlEncode(TPlanName)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "CM_03_012", MyValue)
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

    End Sub

End Class
