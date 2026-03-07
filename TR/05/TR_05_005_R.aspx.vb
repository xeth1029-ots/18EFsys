Partial Class TR_05_005_R
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        If Not IsPostBack Then
            CreateItem()
        End If
        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If
        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()

        Syear = TIMS.GetSyear(Syear) '年度
        Common.SetListItem(Syear, Now.Year)

        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '選擇轄區

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
                If DistID1 <> "" Then DistID1 = ","
                DistID1 &= Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
                If DistName <> "" Then DistName = ","
                DistName &= Convert.ToString("\'" & Me.DistID.Items(i).Text & "\'")
            End If
        Next

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 = ","
                TPlanID1 &= Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
                If TPlanName <> "" Then TPlanName = ","
                TPlanName &= Convert.ToString("\'" & Me.TPlanID.Items(i).Text & "\'")
            End If
        Next

        Dim Years As String = ""
        Years = Syear.SelectedValue
        If Syear.SelectedValue = "" Then
            Years = sm.UserInfo.Years
        End If

        Dim MyValue As String = ""
        MyValue = "jkl=jlk"
        MyValue += "&Years=" & Years
        MyValue += "&STDate1=" & STDate1.Text
        MyValue += "&STDate2=" & STDate2.Text
        MyValue += "&FTDate1=" & FTDate1.Text
        MyValue += "&FTDate2=" & FTDate2.Text
        MyValue += "&DistID=" & DistID1 'newDistID
        MyValue += "&DistName=" & DistName 'newDistName
        MyValue += "&TPlanID=" & TPlanID1 'newTPlanID
        MyValue += "&PlanName=" & TPlanName 'newTPlanIDName
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_05_005_R", MyValue)
        ReportQuery.PrintReport(Me, "Report2011", "TR_05_005_R2", MyValue)

    End Sub
End Class
