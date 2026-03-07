Partial Class TR_05_006_R
    Inherits AuthBasePage

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

        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String

        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year) '年

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        'Dim TPlanID1 As String
        'Dim newTPlanID As String
        'Dim TPlanName As String
        'Dim newTPlanIDName As String
        'Dim i As Integer
        'Dim objitem As ListItem

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

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += "\'" & Me.TPlanID.Items(i).Value & "\'"
                If TPlanName <> "" Then TPlanName += ","
                TPlanName += Me.TPlanID.Items(i).Text
            End If
        Next

        'If TPlanID1 <> "" Then
        '    newTPlanID = Mid(TPlanID1, 1, TPlanID1.Length - 1)
        '    newTPlanIDName = Mid(TPlanName, 1, TPlanName.Length - 1)
        'End If

        Dim myValue As String = ""
        myValue &= "&Years=" & Syear.SelectedValue
        myValue &= "&STDate1=" & STDate1.Text
        myValue &= "&STDate2=" & STDate2.Text
        myValue &= "&FTDate1=" & FTDate1.Text
        myValue &= "&FTDate2=" & FTDate2.Text
        myValue &= "&TPlanID=" & TPlanID1
        myValue &= "&PlanName=" & TPlanName
        ReportQuery.PrintReport(Me, "TR", "TR_05_006_R", myValue)

    End Sub
End Class
