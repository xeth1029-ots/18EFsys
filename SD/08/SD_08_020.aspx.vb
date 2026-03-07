Partial Class SD_08_020
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_08_020"

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁'檢查Session是否存在 Start' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)'檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
            'FTDate2.Text = Now.Date
        End If

        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        Print.Attributes("onclick") = "return search();"

    End Sub

    Sub CreateItem()
        TPlanID = TIMS.Get_TPlan(TPlanID)
        TPlanID.Items.Insert(0, New ListItem("全部", ""))
        'TPlanID.Items(0).Selected = True
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected AndAlso Me.TPlanID.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= "\'" & Me.TPlanID.Items(i).Value & "\'"

                If TPlanName <> "" Then TPlanName &= ","
                TPlanName &= Me.TPlanID.Items(i).Text
            End If
        Next

        Dim RID As String = sm.UserInfo.RID
        Dim NewRID As String = Left(RID, 1)

        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "STDate1", STDate1.Text)
        TIMS.SetMyValue(MyValue, "STDate2", STDate2.Text)
        TIMS.SetMyValue(MyValue, "TPlanID", TPlanID1)
        TIMS.SetMyValue(MyValue, "PlanName", TPlanName)
        TIMS.SetMyValue(MyValue, "RID", NewRID)
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

    End Sub

End Class
