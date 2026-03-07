Partial Class TR_05_007_R
    Inherits AuthBasePage

    'ReportQuery
    'TR_05_007_R
    Const cst_printFN1 As String = "TR_05_007_R"

    '外部帶入 年度訓練計畫特定對象
    Const Cst_MIdentityID As String = "'02','03','04','05','06','07','10','13','14','27','28','35','36'"

    'Dim objconn As SqlConnection
    'Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Call TIMS.CloseDbConn(objconn)
    'End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        'objconn = DbAccess.GetConnection()
        'AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
        End If

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        ''選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        Syear = TIMS.GetSyear(Syear)
        Common.SetListItem(Syear, Now.Year) '年度

        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim msg As String = ""
        If Me.STDate1.Text = "" AndAlso Me.STDate2.Text = "" Then
            If Me.FTDate1.Text = "" AndAlso Me.FTDate2.Text = "" Then
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
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then
                If DistID1 <> "" Then DistID1 += ","
                DistID1 += "\'" & Me.DistID.Items(i).Value & "\'"
            End If
        Next

        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            If Me.TPlanID.Items(i).Selected Then
                If TPlanID1 <> "" Then TPlanID1 += ","
                TPlanID1 += Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
            End If
        Next

        Dim MyValue As String = ""
        MyValue += "&Years=" & Syear.SelectedValue
        MyValue += "&STDate1=" & STDate1.Text
        MyValue += "&STDate2=" & STDate2.Text
        MyValue += "&FTDate1=" & FTDate1.Text
        MyValue += "&FTDate2=" & FTDate2.Text
        MyValue += "&DistID=" & DistID1
        MyValue += "&TPlanID=" & TPlanID1
        MyValue += "&MIdentityID=" & Replace(Cst_MIdentityID, "'", "\'")
        ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

    End Sub
End Class
