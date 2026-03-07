Partial Class CP_04_006
    Inherits AuthBasePage

    'rpt: CP_04_006
    'CP_04_006*.jrxml
    Const cst_printFN1 As String = "CP_04_006"

    Const cst_TPlanIDs As String = "'17','22','38','39','63','59'"

    '增加 63,59 20150409 BY AMU 
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not Page.IsPostBack Then
            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)

            Dim dt As DataTable
            Dim sqlstr As String = ""
            sqlstr = "SELECT NAME,DISTID FROM ID_DISTRICT ORDER BY DistID"
            dt = DbAccess.GetDataTable(sqlstr, objconn)
            With DistrictList
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "DistID"
                .DataBind()
                .Items.Insert(0, New ListItem("全部", ""))
            End With

            PlanList = TIMS.Get_TPlan(PlanList, , 1, "Y", "TPlanID IN (" & cst_TPlanIDs & ")")
        End If

        If sm.UserInfo.DistID = "000" Then
            DistrictList.Enabled = True
        Else
            DistrictList.SelectedValue = sm.UserInfo.DistID
            DistrictList.Enabled = False
        End If

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        '選擇全部訓練計畫
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"

    End Sub

    Private Sub print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles print.Click
        ''選擇轄區

        'Dim i As Integer = 0
        '報表要用的轄區參數
        Dim DistID1 As String = ""
        Dim DistName As String = ""
        'DistID1 = ""
        'DistName = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected AndAlso Me.DistrictList.Items(i).Value <> "" Then
                If DistID1 <> "" Then DistID1 &= ","
                DistID1 &= Convert.ToString("\'" & Me.DistrictList.Items(i).Value & "\'")

                If DistName <> "" Then DistName &= ","
                DistName &= Convert.ToString(Me.DistrictList.Items(i).Text)
            End If
        Next


        '報表要用的訓練計畫參數
        Dim TPlanID1 As String = ""
        Dim TPlanName As String = ""
        'TPlanID1 = ""
        'TPlanName = ""
        For i As Integer = 1 To Me.PlanList.Items.Count - 1
            If Me.PlanList.Items(i).Selected AndAlso Me.PlanList.Items(i).Value <> "" Then
                If TPlanID1 <> "" Then TPlanID1 &= ","
                TPlanID1 &= Convert.ToString("\'" & Me.PlanList.Items(i).Value & "\'")

                If TPlanName <> "" Then TPlanName &= ","
                TPlanName &= Convert.ToString(Me.PlanList.Items(i).Text)
            End If
        Next

        Dim str As String = "rs=1"
        str += "&Years=" & yearlist.SelectedValue
        str += "&DistID=" & DistID1
        str += "&DistName=" & Server.UrlEncode(DistName)
        str += "&STDate1=" & STDate1.Text
        str += "&STDate2=" & STDate2.Text
        str += "&FTDate1=" & FTDate1.Text
        str += "&FTDate2=" & FTDate2.Text
        str += "&TPlanID=" & TPlanID1
        str += "&PlanName=" & Server.UrlEncode(TPlanName)

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, str)
    End Sub

End Class
