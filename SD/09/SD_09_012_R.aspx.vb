Partial Class SD_09_012_R
    Inherits AuthBasePage

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

        Me.start_date.Attributes("onchange") = "check_date();"
        Me.end_date.Attributes("onchange") = "check_date();"

        If Not Page.IsPostBack Then
            Call Create1()
        End If

        Button1.Attributes("onclick") = "javascript:return print();"

        '選擇全部轄區
        DistrictList.Attributes("onclick") = "SelectAll('DistrictList','DistHidden');"

        '選擇全部訓練計畫
        PlanList.Attributes("onclick") = "SelectAll('PlanList','TPlanHidden');"

    End Sub

    Sub Create1()

        Dim sqlstr As String = ""
        Dim dt As DataTable = Nothing
        sqlstr = "SELECT NAME,DISTID FROM ID_DISTRICT ORDER BY DISTID"
        dt = DbAccess.GetDataTable(sqlstr, objconn)

        Me.DistrictList.DataSource = dt
        Me.DistrictList.DataTextField = "Name"
        Me.DistrictList.DataValueField = "DistID"
        Me.DistrictList.DataBind()
        Me.DistrictList.Items.Insert(0, New ListItem("全部", ""))

        sqlstr = "SELECT PLANNAME,TPLANID FROM KEY_PLAN ORDER BY TPLANID"
        dt = DbAccess.GetDataTable(sqlstr, objconn)
        Me.PlanList.DataSource = dt
        Me.PlanList.DataTextField = "PlanName"
        Me.PlanList.DataValueField = "TPlanID"
        Me.PlanList.DataBind()
        Me.PlanList.Items.Insert(0, New ListItem("全部", ""))
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim msg As String = ""
        Dim STYear1 As String = ""
        Dim STYear2 As String = ""
        If (Me.start_date.Text <> "") Then STYear1 = Year(Me.start_date.Text)
        If (Me.end_date.Text <> "") Then STYear2 = Year(Me.end_date.Text)
        If (Me.start_date.Text <> "") And (STYear1 < 2008) Then msg += "退訓起日年度請設在西元2008年之後!" & vbCrLf
        If (Me.end_date.Text <> "") And (STYear2 < 2008) Then msg += "退訓迄日年度請設在西元2008年之後!" & vbCrLf
        If msg <> "" Then
            Common.MessageBox(Me, msg)
            Exit Sub
        End If

        '報表要用的轄區參數
        Dim DistID As String = ""
        DistID = ""
        For i As Integer = 1 To Me.DistrictList.Items.Count - 1
            If Me.DistrictList.Items(i).Selected Then
                If DistID <> "" Then DistID += ","
                DistID += Convert.ToString("\'" & Me.DistrictList.Items(i).Value & "\'")
            End If
        Next
        'If DistID <> "" Then
        '    newDistID = Mid(DistID, 1, DistID.Length - 1)
        'End If

        '報表要用的訓練計畫參數
        Dim TPlanID As String = ""
        TPlanID = ""
        For i As Integer = 1 To Me.PlanList.Items.Count - 1
            If Me.PlanList.Items(i).Selected Then
                If TPlanID <> "" Then TPlanID += ","
                TPlanID += Convert.ToString("\'" & Me.PlanList.Items(i).Value & "\'")
            End If
        Next
        'If TPlanID <> "" Then
        '    newTPlanID = Mid(TPlanID, 1, TPlanID.Length - 1)
        'End If

        Dim MyValue As String = ""
        MyValue = "&start_date=" & start_date.Text & "&end_date=" & end_date.Text & "&TPlanID=" & TPlanID & "&DistID=" & DistID & ""
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "SD_09_012_R", MyValue)

    End Sub
End Class
