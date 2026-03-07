Partial Class TR_04_002_R
    Inherits AuthBasePage

    'smartQuery
    'TR_04_002_R @TR
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

        If Not IsPostBack Then
            CreateItem()
        End If
        ' Button1.Attributes("onclick") = "javascript:return print();"
    End Sub

    Sub CreateItem()
        For i As Integer = Now.Year To 2005 Step -1
            SYear.Items.Add(i)
            FYear.Items.Add(i)
        Next
        For i As Integer = 1 To 12
            SMonth.Items.Add(i)
            FMonth.Items.Add(i)
        Next
        Common.SetListItem(SMonth, Now.Month - 3)
        Common.SetListItem(FMonth, Now.Month)


        Dim RGSTNRound As String
        RGSTNRound = ""
        RGSTNRound += " AND (Station_Scheme_ID='A' and Station_Unit_ID='31' and Station_ID='000')"
        RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='32' and Station_ID='000')"
        RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='33' and Station_ID='000')"
        RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='34' and Station_ID='000')"
        RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='35' and Station_ID='000')"
        RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='41' and Station_ID='000')"
        RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='51' and Station_ID='000')"

        Dim dt As DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT Station_Name" & vbCrLf
        sql += " ,Station_Scheme_ID+Station_Unit_ID Station_ID" & vbCrLf
        sql += " FROM Adp_WorkStation" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += RGSTNRound
        dt = DbAccess.GetDataTable(sql, objconn)

        With Station
            .DataSource = dt
            .DataTextField = "Station_Name"
            .DataValueField = "Station_ID"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", "%"))
            .Items(0).Selected = True
        End With
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim stdate_start, stdate_end, title, range_list As String
        stdate_start = Convert.ToString(SYear.SelectedValue) & "/" & Convert.ToString(SMonth.SelectedValue) & "/1"
        stdate_end = CStr(DateAdd(DateInterval.Day, -1, DateAdd(DateInterval.Month, 1, CDate(Convert.ToString(FYear.SelectedValue) & "/" & Convert.ToString(FMonth.SelectedValue) & "/1"))))
        range_list = "統計區間: " & Convert.ToString(SYear.SelectedValue) & "/" & Convert.ToString(SMonth.SelectedValue) & "~" & Convert.ToString(FYear.SelectedValue) & "/" & Convert.ToString(FMonth.SelectedValue)
        If Station.SelectedValue <> "%" Then
            title = Convert.ToString(Station.SelectedItem.Text) & "  推介失業者參加職業訓練成果統計表(一)"
        Else
            title = "推介失業者參加職業訓練成果統計表(一)"
        End If

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&FTDate=" & stdate_start
        MyValue += "&FTDate2=" & stdate_end
        MyValue += "&CREATE_RGSTN=" & Station.SelectedValue
        MyValue += "&title=" & Server.UrlEncode(title)
        MyValue += "&range_list=" & Server.UrlEncode(range_list)

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "TR", "TR_04_002_R", MyValue)
    End Sub
End Class
