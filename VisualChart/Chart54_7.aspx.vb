Partial Class Chart54_7
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    Dim strSS As String = ""

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call sCreate() '頁面初始化
        End If

    End Sub

    '頁面初始化
    Sub sCreate()
        Call creatChartA_7()
    End Sub


    Sub creatChartA_7() '產投：指標7_政策性產業經費統計
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""
        Dim vHid_data6 As String = ""

        Dim dt As New DataTable
        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART54_7 b" & vbCrLf
        sql &= " ORDER BY b.DISTID" & vbCrLf ', b.ORGKIND2
        dt = DbAccess.GetDataTable(sql, objconn)

        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("DISTPLAN_N")) '分署名稱
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT6_1")) '提案班數
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT6_2")) '核定班數
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("CNT6_3")) '開訓班數
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("CNT6_4")) '結訓班數
            If vHid_data6 <> "" Then vHid_data6 &= ","
            vHid_data6 &= Convert.ToString(dr("RATE1")) '開訓率
        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
        Hid_data6.Value = vHid_data6

    End Sub

End Class