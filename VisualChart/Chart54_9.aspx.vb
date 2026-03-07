Partial Class Chart54_9
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
        Call creatChartA_9()
    End Sub


    Sub creatChartA_9() '產投：指標9_三年補助使用情形
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""

        Dim dt As New DataTable
        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART54_9 b" & vbCrLf
        sql &= " ORDER BY b.DISTID" & vbCrLf ', b.ORGKIND2
        dt = DbAccess.GetDataTable(sql, objconn)

        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("DISTPLAN_N")) '分署名稱
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT1")) '三年內第一次使用 (分列五分署)
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT2")) '第二~五次使用 (分列五分署)
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("CNT4")) '第六次(含)以上使用 (分列五分署)
        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4

    End Sub

End Class