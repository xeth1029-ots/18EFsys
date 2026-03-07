Partial Class Chart06_3
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
            Call sCreate3() '頁面初始化
        End If

    End Sub

    '頁面初始化
    Sub sCreate3()
        Call createMFChartB_3()
    End Sub


    Sub createMFChartB_3() '自辦在職：指標3_參訓年齡/性別分佈
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""

        Dim dt As New DataTable
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " WITH WC1 AS (" & vbCrLf
        'sql &= "    Select YEARSOLDT4" & vbCrLf
        'sql &= "    ,COUNT(CASE WHEN SEX='M' THEN 1 END) CNT_M" & vbCrLf
        'sql &= " 	,COUNT(CASE WHEN SEX='F' THEN 1 END) CNT_F" & vbCrLf
        'sql &= " 	From V_STUDENTINFO s WITH(NOLOCK) " & vbCrLf
        'sql &= " 	Where 1 = 1" & vbCrLf
        'sql &= " 	And YEARS ='2018'" & vbCrLf
        ''sql &= " 	--And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " 	And TPLANID IN ('06','07')" & vbCrLf
        'sql &= " 	GROUP BY YEARSOLDT4" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " Select m.YID,m.YNAME " & vbCrLf
        'sql &= " ,ISNULL(a.CNT_M * -1,0) CNT_M ---男" & vbCrLf
        'sql &= " ,ISNULL(a.CNT_F ,0) CNT_F ---女" & vbCrLf
        'sql &= " From V_YEARSOLD4 m WITH(NOLOCK) " & vbCrLf
        'sql &= " Left Join WC1 a on a.YEARSOLDT4=m.YID" & vbCrLf
        'sql &= " ORDER BY m.YID" & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART06_3 b" & vbCrLf
        sql &= " ORDER BY b.YID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("YNAME"))
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("CNT_M"))
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT_F"))
        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
    End Sub

End Class