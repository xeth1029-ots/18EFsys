Partial Class Chart28_3
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
        Call creatChartA_3()
    End Sub


    Sub creatChartA_3() '產投：指標3_總體補助費指標
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""
        Dim vHid_data6 As String = ""

        Dim dt As New DataTable
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        ''多重長條圖:申請補助費、核定補助費、預估補助費、核撥補助費
        'sql &= " With WC1 As (Select NAME2 PLANNAME,VALUE ORGKIND2 FROM V_ORGKIND1 WITH(NOLOCK))" & vbCrLf
        'sql &= " ,WC2 AS (" & vbCrLf
        'sql &= " Select a.DISTID,b.ORGKIND2,a.DISTNAME3+'_'+b.PLANNAME COLLATE Chinese_Taiwan_Stroke_CS_AS DISTPLAN_N" & vbCrLf
        'sql &= " From V_DISTRICT a WITH(NOLOCK) " & vbCrLf
        'sql &= " CROSS Join WC1 b" & vbCrLf
        'sql &= " WHERE a.DISTID Not IN ('000','002')" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC3 AS (" & vbCrLf
        'sql &= " Select DISTID,ORGKIND2" & vbCrLf
        'sql &= " ,SUM(SUMOFMONEY) SUMOFMONEY1" & vbCrLf '申請補助費
        'sql &= " ,SUM(CASE WHEN APPLIEDSTATUSM='Y' THEN SUMOFMONEY END) SUMOFMONEY2" & vbCrLf '核定補助費
        'sql &= " ,SUM(CASE WHEN APPLIEDSTATUSM='Y' AND APPLIEDSTATUS=1 THEN SUMOFMONEY END) SUMOFMONEY3" & vbCrLf '核撥補助費
        'sql &= " From VIEW_SUBSIDYCOST WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " --And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC32 AS (" & vbCrLf
        ''預估補助費
        'sql &= " Select DISTID,ORGKIND2, SUM(DEFGOVCOST) DEFGOVCOST" & vbCrLf
        'sql &= " From VIEW2 WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " --And CONVERT(Date,STDATE)<=CONVERT(Date,GETDATE())" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " Select a.DISTID,a.ORGKIND2,a.DISTPLAN_N" & vbCrLf
        'sql &= " ,ISNULL(b.SUMOFMONEY1,0) CNT1" & vbCrLf '申請補助費
        'sql &= " ,ISNULL(b.SUMOFMONEY2,0) CNT2" & vbCrLf '核定補助費
        'sql &= " ,ISNULL(b2.DEFGOVCOST,0) CNT3" & vbCrLf '預估補助費
        'sql &= " ,ISNULL(b.SUMOFMONEY3,0) CNT4" & vbCrLf '核撥補助費
        'sql &= " ,CASE WHEN ISNULL(b.SUMOFMONEY2,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(b.SUMOFMONEY3,0))/CONVERT(float,ISNULL(b.SUMOFMONEY2,0)) ,2) End RATE1" & vbCrLf '執行率=核撥補助費/核定補助費
        'sql &= " From WC2 a" & vbCrLf
        'sql &= " Left Join WC3 b on a.DISTID=b.DISTID And a.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS " & vbCrLf
        'sql &= " Left Join WC32 b2 on a.DISTID=b2.DISTID And a.ORGKIND2=b2.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS " & vbCrLf
        'sql &= " ORDER BY a.DISTID, a.ORGKIND2" & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART28_3 b" & vbCrLf
        sql &= " ORDER BY b.DISTID, b.ORGKIND2" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("DISTPLAN_N")) '分署名稱
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT1")) '申請補助費
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT2")) '核定補助費
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("CNT3")) '預估補助費
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("CNT4")) '核撥補助費
            If vHid_data6 <> "" Then vHid_data6 &= ","
            vHid_data6 &= Convert.ToString(dr("RATE1")) '執行率

        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
        Hid_data6.Value = vHid_data6

    End Sub

End Class