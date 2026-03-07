Partial Class VisualCHart
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
            Call sCreate1() '頁面初始化
        End If

    End Sub

    '頁面初始化
    Sub sCreate1()
        Call creatChartB_1()
    End Sub

    Sub creatChartB_1() '自辦在職：指標1_各分署辦理訓練人次統計
        '[100, 150, 120, 110, 100, 200, 150, 130, 100, 140], //訓練目標人數
        '   [60, 70, 62, 55, 45, 115, 72, 68, 92, 88],	//開訓人數
        '   [52, 38, 45, 29, 45, 96, 63, 67, 82, 54],	//結訓人數
        '   [0.22, 0.33, 0.66, 0.44, 0.36, 0.68, 0.99, 0.86, 0.57, 0.66],	//訓練達成率
        'Dim vHid_data1 As String = "100,150,120,110,100,200,150,130,100,140"
        'Dim vHid_data2 As String = "60,70,62,55,45,115,72,68,92,88"
        'Dim vHid_data3 As String = "52,38,45,29,45,96,63,67,82,54"
        'Dim vHid_data4 As String = "0.22,0.33,0.66,0.44,0.36,0.68,0.99,0.86,0.57,0.66"
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""

        Dim dt As New DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (select TPLANID,PLANNAME from key_plan where tplanid in('06','07'))" & vbCrLf
        sql &= " ,WC2 AS (" & vbCrLf
        sql &= " SELECT a.DISTID,b.TPLANID,a.DISTNAME+'_'+b.PLANNAME  DISTPLAN_N" & vbCrLf
        sql &= " FROM V_DISTRICT a" & vbCrLf
        sql &= " CROSS JOIN WC1 b" & vbCrLf
        sql &= " WHERE a.DISTID NOT IN ('000','002')" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WC3 AS (" & vbCrLf
        sql &= " SELECT DISTID,TPLANID,SUM(TNUM) CNT1" & vbCrLf
        sql &= " FROM VIEW2" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND YEARS ='2018'" & vbCrLf
        sql &= " AND CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        sql &= " AND TPLANID IN ('06','07')" & vbCrLf
        sql &= " GROUP BY DISTID,TPLANID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WC4 AS (" & vbCrLf
        sql &= " SELECT DISTID,TPLANID" & vbCrLf
        sql &= " ,COUNT(1) CNT1 --開訓人數" & vbCrLf
        sql &= " ,COUNT(CASE WHEN STUDSTATUS NOT IN (2,3) AND FTDATE<=GETDATE() THEN 1 END) CNT2 --結訓人數" & vbCrLf
        sql &= " ,ROUND(CONVERT(float,COUNT(CASE WHEN STUDSTATUS NOT IN (2,3) AND FTDATE<=GETDATE() THEN 1 END))/CONVERT(float,COUNT(1)),2) RATE1 --達成率" & vbCrLf
        sql &= " FROM V_STUDENTINFO" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND YEARS ='2018'" & vbCrLf
        sql &= " AND TPLANID IN ('06','07')" & vbCrLf
        sql &= " GROUP BY DISTID,TPLANID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT b.DISTID,b.TPLANID,b.DISTPLAN_N" & vbCrLf
        sql &= " ,ISNULL(c.CNT1,0) CNT1" & vbCrLf
        sql &= " ,ISNULL(c4.CNT1,0) CNT41" & vbCrLf
        sql &= " ,ISNULL(c4.CNT2,0) CNT42" & vbCrLf
        sql &= " ,ISNULL(c4.RATE1,0) RATE41" & vbCrLf
        sql &= " FROM WC2 b" & vbCrLf
        sql &= " LEFT JOIN WC3 c ON c.DISTID=b.DISTID and c.TPLANID=b.TPLANID" & vbCrLf
        sql &= " LEFT JOIN WC4 c4 ON c4.DISTID=b.DISTID and c4.TPLANID=b.TPLANID" & vbCrLf
        sql &= " ORDER BY b.DISTID,b.TPLANID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("DISTPLAN_N"))
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("CNT1"))
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT41"))
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT42"))
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("RATE41"))
        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
    End Sub

    Sub creatChartB_2() '自辦在職：指標2_各分署辦理訓練班次統計
        'Dim vHid_data1 As String = "100,150,120,110,100,200,150,130,100,140"
        'Dim vHid_data2 As String = "60,70,62,55,45,115,72,68,92,88"
        'Dim vHid_data3 As String = "52,38,45,29,45,96,63,67,82,54"
        'Dim vHid_data4 As String = "0.22,0.33,0.66,0.44,0.36,0.68,0.99,0.86,0.57,0.66"
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""

        Dim dt As New DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (select TPLANID,PLANNAME from key_plan where tplanid in('06','07'))" & vbCrLf
        sql &= " ,WC2 AS (" & vbCrLf
        sql &= " SELECT a.DISTID,b.TPLANID,a.DISTNAME+'_'+b.PLANNAME  DISTPLAN_N" & vbCrLf
        sql &= " FROM V_DISTRICT a" & vbCrLf
        sql &= " CROSS JOIN WC1 b" & vbCrLf
        sql &= " WHERE a.DISTID NOT IN ('000','002')" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WC3 AS (" & vbCrLf
        sql &= " SELECT DISTID,TPLANID,SUM(TNUM) CNT1" & vbCrLf
        sql &= " FROM VIEW2" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND YEARS ='2018'" & vbCrLf
        sql &= " AND CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        sql &= " AND TPLANID IN ('06','07')" & vbCrLf
        sql &= " GROUP BY DISTID,TPLANID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WC4 AS (" & vbCrLf
        sql &= " SELECT DISTID,TPLANID" & vbCrLf
        sql &= " ,COUNT(1) CNT1 --開訓人數" & vbCrLf
        sql &= " ,COUNT(CASE WHEN STUDSTATUS NOT IN (2,3) AND FTDATE<=GETDATE() THEN 1 END) CNT2 --結訓人數" & vbCrLf
        sql &= " ,ROUND(CONVERT(float,COUNT(CASE WHEN STUDSTATUS NOT IN (2,3) AND FTDATE<=GETDATE() THEN 1 END))/CONVERT(float,COUNT(1)),2) RATE1 --達成率" & vbCrLf
        sql &= " FROM V_STUDENTINFO" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND YEARS ='2018'" & vbCrLf
        sql &= " AND TPLANID IN ('06','07')" & vbCrLf
        sql &= " GROUP BY DISTID,TPLANID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT b.DISTID,b.TPLANID,b.DISTPLAN_N" & vbCrLf
        sql &= " ,ISNULL(c.CNT1,0) CNT1" & vbCrLf
        sql &= " ,ISNULL(c4.CNT1,0) CNT41" & vbCrLf
        sql &= " ,ISNULL(c4.CNT2,0) CNT42" & vbCrLf
        sql &= " ,ISNULL(c4.RATE1,0) RATE41" & vbCrLf
        sql &= " FROM WC2 b" & vbCrLf
        sql &= " LEFT JOIN WC3 c ON c.DISTID=b.DISTID and c.TPLANID=b.TPLANID" & vbCrLf
        sql &= " LEFT JOIN WC4 c4 ON c4.DISTID=b.DISTID and c4.TPLANID=b.TPLANID" & vbCrLf
        sql &= " ORDER BY b.DISTID,b.TPLANID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("DISTPLAN_N"))
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("CNT1"))
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT41"))
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT42"))
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("RATE41"))
        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
    End Sub





End Class