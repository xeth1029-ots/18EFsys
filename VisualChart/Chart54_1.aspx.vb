Partial Class Chart54_1
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
            Call sCreate2() '頁面初始化
        End If

    End Sub

    '頁面初始化
    Sub sCreate2()
        Call creatChartA_1()
    End Sub


    Sub creatChartA_1() '產投：指標1_總體參訓人數指標
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""
        Dim vHid_data6 As String = ""
        Dim vHid_data7 As String = ""
        Dim vHid_data8 As String = ""

        Dim dt As New DataTable
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        ''多重長條圖:核定人數、報名人數、開訓人數、結訓人數、撥款人數
        'sql &= " With WC1 As (Select NAME2 PLANNAME,VALUE ORGKIND2 FROM V_ORGKIND1 WITH(NOLOCK))" & vbCrLf
        'sql &= " ,WC2 AS (" & vbCrLf
        'sql &= " Select a.DISTID,b.ORGKIND2,a.DISTNAME3+'_'+b.PLANNAME COLLATE Chinese_Taiwan_Stroke_CS_AS DISTPLAN_N" & vbCrLf
        'sql &= " From V_DISTRICT a WITH(NOLOCK) " & vbCrLf
        'sql &= " CROSS Join WC1 b" & vbCrLf
        'sql &= " WHERE a.DISTID Not IN ('000','002')" & vbCrLf
        'sql &= " --ORDER BY a.DISTID,b.ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC3 AS (" & vbCrLf
        ''核定人數
        'sql &= " Select DISTID,ORGKIND2,SUM(TNUM) CNT1" & vbCrLf
        'sql &= " From VIEW2 WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC32 AS (" & vbCrLf
        ''報名人數
        'sql &= " Select DISTID,ORGKIND2,SUM(TNUM) CNT1" & vbCrLf
        'sql &= " From V_ENTERTYPE2 WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC33 AS (" & vbCrLf
        ''撥款人數
        'sql &= " Select DISTID,ORGKIND2,COUNT(1) CNT1" & vbCrLf
        'sql &= " From VIEW_SUBSIDYCOST WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And APPLIEDSTATUS=1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC4 AS (" & vbCrLf
        'sql &= " Select DISTID,ORGKIND2" & vbCrLf
        'sql &= " ,COUNT(1) CNT1 --開訓人數" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN STUDSTATUS Not IN (2,3) And FTDATE<=GETDATE() THEN 1 END) CNT2 --結訓人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(CONVERT(float,COUNT(1)),0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,COUNT(CASE WHEN STUDSTATUS Not IN (2,3) And FTDATE<=GETDATE() THEN 1 END))/CONVERT(float,COUNT(1)),2) End RATE1 --達成率" & vbCrLf
        'sql &= " FROM V_STUDENTINFO WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " --And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " --And DISTID ='001' AND RID ='B'" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        ''多重長條圖:核定人數、報名人數、開訓人數、結訓人數、撥款人數
        'sql &= " Select b.DISTID,b.ORGKIND2,b.DISTPLAN_N" & vbCrLf
        'sql &= " ,ISNULL(c.CNT1,0) CNT1 --核定人數" & vbCrLf
        'sql &= " ,ISNULL(c2.CNT1,0) CNT2 --報名人數" & vbCrLf
        'sql &= " ,ISNULL(c3.CNT1,0) CNT3 --撥款人數" & vbCrLf
        'sql &= " ,ISNULL(c4.CNT1,0) CNT4 --開訓人數" & vbCrLf
        'sql &= " ,ISNULL(c4.CNT2,0) CNT5 --結訓人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c2.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c4.CNT1,0))/CONVERT(float,ISNULL(c2.CNT1,0)) ,2) End RATE21 --錄取率 = 開訓人數/報名人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c4.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c4.CNT2,0))/CONVERT(float,ISNULL(c4.CNT1,0)) ,2) End RATE22 --結訓率 = 結訓人數/開訓人數" & vbCrLf
        'sql &= " --,ISNULL(c4.RATE1,0) RATE41" & vbCrLf
        'sql &= " From WC2 b" & vbCrLf
        'sql &= " Left Join WC3 c ON c.DISTID=b.DISTID And c.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC32 c2 ON c2.DISTID=b.DISTID And c2.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC33 c3 ON c3.DISTID=b.DISTID And c3.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC4 c4 ON c4.DISTID=b.DISTID And c4.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " ORDER BY b.DISTID, b.ORGKIND2" & vbCrLf


        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART54_1 b" & vbCrLf
        sql &= " ORDER BY b.DISTID" & vbCrLf ', b.ORGKIND2
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("DISTPLAN_N")) '分署名稱
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT1")) '核定人數
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT2")) '報名人數
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("CNT4")) '開訓人數
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("CNT5")) '結訓人數
            If vHid_data6 <> "" Then vHid_data6 &= ","
            vHid_data6 &= Convert.ToString(dr("CNT3")) '撥款人數
            If vHid_data7 <> "" Then vHid_data7 &= ","
            vHid_data7 &= Convert.ToString(dr("RATE21")) '錄取率
            If vHid_data8 <> "" Then vHid_data8 &= ","
            vHid_data8 &= Convert.ToString(dr("RATE22")) '結訓率

        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
        Hid_data6.Value = vHid_data6
        Hid_data7.Value = vHid_data7
        Hid_data8.Value = vHid_data8

    End Sub

End Class