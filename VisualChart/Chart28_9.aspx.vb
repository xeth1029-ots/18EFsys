Partial Class Chart28_9
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
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        ''多重長條圖:核定人數、報名人數、開訓人數、結訓人數、撥款人數
        'sql &= " With WC1 As (Select NAME2 PLANNAME,VALUE ORGKIND2 FROM V_ORGKIND1)" & vbCrLf
        'sql &= " ,WC2 AS (" & vbCrLf
        'sql &= " Select a.DISTID,b.ORGKIND2,a.CONTACTNAME+'_'+b.PLANNAME COLLATE Chinese_Taiwan_Stroke_CS_AS DISTPLAN_N" & vbCrLf
        'sql &= " From V_DISTRICT a " & vbCrLf
        'sql &= " CROSS Join WC1 b" & vbCrLf
        'sql &= " WHERE a.DISTID Not IN ('000','002')" & vbCrLf
        'sql &= " --ORDER BY a.DISTID,b.ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC3 AS (" & vbCrLf
        ''核定人數
        'sql &= " Select DISTID,ORGKIND2,SUM(TNUM) CNT1" & vbCrLf
        'sql &= " From VIEW2" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC32 AS (" & vbCrLf
        ''報名人數
        'sql &= " Select DISTID,ORGKIND2,SUM(TNUM) CNT1" & vbCrLf
        'sql &= " From V_ENTERTYPE2" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC33 AS (" & vbCrLf
        ''撥款人數
        'sql &= " Select DISTID,ORGKIND2,COUNT(1) CNT1" & vbCrLf
        'sql &= " From VIEW_SUBSIDYCOST" & vbCrLf
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
        'sql &= " ,ROUND(CONVERT(float,COUNT(CASE WHEN STUDSTATUS Not IN (2,3) And FTDATE<=GETDATE() THEN 1 END))/CONVERT(float,COUNT(1)),2) RATE1 --達成率" & vbCrLf
        'sql &= " From V_STUDENTINFO" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        ''sql &= " --And CONVERT(date,STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        ''sql &= " --And DISTID ='001' AND RID ='B'" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        ''多重長條圖:核定人數、報名人數、開訓人數、結訓人數、撥款人數
        'sql &= " Select b.DISTID,b.ORGKIND2,b.DISTPLAN_N" & vbCrLf
        'sql &= " ,ISNULL(c.CNT1,0) CNT1 --核定人數" & vbCrLf
        'sql &= " ,ISNULL(c2.CNT1,0) CNT2 --報名人數" & vbCrLf
        'sql &= " ,ISNULL(c3.CNT1,0) CNT3 --撥款人數" & vbCrLf
        'sql &= " ,ISNULL(c4.CNT1,0) CNT4 --開訓人數" & vbCrLf
        'sql &= " ,ISNULL(c4.CNT2,0) CNT5 --結訓人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c4.CNT1,0))/CONVERT(float,ISNULL(c2.CNT1,0)) * 100,1) End RATE21 --錄取率 = 開訓人數/報名人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c4.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c4.CNT2,0))/CONVERT(float,ISNULL(c4.CNT1,0)) * 100,1) End RATE22 --結訓率 = 結訓人數/開訓人數" & vbCrLf
        'sql &= " --,ISNULL(c4.RATE1,0) RATE41" & vbCrLf
        'sql &= " From WC2 b" & vbCrLf
        'sql &= " Left Join WC3 c ON c.DISTID=b.DISTID And c.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC32 c2 ON c2.DISTID=b.DISTID And c2.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC33 c3 ON c3.DISTID=b.DISTID And c3.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC4 c4 ON c4.DISTID=b.DISTID And c4.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " ORDER BY b.DISTID, b.ORGKIND2" & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART28_9 b" & vbCrLf
        sql &= " ORDER BY b.DISTID, b.ORGKIND2" & vbCrLf
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