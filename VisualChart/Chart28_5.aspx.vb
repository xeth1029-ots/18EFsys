Partial Class Chart28_5
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
        Call creatChartA_5()
    End Sub


    Sub creatChartA_5() '產投：指標5_政策性產業參訓人數統計
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
        'sql &= " With WC1 As (Select NAME2 PLANNAME,VALUE ORGKIND2 FROM V_ORGKIND1 With(NOLOCK))" & vbCrLf
        'sql &= " ,WC1B AS (SELECT a.DISTID,b.ORGKIND2,a.DISTNAME3+'_'+b.PLANNAME COLLATE Chinese_Taiwan_Stroke_CS_AS DISTPLAN_N FROM V_DISTRICT a WITH(NOLOCK) CROSS JOIN WC1 b WHERE a.DISTID NOT IN ('000','002'))" & vbCrLf
        'sql &= " ,WC2 AS (" & vbCrLf
        'sql &= " Select cc.DISTID,cc.ORGKIND2,COUNT(1) CNT1 " & vbCrLf
        'sql &= " From V_SELRESULT a WITH(NOLOCK) " & vbCrLf
        'sql &= " Join VIEW2 cc With(NOLOCK) on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And a.ADMISSION ='Y'" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID In ('28')" & vbCrLf
        'sql &= " GROUP BY cc.DISTID, cc.ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC3 AS (" & vbCrLf
        'sql &= " Select cc.DISTID,cc.ORGKIND2,COUNT(1) CNT1 " & vbCrLf
        'sql &= " From V_ENTERTYPE2 A WITH(NOLOCK) " & vbCrLf
        'sql &= " Join VIEW2 cc With(NOLOCK) on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY cc.DISTID, cc.ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC4 AS (" & vbCrLf
        'sql &= " Select cc.DISTID,cc.ORGKIND2,COUNT(1) CNT1 " & vbCrLf
        'sql &= " ,COUNT(CASE WHEN a.STUDSTATUS Not IN (2,3) And a.FTDATE<=GETDATE() THEN 1 END) CNT2" & vbCrLf
        'sql &= " From V_STUDENTINFO A WITH(NOLOCK) " & vbCrLf
        'sql &= " Join VIEW2 cc With(NOLOCK) on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY cc.DISTID, cc.ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC5 AS (" & vbCrLf
        'sql &= " Select cc.DISTID,cc.ORGKIND2,COUNT(1) CNT1 " & vbCrLf
        'sql &= " ,COUNT(CASE WHEN a.STUDSTATUS Not IN (2,3) And a.FTDATE<=GETDATE() THEN 1 END) CNT2" & vbCrLf
        'sql &= " From VIEW_SUBSIDYCOST A WITH(NOLOCK) " & vbCrLf
        'sql &= " Join VIEW2 cc With(NOLOCK) on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And A.APPLIEDSTATUS=1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID In ('28')" & vbCrLf
        'sql &= " GROUP BY cc.DISTID, cc.ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf

        'sql &= " Select t.DISTID,t.DISTPLAN_N " & vbCrLf
        'sql &= " ,ISNULL(c2.CNT1,0) CNT1 --核定人數" & vbCrLf
        'sql &= " ,ISNULL(c3.CNT1,0) CNT2 --報名人數" & vbCrLf
        'sql &= " ,ISNULL(c4.CNT1,0) CNT3 --開訓人數" & vbCrLf
        'sql &= " ,ISNULL(c4.CNT2,0) CNT4 --結訓人數" & vbCrLf
        'sql &= " ,ISNULL(c5.CNT2,0) CNT5 --撥款人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c3.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c4.CNT1,0))/CONVERT(float,ISNULL(c3.CNT1,0)) ,2) End RATE1 --錄取率=開訓人數/報名人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c4.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c4.CNT2,0))/CONVERT(float,ISNULL(c4.CNT1,0)) ,2) End RATE2 --結訓率=結訓人數/開訓人數" & vbCrLf
        'sql &= " From WC1B t" & vbCrLf
        'sql &= " Left Join WC2 c2 on t.DISTID=c2.DISTID And t.ORGKIND2=c2.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC3 c3 on t.DISTID=c3.DISTID And t.ORGKIND2=c3.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC4 c4 on t.DISTID=c4.DISTID And t.ORGKIND2=c4.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " Left Join WC5 c5 on t.DISTID=c5.DISTID And t.ORGKIND2=c5.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        'sql &= " ORDER BY t.DISTID, t.ORGKIND2" & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART28_5 b" & vbCrLf
        sql &= " ORDER BY b.DISTID, b.ORGKIND2" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("DISTPLAN_N")) '分署名稱
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT1")) '核定人數
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT2")) '報名人數
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("CNT3")) '開訓人數
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("CNT4")) '結訓人數
            If vHid_data6 <> "" Then vHid_data6 &= ","
            vHid_data6 &= Convert.ToString(dr("CNT5")) '撥款人數
            If vHid_data7 <> "" Then vHid_data7 &= ","
            vHid_data7 &= Convert.ToString(dr("RATE1")) '錄取率
            If vHid_data8 <> "" Then vHid_data8 &= ","
            vHid_data8 &= Convert.ToString(dr("RATE2")) '結訓率

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