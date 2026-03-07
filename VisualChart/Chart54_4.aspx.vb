Partial Class Chart54_4
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
        Call creatChartA_4()
    End Sub


    Sub creatChartA_4() '產投：指標4_19大類指標統計
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
        'sql &= " With WC1 As (Select TMID,JOBID,JOBNAME FROM VIEW_TRAINTYPE WITH(NOLOCK) WHERE BUSID ='H' AND JOBID IS NOT NULL )" & vbCrLf
        'sql &= " ,WC1B AS (SELECT JOBID,JOBNAME FROM VIEW_TRAINTYPE WITH(NOLOCK) WHERE BUSID ='H' AND JOBID IS NOT NULL AND TRAINID IS NULL)" & vbCrLf
        'sql &= " ,WC2 AS (" & vbCrLf
        'sql &= " --核定人數" & vbCrLf
        'sql &= " Select t.JOBID,t.JOBNAME,SUM(cc.TNUM) CNT1 " & vbCrLf
        'sql &= " From VIEW2 cc WITH(NOLOCK)" & vbCrLf
        'sql &= " Join WC1 t on t.TMID=cc.TMID" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID In ('28')" & vbCrLf
        'sql &= " GROUP BY t.JOBID, t.JOBNAME" & vbCrLf
        'sql &= " )" & vbCrLf

        'sql &= " ,WC2A AS (" & vbCrLf
        'sql &= " --報名人數" & vbCrLf
        'sql &= " Select t.JOBID,t.JOBNAME,COUNT(1) CNT1 " & vbCrLf
        'sql &= " From V_ENTERTYPE2 A WITH(NOLOCK)" & vbCrLf
        'sql &= " Join VIEW2 cc on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Join WC1 t on t.TMID=cc.TMID" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY t.JOBID, t.JOBNAME" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC2B AS (" & vbCrLf
        ''sql &= " --開訓人數、結訓人數" & vbCrLf
        'sql &= " Select t.JOBID,t.JOBNAME,COUNT(1) CNT1 " & vbCrLf
        'sql &= " ,COUNT(CASE WHEN a.STUDSTATUS Not IN (2,3) And a.FTDATE<=GETDATE() THEN 1 END) CNT2" & vbCrLf
        'sql &= " From V_STUDENTINFO a WITH(NOLOCK)" & vbCrLf
        'sql &= " Join VIEW2 cc on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Join WC1 t on t.TMID=cc.TMID" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID IN ('28')" & vbCrLf
        'sql &= " GROUP BY t.JOBID, t.JOBNAME" & vbCrLf
        'sql &= " )" & vbCrLf

        'sql &= " ,WC2C AS (" & vbCrLf
        ''sql &= " --撥款人數" & vbCrLf
        'sql &= " Select t.JOBID,t.JOBNAME,COUNT(1) CNT1 " & vbCrLf
        'sql &= " ,COUNT(CASE WHEN a.STUDSTATUS Not IN (2,3) And a.FTDATE<=GETDATE() THEN 1 END) CNT2" & vbCrLf
        'sql &= " From VIEW_SUBSIDYCOST a WITH(NOLOCK)" & vbCrLf
        'sql &= " Join VIEW2 cc on cc.ocid =a.ocid" & vbCrLf
        'sql &= " Join WC1 t on t.TMID=cc.TMID" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And A.APPLIEDSTATUS=1" & vbCrLf
        'sql &= " And cc.YEARS ='2018'" & vbCrLf
        'sql &= " And cc.TPLANID In ('28')" & vbCrLf
        'sql &= " GROUP BY t.JOBID, t.JOBNAME" & vbCrLf
        'sql &= " )" & vbCrLf

        'sql &= " Select t.JOBID,t.JOBNAME" & vbCrLf
        'sql &= " ,ISNULL(c2.CNT1,0) CNT1 --核定人數" & vbCrLf
        'sql &= " ,ISNULL(c2a.CNT1,0) CNT2 --報名人數" & vbCrLf
        'sql &= " ,ISNULL(c2b.CNT1,0) CNT3 --開訓人數" & vbCrLf
        'sql &= " ,ISNULL(c2b.CNT2,0) CNT4 --結訓人數" & vbCrLf
        'sql &= " ,ISNULL(c2c.CNT1,0) CNT5 --撥款人數" & vbCrLf
        ''sql &= " --,ISNULL(c2c.CNT2,0) 結訓撥款人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c2a.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c2b.CNT1,0))/CONVERT(float,ISNULL(c2a.CNT1,0)) ,2) End RATE1 --錄取率=開訓人數/報名人數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(c2b.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(c2b.CNT2,0))/CONVERT(float,ISNULL(c2b.CNT1,0)) ,2) End RATE2 --結訓率=結訓人數/開訓人數" & vbCrLf
        'sql &= " From WC1B t" & vbCrLf
        'sql &= " Left Join WC2 c2 on c2.JOBID=t.JOBID" & vbCrLf
        'sql &= " Left Join WC2a c2a on c2a.JOBID=t.JOBID" & vbCrLf
        'sql &= " Left Join WC2b c2b on c2b.JOBID=t.JOBID" & vbCrLf
        'sql &= " Left Join WC2c c2c on c2c.JOBID=t.JOBID" & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART54_4 b" & vbCrLf
        sql &= " ORDER BY b.JOBID, b.JOBNAME" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("JOBNAME")) '19大類名稱
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