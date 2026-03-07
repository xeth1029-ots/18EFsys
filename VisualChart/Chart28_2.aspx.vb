Partial Class Chart28_2
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
        Call creatChartA_2()
    End Sub

    Sub creatChartA_2() '產投：指標2_總體開班數指標

        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""
        Dim vHid_data6 As String = ""
        Dim vHid_data7 As String = ""

        Dim dt As New DataTable
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " With WC1 As (Select NAME2 PLANNAME,VALUE ORGKIND2 FROM V_ORGKIND1 WITH(NOLOCK))" & vbCrLf
        'sql &= " ,WC2 AS (" & vbCrLf
        'sql &= " Select a.DISTID,b.ORGKIND2,a.DISTNAME3+'_'+b.PLANNAME COLLATE Chinese_Taiwan_Stroke_CS_AS DISTPLAN_N" & vbCrLf
        'sql &= " From V_DISTRICT a WITH(NOLOCK) " & vbCrLf
        'sql &= " CROSS Join WC1 b" & vbCrLf
        'sql &= " WHERE a.DISTID Not IN ('000','002')" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " ,WC3 AS (" & vbCrLf
        'sql &= " Select DISTID,ORGKIND2" & vbCrLf
        'sql &= " ,COUNT(1) CNT1 --提案班數" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN APPLIEDRESULT ='Y' THEN 1 END) CNT2 --核定班數" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN ISSUCCESS='Y' AND NOTOPEN='N' AND OCID IS NOT NULL THEN 1 END) CNT3 --開訓班數" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN ISSUCCESS='Y' AND NOTOPEN='N' AND OCID IS NOT NULL AND FTDATE<=CONVERT(date,GETDATE()) THEN 1 END) CNT4 --結訓班數" & vbCrLf
        'sql &= " From VIEW2B WITH(NOLOCK) " & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And YEARS ='2018'" & vbCrLf
        'sql &= " And TPLANID IN ('28')" & vbCrLf
        'sql &= " --And CONVERT(Date,STDATE)<=CONVERT(Date,GETDATE())" & vbCrLf
        'sql &= " GROUP BY DISTID, ORGKIND2" & vbCrLf
        'sql &= " --ORDER BY DISTID,ORGKIND2" & vbCrLf
        'sql &= " )" & vbCrLf
        'sql &= " Select a.DISTID,a.ORGKIND2,a.DISTPLAN_N" & vbCrLf
        'sql &= " ,ISNULL(b.CNT1,0) CNT1" & vbCrLf '提案班數
        'sql &= " ,ISNULL(b.CNT2,0) CNT2" & vbCrLf '核定班數
        'sql &= " ,ISNULL(b.CNT3,0) CNT3" & vbCrLf '開訓班數
        'sql &= " ,ISNULL(b.CNT4,0) CNT4" & vbCrLf '結訓班數
        'sql &= " ,CASE WHEN ISNULL(b.CNT1,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(b.CNT3,0))/CONVERT(float,ISNULL(b.CNT1,0)) ,2) End RATE1 --開訓率=開訓班數/提案班數" & vbCrLf
        'sql &= " ,CASE WHEN ISNULL(b.CNT3,0) =0 THEN 0" & vbCrLf
        'sql &= " Else ROUND(CONVERT(float,ISNULL(b.CNT4,0))/CONVERT(float,ISNULL(b.CNT3,0)) ,2) End RATE2 --結訓率=結訓班數/開訓班數" & vbCrLf
        'sql &= " From WC2 a" & vbCrLf
        'sql &= " Left Join WC3 b on a.DISTID=b.DISTID And a.ORGKIND2=b.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS " & vbCrLf
        'sql &= " ORDER BY a.DISTID, a.ORGKIND2" & vbCrLf

        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART28_2 b" & vbCrLf
        sql &= " ORDER BY b.DISTID, b.ORGKIND2" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("DISTPLAN_N"))
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT1"))
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("CNT2"))
            If vHid_data4 <> "" Then vHid_data4 &= ","
            vHid_data4 &= Convert.ToString(dr("CNT3"))
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("CNT4"))
            If vHid_data6 <> "" Then vHid_data6 &= ","
            vHid_data6 &= Convert.ToString(dr("RATE1"))
            If vHid_data7 <> "" Then vHid_data7 &= ","
            vHid_data7 &= Convert.ToString(dr("RATE2"))

        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
        Hid_data6.Value = vHid_data6
        Hid_data7.Value = vHid_data7

    End Sub

End Class