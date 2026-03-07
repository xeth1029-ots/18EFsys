Partial Class Chart06_2
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
        Call creatChartB_2()
    End Sub


    Sub creatChartB_2() '自辦在職：指標2_各分署辦理訓練班次統計
        Dim vHid_data1 As String = ""
        Dim vHid_data2 As String = ""
        Dim vHid_data3 As String = ""
        Dim vHid_data4 As String = ""
        Dim vHid_data5 As String = ""

        Dim dt As New DataTable
        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT a.DISTID,(b.DISTNAME3+'_'+ Case when a.TPLANID='06' then '在職' when a.TPLANID='07' then '企委' else null End) as DISTNAME" & vbCrLf
        'sql &= " ,COUNT(1) CNT1 --核定開訓班數" & vbCrLf
        'sql &= " ,COUNT(CASE WHEN CONVERT(date,a.STDATE)<=CONVERT(date,GETDATE()) THEN 1 END) CNT2 --已開訓班數" & vbCrLf
        'sql &= " ,ROUND(CONVERT(float,COUNT(CASE WHEN CONVERT(date,a.STDATE)<=CONVERT(date,GETDATE()) THEN 1 END))/CONVERT(float,COUNT(1)),2) RATE1 --訓練班數達成率" & vbCrLf
        'sql &= " From VIEW2 a WITH(NOLOCK)" & vbCrLf
        'sql &= " Join V_DISTRICT b WITH(NOLOCK) on a.DISTID=b.DISTID" & vbCrLf
        'sql &= " Where 1 = 1" & vbCrLf
        'sql &= " And a.YEARS ='2018'" & vbCrLf
        'sql &= "--And CONVERT(date,a.STDATE)<=CONVERT(date,GETDATE())" & vbCrLf
        'sql &= "And a.TPLANID IN ('06','07')" & vbCrLf
        'sql &= "GROUP BY a.DISTID,b.DISTNAME3, a.TPLANID" & vbCrLf
        'sql &= "ORDER BY a.DISTID" & vbCrLf

        'Dim sql As String = ""
        'sql = "" & vbCrLf
        'sql &= " SELECT b.DISTID" & vbCrLf
        'sql &= " ,b.DISTNAME" & vbCrLf
        'sql &= " ,b.CNT1" & vbCrLf
        'sql &= " ,b.CNT2" & vbCrLf
        'sql &= " ,b.RATE1" & vbCrLf
        Dim sql As String = ""
        sql &= " SELECT * FROM ADP_CHART06_2 b" & vbCrLf
        sql &= " ORDER BY b.DISTID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each dr As DataRow In dt.Rows
            If vHid_data5 <> "" Then vHid_data5 &= ","
            vHid_data5 &= Convert.ToString(dr("DISTNAME"))
            If vHid_data1 <> "" Then vHid_data1 &= ","
            vHid_data1 &= Convert.ToString(dr("CNT1"))
            If vHid_data2 <> "" Then vHid_data2 &= ","
            vHid_data2 &= Convert.ToString(dr("CNT2"))
            If vHid_data3 <> "" Then vHid_data3 &= ","
            vHid_data3 &= Convert.ToString(dr("RATE1"))
            'If vHid_data4 <> "" Then vHid_data4 &= ","
            'vHid_data4 &= Convert.ToString(dr("RATE41"))
        Next

        Hid_data1.Value = vHid_data1
        Hid_data2.Value = vHid_data2
        Hid_data3.Value = vHid_data3
        Hid_data4.Value = vHid_data4
        Hid_data5.Value = vHid_data5
    End Sub

End Class