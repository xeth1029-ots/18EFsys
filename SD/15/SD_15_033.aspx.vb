Partial Class SD_15_033
    Inherits AuthBasePage

    Const cst_CalcMode_依參訓人次 As String = "1"
    Const cst_CalcMode_依百分比 As String = "2"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            cCreate1()
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Sub cCreate1()
        msg1.Text = ""

        ddlYEARS_SCH1 = TIMS.GetSyear(ddlYEARS_SCH1)
        ddlYEARS_SCH2 = TIMS.GetSyear(ddlYEARS_SCH2)
        Common.SetListItem(ddlYEARS_SCH1, sm.UserInfo.Years)
        Common.SetListItem(ddlYEARS_SCH2, sm.UserInfo.Years)

        OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
        Common.SetListItem(OrgKind2, "A")
        'table3,'ddlYEARS_SCH1,'ddlYEARS_SCH2,'trPlanKind,'OrgKind2,'trCalcMode1 計算方式
        'CalcMode1 計算方式 1：依參訓人次／2：依百分比,'RBListExpType,'BTN_EXPORT1,'匯出,'msg1,
    End Sub

    Protected Sub BTN_EXPORT1_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT1.Click
        Dim v_RblCalcMode1 As String = TIMS.GetListValue(RblCalcMode1)
        Dim iCalcMode As Integer = Val(v_RblCalcMode1)
        Select Case v_RblCalcMode1
            Case cst_CalcMode_依參訓人次
                Call EXPORT_1(iCalcMode)

            Case cst_CalcMode_依百分比
                Call EXPORT_2(iCalcMode)

            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return

        End Select
    End Sub

    Function GET_SQL_WC2() As String
        Dim sqlwc2 As String = "SELECT SM1_RANGE, SM1_TXT
FROM (VALUES (0,'補助費用10,000以下')
,(1,'補助費用10,001－20,000')
,(2,'補助費用20,001－30,000')
,(3,'補助費用30,001－40,000')
,(4,'補助費用40,001－50,000')
,(5,'補助費用50,001－60,000')
,(6,'補助費用60,001－70,000')
,(7,'補助費用70,001－80,000')
,(8,'補助費用80,001－90,000')
,(9,'補助費用90,001－99,999')
,(10,'補助費用達100,000')) AS T(SM1_RANGE, SM1_TXT)"
        Return sqlwc2
    End Function

    Function sSearch1_DATA_dt1() As DataTable

        Dim V_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim V_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim V_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim sPMS As New Hashtable From {
            {"TPLANID", sm.UserInfo.TPlanID},
            {"YEARS1", V_ddlYEARS_SCH1},
            {"YEARS2", V_ddlYEARS_SCH2}
        }
        Dim s_WC2 As String = GET_SQL_WC2()
        Dim sSql As String = ""
        sSql &= " WITH WC1 AS ( select (SELECT CASE WHEN cct.SUMOFMONEY1 <=10000 THEN 0 WHEN cct.SUMOFMONEY1 <=20000 THEN 1 WHEN cct.SUMOFMONEY1<=30000 THEN 2" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<=40000 THEN 3 WHEN cct.SUMOFMONEY1<=50000 THEN 4 WHEN cct.SUMOFMONEY1<=60000 THEN 5" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<=70000 THEN 6 WHEN cct.SUMOFMONEY1<=80000 THEN 7 WHEN cct.SUMOFMONEY1<=90000 THEN 8" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<100000 THEN 9 ELSE 10 END) SM1_RANGE" & vbCrLf '/*已撥款總累積補助級距*/
        'sSql &= " WHEN cct.SUMOFMONEY1<70000 THEN 6 WHEN cct.SUMOFMONEY1=70000 THEN 7 ELSE 8 END) SM1_RANGE /*已撥款總累積補助級距*/" & vbCrLf
        'sSql &= " WHEN cct.SUMOFMONEY1<70000 THEN 6 ELSE 7 END) SM1_RANGE" & vbCrLf '/*已撥款總累積補助級距*/
        sSql &= " ,ss.ORGKINDGW,ss.IDNO,ig3.GCODE31" & vbCrLf
        sSql &= " ,ISNULL(CASE WHEN ct.APPLIEDSTATUS=1 THEN ct.SUMOFMONEY END,0) SUMOFMONEY"
        sSql &= " FROM dbo.V_STUDENTINFO ss" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBSIDYCOST ct WITH(NOLOCK) on ct.SOCID =ss.SOCID" & vbCrLf
        sSql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ss.GCID3=ig3.GCID3" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBCOST cct WITH(NOLOCK) on cct.COSTTYPE=4 and cct.idno=ss.idno and cct.stdate1<=ss.stdate and ss.stdate<=cct.stdate2" & vbCrLf
        sSql &= " WHERE (ct.APPLIEDSTATUSM='Y' OR ct.APPLIEDSTATUS=1)" & vbCrLf
        sSql &= " AND ss.TPLANID=@TPLANID AND ss.YEARS>=@YEARS1 AND ss.YEARS<=@YEARS2" & vbCrLf
        Select Case V_OrgKind2
            Case "G", "W"
                sPMS.Add("ORGKIND2", V_OrgKind2)
                sSql &= " AND ss.ORGKIND2>=@ORGKIND2" & vbCrLf
        End Select
        sSql &= " )" & vbCrLf
        'sSql &= " AND ss.TPLANID='28' AND ss.YEARS='2023' )" & vbCrLf
        sSql &= String.Concat(" ,WC2 AS (", s_WC2, ")", vbCrLf)

        sSql &= " ,WC3 AS ( SELECT SM1_RANGE" & vbCrLf
        sSql &= " ,COUNT(DISTINCT IDNO) STDCNT1" & vbCrLf
        sSql &= " ,COUNT(1) STDCNT2" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '01' THEN 1 END) GC01" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '02' THEN 1 END) GC02" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '03' THEN 1 END) GC03" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '04' THEN 1 END) GC04" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '05' THEN 1 END) GC05" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '06' THEN 1 END) GC06" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '07' THEN 1 END) GC07" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '08' THEN 1 END) GC08" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '09' THEN 1 END) GC09" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '10' THEN 1 END) GC10" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '11' THEN 1 END) GC11" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '12' THEN 1 END) GC12" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '13' THEN 1 END) GC13" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '14' THEN 1 END) GC14" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '15' THEN 1 END) GC15" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '16' THEN 1 END) GC16" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '17' THEN 1 END) GC17" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '18' THEN 1 END) GC18" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '19' THEN 1 END) GC19" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN GCODE31 IS NOT NULL THEN 1 END) SUBTOTAL" & vbCrLf
        sSql &= " FROM WC1" & vbCrLf
        sSql &= " GROUP BY SM1_RANGE )" & vbCrLf

        sSql &= " ,WC4 AS ( select 91 SM1_RANGE,'合計' SM1_TXT" & vbCrLf
        sSql &= " ,SUM(STDCNT1) STDCNT1" & vbCrLf
        sSql &= " ,SUM(STDCNT2) STDCNT2" & vbCrLf
        sSql &= " ,SUM(GC01) GC01" & vbCrLf
        sSql &= " ,SUM(GC02) GC02" & vbCrLf
        sSql &= " ,SUM(GC03) GC03" & vbCrLf
        sSql &= " ,SUM(GC04) GC04" & vbCrLf
        sSql &= " ,SUM(GC05) GC05" & vbCrLf
        sSql &= " ,SUM(GC06) GC06" & vbCrLf
        sSql &= " ,SUM(GC07) GC07" & vbCrLf
        sSql &= " ,SUM(GC08) GC08" & vbCrLf
        sSql &= " ,SUM(GC09) GC09" & vbCrLf
        sSql &= " ,SUM(GC10) GC10" & vbCrLf
        sSql &= " ,SUM(GC11) GC11" & vbCrLf
        sSql &= " ,SUM(GC12) GC12" & vbCrLf
        sSql &= " ,SUM(GC13) GC13" & vbCrLf
        sSql &= " ,SUM(GC14) GC14" & vbCrLf
        sSql &= " ,SUM(GC15) GC15" & vbCrLf
        sSql &= " ,SUM(GC16) GC16" & vbCrLf
        sSql &= " ,SUM(GC17) GC17" & vbCrLf
        sSql &= " ,SUM(GC18) GC18" & vbCrLf
        sSql &= " ,SUM(GC19) GC19" & vbCrLf
        sSql &= " ,SUM(SUBTOTAL) SUBTOTAL" & vbCrLf
        sSql &= " FROM WC3 )" & vbCrLf

        sSql &= " ,WC4B AS ( select 92 SM1_RANGE,'總補助金額' SM1_TXT" & vbCrLf
        sSql &= " ,'' STDCNT1" & vbCrLf
        sSql &= " ,'' STDCNT2" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '01' THEN SUMOFMONEY END) GC01" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '02' THEN SUMOFMONEY END) GC02" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '03' THEN SUMOFMONEY END) GC03" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '04' THEN SUMOFMONEY END) GC04" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '05' THEN SUMOFMONEY END) GC05" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '06' THEN SUMOFMONEY END) GC06" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '07' THEN SUMOFMONEY END) GC07" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '08' THEN SUMOFMONEY END) GC08" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '09' THEN SUMOFMONEY END) GC09" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '10' THEN SUMOFMONEY END) GC10" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '11' THEN SUMOFMONEY END) GC11" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '12' THEN SUMOFMONEY END) GC12" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '13' THEN SUMOFMONEY END) GC13" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '14' THEN SUMOFMONEY END) GC14" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '15' THEN SUMOFMONEY END) GC15" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '16' THEN SUMOFMONEY END) GC16" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '17' THEN SUMOFMONEY END) GC17" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '18' THEN SUMOFMONEY END) GC18" & vbCrLf
        sSql &= " ,SUM(CASE GCODE31 WHEN '19' THEN SUMOFMONEY END) GC19" & vbCrLf
        sSql &= " ,SUM(CASE WHEN GCODE31 IS NOT NULL THEN SUMOFMONEY END) SUBTOTAL" & vbCrLf
        sSql &= " FROM WC1 )" & vbCrLf

        'sSql &= " ,WC4 AS (" & vbCrLf
        'sSql &= " select 99 SM1_RANGE,'合計' SM1_TXT" & vbCrLf
        'sSql &= " ,COUNT(DISTINCT IDNO) STDCNT1" & vbCrLf
        'sSql &= " ,COUNT(1) STDCNT2" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '01' THEN 1 END) GC01" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '02' THEN 1 END) GC02" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '03' THEN 1 END) GC03" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '04' THEN 1 END) GC04" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '05' THEN 1 END) GC05" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '06' THEN 1 END) GC06" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '07' THEN 1 END) GC07" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '08' THEN 1 END) GC08" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '09' THEN 1 END) GC09" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '10' THEN 1 END) GC10" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '11' THEN 1 END) GC11" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '12' THEN 1 END) GC12" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '13' THEN 1 END) GC13" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '14' THEN 1 END) GC14" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '15' THEN 1 END) GC15" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '16' THEN 1 END) GC16" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '17' THEN 1 END) GC17" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '18' THEN 1 END) GC18" & vbCrLf
        'sSql &= " ,COUNT(CASE GCODE31 WHEN '19' THEN 1 END) GC19" & vbCrLf
        'sSql &= " ,COUNT(CASE WHEN GCODE31 IS NOT NULL THEN 1 END) SUBTOTAL" & vbCrLf
        'sSql &= " from WC1 )" & vbCrLf

        sSql &= " SELECT b.SM1_RANGE,b.SM1_TXT" & vbCrLf
        sSql &= " ,concat('',c.STDCNT1) STDCNT1" & vbCrLf
        sSql &= " ,concat('',c.STDCNT2) STDCNT2" & vbCrLf
        sSql &= " ,c.GC01,c.GC02,c.GC03,c.GC04,c.GC05,c.GC06,c.GC07,c.GC08" & vbCrLf
        sSql &= " ,c.GC09,c.GC10,c.GC11,c.GC12,c.GC13,c.GC14,c.GC15,c.GC16" & vbCrLf
        sSql &= " ,c.GC17,c.GC18,c.GC19,c.SUBTOTAL" & vbCrLf
        sSql &= " FROM WC2 b" & vbCrLf
        sSql &= " LEFT JOIN WC3 c on c.SM1_RANGE=b.SM1_RANGE" & vbCrLf
        sSql &= " UNION" & vbCrLf
        sSql &= " select c.SM1_RANGE,c.SM1_TXT" & vbCrLf
        sSql &= " ,concat('',c.STDCNT1) STDCNT1" & vbCrLf
        sSql &= " ,concat('',c.STDCNT2) STDCNT2" & vbCrLf
        sSql &= " ,c.GC01,c.GC02,c.GC03,c.GC04,c.GC05,c.GC06,c.GC07,c.GC08" & vbCrLf
        sSql &= " ,c.GC09,c.GC10,c.GC11,c.GC12,c.GC13,c.GC14,c.GC15,c.GC16" & vbCrLf
        sSql &= " ,c.GC17,c.GC18,c.GC19,c.SUBTOTAL" & vbCrLf
        sSql &= " FROM WC4 c" & vbCrLf
        sSql &= " UNION" & vbCrLf
        sSql &= " select c.SM1_RANGE,c.SM1_TXT" & vbCrLf
        sSql &= " ,concat('',c.STDCNT1) STDCNT1" & vbCrLf
        sSql &= " ,concat('',c.STDCNT2) STDCNT2" & vbCrLf
        sSql &= " ,c.GC01,c.GC02,c.GC03,c.GC04,c.GC05,c.GC06,c.GC07,c.GC08" & vbCrLf
        sSql &= " ,c.GC09,c.GC10,c.GC11,c.GC12,c.GC13,c.GC14,c.GC15,c.GC16" & vbCrLf
        sSql &= " ,c.GC17,c.GC18,c.GC19,c.SUBTOTAL" & vbCrLf
        sSql &= " FROM WC4B c" & vbCrLf

        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        If TIMS.dtNODATA(dt1) Then Return dt1

        Dim STDCNT1 As Double = 0 '學員人數 
        Dim STDCNT2 As Double = 0 '參訓次數 
        Dim SUBTOTAL As Double = 0 '總補助金額
        Dim AVGSUMOFMONEY As Double = 0 '平均每人補助
        Dim AVGSTDCNT2 As Double = 0 '平均每人參訓次數
        Dim AVGSUBSIDYPER As Double = 0 '平均每次補助
        Dim dr91 As DataRow = Nothing
        Dim dr92 As DataRow = Nothing

        dr91 = If(dt1.Select("SM1_RANGE=91").Length > 0, dt1.Select("SM1_RANGE=91")(0), Nothing)
        If dr91 Is Nothing Then Return dt1
        dr92 = If(dt1.Select("SM1_RANGE=92").Length > 0, dt1.Select("SM1_RANGE=92")(0), Nothing)
        If dr92 Is Nothing Then Return dt1

        STDCNT1 = Val(dr91("STDCNT1"))
        STDCNT2 = Val(dr91("STDCNT2"))
        SUBTOTAL = Val(dr92("SUBTOTAL"))
        If (STDCNT1 > 0) Then AVGSUMOFMONEY = TIMS.ROUND(SUBTOTAL / STDCNT1)
        If (STDCNT1 > 0) Then AVGSTDCNT2 = TIMS.ROUND(STDCNT2 / STDCNT1, 1)
        If (STDCNT2 > 0) Then AVGSUBSIDYPER = TIMS.ROUND(SUBTOTAL / STDCNT2, 1)

        Dim t_AVGSUMOFMONEY As String = String.Concat("平均每人補助", AVGSUMOFMONEY, "元")
        Dim t_AVGSTDCNT2 As String = String.Concat("平均每人參訓", AVGSTDCNT2, "次")
        Dim t_AVGSUBSIDYPER As String = String.Concat("平均每次補助", AVGSUBSIDYPER, "元")
        dt1.Select("SM1_RANGE=92")(0)("STDCNT1") = String.Concat(t_AVGSUMOFMONEY, "<br/>", t_AVGSTDCNT2)
        dt1.Select("SM1_RANGE=92")(0)("STDCNT2") = String.Concat(t_AVGSUBSIDYPER)


        Dim dr93 As DataRow = dt1.NewRow
        dt1.Rows.Add(dr93)
        dr93("SM1_RANGE") = 93
        dr93("SM1_TXT") = "各職類平均每次補助金額"
        For ix As Integer = 1 To 19
            Dim COLNM1 As String = String.Concat("GC", ix.ToString("00"))
            If Not IsDBNull(dr92(COLNM1)) AndAlso Not IsDBNull(dr91(COLNM1)) AndAlso Val(dr91(COLNM1)) > 0 Then
                dr93(COLNM1) = TIMS.ROUND(Val(dr92(COLNM1)) / Val(dr91(COLNM1)))
            End If
        Next
        If Not IsDBNull(dr92("SUBTOTAL")) AndAlso Not IsDBNull(dr91("SUBTOTAL")) AndAlso Val(dr91("SUBTOTAL")) > 0 Then
            dr93("SUBTOTAL") = TIMS.ROUND(dr92("SUBTOTAL") / dr91("SUBTOTAL"))
        End If

        Dim dr94o As DataRow = sSearch1_DATA_dr94o()
        If dr94o Is Nothing Then Return dt1

        Dim dr94 As DataRow = dt1.NewRow
        dt1.Rows.Add(dr94)
        dr94("SM1_RANGE") = 94
        dr94("SM1_TXT") = "各職類參訓人數"
        For ix As Integer = 1 To 19
            Dim COLNM1 As String = String.Concat("GC", ix.ToString("00"))
            dr94(COLNM1) = dr94o(COLNM1)
        Next
        dr94("SUBTOTAL") = dr94o("SUBTOTAL")

        Dim dr95 As DataRow = dt1.NewRow
        dt1.Rows.Add(dr95)
        dr95("SM1_RANGE") = 95
        dr95("SM1_TXT") = "各職類平均每人參訓次數"
        For ix As Integer = 1 To 19
            Dim COLNM1 As String = String.Concat("GC", ix.ToString("00"))
            If Not IsDBNull(dr91(COLNM1)) AndAlso Not IsDBNull(dr94o(COLNM1)) AndAlso Val(dr94o(COLNM1)) > 0 Then
                dr95(COLNM1) = TIMS.ROUND(Val(dr91(COLNM1)) / Val(dr94o(COLNM1)), 2)
            End If
        Next
        If Not IsDBNull(dr91("SUBTOTAL")) AndAlso Not IsDBNull(dr94o("SUBTOTAL")) AndAlso Val(dr94o("SUBTOTAL")) > 0 Then
            dr95("SUBTOTAL") = TIMS.ROUND(dr91("SUBTOTAL") / dr94o("SUBTOTAL"), 2)
        End If
        Return dt1
    End Function

    Function sSearch1_DATA_dr94o() As DataRow
        Dim V_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim V_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim V_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim sPMS As New Hashtable From {
            {"TPLANID", sm.UserInfo.TPlanID},
            {"YEARS1", V_ddlYEARS_SCH1},
            {"YEARS2", V_ddlYEARS_SCH2}
        }
        Dim sSql As String = ""
        sSql &= " WITH WC1 AS ( select (SELECT CASE WHEN cct.SUMOFMONEY1 <=10000 THEN 0 WHEN cct.SUMOFMONEY1 <=20000 THEN 1 WHEN cct.SUMOFMONEY1<=30000 THEN 2" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<=40000 THEN 3 WHEN cct.SUMOFMONEY1<=50000 THEN 4 WHEN cct.SUMOFMONEY1<=60000 THEN 5" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<=70000 THEN 6 WHEN cct.SUMOFMONEY1<=80000 THEN 7 WHEN cct.SUMOFMONEY1<=90000 THEN 8" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<100000 THEN 9 ELSE 10 END) SM1_RANGE" & vbCrLf '/*已撥款總累積補助級距*/
        'sSql &= " WHEN cct.SUMOFMONEY1<70000 THEN 6 WHEN cct.SUMOFMONEY1=70000 THEN 7 ELSE 8 END) SM1_RANGE /*已撥款總累積補助級距*/" & vbCrLf
        'sSql &= " WHEN cct.SUMOFMONEY1<70000 THEN 6 ELSE 7 END) SM1_RANGE /*已撥款總累積補助級距*/" & vbCrLf
        sSql &= " ,ss.ORGKINDGW,ss.IDNO,ig3.GCODE31" & vbCrLf
        sSql &= " ,ISNULL(CASE WHEN ct.APPLIEDSTATUS=1 THEN ct.SUMOFMONEY END,0) SUMOFMONEY"
        sSql &= " FROM dbo.V_STUDENTINFO ss" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBSIDYCOST ct WITH(NOLOCK) on ct.SOCID =ss.SOCID" & vbCrLf
        sSql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ss.GCID3=ig3.GCID3" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBCOST cct WITH(NOLOCK) on cct.COSTTYPE=4 and cct.idno=ss.idno and cct.stdate1<=ss.stdate and ss.stdate<=cct.stdate2" & vbCrLf
        sSql &= " WHERE (ct.APPLIEDSTATUSM='Y' OR ct.APPLIEDSTATUS=1)" & vbCrLf
        sSql &= " AND ss.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND ss.YEARS>=@YEARS1 AND ss.YEARS<=@YEARS2" & vbCrLf
        Select Case V_OrgKind2
            Case "G", "W"
                sPMS.Add("ORGKIND2", V_OrgKind2)
                sSql &= " AND ss.ORGKIND2>=@ORGKIND2" & vbCrLf
        End Select
        sSql &= " )" & vbCrLf

        sSql &= " SELECT 1 N1" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='01') GC01" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='02') GC02" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='03') GC03" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='04') GC04" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='05') GC05" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='06') GC06" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='07') GC07" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='08') GC08" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='09') GC09" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='10') GC10" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='11') GC11" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='12') GC12" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='13') GC13" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='14') GC14" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='15') GC15" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='16') GC16" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='17') GC17" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='18') GC18" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31='19') GC19" & vbCrLf
        sSql &= " ,(SELECT COUNT(DISTINCT IDNO) FROM WC1 WHERE GCODE31 IS NOT NULL) SUBTOTAL" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        If TIMS.dtNODATA(dt1) Then Return Nothing
        Return dt1.Rows(0)
    End Function


    Sub CHK_EXPORT_1_VAL(ByRef sERRMSG As String)
        sERRMSG = ""

        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)

        If v_ddlYEARS_SCH1 = "" Then sERRMSG &= "請選擇 年度起始區間(起始) 不可為空!" & vbCrLf
        If v_ddlYEARS_SCH2 = "" Then sERRMSG &= "請選擇 年度起始區間(起始) 不可為空!" & vbCrLf

        If v_ddlYEARS_SCH1 <> "" AndAlso v_ddlYEARS_SCH2 <> "" Then
            If Val(v_ddlYEARS_SCH1) > Val(v_ddlYEARS_SCH2) Then sERRMSG &= String.Concat("年度起始區間順序有誤!.", v_ddlYEARS_SCH1, "~", v_ddlYEARS_SCH2, vbCrLf)
        End If
    End Sub

    Function GET_GOVCLASSCAST3_TB(oConn As SqlConnection) As DataTable
        Dim sSql As String = ""
        sSql &= " SELECT ig3.GCODE31,CONCAT('GC',ig3.GCODE31) CODE1,concat(ig3.GCODE31,'.',ig3.CNAME) GCODE31PN FROM V_GOVCLASSCAST3 ig3 WHERE PNAME IS NULL ORDER BY ig3.GCODE31" & vbCrLf
        Dim dt As New DataTable
        TIMS.OpenDbConn(oConn)
        Dim sCmd As New SqlCommand(sSql, oConn)
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        Return dt
    End Function

    Function GET_GOVCLASSCAST3_COMBIN(dt As DataTable, COLUMN_N As String) As String
        Dim rst As String = ""
        If TIMS.dtNODATA(dt) Then Return rst
        For Each dr1 As DataRow In dt.Rows
            rst &= String.Concat(If(rst <> "", ",", ""), dr1(COLUMN_N))
        Next
        Return rst
    End Function

    Sub EXPORT_1(iCalcMode As Integer)
        Dim sERRMSG As String = ""
        Call CHK_EXPORT_1_VAL(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Exit Sub
        End If
        Dim dtXls As DataTable = sSearch1_DATA_dt1()
        If dtXls Is Nothing OrElse dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim YEAR_N As String = (v_ddlYEARS_SCH1 - 1911)
        If v_ddlYEARS_SCH1 <> v_ddlYEARS_SCH2 Then
            YEAR_N = String.Concat((v_ddlYEARS_SCH1 - 1911), "~", (v_ddlYEARS_SCH2 - 1911))
        End If
        Dim TXT_ORGKIND2 As String = TIMS.GetListText(OrgKind2)
        TXT_ORGKIND2 = If(TXT_ORGKIND2 <> "", String.Concat("(", TXT_ORGKIND2, ")"), "")

        Dim ss_TitleS1 As String = String.Concat((sm.UserInfo.Years - 1911), "年度x產投方案學員參訓情形分析")
        Dim strFilename1 As String = String.Concat(ss_TitleS1, TIMS.GetDateNo2())
        Dim sTitle1 As String = String.Concat("產業人才投資方案", YEAR_N, "年度參訓學員補助費使用級距", TXT_ORGKIND2)

        Dim dt3 As DataTable = GET_GOVCLASSCAST3_TB(objconn)
        Dim sGCODE31PN As String = GET_GOVCLASSCAST3_COMBIN(dt3, "GCODE31PN")
        Dim sGOV3CODE1 As String = GET_GOVCLASSCAST3_COMBIN(dt3, "CODE1")
        'Dim sPattern As String = String.Concat("補助費使用級距,學員人數(ID不重複),參訓次數,", sGCODE31PN, ",小計")
        Dim sColumn As String = String.Concat("SM1_TXT,STDCNT1,STDCNT2,", sGOV3CODE1, ",SUBTOTAL")

        'Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim s_TitleHtml2 As String = ""
        s_TitleHtml2 &= "<tr>"
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "補助費使用級距")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "學員人數(ID不重複)")
        's_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "學員人數(累計)")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "參訓次數")
        s_TitleHtml2 &= String.Format("<td colspan=19>{0}</td>", "課程分類")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "小計")
        s_TitleHtml2 &= "</tr>"
        s_TitleHtml2 &= "<tr>"
        For Each SPN1 As String In sGCODE31PN.Split(",")
            s_TitleHtml2 &= String.Format("<td>{0}</td>", SPN1)
        Next
        s_TitleHtml2 &= "</tr>"

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        'parms.Add("sPatternA", sPatternA)
        parms.Add("TitleHtml2", s_TitleHtml2)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Function sSearch1_DATA_dt2() As DataTable
        Dim V_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim V_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim V_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim sPMS As New Hashtable From {
            {"TPLANID", sm.UserInfo.TPlanID},
            {"YEARS1", V_ddlYEARS_SCH1},
            {"YEARS2", V_ddlYEARS_SCH2}
        }

        Dim s_WC2 As String = GET_SQL_WC2()
        Dim sSql As String = ""
        sSql &= " WITH WC1 AS ( select (SELECT CASE WHEN cct.SUMOFMONEY1 <=10000 THEN 0 WHEN cct.SUMOFMONEY1 <=20000 THEN 1 WHEN cct.SUMOFMONEY1<=30000 THEN 2" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<=40000 THEN 3 WHEN cct.SUMOFMONEY1<=50000 THEN 4 WHEN cct.SUMOFMONEY1<=60000 THEN 5" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<=70000 THEN 6 WHEN cct.SUMOFMONEY1<=80000 THEN 7 WHEN cct.SUMOFMONEY1<=90000 THEN 8" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1<100000 THEN 9 ELSE 10 END) SM1_RANGE" & vbCrLf '/*已撥款總累積補助級距*/
        'sSql &= " WHEN cct.SUMOFMONEY1<70000 THEN 6 WHEN cct.SUMOFMONEY1=70000 THEN 7 ELSE 8 END) SM1_RANGE /*已撥款總累積補助級距*/" & vbCrLf
        'sSql &= " WHEN cct.SUMOFMONEY1<70000 THEN 6 ELSE 7 END) SM1_RANGE /*已撥款總累積補助級距*/" & vbCrLf
        sSql &= " ,ss.ORGKINDGW,ss.IDNO,ig3.GCODE31" & vbCrLf
        sSql &= " FROM dbo.V_STUDENTINFO ss" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBSIDYCOST ct WITH(NOLOCK) on ct.SOCID =ss.SOCID" & vbCrLf
        sSql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ss.GCID3=ig3.GCID3" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBCOST cct WITH(NOLOCK) on cct.COSTTYPE=4 and cct.idno=ss.idno and cct.stdate1<=ss.stdate and ss.stdate <=cct.stdate2" & vbCrLf
        sSql &= " WHERE (ct.APPLIEDSTATUSM='Y' OR ct.APPLIEDSTATUS=1)" & vbCrLf
        sSql &= " AND ss.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND ss.YEARS>=@YEARS1 AND ss.YEARS<=@YEARS2" & vbCrLf
        Select Case V_OrgKind2
            Case "G", "W"
                sPMS.Add("ORGKIND2", V_OrgKind2)
                sSql &= " AND ss.ORGKIND2>=@ORGKIND2" & vbCrLf
        End Select
        sSql &= " )" & vbCrLf
        'sSql &= " AND ss.TPLANID='28' AND ss.YEARS='2023' )" & vbCrLf
        sSql &= String.Concat(" ,WC2 AS (", s_WC2, ")", vbCrLf)

        sSql &= " ,WC3 AS ( SELECT SM1_RANGE" & vbCrLf
        sSql &= " ,COUNT(DISTINCT IDNO) STDCNT1" & vbCrLf
        sSql &= " ,COUNT(1) STDCNT2" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '01' THEN 1 END) GC01" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '02' THEN 1 END) GC02" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '03' THEN 1 END) GC03" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '04' THEN 1 END) GC04" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '05' THEN 1 END) GC05" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '06' THEN 1 END) GC06" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '07' THEN 1 END) GC07" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '08' THEN 1 END) GC08" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '09' THEN 1 END) GC09" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '10' THEN 1 END) GC10" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '11' THEN 1 END) GC11" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '12' THEN 1 END) GC12" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '13' THEN 1 END) GC13" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '14' THEN 1 END) GC14" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '15' THEN 1 END) GC15" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '16' THEN 1 END) GC16" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '17' THEN 1 END) GC17" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '18' THEN 1 END) GC18" & vbCrLf
        sSql &= " ,COUNT(CASE GCODE31 WHEN '19' THEN 1 END) GC19" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN GCODE31 IS NOT NULL THEN 1 END) SUBTOTAL" & vbCrLf
        sSql &= " FROM WC1" & vbCrLf
        sSql &= " GROUP BY SM1_RANGE )" & vbCrLf

        sSql &= " ,WC4 AS ( select 99 SM1_RANGE,'合計' SM1_TXT" & vbCrLf
        sSql &= " ,SUM(STDCNT1) STDCNT1" & vbCrLf
        sSql &= " ,SUM(STDCNT2) STDCNT2" & vbCrLf
        sSql &= " ,SUM(GC01) GC01" & vbCrLf
        sSql &= " ,SUM(GC02) GC02" & vbCrLf
        sSql &= " ,SUM(GC03) GC03" & vbCrLf
        sSql &= " ,SUM(GC04) GC04" & vbCrLf
        sSql &= " ,SUM(GC05) GC05" & vbCrLf
        sSql &= " ,SUM(GC06) GC06" & vbCrLf
        sSql &= " ,SUM(GC07) GC07" & vbCrLf
        sSql &= " ,SUM(GC08) GC08" & vbCrLf
        sSql &= " ,SUM(GC09) GC09" & vbCrLf
        sSql &= " ,SUM(GC10) GC10" & vbCrLf
        sSql &= " ,SUM(GC11) GC11" & vbCrLf
        sSql &= " ,SUM(GC12) GC12" & vbCrLf
        sSql &= " ,SUM(GC13) GC13" & vbCrLf
        sSql &= " ,SUM(GC14) GC14" & vbCrLf
        sSql &= " ,SUM(GC15) GC15" & vbCrLf
        sSql &= " ,SUM(GC16) GC16" & vbCrLf
        sSql &= " ,SUM(GC17) GC17" & vbCrLf
        sSql &= " ,SUM(GC18) GC18" & vbCrLf
        sSql &= " ,SUM(GC19) GC19" & vbCrLf
        sSql &= " ,SUM(SUBTOTAL) SUBTOTAL" & vbCrLf
        sSql &= " from WC3 )" & vbCrLf

        sSql &= " ,WC5 AS ( SELECT b.SM1_RANGE,b.SM1_TXT" & vbCrLf
        sSql &= " ,c.STDCNT1,c.STDCNT2" & vbCrLf
        sSql &= " ,c.GC01,c.GC02,c.GC03,c.GC04,c.GC05,c.GC06,c.GC07,c.GC08" & vbCrLf
        sSql &= " ,c.GC09,c.GC10,c.GC11,c.GC12,c.GC13,c.GC14,c.GC15,c.GC16" & vbCrLf
        sSql &= " ,c.GC17,c.GC18,c.GC19,c.SUBTOTAL" & vbCrLf
        sSql &= " FROM WC2 b" & vbCrLf
        sSql &= " LEFT JOIN WC3 c on c.SM1_RANGE=b.SM1_RANGE )" & vbCrLf

        sSql &= " ,WC5B AS ( SELECT b.SM1_RANGE" & vbCrLf
        sSql &= " ,(SELECT SUM(x.STDCNT1) STDCNT1 FROM WC5 x WHERE x.SM1_RANGE<=b.SM1_RANGE ) STDCNT1B" & vbCrLf
        sSql &= " FROM WC5 b" & " )" & vbCrLf

        sSql &= " SELECT b.SM1_RANGE,b.SM1_TXT" & vbCrLf
        'sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(c.STDCNT1,0)/a.STDCNT1*100,'0.00'),'%', c.STDCNT1) else '0%' end STDCNT1" & vbCrLf
        'sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(bb.STDCNT1B,0)/a.STDCNT1*100,'0.00'),'%', bb.STDCNT1B) else '0%' end STDCNT1B" & vbCrLf
        sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(c.STDCNT1,0)/a.STDCNT1*100,'0.00'),'%') else '0%' end STDCNT1" & vbCrLf
        sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(bb.STDCNT1B,0)/a.STDCNT1*100,'0.00'),'%') else '0%' end STDCNT1B" & vbCrLf
        sSql &= " ,case when a.STDCNT2>0 then concat(format(1.0*isnull(c.STDCNT2,0)/a.STDCNT2*100,'0.00'),'%') else '0%' end STDCNT2" & vbCrLf
        sSql &= " ,case when a.GC01>0 then concat(format(1.0*isnull(c.GC01,0)/a.GC01*100,'0.00'),'%') else '0%' end GC01" & vbCrLf
        sSql &= " ,case when a.GC02>0 then concat(format(1.0*isnull(c.GC02,0)/a.GC02*100,'0.00'),'%') else '0%' end GC02" & vbCrLf
        sSql &= " ,case when a.GC03>0 then concat(format(1.0*isnull(c.GC03,0)/a.GC03*100,'0.00'),'%') else '0%' end GC03" & vbCrLf
        sSql &= " ,case when a.GC04>0 then concat(format(1.0*isnull(c.GC04,0)/a.GC04*100,'0.00'),'%') else '0%' end GC04" & vbCrLf
        sSql &= " ,case when a.GC05>0 then concat(format(1.0*isnull(c.GC05,0)/a.GC05*100,'0.00'),'%') else '0%' end GC05" & vbCrLf
        sSql &= " ,case when a.GC06>0 then concat(format(1.0*isnull(c.GC06,0)/a.GC06*100,'0.00'),'%') else '0%' end GC06" & vbCrLf
        sSql &= " ,case when a.GC07>0 then concat(format(1.0*isnull(c.GC07,0)/a.GC07*100,'0.00'),'%') else '0%' end GC07" & vbCrLf
        sSql &= " ,case when a.GC08>0 then concat(format(1.0*isnull(c.GC08,0)/a.GC08*100,'0.00'),'%') else '0%' end GC08" & vbCrLf
        sSql &= " ,case when a.GC09>0 then concat(format(1.0*isnull(c.GC09,0)/a.GC09*100,'0.00'),'%') else '0%' end GC09" & vbCrLf
        sSql &= " ,case when a.GC10>0 then concat(format(1.0*isnull(c.GC10,0)/a.GC10*100,'0.00'),'%') else '0%' end GC10" & vbCrLf
        sSql &= " ,case when a.GC11>0 then concat(format(1.0*isnull(c.GC11,0)/a.GC11*100,'0.00'),'%') else '0%' end GC11" & vbCrLf
        sSql &= " ,case when a.GC12>0 then concat(format(1.0*isnull(c.GC12,0)/a.GC12*100,'0.00'),'%') else '0%' end GC12" & vbCrLf
        sSql &= " ,case when a.GC13>0 then concat(format(1.0*isnull(c.GC13,0)/a.GC13*100,'0.00'),'%') else '0%' end GC13" & vbCrLf
        sSql &= " ,case when a.GC14>0 then concat(format(1.0*isnull(c.GC14,0)/a.GC14*100,'0.00'),'%') else '0%' end GC14" & vbCrLf
        sSql &= " ,case when a.GC15>0 then concat(format(1.0*isnull(c.GC15,0)/a.GC15*100,'0.00'),'%') else '0%' end GC15" & vbCrLf
        sSql &= " ,case when a.GC16>0 then concat(format(1.0*isnull(c.GC16,0)/a.GC16*100,'0.00'),'%') else '0%' end GC16" & vbCrLf
        sSql &= " ,case when a.GC17>0 then concat(format(1.0*isnull(c.GC17,0)/a.GC17*100,'0.00'),'%') else '0%' end GC17" & vbCrLf
        sSql &= " ,case when a.GC18>0 then concat(format(1.0*isnull(c.GC18,0)/a.GC18*100,'0.00'),'%') else '0%' end GC18" & vbCrLf
        sSql &= " ,case when a.GC19>0 then concat(format(1.0*isnull(c.GC19,0)/a.GC19*100,'0.00'),'%') else '0%' end GC19" & vbCrLf
        sSql &= " ,case when a.SUBTOTAL>0 then concat(format(1.0*isnull(c.SUBTOTAL,0)/a.SUBTOTAL*100,'0.00'),'%') else '0%' end SUBTOTAL" & vbCrLf
        sSql &= " FROM WC5 b" & vbCrLf
        sSql &= " LEFT JOIN WC3 c on c.SM1_RANGE=b.SM1_RANGE" & vbCrLf
        sSql &= " LEFT JOIN WC5B bb on bb.SM1_RANGE=b.SM1_RANGE" & vbCrLf
        sSql &= " CROSS JOIN WC4 a" & vbCrLf

        'sSql &= " UNION" & vbCrLf
        'sSql &= " select c.SM1_RANGE,c.SM1_TXT" & vbCrLf
        'sSql &= " ,CONCAT('',c.STDCNT1) STDCNT1" & vbCrLf
        'sSql &= " ,CONCAT('',c.STDCNT1) STDCNT1B" & vbCrLf
        'sSql &= " ,CONCAT('',c.STDCNT2) STDCNT2" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC01) GC01" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC02) GC02" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC03) GC03" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC04) GC04" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC05) GC05" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC06) GC06" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC07) GC07" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC08) GC08" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC09) GC09" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC10) GC10" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC11) GC11" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC12) GC12" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC13) GC13" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC14) GC14" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC15) GC15" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC16) GC16" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC17) GC17" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC18) GC18" & vbCrLf
        'sSql &= " ,CONCAT('',c.GC19) GC19" & vbCrLf
        'sSql &= " ,CONCAT('',c.SUBTOTAL) SUBTOTAL" & vbCrLf
        ''sSql &= " ,c.STDCNT1,c.STDCNT1 STDCNT1B,c.STDCNT2" & vbCrLf
        ''sSql &= " ,c.GC01,c.GC02,c.GC03,c.GC04,c.GC05,c.GC06,c.GC07,c.GC08" & vbCrLf
        ''sSql &= " ,c.GC09,c.GC10,c.GC11,c.GC12,c.GC13,c.GC14,c.GC15,c.GC16" & vbCrLf
        ''sSql &= " ,c.GC17,c.GC18,c.GC19,c.SUBTOTAL" & vbCrLf
        'sSql &= " FROM WC4 c" & vbCrLf

        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        Return dt1
    End Function

    Private Sub EXPORT_2(iCalcMode As Integer)

        Dim sERRMSG As String = ""
        Call CHK_EXPORT_1_VAL(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Exit Sub
        End If
        Dim dtXls As DataTable = sSearch1_DATA_dt2()
        If dtXls Is Nothing OrElse dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim YEAR_N As String = (v_ddlYEARS_SCH1 - 1911)
        If v_ddlYEARS_SCH1 <> v_ddlYEARS_SCH2 Then
            YEAR_N = String.Concat((v_ddlYEARS_SCH1 - 1911), "~", (v_ddlYEARS_SCH2 - 1911))
        End If
        Dim TXT_ORGKIND2 As String = TIMS.GetListText(OrgKind2)
        TXT_ORGKIND2 = If(TXT_ORGKIND2 <> "", String.Concat("(", TXT_ORGKIND2, ")"), "")

        Dim ss_TitleS1 As String = String.Concat((sm.UserInfo.Years - 1911), "年度x產投方案學員參訓情形分析")
        Dim strFilename1 As String = String.Concat(ss_TitleS1, TIMS.GetDateNo2())
        Dim sTitle1 As String = String.Concat("產業人才投資方案", YEAR_N, "年度參訓學員補助費使用級距", TXT_ORGKIND2)

        Dim dt3 As DataTable = GET_GOVCLASSCAST3_TB(objconn)
        Dim sGCODE31PN As String = GET_GOVCLASSCAST3_COMBIN(dt3, "GCODE31PN")
        Dim sGOV3CODE1 As String = GET_GOVCLASSCAST3_COMBIN(dt3, "CODE1")
        'Dim sPattern As String = String.Concat("補助費使用級距,學員人數(ID不重複),學員人數(累計),參訓次數,", sGCODE31PN, ",小計")
        Dim sColumn As String = String.Concat("SM1_TXT,STDCNT1,STDCNT1B,STDCNT2,", sGOV3CODE1, ",SUBTOTAL")

        Dim s_TitleHtml2 As String = ""
        s_TitleHtml2 &= "<tr>"
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "補助費使用級距")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "學員人數(ID不重複)")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "學員人數(累計)")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "參訓次數")
        s_TitleHtml2 &= String.Format("<td colspan=19>{0}</td>", "課程分類(百分比)")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "小計")
        s_TitleHtml2 &= "</tr>"
        s_TitleHtml2 &= "<tr>"
        For Each SPN1 As String In sGCODE31PN.Split(",")
            s_TitleHtml2 &= String.Format("<td>{0}</td>", SPN1)
        Next
        s_TitleHtml2 &= "</tr>"

        'Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        'parms.Add("sPatternA", sPatternA)
        parms.Add("TitleHtml2", s_TitleHtml2)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")

    End Sub

    Protected Sub BTN_EXPORT3_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT3.Click
        EXPORT_3()
    End Sub

    Function sSearch1_DATA_dt3() As DataTable
        Dim V_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim V_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim V_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim sPMS As New Hashtable From {
            {"TPLANID", sm.UserInfo.TPlanID},
            {"YEARS1", V_ddlYEARS_SCH1},
            {"YEARS2", V_ddlYEARS_SCH2}
        }

        Dim sSql As String = ""
        'sSql &= " WITH WC1 AS ( select (SELECT CASE WHEN cct.SUMOFMONEY1 >=70000 THEN 1" & vbCrLf
        'sSql &= " WHEN cct.SUMOFMONEY1 >=65000 THEN 2 WHEN cct.SUMOFMONEY1 >=60000 THEN 3 ELSE 4 END) SM2_RANGE /*滿額計算*/" & vbCrLf

        sSql &= " WITH WC1 AS ( select (SELECT CASE WHEN cct.SUMOFMONEY1 >=100000 THEN 1" & vbCrLf
        sSql &= " WHEN cct.SUMOFMONEY1 >=95000 THEN 2 WHEN cct.SUMOFMONEY1 >=90000 THEN 3 ELSE 4 END) SM2_RANGE /*滿額計算*/" & vbCrLf
        sSql &= " ,ss.YEARS,ss.ORGKINDGW,ss.IDNO,ig3.GCODE31" & vbCrLf
        sSql &= " FROM dbo.V_STUDENTINFO ss" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBSIDYCOST ct WITH(NOLOCK) on ct.SOCID =ss.SOCID" & vbCrLf
        sSql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ss.GCID3=ig3.GCID3" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBCOST cct WITH(NOLOCK) on cct.COSTTYPE=4 and cct.idno=ss.idno and cct.stdate1<=ss.stdate and ss.stdate <=cct.stdate2" & vbCrLf
        sSql &= " WHERE (ct.APPLIEDSTATUSM='Y' OR ct.APPLIEDSTATUS=1)" & vbCrLf
        sSql &= " AND ss.TPLANID=@TPLANID" & vbCrLf
        sSql &= " AND ss.YEARS>=@YEARS1 AND ss.YEARS<=@YEARS2" & vbCrLf
        Select Case V_OrgKind2
            Case "G", "W"
                sPMS.Add("ORGKIND2", V_OrgKind2)
                sSql &= " AND ss.ORGKIND2>=@ORGKIND2" & vbCrLf
        End Select
        sSql &= " )" & vbCrLf
        'sSql &= " AND ss.TPLANID='28' AND ss.YEARS<='2023' AND ss.YEARS>='2020' )" & vbCrLf

        sSql &= " ,WC2 AS ( SELECT MIN(YEARS) MIN_YEARS" & vbCrLf
        sSql &= " ,MAX(YEARS) MAX_YEARS" & vbCrLf
        sSql &= " ,COUNT(DISTINCT IDNO) STDCNT1" & vbCrLf
        sSql &= " ,COUNT(CASE SM2_RANGE WHEN 1 THEN 1 END) SM2_RANGE_1" & vbCrLf
        sSql &= " ,COUNT(CASE SM2_RANGE WHEN 2 THEN 1 END) SM2_RANGE_2" & vbCrLf
        sSql &= " ,COUNT(CASE SM2_RANGE WHEN 3 THEN 1 END) SM2_RANGE_3" & vbCrLf
        sSql &= " FROM WC1 )" & vbCrLf
        sSql &= " select concat(dbo.FN_CYEAR2(a.MIN_YEARS),'-', dbo.FN_CYEAR2(a.MAX_YEARS)) YEARSR" & vbCrLf
        sSql &= " ,a.STDCNT1" & vbCrLf
        sSql &= " ,a.SM2_RANGE_1" & vbCrLf
        sSql &= " ,a.SM2_RANGE_2" & vbCrLf
        sSql &= " ,a.SM2_RANGE_3" & vbCrLf
        sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(a.SM2_RANGE_1,0)/a.STDCNT1*100,'0.00'),'%') else '0%' end SM2_RANGE_1P" & vbCrLf
        sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(a.SM2_RANGE_2,0)/a.STDCNT1*100,'0.00'),'%') else '0%' end SM2_RANGE_2P" & vbCrLf
        sSql &= " ,case when a.STDCNT1>0 then concat(format(1.0*isnull(a.SM2_RANGE_3,0)/a.STDCNT1*100,'0.00'),'%') else '0%' end SM2_RANGE_3P" & vbCrLf
        sSql &= " FROM WC2 a" & vbCrLf

        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        Return dt1
    End Function

    Private Sub EXPORT_3()
        Dim sERRMSG As String = ""
        Call CHK_EXPORT_1_VAL(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Exit Sub
        End If
        Dim dtXls As DataTable = sSearch1_DATA_dt3()
        If dtXls Is Nothing OrElse dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        Dim v_ddlYEARS_SCH1 As String = TIMS.GetListValue(ddlYEARS_SCH1)
        Dim v_ddlYEARS_SCH2 As String = TIMS.GetListValue(ddlYEARS_SCH2)
        Dim YEAR_N As String = (v_ddlYEARS_SCH1 - 1911)
        If v_ddlYEARS_SCH1 <> v_ddlYEARS_SCH2 Then
            YEAR_N = String.Concat((v_ddlYEARS_SCH1 - 1911), "~", (v_ddlYEARS_SCH2 - 1911))
        End If
        Dim TXT_ORGKIND2 As String = TIMS.GetListText(OrgKind2)
        TXT_ORGKIND2 = If(TXT_ORGKIND2 <> "", String.Concat("(", TXT_ORGKIND2, ")"), "")

        Dim ss_TitleS1 As String = "3年10萬額度使用情形"
        Dim strFilename1 As String = String.Concat(ss_TitleS1, "統計總表", TIMS.GetDateNo2())
        Dim sTitle1 As String = String.Concat(ss_TitleS1, TXT_ORGKIND2)

        'Dim dt3 As DataTable = GET_GOVCLASSCAST3_TB(objconn)
        'Dim sGCODE31PN As String = GET_GOVCLASSCAST3_COMBIN(dt3, "GCODE31PN")
        'Dim sGOV3CODE1 As String = GET_GOVCLASSCAST3_COMBIN(dt3, "CODE1")
        'Dim sPattern As String = String.Concat("補助費使用級距,學員人數(ID不重複),學員人數(累計),參訓次數,", sGCODE31PN, ",小計")
        Dim sColumn As String = "YEARSR,STDCNT1,SM2_RANGE_1,SM2_RANGE_1p,SM2_RANGE_2,SM2_RANGE_2p,SM2_RANGE_3,SM2_RANGE_3p"

        Dim s_TitleHtml2 As String = ""
        s_TitleHtml2 &= "<tr>"
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "年度區間")
        s_TitleHtml2 &= String.Format("<td rowspan=2>{0}</td>", "總訓練人數(扣重複)")
        s_TitleHtml2 &= String.Format("<td colspan=6>{0}</td>", "各額度訓練人數(扣重複)")
        s_TitleHtml2 &= "</tr>"
        s_TitleHtml2 &= "<tr>"
        's_TitleHtml2 &= String.Format("<td>{0}</td>", "滿7萬")'s_TitleHtml2 &= String.Format("<td>{0}</td>", "比率")
        's_TitleHtml2 &= String.Format("<td>{0}</td>", "滿6萬5")'s_TitleHtml2 &= String.Format("<td>{0}</td>", "比率")
        's_TitleHtml2 &= String.Format("<td>{0}</td>", "滿6萬")'s_TitleHtml2 &= String.Format("<td>{0}</td>", "比率")
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "滿9萬") : s_TitleHtml2 &= String.Format("<td>{0}</td>", "比率")
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "滿9萬5") : s_TitleHtml2 &= String.Format("<td>{0}</td>", "比率")
        s_TitleHtml2 &= String.Format("<td>{0}</td>", "滿10萬") : s_TitleHtml2 &= String.Format("<td>{0}</td>", "比率")
        s_TitleHtml2 &= "</tr>"

        'Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        'parms.Add("sPatternA", sPatternA)
        parms.Add("TitleHtml2", s_TitleHtml2)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")

    End Sub

End Class

