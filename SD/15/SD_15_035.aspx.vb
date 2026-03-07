Partial Class SD_15_035
    Inherits AuthBasePage

    '管考類別:STEM,ICT
    Const CST_STEM As String = "STEM"
    Const CST_ICT As String = "ICT"
    'Dim v_RBL_MECTYPE As String = TIMS.GetListValue(RBL_MECTYPE)

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        If Not IsPostBack Then
            CCreate1()
        End If
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    Sub CCreate1()
        Labmsg1.Text = "(依目前登入計畫搜尋)"
        msg1.Text = ""

        ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)
    End Sub

    Function SSearch1_DATA_dt() As DataTable
        STDATE_SCH1.Text = TIMS.Cdate3(STDATE_SCH1.Text)
        STDATE_SCH2.Text = TIMS.Cdate3(STDATE_SCH2.Text)
        FTDATE_SCH1.Text = TIMS.Cdate3(FTDATE_SCH1.Text)
        FTDATE_SCH2.Text = TIMS.Cdate3(FTDATE_SCH2.Text)

        '管考類別:STEM,ICT 'Const CST_STEM As String = "STEM" 'Const CST_ICT As String = "ICT"
        Dim v_RBL_MECTYPE As String = TIMS.GetListValue(RBL_MECTYPE)
        Dim v_ddlYEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH)
        Dim v_STDATE_SCH1 As String = TIMS.Cdate3(STDATE_SCH1.Text)
        Dim v_STDATE_SCH2 As String = TIMS.Cdate3(STDATE_SCH2.Text)
        Dim v_FTDATE_SCH1 As String = TIMS.Cdate3(FTDATE_SCH1.Text)
        Dim v_FTDATE_SCH2 As String = TIMS.Cdate3(FTDATE_SCH2.Text)

        Dim sPMS As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}}
        If (v_ddlYEARS_SCH <> "") Then sPMS.Add("YEARS", v_ddlYEARS_SCH)
        If (v_STDATE_SCH1 <> "") Then sPMS.Add("STDATE1", TIMS.Cdate2(v_STDATE_SCH1))
        If (v_STDATE_SCH2 <> "") Then sPMS.Add("STDATE2", TIMS.Cdate2(v_STDATE_SCH2))
        If (v_FTDATE_SCH1 <> "") Then sPMS.Add("FTDATE1", TIMS.Cdate2(v_FTDATE_SCH1))
        If (v_FTDATE_SCH2 <> "") Then sPMS.Add("FTDATE2", TIMS.Cdate2(v_FTDATE_SCH2))

        Dim SSQL As String = ""
        SSQL &= " WITH WG3 AS ( SELECT ig3.GCID3,ig3.GCODE31,ig3.GCODE32,ig3.GCODE2,ig3.PNAME,ig3.CNAME FROM V_GOVCLASSCAST3 ig3" & vbCrLf
        Select Case v_RBL_MECTYPE
            Case CST_STEM
                SSQL &= " WHERE ig3.GCODE2 IN ('01-15','06-01','07-10','12-01','17-01','20-02','23-01','23-02','23-03') )" & vbCrLf
            Case CST_ICT
                SSQL &= " WHERE ig3.GCODE2 IN ('07-10','11-01','11-02','12-01','17-01','17-02','20-02','26-01','27-02','32-01','32-02') )" & vbCrLf
            Case Else
                SSQL &= " WHERE ig3.GCODE2 IN ('01-15','11-01','11-02','12-01','17-01','17-02','27-02','06-01','07-10','23-01','23-02','23-03','15-01','15-02','16-01','26-01','20-01','20-02','22-02','32-01','32-02') )" & vbCrLf
        End Select
        SSQL &= " ,WC1 AS ( SELECT cc.YEARS,cc.DISTID,cc.DISTNAME,CC.OCID,CC.CLASSCNAME,CC.CLASSCNAME2,cc.TNUM,cc.STDATE,cc.FTDATE" & vbCrLf
        SSQL &= " ,ig3.GCODE2,concat('(',ig3.GCODE31,')',ig3.PNAME) T1N,concat('(',ig3.GCODE32,')',ig3.CNAME) T2N,concat('[',ig3.GCODE2,']',ig3.CNAME) T3N" & vbCrLf
        SSQL &= " FROM dbo.VIEW2 cc" & vbCrLf
        SSQL &= " JOIN WG3 ig3 ON ig3.GCID3=CC.GCID3" & vbCrLf
        SSQL &= " WHERE cc.TPLANID=@TPLANID" & vbCrLf
        'sSql &= " AND cc.STDATE >= convert(date,'2025/01/01') AND cc.STDATE <= convert(date,'2025/07/31')" & vbCrLf
        If (v_ddlYEARS_SCH <> "") Then SSQL &= " AND cc.YEARS=@YEARS" & vbCrLf
        If (v_STDATE_SCH1 <> "") Then SSQL &= " AND cc.STDATE>=@STDATE1" & vbCrLf
        If (v_STDATE_SCH2 <> "") Then SSQL &= " AND cc.STDATE<=@STDATE2" & vbCrLf
        If (v_FTDATE_SCH1 <> "") Then SSQL &= " AND cc.FTDATE>=@FTDATE1" & vbCrLf
        If (v_FTDATE_SCH2 <> "") Then SSQL &= " AND cc.FTDATE<=@FTDATE2" & vbCrLf
        SSQL &= " )" & vbCrLf
        SSQL &= " ,WS1 AS (SELECT cc.OCID,count(1) TOTAL" & vbCrLf
        SSQL &= " ,count(CASE WHEN cs.SEX='M' THEN 1 END) TOTAL_M" & vbCrLf
        SSQL &= " ,count(CASE WHEN cs.SEX='F' THEN 1 END) TOTAL_F" & vbCrLf
        SSQL &= " FROM WC1 cc" & vbCrLf
        SSQL &= " JOIN dbo.V_STUDENTINFO cs on cs.OCID=cc.OCID AND cs.BUDGETID IN ('02','03')" & vbCrLf
        SSQL &= " GROUP BY cc.OCID )" & vbCrLf

        SSQL &= " SELECT cc.YEARS,cc.DISTID,cc.DISTNAME,CC.CLASSCNAME,CC.OCID,cc.T1N,cc.T2N,cc.T3N,cc.TNUM" & vbCrLf
        SSQL &= " ,ISNULL(s1.TOTAL,0) TOTAL" & vbCrLf
        SSQL &= " ,ISNULL(s1.TOTAL_M,0) TOTAL_M" & vbCrLf
        SSQL &= " ,ISNULL(s1.TOTAL_F,0) TOTAL_F" & vbCrLf
        SSQL &= " FROM WC1 cc" & vbCrLf
        SSQL &= " LEFT JOIN WS1 s1 on s1.OCID=cc.OCID" & vbCrLf
        SSQL &= " ORDER BY cc.YEARS,cc.DISTID" & vbCrLf

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, $"--sPMS:{TIMS.GetMyValue5(sPMS)}{vbCrLf},#SD_15_035: {vbCrLf}{SSQL}{vbCrLf}")
        End If

        Return DbAccess.GetDataTable(SSQL, objconn, sPMS)
    End Function

    Sub CHK_EXPORT_1_VAL(ByRef sERRMSG As String)
        sERRMSG = ""

        STDATE_SCH1.Text = TIMS.Cdate3(STDATE_SCH1.Text)
        STDATE_SCH2.Text = TIMS.Cdate3(STDATE_SCH2.Text)
        FTDATE_SCH1.Text = TIMS.Cdate3(FTDATE_SCH1.Text)
        FTDATE_SCH2.Text = TIMS.Cdate3(FTDATE_SCH2.Text)
        Dim v_ddlYEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH)
        Dim v_STDATE_SCH1 As String = TIMS.Cdate3(STDATE_SCH1.Text)
        Dim v_STDATE_SCH2 As String = TIMS.Cdate3(STDATE_SCH2.Text)
        Dim v_FTDATE_SCH1 As String = TIMS.Cdate3(FTDATE_SCH1.Text)
        Dim v_FTDATE_SCH2 As String = TIMS.Cdate3(FTDATE_SCH2.Text)
        'If v_ddlYEARS_SCH = "" AndAlso v_cblDISTID = "" Then sERRMSG &= "年度與轄區分署 至少要有資料!" & vbCrLf
        If v_ddlYEARS_SCH = "" AndAlso v_STDATE_SCH1 = "" AndAlso v_STDATE_SCH2 = "" AndAlso v_FTDATE_SCH1 = "" AndAlso v_FTDATE_SCH2 = "" Then sERRMSG &= "年度、開訓期間、結訓期間 至少擇一填寫!" & vbCrLf
    End Sub

    Sub EXPORT_1()
        Dim sERRMSG As String = ""
        Call CHK_EXPORT_1_VAL(sERRMSG)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Exit Sub
        End If
        Dim dtXls As DataTable = SSearch1_DATA_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        '管考類別:STEM,ICT 'Const CST_STEM As String = "STEM" 'Const CST_ICT As String = "ICT"
        Dim v_RBL_MECTYPE As String = TIMS.GetListValue(RBL_MECTYPE)
        Dim strFilename1 As String = $"參訓STEM及ICT人數{TIMS.GetDateNo2()}"
        Dim sTitle1 As String = "參訓STEM及ICT人數"
        Select Case v_RBL_MECTYPE
            Case CST_STEM
                sTitle1 = "參訓STEM人數"
                strFilename1 = $"{sTitle1}{TIMS.GetDateNo2()}"
            Case CST_ICT
                sTitle1 = "參訓ICT人數"
                strFilename1 = $"{sTitle1}{TIMS.GetDateNo2()}"
        End Select
        Dim sPattern As String = "年度,分署,課程名稱,課程代碼,訓練業別編碼,訓練業別,核定人次,開訓人數,開訓人數男,開訓人數女"
        Dim sColumn As String = "YEARS,DISTNAME,CLASSCNAME,OCID,T1N,T3N,TNUM,TOTAL,TOTAL_M,TOTAL_F"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)},
            {"FileName", strFilename1},
            {"TitleName", TIMS.ClearSQM(sTitle1)},
            {"TitleColSpanCnt", iColSpanCount},
            {"sPatternA", sPatternA},
            {"sColumnA", sColumnA}
        }
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub BTN_EXPORT1_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT1.Click
        Call EXPORT_1()
    End Sub
End Class
