Partial Class SD_15_031
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

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

        cblDISTID = TIMS.Get_DistID(cblDISTID, TIMS.Get_DISTIDT2(objconn))
        cblDISTID.Items.Insert(0, New ListItem("全部", 0))
        Common.SetListItem(cblDISTID, sm.UserInfo.DistID)
        cblDISTID.Enabled = If(sm.UserInfo.LID <> 0, False, True)
        cblDISTIDHidden.Value = "0"
        cblDISTID.Attributes("onclick") = "SelectAll('cblDISTID','cblDISTIDHidden');"

        'ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, TIMS.Get_DISTIDT2(objconn)) 'Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID)
    End Sub

    Function SSearch1_DATA_dt() As DataTable
        'Dim dt1 As DataTable = Nothing
        Dim v_ddlYEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH)
        'Dim v_cblDISTID As String = TIMS.GetCblValue(cblDISTID)
        Dim v_cblDISTIDin As String = TIMS.GetCblValueIn(cblDISTID)
        'Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH)
        STDATE_SCH1.Text = TIMS.Cdate3(STDATE_SCH1.Text)
        STDATE_SCH2.Text = TIMS.Cdate3(STDATE_SCH2.Text)
        Dim v_STDATE_SCH1 As String = TIMS.Cdate3(STDATE_SCH1.Text)
        Dim v_STDATE_SCH2 As String = TIMS.Cdate3(STDATE_SCH2.Text)
        FTDATE_SCH1.Text = TIMS.Cdate3(FTDATE_SCH1.Text)
        FTDATE_SCH2.Text = TIMS.Cdate3(FTDATE_SCH2.Text)
        Dim v_FTDATE_SCH1 As String = TIMS.Cdate3(FTDATE_SCH1.Text)
        Dim v_FTDATE_SCH2 As String = TIMS.Cdate3(FTDATE_SCH2.Text)

        'If v_ddlYEARS_SCH = "" OrElse v_ddlDISTID_SCH = "" Then Return dt1

        Dim sPMS As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}}
        If (v_ddlYEARS_SCH <> "") Then sPMS.Add("YEARS", v_ddlYEARS_SCH)
        'If (v_cblDISTID <> "") Then sPMS.Add("DISTID", v_cblDISTID)
        If (v_STDATE_SCH1 <> "") Then sPMS.Add("STDATE1", TIMS.Cdate2(v_STDATE_SCH1))
        If (v_STDATE_SCH2 <> "") Then sPMS.Add("STDATE2", TIMS.Cdate2(v_STDATE_SCH2))
        If (v_FTDATE_SCH1 <> "") Then sPMS.Add("FTDATE1", TIMS.Cdate2(v_FTDATE_SCH1))
        If (v_FTDATE_SCH2 <> "") Then sPMS.Add("FTDATE2", TIMS.Cdate2(v_FTDATE_SCH2))

        '新住民參訓明細SQL" & vbCrLf
        Dim sSql As String = "
SELECT cc.YEARS,cc.ORGKIND2,cc.ORGPLANNAME2 ,cc.OCID,cc.TOTALCOST
,format(cc.STDATE,'yyyy/MM/dd') STDATE ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE
,cc.ORGNAME,cc.DISTID,cc.DISTNAME,cc.TNUM
,'' CJOBNAME2,cc.CJOB_UNKEY,cc.GCID,cc.GCID2,cc.GCID3 ,ISNULL(g2.GCODE2,g3.GCODE32) GCODE32
,cc.CLASSCNAME,cc.CLASSCNAME2,cc.JOBNAME
,cs.NAME,cs.IDNO,dbo.FN_GET_MASK1(cs.IDNO) IDNO_MK
,cs.PASSPORTNO,cs.Nationality
,cs.SEX2,YEAR(cs.BIRTHDAY)-1911 BIRTHY,cs.AGE
,cs.BUDGETID,cs.BUDGETIDN,sb.SUMOFMONEY,sb.PAYMONEY
,ISNULL(sb.SUMOFMONEY,0)+ISNULL(sb.PAYMONEY,0) PAYMONEY2 
FROM VIEW2 cc
JOIN V_STUDENTINFO cs on cs.OCID=cc.OCID
LEFT JOIN V_GOVCLASSCAST2 g2 ON g2.GCID2=cc.GCID2
LEFT JOIN V_GOVCLASSCAST3 g3 on g3.GCID3=cc.GCID3
LEFT JOIN VIEW_SUBSIDYCOST sb ON sb.SOCID=cs.SOCID
WHERE cc.TPLANID=@TPLANID AND SUBSTRING(cs.IDNO,2,1) NOT IN ('1','2') AND cs.BUDGETID !='97'
"
        If (v_ddlYEARS_SCH <> "") Then sSql &= " AND cc.YEARS=@YEARS" & vbCrLf
        If (v_cblDISTIDin <> "") Then sSql &= $" AND cs.DISTID IN ({v_cblDISTIDin}){vbCrLf}" 'sSql &= " AND cc.DISTID =@DISTID" & vbCrLf
        If (v_STDATE_SCH1 <> "") Then sSql &= " AND cc.STDATE >= @STDATE1" & vbCrLf
        If (v_STDATE_SCH2 <> "") Then sSql &= " AND cc.STDATE <= @STDATE2" & vbCrLf
        If (v_FTDATE_SCH1 <> "") Then sSql &= " AND cc.FTDATE >= @FTDATE1" & vbCrLf
        If (v_FTDATE_SCH2 <> "") Then sSql &= " AND cc.FTDATE <= @FTDATE2" & vbCrLf

        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, $"--#SD_15_031:{vbCrLf}{TIMS.GetMyValue5(sPMS)}{vbCrLf}{vbCrLf}{sSql}")

        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
        If TIMS.dtHaveDATA(dt1) Then
            '通俗職類-table-含命名
            Dim dtSHARECJOB As DataTable = TIMS.Get_SHARECJOBdtV(objconn)
            For Each dr1 As DataRow In dt1.Rows
                dr1("CJOBNAME2") = TIMS.Get_CJOBNAME(dtSHARECJOB, $"{dr1("CJOB_UNKEY")}", 2) 'CJOBNAME2:"通俗職類-小類"
            Next
        End If
        Return dt1
    End Function

    Sub CHK_EXPORT_1_VAL(ByRef sERRMSG As String)
        sERRMSG = ""

        Dim v_ddlYEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH)
        'Dim v_cblDISTIDin As String = TIMS.GetCblValueIn(cblDISTID)
        'Dim v_cblDISTID As String = TIMS.GetCblValue(cblDISTID)
        'Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH)
        STDATE_SCH1.Text = TIMS.Cdate3(STDATE_SCH1.Text)
        STDATE_SCH2.Text = TIMS.Cdate3(STDATE_SCH2.Text)
        Dim v_STDATE_SCH1 As String = TIMS.Cdate3(STDATE_SCH1.Text)
        Dim v_STDATE_SCH2 As String = TIMS.Cdate3(STDATE_SCH2.Text)
        FTDATE_SCH1.Text = TIMS.Cdate3(FTDATE_SCH1.Text)
        FTDATE_SCH2.Text = TIMS.Cdate3(FTDATE_SCH2.Text)
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
        If dtXls Is Nothing OrElse dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If

        Const cst_TitleS1 As String = "新住民參訓明細"
        Dim strFilename1 As String = String.Concat(cst_TitleS1, TIMS.GetDateNo2())
        Dim sTitle1 As String = "新住民參訓明細"

        Dim sPattern As String = ""
        sPattern &= "計劃年度,姓名,身分證統一編號,性別,國籍,出生年,年齡,計畫別,分署別,訓練單位,課程分類,課程名稱,課程代碼,通俗職類-小類,開訓日期,結訓日期"
        sPattern &= ",核定人數,核定補助費用,課程業別代碼,預算來源,核定補助費個人,撥款金額個人"
        Dim sColumn As String = ""
        sColumn &= "YEARS,NAME,IDNO_MK,SEX2,Nationality,BIRTHY,AGE,ORGPLANNAME2,DISTNAME,ORGNAME,JOBNAME,CLASSCNAME2,OCID,CJOBNAME2,STDATE,FTDATE"
        sColumn &= ",TNUM,TOTALCOST,GCODE32,BUDGETIDN,PAYMONEY2,SUMOFMONEY"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim parms As New Hashtable
        parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parms.Add("FileName", strFilename1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub BTN_EXPORT1_Click(sender As Object, e As EventArgs) Handles BTN_EXPORT1.Click
        Call EXPORT_1()
    End Sub
End Class
