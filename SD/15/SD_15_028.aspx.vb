Public Class SD_15_028
    Inherits AuthBasePage

    '主旨:    '【確定性需求】訓練單位辦訓相關報表：4-1 訓練單位辦訓情形
    '從:    '張瑩珊 <sammychang@turbotech.com.tw>    '日期:    '2023/3/14 下午 04:42    '到:    '東柏_AMU(丁正中) <amu_ting@turbotech.com.tw>
    '副本:    '洪健智 <C7300052@wda.gov.tw>, 東柏-健智 <jameshung@turbotech.com.tw>
    '確定性需求    '在職系統：產投計畫    '預計完成日期：2023/04/10
    '需求說明：    '此為需求書內之確定性需求：新增產業人才投資方案之訓練單位辦訓相關報表，於112年4月10日前完成。
    '此封信將說明  4-1：訓練單位辦訓情形  需求內容。
    '===========================================================
    '功能路徑： 首頁>>學員動態管理>>統計表>>訓練單位辦訓情形
    '使用者：署 (系統管理者、承辦人)
    '系統介面    '計畫：鎖定產投 (看你覺得有需要放這欄嗎?)
    '查詢條件：年度    '其他邏輯補充：不分階段，不考慮停辦    '表頭名稱：XXX年度產投方案各類訓練單位辦訓情形
    '產出報表範例檔案請參附件，匯出欄位說明如下：
    'A. 計畫：產投/自主
    'B. 單位屬性：訓練機構設定功能之【機構別】欄位，如圖一
    'C. 訓練家數：有提案申請班級的單位數
    'D. 申請班數
    'E. 核定班數
    'F. 屬性占比：該單位屬性核定班數(E) / 所有核定班數加總 (所有E加總) % 取到小數2位
    'G. 核定比率：該單位屬性核定班數(E) / 該單位屬性申請班數(D)  % 取到小數2位
    'H. 開訓班數
    'I. 開訓比率：該單位屬性開訓班數(H) / 該單位屬性核定班數(E)  % 取到小數2位
    'J. 核定人次
    'K.實際開訓人次
    'L. 實際開訓人次(不含協助)：僅算就安、就保
    'M. 開訓人次比率：該單位屬性實際開訓人次(K) / 該單位屬性核定人次(J)  % 取到小數2位
    'N. 開訓人次比率(不含協助)：該單位屬性實際開訓人次(不含協助)(L) / 該單位屬性核定人次(J)  % 取到小數2位
    'O. 核定補助費
    'P. 總撥款金額
    'Q. 補助費比率：該單位屬性總撥款金額(P) / 該單位屬性核定補助費(O)  % 取到小數2位
    '4.110年度產投方案各類訓練單位辦訓情形.xlsx	14.4 KB
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        TIMS.OpenDbConn(objconn)
        TIMS.Get_TitleLab(objconn, Request("ID"), TitleLab1, TitleLab2)

        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            bt_EXPORT.Enabled = False : TIMS.Tooltip(bt_EXPORT, TIMS.cst_ErrorMsg17, True)
            Common.MessageBox(Me, TIMS.cst_ErrorMsg17)
            Return 'Exit Sub
        End If

        If Not Me.IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        msg.Text = ""

        'Dim where_ff3 As String = "TPLANID IN ('06','28','54','70')"
        Dim where_ff3 As String = "TPLANID='28'"
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, where_ff3, objconn)
        Common.SetListItem(TPlanlist1, sm.UserInfo.TPlanID)

        TPlanlist1.Enabled = If(TPlanlist1.SelectedValue <> "", False, True)
        If (Not TPlanlist1.Enabled) Then TIMS.Tooltip(TPlanlist1, "限定計畫")

        'Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        'TPlanlist1.Enabled = False
        'If flagS1 AndAlso flag_test Then TPlanlist1.Enabled = True

        Dim iSYears As Integer = Year(Now) - 3 '2021
        Dim iEYearsNb1 As Integer = Year(Now)
        Dim iEYears As Integer = If(iEYearsNb1 > iSYears, iEYearsNb1, iSYears)
        yearlist = TIMS.GetSyear(yearlist, iSYears, iEYears, False)
        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist, "")
    End Sub

    Protected Sub bt_EXPORT_Click(sender As Object, e As EventArgs) Handles bt_EXPORT.Click
        Call EXPORT1_XLS()
    End Sub

    Private Sub EXPORT1_XLS()
        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        Call ExpReport1(dtXls)
    End Sub

    Private Sub ExpReport1(dtXls As DataTable)
        'Dim YEAR_ROC As String = TIMS.GetListText(yearlist)
        Dim YEAR_ROC As String = TIMS.ClearSQM(Hid_YEAR_ROC.Value)
        Dim sPattern As String = "" '序號,
        Dim sColumn As String = ""
        sPattern = "計畫,單位屬性,訓練家數,申請班數,核定班數,屬性占比,核定比率,開訓班數,開訓比率,核定人次,實際開訓人次,實際開訓人次(不含協助),開訓人次比率,開訓人次比率(不含協助),核定補助費,總撥款金額,補助費比率"
        sColumn = "PLAN1,OTA,OTN,CLSN1,CLSN2,AP1,AP2,CLSN3,AP3,STDN1,STDN2,STDN3,AP4,AP5,SUB1,SUB2,AP6"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        Dim sFileName1 As String = String.Concat(YEAR_ROC, "年度產投方案各類訓練單位辦訓情形")
        '套CSS值
        Dim strSTYLE As String = String.Concat("<style>", "td{mso-number-format:""\@"";}", ".noDecFormat{mso-number-format:""0"";}", "</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '標題抬頭1
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFileName1) '& vbTab
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))
        '標題抬頭2
        ExportStr = ""
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '建立資料面
        Dim iOTN, iCLSN1, iCLSN2, iAP1, iAP2, iCLSN3, iAP3, iSTDN1, iSTDN2, iSTDN3, iAP4, iAP5, iSUB1, iSUB2, iAP6 As Double
        Dim iNum1 As Integer = 0
        For Each dr As DataRow In dtXls.Rows
            iNum1 += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                Dim sCOLTXT As String = Convert.ToString(dr(sColumnA(i)))
                Select Case sColumnA(i)
                    Case "OTN"
                        iOTN += TIMS.VAL1(dr(sColumnA(i)))
                    Case "CLSN1"
                        iCLSN1 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "CLSN2"
                        iCLSN2 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "CLSN3"
                        iCLSN3 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "AP1"
                        iAP1 += TIMS.VAL1(Convert.ToString(dr(sColumnA(i))).Replace("%", ""))
                    Case "AP2"
                        iAP2 += TIMS.VAL1(Convert.ToString(dr(sColumnA(i))).Replace("%", ""))
                    Case "AP3"
                        iAP3 += TIMS.VAL1(Convert.ToString(dr(sColumnA(i))).Replace("%", ""))
                    Case "STDN1"
                        iSTDN1 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "STDN2"
                        iSTDN2 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "STDN3"
                        iSTDN3 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "SUB1"
                        sCOLTXT = TIMS.VAL2N0(sCOLTXT)
                        iSUB1 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "SUB2"
                        sCOLTXT = TIMS.VAL2N0(sCOLTXT)
                        iSUB2 += TIMS.VAL1(dr(sColumnA(i)))
                    Case "AP4"
                        iAP4 += TIMS.VAL1(Convert.ToString(dr(sColumnA(i))).Replace("%", ""))
                    Case "AP5"
                        iAP5 += TIMS.VAL1(Convert.ToString(dr(sColumnA(i))).Replace("%", ""))
                    Case "AP6"
                        iAP6 += TIMS.VAL1(Convert.ToString(dr(sColumnA(i))).Replace("%", ""))
                End Select
                ExportStr &= String.Format("<td>{0}</td>", sCOLTXT)
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next

        If iNum1 > 0 Then
            ExportStr = "<tr>"
            ExportStr &= "<td colspan=2>總計</td>"
            ExportStr &= String.Format("<td>{0}</td>", iOTN) 'OTN
            ExportStr &= String.Format("<td>{0}</td>", iCLSN1) 'CLSN1
            ExportStr &= String.Format("<td>{0}</td>", iCLSN2) 'CLSN2
            ExportStr &= String.Format("<td>{0}%</td>", iAP1) 'AP1
            ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP2 / iNum1, 2)) 'AP2
            ExportStr &= String.Format("<td>{0}</td>", iCLSN3) 'CLSN3
            ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP3 / iNum1, 2)) 'AP3
            ExportStr &= String.Format("<td>{0}</td>", iSTDN1) 'STDN1
            ExportStr &= String.Format("<td>{0}</td>", iSTDN2) 'STDN2
            ExportStr &= String.Format("<td>{0}</td>", iSTDN3) 'STDN3
            ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP4 / iNum1, 2)) 'AP4
            ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP5 / iNum1, 2)) 'AP5
            ExportStr &= String.Format("<td>{0}</td>", TIMS.VAL2N0(iSUB1)) 'SUB1
            ExportStr &= String.Format("<td>{0}</td>", TIMS.VAL2N0(iSUB2)) 'SUB2
            ExportStr &= String.Format("<td>{0}%</td>", TIMS.ROUND(iAP6 / iNum1, 2)) 'AP6
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        End If

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

    Private Function SEARCH_DATA1_dt() As DataTable
        Hid_YEAR_ROC.Value = TIMS.GetListText(yearlist)
        Dim v_Yearlist As String = TIMS.GetListValue(yearlist)
        Dim sParms As New Hashtable
        sParms.Add("YEARS", v_Yearlist)
        sParms.Add("TPlanID", sm.UserInfo.TPlanID)

        Dim sSql As String = ""
        sSql &= " WITH WC1 AS (SELECT OCID,PLANID,COMIDNO,SEQNO,ORGKIND1,ISSUCCESS,ISAPPRPAPER,NOTOPEN,PVR_ISAPPRPAPER,FIXSUMCOST,DEFGOVCOST,TNUM,STDATE,FTDATE FROM VIEW2B" & vbCrLf
        sSql &= " WHERE YEARS=@YEARS AND TPLANID=@TPlanID)" & vbCrLf
        'sSql &= " WHERE YEARS='2022' AND TPLANID='28' AND DISTID='001')" & vbCrLf
        sSql &= " ,WC2 AS (SELECT COUNT(1) ALLCNT1 FROM WC1 WHERE ISSUCCESS='Y' AND ISAPPRPAPER='Y' AND PVR_ISAPPRPAPER='Y')" & vbCrLf

        'V_STUDENTINFO
        sSql &= " ,WS1 AS (SELECT cc.ORGKIND1,cs.SOCID,cs.BUDGETID" & vbCrLf
        sSql &= " ,case when cs.ISAPPRPAPER='Y' AND cs.CREDITPOINTS IS NOT NULL AND ct.APPLIEDSTATUS = '1' AND cs.STUDSTATUS NOT IN (2,3) AND cc.FTDATE < GETDATE() then ct.SUMOFMONEY end SUMOFMONEY" & vbCrLf
        sSql &= " ,dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cc.STDATE) STUDCNT14B" & vbCrLf
        sSql &= " FROM WC1 cc JOIN V_STUDENTINFO cs on cs.OCID=cc.OCID LEFT JOIN STUD_SUBSIDYCOST ct on ct.SOCID=cs.SOCID)" & vbCrLf

        sSql &= " ,WO1 AS (" & vbCrLf
        sSql &= " SELECT kt.ORGTYPEID1" & vbCrLf
        sSql &= " ,CASE WHEN kt.TYPEID1='1' THEN '產投' ELSE '自主' END PLAN1" & vbCrLf 'PLAN1 計畫
        sSql &= " ,concat(kt.TypeID2,'-',kt.TypeID2Name) OTA" & vbCrLf   'OTA 單位屬性
        sSql &= " ,(SELECT COUNT(DISTINCT COMIDNO) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) OTN" & vbCrLf 'OTN 訓練家數
        sSql &= " ,(SELECT COUNT(1) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) CLSN1" & vbCrLf  'CLSN1 申請班數
        sSql &= " ,(SELECT COUNT(CASE WHEN ISSUCCESS='Y' AND ISAPPRPAPER='Y' AND PVR_ISAPPRPAPER='Y' THEN 1 END) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) CLSN2" & vbCrLf   'CLSN2 核定班數
        sSql &= " ,(SELECT COUNT(CASE WHEN ISSUCCESS='Y' AND NOTOPEN='N' AND  ISAPPRPAPER='Y' AND PVR_ISAPPRPAPER='Y' THEN 1 END) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) CLSN3" & vbCrLf 'CLSN3 開訓班數
        sSql &= " ,(select SUM(CASE WHEN ISSUCCESS='Y' AND ISAPPRPAPER='Y' AND PVR_ISAPPRPAPER='Y' THEN TNUM END) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) STDN1" & vbCrLf  'STDN1 核定人次
        'STDN2 實際開訓人次
        'sSql &= " ,(select COUNT(STUDCNT14B) FROM WS1 WHERE ORGKIND1=kt.ORGTYPEID1) STDN2" & vbCrLf 
        'STDN2 實際開訓人次
        sSql &= " ,(select COUNT(CASE WHEN BUDGETID IN ('02','03','97') THEN STUDCNT14B END) FROM WS1 WHERE ORGKIND1=kt.ORGTYPEID1) STDN2" & vbCrLf
        sSql &= " ,(select COUNT(CASE WHEN BUDGETID !='97' AND BUDGETID IN ('02','03') THEN STUDCNT14B END) FROM WS1 WHERE ORGKIND1=kt.ORGTYPEID1) STDN3" & vbCrLf 'STDN3 實際開訓人次(不含協助)
        'sSql &= " ,(select SUM(CASE WHEN ISSUCCESS='Y' AND ISAPPRPAPER='Y' AND PVR_ISAPPRPAPER='Y' THEN FIXSUMCOST END) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) SUB1" & vbCrLf
        sSql &= " ,(select SUM(CASE WHEN ISSUCCESS='Y' AND ISAPPRPAPER='Y' AND PVR_ISAPPRPAPER='Y' THEN DEFGOVCOST END) FROM WC1 WHERE ORGKIND1=kt.ORGTYPEID1) SUB1" & vbCrLf  'SUB1 核定補助費
        sSql &= " ,(select SUM(SUMOFMONEY) FROM WS1 WHERE ORGKIND1=kt.ORGTYPEID1) SUB2" & vbCrLf 'SUB2  總撥款金額
        sSql &= " ,c2.ALLCNT1" & vbCrLf
        sSql &= " FROM dbo.KEY_ORGTYPE1 kt" & vbCrLf
        sSql &= " CROSS JOIN WC2 c2)" & vbCrLf
        'PLAN1 計畫	單位屬性	 OTN 訓練家數	 CLSN1 申請班數	CLSN2 核定班數	屬性占比	
        '核定比率	開訓班數	開訓比率	核定人次	實際開訓人次	實際開訓人次(不含協助)	
        '開訓人次比率	開訓人次比率(不含協助)	核定補助費	總撥款金額	補助費比率

        sSql &= " SELECT ORGTYPEID1,PLAN1,OTA,OTN,CLSN1,CLSN2" & vbCrLf 'PLAN1 計畫	
        sSql &= " ,CASE WHEN ALLCNT1>0 THEN CONCAT(FORMAT(ROUND(100.0*CLSN2/ALLCNT1,2),'#0.00'),'%') END AP1" & vbCrLf 'AP1 屬性占比
        sSql &= " ,CASE WHEN CLSN1>0 THEN CONCAT(FORMAT(ROUND(100.0*CLSN2/CLSN1,2),'#0.00'),'%') END AP2" & vbCrLf 'AP2 核定比率
        sSql &= " ,CLSN3" & vbCrLf
        sSql &= " ,CASE WHEN CLSN2>0 THEN CONCAT(FORMAT(ROUND(100.0*CLSN3/CLSN2,2),'#0.00'),'%') END AP3" & vbCrLf 'AP3 開訓比率
        sSql &= " ,STDN1,STDN2,STDN3" & vbCrLf
        sSql &= " ,CASE WHEN STDN1>0 THEN CONCAT(FORMAT(ROUND(100.0*STDN2/STDN1,2),'#0.00'),'%') END AP4" & vbCrLf 'AP4  開訓人次比率
        sSql &= " ,CASE WHEN STDN1>0 THEN CONCAT(FORMAT(ROUND(100.0*STDN3/STDN1,2),'#0.00'),'%') END AP5" & vbCrLf 'AP5  開訓人次比率(不含協助)
        sSql &= " ,SUB1" & vbCrLf
        sSql &= " ,SUB2" & vbCrLf
        sSql &= " ,CASE WHEN SUB1>0 THEN CONCAT(FORMAT(ROUND(100.0*isnull(SUB2,0)/SUB1,2),'#0.00'),'%') END AP6" & vbCrLf 'AP6 補助費比率
        'sSql &= " ,CASE WHEN SUB1B>0 THEN CONCAT(FORMAT(ROUND(100.0*isnull(SUB2,0)/SUB1B,2),'#0.00'),'%') END AP6B" & vbCrLf
        sSql &= " FROM WO1" & vbCrLf
        sSql &= " ORDER BY ORGTYPEID1" & vbCrLf
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sSql, objconn)
        Call DbAccess.HashParmsChange(sCmd, sParms)
        dt.Load(sCmd.ExecuteReader())
        Call TIMS.CHG_dtReadOnly(dt)
        Return dt
    End Function
End Class
