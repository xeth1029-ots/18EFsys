Public Class SD_15_030
    Inherits AuthBasePage

    '功能路徑： 首頁>>學員動態管理>>統計表>>單位跨年度課程審查結果   (先暫定，它們定的路徑但我還是覺得路徑怪怪....)
    '使用者：署 (系統管理者、承辦人)
    '系統介面
    '(鎖定產投 )
    '訓練機構：必選，因為是要單一單位跨年度，看你覺得應該是要用選單的方式(如圖)，還是做成單位名稱關鍵字 +統編搜尋? (同訓練機構設定功能)
    '年度區間：必選，XXX年~XXX年 (109年起算，因為系統內的審查計分表資料從109年起才有資料)
    '年度的部分我們可以再討論做法，我是覺得不用再選擇上下半年了，反正每年都固定撈取上跟下，沒有資料的就空著
    '可增加檢核：後面欄位不可<前面欄位
    '示意圖如下：
    '邏輯補充：
    '(1) 篩選上下半年(不含政策性階段)
    '(2) 都是審核通過的班級，不管是否停辦
    '按鈕：匯出明細表、匯出總表
    '<明細表>
    '參考附件：產投方案【基隆市美容業職業工會】課程審查結果明細表.xlsx
    '表頭名稱：
    '109~112年【XXXXX (訓練單位名稱)】課程審查結果明細表
    '產出報表範例格式檔，請參附件：
    '【User提供範本】：為現階段業務單位自行人工處理的版本
    '【調整後】：是我有調整版面，讓系統比較方便產出的格式，可先用這版做
    '匯出欄位說明：
    '年度、申請階段
    '等級：以單位在 該年度/階段往前半年的【審查計分區間】資料檔為母體，撈取該單位【複審審核】=通過的【複審等級】
    '舉例說明：110年度上半年XXX單位的等級資料，使用109年7月那份【審查計分區間】資料檔，撈取該單位【複審審核】=通過的【複審等級】
    '理事長：該單位/年度/產投：訓練機構設定中【負責人姓名】欄位資料
    '申請課程名稱：顯示班級名稱，若當年度階段無申請課程，顯示空白或顯示未申請
    '是否通過：是/否
    '訓練人數、訓練費用(元)
    '備註：留空白即可
    '最後一行顯示：總計申請XX班，通過XX班，核定訓練人數：XXX(審核通過班級的訓練人數加總)人，核定訓練費用：$XXXX((審核通過班級的總訓練費用總額加總)
    '<總表>
    '參考附件：產投方案【基隆市美容業職業工會】課程審查結果總表.xlsx
    '表頭名稱：
    '109-112年度產投方案【XXXXX (訓練單位名稱)】申請及核定統計
    '產出報表範例格式檔，請參附件：
    '【User提供範本】：為現階段業務單位自行人工處理的版本
    '【調整後】：是我有調整版面，讓系統比較方便產出的格式，可先用這版做
    '匯出欄位說明：
    '理事長：該單位/年度/產投：訓練機構設定中【負責人姓名】欄位資料
    '年度：一年一格
    '審查計分等級：顯示範例：AB，上半年等級、下半年等級
    '以單位在 該年度/階段往前半年的【審查計分區間】資料檔為母體，撈取該單位【複審審核】=通過的【複審等級】
    '舉例說明：110年上半年等級，使用109年7月那份【審查計分區間】資料檔，撈取該單位【複審審核】=通過的【複審等級】
    '110年下半年等級，使用110年1月那份【審查計分區間】資料檔，撈取該單位【複審審核】=通過的【複審等級】
    '申請
    '班數、訓練人次、補助經費(元)
    '核定
    '班數、訓練人次、補助經費(元)
    '小計
    '上述欄位各年度加總
    '--
    'Best Regards
    '張瑩珊 Sammy Chang 　專案經理
    '東柏資訊科技 Turbo Technologies Co., Ltd
    '臺北市松山區復興北路1號14樓之3
    'Email： sammychang@turbotech.com.tw
    'Tel：02-2776-9993#116
    '附件:
    '產投方案【基隆市美容業職業工會】課程審查結果總表.xlsx	20.1 KB
    '產投方案【基隆市美容業職業工會】課程審查結果明細表.xlsx	21.2 KB

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
            bt_EXPORT1.Enabled = False : TIMS.Tooltip(bt_EXPORT1, TIMS.cst_ErrorMsg17, True)
            bt_EXPORT2.Enabled = False : TIMS.Tooltip(bt_EXPORT2, TIMS.cst_ErrorMsg17, True)
            Common.MessageBox(Me, TIMS.cst_ErrorMsg17)
            Return 'Exit Sub
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Not Me.IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        msg.Text = ""

        Org.Attributes("onclick") = If((sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1), "javascript:openOrg('../../Common/LevOrg.aspx');", "javascript:openOrg('../../Common/LevOrg1.aspx');")

        Dim iSYears As Integer = 2020
        Dim iEYearsNb1 As Integer = Year(Now)
        Dim iEYears As Integer = If(iEYearsNb1 > iSYears, iEYearsNb1, iSYears)
        yearlist1 = TIMS.GetSyear(yearlist1, iSYears, iEYears, False)
        yearlist2 = TIMS.GetSyear(yearlist2, iSYears, iEYears, False)
        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist1, Year(Now) - 2)
        Common.SetListItem(yearlist2, Year(Now))
    End Sub

    '匯出明細表
    Private Sub ExpReport1()
        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If dtXls Is Nothing Then Return
        If dtXls.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dtXls.Rows(0)
        Dim pORGNAME As String = Convert.ToString(dr1("ORGNAME"))
        Dim YEAR_ROC1 As String = TIMS.ClearSQM(Hid_YEAR_ROC1.Value)
        Dim YEAR_ROC2 As String = TIMS.ClearSQM(Hid_YEAR_ROC2.Value)

        Dim sP1 As String = "年度,申請階段,等級,理事長"
        Dim sP2 As String = ",申請課程名稱,是否通過,訓練人數,訓練費用(元)"
        Dim sP3 As String = ",備註"

        'Dim sFOOTER As String = String.Format("總計申請{0}班，通過{1}班，核定訓練人數：{2}人，核定訓練費用：${3}", dr1("CLST1"), dr1("CLSOKT2"), dr1("TNUMS1"), dr1("DEFGOVCOSTS1"))
        Dim sFOOTER_1 As String = String.Format("總計申請{0}班，通過{1}班，申請總訓練人數：{2}人，申請總訓練費用：${3}", dr1("CLST1"), dr1("CLSOKT2"), dr1("TNUMS1"), dr1("DEFGOVCOSTS1"))
        Dim sFOOTER_2 As String = String.Format("，核定訓練人數：{0}人，核定訓練費用：${1}", dr1("TNUMS1Y"), dr1("DEFGOVCOSTS1Y"))
        Dim sFOOTER As String = String.Concat(sFOOTER_1, sFOOTER_2)

        Dim sP1A() As String = Split(sP1, ",")
        Dim sP2A() As String = Split(sP2, ",")
        Dim sP3A() As String = Split(sP3, ",")

        'Dim sPattern As String = "" '序號,
        'Dim sColumn As String = ""
        Dim sPattern As String = String.Concat(sP1, sP2, sP3)
        Dim sColumn As String = "YEARS_ROC,APPSTAGE_N,RLEVEL2,MASTERNAME,CLASSCNAME2,APPLIEDRESULT,TNUM,DEFGOVCOST,MEMO1"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        'Dim sFileName1 As String = String.Concat(YEAR_ROC, "-", YEAR_ROC3, "年度上半年產投方案核定課程統計", v_OrgKind2)
        Dim sFileName1 As String = String.Format("{0}~{1}年【{2}】課程審查結果明細表", YEAR_ROC1, YEAR_ROC2, pORGNAME)
        '套CSS值
        Dim strSTYLE As String = String.Concat("<style>", "td{mso-number-format:""\@"";}", ".noDecFormat{mso-number-format:""0"";}", "</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = "" '建立輸出文字
        '標題抬頭1
        ExportStr = String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFileName1) '& vbTab
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭2
        ExportStr = ""
        For i As Integer = 0 To sP1A.Length - 1
            ExportStr &= String.Format("<td rowspan=2>{0}</td>", sP1A(i)) '& vbTab
        Next
        ExportStr &= String.Format("<td colspan={0}>{1}</td>", 1, "受理申請") '& vbTab
        ExportStr &= String.Format("<td colspan={0}>{1}</td>", 3, "審查結果") '& vbTab
        For i As Integer = 1 To sP3A.Length - 1
            ExportStr &= String.Format("<td rowspan=2>{0}</td>", sP3A(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭3
        ExportStr = ""
        For i As Integer = 1 To sP2A.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sP2A(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '建立資料面
        Dim iRows As Integer = 0
        For Each dr As DataRow In dtXls.Rows
            iRows += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                Dim sCOLTXT As String = Convert.ToString(dr(sColumnA(i)))
                ExportStr &= String.Format("<td>{0}</td>", sCOLTXT)
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next

        '尾部註1
        ExportStr = String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFOOTER) '& vbTab
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        'parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        parmsExp.Add("ExpType", "EXCEL") 'EXCEL/PDF/ODS
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
        Hid_YEAR_ROC1.Value = TIMS.GetListText(yearlist1) '民國年1
        Hid_YEAR_ROC2.Value = TIMS.GetListText(yearlist2) '民國年2
        Dim v_yearlist1 As String = TIMS.GetListValue(yearlist1) '西元年1
        Dim v_yearlist2 As String = TIMS.GetListValue(yearlist2) '西元年2
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        ComidValue.Value = TIMS.ClearSQM(ComidValue.Value)

        Dim sParms As New Hashtable
        sParms.Add("YEARS1", v_yearlist1)
        sParms.Add("YEARS2", v_yearlist2)
        sParms.Add("RID", RIDValue.Value)
        sParms.Add("COMIDNO", ComidValue.Value)
        sParms.Add("TPlanID", sm.UserInfo.TPlanID)

        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " WITH WO1 AS (SELECT r.ORGNAME,r.TPLANID,r.ORGKIND2,r.COMIDNO,r.DISTID FROM dbo.VIEW_RIDNAME r WHERE r.RID=@RID AND r.COMIDNO=@COMIDNO AND r.TPLANID=@TPLANID)" & vbCrLf

        sSql &= " ,WC1 AS (SELECT cc.RID,cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE,cc.FIXSUMCOST,cc.CLASSCNAME2,cc.STDATE" & vbCrLf
        sSql &= " ,o1.ORGNAME,cc.APPLIEDRESULT,cc.TNUM,cc.DEFGOVCOST" & vbCrLf
        sSql &= " FROM WO1 o1" & vbCrLf
        sSql &= " JOIN dbo.VIEW2B cc on cc.COMIDNO=o1.COMIDNO AND cc.DISTID =o1.DISTID AND cc.TPLANID=o1.TPLANID AND cc.ORGKIND2=o1.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sSql &= " WHERE cc.YEARS>=@YEARS1 AND cc.YEARS<=@YEARS2 AND cc.APPSTAGE IN (1,2))" & vbCrLf

        sSql &= " ,WC2 AS (SELECT COUNT(1) CLST1,COUNT(CASE WHEN APPLIEDRESULT='Y' THEN 1 END) CLSOKT2" & vbCrLf
        sSql &= " ,SUM(TNUM) TNUMS1,SUM(DEFGOVCOST) DEFGOVCOSTS1" & vbCrLf
        sSql &= " ,SUM(CASE WHEN APPLIEDRESULT='Y' THEN TNUM END) TNUMS1Y" & vbCrLf
        sSql &= " ,SUM(CASE WHEN APPLIEDRESULT='Y' THEN DEFGOVCOST END) DEFGOVCOSTS1Y FROM WC1)" & vbCrLf

        sSql &= " SELECT cc.ORGNAME" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(cc.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,dbo.FN_GET_APPSTAGE(cc.APPSTAGE) APPSTAGE_N" & vbCrLf
        sSql &= " ,dbo.FN_SCORING2_RLEVEL_2(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE) RLEVEL2" & vbCrLf
        sSql &= " ,(SELECT x.MASTERNAME FROM dbo.VIEW_RIDNAME x WHERE x.RID=cc.RID) MASTERNAME" & vbCrLf
        sSql &= " ,cc.CLASSCNAME2,cc.STDATE" & vbCrLf
        sSql &= " ,CASE WHEN cc.APPLIEDRESULT='Y' THEN '是' ELSE '否' END APPLIEDRESULT" & vbCrLf
        sSql &= " ,cc.TNUM,cc.DEFGOVCOST,'' MEMO1" & vbCrLf
        sSql &= " ,c2.CLST1,c2.CLSOKT2,c2.TNUMS1,c2.DEFGOVCOSTS1,c2.TNUMS1Y,c2.DEFGOVCOSTS1Y"
        sSql &= " FROM WC1 cc" & vbCrLf
        sSql &= " CROSS JOIN WC2 c2"
        sSql &= " ORDER BY cc.YEARS,cc.APPSTAGE,cc.STDATE"

        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sSql, objconn)
        Call DbAccess.HashParmsChange(sCmd, sParms)
        dt.Load(sCmd.ExecuteReader())
        Call TIMS.CHG_dtReadOnly(dt)
        Return dt
    End Function

    Function cCheckData1() As String
        '可增加檢核：後面欄位不可<前面欄位
        Dim ErrMessage1 As String = ""

        If ComidValue.Value = "" AndAlso RIDValue.Value <> "" Then
            Dim drR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
            ComidValue.Value = If(drR IsNot Nothing, Convert.ToString(drR("COMIDNO")), "")
        End If
        If RIDValue.Value = "" OrElse ComidValue.Value = "" Then ErrMessage1 &= "請選擇一訓練機構!" & vbCrLf

        Dim v_yearlist1 As String = TIMS.GetListValue(yearlist1)
        Dim v_yearlist2 As String = TIMS.GetListValue(yearlist2)
        If v_yearlist1 = "" Then ErrMessage1 &= "請選擇起始年度!" & vbCrLf
        If v_yearlist2 = "" Then ErrMessage1 &= "請選擇迄止年度!" & vbCrLf
        If v_yearlist1 <> "" AndAlso v_yearlist1 <> "" AndAlso Val(v_yearlist1) > Val(v_yearlist2) Then ErrMessage1 &= "起始年度 不可大於 迄止年度!" & vbCrLf
        'If ErrMessage1 <> "" Then
        '    Common.MessageBox(Me, ErrMessage1)
        '    Return '    Exit Sub
        'End If

        Return ErrMessage1
    End Function

    '匯出明細表
    Protected Sub bt_EXPORT1_Click(sender As Object, e As EventArgs) Handles bt_EXPORT1.Click
        Dim ErrMsg1 As String = cCheckData1()
        If ErrMsg1 <> "" Then Common.MessageBox(Me, ErrMsg1)
        If ErrMsg1 <> "" Then Return

        '匯出明細表
        Call ExpReport1()
    End Sub

    '匯出總表
    Protected Sub bt_EXPORT2_Click(sender As Object, e As EventArgs) Handles bt_EXPORT2.Click
        Dim ErrMsg1 As String = cCheckData1()
        If ErrMsg1 <> "" Then Common.MessageBox(Me, ErrMsg1)
        If ErrMsg1 <> "" Then Return

        '匯出總表
        Call ExpReport2()
    End Sub

    '匯出總表
    Private Sub ExpReport2()
        Dim dtXls As DataTable = SEARCH_DATA2_dt()
        If dtXls Is Nothing Then Return
        If dtXls.Rows.Count = 0 Then Return

        Dim dr1 As DataRow = dtXls.Rows(0)
        Dim pORGNAME As String = Convert.ToString(dr1("ORGNAME"))
        Dim YEAR_ROC1 As String = TIMS.ClearSQM(Hid_YEAR_ROC1.Value)
        Dim YEAR_ROC2 As String = TIMS.ClearSQM(Hid_YEAR_ROC2.Value)

        Dim sP1 As String = "理事長,年度,審查計分等級"
        Dim sP2 As String = ",班數,訓練人次,補助經費(元),班數,訓練人次,補助經費(元)"
        Dim sP3 As String = ",備註"

        Dim sP1A() As String = Split(sP1, ",")
        Dim sP2A() As String = Split(sP2, ",")
        Dim sP3A() As String = Split(sP3, ",")

        Dim sPattern As String = "" '序號,
        Dim sColumn As String = ""
        sPattern = String.Concat(sP1, sP2, sP3)
        sColumn = "MASTERNAME,YEARS_ROC,RLEVEL2,CLST1,TNUMS1,DEFGOVCOSTS1,CLST2,TNUMS2,DEFGOVCOSTS2,MEMO1"
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        'Dim sFileName1 As String = String.Concat(YEAR_ROC, "-", YEAR_ROC3, "年度上半年產投方案核定課程統計", v_OrgKind2)
        Dim sFileName1 As String = String.Format("{0}~{1}年度產投方案【{2}】申請及核定統計", YEAR_ROC1, YEAR_ROC2, pORGNAME)
        '套CSS值
        Dim strSTYLE As String = String.Concat("<style>", "td{mso-number-format:""\@"";}", ".noDecFormat{mso-number-format:""0"";}", "</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = "" '建立輸出文字
        '標題抬頭1
        ExportStr = String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFileName1) '& vbTab
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭2
        ExportStr = ""
        For i As Integer = 0 To sP1A.Length - 1
            ExportStr &= String.Format("<td rowspan=2>{0}</td>", sP1A(i)) '& vbTab
        Next
        ExportStr &= String.Format("<td colspan={0}>{1}</td>", 3, "申請") '& vbTab
        ExportStr &= String.Format("<td colspan={0}>{1}</td>", 3, "核定") '& vbTab
        For i As Integer = 1 To sP3A.Length - 1
            ExportStr &= String.Format("<td rowspan=2>{0}</td>", sP3A(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭3
        ExportStr = ""
        For i As Integer = 1 To sP2A.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sP2A(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '建立資料面
        Dim CLST1, TNUMS1, DEFGOVCOSTS1, CLST2, TNUMS2, DEFGOVCOSTS2 As Double
        Dim iRows As Integer = 0
        For Each dr As DataRow In dtXls.Rows
            iRows += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                Dim sCOLTXT As String = Convert.ToString(dr(sColumnA(i)))
                Select Case sColumnA(i)
                    Case "CLST1"
                        CLST1 += TIMS.VAL1(sCOLTXT)
                    Case "TNUMS1"
                        TNUMS1 += TIMS.VAL1(sCOLTXT)
                    Case "DEFGOVCOSTS1"
                        DEFGOVCOSTS1 += TIMS.VAL1(sCOLTXT)
                    Case "CLST2"
                        CLST2 += TIMS.VAL1(sCOLTXT)
                    Case "TNUMS2"
                        TNUMS2 += TIMS.VAL1(sCOLTXT)
                    Case "DEFGOVCOSTS2"
                        DEFGOVCOSTS2 += TIMS.VAL1(sCOLTXT)
                End Select
                ExportStr &= String.Format("<td>{0}</td>", sCOLTXT)
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next

        If iRows > 0 Then
            ExportStr = "<tr>"
            ExportStr &= "<td colspan=3>小計</td>"
            ExportStr &= String.Format("<td>{0}</td>", CLST1)
            ExportStr &= String.Format("<td>{0}</td>", TNUMS1)
            ExportStr &= String.Format("<td>{0}</td>", DEFGOVCOSTS1)
            ExportStr &= String.Format("<td>{0}</td>", CLST2)
            ExportStr &= String.Format("<td>{0}</td>", TNUMS2)
            ExportStr &= String.Format("<td>{0}</td>", DEFGOVCOSTS2)

            ExportStr &= String.Format("<td>{0}</td>", "")
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        End If

        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        'parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        parmsExp.Add("ExpType", "EXCEL") 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

    Private Function SEARCH_DATA2_dt() As DataTable
        Hid_YEAR_ROC1.Value = TIMS.GetListText(yearlist1) '民國年1
        Hid_YEAR_ROC2.Value = TIMS.GetListText(yearlist2) '民國年2
        Dim v_yearlist1 As String = TIMS.GetListValue(yearlist1) '西元年1
        Dim v_yearlist2 As String = TIMS.GetListValue(yearlist2) '西元年2
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        ComidValue.Value = TIMS.ClearSQM(ComidValue.Value)

        Dim sParms As New Hashtable
        sParms.Add("YEARS1", v_yearlist1)
        sParms.Add("YEARS2", v_yearlist2)
        sParms.Add("RID", RIDValue.Value)
        sParms.Add("COMIDNO", ComidValue.Value)
        sParms.Add("TPlanID", sm.UserInfo.TPlanID)

        Dim sSql As String = ""
        sSql &= " WITH WO1 AS (SELECT r.ORGNAME,r.TPLANID,r.ORGKIND2,r.COMIDNO,r.DISTID FROM dbo.VIEW_RIDNAME r WHERE r.RID=@RID AND r.COMIDNO=@COMIDNO AND r.TPLANID=@TPLANID)" & vbCrLf

        sSql &= " ,WC1 AS (SELECT cc.RID,cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,cc.APPSTAGE,cc.FIXSUMCOST,cc.CLASSCNAME2,cc.STDATE" & vbCrLf
        sSql &= " ,o1.ORGNAME,cc.APPLIEDRESULT,cc.TNUM,cc.DEFGOVCOST" & vbCrLf
        sSql &= " FROM WO1 o1" & vbCrLf
        sSql &= " JOIN dbo.VIEW2B cc on cc.COMIDNO=o1.COMIDNO AND cc.DISTID=o1.DISTID AND cc.TPLANID=o1.TPLANID AND cc.ORGKIND2=o1.ORGKIND2 COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sSql &= " WHERE cc.YEARS>=@YEARS1 AND cc.YEARS<=@YEARS2 AND cc.APPSTAGE IN (1,2))" & vbCrLf

        sSql &= " ,WC2 AS (SELECT cc.YEARS" & vbCrLf
        sSql &= " ,COUNT(1) CLST1,SUM(TNUM) TNUMS1,SUM(DEFGOVCOST) DEFGOVCOSTS1" & vbCrLf '申請補助經費(元)
        sSql &= " ,COUNT(CASE WHEN APPLIEDRESULT='Y' THEN 1 END) CLST2" & vbCrLf
        sSql &= " ,SUM(CASE WHEN APPLIEDRESULT='Y' THEN TNUM END) TNUMS2" & vbCrLf
        sSql &= " ,SUM(CASE WHEN APPLIEDRESULT='Y' THEN DEFGOVCOST END) DEFGOVCOSTS2" & vbCrLf '核定補助經費(元)
        sSql &= " ,MIN(cc.RID) RID" & vbCrLf
        sSql &= " ,MIN(dbo.FN_SCORING2_RLEVEL_2(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,1)) RLEVEL21" & vbCrLf
        sSql &= " ,MIN(dbo.FN_SCORING2_RLEVEL_2(cc.COMIDNO,cc.TPLANID,cc.DISTID,cc.YEARS,2)) RLEVEL22" & vbCrLf
        sSql &= " FROM WC1 cc" & vbCrLf
        sSql &= " GROUP BY cc.YEARS)" & vbCrLf

        sSql &= " SELECT (SELECT x.MASTERNAME FROM dbo.VIEW_RIDNAME x WHERE x.RID=cc.RID) MASTERNAME" & vbCrLf
        sSql &= " ,(SELECT x.ORGNAME FROM WO1 x) ORGNAME" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(cc.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,concat(RLEVEL21,RLEVEL22) RLEVEL2" & vbCrLf
        sSql &= " ,cc.CLST1,cc.TNUMS1,cc.DEFGOVCOSTS1" & vbCrLf
        sSql &= " ,cc.CLST2,cc.TNUMS2,cc.DEFGOVCOSTS2,'' MEMO1" & vbCrLf
        sSql &= " FROM WC2 cc" & vbCrLf
        sSql &= " ORDER BY cc.YEARS" & vbCrLf

        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(sSql, objconn)
        Call DbAccess.HashParmsChange(sCmd, sParms)
        dt.Load(sCmd.ExecuteReader())
        Call TIMS.CHG_dtReadOnly(dt)
        Return dt
    End Function
End Class