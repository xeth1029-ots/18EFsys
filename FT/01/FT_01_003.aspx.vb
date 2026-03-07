Public Class FT_01_003
    'Inherits System.Web.UI.Page
    Inherits AuthBasePage

    'ADP_FDOWNLOAD
    'Batch\dbt_20210217 'dbt_20210217 '綜合查詢統計表(定版)

    'SELECT TOP 10 * FROM ADP_FDOWNLOAD
    'SELECT TOP 10 * FROM ADP_FDOWNLOAD

    '增修需求(提案單)
    '需求編號：OJT-20051204(產投/充飛)
    '處理等級：普
    '預計完成日期：2月底    (總共會有3種定版數據統計表，我陸續提供給你)
    '提案人：發展署 - 金麗芬
    '開發完成後，請先上版至測試環境，待提案人確認OK後，再上版至正式環境
    '以下需求，有疑問再請提出討論，感謝
    '*以資料庫匯出或實體檔案下載方式，待會再電話討論一下!
    '======================================================
    '系統：在職系統
    '計畫：產投、充飛
    '新功能路徑 ：首頁>>定版數據統計表>>綜合查詢統計表(定版)
    '需求：
    '一、資料內容：
    '使用系統排程每月1日產生綜合查詢統計表。
    '1.每年1-6月，可匯出(下載)兩個檔：
    '當年度    ：撈取當年度開訓日為1/1-12/31之資料
    '前一年度：撈取前一年度開訓日為1/1-12/31之資料

    '2.每年7-12月，可匯出(下載)一個檔：
    '當年度    ：撈取當年度開訓日1/1-12/31之資料

    '二、撈取語法：
    '同「綜合查詢統計表」功能，條件僅下開訓日區間 + 下方條件： (包含已核定/未核定、開班/停辦)
    '產投：撈取非包班。
    '充飛：撈取企業包班、聯合企業包班。
    '二、匯出為excel檔，所需欄位為「綜合查詢統計表」全部欄位，如附件。
    '三、資料匯出：
    '年度：依【年度】欄位選擇要匯出之年度
    '計畫：依登入計畫
    '資料版本：依【資料版本】欄位選擇月份，當選1-6月，可另挑選當年度 or 前一年度版。當選7-12月，前一年度版反灰。
    '匯出檔名(範例)：108年1月1日_產投_綜合查詢統計表_當年度定版數據+日期.xlsx、108年1月1日_產投_綜合查詢統計表_前一年度定版數據+日期.ods   (依計畫分：充飛)
    '資料為定版數據(不因撈取時間不同而變動)。
    '四、權限：此功能僅開放給署，不開放給分署及訓練單位。
    '五、功能介面示意圖如下：

    Dim sCJOB_UNKEY As String = ""
    Dim dtSHARECJOB As DataTable
    Dim dtIdentity As DataTable 'key_identity
    Dim dtZip As DataTable

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        msg.Text = ""
        If Not Me.IsPostBack Then
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        TPlanlist1 = TIMS.Get_TPlan(TPlanlist1,,,, "TPLANID IN ('28','54')", objconn)
        Common.SetListItem(TPlanlist1, sm.UserInfo.TPlanID)

        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        TPlanlist1.Enabled = False
        If flagS1 AndAlso flag_test Then TPlanlist1.Enabled = True

        'Dim sql As String
        'sql = "SELECT DISTID,DISTNAME3 NAME FROM V_DISTRICT ORDER BY DISTID"
        'Dim dtDIST As DataTable = DbAccess.GetDataTable(sql, objconn)
        'ddlDISTID = TIMS.Get_DistID(ddlDISTID, dtDIST)
        'Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)

        Dim iSYears As Integer = (Now.Year - 1)
        Dim iEYears As Integer = (Now.Year - 0)
        'Dim iEYearsNowb1 As Integer = Year(Now) - 0
        'Dim iEYears As Integer = If(iEYearsNowb1 > 2020, iEYearsNowb1, 2020)
        yearlist = TIMS.GetSyear(yearlist, iSYears, iEYears, False)
        'yearlist = TIMS.Get_Years(yearlist, objconn)
        'yearlist.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        Common.SetListItem(yearlist, "")

        monthlist = TIMS.Get_Month2(monthlist, "") '補0
        Common.SetListItem(monthlist, Now.ToString("MM"))
    End Sub

    ''' <summary>
    ''' 取得資料
    ''' </summary>
    ''' <returns></returns>
    Function Get_DATATABLE1(ByRef objdt1 As DataTable) As DataTable
        'Dim objdt1 As DataTable = New DataTable
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        If (v_TPlanlist1 = "") Then Return objdt1
        'Dim v_DISTID As String = TIMS.GetListValue(ddlDISTID)
        'If (v_DISTID = "") Then Return objdt1

        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        If (v_yearlist = "") Then Return objdt1
        Dim v_monthlist As String = TIMS.GetListValue(monthlist)
        If (v_monthlist = "") Then Return objdt1
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER) 'N/B
        If (v_BDATAVER = "") Then Return objdt1
        Dim BY1 As String = DateTime.Now.Year.ToString()
        Dim BY2 As String = (DateTime.Now.Year - 1).ToString()
        Dim v_STDATE1 As String = String.Format("{0}/01/01", BY1)
        Dim v_STDATE2 As String = String.Format("{0}/12/31", BY1)
        If v_BDATAVER = "B" Then
            v_STDATE1 = String.Format("{0}/01/01", BY2)
            v_STDATE2 = String.Format("{0}/12/31", BY2)
        End If

        '綜合查詢
        Dim sql As String = ""
        sql &= " SELECT a.*" & vbCrLf
        sql &= " FROM dbo.MV_STUD003 a" & vbCrLf
        sql &= " WHERE a.TPLANID =@TPLANID" & vbCrLf
        'sql &= " AND a.DISTID =@DISTID" & vbCrLf
        sql &= " AND a.BYEAR=@BYEAR" & vbCrLf
        sql &= " AND a.BMONTH=@BMONTH" & vbCrLf
        sql &= " AND a.STDATE >=@STDATE1" & vbCrLf
        sql &= " AND a.STDATE <=@STDATE2" & vbCrLf
        TIMS.OpenDbConn(objconn)
        Dim sCmd As New SqlCommand(sql, objconn)
        sCmd.CommandTimeout = 300
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = v_TPlanlist1
            '.Parameters.Add("DISTID", SqlDbType.VarChar).Value = v_DISTID
            .Parameters.Add("BYEAR", SqlDbType.VarChar).Value = v_yearlist
            .Parameters.Add("BMONTH", SqlDbType.VarChar).Value = v_monthlist
            .Parameters.Add("STDATE1", SqlDbType.DateTime).Value = TIMS.Cdate2(v_STDATE1)
            .Parameters.Add("STDATE2", SqlDbType.DateTime).Value = TIMS.Cdate2(v_STDATE2)
            objdt1.Load(.ExecuteReader())
        End With

        'sql &= " AND a.BDATAVER=@BDATAVER" & vbCrLf
        'Dim parms As New Hashtable
        'parms.Clear()
        'parms.Add("TPLANID", v_TPlanlist1)
        'parms.Add("BYEAR", v_yearlist)
        'parms.Add("BMONTH", v_monthlist)
        'parms.Add("STDATE1", TIMS.cdate2(v_STDATE1))
        'parms.Add("STDATE2", TIMS.cdate2(v_STDATE2))
        ''Dim s_log1 As String = ""
        ''s_log1 &= String.Format("parms:{0}", TIMS.GetMyValue3(parms)) & vbCrLf
        'TIMS.writeLog_1(Me, "FT_01_003", sql, parms)
        ''Dim objdt1 As DataTable = Nothing
        'objdt1 = DbAccess.GetDataTable(sql, objconn, parms)

        Return objdt1
    End Function

    ''' <summary> 組合 Cst_上課地址及教室 SELECT TOP 10 * FROM MV_STUD003 </summary>
    ''' <param name="dtZip"></param>
    ''' <param name="dr"></param>
    ''' <returns></returns>
    Public Shared Function Get_AddressPlaceName(ByRef dtZip As DataTable, ByVal dr As DataRow) As String
        Dim rst As String = ""
        Dim tmpAddr As String = ""
        Const cst_spTag As String = "、"
        Dim strTag As String = ""
        For i As Integer = 1 To 4
            Select Case i
                Case 1
                    strTag = "s1"
                Case 2
                    strTag = "s2"
                Case 3
                    strTag = "t1"
                Case 4
                    strTag = "t2"
            End Select
            Dim SZIPCODE As String = If(Convert.ToString(dr(strTag & "ZIP6W")) <> "", Convert.ToString(dr(strTag & "ZIP6W")), Convert.ToString(dr(strTag & "ZIPCODE")))
            tmpAddr = TIMS.getZipName6(SZIPCODE, Convert.ToString(dr(strTag & "ADDRESS")), Convert.ToString(dr(strTag & "PLACENAME")), dtZip)
            If tmpAddr <> "" Then
                If rst <> "" Then rst &= cst_spTag
                rst &= tmpAddr
            End If
        Next
        Return rst
    End Function

#Region "NOUSE"
    ''' <summary>匯出</summary>
    Sub Utl_EXPORT1()
        msg.Text = "查無資料!!"
        Dim flag_NG_1 As Boolean = False
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        If (v_TPlanlist1 = "") Then flag_NG_1 = True
        Dim t_TPlanlist1 As String = TIMS.GetListText(TPlanlist1)
        If (t_TPlanlist1 = "") Then flag_NG_1 = True
        'Dim t_DISTID As String = TIMS.GetListText(ddlDISTID)
        'If (t_DISTID = "") Then flag_NG_1 = True

        Dim t_yearlist As String = TIMS.GetListText(yearlist)
        If (t_yearlist = "") Then flag_NG_1 = True
        Dim t_monthlist As String = TIMS.GetListText(monthlist)
        If (t_monthlist = "") Then flag_NG_1 = True
        Dim t_BDATAVER As String = TIMS.GetListText(rbl_BDATAVER)
        If (t_BDATAVER = "") Then flag_NG_1 = True
        If flag_NG_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Dim objtable As DataTable = New DataTable
        objtable = Get_DATATABLE1(objtable)
        If objtable Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return 'Exit Sub
        End If
        If objtable.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return 'Exit Sub
        End If
        Dim s_log1 As String = ""
        s_log1 = String.Format("Total rows={0}, columns={1}", objtable.Rows.Count, objtable.Columns.Count)
        TIMS.LOG.Debug(s_log1)
        msg.Text = ""

        '通俗職類-table-含命名
        dtSHARECJOB = TIMS.Get_SHARECJOBdtV(objconn)
        Dim sql As String = ""
        sql = "SELECT * FROM dbo.VIEW_ZIPNAME WITH(NOLOCK) ORDER BY ZIPCODE"
        dtZip = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT IDENTITYID,NAME FROM dbo.KEY_IDENTITY WITH(NOLOCK) WHERE 1=1"
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            sql &= " AND IDENTITYID IN (" & TIMS.Cst_Identity06_2019_11 & ")" & vbCrLf
        Else
            sql &= " AND IDENTITYID IN (" & TIMS.Cst_Identity28_2019_11 & ")" & vbCrLf
        End If
        sql &= " ORDER BY SORT28"
        dtIdentity = DbAccess.GetDataTable(sql, objconn)

        Dim HMSSTR1x As String = DateTime.Now.ToString("yyyyMMddHHmm")
        '欄位設計中文／對應欄位
        '108年1月1日_產投_綜合查詢統計表_當年度定版數據+日期.xlsx
        Const s_FileNamefmt As String = "{0}年_{1}月_{2}_{3}_綜合查詢統計表_{4}"
        Dim sFileName1 As String = String.Format(s_FileNamefmt, t_yearlist, Val(t_monthlist), t_BDATAVER, t_TPlanlist1, HMSSTR1x)
        sFileName1 = TIMS.ClearSQM(sFileName1)

        'Const s_title1 As String = "計畫年度,計畫名稱,計畫別,分署,課程申請階段,訓練單位,班別名稱,課程辦訓縣市,訓練時數,訓練費用,開訓日期,結訓日期,訓練人數,實際報名人數,訓練單位類別,課程分類(一階),課程分類代碼(二階),課程分類(二階),訓練職能,政府政策性產業,新南向政策,參加課成型態,學員姓名,身分證號碼,性別,年齡,年齡級距,最高學歷,身分別,保險證號,投保薪資級距,投保單位,公司統一編號,服務單位,服務部門,職稱,預算別,補助金額,學員經費審核申請,學員經費撥款狀態,參訓狀態,報名時間,聯絡電話(日),聯絡電話(夜),手機號碼,加強原專長相關技能,培育第二專長或轉換其他行職業所需技能,考取相關檢定或證照,拓展人脈,使用政府提供之訓練費用補助,其他（請說明）,是否為第1次參加產業人才投資方案課程,服務單位員工人數,本署或分署網站,就業服務中心,訓練單位,搜尋網站,報紙,廣播,電視,親友介紹,社群媒體,其他,參加本次課程的主要原因,選擇本訓練單位的主要原因,其他,沒有參加本方案訓練之前，每年參加訓練支出的費用,如果沒有補助訓練費用，你每年願意自費參加訓練課程的金額,您認為本次課程的訓練費用是否合理,結訓後對於工作的規劃,課程內容符合期望,課程難易安排適當,課程總時數適當,課程符合實務需求,課程符合產業發展趨勢,滿意講師的教學態度,滿意講師的教學方法,滿意講師的課程專業度,對於訓練教材感到滿意,訓練教材能夠輔助課程學習,您對於訓練場地感到滿意,您對於訓練設備感到滿意,您認為實作設備的數量適當,您認為實作設備新穎,訓練評量（如：訓後測驗、專題報告、作品展示等）能促進學習效果,您認為在訓練課程中，課程內容能讓您專注,您在完成訓練後，已充份學習訓練課程所教授知識或技能,您在完成訓練後，有學習到新的知識或技能,您對於訓練單位的課程安排與授課情形感到滿意,您對於訓練單位的行政服務感到滿意,您對於產業人才投資方案感到滿意,您認為完成本訓練課程對於目前或未來工作有幫助,若本訓練課程沒有補助，是否會全額自費參訓,其他建議,請問您目前的就業狀況為何,請問您的薪資於結訓後有提升嗎,請問您擔任的職位有變化嗎,請問您對目前工作的滿意度是否有變化,請問您目前工作內容是否與本次參訓課程有相關,請問您是否有繼續參與本方案的意願,結訓後有與下列人員保持聯絡-講師,結訓後有與下列人員保持聯絡-學員,結訓後有與下列人員保持聯絡-工作人員,結訓後有與人員無保持聯絡,對工作能力更有信心,有助於提升工作技能,有助於提升工作效率,能增進我的問題解決能力,能將所學應用到工作上,能將所學應用於日常生活中,是否同意參加訓練對第二專長有幫助,是否同意參加訓練對目前工作表現有幫助,有助於提升我的績效考核,有助於薪資的調升,有助於職位的升遷,有助於獲得證照,有助於發展職涯,有助於強化個人職場競爭力,其他建議_"
        'Const s_data1 As String = "YEARS,PLANNAME,ORGKIND2N,DISTNAME,APPSTAGE,ORGNAME,CLASSCNAME,CTNAME,THOURS,TOTAL,STDATE,FTDATE,TNUM,ENTEROUT,TYPEID2NAME,COURSE01,COURSE02ID,COURSE02,CLASSCATEN,KNAME19,KNAME18,PACKAGETYPEN,STDNAME,IDNO,SEX2,YEARSOLD,YEARSOLDT4N,DEGREENAME,MINAME,ACTNO,JOBSALID,ACTNAME,INTAXNO,UNAME,SERVDEPT,JOBTITLE,BUDGETIDN,SUMOFMONEY,APPLIEDSTATUSM2N,APPLIEDSTATUS2N,STUDSTATUS2,RELENTERDATE,PHONED,PHONEN,CELLPHONE,S11,S12,S13,S14,S15,S16_NOTE,S2,S3,A1_1,A1_2,A1_3,A1_4,A1_5,A1_6,A1_7,A1_8,A1_9,A1_10_NOTE,A2,A3,A3_5_NOTE,A4,A5,A6,A7,B11,B12,B13,B14,B15,B21,B22,B23,B31,B32,B41,B42,B43,B44,B51,B61,B62,B63,B71,B72,B73,B74,C11,C21_NOTE,Q1,Q2,Q3,Q4,Q5,Q8,Q7MR1,Q7MR2,Q7MR3,Q7MR4,Q211,Q212,Q213,Q214,Q215,Q216,Q217,Q218,Q221,Q222,Q223,Q224,Q225,Q226,Q3_NOTE"
        'Dim As_title1() As String = s_title1.Split(",")
        'Dim As_data1() As String = s_data1.Split(",")

        '套CSS值
        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim sbHTML As New System.Text.StringBuilder()
        'Dim strHTML As String = ""
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = ""
        ExportStr &= "<tr>"
        If TIMS.Cst_TPlanID28.IndexOf(v_TPlanlist1) > -1 Then
            ExportStr &= "<td>計畫</td>" & vbTab  '自主/產投(產投計畫別)
        End If
        ExportStr &= "<td>各分署</td>" & vbTab
        ExportStr &= "<td>單位屬性</td>" & vbTab
        ExportStr &= "<td>訓練機構</td>" & vbTab

        'Case Cst_統一編號
        ExportStr &= "<td>統一編號</td>" & vbTab
        'Case Cst_立案縣市
        ExportStr &= "<td>立案縣市</td>" & vbTab

        ExportStr &= "<td>班別名稱</td>" & vbTab
        ExportStr &= "<td>期別</td>" & vbTab
        ExportStr &= "<td>課程代碼</td>" & vbTab
        ExportStr &= "<td>申請階段</td>" & vbTab 'APPSTAGE

        ExportStr &= "<td>開訓日期</td>" & vbTab
        ExportStr &= "<td>結訓日期</td>" & vbTab
        ExportStr &= "<td>課程分類</td>" & vbTab '課程分類

        'If gflag_SHOW_2019_1 Then End If
        '政府政策性產業
        '「5+2」產業創新計畫 5+2產業
        '【台灣AI行動計畫】 KID='08'
        '【數位國家創新經濟發展方案】KID='09'
        '【國家資通安全發展方案】KID='10'
        '【前瞻基礎建設計畫】
        '【新南向政策】KID='19'
        ExportStr &= "<td>5+2產業創新計畫</td>" & vbTab
        ExportStr &= "<td>台灣AI行動計畫</td>" & vbTab
        ExportStr &= "<td>數位國家創新經濟發展方案</td>" & vbTab
        ExportStr &= "<td>國家資通安全發展方案</td>" & vbTab
        ExportStr &= "<td>前瞻基礎建設計畫</td>" & vbTab
        ExportStr &= "<td>新南向政策</td>" & vbTab

        ExportStr &= "<td>新興產業</td>" & vbTab '※六大新興產業			
        ExportStr &= "<td>重點服務業</td>" & vbTab '※十大重點服務業			
        'Select Case hid_ssYears.Value 'sCaseYears
        '    Case cst_y2017
        '    Case Else
        '        ExportStr &= "<td>新興智慧型產業</td>" & vbTab '※四大新興智慧型產業			
        'End Select
        ExportStr &= "<td>新興智慧型產業</td>" & vbTab '※四大新興智慧型產業			

        ExportStr &= "<td>訓練業別編碼</td>" & vbTab
        ExportStr &= "<td>訓練業別</td>" & vbTab
        ExportStr &= "<td>通俗職類-大類</td>" & vbTab
        ExportStr &= "<td>通俗職類-小類</td>" & vbTab
        ExportStr &= "<td>訓練職能</td>" & vbTab 'CCName

        ExportStr &= "<td>學科辦訓地縣市</td>" & vbTab
        ExportStr &= "<td>術科辦訓地縣市</td>" & vbTab
        ExportStr &= "<td>包班總類</td>" & vbTab '包班總類 包班種類

        'Case Cst政府政策性產業_108NOUSE '(108年之後不使用此欄)
        ExportStr &= "<td>政府政策性產業(108年之後不使用此欄)</td>" & vbTab
        'Case Cst新南向政策
        ExportStr &= "<td>新南向政策</td>" & vbTab
        'Case Cst轄區重點產業
        ExportStr &= "<td>轄區重點產業</td>" & vbTab
        'Case Cst生產力40
        ExportStr &= "<td>生產力4.0</td>" & vbTab
        'Case Cst申請人次
        ExportStr &= "<td>申請人次</td>" & vbTab
        'Case Cst申請補助費
        ExportStr &= "<td>申請補助費</td>" & vbTab
        'Case Cst核定人次
        ExportStr &= "<td>核定人次</td>" & vbTab
        'Case Cst核定補助費
        ExportStr &= "<td>核定補助費</td>" & vbTab
        'Case Cst實際開訓人次
        ExportStr &= "<td>實際就保開訓人次</td>" & vbTab
        ExportStr &= "<td>實際就安開訓人次</td>" & vbTab
        ExportStr &= "<td>實際公務開訓人次</td>" & vbTab
        ExportStr &= "<td>實際協助開訓人次</td>" & vbTab
        'Case Cst實際開訓人次加總
        ExportStr &= "<td>實際合計開訓人次</td>" & vbTab
        'Case Cst預估補助費
        ExportStr &= "<td>就保預估補助費金額</td>" & vbTab
        ExportStr &= "<td>就安預估補助費金額</td>" & vbTab
        ExportStr &= "<td>公務預估補助費金額</td>" & vbTab
        ExportStr &= "<td>協助預估補助費金額</td>" & vbTab
        'Case Cst預估補助費加總
        ExportStr &= "<td>合計預估補助費金額</td>" & vbTab
        'Case Cst結訓人次 '合計結訓人次
        ExportStr &= "<td>就保結訓人次</td>" & vbTab
        ExportStr &= "<td>就安結訓人次</td>" & vbTab
        ExportStr &= "<td>公務結訓人次</td>" & vbTab
        ExportStr &= "<td>協助結訓人次</td>" & vbTab
        ExportStr &= "<td>合計結訓人次</td>" & vbTab
        'Case Cst撥款人次 'Key_Identity
        ExportStr &= "<td>就保撥款人次</td>" & vbTab
        ExportStr &= "<td>就安撥款人次</td>" & vbTab
        ExportStr &= "<td>公務撥款人次</td>" & vbTab
        ExportStr &= "<td>協助撥款人次</td>" & vbTab
        'Case Cst各身分別撥款人次
        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>就保一般身分撥款人次</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>就保特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>就保特殊身分總撥款人次</td>" & vbTab

        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>就安一般身分撥款人次</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>就安特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>就安特殊身分總撥款人次</td>" & vbTab

        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>公務一般身分撥款人次</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>公務特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>公務特殊身分總撥款人次</td>" & vbTab

        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>協助一般身分撥款人次</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>協助特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款人次</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>協助特殊身分總撥款人次</td>" & vbTab

        'Case Cst撥款補助費
        ExportStr &= "<td>就保撥款補助費</td>" & vbTab
        ExportStr &= "<td>就安撥款補助費</td>" & vbTab
        'ExportStr &= "<td>公務撥款補助費</td>" & vbTab 'budmoneyall3
        ExportStr &= "<td>協助撥款補助費</td>" & vbTab
        'Case Cst各身分別撥款補助費
        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>就保一般身分撥款補助費</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>就保特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>就保特殊身分總撥款補助費</td>" & vbTab

        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>就安一般身分撥款補助費</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>就安特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>就安特殊身分總撥款補助費</td>" & vbTab

        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>公務一般身分撥款補助費</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>公務特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>公務特殊身分總撥款補助費</td>" & vbTab

        For Each dr1 As DataRow In dtIdentity.Rows
            Select Case Convert.ToString(dr1("IdentityID"))
                Case "01"
                    ExportStr &= "<td>協助一般身分撥款補助費</td>" & vbTab '01
                Case Else
                    ExportStr &= "<td>協助特殊身分(" & Convert.ToString(dr1("NAME")) & ")撥款補助費</td>" & vbTab '07
            End Select
        Next
        ExportStr &= "<td>協助特殊身分總撥款補助費</td>" & vbTab

        'Case Cst不預告訪視次數_實地抽訪
        ExportStr &= "<td>累計不預告實地抽訪次數</td>" & vbTab
        ExportStr &= "<td>實地訪視日期</td>" & vbTab
        'Case Cst不預告訪視次數_電話抽訪
        ExportStr &= "<td>累計不預告電話抽訪次數</td>" & vbTab 'CTALL3
        ExportStr &= "<td>電話訪視日期</td>" & vbTab 'cuAPPLYDATE3
        ExportStr &= "<td>累計不預告電話抽訪次數(實地抽訪未到)</td>" & vbTab 'CTALL4
        ExportStr &= "<td>電話訪視日期(實地抽訪未到)</td>" & vbTab 'cuAPPLYDATE4
        'Case Cst累積訪視異常次數
        ExportStr &= "<td>累計不預告實地抽訪異常次數</td>" & vbTab
        ExportStr &= "<td>累計不預告電話抽訪異常次數</td>" & vbTab
        'Case Cst累計實地訪視異常原因
        ExportStr &= "<td>出席率不佳</td>" & vbTab 'It22b01N
        ExportStr &= "<td>簽到退未落實</td>" & vbTab 'It22b02N
        ExportStr &= "<td>師資不符</td>" & vbTab 'It22b03N
        ExportStr &= "<td>助教不符</td>" & vbTab 'It22b06N
        ExportStr &= "<td>課程內容不符</td>" & vbTab 'It22b04N
        ExportStr &= "<td>上課地點不符</td>" & vbTab 'It22b05N
        ExportStr &= "<td>其他</td>" & vbTab 'IT22B99NOTE
        ExportStr &= "<td>其他補充說明</td>" & vbTab 'LITEM23NOTE
        'Case Cst會計查帳次數
        ExportStr &= "<td>會計查帳次數</td>" & vbTab
        'Case Cst離訓人次
        ExportStr &= "<td>離訓人次</td>" & vbTab
        'Case Cst退訓人次
        ExportStr &= "<td>退訓人次</td>" & vbTab
        'Case Cst訓練時數
        ExportStr &= "<td>訓練時數</td>" & vbTab

        'Case Cst固定費用總額
        ExportStr &= "<td>固定費用總額</td>" & vbTab
        'Case Cst固定費用單一人時成本
        ExportStr &= "<td>固定費用單一人時成本</td>" & vbTab
        'Case Cst人時成本超出原因說明
        ExportStr &= "<td>人時成本超出原因說明</td>" & vbTab
        'Case Cst材料費總額
        ExportStr &= "<td>材料費總額</td>" & vbTab
        'Case Cst材料費占比
        ExportStr &= "<td>材料費占比</td>" & vbTab
        'Case Cst超出材料費比率上限原因說明
        ExportStr &= "<td>超出材料費比率上限原因說明</td>" & vbTab
        ''Case Cst人時成本
        'ExportStr &= "<td>人時成本</td>" & vbTab
        'Case Cst上課時間
        ExportStr &= "<td>上課時間</td>" & vbTab
        'Case Cst撥款日期
        ExportStr &= "<td>撥款日期</td>" & vbTab

        'Case Cst包班事業單位
        ExportStr &= "<td>包班事業單位</td>" & vbTab
        'Case Cst師資名單
        ExportStr &= "<td>師資名單</td>" & vbTab
        'Case Cst上課地址及教室
        ExportStr &= "<td>上課地址及教室</td>" & vbTab
        ''Case Cst上課地址及教室2
        '    ExportStr &= "<td>上課地址及教室2</td>" & vbTab
        'Case Cst包班事業單位保險證號
        ExportStr &= "<td>包班事業單位保險證號</td>" & vbTab
        'Case Cst包班事業單位統一編號
        ExportStr &= "<td>包班事業單位統一編號</td>" & vbTab
        'Case Cst協助性別人數
        ExportStr &= "<td>協助男性人數</td>" & vbTab
        ExportStr &= "<td>協助女性人數</td>" & vbTab
        'Case Cst課程申請流水號
        ExportStr &= "<td>課程申請流水號</td>" & vbTab
        'Case Cst上架日期
        ExportStr &= "<td>上架日期</td>" & vbTab
        'Case Cst開放報名日期
        ExportStr &= "<td>開放報名日期</td>" & vbTab
        'Case Cst課程備註
        ExportStr &= "<td>課程備註1</td>" & vbTab
        ExportStr &= "<td>課程備註2</td>" & vbTab
        'Case Cst術科時數
        ExportStr &= "<td>術科時數</td>" & vbTab
        'Case Cst聯絡人
        ExportStr &= "<td>聯絡人</td>" & vbTab
        'Case Cst聯絡電話
        ExportStr &= "<td>聯絡電話</td>" & vbTab
        'Case Cst是否停辦
        ExportStr &= "<td>是否停辦</td>" & vbTab
        'Case CstiCAP標章證號
        ExportStr &= "<td>iCAP標章證號</td>" & vbTab
        ''Case Cst政策性產業課程可辦理班數
        '    ExportStr &= "<td>" & T_YR1 & "</td>" & vbTab
        '    ExportStr &= "<td>" & T_YR2 & "</td>" & vbTab
        '    ExportStr &= "<td>" & T_YR3 & "</td>" & vbTab
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr) '(TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        ExportStr = ""
        For Each dr As DataRow In objtable.Rows
            ExportStr = "<tr>"
            If TIMS.Cst_TPlanID28.IndexOf(v_TPlanlist1) > -1 Then
                ExportStr &= "<td>" & dr("OrgPlanName2") & "</td>" & vbTab  '自主/產投(產投計畫別)
            End If

            ExportStr &= "<td>" & dr("DistName") & "</td>" & vbTab  '分署
            ExportStr &= "<td>" & Convert.ToString(dr("OrgTypeName")) & "</td>" & vbTab    '單位屬性
            ExportStr &= "<td>" & dr("orgname") & "</td>" & vbTab  '訓練機構

            'Case Cst_統一編號
            ExportStr &= "<td>" & Convert.ToString(dr("ComIDNO")) & "</td>" & vbTab
            'Case Cst_立案縣市
            ExportStr &= "<td>" & Convert.ToString(dr("CTName2")) & "</td>" & vbTab

            ExportStr &= "<td>" & dr("ClassName") & "</td>" & vbTab  '班別名稱
            ExportStr &= "<td>" & dr("CyclType") & "</td>" & vbTab  '期別
            ExportStr &= "<td>" & dr("ClassID") & "</td>" & vbTab  '課程代碼
            'APPSTAGE
            ExportStr &= "<td>" & Convert.ToString(dr("APPSTAGE")) & "</td>" & vbTab '申請階段

            ExportStr &= "<td>" & dr("STDate") & "</td>" & vbTab  '開訓日期
            ExportStr &= "<td>" & dr("FDDate") & "</td>" & vbTab  '結訓日期
            ExportStr &= "<td>" & dr("Pkname12") & "</td>" & vbTab  '課程分類

            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME1")) & "</td>" & vbTab '5+2產業創新計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME2")) & "</td>" & vbTab '台灣AI行動計畫
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME3")) & "</td>" & vbTab '"數位國家創新經濟發展方案</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME4")) & "</td>" & vbTab '"國家資通安全發展方案</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME5")) & "</td>" & vbTab '"前瞻基礎建設計畫</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D20KNAME6")) & "</td>" & vbTab '"新南向政策</td>" & vbTab

            '20100114 add andy
            ExportStr &= "<td>" & dr("kname1") & "</td>" & vbTab '※六大新興產業	
            ExportStr &= "<td>" & dr("kname2") & "</td>" & vbTab '※十大重點服務業

            ExportStr &= "<td>" & dr("kname3") & "</td>" & vbTab  '※四大新興智慧型產業

            ExportStr &= "<td>" & dr("GCodeName") & "</td>" & vbTab   '訓練業別編碼
            ExportStr &= "<td>" & Convert.ToString(dr("GCNAME")) & "</td>" & vbTab  '訓練業別
            'ExportStr &= "<td>" & Convert.ToString(dr("TJOBNAME")) & "</td>" & vbTab  '訓練業別(職訓業別)

            sCJOB_UNKEY = Convert.ToString(dr("CJOB_UNKEY"))
            ExportStr &= "<td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 1) & "</td>" & vbTab  '通俗職類-大類
            ExportStr &= "<td>" & TIMS.Get_CJOBNAME(dtSHARECJOB, sCJOB_UNKEY, 2) & "</td>" & vbTab  '通俗職類-小類
            ExportStr &= "<td>" & Convert.ToString(dr("CCName")) & "</td>" & vbTab   '訓練職能

            ExportStr &= "<td>" & dr("AddressSciPTID") & "</td>" & vbTab  '學科辦訓地縣市
            ExportStr &= "<td>" & dr("AddressTechPTID") & "</td>" & vbTab '術科辦訓地縣市
            'ExportStr &= "<td>" & dr("PackageType") & "</td>" & vbTab  '包班種類
            ExportStr &= "<td>" & dr("PackageTypeN") & "</td>" & vbTab  '包班總類 包班種類

            'Case Cst_政府政策性產業_108NOUSE '(108年之後不使用此欄)
            ExportStr &= "<td>" & dr("KNAME19") & "</td>" & vbTab  '政府政策性產業(108年之後不使用此欄)
            'Case Cst_新南向政策
            ExportStr &= "<td>" & dr("KNAME18") & "</td>" & vbTab  '新南向政策
            'Case Cst_轄區重點產業
            '空白或新年度，使用新欄位
            'Dim s_KNAME1315 As String = Convert.ToString(dr("KNAME13"))
            Dim s_KNAME1315 As String = Convert.ToString(dr("KNAME15"))
            ExportStr &= "<td>" & s_KNAME1315 & "</td>" & vbTab  '轄區重點產業
            'Case Cst_生產力40
            ExportStr &= "<td>" & dr("KNAME14") & "</td>" & vbTab  '生產力4.0
            'Case Cst_申請人次
            ExportStr &= "<td class=""noDecFormat"">" & dr("ATNum") & "</td>" & vbTab  '申請人數
            'Case Cst_申請補助費
            ExportStr &= "<td class=""noDecFormat"">" & dr("ADefGovCost") & "</td>" & vbTab  '申請補助費
            'Case Cst_核定人次
            ExportStr &= "<td class=""noDecFormat"">" & dr("TNum") & "</td>" & vbTab  '核定人數
            'Case Cst_核定補助費
            ExportStr &= "<td class=""noDecFormat"">" & dr("DefGovCost") & "</td>" & vbTab '核定補助費
            'Case Cst_實際開訓人次
            ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount1") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount2") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount3") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcount97") & "</td>" & vbTab

            'Case Cst_實際開訓人次加總 'openstudcountall
            ExportStr &= "<td class=""noDecFormat"">" & dr("openstudcountall") & "</td>" & vbTab  '開訓人次加總
            'Case Cst_預估補助費
            ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost1"), 2) & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost2"), 2) & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost3"), 2) & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("cost97"), 2) & "</td>" & vbTab
            'Case Cst_預估補助費加總
            ExportStr &= "<td class=""noDecFormat"">" & TIMS.ROUND(dr("costAll"), 2) & "</td>" & vbTab '預估補助費加總
            'Case Cst_結訓人次 '合計結訓人次
            ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout03") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout02") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout01") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcout97") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("closestudcoutall") & "</td>" & vbTab

            'Case Cst_撥款人次
            ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall2") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall3") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("budcountall97") & "</td>" & vbTab
            'Case Cst_各身分別撥款人次
            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud03count" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount01") & "</td>" & vbTab

            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud02count" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount02") & "</td>" & vbTab

            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud01count" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount03") & "</td>" & vbTab

            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud97count" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudcount97") & "</td>" & vbTab

            'Case Cst_撥款補助費
            ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall2") & "</td>" & vbTab
            'ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall3") & "</td>" & vbTab
            ExportStr &= "<td class=""noDecFormat"">" & dr("budmoneyall97") & "</td>" & vbTab
            'Case Cst_各身分別撥款補助費
            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud03money" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney01") & "</td>" & vbTab

            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud02money" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney02") & "</td>" & vbTab

            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud01money" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney03") & "</td>" & vbTab

            For Each dr1 As DataRow In dtIdentity.Rows
                ExportStr &= "<td class=""noDecFormat"">" & dr("bud97money" & dr1("IdentityID")) & "</td>" & vbTab
            Next
            ExportStr &= "<td class=""noDecFormat"">" & dr("Sbudmoney97") & "</td>" & vbTab

            'Case Cst_不預告訪視次數_實地抽訪
            ExportStr &= "<td class=""noDecFormat"">" & dr("cuall") & "</td>" & vbTab '不預告訪視次數-實地抽訪
            Dim ItAPPLYDATE As String = Convert.ToString(dr("ItAPPLYDATE"))
            If ItAPPLYDATE <> "" Then ItAPPLYDATE = Replace(ItAPPLYDATE, ",", ";")
            ExportStr &= "<td>" & ItAPPLYDATE & "</td>" & vbTab '訪視日期

            'Case Cst_不預告訪視次數_電話抽訪
            '僅計算電話抽訪原因為非「實地抽訪時未到」的件數
            ExportStr &= "<td class=""noDecFormat"">" & dr("CTALL3") & "</td>" & vbTab '不預告訪視次數-電話抽訪
            Dim cuAPPLYDATE3 As String = Convert.ToString(dr("cuAPPLYDATE3"))
            If cuAPPLYDATE3 <> "" Then cuAPPLYDATE3 = Replace(cuAPPLYDATE3, ",", ";")
            ExportStr &= "<td>" & cuAPPLYDATE3 & "</td>" & vbTab '訪視日期

            '僅計算電話抽訪原因=「實地抽訪時未到」的件數
            ExportStr &= "<td class=""noDecFormat"">" & dr("CTALL4") & "</td>" & vbTab '不預告訪視次數-電話抽訪
            Dim cuAPPLYDATE4 As String = Convert.ToString(dr("cuAPPLYDATE4"))
            If cuAPPLYDATE4 <> "" Then cuAPPLYDATE4 = Replace(cuAPPLYDATE4, ",", ";")
            ExportStr &= "<td>" & cuAPPLYDATE4 & "</td>" & vbTab '訪視日期
            'Case Cst_累積訪視異常次數
            'ExportStr &= "<td class=""noDecFormat"">" & dr("vtn") & "</td>" & vbTab    '累計訪視異常次數
            ExportStr &= "<td class=""noDecFormat"">" & dr("vitN") & "</td>" & vbTab    '累計訪視異常次數/累計不預告實地抽訪異常次數
            ExportStr &= "<td class=""noDecFormat"">" & dr("VitTelN") & "</td>" & vbTab    '累計訪視異常次數/累計不預告電話抽訪異常次數
            'Case Cst_累計實地訪視異常原因
            ExportStr &= "<td class=""noDecFormat"">" & dr("It22b01N") & "</td>" & vbTab    '累計實地訪視異常原因-出席率不佳
            ExportStr &= "<td class=""noDecFormat"">" & dr("It22b02N") & "</td>" & vbTab    '累計實地訪視異常原因-簽到退未落實
            ExportStr &= "<td class=""noDecFormat"">" & dr("It22b03N") & "</td>" & vbTab    '累計實地訪視異常原因-師資不符
            ExportStr &= "<td class=""noDecFormat"">" & dr("It22b06N") & "</td>" & vbTab    '累計實地訪視異常原因-助教不符
            ExportStr &= "<td class=""noDecFormat"">" & dr("It22b04N") & "</td>" & vbTab    '累計實地訪視異常原因-課程內容不符
            ExportStr &= "<td class=""noDecFormat"">" & dr("It22b05N") & "</td>" & vbTab    '累計實地訪視異常原因-上課地點不符
            ExportStr &= "<td>" & Convert.ToString(dr("IT22B99NOTE")) & "</td>" & vbTab    '累計實地訪視異常原因-其他
            ExportStr &= "<td>" & Convert.ToString(dr("LITEM23NOTE")) & "</td>" & vbTab    '累計實地訪視異常原因-其他補充說明
            'Case Cst_會計查帳次數
            ExportStr &= "<td class=""noDecFormat"">" & "" & "</td>" & vbTab '會計查帳次數
            'Case Cst_離訓人次
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("std_cnt2")) & "</td>" & vbTab   '離訓人次
            'Case Cst_退訓人次
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("std_cnt3")) & "</td>" & vbTab   '退訓人次
            'Case Cst_訓練時數
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("THours")) & "</td>" & vbTab   '訓練時數
            'Case Cst_固定費用總額
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("FIXSUMCOST")) & "</td>" & vbTab
            'Case Cst_固定費用單一人時成本
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("ACTHUMCOST")) & "</td>" & vbTab
            'Select Case hid_ssYears.Value 'sCaseYears
            '    Case Is >= cst_y2018
            '        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("ACTHUMCOST")) & "</td>" & vbTab
            '    Case Else
            '        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("PHCOST")) & "</td>" & vbTab
            'End Select

            'Case Cst_人時成本超出原因說明
            ExportStr &= "<td>" & Convert.ToString(dr("FIXExceeDesc")) & "</td>" & vbTab
            'Case Cst_材料費總額
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("METSUMCOST")) & "</td>" & vbTab
            'Case Cst_材料費占比
            ExportStr &= "<td>" & Convert.ToString(dr("METCOSTPER")) & "</td>" & vbTab
            'Case Cst_超出材料費比率上限原因說明
            ExportStr &= "<td>" & Convert.ToString(dr("METExceeDesc")) & "</td>" & vbTab
            ''Case Cst_人時成本
            'ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("PhCost")) & "</td>" & vbTab   '人時成本
            'Case Cst_上課時間
            ExportStr &= "<td>" & Convert.ToString(dr("WEEKSTIME")) & "</td>" & vbTab   '上課時間
            'Case Cst_撥款日期
            ExportStr &= "<td>" & Convert.ToString(dr("AllotDate")) & "</td>" & vbTab

            'Case Cst_包班事業單位
            ExportStr &= "<td>" & Convert.ToString(dr("BusPackage")) & "</td>" & vbTab
            'Case Cst_師資名單
            ExportStr &= "<td>" & Convert.ToString(dr("PlanTeacher")) & "</td>" & vbTab
            'Case Cst_上課地址及教室
            Dim spAddress As String = "" '組合地址用
            '組合 Cst_上課地址及教室
            spAddress = Get_AddressPlaceName(dtZip, dr)
            ExportStr &= "<td>" & spAddress & "</td>" & vbTab
            'Case Cst_包班事業單位保險證號
            ExportStr &= "<td>" & Convert.ToString(dr("BusPackage2")) & "</td>" & vbTab
            'Case Cst_包班事業單位統一編號
            ExportStr &= "<td>" & Convert.ToString(dr("BusPackage3")) & "</td>" & vbTab
            'Case Cst_協助性別人數
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("SexNumxM")) & "</td>" & vbTab  '/*男性人數*/
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("SexNumxF")) & "</td>" & vbTab  '/*女性人數*/
            'Case Cst_課程申請流水號
            ExportStr &= "<td>" & Convert.ToString(dr("PSNO28")) & "</td>" & vbTab
            'Case Cst_上架日期
            ExportStr &= "<td>" & Convert.ToString(dr("ONSHELLDATE")) & "</td>" & vbTab
            'Case Cst_開放報名日期
            ExportStr &= "<td>" & Convert.ToString(dr("SENTERDATE")) & "</td>" & vbTab
            'Case Cst_課程備註
            ExportStr &= "<td>" & Convert.ToString(dr("memo8")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("memo82")) & "</td>" & vbTab
            'Case Cst_術科時數 'ProTechHours
            ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr("ProTechHours")) & "</td>" & vbTab
            'Case Cst_聯絡人
            ExportStr &= "<td>" & Convert.ToString(dr("ContactName")) & "</td>" & vbTab
            'Case Cst_聯絡電話
            ExportStr &= "<td>" & Convert.ToString(dr("ContactPhone")) & "</td>" & vbTab
            'Case Cst_是否停辦
            ExportStr &= "<td>" & Convert.ToString(dr("NotOpenN")) & "</td>" & vbTab
            'Case Cst_iCAP標章證號 'ICAPNUM-iCAP標章證號
            ExportStr &= "<td>" & Convert.ToString(dr("ICAPNUM")) & "</td>" & vbTab
            ''Case Cst_政策性產業課程可辦理班數
            '    ExportStr &= "<td>" & Convert.ToString(dr("PCNT11")) & "</td>" & vbTab
            '    ExportStr &= "<td>" & Convert.ToString(dr("PCNT12")) & "</td>" & vbTab
            '    ExportStr &= "<td>" & Convert.ToString(dr("PCNT13")) & "</td>" & vbTab
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")
        objtable = Nothing

        Const cst_ExpType_XLSX As String = "XLSX"
        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", cst_ExpType_XLSX) ' TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("strPROG", "FT_01_003")
        parmsExp.Add("saveFile", "Y")
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

#End Region

    '檢核資料表-取得匯入匯出檔名
    Function GET_FDOWNLOAD(ByRef s_attachmn As String) As String
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_monthlist As String = TIMS.GetListValue(monthlist)
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)

        'Dim s_attachmn As String = ""
        Dim s_filename As String = ""

        Const cst_FTNUM_003 As String = "003"
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("FTNUM", cst_FTNUM_003)
        parms.Add("TPLANID", v_TPlanlist1)
        parms.Add("BYEAR", v_yearlist)
        parms.Add("BMONTH", v_monthlist)
        parms.Add("BDATAVER", v_BDATAVER)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT FDID " & vbCrLf '/*PK*/
        sql &= " ,FNAME1 ,FNAMEIN ,FNAMEOUT" & vbCrLf
        sql &= " ,FTNUM ,EXPDATE ,TPLANID" & vbCrLf
        sql &= " ,BYEAR ,BMONTH ,BDATAVER" & vbCrLf
        sql &= " ,FSN,FSTATUS,FCOUNT" & vbCrLf
        'sql &= " ,MODIFYACCT,MODIFYDATE" & vbCrLf
        sql &= " FROM ADP_FDOWNLOAD" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND FTNUM=@FTNUM " & vbCrLf
        sql &= " AND TPLANID=@TPLANID " & vbCrLf
        sql &= " AND BYEAR=@BYEAR " & vbCrLf
        sql &= " AND BMONTH=@BMONTH " & vbCrLf
        sql &= " AND BDATAVER=@BDATAVER " & vbCrLf
        sql &= " AND FSTATUS IS NULL" & vbCrLf
        sql &= " ORDER BY MODIFYDATE DESC " & vbCrLf
        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        If dt.Rows.Count = 0 Then Return s_filename
        Dim dr1 As DataRow = dt.Rows(0)
        s_attachmn = Convert.ToString(dr1("FNAMEOUT"))
        s_filename = Convert.ToString(dr1("FNAMEIN"))
        Return s_filename
    End Function

    '下載匯出檔-取得匯入匯出檔名
    Function GET_FILENAME1(ByRef s_attachmn As String) As String
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_monthlist As String = TIMS.GetListValue(monthlist)
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)

        'Dim s_attachmn As String = ""
        Dim s_filename As String = ""
        s_filename = GET_FDOWNLOAD(s_attachmn)
        '若有查詢到值，下面可略過
        If s_filename <> "" AndAlso s_attachmn <> "" Then Return s_filename

        Dim s_PLANNAME2 As String = ""
        Select Case v_TPlanlist1
            Case "28"
                s_PLANNAME2 = "產投"
            Case "54"
                s_PLANNAME2 = "充電起飛"
        End Select
        Dim s_BDATAVER2 As String = ""
        Select Case v_BDATAVER
            Case "N"
                s_BDATAVER2 = "當年度"
            Case "B"
                s_BDATAVER2 = "前一年度"
        End Select
        Dim s_ROCYEAR As String = CStr(Val(v_yearlist) - 1911)
        Dim fileN1 As String = ""
        fileN1 = String.Format("{0}_{1}_{2}_{3}_綜合查詢統計表", s_ROCYEAR, v_monthlist, s_PLANNAME2, s_BDATAVER2)

        Select Case v_yearlist
            Case "2021"
                Select Case v_monthlist
                    Case "04"
                        Select Case v_TPlanlist1
                            Case "28" 's_PLANNAME2 = "產投"
                                Select Case v_BDATAVER
                                    Case "N" 's_BDATAVER2 = "當年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202104010312")
                                    Case "B" 's_BDATAVER2 = "前一年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202104010313")
                                End Select
                            Case "54" 's_PLANNAME2 = "充電起飛"
                                Select Case v_BDATAVER
                                    Case "N" 's_BDATAVER2 = "當年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202104010314")
                                    Case "B" 's_BDATAVER2 = "前一年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202104010314")
                                End Select
                        End Select
                    Case "05"
                        Select Case v_TPlanlist1
                            Case "28" 's_PLANNAME2 = "產投"
                                Select Case v_BDATAVER
                                    Case "N" 's_BDATAVER2 = "當年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202105010318")
                                    Case "B" 's_BDATAVER2 = "前一年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202105010318")
                                End Select
                            Case "54"

                                Select Case v_BDATAVER's_PLANNAME2 = "充電起飛"
                                    Case "N" 's_BDATAVER2 = "當年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202105010319")
                                    Case "B" 's_BDATAVER2 = "前一年度"
                                        s_attachmn = String.Format("{0}.xlsx", fileN1)
                                        s_filename = String.Format("{0}_{1}.xlsx", fileN1, "202105010319")
                                End Select
                        End Select

                End Select
        End Select

        'If s_filename = "" Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
        '    Exit Sub
        'End If
        Return s_filename
    End Function

    '檢核查詢參數，有誤為true
    Sub CHK_SchNGVal(ByRef flag_NG_1 As Boolean)
        'Dim flag_NG_1 As Boolean = False
        Dim v_TPlanlist1 As String = TIMS.GetListValue(TPlanlist1)
        Dim v_yearlist As String = TIMS.GetListValue(yearlist)
        Dim v_monthlist As String = TIMS.GetListValue(monthlist)
        Dim v_BDATAVER As String = TIMS.GetListValue(rbl_BDATAVER)
        If v_TPlanlist1 = "" Then flag_NG_1 = True
        If v_yearlist = "" Then flag_NG_1 = True
        If v_monthlist = "" Then flag_NG_1 = True
        If v_BDATAVER = "" Then flag_NG_1 = True
        '    3832339 110_04_產投_當年度_綜合查詢統計表_202104010312.xlsx
        '  6831157 110_04_產投_前一年度_綜合查詢統計表_202104010313.xlsx
        '   86161 110_04_充電起飛_當年度_綜合查詢統計表_202104010314.xlsx
        '39 110_04_充電起飛_前一年度_綜合查詢統計表_202104010314.xlsx

        'If flag_NG_1 Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
        '    Exit Sub
        'End If
        'Return flag_NG_1
    End Sub

    ''' <summary>
    ''' 下載匯出檔
    ''' </summary>
    Sub Utl_DOWNLOAD1()
        Dim flag_NG_1 As Boolean = False
        Call CHK_SchNGVal(flag_NG_1)
        If flag_NG_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If
        Dim s_attachmn As String = ""
        Dim s_filename As String = ""
        s_filename = GET_FILENAME1(s_attachmn)
        If s_filename = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Exit Sub
        End If

        Dim s_MapFTXLSXPath As String = TIMS.Utl_GetConfigSet("MapFTXLSXPath")
        If s_MapFTXLSXPath = "" Then s_MapFTXLSXPath = "~/XLSX/"
        Dim s_mapfile As String = Server.MapPath(s_MapFTXLSXPath & s_filename)
        If Not IO.File.Exists(s_mapfile) Then
            Dim s_ERR As String = TIMS.cst_NODATAMsg1
            If s_filename <> "" Then s_ERR = (TIMS.cst_NODATAMsg1 & s_filename)
            If s_attachmn <> "" Then s_ERR = (TIMS.cst_NODATAMsg1 & s_attachmn)
            Common.MessageBox(Me, s_ERR)
            Exit Sub
        End If

        Dim Response As System.Web.HttpResponse = System.Web.HttpContext.Current.Response
        Response.ClearContent()
        Response.Clear()
        Response.ContentType = "application/vnd.ms-excel" '"text/plain"
        Response.AddHeader("Content-Disposition", String.Format("attachment; filename={0};", s_attachmn))
        Response.TransmitFile(s_mapfile)
        Response.Flush()
        Response.End()
    End Sub

    Protected Sub bt_DOWNLOADFILE_Click(sender As Object, e As EventArgs) Handles bt_DOWNLOADFILE.Click
        Call Utl_DOWNLOAD1()
    End Sub

    'Protected Sub bt_EXPORT_Click(sender As Object, e As EventArgs) Handles bt_EXPORT.Click
    '    Utl_EXPORT1()
    'End Sub

End Class