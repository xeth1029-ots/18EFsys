Partial Class SD_01_004
    Inherits AuthBasePage

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    '輸入個資安全密碼()
    Const cst_btn_button1 As String = "button1" '查詢
    Const cst_btn_button13 As String = "button13" '匯出
    Const cst_btn_btndivPwdSubmit As String = "btndivpwdsubmit" ' hidSchBtnNum.value: 1.正常查詢 2.正常匯出
    'Const cst_msgA1 As String="已過最晚可e網報名審核作業時間"
    Dim vMsg As String = ""
    Dim CPdt As DataTable 'Session(cst_ssDataTable)=dt 傳遞搜尋後TABLE使用。
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    Const cst_printFN1 As String = "SD_01_001_1"
    'Session(cst_ssDDCHKVIEW)=Nothing '登入後就永不清除SESSION。 
    Const cst_ssDDCHKVIEW As String = "ssDDCHKVIEW"

    'If Convert.ToString(drv("signUpStatus"))="0" AndAlso Convert.ToString(drv("DDCHKVIEW"))="Y" Then
    'eSerNum
    Dim str_addedit_aspx As String = "" '"SD_01_004_add.aspx?ID=" & Request("ID")
    'Const cst_IJC As String="民眾-XXX的參訓資格，因與委外實施基準條款JJJ有抵觸，請確認是否要同意此民眾的報名?"
    '產投政府已補助經費
    'Const Cst_SupplyMoney As Integer=50000 '99年(由3萬改為 3年5萬喔) by AMU 20100607

    '(「職場續航」之課程勾稽投保年資)
    Dim gfg_WYROLE As Boolean = False ' CHECK_WYROLE()

    Dim sMemo As String = ""
    Dim ss As String = ""
    Dim IDNOArray As New ArrayList
    Dim dt_Key_Degree As DataTable
    Dim dt_Key_Identity As DataTable
    Dim dt_Key_Trade As DataTable
    Dim dtClassInfo As DataTable
    Dim dtZipCode As DataTable
    Dim aCTID As String '縣市別

    '144、Stud_EnterTemp2-e網報名資料暫存檔
    Dim aIDNO As String             '身分證號碼 --varchar(15)
    Dim aName As String             '姓名--nvarchar(30)
    Dim aSex As String              '性別--男@M 女@F
    Dim aBirthday As String         '出生日期--yyyy/MM/dd
    Dim aPassPortNO As String       '身分別--1:本國2:外籍
    'Dim aMaritalStatus As String   '婚姻狀況--1.已;2.未
    Dim aDegreeID As String         '學歷代碼	varchar(3)
    Dim aGradID As String           '畢業狀況代碼	varchar(3)
    'Dim aSchool As String          '學校名稱 	nvarchar(30)
    'Dim aDepartment As String      '科系名稱 	nvarchar(128)
    Dim aMilitaryID As String       '兵役代碼    	varchar(3)
    Dim aZipCode As String          '通訊郵遞區號 	int
    Dim aZipCODE6W As String        '通訊郵遞區號後2碼 	int
    Dim aAddress As String          '通訊地址 	nvarchar(200)
    Dim aPhone1 As String           '聯絡電話(日)	varchar(25)
    Dim aPhone2 As String           '聯絡電話(夜)	varchar(25) null
    Dim aCellPhone As String        '行動電話 	varchar(25) null
    Dim aEmail As String            'Email 	varchar(30) null(最好是可以填若沒有填'無')
    Dim aIsAgree As String          '同意否	char(1)		null (Y/N)

    '146、Stud_EnterType2-e網報名職類檔
    Dim aEnterDate As String        '輸入日期 	datetime	
    Dim aSerNum As String           '序號	int	PK
    Dim aRelEnterDate As String     '報名日期	datetime	
    Dim aExamNo As String           '准考證號	varchar(14 )	 
    Dim aOCID1 As String            '報考班別代碼1	int
    Dim aTMID1 As String            '報考職類ID1	int
    'Dim aOCID2 As String           '報考班別代碼2 int	 
    'Dim aTMID2 As String           '報考職類ID2	int
    'Dim aOCID3 As String           '報考班別代碼3	int
    'Dim aTMID3 As String           '報考職類ID3	int
    Dim aIdentityID As String       '參訓身分別代碼	varchar(50 )	 
    Dim aRID As String              'RID 	varcha r(10)	
    Dim aPlanID As String           '計畫代碼	int
    Dim aCCLID As String            'CCLID	int 
    Dim aSIGNUPSTATUS As String     '報名狀態	int SIGNUPSTATUS '0 :收件完成'1 :報名成功'2 :報名失敗'3 :正取(Key_SelResult)'4 :備取'5 :未錄取
    Dim asignUpMemo As String       '報名備註	nvarchar(250) varchar(128)'備註(失敗原因)
    Dim aIsOut As String            '是否轉出(e網)	bit 
    Dim aSupplyID As String         '補助比例代碼	varchar(1)
    Dim aBudID As String            '預算別代碼	varchar(3)

    '147、STUD_ENTERTRAIN2-線上報名資料(產學訓)
    Dim aSEID As String               '流水號PK  int
    Dim aeSerNum As String            '線上報名職類流水號 int
    Dim aZipCode2 As String           '戶籍地址-郵遞區號  int
    Dim aZipCode2_6W As String        '戶籍地址-郵遞區號後2碼 int
    Dim aHouseholdAddress As String   '戶籍地址-地址 nvarchar(200)
    Dim aMIdentityID As String        '主要參訓身分別 varchar(3)

    'Dim aHandTypeID As String         '障礙類別 varchar(3)
    'Dim aHandLevelID As String        '障礙等級 varchar(3)
    'Dim aPriorWorkOrg1 As String      '受訓前工作單位名稱1 nvarchar(30)
    'Dim aTitle1 As String             '職稱1 nvarchar(20)
    'Dim aPriorWorkOrg2 As String      '受訓前工作單位名稱2 nvarchar(30)
    'Dim aTitle2 As String             '職稱2 nvarchar(20)
    'Dim aSOfficeYM1 As String         '任職起日1 Datetime
    'Dim aFOfficeYM1 As String         '任職迄日1 Datetime
    'Dim aSOfficeYM2 As String         '任職起日2 Datetime
    'Dim aFOfficeYM2 As String         '任職迄日2 Datetime
    Dim aPriorWorkPay As String        '受訓前薪資 Int
    'Dim aRealJobless As String        '失業週數   varchar(3)
    'Dim aJoblessID As String          '失業週數代碼 varchar(3)
    'Dim aTraffic As String            '交通方式     Int
    'Dim aShowDetail As String         '是否供求才廠商查詢 char(1)
    Dim aAcctMode As String            '郵政或金融         Bit
    Dim aPostNo As String              '郵政-局號          nvarchar(50)

    Dim aBankName As String            '銀行名稱           nvarchar(100)
    Dim aAcctHeadNo As String          '金融-總代號        nvarchar(50)
    Dim aExBankName As String          '分行名稱                 nvarchar(200)
    Dim aAcctExNo As String            '分行代碼                 nvarchar(100)
    Dim aAcctNo As String              'AcctNo(郵政)帳號    nvarchar(50)
    Dim aAcctNo2 As String             'AcctNo(銀行)帳號    nvarchar(50)

    Dim aFirDate As String             '第一次投保勞保日   Datetime
    Dim aUname As String               '公司名稱           nvarchar(50)
    Dim aIntaxno As String             '服務單位統一編號 varchar(10)
    'Dim aServDept As String           '服務部門           nvarchar(50)
    Dim aJobTitle As String            '職稱               nvarchar(50)
    'Dim aZip As String                '郵遞區號           Int
    'Dim aAddr As String               '公司地址           nvarchar(100)
    'Dim aTel As String                '公司電話           nvarchar(30)
    'Dim aFax As String                '公司傳真           nvarchar(30)
    'Dim aSDate As String              '個人到任目前任職公司起日 Datetime
    'Dim aSJDate As String             '個人到任目前職務起日     Datetime
    'Dim aSPDate As String             '最近升遷日期           Datetime
    Dim aQ1 As String                  'Q1是否由公司推薦參訓    Bit
    Dim aQ2 As String '參訓動機
    'Dim aQ2_1 As String               'Q2_1參訓動機1            Int
    'Dim aQ2_2 As String               'Q2_2參訓動機2            Int
    'Dim aQ2_3 As String               'Q2_3參訓動機3            Int
    'Dim aQ2_4 As String               'Q2_4參訓動機4            Int
    Dim aQ3 As String                  'Q3訓後動向               Int
    Dim aQ3_Other As String            'Q3訓後動向其他           nvarchar(50)
    Dim aQ4 As String                  'Q4服務單位行業別         varchar(2)
    Dim aQ5 As String                  'Q5服務單位是否屬於中小企業 Bit
    Dim aQ61 As String                 '個人年資                 Int
    Dim aQ62 As String                 '公司年資                 Int
    Dim aQ63 As String                 '職位年資                 Int
    Dim aQ64 As String                 '升遷年資                 Int
    Dim aActNo As String               '保險證號                 varchar(20)
    Dim aActname As String             '投保公司名稱             nvarchar(100)
    Dim actTel As String               '投保單位電話
    Dim actZipCode As String           '投保單位郵遞區號前三碼
    Dim actZipCODE6W As String         '投保單位郵遞區號後二碼
    Dim actAddress As String           '投保單位地址

    Dim aIseMail As String             '是否願意收到職訓通知     Char(1)
    Dim aActType As String             '投保類別               Char(1)
    Dim aScale As String               '服務單位規模            Char(1)
    Dim orgKind As String              '機構別代碼，用來判斷是否為勞工團體，延伸判斷帳戶類別是否可以使用代碼2(訓練單位代轉現金)

    'eSerNum=DataGrid1.DataKeys(item.ItemIndex)
    'Const cst_序號 As Integer=0
    'Const cst_姓名 As Integer=1          'Cells(cst_姓名)
    Const cst_身分證號碼 As Integer = 2     'Cells(cst_身分證號碼)
    Const cst_報名機構 As Integer = 3
    Const cst_報名班級 As Integer = 4
    Const cst_報名日期 As Integer = 5       'Cells(5)=Cells(cst_報名日期)
    Const cst_報名審核 As Integer = 6

    Const cst_報名路徑 As Integer = 7       'Const cst_預算別 As Integer=8
    Const cst_是否為在職者補助身分 As Integer = 8
    'Const cst_協助基金 As Integer=9
    Const cst_保險證號 As Integer = 9 '10
    Const cst_失敗原因 As Integer = 10 '11
    'Const Cst_請選擇 As String="請選擇" '補助比例寫死了

    '"補助費時段重疊情形" '"補助費及同時段重疊報名情形" '重複參訓
    'Const cst_tims28_BtnHistory As String="報名及補助查詢" ' "補助費時段重疊情形"
    'Const cst_tims28_BtnHis_confirm As String="補助費及同時段重疊報名情形?"
    'Const cst_tims28_Double As String="有參訓時段重疊的報名情形，請與學員確認。"
    'Const cst_tims28_over6w As String="預估目前補助費使用已達6萬元（包含已核撥、參訓中、報名中的課程）請再提醒學員。"
    Const cst_tims28_labdouble As String = "報名時段重疊名單："
    'Const cst_tims28_labover6w As String="補助費已達6萬名單："
    Const cst_tims28_SubsidyWarningCost9W As String = "90000"
    Const cst_tims28_labover9w As String = "補助費已達9萬名單："
    Const cst_tims28_SubsidyWarningCost6W As String = "60000"
    Const cst_tims28_labover6w As String = "補助費已達6萬名單："

    'divEnterDouble/divEnterMoney/labEnterDouble/labEnterMoney
    'Const cst_tims28_label1 As String="* 表示該學員有其他班級仍在訓中,請點選檢視功能查詢"
    'Const cst_tims28_labsubsidycost As String="* 表示該學員已申請職訓生活津貼,請點選檢視功能查詢"
    'Const cst_tims28_label2 As String=""
    Const cst_tims28_label1 As String = "* 表示該學員尚有產投或自辦在職課程仍在訓中，請查詢學員參訓歷史"
    Const cst_tims28_labsubsidycost As String = ""
    Const cst_tims28_label2 As String = ""

    '公司/商業負責人 (Master)
    Const cst_xMaster As String = "民眾XXX具公司/商業負責人身分，非屬失業勞工，不得報名失業者職前訓練。"

    '是否啟用日期欄位資訊西元轉民國機制
    Dim flag_ROC As Boolean = False

#Region "FUNCTION1"

    ''' <summary>
    '''  'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
    ''' </summary>
    ''' <returns></returns>
    Function GET_SsignUpStatus_VAL() As String
        '970520 Andy 修正審核狀態 
        'Stud_EnterType2  [signUpStatus] 'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim status As String = ""
        'SearchStr2 += " and  signUpStatus in (0,1,2,3,4,5)" & vbCrLf '尚未審核
        If SsignUpStatus.Items(0).Selected Then status &= ",0" '尚未審核
        If SsignUpStatus.Items(1).Selected Then status &= ",1,3,4,5" '審核成功
        If SsignUpStatus.Items(2).Selected Then status &= ",2" '審核失敗
        If status <> "" Then status = Right(status, Len(status) - 1)
        Return status
    End Function

    '查詢 (SQL) 語法匯出
    Function Search_SQL(ByRef parms As Hashtable) As String
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Return ""
        If Not sm.IsLogin Then Return ""
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = Convert.ToString(sm.UserInfo.RID)
        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim sqlWSB As String = TIMS.Get_StdBlackWSB(Me, iStdBlackType, stdBLACK2TPLANID, 1)

        Dim Relship As String = ""
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        If IDNO.Text <> "" Then IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        Dim flagNUCdition1 As Boolean = False
        If IDNO.Text <> "" AndAlso flgROLEIDx0xLIDx0 Then flagNUCdition1 = True '若為 SUPER UESR 且有輸入IDNO 可不判斷 此條件
        'Get Relship, Me.DistValue.Value 
        DistValue.Value = ""
        'RIDValue.Value 若沒有值代入 ' Convert.ToString(sm.UserInfo.RID)
        Dim sql_r As String = " SELECT RELSHIP, DISTID FROM AUTH_RELSHIP WHERE RID=@RID"
        Call TIMS.OpenDbConn(objconn)
        Dim dt_r As New DataTable
        Dim sCmd_r As New SqlCommand(sql_r, objconn)
        With sCmd_r
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
            dt_r.Load(.ExecuteReader())
        End With
        'parms.Clear()
        'parms.Add("RID", RIDValue.Value)
        'dt=DbAccess.GetDataTable(sql, objconn, parms)
        If dt_r.Rows.Count = 0 Then Return ""
        Relship = $"{dt_r.Rows(0)("Relship")}"
        DistValue.Value = $"{dt_r.Rows(0)("DistID")}"
        'Stud_EnterType2  [signUpStatus] 'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim status As String = GET_SsignUpStatus_VAL()

        '(「職場續航」之課程勾稽投保年資)
        'Dim fg_WYROLE As Boolean=CHECK_WYROLE()

        'BudID (預算別) by AMU 20080602
        '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
        '01、04、05、15及08者其補助經費來源歸屬為 03:就保基金
        '02、03、06、07者其經費來源歸屬為 02:就安基金
        '09與無法辨視者為 99:不予補助對象
        '20090324(Milor)補充判斷未審核的資料，才列入判斷預算別狀態，如已審核，則忠實呈現預算別。  'by AMU 20100414
        '例外,就是報名學員的保險證號前三碼為075時,預算別是"不補助",補助比例是"0%".
        'SupplyID (補助比例) by AMU 20080602
        '9:不補助者0%        '1:為 01:一般身分者，補助80%        '2:不為一般身分者，補助100%
        '20090324(Milor)補充判斷未審核的資料，才列入判斷補助比例狀態，如已審核，則忠實呈現補助比例。
        parms.Clear()

        Dim sql As String = ""
        sql &= sqlWSB
        '(「職場續航」之課程勾稽投保年資)
        sql &= String.Concat(" SELECT ", If(gfg_WYROLE, "a3.WYROLE,a3.AGE,a3.ITRMY,", ""), "b2.MIdentityID")
        sql &= " ,a.IdentityID ,a.eSerNum ,a.eSETID ,b.Name ,b.IDNO ,b.Email ,b.Phone1 ,b.Phone2 ,b.CellPhone" & vbCrLf
        sql &= " ,CONVERT(varchar, b.BirthDay, 111) BirthDay" & vbCrLf
        'ACTNO產投跟職前不同欄位 a2.STUD_ENTERSUBDATA2 / b2.STUD_ENTERTRAIN2
        'SELECT DISTINCT ACTNO FROM STUD_ENTERTRAIN2 WHERE MODIFYDATE >=DATEADD(DAY, getdate(), -100)
        sql &= " ,IsNull(b2.ActNo,IsNull(a2.ActNo,'無資料')) ActNo" & vbCrLf
        'Public Const cst_Actno_NG2 As String="'09'"
        'Public Const cst_Actno_NG3 As String="'075','175','076','176'"
        'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql &= " ,CASE WHEN a.SignUpStatus=0 AND a.BudID IS NULL THEN" & vbCrLf
        '例外,就是報名學員的保險證號前三碼為075時,預算別是"不補助",補助比例是"0%".
        '1.目前投保證號設定填寫時必須為數字0開頭，勞保局因07（自願加保者）數字額滿後新增17，故開放新增填寫時開頭數字可為1
        '2.開頭數字為075、175（裁減續保）、076、176（職災續保）、09（訓）皆為不予補助對象，並設定阻擋。
        '03:就保基金 02:就安基金 99:不予補助對象
        sql &= "  CASE WHEN SUBSTRING(b2.ActNo,1,3) IN ('075','175','076','176') THEN '99'" & vbCrLf
        '於「e網報名審核」增加 投保證號170 預算別為 就安
        sql &= "  WHEN SUBSTRING(b2.ActNo,1,3) IN ('170') THEN '02'" & vbCrLf
        sql &= "  WHEN SUBSTRING(b2.ActNo,1,2) IN ('09') THEN '99'" & vbCrLf
        '險證號前2碼判讀
        sql &= "  WHEN SUBSTRING(b2.ActNo,1,2) IN ('01','04','05','15','08') THEN '03'" & vbCrLf
        sql &= "  WHEN SUBSTRING(b2.ActNo,1,2) IN ('02','03','06','07') THEN '02'" & vbCrLf
        '(其餘為)'99:不予補助對象 / '不為null照原顯示
        sql &= "  ELSE '99' END ELSE a.BudID END BudID" & vbCrLf

        '若補助比例是null
        'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql &= " ,CASE WHEN a.SignUpStatus=0 AND a.SupplyID IS NULL THEN" & vbCrLf
        '例外,就是報名學員的保險證號前三碼為075時,預算別是"不補助",補助比例是"0%".
        sql &= "  CASE WHEN SUBSTRING(b2.ActNo,1,3) IN ('075','175','076','176') THEN 9" & vbCrLf '9:0%
        sql &= "  WHEN SUBSTRING(b2.ActNo,1,2) IN ('09') THEN 9" & vbCrLf '9:0%
        '險證號前2碼判讀or前3碼判讀 ('1:80% 2:100%)
        '1:01一般身分者，補助80% / 2:不為一般身分者，補助100% / 9:不補助者0%
        sql &= "  WHEN SUBSTRING(b2.ActNo,1,2) IN ('01','04','05','15','08','02','03','06','07')" & vbCrLf
        sql &= "  OR SUBSTRING(b2.ActNo,1,3) IN ('170')" & vbCrLf
        sql &= "  THEN (CASE WHEN b2.MIdentityID='01' THEN 1 ELSE 2 END)" & vbCrLf '1:補助80% /2:補助100%
        '(其餘為) 9:不補助者0% / '不為null照原顯示
        sql &= "  ELSE 9 END ELSE CONVERT(NUMERIC, a.SupplyID) END SupplyID" & vbCrLf

        'cst_是否為在職者補助身分
        sql &= " ,a.WorkSuppIdent ,mi.Name MIdentityName ,a.RID" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.EnterDate, 111) EnterDate ,a.RelEnterDate" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.RelEnterDate, 111) RelEnterDate2" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.RelEnterDate, 108) RelEnterDate3" & vbCrLf
        sql &= " ,c.ClassCName ClassCName1 ,c.CyclType CyclType1" & vbCrLf
        'sql &= " ,d.ClassCName ClassCName2 ,d.CyclType CyclType2" & vbCrLf
        'sql &= " ,e.ClassCName ClassCName3 ,e.CyclType CyclType3" & vbCrLf
        'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        sql &= " ,a.signUpStatus" & vbCrLf
        sql &= " ,CASE WHEN a2.FOfficeYM1 IS NOT NULL THEN a2.FOfficeYM1 ELSE b2.FOfficeYM1 END FOfficeYM1" & vbCrLf
        '備註(失敗原因)
        sql &= " ,b2.Uname ,a.ExamNo ,a.MODIFYACCT ,f.OrgName ,f.Relship" & vbCrLf
        sql &= " ,a.OCID1 ,o2.OrgName ORGNAME2 ,f.OrgLevel" & vbCrLf
        'sql &= " ,IsNull(a.OCID2,0) AS OCID2 ,IsNull(a.OCID3,0) AS OCID3" & vbCrLf
        sql &= " ,g.LevelName ,CONVERT(VARCHAR, c.STDate, 111) STDate1 ,CONVERT(VARCHAR, c.FTDate, 111) FTDate" & vbCrLf
        sql &= " ,dbo.FN_STDAGE(b.BirthDay,c.STDate) STDAGE" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, c.SENTERDATE, 111) SENTERDATE" & vbCrLf '報名開始日
        sql &= " ,CONVERT(VARCHAR, c.FENTERDATE, 111) FENTERDATE" & vbCrLf '報名截止日
        sql &= " ,CONVERT(VARCHAR, c.ExamDate, 111) ExamDate ,c.Thours" & vbCrLf
        '有多筆 學員處分資料
        sql &= " ,CASE WHEN WSB.IDNO IS NOT NULL THEN 'Y' ELSE 'N' END IsStdBlack" & vbCrLf
        '倘報名人數(已扣除自行取消報名及報名審核失敗)超過該班訓練人數+備取預設人數(目前為10人)後之名單
        sql &= " ,a.SignNo ,CASE WHEN ISNULL(a.SignNo,0)<=ISNULL(c.TNUM,0)+10 THEN 'Y' END DDCHKVIEW" & vbCrLf

        '(「職場續航」之課程勾稽投保年資)
        If gfg_WYROLE Then
            sql &= " , a.SIGNUPMEMO" & vbCrLf
        Else
            'OJT-24061203 'a.signUpMemo
            sql &= " ,CASE WHEN ISNULL(a.SignNo,0)>ISNULL(c.TNUM,0)+10 AND a.signUpStatus=0 AND a.SIGNUPMEMO IS NULL THEN '本班報名已額滿' ELSE a.SIGNUPMEMO END SIGNUPMEMO" & vbCrLf
        End If
        'sql &= " ,dbo.fn_GET_GOVCOST(upper(b.idno),CONVERT(varchar, c.STDate, 111)) GovCost" & vbCrLf
        sql &= " ,0 GovCost" & vbCrLf
        sql &= " ,a.EnterPath,dbo.DECODE6(a.EnterPath,'O','外網','o','內網','未知') EnterPath_N" & vbCrLf
        sql &= " ,a.CMASTER1" & vbCrLf '認定為公司負責人
        sql &= " ,a.CMASTER1NT" & vbCrLf '已切結
        ' dbo.fn_GET_GOVCOST('" & UCase(drv("IDNO")) & "','" & drv("STDate1").ToString & "') GovCost
        '.e網報名審核的查詢結果列表，若顯示的報名資料的報名班級已過最晚可e網報名審核作業時間
        '，則報名審核欄位的o成功o失敗的選擇，就變成灰色不可選擇狀態。
        sql &= " ,CASE WHEN c.FENTERDATE2 < GETDATE() THEN 'Y' END LOCK1" & vbCrLf 'LOCK1()
        '屆退官兵身分者 'Session retreat Soldiers The identity of persons
        sql &= " ,CASE WHEN dbo.TRUNC_DATETIME(a.PREEXDATE) > dbo.TRUNC_DATETIME(GETDATE()) THEN 'Y' END SRSOLDIERS ,b.ZIPCODE ,b2.ZIPCODE2" & vbCrLf
        sql &= " ,ip.TPlanID ,ip.YEARS" & vbCrLf
        sql &= " FROM STUD_ENTERTYPE2 a" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP2 b ON a.eSETID=b.eSETID" & vbCrLf

        If IDNO.Text <> "" Then
            sql &= " AND b.IDNO=@IDNO" & vbCrLf
            parms.Add("IDNO", IDNO.Text)
        End If

        sql &= " LEFT JOIN STUD_ENTERTRAIN2 b2 ON b2.eSerNum=a.eSerNum" & vbCrLf
        sql &= " LEFT JOIN STUD_ENTERSUBDATA2 a2 ON a2.eSerNum=a.eSerNum" & vbCrLf
        '(「職場續航」之課程勾稽投保年資) dbo.V_BLIGATEDATA28YM,,V_BLIGATEDATA28YM,,STUD_BLIGATEDATA28YM
        If gfg_WYROLE Then sql &= " LEFT JOIN dbo.V_BLIGATEDATA28YM a3 ON a3.eSerNum=a.eSerNum" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON a.OCID1=c.OCID" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PLANID=c.PLANID" & vbCrLf
        'sql &= " LEFT JOIN Class_ClassInfo d ON a.OCID2=d.OCID" & vbCrLf
        'sql &= " LEFT JOIN Class_ClassInfo e ON a.OCID3=e.OCID" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME f ON a.RID=f.RID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO o2 ON o2.comidno=c.comidno" & vbCrLf
        sql &= " LEFT JOIN CLASS_CLASSLEVEL g ON a.OCID1=g.OCID AND a.CCLID=g.CCLID" & vbCrLf
        sql &= " LEFT JOIN KEY_IDENTITY mi ON mi.IdentityID=b2.MIdentityID" & vbCrLf
        '有多筆 學員處分資料 (依系統日期1年內處分。)
        sql &= " LEFT JOIN WSB ON WSB.IDNO=b.IDNO" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            If IDNO.Text = "" Then
                If Relship <> "" Then
                    sql &= " AND f.Relship LIKE @Relship" & vbCrLf
                    parms.Add("Relship", Relship & "%")
                End If
            End If
        Else
            If Relship <> "" Then
                sql &= " AND f.Relship LIKE @Relship" & vbCrLf
                parms.Add("Relship", Relship & "%")
            End If
        End If

        If status <> "" Then
            'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
            sql &= $" AND a.signUpStatus IN ({status})" & vbCrLf
        End If

        If OCIDValue1.Value <> "" Then
            sql &= " AND a.OCID1=@OCID1" & vbCrLf
            'sql &= " AND (a.OCID1='" & OCIDValue1.Value & "'" & vbCrLf
            'sql &= "  OR a.OCID2='" & OCIDValue1.Value & "'" & vbCrLf
            'sql &= "  OR a.OCID3='" & OCIDValue1.Value & "')" & vbCrLf
            parms.Add("OCID1", OCIDValue1.Value)
        End If

        If cjobValue.Value <> "" Then
            sql &= " AND c.CJOB_UNKEY=@CJOB_UNKEY" & vbCrLf
            parms.Add("CJOB_UNKEY", cjobValue.Value)
        End If

        If start_date.Text <> "" Then
            sql &= " AND dbo.TRUNC_DATETIME(a.RelEnterDate) >= @RelEnterDate1" & vbCrLf
            parms.Add("RelEnterDate1", If(flag_ROC, TIMS.Cdate18(start_date.Text), TIMS.Cdate3(start_date.Text)))
        End If

        Dim v_end_date As String = If(flag_ROC, TIMS.Cdate18(end_date.Text), TIMS.Cdate3(end_date.Text))
        Dim v_end_date_a1 As String = ""
        If v_end_date <> "" Then v_end_date_a1 = TIMS.Cdate3(CDate(v_end_date).AddDays(1)) '+1天
        If v_end_date_a1 <> "" Then
            sql &= " AND dbo.TRUNC_DATETIME(a.RelEnterDate) < @RelEnterDate2" & vbCrLf
            parms.Add("RelEnterDate2", v_end_date_a1)
        End If

        'If end_date.Text <> "" Then
        '    sql &= " AND dbo.TRUNC_DATETIME(a.RelEnterDate) < @RelEnterDate2" & vbCrLf
        '    parms.Add("RelEnterDate1", If(flag_ROC, TIMS.cdate18(end_date.Text), TIMS.cdate3(end_date.Text)))
        '    If flag_ROC Then
        '        parms.Add("RelEnterDate2", CDate(TIMS.cdate18(end_date.Text)).AddDays(1).ToString("yyyy/MM/dd"))
        '    Else
        '        parms.Add("RelEnterDate2", CDate(end_date.Text).AddDays(1).ToString("yyyy/MM/dd"))
        '    End If
        'End If

        If sm.UserInfo.RID = "A" Then
            If Not flagNUCdition1 Then
                '若為 SUPER UESR 且有輸入IDNO 可不判斷 此條件
                sql &= " AND ip.TPlanID=@TPlanID AND ip.Years=@Years" & vbCrLf
                'sql &= " AND ip.Years IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' AND Years='" & sm.UserInfo.Years & "')" & vbCrLf
                parms.Add("TPlanID", sm.UserInfo.TPlanID)
                parms.Add("Years", sm.UserInfo.Years)
            End If
        Else
            sql &= " AND ip.PlanID=@PlanID" & vbCrLf
            parms.Add("PlanID", sm.UserInfo.PlanID)
        End If
        'select distinct MIdentityID,count(1) cnt from Stud_EnterTrain2 group by MIdentityID order by 1
        'select distinct IdentityID,count(1) cnt from STUD_ENTERTYPE2 group by IdentityID order by 1

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, $"--{vbCrLf}{TIMS.GetMyValue5(parms)}{vbCrLf}--##SD_01_004.aspx,sSql:{vbCrLf}{sql}")
        End If
        Return sql
    End Function

    '(「職場續航」之課程勾稽投保年資) TRUE: 職場續航 FALSE:(非)
    'Private Function CHECK_WYROLE() As Boolean
    '    'Dim fg_WYROLE As Boolean=False '(「職場續航」之課程勾稽投保年資)
    '    Dim drCC_2854 As DataRow=TIMS.GetOCIDDate2854(OCIDValue1.Value, objconn)
    '    Dim bfg1_OK As Boolean=(OCIDValue1.Value <> "" AndAlso drCC_2854 IsNot Nothing) '有輸入班級資料-檢核班級是否OK(範圍內)
    '    If Not bfg1_OK Then Return False '(查無資料)
    '    'If bfg1_OK Then fg_WYROLE=(Convert.ToString(drCC_2854("KID25_8")) <> "") '(「職場續航」之課程勾稽投保年資) 
    '    '(「職場續航」之課程勾稽投保年資)
    '    Return (Convert.ToString(drCC_2854("KID25_8")) <> "")
    'End Function

    ''#Region "FUNCTION1"

    '刪除 (含還原刪除)
    Public Shared Sub DEL_STUDENTERTYPE(ByVal tmpSETID As Integer, ByVal tmpOCID As Integer, ByVal sm As SessionModel, ByVal tmpTran As SqlTransaction)
        Dim myParam_u As New Hashtable From {{"ModifyAcct", Convert.ToString(sm.UserInfo.UserID)}, {"SETID", tmpSETID}, {"OCID", tmpOCID}}
        Dim sqlStr_u As String = ""
        sqlStr_u &= " UPDATE STUD_ENTERTYPE SET ModifyAcct=@ModifyAcct, ModifyDate=GETDATE() WHERE SETID=@SETID AND OCID1=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(sqlStr_u, tmpTran, myParam_u)

        Dim myParam_i As New Hashtable From {{"SETID", tmpSETID}, {"OCID", tmpOCID}}
        Dim sqlStr_i As String = ""
        sqlStr_i &= " INSERT INTO STUD_ENTERTYPEDELDATA (SETID ,ENTERDATE ,SERNUM ,EXAMNO ,OCID1 ,TMID1 ,OCID2 ,TMID2 ,OCID3 ,TMID3 ,WRITERESULT ,ORALRESULT ,TOTALRESULT ,ENTERCHANNEL ,IDENTITYID ,RID ,PLANID ,TRNDMODE ,TRNDTYPE ,Q1_1 ,Q1_2 ,Q1_2OTHER ,Q1_3 ,Q1_3OTHER ,Q1_4 ,Q1_4OTHER ,Q1_5 ,Q2_3 ,Q2_4 ,Q2_5OTHER ,MODIFYACCT ,MODIFYDATE ,TICKET_NO ,RELENTERDATE ,NOTEXAM ,CCLID ,ESETID ,ESERNUM ,TRANSDATE ,SEID ,SUPPLYID ,BUDID ,ENTERPATH ,HIGHEDUBG ,WORKSUPPIDENT ,USERNOSHOW ,NOTES ,PRIORWORKTYPE1 ,PRIORWORKORG1 ,SOFFICEYM1 ,FOFFICEYM1 ,ACTNO ,WSID ,INVOLLEAVER ,CFIRE1 ,CFIRE1NS ,CFIRE1REASON ,CFIRE1MACCT ,CFIRE1MDATE ,ENTERPATH2 ,CMASTER1 ,CMASTER1NS ,CMASTER1REASON ,CMASTER1MACCT ,CMASTER1MDATE ,CMASTER1NT ,CFIRE1R2 ,EXAMPLUS ,PREEXDATE)" & vbCrLf
        sqlStr_i &= " SELECT SETID ,ENTERDATE ,SERNUM ,EXAMNO ,OCID1 ,TMID1 ,OCID2 ,TMID2 ,OCID3 ,TMID3 ,WRITERESULT ,ORALRESULT ,TOTALRESULT ,ENTERCHANNEL ,IDENTITYID ,RID ,PLANID ,TRNDMODE ,TRNDTYPE ,Q1_1 ,Q1_2 ,Q1_2OTHER ,Q1_3 ,Q1_3OTHER ,Q1_4 ,Q1_4OTHER ,Q1_5 ,Q2_3 ,Q2_4 ,Q2_5OTHER ,MODIFYACCT ,MODIFYDATE ,TICKET_NO ,RELENTERDATE ,NOTEXAM ,CCLID ,ESETID ,ESERNUM ,TRANSDATE ,SEID ,SUPPLYID ,BUDID ,ENTERPATH ,HIGHEDUBG ,WORKSUPPIDENT ,USERNOSHOW ,NOTES ,PRIORWORKTYPE1 ,PRIORWORKORG1 ,SOFFICEYM1 ,FOFFICEYM1 ,ACTNO ,WSID ,INVOLLEAVER ,CFIRE1 ,CFIRE1NS ,CFIRE1REASON ,CFIRE1MACCT ,CFIRE1MDATE ,ENTERPATH2 ,CMASTER1 ,CMASTER1NS ,CMASTER1REASON ,CMASTER1MACCT ,CMASTER1MDATE ,CMASTER1NT ,CFIRE1R2 ,EXAMPLUS ,PREEXDATE" & vbCrLf
        sqlStr_i &= " FROM STUD_ENTERTYPE WHERE SETID=@SETID AND OCID1=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(sqlStr_i, tmpTran, myParam_i)

        Dim myParam_d As New Hashtable From {{"SETID", tmpSETID}, {"OCID", tmpOCID}}
        Dim sqlStr_d As String = " DELETE STUD_ENTERTYPE WHERE SETID=@SETID and OCID1=@OCID "
        DbAccess.ExecuteNonQuery(sqlStr_d, tmpTran, myParam_d)
    End Sub

    '刪除 (含還原刪除)
    Public Shared Sub DEL_STUDSELRESULT(ByVal tmpSETID As Integer, ByVal tmpOCID As Integer, ByVal sm As SessionModel, ByVal tmpTran As SqlTransaction)
        Dim myParam_u As New Hashtable From {{"ModifyAcct", sm.UserInfo.UserID}, {"SETID", tmpSETID}, {"OCID", tmpOCID}}
        Dim sqlStr_u As String = " UPDATE STUD_SELRESULT SET ModifyAcct=@ModifyAcct, ModifyDate=GETDATE() WHERE SETID=@SETID AND OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(sqlStr_u, tmpTran, myParam_u)

        Dim myParam_i As New Hashtable From {{"SETID", tmpSETID}, {"OCID", tmpOCID}}
        Dim sqlStr_i As String = ""
        sqlStr_i &= " INSERT INTO Stud_SelResultDelData" & vbCrLf
        sqlStr_i &= " SELECT * FROM Stud_SelResult WHERE SETID=@SETID AND OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(sqlStr_i, tmpTran, myParam_i)

        Dim myParam_d As New Hashtable From {{"SETID", tmpSETID}, {"OCID", tmpOCID}}
        Dim sqlStr_d As String = ""
        sqlStr_d = " DELETE STUD_SELRESULT WHERE SETID=@SETID AND OCID=@OCID "
        DbAccess.ExecuteNonQuery(sqlStr_d, tmpTran, myParam_d)
    End Sub

    '刪除
    Public Shared Sub DEL_STUDENTERTRAIN2(ByVal tmpeSerNum As Integer, ByVal sm As SessionModel, ByVal tmpTran As SqlTransaction)
        Dim sqlStrU As String = " UPDATE STUD_ENTERTRAIN2 SET ModifyAcct=@ModifyAcct, ModifyDate=GETDATE() WHERE eSerNum=@eSerNum"
        Dim myPmsU As New Hashtable From {{"ModifyAcct", Convert.ToString(sm.UserInfo.UserID)}, {"eSerNum", tmpeSerNum}}
        DbAccess.ExecuteNonQuery(sqlStrU, tmpTran, myPmsU)

        Dim s_COLUMN As String = "SEID,ESERNUM,ZIPCODE2,HOUSEHOLDADDRESS,MIDENTITYID,HANDTYPEID,HANDLEVELID,PRIORWORKORG1,TITLE1,PRIORWORKORG2,TITLE2,SOFFICEYM1,FOFFICEYM1
,SOFFICEYM2,FOFFICEYM2,PRIORWORKPAY,REALJOBLESS,JOBLESSID,TRAFFIC,SHOWDETAIL,ACCTMODE,POSTNO,ACCTHEADNO,BANKNAME,ACCTEXNO,EXBANKNAME,ACCTNO,FIRDATE,UNAME,INTAXNO
,ACTNO,ACTNAME,SERVDEPT,JOBTITLE,ZIP,ADDR,TEL,FAX,SDATE,SJDATE,SPDATE,Q1,Q2_1,Q2_2,Q2_3,Q2_4,Q3,Q3_OTHER,Q4,Q5,Q61,Q62,Q63,Q64,ISEMAIL,MODIFYACCT,MODIFYDATE
,ACTTYPE,SCALE,ZIPCODE2_2W,ACTTEL,ZIPCODE3,ZIPCODE3_2W,ACTADDRESS,INSURED,SERVDEPTID,JOBTITLEID,ZIPCODE3_N,ZIPCODE2_N,ZIP2W,ZIP_N,ZIP6W,ZIPCODE2_6W,ZIPCODE3_6W"
        Dim sqlStrI As String = String.Concat(" INSERT INTO STUD_ENTERTRAIN2DELDATA(", s_COLUMN, ") SELECT ", s_COLUMN, " FROM STUD_ENTERTRAIN2 WHERE eSerNum=@eSerNum")
        Dim myPmsI As New Hashtable From {{"eSerNum", tmpeSerNum}}
        DbAccess.ExecuteNonQuery(sqlStrI, tmpTran, myPmsI)

        Dim sqlStrD As String = " DELETE STUD_ENTERTRAIN2 WHERE eSerNum=@eSerNum"
        Dim myPmsD As New Hashtable From {{"eSerNum", tmpeSerNum}}
        DbAccess.ExecuteNonQuery(sqlStrD, tmpTran, myPmsD)
    End Sub

    '刪除
    Public Shared Sub DEL_STUDENTERTYPE2(ByVal tmpeSerNum As Integer, ByVal tmpeSETID As Integer, ByVal sm As SessionModel, ByVal tmpTran As SqlTransaction)
        Dim s_ModifyAcct As String = sm.UserInfo.UserID
        Dim myParam_u As New Hashtable From {{"ModifyAcct", s_ModifyAcct}, {"eSerNum", tmpeSerNum}, {"eSETID", tmpeSETID}}
        Dim sqlStr_u As String = " UPDATE STUD_ENTERTYPE2 SET ModifyAcct=@ModifyAcct, ModifyDate=GETDATE() WHERE eSerNum=@eSerNum AND eSETID=@eSETID"
        DbAccess.ExecuteNonQuery(sqlStr_u, tmpTran, myParam_u)

        Dim s_COLUMN As String = "ESERNUM ,ESETID,SETID,ENTERDATE,SERNUM,RELENTERDATE,EXAMNO,OCID1,TMID1,OCID2,TMID2,OCID3,TMID3,IDENTITYID,RID,PLANID,CCLID
,SIGNUPSTATUS,SIGNUPMEMO,ISOUT,SUPPLYID,BUDID,MODIFYACCT,MODIFYDATE ,ENTERPATH,WORKSUPPIDENT,USERNOSHOW,NOTES,ISEMAILFAIL,SIGNNO,INVOLLEAVER,CFIRE1,CFIRE1NS
,CFIRE1REASON,CFIRE1MACCT,CFIRE1MDATE,CMASTER1,CMASTER1NS,CMASTER1REASON,CMASTER1MACCT,CMASTER1MDATE,CMASTER1NT,CFIRE1R2,PREEXDATE,APID1,MIDENTITYID,ABANDON
,ABANDONReason,ABANDONACCT,ABANDONDATE"
        Dim myParam_i As New Hashtable From {{"eSerNum", tmpeSerNum}, {"eSETID", tmpeSETID}}
        Dim sqlStr_i As String = ""
        sqlStr_i &= String.Concat(" INSERT INTO STUD_ENTERTYPE2DelData(", s_COLUMN, ")")
        sqlStr_i &= String.Concat(" SELECT ", s_COLUMN, " FROM STUD_ENTERTYPE2 WHERE eSerNum=@eSerNum AND eSETID=@eSETID")
        DbAccess.ExecuteNonQuery(sqlStr_i, tmpTran, myParam_i)

        Dim myParam_d As New Hashtable From {{"eSerNum", tmpeSerNum}, {"eSETID", tmpeSETID}}
        Dim sqlStr_d As String = " DELETE STUD_ENTERTYPE2 WHERE eSerNum=@eSerNum AND eSETID=@eSETID "
        DbAccess.ExecuteNonQuery(sqlStr_d, tmpTran, myParam_d)
    End Sub

    '檢查 班級學員資料是否存在
    Function Check_Student(ByVal IDNO As String, ByVal OCID As String) As Boolean
        Dim rst As Boolean = False
        IDNO = TIMS.ClearSQM(IDNO)
        OCID = TIMS.ClearSQM(OCID)
        If IDNO = "" OrElse OCID = "" Then Return rst

        Dim sql As String = ""
        sql &= " SELECT 'x'" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO SS" & vbCrLf
        sql &= " JOIN Class_StudentsOfClass cs ON cs.SID=ss.SID" & vbCrLf
        sql &= " WHERE SS.IDNO=@IDNO AND cs.OCID=@OCID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    '檢查是否有內部報名
    Function Chk_StudEnterType(ByVal IDNO As String, ByVal OCID As String) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = " SELECT 1 FROM STUD_ENTERTYPE a JOIN STUD_ENTERTEMP b ON a.SETID=b.SETID WHERE b.IDNO=@IDNO AND a.OCID1=@OCID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO
            .Parameters.Add("OCID", SqlDbType.BigInt).Value = TIMS.GetValue2(OCID)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst
    End Function

    Sub UseKeepSearch()
        If Session("_SearchStr") Is Nothing Then Return
        'If Session("_SearchStr") IsNot Nothing Then
        'End If
        Dim MyValue As String = ""
        Me.ViewState("_SearchStr") = Session("_SearchStr")
        Session("_SearchStr") = Nothing
        MyValue = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "prg")
        If MyValue = "SD_01_004" Then
            center.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "center")
            RIDValue.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "RIDValue")
            TMID1.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "TMID1")
            OCID1.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "OCID1")
            TMIDValue1.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "OCIDValue1")
            IDNO.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "IDNO")
            IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
            start_date.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "start_date")
            end_date.Text = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "end_date")
            MyValue = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "SsignUpStatus")
            If MyValue <> "" Then Common.SetListItem(SsignUpStatus, MyValue)
            Me.ViewState("PageIndex") = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "PageIndex")
            MyValue = TIMS.GetMyValue(Me.ViewState("_SearchStr"), "submit")
            If MyValue = "1" Then
                'Button1_Click(sender, e)
                Call Search1()

                If DataGridTable.Visible AndAlso IsNumeric(Me.ViewState("PageIndex")) Then
                    '有資料SHOW出 跳頁
                    PageControler1.PageIndex = Me.ViewState("PageIndex")
                    'PageControler1.CreateData() 'Dim CPdt As DataTable 'CPdt=dt.Copy()
                    If CPdt IsNot Nothing Then
                        PageControler1.DataTableCreate(CPdt, PageControler1.Sort, PageControler1.PageIndex)
                    End If
                End If
            End If
        End If
    End Sub

    Sub KeepSearch()
        Session("_SearchStr") = Nothing
        Dim xSearchStr As String = ""
        xSearchStr = "prg=SD_01_004"
        xSearchStr &= "&center=" & TIMS.ClearSQM(center.Text)
        xSearchStr &= "&RIDValue=" & TIMS.ClearSQM(RIDValue.Value)
        xSearchStr &= "&TMID1=" & TIMS.ClearSQM(TMID1.Text)
        xSearchStr &= "&OCID1=" & TIMS.ClearSQM(OCID1.Text)
        xSearchStr &= "&OCIDValue1=" & TIMS.ClearSQM(OCIDValue1.Value)
        xSearchStr &= "&TMIDValue1=" & TIMS.ClearSQM(TMIDValue1.Value)
        xSearchStr &= "&IDNO=" & TIMS.ClearSQM(TIMS.ChangeIDNO(IDNO.Text))
        xSearchStr &= "&start_date=" & TIMS.ClearSQM(start_date.Text)
        xSearchStr &= "&end_date=" & TIMS.ClearSQM(end_date.Text)
        xSearchStr &= "&SsignUpStatus=" & TIMS.GetListValue(SsignUpStatus) 'TIMS.ClearSQM(SsignUpStatus.SelectedValue)
        xSearchStr &= "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        xSearchStr &= "&submit=" & If(DataGridTable.Visible, "1", "0")
        Session("_SearchStr") = xSearchStr
    End Sub

    '檢查資料正確性 (匯入功能)
    Function CheckImportData(ByVal colArray As Array) As String
        Dim Reason As String = ""
        Dim sql As String = ""
        'Dim dr As DataRow
        'Dim ocidFlag As Boolean
        Dim IDNOFlag As Boolean '**by Milor 20080529--用來判斷班級代碼與身分證號都正確時，才判斷匯入資料的身分證號是否重複
        Const cst_Len As Integer = 50
        Dim DoDataValid As Boolean '20090210 andy add 
        Reason = ""
        If colArray.Length < cst_Len Then
            'Reason &= "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Reason &= "欄位對應有誤<BR>"
            Reason &= "請注意欄位中是否有半形逗點<BR>"
        Else
            aCTID = colArray(0).ToString                  '縣市別代碼
            aOCID1 = TIMS.ClearSQM(colArray(1).ToString)                 '報名課程代碼
            aPassPortNO = colArray(2).ToString            '身分別
            aIDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(colArray(3))) '身分證號碼
            aBirthday = colArray(4).ToString            '出生日期
            aName = colArray(5).ToString                '姓名
            aSex = colArray(6).ToString                 '性別
            aDegreeID = colArray(7).ToString            '最高學歷代碼
            aPhone1 = colArray(8).ToString              '聯絡電話(日)
            aPhone2 = colArray(9).ToString              '聯絡電話(夜)
            aCellPhone = colArray(10).ToString          '行動電話
            aZipCode = colArray(11).ToString            '通訊郵遞區號
            aZipCODE6W = colArray(12).ToString          '通訊郵遞區號6碼
            aAddress = colArray(13).ToString            '通訊地址   
            aZipCode2 = colArray(14).ToString           '戶藉郵遞區號
            aZipCode2_6W = colArray(15).ToString        '戶藉郵遞區號6碼
            aHouseholdAddress = colArray(16).ToString   '戶藉地址
            aEmail = colArray(17).ToString              'Email 
            aMIdentityID = TIMS.ClearSQM(colArray(18))  '主要參訓身分別 varchar(3)
            aIdentityID = aMIdentityID                  '參訓身分別代碼	varchar(50)	 
            aUname = colArray(19).ToString              '公司名稱       nvarchar(50)
            aIntaxno = colArray(20).ToString            '服務單位統一編號 varchar(10)
            aQ4 = colArray(21).ToString                 'Q4服務單位行業別 
            aScale = colArray(22).ToString              '服務單位規模 
            aJobTitle = colArray(23).ToString           '職稱 
            aQ61 = colArray(24).ToString                '個人年資
            aPriorWorkPay = colArray(25).ToString       '受訓前薪資
            aActType = colArray(26).ToString            '投保類別  Char(1)
            aActname = colArray(27).ToString            '投保公司名稱 
            aActNo = colArray(28).ToString              '投保單位(公司)保險證號
            aAcctMode = colArray(29).ToString           '帳戶類別(郵政或金融)
            aPostNo = colArray(30).ToString             '郵政-局號  
            aAcctNo = colArray(31).ToString             '(郵政/銀行)帳號    
            aBankName = colArray(32).ToString           '銀行名稱
            aAcctHeadNo = colArray(33).ToString         '金融-總代號 
            aExBankName = colArray(34).ToString         '分行名稱
            aAcctExNo = colArray(35).ToString           '分行代碼
            aAcctNo2 = colArray(36).ToString            '(郵政/銀行)帳號    
            aQ2 = colArray(37).ToString                 'Q2參訓動機 
            aQ3 = colArray(38).ToString                 'Q3訓後動向       
            aQ3_Other = colArray(39).ToString           'Q3訓後動向其他   
            aIseMail = colArray(40).ToString            ' 是否願意收到職訓通知 
            'aIsAgree=colArray(41).ToString           '同意否
            aIsAgree = "Y"  '產投都帶同意
            aQ5 = colArray(41).ToString                 'Q5服務單位是否屬於中小企業 Bit
            aQ62 = colArray(42).ToString                '公司年資                Int
            aQ63 = colArray(43).ToString                '職位年資                Int
            aQ64 = colArray(44).ToString                '升遷年資                Int
            '**by Milor 20081016--加入投保單位電話與地址
            actTel = Convert.ToString(colArray(45))       '投保單位電話
            actZipCode = Convert.ToString(colArray(46))   '投保單位郵遞區號前三碼
            actZipCODE6W = Convert.ToString(colArray(47)) '投保單位郵遞區號6碼
            actAddress = Convert.ToString(colArray(48))   '投保單位地址
            aEnterDate = Common.FormatDate(aNow)          '報名日期
            If aEnterDate = "" Then
                Reason &= "必須填寫報名日期 <BR>"
            Else
                'If flag_ROC Then
                '    If TIMS.IsDate7(aEnterDate)=False Then
                '        Reason &= "填寫報名日期必須是民國年格式(yyy/MM/dd)<BR>"  'edit，by:20181018
                '    Else
                '        aEnterDate=Common.FormatDate(TIMS.cdate18(aEnterDate))  'edit，by:20181018
                '        If CDate(aEnterDate) < "1900/1/1" Or CDate(aEnterDate) > "2100/1/1" Then Reason &= "填寫報名日期範圍有誤<BR>"  'edit，by:20181018
                '    End If
                'End If
                If IsDate(aEnterDate) = False Then
                    Reason &= "填寫報名日期必須是西元年格式(yyyy/MM/dd)<BR>"
                Else
                    aEnterDate = Common.FormatDate(aEnterDate)
                    If CDate(aEnterDate) < "1900/1/1" Or CDate(aEnterDate) > "2100/1/1" Then Reason &= "填寫報名日期範圍有誤<BR>"
                End If
            End If


            '檢查縣市別是否正確
            'Dim subsql As String=""
            'Dim subdt As DataTable
            'subsql &= " select * from Class_ClassInfo cc JOIN ID_ZIP iz ON iz.ZipCode=cc.TaddressZip "
            'subsql &= " where 1=1 "
            'subsql &= " AND cc.OCID='" & aOCID1 & "'"
            'subsql &= " AND iz.CTID ='" & aCTID & "'"
            'subdt=DbAccess.GetDataTable(subsql)
            'If Not subdt.Rows.Count > 0 Then Reason &= "請選擇正確的縣市別代碼<BR>"
            'Dim drCC1 As DataRow=TIMS.GetOCIDDate(aOCID1, objconn)
            Dim fg_OCID As Boolean = False
            If aOCID1 <> "" AndAlso aOCID1 <> OCIDValue1.Value Then
                Reason &= "匯入報考班別代碼1與查詢資料不符 <BR>"
            ElseIf aOCID1 = "" Then
                Reason &= "必須填寫報考班別代碼1 <BR>"
            Else
                Dim MyKey As String = aOCID1
                If IsNumeric(MyKey) Then
                    MyKey = CInt(MyKey)
                    '**by Milor 20080512--班級如果存在順帶取出機構別，補充判斷縣市別代碼是否填寫-start
                    Dim s_FINDTMP As String = $"OCID='{MyKey}'"
                    Dim drCI As DataRow = Nothing
                    If dtClassInfo Is Nothing OrElse dtClassInfo.Select(s_FINDTMP).Length = 0 Then
                        Reason &= String.Concat("報名課程代碼有誤，(", MyKey, ")不符合鍵詞<BR>")
                    Else
                        drCI = dtClassInfo.Select(s_FINDTMP)(0)
                        orgKind = TIMS.Get_OrgKind2(MyKey, TIMS.c_OCID, objconn)
                        '20090507 by岡 班級開訓14天後，依規定不得再匯入該班學員資料-start
                        'If TIMS.Server_Path()="DEMO" Then
                        '    If Key_Class_ClassInfo.Select("OCID='" & MyKey & "' and ddiff>='14'").Length > 0 Then Reason &= "班級開訓14天後，依規定不得再匯入該班學員資料<BR>"
                        'End If
                        '20090507 by岡-end
                    End If
                    If aCTID = "" Then
                        Reason &= "必須填寫縣市別代碼<BR>"
                    Else
                        If IsNumeric(aCTID) Then
                            '(任1為真即可)
                            Dim ocidFlag0 As Boolean = False
                            Dim ocidFlag1 As Boolean = False
                            Dim ocidFlag2 As Boolean = False
                            If drCI IsNot Nothing Then
                                ocidFlag0 = (Convert.ToString(drCI("CTID")) = aCTID)
                                ocidFlag1 = (Convert.ToString(drCI("CTID1")) = aCTID)
                                ocidFlag2 = (Convert.ToString(drCI("CTID2")) = aCTID)
                            End If
                            fg_OCID = (ocidFlag0 OrElse ocidFlag1 OrElse ocidFlag2)
                            If Not fg_OCID Then Reason &= "縣市別代碼與班級位置不符<BR>"
                        Else
                            Reason &= "請選擇正確的縣市別代碼" & aCTID & "<BR>"
                        End If
                    End If
                    '**by Milor 20080512-end
                Else
                    Reason &= "報名課程代碼有誤，(" & MyKey & ")不符合鍵詞<BR>"
                End If
                aOCID1 = MyKey
            End If
            '**by Milor 20080529--本國籍要判斷身分證號正確性，外籍不用-start
            IDNOFlag = True
            If aPassPortNO = "" Then
                Reason &= "必須填寫身分別代碼 1:本國 2:外籍<BR>"
            Else
                Select Case aPassPortNO
                    Case "1"
                        '身分證驗証
                        aIDNO = TIMS.ChangeIDNO(aIDNO)
                        If aIDNO = "" Then
                            Reason &= "必須填寫身分證號碼<BR>"
                        Else
                            If TIMS.CheckIDNO(aIDNO) Then '一般驗証
                                If sm.UserInfo.RoleID = 1 Then '角色代碼為1 可執行安全性規則確認
                                    'Dim IDNOFlag As Boolean=True
                                    Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                                    If aIDNO.Length <> 10 Then
                                        IDNOFlag = False
                                    ElseIf aIDNO.Chars(1) <> "1" And aIDNO.Chars(1) <> "2" Then
                                        IDNOFlag = False
                                    ElseIf EngStr.IndexOf(aIDNO.ToUpper.Chars(0)) = -1 Then
                                        IDNOFlag = False
                                    ElseIf aIDNO = "A123456789" Then
                                        IDNOFlag = False
                                    End If
                                    If Not IDNOFlag Then Reason &= "身分證號碼錯誤!<BR>"
                                End If
                                If IDNOFlag = True Then 'Reason="" Then
                                    '=========== 驗証匯入檔案時不要有相同的身分證號碼 Start =============
                                    Dim Flag As Boolean = True
                                    For i As Integer = 0 To IDNOArray.Count - 1
                                        If IDNOArray(i) = aIDNO Then
                                            Reason &= "檔案中有相同的身分證號碼<BR>"
                                            Flag = False
                                        End If
                                    Next
                                    If Flag Then IDNOArray.Add(aIDNO)
                                    '=========== 驗証匯入檔案時不要有相同的身分證號碼 -End- =============
                                End If
                            Else
                                Reason &= "身分證號碼錯誤!請聯絡系統管理員<BR>"
                                IDNOFlag = False
                            End If
                        End If
                    Case "2" '外國籍只判斷是否有填IDNO
                        aIDNO = TIMS.ChangeIDNO(aIDNO)
                        If aIDNO = "" Then Reason &= "必須填寫(外國籍)身分證號碼<BR>"
                    Case Else
                        Reason &= "身分別代碼只能是1或者是2<BR>"
                End Select
            End If
            '當身分證號與課程代碼檢核都通過時，判斷要匯入的身分證號已經存在同一班。
            If fg_OCID AndAlso IDNOFlag Then
                sql = $"SELECT a.eSETID,a.IDNO,b.OCID1 FROM STUD_ENTERTEMP2 a JOIN STUD_ENTERTYPE2 b on a.eSETID=b.eSETID where a.IDNO ='{aIDNO}' and b.OCID1={aOCID1}"
                If DbAccess.GetCount(sql, objconn) > 0 Then Reason &= "身分證號碼已經存在該班級，請確認身分證號碼是否正確<BR>"
            End If
            '**by Milor 20080529-end
            If aBirthday = "" Then
                Reason &= "必須填寫出生日期<BR>"
            Else
                'If flag_ROC Then
                '    If TIMS.IsDate7(aBirthday)=False Then
                '        Reason &= "出生日期必須是西元年格式(yyyy/MM/dd)<BR>"  'edit，by:20181018
                '    Else
                '        aBirthday=Common.FormatDate(TIMS.cdate18(aBirthday))  'edit，by:20181018
                '        If CDate(aBirthday) < "1900/1/1" Or CDate(aBirthday) > "2100/1/1" Then Reason &= "出生日期範圍有誤<BR>"  'edit，by:20181018
                '    End If
                'End If
                If IsDate(aBirthday) = False Then
                    Reason &= "出生日期必須是西元年格式(yyyy/MM/dd)<BR>"
                Else
                    aBirthday = Common.FormatDate(aBirthday)
                    If CDate(aBirthday) < "1900/1/1" Or CDate(aBirthday) > "2100/1/1" Then Reason &= "出生日期範圍有誤<BR>"
                End If
            End If
            If aName = "" Then
                Reason &= "必須填寫中文姓名<BR>"
            Else
                If aName.Length > 30 Then Reason &= "中文姓名長度必須小於15<BR>"
            End If
            If aSex = "" Then
                Reason &= "必須填寫性別<BR>"
            Else
                Select Case aSex
                    Case "M", "男"
                        aSex = "M"
                    Case "F", "女"
                        aSex = "F"
                    Case Else
                        Reason &= "性別代號只能是M/男或者是F/女<BR>"
                End Select
            End If
            '**by Milor 20080530--外籍人士不進行性別與身分證號的判斷
            'PassPortNO:1:本國/2:外籍
            If (Mid(aIDNO, 2, 1) = "1" AndAlso aSex = "M") OrElse (Mid(aIDNO, 2, 1) = "2" And aSex = "F") OrElse aPassPortNO = "2" Then
            Else
                Reason &= "性別代號與身分證號碼不符<BR>"
            End If
            '最高學歷
            If aDegreeID = "" Then
                Reason &= "必須填寫最高學歷<BR>"
            Else
                Dim MyKey As String = aDegreeID
                If MyKey.Length < 2 Then MyKey = "0" & MyKey
                If dt_Key_Degree.Select("DegreeID='" & MyKey & "'").Length = 0 Then Reason &= "最高學歷代碼有錯，不符合鍵詞<BR>"
                aDegreeID = MyKey
            End If
            If aPhone1 = "" Then
                Reason &= "必須填寫聯絡電話(日)<BR>"
            Else
                If aPhone1.Length > 25 Then Reason &= "聯絡電話(日)長度大於儲存範圍(25)<BR>"
            End If
            If aPhone2 = "" Then
                'Reason &= "必須填寫聯絡電話(夜)<BR>"
            Else
                If aPhone2.Length > 25 Then Reason &= "聯絡電話(夜)長度大於儲存範圍(25)<BR>"
            End If
            If aCellPhone = "" Then
                'Reason &= "必須填寫聯絡電話(夜)<BR>"
            Else
                If aCellPhone.Length > 25 Then Reason &= "行動電話長度大於儲存範圍(25)<BR>"
            End If
            If aZipCode = "" Then
                Reason &= "必須填寫郵遞區號<BR>"
            Else
                If Len(aZipCode) = 3 Then
                    If aZipCode < "000" Or aZipCode > "999" Then
                        Reason &= "郵遞區號填寫有誤 <BR>"
                    Else
                        If dtZipCode.Select("ZipCode=" & aZipCode).Length = 0 Then Reason &= "郵遞區號填寫範圍有誤 <BR>"
                    End If
                Else
                    Reason &= "郵遞區號填寫有誤 <BR>"
                End If
            End If

            '郵遞區號 5碼或6碼
            Dim TMPERR1 As String = TIMS.CHK_ZIPCODE6W(aZipCODE6W, "通訊郵遞區號")
            If TMPERR1 <> "" Then Reason &= TMPERR1

            If aAddress = "" Then
                Reason &= "必須填寫通訊地址<BR>"
            Else
                If aAddress.Length > 100 Then Reason &= "通訊地址長度大於儲存範圍(100)<BR>"
            End If
            If aZipCode2 = "" Then
                Reason &= "必須填寫戶藉地址郵遞區號<BR>"
            Else
                If Len(aZipCode2) = 3 Then
                    If aZipCode2 < "000" OrElse aZipCode2 > "999" Then
                        Reason &= "戶藉地址郵遞區號填寫有誤 <BR>"
                    Else
                        If dtZipCode.Select("ZipCode=" & aZipCode2).Length = 0 Then Reason &= "戶藉地址郵遞區號填寫範圍有誤 <BR>"
                    End If
                Else
                    Reason &= "戶藉地址郵遞區號填寫有誤 <BR>"
                End If
            End If

            '郵遞區號 5碼或6碼
            TMPERR1 = TIMS.CHK_ZIPCODE6W(aZipCode2_6W, "戶藉地址郵遞區號")
            If TMPERR1 <> "" Then Reason &= TMPERR1

            If aHouseholdAddress = "" Then
                Reason &= "必須填寫戶藉地址<BR>"
            Else
                If aHouseholdAddress.Length > 100 Then Reason &= "戶藉地址長度大於儲存範圍(100)<BR>"
            End If
            '20090212 andy add email 格式檢查 start
            aEmail = TIMS.ChangeEmail(TIMS.ClearSQM(aEmail))
            If Len(aEmail) = 0 Then
                Reason &= "必須填寫Email<BR>"
            ElseIf aEmail <> "無" Then
                If aEmail.Length > 60 Then
                    Reason &= String.Format("Email長度大於儲存範圍(60,{0})<BR>", aEmail.Length)
                ElseIf Not TIMS.CheckEmail(aEmail) Then
                    Reason &= String.Format("Email格式有誤({0})<BR>", aEmail)
                End If
            End If
            '20090212 andy add email 格式檢查 end
            If aMIdentityID = "" Then
                Reason &= "必須填寫主要參訓身分別代碼<BR>"
            Else
                Dim MyKey As String = aMIdentityID
                If MyKey.Length < 2 Then MyKey = "0" & MyKey Else MyKey = MyKey
                If dt_Key_Identity.Select("IdentityID='" & MyKey & "'").Length = 0 Then Reason &= "主要參訓身分別代碼不符合鍵詞<BR>"
                aMIdentityID = MyKey
            End If
            '20090210 2009年 身分為「非自願離職者」不檢查投保相關資料 andy  edit   
            '2009年 「非自願離職者」不檢查投保相關資料 by AMU 20100414
            '2010年 取消「非自願離職者」 by AMU 20100414
            If Not (aMIdentityID = "02" AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 _
                AndAlso CInt(sm.UserInfo.Years) > 2008 AndAlso CInt(sm.UserInfo.Years) < 2010) Then
                DoDataValid = True '要檢查
            Else
                DoDataValid = False '不檢查
            End If
            If aUname <> "" Then
                If aUname.Length > 30 Then Reason &= "服務單位長度大於儲存範圍<BR>"
            Else
                If DoDataValid = True Then Reason &= "服務單位必須填寫<BR>" '20090211非自願離職者此欄位為非必填
            End If
            '**by Milor 20080512--PM應客戶要求，取消統編匯入必填的檢查-start 
            aIntaxno = TIMS.ClearSQM(aIntaxno)
            If aIntaxno <> "" Then
                If Not IsNumeric(aIntaxno) Then
                    Reason &= "填寫的服務單位統一編號有誤(統一編號皆為數字，不可含空白或任何符號)<BR>"
                Else
                    If aIntaxno.Length > 10 Then Reason &= "服務單位統一編號大於儲存範圍<BR>"
                End If
                'Reason &= "服務單位統一編號必須填寫<BR>"
            End If
            '**by Milor 20080512-end
            If aQ4 <> "" Then
                Dim MyKey As String = aQ4
                If DoDataValid = False Then   '20090211非自願離職者此欄位填[30]其他服務業
                    'If MyKey <> 30 Then Reason &= "主要參訓身分別代碼為[02]非自願離職者時,服務單位行業別代碼須填[30]其他服務業<BR>"  '20090312 andy 當主要參訓身分別為"非自願離職者","服務單位行業別"改為非必填
                Else
                    If MyKey.Length < 2 Then MyKey = "0" & MyKey Else MyKey = MyKey
                    If dt_Key_Trade.Select("TradeID='" & MyKey & "'").Length = 0 Then Reason &= "服務單位行業別代碼不符合鍵詞<BR>"
                    aQ4 = MyKey
                End If
            Else
                Reason &= "服務單位行業別 必須填寫<BR>"
            End If
            If aScale <> "" Then
                Select Case aScale
                    Case "1", "2", "3", "4"
                    Case Else
                        Reason &= "填寫服務單位規模範圍有誤<BR>"
                End Select
            End If
            If aJobTitle = "" Then
                'If DoDataValid=True Then Reason &= "必須填寫工作職稱 <BR>" '20090211非自願離職者此欄位為非必填
                '20090302 andy  改為非必填
            Else
                If aJobTitle.Length > 50 Then
                    Reason &= "工作職稱大於儲存範圍<BR>"
                End If
            End If
            aQ61 = TIMS.ClearSQM(aQ61)
            If aQ61 = "" Then
                'Reason &= "必須填寫工作年資(個人) <BR>"
                '20090302 andy  改為非必填
                aQ61 = ""
            Else
                If Not IsNumeric(aQ61) Then
                    Reason &= "工作年資(個人)必須填寫為數字 <BR>"
                Else
                    If CInt(aQ61).ToString <> aQ61.ToString Then Reason &= "工作年資(個人)必須填寫為整數數字 <BR>"
                End If
            End If
            '**by Milor 20080512--PM應客戶要求，取消受訓前薪資必填的檢查-start 
            'If aPriorWorkPay="" Then
            'Reason &= "必須填寫受訓前薪資 <BR>"
            'Else
            If Not aPriorWorkPay = "" Then
                If Not IsNumeric(aPriorWorkPay) Then
                    Reason &= "受訓前薪資必須填寫為數字 <BR>"
                Else
                    If CInt(aPriorWorkPay).ToString <> aPriorWorkPay.ToString Then Reason &= "受訓前薪資必須填寫為整數數字 <BR>"
                End If
            End If
            '**by Milor 20080512-end
            '**by Milor 20080512--加入帳戶類別2-訓練單位代轉現金的判斷-start
            aActType = TIMS.ClearSQM(aActType)
            If DoDataValid = False Then  '20090210 andy「非自願離職者」無投保類別資料 設為0,資料null
                aActType = "0"
            Else
                If aActType.Length <> 1 Then
                    Reason &= "投保類別填寫有誤, 必須填寫為個位數字<BR>"
                Else
                    Select Case aActType
                        Case 1, 2
                        Case Else
                            Reason &= "投保類別填寫有誤, 必須填寫為個位數字(1.勞2.農)<BR>"
                    End Select
                End If
            End If
            aAcctMode = TIMS.ClearSQM(aAcctMode)
            If aAcctMode <> "" Then
                Select Case aAcctMode.Substring(0, 1)
                    Case "0", "郵"
                        aAcctMode = 0
                    Case "1", "金", "銀"
                        aAcctMode = 1
                    Case "2", "代"
                        aAcctMode = 2
                    Case Else
                        Reason &= "帳戶類別(0:郵政1:金融2:訓練單位代轉現金)超過鍵詞範圍<BR>"
                End Select
            End If
            Select Case aAcctMode
                Case 0
                    If aPostNo = "" Then Reason &= "必須填寫郵政局號<BR>"
                    If aAcctNo = "" Then Reason &= "必須填寫郵政帳號<BR>"
                Case 1
                    If aBankName = "" Then Reason &= "必須總行名稱(銀行)<BR>"
                    If aAcctHeadNo = "" Then Reason &= "必須填寫總行代號<BR>"
                    If aExBankName = "" Then Reason &= "必須填寫分行名稱(銀行)<BR>"
                    If aAcctExNo = "" Then Reason &= "必須填寫分行代號<BR>"
                    If aAcctNo2 = "" Then Reason &= "必須填寫銀行帳號<BR>"
                Case 2
                    '因為已經限定只能匯入產學訓的班級，所以不需要再判斷是否為產學訓
                    If orgKind <> "W" Then Reason &= "非勞工團體不能使用帳戶類別-2訓練單位代轉現金<BR>"
                Case Else
                    Reason &= "必須填寫帳戶類別(0:郵政1:金融2:訓練單位代轉現金)<BR>"
            End Select
            '**by Milor 20080512-end
            '參訓動機代碼
            If aQ2 = "" Then
                Reason &= "必須填寫參訓動機代碼 <BR>"
            Else
                Select Case aQ2
                    Case "1", "2", "3", "4"
                    Case Else
                        Reason &= "參訓動機代碼只能是1.2.3.4<BR>"
                End Select
            End If
            If aQ3 = "" Then
                'Reason &= "必須填寫訓後動向代碼 <BR>"
                '20090302 andy 改為非必填
                aQ3 = ""
            Else
                Select Case aQ3
                    Case "1", "2", "3"
                    Case Else
                        Reason &= "訓後動向代碼只能是1.2.3<BR>"
                End Select
            End If
            If aQ3_Other <> "" Then
                If aQ3_Other.Length > 50 Then Reason &= "訓後動向其他長度大於儲存範圍<BR>"
            End If
            Select Case aIseMail
                Case "Y", "N"
                Case Else
                    Reason &= "希望收到最新課程資訊只能是Y或者是N<BR>"
            End Select
            'If aIsAgree="" Then Reason &= "必須填寫是否同意個人基本資料供所屬機關運用 <BR>"
            Select Case aIsAgree
                Case "Y", "N"
                Case Else
                    Reason &= "同意所屬機關使用本人資料只能是Y或者是N<BR>"
            End Select
            aQ5 = TIMS.ClearSQM(aQ5)
            If aQ5 = "" Then
                'Reason &= "必須填寫服務單位是否屬於中小企業 <BR>"
                '20090302 andy  改為非必填
                aQ5 = "0"
            Else
                If DoDataValid = False Then   '20090210 [02]非自願離職者 填「N」
                    If aQ5 = "Y" Then Reason &= "服務單位是否屬於中小企業 身分別為 [02]非自願離職者 請填「N」<BR>"
                End If
                If aQ5 = "Y" Then
                    aQ5 = "1"
                Else
                    aQ5 = "0"
                End If
            End If
            aQ62 = TIMS.ClearSQM(aQ62)
            If aQ62 = "" Then
                'If DoDataValid=True Then Reason &= "必須填寫公司年資 <BR>" '20090210 [02]非自願離職者此欄位非必填
                '20090302 andy  改為非必填
            Else
                If Not IsNumeric(aQ62) Then
                    Reason &= "公司年資必須填寫為數字 <BR>"
                Else
                    If CInt(aQ62).ToString <> aQ62.ToString Then Reason &= "公司年資必須填寫為整數數字 <BR>"
                End If
            End If
            aQ63 = TIMS.ClearSQM(aQ63)
            If aQ63 = "" Then
                'If DoDataValid=True Then Reason &= "必須填寫職位年資<BR>" '20090210 [02]非自願離職者此欄位非必填
                '20090302 andy  改為非必填
            Else
                If Not IsNumeric(aQ63) Then
                    Reason &= "職位年資必須填寫為數字 <BR>"
                Else
                    If CInt(aQ63).ToString <> aQ63.ToString Then Reason &= "職位年資必須填寫為整數數字 <BR>"
                End If
            End If
            aQ64 = TIMS.ClearSQM(aQ64)
            If aQ64 = "" Then
                'If DoDataValid=True Then Reason &= "必須填寫升遷離本職幾年 <BR>" '20090210 [02]非自願離職者此欄位非必填
                '20090302 andy  改為非必填
            Else
                If Not IsNumeric(aQ64) Then
                    Reason &= "升遷離本職幾年必須填寫為數字 <BR>"
                Else
                    If CInt(aQ64).ToString <> aQ64.ToString Then Reason &= "升遷離本職幾年必須填寫為整數數字 <BR>"
                End If
            End If
            '**by Milor 20081016--加入投保單位電話、地址
            If actTel = "" Then
                If DoDataValid = True Then Reason &= "必須填寫投保單位電話 <BR>" '20090210 [02]非自願離職者此欄位非必填
            End If
            actZipCode = TIMS.ClearSQM(actZipCode)
            If actZipCode = "" Then
                If DoDataValid = True Then Reason &= "必須填寫投保單位郵遞區號前三碼 <BR>" '20090210 [02]非自願離職者此欄位非必填
            Else
                If Not IsNumeric(actZipCode) Then Reason &= "投保單位郵遞區號前三碼必須為數字 <BR>"
                If Len(actZipCode) > 3 Then Reason &= "投保單位郵遞區號前三碼，只能為三碼數字 <BR>"
            End If

            '投保單位郵遞區號 5碼或6碼
            TMPERR1 = TIMS.CHK_ZIPCODE6W(actZipCODE6W, "投保單位郵遞區號")
            If TMPERR1 <> "" Then Reason &= TMPERR1

            If actAddress = "" Then
                If DoDataValid = True Then Reason &= "必須填寫投保單位地址 <BR>" '20090210 [02]非自願離職者此欄位非必填
            End If
        End If
        Return Reason
    End Function

    ''' <summary>
    ''' 是否己過甄試日期
    ''' </summary>
    ''' <param name="str_OCID"></param>
    ''' <param name="ExamDate"></param>
    ''' <returns></returns>
    Function CheckExamDate(ByVal str_OCID As String, ByVal ExamDate As String) As Boolean
        Dim rst As Boolean = False '表示己過甄試日期,不能修改 'Dim sql As String="" 'Dim dr As DataRow
        ExamDate = TIMS.ClearSQM(ExamDate)
        str_OCID = TIMS.ClearSQM(str_OCID)
        If str_OCID <> "" Then
            If ExamDate.ToString <> "" Then
                Dim pms1 As New Hashtable From {{"OCID", CInt(str_OCID)}}
                Dim sql As String = " SELECT 'x' x FROM dbo.CLASS_CLASSINFO cc WHERE cc.ExamDate>GETDATE() AND cc.OCID=@OCID"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms1) 'rst=False '表示己過甄試日期,不能修改
                If dr IsNot Nothing Then rst = True '表示未過甄試日期可以修改
            ElseIf ExamDate = "" Then
                Dim pms1 As New Hashtable From {{"OCID", CInt(str_OCID)}}
                Dim sql As String = " SELECT OCID FROM dbo.STUD_SELRESULT WHERE OCID=@OCID"
                Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pms1) 'rst=False '表示己有試算檔的資料,不能修改
                If dr Is Nothing Then rst = True '表示沒有試算檔的資料可以修改
            End If
        End If
        Return rst
    End Function

#End Region

    'confidential information
    'Dim flgCIShow As Boolean=False '是否可正常顯示個資。
    Dim drOCID As DataRow '班級資料查詢。
    Dim aNow As Date
    Dim objconn As SqlConnection

    '若有調整 SD_01_004之e網報名者待審核，煩請一併調整位置 主頁搜尋條件。
    Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        AddHandler Button1.Click, AddressOf sUtl_btnSearchData1 '查詢  
        AddHandler Button13.Click, AddressOf sUtl_btnSearchData1 '匯出 Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
        AddHandler btndivPwdSubmit.Click, AddressOf sUtl_btnSearchData1
        'If TIMS.sUtl_ChkTest() Then Hid_oTest.Value="Y" 'TIMS.sUtl_ChkTest()
        'If TIMS.sUtl_ChkTest() Then oTest_flag=True
        Hid_TV1_MSG1S.Value = TIMS.Cst_TV1_MSG1S '尼伯特颱風臺東地區受災者
        Hid_TV1_MSG2S.Value = TIMS.Cst_TV1_MSG2S '尼伯特颱風受災者,為屏東地區或臺南市七股區民眾
        Hid_TV2_MSG3S.Value = TIMS.Cst_TV2_MSG3S '梅姬颱風受災者
        Hid_DIS2ALARMMSG1.Value = TIMS.Cst_DIS_MSG1S '屬於重大災害受災地區範圍
        Dim work2015 As String = TIMS.Utl_GetConfigSet("work2015")
        hidLockTime2.Value = "2"
        If work2015 = "Y" Then hidLockTime2.Value = "1" '啟用鎖定。
        str_addedit_aspx = "SD_01_004_add.aspx?ID=" & TIMS.Get_MRqID(Me) 'Request("ID") '"SD_01_004_add.aspx?ID=" & Request("ID")

        'TPLANID28_TR1.Visible=False '28:產業人才投資計劃-暫不啟用匯入
        'BtnImport28.Enabled=False '28:產業人才投資計劃-暫不啟用匯入

        '啟動個資法。
        Button1.Attributes.Add("onclick", "return showLoginPwdDiv(1);")
        Button13.Attributes.Add("onclick", "return showLoginPwdDiv(2);")
        Button1.CommandName = cst_btn_button1 '"Button1"
        Button13.CommandName = cst_btn_button13 '"Button13"
        Button3.Attributes.Add("onclick", "return confirm('確認是否儲存資料?');")

        PageControler1.PageDataGrid = DataGrid1

        Call TIMS.OpenDbConn(objconn)
        aNow = TIMS.GetSysDateNow(objconn)
        '檢查Session是否存在 End

        '非 ROLEID=0 LID=0 'Dim flgROLEIDx0xLIDx0 As Boolean=False '判斷登入者的權限。
        '如果是系統管理者開啟功能。'ROLEID=0 LID=0 '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1) 'False 'If TIMS.IsSuperUser(Me, 1) Then flgROLEIDx0xLIDx0=True

        '是否開啟民國年日期顯示機制(是true/否false)
        flag_ROC = TIMS.CHK_REPLACE2ROC_YEARS()

        If Not Page.IsPostBack Then
            Call cCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '28:產業人才投資計劃
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Label1.Text = cst_tims28_label1
            LabSubsidyCost.Text = cst_tims28_labsubsidycost
            Label2.Text = cst_tims28_label2
        End If
        'BtnImport28.Attributes("onClick")="checkSizeA();"
        'onChange="CheckSize(this.value)"

        '090914 當班級代碼欄位有變更時檢查登入帳號是否有授權
        '※目前匯入名冊功能平時不開放，只提供給已結訓班級使用授權檔內有指定之授權帳號使用()
        '請勿mark掉此段程式及 javascript function  chkOCID(){ __doPostBack('LinkButton2','');}	
        '-start
        'BtnImport28.Enabled=False
        'TIMS.Tooltip(BtnImport28, "停用匯入功能。")
        'If chkAcctRight(sm.UserInfo.UserID, OCIDValue1.Value) Then BtnImport28.Enabled=True
        'If BtnImport28.Enabled=False Then
        '    'TIMS.Tooltip(BtnImport28, "已提供線上報名功能，匯入名冊功能停止使用(授權檔內有指定之授權帳號(補登使用))", True)
        '    TIMS.Tooltip(BtnImport28, "補登使用(授權檔內有指定之授權帳號,且在補登期間)", True)
        'End If
        '- end
        'If TIMS.Server_Path="DEMO" Then
        '    BtnImport28.Enabled=True
        '    TIMS.Tooltip(BtnImport28, "測試機暫時開放") ' by AMU 20100329
        'End If

        '確認機構是否為黑名單
        Dim vsMsg2 As String = "" '確認機構是否為黑名單
        'vsMsg2=""
        If Chk_OrgBlackList(vsMsg2) Then
            Button3.Enabled = False
            TIMS.Tooltip(Button3, vsMsg2)
            Button9.Enabled = False
            TIMS.Tooltip(Button9, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    Sub cCreate1()
        divEnterDouble.Visible = False
        divEnterMoney.Visible = False
        dtgAddresses1.Visible = False '匯入名冊產投用(檢視)報名日期 迄止日期格式有誤
        'LinkButton1.Visible=False '測試寄送信件

        TPLANID28_TR1.Visible = False '28:產業人才投資計劃-暫不啟用匯入
        BtnImport28.Enabled = False '28:產業人才投資計劃-暫不啟用匯入
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then
            '依計畫改變名稱
            labPlanTxt.Text = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        End If
        'BtnImport28.Enabled=False '非產業人才投資計劃(匯入e網報名名冊(產業人才投資)) 'TPLANID28_TR1.Visible=False

        '取出鍵詞-查詢原因
        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        '54:充電起飛計畫（在職）判斷方式
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            TPLANID28_TR1.Visible = True
            'Select Case sm.UserInfo.LID '委訓單位 開放
            '    Case "2" '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            '    Case Else
            '        TPLANID28_TR1.Visible=True
            'End Select
        End If
        If Not BtnImport28.Enabled Then TIMS.Tooltip(BtnImport28, "查詢有效匯入班級，即可啟用!", True)

        '查詢參訓歷史 'open_StudentList 'SD_05_010'Session(cst_ssDDCHKVIEW)=Nothing '登入後就永不清除SESSION。  
        Dim rqID As String = TIMS.Get_MRqID(Me)
        Button9.Attributes.Add("onclick", $"return open_StudentList('{Button1.ClientID}','{rqID}');")
        DataGridTable.Visible = False
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        end_date.Text = Common.FormatDate(aNow)
        end_date.Text = If(flag_ROC, TIMS.Cdate17(end_date.Text), Common.FormatDate(aNow))

        Call TIMS.GET_OnlyOne_OCID3(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
        'If sm.UserInfo.LID <> "2" Then
        '    TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        'Else
        '    Button7_Click(sender, e)
        'End If
        Call UseKeepSearch()

        Hid_PreUseLimited18a.Value = ""
        If TIMS.Cst_TPlanID_PreUseLimited18a.IndexOf(sm.UserInfo.TPlanID) > -1 Then Hid_PreUseLimited18a.Value = "Y"
    End Sub


    '機構黑名單內容(訓練單位處分功能)
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            Me.isBlack.Value = "Y"
            Me.Blackorgname.Value = sm.UserInfo.OrgName
            'btnAdd.Visible=False 'Button8.Visible=False
        End If
        Return rst
    End Function

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Select Case e.CommandName
            Case "view" '檢視
                'btn.CommandArgument="&eSerNum=" & drv("eSerNum") & "&BudID=" & BudID.SelectedValue & "&SupplyID=" & SupplyID.SelectedValue
                Dim eSerNum As String = TIMS.GetMyValue(e.CommandArgument, "eSerNum")
                Dim BudID As String = TIMS.GetMyValue(e.CommandArgument, "BudID")
                Dim SupplyID As String = TIMS.GetMyValue(e.CommandArgument, "SupplyID")
                Dim strMsgBox As String = ""
                If Val(eSerNum) = 0 Then
                    strMsgBox = "操作失敗，資料有誤!!"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                End If

                '-start
                Dim eSerNumX1 As String = String.Concat("'", eSerNum, "'")
                Dim strDDCHKVIEW As String = Convert.ToString(Session(cst_ssDDCHKVIEW))
                If Session(cst_ssDDCHKVIEW) IsNot Nothing Then strDDCHKVIEW = Convert.ToString(Session(cst_ssDDCHKVIEW))
                '沒有才做事
                If strDDCHKVIEW <> "" Then
                    If strDDCHKVIEW.IndexOf(eSerNumX1) = -1 Then strDDCHKVIEW &= String.Concat(",", eSerNumX1)
                Else
                    strDDCHKVIEW &= eSerNumX1 '沒有才做事
                End If
                Session(cst_ssDDCHKVIEW) = strDDCHKVIEW '儲存
                '-end

                Dim eSETID As String = ""
                Dim tmpNAME As String = ""
                Dim tmpIDNO As String = ""
                Dim tmpOCID As String = ""

                Dim sql As String = ""
                'sql="" & vbCrLf
                sql &= " SELECT a.SETID ,a.IDNO ,a.NAME ,b.OCID1 ,b.eSerNum ,b.eSETID ,b.SignUpStatus" & vbCrLf
                sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.eSETID=a.eSETID" & vbCrLf
                sql &= " WHERE b.eSerNum=@eSerNum" & vbCrLf
                Call TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("eSerNum", SqlDbType.Int).Value = Val(eSerNum)
                    dt1.Load(.ExecuteReader())
                End With
                '只能有1筆資料其他是錯誤。
                If dt1.Rows.Count <> 1 Then
                    strMsgBox = "操作失敗，資料有誤!!"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                End If

                tmpNAME = TIMS.HtmlDecode1(Convert.ToString(dt1.Rows(0)("NAME")))
                tmpIDNO = UCase(Convert.ToString(dt1.Rows(0)("IDNO")))
                tmpOCID = Convert.ToString(dt1.Rows(0)("OCID1"))
                'tSignUpStatus=Convert.ToString(dt1.Rows(0)("SignUpStatus"))
                'tmpeSerNum=Convert.ToString(dt1.Rows(0)("eSerNum"))
                eSETID = Convert.ToString(dt1.Rows(0)("eSETID"))
                'tmpSETID=Convert.ToString(dt1.Rows(0)("SETID"))

                sMemo = String.Concat("&動作=檢視", "&NAME=", tmpNAME, "&IDNO=", tmpIDNO)
                Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip1, tmpOCID, sMemo)

                Call KeepSearch()
                'btn.CommandArgument="&eSerNum=" & drv("eSerNum") & "&BudID=" & BudID.SelectedValue & "&SupplyID=" & SupplyID.SelectedValue
                Dim MyValue As String = ""
                TIMS.SetMyValue(MyValue, "eSETID", eSETID)
                TIMS.SetMyValue(MyValue, "eSerNum", eSerNum)
                TIMS.SetMyValue(MyValue, "BudID", BudID)
                TIMS.SetMyValue(MyValue, "SupplyID", SupplyID)
                TIMS.SetMyValue(MyValue, "IDNO", tmpIDNO)
                TIMS.SetMyValue(MyValue, "OCID1", tmpOCID)
                Session(TIMS.gcst_rblWorkMode) = TIMS.cst_wmdip1 'rblWorkMode.SelectedValue
                'TIMS.SetMyValue(MyValue, "rblWorkMode", rblWorkMode.SelectedValue)
                'Session("SD01004addMyValue")=MyValue
                TIMS.Utl_Redirect1(Me, str_addedit_aspx & MyValue)

            Case "del" '刪除
                Dim tSignUpStatus As String = ""
                Dim tmpNAME As String = ""
                Dim tmpIDNO As String = ""
                Dim tmpOCID As String = ""
                Dim tmpeSerNum As String = ""
                Dim tmpeSETID As String = ""
                Dim tmpSETID As String = ""
                Dim strMsgBox As String = ""
                If Val(e.CommandArgument) = 0 Then
                    strMsgBox = "刪除失敗，資料有誤!!"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                End If

                Dim sql As String = ""
                'sql="" & vbCrLf
                sql &= " SELECT a.SETID ,a.IDNO ,A.NAME ,b.OCID1 ,b.eSerNum ,b.eSETID ,b.SignUpStatus" & vbCrLf
                sql &= " FROM STUD_ENTERTEMP2 a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN STUD_ENTERTYPE2 b WITH(NOLOCK) ON b.eSETID=a.eSETID" & vbCrLf
                sql &= " WHERE b.eSerNum=@eSerNum" & vbCrLf
                Call TIMS.OpenDbConn(objconn)
                Dim dt1 As New DataTable
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("eSerNum", SqlDbType.Int).Value = Val(e.CommandArgument)
                    dt1.Load(.ExecuteReader())
                End With
                If dt1.Rows.Count <> 1 Then
                    strMsgBox = "刪除失敗，資料有誤!!"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                Else
                    tmpNAME = TIMS.HtmlDecode1(Convert.ToString(dt1.Rows(0)("NAME")))
                    tmpIDNO = UCase(Convert.ToString(dt1.Rows(0)("IDNO")))
                    tmpOCID = Convert.ToString(dt1.Rows(0)("OCID1"))
                    tSignUpStatus = Convert.ToString(dt1.Rows(0)("SignUpStatus"))
                    tmpeSerNum = Convert.ToString(dt1.Rows(0)("eSerNum"))
                    tmpeSETID = Convert.ToString(dt1.Rows(0)("eSETID"))
                    tmpSETID = Convert.ToString(dt1.Rows(0)("SETID"))
                End If

                '寫入Log查詢 (Auth_Accountlog)
                sMemo = String.Concat("&動作=刪除", "&NAME=", tmpNAME, "&IDNO=", tmpIDNO)
                Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm刪除, TIMS.cst_wmdip1, tmpOCID, sMemo)

                If tmpIDNO <> "" And tmpOCID <> "" Then
                    If Check_Student(tmpIDNO, tmpOCID) = True Then
                        strMsgBox &= $"此報名學員({tmpNAME})己有班級學員資料，不能刪除，請先刪除學員資料。{vbCrLf}"
                    Else
                        '因為目前Type2的KEY不見得一定會被存到Type中，所以常常導致Type沒有被清除，
                        '而讓報名登錄或參訓在Join的時候出現兩筆以上的紀錄，所以改為廣泛定義，
                        '同一個人，不應該存在兩筆以上同一班的Type資料，一但發生，則所有此人此班的資料將一併被刪除。
                        Dim sqlStr As String = ""
                        'sqlStr=""
                        sqlStr &= " SELECT a.SETID ,a.IDNO ,a.NAME ,a.SEX ,a.BIRTHDAY ,a.PASSPORTNO ,a.MARITALSTATUS" & vbCrLf
                        sqlStr &= " ,a.DEGREEID ,a.GRADID ,a.SCHOOL ,a.DEPARTMENT ,a.MILITARYID ,a.ZIPCODE" & vbCrLf
                        sqlStr &= " ,a.ADDRESS ,a.PHONE1 ,a.PHONE2 ,a.CELLPHONE ,a.EMAIL" & vbCrLf
                        sqlStr &= " ,a.NOTES ,a.ISAGREE ,a.LAINFLAG ,a.MODIFYACCT ,a.MODIFYDATE ,a.ESETID ,a.ZIPCODE6W" & vbCrLf
                        sqlStr &= " ,b.Q2_5OTHER ,b.Q2_3 ,b.Q2_4 ,b.MODIFYACCT ,b.MODIFYDATE ,b.TICKET_NO ,b.RELENTERDATE ,b.NOTEXAM ,b.CCLID ,b.ESETID ,b.ESERNUM" & vbCrLf
                        sqlStr &= " ,b.TRANSDATE ,b.SEID ,b.SUPPLYID ,b.BUDID ,b.ENTERPATH ,b.HIGHEDUBG ,b.WORKSUPPIDENT" & vbCrLf
                        sqlStr &= " ,b.USERNOSHOW ,b.NOTES ,b.PRIORWORKTYPE1 ,b.PRIORWORKORG1 ,b.SOFFICEYM1 ,b.FOFFICEYM1" & vbCrLf
                        sqlStr &= " ,b.ACTNO ,b.SETID ,b.ENTERDATE ,b.SERNUM ,b.EXAMNO ,b.OCID1 ,b.TMID1 ,b.OCID2 ,b.TMID2" & vbCrLf
                        sqlStr &= " ,b.OCID3 ,b.TMID3 ,b.WRITERESULT ,b.ORALRESULT ,b.TOTALRESULT ,b.ENTERCHANNEL" & vbCrLf
                        sqlStr &= " ,b.IDENTITYID ,b.RID ,b.PLANID ,b.TRNDMODE ,b.TRNDTYPE ,b.Q1_1 ,b.Q1_2 ,b.Q1_2OTHER" & vbCrLf
                        sqlStr &= " ,b.Q1_3 ,b.Q1_3OTHER ,b.Q1_4 ,b.Q1_4OTHER ,b.Q1_5" & vbCrLf
                        sqlStr &= " FROM STUD_ENTERTEMP a" & vbCrLf
                        sqlStr &= " JOIN STUD_ENTERTYPE b ON b.SETID=a.SETID" & vbCrLf
                        sqlStr &= " WHERE a.IDNO=@IDNO AND b.OCID1=@OCID1" & vbCrLf
                        Dim oCmd2 As New SqlCommand(sqlStr, objconn)
                        Dim dt2 As New DataTable
                        With oCmd2
                            .Parameters.Clear()
                            .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = tmpIDNO
                            .Parameters.Add("@OCID1", SqlDbType.Int).Value = Val(tmpOCID)
                            dt2.Load(.ExecuteReader())
                        End With

                        Dim tmpConn As SqlConnection = DbAccess.GetConnection()
                        Dim tmpTrans As SqlTransaction = DbAccess.BeginTrans(tmpConn)
                        Try
                            '未審核與報名失敗的資料，只進行Type2的刪除，因為理論上而言，未審核與報名失敗都不該有Type的資料，
                            '如果以Type2廣泛的條件進行刪除Type時，就會導致原本不該刪的資料被刪除。
                            If tSignUpStatus <> "0" AndAlso tSignUpStatus <> "2" Then
                                If dt2.Rows.Count > 0 Then
                                    '判斷如果學員是從三合一帶近來的則把職訓卷記錄檔,學習卷記錄檔,推介記錄檔的 transToTIMS 改成'N'
                                    For Each dr_T1 As DataRow In dt2.Rows
                                        'Select Case Convert.ToString(dr_T1("TRNDMode"))
                                        '    Case "1"
                                        '        sql="UPDATE Adp_TRNData SET TransToTIMS='N' WHERE TICKET_NO='" & Convert.ToString(dr_T1("TICKET_NO")) & "'"
                                        '        DbAccess.ExecuteNonQuery(sql, objconn)
                                        '    Case "2"
                                        '        sql="UPDATE Adp_DGTRNData SET TransToTIMS='N' WHERE TICKET_NO='" & Convert.ToString(dr_T1("TICKET_NO")) & "'"
                                        '        DbAccess.ExecuteNonQuery(sql, objconn)
                                        '    Case "3"
                                        '        sql="UPDATE Adp_GOVTRNData SET TransToTIMS='N' WHERE TICKET_NO='" & Convert.ToString(dr_T1("TICKET_NO")) & "'"
                                        '        DbAccess.ExecuteNonQuery(sql, objconn)
                                        'End Select
                                        DEL_STUDENTERTYPE(dr_T1("SETID"), dr_T1("OCID1"), sm, tmpTrans)
                                        DEL_STUDSELRESULT(dr_T1("SETID"), dr_T1("OCID1"), sm, tmpTrans)
                                    Next
                                    DEL_STUDENTERTRAIN2(tmpeSerNum, sm, tmpTrans)
                                    DEL_STUDENTERTYPE2(tmpeSerNum, tmpeSETID, sm, tmpTrans)
                                    strMsgBox = "刪除成功!"
                                Else
                                    If Convert.ToString(tmpSETID) <> "" Then DEL_STUDSELRESULT(tmpSETID, tmpOCID, sm, tmpTrans)
                                    DEL_STUDENTERTRAIN2(tmpeSerNum, sm, tmpTrans)
                                    DEL_STUDENTERTYPE2(tmpeSerNum, tmpeSETID, sm, tmpTrans)
                                    strMsgBox = "刪除成功!"
                                End If
                            Else
                                DEL_STUDENTERTRAIN2(tmpeSerNum, sm, tmpTrans)
                                DEL_STUDENTERTYPE2(tmpeSerNum, tmpeSETID, sm, tmpTrans)
                                strMsgBox = "刪除成功!"
                            End If
                            DbAccess.CommitTrans(tmpTrans)

                        Catch ex As Exception
                            DbAccess.RollbackTrans(tmpTrans)
                            DbAccess.CloseDbConn(tmpConn)

                            Common.MessageBox(Me, ex.ToString)
                            Throw ex
                        End Try
                        DbAccess.CloseDbConn(tmpConn)

                        Common.MessageBox(Me, strMsgBox)
                    End If
                End If
                'Button1_Click(Button1, e)
                Call Search1()

            Case "rev" '還原
                Dim veSerNum As String = TIMS.ClearSQM(e.CommandArgument)
                Dim tmpIDNO As String = ""
                Dim revFlag As Boolean = False
                Dim strMsgBox As String = ""
                Dim dr1 As DataRow = Nothing
                If veSerNum <> "" AndAlso TIMS.IsNumberStr(veSerNum) Then
                    Dim PMS_s1 As New Hashtable From {{"eSerNum", Val(veSerNum)}}
                    Dim sql_s1 As String = ""
                    sql_s1 &= " SELECT sp.IDNO, st.* "
                    sql_s1 &= " FROM STUD_ENTERTEMP2 sp "
                    sql_s1 &= " JOIN STUD_ENTERTYPE2 st ON sp.esetid=st.esetid "
                    sql_s1 &= " WHERE st.eSerNum =@eSerNum"
                    dr1 = DbAccess.GetOneRow(sql_s1, objconn, PMS_s1)
                End If
                If veSerNum = "" Then
                    strMsgBox = "資料錯誤，無法還原報名狀態!(查無資料)"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                ElseIf dr1 Is Nothing Then
                    strMsgBox = "資料錯誤，無法還原報名狀態!!(查無資料)"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                End If
                If Check_Student(dr1("IDNO").ToString, dr1("OCID1")) Then
                    strMsgBox = "此報名學員己有班級學員資料，不能還原報名狀態!!"
                    Common.MessageBox(Me, strMsgBox)
                    Exit Sub
                End If
                tmpIDNO = Convert.ToString(dr1("IDNO"))

                Dim tmpConn As SqlConnection = DbAccess.GetConnection()
                Dim tmpTrans As SqlTransaction = DbAccess.BeginTrans(tmpConn)
                Try
                    '未審核與報名失敗的資料，只進行Type2的刪除，因為理論上而言，未審核與報名失敗都不該有Type的資料，
                    '如果以Type2廣泛的條件進行刪除Type時，就會導致原本不該刪的資料被刪除。
                    If Convert.ToString(dr1("SignUpStatus")) <> "0" And Convert.ToString(dr1("SignUpStatus")) <> "2" Then
                        '因為目前Type2的KEY不見得一定會被存到Type中，所以常常導致Type沒有被清除，
                        '而讓報名登錄或參訓在Join的時候出現兩筆以上的紀錄，所以改為廣泛定義，
                        '同一個人，不應該存在兩筆以上同一班的Type資料，一但發生，則所有此人此班的資料將一併被刪除。
                        Dim sql_t As String = ""
                        sql_t &= " SELECT a.SETID ,a.IDNO ,a.NAME ,a.SEX ,a.BIRTHDAY ,a.PASSPORTNO ,a.MARITALSTATUS" & vbCrLf
                        sql_t &= " ,a.DEGREEID ,a.GRADID ,a.SCHOOL ,a.DEPARTMENT ,a.MILITARYID ,a.ZIPCODE" & vbCrLf
                        sql_t &= " ,a.ADDRESS ,a.PHONE1 ,a.PHONE2 ,a.CELLPHONE ,a.EMAIL" & vbCrLf
                        sql_t &= " ,a.NOTES ,a.ISAGREE ,a.LAINFLAG ,a.MODIFYACCT ,a.MODIFYDATE ,a.ESETID ,a.ZIPCODE6W" & vbCrLf
                        sql_t &= " ,b.Q2_5OTHER ,b.Q2_3,b.Q2_4 ,b.MODIFYACCT ,b.MODIFYDATE ,b.TICKET_NO ,b.RELENTERDATE ,b.NOTEXAM ,b.CCLID ,b.ESETID ,b.ESERNUM" & vbCrLf
                        sql_t &= " ,b.TRANSDATE ,b.SEID ,b.SUPPLYID ,b.BUDID ,b.ENTERPATH ,b.HIGHEDUBG ,b.WORKSUPPIDENT" & vbCrLf
                        sql_t &= " ,b.USERNOSHOW ,b.NOTES ,b.PRIORWORKTYPE1 ,b.PRIORWORKORG1 ,b.SOFFICEYM1 ,b.FOFFICEYM1" & vbCrLf
                        sql_t &= " ,b.ACTNO ,b.SETID ,b.ENTERDATE ,b.SERNUM ,b.EXAMNO ,b.OCID1 ,b.TMID1 ,b.OCID2 ,b.TMID2" & vbCrLf
                        sql_t &= " ,b.OCID3 ,b.TMID3 ,b.WRITERESULT ,b.ORALRESULT ,b.TOTALRESULT ,b.ENTERCHANNEL" & vbCrLf
                        sql_t &= " ,b.IDENTITYID ,b.RID ,b.PLANID ,b.TRNDMODE ,b.TRNDTYPE ,b.Q1_1 ,b.Q1_2 ,b.Q1_2OTHER" & vbCrLf
                        sql_t &= " ,b.Q1_3 ,b.Q1_3OTHER ,b.Q1_4 ,b.Q1_4OTHER ,b.Q1_5" & vbCrLf
                        sql_t &= " FROM STUD_ENTERTEMP a" & vbCrLf
                        sql_t &= " JOIN STUD_ENTERTYPE b ON b.SETID=a.SETID" & vbCrLf
                        sql_t &= " WHERE a.IDNO=@IDNO AND b.OCID1=@OCID" & vbCrLf
                        Dim objdt As New DataTable
                        Dim sCmd As New SqlCommand(sql_t, tmpConn, tmpTrans)
                        With sCmd
                            .Parameters.Clear()
                            .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = dr1("IDNO")
                            .Parameters.Add("@OCID", SqlDbType.Int).Value = dr1("OCID1")
                            objdt.Load(.ExecuteReader())
                        End With

                        'With sqlAdpRev
                        '    .SelectCommand=New SqlCommand(sql, tmpConn, tmpTrans)
                        '    .SelectCommand.Parameters.Clear()
                        '    .SelectCommand.Parameters.Add("@IDNO", SqlDbType.VarChar).Value=dr1("IDNO")
                        '    .SelectCommand.Parameters.Add("@OCID", SqlDbType.Int).Value=dr1("OCID1")
                        '    .Fill(objDSRev, "Data")
                        'End With

                        If objdt.Rows.Count > 0 Then
                            For Each dr_Data As DataRow In objdt.Rows
                                'Select Case Convert.ToString(dr_Data("TRNDMode")) '判斷如果學員是從三合一帶近來的則把職訓卷記錄檔,學習卷記錄檔,推介記錄檔的 transToTIMS 改成'N'
                                '    Case "1"
                                '        sql="UPDATE Adp_TRNData SET TransToTIMS='N' WHERE TICKET_NO='" & Convert.ToString(dr_Data("TICKET_NO")) & "'"
                                '        DbAccess.ExecuteNonQuery(sql, tmpTrans)
                                '    Case "2"
                                '        sql="UPDATE Adp_DGTRNData SET TransToTIMS='N' WHERE TICKET_NO='" & Convert.ToString(dr_Data("TICKET_NO")) & "'"
                                '        DbAccess.ExecuteNonQuery(sql, tmpTrans)
                                '    Case "3"
                                '        sql="UPDATE Adp_GOVTRNData SET TransToTIMS='N' WHERE TICKET_NO='" & Convert.ToString(dr_Data("TICKET_NO")) & "'"
                                '        DbAccess.ExecuteNonQuery(sql, tmpTrans)
                                'End Select
                                '還原刪除 Type Result
                                DEL_STUDENTERTYPE(dr_Data("SETID"), dr_Data("OCID1"), sm, tmpTrans)
                                DEL_STUDSELRESULT(dr_Data("SETID"), dr_Data("OCID1"), sm, tmpTrans)
                            Next
                        End If
                        If Convert.ToString(dr1("SETID")) <> "" Then
                            DEL_STUDSELRESULT(dr1("SETID"), dr1("OCID1"), sm, tmpTrans)
                        End If
                    End If

                    Dim sql_u As String = ""
                    sql_u &= " UPDATE STUD_ENTERTYPE2" & vbCrLf
                    sql_u &= " SET signUpStatus=0" & vbCrLf
                    sql_u &= " ,signUpMemo=NULL" & vbCrLf '備註(失敗原因)
                    sql_u &= " ,isEmailFail=NULL" & vbCrLf
                    sql_u &= " ,modifydate=GETDATE()" & vbCrLf
                    sql_u &= " ,modifyacct=@modifyacct" & vbCrLf
                    sql_u &= " WHERE ESERNUM IN (" & vbCrLf
                    sql_u &= " 	SELECT st.ESERNUM" & vbCrLf
                    sql_u &= " 	FROM stud_entertemp2 sp" & vbCrLf
                    sql_u &= " 	JOIN STUD_ENTERTYPE2 st ON sp.esetid=st.esetid" & vbCrLf
                    sql_u &= " 	WHERE sp.IDNO=@IDNO" & vbCrLf
                    sql_u &= " 	AND st.OCID1=@OCID1" & vbCrLf
                    sql_u &= "  AND st.eSerNum=@eSerNum" & vbCrLf
                    sql_u &= " )" & vbCrLf
                    Dim uCmd As New SqlCommand(sql_u, tmpConn, tmpTrans)
                    With uCmd
                        .Parameters.Clear()
                        .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .Parameters.Add("IDNO", SqlDbType.VarChar).Value = UCase(dr1("IDNO"))
                        .Parameters.Add("OCID1", SqlDbType.VarChar).Value = dr1("OCID1")
                        .Parameters.Add("eSerNum", SqlDbType.Int).Value = e.CommandArgument
                        .ExecuteNonQuery()
                    End With
                    DbAccess.CommitTrans(tmpTrans)

                    'flagtmpTransCommit=False '未完成 Commit 'tmpTrans.Commit() 'flagtmpTransCommit=True '已經完成 Commit
                    revFlag = True
                    sMemo = String.Concat("&動作=還原", "&IDNO=", tmpIDNO) '寫入Log查詢 (Auth_Accountlog)
                    Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm修改, TIMS.cst_wmdip1, Me.OCIDValue1.Value, sMemo)

                Catch ex As Exception
                    '交易已開始，但未完成 Commit
                    DbAccess.RollbackTrans(tmpTrans)
                    DbAccess.CloseDbConn(tmpConn)
                    TIMS.LOG.Error(ex.Message, ex)
                    'If flagtmpTransBeginCommit AndAlso flagtmpTransCommit=False Then
                    '    If tmpTrans IsNot Nothing Then tmpTrans.Rollback()
                    'End If
                    strMsgBox = "資料錯誤，無法還原報名狀態!!"
                    Common.MessageBox(Me, strMsgBox) 'Common.MessageBox(Me, ex.ToString)
                    Exit Sub
                End Try
                DbAccess.CloseDbConn(tmpConn)

                If revFlag Then
                    strMsgBox = "報名資料還原成功!!"
                    Common.MessageBox(Me, strMsgBox)
                End If
                '查詢
                'Call Button1_Click(Button1, e)
                Call Search1()
        End Select
    End Sub

    'PageControler1
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Select Case SsignUpStatus.SelectedValue
                    Case "2", "3" '2:審核成功 '3:審核失敗'Case "0","1" '不區分 '尚未審核
                        e.Item.Cells(cst_報名審核).Text = "報名審核"
                End Select
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                    e.Item.Cells(cst_報名審核).Text = "報名審核"
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim star1 As TextBox=e.Item.FindControl("star1")
                Dim star2 As Label = e.Item.FindControl("star2")
                Dim star3 As Label = e.Item.FindControl("star3")
                Dim star4 As Label = e.Item.FindControl("star4")
                Dim hstar3 As HtmlInputHidden = e.Item.FindControl("hstar3") 'hstar3:重複參訓
                Dim stud1 As TextBox = e.Item.FindControl("stud1") '序號
                'Dim LabGovCost As Label=e.Item.FindControl("LabGovCost")
                Dim LabEnterPath As Label = e.Item.FindControl("LabEnterPath") '報名路徑
                'Dim BudID As DropDownList=e.Item.FindControl("BudID")
                'Dim SupplyID As DropDownList=e.Item.FindControl("SupplyID")
                Dim labSTNAME As Label = e.Item.FindControl("labSTNAME")
                Dim BtnHistory As Button = e.Item.FindControl("BtnHistory") '近2年參訓
                'Dim BudgetID97 As Label = e.Item.FindControl("BudgetID97")
                Dim HidBirthDay As HtmlInputHidden = e.Item.FindControl("HidBirthDay")
                Dim HidSTDate As HtmlInputHidden = e.Item.FindControl("HidSTDate")
                Dim Hid_eSerNum As HtmlInputHidden = e.Item.FindControl("Hid_eSerNum")
                Dim HidCMASTER1 As HtmlInputHidden = e.Item.FindControl("HidCMASTER1") '認定為公司負責人
                Dim HidCMASTER1NT As HtmlInputHidden = e.Item.FindControl("HidCMASTER1NT") '已切結
                Dim Hid_IJC As HtmlInputHidden = e.Item.FindControl("Hid_IJC")
                '該民眾為臺東地區民眾，若符合「勞動部因應重大災害職業訓練協助計畫」受災者，請其提供證明文件，得免試入訓。
                '<input id="Hid_MSG1A" type="hidden" runat="server" />
                Dim Hid_MSG1NAME As HtmlInputHidden = e.Item.FindControl("Hid_MSG1NAME")
                Dim Hid_MSGTYPEN As HtmlInputHidden = e.Item.FindControl("Hid_MSGTYPEN")
                Dim Hid_MSGADIDN As HtmlInputHidden = e.Item.FindControl("Hid_MSGADIDN")
                Dim labDiffYears As Label = e.Item.FindControl("labDiffYears")
                labDiffYears.Text = ""
                If Not Convert.ToString(drv("YEARS")).Equals(sm.UserInfo.Years) Then
                    labDiffYears.Text &= String.Concat("(", drv("YEARS"), ")")
                End If
                If Not Convert.ToString(drv("TPLANID")).Equals(sm.UserInfo.TPlanID) Then
                    labDiffYears.Text &= String.Concat("P:", drv("TPLANID"))
                End If

                Dim iTYPEN As Integer = 0 'iTYPEN:返回颱風類型。
                Dim iADID As Integer = 0 'iADID:返回 重大災害受災地區範圍 序號。
                Dim sENTERDATE As String = $"{drv("ENTERDATE")}"
                Dim sZIPCODE As String = $"{drv("ZIPCODE")}" 'Convert.ToString(drv("ZIPCODE")) '通訊郵遞區號
                Dim flagMSG1 As Boolean = False '無資訊
                If sZIPCODE <> "" AndAlso Not flagMSG1 Then
                    flagMSG1 = TIMS.CHK_ZIP2MSG(Me, sZIPCODE, sENTERDATE, objconn, iTYPEN)
                    If flagMSG1 Then
                        Hid_MSG1NAME.Value = TIMS.HtmlDecode1(Convert.ToString(drv("NAME"))) '有(資訊)值塞入姓名
                        Hid_MSGTYPEN.Value = iTYPEN
                    End If
                End If
                '查無(資訊)資料 再判斷1次
                sZIPCODE = Convert.ToString(drv("ZIPCODE2")) '戶籍地址-郵遞區號  int
                If sZIPCODE <> "" AndAlso Not flagMSG1 Then
                    flagMSG1 = TIMS.CHK_ZIP2MSG(Me, sZIPCODE, sENTERDATE, objconn, iTYPEN)
                    If flagMSG1 Then
                        Hid_MSG1NAME.Value = TIMS.HtmlDecode1(Convert.ToString(drv("NAME"))) '有(資訊)值塞入姓名
                        Hid_MSGTYPEN.Value = iTYPEN
                    End If
                End If
                'https://jira.turbotech.com.tw/browse/TIMSC-150
                sZIPCODE = Convert.ToString(drv("ZIPCODE")) '通訊郵遞區號
                If sZIPCODE <> "" AndAlso Not flagMSG1 Then
                    flagMSG1 = TIMS.CHK_DIS2MSG(Me, sZIPCODE, sENTERDATE, objconn, iADID)
                    If flagMSG1 Then
                        Hid_MSG1NAME.Value = TIMS.HtmlDecode1(Convert.ToString(drv("NAME"))) '有(資訊)值塞入姓名
                        Hid_MSGADIDN.Value = iADID
                    End If
                End If
                sZIPCODE = Convert.ToString(drv("ZIPCODE2")) '戶籍地址-郵遞區號
                If sZIPCODE <> "" AndAlso Not flagMSG1 Then
                    flagMSG1 = TIMS.CHK_DIS2MSG(Me, sZIPCODE, sENTERDATE, objconn, iADID)
                    If flagMSG1 Then
                        Hid_MSG1NAME.Value = TIMS.HtmlDecode1(Convert.ToString(drv("NAME"))) '有(資訊)值塞入姓名
                        Hid_MSGADIDN.Value = iADID
                    End If
                End If
                Dim HidIdentityID As HtmlInputHidden = e.Item.FindControl("HidIdentityID")
                HidIdentityID.Value = Convert.ToString(drv("IdentityID"))
                If Convert.ToString(drv("SRSOLDIERS")) = "Y" Then
                    '12:屆退官兵(須單位將級以上長官薦送函)
                    Const cst_id12 As String = "12"
                    If HidIdentityID.Value.IndexOf(cst_id12) = -1 Then
                        If HidIdentityID.Value <> "" Then HidIdentityID.Value &= ","
                        HidIdentityID.Value &= cst_id12
                    End If
                End If
                LabEnterPath.Text = Convert.ToString(drv("EnterPath_N")) '報名路徑
                Hid_eSerNum.Value = Convert.ToString(drv("eSerNum"))
                HidCMASTER1.Value = Convert.ToString(drv("CMASTER1")) '認定為公司負責人
                HidCMASTER1NT.Value = Convert.ToString(drv("CMASTER1NT")) '已切結
                HidBirthDay.Value = Convert.ToString(drv("BirthDay"))
                HidSTDate.Value = Convert.ToString(drv("STDate1"))

                '預算別
                'BudID.Attributes("onchange")="return Change('" & BudID.ClientID & "','" & SupplyID.ClientID & "');"
                e.Item.Cells(cst_保險證號).ToolTip &= String.Concat("主要身分別: ", drv("MIdentityName"))
                'If Me.DistValue.Value <> "001" Then BudID.Items.RemoveAt(0) '排除北區
                '檢查此身分證號碼是否有重複報名- --Start
                'Dim sql, sqlGC As String
                'Dim dr, drGC As DataRow
                '檢查此身分證號碼是否有重複報名-- --End
                stud1.Text = TIMS.Get_DGSeqNo(sender, e)

                Dim vTitle As String = String.Concat(" 報名序號：", drv("eSerNum"), vbCrLf, " 班級代號：", drv("OCID1"))
                TIMS.Tooltip(stud1, vTitle)
                If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '您的報名序號SignNo
                    stud1.Text = Convert.ToString(drv("SignNo"))
                    stud1.ForeColor = Color.Blue
                    stud1.Enabled = True
                    stud1.ReadOnly = True
                End If

                '將身分證號碼存入陣列-start
                IDNOArray.Add(TIMS.ChangeIDNO(drv("IDNO").ToString))
                Session("IDNOArray") = IDNOArray
                '將身分證號碼存入陣列-- --end

                ''政府已補助經費   Start
                'Dim sqlGC As String=""
                'sqlGC="SELECT dbo.fn_GET_GOVCOST('" & UCase(drv("IDNO")) & "','" & drv("STDate1").ToString & "') GovCost "
                'Call TIMS.OpenDbConn(objconn)
                'LabGovCost.Text=DbAccess.ExecuteScalar(sqlGC, objconn)
                'If LabGovCost.Text="" Then LabGovCost.Text="0"
                ''LabGovCost.Text=Val(drv("GovCost"))
                'If LabGovCost.Text.Trim <> "" Then
                '    'LabGovCost.Text=drv("GovCost").ToString
                '    If CInt(LabGovCost.Text) > TIMS.Get_3Y_SupplyMoney(Me) Then
                '        LabGovCost.ForeColor=LabGovCost.ForeColor.Red '超過三萬(政府已補助經費)的提示，將字變為紅色的
                '        LabGovCost.Font.Bold=True
                '    End If
                '    LabGovCost.ToolTip="此班開訓日期前的補助經費"
                'End If
                ''政府已補助經費   End

                'Dim ParentRID As String
                'Dim ParentOrgName As String
                'If Split(drv("Relship"), "/").Length >= 3 Then
                '    ParentRID=Split(drv("Relship"), "/")(Split(drv("Relship"), "/").Length - 3)
                '    If ParentRID.Length=1 Then
                '        e.Item.Cells(cst_報名機構).Text=drv("OrgName")
                '    Else
                '        ParentOrgName=DbAccess.ExecuteScalar("SELECT OrgName FROM view_RIDName WHERE RID='" & ParentRID & "'", objconn)
                '        e.Item.Cells(cst_報名機構).Text=ParentOrgName & "-" & drv("OrgName")
                '    End If
                'End If

                e.Item.Cells(cst_報名機構).Text = Convert.ToString(drv("OrgName"))
                If CInt(drv("OrgLevel")) >= 2 Then e.Item.Cells(cst_報名機構).Text = Convert.ToString(drv("OrgName2"))

                e.Item.Cells(cst_報名班級).Text = drv("ClassCName1").ToString
                If IsNumeric(drv("CyclType1")) AndAlso Int(drv("CyclType1")) <> 0 Then
                    e.Item.Cells(cst_報名班級).Text &= String.Concat("第", drv("CyclType1"), "期")
                End If

                If drv("LevelName").ToString <> "" AndAlso Int(drv("LevelName")) <> 0 Then
                    e.Item.Cells(cst_報名班級).Text &= String.Concat("(第", drv("LevelName"), "階段)")
                End If

                If DataGrid1.Columns(cst_報名路徑).Visible Then
                    e.Item.Cells(cst_報名班級).Attributes("onmouseover") = "document.getElementById('ClassData" & e.Item.ItemIndex & "').style.display='inline';"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseover") += "if(document.getElementById('" & BudID.ClientID & "'))"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseover") += "document.getElementById('" & BudID.ClientID & "').style.display='none';"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseover") += "if(document.getElementById('" & SupplyID.ClientID & "'))"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseover") += "document.getElementById('" & SupplyID.ClientID & "').style.display='none';"
                    e.Item.Cells(cst_報名班級).Attributes("onmouseout") = "document.getElementById('ClassData" & e.Item.ItemIndex & "').style.display='none';"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseout") += "if(document.getElementById('" & BudID.ClientID & "'))"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseout") += "document.getElementById('" & BudID.ClientID & "').style.display='inline';"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseout") += "if(document.getElementById('" & SupplyID.ClientID & "'))"
                    'e.Item.Cells(cst_報名班級).Attributes("onmouseout") += "document.getElementById('" & SupplyID.ClientID & "').style.display='inline';"
                Else
                    e.Item.Cells(cst_報名班級).Attributes("onmouseover") = "document.getElementById('ClassData" & e.Item.ItemIndex & "').style.display='inline';"
                    e.Item.Cells(cst_報名班級).Attributes("onmouseout") = "document.getElementById('ClassData" & e.Item.ItemIndex & "').style.display='none';"
                End If

                '20121129 個資顯示
                'If Not flgCIShow Then e.Item.Cells(cst_身分證號碼).Text=TIMS.strMask(Convert.ToString(drv("IDNO")), 1)

                '模楜顯示，且身分證號為空白
                'e.Item.Cells(cst_身分證號碼).Text=Convert.ToString(drv("IDNO"))
                'If rblWorkMode.SelectedValue="1" AndAlso IDNO.Text="" Then e.Item.Cells(cst_身分證號碼).Text=TIMS.strMask(Convert.ToString(drv("IDNO")), 1) '使用 '個資法遮罩
                e.Item.Cells(cst_身分證號碼).Text = TIMS.strMask(Convert.ToString(drv("IDNO")), 1) '使用 '個資法遮罩
                'Cells(cst_姓名)
                labSTNAME.Text = TIMS.HtmlDecode1(Convert.ToString(drv("NAME")))
                '(「職場續航」之課程勾稽投保年資)
                star4.Visible = False
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso ViewState("WYROLE") = "Y" Then
                    '(「職場續航」之課程勾稽投保年資)
                    star4.Visible = (Convert.ToString(drv("WYROLE")) <> "")
                    '* 表示該學員符合「職場續航」優先錄訓條件，可將滑鼠移至姓名處查看年資、年齡。
                    '職場續航-(優先錄訓條件1-工作15年以上年滿55歲、2-工作25年以上、3-工作10年以上年滿60歲、4-年滿65歲),,'4-強制退休年齡前2年內之63-64歲者
                    Dim titWYROLE As String = If(star4.Visible, String.Concat("(", drv("WYROLE"), ")"), "")
                    Dim titname As String = String.Concat("年資", drv("ITRMY"), "年、", drv("AGE"), "歲", titWYROLE) '年資年、歲
                    If titname = "年資年、歲" Then titname = String.Concat("無勾稽資料,", drv("STDAGE"), "歲")
                    TIMS.Tooltip(labSTNAME, titname, True)

                ElseIf TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 AndAlso Convert.ToString(drv("ExamDate")) <> "" Then
                    Dim vstitle As String = String.Concat("甄試日期", drv("ExamDate")) '"甄試日期未定義"
                    'If Convert.ToString(drv("ExamDate")) <> "" Then vstitle=String.Concat("甄試日期為", drv("ExamDate"))
                    TIMS.Tooltip(labSTNAME, vstitle, True)

                    '表格繪製-start 'DISPLAY: none; '非本班同一天甄試日期 ExamDate 'If Convert.ToString(drv("ExamDate")) <> "" Then 'End If
                    Dim pms_it2 As New Hashtable From {{"ExamDate", drv("ExamDate")}, {"IDNO", drv("IDNO")}, {"OCID", drv("OCID1")}}
                    Dim sSql_it2 As String = ""
                    sSql_it2 &= " SELECT p.OCID,p.DISTNAME2,p.PLANNAME,p.ORGNAME,dbo.FN_GET_CLASSCNAME(p.CLASSCNAME,p.CYCLTYPE) CLASSCNAME2,p.ExamDate" & vbCrLf
                    sSql_it2 &= " FROM dbo.VIEW_ENTERCHANNEL5 p WHERE CONVERT(DATE,p.ExamDate)=CONVERT(DATE,@ExamDate) AND p.IDNO=@IDNO AND p.OCID!=@OCID"
                    Dim dt2 As DataTable = DbAccess.GetDataTable(sSql_it2, objconn, pms_it2)

                    If dt2.Rows.Count > 0 Then
                        'labSTNAME.Text=String.Concat("<FONT color='blue'>", drv("NAME"), "<FONT>")
                        labSTNAME.ForeColor = System.Drawing.Color.Blue
                        With labSTNAME
                            If DataGrid1.Columns(cst_報名路徑).Visible Then
                                .Attributes("onmouseover") = "document.getElementById('ExamData" & e.Item.ItemIndex & "').style.display='inline';"
                                .Attributes("onmouseout") = "document.getElementById('ExamData" & e.Item.ItemIndex & "').style.display='none';"
                            Else
                                .Attributes("onmouseover") = "document.getElementById('ExamData" & e.Item.ItemIndex & "').style.display='inline';"
                                .Attributes("onmouseout") = "document.getElementById('ExamData" & e.Item.ItemIndex & "').style.display='none';"
                            End If
                        End With
                        'Cells(cst_身分證號碼) ""
                        Dim s_ExamData_Double As String = Get_ExamData_Double(e, dt2)
                        '個資法遮罩 '模楜顯示，且身分證號為空白
                        e.Item.Cells(cst_身分證號碼).Text = String.Concat(TIMS.strMask(Convert.ToString(drv("IDNO")), 1), s_ExamData_Double)
                        '表格繪製-end
                    End If

                End If

                '(內網報名資訊查詢)
                Dim RedFlag As Boolean = Chk_StudEnterType(Convert.ToString(drv("IDNO")), Convert.ToString(drv("OCID1")))
                e.Item.Cells(cst_報名日期).Text = "<Table class='font' bgcolor='#FFFFE6' width='300' id='ClassData" & e.Item.ItemIndex & "' style='DISPLAY: none; POSITION: absolute;BORDER-COLLAPSE: collapse' border=1>"
                e.Item.Cells(cst_報名日期).Text &= "<TR>"
                e.Item.Cells(cst_報名日期).Text &= "<TD>"
                If RedFlag = True Then e.Item.Cells(cst_報名日期).Text &= "<font color='red'>"
                e.Item.Cells(cst_報名日期).Text &= "第一志願:" & drv("ClassCName1").ToString
                If IsNumeric(drv("CyclType1")) AndAlso Int(drv("CyclType1")) <> 0 Then
                    e.Item.Cells(cst_報名日期).Text &= "第" & Int(drv("CyclType1")) & "期"
                End If
                If drv("LevelName").ToString <> "" AndAlso Int(drv("LevelName")) <> 0 Then
                    e.Item.Cells(cst_報名日期).Text &= "第" & Int(drv("LevelName")) & "階段"
                End If
                If RedFlag = True Then e.Item.Cells(cst_報名日期).Text &= "(內網已有紀錄)</font>"
                e.Item.Cells(cst_報名日期).Text &= "</TD>"
                e.Item.Cells(cst_報名日期).Text &= "</TR>"
                e.Item.Cells(cst_報名日期).Text &= "</Table>"
                'e.Item.Cells(cst_報名日期).Text += FormatDateTime(drv("RelEnterDate"), 2)
                'e.Item.Cells(cst_報名日期).Text += FormatDateTime(drv("RelEnterDate"), DateFormat.GeneralDate)
                '表格繪製-end

                '西元轉民國年顯示
                Dim relEnterDate As String = FormatDateTime(drv("RelEnterDate"), DateFormat.GeneralDate)
                If flag_ROC Then
                    relEnterDate = TIMS.Cdate17t1(drv("RelEnterDate"))
                End If
                e.Item.Cells(cst_報名日期).Text &= relEnterDate

                Dim signUpStatus1 As HtmlInputRadioButton = e.Item.FindControl("signUpStatus1")
                Dim signUpStatus2 As HtmlInputRadioButton = e.Item.FindControl("signUpStatus2")
                Dim signUpStatus As HtmlInputHidden = e.Item.FindControl("signUpStatus")
                Dim signUpMemo As TextBox = e.Item.FindControl("signUpMemo") '備註(失敗原因)
                Dim btn As LinkButton = e.Item.FindControl("Button2") '檢視
                Dim btn1 As LinkButton = e.Item.FindControl("Button4") '刪除
                Dim Button6 As LinkButton = e.Item.FindControl("Button6") '還原
                'Dim labSignUp As Label=e.Item.FindControl("lab_SignUp")
                'Dim spanSignUp As HtmlControl=e.Item.FindControl("span_SignUp")

                signUpStatus.Value = Convert.ToString(drv("signUpStatus")) '審核狀態
                Dim WorkSuppIdent1 As HtmlInputRadioButton = e.Item.FindControl("WorkSuppIdent1")
                Dim WorkSuppIdent2 As HtmlInputRadioButton = e.Item.FindControl("WorkSuppIdent2")
                If Not Convert.ToString(drv("signUpStatus")) = "0" Then '0:收件完成 非0:已審
                    'cst_是否為在職者補助身分
                    Select Case Convert.ToString(drv("WorkSuppIdent")) '轉為文字
                        Case "Y"
                            e.Item.Cells(cst_是否為在職者補助身分).Text = "是"
                        Case "N"
                            e.Item.Cells(cst_是否為在職者補助身分).Text = "否"
                        Case Else
                            e.Item.Cells(cst_是否為在職者補助身分).Text = "無資料"
                    End Select
                Else
                    Select Case Convert.ToString(drv("WorkSuppIdent"))
                        Case "Y"
                            WorkSuppIdent1.Checked = True
                        Case "N"
                            WorkSuppIdent2.Checked = True
                    End Select
                End If

                '970520 Andy 修正審核狀態 
                'Stud_EnterType2  [signUpStatus] 
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                Hid_IJC.Value = ""
                Select Case drv("signUpStatus")'為數字
                    Case 0 '尚未審核。
                        'Dim flagCanUseChk1 As Boolean=False '是否可執行檢核(true:可 false:不可)
                        'If TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID)=-1 Then flagCanUseChk1=True
                        'If TIMS.Cst_TPlanID_PreUseLimited17e.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagCanUseChk1=True

                        'TIMS 非產投計畫(在職)。
                        'If TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID)=-1 Then

                        '    '(職前課程邏輯)若為下列計畫, 則依4項不予錄訓規定設定邏輯判斷學員是否可參訓:
                        '    ' https://jira.turbotech.com.tw/browse/TIMSC-142
                        '    ' 呼叫 TIMS.Get_ChkIsJobsCounse44() 進行檢查

                        '    Dim vERRMSG As String=""
                        '    Dim IDNOt As String=Convert.ToString(drv("IDNO"))
                        '    Dim OCIDVal As String=Convert.ToString(drv("OCID1"))
                        '    If OCIDVal <> "" Then
                        '        Dim htSS As New Hashtable 'htSS Hashtable() '
                        '        htSS.Add("IDNOt", IDNOt)
                        '        htSS.Add("OCIDVal", OCIDVal)
                        '        htSS.Add("SENTERDATE", TIMS.cdate3(drv("SENTERDATE")))
                        '        vERRMSG &= TIMS.Get_ChkIsJobsCounse44(Me, htSS, TIMS.cst_FunID_e網報名審核, objconn)
                        '        If vERRMSG <> "" Then
                        '            'Common.MessageBox(Me, vERRMSG)
                        '            'Exit Sub
                        '            signUpStatus2.Checked=True '失敗
                        '            signUpStatus1.Disabled=True
                        '            signUpStatus2.Disabled=True
                        '            TIMS.Tooltip(signUpStatus1, vERRMSG)
                        '            TIMS.Tooltip(signUpStatus2, vERRMSG)
                        '        End If
                        '    End If

                        'End If


                        '非產投計畫
                        '.e網報名審核的查詢結果列表，若顯示的報名資料的報名班級已過最晚可e網報名審核作業時間
                        '，則報名審核欄位的o成功o失敗的選擇，就變成灰色不可選擇狀態。
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                            If Convert.ToString(drv("LOCK1")) = "Y" Then
                                signUpStatus1.Disabled = True
                                signUpStatus2.Disabled = True
                                vMsg = "已過最晚可e網報名審核作業時間" 'Const cst_msgA1 As String="已過最晚可e網報名審核作業時間"
                                TIMS.Tooltip(signUpStatus1, vMsg, True)
                                TIMS.Tooltip(signUpStatus2, vMsg, True)
                                If TIMS.sUtl_ChkTest() Then
                                    signUpStatus1.Disabled = False
                                    signUpStatus2.Disabled = False
                                End If
                            End If
                        End If
                    Case 1, 3, 4, 5
                        'labSignUp.Text="成功"
                        'spanSignUp.Visible=False
                        'SELECT signUpStatus,COUNT(1) CNT FROM STUD_ENTERTYPE2 WHERE OCID1 =135705  GROUP BY signUpStatus
                        'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                        Dim S_REVIEW As String = "成功"
                        e.Item.Cells(cst_報名審核).Text = S_REVIEW
                        signUpMemo.ReadOnly = True '備註(失敗原因)
                    Case 2 '失敗 'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                            'e.Item.Cells(cst_報名審核).Text="失敗"
                            signUpStatus2.Checked = True '失敗
                            'labSignUp.Text="失敗"
                            'spanSignUp.Visible=False
                            signUpMemo.ReadOnly = True '備註(失敗原因)
                        Else
                            If Not CheckExamDate(Convert.ToString(drv("OCID1")), Convert.ToString(drv("ExamDate"))) Then
                                e.Item.Cells(cst_報名審核).Text = "失敗"
                                'labSignUp.Text="失敗"
                                'spanSignUp.Visible=False
                                signUpMemo.ReadOnly = True '備註(失敗原因)
                            Else
                                signUpStatus2.Checked = True '失敗
                            End If
                        End If
                End Select

                '資訊可能異常判斷
                If Convert.ToString(drv("signUpStatus")) <> "0" AndAlso IsDBNull(drv("BudID")) AndAlso
                    IsDBNull(drv("SupplyID")) AndAlso IsDBNull(drv("ExamNo")) AndAlso
                    UCase(Convert.ToString(drv("MODIFYACCT"))) = UCase(Convert.ToString(drv("IDNO"))) Then
                    e.Item.Cells(cst_報名審核).Text = "<font color=red>成功</font>"
                    TIMS.Tooltip(e.Item.Cells(cst_報名審核), "(資訊可能異常)請確認該資料，執行還原再次審核!!")
                End If

                signUpStatus1.Attributes("onclick") = "ChangeStatus(1,'" & signUpStatus.ClientID & "');"
                signUpStatus2.Attributes("onclick") = "ChangeStatus(2,'" & signUpStatus.ClientID & "');"

                '---預算及補助比例start---
                'If Convert.ToString(drv("ActNo"))="無資料" Then  '判斷如果產投的投保單位保險證號沒有資料就判斷職前的投保單位保險證號有無資料
                '    If Convert.ToString(drv("ActNo2")) <> "" Then
                '        ActNo=Convert.ToString(drv("ActNo2"))
                '    End If
                'Else
                '    ActNo=Convert.ToString(drv("ActNo"))
                'End If

                '2011/4/15日'充電起飛計畫公告為 4/15日後才可使用協助基金 
                'Const Cst_20110415 As String = "2011/04/15"
                'BudgetID97.Text = "否" '是否是協助基金
                'Dim vsFOfficeYM1 As String = ""
                'Dim vsSTDate1 As String = ""
                'If Convert.ToString(drv("FOfficeYM1")) <> "" Then vsFOfficeYM1 = CDate(drv("FOfficeYM1")).ToString("yyyy/MM/dd")
                'If Convert.ToString(drv("STDate1")) <> "" Then vsSTDate1 = CDate(drv("STDate1")).ToString("yyyy/MM/dd")
                'Dim isEcfa_Flag1 As Boolean = False '如果是 Ecfa 就不再顯示原資料
                'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '    If DateDiff(DateInterval.Day, CDate(Cst_20110415), CDate(vsSTDate1)) >= 0 Then
                '        If TIMS.CheckIsECFA(Me, drv("ActNo"), "", vsSTDate1, objconn) = True Then  '2011/05/20 新增ECFA判斷
                '            'BudID.SelectedValue="97"    '協助 'SupplyID.SelectedValue="2"  '2.補助100%
                '            BudgetID97.Text = "是" '是否是協助基金
                '            isEcfa_Flag1 = True
                '        End If
                '    End If
                'Else
                '    If TIMS.CheckIsECFA(Me, drv("ActNo"), vsFOfficeYM1, "", objconn) = True Then  '2011/05/20 新增ECFA判斷
                '        'BudID.SelectedValue="97"    '協助 'SupplyID.SelectedValue="2"  '2.補助100%
                '        BudgetID97.Text = "是" '是否是協助基金
                '        isEcfa_Flag1 = True
                '    End If
                'End If

                '不是 Ecfa 就顯示原資料
                'If Not isEcfa_Flag1 Then
                '    If Not IsDBNull(drv("BudID")) Then Common.SetListItem(BudID, drv("BudID").ToString)
                '    If Not IsDBNull(drv("SupplyID")) Then Common.SetListItem(SupplyID, drv("SupplyID").ToString)
                'End If
                'If BudID.SelectedValue="97" Then BudgetID97.Text="是" '是否是協助基金
                '---預算及補助比例end---

                signUpMemo.Text = drv("signUpMemo").ToString '備註(失敗原因)
                'BudID2.Value=BudID.SelectedValue 'SupplyID2.Value=SupplyID.SelectedValue '('檢視 '刪除 '還原)
                Dim sCmdArg As String = ""
                sCmdArg &= "&eSerNum=" & Convert.ToString(drv("eSerNum"))
                'sCmdArg &= "&BudID=" & BudID.SelectedValue 'sCmdArg &= "&SupplyID=" & SupplyID.SelectedValue
                btn.CommandArgument = sCmdArg '"&eSerNum=" & drv("eSerNum") & "&BudID=" & BudID.SelectedValue & "&SupplyID=" & SupplyID.SelectedValue
                btn1.CommandArgument = Convert.ToString(drv("eSerNum"))
                Button6.CommandArgument = Convert.ToString(drv("eSerNum"))

                'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then '28:產業人才投資計劃
                '    '補助費及同時段重疊報名情形
                '    Dim sCmdArg As String=""
                '    TIMS.SetMyValue(sCmdArg, "IDNO", Convert.ToString(drv("IDNO")))
                '    TIMS.SetMyValue(sCmdArg, "OCID1", Convert.ToString(drv("OCID1")))
                '    TIMS.SetMyValue(sCmdArg, "eSerNum", Convert.ToString(drv("eSerNum")))
                '    TIMS.SetMyValue(sCmdArg, "eSETID", Convert.ToString(drv("eSETID")))
                '    'BtnHistory.CommandName="tims28"
                '    BtnHistory.CommandArgument=sCmdArg
                '    BtnHistory.Text=cst_tims28_BtnHistory '"補助費時段重疊情形" '"補助費及同時段重疊報名情形"
                '    BtnHistory.Attributes("onclick")="window.open('../01/SD_01_004_dbl.aspx?" & sCmdArg & "' ,'history','width=770,height=600,scrollbars=1'); return false;"
                'Else
                '    'CommandName="History"
                '    BtnHistory.CommandArgument=drv("IDNO") '近2年參訓
                '    '改成近二年的學員參訓歷史
                '    Dim TwoYears As Integer=0
                '    TwoYears=CInt(Year(aNow)) - 2 '取得去年的年度(yyyy)
                '    BtnHistory.Attributes("onclick")="window.open('../05/SD_05_010.aspx?SD_01_004_Type=Student&IDNO=" & drv("IDNO") & "&TwoYears=" & TwoYears & "&BtnHistory=" & Button1.ClientID & "' ,'history','width=700,height=500,scrollbars=1'); return false;"
                'End If

                'CommandName="History"
                BtnHistory.CommandArgument = drv("IDNO") '近2年參訓
                '改成近二年的學員參訓歷史 SD_05_010
                Dim iTwoYears As Integer = CInt(Year(aNow)) - 2 '取得去年的年度(yyyy)
                Call GSet_BtnHistory(BtnHistory, drv, iTwoYears)

                'hstar3:重複參訓
                star3.Visible = False
                If TIMS.Chk_StudStatus(drv("IDNO"), drv("STDate1"), drv("FTDate"), drv("OCID1").ToString, objconn) Then star3.Visible = True
                If star3.Visible Then hstar3.Value = "1"
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then '28:產業人才投資計劃
                    Select Case Convert.ToString(drv("signUpStatus"))
                        Case "0"
                            '2.人數限制於該班訓練人數加10位，報名序號在該班訓練人數加10名以內者，受此限制;
                            '第11名後「報名審核」欄位不受此限，針對報名民眾於不同階段發生參訓時段重疊(報名)情形
                            '，設立無法儲存階段及彈跳提醒視窗，彈跳視窗文字內容:「學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！」
                            '，比對無法進行報名審核之邏輯，請參考-參訓期間重疊比對邏輯.docx與圖3.jpg。
                            If Convert.ToString(drv("signUpStatus")) = "0" AndAlso Convert.ToString(drv("DDCHKVIEW")) = "Y" Then
                                Dim eSerNumX1 As String = String.Concat("'", drv("eSerNum"), "'")
                                Dim strDDCHKVIEW As String = Convert.ToString(Session(cst_ssDDCHKVIEW))
                                If Session(cst_ssDDCHKVIEW) IsNot Nothing Then strDDCHKVIEW = Convert.ToString(Session(cst_ssDDCHKVIEW))
                                If strDDCHKVIEW <> "" Then
                                    If strDDCHKVIEW.IndexOf(eSerNumX1) = -1 Then
                                        signUpStatus1.Disabled = True
                                        signUpStatus2.Disabled = True
                                        vMsg = "不可略過，請點檢視"
                                        TIMS.Tooltip(signUpStatus1, vMsg, True)
                                        TIMS.Tooltip(signUpStatus2, vMsg, True)
                                    End If
                                Else
                                    signUpStatus1.Disabled = True
                                    signUpStatus2.Disabled = True
                                    vMsg = "不可略過，請點檢視"
                                    TIMS.Tooltip(signUpStatus1, vMsg, True)
                                    TIMS.Tooltip(signUpStatus2, vMsg, True)
                                End If
                            End If
                        Case "2"
                            '7.「報名審核」欄位，經點選審核失敗後，即反灰無法再做修正，
                            '如欲更改「報名審核」欄位內容，須經點選「檢視」後，開放可供點選。
                            If Convert.ToString(drv("signUpStatus")) = "2" Then
                                signUpStatus1.Disabled = True '即反灰無法再做修正 
                                signUpStatus2.Disabled = True '即反灰無法再做修正 
                                vMsg = "審核失敗"
                                TIMS.Tooltip(signUpStatus1, vMsg, True)
                                TIMS.Tooltip(signUpStatus2, vMsg, True)

                                Dim eSerNumX1 As String = String.Concat("'", drv("eSerNum"), "'")
                                Dim strDDCHKVIEW As String = Convert.ToString(Session(cst_ssDDCHKVIEW))
                                If Session(cst_ssDDCHKVIEW) IsNot Nothing Then strDDCHKVIEW = Convert.ToString(Session(cst_ssDDCHKVIEW))
                                If strDDCHKVIEW <> "" Then
                                    If strDDCHKVIEW.IndexOf(eSerNumX1) > -1 Then
                                        signUpStatus1.Disabled = False '須經點選「檢視」後，開放可供點選。
                                        signUpStatus2.Disabled = False '須經點選「檢視」後，開放可供點選。
                                    End If
                                End If
                            End If
                    End Select

                    '20090123 andy  edit 產投、在職 2009年 身分別為「非自願失業者」時
                    '1.預算來源設定為 02:就安基金 ； 2.補助比例為100%
                    '20090224 penny 身分別為[非自願失業者]時 預算來源設定改為 03:就保
                    '---  start
                    'If CInt(Me.sm.UserInfo.Years) > 2008 Then
                    '    For i As Integer=0 To Split(Convert.ToString(drv("IdentityID")), ",").Length - 1
                    '        If Split(drv("IdentityID").ToString, ",")(i)="02" Then
                    '            BudID.ClearSelection()
                    '            Common.SetListItem(BudID, "02")
                    '            Common.SetListItem(BudID, "03")
                    '            Common.SetListItem(SupplyID, "2")
                    '        End If
                    '    Next
                    '    If SupplyID.SelectedValue Is Nothing Then Common.SetListItem(SupplyID, drv("SupplyID").ToString)
                    'Else
                    '    Common.SetListItem(SupplyID, drv("SupplyID").ToString)
                    'End If

                    '--- end
                    If Convert.ToString(drv("signUpStatus")) = "0" Then
                        '2010年 取消「非自願離職者」 by AMU 20100414
                        'If Convert.ToString(drv("IdentityID"))="02" Then
                        '    Common.SetListItem(BudID, "03")
                        '    Common.SetListItem(SupplyID, "2")
                        'End If
                        If Convert.ToString(drv("IsStdBlack")) = "Y" Then '判斷是否為黑名單遭處分學員
                            signUpStatus1.Disabled = True
                            signUpStatus2.Checked = True
                            signUpStatus2.Disabled = True
                            signUpStatus.Value = 1
                            'SupplyID.SelectedValue="9" '不補助
                            'SupplyID.Enabled=False
                            'BudID.SelectedValue="99"
                            'BudID.Enabled=False
                            signUpMemo.Text = "此學員己遭處分,系統帶審核失敗" '備註(失敗原因)
                            signUpMemo.ReadOnly = True '備註(失敗原因)
                            hidBlackMsg.Value &= "序號" + stud1.Text + "." + drv("IDNO") + " " + TIMS.HtmlDecode1(drv("Name")) + "已受處分" & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                        End If
                    End If
                    'If drv("BIEPTBL") > 0 Then star1.Visible=True Else star1.Visible=False
                    'star2.Visible=False
                    If TIMS.Check_Sub_SubSidyApply(Convert.ToString(drv("IDNO")), objconn) Then star2.Visible = True Else star2.Visible = False
                    'If drv("SubsidyCost") > 0 Then star2.Visible=True Else star2.Visible=False

                    If sm.UserInfo.LID < 2 Then
                        btn1.Visible = True
                        btn1.Attributes("onclick") = "return confirm('這樣會刪除此學員的報名資料,\n確定要繼續刪除?');"
                    Else
                        btn1.Visible = False
                    End If
                    If drv("signUpStatus").ToString <> "0" Then Button6.Enabled = True Else Button6.Enabled = False
                    'Button6.Attributes("onclick")="return confirm('這樣會還原此學員的報名資料,\n確定要繼續還原?');"
                    If Button6.Enabled = True Then Button6.Attributes("onclick") = "return confirm('這樣會還原此學員的報名資料,\n確定要繼續還原?');"
                Else
                    'star1.Visible=False
                    star2.Visible = False
                    btn1.Visible = False
                    Button6.Visible = False
                    If Convert.ToString(drv("IsStdBlack")) = "Y" Then hidBlackMsg.Value += "序號" + stud1.Text + "." + drv("IDNO") + " " + TIMS.HtmlDecode1(drv("Name")) + "已受處分" & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                    If flgROLEIDx0xLIDx0 Then '判斷登入者的權限。
                        If drv("signUpStatus").ToString <> "0" Then Button6.Visible = True Else Button6.Visible = False
                        If drv("signUpStatus").ToString <> "0" Then Button6.Enabled = True Else Button6.Enabled = False
                        'Button6.Attributes("onclick")="return confirm('這樣會還原此學員的報名資料,\n確定要繼續還原?');"
                        If Button6.Enabled = True Then Button6.Attributes("onclick") = "return confirm('這樣會還原此學員的報名資料,\n確定要繼續還原?');"
                    End If

                    If TIMS.sUtl_ChkTest() Then
                        If drv("signUpStatus").ToString <> "0" Then Button6.Visible = True Else Button6.Visible = False
                        If drv("signUpStatus").ToString <> "0" Then Button6.Enabled = True Else Button6.Enabled = False
                        'Button6.Attributes("onclick")="return confirm('這樣會還原此學員的報名資料,\n確定要繼續還原?');"
                        If Button6.Enabled = True Then Button6.Attributes("onclick") = "return confirm('這樣會還原此學員的報名資料,\n確定要繼續還原?');"
                    End If

                    If Convert.ToString(drv("signUpStatus")) = "0" Then
                        '2010年 取消「非自願離職者」 by AMU 20100414
                        'If Convert.ToString(drv("IdentityID"))="02" Then
                        '    Common.SetListItem(BudID, "03")
                        '    Common.SetListItem(SupplyID, "2")
                        'End If
                        If Convert.ToString(drv("IsStdBlack")) = "Y" Then '判斷是否為黑名單遭處分學員
                            signUpStatus1.Disabled = True
                            signUpStatus2.Checked = True
                            signUpStatus2.Disabled = True
                            signUpStatus.Value = 1
                            'SupplyID.SelectedValue="9" '不補助
                            'SupplyID.Enabled=False
                            'BudID.SelectedValue="99"
                            'BudID.Enabled=False
                            signUpMemo.Text = "此學員己遭處分,系統帶審核失敗" '備註(失敗原因)
                            signUpMemo.ReadOnly = True '備註(失敗原因)
                            hidBlackMsg.Value += "序號" + stud1.Text + "." + drv("IDNO") + " " + TIMS.HtmlDecode1(drv("Name")) + "已受處分" & vbCrLf '加入單名單暫存(2009/07/28 判斷黑名單)
                        End If
                    End If
                End If

                If Convert.ToString(drv("signUpStatus")) <> "0" Then
                    '已審不可再修改預算，補助
                    'BudID.Enabled=False
                    'SupplyID.Enabled=False
                End If
        End Select
    End Sub

    ''' <summary>組html table</summary>
    ''' <param name="e"></param>
    ''' <param name="dt2"></param>
    ''' <returns></returns>
    Function Get_ExamData_Double(ByRef e As System.Web.UI.WebControls.DataGridItemEventArgs, ByRef dt2 As DataTable) As String
        Dim rst As String = ""
        If TIMS.dtNODATA(dt2) Then Return rst 'If dt2.Rows.Count=0 Then Return rst
        rst = "<Table class='font' bgcolor='#FFFFE6' width='300' id='ExamData" & e.Item.ItemIndex & "' style='DISPLAY: none; POSITION: absolute;BORDER-COLLAPSE: collapse' border=1>"
        rst &= "<TR><TD>"
        'SELECT p.OCID,p.DISTNAME2,p.PLANNAME,p.ORGNAME,dbo.FN_GET_CLASSCNAME(p.CLASSCNAME,p.CYCLTYPE) CLASSCNAME2
        Dim xi As Integer = 0  'If xi <> 0 Then rst &= "<BR>" '換行
        For Each dr As DataRow In dt2.Rows
            rst &= String.Concat(If(xi <> 0, "<BR>", ""), dr("DISTNAME2"), ".", dr("PLANNAME"), ".", dr("ORGNAME"), ".", dr("CLASSCNAME2")) ', "&nbsp;第", dr("cycltype"), "期"
            xi += 1
        Next
        rst &= String.Concat("</TD></TR>", "</Table>")
        Return rst
    End Function

    '查詢檢核
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        If start_date.Text <> "" Then
            If flag_ROC Then
                If Not TIMS.IsDate7(start_date.Text) Then Errmsg += "報名日期 起始日期格式有誤" & vbCrLf
                If Errmsg = "" Then start_date.Text = TIMS.Cdate7(start_date.Text)
            Else
                If Not IsDate(start_date.Text) Then Errmsg += "報名日期 起始日期格式有誤" & vbCrLf
                If Errmsg = "" Then start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            End If
        End If
        'Errmsg += "報名日期 起始日期 為必填" & vbCrLf
        If end_date.Text <> "" Then
            If flag_ROC Then
                If Not TIMS.IsDate7(end_date.Text) Then Errmsg += "報名日期 迄止日期格式有誤" & vbCrLf
                If Errmsg = "" Then end_date.Text = TIMS.Cdate7(end_date.Text)
            Else
                If Not IsDate(end_date.Text) Then Errmsg += "報名日期 迄止日期格式有誤" & vbCrLf
                If Errmsg = "" Then end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If
        End If
        'Errmsg += "報名日期 迄止日期 為必填" & vbCrLf

        If Errmsg = "" Then
            If start_date.Text <> "" AndAlso end_date.Text <> "" Then
                If flag_ROC Then
                    If CDate(TIMS.Cdate18(start_date.Text)) > CDate(TIMS.Cdate18(end_date.Text)) Then Errmsg += "【報名日期】的起日不得大於【報名日期】的迄日!!" & vbCrLf
                Else
                    If CDate(start_date.Text) > CDate(end_date.Text) Then Errmsg += "【報名日期】的起日不得大於【報名日期】的迄日!!" & vbCrLf
                End If
            End If
        End If
        'If Convert.ToString(OCID.SelectedValue)="" OrElse Not IsNumeric(OCID.SelectedValue) Then Errmsg += "統計對象 為必選" & vbCrLf
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢鈕 [SQL@Search_SQL]
    Sub Search1()
        'Session(cst_ssDDCHKVIEW)=Nothing '登入後就永不清除SESSION。 
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        divEnterDouble.Visible = False
        divEnterMoney.Visible = False
        dtgAddresses1.Visible = False '匯入名冊產投用(檢視)
        'LinkButton1.Visible=False '測試寄送信件
        'BtnImport28.Enabled=False '(匯入功能)
        'TIMS.Tooltip(BtnImport28, "停用匯入功能。", True)

        Hid_impOCID.Value = "" '要匯入的資訊
        Dim drCC_2854 As DataRow = TIMS.GetOCIDDate2854(OCIDValue1.Value, objconn)
        Dim bfg1_OK As Boolean = (OCIDValue1.Value <> "" AndAlso drCC_2854 IsNot Nothing) '有輸入班級資料-檢核班級是否OK(範圍內)
        If bfg1_OK Then
            Hid_impOCID.Value = OCIDValue1.Value '確認匯入資訊
            TIMS.Tooltip(BtnImport28, "有輸入班級提供匯入功能", True)
        End If

        '(「職場續航」之課程勾稽投保年資)
        gfg_WYROLE = TIMS.CHECK_WYROLE(objconn, OCIDValue1.Value) 'Dim gfg_WYROLE As Boolean=CHECK_WYROLE()
        ViewState("WYROLE") = If(gfg_WYROLE, "Y", "") 'DG使用'(「職場續航」之課程勾稽投保年資)
        LabWYROLE.Visible = gfg_WYROLE
        LabWYROLE2.Visible = gfg_WYROLE

        TPLANID28_TR1.Visible = False '28:產業人才投資計劃
        If bfg1_OK AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case sm.UserInfo.LID
                Case "2" '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                Case Else
                    '暫時權限Table---Start
                    Dim dtArc As DataTable
                    dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
                    '暫時權限Table---End
                    '產業人才投資計劃(匯入e網報名名冊(產業人才投資))
                    'TPLANID28_TR1.Visible=True
                    If TIMS.Check_Auth_RendClass(OCIDValue1.Value, dtArc) Then
                        TPLANID28_TR1.Visible = True '有補登權限
                        BtnImport28.Enabled = True
                        TIMS.Tooltip(BtnImport28, "補登使用(授權檔內有指定之授權帳號,且在補登期間)", True)
                    End If
            End Select
        End If
        '54:充飛-全部的人都可以匯入
        If bfg1_OK AndAlso TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            TPLANID28_TR1.Visible = True '可匯入
            BtnImport28.Enabled = True '可匯入
        End If

        'If flgROLEIDx0xLIDx0 Then LinkButton1.Visible=True '測試寄送信件
        'If sm.UserInfo.RoleID=1 AndAlso sm.UserInfo.LID=1 AndAlso sm.UserInfo.OrgID=214 Then LinkButton1.Visible=True '測試寄送信件
        If Not BtnImport28.Enabled Then
            '未啟用-再檢測測試參數
            If TIMS.Utl_GetConfigSet("SD_01_004_IMPORT") = "Y" Then
                BtnImport28.Enabled = True '(匯入功能)
                TIMS.Tooltip(BtnImport28, "測試機暫時開放 匯入e網報名名冊") ' by AMU 20140623
            End If
        End If

        msg.Text = "查無資料"
        DataGridTable.Visible = False

        Dim drCC As DataRow = Nothing
        If OCIDValue1.Value <> "" Then
            drCC = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
            If drCC Is Nothing Then
                Common.MessageBox(Page, TIMS.cst_NODATAMsg1)
                Exit Sub
            End If
        End If
        hidBlackMsg.Value = "" '清空黑名單暫存記錄(2009/07/28 判斷黑名單)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim parms As New Hashtable()
        Dim sql As String = Search_SQL(parms)
        If sql = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        Dim s_PageSort As String = "RelEnterDate,eSerNum"
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso gfg_WYROLE Then
            s_PageSort = "WYROLE DESC,OCID1,SignNo,RelEnterDate,eSerNum"
        ElseIf TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            s_PageSort = "OCID1,SignNo,RelEnterDate,eSerNum"
        End If

        dt.DefaultView.Sort = s_PageSort
        dt = TIMS.dv2dt(dt.DefaultView)
        CPdt = dt.Copy()
        'If TIMS.Get_SQLRecordCount(sql)=0 Then

        msg.Text = "查無資料"
        DataGridTable.Visible = False
        Button3.Enabled = False
        If dt.Rows.Count = 0 Then Button3.Attributes.Remove("onclick")

        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SIGNNO,CLASSCNAME1,CYCLTYPE1,NAME,IDNO,ORGNAME2,RELENTERDATE,SIGNUPSTATUS,SIGNUPMEMO,ENTERPATH_N")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip1, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        ' dt.Rows.Count > 0
        If TIMS.dtHaveDATA(dt) Then
            Button3.Enabled = True
            'Button3.Attributes.Add("onclick", "return confirm('確認是否儲存資料?');")
            msg.Text = ""
            DataGridTable.Visible = True

            'OJT-21071601：自辦、接受企業委託、區域 - e網報名審：隱藏【協助基金】欄位
            '針對在職進修訓練、接受企業委託、區域產業據點計畫三個計畫，於e網報名審核查詢結果清單， 隱藏【協助基金】欄位，
            '因這三個計畫無公務ECFA預算別， 但產投、充飛不變， 仍保留【協助基金】欄位。
            'Dim flag_ojt_21072601 As Boolean = True
            'If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_ojt_21072601 = False
            'If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_ojt_21072601 = False
            'If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_ojt_21072601 = False
            'DataGrid1.Columns(cst_協助基金).Visible = flag_ojt_21072601

            'DataGridTable 欄位控制 '儲存鈕的錯誤訊息設置
            DataGrid1.Columns(cst_是否為在職者補助身分).Visible = False
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '28:產業人才投資計劃
                DataGrid1.Columns(cst_報名路徑).Visible = True
                'DataGrid1.Columns(cst_預算別).Visible=True
                Button3.Attributes("onclick") = "if(confirm('確認是否儲存資料?')){return CheckData(" & cst_失敗原因 - 1 & ");}else{return false;}"
                '2007北區職業訓練中心產業人才投資方案001 供測試
                'BIEPTBL.Visible=True
                LabSubsidyCost.Visible = True
            Else
                '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '    DataGrid1.Columns(cst_報名路徑).Visible=True
                '    'DataGrid1.Columns(cst_預算別).Visible=False
                '    DataGrid1.Columns(cst_是否為在職者補助身分).Visible=True
                '    Button3.Attributes("onclick")="if(confirm('確認是否儲存資料?')){return ErrmsgShow(" & cst_失敗原因 - 1 & ");}else{return false;}"
                '    '2008北區職業訓練中心在職進修訓練001 供測試
                '    'BIEPTBL.Visible=False
                '    LabSubsidyCost.Visible=False
                'End If

                '其他TIMS計畫。
                DataGrid1.Columns(cst_報名路徑).Visible = False
                'DataGrid1.Columns(cst_預算別).Visible=False
                'DataGrid1.Columns(cst_是否為在職者補助身分).Visible=False
                Button3.Attributes("onclick") = "if(confirm('確認是否儲存資料?')){return ErrmsgShow(" & cst_失敗原因 - 3 & ");}else{return false;}"
                '2008北區職業訓練中心在職進修訓練001 供測試
                'BIEPTBL.Visible=False
                LabSubsidyCost.Visible = False
            End If
            DataGrid1.DataKeyField = "eSerNum"
            'PageControler1.SqlPrimaryKeyDataCreate(sql, "eSerNum")
            PageControler1.PrimaryKey = "eSerNum"
            PageControler1.Sort = s_PageSort 'ORDER BY "WYROLE DESC,OCID1,SignNo,RelEnterDate,eSerNum"
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

        '產投報名班級增加判斷
        divEnterDouble.Visible = False
        divEnterMoney.Visible = False
        'If TIMS.sUtl_ChkTest() Then
        '    divEnterDouble.Visible=True
        '    labEnterDouble.Text=cst_tims28_labdouble & "王小明,林小美"
        '    divEnterMoney.Visible=True
        '    labEnterMoney.Text=cst_tims28_labover6w & "王小明,林小美"
        'End If

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If OCIDValue1.Value <> "" Then
                Dim Rst As String = "" '取得學員姓名 Rst="" '取得學員姓名
                If TIMS.Chk_EnterDouble(OCIDValue1.Value, Rst, objconn) Then
                    'Dim sMsg As String=String.Concat("本班次學員", Rst)
                    'sMsg &= cst_tims28_Double '"有參訓時段重疊的報名情形，請與學員確認。"
                    If Rst <> "" Then
                        divEnterDouble.Visible = True
                        labEnterDouble.Text = String.Concat(cst_tims28_labdouble, Rst)
                        'Common.MessageBox(Me, sMsg)
                    End If
                End If

                Rst = "" '取得學員姓名
                Dim fg_test As Boolean = TIMS.CHK_IS_TEST_ENVC() 'fg_test OrElse fg_use1 
                Dim fg_use1 As Boolean = TIMS.CanUse3Y10WCost()
                If (fg_test OrElse fg_use1) Then
                    If TIMS.Chk_EnterMoney(OCIDValue1.Value, Rst, cst_tims28_SubsidyWarningCost9W, objconn) Then
                        'Dim sMsg As String=String.Concat("本班次學員", Rst)
                        'sMsg &= cst_tims28_over6w '"預估目前補助費使用已達6萬元（包含已核撥、參訓中、已報名的課程）請再提醒學員。"
                        If Rst <> "" Then
                            divEnterMoney.Visible = True
                            labEnterMoney.Text = String.Concat(cst_tims28_labover9w, Rst)
                            'Common.MessageBox(Me, sMsg)
                        End If
                    End If
                Else
                    If TIMS.Chk_EnterMoney(OCIDValue1.Value, Rst, cst_tims28_SubsidyWarningCost6W, objconn) Then
                        'Dim sMsg As String=String.Concat("本班次學員", Rst)
                        'sMsg &= cst_tims28_over6w '"預估目前補助費使用已達6萬元（包含已核撥、參訓中、已報名的課程）請再提醒學員。"
                        If Rst <> "" Then
                            divEnterMoney.Visible = True
                            labEnterMoney.Text = String.Concat(cst_tims28_labover6w, Rst)
                            'Common.MessageBox(Me, sMsg)
                        End If
                    End If
                End If

            End If
        End If
    End Sub

    ''' <summary>匯出鈕</summary>
    Sub Export1()
        Dim parms As New Hashtable()
        Dim sql As String = Search_SQL(parms)
        If sql = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt.DefaultView.Sort = "RelEnterDate,eSerNum"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then dt.DefaultView.Sort = "SignNo,RelEnterDate,eSerNum" '產投排序
        dt = TIMS.dv2dt(dt.DefaultView)

        Dim sPattern As String = "班別名稱,期別,課程代碼,報名日期,姓名,電話日,電話夜,行動電話,服務單位,Email,備註"
        Dim sColumn As String = "ClassCName1,CyclType1,OCID1,RelEnterDate2,Name,Phone1,Phone2,CellPhone,Uname,Email,signUpMemo"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投欄位
            sPattern = "序號,班別名稱,期別,課程代碼,報名日期,報名時間,姓名,電話日,電話夜,行動電話,服務單位,Email,備註"
            sColumn = "SignNo,ClassCName1,CyclType1,OCID1,RelEnterDate2,RelEnterDate3,Name,Phone1,Phone2,CellPhone,Uname,Email,signUpMemo"
        End If

        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, sColumn)
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip1, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        Dim sFileName1 As String = String.Concat("student", TIMS.GetDateNo2())
        'Response.Clear()
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'Common.RespWrite(Me, "<tr>")
        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = "<tr>"
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= "<td>" & sPatternA(i) & "</td>" '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        strHTML &= (TIMS.sUtl_AntiXss(ExportStr))

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dt.DefaultView.Table.Rows
            iNum += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                'Select Case CStr(sColumnA(i))
                '    Case "Phone1", "Phone2", "CellPhone"
                '        ExportStr &= "<td class=""noDecFormat"">" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
                '    Case Else
                '        ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
                'End Select

                Select Case sColumnA(i)
                    Case "RelEnterDate2" '日期欄位(西元轉民國年顯示)
                        ExportStr &= "<td>" & TIMS.Cdate17(dr(sColumnA(i))) & "</td>" '& vbTab
                    Case Else
                        ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
                End Select
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= (TIMS.sUtl_AntiXss(ExportStr))
        Next
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable From {
            {"ExpType", TIMS.GetListValue(RBListExpType)},
            {"FileName", sFileName1},
            {"strSTYLE", strSTYLE},
            {"strHTML", strHTML},
            {"ResponseNoEnd", "Y"}
        }
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Response.End()
    End Sub

    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        'Stud_EnterType2  [signUpStatus] 'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim status As String = GET_SsignUpStatus_VAL()
        If cjobValue.Value <> "" Then RstMemo &= String.Concat("&通俗職類代碼=", cjobValue.Value)
        If IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", IDNO.Text)
        If start_date.Text <> "" Then RstMemo &= String.Concat("&報名日期起日=", start_date.Text)
        If end_date.Text <> "" Then RstMemo &= String.Concat("&報名日期迄日=", end_date.Text)
        If status <> "" Then RstMemo &= String.Concat("&signUpStatus=", status)
        Return RstMemo
    End Function

    ''' <summary> 儲存 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '2006/03/ add conn by matt
        aNow = TIMS.GetSysDateNow(objconn)
        'http://163.29.199.211/Check_ws/Check_ws.asmx
        'Dim Chkws1 As New Check_ws.Check_ws
        Dim sEmailSend As String = TIMS.CheckEmailSend(Me, "", "", objconn)
        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim sBlackIDNO As String = TIMS.Get_StdBlackIDNO(Me, iStdBlackType, stdBLACK2TPLANID, objconn) '學員處分

        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12
        '檢核學員重複參訓。
        'http://163.29.199.211/TIMSWS/timsService1.asmx
        'https://wltims.wda.gov.tw/TIMSWS/timsService1.asmx
        Dim timsSer1 As New timsService1.timsService1

        'CHECK
        Dim tmpMaster As String = "" '如為master(公司負責人搜集姓名)'具公司/商業負責人身分
        Dim ERRMSG As String = ""
        ERRMSG = ""
        Dim i As Integer = 0
        'Dim flagYearsOld65 As Boolean=False '判斷是否為 六十五歲以上者資格
        'BUDID 02
        'Const Cst_Msg65 As String="參訓學員為65歲以上者, 其預算別一律運用就安預算!!預算別，(非就安)有誤!"  '產投

        '檢核所有審核成功的學員資訊
        For Each eItem As DataGridItem In DataGrid1.Items
            i += 1
            Dim signUpStatus1 As HtmlInputRadioButton = eItem.FindControl("signUpStatus1") '成功。
            Dim signUpStatus2 As HtmlInputRadioButton = eItem.FindControl("signUpStatus2") '失敗。
            Dim signUpStatus As HtmlInputHidden = eItem.FindControl("signUpStatus") '報名狀態。
            Dim signUpMemo As TextBox = eItem.FindControl("signUpMemo") '備註(失敗原因)
            'Dim BudID As DropDownList=eItem.FindControl("BudID") '預算別
            'Dim SupplyID As DropDownList=eItem.FindControl("SupplyID") '補助比例
            '是否為在職者補助身分。
            Dim WorkSuppIdent1 As HtmlInputRadioButton = eItem.FindControl("WorkSuppIdent1") '是
            Dim WorkSuppIdent2 As HtmlInputRadioButton = eItem.FindControl("WorkSuppIdent2") '否
            Dim HidBirthDay As HtmlInputHidden = eItem.FindControl("HidBirthDay")
            Dim HidSTDate As HtmlInputHidden = eItem.FindControl("HidSTDate")
            Dim Hid_eSerNum As HtmlInputHidden = eItem.FindControl("Hid_eSerNum")
            Dim HidCMASTER1 As HtmlInputHidden = eItem.FindControl("HidCMASTER1")
            Dim HidCMASTER1NT As HtmlInputHidden = eItem.FindControl("HidCMASTER1NT")
            Dim labSTNAME As Label = eItem.FindControl("labSTNAME")
            'Dim vSTDNAME As String=TIMS.ClearSQM(If(eItem.Cells(cst_姓名) IsNot Nothing AndAlso eItem.Cells(cst_姓名).Text <> "", eItem.Cells(cst_姓名).Text, ""))
            Dim vSTDNAME As String = TIMS.ClearSQM(If(labSTNAME IsNot Nothing AndAlso labSTNAME.Text <> "", labSTNAME.Text, ""))

            'If oTest_flag Then
            '    ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")，依處分日期及年限，仍在處分期間者，e網報名審核，只能審核為失敗。" & vbCrLf
            '    Exit For
            'End If
            Dim xStudInfo As String = ""
            Dim agIDNO As String = "" 'CStr(drET2("IDNO"))
            Dim agOCID1 As String = "" 'CStr(drET2("OCID1"))
            If signUpStatus.Value <> "0" Then
                xStudInfo = ""
                Dim drET2 As DataRow = TIMS.Get_ENTERTYPE2(Hid_eSerNum.Value, objconn)
                If drET2 Is Nothing Then ERRMSG += "(第" & CStr(i) & "行:" & vSTDNAME & ") 查無此報名學員的報名資料，請重新查詢!" & vbCrLf
                If drET2 IsNot Nothing Then
                    agIDNO = TIMS.ClearSQM(drET2("IDNO"))
                    agOCID1 = TIMS.ClearSQM(drET2("OCID1"))
                    TIMS.SetMyValue(xStudInfo, "IDNO", agIDNO)
                    TIMS.SetMyValue(xStudInfo, "OCID1", agOCID1)
                End If

                If signUpStatus1.Checked Then '成功。
                    If sBlackIDNO <> "" AndAlso sBlackIDNO.IndexOf(agIDNO) > -1 Then
                        ERRMSG &= "(第" & CStr(i) & "行:" & vSTDNAME & ")，依處分日期及年限，仍在處分期間者，e網報名審核，只能審核為失敗。" & vbCrLf
                        Exit For
                    End If
                    '具公司/商業負責人身分 '限定計畫執行
                    If TIMS.Cst_NotTPlanID5.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                        '認定為公司負責人且未切結
                        If HidCMASTER1.Value = "Y" AndAlso HidCMASTER1NT.Value = "" Then
                            If tmpMaster <> "" Then tmpMaster &= "、"
                            tmpMaster &= vSTDNAME
                        End If
                    End If
                    '28:產業人才投資計劃
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '檢核學員OOO有參訓時段重疊報名情形，無法儲存，請再確認！
                        'Call TIMS.ChkStudDouble(ERRMSG, eItem.Cells(cst_姓名).Text, xStudInfo)
                        Call TIMS.ChkStudDouble(timsSer1, ERRMSG, vSTDNAME, xStudInfo)

                        'flagYearsOld65=False 'false:否 判斷是否為 六十五歲以上者資格
                        'If TIMS.Check_YearsOld65(HidBirthDay.Value, HidSTDate.Value) Then flagYearsOld65=True 'true:是 判斷是否為 六十五歲以上者資格
                        'If BudID.SelectedValue=Cst_請選擇 OrElse BudID.SelectedValue="" Then ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")審核清單中,有報名學員的預算別與補助比例尚未選擇,請設定!" & vbCrLf
                        'If SupplyID.SelectedValue=Cst_請選擇 OrElse SupplyID.SelectedValue="" Then ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")審核清單中,有報名學員的預算別與補助比例尚未選擇,請設定!" & vbCrLf
                        'If BudID.SelectedValue="97" AndAlso SupplyID.SelectedValue <> "2" Then ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")預算別為協助,補助比例應為100% !" & vbCrLf
                        'If BudID.SelectedValue="99" AndAlso SupplyID.SelectedValue <> "9" Then ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")預算別為不補助,補助比例應為0% !" & vbCrLf
                        '修改說明:有關參訓學員已65歲以上者（依該參訓學員出生年月日及開訓日期判斷），其預算別一律運用就安預算 2013/11/20
                        'true:是 判斷是否為 六十五歲以上者資格
                        'If flagYearsOld65 AndAlso BudID.SelectedValue <> "02" Then ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")" & Cst_Msg65 & vbCrLf
                    Else
                        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                        'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '    If WorkSuppIdent1.Checked=False And WorkSuppIdent2.Checked=False Then ERRMSG &= String.Format("(第{0}行: {1})審核清單中,有報名學員的「是否為在職者補助身分」尚未選擇,請設定!", CStr(i), eItem.Cells(cst_姓名).Text) & vbCrLf
                        'End If
                        'If TIMS.Chk_Master(Me, Chkws1, Hid_eSerNum.Value, objconn)="Y" Then
                        '    If tmpMaster <> "" Then tmpMaster &= ","
                        '    tmpMaster &= eItem.Cells(cst_姓名).Text
                        'End If
                    End If
                End If

                '備註(失敗原因)
                signUpMemo.Text = TIMS.ClearSQM(signUpMemo.Text)
                Const cst_sign_max_text_length As Integer = 150
                If signUpMemo.Text <> "" AndAlso signUpMemo.Text.Length > cst_sign_max_text_length Then
                    ERRMSG &= String.Format("(第{0}行: {1}) 備註(失敗原因),文字敘述過長限 {2}個字元,請重新設定!!", CStr(i), vSTDNAME, cst_sign_max_text_length) & vbCrLf
                End If
                If signUpStatus2.Checked Then '失敗。
                    ''A226974437' 
                    ' (這應是職前邏輯)該學員為非自願離職者，不可審核為失敗
                    'If TIMS.Chk_ENTERTYPEW(agIDNO, agOCID1, objconn) Then ERRMSG += "(第" & CStr(i) & "行:" & eItem.Cells(cst_姓名).Text & ")該學員為非自願離職者，不可審核為失敗!!" & vbCrLf
                    '備註(失敗原因)
                    If signUpMemo.Text = "" Then ERRMSG &= String.Format("(第{0}行: {1}) 請輸入,備註(失敗原因)!!", CStr(i), vSTDNAME) & vbCrLf
                End If
            End If
        Next
        '具公司/商業負責人身分
        If tmpMaster <> "" Then
            tmpMaster = Replace(cst_xMaster, "XXX", tmpMaster)
            ERRMSG += tmpMaster & vbCrLf
        End If
        If ERRMSG <> "" Then
            Common.MessageBox(Me, ERRMSG)
            Exit Sub
        End If
        '====== ERRMSG ====== 

        '====== SAVE START ====== 
        'If TestStr="AmuTest" Then Exit Sub'測試用
        'Dim sqlAdp As New SqlDataAdapter
        Dim path3 As String = TIMS.Utl_GetConfigSet("from_emailaddress")
        Dim vpath3 As String = If(String.IsNullOrEmpty(path3), TIMS.Cst_SendMail3_from_emailaddress, path3)
        Dim flagError1 As Boolean = False '儲存修改資料錯誤
        Dim flagError2 As Boolean = False '異常錯誤
        Dim sErrMsg As String = ""
        Dim sErrMsg1 As String = "" 'sErrMsg
        Dim sErrMsg2 As String = "" 'ex.ToString 
        flagError1 = False
        flagError2 = False
        sErrMsg = ""
        sErrMsg1 = ""
        sErrMsg2 = ""

        Dim tmpConn As SqlConnection = DbAccess.GetConnection()
        Dim tmpTrans As SqlTransaction = DbAccess.BeginTrans(tmpConn)
        Try
            For Each item As DataGridItem In DataGrid1.Items
                '更新e網報名資料。
                sErrMsg = ""
                Call UPDATE_STUD_ENTERTYPE2(tmpConn, tmpTrans, item, sEmailSend, sErrMsg, vpath3, aNow, flag_ROC, flgROLEIDx0xLIDx0, Me, sm)
                If sErrMsg <> "" Then
                    sErrMsg1 += sErrMsg
                    flagError1 = True '儲存修改資料錯誤 'Call DbAccess.RollbackTrans(tmpTrans)
                    Exit For
                End If
            Next
            Call DbAccess.CommitTrans(tmpTrans)
        Catch ex As Exception
            flagError2 = True '異常錯誤
            sErrMsg2 += ex.ToString & vbCrLf
            DbAccess.RollbackTrans(tmpTrans)
            DbAccess.CloseDbConn(tmpConn)
            'Throw ex
        End Try
        DbAccess.CloseDbConn(tmpConn)

        '儲存修改資料錯誤
        If flagError1 Then
            If sErrMsg1 = "" Then sErrMsg1 = TIMS.cst_ErrorMsg11
            Common.MessageBox(Me, sErrMsg1)

            sErrMsg1 += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'sErrMsg1=Replace(sErrMsg1, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(sErrMsg1)
            Exit Sub
        End If
        '異常錯誤
        If flagError2 Then
            If sErrMsg2 = "" Then sErrMsg2 = TIMS.cst_ErrorMsg11
            Common.MessageBox(Me, sErrMsg2)

            sErrMsg2 += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'sErrMsg2=Replace(sErrMsg2, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(sErrMsg2)
            Exit Sub
        End If

        '任何無錯誤
        If Not flagError1 AndAlso Not flagError2 Then
            Common.MessageBox(Me, "儲存成功")
            'Call Button1_Click(sender, e)
            Call Search1()
        Else
            Dim strErrmsg As String = ""
            strErrmsg += "*****儲存異常，請檢查資料是否正確 (請重試)!!*****" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            Exit Sub
        End If
    End Sub

    ''' <summary>INSERT_STUD_ENTERTEMP2</summary>
    ''' <param name="dt"></param>
    ''' <param name="ieSETID"></param>
    Sub INSERT_STUD_ENTERTEMP2(ByRef dt As DataTable, ByRef ieSETID As Integer)
        'ieSETID=TIMS.Get_eSETID_MaxID(aIDNO, objconn, trans)
        '如未有此學員的線上報名資料.則新增一筆報名學員資料
        Dim dr As DataRow = dt.NewRow()
        dt.Rows.Add(dr)
        dr("eSETID") = ieSETID 'sqldr("MaxID")
        dr("IDNO") = aIDNO
        Call ImpDr_STUD_ENTERTEMP2(dr)
    End Sub

    ''' <summary>UPDATE STUD_ENTERTEMP2</summary>
    ''' <param name="dt"></param>
    ''' <param name="ieSETID"></param>
    ''' <param name="iSETID"></param>
    Sub UPDATE_STUD_ENTERTEMP2(ByRef dt As DataTable, ByRef ieSETID As Integer, ByRef iSETID As Integer)
        For y As Integer = 0 To dt.Rows.Count - 1
            Dim dr As DataRow = dt.Rows(y)
            ieSETID = dr("eSETID") '取得最後1筆(eSETID)
            If Convert.ToString(dr("SETID")) <> "" Then iSETID = dr("SETID") '取得最後1筆(SETID)
            Call ImpDr_STUD_ENTERTEMP2(dr)
        Next
    End Sub

    Sub ImpDr_STUD_ENTERTEMP2(ByRef dr As DataRow)
        dr("Name") = TIMS.HtmlDecode1(aName)
        dr("Sex") = aSex
        dr("Birthday") = aBirthday
        Dim vPassPortNo As String = "2"
        Select Case Convert.ToString(aPassPortNO)
            Case "1", "2"
                vPassPortNo = aPassPortNO
        End Select
        dr("PassPortNo") = vPassPortNo
        dr("MaritalStatus") = Convert.DBNull '婚姻狀況
        dr("DegreeID") = aDegreeID
        dr("GradID") = "01" '畢業'Me.GraduateStatus.SelectedValue
        '若為空值則輸入不詳
        '若為空值則輸入不詳
        dr("School") = If(Convert.ToString(dr("School")) = "", TIMS.cst_未填寫, dr("School"))
        dr("Department") = If(Convert.ToString(dr("Department")) = "", TIMS.cst_未填寫, dr("Department"))
        dr("MilitaryID") = "" '兵役'Me.MilitaryID.SelectedValue
        dr("ZipCode") = aZipCode
        dr("ZipCODE6W") = If(aZipCODE6W <> "", aZipCODE6W, Convert.DBNull)
        dr("Address") = aAddress
        aPhone1 = TIMS.ChangeIDNO(aPhone1)
        aPhone2 = TIMS.ChangeIDNO(aPhone2)
        aCellPhone = TIMS.ChangeIDNO(aCellPhone)
        dr("Phone1") = If(aPhone1 <> "", aPhone1, Convert.DBNull) 'TIMS.ChangeIDNO(aPhone1)
        dr("Phone2") = If(aPhone2 <> "", aPhone2, Convert.DBNull) 'TIMS.ChangeIDNO(aPhone2)
        dr("CellPhone") = If(aCellPhone <> "", aCellPhone, Convert.DBNull)
        '**by AMU --資安問題過濾Email文字
        Dim vEmail As String = TIMS.ChangeEmail(TIMS.ClearSQM(dr("Email")))
        dr("Email") = If(aEmail <> "", aEmail, vEmail)
        dr("IsAgree") = aIsAgree
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = aNow 'Now
    End Sub

    ''' <summary> INSERT STUD_ENTERTEMP</summary>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    ''' <param name="sm"></param>
    ''' <param name="dr1">STUD_ENTERTEMP2 / STUD_ENTERTYPE2  GET_STUDENTERTYPE2(tmpConn, tmpTrans, Hid_eSerNum.Value) </param>
    ''' <param name="iSETID">STUD_ENTERTEMP-SETID </param>
    Public Shared Sub INSERT_STUD_ENTERTEMP(ByRef tmpConn As SqlConnection, ByRef tmpTrans As SqlTransaction, ByRef sm As SessionModel,
                                            ByRef dr1 As DataRow, ByRef iSETID As Integer)
        'Dim sqlAdp As New SqlDataAdapter
        'Dim qryAdp As New SqlDataAdapter
        If iSETID > 0 Then Return

        Dim sqlStr As String = ""
        sqlStr &= " INSERT INTO STUD_ENTERTEMP (SETID,IDNO,Name,Sex,Birthday,PassPortNO,MaritalStatus,DegreeID,GradID,School,Department,MilitaryID,ZipCode,Address,Phone1,Phone2,CellPhone,Email,eSETID,ModifyAcct,ModifyDate,ZipCODE6W,LAINFLAG) "
        sqlStr &= " VALUES(@SETID,@IDNO,@Name,@Sex,@Birthday,@PassPortNO,@MaritalStatus,@DegreeID,@GradID,@School,@Department,@MilitaryID,@ZipCode,@Address,@Phone1,@Phone2,@CellPhone,@Email,@eSETID,@ModifyAcct,getdate(),@ZipCODE6W,@LAINFLAG) "
        Dim sqlAdp As New SqlDataAdapter
        sqlAdp.InsertCommand = New SqlCommand(sqlStr, tmpConn, tmpTrans)

        Dim v_PassPortNO As Integer = 2
        Select Case Convert.ToString(dr1("PassPortNO"))
            Case "1", "2"
                v_PassPortNO = Convert.ToInt32(dr1("PassPortNO"))
        End Select
        With sqlAdp.InsertCommand
            iSETID = DbAccess.GetNewId(tmpTrans, "STUD_ENTERTEMP_SETID_SEQ,STUD_ENTERTEMP,SETID")
            '.InsertCommand=New SqlCommand(sqlStr, tmpConn, tmpTrans)
            .Parameters.Clear()
            .Parameters.Add("@SETID", SqlDbType.VarChar).Value = iSETID
            .Parameters.Add("@IDNO", SqlDbType.VarChar).Value = TIMS.ChangeIDNO(dr1("IDNO"))
            .Parameters.Add("@Name", SqlDbType.NVarChar).Value = TIMS.HtmlDecode1(Convert.ToString(dr1("Name")))
            .Parameters.Add("@Sex", SqlDbType.Char).Value = Convert.ToString(dr1("Sex"))
            .Parameters.Add("@Birthday", SqlDbType.DateTime).Value = Convert.ToDateTime(dr1("Birthday"))
            .Parameters.Add("@PassPortNO", SqlDbType.Int).Value = v_PassPortNO
            .Parameters.Add("@MaritalStatus", SqlDbType.Int).Value = If(Convert.ToString(dr1("MaritalStatus")) <> "", Convert.ToInt32(dr1("MaritalStatus")), Convert.DBNull)
            .Parameters.Add("@DegreeID", SqlDbType.VarChar).Value = Convert.ToString(dr1("DegreeID"))
            .Parameters.Add("@GradID", SqlDbType.VarChar).Value = Convert.ToString(dr1("GradID"))
            .Parameters.Add("@School", SqlDbType.NVarChar).Value = Convert.ToString(dr1("School"))
            .Parameters.Add("@Department", SqlDbType.NVarChar).Value = Convert.ToString(dr1("Department"))
            .Parameters.Add("@MilitaryID", SqlDbType.VarChar).Value = Convert.ToString(dr1("MilitaryID"))
            .Parameters.Add("@ZipCode", SqlDbType.Int).Value = Convert.ToString(dr1("ZipCode"))
            .Parameters.Add("@Address", SqlDbType.NVarChar).Value = Convert.ToString(dr1("Address"))

            .Parameters.Add("@Phone1", SqlDbType.VarChar).Value = If(Convert.ToString(dr1("Phone1")) <> "", Convert.ToString(dr1("Phone1")), Convert.DBNull)
            .Parameters.Add("@Phone2", SqlDbType.VarChar).Value = If(Convert.ToString(dr1("Phone2")) <> "", Convert.ToString(dr1("Phone2")), Convert.DBNull)
            .Parameters.Add("@CellPhone", SqlDbType.VarChar).Value = If(Convert.ToString(dr1("CellPhone")) <> "", Convert.ToString(dr1("CellPhone")), Convert.DBNull)

            .Parameters.Add("@Email", SqlDbType.VarChar).Value = If(Convert.ToString(dr1("Email")) <> "", Convert.ToString(dr1("Email")), Convert.DBNull)
            .Parameters.Add("@eSETID", SqlDbType.Int).Value = If(Convert.ToString(dr1("eSETID")) <> "", Convert.ToInt32(dr1("eSETID")), Convert.DBNull)
            .Parameters.Add("@ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
            .Parameters.Add("@ZipCODE6W", SqlDbType.VarChar).Value = If(Convert.ToString(dr1("ZipCODE6W")) <> "", Convert.ToString(dr1("ZipCODE6W")), Convert.DBNull)
            .Parameters.Add("@LAINFLAG", SqlDbType.Int).Value = 0
            .ExecuteNonQuery()

            'DbAccess.ExecuteNonQuery(sqlAdp.InsertCommand.CommandText, tmpConn, sqlAdp.InsertCommand.Parameters)
            'SETID=DbAccess.GetId(objconn, "AUTH_ACCRWPLANTEMP_ACCTPID_SEQ")
            'SETID=DbAccess.GetId(tmpTrans, "STUD_ENTERTEMP_SETID_SEQ")
        End With
        'vsSETID=SETID
        'sqlAdp.Dispose()
        'sqlAdp=Nothing
    End Sub

    ''' <summary>INSERT STUD_ENTERTYPE</summary>
    ''' <param name="tmpTrans"></param>
    ''' <param name="sm"></param>
    ''' <param name="dr1">STUD_ENTERTEMP2 / STUD_ENTERTYPE2  GET_STUDENTERTYPE2(tmpConn, tmpTrans, Hid_eSerNum.Value)</param>
    ''' <param name="iSETID">STUD_ENTERTEMP-SETID </param>
    ''' <param name="iSEID">STUD_ENTERTRAIN2-SEID</param>
    ''' <param name="NewExamNO"></param>
    ''' <param name="V_HidIdentityID"></param>
    ''' <param name="aNow"></param>
    ''' <returns>iSerNum</returns>
    Public Shared Function INSERT_STUD_ENTERTYPE(ByRef tmpTrans As SqlTransaction, ByRef sm As SessionModel, ByRef dr1 As DataRow,
                                                 ByRef iSETID As Integer, ByRef iSEID As Integer, ByVal NewExamNO As String,
                                                 ByVal V_HidIdentityID As String, ByVal aNow As Date) As Integer
        Dim iSerNum As Integer = 1

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing

        '某日某學員
        sql = " SELECT * FROM STUD_ENTERTYPE WHERE SETID='" & iSETID & "' AND EnterDate=" & TIMS.To_date(dr1("EnterDate")) & " ORDER BY SerNum DESC "
        dt = DbAccess.GetDataTable(sql, da, tmpTrans)

        If dt.Rows.Count = 0 Then
            '完全沒有當日資料。
            iSerNum = 1 'SerNum 起始為1
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SETID") = iSETID
            dr("EnterDate") = dr1("EnterDate")
            dr("SerNum") = iSerNum
        Else
            'Dim double_OCID As Boolean=False '同一SETID(學員)，產生重複報名同一班(OCID)
            'double_OCID=False 'Dim i As Integer=0 'i=0 '某班
            Dim s_FINDT1 As String = String.Concat("OCID1=", dr1("OCID1"))
            If dt.Select(s_FINDT1).Length > 0 Then
                'double_OCID=True
                '同一SETID(學員)，產生重複報名同一班(OCID)
                'SerNum=dt.Select("OCID1='" & dr1("OCID1") & "'")(0)("SerNum") + 1
                dr = dt.Select(s_FINDT1)(0)
                iSerNum = dr("SerNum")
            Else
                '查無 同OCID
                'Not 同一SETID(學員)，產生重複報名 非同一班(OCID)
                iSerNum = dt.Rows(0)("SerNum") + 1
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("SETID") = iSETID
                dr("EnterDate") = dr1("EnterDate")
                dr("SerNum") = iSerNum
            End If
        End If
        'Dim vsRelEnterDate As String=""
        '報名日期(轉民國)/AD
        'If Convert.ToString(dr1("RelEnterDate")) <> "" Then
        '    vsRelEnterDate=If(flag_ROC, Common.FormatDate2Roc(dr1("RelEnterDate")), dr1("RelEnterDate"))
        'End If
        dr("RelEnterDate") = dr1("RelEnterDate")
        dr("ExamNo") = NewExamNO '准考證號
        'vsExamNo=NewExamNO '准考證號
        dr("OCID1") = dr1("OCID1")
        dr("TMID1") = dr1("TMID1")
        dr("OCID2") = dr1("OCID2")
        dr("TMID2") = dr1("TMID2")
        dr("OCID3") = dr1("OCID3")
        dr("TMID3") = dr1("TMID3")
        '就服站不可異動
        Dim v_ENTERCHANNEL As Integer = If(Convert.ToString(dr1("EnterPath")) = "o", 2, 1) '(o內部報名：2:現場／(其它)外部：1:網路))
        dr("EnterChannel") = v_ENTERCHANNEL 'If(Convert.ToString(dr1("EnterPath")) <> "W", 1, Convert.DBNull)
        dr("EnterPath") = If(Convert.ToString(dr1("EnterPath")) <> "W", "B", Convert.DBNull) 'E網批次審核
        'V_HidIdentityID=TIMS.ClearSQM(HidIdentityID.Value)
        dr("IdentityID") = If(V_HidIdentityID <> "", V_HidIdentityID, dr1("IdentityID"))
        dr("MIDENTITYID") = dr1("MIDENTITYID")
        dr("RID") = dr1("RID")
        dr("PlanID") = dr1("PlanID")
        dr("CCLID") = dr1("CCLID")
        dr("eSerNum") = dr1("eSerNum")
        dr("eSETID") = dr1("eSETID")
        '-受訓前任職資料start-
        dr("ActNo") = If(Convert.ToString(dr1("ActNo")) = "", Convert.DBNull, Convert.ToString(dr1("ActNo")))
        dr("PriorWorkType1") = If(Convert.ToString(dr1("PriorWorkType1")) = "", Convert.DBNull, Convert.ToString(dr1("PriorWorkType1")))
        dr("PriorWorkOrg1") = If(Convert.ToString(dr1("PriorWorkOrg1")) = "", Convert.DBNull, Convert.ToString(dr1("PriorWorkOrg1")))
        dr("SOfficeYM1") = If(Convert.ToString(dr1("SOfficeYM1")) = "", Convert.DBNull, Convert.ToString(dr1("SOfficeYM1")))
        dr("FOfficeYM1") = If(Convert.ToString(dr1("FOfficeYM1")) = "", Convert.DBNull, Convert.ToString(dr1("FOfficeYM1")))
        '-受訓前任職資料end-
        dr("CMASTER1") = dr1("CMASTER1") '(公司負責人勾稽)
        dr("TransDate") = aNow 'Now
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = aNow 'Now

        '28:產業人才投資計劃
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If BudID.SelectedValue=Cst_請選擇 Then dr("BudID")=Convert.DBNull Else dr("BudID")=BudID.SelectedValue
        '    If SupplyID.SelectedValue=Cst_請選擇 Then dr("SupplyID")=Convert.DBNull Else dr("SupplyID")=SupplyID.SelectedValue
        'Else
        '    是否為在職者補助身分 46: 補助辦理保母職業訓練 '47:補助辦理照顧服務員職業訓練
        '    If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        Dim v_WorkSuppIdent As String=""
        '        If WorkSuppIdent1.Checked Then v_WorkSuppIdent="Y"
        '        If WorkSuppIdent2.Checked Then v_WorkSuppIdent="N"
        '        dr("WorkSuppIdent")=If(v_WorkSuppIdent <> "", v_WorkSuppIdent, Convert.DBNull)
        '        dr("BudID")=Convert.DBNull
        '        dr("SupplyID")=Convert.DBNull
        '    Else
        '        dr("WorkSuppIdent")=Convert.DBNull
        '        dr("BudID")=Convert.DBNull
        '        dr("SupplyID")=Convert.DBNull
        '    End If
        'End If
        '非 28:產業人才投資計劃 清空資料
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            dr("WorkSuppIdent") = Convert.DBNull
            dr("BudID") = Convert.DBNull
            dr("SupplyID") = Convert.DBNull
        End If
        If iSEID <> 0 Then dr("SEID") = iSEID
        'Dim sTemp1 As String=""
        'sTemp1="CFIRE1,CFIRE1NS,CFIRE1REASON,CFIRE1MACCT,CFIRE1MDATE"
        'sTemp1 &= ",CMASTER1,CMASTER1NS,CMASTER1REASON,CMASTER1MACCT,CMASTER1MDATE,CMASTER1NT,CFIRE1R2"
        'Dim aTemp1 As String()=sTemp1.Split(",")
        'For i As Integer=0 To aTemp1.Length - 1
        '    Dim tmpCT1 As String=aTemp1(i)
        '    'type2 有值 type 無值 才動作
        '    If Convert.ToString(dr1(tmpCT1)) <> "" AndAlso Convert.ToString(dr(tmpCT1))="" Then
        '        dr(tmpCT1)=dr1(tmpCT1)
        '    End If
        'Next
        DbAccess.UpdateDataTable(dt, da, tmpTrans)
        Return iSerNum
    End Function

    ''' <summary>(儲存) 更新e網報名資料。</summary>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    ''' <param name="item"></param>
    ''' <param name="sEmailSend"></param>
    ''' <param name="sErrMsg"></param>
    ''' <param name="vpath3"></param>
    ''' <param name="aNow"></param>
    ''' <param name="MyPage"></param>
    Public Shared Sub UPDATE_STUD_ENTERTYPE2(ByRef tmpConn As SqlConnection, ByRef tmpTrans As SqlTransaction, ByRef item As DataGridItem,
                                             ByVal sEmailSend As String, ByRef sErrMsg As String, vpath3 As String, aNow As Date,
                                             ByRef flag_ROC As Boolean, ByRef flgROLEIDx0xLIDx0 As Boolean,
                                             ByVal MyPage As Page, ByRef sm As SessionModel)
        'sErrMsg=""
        Dim signUpStatus1 As HtmlInputRadioButton = item.FindControl("signUpStatus1")
        Dim signUpStatus2 As HtmlInputRadioButton = item.FindControl("signUpStatus2")
        Dim signUpStatus As HtmlInputHidden = item.FindControl("signUpStatus")
        Dim Hid_eSerNum As HtmlInputHidden = item.FindControl("Hid_eSerNum")
        Dim signUpMemo As TextBox = item.FindControl("signUpMemo") '備註(失敗原因)
        'Dim BudID As DropDownList=item.FindControl("BudID")
        'Dim SupplyID As DropDownList=item.FindControl("SupplyID")
        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        'TIMS.Cst_TPlanID46AppPlan5
        Dim WorkSuppIdent1 As HtmlInputRadioButton = item.FindControl("WorkSuppIdent1")
        Dim WorkSuppIdent2 As HtmlInputRadioButton = item.FindControl("WorkSuppIdent2")
        Dim HidIdentityID As HtmlInputHidden = item.FindControl("HidIdentityID")
        Dim Hid_MSG1NAME As HtmlInputHidden = item.FindControl("Hid_MSG1NAME")
        Dim Hid_MSGTYPEN As HtmlInputHidden = item.FindControl("Hid_MSGTYPEN")
        Dim Hid_MSGADIDN As HtmlInputHidden = item.FindControl("Hid_MSGADIDN")

        Dim sql As String = ""
        Dim iSETID As Integer = 0 '(初值為0)

        Dim s_tmpErr1 As String = ""
        If signUpStatus.Value <> "0" Then
            'Dim dr1 As DataRow=Nothing 'STUD_ENTERTEMP2 / STUD_ENTERTYPE2
            'Dim dr2 As DataRow=Nothing 'STUD_ENTERTEMP
            '取得單1報名資料。 STUD_ENTERTEMP2,STUD_ENTERTYPE2,STUD_ENTERSUBDATA2
            Dim dr1 As DataRow = GET_STUDENTERTYPE2(tmpConn, tmpTrans, Hid_eSerNum.Value)
            'Hid_eSerNum.Value 'dr1=Get_StudEnterType2(tmpConn, tmpTrans,DataGrid1.DataKeys(item.ItemIndex))
            If dr1 Is Nothing Then
                s_tmpErr1 = String.Format("第 {0} 筆 該筆資料已被刪除，請重新查詢資料狀態({1})", (item.ItemIndex + 1), Hid_eSerNum.Value)
                sErrMsg &= s_tmpErr1 & vbCrLf
                Exit Sub
            End If
            'If Hid_MSGADIDN.Value <> "" Then
            '    Dim ss As String=""
            '    Call TIMS.SetMyValue(ss, "ADID", Hid_MSGADIDN.Value)
            '    Call TIMS.SetMyValue(ss, "OCID", dr1("OCID1"))
            '    Call TIMS.SetMyValue(ss, "IDNO", dr1("IDNO"))
            '    Call TIMS.SetMyValue(ss, "ESETID", dr1("ESETID"))
            '    Call TIMS.SUtl_AddDISASTER(MyPage, ss, tmpTrans)
            'End If
            Dim iSEID As Integer = 0 'STUD_ENTERTRAIN2-SEID
            If dr1 IsNot Nothing Then
                '把SEID(線上報名資料的流水號-產學訓)寫入Stud_Entertype中---start
                Dim objstr As String = String.Concat(" SELECT SEID FROM STUD_ENTERTRAIN2 WHERE eSerNum=", dr1("eSerNum"))
                Dim objdr As DataRow = DbAccess.GetOneRow(objstr, tmpTrans)
                If objdr IsNot Nothing Then iSEID = objdr("SEID")
            End If

            Dim vIDNO As String = TIMS.ChangeIDNO(TIMS.ClearSQM(dr1("IDNO")))

            Dim sql2 As String = String.Format("SELECT * FROM STUD_ENTERTEMP WHERE IDNO='{0}'", vIDNO)
            Dim dr2 As DataRow = DbAccess.GetOneRow(sql2, tmpTrans) 'STUD_ENTERTEMP
            iSETID = If(dr2 IsNot Nothing, Val(dr2("SETID")), 0)

            Dim vsStud_name As String = ""
            Dim vsSubject As String = ""
            Dim vsRID As String = ""
            Dim vsCheckInDate As String = ""
            Dim vsExamDate As String = ""
            Dim vsExamNo As String = ""
            Dim vsRelEnterDate As String = ""
            Dim vssignUpMemo As String = "" '備註(失敗原因)
            vsStud_name = TIMS.HtmlDecode1(dr1("Name")) '報考人姓名
            Dim drCC As DataRow = TIMS.GetOCIDDate(Convert.ToString(dr1("OCID1")), tmpConn, tmpTrans)
            If drCC Is Nothing Then
                s_tmpErr1 = String.Format("第 {0} 筆 查無班級資料， 班級資料有誤!({1}, {2})", (item.ItemIndex + 1), dr1("OCID1"), Hid_eSerNum.Value)
                sErrMsg &= s_tmpErr1 & vbCrLf
                Exit Sub
            End If
            vsSubject = TIMS.Get_Subject(drCC)
            vsRID = Convert.ToString(drCC("OrgName")) 'TIMS.Get_OrgNameInputRID(dr1("RID").ToString, objconn)
            'vssignUpMemo=Convert.ToString(dr1("signUpMemo"))
            'Dim drX As DataRow=TIMS.GetOCIDDate(dr1("OCID1"), objconn)
            vsCheckInDate = ""
            If Convert.ToString(drCC("CheckInDate")) <> "" Then
                vsCheckInDate = If(flag_ROC, Common.FormatDate2Roc(drCC("CheckInDate")), TIMS.Cdate3(drCC("CheckInDate")))
            End If
            vsExamDate = ""
            If Convert.ToString(drCC("ExamDate")) <> "" Then
                vsExamDate = If(flag_ROC, Common.FormatDate2Roc(drCC("ExamDate")), TIMS.Cdate3(drCC("ExamDate")))
            End If

            '取出准考證號   Start
            Dim flagChkPlanT As Boolean = True
            Select Case Convert.ToString(sm.UserInfo.LID)
                Case "0"
                    If Convert.ToString(sm.UserInfo.TPlanID) <> Convert.ToString(drCC("TPlanID")) Then flagChkPlanT = False
                Case Else
                    If Convert.ToString(sm.UserInfo.PlanID) <> Convert.ToString(dr1("PlanID")) Then flagChkPlanT = False
            End Select
            If Not flgROLEIDx0xLIDx0 AndAlso Not flagChkPlanT Then
                sErrMsg &= String.Concat("第 ", item.ItemIndex + 1, " 筆")
                sErrMsg &= "班級的計畫代號 與登入者不同， 請選擇正確計畫登入審核。" & vbCrLf
                Exit Sub
            End If

            Dim ExamOcid1 As String = dr1("OCID1").ToString
            'TIMS.Get_ExamNo1(ExamOcid1, objconn)
            Dim ExamNo1 As String = drCC("ExamNO1") '取出班級的CLASSID +期別 成為准考證編碼的前面的固定碼
            If ExamNo1 = "" OrElse ExamNo1.Length < 6 Then '防呆
                sErrMsg &= String.Concat("第 ", item.ItemIndex + 1, " 筆")
                sErrMsg &= "班級的代號與期別有誤，請確認班級狀態" & vbCrLf
                Exit Sub
            End If

            Dim NewExamNO As String = ""
            Dim ExamPlanID As String = sm.UserInfo.PlanID '.ToString '"" 'dr("PlanID").ToString
            '超級使用者 
            If sm.UserInfo.LID = 0 AndAlso flgROLEIDx0xLIDx0 Then
                ExamPlanID = Convert.ToString(drCC("PlanID")) '.ToString '"" 'dr("PlanID").ToString
            End If
            'Dim flgChkExamNo As Boolean=TIMS.Chk_NewExamNOc(ExamPlanID, ExamOcid1, objconn)
            Dim flgChkExamNo As Boolean = True
            If ExamPlanID <> Convert.ToString(drCC("PlanID")) Then flgChkExamNo = False
            If Not flgChkExamNo Then
                s_tmpErr1 = String.Format("第 {0} 筆 班級的代號 與計畫不符， 請確認班級狀態(取出准考證號)", (item.ItemIndex + 1))
                sErrMsg &= s_tmpErr1 & vbCrLf
                Exit Sub
            End If
            '准考證號
            NewExamNO = TIMS.Get_NewExamNOt(ExamPlanID, ExamNo1, ExamOcid1, tmpTrans)
            If NewExamNO = "" Then
                s_tmpErr1 = String.Format("第 {0} 筆 班級的代號 與計畫不符， 請確認班級狀態(取出准考證號)", (item.ItemIndex + 1))
                sErrMsg &= s_tmpErr1 & vbCrLf
                Exit Sub
            End If
            '取出准考證號   End

            '錯誤離開
            If sErrMsg <> "" Then Exit Sub

            If signUpStatus1.Checked Then
                '錄取 報名成功
                '**by Milor 20081016--Stud_EnterTemp的SETID有做Identity，所以不能用這種方法來處理。
                ' 正確應該只在有過報名資料才取得過去的SETID，沒有就直接Insert取得。
                '**by Milor 20081016--當SETID沒有的時候，表示新報名的學員，從Insert Stud_EnterTemp中取得SETID。
                '20090507(Milor)--當有SETID存在時，將Stud_EnterTemp的資料備份，並更新成Stud_EnterTemp2的資料。
                If iSETID <= 0 Then
                    Call INSERT_STUD_ENTERTEMP(tmpConn, tmpTrans, sm, dr1, iSETID)
                Else
                    '有 Stud_EnterTemp資料做修改
                    If Convert.ToString(dr1("IDNO")) <> "" Then Call TIMS.UPDATE_STUDENTERTEMP(MyPage, Convert.ToString(dr1("IDNO")), tmpConn, tmpTrans)
                End If
                '取出准考證號   End

                '有特定的參訓身份
                HidIdentityID.Value = TIMS.ClearSQM(HidIdentityID.Value)
                '報名日期(轉民國)/AD
                vsRelEnterDate = If(Convert.ToString(dr1("RelEnterDate")) <> "", If(flag_ROC, Common.FormatDate2Roc(dr1("RelEnterDate")), dr1("RelEnterDate")), "")
                '准考證號 寫入vsExamNo
                vsExamNo = NewExamNO '准考證號
                '新增報名登錄
                Dim iSerNum As Integer = INSERT_STUD_ENTERTYPE(tmpTrans, sm, dr1, iSETID, iSEID, NewExamNO, HidIdentityID.Value, aNow)

                'UPDATE STUD_ENTERTEMP2 
                Dim parms As New Hashtable From {{"SETID", iSETID}, {"ModifyAcct", sm.UserInfo.UserID}, {"ModifyDate", aNow}, {"eSETID", Val(dr1("eSETID"))}}
                Dim sql_u As String = ""
                sql_u &= " UPDATE STUD_ENTERTEMP2"
                sql_u &= " SET SETID=@SETID,ModifyAcct=@ModifyAcct, ModifyDate=@ModifyDate"
                sql_u &= " WHERE eSETID=@eSETID"
                DbAccess.ExecuteNonQuery(sql_u, tmpTrans, parms)

                'UPDATE STUD_ENTERTYPE2 
                Dim parms2 As New Hashtable From {{"SETID", iSETID}, {"ExamNo", NewExamNO}, {"SerNum", iSerNum},
                    {"ModifyAcct", sm.UserInfo.UserID}, {"ModifyDate", aNow}, {"eSerNum", Val(dr1("eSerNum"))}}
                Dim sql_u2 As String = ""
                sql_u2 &= " UPDATE STUD_ENTERTYPE2"
                sql_u2 &= " SET SETID=@SETID, ExamNo =@ExamNo, SerNum =@SerNum"
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                sql_u2 &= " ,signUpStatus=1 ,ModifyAcct=@ModifyAcct ,ModifyDate=@ModifyDate"
                sql_u2 &= " WHERE eSerNum =@eSerNum"
                DbAccess.ExecuteNonQuery(sql_u2, tmpTrans, parms2)

                '假如是插班,則直接進入參訓狀態
                '如果是產投，則直接進入錄取狀態(Admission=null,SelResultID='03')
                If Convert.ToString(dr1("CCLID")) <> "" OrElse TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If iSerNum <> 0 Then
                        'If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then dr("Admission")="Y"
                        Dim vAdmission As String = "Y" '錄取
                        'SELRESULTID: 01:正取 02:備取 03:未錄取 04:缺考 05:審核中
                        Dim vSelResultID As String = TIMS.cst_SelResultID_正取 '01:正取
                        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                            vAdmission = "" 'Convert.DBNull '是否錄取 未填寫
                            vSelResultID = TIMS.cst_SelResultID_審核中 '05:審核中 03:不錄取(未錄取)
                        End If
                        Dim dr As DataRow = Nothing
                        Dim dt As DataTable = Nothing
                        Dim da As SqlDataAdapter = Nothing
                        sql = String.Concat("SELECT * FROM STUD_SELRESULT WHERE SETID=", iSETID, "  AND EnterDate=", TIMS.To_date(dr1("EnterDate")), " AND SerNum=", iSerNum)
                        dt = DbAccess.GetDataTable(sql, da, tmpTrans)
                        If dt.Rows.Count = 0 Then
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("SETID") = iSETID
                            dr("EnterDate") = dr1("EnterDate")
                            dr("SerNum") = iSerNum
                        Else
                            dr = dt.Rows(0)
                        End If
                        dr("OCID") = dr1("OCID1")
                        dr("Admission") = If(vAdmission <> "", vAdmission, Convert.DBNull) '"Y" '錄取
                        dr("SelResultID") = If(vSelResultID <> "", vSelResultID, Convert.DBNull) '01:正取
                        dr("RID") = sm.UserInfo.RID
                        dr("PlanID") = sm.UserInfo.PlanID
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = aNow 'Now
                        DbAccess.UpdateDataTable(dt, da, tmpTrans)
                    End If
                    Call TIMS.Update_ClassInfoIsCalculate("Y", dr1("OCID1"), tmpTrans) '是否試算 SqlTransaction
                End If
                'flagtmpTransCommit=False '未完成 Commit 'tmpTrans.Commit() 'flagtmpTransCommit=True '已經完成 Commit

                '20090601 by Jimmy 取得e網報名審核成功「說明事項」內容 -- begin
                Dim tmpdt As DataTable = Nothing
                Dim tmpdr As DataRow = Nothing
                Dim tmpDistID As String = ""
                Dim tmpOrgID As String = ""

                sql = String.Concat("SELECT DISTID FROM AUTH_RELSHIP WHERE RID='", dr1("RID"), "'")
                tmpDistID = DbAccess.ExecuteScalar(sql, tmpTrans)
                If tmpDistID Is Nothing Then tmpDistID = ""

                sql = String.Concat("SELECT b.OrgID FROM AUTH_RELSHIP a JOIN ORG_ORGINFO b ON a.ORGID=b.ORGID WHERE a.RID='", dr1("RID"), "'")
                tmpOrgID = DbAccess.ExecuteScalar(sql, tmpTrans)
                If tmpOrgID Is Nothing Then tmpOrgID = ""

                '20090618 by Jimmy 修改 依需求訓練機構(中心)都要for不同年度、計畫、轄區
                tmpdt = TIMS.Get_FinalEComment("Class", Convert.ToString(tmpOrgID), Convert.ToString(dr1("OCID1")), Convert.ToString(dr1("RID")), Convert.ToString(tmpDistID), Convert.ToString(sm.UserInfo.PlanID), Nothing, tmpConn, tmpTrans)

                Dim vsEComment As String = ""
                If tmpdt IsNot Nothing AndAlso tmpdt.Rows.Count > 0 Then
                    tmpdr = tmpdt.Rows(0)
                    vsEComment = Convert.ToString(tmpdr("eComment"))
                End If
                '20090601 by Jimmy 取得e網報名審核成功「說明事項」內容 -- end

                'Me.vsEmailSend 為發送或不發送
                Dim vEmail As String = TIMS.ChangeEmail(TIMS.ClearSQM(dr1("Email")))
                If vEmail <> "" AndAlso sEmailSend Then
                    'If dr1("Email").ToString <> "" Then '測試用上一句正確
                    '修正錯誤的EMAIL增加發信成功率 BY AMU
                    dr1("Email") = vEmail
                    '20090601 by Jimmy add 依需求將「甄試日期」改為「說明事項」，新增 eComment 參數
                    Dim htSS As New Hashtable From {
                        {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
                        {"Stud_Name", vsStud_name},
                        {"Subject", vsSubject},
                        {"ExamNo", vsExamNo},
                        {"RelEnterDate", vsRelEnterDate},
                        {"ExamDate", vsExamDate},
                        {"CheckInDate", vsCheckInDate},
                        {"EComment", vsEComment},
                        {"Email", TIMS.ChangeEmail(Convert.ToString(dr1("Email")))},
                        {"from_emailaddress", vpath3},
                        {"signUpMemo", ""},
                        {"sRIDOrgName", ""},
                        {"sType", TIMS.Cst_SendMail3_CheckedOK}
                    } 'htSS Hashtable() 
                    Dim mail_msg As String = TIMS.SendMail3(htSS)
                    'If mail_msg <> "" Then Common.RespWrite(Me, "<script>alert('" & mail_msg & "');</script>")
                    If mail_msg <> "" Then
                        Common.RespWrite(MyPage, "<script>alert('" & mail_msg & "');</script>")
                    Else
                        Dim pms_u As New Hashtable From {{"eSerNum", Val(Hid_eSerNum.Value)}}
                        Dim sqlstr_u As String = " UPDATE STUD_ENTERTYPE2 SET isEmailFail='O' WHERE eSerNum=@eSerNum"
                        DbAccess.ExecuteNonQuery(sqlstr_u, tmpTrans, pms_u)
                    End If
                End If
                vsStud_name = Nothing
                vsSubject = Nothing
                vsExamNo = Nothing
                vsRelEnterDate = Nothing
                vsExamDate = Nothing
                vsCheckInDate = Nothing
                'Me.ViewState("Email")=Nothing
                'Me.vssignUpMemo=Nothing
                vsRID = Nothing

            ElseIf signUpStatus2.Checked Then
                '不錄取 2 : 報名失敗
                'Hid_eSerNum.Value
                Dim vsisEmailFail As String = Convert.ToString(dr1("isEmailFail"))  '發送失敗E-MAIL
                Dim vsECommentEmpty As String = "" '空值
                'UPDATE STUD_ENTERTYPE2 
                signUpMemo.Text = TIMS.ClearSQM(signUpMemo.Text)
                vssignUpMemo = signUpMemo.Text 'Replace(signUpMemo.Text.ToString, "'", "''") '備註(失敗原因)
                Dim parms2 As New Hashtable
                parms2.Clear()
                parms2.Add("signUpMemo", vssignUpMemo)

                parms2.Add("ModifyAcct", sm.UserInfo.UserID)
                parms2.Add("ModifyDate", aNow)
                parms2.Add("eSerNum", Val(Hid_eSerNum.Value))
                sql = ""
                sql &= " UPDATE STUD_ENTERTYPE2"
                sql &= " SET signUpMemo=@signUpMemo"
                'signUpStatus 0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                sql &= " ,signUpStatus=2" '2 : 報名失敗
                sql &= " ,ModifyAcct=@ModifyAcct"
                sql &= " ,ModifyDate=@ModifyDate"
                sql &= " WHERE eSerNum =@eSerNum"
                DbAccess.ExecuteNonQuery(sql, tmpTrans, parms2)

                If iSETID > 0 Then
                    '有現場報名資料產生。
                    sql = "" & vbCrLf
                    sql &= " SELECT b.SETID ,b.enterdate ,b.sernum ,b.OCID1" & vbCrLf
                    sql &= " FROM Stud_EnterType b" & vbCrLf
                    sql &= " JOIN Stud_EnterTemp a ON a.setid=b.setid" & vbCrLf
                    sql &= " WHERE b.OCID1=@OCID1 AND a.IDNO=@IDNO AND a.SETID=@SETID" & vbCrLf
                    Dim sCmd3 As New SqlCommand(sql, tmpConn, tmpTrans)

                    Dim dt3 As New DataTable
                    With sCmd3
                        .Parameters.Clear()
                        .Parameters.Add("OCID1", SqlDbType.VarChar).Value = dr1("OCID1")
                        .Parameters.Add("IDNO", SqlDbType.VarChar).Value = dr1("IDNO")
                        .Parameters.Add("SETID", SqlDbType.Int).Value = iSETID
                        dt3.Load(.ExecuteReader())
                    End With
                    If dt3.Rows.Count > 0 Then
                        For Each dr3 As DataRow In dt3.Rows
                            '審核失敗刪除錄取資料。
                            Call DEL_STUDSELRESULT(dr3("SETID"), dr3("OCID1"), sm, tmpTrans)
                        Next
                    End If
                End If

                Dim vEmail As String = TIMS.ChangeEmail(TIMS.ClearSQM(dr1("Email")))
                If vEmail <> "" And sEmailSend And vsisEmailFail <> "Y" Then
                    'If dr1("Email").ToString <> "" And vsisEmailFail <> "Y" Then '測試用上一句正確
                    '修正錯誤的EMAIL增加發信成功率 BY AMU
                    dr1("Email") = vEmail
                    '20090601 by Jimmy add 依需求將「甄試日期」改為「說明事項」，新增 eComment 參數
                    Dim htSS As New Hashtable From {
                        {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
                        {"Stud_Name", vsStud_name},
                        {"Subject", vsSubject},
                        {"ExamNo", vsExamNo},
                        {"RelEnterDate", vsRelEnterDate},
                        {"ExamDate", vsExamDate},
                        {"CheckInDate", vsCheckInDate},
                        {"EComment", vsECommentEmpty},
                        {"Email", vEmail},
                        {"from_emailaddress", vpath3},
                        {"signUpMemo", vssignUpMemo},
                        {"sRIDOrgName", vsRID},
                        {"sType", TIMS.Cst_SendMail3_CheckedFalse}
                    } 'htSS Hashtable() 
                    Dim mail_msg As String = TIMS.SendMail3(htSS)
                    If mail_msg <> "" Then
                        Common.RespWrite(MyPage, "<script>alert('" & mail_msg & "');</script>")
                    Else
                        Dim pms_u As New Hashtable From {{"eSerNum", Val(Hid_eSerNum.Value)}}
                        Dim sqlstr_u As String = " UPDATE STUD_ENTERTYPE2 SET isEmailFail='Y' WHERE eSerNum=@eSerNum"
                        DbAccess.ExecuteNonQuery(sqlstr_u, tmpTrans, pms_u)
                    End If
                End If
                vsStud_name = Nothing
                vsSubject = Nothing
                vsExamNo = Nothing
                vsRelEnterDate = Nothing
                vsExamDate = Nothing
                vsCheckInDate = Nothing
                vssignUpMemo = Nothing '備註(失敗原因)
                vsRID = Nothing
            End If
        End If
    End Sub

    ''' <summary>匯入名冊產投用</summary>
    Sub Import_ENTERTYPE2()
        'Hid_impOCID.Value=""
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇 班級名稱!!")
            Exit Sub
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate2854(OCIDValue1.Value, objconn)
        Dim bflag1_OK As Boolean = If(drCC IsNot Nothing, True, False) '有輸入班級資料-檢核班級是否OK(範圍內)
        If Not bflag1_OK Then
            BtnImport28.Enabled = False
            Common.MessageBox(Me, "該班級不提供匯入!!")
            Exit Sub
        End If
        If Hid_impOCID.Value <> OCIDValue1.Value Then
            BtnImport28.Enabled = False
            Common.MessageBox(Me, "請先輸入班級，後按查詢才可確認匯入班級!")
            Exit Sub
        End If

        Dim Upload_Path As String = "~/SD/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Const cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, cst_Filetype) Then Return

        Dim iSerNum_MaxID As Integer = 0 ' String="" '暫時用序號
        Dim ieSETID As Integer = 0
        Dim ieSerNum As Integer = 0
        Dim OKFlag As Boolean = False '報名成功 True/ 失敗 False
        Dim SID As String = "" '學員編號

        If File1.Value = "" Then
            Common.MessageBox(Me, "請選擇匯入檔案!!")
            Exit Sub
        ElseIf File1.PostedFile.ContentLength = 0 Then
            '檢查檔案格式與大小
            Common.MessageBox(Me, "檔案位置錯誤!!")
            Exit Sub
        End If
        '取出檔案名稱
        Dim MyFileName As String = ""
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        Dim MyFileType As String = ""
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> LCase(cst_Filetype) Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(cst_Filetype) & "檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小

        '固定關鍵字
        Const cst_keyword1 As String = "身份證字號"

        Dim dt_xls As DataTable
        Dim Errmag As String = ""
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
        '上傳檔案
        File1.PostedFile.SaveAs(filePath1)
        '取得內容
        dt_xls = TIMS.GetDataTable_XlsFile(filePath1, "", Errmag, cst_keyword1)
        If dt_xls Is Nothing Then
            Errmag = "" '(再試一次)
            dt_xls = TIMS.GetDataTable_XlsFile(filePath1, "", Errmag, cst_keyword1) '取得內容
        End If
        '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        TIMS.MyFileDelete(filePath1)

        If CheckBox1.Checked Then '匯入名冊產投用(檢視)
            dtgAddresses1.Visible = True
            dtgAddresses1.DataSource = dt_xls   '匯入名冊產投用(檢視)
            dtgAddresses1.DataBind()
            Exit Sub
        End If

        If Errmag <> "" Then
            'Common.MessageBox(Me, Errmag)
            Dim s_Errmag1 As String = ""
            s_Errmag1 = "資料有誤，故無法匯入，請修正Excel檔案，謝謝" & vbCrLf
            s_Errmag1 &= Errmag & vbCrLf
            Common.MessageBox(Me, s_Errmag1)
            Exit Sub
        End If
        Dim RowIndex As Integer = 1 '讀取行累計數

        'Dim col As String          '欄位
        Dim colArray As Array
        '取出資料庫的所有欄位--------   Start
        'Dim sql As String
        ''建立Next_Dlid值 (新值，不會重複)
        'sql="Select distinct max(dlid)+1 Next_Dlid from stud_resultstuddata "
        'Dim Next_Dlid As String=DbAccess.ExecuteScalar(sql)
        'Dim BasicSID As String=TIMS.Get_DateNo
        'Dim SIDNum As Integer=1
        'Dim SID As String
        Dim Reason As String = ""           '儲存錯誤的原因
        Dim dtWrong As New DataTable        '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        '建立錯誤資料格式Table-Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Name"))
        dtWrong.Columns.Add(New DataColumn("IDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table-End

        '取出所有鍵值當判斷-Start
        Dim sql As String = " SELECT * FROM KEY_DEGREE WHERE DEGREETYPE IN ('0','1') "
        dt_Key_Degree = DbAccess.GetDataTable(sql, objconn)

        '依計畫
        'sql=""
        'sql &= " SELECT IdentityID, Name" & vbCrLf
        'sql &= " FROM Key_Identity" & vbCrLf
        'sql &= " WHERE 1=1" & vbCrLf
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '產投／充飛
        '    '新增02:非自願離職者身分 by Andy 20090210 
        '    '取消02:非自願離職者身分 by AMU 20100414
        '    '新增33:中低收入戶 by AMU 201106
        '    'sql="SELECT IdentityID, case when IdentityID='02' then '非自願離職者' else Name end as Name" & vbCrLf
        '    'sql &= " FROM Key_Identity WHERE 1=1" & vbCrLf
        '    'sql &= " AND IdentityID IN ('01','04','05','06','07','26','10','28','33')" & vbCrLf
        '    'sql &= " AND IdentityID IN (" & TIMS.cst_Identity28 & ")" & vbCrLf
        'End If
        ''改為依計畫顯示 by AMU 201106
        ''依參數設定
        'sql &= " AND IdentityID NOT IN (" & vbCrLf
        'sql &= "  SELECT IdentityID FROM Plan_Identity WHERE TPlanID='" & Convert.ToString(sm.UserInfo.TPlanID) & "'" & vbCrLf
        'sql &= "  AND isEnabled='N'" & vbCrLf '排除不可用
        'sql &= " )" & vbCrLf

        dt_Key_Identity = TIMS.Get_dtIdentity(9, objconn, sm)

        sql = " SELECT TradeID,concat('[',TradeID,']', TradeName) TradeName FROM KEY_TRADE"
        dt_Key_Trade = DbAccess.GetDataTable(sql, objconn)

        '28/54 TIMS28.OCID
        sql = ""
        sql &= " SELECT cc.OCID ,cc.ClassCName" & vbCrLf  ', dbo.TRUNC_DATETIME(getdate() - cc.STDate,0) as ddiff
        sql &= " ,ISNULL(iz.CTID,0) CTID" & vbCrLf
        sql &= " ,ISNULL(iz1.CTID,0) CTID1" & vbCrLf
        sql &= " ,ISNULL(iz2.CTID,0) CTID2" & vbCrLf
        sql &= " ,cc.TMID,cc.RID,cc.PlanID" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " JOIN PLAN_PLANINFO pp ON pp.planid=cc.planid AND pp.comidno=cc.comidno AND pp.seqno=cc.seqno" & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID=cc.PlanID" & vbCrLf
        sql &= " LEFT JOIN ID_ZIP iz ON iz.ZipCode=cc.TaddressZip" & vbCrLf '2011 已不使用(產投)
        '產投上課地址學科場地代碼
        sql &= " LEFT JOIN PLAN_TRAINPLACE sp ON sp.PTID=pp.AddressSciPTID" & vbCrLf
        sql &= " LEFT JOIN ID_ZIP iz1 ON iz1.zipCode=sp.ZipCode" & vbCrLf
        sql &= " LEFT JOIN ID_CITY ic1 ON ic1.CTID=iz1.CTID" & vbCrLf
        '產投上課地址術科場地代碼
        sql &= " LEFT JOIN PLAN_TRAINPLACE tp ON tp.PTID=pp.AddressTechPTID" & vbCrLf
        sql &= " LEFT JOIN ID_ZIP iz2 ON iz2.zipCode=tp.ZipCode" & vbCrLf
        sql &= " LEFT JOIN ID_CITY ic2 ON ic2.CTID=iz2.CTID" & vbCrLf
        sql &= " WHERE cc.NotOpen='N' AND cc.IsClosed <> 'Y'" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then sql &= " AND ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf '限定轄區
        sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf '限定登入年度
        sql &= " AND ip.TPlanID IN (" & TIMS.Cst_TPlanID28_2 & ")" & vbCrLf
        dtClassInfo = DbAccess.GetDataTable(sql, objconn)

        sql = "SELECT ZIPCODE, ZIPNAME FROM dbo.ID_ZIP ORDER BY 1"
        dtZipCode = DbAccess.GetDataTable(sql, objconn)
        '取出所有鍵值當判斷-End

        aNow = TIMS.GetSysDateNow(objconn)
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            Reason = ""
            colArray = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
            If Reason = "" Then Reason &= CheckImportData(colArray) '檢查資料正確性
            If Reason = "" Then
                Try
                    Dim Str As String = "" '存sql語法用
                    Dim da As SqlDataAdapter = Nothing
                    Dim dt As DataTable = Nothing
                    Dim dr As DataRow = Nothing
                    'Dim x As Integer=0
                    'Dim y As Integer=0
                    Dim iSETID As Integer = 0
                    '線上報名資料寫入Stud_EnterTemp2中---start
                    '20090511(Milor)身分證號相同才視為同一人，不再檢核生日。
                    Str = " SELECT * FROM STUD_ENTERTEMP2 WHERE IDNO='" & aIDNO & "'" & vbCrLf
                    dt = DbAccess.GetDataTable(Str, da, objconn)
                    If dt.Rows.Count = 0 Then
                        '未有此學員的線上報名資料.則新增一筆報名學員資料
                        ieSETID = TIMS.Get_eSETID_MaxID(aIDNO, objconn, Nothing)
                        Call INSERT_STUD_ENTERTEMP2(dt, ieSETID)
                    Else
                        'SETID=dt.Rows(0)("SETID") '任1個 SETID '多筆資料修正
                        Call UPDATE_STUD_ENTERTEMP2(dt, ieSETID, iSETID)
                    End If
                    DbAccess.UpdateDataTable(dt, da)

                    'flag_i=0 '線上報名資料寫入Stud_EnterType2中---start
                    Str = $" SELECT TOP 1 * FROM STUD_ENTERTYPE2 WHERE eSETID={ieSETID} AND OCID1={aOCID1}"
                    dt = DbAccess.GetDataTable(Str, da, objconn)
                    If dt.Rows.Count = 0 Then
                        ieSerNum = TIMS.GET_ESERNUM_MAXID(objconn, Nothing)
                        dr = dt.NewRow()
                        dt.Rows.Add(dr)
                        dr("eSerNum") = ieSerNum 'PK
                        dr("eSETID") = ieSETID
                    Else
                        dr = dt.Rows(0)
                        ieSerNum = dr("eSerNum")
                        'flag_i=1
                    End If
                    '取得新eSerNum@Stud_EnterType2 (e網報名)
                    iSerNum_MaxID = TIMS.GET_TYPE2_MAXSerNum1(ieSETID, objconn, Nothing)
                    dr("SETID") = If(iSETID > 0, iSETID, Convert.DBNull)
                    dr("EnterDate") = aNow.ToString("yyyy/MM/dd") 'Today()
                    dr("SerNum") = iSerNum_MaxID
                    dr("RelEnterDate") = aNow ' Now()
                    dr("OCID1") = aOCID1
                    'Dim drCC1 As DataRow=TIMS.GetOCIDDate(aOCID1, objconn)
                    Dim s_FINDT1 As String = String.Concat("OCID='", aOCID1, "'")
                    If (dtClassInfo.Select(s_FINDT1).Length > 0) Then
                        Dim drCF As DataRow = dtClassInfo.Select(s_FINDT1)(0)
                        dr("TMID1") = drCF("TMID") '暫時用
                        dr("RID") = drCF("RID") '暫時用
                        dr("PlanID") = drCF("PlanID") '暫時用
                    End If
                    dr("identityID") = aMIdentityID  'IdentityID.SelectedValue
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = aNow 'Now
                    DbAccess.UpdateDataTable(dt, da)

                    '---end
                    'flag_i=0
                    '線上報名資料寫入Stud_EnterTrain2中---start
                    Dim iSEID As Integer = -1
                    Str = $" SELECT * FROM STUD_ENTERTRAIN2 WHERE eSerNum={ieSerNum } "
                    dt = DbAccess.GetDataTable(Str, da, objconn)
                    If dt.Rows.Count = 0 Then
                        iSEID = DbAccess.GetNewId(objconn, "STUD_ENTERTRAIN2_SEID_SEQ,STUD_ENTERTRAIN2,SEID")
                        dr = dt.NewRow()
                        dt.Rows.Add(dr)
                        dr("SEID") = iSEID
                        dr("eSerNum") = ieSerNum
                    Else
                        dr = dt.Rows(0)
                        iSEID = dr("SEID")
                        'flag_i=1
                    End If
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = aNow 'Now
                    dr("ZipCode2") = If(aZipCode2 <> "", aZipCode2, Convert.DBNull)
                    dr("ZipCode2_6W") = If(aZipCode2_6W <> "", aZipCode2_6W, Convert.DBNull)
                    dr("HouseholdAddress") = aHouseholdAddress
                    dr("MidentityID") = aMIdentityID
                    dr("PriorWorkPay") = If(aPriorWorkPay <> "", aPriorWorkPay, Convert.DBNull)
                    dr("AcctMode") = TIMS.ClearSQM(aAcctMode)

                    '如果帳戶類別是2時，因為沒有任何的帳號，所以不需要存入帳號資訊
                    If aAcctMode = 0 Then
                        dr("PostNo") = TIMS.ClearSQM(aPostNo)
                        dr("AcctNo") = TIMS.ClearSQM(aAcctNo)
                    ElseIf aAcctMode = 1 Then
                        dr("BankName") = TIMS.ClearSQM(aBankName)
                        dr("AcctHeadNo") = TIMS.ClearSQM(aAcctHeadNo)
                        dr("ExBankName") = TIMS.ClearSQM(aExBankName)
                        dr("AcctExNo") = TIMS.ClearSQM(aAcctExNo)
                        dr("AcctNo") = TIMS.ClearSQM(aAcctNo2)
                    End If
                    aUname = TIMS.ClearSQM(aUname)
                    aIntaxno = TIMS.ClearSQM(aIntaxno)
                    aActNo = TIMS.ClearSQM(aActNo)
                    aActname = TIMS.ClearSQM(aActname)
                    dr("Uname") = If(aUname <> "", aUname, Convert.DBNull) '公司名稱
                    dr("Intaxno") = If(aIntaxno <> "", aIntaxno, Convert.DBNull) '服務單位統一編號
                    dr("ActNo") = If(aActNo <> "", aActNo, Convert.DBNull) '保險證號
                    dr("Actname") = If(aActname <> "", aActname, Convert.DBNull) '投保公司名稱
                    aActType = TIMS.ClearSQM(aActType)
                    If aActType.Length > 1 Then
                        aActType = If(Right(aActType, 1) = "0", "", Right(aActType, 1))
                    Else
                        aActType = If(aActType = "0", "", aActType)
                    End If
                    dr("ActType") = If(aActType <> "", aActType, Convert.DBNull)
                    '20090210 andy  edit --- end
                    aScale = TIMS.ClearSQM(aScale)
                    dr("Scale") = aScale

                    aJobTitle = TIMS.ClearSQM(aJobTitle)
                    dr("JobTitle") = If(aJobTitle <> "", aJobTitle, Convert.DBNull)
                    '=====任職公司其他資料地址=====
                    dr("Zip") = "-1" '97產業人才投資方案取消輸入
                    dr("Addr") = " " '97產業人才投資方案取消輸入
                    dr("Tel") = " " '97產業人才投資方案取消輸入
                    dr("ShowDetail") = "N" '97產業人才投資方案取消輸入
                    dr("Q1") = 0
                    dr("Q2_1") = If(aQ2 = "1", 1, 2)
                    dr("Q2_2") = If(aQ2 = "2", 1, 2)
                    dr("Q2_3") = If(aQ2 = "3", 1, 2)
                    dr("Q2_4") = If(aQ2 = "4", 1, 2)

                    dr("Q3") = If(aQ3 <> "", aQ3, Convert.DBNull)
                    dr("Q3_Other") = If(aQ3_Other <> "", aQ3_Other, dr("Q3_Other"))
                    dr("Q4") = aQ4
                    dr("Q5") = If(IsNumeric(aQ5), Convert.ToInt32(aQ5), 0)
                    dr("Q61") = If(aQ61 <> "", Convert.ToInt32(aQ61), Convert.DBNull)
                    dr("Q62") = If(aQ62 <> "", Convert.ToInt32(aQ62), Convert.DBNull)
                    dr("Q63") = If(aQ63 <> "", Convert.ToInt32(aQ63), Convert.DBNull)
                    dr("Q64") = If(aQ64 <> "", Convert.ToInt32(aQ64), Convert.DBNull)
                    '20090210 andy edit - end

                    aIseMail = TIMS.ClearSQM(aIseMail)
                    dr("IseMail") = If(aIseMail <> "", aIseMail, dr("IseMail"))
                    '投保單位電話、地址
                    dr("ActTel") = If(actTel <> "", actTel, Convert.DBNull)
                    dr("ZipCode3") = If(actZipCode <> "", actZipCode, Convert.DBNull)
                    dr("ZipCode3_6W") = If(actZipCODE6W <> "", actZipCODE6W, Convert.DBNull)
                    dr("ActAddress") = If(actAddress <> "", actAddress, Convert.DBNull)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = aNow 'Now
                    DbAccess.UpdateDataTable(dt, da)
                    'DbAccess.CommitTrans(trans) 'Common.RespWrite(Me, "<script>alert('報名成功');</script>")
                    OKFlag = True
                Catch ex As Exception
                    'DbAccess.RollbackTrans(trans) 'Common.RespWrite(Me, "<script>alert('報名資料上傳失敗');</script>")
                    OKFlag = False
                    Throw ex
                End Try
            Else
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                drWrong("Index") = RowIndex
                If colArray.Length > 5 Then
                    drWrong("Name") = TIMS.HtmlDecode1(aName)
                    drWrong("IDNO") = TIMS.ChangeIDNO(aIDNO)
                    drWrong("Reason") = Reason
                End If
            End If
            RowIndex += 1 '讀取行累計數
        Next
        '開始判別欄位存入------------   End

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        explain += "匯入資料共" & dt_xls.Rows.Count & "筆" & vbCrLf
        explain += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆" & vbCrLf
        explain += "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
        Dim explain2 As String = ""
        explain2 += "匯入資料共" & dt_xls.Rows.Count & "筆\n"
        explain2 += "成功：" & (dt_xls.Rows.Count - dtWrong.Rows.Count) & "筆\n"
        explain2 += "失敗：" & dtWrong.Rows.Count & "筆\n"
        '開始判別欄位存入------------   End
        If dtWrong.Rows.Count = 0 Then
            If Reason <> "" Then
                Common.MessageBox(Me, explain & Reason)
                Exit Sub
            End If
            Common.MessageBox(Me, explain)
            'Exit Sub
        Else
            Session("MyWrongTable") = dtWrong
            Dim x_script As String = String.Concat("<script>if(confirm('", explain2, "是否要檢視原因?')){window.open('SD_01_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
            Page.RegisterStartupScript("", x_script)
        End If
    End Sub

    ''' <summary>
    ''' 匯入名冊-checkSizeA
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnImport28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnImport28.Click
        Call Import_ENTERTYPE2() '匯入名冊產投用
    End Sub

    '列印匯入學員報名名冊用的班級代碼
    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = Convert.ToString(sm.UserInfo.RID)
        Years.Value = Convert.ToString(sm.UserInfo.Years)

        Dim myValue As String = ""
        myValue = "RID=" & RIDValue.Value
        myValue &= "&Years=" & Years.Value
        myValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        DataGridTable.Visible = False
        Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
    End Sub

    '測試信發送
    'Private Sub LinkButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton1.Click
    '    Call Utl_TestEmailSend_1()
    'End Sub

    '測試信發送
    Sub Utl_TestEmailSend_1()
        Dim path3 As String = TIMS.Utl_GetConfigSet("from_emailaddress")
        Dim vpath3 As String = If(String.IsNullOrEmpty(path3), TIMS.Cst_SendMail3_from_emailaddress, path3)
        '20090601 by Jimmy add 依需求將「甄試日期」改為「說明事項」，新增 eComment 參數
        Dim htSS As New Hashtable From {
            {"TPlanID", Convert.ToString(sm.UserInfo.TPlanID)},
            {"Stud_Name", "STUD_NAME"},
            {"Subject", "Subject-測試信"},
            {"ExamNo", "ExamNo"},
            {"RelEnterDate", "RelEnterDate"},
            {"ExamDate", "ExamDate"},
            {"CheckInDate", "CheckInDate"},
            {"EComment", "EComment"},
            {"Email", TIMS.Cst_EmailtoMe},
            {"from_emailaddress", vpath3},
            {"signUpMemo", ""},
            {"sRIDOrgName", ""},
            {"sType", TIMS.Cst_SendMail3_CheckedOK}
        } 'htSS Hashtable() 
        Dim mail_msg As String = TIMS.SendMail3(htSS)
        If mail_msg <> "" Then Common.RespWrite(Me, "<script>alert('" & mail_msg & "');</script>")
    End Sub

    ''' <summary>
    ''' '取得單1報名資料。 Stud_EnterTemp2,Stud_EnterType2,Stud_EnterSubData2
    ''' </summary>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    ''' <param name="eSerNum"></param>
    ''' <returns></returns>
    Public Shared Function GET_STUDENTERTYPE2(ByRef tmpConn As SqlConnection, ByRef tmpTrans As SqlTransaction, ByVal eSerNum As String) As DataRow
        eSerNum = TIMS.ClearSQM(eSerNum)
        Dim rst As DataRow = Nothing
        If eSerNum = "" Then Return rst
        Dim pms As New Hashtable From {{"eSerNum", Val(eSerNum)}}
        Dim sql As String = "" 'isEmailFail
        sql &= " SELECT b.RID ,b.OCID1 ,b.TMID1 ,b.OCID2 ,b.TMID2 ,b.OCID3 ,b.TMID3 ,b.EnterDate ,b.RelEnterDate,b.ENTERPATH"
        sql &= " ,b.IdentityID ,b.MIDENTITYID,b.PlanID ,b.CCLID ,b.eSerNum ,b.isEmailFail,b.CMASTER1"
        sql &= " ,se2.ActNo ,se2.PriorWorkType1 ,se2.PriorWorkOrg1 ,se2.SOfficeYM1 ,se2.FOfficeYM1 ,a.eSETID"
        sql &= " ,a.IDNO ,a.Name ,a.Sex ,a.Birthday ,a.PassPortNO ,a.MaritalStatus ,a.PassPortNO ,a.DegreeID ,a.GradID ,a.School"
        sql &= " ,a.Department ,a.MilitaryID ,a.ZipCode ,a.Address ,a.Phone1 ,a.Phone2 ,a.CellPhone ,a.Email ,a.ZipCODE6W"
        sql &= " FROM dbo.STUD_ENTERTEMP2 a"
        sql &= " JOIN dbo.STUD_ENTERTYPE2 b ON a.eSETID=b.eSETID"
        sql &= " LEFT JOIN dbo.STUD_ENTERSUBDATA2 se2 ON b.eSerNum=se2.eSerNum"
        sql &= " WHERE b.eSerNum =@eSerNum"
        rst = DbAccess.GetOneRow(sql, tmpTrans, pms)   'Stud_EnterTemp2
        Return rst
    End Function

    '查詢鈕  '匯出鈕 'hidSchBtnNum.value: 1.正常查詢 2.正常匯出
    Sub sUtl_btnSearchData1(ByVal sender As System.Object, ByVal e As System.EventArgs)
        Dim objLock As New Object
        SyncLock objLock
            SyncLock TIMS.objLock_SD01004
                Call SearchAll_1(sender, e) '查詢
            End SyncLock
        End SyncLock
    End Sub

    Sub SearchAll_1(sender As System.Object, e As System.EventArgs)
        Dim BtnObj As Button = CType(sender, Button)
        Dim sMsg As String = ""
        Select Case LCase(BtnObj.CommandName)
            Case cst_btn_button1 '查詢鈕
                Call Search1() '查詢

            Case cst_btn_button13 '匯出鈕
                Call Export1() '匯出

            Case cst_btn_btndivPwdSubmit
                '正常顯示 '查詢或匯出。輸入個資安全密碼()
                If Not TIMS.sUtl_ChkPlanPwd(sm.UserInfo.PlanID, objconn) Then
                    sMsg = "未設定計畫密碼!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                If Not TIMS.sUtl_ChkPlanPwdOK(objconn, sm.UserInfo.PlanID, txtdivPxssward.Text) Then
                    sMsg = "個資安全密碼錯誤!!"
                    labChkMsg.Text = sMsg
                    Common.MessageBox(Me, sMsg)
                    Exit Sub
                End If
                'If rblWorkMode.SelectedValue="2" Then flgCIShow=True '可正常顯示個資。
                txtdivPxssward.Text = ""
                Select Case hidSchBtnNum.Value
                    Case "1"
                        Call Search1()
                    Case "2"
                        Call Export1()
                End Select
        End Select
    End Sub

    Sub GSet_BtnHistory(ByRef BtnHistory As Button, ByRef drv As DataRowView, ByRef iTwoYears As Integer)
        Const cst_wo_name1 As String = "history"
        Const cst_wo_specs1 As String = "width=1400,height=820,scrollbars=1"

        Dim wo_url1 As String = TIMS.Get_Url1(Me, "../05/SD_05_010_pop.aspx")  '20180921
        wo_url1 &= String.Concat("&SD_01_004_Type=Student&IDNO=", drv("IDNO"), "&TwoYears=", iTwoYears, "&BtnHistory=", Button1.ClientID)

        Dim win_open_script As String = String.Concat("window.open(", "'", wo_url1, "','", cst_wo_name1, "','", cst_wo_specs1, "');return false;")
        BtnHistory.Attributes("onclick") = win_open_script
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

#Region "NO USE"
    '匯入名冊 是否開放檢核鈕
    'Private Sub LinkButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LinkButton2.Click
    '    '090914 當班級代碼欄位有變更時檢查登入帳號是否有授權
    '    '※目前匯入名冊功能平時不開放，只提供給已結訓班級使用授權檔內有指定之授權帳號使用()
    '    '請勿mark掉此段程式及 javascript function  chkOCID(){ __doPostBack('LinkButton2','');}	
    '    'chkAcctRight()
    '    '檢查Session是否存在 Start
    '    ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
    '    '檢查Session是否存在 End
    '    BtnImport28.Enabled=False '匯入名冊
    '    If chkAcctRight(sm.UserInfo.UserID, OCIDValue1.Value) Then BtnImport28.Enabled=True
    'End Sub

    '將Temp2的資料更新到TEMP
    'Private Sub Update_StudEnterTemp(ByVal IDNO2 As String, ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction)
    '    Dim sqlAdp As New SqlDataAdapter
    '    Try
    '        Dim sqlStr As String
    '        Dim BKID As Integer=0
    '        '將Temp2的資料更新到TEMP
    '        sqlStr="update Stud_EnterTemp set Stud_EnterTemp.Name=Stud_EnterTemp2.Name,Stud_EnterTemp.Sex=Stud_EnterTemp2.Sex" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.Birthday=Stud_EnterTemp2.Birthday,Stud_EnterTemp.PassPortNO=Stud_EnterTemp2.PassPortNO" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.MaritalStatus=Stud_EnterTemp2.MaritalStatus,Stud_EnterTemp.DegreeID=Stud_EnterTemp2.DegreeID" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.GradID=Stud_EnterTemp2.GradID,Stud_EnterTemp.School=Stud_EnterTemp2.School" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.Department=Stud_EnterTemp2.Department,Stud_EnterTemp.MilitaryID=Stud_EnterTemp2.MilitaryID" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.ZipCode=Stud_EnterTemp2.ZipCode,Stud_EnterTemp.Address=Stud_EnterTemp2.Address" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.Phone1=Stud_EnterTemp2.Phone1,Stud_EnterTemp.Phone2=Stud_EnterTemp2.Phone2" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.CellPhone=Stud_EnterTemp2.CellPhone,Stud_EnterTemp.Email=Stud_EnterTemp2.Email" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.IsAgree=Stud_EnterTemp2.IsAgree,Stud_EnterTemp.ZipCODE6W=Stud_EnterTemp2.ZipCODE6W" & vbCrLf
    '        sqlStr += ",Stud_EnterTemp.ModifyAcct=@ModifyAcct,Stud_EnterTemp.ModifyDate=getdate()" & vbCrLf
    '        sqlStr += "from Stud_EnterTemp join Stud_EnterTemp2 on Stud_EnterTemp.SETID=Stud_EnterTemp2.SETID" & vbCrLf
    '        sqlStr += "where upper(Stud_EnterTemp.IDNO)=@IDNO2 "
    '        With sqlAdp
    '            .UpdateCommand=New SqlCommand(sqlStr, tmpConn, tmpTrans)
    '            .UpdateCommand.Parameters.Clear()
    '            .UpdateCommand.Parameters.Add("@ModifyAcct", SqlDbType.VarChar).Value=Convert.ToString(sm.UserInfo.UserID)
    '            .UpdateCommand.Parameters.Add("@IDNO2", SqlDbType.VarChar).Value=TIMS.ChangeIDNO(IDNO2)
    '            .UpdateCommand.ExecuteNonQuery()
    '        End With
    '    Catch ex As Exception
    '        tmpTrans.Rollback()
    '        DbAccess.CloseDbConn(tmpConn)
    '        sqlAdp.Dispose()
    '        Throw ex
    '        Common.MessageBox(Me, ex.ToString)
    '    End Try
    '    sqlAdp.Dispose()
    'End Sub

    '查詢
    'Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    '    'sUtl_btnSearchData1
    'End Sub

    '查詢參訓歷史
    'Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
    '    Page.RegisterStartupScript("History", "<script>window.open('../05/SD_05_010.aspx?SD_01_004_Type=StudentList&BtnHistory=" & Button1.ClientID & "' ,'history','width=700,height=500,scrollbars=1')</script>")
    'End Sub

    ''判斷帳號是否有權限使用補登功能 'True:有權限使用'False:沒有權限
    'Function chkAcctRight(ByVal SessionUserID As String, ByVal vOCID As String) As Boolean
    '    Dim Rst As Boolean=False
    '    'BtnImport28.Enabled=False'OCIDValue1.Value'sm.UserInfo.UserID
    '    If vOCID <> "" AndAlso SessionUserID <> "" Then
    '        If ChkIsEndDate(SessionUserID, vOCID)=False Then    '判斷帳號是否已到期 'True:已到期 False:未到期
    '            Rst=True
    '            'BtnImport28.Enabled=True
    '        End If
    '    End If
    '    Return Rst
    'End Function

    ''判斷帳號是否已到期  '20090226 andy add 檢查 SYS_03_006 已結訓班級使用授權
    'Function ChkIsEndDate(ByVal UserID As String, ByVal OCID As String) As Boolean
    '    'Dim FunIDstr, sql As String
    '    'Dim dr As DataRow
    '    Dim Rst As Boolean=True 'True:已到期 False:未到期
    '    Dim dr As DataRow=Nothing
    '    Dim FunIDstr As String=""
    '    Dim sql As String=""
    '    sql=" SELECT FunID FROM Auth_REndClass where 0=0 and OCID=" & OCID & " and  Account='" & UserID & "'"
    '    Try
    '        FunIDstr=Convert.ToString(DbAccess.ExecuteScalar(sql, objconn))
    '    Catch ex As Exception
    '        FunIDstr=""
    '    End Try
    '    Try
    '        If FunIDstr <> "" Then
    '            sql=""
    '            sql &= " SELECT FunID from ID_Function "
    '            sql &= " WHERE FunID in ( " & FunIDstr & " ) "
    '            sql &= " and FunID ='262'"
    '            dr=DbAccess.GetOneRow(sql, objconn)
    '        End If
    '        If Not dr Is Nothing Then
    '            sql="" & vbCrLf
    '            sql &= " select *" & vbCrLf
    '            sql &= " from (" & vbCrLf
    '            sql &= " select Account,OCID,UseAble,EndDate,FunID" & vbCrLf
    '            sql &= " ,case when EndDate>=getdate() and UseAble='Y' then 'N'" & vbCrLf
    '            sql &= "   when UseAble='N' then 'Y' else 'Y' end IsEndDate" & vbCrLf
    '            sql &= " from Auth_REndClass" & vbCrLf
    '            sql &= " where 0=0" & vbCrLf
    '            sql &= " and OCID=" & OCID & "" & vbCrLf
    '            sql &= " and Account='" & UserID & "'" & vbCrLf
    '            sql &= " ) aa" & vbCrLf
    '            sql &= " where IsEndDate='N'" & vbCrLf
    '            dr=DbAccess.GetOneRow(sql, objconn)
    '            If Not dr Is Nothing Then Rst=False
    '        End If
    '    Catch ex As Exception
    '        Rst=True
    '    End Try
    '    Return Rst
    'End Function

    'Call Search1()  / sUtl_btnSearchData1 '查詢
    'Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
    'End Sub

    ' sUtl_btnSearchData1 '匯出
    'Protected Sub Button13_Click(sender As Object, e As EventArgs) Handles Button13.Click
    'End Sub

    'Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged
    'End Sub

    'window.open('../05/SD_05_010.aspx?SD_01_004_Type=StudentList&BtnHistory=' + Button1ClientID + '', 'history', 'width=900,height=700,scrollbars=1')
    'Protected Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
    'Button9.Attributes.Add("onclick", "return open_StudentList('" & Button1.ClientID & "');")
    'End Sub
#End Region

End Class
