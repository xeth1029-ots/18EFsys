Partial Class SD_01_001
    Inherits AuthBasePage

    'Stud_EnterTypeDelData 'Stud_EnterType 'Stud_DelEnterType
    'Const cst_SearchSqlStr As String="SearchSqlStr" 'Me.ViewState(cst_SearchSqlStr)=sql  '存sql語法。 

    'print: SD_01_001
    Const cst_printFN1 As String = "SD_01_001_1"
    Const CST_KD_STUDENTLIST As String = "StudentList" 'Session("IDNOArray")
    Const cst_SD01001_addaspx As String = "SD_01_001_add.aspx"
    'Const cst_SD_16_addaspx As String="SD_01_001_add_16.aspx"
    'Dim str_SD_addaspx As String=""

    'SD_01_001_add.aspx SD_01_001_3in1.aspx
    '判斷功能ID  -- SELECT * FROM ID_FUNCTION WHERE FUNID IN (701,70,764)
    Const cst_funid報名登錄 As String = "70"               'EnterPath2@N
    Const cst_funid專案核定報名登錄 As String = "701"        'EnterPath2@P
    Const cst_funid特例專案核定報名登錄 As String = "764"     'EnterPath2@S '764 特例專案核定報名登錄

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Const cst_msg1 As String = "該民眾為就保非自願離職者，請通知民眾於該訓練班次報名截止日前，先至公立就業服務機構辦理求職登記，並經適訓評估後，推介參訓。"
    Const cst_xMaster2 As String = "該民眾 具公司/商業負責人身分，非屬失業勞工，不得報名失業者職前訓練。"
    'Const cst_xMaster3 As String="該民眾 具公司/商業負責人身分，非屬失業勞工，認定為在職者。"
    Const cst_msgERR1 As String = "查詢該民眾 為就保非自願離職者 連線有誤，請重新查詢!!"
    Const cst_msgERR2 As String = "查詢該民眾 具公司/商業負責人身分 連線有誤，請重新查詢!!"
    Const cst_ExportxMsg As String = "請遵照「個人資料保護法」相關法令規定，確實謹慎使用及保管本資料！"
    '請先確認民眾是否非公司／商業負責人及非就保非自願離職者，須符合參訓資格才能繼續報名。
    '另請提醒民眾於開訓日日前須符合參訓資格。

    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。

    'Dim sUrl As String=""                  '暫存。
    'Dim cmdArg As String=""                 '暫存。
    Dim ff As String = ""                     '過濾字
    'Dim dtGetjobc1 As DataTable=Nothing    '(就業情形)
    Dim blnExport As Boolean = False
    'Dim FunDr As DataRow=Nothing
    'Dim Auth_Relship As DataTable=Nothing
    Dim IDNOArray As New ArrayList            '身分證號碼陣列
    'Dim s_log1 As String=""

    'Dim aSETID As String            '流水ID--int
    Dim aIDNO As String              '身分證號碼 --varchar(15)
    Dim aName As String              '姓名--nvarchar(30)
    Dim aSex As String               '性別--男@M 女@F
    Dim aBirthday As String          '出生日期--yyyy/MM/dd
    Dim aPassPortNO As String        '身分別--1:本國2:外籍
    Dim aMaritalStatus As String     '婚姻狀況--1.已;2.未.3.暫不提供
    Dim aDegreeID As String          '學歷代碼	varchar(3)
    Dim aGradID As String            '畢業狀況代碼	varchar(3)
    Dim aSchool As String            '學校名稱 	nvarchar(30)
    Dim aDepartment As String        '科系名稱 	nvarchar(128)
    Dim aMilitaryID As String = ""   '兵役代碼	varchar(3)
    Dim aZipCode As String           '郵遞區號前3碼 	int
    Dim aZipCODE6W As String         '郵遞區號6碼 	int
    Dim aAddress As String           '通訊地址 	nvarchar(200)
    '*戶籍地址
    Dim aZIPCODE2 As String           '戶籍郵遞區號前3碼 	int
    Dim aZIPCODE2_6W As String        '戶籍郵遞區號6碼 	int
    Dim aHOUSEHOLDADDRESS As String   '戶籍地址 	nvarchar(200)

    Dim aEmail As String             'Email 	varchar(30) null(最好是可以填若沒有填'無')
    Dim aPhone1 As String            '聯絡電話(日)	varchar(25)
    Dim aPhone2 As String            '聯絡電話(夜)	varchar(25) null
    Dim aCellPhone As String         '行動電話 	varchar(25) null
    Dim aIsAgree As String           '同意否	char(1)		null (Y/N)

    ' TABLE 107、 Stud_EnterType
    Dim aMIdentityID As String      '主要參訓身分別代碼
    Dim aIdentityID1 As String      '參訓身分別代碼1
    Dim aIdentityID2 As String      '參訓身分別代碼2
    Dim aIdentityID3 As String      '參訓身分別代碼3
    Dim aIdentityID4 As String      '參訓身分別代碼3
    Dim aIdentityID5 As String      '參訓身分別代碼3

    Dim aEnterDate As String     '輸入日期(報名日期)yyyy/MM/dd PK (NOW)
    'Dim aSerNum As String       '序號 int PK (SYS)

    Dim aOCID1 As String        '依匯入資料取得'報考班別代碼1	int
    Dim aNewExamNO As String    '依匯入資料取得'(傳輸用) '准考證號	varchar(9) (SYS)
    Dim aTMID1 As String        '依匯入資料取得'(傳輸用)'報考職類ID1	int

    Dim aEnterChannel As String '報名管道 	int  1.網;2.現;3.通;4.推
    'IdentityID	參訓身分別代碼	varchar(50)	用半型, 分開
    'Dim aRID As String         'RID 	varchar(10) (SESSION)
    'Dim aPlanID As String      '計畫代碼 int       (SESSION)
    Dim aRID As String          '依匯入資料取得RID
    Dim aPlanID As String       '依匯入資料取得PlanID
    'Dim aTRNDMode As String    '推介種類	int		NULL	1.職2.學3.推
    'Dim aTRNDType As String    '職訓卷種類	int		NULL	1.甲式2.乙式
    'Dim aTicket_NO As String   '職訓券編號	nvarchar (20)		V	編碼規則 [登記編號(18)]-[流水號(2)]
    'Dim aNotExam As String     '是否免試 bit		Default ‘0’	0@NO;1@YES
    'Dim aNotExamID As String   '是否免試 bit		Default ‘0’	0@NO;1@YES

    Dim aUNAME As String '_服務單位
    Dim aINTAXNO As String '統一編號
    Dim aSERVDEPT As String '服務部門
    Dim aSERVDEPTID As String '服務部門ID

    Dim aACTNAME As String '投保單位名稱 
    Dim aACTTYPE As String '投保類別
    Dim aACTNO As String '投保單位保險證號
    Dim aACTTEL As String '投保單位保險證號
    Dim aZIPCODE3 As String    '戶籍郵遞區號前3碼 	int
    Dim aZIPCODE3_6W As String '戶籍郵遞區號6碼 	int
    Dim aACTADDRESS As String  '戶籍地址 	nvarchar(200)

    Dim aJOBTITLE As String '職稱/職務
    Dim aJOBTITLEID As String '職稱/職務ID

    Dim aAVTCP1_all As String = ""
    Dim aAVTCP1_01 As String   '獲得課程管道_01_本署或分署網站
    Dim aAVTCP1_02 As String   '獲得課程管道_01_本署或分署網站 As Integer 
    Dim aAVTCP1_03 As String   '獲得課程管道_02_就業服務中心 As Integer=
    Dim aAVTCP1_04 As String   '獲得課程管道_03_訓練單位 As Integer=44
    Dim aAVTCP1_05 As String   '獲得課程管道_04_搜尋網站 As Integer=45
    Dim aAVTCP1_06 As String   '獲得課程管道_05_報紙 As Integer=46
    Dim aAVTCP1_07 As String   '獲得課程管道_06_廣播 As Integer=47
    Dim aAVTCP1_08 As String   '獲得課程管道_07_電視 As Integer=48
    Dim aAVTCP1_09 As String   '獲得課程管道_08_朋友介紹 As Integer=49
    Dim aAVTCP1_99 As String   '獲得課程管道_09_社群媒體 As Integer=50

    Dim aNotes As String       '備註	nvarchar(500) '(x)ntext

    Const cst_i_c_身分證號碼 As Integer = 0
    Const cst_i_c_姓名 As Integer = 1
    Const cst_i_c_性別 As Integer = 2
    Const cst_i_c_出生日期 As Integer = 3
    Const cst_i_c_身份別 As Integer = 4

    Const cst_i_c_婚姻狀況 As Integer = 5
    Const cst_i_c_學歷代碼 As Integer = 6
    Const cst_i_c_畢業狀況代碼 As Integer = 7
    Const cst_i_c_學校名稱 As Integer = 8
    Const cst_i_c_科系名稱 As Integer = 9
    Const cst_i_c_兵役代碼 As Integer = 10

    Const cst_i_c_通訊郵遞區號前3碼 As Integer = 11
    Const cst_i_c_通訊郵遞區號5或6碼 As Integer = 12
    Const cst_i_c_通訊地址 As Integer = 13
    Const cst_i_c_戶籍郵遞區號前3碼 As Integer = 14
    Const cst_i_c_戶籍郵遞區號5或6碼 As Integer = 15
    Const cst_i_c_戶籍地址 As Integer = 16

    Const cst_i_c_Email As Integer = 17
    Const cst_i_c_聯絡電話_日 As Integer = 18
    Const cst_i_c_聯絡電話_夜 As Integer = 19
    Const cst_i_c_行動電話 As Integer = 20

    Const cst_i_c_主要參訓身份別代碼 As Integer = 21
    Const cst_i_c_參訓身份別代碼1 As Integer = 22
    Const cst_i_c_參訓身份別代碼2 As Integer = 23
    Const cst_i_c_參訓身份別代碼3 As Integer = 24
    Const cst_i_c_參訓身份別代碼4 As Integer = 25
    Const cst_i_c_參訓身份別代碼5 As Integer = 26

    Const cst_i_c_報名日期 As Integer = 27
    Const cst_i_c_報考班別代碼1 As Integer = 28    'Const cst_i_c_報考班別代碼2 As Integer=29    'Const cst_i_c_報考班別代碼3 As Integer=30
    Const cst_i_c_報名管道 As Integer = 29
    Const cst_i_c_同意共開資料 As Integer = 30

    Const cst_i_c_服務單位 As Integer = 31 'SERVDEPT
    Const cst_i_c_統一編號 As Integer = 32 'ACTNO
    Const cst_i_c_服務部門 As Integer = 33 'SERVDEPT
    Const cst_i_c_投保單位名稱 As Integer = 34 'SERVDEPT
    Const cst_i_c_投保類別 As Integer = 35 ''(1:勞保/2:農保)
    Const cst_i_c_投保單位保險證號 As Integer = 36
    Const cst_i_c_投保單位電話 As Integer = 37
    Const cst_i_c_投保單位郵遞區號前3碼 As Integer = 38
    Const cst_i_c_投保單位郵遞區號5或6碼 As Integer = 39
    Const cst_i_c_投保單位地址 As Integer = 40

    Const cst_i_c_職稱 As Integer = 41 ''(01:基層員工,02:基層管理者,03:中階管理者,04:高階管理者,05:負責人,99:其他)
    Const cst_i_c_獲得課程管道_01_本署或分署網站 As Integer = 42
    Const cst_i_c_獲得課程管道_02_就業服務中心 As Integer = 43
    Const cst_i_c_獲得課程管道_03_訓練單位 As Integer = 44
    Const cst_i_c_獲得課程管道_04_搜尋網站 As Integer = 45
    Const cst_i_c_獲得課程管道_05_報紙 As Integer = 46
    Const cst_i_c_獲得課程管道_06_廣播 As Integer = 47
    Const cst_i_c_獲得課程管道_07_電視 As Integer = 48
    Const cst_i_c_獲得課程管道_08_朋友介紹 As Integer = 49
    Const cst_i_c_獲得課程管道_09_社群媒體 As Integer = 50

    Const cst_i_c_獲得課程管道_99_其他 As Integer = 51
    Const cst_i_c_備註 As Integer = 52
    Const cst_Max_a_Len As Integer = 53

    Dim dtSERVDEPT As DataTable = Nothing
    Dim dtJOBTITLE As DataTable = Nothing
    Dim Key_Degree As DataTable
    Dim Key_GradState As DataTable
    Dim Key_Military As DataTable
    Dim Key_Identity As DataTable
    Dim dt_CLASS_CLASSINFO As DataTable
    Dim ID_ZipCode As DataTable
    'Const CST_身分證號=3

    Const cst_chk1 As Integer = 0
    Const cst_編號 As Integer = 1
    Const cst_姓名 As Integer = 2
    Const cst_身分證號碼 As Integer = 3
    Const cst_報名機構 As Integer = 4
    Const cst_報名班級 As Integer = 5
    Const cst_開訓日期 As Integer = 6
    Const cst_結訓日期 As Integer = 7
    Const cst_報名日期 As Integer = 8
    Const cst_准考證號碼 As Integer = 9
    Const cst_報名管道 As Integer = 10
    Const cst_是否試算 As Integer = 11
    Const cst_錄取結果 As Integer = 12
    'Const cst_協助基金 As Integer=13
    Const cst_結訓情形 As Integer = 13
    'Const cst_就業情形 As Integer=15
    Const cst_功能 As Integer = 14

    Const cst_EnterPathW As String = "W"                          '就服站代碼
    Const cst_EnterPathR As String = "R"                          '就服站代碼
    Const cst_EnterPathNameW As String = "<br />(就服單位協助報名)"  '說明
    'Const cst_EnterPathNameR As String="<br />(就服單位協助報名)" '說明
    Const cst_rw2不區分 As String = "A"
    Const cst_rw2一般推介單 As String = "CH4"
    Const cst_rw2免試推介單 As String = "EPW"
    Const cst_rw2專案核定報名 As String = "EP2P"                    'EP2PY

    Const cst_msg219 As String = "※ 姓名前標記「x-」表示民眾已註銷推介"
    Const cst_fgb219 As String = "x-"
    Const cst_Mgc219 As String = "民眾已註銷推介"

    Dim oflag_Test As Boolean = False  '測試
    Dim blnP0 As Boolean = False       '報名管道(職前計畫顯示)

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt)  '2011取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid2
        '分頁設定 End
        '判斷是否顯示准考證號碼
        PrintShow.Value = TIMS.GetGlobalVar(Me, "21", "1", objconn)
        blnExport = False '匯出為false;

        trImport1.Visible = False
        '接受企業委託訓練07、產學訓攜手合作計畫45
        '產學訓攜手合作計畫 接受企業委託訓練
        'SELECT * FROM VIEW_PLAN WHERE TPLANID IN ('07','45') AND YEARS ='2015'
        'Const Cst_TPlanID07AppPlan As String="07,45"
        'If sm.UserInfo.Years >= 2015 AndAlso Cst_TPlanID07AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    trImport1.Visible=True
        'End If
        trImport1.Visible = False
        If sm.UserInfo.Years >= 2015 AndAlso TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trImport1.Visible = True
        End If
        If sm.UserInfo.Years >= 2015 AndAlso TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso DateDiff(DateInterval.Day, Now, CDate(TIMS.cst_TPlanID70_end_date_1)) >= 0 Then
            trImport1.Visible = True
        End If

        blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        If blnP0 Then Trwork2013a.Visible = True

        '就服單位協助報名
        'Trwork2013a.Visible=False
        'If sm.UserInfo.Years >= 2013 AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013")="Y" Then Trwork2013a.Visible=True
        'End If

        'Dim oflag_Test As Boolean=False '測試
        If TIMS.sUtl_ChkTest() Then oflag_Test = True '測試
        '非 ROLEID=0 LID=0
        'Dim flgROLEIDx0xLIDx0 As Boolean=False '判斷登入者的權限。
        '如果是系統管理者開啟功能。'判斷登入者的權限。'ROLEID=0 LID=0
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1) 'False

        'btnIMPORT07_Click 'Utl_IMPORT07()
        'Const cst_Imp_v19 As String="../../Doc/Stud_Temp_v19.zip"
        Const cst_Stud_Temp_v22 As String = "../../Doc/Stud_Temp_v22.zip"
        HyperLink1.NavigateUrl = cst_Stud_Temp_v22 'cst_Imp_v19

        If Not IsPostBack Then
            If TIMS.StopEnterTempMsg_ANM2(Me, objconn, True) Then Exit Sub
            msg.Text = ""
            add_but.Attributes("onclick") = "javascript:return chkblank(1);"  '新增鈕
            Button2.Attributes("onclick") = "javascript:return chkblank(2);"  '查詢鈕
            table4.Visible = False
            table5.Visible = False
            end_date.Text = Common.FormatDate(Now.Date)
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Hid_PreUseLimited18a.Value = ""
            If TIMS.Cst_TPlanID_PreUseLimited18a.IndexOf(sm.UserInfo.TPlanID) > -1 Then Hid_PreUseLimited18a.Value = "Y"
        End If

        Dim s_LevOrg_aspx As String = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "LevOrg.aspx", "LevOrg1.aspx")
        Button8.Attributes("onclick") = String.Format("javascript:openOrg('../../Common/{0}');", s_LevOrg_aspx)

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '依sm.UserInfo.PlanID取得PlanKind
        Dim PlanKind As String = TIMS.Get_PlanKind(Me, objconn)
        If PlanKind = "1" Then
            Button5.Attributes("onclick") = "choose_class(2);"
        Else
            Button5.Attributes("onclick") = "choose_class(1);"
        End If

        '列印報名表(ReportQuery)
        'Button7.Attributes("onclick") +=     ReportQuery.ReportScript(Me, "list", "SD_01_001", "SETID='+doucment.getElementById('SETID').value +'&ExamNo='+doucment.getElementById('ExamNo').value +'&SerNum='+ doucment.getElementById('SerNum').value +'")
        'Me.Button7.Attributes.Add("OnClick", "return CheckPrint('" & ReportQuery.GetSmartQueryPath & "','" & sm.UserInfo.UserID & "');")

        If Not IsPostBack Then
            'Session("xx_DataTable")=Nothing
            If Session("_SearchStr") IsNot Nothing Then
                Dim str_SearchStr_x1 As String = Convert.ToString(Session("_SearchStr"))
                Session("_SearchStr") = Nothing

                Dim MyValue As String = ""
                center.Text = TIMS.GetMyValue(str_SearchStr_x1, "center")
                RIDValue.Value = TIMS.GetMyValue(str_SearchStr_x1, "RIDValue")
                OCID1.Text = TIMS.GetMyValue(str_SearchStr_x1, "OCID1")
                TMID1.Text = TIMS.GetMyValue(str_SearchStr_x1, "TMID1")
                OCIDValue1.Value = TIMS.GetMyValue(str_SearchStr_x1, "OCIDValue1")
                TMIDValue1.Value = TIMS.GetMyValue(str_SearchStr_x1, "TMIDValue1")
                IDNO.Text = TIMS.ChangeIDNO(TIMS.GetMyValue(str_SearchStr_x1, "IDNO"))
                start_date.Text = TIMS.GetMyValue(str_SearchStr_x1, "start_date")
                end_date.Text = TIMS.GetMyValue(str_SearchStr_x1, "end_date")
                Me.ViewState("PageIndex") = TIMS.GetMyValue(str_SearchStr_x1, "PageIndex")
                MyValue = TIMS.GetMyValue(str_SearchStr_x1, "submit")
                If MyValue = "1" Then
                    Call Search1()  '查詢鈕
                    If IsNumeric(Me.ViewState("PageIndex")) Then
                        PageControler1.PageIndex = Me.ViewState("PageIndex")
                        PageControler1.CreateData()
                    End If
                End If
            End If

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Call Search1()  '查詢鈕
                'Button12_Click(sender, e)
            End If

            btnExport1.Attributes("onclick") = "return confirm('" & cst_ExportxMsg & "');"
        End If
        If Session("xx_DataTable") IsNot Nothing Then
            Session("_DataTable") = Session("xx_DataTable")
            Session("xx_DataTable") = Nothing
        End If

        If Button3.Visible Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Me.Button3.Visible = False  '產投不顯示
        End If

        '確認機構是否為黑名單
        Dim vsMsg2 As String = "" '確認機構是否為黑名單
        vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            add_but.Enabled = False
            TIMS.Tooltip(add_but, vsMsg2)
            Button2.Enabled = False
            TIMS.Tooltip(Button2, vsMsg2)
            Button3.Enabled = False
            TIMS.Tooltip(Button3, vsMsg2)
            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
    End Sub

    ''' <summary> 機構黑名單內容(訓練單位處分功能) </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = sm.UserInfo.OrgName & "，已列入處分名單!!"
            Me.isBlack.Value = "Y"
            Me.Blackorgname.Value = sm.UserInfo.OrgName
        End If
        Return rst
    End Function

    ''' <summary>學員資料2更新／新增--START </summary>
    ''' <param name="trans"></param>
    Sub UPDATE_STUD_ENTERTEMP2(ByRef trans As SqlTransaction, ByRef iSETID As Integer)
        '學員資料2更新／新增--START 
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        'trans=DbAccess.BeginTrans(conn)
        Dim sql As String = " SELECT * FROM STUD_ENTERTEMP2 WHERE IDNO='" & TIMS.ChangeIDNO(aIDNO) & "'"
        dt = DbAccess.GetDataTable(sql, da, trans)
        '(查無資料離開)
        If dt.Rows.Count = 0 Then Return

        'dr("ESETID")=iESETID
        If dt.Rows.Count <> 0 Then
            aPhone1 = TIMS.ChangeIDNO(aPhone1)
            aPhone2 = TIMS.ChangeIDNO(aPhone2)
            aCellPhone = TIMS.ChangeIDNO(aCellPhone)

            For z As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(z)
                'dr("SETID")=dt.Rows(z).Item(1) 'dr("IDNO")=TIMS.ChangeIDNO(aIDNO)
                If (Convert.ToString(dr("SETID")) = "") Then dr("SETID") = iSETID
                If aName <> "" Then aName = Trim(aName) 'aName=TIMS.ClearSQM(aName)
                dr("Name") = aName
                dr("Sex") = aSex
                dr("Birthday") = aBirthday
                dr("PassPortNO") = aPassPortNO
                If Not (New List(Of String) From {"1", "2", "3"}).Contains(aMaritalStatus) Then aMaritalStatus = "" 'Convert.DBNull /MaritalStatus
                dr("MaritalStatus") = If(aMaritalStatus <> "", aMaritalStatus, Convert.DBNull)
                dr("DegreeID") = If(aDegreeID.Length < 2, "0" & aDegreeID, aDegreeID)
                dr("GradID") = If(aGradID.Length < 2, "0" & aGradID, aGradID)
                dr("School") = aSchool
                dr("Department") = aDepartment
                If aMilitaryID.ToString <> "" Then
                    If Convert.ToString(aMilitaryID) <> "" AndAlso aMilitaryID.Length < 2 Then aMilitaryID = String.Concat("0", aMilitaryID)
                    ff = "MilitaryID='" & aMilitaryID & "'"
                    If Key_Military.Select(ff).Length = 0 Then aMilitaryID = ""
                End If
                dr("MilitaryID") = If(aMilitaryID <> "", aMilitaryID, Convert.DBNull)
                dr("ZipCode") = aZipCode
                dr("ZipCODE6W") = aZipCODE6W
                dr("Address") = aAddress
                dr("Phone1") = If(aPhone1 <> "", aPhone1, "")
                dr("Phone2") = If(aPhone2 <> "", aPhone2, "")
                dr("CellPhone") = If(aCellPhone <> "", aCellPhone, "")
                dr("Email") = aEmail '必填
                'If aNotes <> "" Then dr("Notes")=aNotes Else dr("Notes")=""
                dr("IsAgree") = If(aIsAgree <> "", aIsAgree, Convert.DBNull)
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
            Next
        End If
        DbAccess.UpdateDataTable(dt, da, trans)
    End Sub

    ''' <summary>'學員資料更新／新增--START </summary>
    ''' <param name="NewSetID_flag"></param>
    ''' <param name="trans"></param>
    ''' <param name="iSETID"></param>
    Sub UPDATE_STUD_ENTERTEMP(ByRef NewSetID_flag As Boolean, ByRef trans As SqlTransaction, ByRef iSETID As Integer)
        '學員資料更新／新增--START 
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing

        Dim sql As String = " SELECT * FROM STUD_ENTERTEMP WHERE IDNO='" & TIMS.ChangeIDNO(aIDNO) & "'"
        dt = DbAccess.GetDataTable(sql, da, trans)

        If dt.Rows.Count = 0 Then
            iSETID = DbAccess.GetNewId(trans, "STUD_ENTERTEMP_SETID_SEQ,STUD_ENTERTEMP,SETID")
            NewSetID_flag = True
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SETID") = iSETID  'NewSetID
            dr("IDNO") = TIMS.ChangeIDNO(aIDNO)
            Call sUtl_UpdateEnterTemp(dr)
            DbAccess.UpdateDataTable(dt, da, trans)
        Else
            For x As Integer = 0 To dt.Rows.Count - 1
                dr = dt.Rows(x) 'NewSetID=dr("SETID")
                iSETID = dr("SETID")
                Call sUtl_UpdateEnterTemp(dr)
            Next
            DbAccess.UpdateDataTable(dt, da, trans)
        End If
    End Sub

    ''' <summary>報名職類檔 更新／新增-Start </summary>
    ''' <param name="trans"></param>
    ''' <param name="iSETID"></param>
    ''' <param name="iSERNUM"></param>
    Sub UPDATE_STUD_ENTERTYPE(ByRef trans As SqlTransaction, ByRef iSETID As Integer, ByRef iSERNUM As Integer)
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing

        Dim sql As String = "SELECT * FROM STUD_ENTERTYPE WHERE SETID ='" & iSETID & "' AND EnterDate=" & TIMS.To_date(aEnterDate) & " AND SerNum ='" & iSERNUM & "'"
        dt = DbAccess.GetDataTable(sql, da, trans)
        If dt.Rows.Count <> 0 Then
            dr = dt.Rows(0) '原資料修改
        Else
            dr = dt.NewRow
            dt.Rows.Add(dr) '新增資料
            dr("SETID") = iSETID 'NewSetID
            dr("EnterDate") = aEnterDate 'dr("EnterDate")=Now().ToShortDateString '輸入日期
            dr("SerNum") = iSERNUM 'NewSerNum
            dr("RID") = aRID 'Session("RID") 'aRID '-- '別的單位執行匯入功能
            dr("PlanID") = aPlanID 'Session("PlanID") 'aPlanID '-- '別的單位執行匯入功能
            dr("RelEnterDate") = aEnterDate '報名日期
        End If
        'dr("RelEnterDate")=aRelEnterDate'已經有預設值，不須在輸入
        dr("ExamNo") = aNewExamNO 'NewExamNo1
        dr("OCID1") = aOCID1
        dr("TMID1") = aTMID1
        'If aOCID2 <> "" Then'    dr("OCID2")=aOCID2'    dr("TMID2")=NewTMID2'End If
        'If aOCID3 <> "" Then'    dr("OCID3")=aOCID3'    dr("TMID3")=NewTMID3'End If
        '1.網;2.現;3.通;4.推
        dr("EnterChannel") = If(aEnterChannel <> "", aEnterChannel, "2")
        dr("EnterPath") = "I" 'I匯入(報名登錄)
        dr("MIDENTITYID") = If(aMIdentityID <> "", aMIdentityID, Convert.DBNull)
        Dim vIdentityID As String = ""
        Dim IdentityID_A As String() = {aIdentityID1, aIdentityID2, aIdentityID3, aIdentityID4, aIdentityID5}
        For i2 As Integer = 0 To IdentityID_A.Length - 1
            IdentityID_A(i2) = TIMS.ClearSQM(IdentityID_A(i2))
            If IdentityID_A(i2) <> "" Then
                If vIdentityID <> "" Then vIdentityID &= ","
                vIdentityID &= IdentityID_A(i2)
            End If
        Next
        dr("IdentityID") = If(vIdentityID <> "", vIdentityID, Convert.DBNull)

        dr("TransDate") = If(aEnterDate <> "", TIMS.Cdate2(Now), Convert.DBNull)
        dr("Notes") = If(aNotes <> "", aNotes, Convert.DBNull)
        dr("ModifyAcct") = sm.UserInfo.UserID 'Session("UserID")
        dr("ModifyDate") = Now
        dr("APID1") = If(aAVTCP1_all <> "", aAVTCP1_all, Convert.DBNull)
        DbAccess.UpdateDataTable(dt, da, trans)

        'dr("TRNDMode")=If(aTRNDMode <> "", aTRNDMode, Convert.DBNull)
        'dr("TRNDType")=If(aTRNDType <> "", aTRNDType, Convert.DBNull)
        'dr("Ticket_NO")=If(aTicket_NO <> "", aTicket_NO, Convert.DBNull)

        'Dim v_NotExamID As String="0"
        'aNotExamID=TIMS.ClearSQM(aNotExamID)
        'Select Case aNotExamID
        '    Case "1", "Y", "YES"
        '        v_NotExamID=1
        '    Case "0", "N", "NO"
        '        v_NotExamID=0
        '    Case Else
        '        v_NotExamID=0
        'End Select
        'dr("NotExam")=If(v_NotExamID <> "", v_NotExamID, Convert.DBNull)
        'dr("TransDate")=aEnterDate '使用 aEnterDate 相同值
        '轉入日期 --判斷 aEnterDate <> ""
    End Sub

    ''' <summary>報名職類檔2 更新／新增-Start</summary>
    ''' <param name="trans"></param>
    ''' <param name="iSETID"></param>
    ''' <param name="iSERNUM"></param>
    Sub UPDATE_STUD_ENTERTRAIN(ByRef trans As SqlTransaction, ByRef iSETID As Integer, ByRef iSERNUM As Integer)
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim sql As String = ""

        sql = "" & vbCrLf
        sql &= " SELECT 'X'" & vbCrLf ' a.SENID  /*PK*/
        sql &= " FROM STUD_ENTERTRAIN a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.SETID =@SETID " & vbCrLf
        sql &= " AND a.EnterDate =@EnterDate " & vbCrLf
        sql &= " AND a.SerNum =@SerNum " & vbCrLf
        Dim parms As New Hashtable
        parms.Add("SETID", iSETID)
        parms.Add("EnterDate", CDate(aEnterDate))
        parms.Add("SerNum", iSERNUM)
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, trans, parms)

        If dt1.Rows.Count = 0 Then
            Dim i_sql As String = ""
            i_sql = "" & vbCrLf
            i_sql &= " INSERT INTO STUD_ENTERTRAIN ( SENID,SETID,ENTERDATE,SERNUM,ZIPCODE2,ZIPCODE2_6W,HOUSEHOLDADDRESS" & vbCrLf
            i_sql &= " ,UNAME,INTAXNO,SERVDEPT,JOBTITLE,ACTNAME,ACTTYPE,ACTNO,ACTTEL" & vbCrLf
            i_sql &= " ,ZIPCODE3,ZIPCODE3_6W,ACTADDRESS,SERVDEPTID,JOBTITLEID ,MODIFYACCT,MODIFYDATE)" & vbCrLf
            i_sql &= " VALUES ( @SENID,@SETID,@ENTERDATE,@SERNUM,@ZIPCODE2,@ZIPCODE2_6W,@HOUSEHOLDADDRESS" & vbCrLf
            i_sql &= " ,@UNAME,@INTAXNO,@SERVDEPT,@JOBTITLE,@ACTNAME,@ACTTYPE,@ACTNO,@ACTTEL" & vbCrLf
            i_sql &= " ,@ZIPCODE3,@ZIPCODE3_6W,@ACTADDRESS,@SERVDEPTID,@JOBTITLEID ,@MODIFYACCT,GETDATE())" & vbCrLf

            Dim iSENID As Integer = DbAccess.GetNewId(trans, "STUD_ENTERTRAIN_SENID_SEQ,STUD_ENTERTRAIN,SENID")

            Dim i_parms As New Hashtable
            i_parms.Add("SENID", iSENID)
            i_parms.Add("SETID", iSETID)
            i_parms.Add("EnterDate", CDate(aEnterDate))
            i_parms.Add("SerNum", iSERNUM)

            i_parms.Add("ZIPCODE2", TIMS.GetValue1(aZIPCODE2))
            i_parms.Add("ZIPCODE2_6W", TIMS.GetValue1(aZIPCODE2_6W))
            i_parms.Add("HOUSEHOLDADDRESS", If(aHOUSEHOLDADDRESS <> "", aHOUSEHOLDADDRESS, Convert.DBNull))

            i_parms.Add("UNAME", If(aUNAME <> "", aUNAME, Convert.DBNull))
            i_parms.Add("INTAXNO", If(aINTAXNO <> "", aINTAXNO, Convert.DBNull))
            i_parms.Add("SERVDEPT", TIMS.GetValue1(aSERVDEPT))
            i_parms.Add("JOBTITLE", TIMS.GetValue1(aJOBTITLE))
            i_parms.Add("ACTNAME", If(aACTNAME <> "", aACTNAME, Convert.DBNull))
            i_parms.Add("ACTTYPE", If(aACTTYPE <> "", aACTTYPE, Convert.DBNull))
            i_parms.Add("ACTNO", If(aACTNO <> "", aACTNO, Convert.DBNull))
            i_parms.Add("ACTTEL", If(aACTTEL <> "", aACTTEL, Convert.DBNull))

            i_parms.Add("ZIPCODE3", TIMS.GetValue1(aZIPCODE3))
            i_parms.Add("ZIPCODE3_6W", TIMS.GetValue1(aZIPCODE3_6W))
            i_parms.Add("ACTADDRESS", If(aACTADDRESS <> "", aACTADDRESS, Convert.DBNull))
            i_parms.Add("SERVDEPTID", TIMS.GetValue1(aSERVDEPTID))
            i_parms.Add("JOBTITLEID", TIMS.GetValue1(aJOBTITLEID))
            i_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            DbAccess.ExecuteNonQuery(i_sql, trans, i_parms)
        Else
            Dim u_sql As String = ""
            u_sql = "" & vbCrLf
            u_sql &= " UPDATE STUD_ENTERTRAIN" & vbCrLf ' a.SENID  /*PK*/
            u_sql &= " SET ZIPCODE2=@ZIPCODE2,ZIPCODE2_6W=@ZIPCODE2_6W" & vbCrLf
            u_sql &= " ,HOUSEHOLDADDRESS=@HOUSEHOLDADDRESS" & vbCrLf
            'u_sql &= " ,HANDTYPEID=@HANDTYPEID" & vbCrLf
            'u_sql &= " ,HANDLEVELID=@HANDLEVELID" & vbCrLf
            u_sql &= " ,UNAME=@UNAME" & vbCrLf
            u_sql &= " ,INTAXNO=@INTAXNO" & vbCrLf
            u_sql &= " ,SERVDEPT=@SERVDEPT" & vbCrLf
            u_sql &= " ,JOBTITLE=@JOBTITLE" & vbCrLf
            u_sql &= " ,ACTNAME=@ACTNAME" & vbCrLf
            u_sql &= " ,ACTTYPE=@ACTTYPE" & vbCrLf
            u_sql &= " ,ACTNO=@ACTNO" & vbCrLf
            u_sql &= " ,ACTTEL=@ACTTEL" & vbCrLf

            u_sql &= " ,ZIPCODE3=@ZIPCODE3" & vbCrLf
            u_sql &= " ,ZIPCODE3_6W=@ZIPCODE3_6W" & vbCrLf
            u_sql &= " ,ACTADDRESS=@ACTADDRESS" & vbCrLf
            u_sql &= " ,SERVDEPTID=@SERVDEPTID" & vbCrLf
            u_sql &= " ,JOBTITLEID=@JOBTITLEID" & vbCrLf
            u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
            u_sql &= " FROM STUD_ENTERTRAIN a" & vbCrLf
            u_sql &= " WHERE 1=1" & vbCrLf
            u_sql &= " AND SETID =@SETID " & vbCrLf
            u_sql &= " AND EnterDate =@EnterDate " & vbCrLf
            u_sql &= " AND SerNum =@SerNum " & vbCrLf

            Dim u_parms As New Hashtable
            u_parms.Add("ZIPCODE2", TIMS.GetValue1(aZIPCODE2))
            u_parms.Add("ZIPCODE2_6W", TIMS.GetValue1(aZIPCODE2_6W))
            u_parms.Add("HOUSEHOLDADDRESS", If(aHOUSEHOLDADDRESS <> "", aHOUSEHOLDADDRESS, Convert.DBNull))

            u_parms.Add("UNAME", If(aUNAME <> "", aUNAME, Convert.DBNull))
            u_parms.Add("INTAXNO", If(aINTAXNO <> "", aINTAXNO, Convert.DBNull))
            u_parms.Add("SERVDEPT", TIMS.GetValue1(aSERVDEPT))
            u_parms.Add("JOBTITLE", TIMS.GetValue1(aJOBTITLE))
            u_parms.Add("ACTNAME", If(aACTNAME <> "", aACTNAME, Convert.DBNull))
            u_parms.Add("ACTTYPE", If(aACTTYPE <> "", aACTTYPE, Convert.DBNull))
            u_parms.Add("ACTNO", If(aACTNO <> "", aACTNO, Convert.DBNull))
            u_parms.Add("ACTTEL", If(aACTTEL <> "", aACTTEL, Convert.DBNull))

            u_parms.Add("ZIPCODE3", TIMS.GetValue1(aZIPCODE3))
            u_parms.Add("ZIPCODE3_6W", TIMS.GetValue1(aZIPCODE3_6W))
            u_parms.Add("ACTADDRESS", If(aACTADDRESS <> "", aACTADDRESS, Convert.DBNull))
            u_parms.Add("SERVDEPTID", TIMS.GetValue1(aSERVDEPTID))
            u_parms.Add("JOBTITLEID", TIMS.GetValue1(aJOBTITLEID))
            u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)

            u_parms.Add("SETID", iSETID)
            u_parms.Add("EnterDate", CDate(aEnterDate))
            u_parms.Add("SerNum", iSERNUM)
            DbAccess.ExecuteNonQuery(u_sql, trans, u_parms)
        End If
    End Sub

    'Stud_EnterTemp
    Sub sUtl_UpdateEnterTemp(ByRef dr As DataRow)
        If aName <> "" Then aName = Trim(aName) 'aName=TIMS.ClearSQM(aName)
        dr("Name") = aName
        dr("Sex") = aSex
        dr("Birthday") = aBirthday
        dr("PassPortNO") = aPassPortNO
        dr("MaritalStatus") = If((New List(Of String) From {"1", "2", "3"}).Contains(aMaritalStatus), aMaritalStatus, Convert.DBNull)
        'dr("MaritalStatus")=If(aMaritalStatus="1", aMaritalStatus, If(aMaritalStatus="2", aMaritalStatus, If(aMaritalStatus="3", aMaritalStatus, Convert.DBNull)))
        aDegreeID = TIMS.ChangeIDNO(aDegreeID)
        dr("DegreeID") = If(aDegreeID <> "" AndAlso aDegreeID.Length < 2, "0" & aDegreeID, aDegreeID)
        aGradID = TIMS.ChangeIDNO(aGradID)
        dr("GradID") = If(aGradID <> "" AndAlso aGradID.Length < 2, "0" & aGradID, aGradID)
        dr("School") = aSchool
        dr("Department") = aDepartment

        If aMilitaryID.ToString <> "" Then
            If aMilitaryID.ToString.Length < 2 Then aMilitaryID = "0" & aMilitaryID
            ff = "MilitaryID='" & aMilitaryID & "'"
            If Key_Military.Select(ff).Length = 0 Then aMilitaryID = ""
        End If
        dr("MilitaryID") = If(aMilitaryID <> "", aMilitaryID, Convert.DBNull)

        dr("ZipCode") = TIMS.ChangeIDNO(aZipCode)
        dr("ZipCODE6W") = TIMS.ChangeIDNO(aZipCODE6W)
        dr("Address") = aAddress
        aPhone1 = TIMS.ChangeIDNO(aPhone1)
        aPhone2 = TIMS.ChangeIDNO(aPhone2)
        aCellPhone = TIMS.ChangeIDNO(aCellPhone)
        dr("Phone1") = If(aPhone1 <> "", aPhone1, "")
        dr("Phone2") = If(aPhone2 <> "", aPhone2, "")
        dr("CellPhone") = If(aCellPhone <> "", aCellPhone, "")

        dr("Email") = TIMS.ChangeIDNO(aEmail) '必填
        'If aNotes <> "" Then dr("Notes")=aNotes Else dr("Notes")=""
        dr("IsAgree") = If(aIsAgree <> "", aIsAgree, "")
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandName = "" Then Exit Sub
        If e.CommandArgument = "" Then Exit Sub
        '產業人才投資方案，暫停使用報名登錄功能
        Select Case e.CommandName
            Case "add"
                Dim aNow As Date = TIMS.GetSysDateNow(objconn)
                Dim sAltMsg As String = "" '訊息
                Dim flag_stopEnterH4 As Boolean = TIMS.StopEnterTempMsgH4(objconn, sAltMsg)
                If flag_stopEnterH4 Then
                    Common.MessageBox(Me, sAltMsg)
                    Return 'Exit Sub
                End If

                '新增
                GetSearchStr()
                Dim SETID As String = TIMS.GetMyValue(e.CommandArgument, "SETID")
                If SETID = "" Then Exit Sub

                If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Dim url1 As String = String.Concat(cst_SD01001_addaspx, "?ID=", TIMS.Get_MRqID(Me), "&serial=", SETID, "&proecess=add")
                    Call TIMS.Utl_Redirect(Me, objconn, url1)
                End If
                '產業人才投資方案，暫停使用報名登錄功能
                'Response.Redirect("SD_01_001_1_add.aspx?ID=" & Request("ID") & "&serial=" & e.CommandArgument & "&proecess=" & e.CommandName & "")
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Button1 As LinkButton = e.Item.FindControl("Button1")
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                'but=e.Item.Cells(4).Controls(1)
                Dim vSETID As String = TIMS.ClearSQM(drv("SETID"))
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SETID", vSETID)
                Button1.CommandArgument = sCmdArg
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim tmpConn As SqlConnection
        'Dim tmpTrans As SqlTransaction
        'Dim vsSETID As String="" '該學員
        'Dim vsEnterDate As String=""
        'Dim vsSerNum As String=""
        'Dim vsOCID1 As String="" '報名班級
        Dim flagNG1 As Boolean = If(e.CommandArgument = "", True, False) ' False
        'If e.CommandArgument="" Then flagNG1=True
        If flagNG1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim sCmdArg As String = e.CommandArgument
        Select Case e.CommandName
            Case "edit"
                Dim CM_proecess As String = TIMS.GetMyValue(sCmdArg, "proecess")
                Dim CM_serial As String = TIMS.GetMyValue(sCmdArg, "serial")
                Dim CM_EnterDate As String = TIMS.GetMyValue(sCmdArg, "EnterDate")
                Dim CM_SerNum As String = TIMS.GetMyValue(sCmdArg, "SerNum")
                Dim CM_STDate As String = TIMS.GetMyValue(sCmdArg, "STDate")
                If CM_proecess = "" Then flagNG1 = True
                If CM_serial = "" Then flagNG1 = True
                If CM_EnterDate = "" Then flagNG1 = True
                If CM_SerNum = "" Then flagNG1 = True
                If CM_STDate = "" Then flagNG1 = True
                If flagNG1 Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
                '修改
                Call GetSearchStr()
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '產業人才投資方案，暫停使用報名登錄功能
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
                Dim url1 As String = String.Concat(cst_SD01001_addaspx, "?ID=", TIMS.Get_MRqID(Me), e.CommandArgument)
                Call TIMS.Utl_Redirect(Me, objconn, url1)

            Case "del"
                '刪除
                Dim vsSETID As String = "" '該學員
                Dim vsOCID1 As String = "" '報名班級
                Dim sql As String
                Dim msgbox As String = ""
                Dim dr As DataRow
                Dim dt As DataTable
                Dim da As SqlDataAdapter = Nothing
                'Call TIMS.OpenDbConn(objconn)

                sql = String.Concat(" SELECT * FROM STUD_SELRESULT ", e.CommandArgument)
                dr = DbAccess.GetOneRow(sql, objconn)
                If dr IsNot Nothing Then msgbox += "此學員所參加的班級已經試算過，無法進行刪除!" & vbCrLf

                '無意外訊息可進行刪除
                If msgbox = "" Then
                    sql = String.Concat(" SELECT * FROM STUD_ENTERTYPE ", e.CommandArgument)
                    dr = DbAccess.GetOneRow(sql, objconn)
                    If dr IsNot Nothing Then
                        '為了修改E網為報名失敗而使用
                        vsSETID = Convert.ToString(dr("SETID"))
                        vsOCID1 = Convert.ToString(dr("OCID1"))
                        'vsEnterDate=CDate(dr("EnterDate")).ToString("yyyy/MM/dd")
                        'vsSerNum=Convert.ToString(dr("SerNum"))
                        '為了修改E網為報名失敗而使用
                        'Me.ViewState("SETID")=dr("SETID")
                        'Me.ViewState("OCID1")=dr("OCID1")
                        'Select Case dr("TRNDMode").ToString
                        '    Case "1"
                        '        sql="UPDATE Adp_TRNData SET TransToTIMS='N' WHERE TICKET_NO='" & dr("TICKET_NO") & "'"
                        '        DbAccess.ExecuteNonQuery(sql, objconn)
                        '    Case "2"
                        '        sql="UPDATE Adp_DGTRNData SET TransToTIMS='N' WHERE TICKET_NO='" & dr("TICKET_NO") & "'"
                        '        DbAccess.ExecuteNonQuery(sql, objconn)
                        '    Case "3"
                        '        sql="UPDATE Adp_GOVTRNData SET TransToTIMS='N' WHERE TICKET_NO='" & dr("TICKET_NO") & "'"
                        '        DbAccess.ExecuteNonQuery(sql, objconn)
                        'End Select
                    End If

                    '20090511(Milor)目前先保留資料備份到Stud_DelEnterType的機制，並將資料再備一份到Stud_EnterTypeDelData，
                    '避免同時需要修改的程式過多。
                    'Dim dr1, dr5 As DataRow
                    sql = ""
                    sql &= " SELECT * " & vbCrLf
                    sql &= " FROM STUD_ENTERTYPE" & e.CommandArgument
                    Dim dr5 As DataRow = DbAccess.GetOneRow(sql, objconn)
                    'sql="SELECT * FROM Stud_DelEnterType WHERE 1<>1" '把資料從Stud_EnterType 寫到 Stud_DelEnterType 開始
                    '刪除列為dr5
                    sql = ""
                    sql &= " SELECT * " & vbCrLf
                    sql &= " FROM STUD_DELENTERTYPE" & e.CommandArgument '把資料從Stud_EnterType 寫到 Stud_DelEnterType 開始
                    dt = DbAccess.GetDataTable(sql, da, objconn)
                    If dr5 IsNot Nothing Then
                        Dim dr1 As DataRow
                        If dt.Rows.Count > 0 Then
                            dr1 = dt.Rows(0)
                        Else
                            dr1 = dt.NewRow
                            dt.Rows.Add(dr1)
                            dr1("SETID") = dr5("SETID")
                            dr1("EnterDate") = dr5("EnterDate")
                            dr1("SerNum") = dr5("SerNum")
                        End If

                        dr1("ExamNo") = dr5("ExamNo")
                        dr1("OCID1") = dr5("OCID1")
                        dr1("TMID1") = dr5("TMID1")
                        dr1("OCID2") = dr5("OCID2")
                        dr1("TMID2") = dr5("TMID2")
                        dr1("OCID3") = dr5("OCID3")
                        dr1("TMID3") = dr5("TMID3")
                        dr1("WriteResult") = dr5("WriteResult")
                        dr1("OralResult") = dr5("OralResult")
                        dr1("TotalResult") = dr5("TotalResult")
                        dr1("EnterChannel") = dr5("EnterChannel")
                        dr1("EnterPath") = dr5("EnterPath")
                        dr1("IdentityID") = dr5("IdentityID")
                        dr1("RID") = dr5("RID")
                        dr1("PlanID") = dr5("PlanID")
                        dr1("TRNDMode") = dr5("TRNDMode")
                        dr1("TRNDType") = dr5("TRNDType")
                        dr1("Q1_1") = dr5("Q1_1")
                        dr1("Q1_2") = dr5("Q1_2")
                        dr1("Q1_2Other") = dr5("Q1_2Other")
                        dr1("Q1_3") = dr5("Q1_3")
                        dr1("Q1_3Other") = dr5("Q1_3Other")
                        dr1("Q1_4") = dr5("Q1_4")
                        dr1("Q1_4Other") = dr5("Q1_4Other")
                        dr1("Q1_5") = dr5("Q1_5")
                        dr1("Q2_3") = dr5("Q2_3")
                        dr1("Q2_4") = dr5("Q2_4")
                        dr1("Q2_5Other") = dr5("Q2_5Other")
                        dr1("ModifyAcct") = sm.UserInfo.UserID
                        dr1("ModifyDate") = Now
                        dr1("Ticket_NO") = dr5("Ticket_NO")
                        dr1("RelEnterDate") = dr5("RelEnterDate")
                        dr1("NotExam") = dr5("NotExam")
                        dr1("CCLID") = dr5("CCLID")
                        dr1("eSETID") = dr5("eSETID")
                        dr1("eSerNum") = dr5("eSerNum")
                        dr1("TransDate") = dr5("TransDate")
                        dr1("SEID") = dr5("SEID")
                        dr1("SupplyID") = dr5("SupplyID")
                        dr1("BudID") = dr5("BudID")
                        DbAccess.UpdateDataTable(dt, da)

                        Dim tConn As SqlConnection = DbAccess.GetConnection()
                        Dim tmpTrans As SqlTransaction = DbAccess.BeginTrans(tConn)
                        'Call TIMS.TestDbConn(Me, tmpConn, True)
                        'tmpTrans=tmpConn.BeginTransaction()
                        'Del_StudEnterType(Convert.ToString(dr5("SETID")), Convert.ToString(dr5("OCID1")), tmpConn, tmpTrans)
                        'Del_StudSelResult(Convert.ToString(dr5("SETID")), Convert.ToString(dr5("OCID1")), tmpConn, tmpTrans)
                        Call Del_StudEnterType(Convert.ToString(e.CommandArgument), tConn, tmpTrans)
                        Call Del_StudSelResult(Convert.ToString(e.CommandArgument), tConn, tmpTrans)
                        tmpTrans.Commit()
                        tmpTrans.Dispose()
                        Call TIMS.CloseDbConn(tConn)
                    End If



                    'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
                    sql = "UPDATE STUD_ENTERTYPE2 SET signUpStatus='2' " & e.CommandArgument
                    DbAccess.ExecuteNonQuery(sql, objconn)
                    If vsSETID <> "" AndAlso vsOCID1 <> "" Then
                        sql = " UPDATE STUD_ENTERTYPE2 SET signUpStatus='2' " & vbCrLf 'E網: 2:報名失敗 '5:未錄取
                        sql &= " ,ModifyAcct='" & Convert.ToString(sm.UserInfo.UserID) & "'" & vbCrLf
                        sql &= " ,ModifyDate=getdate() " & vbCrLf
                        sql &= " WHERE SETID=" & vsSETID & vbCrLf
                        'sql += "  AND EnterDate=" & vsEnterDate & vbCrLf
                        'sql += "  AND SerNum=" & vsSerNum & vbCrLf
                        sql &= "  AND OCID1=" & vsOCID1 & vbCrLf
                        DbAccess.ExecuteNonQuery(sql, objconn)
                    End If

                    'Try
                    'Catch ex As Exception
                    '    Dim strErrmsg1 As String=""
                    '    strErrmsg1 &= "Private Sub DataGrid2_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand" & vbCrLf
                    '    strErrmsg1 &= TIMS.GetErrorMsg(Me) & vbCrLf '取得錯誤資訊寫入
                    '    'strErrmsg &= "SQL:" & vbCrLf & Sql & vbCrLf
                    '    strErrmsg1 &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
                    '    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    '    Call TIMS.WriteTraceLog(strErrmsg1)
                    '    Throw ex
                    'End Try
                    msgbox = "刪除成功!"
                End If

                Common.MessageBox(Me, msgbox)
                'Button2_Click(Button2, e)
                Call Search1()  '查詢鈕
            Case "print"
                'Dim cGuid As String=  ReportQuery.GetGuid(Page)
                'Dim Url As String=  ReportQuery.GetUrl(Page)
                'Dim strScript As String
                'strScript="<script language=""javascript"">" + vbCrLf
                'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=SD_01_001&path=TIMS&SETID=" & e.CommandArgument & "&ExamNo=" & e.Item.Cells(7).Text & "&SerNum=" & e.Item.Cells(8).Text & "');" + vbCrLf
                'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=SD_01_001&path=TIMS" & e.CommandArgument & "');" + vbCrLf
                'strScript += "</script>"
                'Page.RegisterStartupScript("window_onload", strScript)
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                If Not blnExport AndAlso Me.ViewState("sort") IsNot Nothing Then
                    Dim img As New UI.WebControls.Image
                    Dim i As Integer
                    Select Case (Me.ViewState("sort"))
                        Case "NAME", "NAME desc"
                            i = cst_姓名 '2
                        Case "IDNO_MK", "IDNO_MK desc"
                            i = cst_身分證號碼 '3
                        Case "ORGNAME", "ORGNAME desc"
                            i = cst_報名機構 '4
                        Case "CLASSCNAME1B", "CLASSCNAME1B desc"
                            i = cst_報名班級 '5
                        Case "RELENTERDATE", "RELENTERDATE desc"
                            i = cst_報名日期 '6
                        Case "EXAMNO", "EXAMNO desc" 'ExamNO
                            i = cst_准考證號碼 '7
                    End Select

                    Dim flag_vs_sort_desc As Boolean = True
                    If Me.ViewState("sort").ToString.IndexOf("desc") = -1 Then flag_vs_sort_desc = False
                    img.ImageUrl = If(flag_vs_sort_desc, "../../images/SortDown.gif", "../../images/SortUp.gif")
                    e.Item.Cells(i).Controls.Add(img)  '不是匯出
                End If

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btnEditView6 As LinkButton = e.Item.FindControl("btnEditView6") '修改／檢視鈕
                Dim Button4 As LinkButton = e.Item.FindControl("Button4") '刪除鈕
                'Dim btn3 As Button=e.Item.FindControl("Button7") '列印按鈕
                Dim RelEnterDate As Label = e.Item.FindControl("RelEnterDate")
                Dim LNO As Label = e.Item.FindControl("LNO")
                Dim star3 As Label = e.Item.FindControl("star3")
                'Dim LBudgetID97 As Label=e.Item.FindControl("LBudgetID97")
                LNO.Text = TIMS.Get_DGSeqNo(sender, e)
                star3.Visible = False
                If TIMS.Chk_StudStatus(drv("IDNO").ToString, drv("STDate").ToString, drv("FTDate").ToString, drv("OCID1").ToString, objconn) Then star3.Visible = True  '學員是否在訓的檢核 2008/5/14

                Dim bln_View As Boolean = False '修改／檢視鈕
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "proecess", "edit")
                TIMS.SetMyValue(sCmdArg, "serial", Convert.ToString(drv("SETID")))
                TIMS.SetMyValue(sCmdArg, "EnterDate", Convert.ToString(drv("EnterDate")))
                TIMS.SetMyValue(sCmdArg, "SerNum", Convert.ToString(drv("SerNum")))
                TIMS.SetMyValue(sCmdArg, "STDate", Convert.ToString(drv("STDate")))

                '如果過了開訓日期,或是已做參訓功能,或是不是這個報名單位的使用者登入,就顯示檢視
                If drv("RID") <> sm.UserInfo.RID Or drv("STDate") < DateTime.Now Or drv("AppliedStatus") = "Y" Then bln_View = True

                If oflag_Test Then bln_View = False '測試用

                '檢視/ '修改
                Const cst_檢視 As String = "檢視"
                If bln_View Then
                    TIMS.SetMyValue(sCmdArg, "view", "1") '檢視 'cmdArg += "&view=1" '檢視 
                    btnEditView6.Text = cst_檢視 '"檢視"
                Else
                    '修改'(再檢測一次修改權限)
                    'If btnEditView6.Enabled Then
                    '    btnEditView6.Enabled=False
                    '    If au.blnCanMod Then btnEditView6.Enabled=True
                    '    If Not btnEditView6.Enabled Then TIMS.Tooltip(btnEditView6, "停用 修改功能!")
                    'End If
                End If
                btnEditView6.CommandArgument = sCmdArg

                Dim flag_can_delete As Boolean = True
                If sm.UserInfo.RoleID = 0 Then
                    '系統管理者不可刪除-避免問題
                    flag_can_delete = False
                End If

                If Not flag_can_delete Then
                    '系統管理者不可刪除
                    Button4.Enabled = False
                    If Not Button4.Enabled Then TIMS.Tooltip(Button4, "停用 刪除功能!")
                End If
                Button4.Attributes("onclick") = "return confirm('這樣會刪除這個人的報名資料，\n但不會刪除個人基本資料，\n確定要刪除?');"

                '就服站不管如何都不可刪除
                If Button4.Enabled AndAlso Convert.ToString(drv("EnterPath")) = cst_EnterPathW Then
                    '不可刪除
                    Button4.Enabled = False
                    If Not Button4.Enabled Then TIMS.Tooltip(Button4, "停用 刪除功能!")
                End If

                If Button4.Enabled AndAlso Convert.ToString(drv("EnterPath")) = cst_EnterPathR Then
                    '不可刪除
                    Button4.Enabled = False
                    If Not Button4.Enabled Then TIMS.Tooltip(Button4, "停用 刪除功能!")
                    If Convert.ToString(drv("WSIDName")) <> "" Then
                        TIMS.Tooltip(Button4, Convert.ToString(drv("WSIDName")))
                    End If
                End If

                'If Convert.ToString(drv("ORGNAME2")) <> "" Then e.Item.Cells(cst_報名機構).Text="<font color='Blue'>" & TIMS.ClearSQM(drv("ORGNAME2")) & "</font>-" & TIMS.ClearSQM(drv("ORGNAME"))

                e.Item.Cells(cst_報名日期).Text = ""
                Dim str_HTML1 As String = ""
                If Not blnExport Then '不是匯出
                    str_HTML1 = String.Concat("<Table class='font' bgcolor='#FFFFE6' width='300' id='ClassData", e.Item.ItemIndex, "' style='DISPLAY: none; POSITION: absolute;BORDER-COLLAPSE: collapse' border=1>")
                    str_HTML1 &= String.Concat("<TR><TD>", "報名班級:", Convert.ToString(drv("CLASSCNAME1B")), "</TD></TR>")
                    str_HTML1 &= "</Table>"
                End If
                e.Item.Cells(cst_報名日期).Text = String.Concat(str_HTML1, TIMS.Cdate3(drv("RelEnterDate"))) ' FormatDateTime(drv("RelEnterDate"), 2)

                Dim sEnterChannel As String = GET_EnterChannel_N(Convert.ToString(drv("EnterChannel")))
                If Convert.ToString(drv("EnterPath")) = cst_EnterPathW Then
                    If Convert.ToString(drv("WSIDName")) <> "" Then
                        sEnterChannel &= String.Concat("<br />", Convert.ToString(drv("WSIDName")))
                    Else
                        sEnterChannel &= cst_EnterPathNameW
                    End If
                End If
                If Convert.ToString(drv("EP2PY")) = "Y" Then sEnterChannel = "專案核定報名" '置換

                'Dim flagPS1 As Boolean=TIMS.Chk_TICKETPS1(objconn, Convert.ToString(drv("OCID1")), Convert.ToString(drv("IDNO")))
                'If flagPS1 Then sEnterChannel="推介(註銷)"
                e.Item.Cells(cst_報名管道).Text = sEnterChannel  '報名管道

                If drv("OCID").ToString = "" Then
                    e.Item.Cells(cst_是否試算).Text = "否"
                    e.Item.Cells(cst_錄取結果).Text = ""
                Else
                    Select Case Convert.ToString(drv("Admission"))
                        Case "Y"
                            e.Item.Cells(cst_是否試算).Text = "是"
                            e.Item.Cells(cst_錄取結果).Text = "是"
                        Case "N"
                            e.Item.Cells(cst_是否試算).Text = "是"
                            e.Item.Cells(cst_錄取結果).Text = "否"
                            e.Item.Cells(cst_錄取結果).ForeColor = Color.Red
                        Case Else
                            e.Item.Cells(cst_是否試算).Text = "是"
                            e.Item.Cells(cst_錄取結果).Text = ""
                    End Select
                End If

                '給予 Sql搜尋條件，直接使用
                Button4.CommandArgument = ""
                If Convert.ToString(drv("SETID")) <> "" AndAlso Convert.ToString(drv("EnterDate")) <> "" AndAlso Convert.ToString(drv("SerNum")) <> "" Then
                    'Dim tdEnterDate As String=TIMS.to_date(FormatDateTime(drv("EnterDate"), DateFormat.ShortDate))
                    'Dim tdEnterDate As String=TIMS.to_date(TIMS.cdate3(drv("EnterDate")))
                    Button4.CommandArgument = String.Concat(" WHERE SETID='", drv("SETID"), "' AND EnterDate=", TIMS.To_date(TIMS.Cdate3(drv("EnterDate"))), " AND SerNum='", drv("SerNum"), "'")
                End If
                'btn3.CommandArgument="&SETID=" & drv("SETID") & "&ExamNo=" & drv("ExamNo") & "&SerNum=" & drv("SerNum")
                'btn3.Attributes("onclick")=    ReportQuery.ReportScript(Me, "list", "SD_01_001", "&SETID=" & drv("SETID") & "&ExamNo=" & drv("ExamNo") & "&SerNum=" & drv("SerNum"))
                e.Item.Cells(cst_報名班級).Attributes("onmouseover") = "document.getElementById('ClassData" & e.Item.ItemIndex & "').style.display='inline';"
                e.Item.Cells(cst_報名班級).Attributes("onmouseout") = "document.getElementById('ClassData" & e.Item.ItemIndex & "').style.display='none';"

                Dim Hid_IDNO_MK As HtmlInputHidden = e.Item.FindControl("Hid_IDNO_MK")
                Hid_IDNO_MK.Value = TIMS.EncryptAes(drv("IDNO")) '加密 Aes
                Dim ExamNO As HtmlInputHidden = e.Item.FindControl("ExamNO")
                Dim SETID As HtmlInputHidden = e.Item.FindControl("SETID")
                Dim SerNum As HtmlInputHidden = e.Item.FindControl("SerNum")
                ExamNO.Value = Convert.ToString(drv("ExamNO"))
                SETID.Value = drv("SETID")
                SerNum.Value = drv("SerNum")
                'LBudgetID97.Text=If(Convert.ToString("BudID")="97", "是", "否")

                'Dim drJ As DataRow=Nothing '(VIEW_GETJOBC1)
                'drJ=TIMS.Getjobc1(Convert.ToString(drv("IDNO")), Convert.ToString(drv("OCID1")), objconn)
                'If drJ IsNot Nothing Then
                '    e.Item.Cells(cst_結訓情形).Text=Convert.ToString(drJ("StudStatusN"))
                '    e.Item.Cells(cst_就業情形).Text=Convert.ToString(drJ("IsGetJobN"))
                'End If

                Dim drStud As DataRow = Nothing '(VIEW_GETJOBC1)
                drStud = TIMS.Get_V_STUDENTINFO(Convert.ToString(drv("IDNO")), Convert.ToString(drv("OCID1")), objconn)
                If drStud IsNot Nothing Then e.Item.Cells(cst_結訓情形).Text = Convert.ToString(drStud("STUDSTATUS2"))

                'ff="IDNO='" & Convert.ToString(drv("IDNO")) & "' AND OCID=" & Convert.ToString(drv("OCID1"))
                'If dtGetjobc1 Is Nothing Then
                '    dtGetjobc1=PageControler2.PageDataTable2
                'End If
                'If Not dtGetjobc1 Is Nothing Then
                '    If dtGetjobc1.Select(ff).Length > 0 Then
                '        e.Item.Cells(cst_結訓情形).Text=Convert.ToString(dtGetjobc1.Select(ff)(0)("StudStatusN"))
                '        e.Item.Cells(cst_就業情形).Text=Convert.ToString(dtGetjobc1.Select(ff)(0)("IsGetJobN"))
                '    End If
                'End If

        End Select
    End Sub

    Private Sub DataGrid2_SortCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridSortCommandEventArgs) Handles DataGrid2.SortCommand
        Me.ViewState("sort") = If(e.SortExpression = Me.ViewState("sort"), e.SortExpression & " desc", e.SortExpression)

        PageControler1.Sort = Me.ViewState("sort")
        PageControler1.ChangeSort()
    End Sub

    Sub GetSearchStr()
        Dim str_SearchStr_x1 As String = ""
        str_SearchStr_x1 = "center=" & center.Text
        str_SearchStr_x1 += "&RIDValue=" & RIDValue.Value
        str_SearchStr_x1 += "&OCID1=" & OCID1.Text
        str_SearchStr_x1 += "&TMID1=" & TMID1.Text
        str_SearchStr_x1 += "&OCIDValue1=" & OCIDValue1.Value
        str_SearchStr_x1 += "&TMIDValue1=" & TMIDValue1.Value
        str_SearchStr_x1 += "&IDNO=" & TIMS.ChangeIDNO(IDNO.Text)
        str_SearchStr_x1 += "&start_date=" & start_date.Text
        str_SearchStr_x1 += "&end_date=" & end_date.Text
        str_SearchStr_x1 += "&PageIndex=" & DataGrid2.CurrentPageIndex + 1
        str_SearchStr_x1 += If(table5.Visible, "&submit=1", "&submit=0")

        Session("_SearchStr") = str_SearchStr_x1
    End Sub

    '檢查e網前台 報名重複資料 MEM_060    'TIMS SD_01_001_add 報名登錄功能    'TIMS SD_01_001 報名登錄功能
    Function Check_E_ClsTrace(ByRef Errmsg As String, ByVal tmpIDNO As String, ByVal tmpOCID1 As String) As Boolean
        ', ByVal tmpOCID2 As String, ByVal tmpOCID3 As String
        Check_E_ClsTrace = False
        Dim OCIDs As String = ""
        OCIDs = ""
        If tmpOCID1 <> "" Then
            If OCIDs <> "" Then OCIDs &= ","
            OCIDs &= tmpOCID1
        End If
        'If tmpOCID2 <> "" Then
        '    If OCIDs <> "" Then OCIDs &= ","
        '    OCIDs &= tmpOCID2
        'End If
        'If tmpOCID3 <> "" Then
        '    If OCIDs <> "" Then OCIDs &= ","
        '    OCIDs &= tmpOCID3
        'End If

        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Dim dt2 As DataTable

        Errmsg = ""
        If OCIDs <> "" Then
            sql = "" & vbCrLf
            sql &= " select i.ORGID " & vbCrLf
            sql &= " ,i.OrgName " & vbCrLf
            sql &= " ,a.PlanID " & vbCrLf
            sql &= " ,a.ComIDNO " & vbCrLf
            sql &= " ,a.SeqNo " & vbCrLf
            sql &= " ,a.RID " & vbCrLf
            sql &= " ,a.OCID " & vbCrLf
            sql &= " ,a.ExamDate " & vbCrLf
            sql &= " ,a.ExamPeriod " & vbCrLf
            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
            sql &= " FROM CLASS_CLASSINFO a " & vbCrLf
            sql &= " JOIN ORG_ORGINFO i on i.ComIDNO=a.ComIDNO " & vbCrLf
            sql &= " WHERE 1=1 " & vbCrLf
            sql &= " AND a.OCID IN (" & OCIDs & ")" & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)

            If dt.Rows.Count > 0 Then
#Region "(No Use)"

                '狀況一(當本次報名課程清單中,有同一培訓單位,且甄試日期為同一天的報名課程),顯示訊息如下
                'For i As Integer=0 To dt.Rows.Count - 1 '本次報名課程，各個課程判斷
                '    dr=dt.Rows(i)
                '    If dr("ExamDate").ToString <> "" And IsDate(dr("ExamDate")) Then
                '        dt2=Nothing
                '        sql="" & vbCrLf
                '        sql &= " 	SELECT  i.ORGID, i.OrgName" & vbCrLf
                '        sql &= "  	,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate, a.ExamPeriod" & vbCrLf
                '        sql &= "    ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                '        sql &= " 	FROM CLASS_CLASSINFO a  " & vbCrLf
                '        sql &= " 	join Org_OrgInfo i on i.ComIDNO=a.ComIDNO  " & vbCrLf
                '        sql &= " 	WHERE 1=1 " & vbCrLf
                '        sql &= " 	AND a.OCID IN (" & OCIDs & ")" & vbCrLf
                '        sql &= " 	and i.ORGID=" & dr("ORGID") & vbCrLf
                '        sql &= " 	and a.ExamDate='" & Common.FormatDate(dr("ExamDate")) & "'" & vbCrLf
                '        Select Case Convert.ToString(dr("ExamPeriod"))
                '            Case "02" '02:上午 (01:全天或02:同上午)
                '                sql &= " 	and a.ExamPeriod IN ('01','02') " & vbCrLf
                '            Case "03" '03:下午 (01:全天或03:同下午)
                '                sql &= " 	and a.ExamPeriod IN ('01','03') " & vbCrLf
                '            Case Else '01:全天 (其他全天)
                '        End Select
                '        dt2=DbAccess.GetDataTable(sql)
                '        If dt2.Rows.Count > 1 Then '大於兩筆,有同一培訓單位,且甄試日期為同一天的報名課程
                '            Errmsg += " 您所選擇由 「" & dt2.Rows(0)("OrgName").ToString & "」 (" & dt2.Rows(0)("ORGID").ToString & ")(培訓單位)\r\n"
                '            Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ")\r\n"
                '            Errmsg += " 報名課程2: 「" & dt2.Rows(1)("ClassName").ToString & "」 (" & dt2.Rows(1)("OCID").ToString & ")\r\n"
                '            Errmsg += " 甄試作業為同一天舉辦 (" & Common.FormatDate(dr("ExamDate")) & ")，請擇一報名，以確保您後續的權益，謝謝!"
                '            Exit For
                '        End If
                '    End If
                'Next

                ' IDNO@Session("UsrMemberID")@tmpIDNO

#End Region

                '狀況二(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
                ',且甄試日期為同一天的報名課程,顯示訊息如下

                '您所選擇由XXXXX(訓練單位名稱)開訓之課程XXXXX(課程名稱)
                '，與您日前已完成報名之XXXXX(課程名稱)甄試作業為同一天舉辦
                '，因故無法受理您此次報名需求，請見諒。

                '狀況三(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
                ',且甄試日期為同一天的報名課程,但已報名課程之報名資料被審核失敗,則容許報名本次)
                '(e網審核成功,錄取作業為未選擇,正取與備取者.)
                '請查閱： Get_Stud_EnterType2_OCIDs
                'If Errmsg="" Then OCIDs2=Get_Stud_EnterType2_OCIDs(tmpIDNO) '收件完成(審核中)，報名成功


                '狀況四(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
                ',但已報名課程之報名資料被審核失敗,則容許報名本次，否則不可報名)
                If Errmsg = "" Then
                    Dim OCIDs2 As String = Get_Stud_EnterType2_OCIDs(tmpIDNO) '收件完成(審核中)，報名成功
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dr = dt.Rows(i) '要報名的班級
                        If OCIDs2 <> "" Then
                            sql = "" & vbCrLf
                            sql &= " select  i.ORGID " & vbCrLf
                            sql &= " ,i.OrgName " & vbCrLf
                            sql &= " ,a.PlanID " & vbCrLf
                            sql &= " ,a.ComIDNO " & vbCrLf
                            sql &= " ,a.SeqNo " & vbCrLf
                            sql &= " ,a.RID " & vbCrLf
                            sql &= " ,a.OCID " & vbCrLf
                            sql &= " ,a.ExamDate " & vbCrLf
                            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                            sql &= " FROM CLASS_CLASSINFO a " & vbCrLf
                            sql &= " join Org_OrgInfo i on i.ComIDNO=a.ComIDNO " & vbCrLf
                            sql &= " WHERE 1=1 " & vbCrLf
                            sql &= " AND a.OCID IN (" & OCIDs2 & ") " & vbCrLf
                            sql &= " AND i.ORGID=" & dr("ORGID") & " " & vbCrLf
                            sql &= " AND a.OCID='" & dr("OCID").ToString & "' " & vbCrLf
                            dt2 = DbAccess.GetDataTable(sql, objconn)

                            If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,尚在審核中或報名成功
                                Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位)\r\n"
                                Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") \r\n"
                                Errmsg += " 與您日前已完成報名之 \r\n 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") \r\n"
                                Errmsg += " 已在e網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                                Exit For
                            End If
                        End If
                    Next
                End If

#Region "(No Use)"

                '同狀況三
                '狀況四-2(當本次報名課程清單中,與過去已報名課程,為同一培訓單位(甄試作業為同一天舉辦)
                ',但已報名課程之報名資料被審核失敗,則容許報名本次，否則不可報名)
                'If Errmsg="" Then
                '    If OCIDs2="" Then
                '        OCIDs2=Get_Stud_EnterType2_OCIDs(TMPIDNO) '收件完成(審核中)，報名成功
                '    End If
                '    For i As Integer=0 To dt.Rows.Count - 1
                '        dr=dt.Rows(i) '要報名的班級
                '        If OCIDs2 <> "" Then
                '            If dr("ExamDate").ToString <> "" And IsDate(dr("ExamDate")) Then
                '                sql="" & vbCrLf
                '                sql &= " 	SELECT  i.ORGID, i.OrgName" & vbCrLf
                '                sql &= "  	,a.PlanID, a.ComIDNO, a.SeqNo, a.RID, a.OCID, a.ExamDate" & vbCrLf
                '                sql &= "    ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                '                sql &= " 	FROM CLASS_CLASSINFO a  " & vbCrLf
                '                sql &= " 	join Org_OrgInfo i on i.ComIDNO=a.ComIDNO  " & vbCrLf
                '                sql &= " 	WHERE 1=1 " & vbCrLf
                '                sql &= " 	AND a.OCID IN (" & OCIDs2 & ")" & vbCrLf
                '                sql &= " 	and i.ORGID=" & dr("ORGID") & vbCrLf
                '                'sql += " 	AND a.OCID='" & dr("OCID").ToString & "'" & vbCrLf
                '                sql &= " 	and a.ExamDate='" & Common.FormatDate(dr("ExamDate")) & "'" & vbCrLf
                '                dt2=DbAccess.GetDataTable(sql)
                '                If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,尚在審核中或報名成功
                '                    Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位)\r\n"
                '                    Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") \r\n"
                '                    Errmsg += " 與您日前已完成報名之 \r\n 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") \r\n"
                '                    Errmsg += " 甄試作業為同一天舉辦 (" & Common.FormatDate(dr("ExamDate")) & ") "
                '                    Errmsg += " 已在e網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                '                    Exit For
                '                End If
                '            End If
                '        End If
                '    Next
                'End If

#End Region

                '狀況五(已在內網有報名資料)
                'Dim OCIDs1 As String=""
                'OCIDs1=""
                If Errmsg = "" Then
                    Dim OCIDs1 As String = Get_Stud_EnterType_OCIDs(tmpIDNO) '已在內網有報名資料
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dr = dt.Rows(i) '要報名的班級
                        If OCIDs1 <> "" Then
                            sql = "" & vbCrLf
                            sql &= " select i.ORGID " & vbCrLf
                            sql &= " ,i.OrgName " & vbCrLf
                            sql &= " ,a.PlanID " & vbCrLf
                            sql &= " ,a.ComIDNO " & vbCrLf
                            sql &= " ,a.SeqNo " & vbCrLf
                            sql &= " ,a.RID " & vbCrLf
                            sql &= " ,a.OCID " & vbCrLf
                            sql &= " ,a.ExamDate " & vbCrLf
                            sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSNAME" & vbCrLf
                            sql &= " FROM CLASS_CLASSINFO a  " & vbCrLf
                            sql &= " join Org_OrgInfo i on i.ComIDNO=a.ComIDNO " & vbCrLf
                            sql &= " WHERE 1=1 " & vbCrLf
                            sql &= " AND a.OCID IN (" & OCIDs1 & ") " & vbCrLf
                            sql &= " and i.ORGID=" & dr("ORGID") & " " & vbCrLf
                            sql &= " AND a.OCID='" & dr("OCID").ToString & "' " & vbCrLf
                            dt2 = DbAccess.GetDataTable(sql, objconn)
                            If dt2.Rows.Count > 0 Then '大於兩筆,為同一培訓單位,已在內網有報名資料
                                Errmsg += " 您所選擇由 「" & dr("OrgName").ToString & "」 (" & dr("ORGID").ToString & ")(培訓單位)\r\n"
                                Errmsg += " 開訓之課程 \r\n 報名課程1: 「" & dr("ClassName").ToString & "」 (" & dr("OCID").ToString & ") \r\n"
                                Errmsg += " 與您日前已完成報名之 \r\n 報名課程2: 「" & dt2.Rows(0)("ClassName").ToString & "」 (" & dt2.Rows(0)("OCID").ToString & ") \r\n"
                                Errmsg += " 已在內網有報名資料，尚在審核中或報名成功，因故無法受理您此次報名需求，請見諒。"
                                Exit For
                            End If
                        End If
                    Next
                End If

                '狀況五-2(已在內網有報名資料)(甄試作業為同一天舉辦)
                'If Errmsg="" Then
                '    If OCIDs1="" Then OCIDs1=Get_Stud_EnterType_OCIDs(tmpIDNO) '已在內網有報名資料
                'End If
            End If
        End If


        If Errmsg <> "" Then
            Errmsg = Replace(Errmsg, "\r\n", vbCrLf)
            Exit Function
        End If
        Check_E_ClsTrace = True
    End Function

    Function Get_Stud_EnterType2_OCIDs(ByVal IDNO As String) As String
        '狀況三(當本次報名課程清單中,與過去已報名課程,為同一培訓單位
        ',且甄試日期為同一天的報名課程,但已報名課程之報名資料被審核失敗,則容許報名本次)
        '(e網審核成功,錄取作業為 0:未選擇, 1:正取, 4:備取者.) 排除：(NOT IN) 2:報名失敗/5:未錄取
        'signUpStatus: 2:報名失敗/5:未錄取
        'signUpStatus-0:收件完成 1:報名成功 2:報名失敗 3:正取(Key_SelResult) 4:備取 5:未錄取
        Dim OCIDs As String = ""
        Dim sql As String
        sql = "" & vbCrLf
        sql &= " SELECT DISTINCT b.OCID1 OCID " & vbCrLf
        sql &= " FROM STUD_ENTERTEMP2 a " & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE2 b on a.eSETID=b.eSETID and b.signUpStatus NOT IN (2,5) " & vbCrLf
        sql &= " WHERE a.IDNO='" & IDNO & "'" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            OCIDs = ""
            For i As Integer = 0 To dt.Rows.Count - 1
                If OCIDs.IndexOf(dt.Rows(i)("OCID").ToString) = -1 Then
                    If OCIDs <> "" Then OCIDs &= ","
                    OCIDs &= dt.Rows(i)("OCID").ToString
                End If
            Next
        End If
        Return OCIDs
    End Function

    Function Get_Stud_EnterType_OCIDs(ByVal IDNO As String) As String
        '狀況五(已在內網有報名資料)

        Dim OCIDs As String = ""
        Dim sql As String
        sql = "" & vbCrLf
        sql &= "  SELECT DISTINCT b.OCID1 OCID " & vbCrLf 'f.Admission 是否錄取 (N:不通過, Y:通過, null:尚未審核、審核中)
        sql &= "  FROM STUD_ENTERTEMP a " & vbCrLf
        sql &= "  JOIN STUD_ENTERTYPE b on a.SETID=b.SETID" & vbCrLf
        sql &= "  LEFT JOIN STUD_SELRESULT f ON b.SETID=f.SETID " & vbCrLf '可能有資料，若做了試算以後。
        sql &= "  WHERE a.IDNO='" & IDNO & "'" & vbCrLf
        sql &= "  AND (f.Admission ='Y' OR f.Admission IS NULL)" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            For i As Integer = 0 To dt.Rows.Count - 1
                If OCIDs.IndexOf(dt.Rows(i)("OCID").ToString) = -1 Then
                    If OCIDs <> "" Then OCIDs &= ","
                    OCIDs &= dt.Rows(i)("OCID").ToString
                End If
            Next
        End If

        Return OCIDs
    End Function

    '新增按鈕
    Private Sub add_but_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles add_but.Click
        Dim aNow As Date = TIMS.GetSysDateNow(objconn)
        Dim sAltMsg As String = "" '訊息
        Dim flag_stopEnterH4 As Boolean = TIMS.StopEnterTempMsgH4(objconn, sAltMsg)
        If flag_stopEnterH4 Then
            Common.MessageBox(Me, sAltMsg)
            Exit Sub
        End If

        'Common.MessageBox(Me, "此學員含有推介單報名資料，請到3合1查詢")
        'Exit Sub
        'If IDNO.Text <> "" Then IDNO.Text=Trim(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text=UCase(IDNO.Text)
        'If IDNO.Text <> "" Then IDNO.Text=TIMS.ChangeIDNO(IDNO.Text)
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text = "" OrElse IDNO.Text.Length <> 10 Then
            Common.MessageBox(Me, "請輸入正確身分證號碼!")
            Exit Sub
        End If

        '1:國民身分證 2:居留證 4:居留證2021
        Dim rqIDNO As String = IDNO.Text
        Dim flag1 As Boolean = TIMS.CheckIDNO(rqIDNO)
        Dim flag2 As Boolean = TIMS.CheckIDNO2(rqIDNO, 2)
        Dim flag4 As Boolean = TIMS.CheckIDNO2(rqIDNO, 4)
        If Not flag1 AndAlso Not flag2 AndAlso Not flag4 Then
            Common.MessageBox(Me, "請輸入正確身分證號碼!")
            Exit Sub
        End If

        Dim stdBLACK2TPLANID As String = ""
        Dim iStdBlackType As Integer = TIMS.Chk_StdBlackType(Me, objconn, stdBLACK2TPLANID)
        Dim flgStdBlack As Boolean = TIMS.Get_StdBlackIDNO1(Me, iStdBlackType, stdBLACK2TPLANID, IDNO.Text, objconn)
        If flgStdBlack Then
            '身分證號被處分了
            '依處分日期及年限，仍在處分期間者，報名登錄，不能新增處分中的報名者。
            Common.MessageBox(Me, "依處分日期及年限，仍在處分期間者，報名登錄，不能新增處分中的報名者。")
            Exit Sub
        End If
        'If oflag_Test Then
        '    Common.MessageBox(Me, "依處分日期及年限，仍在處分期間者，報名登錄，不能新增處分中的報名者。")
        '    Exit Sub
        'End If

#Region "(No Use)"

        ''2015年執行(未決定執行)--這是職前課程邏輯, 忽略
        'Dim sMsg As String=""
        ''判斷功能ID  --SELECT * FROM ID_FUNCTION WHERE FUNID IN (701,70,764)
        'Select Case Convert.ToString(Request("ID"))
        '    Case cst_funid報名登錄 '報名登錄(SD_01_001_add)
        '        '具公司/商業負責人身分 '限定計畫執行
        '        'http://163.29.199.211/Check_ws/Check_ws.asmx
        '        Dim Chkws1 As New Check_ws.Check_ws
        '        'TIMS.Cst_NotTPlanID5
        '        Select Case TIMS.Chk_Master(Me, Chkws1, IDNO.Text)
        '            Case "Y"
        '                Common.MessageBox(Me, cst_xMaster2)
        '                Exit Sub
        '            Case TIMS.cst_Error
        '                Common.MessageBox(Me, cst_msgERR2)
        '                Exit Sub
        '        End Select

        '        '有身分證號, (系統)依身分證號 判斷就保非自願離職者
        '        Select Case TIMS.sUtl_ChkFire(Me, IDNO.Text)
        '            Case "Y"
        '                'sMsg=cst_msg1 ' "該民眾為就保非自願離職者，請通知民眾於該訓練班次報名截止日前，先至公立就業服務機構辦理求職登記，並經適訓評估後，推介參訓。"
        '                Common.MessageBox(Me, cst_msg1)
        '                Exit Sub
        '            Case TIMS.cst_Error
        '                Common.MessageBox(Me, cst_msgERR1)

        '                If Not oflag_Test Then
        '                    Exit Sub
        '                End If

        '        End Select

        '    Case cst_funid專案核定報名登錄 '專案核定報名登錄
        '        'HidMaster.Value=""
        '        '具公司/商業負責人身分 '限定計畫執行
        '        'http://163.29.199.211/Check_ws/Check_ws.asmx
        '        'Dim Chkws1 As New Check_ws.Check_ws
        '        'If TIMS.Chk_Master(Me, Chkws1, IDNO.Text)="Y" Then
        '        '    Common.MessageBox(Me, cst_xMaster3)
        '        '    'Exit Sub '同意繼續報名
        '        '    HidMaster.Value="Y"
        '        'End If
        '        '47 補助辦理照顧服務員職業訓練 '58 補助辦理托育人員職業訓練 20150706 BY AMU
        '        'If TIMS.Cst_TPlanID47AppPlan6.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        'End If
        '    Case cst_funid特例專案核定報名登錄  'EnterPath2@S '764 特例專案核定報名登錄
        '        '取消開訓14天不能報名的限制,排除產投的所有計畫
        'End Select
        ''If TestStr="AmuTest" Then '測試用
        ''    sMsg="該民眾為就保非自願離職者，請通知民眾於該訓練班次報名截止日前，先至公立就業服務機構辦理求職登記，並經適訓評估後，推介免試參訓。"
        ''    Common.MessageBox(Me, sMsg)
        ''    Exit Sub
        ''End If

#End Region

        table5.Visible = False
        DataGrid1.CurrentPageIndex = 0

        Dim url1 As String = ""
        '此功能應該是產投類無法使用。
        '(非產投類計畫) 可插入 推介單 檢核
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產學訓(走e網)
            '產業人才投資方案，暫停使用報名登錄功能
            'Response.Redirect("SD_01_001_1_add.aspx?ID=" & Request("ID") & "&proecess=add&IDNO=" & TIMS.ChangeIDNO(IDNO.Text) & "")
            Common.MessageBox(Me, "該計畫暫停使用報名登錄功能")
            Exit Sub
        End If

#Region "(No Use)"

        'If TIMS.Check_Adp_GOVTRNData(IDNO.Text) Then
        '    '有三合一資料 'by mick (proecess=shift)
        '    Call GetSearchStr()
        '    url1=cst_SD01001_addaspx & "?ID=" & Request("ID") & "&proecess=shift&ticket=3&from_type=add&IDNO=" & TIMS.ChangeIDNO(IDNO.Text) & ""
        '    Call TIMS.Utl_Redirect(Me, objconn, url1)
        'End If

#End Region

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        'Dim sql As String=""
        'sql="" & vbCrLf
        'sql &= " SELECT a.*" & vbCrLf
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT DISTINCT a.SETID " & vbCrLf '/*PK*/
        sql &= " ,a.IDNO " & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,a.ESETID " & vbCrLf
        sql &= " ,a.NAME " & vbCrLf
        sql &= " ,CONVERT(varchar, a.Birthday, 111) Birthday " & vbCrLf
        sql &= " FROM STUD_ENTERTEMP a " & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE b on a.SETID=b.SETID " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND a.IDNO='" & IDNO.Text & "' " & vbCrLf
        sql &= " AND b.RID like '" & RIDValue.Value & "%' " & vbCrLf
        sql &= " AND b.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        Dim oCmd As New SqlCommand(sql, objconn)
        'Call TIMS.OpenDbConn(objconn)
        Dim odt As New DataTable
        odt = DbAccess.GetDataTable(sql, objconn)

        If odt.Rows.Count = 0 Then
            '沒有資料 (XX 或 有1筆資料)
            '產業人才投資方案，暫停使用報名登錄功能
            Call GetSearchStr()
            url1 = String.Concat(cst_SD01001_addaspx, "?ID=", TIMS.Get_MRqID(Me), "&proecess=add&IDNO=", IDNO.Text)
            Call TIMS.Utl_Redirect(Me, objconn, url1)
        End If
        If odt.Rows.Count = 1 Then
            '有1筆資料
            Call GetSearchStr()
            Dim SETID As String = Convert.ToString(odt.Rows(0)("SETID"))
            url1 = String.Concat(cst_SD01001_addaspx, "?ID=", TIMS.Get_MRqID(Me), "&serial=", SETID, "&proecess=add")
            Call TIMS.Utl_Redirect(Me, objconn, url1)
        End If
        If odt.Rows.Count > 1 Then
            '大於2筆，採用list
            table4.Visible = True
            DataGrid1.DataSource = odt
            DataGrid1.DataBind()
        End If
    End Sub

    '查詢 SQL
    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確
        If TxtPageSize.Text <> DataGrid2.PageSize Then DataGrid2.PageSize = TxtPageSize.Text

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        HidOCID1.Value = OCIDValue1.Value
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)

        Dim flagNUCdition1 As Boolean = False
        If IDNO.Text <> "" AndAlso flgROLEIDx0xLIDx0 Then
            '若為 SUPER UESR 且有輸入IDNO 可不判斷 此條件
            flagNUCdition1 = True
        End If

        Dim vs_start_date As String = "" '
        Dim vs_end_date As String = "" 'DateAdd(DateInterval.Day, 1, CDate(end_date.Text))
        Dim vs_transDate1 As String = "" '
        Dim vs_transDate2 As String = "" '

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        transDate1.Text = TIMS.ClearSQM(transDate1.Text)
        transDate2.Text = TIMS.ClearSQM(transDate2.Text)
        If start_date.Text <> "" Then vs_start_date = TIMS.Cdate3(start_date.Text)
        If end_date.Text <> "" Then vs_end_date = TIMS.Cdate3(DateAdd(DateInterval.Day, 1, TIMS.Cdate2(end_date.Text)))
        If transDate1.Text <> "" Then vs_transDate1 = TIMS.Cdate3(transDate1.Text)
        If transDate2.Text <> "" Then vs_transDate2 = TIMS.Cdate3(DateAdd(DateInterval.Day, 1, TIMS.Cdate2(transDate2.Text)))

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.SETID " & vbCrLf
        sql &= " ,a.Name " & vbCrLf
        sql &= " ,a.IDNO " & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,d.OrgName " & vbCrLf
        'sql &= " ,R23.ORGNAME2 " & vbCrLf
        '下列資料暫不停用
        sql &= " ,e.ClassCName ClassCName1" & vbCrLf
        sql &= " ,e.CyclType CyclType1 " & vbCrLf
        'sql &= " ,h.ClassCName ClassCName2, h.CyclType CyclType2 " & vbCrLf
        'sql &= " ,i.ClassCName ClassCName3, i.CyclType CyclType3 " & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(e.CLASSCNAME,e.CYCLTYPE) CLASSCNAME1B" & vbCrLf
        sql &= " ,b.EnterPath " & vbCrLf
        sql &= " ,b.WSID " & vbCrLf
        sql &= " ,vw.StName WSIDName " & vbCrLf
        sql &= " ,FORMAT(b.EnterDate,'yyyy/MM/dd') EnterDate" & vbCrLf
        ' dbo.NVL(b.ENTERPATH2,' ')='P' 專案核定報名
        sql &= " ,CASE WHEN ISNULL(b.ENTERPATH2,' ')='P' then 'Y' end EP2PY" & vbCrLf
        sql &= " ,b.RelEnterDate " & vbCrLf
        sql &= " ,b.EnterChannel " & vbCrLf
        sql &= " ,b.ExamNo " & vbCrLf
        sql &= " ,b.SerNum " & vbCrLf
        sql &= " ,b.OCID1 " & vbCrLf
        sql &= " ,c.Relship " & vbCrLf
        sql &= " ,c.RID " & vbCrLf
        sql &= " ,f.OCID " & vbCrLf
        sql &= " ,f.Admission " & vbCrLf
        'sql += " ,e.OCID OCID1" & vbCrLf
        sql &= " ,ISNULL(f.AppliedStatus,'N') AppliedStatus " & vbCrLf
        sql &= " ,CONVERT(varchar, e.STDate, 111) STDate " & vbCrLf
        sql &= " ,CONVERT(varchar, e.FTDate, 111) FTDate " & vbCrLf
        sql &= " FROM dbo.STUD_ENTERTEMP a " & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTYPE b on b.SETID=a.SETID " & vbCrLf

        sql &= " JOIN dbo.AUTH_RELSHIP c ON b.RID=c.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO d ON c.OrgID=d.OrgID " & vbCrLf
        sql &= " JOIN dbo.CLASS_CLASSINFO e ON b.OCID1=e.OCID " & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN ip on ip.PlanID=e.PlanID " & vbCrLf
        'sql &= " LEFT JOIN dbo.MVIEW_RELSHIP23 R23 on R23.RID3=e.RID " & vbCrLf
        sql &= " LEFT JOIN dbo.V_WORKSTATION vw ON vw.WSID=b.WSID " & vbCrLf
        'sql &= " LEFT JOIN dbo.CLASS_CLASSINFO h ON b.OCID2=h.OCID " & vbCrLf
        'sql &= " LEFT JOIN dbo.CLASS_CLASSINFO i ON b.OCID3=i.OCID " & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SELRESULT f ON b.SETID=f.SETID AND b.EnterDate=f.EnterDate AND b.SerNum=f.SerNum " & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        '准考證號碼區間
        SExamNo.Text = TIMS.ClearSQM(SExamNo.Text)
        FExamNo.Text = TIMS.ClearSQM(FExamNo.Text)
        If SExamNo.Text <> "" Then sql &= " AND b.ExamNo >= '" & SExamNo.Text & "' " & vbCrLf
        If FExamNo.Text <> "" Then sql &= " AND b.ExamNo <= '" & FExamNo.Text & "' " & vbCrLf

        'Select Case rblEnterPathW.SelectedValue
        '    Case "Y" '是 就服單位協助報名
        '        sql &= " and dbo.NVL(b.EnterPath,' ')='" & cst_EnterPathW & "'" & vbCrLf
        '    Case "N" '不是 就服單位協助報名
        '        sql &= " and dbo.NVL(b.EnterPath,' ') != '" & cst_EnterPathW & "'" & vbCrLf
        'End Select
        Select Case rblEnterPathW2.SelectedValue
            Case cst_rw2不區分
            Case cst_rw2一般推介單
                sql &= " AND ISNULL(b.ENTERCHANNEL, 0)=4 " & vbCrLf
                sql &= " AND ISNULL(b.ENTERPATH, ' ') != 'W' " & vbCrLf '排除免試
                sql &= " AND ISNULL(b.ENTERPATH2, ' ') != 'P' " '專案核定
            Case cst_rw2免試推介單
                sql &= " AND ISNULL(b.ENTERPATH, ' ')='W' " & vbCrLf
            Case cst_rw2專案核定報名
                sql &= " AND ISNULL(b.ENTERPATH2, ' ')='P' " & vbCrLf
        End Select
        If vs_start_date <> "" Then sql &= " AND b.RelEnterDate >=" & TIMS.To_date(vs_start_date) & vbCrLf
        If vs_end_date <> "" Then sql &= " AND b.RelEnterDate <=" & TIMS.To_date(vs_end_date) & vbCrLf
        If vs_transDate1 <> "" Then sql &= " AND b.TransDate >=" & TIMS.To_date(vs_transDate1) & vbCrLf
        If vs_transDate2 <> "" Then sql &= " AND b.TransDate <" & TIMS.To_date(vs_transDate2) & vbCrLf
        'OCIDValue1.Value=TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value <> "" Then sql &= " AND (b.OCID1='" & OCIDValue1.Value & "')" & vbCrLf
        '7CE81FD5F9A9D0EDD83FC32835674B7E792D98970DD607F416F7111AC182E62C73B5D6979C58A2B8E6C241BA38408AE31514CAFA59BDED6ADE7C9003927550FF24E6BECF6480201D2D05C8F88250403031D20B38433F9B1F29EE9087761A85192E30CAD3561502BF246A09FFFF5A798A69D1CA5397B7F549D44ED2CE0EDFCFB9702A9A3CD023CB225E775A552816779E6B10BC7C93EBA04769918DDDA41B82180A066575E6EB43897365462F4A48261CE8F3ADCA5BF8E22CEE787E28AD8F453BA19326F2267A5558E485343C809F8558AEE65BD3FBCB5CA9C70A9AEAE5A0CA9BCF7AF12224F507D69FC7BF7536D8CF712D8FF7AD286FC55C7C1CE26B1CDBFDBAA275531C9F7C7714D0AEFAF98107F3F6753727999A6FE8C2F6AB203565AA5B5E922784F15097B2A2B695A5171BDD868FFE16D5D72BE2BADB554474DB911069EFEBE2129BA54DAB8AE2F645902257A18CD17BE29BA67599ED932339BE2D11C80B633BD09FE165CCB35F4FA57AFDFC5409CC9119F22352DAFC50052A7D119D134D731AFF6FEE66B20A9388F09EF750A33934B046BC3101F2E87489B7D24273733F8E98011D70C8BFE817F33D3E9F80BDF77BEB89A97D6ADA3C0AE56402D96DAA50E0438364B8914F8B4ADD4CAF91A2D8214D334601E5832752F31E0B8668BE8FD30F2A8AE040CEBD5A9F262B4BB0B2A244F658DF146638DCC8545A9B815DCDD28A3F80FC6F87D802E56CE97BAB3B5CE94A497E1DEA118B2F700ED88818FC1D8F4E244E32BAB636DF63D506C982F34AE707
        If IDNO.Text <> "" Then sql &= " AND a.IDNO='" & IDNO.Text & "' " & vbCrLf

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Select Case sm.UserInfo.LID
            Case 0
                '限定年度計畫
                If Not flagNUCdition1 Then
                    '若為 SUPER UESR 且有輸入IDNO 可不判斷 此條件
                    sql &= " AND ip.Years='" & sm.UserInfo.Years & "' " & vbCrLf
                    sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "' " & vbCrLf
                    Dim flagNUCdition2 As Boolean = False
                    If IDNO.Text = "" Then flagNUCdition2 = True '(須限定RID)
                    If sm.UserInfo.LID <> "0" Then flagNUCdition2 = True '(須限定RID)
                    If flagNUCdition2 Then
                        '(須限定RID)'沒有選擇特定機構
                        sql &= If(Len(RIDValue.Value) = 1, " AND c.RID LIKE '" & RIDValue.Value & "%'", " AND c.RID ='" & RIDValue.Value & "'")
                    End If
                End If

            Case Else
                '限定年度計畫
                sql &= " AND ip.Years='" & sm.UserInfo.Years & "' " & vbCrLf
                sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
                If Len(sm.UserInfo.RID) = 1 Then
                    '沒有選擇特定機構
                    sql &= If(Len(RIDValue.Value) = 1, " AND c.RID LIKE '" & RIDValue.Value & "%'", " AND c.RID ='" & RIDValue.Value & "'")
                Else
                    sql &= If(RIDValue.Value <> "", " AND c.RID='" & RIDValue.Value & "'", " AND c.RID='" & sm.UserInfo.RID & "'")
                End If
        End Select

        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        '增加通俗職類
        'sql += " where 1=1 and e.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        If cjobValue.Value <> "" Then sql &= " AND e.CJOB_UNKEY=" & cjobValue.Value & vbCrLf

        If stdate1.Text <> "" Then sql &= " AND e.STDATE >= " & TIMS.To_date(stdate1.Text) & vbCrLf

        If stdate2.Text <> "" Then sql &= " AND e.STDATE <= " & TIMS.To_date(stdate2.Text) & vbCrLf

        '依計畫別及RID取得相關資料。
        'dtGetjobc1=TIMS.Getjobc1(Me, RIDValue.Value, objconn)
        Dim dt As DataTable
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
            'Me.ViewState(cst_SearchSqlStr)=sql '存sql語法。
            msg.Text = "查無資料!!"
            table5.Visible = False
            If dt.Rows.Count > 0 Then
                msg.Text = ""
                table5.Visible = True
                If Me.ViewState("sort") = "" Then Me.ViewState("sort") = "RELENTERDATE"
                'PageControler2.SqlString=sql
                'PageControler2.PageDataTable2=dtGetjobc1
                PageControler1.Sort = Me.ViewState("sort")
                PageControler1.PageDataTable = dt
                PageControler1.ControlerLoad()
            End If
        Catch ex As Exception
            '取得錯誤資訊寫入
            Dim strErrmsg As String = $"{TIMS.GetErrorMsg(Me)}{vbCrLf}SQL:{vbCrLf}{sql}{vbCrLf}ex.ToString:{vbCrLf}{ex.ToString}{vbCrLf}"
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            'Throw ex
        End Try
    End Sub

    '查詢鈕
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Search1()  '查詢 SQL
    End Sub

    '查詢三合一資料
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        'TIMS.SUtl_TxtPageSize(Me, Me.TxtPageSize, Me.DataGrid1)
        TxtPageSize.Text = TIMS.ClearSQM(TxtPageSize.Text)
        Dim PageSize As String = TxtPageSize.Text
        Dim s_MRqID As String = TIMS.ClearSQM(Request("ID"))
        GetSearchStr()
        Dim url1 As String = $"SD_01_001_3in1.aspx?PageSize={PageSize}&ID={s_MRqID}"
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '查詢參訓歷史
    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '承訓單位於上方路徑欲利用「查詢歷史紀錄」查詢時，無法顯示「身心障礙者職業重建服務窗口計畫」歷史紀錄
        '，雖於報名完成後可至「首頁>>學員動態管理>>教務管理>>學員參訓歷史」查詢
        '，但有資訊不一致性的問題，故建議在「查詢歷史紀錄」中也連結「身心障礙者職業重建服務窗口計畫」
        '，讓承訓單位更方便查詢。 by AMU 201511103

        Dim IDNOArray As New ArrayList
        For Each eItem As DataGridItem In DataGrid2.Items
            Dim Hid_IDNO_MK As HtmlInputHidden = eItem.FindControl("Hid_IDNO_MK")
            'Hid_IDNO_MK.Value=TIMS.EncryptAes(drv("IDNO")) '加密 Aes
            Dim vIDNO As String = TIMS.DecryptAes(Hid_IDNO_MK.Value) '解密 Aes
            Dim Checkbox1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1")
            If Checkbox1.Checked AndAlso vIDNO <> "" Then IDNOArray.Add(vIDNO)
        Next
        '排序方式
        Session("IDNOArray") = IDNOArray
        If Session("_DataTable") IsNot Nothing Then
            Session("xx_DataTable") = Session("_DataTable")
            Session("_DataTable") = Nothing
        End If

        'Dim ENCIDNO As String=RSA20031.AesEncrypt2(Convert.ToString(drv("IDNO")))
        Dim rqID As String = TIMS.Get_MRqID(Me)
        Dim Script2 As String = $"<script>window.open('../05/SD_05_010_pop.aspx?ID={rqID}&SD_01_004_Type={CST_KD_STUDENTLIST}' ,'history','width=1400,height=820,scrollbars=1')</script>"
        Page.RegisterStartupScript("History", Script2)
    End Sub

    '近兩年參訓資料
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim TwoYears As Integer = 0
        TwoYears = CInt(Year(Now)) - 2 '取得去年的年度

        '排序方式
        Dim IDNOArray As New ArrayList
        For Each eItem As DataGridItem In DataGrid2.Items
            Dim Hid_IDNO_MK As HtmlInputHidden = eItem.FindControl("Hid_IDNO_MK")
            'Hid_IDNO_MK.Value=TIMS.EncryptAes(drv("IDNO")) '加密 Aes
            Dim vIDNO As String = TIMS.DecryptAes(Hid_IDNO_MK.Value) '解密 Aes
            Dim Checkbox1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1")
            If Checkbox1.Checked AndAlso vIDNO <> "" Then IDNOArray.Add(vIDNO)
        Next

        '排序方式
        Session("IDNOArray") = IDNOArray
        If Session("_DataTable") IsNot Nothing Then
            Session("xx_DataTable") = Session("_DataTable")
            Session("_DataTable") = Nothing
        End If

        Dim rqID As String = TIMS.Get_MRqID(Me)
        Dim Script2 As String = $"<script>window.open('../05/SD_05_010_pop.aspx?ID={rqID}&SD_01_004_Type={CST_KD_STUDENTLIST}&TwoYears={TwoYears}','history','width=1400,height=820,scrollbars=1')</script>"
        Page.RegisterStartupScript("History", Script2)
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        table4.Visible = False
        table5.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        table4.Visible = False
        table5.Visible = False

    End Sub

    Sub Del_StudEnterType(ByVal whereCmd As String, ByVal tmpConn As SqlConnection, ByVal tmpTran As SqlTransaction)
        If whereCmd = "" Then Exit Sub
        Dim parms As New Hashtable
        parms.Add("ModifyAcct", sm.UserInfo.UserID)

        Dim sqlStr As String = String.Concat("SELECT 'x' FROM STUD_ENTERTYPE ", whereCmd)
        Dim dtS As DataTable = DbAccess.GetDataTable(sqlStr, tmpTran, parms)
        If dtS.Rows.Count <> 1 Then Exit Sub

        sqlStr = ""
        sqlStr &= " UPDATE STUD_ENTERTYPE " & vbCrLf
        sqlStr &= String.Concat(" SET ModifyAcct=@ModifyAcct, MODIFYDATE=GETDATE() ", whereCmd)
        DbAccess.ExecuteNonQuery(sqlStr, tmpTran, parms)

        sqlStr = " INSERT INTO STUD_ENTERTYPEDELDATA " & vbCrLf
        sqlStr &= String.Concat(" SELECT * FROM STUD_ENTERTYPE ", whereCmd, " AND ModifyAcct=@ModifyAcct")
        DbAccess.ExecuteNonQuery(sqlStr, tmpTran, parms)

        sqlStr = String.Concat(" DELETE STUD_ENTERTYPE ", whereCmd, " AND ModifyAcct=@ModifyAcct")
        DbAccess.ExecuteNonQuery(sqlStr, tmpTran, parms)
    End Sub

    Sub Del_StudSelResult(ByVal whereCmd As String, ByVal tmpConn As SqlConnection, ByVal tmpTran As SqlTransaction)
        If whereCmd = "" Then Exit Sub
        Dim parms As New Hashtable
        parms.Add("ModifyAcct", sm.UserInfo.UserID)

        Dim sqlStr As String = String.Concat("SELECT 'X' FROM STUD_SELRESULT ", whereCmd)
        Dim dtS As DataTable = DbAccess.GetDataTable(sqlStr, tmpTran, parms)
        If dtS.Rows.Count <> 1 Then Exit Sub

        sqlStr = " UPDATE STUD_SELRESULT"
        sqlStr &= String.Concat(" SET ModifyAcct=@ModifyAcct,MODIFYDATE=GETDATE() ", whereCmd) & vbCrLf
        DbAccess.ExecuteNonQuery(sqlStr, tmpTran, parms)

        sqlStr = " INSERT INTO STUD_SELRESULTDELDATA " & vbCrLf
        sqlStr &= String.Concat(" SELECT * FROM STUD_SELRESULT ", whereCmd, " AND ModifyAcct=@ModifyAcct")
        DbAccess.ExecuteNonQuery(sqlStr, tmpTran, parms)

        sqlStr = String.Concat(" DELETE STUD_SELRESULT ", whereCmd, " AND ModifyAcct=@ModifyAcct")
        DbAccess.ExecuteNonQuery(sqlStr, tmpTran, parms)
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    ''' <summary>
    ''' 匯出
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        'Const cst_chk1 As Integer=0
        'Const cst_功能  As Integer=12
        blnExport = True
        'Const Cst_FileName As String="報名登錄查詢匯出"

        Dim objDG As DataGrid
        objDG = DataGrid2
        objDG.AllowPaging = False '關閉分頁
        objDG.EnableViewState = False  '把ViewState給關了
        'objDG.Items.Item(0).sor
        Call Search1()

        'Dim sFileName As String=""
        'sFileName=HttpUtility.UrlEncode(Cst_FileName & ".xls", System.Text.Encoding.UTF8)
        Dim sFileName1 As String = "報名登錄查詢匯出"
        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= ("</style>")

        objDG.AllowPaging = False
        objDG.Columns(cst_chk1).Visible = False '關閉不列印欄位
        objDG.Columns(cst_功能).Visible = False '關閉不列印欄位
        objDG.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div2.RenderControl(objHtmlTextWriter)
        Dim strHTML As String = ""
        'strHTML &= ("<div>")
        strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        'strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        objDG.AllowPaging = True '結束開啟分頁
        objDG.Columns(cst_chk1).Visible = True '不列印欄位
        objDG.Columns(cst_功能).Visible = True '不列印欄位
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'Call TIMS.CloseDbConn(objconn)
        'Response.End()
    End Sub

    Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged
    End Sub

    Protected Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
    End Sub

    '列印匯入學員報名名冊用的班級代碼
    Protected Sub btnPrintOCID1_Click(sender As Object, e As EventArgs) Handles btnPrintOCID1.Click
        Years.Value = TIMS.ClearSQM(CStr(sm.UserInfo.Years))
        If Years.Value = "" Then Exit Sub
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then Exit Sub

        Dim myValue As String = ""
        myValue &= "RID=" & RIDValue.Value
        myValue &= "&Years=" & Years.Value
        myValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub

    '檢查匯入資料
    Function CheckImportData(ByVal colArray As Array) As String
        Dim Reason As String = ""
        'Dim SearchEngStr As String="ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        'Dim sql As String
        'Dim dr As DataRow
        If colArray Is Nothing Then
            Reason = "匯入資料有誤!"
            Return Reason
        End If
        'TIMS.LOG.Debug(String.Format("#colArray.Length :{0}", colArray.Length))
        'TIMS.LOG.Debug(String.Format("#cst_Max_a_Len :{0}", cst_Max_a_Len))
        If colArray.Length <> cst_Max_a_Len Then
            'Reason &= "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Reason &= "欄位對應有誤<BR>"
            Reason &= "請注意欄位中是否有半形逗點<BR>"
            If colArray.Length > cst_i_c_身分證號碼 Then aIDNO = colArray(cst_i_c_身分證號碼)
            If colArray.Length > cst_i_c_姓名 Then aName = colArray(cst_i_c_姓名)
            Return Reason
        End If

        'aSETID="" '流水ID
        aIDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(colArray(cst_i_c_身分證號碼))) '身分證號碼
        aName = colArray(cst_i_c_姓名).ToString            '姓名
        aSex = TIMS.ClearSQM(colArray(cst_i_c_性別)) '性別
        aBirthday = colArray(cst_i_c_出生日期).ToString    '出生日期
        aPassPortNO = colArray(cst_i_c_身份別).ToString    '身分別

        aMaritalStatus = colArray(cst_i_c_婚姻狀況).ToString  '婚姻狀況
        aDegreeID = colArray(cst_i_c_學歷代碼).ToString       '學歷代碼
        aGradID = colArray(cst_i_c_畢業狀況代碼).ToString     '畢業狀況代碼
        aSchool = colArray(cst_i_c_學校名稱).ToString         '學校名稱
        aDepartment = colArray(cst_i_c_科系名稱).ToString     '科系名稱
        aMilitaryID = colArray(cst_i_c_兵役代碼).ToString     '兵役代碼

        aZipCode = TIMS.ClearSQM(colArray(cst_i_c_通訊郵遞區號前3碼)) '郵遞區號前3碼
        aZipCODE6W = TIMS.ClearSQM(colArray(cst_i_c_通訊郵遞區號5或6碼)) '郵遞區號6碼
        aAddress = TIMS.ClearSQM(colArray(cst_i_c_通訊地址)) '通訊地址

        aZIPCODE2 = TIMS.ClearSQM(colArray(cst_i_c_戶籍郵遞區號前3碼)) '戶籍郵遞區號前3碼
        aZIPCODE2_6W = TIMS.ClearSQM(colArray(cst_i_c_戶籍郵遞區號5或6碼)) '戶籍郵遞區號6碼
        aHOUSEHOLDADDRESS = TIMS.ClearSQM(colArray(cst_i_c_戶籍地址)) '戶籍地址   

        aEmail = colArray(cst_i_c_Email).ToString          'Email 
        aPhone1 = colArray(cst_i_c_聯絡電話_日).ToString   '聯絡電話(日)
        aPhone2 = colArray(cst_i_c_聯絡電話_夜).ToString   '聯絡電話(夜)
        aCellPhone = colArray(cst_i_c_行動電話).ToString   '行動電話

        aMIdentityID = colArray(cst_i_c_主要參訓身份別代碼).ToString
        aIdentityID1 = colArray(cst_i_c_參訓身份別代碼1).ToString
        aIdentityID2 = colArray(cst_i_c_參訓身份別代碼2).ToString
        aIdentityID3 = colArray(cst_i_c_參訓身份別代碼3).ToString
        aIdentityID4 = colArray(cst_i_c_參訓身份別代碼4).ToString
        aIdentityID5 = colArray(cst_i_c_參訓身份別代碼5).ToString

        aEnterDate = TIMS.ClearSQM(colArray(cst_i_c_報名日期))  '輸入日期(報名日期)yyyy/MM/dd PK (NOW) '2006/06/08 改成報名日期 RelEnterDate
        aOCID1 = colArray(cst_i_c_報考班別代碼1).ToString       '報考班別代碼1	int 'aOCID2=colArray(23).ToString 'aOCID3=colArray(24).ToString
        aEnterChannel = colArray(cst_i_c_報名管道).ToString   '報考管道
        aIsAgree = colArray(cst_i_c_同意共開資料).ToString        '同意共開資料-同意否

        aUNAME = colArray(cst_i_c_服務單位).ToString '      Const cst_i_c_服務單位 As Integer=31 'SERVDEPT
        aINTAXNO = colArray(cst_i_c_統一編號).ToString '  Const cst_i_c_統一編號 As Integer=32 'ACTNO

        aSERVDEPTID = TIMS.AddZero(colArray(cst_i_c_服務部門), 2) '  Const cst_i_c_服務部門 As Integer=33 'SERVDEPT
        ff = String.Format("SERVDEPTID='{0}'", aSERVDEPTID)
        aSERVDEPT = If(dtSERVDEPT.Select(ff).Length > 0, dtSERVDEPT.Select(ff)(0)("SDNAME"), "")

        aACTNAME = TIMS.ClearSQM(colArray(cst_i_c_投保單位名稱)) '  Const cst_i_c_投保單位名稱 As Integer=34 'SERVDEPT
        aACTTYPE = TIMS.ClearSQM(colArray(cst_i_c_投保類別)) '  Const cst_i_c_投保類別 As Integer=35 ''(1:勞保/2:農保)
        aACTNO = TIMS.ClearSQM(colArray(cst_i_c_投保單位保險證號)) '  Const cst_i_c_投保單位保險證號 As Integer=36
        aACTTEL = TIMS.ClearSQM(colArray(cst_i_c_投保單位電話)) '  Const cst_i_c_投保單位電話 As Integer=37
        aZIPCODE3 = TIMS.ClearSQM(colArray(cst_i_c_投保單位郵遞區號前3碼)) '  Const cst_i_c_投保單位郵遞區號前3碼 As Integer=38
        aZIPCODE3_6W = TIMS.ClearSQM(colArray(cst_i_c_投保單位郵遞區號5或6碼)) '  Const cst_i_c_投保單位郵遞區號5或6碼 As Integer=39
        aACTADDRESS = TIMS.ClearSQM(colArray(cst_i_c_投保單位地址)) '  Const cst_i_c_投保單位地址 As Integer=40

        aJOBTITLEID = TIMS.AddZero(colArray(cst_i_c_職稱), 2) 'Const cst_i_c_職稱 As Integer=41 ''(01:基層員工,02:基層管理者,03:中階管理者,04:高階管理者,05:負責人,99:其他)
        ff = String.Format("JOBTITLEID='{0}'", aJOBTITLEID)
        aJOBTITLE = If(dtJOBTITLE.Select(ff).Length > 0, dtJOBTITLE.Select(ff)(0)("JTNAME"), "")

        aAVTCP1_01 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_01_本署或分署網站))
        aAVTCP1_02 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_02_就業服務中心))
        aAVTCP1_03 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_03_訓練單位))
        aAVTCP1_04 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_04_搜尋網站))
        aAVTCP1_05 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_05_報紙))
        aAVTCP1_06 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_06_廣播))
        aAVTCP1_07 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_07_電視))
        aAVTCP1_08 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_08_朋友介紹))
        aAVTCP1_09 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_09_社群媒體))
        aAVTCP1_99 = TIMS.ClearSQM(colArray(cst_i_c_獲得課程管道_99_其他))
        aAVTCP1_all = ""
        If aAVTCP1_01 = "Y" Then TIMS.Str1V(aAVTCP1_all, "01")
        If aAVTCP1_02 = "Y" Then TIMS.Str1V(aAVTCP1_all, "02")
        If aAVTCP1_03 = "Y" Then TIMS.Str1V(aAVTCP1_all, "03")
        If aAVTCP1_04 = "Y" Then TIMS.Str1V(aAVTCP1_all, "04")
        If aAVTCP1_05 = "Y" Then TIMS.Str1V(aAVTCP1_all, "05")
        If aAVTCP1_06 = "Y" Then TIMS.Str1V(aAVTCP1_all, "06")
        If aAVTCP1_07 = "Y" Then TIMS.Str1V(aAVTCP1_all, "07")
        If aAVTCP1_08 = "Y" Then TIMS.Str1V(aAVTCP1_all, "08")
        If aAVTCP1_09 = "Y" Then TIMS.Str1V(aAVTCP1_all, "09")
        If aAVTCP1_99 = "Y" Then TIMS.Str1V(aAVTCP1_all, "99")

        aNotes = colArray(cst_i_c_備註).ToString '備註
        'aRelEnterDate=colArray(33).ToString

        '身分證驗証
        aIDNO = TIMS.ChangeIDNO(aIDNO)
        If aIDNO = "" Then
            Reason &= "必須填寫身分證號碼<BR>"
        Else
            Select Case aPassPortNO
                Case "2" '身分別為外籍將無驗証功能
                    If Reason = "" Then
                        '驗証匯入檔案時不要有相同的身分證號碼 Start
                        Dim Flag As Boolean = True
                        For i As Integer = 0 To IDNOArray.Count - 1
                            If IDNOArray(i) = aIDNO Then
                                Reason &= "檔案中有相同的身分證號碼<BR>"
                                Flag = False
                            End If
                        Next
                        If Flag Then IDNOArray.Add(aIDNO)
                    End If
                Case "1"
                    If TIMS.CheckIDNO(aIDNO) Then '一般驗証
                        If sm.UserInfo.RoleID = "1" Then '角色代碼為1 可執行安全性規則確認
                            Dim IDNOFlag As Boolean = True
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

                            If Not IDNOFlag Then
                                Reason &= "身分證號碼錯誤!<BR>"
                            End If
                        End If
                        If Reason = "" Then
                            '驗証匯入檔案時不要有相同的身分證號碼 Start 
                            Dim Flag As Boolean = True
                            For i As Integer = 0 To IDNOArray.Count - 1
                                If IDNOArray(i) = aIDNO Then
                                    Reason &= "檔案中有相同的身分證號碼<BR>"
                                    Flag = False
                                End If
                            Next
                            If Flag Then IDNOArray.Add(aIDNO)
                        End If
                    Else
                        Reason &= "身分證號碼錯誤!請聯絡系統管理員<BR>"
                    End If
                Case Else
                    Reason &= "身分別代號只能是1或者是2<BR>"
            End Select

        End If

        If aName = "" Then
            Reason &= "必須填寫中文姓名<BR>"
        Else
            If aName.Length > 30 Then
                Reason &= "中文姓名長度必須小於15<BR>"
            End If
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

        If aPassPortNO <> "" Then
            Select Case aPassPortNO
                Case "1", "2"
                Case Else
                    Reason &= "身分別 代號只能是1或者是2<BR>"
            End Select
        Else
            Reason &= "必須填寫身分別 1:本國 2:外籍<BR>"
        End If

        '1:本國 非'2:外籍
        Select Case aPassPortNO
            Case "2" '2:外籍 
            Case Else
                '1:本國
                If Not ((Mid(aIDNO, 2, 1) = "1" AndAlso aSex = "M") OrElse (Mid(aIDNO, 2, 1) = "2" AndAlso aSex = "F")) Then
                    Reason &= "性別代號與身分證號碼不符<BR>"
                End If
                '1:本國'2:外籍  
                If Reason = "" Then
                    If Not TIMS.CheckMemberSex(aIDNO, aSex) Then
                        'Reason &= "依身分證(或居留證)號判斷 性別選項 不正確<BR>"
                        Reason &= "依身分證號判斷 性別選項 不正確<BR>"
                    End If
                End If
        End Select

        If aBirthday = "" Then
            Reason &= "必須填寫出生日期<BR>"
        Else
            If IsDate(aBirthday) = False Then
                Reason &= "出生日期必須是西元年格式(yyyy/MM/dd)<BR>"
            Else
                If CDate(aBirthday) < "1900/1/1" Or CDate(aBirthday) > "2100/1/1" Then
                    Reason &= "出生日期範圍有誤<BR>"
                End If
            End If
        End If

        If aMaritalStatus = "" Then
            'Reason &= "必須填寫婚姻狀況 1.已;2.未<BR>"
        Else
            Select Case aMaritalStatus
                Case "1", "2", "3"
                Case Else
                    Reason &= "婚姻狀況代號只能是1.2或者是3<BR>"
            End Select
        End If


        '學歷
        If aDegreeID.Length = 1 Then aDegreeID = "0" & aDegreeID '補0
        'CheckIDValeuErr
        Call TIMS.CheckIDValeuErr(aDegreeID, "最高學歷", True, "DegreeID", Key_Degree, Reason)

        '畢業狀況
        If aGradID.Length = 1 Then aGradID = "0" & aGradID '補0
        'CheckIDValeuErr
        Call TIMS.CheckIDValeuErr(aGradID, "畢業狀況", True, "GradID", Key_GradState, Reason)


        If aSchool = "" Then
            Reason &= "必須填寫 學校名稱<BR>"
        Else
            If aSchool.Length > 30 Then
                Reason &= "學校名稱 長度大於儲存範圍(30)<BR>"
            End If
        End If

        If aDepartment = "" Then
            Reason &= "必須填寫科系<BR>"
        Else
            If aDepartment.Length > 128 Then
                Reason &= "填寫科系 長度大於儲存範圍(128)<BR>"
            End If
        End If

        'CheckIDValeuErr
        If aMilitaryID <> "" Then
            If aMilitaryID.Length = 1 Then aMilitaryID = "0" & aMilitaryID '補0
            Call TIMS.CheckIDValeuErr(aMilitaryID, "兵役狀況", True, "MilitaryID", Key_Military, Reason)
        End If
        '必須填寫郵遞區號前3碼 'CheckIDValeuErr
        Call TIMS.CheckIDValeuErr(aZipCode, "通訊郵遞區號填前3碼", True, "ZipCode", ID_ZipCode, Reason)
        '必須填寫郵遞區號 6碼 '郵遞區號 5碼或6碼
        Dim TMPERR1 As String = TIMS.CHK_ZIPCODE6W(aZipCODE6W, "通訊郵遞區號")
        If TMPERR1 <> "" Then Reason &= TMPERR1
        If aAddress = "" Then
            Reason &= "必須填寫通訊地址<BR>"
        Else
            If aAddress.Length > 100 Then Reason &= "通訊地址 長度大於儲存範圍(100)<BR>"
        End If

        '必須填寫郵遞區號前3碼 'CheckIDValeuErr
        Call TIMS.CheckIDValeuErr(aZIPCODE2, "戶籍郵遞區號填前3碼", True, "ZipCode", ID_ZipCode, Reason)
        '必須填寫郵遞區號 6碼 '郵遞區號 5碼或6碼
        TMPERR1 = TIMS.CHK_ZIPCODE6W(aZIPCODE2_6W, "戶籍郵遞區號")
        If TMPERR1 <> "" Then Reason &= TMPERR1
        If aHOUSEHOLDADDRESS = "" Then
            Reason &= "必須填寫戶籍地址<BR>"
        Else
            If aHOUSEHOLDADDRESS.Length > 100 Then Reason &= "戶籍地址 長度大於儲存範圍(100)<BR>"
        End If

        If aPhone1 = "" Then
            Reason &= "必須填寫聯絡電話(日)<BR>"
        Else
            If aPhone1.Length > 25 Then Reason &= "聯絡電話(日)長度大於儲存範圍(25)<BR>"
        End If

        If aPhone2 <> "" AndAlso aPhone2.Length > 25 Then Reason &= "聯絡電話(夜)長度大於儲存範圍(25)<BR>"
        If aCellPhone <> "" AndAlso aCellPhone.Length > 25 Then Reason &= "行動電話 長度大於儲存範圍(25)<BR>"

        aEmail = TIMS.ChangeEmail(TIMS.ClearSQM(aEmail))
        If aEmail = "" Then
            Reason &= "必須填寫Email <BR>"
        Else
            If aEmail.Length > 60 Then
                Reason &= String.Format("Email長度大於儲存範圍(60,{0})<BR>", aEmail.Length)
            ElseIf aEmail <> "無" AndAlso Not TIMS.CheckEmail(aEmail) Then
                Reason &= String.Format("Email格式有誤({0})<BR>", aEmail)
            End If
        End If

        If aMIdentityID.Length = 1 Then aMIdentityID = "0" & aMIdentityID '補0
        If aIdentityID1.Length = 1 Then aIdentityID1 = "0" & aIdentityID1 '補0
        If aIdentityID2.Length = 1 Then aIdentityID2 = "0" & aIdentityID2 '補0
        If aIdentityID3.Length = 1 Then aIdentityID3 = "0" & aIdentityID3 '補0
        If aIdentityID4.Length = 1 Then aIdentityID4 = "0" & aIdentityID3 '補0
        If aIdentityID5.Length = 1 Then aIdentityID5 = "0" & aIdentityID3 '補0

        Call TIMS.CheckIDValeuErr(aMIdentityID, "主要身分別代碼", True, "IdentityID", Key_Identity, Reason)
        Call TIMS.CheckIDValeuErr(aIdentityID1, "身分別代碼1", True, "IdentityID", Key_Identity, Reason)
        Call TIMS.CheckIDValeuErr(aIdentityID2, "身分別代碼2", False, "IdentityID", Key_Identity, Reason)
        Call TIMS.CheckIDValeuErr(aIdentityID3, "身分別代碼3", False, "IdentityID", Key_Identity, Reason)
        Call TIMS.CheckIDValeuErr(aIdentityID4, "身分別代碼4", False, "IdentityID", Key_Identity, Reason)
        Call TIMS.CheckIDValeuErr(aIdentityID5, "身分別代碼5", False, "IdentityID", Key_Identity, Reason)

        If aUNAME = "" Then Reason &= ".必須填寫 服務單位<BR>"
        'If aINTAXNO="" Then Reason &= ".必須填寫 統一編號<BR>"
        If aSERVDEPTID = "" Then
            Reason &= ".必須填寫 服務部門<BR>"
        Else
            If aSERVDEPT = "" Then Reason &= String.Format(".查無 服務部門 代號有誤:{0}<BR>", aSERVDEPTID)
        End If
        If aACTNAME = "" Then Reason &= ".必須填寫 投保單位名稱<BR>"

        If aACTTYPE = "" Then
            Reason &= ".必須填寫 投保類別<BR>"
        Else
            Select Case aACTTYPE
                Case "1", "2"
                Case Else
                    Reason &= String.Format(".查無 投保類別 代號有誤:{0}<BR>", aACTTYPE)
            End Select
        End If

        'aACTNO=colArray(cst_i_c_投保單位保險證號).ToString '  Const cst_i_c_投保單位保險證號 As Integer=36
        'aACTTEL=colArray(cst_i_c_投保單位電話).ToString '  Const cst_i_c_投保單位電話 As Integer=37
        'aZIPCODE3=colArray(cst_i_c_投保單位郵遞區號前3碼).ToString '  Const cst_i_c_投保單位郵遞區號前3碼 As Integer=38
        'aZIPCODE3_6W=colArray(cst_i_c_投保單位郵遞區號5或6碼).ToString '  Const cst_i_c_投保單位郵遞區號5或6碼 As Integer=39
        'aACTADDRESS=colArray(cst_i_c_投保單位地址).ToString '  Const cst_i_c_投保單位地址 As Integer=40

        If aJOBTITLEID = "" Then
            Reason &= ".必須填寫 職稱<BR>"
        Else
            If aJOBTITLE = "" Then Reason &= String.Format(".查無 職稱 代號有誤:{0}<BR>", aJOBTITLEID)
        End If

        If aEnterDate = "" Then
            Reason &= "必須填寫報名日期 <BR>"
        Else
            If IsDate(aEnterDate) = False Then
                Reason &= "填寫報名日期必須是西元年格式(yyyy/MM/dd)<BR>"
            Else
                If CDate(aEnterDate) < "1900/1/1" Or CDate(aEnterDate) > "2100/1/1" Then
                    Reason &= "填寫報名日期範圍有誤<BR>"
                End If
                If Reason = "" Then
                    Dim flagYears1Ok1 As Boolean = False
                    Dim flagYears2Ok2 As Boolean = False
                    Dim iYears1 As Integer = Val(sm.UserInfo.Years) - 1
                    Dim iYears2 As Integer = Val(sm.UserInfo.Years)
                    If iYears1 = Year(aEnterDate) Then flagYears1Ok1 = True
                    If iYears2 = Year(aEnterDate) Then flagYears2Ok2 = True
                    If Not flagYears1Ok1 AndAlso Not flagYears2Ok2 Then
                        Reason &= String.Format("填寫報名日期範圍跨年有誤({0})<BR>", aEnterDate)
                    End If
                End If
                If Reason = "" Then aEnterDate = Common.FormatDate(aEnterDate)
            End If
        End If

        Call TIMS.CheckIDValeuErr(aOCID1, "報考班別代碼1", True, "OCIDN", dt_CLASS_CLASSINFO, Reason)

        If aEnterChannel = "" Then
            Reason &= "必須填寫報名管道 1.網;2.現;3.通;4.推<BR>"
        Else
            Select Case aEnterChannel
                Case "1", "2", "3", "4"
                Case Else
                    Reason &= "報名管道代號只能是 1.網;2.現;3.通;4.推<BR>"
            End Select
        End If

        If aIsAgree = "" Then
            Reason &= "必須填寫是否同意個人基本資料供所屬機關運用 <BR>"
        Else
            Select Case aIsAgree
                Case "Y", "N"
                Case Else
                    Reason &= "是否同意個人基本資料供所屬機關運用代號只能是Y或者是N<BR>"
            End Select
        End If

        'If aTRNDMode <> "" Then'    Select Case aTRNDMode'        Case "1", "2", "3"'        Case Else'            Reason &= "推介種類代號只能是 1.職2.學3.推<BR>"'    End Select'End If
        'If aTRNDType <> "" Then'    Select Case aTRNDType'        Case "1", "2"'        Case Else'            Reason &= "職訓卷種類代號只能是1.甲式2.乙式<BR>"'    End Select'End If
        'If aTicket_NO <> "" Then'    If aTicket_NO.Length <> 18 Then'        Reason &= "職訓券編號長度錯誤 <BR>"'    End If'End If
        'aNotExamID=""
        'aNotExam=""
        'If aNotExam <> "" Then
        '    Select Case aNotExam.ToUpper
        '        Case "NO", "N", "0"
        '            aNotExamID="0"
        '            aNotExam="0"
        '        Case "YES", "Y", "1"
        '            aNotExamID="1"
        '            aNotExam="1"
        '        Case Else
        '            Reason &= "是否免試代號只能是 N:NO Y:YES<BR>"
        '    End Select
        'End If

        If Reason = "" Then
            'Call Check_E_ClsTrace(Reason, aIDNO, aOCID1, aOCID2, aOCID3)
            Call Check_E_ClsTrace(Reason, aIDNO, aOCID1)
            'If Reason <> "" Then
            '    Reason=Replace(Reason, "\r\n", "<br>")
            'End If
        End If

        Return Reason
    End Function

    Sub Utl_IMPORT07()
        'Dim flag_test As Boolean=TIMS.sUtl_ChkTest("testWebForm1")
        Dim Upload_Path As String = "~/SD/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return
        'Const cst_flag As String=","
        'Dim NewSetID, NewSerNum, NewModifyDateS, NewModifyDateE As String
        'Dim NewTMID1, NewTMID2, NewTMID3 As String
        'Dim NewTMID1 As String
        'Dim MyFile As System.IO.File
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        If File1.Value = "" Then
            Common.MessageBox(Me, "檔案位置不可為空 請輸入檔案位置!")
            Exit Sub
        End If
        If File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If

        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> LCase(Cst_Filetype) Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小----------   End

        Const cst_colID_IDNO As String = "身分證號碼" '身分證號碼
        Dim dt_xls As DataTable = Nothing
        Dim Errmag As String = ""

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
        '上傳檔案
        File1.PostedFile.SaveAs(filePath1)
        '取得內容
        dt_xls = TIMS.GetDataTable_XlsFile(filePath1, "", Errmag, cst_colID_IDNO)
        '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        TIMS.MyFileDelete(filePath1)

        If Errmag <> "" Then
            Dim strErrmsg As String = ""
            strErrmsg += "/* Upload_Path & MyFileName: */" & vbCrLf
            strErrmsg += Upload_Path & MyFileName & vbCrLf
            strErrmsg += "/* SD_01_001.Sub Utl_IMPORT07(): */" & vbCrLf
            strErrmsg += Errmag & vbCrLf
            'strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Errmag)
            'Common.MessageBox(Me, Errmag)
            Dim v_msg As String = ""
            If oflag_Test Then v_msg &= Errmag & vbCrLf
            v_msg &= "資料有誤，故無法匯入，請修正Excel檔案，謝謝!"
            Common.MessageBox(Me, v_msg)
            Exit Sub
        End If

        Dim RowIndex As Integer = 1 '讀取行累計數
        'Dim col As String           '欄位
        Dim colArray As Array

        '取出資料庫的所有欄位--------   Start
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        '建立Next_Dlid值 (新值，不會重複)
        'sql="Select distinct max(dlid)+1 Next_Dlid from stud_resultstuddata "
        'Dim Next_Dlid As String=DbAccess.ExecuteScalar(sql)

        'Dim BasicSID As String=TIMS.Get_DateNo
        'Dim SIDNum As Integer=1
        'Dim SID As String
        Dim Reason As String = ""               '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow = Nothing

        '建立錯誤資料格式Table---Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("Name"))
        dtWrong.Columns.Add(New DataColumn("IDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table---End

        '取出所有鍵值當判斷---Start
        'sql="SELECT * FROM Key_Degree"
        sql = "SELECT SERVDEPTID,SDNAME FROM dbo.KEY_SERVDEPT ORDER BY SERVDEPTID "
        dtSERVDEPT = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT JOBTITLEID,JTNAME FROM dbo.KEY_JOBTITLE ORDER BY JOBTITLEID "
        dtJOBTITLE = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT DEGREEID,NAME,DEGREETYPE,SORT FROM KEY_DEGREE WHERE DEGREETYPE IN ('0','1') "
        Key_Degree = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT GRADID,NAME FROM KEY_GRADSTATE "
        Key_GradState = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT MILITARYID,NAME FROM KEY_MILITARY "
        Key_Military = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT IDENTITYID,NAME FROM KEY_IDENTITY "
        Key_Identity = DbAccess.GetDataTable(sql, objconn)
        sql = "SELECT ZIPCODE,ZIPNAME FROM ID_ZIP ORDER BY ZIPCODE"
        ID_ZipCode = DbAccess.GetDataTable(sql, objconn)
        '取出所有鍵值當判斷---End

        '建立課程鍵值--1
        Dim s_AAOCID As String = ""
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            colArray = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
            Dim s_OCID1 As String = "" 'ChkImpDataRstOCID(colArray, Me)
            If colArray.Length <= cst_Max_a_Len Then
                s_OCID1 = TIMS.ClearSQM(colArray(cst_i_c_報考班別代碼1)) '報考班別代碼1 int
                If Not TIMS.IsNumeric2(s_OCID1) Then s_OCID1 = ""
            End If
            Dim drC As DataRow = TIMS.GetOCIDDate(s_OCID1, objconn)
            s_OCID1 = If(drC IsNot Nothing, Convert.ToString(drC("OCID")), "")
            If s_OCID1 <> "" Then
                If s_AAOCID <> "" Then
                    '第2次以後
                    If String.Format(",{0},", s_AAOCID).IndexOf(String.Format(",{0},", s_OCID1)) = -1 Then
                        s_AAOCID &= String.Concat(If(s_AAOCID <> "", ",", ""), s_OCID1)
                    End If
                Else
                    '第1次加入ocid
                    s_AAOCID &= String.Concat(If(s_AAOCID <> "", ",", ""), s_OCID1)
                End If
            End If
        Next

        '建立課程鍵值
        sql = "" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,convert(varchar,a.OCID) OCIDN" & vbCrLf
        sql &= " ,a.ClassCName" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,a.PlanID " & vbCrLf
        'sql += " ,b.ClassID+a.CyclType as NewExam1 " & vbCrLf
        sql &= " ,a.TMID TMID1" & vbCrLf
        sql &= " FROM Class_ClassInfo a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN ID_Class b WITH(NOLOCK) on a.CLSID=b.CLSID" & vbCrLf
        sql &= " WHERE a.IsClosed<>'Y'" & vbCrLf
        If s_AAOCID <> "" Then sql &= " AND a.OCID IN (" & s_AAOCID & ")" & vbCrLf

        dt_CLASS_CLASSINFO = DbAccess.GetDataTable(sql, objconn)
        If dt_CLASS_CLASSINFO.Rows.Count = 0 Then
            Common.MessageBox(Me, "班級的代號 有誤，請確認班級狀態!")
            Exit Sub
        End If
        '建立課程鍵值--1

        For i As Integer = 0 To dt_xls.Rows.Count - 1
            '取得匯入資料
            colArray = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
            '(取得匯入資料) 檢查資料正確性
            Reason = CheckImportData(colArray)

            If Reason = "" Then
                Dim ff3 As String = "OCID='" & TIMS.ChangeSQM(aOCID1) & "'"
                If dt_CLASS_CLASSINFO.Select(ff3).Length = 0 Then
                    Common.MessageBox(Me, "班級的代號 有誤，請確認班級狀態!!")
                    Exit Sub
                End If
                Dim drCC As DataRow = dt_CLASS_CLASSINFO.Select(ff3)(0)
                aTMID1 = drCC("TMID1")
                aRID = drCC("RID")
                aPlanID = drCC("PlanID")

                Dim NewSetID_flag As Boolean = False
                Dim strfield As String = ""
                Dim strErrmsg As String = ""
                'strfield=""
                'strErrmsg=""
                'If aIDNO="B120774236" Then
                '    Dim xxx As String="x"
                '    xxx="x"
                'End If
                Dim iSETID As Integer = 0
                Dim tConn As SqlConnection = DbAccess.GetConnection()
                Dim trans As SqlTransaction = Nothing
                Try
                    trans = DbAccess.BeginTrans(tConn)
                    '學員資料更新／新增
                    Call UPDATE_STUD_ENTERTEMP(NewSetID_flag, trans, iSETID)
                    '學員資料2更新／新增
                    Call UPDATE_STUD_ENTERTEMP2(trans, iSETID)
                    DbAccess.CommitTrans(trans)
                Catch ex As Exception
                    Dim strErrmsg1 As String = ""
                    strErrmsg1 &= "Sub Utl_IMPORT07()" & vbCrLf
                    strErrmsg1 &= TIMS.GetErrorMsg(Me) & vbCrLf '取得錯誤資訊寫入
                    strErrmsg1 &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
                    Call TIMS.WriteTraceLog(strErrmsg1) '錯誤回報'

                    DbAccess.RollbackTrans(trans)
                    TIMS.CloseDbConn(tConn)
                    Throw ex
                End Try

                '當日報名新序號
                Dim iSERNUM As Integer = TIMS.GET_ENTERTYPESERNUM(iSETID, aEnterDate, objconn)
                '准考證編碼
                Dim E_OCID1 As String = aOCID1 '.ToString
                Dim ExamNo1 As String = TIMS.Get_ExamNo1(E_OCID1, objconn) '取出班級的CLASSID +期別 成為准考證編碼的前面的固定碼
                If ExamNo1 = "" OrElse ExamNo1.Length < 6 Then '防呆
                    Common.MessageBox(Me, "班級的代號 與期別有誤，請確認班級狀態")
                    Exit Sub
                End If
                Dim ExamPlanID As String = If(sm.UserInfo.LID = 0, aPlanID, sm.UserInfo.PlanID)
                Dim flgChkExamNo As Boolean = TIMS.Chk_NewExamNOc(ExamPlanID, E_OCID1, objconn)
                If Not flgChkExamNo Then
                    Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)!")
                    Exit Sub
                End If

                '=================== 報名職類檔 更新／新增-Start ===================
                'Call TIMS.OpenDbConn(tConn)
                Try
                    trans = DbAccess.BeginTrans(tConn)
                    'trans=DbAccess.BeginTrans(conn)
                    '取出准考證號   Start 'aNewExamNO'Dim NewExamNO As String 
                    aNewExamNO = TIMS.Get_NewExamNOt(ExamPlanID, ExamNo1, E_OCID1, trans)
                    If aNewExamNO = "" Then
                        DbAccess.RollbackTrans(trans)
                        DbAccess.CloseDbConn(tConn)
                        Common.MessageBox(Me, "班級的代號 與計畫不符，請確認班級狀態(取出准考證號)!!!")
                        Exit Sub
                    End If
                    '取出准考證號   End

                    '報名職類檔 更新／新增-Start
                    Call UPDATE_STUD_ENTERTYPE(trans, iSETID, iSERNUM)
                    '報名職類檔2 更新／新增-Start
                    Call UPDATE_STUD_ENTERTRAIN(trans, iSETID, iSERNUM)
                    DbAccess.CommitTrans(trans)
                Catch ex As Exception
                    Dim strErrmsg1 As String = ""
                    strErrmsg1 &= "Sub Utl_IMPORT07()" & vbCrLf
                    strErrmsg1 &= TIMS.GetErrorMsg(Me) & vbCrLf '取得錯誤資訊寫入
                    'strErrmsg &= "SQL:" & vbCrLf & Sql & vbCrLf
                    strErrmsg1 &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg1)

                    DbAccess.RollbackTrans(trans)
                    TIMS.CloseDbConn(tConn)
                    'Throw ex
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)

                    drWrong("Index") = RowIndex
                    If colArray.Length > 5 Then
                        drWrong("Name") = aName
                        drWrong("IDNO") = TIMS.ChangeIDNO(aIDNO)
                        drWrong("Reason") = ex.Message
                    End If
                End Try
                Call TIMS.CloseDbConn(tConn)


            Else
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)

                drWrong("Index") = RowIndex
                If colArray.Length > 5 Then
                    drWrong("Name") = aName
                    drWrong("IDNO") = TIMS.ChangeIDNO(aIDNO)
                    drWrong("Reason") = Reason
                End If
            End If
            RowIndex += 1 '讀取行累計數

        Next
        '開始判別欄位存入------------   End


        '判斷匯出資料是否有誤
        'Dim explain, explain2 As String
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
            '沒有錯誤資料
            If Reason <> "" Then
                Common.MessageBox(Me, explain & Reason)
                Exit Sub
            End If
            If explain <> "" Then
                Common.MessageBox(Me, explain)
                Exit Sub
            End If
        End If

        Session("MyWrongTable") = dtWrong
        Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('SD_01_001_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
    End Sub

    '匯入名冊
    Protected Sub btnIMPORT07_Click(sender As Object, e As EventArgs) Handles btnIMPORT07.Click
        Call Utl_IMPORT07()
    End Sub

    Function GET_EnterChannel_N(ByVal vEnterChannel As String) As String
        Dim rst As String = ""
        If vEnterChannel = "" Then Return rst
        If vEnterChannel = "1" Then Return "網路"
        If vEnterChannel = "2" Then Return "現場"
        If vEnterChannel = "3" Then Return "通訊"
        If vEnterChannel = "4" Then Return "推介"
        Return rst
    End Function
End Class