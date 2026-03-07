Partial Class SD_03_002_add
    Inherits AuthBasePage

#Region "REM1"
    '28:產業人才投資方案 'Const cst_SD03002_addaspx As String="SD_03_002_add.aspx"
    '06:在職進修訓練/'70:區域產業據點職業訓練計畫(在職) 'Const cst_SD03002_add2aspx As String="SD_03_002_add2.aspx"

    '首頁>>學員動態管理>>學員資料管理>> 學員資料維護 '/SD/03/SD_03_002
    '本作業僅限自辦職前訓練計畫,'參訓身分別
    '於學員資料維護功能，點選修改後，依據屆退官兵荐訓名冊資料，
    '身分勾稽為屆退官兵者(檢核當日檢核"預定退伍日"是否已過開訓日，若未過開訓日，即為屆退官兵者)，
    '於參訓身分別自動勾選"屆退官兵(須單位將級以上長官薦送函)"選項，並於儲存時，
    '檢核參訓身分別是否已勾選"屆退官兵(須單位將級以上長官薦送函)"選項，若未勾選，
    '則顯示告警訊息,'"此訓練學員為屆退官兵，請於參訓身分別勾選"屆退官兵(須單位將級以上長官薦送函)"選項，
    '否則，不能儲存!"，並不予儲存，除非參訓身分別已勾選"屆退官兵(須單位將級以上長官薦送函)"選項,'，才可儲存。
    'SELECT * FROM KEY_IDENTITY WHERE NAME LIKE '%屆退%'--12	 屆退官兵(須單位將級以上長官薦送函)
    'SELECT * FROM PLAN_BUDGET WHERE SYEAR>='2023'
    'select TOP 1 * FROM STUD_TRAINBGQ2 
#End Region

    Const CST_SumOfMoneyMax As Integer = 70000
    Const CST_ToolTipSumOfMoneyMaxMSG1 As String = "紅色為超過7萬之提醒"

    '保險證號(ActNo)前二碼意義：
    '01:工廠'02:公會(工會)'03:漁會'04:政府機關'05:公司'06:農會'07:自由業'08:自由業'09:職訓保
    '02:公會/'03:漁會/'06:農會
    '若是登入年度為 2017年以後，則傳回2，其餘為1
    Dim iPYNum17 As Integer = 1 'iPYNum17=TIMS.sUtl_GetPYNum17(Me)
    Dim flag_show_actno_budid As Boolean = False '保險證號/預算別代碼 false:不顯示 true:顯示

    'ECFA 005 勞動力發展署雲嘉南分署 2024 暫不限定ECFA
    'ECFA 003 勞動力發展署桃竹苗分署 2024 暫不限定ECFA (stop 20241213)
    Dim flag_BudID_ECFA_NoLock As Boolean = False
    'flag_BudID_ECFA_NoLock=((sm.UserInfo.DistID="005" OrElse sm.UserInfo.DistID="003") AndAlso sm.UserInfo.Years="2024")

    Dim gstr_COLUMN_1 As String = ""
    Dim gstr_ROWVAL_1 As String = ""

    Dim tZipLName As String = "" '暫存資訊
    Dim tZipNameN As String = "" '暫存資訊
    'Const cst_inline1 As String="inline"
    Const cst_inline1 As String = ""
    Const cst_none1 As String = "none"

    '該民眾不具失、待業身分，不得參加失業者職前訓練。
    'Dim dtBLIDET1 As DataTable=Nothing
    'dtBLIDET1=TIMS.Get_dtBLIDET1(rqOCID, objconn)

    Dim strScript1 As String = ""
    Const cst_SearchSOCID As String = "SearchSOCID"
    Const cst_Msg1 As String = "此學員現在有職訓生活津貼資料，不能修改 "
    Const cst_Msg1b As String = "(系統管理者開放修改) "
    Const cst_Msg2 As String = "學員資料確定，不可修改"
    'Const cst_Msg2b As String="不提供訓練單位及分署修改"
    Const cst_Msg2c As String = "不提供訓練單位及分署輸入"
    Const cst_Msg2d As String = "委訓單位不可修改，分署可修改"

    Const cst_TITC1 As String = "(預算別移至參訓背景頁籤)"
    'Const cst_Msg14 As String="學員資料維護於訓後14日鎖定" '學員資料維護於訓後14日鎖定
    'Const cst_Msg14ok As String="中心承辦(預算別)於訓後14日開放修改" '學員資料維護於訓後14日鎖定 (產投)
    '28:產業人才投資計劃 54:充電起飛計畫（在職）15:學習券 '學員資料維護於訓後21日鎖定(排除計畫:28.54.15)
    Const cst_limitDay21st As Integer = 21 '委外、中心開訓後可修改期限(開訓 產投、TIMS限定天數)
    Const cst_limitDay21ft As Integer = 21 '委外、中心結訓後可修改期限(產投、TIMS限定天數)
    Const cst_Msg21 As String = "學員資料維護於開訓後21日鎖定" '學員資料維護於訓後21日鎖定(委訓)
    'Const cst_Msg21ok As String="中心承辦(預算別)於訓後21日開放修改" '學員資料維護於訓後21日鎖定 (產投)
    Const cst_Msg21ok As String = "分署承辦(預算別)於訓後21日開放修改" '學員資料維護於訓後21日鎖定 (產投)
    Const cst_Msg30 As String = "學員資料維護於訓後超過30日不能修改" '學員資料維護於訓後30日不能修改
    '委外職前訓練

    '針對委外職前訓練計畫，系統權限 限制 (Page Load/儲存鈕做限制)
    '針對委外職前訓練，系統開放各分署(中心)修改權限為21日，委外訓練單位仍保持開訓後14日後，限制資料修改之邏輯。
    'Const cst_TPlanIDCanEditStud_id37 As String = "37" '委外職前訓練
    '--cst_MsgTPlanID37a="針對委外職前訓練，系統權限 限制各中心修改權限為21日內之後不能修改"
    'Const cst_MsgTPlanID37a As String = "針對委外職前訓練，系統權限 限制各分署修改權限為21日內之後不能修改"
    'Const cst_MsgTPlanID37b As String = "針對委外職前訓練，系統權限 限制委外訓練單位修改權限為21日內之後不能修改"
    Const cst_Msg30x28 As String = "學員資料維護於訓後超過3個月不能修改" '學員資料維護於訓後3個月不能修改 (限中心)

    '協助 -> 公務(ECFA) BudID: SELECT * FROM KEY_BUDGET 
    Const cst_ECFA As String = "公務(ECFA)"
    Const cst_Msg3 As String = "投保單位保險證號為 受ECFA影響之單位,預算別必須為 公務(ECFA)，補助比例必須為特定100%!" '產投
    'BUDID 97
    Const cst_Msg3b As String = "預算別為公務(ECFA)，補助比例須為特定100%!" '產投
    Const cst_Msg4 As String = "預算別為公務(ECFA)，補助比例須為特定100%!" '產投
    Const cst_Msg5 As String = "投保單位保險證號為 非受ECFA影響之單位,預算別不可為公務(ECFA)!" '產投、TIMS
    Const cst_Msg6 As String = "投保單位保險證號為必填資料" '產投、TIMS
    Const cst_Msg7 As String = "該計畫開訓日期為 2011/4/15 日後才可使用 公務(ECFA)基金補助對象!" '產投
    'BUDID 04
    Const cst_Msg8 As String = "預算別為再出發，補助比例須為特定100%!" '產投
    'BUDID 99
    Const cst_Msg9 As String = "預算別為不補助，補助比例大於0，有誤!"  '產投
    'BUDID 02
    Const cst_Msg65 As String = "參訓學員為逾65歲者(滿65歲生日隔天), 其預算別一律運用就安預算!!預算別，(非就安)有誤!"  '產投

    Const cst_str45yearsOld As String = "(此學員為中高齡)" '45歲~65歲
    Const cst_str65yearsOld As String = "此學員為逾65歲者(滿65歲生日隔天)" '"(此學員為65歲(含)以上)" '65歲(含)以上

    Const cst_Msg3c As String = "身分別為一般身分，補助比例 不可為特定100%!" '產投
    Const cst_msgBirth As String = "出生年月日，不開放訓練單位端修正" '產投
    Const cst_Msg4a As String = "主要參訓身分別屬「一般身分」，補助比例應為80%。"
    Const cst_Msg4b As String = "主要參訓身分別屬「特定對象」，補助比例應為100%。"

    '自辦職前訓練，899小時(含)以下，遞補日期為開訓後7日內
    '自辦職前訓練，900小時(含)以上，遞補日期為開訓後12日內
    '非自辦職前訓練:訓練時數120小時(含)以下，遞補日期為開訓後4日內
    '非自辦職前訓練:訓練時數121小時(含)以上，遞補日期為開訓後5日內
    Const cst_master1 As String = "具公司/商業負責人身分，認定為在職者"
    Const cst_workman1 As String = "錄取者被勾稽為在職者，於學員資料維護的就職狀況，直接鎖定為是在職者身份，不可修改。"
    'Const cst_workman2 As String="錄取者被勾稽為在職者(含公會.漁會.農會)，於學員資料維護的就職狀況，直接改為非在職者身份，不可修改。"
    Const cst_workman2 As String = "錄取者被勾稽為在職者(含公會.漁會.農會)，於學員資料維護的就職狀況，直接改為非在職者身份，可修改。"

    'Const cst_errMsg2 As String="資料異常，無法修改生日與姓名，如需修改請將資料提供給中心承辦人。" '產投、TIMS
    Const cst_errMsg2 As String = "資料異常，無法修改生日與姓名，如需修改請將資料提供給分署承辦人。" '產投、TIMS
    Const cst_errMsg2b As String = "(學員基本資料有重複，造成系統無法儲存，請提供相關資料聯繫OJT 窗口。)" '產投、TIMS
    Const cst_errMsg3 As String = "主要參訓身分別，學員資格不符合中高齡者條件！(年齡非介於45歲~65歲之間)"
    Const cst_errMsg4 As String = "參訓身分別，學員資格不符合中高齡者條件！(年齡非介於45歲~65歲之間)"
    Const cst_errMsg5 As String = "津貼身分別，學員資格不符合中高齡者條件！(年齡非介於45歲~65歲之間)"
    Const cst_errMsg6 As String = "主要參訓身分別，學員資格不符合六十五歲以上者條件！(年齡非 65歲以上者)"
    Const cst_errMsg7 As String = "參訓身分別，學員資格不符合六十五歲以上者條件！(年齡非 65歲以上者)"
    Const cst_errMsg8 As String = "津貼身分別，學員資格不符合六十五歲以上者條件！(年齡非 65歲以上者)"
    Const cst_errMsg9 As String = "主要參訓身分別，學員資格不符合屆退官兵身分！"
    Const cst_errMsg10 As String = "報到日期不可晚於學員結訓日期！"
    Const cst_errMsg11 As String = "報到日期不可晚於班級結訓日期！"
    Const cst_errMsg12 As String = "班別代碼有誤，請確認職類/班別！"

    '2011/4/15日'充電起飛計畫公告為 4/15日後才可使用" & cst_ECFA & "基金 
    Const cst_20110415 As String = "2011/04/15"

    'Dim blnCanAdds As Boolean=False
    'Dim blnCanMod As Boolean=False
    'Dim blnCanDel As Boolean=False
    'Dim blnCanSech As Boolean=False
    Dim blnTPlanUseEcfa As Boolean = False '該計畫是否使用ECFA True:使用 False:不使用
    '在職進修 取消必填，學員資料維護 (SD_03_002_add.aspx)
    Dim sTPlan06_G22 As String = ""
    Dim flag_BudIDNoLock As Boolean = False 'flag_BudIDNoLock BudID NoLock'如果是中心承辦人，預算別不鎖定。by AMU 20140328

    'iBudFlag - OUT: 'iBudFlag :0,1,2: 'iFlag :0:未開放 1:21天內修改 2:開放被登功能
    Dim iBudFlag As Integer = 0
    '屆退官兵者 (依系統日期判斷)
    'Dim flagTPlanID02Plan2 As Boolean=False '判斷計畫為自辦職前。

    'Const cst_IdentityCount As Integer=5 '多重身分至多幾項
    Dim gFlagEnv As Boolean = True 'true:正式環境。(false:測試用) / TestStr

    Const Cst_Msg1t As String = "(測試環境開放修改) "

    Const vs_SearchStr As String = "_SearchStr"
    Const vs_HighEduBg As String = "_HighEduBg"
    Const vs_IdentityID As String = "_IdentityID"
    Const vs_STDate As String = "_STDate"

    Sub SUtl_PageInit1()
        Dim sTABLE_NAME As String = "'STUD_STUDENTINFO','STUD_SUBDATA','STUD_SERVICEPLACE','STUD_TRAINBG'"
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS(sTABLE_NAME, objconn) ' DbAccess.GetDataTable(sql)
        If dt.Rows.Count = 0 Then Exit Sub
        Call TIMS.sUtl_SetMaxLen(dt, "IDNO", IDNO)
        Call TIMS.sUtl_SetMaxLen(dt, "NAME", Name)
        Call TIMS.sUtl_SetMaxLen(dt, "RMPNAME", RMPNAME)
        Call TIMS.sUtl_SetMaxLen(dt, "ENGNAME", LName)
        Call TIMS.sUtl_SetMaxLen(dt, "ENGNAME", FName)

        Call TIMS.sUtl_SetMaxLen(dt, "NATIONALITY", Nationality)
        Call TIMS.sUtl_SetMaxLen(dt, "SCHOOL", School)
        Call TIMS.sUtl_SetMaxLen(dt, "DEPARTMENT", Department)
        Call TIMS.sUtl_SetMaxLen(dt, "ADDRESS", Address)
        Call TIMS.sUtl_SetMaxLen(dt, "HOUSEHOLDADDRESS", HouseholdAddress)
        Call TIMS.sUtl_SetMaxLen(dt, "PHONED", PhoneD)
        Call TIMS.sUtl_SetMaxLen(dt, "PHONEN", PhoneN)
        Call TIMS.sUtl_SetMaxLen(dt, "CELLPHONE", CellPhone)
        Call TIMS.sUtl_SetMaxLen(dt, "EMAIL", Email)
        Call TIMS.sUtl_SetMaxLen(dt, "EMERGENCYCONTACT", EmergencyContact)
        Call TIMS.sUtl_SetMaxLen(dt, "EMERGENCYRELATION", EmergencyRelation)
        Call TIMS.sUtl_SetMaxLen(dt, "EMERGENCYPHONE", EmergencyPhone)
        Call TIMS.sUtl_SetMaxLen(dt, "EMERGENCYADDRESS", EmergencyAddress)
        Call TIMS.sUtl_SetMaxLen(dt, "PRIORWORKORG1", PriorWorkOrg1)
        Call TIMS.sUtl_SetMaxLen(dt, "TITLE1", Title1)
        Call TIMS.sUtl_SetMaxLen(dt, "PRIORWORKORG2", PriorWorkOrg2)
        Call TIMS.sUtl_SetMaxLen(dt, "TITLE2", Title2)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVICEID", ServiceID)
        Call TIMS.sUtl_SetMaxLen(dt, "MILITARYAPPOINTMENT", MilitaryAppointment)
        Call TIMS.sUtl_SetMaxLen(dt, "MILITARYRANK", MilitaryRank)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVICEORG", ServiceOrg)
        Call TIMS.sUtl_SetMaxLen(dt, "CHIEFRANKNAME", ChiefRankName)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVICEADDRESS", ServiceAddress)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVICEPHONE", ServicePhone)
        Call TIMS.sUtl_SetMaxLen(dt, "FORENAME", ForeName)
        Call TIMS.sUtl_SetMaxLen(dt, "FORETITLE", ForeTitle)
        Call TIMS.sUtl_SetMaxLen(dt, "FOREIDNO", ForeIDNO)
        Call TIMS.sUtl_SetMaxLen(dt, "FOREADDR", ForeAddr)
        Call TIMS.sUtl_SetMaxLen(dt, "UNAME", Uname)
        Call TIMS.sUtl_SetMaxLen(dt, "INTAXNO", Intaxno)
        Call TIMS.sUtl_SetMaxLen(dt, "ACTNO", ActNo)
        Call TIMS.sUtl_SetMaxLen(dt, "ACTNAME", ActName)
        Call TIMS.sUtl_SetMaxLen(dt, "ACCTNO", AcctNo2)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVDEPT", ServDept)
        Call TIMS.sUtl_SetMaxLen(dt, "JOBTITLE", JobTitle)
        Call TIMS.sUtl_SetMaxLen(dt, "ADDR", Addr)
        Call TIMS.sUtl_SetMaxLen(dt, "TEL", Tel)
        Call TIMS.sUtl_SetMaxLen(dt, "FAX", Fax)
        Call TIMS.sUtl_SetMaxLen(dt, "ACCTHEADNO", AcctheadNo)
        Call TIMS.sUtl_SetMaxLen(dt, "ACCTEXNO", AcctExNo)
        Call TIMS.sUtl_SetMaxLen(dt, "BANKNAME", BankName)
        Call TIMS.sUtl_SetMaxLen(dt, "EXBANKNAME", ExBankName)
        Call TIMS.sUtl_SetMaxLen(dt, "Q3_OTHER", Q3_Other)

        '金融機構代碼查詢 Financial institution code query
        HL_finaCodeQuery.NavigateUrl = TIMS.str_finaCodeQueryUrl
        HL_finaCodeQuery.Target = "_blank"
        HL_finaCodeQuery.ForeColor = Color.Blue
    End Sub

    '補列各資料選項 (排第1順位)(含本班所有學員下拉)
    Sub Add_Items()
        rqOCID = TIMS.ClearSQM(rqOCID)
        If rqOCID = "" Then Exit Sub
        'Dim sqlstr As String '取得該班開訓日期 sql
        'Dim sqlstr2 As String '顯示 03:負擔家計婦女 sql  
        'Dim sqlstr3 As String '顯示 03:負擔家計婦女 判斷

        '增加民族別選項
        'NativeID=TIMS.Get_KeyNative(NativeID)
        DegreeID = TIMS.Get_Degree(DegreeID, 1, objconn)
        GraduateStatus = TIMS.Get_GradState(GraduateStatus, objconn)
        graduatey = TIMS.GetSyear(graduatey, Year(Now) - 110, Year(Now), True) '畢業年份

        '列出兵役下拉選單資料-by Vicient MilitaryID:兵役
        MilitaryID = TIMS.Get_Military(MilitaryID, 2, objconn)
        MilitaryID.Attributes("onchange") = "sol(this.value)"
        MIdentityID.Attributes("onchange") = "MIdentityChg(this.value);ChkMIdentityID();"

        '參訓身分別鍵詞檔2010/08/12 改為用  Plan_Identity table 可依計畫設定不用顯示
        'MIdentityID=TIMS.Get_Identity(MIdentityID, 5, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        MIdentityID = TIMS.Get_Identity(MIdentityID, 52, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        IdentityID = TIMS.Get_Identity(IdentityID, 53, sm.UserInfo.TPlanID, sm.UserInfo.Years, objconn)
        DDL_DISASTER = TIMS.GET_DISASTER(objconn, DDL_DISASTER)
        'Dim xItem1 As ListItem
        'xItem1=IdentityID.Items.FindByValue("04")
        'hide_IdentityID_04.Value=xItem1.Attributes("NAME")
        hide_IdentityID_04.Value = ""
        hide_IdentityID_06.Value = ""
        For i As Integer = 0 To IdentityID.Items.Count - 1
            '04:中高齡者
            If IdentityID.Items.Item(i).Value = "04" Then hide_IdentityID_04.Value = $"IdentityID_{i}"
            '06:身心障礙者
            If IdentityID.Items.Item(i).Value = "06" Then hide_IdentityID_06.Value = $"IdentityID_{i}"
            '都-有值-離開
            If hide_IdentityID_04.Value <> "" AndAlso hide_IdentityID_06.Value <> "" Then Exit For
        Next

        '生活津貼身分別 (SD_03_002_add.aspx)
        SubsidyIdentity = TIMS.Get_SubsidyIdentity(SubsidyIdentity, 1, objconn)
        '參訓身分別
        'IdentityID.Attributes("onclick")="hard(" & IdentityID.ClientID & ")"
        IdentityID.Attributes.Add("onclick", "hard();")
        '身心障礙者
        rblHandType.Attributes.Add("onclick", "hard();")
        ' ------ End
        '津貼類別 
        'SubsidyID=TIMS.Get_SubsidyID(SubsidyID)
        'TIMS.Tooltip(SubsidyID, cst_TITC1)

        '身心障礙者
        HandTypeID = TIMS.Get_HandicatType(HandTypeID)
        HandLevelID = TIMS.Get_HandicatLevel(HandLevelID)
        HandTypeID2 = TIMS.Get_HandicatType2(HandTypeID2)
        HandLevelID2 = TIMS.Get_HandicatLevel2(HandLevelID2)
        'JoblessID=TIMS.Get_JoblessID(JoblessID, Nothing, sm.UserInfo.Years)

        Dim dtDGHR As DataTable = TIMS.Get_DGTHourDT(objconn) '目前系統有4筆資料。
        RelClass_Unit = TIMS.Get_DGTHour(RelClass_Unit, dtDGHR)

        Call TIMS.Get_Trade(Q4)

        '班別學號基本碼 
        StudentIDValue.Value = TIMS.Get_ClassStudentID(Me, rqOCID, objconn)
        If StudentIDValue.Value = "" Then
            Common.MessageBox(Me, "沒有班別學號基本碼")
        End If

        '(兩週內)離退訓 可供遞補
        Call GET_DDL_RejectSOCID(objconn, RejectSOCID, rqOCID)
        '班級學員下拉
        Call GET_DDL_SOCID(objconn, SOCID, rqOCID)

        'Key_Budget PLAN_BUDGET (BudID)
        Dim s_BudIDMsg As String = TIMS.Get_BudIDrbl(sm, BudID, objconn)
        If s_BudIDMsg <> "" Then
            BudIDMsg.Text = s_BudIDMsg '"尚未設定該年度計畫的預算別"
        Else
            '該計畫是否使用ECFA
            If BudID.Items.FindByValue("97") Is Nothing Then
                If blnTPlanUseEcfa Then BudID.Items.Add(New ListItem(cst_ECFA, "97"))
            End If
            If BudID.Items.Count = 1 Then BudID.Items(0).Selected = True '選項為1時選1
        End If

        '津貼類別 'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0% '請選擇變為空。
        SupplyID = TIMS.Get_SupplyID(SupplyID)
        TIMS.Tooltip(SubsidyID, cst_TITC1)
        TIMS.Tooltip(td_wc_2445, cst_TITC1)

        Dim sql As String = ""
        sql = " SELECT SERVDEPTID ,SDNAME FROM KEY_SERVDEPT ORDER BY SERVDEPTID "
        Dim dtSERVDEPT As DataTable = DbAccess.GetDataTable(sql, objconn)
        ddlSERVDEPTID = TIMS.Get_SERVDEPTID(ddlSERVDEPTID, dtSERVDEPT)

        sql = " SELECT JOBTITLEID ,JTNAME FROM KEY_JOBTITLE ORDER BY JOBTITLEID "
        Dim dtJOBTITLE As DataTable = DbAccess.GetDataTable(sql, objconn)
        ddlJOBTITLEID = TIMS.Get_JOBTITLEID(ddlJOBTITLEID, dtJOBTITLE)
    End Sub

    ''' <summary>(兩週內)離退訓 可供遞補</summary>
    ''' <param name="oConn"></param>
    ''' <param name="obj"></param>
    ''' <param name="OCID"></param>
    Public Shared Sub GET_DDL_RejectSOCID(ByRef oConn As SqlConnection, ByRef obj As DropDownList, ByRef OCID As String)
        'objconn, RejectSOCID,  ' (兩週內)離退訓 可供遞補
        Dim pms_s1 As New Hashtable From {{"OCID", TIMS.CINT1(OCID)}}
        Dim sql As String = ""
        sql &= " SELECT a.StudentID ,b.Name+'('+dbo.FN_CSTUDID2(a.StudentID)+')' Name" & vbCrLf
        'sql &= " ,CASE WHEN LEN(a.StudentID)=12 THEN b.Name + '(' + RIGHT(a.StudentID,3) + ')' ELSE b.Name + '(' + RIGHT(a.StudentID,2) + ')' END Name" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.RejectDayIn14" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
        sql &= " WHERE a.OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, pms_s1)
        '(兩週內)離退訓 可供遞補
        dt.DefaultView.RowFilter = "RejectDayIn14='Y'"
        dt.DefaultView.Sort = "StudentID"
        With obj
            .DataSource = dt.DefaultView
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        '=(兩週內)離退訓 可供遞補
    End Sub

    ''' <summary>班級學員下拉</summary>
    ''' <param name="oConn"></param>
    ''' <param name="obj"></param>
    ''' <param name="OCID"></param>
    Public Shared Sub GET_DDL_SOCID(ByRef oConn As SqlConnection, ByRef obj As DropDownList, OCID As String)

        Dim pms_s1 As New Hashtable From {{"OCID", TIMS.CINT1(OCID)}}
        Dim sql As String = ""
        sql &= " SELECT a.StudentID" & vbCrLf
        sql &= " ,b.Name+'('+dbo.FN_CSTUDID2(a.StudentID)+')' Name" & vbCrLf
        'sql &= " ,CASE WHEN LEN(a.StudentID)=12 THEN b.Name + '(' + RIGHT(a.StudentID,3) + ')' ELSE b.Name + '(' + RIGHT(a.StudentID,2) + ')' END Name" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.RejectDayIn14" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
        sql &= " WHERE a.OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, pms_s1)
        'dt.DefaultView.RowFilter="RejectDayIn14 IS NULL"
        dt.DefaultView.RowFilter = ""
        dt.DefaultView.Sort = "StudentID"

        With obj
            .DataSource = dt.DefaultView
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
    End Sub

    '塞入班級資料 (排第2順位) '重要 '新增資料時取得開訓日期 依 Request("OCID")
    Sub GetOpenDate2()
        rqOCID = TIMS.ClearSQM(rqOCID)
        If rqOCID = "" Then Exit Sub
        Dim drCC As DataRow = TIMS.GetOCIDDate(rqOCID, objconn) ' DbAccess.GetOneRow(sql, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, cst_errMsg12)
            Exit Sub
        End If
        Dim pms_s1 As New Hashtable From {{"OCID", TIMS.CINT1(rqOCID)}}
        Dim SQLOP As String = ""
        SQLOP &= " SELECT d.ActNo" & vbCrLf
        SQLOP &= " FROM CLASS_CLASSINFO a" & vbCrLf
        SQLOP &= " JOIN ID_Plan b ON a.PlanID=b.PlanID" & vbCrLf
        SQLOP &= " JOIN Auth_Relship c ON a.RID=c.RID" & vbCrLf
        SQLOP &= " JOIN Org_OrgPlanInfo d ON c.RSID=d.RSID" & vbCrLf
        SQLOP &= $" WHERE a.OCID=@OCID" & vbCrLf
        Dim drOP As DataRow = DbAccess.GetOneRow(SQLOP, objconn, pms_s1)
        If drOP Is Nothing Then
            Common.MessageBox(Me, cst_errMsg12)
            Exit Sub
        End If

        ViewState(vs_STDate) = Common.FormatDate(drCC("STDate"))
        ClassName.Text = $"{drCC("CLASSCNAME2")}" 'TIMS.ClearSQM()
        LevelNo.Items.Clear()
        If $"{drCC("LevelCount")}" <> "" Then
            If Int(drCC("LevelCount")) <> 0 Then
                For i As Integer = 1 To Int(drCC("LevelCount"))
                    LevelNo.Items.Add(New ListItem($"第{i}階段", i))
                Next
                LevelNo.Items.Remove(LevelNo.Items.FindByValue(""))
                LevelNo.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            Else
                LevelNo.Items.Add(New ListItem("無區分階段", 0))
                LevelNo.Enabled = False
                LevelNo.Attributes("title") = "無區分階段"
            End If
        Else
            LevelNo.Items.Add(New ListItem("無區分階段", 0))
            LevelNo.Enabled = False
            LevelNo.Attributes("title") = "無區分階段"
        End If
        If OpenDate.Text = "" Then OpenDate.Text = Common.FormatDate(drCC("STDate"))
        If CloseDate.Text = "" Then CloseDate.Text = Common.FormatDate(drCC("FTDate"))
        'If ActNo.Text="" Then ActNo.Text=TIMS.ChangeIDNO(drCC("ActNo").ToString)
        Hid_OCID.Value = $"{drCC("OCID")}" 'a.OCID
        STDateHidden.Value = Common.FormatDate(drCC("STDate")) 'yyyy/MM/dd
        FTDateHidden.Value = Common.FormatDate(drCC("FTDate")) 'yyyy/MM/dd
        hide_THours.Value = $"{drCC("THours")}"

        If ActNo.Text = "" Then ActNo.Text = TIMS.ChangeIDNO($"{drOP("ActNo")}")

    End Sub

    ''' <summary>
    ''' 學員動態管理>>學員資料管理>>學員資料維護
    ''' 就職狀況自辦職前訓練因參訓者皆為失業狀況，該欄位為必填，預設為失業。
    ''' 在職進修訓練，因參訓者皆須在職狀況，預設為在職。
    ''' </summary>
    ''' <param name="sVal"></param>
    ''' <remarks></remarks>
    Sub SUtl_AutoJobStateType(ByRef MyPage As Page, ByRef oJST As RadioButtonList, ByVal sVal As String)
        'oJST: JobStateType
        Select Case sVal
            Case "0", "1"
                Common.SetListItem(oJST, sVal)
            Case Else
                '0:失業 1:在職 
                Select Case Convert.ToString(sm.UserInfo.TPlanID)
                    Case "02" '02:自辦職前訓練
                        Common.SetListItem(oJST, "0") '0:失業 1:在職 
                    Case Else '"06" '06:在職進修訓練
                        Common.SetListItem(oJST, "1") '0:失業 1:在職 
                End Select
        End Select
    End Sub

    ''' <summary>塞入學員資料 [SQL]</summary>
    ''' <param name="SOCIDStr"></param>
    Sub Create1_Stud(ByVal SOCIDStr As String)
        WSITR.Visible = False
        WSITR2.Visible = False
        'If SOCIDStr="" Then Exit Sub
        SOCIDStr = TIMS.ClearSQM(SOCIDStr)
        If SOCIDStr = "" Then Exit Sub

        'Dim drT2 As DataRow '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
        'Dim dr As DataRow = Nothing 'STUD_STUDENTINFO,STUD_SUBDATA,CLASS_STUDENTSOFCLASS
        Dim parms As New Hashtable() From {{"SOCID", SOCIDStr}}
        'STUD_STUDENTINFO a
        Dim sql As String = ""
        sql &= " SELECT a.SID ,a.IDNO ,a.Name,a.EngName,a.RMPNAME,a.PassPortNO ,a.Sex ,a.Birthday" & vbCrLf
        sql &= " ,a.MaritalStatus ,a.DegreeID ,a.GraduateStatus ,a.MilitaryID ,a.IdentityID ,a.SubsidyID" & vbCrLf
        sql &= " ,a.IsAgree ,a.ChinaOrNot ,a.Nationality ,a.PPNO ,a.JobState ,a.ActNo ,a.GraduateY" & vbCrLf
        'sql &= " ,a.JoblessID ,a.RealJobless" & vbCrLf
        'STUD_SUBDATA b
        sql &= " ,b.School ,b.Department" & vbCrLf
        sql &= " ,b.ZipCode1 ,b.Address" & vbCrLf
        sql &= " ,b.ZipCode2 ,b.HouseholdAddress" & vbCrLf
        sql &= " ,b.ZipCode3 ,b.EmergencyAddress" & vbCrLf
        sql &= " ,b.ZipCode4 ,b.ServiceAddress" & vbCrLf
        sql &= " ,b.Email ,b.PhoneD ,b.PhoneN ,b.CellPhone" & vbCrLf
        sql &= " ,b.EmergencyContact ,b.EmergencyRelation ,b.EmergencyPhone" & vbCrLf
        sql &= " ,b.PriorWorkOrg1 ,b.SOfficeYM1 ,b.FOfficeYM1" & vbCrLf
        sql &= " ,b.PriorWorkOrg2 ,b.SOfficeYM2 ,b.FOfficeYM2" & vbCrLf
        sql &= " ,b.Traffic ,b.ShowDetail ,b.ServiceID ,b.MilitaryAppointment" & vbCrLf
        sql &= " ,b.MilitaryRank ,b.SServiceDate ,b.FServiceDate ,b.ServiceOrg ,b.ChiefRankName" & vbCrLf
        sql &= " ,b.ServicePhone" & vbCrLf
        sql &= " ,b.ForeName ,b.ForeTitle ,b.ForeSex ,b.ForeBirth ,b.ForeIDNO" & vbCrLf
        sql &= " ,b.ForeZip ,b.ForeAddr" & vbCrLf
        sql &= " ,b.ZipCode1_6W,b.ZipCode2_6W,b.ZipCode3_6W,b.ZipCode4_6W,b.ForeZIP6W" & vbCrLf
        sql &= " ,b.ZipCode1_N,b.ZipCode2_N,b.ZipCode3_N,b.ZipCode4_N,b.ForeZip_N" & vbCrLf
        'sql &= " ,b.PriorWorkPay ,b.Title1 ,b.Title2" & vbCrLf
        sql &= " ,b.HandTypeID,b.HandLevelID" & vbCrLf
        sql &= " ,b.HandTypeID2,b.HandLevelID2" & vbCrLf
        'CLASS_STUDENTSOFCLASS c
        sql &= " ,c.SOCID ,c.OCID ,c.StudentID" & vbCrLf
        sql &= " ,c.EnterDate ,c.OpenDate ,c.CloseDate" & vbCrLf
        sql &= " ,c.RejectTDate1 ,c.RejectTDate2 ,c.RTReasonID ,c.StudStatus" & vbCrLf
        sql &= " ,c.TotalResult ,c.BehaviorResult ,c.Rank ,c.IsOnJob" & vbCrLf
        sql &= " ,c.TRNDMode ,c.TRNDType ,c.EnterChannel ,c.BudgetID ,c.SupplyID ,c.LevelNo" & vbCrLf
        sql &= " ,c.GetCertificate ,c.GetSubsidy" & vbCrLf
        sql &= " ,c.RTReasoOther ,c.RelClass_Unit ,c.RelClass_Hour" & vbCrLf
        sql &= " ,c.TrainHours ,c.MIdentityID ,c.Unit1Score ,c.Unit2Score" & vbCrLf
        sql &= " ,c.Unit3Score ,c.Unit4Score ,c.SETID ,c.ETEnterDate ,c.SerNum" & vbCrLf
        sql &= " ,c.PMode ,c.Unit1Hour ,c.Unit2Hour ,c.Unit3Hour ,c.Unit4Hour" & vbCrLf
        sql &= " ,c.CreditPoints ,c.Native ,c.IsApprPaper ,c.AppliedResult" & vbCrLf
        sql &= " ,c.Memo ,c.SubsidyIdentity ,c.MarryLeaveCount ,c.HighEduBg ,c.WkAheadOfSch" & vbCrLf
        sql &= " ,c.JobOrgName ,c.WorkSuppIdent ,c.JobTel" & vbCrLf
        sql &= " ,c.JobZipCode,c.JobZipCODE6W,c.JobZipCode_N" & vbCrLf
        sql &= " ,c.Jobaddress ,c.JobDate ,c.JobSalID ,c.RejectCDate" & vbCrLf
        sql &= " ,c.PWType1 ,c.PWOrg1 ,c.ActNo ActNo2" & vbCrLf
        sql &= " ,c.SOfficeYM1 AS SOfficeYM3" & vbCrLf
        sql &= " ,c.FOfficeYM1 AS FOfficeYM3" & vbCrLf
        sql &= " ,c.IdentityID IdentityIDEX" & vbCrLf
        sql &= " ,c.SubsidyID SubsidyIDEX" & vbCrLf
        sql &= " ,c.MakeSOCID ,c.RejectSOCID" & vbCrLf
        sql &= " ,c.JoblessID ,c.RealJobless" & vbCrLf
        sql &= " ,c.PriorWorkPay ,c.Title1 ,c.Title2,c.ADID" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql &= " JOIN STUD_SUBDATA b ON b.SID=a.SID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS c ON c.SID=a.SID AND c.SID=b.SID" & vbCrLf
        sql &= " WHERE c.SOCID=@SOCID" & vbCrLf
        'Dim dr As DataRow = Nothing 'STUD_STUDENTINFO,STUD_SUBDATA,CLASS_STUDENTSOFCLASS
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr Is Nothing Then
            Dim rqMID As String = TIMS.Get_MRqID(Me)
            Common.MessageBox(Me, "找不到此學員!")
            Page.RegisterStartupScript("SD_03_002", $"<script>location.herf='SD_03_002.aspx?ID={rqMID}';</script>")
            Exit Sub
        End If

        'https://jira.turbotech.com.tw/browse/TIMSC-154
        '參訓學員資料中，若學員資料的通訊地址與戶籍地址若為受災地區範圍，未使用相對應的主要身分別，則需告警
        '若學員資料使用相對應的主要身分別，通訊地址與戶籍地址若非為受災地區範圍，則需告警
        '若使用相對應的主要身分別，通訊地址與戶籍地址為受災地區範圍，則無需告警。
        Common.SetListItem(DDL_DISASTER, dr("ADID"))
        Dim v_DDL_DISASTER As String = TIMS.GetListValue(DDL_DISASTER) 'ADID 重大災害選項
        Dim STDateH1 As String = TIMS.Cdate3(STDateHidden.Value)
        'Dim FTDateH1 As String=TIMS.cdate3(FTDateHidden.Value)
        Dim sZIPCODE1 As String = $"{dr("ZipCode1")}"
        Dim iADID1 As Integer = 0
        Dim sZIPCODE2 As String = $"{dr("ZipCode2")}"
        Dim iADID2 As Integer = 0
        Dim flagMSG1 As Boolean = TIMS.CHK_DIS2MSG(Me, sZIPCODE1, STDateH1, objconn, iADID1)
        Dim flagMSG2 As Boolean = TIMS.CHK_DIS2MSG(Me, sZIPCODE2, STDateH1, objconn, iADID2)
        Dim strMSG3 As String = "" '告警訊息存放
        Select Case $"{dr("MIdentityID")}"
            Case TIMS.cst_Identity_29 '29.重大災害受災者'
                If Not (flagMSG1 OrElse flagMSG2) Then
                    '需告警
                    strMSG3 = "通訊地址與戶籍地址若非為受災地區範圍，主要身分別不可選用「重大災害受災者」!"
                End If
            Case TIMS.cst_Identity_40 '40.重大災害受災者'重大災害選項
                If tr_DDL_DISASTER.Visible AndAlso v_DDL_DISASTER = "" Then  'ADID 重大災害選項
                    strMSG3 = "主要參訓身分別選擇「經公告之重大災害受災者」，須選擇「重大災害選項」不可為空"
                End If
            Case Else
                '29.重大災害受災者,'經公告之重大災害受災者:40
                Dim fg_MI_29 As Boolean = MIdentityID.Items.FindByValue(TIMS.cst_Identity_29) IsNot Nothing
                Dim fg_MI_40 As Boolean = MIdentityID.Items.FindByValue(TIMS.cst_Identity_40) IsNot Nothing
                If fg_MI_29 OrElse fg_MI_40 Then
                    If flagMSG1 OrElse flagMSG2 Then
                        '需告警
                        strMSG3 = "通訊地址與戶籍地址若為受災地區範圍，主要身分別應選用「重大災害受災者」!"
                    End If
                End If
        End Select
        If strMSG3 <> "" Then Common.MessageBox(Me, strMSG3)
        'If Not gFlagEnv Then
        '    strMSG3="通訊地址與戶籍地址若非為受災地區範圍，主要身分別不可選用「重大災害受災者」!"
        '    'strMSG3="通訊地址與戶籍地址若為受災地區範圍，主要身分別應選用「重大災害受災者」!"
        '    Common.MessageBox(Me, strMSG3)
        'End If

        'Dim rOCID As String = $"{dr("OCID")}"
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{dr("OCID")}", objconn) ' DbAccess.GetOneRow(sql, objconn)
        If drCC Is Nothing Then
            Button1.Enabled = False '(儲存1)
            Button2.Enabled = False '(儲存2)
            TIMS.Tooltip(Button1, cst_errMsg12, True)
            TIMS.Tooltip(Button2, cst_errMsg12, True)
            Common.MessageBox(Me, cst_errMsg12)
            Return
        End If

        '檢驗同一身分證號是否有兩筆以上的STUD_STUDENTINFO。
        Dim fg_chkStud As Boolean = TIMS.Check_StudStudentInfo(Convert.ToString(dr("IDNO")), objconn)
        LabErrMsg.Text = ""
        If Not fg_chkStud Then
            'show laberrMsg
            Dim strErrmsg As String = ""
            strErrmsg &= cst_errMsg2b & vbCrLf
            strErrmsg &= $"班級OCID： {dr("OCID")}{vbCrLf}"
            strErrmsg &= $"班級名稱： {ClassName.Text}{vbCrLf}"
            strErrmsg &= $"學員姓名： {dr("Name")}{vbCrLf}"
            strErrmsg &= $"學員出生年月日： {dr("Birthday")}{vbCrLf}"
            strErrmsg &= $"學員身分證號： {dr("IDNO")}{vbCrLf}"
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            LabErrMsg.Text = cst_errMsg2b
            Button1.Enabled = False '(儲存1)
            Button2.Enabled = False '(儲存2)
            TIMS.Tooltip(Button1, cst_errMsg2, True)
            TIMS.Tooltip(Button2, cst_errMsg2, True)
            Common.MessageBox(Me, cst_errMsg2)
        End If

        If rqOCID = "" AndAlso $"{dr("OCID")}" <> "" Then
            rqOCID = $"{dr("OCID")}"
            Dim drCC2 As DataRow = TIMS.GetOCIDDate(rqOCID, objconn) ' DbAccess.GetOneRow(sql, objconn)
            If drCC2 Is Nothing Then
                Button1.Enabled = False '(儲存1)
                Button2.Enabled = False '(儲存2)
                TIMS.Tooltip(Button1, cst_errMsg12, True)
                TIMS.Tooltip(Button2, cst_errMsg12, True)
                Common.MessageBox(Me, cst_errMsg12)
                Return
            End If
            If OpenDate.Text = "" Then OpenDate.Text = Common.FormatDate(drCC2("STDate"))
            If CloseDate.Text = "" Then CloseDate.Text = Common.FormatDate(drCC2("FTDate"))
            'If ActNo.Text="" Then ActNo.Text=TIMS.ChangeIDNO(dr("ActNo").ToString)
            Hid_OCID.Value = $"{drCC2("OCID")}" 'a.OCID
            STDateHidden.Value = Common.FormatDate(drCC2("STDate")) 'yyyy/MM/dd
            FTDateHidden.Value = Common.FormatDate(drCC2("FTDate")) 'yyyy/MM/dd
            hide_THours.Value = $"{drCC2("THours")}"
        End If

        'Dim vTitle As String=""
        'vTitle="授權設定該班級有開放"

        If TIMS.ChkIsEndDate(rqOCID, TIMS.cst_FunID_學員資料維護, dtArc) Then
            '28:產業人才投資計劃 54:充電起飛計畫（在職） '限分署(中心)
            If TIMS.Cst_TPlanID14DayCanEditStud.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID <= 1 Then
                '學員資料維護於訓後3個月不能修改 'Cst_Msg30x28
                If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, 3, CDate(FTDateHidden.Value)), Today) >= 0 Then
                    Button1.Enabled = False '(儲存1)
                    Button2.Enabled = False '(儲存2)
                    TIMS.Tooltip(Button1, cst_Msg30x28, True)
                    TIMS.Tooltip(Button2, cst_Msg30x28, True)
                    'Common.MessageBox(Me, Cst_Msg30x28)
                End If
            Else
                '針對委外職前訓練計畫，系統權限 限制
                Dim flgElseEvent1 As Boolean = True '其它狀況 預設為True
                'If cst_TPlanIDCanEditStud_id37.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '    Select Case sm.UserInfo.LID
                '        Case "1" '分署(中心)
                '            flgElseEvent1 = False '已經設定狀況 其它@False
                '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                '                Button1.Enabled = False '(儲存1)
                '                Button2.Enabled = False '(儲存2)
                '                TIMS.Tooltip(Button1, cst_MsgTPlanID37a, True)
                '                TIMS.Tooltip(Button2, cst_MsgTPlanID37a, True)
                '            End If
                '        Case Else '"2" '委外
                '            flgElseEvent1 = False '已經設定狀況 其它@False
                '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                '                Button1.Enabled = False '(儲存1)
                '                Button2.Enabled = False '(儲存2)
                '                TIMS.Tooltip(Button1, cst_MsgTPlanID37b, True)
                '                TIMS.Tooltip(Button2, cst_MsgTPlanID37b, True)
                '            End If
                '    End Select
                'End If

                '其它狀況 預設為True
                If flgElseEvent1 Then
                    '學員資料維護於訓後30日不能修改'Cst_Msg30
                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 30, CDate(FTDateHidden.Value)), Today) >= 0 Then
                        Button1.Enabled = False '(儲存1)
                        Button2.Enabled = False '(儲存2)
                        TIMS.Tooltip(Button1, cst_Msg30, True)
                        TIMS.Tooltip(Button2, cst_Msg30, True)
                        'Turbo.Common.MessageBox(Me, Cst_Msg30)
                    End If
                End If
            End If
        End If

        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            Button1.Enabled = True '(儲存1)
            Button2.Enabled = True '(儲存2)
            TIMS.Tooltip(Button1, cst_Msg1b)
            TIMS.Tooltip(Button2, cst_Msg1b)
        End If

        If Not gFlagEnv Then '正式環境。(測試用) / TestStr
            If Button1.Enabled = False Then '測試用。
                Button1.Enabled = True '(儲存1)
                TIMS.Tooltip(Button1, Cst_Msg1t)
            End If
            If Button2.Enabled = False Then '測試用。
                Button2.Enabled = True '(儲存1)
                TIMS.Tooltip(Button2, Cst_Msg1t)
            End If
        End If

        labmakesocid.Text = ""
        hide_MakeSOCID.Value = ""
        If $"{dr("MakeSOCID")}" <> "" Then
            hide_MakeSOCID.Value = $"{dr("MakeSOCID")}"
            labmakesocid.Text = $"遞補學員：{TIMS.GetSOCIDName($"{dr("MakeSOCID")}", objconn)}"
        End If

        '遞補學員(被遞補學員)
        'RejectSOCID.Enabled=True
        hide_RejectSOCID.Value = ""
        If $"{dr("RejectSOCID")}" <> "" Then
            hide_RejectSOCID.Value = $"{dr("RejectSOCID")}"
            Common.SetListItem(RejectSOCID, $"{dr("RejectSOCID")}")
        End If

        ''TIMS 計畫 (非產投)
        'If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '於開訓日期後，依據以下職前訓練計畫規定鎖住「遞補者」選項
        '    Call Check_RejectSOCID_Enabled(STDateHidden.Value, hide_THours.Value)
        'End If
        '遞補者 先停掉此規則。BY AMU 20150505

        '專上畢業學歷失業者
        Dim v_rdo_HighEduBg As String = "N"
        If Convert.ToString(dr("HighEduBg")) = "Y" Then v_rdo_HighEduBg = "Y"
        Common.SetListItem(rdo_HighEduBg, v_rdo_HighEduBg)

        '是否為在職者補助身分 '46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        If $"{dr("WorkSuppIdent")}" <> "" Then Common.SetListItem(rblWorkSuppIdent, $"{dr("WorkSuppIdent")}")

        'HidMaster.Value=""
        'If IDNO.Text <> "" Then
        '    '具公司/商業負責人身分 '限定計畫執行
        '    'http://163.29.199.211/Check_ws/Check_ws.asmx
        '    Dim Chkws1 As New Check_ws.Check_ws
        '    If TIMS.Chk_Master(Me, Chkws1, IDNO.Text)="Y" Then
        '        'Common.MessageBox(Me, cst_xMaster3)
        '        'Exit Sub '同意繼續報名
        '        HidMaster.Value="Y"
        '    End If
        'End If

        HidMaster.Value = TIMS.Chk_MasterEnter(dr("IDNO"), dr("OCID"), objconn)
        If HidMaster.Value = "Y" Then
            '具公司/商業負責人身分 '限定計畫執行 '鎖定為在職者
            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                Common.SetListItem(rblWorkSuppIdent, "Y")
                rblWorkSuppIdent.Enabled = False
                TIMS.Tooltip(rblWorkSuppIdent, cst_master1)
            End If
        End If

        Dim ss As String = "IDNO='" & dr("IDNO") & "'"
        '為在職者補助身份
        '該民眾不具失、待業身分，不得參加失業者職前訓練。'限定計畫執行
        'https://jira.turbotech.com.tw/browse/TIMSB-1247
        '僅涉托育人員及照顧服務員2支計畫 而非所有職前計畫
        LabWSImsg.Text = ""
        'Dim flag_WSI As Boolean=False '(非)勾稽為在職者
        'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If dtBLIDET1 IsNot Nothing Then
        '        If dtBLIDET1.Select(ss).Length > 0 Then
        '            '是否為在職者補助身分
        '            flag_WSI=True '勾稽為在職者
        '            Common.SetListItem(rblWorkSuppIdent, TIMS.cst_YES)
        '            rblWorkSuppIdent.Enabled=False '鎖定
        '            TIMS.Tooltip(rblWorkSuppIdent, cst_workman1)
        '        End If

        '        'NGACTNO:02:公會 x '03:漁會 x '06:農會 x (排除保險證號)
        '        ss="IDNO='" & dr("IDNO") & "' AND NGACTNO='Y'"
        '        If Not rblWorkSuppIdent.Enabled _
        '            AndAlso dtBLIDET1.Select(ss).Length > 0 Then
        '            flag_WSI=False '勾稽為非在職者
        '            rblWorkSuppIdent.Enabled=True '解鎖
        '            'rblWorkSuppIdent.Enabled=False '鎖定
        '            Common.SetListItem(rblWorkSuppIdent, TIMS.cst_NO)
        '            TIMS.Tooltip(rblWorkSuppIdent, cst_workman2, True)
        '            'LabWSImsg.Text=cst_workman2
        '        End If

        '        '直接更動為在職者 (顯示時修正資料)
        '        'Call UPDATE_WorkSuppIdent(SOCIDStr, rOCID, flag_WSI, objconn)

        '        '勾稽為在職者
        '        If flag_WSI Then
        '            Select Case Convert.ToString(dr("WorkSuppIdent"))
        '                Case TIMS.cst_YES
        '                    'LabWSImsg.Text &= "目前是在職者"
        '                Case Else
        '                    LabWSImsg.Text="資料庫是「非在職者」"
        '            End Select
        '        End If
        '    End If
        'End If

        '是否為在職者補助身分
        'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        '    Select Case Convert.ToString(dr("WorkSuppIdent"))
        '        Case "Y"
        '            Page.RegisterStartupScript("Change_MBTable1", "<script>Change_MBTable(1);</script>")
        '        Case "N"
        '            Page.RegisterStartupScript("Change_MBTable2", "<script>Change_MBTable(2);</script>")
        '        Case Else
        '            Page.RegisterStartupScript("Change_MBTable2", "<script>Change_MBTable(2);</script>")
        '    End Select
        'End If

        Call GetHistorySumOfMoney(SOCIDStr)
        Call SHOW_IMG12(Convert.ToString(dr("IDNO")))

        Common.SetListItem(SOCID, dr("SOCID").ToString)
        Common.SetListItem(LevelNo, dr("LevelNo").ToString)

        If Len(Convert.ToString(dr("StudentID"))) = 12 Then
            StudentID.Text = Right(Convert.ToString(dr("StudentID")), 3)
            StudentIDstring.Value = Right(Convert.ToString(dr("StudentID")), 3)
        Else
            StudentID.Text = Right(Convert.ToString(dr("StudentID")), 2)
            StudentIDstring.Value = Right(Convert.ToString(dr("StudentID")), 2)
        End If

        Name.Text = Convert.ToString(dr("Name"))
        'OJT-24102401：<系統> 產投、充飛 - 學員資料維護：調整欄位位置與相關輸入卡控
        If sm.UserInfo.LID = 2 Then TIMS.INPUT_ReadOnly(Name, cst_Msg2d) 'cst_Msg2b) 

        RMPNAME.Text = Convert.ToString(dr("RMPNAME"))
        If Not fg_chkStud Then
            '檢驗同一身分證號是否有兩筆以上的STUD_STUDENTINFO。
            Name.ReadOnly = True
            TIMS.Tooltip(Name, "檢驗同一身分證號有兩筆以上，請洽系統管理者!!")
        End If

        If Convert.ToString(dr("EngName")) <> "" Then
            If Split(dr("EngName"), " ", , CompareMethod.Text).Length = 1 Then
                LName.Text = dr("EngName").ToString
            Else
                If (dr("EngName").ToString.IndexOf(" ")) > -1 Then
                    LName.Text = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
                    FName.Text = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - 1 - dr("EngName").ToString.IndexOf(" ")))
                ElseIf (dr("EngName").ToString.IndexOf("　")) > -1 Then
                    LName.Text = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf("　")))
                    FName.Text = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - 1 - dr("EngName").ToString.IndexOf("　")))
                Else
                    LName.Text = ""
                    FName.Text = Trim(dr("EngName").ToString)
                End If
            End If
        End If

        Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)

        If dr("PassPortNO").ToString = "1" Then
            ChinaOrNotTable.Style("display") = cst_none1
            PPNO.Style("display") = cst_none1
            ForeTr1.Style("display") = cst_none1
            ForeTr2.Style("display") = cst_none1
            ForeTr3.Style("display") = cst_none1
            ForeTr4.Style("display") = cst_none1
            ForeTr5.Style("display") = cst_none1
            ChinaOrNot.SelectedIndex = -1
            Nationality.Text = ""
            PPNO.SelectedIndex = -1
        Else
            ChinaOrNotTable.Style("display") = cst_inline1
            PPNO.Style("display") = cst_inline1
            ForeTr1.Style("display") = cst_inline1
            ForeTr2.Style("display") = cst_inline1
            ForeTr3.Style("display") = cst_inline1
            ForeTr4.Style("display") = cst_inline1
            ForeTr5.Style("display") = cst_inline1
            Common.SetListItem(ChinaOrNot, dr("ChinaOrNot").ToString)
            Nationality.Text = dr("Nationality").ToString
            Common.SetListItem(PPNO, dr("PPNO").ToString)
            ForeName.Text = dr("ForeName").ToString
            ForeTitle.Text = dr("ForeTitle").ToString
            Common.SetListItem(ForeSex, dr("ForeSex").ToString)
            If IsDate(dr("ForeBirth")) Then ForeBirth.Text = Common.FormatDate(dr("ForeBirth"))
            ForeIDNO.Text = TIMS.ChangeIDNO(dr("ForeIDNO").ToString)

            'ForeZip  戶籍地址(國內親屬資料)
            ForeZip.Value = Convert.ToString(dr("ForeZip"))
            hidForeZIP6W.Value = Convert.ToString(dr("ForeZIP6W"))
            ForeZIPB3.Value = TIMS.GetZIPCODEB3(hidForeZIP6W.Value)
            ForeZip_N.Value = Convert.ToString(dr("ForeZip_N"))
            City6.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ForeZip")), Convert.ToString(dr("FOREZIP_N")), objconn)
            ForeAddr.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ForeAddr")))
            Hid_JnForeZip.Value = TIMS.GetZipCodeJn(ForeZip.Value, ForeZIPB3.Value, hidForeZIP6W.Value, City6.Text, ForeAddr.Text)
        End If

        IDNO.ReadOnly = True
        IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
        TIMS.Tooltip(IDNO, "身分證號碼修改，請洽系統管理者!!")
        '------for 個資保護------
        txtShowIDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)

        Common.SetListItem(Sex, dr("Sex").ToString)

        If Convert.ToString(dr("Birthday")) <> "" Then
            Birthday.Text = Common.FormatDate(dr("Birthday"))
            'Birthday.Text=Common.FormatDate(Convert.ToString(dr("Birthday")), DateFormat.ShortDate)
            'Me.ViewState("Birthday")=Common.FormatDate(dr("Birthday"))
            '------for 個資保護------
            'txtShowBirthday.Text=Common.FormatDate(dr("Birthday"))
        End If

        '有兩筆以上的STUD_STUDENTINFO,身分證號唯讀。
        If Not fg_chkStud Then
            'Birthday.ReadOnly=True
            Birthday.Attributes.Add("onkeydown", "this.blur()")
            Birthday.Attributes.Add("oncontextmenu", "return false;")
            Birthday.Enabled = False
            Img1.Style("display") = cst_none1 '出生日期選擇功能
            Img1.Disabled = True
            TIMS.Tooltip(Birthday, "檢驗同一身分證號有兩筆以上，請洽系統管理者!!")
        End If

        Select Case Convert.ToString(dr("MaritalStatus"))
            Case "1", "2"
                Common.SetListItem(MaritalStatus, dr("MaritalStatus").ToString)
            Case Else
                Common.SetListItem(MaritalStatus, "3")
        End Select
        Common.SetListItem(EnterChannel, Convert.ToString(dr("EnterChannel")))

        '0.沒有有現場報名'1.有現場報名
        hide_EnterChannel2.Value = "0"
        '1.網;2.現;3.通;4.推
        hide_EnterChannel.Value = Convert.ToString(dr("EnterChannel"))  '將原本的報名管道放在hide給script備用

        '有現場報名
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        'Dim rOCID As String=Convert.ToString(dr("OCID"))
        Dim SingUp As Boolean = TIMS.CheckIfSingUp(IDNO.Text, rqOCID, 0, objconn)
        If SingUp Then
            '0.沒有有現場報名'1.有現場報名
            hide_EnterChannel2.Value = "1"
        End If
        Page.RegisterStartupScript("EnterChannelChange", "<script>EnterChannelChange();</script>")
        Common.SetListItem(TRNDMode, dr("TRNDMode").ToString)

        '就職狀況'0:失業 1:在職  (JobStateType)
        Call SUtl_AutoJobStateType(Me, JobStateType, Convert.ToString(dr("JobState")))

        If Convert.ToString(dr("OpenDate")) <> "" Then OpenDate.Text = Common.FormatDate(Convert.ToString(dr("OpenDate")))
        If Convert.ToString(dr("CloseDate")) <> "" Then CloseDate.Text = Common.FormatDate(Convert.ToString(dr("CloseDate")))
        If Convert.ToString(dr("EnterDate")) <> "" Then EnterDate.Text = Common.FormatDate(Convert.ToString(dr("EnterDate")))

        'Common.SetListItem(DegreeID, dr("DegreeID").ToString)
        Dim DegreeIDValTmp As String = ""
        '修正學歷代碼 (學員資料維護)
        DegreeIDValTmp = TIMS.Fix_DegreeValue(Convert.ToString(dr("DegreeID")))
        Common.SetListItem(DegreeID, DegreeIDValTmp)

        Dim v_BudID As String = ""
        Dim v_SupplyID As String = ""
        Dim s_ActNo2 As String = Convert.ToString(dr("ActNo2")) 'CLASS_STUDENTSOFCLASS
        Dim s_ActNo As String = Convert.ToString(dr("ActNo")) 'STUD_STUDENTINFO
        If (s_ActNo2 <> "") Then s_ActNo = s_ActNo2 'CLASS_STUDENTSOFCLASS

        'BudID (預算別) 
        If Hid_show_actno_budid.Value = "Y" Then
            '產投 設定預算別 create
            Dim v_def_BudgetID As String = ""
            If Convert.ToString(dr("BudgetID")) <> "" Then
                v_def_BudgetID = Convert.ToString(dr("BudgetID"))
            Else
                'BudID (預算別) by AMU 20080602
                '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
                '01、04、05、15、08 其補助經費來源歸屬為 03:就保基金
                '02、03、06、07 其經費來源歸屬為 02:就安基金
                '09與無法辨視者為 99:不予補助對象
                '2.開頭數字為075、175（裁減續保）、076、176（職災續保）、09（訓）皆為不予補助對象，並設定阻擋。
                Select Case Left(s_ActNo, 2) 'CLASS_STUDENTSOFCLASS
                    Case "01", "04", "05", "15", "08"
                        v_def_BudgetID = "03"
                    Case "02", "03", "06", "07"
                        v_def_BudgetID = "02"
                    Case "09"
                        v_def_BudgetID = "99"
                    Case Else
                        v_def_BudgetID = "99"
                End Select
                Select Case Left(s_ActNo, 3) 'CLASS_STUDENTSOFCLASS
                    Case "075", "175", "076", "176"
                        v_def_BudgetID = "99"
                End Select

                HidMaster.Value = TIMS.Chk_MasterEnter(dr("IDNO"), dr("OCID"), objconn)
                If HidMaster.Value = "Y" Then
                    '具公司/商業負責人身分 '限定計畫執行 '201509 
                    '若勾稽出為負責人的學員，「預算別」欄位直接預設為「就安」。
                    'Common.SetListItem(BudID, "02")
                    v_def_BudgetID = "02"
                End If
                '1.在【學員資料維護】，若該學員於於「公法救助」是屬於「M：多元就業計畫進用人員不適用就保」，系統預算別要預設帶「就安」! (圖4)。
                'Hid_BIEF.Value=TIMS.GET_BLIGATEDATA28(Convert.ToString(dr("SOCID")), Convert.ToString(dr("IDNO")), objconn, "BIEF")
                'If Hid_BIEF.Value="M" Then v_def_BudgetID="02"
                '1.在【學員資料維護】，若該學員於於「公法救助」是屬於「M：多元就業計畫進用人員不適用就保」，系統預算別要預設帶「就安」! (圖4)。
                Hid_BIEF.Value = TIMS.GET_BLIGATEDATA28E(Convert.ToString(dr("IDNO")), Convert.ToString(dr("OCID")), objconn, "BIEF")
                If Hid_BIEF.Value = "M" Then v_def_BudgetID = "02"
            End If
            Common.SetListItem(BudID, v_def_BudgetID)

            ''產投 2009年 身分別為「非自願離職者」時
            ''1.預算來源應為 02:就保基金 ； 2.補助比例為100%
            'If CInt(sm.UserInfo.Years) > 2008 Then
            '    If Convert.ToString(dr("BudgetID")) <> "" Then
            '        Common.SetListItem(BudID, dr("BudgetID").ToString)
            '        '直接抓前頭來源資料不另預設帶值
            '    End If
            'End If

            'SupplyID 的 預設選擇 與值的填入
            'Dim v_BudID As String=TIMS.GetListValue(BudID)
            'Dim v_SupplyID As String=TIMS.GetListValue(SupplyID)
            v_BudID = TIMS.GetListValue(BudID)
            v_SupplyID = TIMS.GetListValue(SupplyID)
            Select Case v_BudID'BudID.SelectedValue
                Case "99"
                    Common.SetListItem(SupplyID, "9") '不補助。
                Case Else
                    '產投 2009年 身分別為「非自願離職者」時
                    '1.預算來源應為 02:就保基金 ； 2.補助比例為100%
                    'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
                    If CInt(sm.UserInfo.Years) > 2008 Then
                        If Convert.ToString(dr("IdentityIDEX")) <> "" Then
                            If Convert.ToString(dr("IdentityIDEX")).IndexOf("02") > -1 Then Common.SetListItem(SupplyID, "2") '特定對象 100%"
                        End If
                        If v_SupplyID = "" OrElse v_SupplyID = "0" Then
                            If Convert.ToString(dr("SupplyID")) <> "" Then Common.SetListItem(SupplyID, dr("SupplyID").ToString)
                        End If
                    Else
                        'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
                        If Convert.ToString(dr("SupplyID")) <> "" Then Common.SetListItem(SupplyID, dr("SupplyID").ToString)
                    End If
            End Select
        End If

        School.Text = Convert.ToString(dr("School"))
        Department.Text = Convert.ToString(dr("Department"))
        If Convert.ToString(dr("GraduateStatus")) <> "" Then Common.SetListItem(GraduateStatus, dr("GraduateStatus"))
        If Convert.ToString(dr("GraduateY")) <> "" Then Common.SetListItem(graduatey, dr("GraduateY"))

        SolTR.Style.Item("display") = cst_none1
        'MilitaryID.SelectedIndex=-1
        If Convert.ToString(dr("MilitaryID")) <> "" Then
            Common.SetListItem(MilitaryID, dr("MilitaryID").ToString)
            If Convert.ToString(dr("MilitaryID")) = "04" Then SolTR.Style.Item("display") = cst_inline1
        End If
        ServiceID.Text = Convert.ToString(dr("ServiceID"))
        MilitaryAppointment.Text = Convert.ToString(dr("MilitaryAppointment"))
        MilitaryRank.Text = Convert.ToString(dr("MilitaryRank"))
        ServiceOrg.Text = Convert.ToString(dr("ServiceOrg"))
        ChiefRankName.Text = Convert.ToString(dr("ChiefRankName"))
        ServicePhone.Text = Convert.ToString(dr("ServicePhone"))

        If Convert.ToString(dr("SServiceDate")) <> "" Then SServiceDate.Text = Common.FormatDate(Convert.ToString(dr("SServiceDate")))
        If Convert.ToString(dr("FServiceDate")) <> "" Then FServiceDate.Text = Common.FormatDate(Convert.ToString(dr("FServiceDate")))

        'ZipCode4 服役單位地址
        ZipCode4.Value = Convert.ToString(dr("ZipCode4"))
        hidZipCode4_6W.Value = Convert.ToString(dr("ZipCode4_6W"))
        ZipCode4_B3.Value = TIMS.GetZIPCODEB3(hidZipCode4_6W.Value)
        ZipCode4_N.Value = Convert.ToString(dr("ZipCode4_N"))
        City4.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode4")), Convert.ToString(dr("ZipCode4_N")), objconn)
        ServiceAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ServiceAddress")))
        Hid_JnZipCode4.Value = TIMS.GetZipCodeJn(ZipCode4.Value, ZipCode4_B3.Value, hidZipCode4_6W.Value, City4.Text, ServiceAddress.Text)

        PhoneD.Text = TIMS.ClearSQM(dr("PhoneD"))
        PhoneN.Text = TIMS.ClearSQM(dr("PhoneN"))
        CellPhone.Text = TIMS.ClearSQM(dr("CellPhone"))

        Dim vMobilYN As String = TIMS.cst_NO
        If CellPhone.Text <> "" Then vMobilYN = TIMS.cst_YES
        Common.SetListItem(rblMobil, vMobilYN)

        'ZipCode1 通訊地址
        ZipCode1.Value = Convert.ToString(dr("ZipCode1"))
        hidZipCode1_6W.Value = Convert.ToString(dr("ZipCode1_6W"))
        ZipCode1_B3.Value = TIMS.GetZIPCODEB3(hidZipCode1_6W.Value)
        ZipCode1_N.Value = Convert.ToString(dr("ZipCode1_N"))
        City1.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode1")), Convert.ToString(dr("ZipCode1_N")), objconn)
        Address.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("Address")))
        Hid_JnZipCode1.Value = TIMS.GetZipCodeJn(ZipCode1.Value, ZipCode1_B3.Value, hidZipCode1_6W.Value, City1.Text, Address.Text)

        'ZipCode2 戶籍地址
        ZipCode2.Value = Convert.ToString(dr("ZipCode2"))
        hidZipCode2_6W.Value = Convert.ToString(dr("ZipCode2_6W"))
        ZipCode2_B3.Value = TIMS.GetZIPCODEB3(hidZipCode2_6W.Value)
        ZipCode2_N.Value = Convert.ToString(dr("ZipCode2_N"))
        City2.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode2")), Convert.ToString(dr("ZipCode2_N")), objconn)
        HouseholdAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("HouseholdAddress")))
        Hid_JnZipCode2.Value = TIMS.GetZipCodeJn(ZipCode2.Value, ZipCode2_B3.Value, hidZipCode2_6W.Value, City2.Text, HouseholdAddress.Text)

        'CheckBox1.Checked=False '不管有無勾選
        CheckBox1.Checked = If(Convert.ToString(dr("ZipCode1")) = Convert.ToString(dr("ZipCode2")) AndAlso Convert.ToString(dr("Address")) = Convert.ToString(dr("HouseholdAddress")) AndAlso Convert.ToString(dr("ZipCode1_6W")) = Convert.ToString(dr("ZipCode2_6W")), True, False)

        '判斷緊急聯絡人地址是否有與通訊地址或戶籍地址相同
        CheckBox2.Checked = If(Convert.ToString(dr("ZipCode1")) = Convert.ToString(dr("ZipCode3")) AndAlso Convert.ToString(dr("Address")) = Convert.ToString(dr("EmergencyAddress")) AndAlso Convert.ToString(dr("ZipCode1_6W")) = Convert.ToString(dr("ZipCode3_6W")), True, False)
        If Not CheckBox2.Checked Then
            CheckBox3.Checked = If(Convert.ToString(dr("ZipCode2")) = Convert.ToString(dr("ZipCode3")) AndAlso Convert.ToString(dr("HouseholdAddress")) = Convert.ToString(dr("EmergencyAddress")) AndAlso Convert.ToString(dr("ZipCode2_6W")) = Convert.ToString(dr("ZipCode3_6W")), True, False)
        End If

        Email.Text = Convert.ToString(dr("Email"))

        Common.SetListItem(SubsidyID, dr("SubsidyIDEX").ToString)

        If Convert.ToString(dr("SubsidyIdentity")) <> "" Then Common.SetListItem(SubsidyIdentity, dr("SubsidyIdentity").ToString)
        SubsidyHidden.Value = If(Convert.ToString(dr("SubsidyIDEX")) = "03", "1", "0")
        SubsidyID.Attributes("onchange") = "ChangeSubsidy();"

        Common.SetListItem(MIdentityID, Convert.ToString(dr("MIdentityID")))
        hide_MIdentityID.Value = TIMS.GetListValue(MIdentityID) '.SelectedValue

        If tr_DDL_DISASTER.Visible Then
            If $"{dr("ADID")}" <> "" Then
                Common.SetListItem(DDL_DISASTER, CStr(dr("ADID")))
                tr_DDL_DISASTER.Style("display") = cst_inline1
            ElseIf hide_MIdentityID.Value = TIMS.cst_Identity_40 Then '40.重大災害受災者'重大災害選項
                tr_DDL_DISASTER.Style("display") = cst_inline1
            Else
                tr_DDL_DISASTER.Style("display") = cst_none1
            End If
        End If

        Dim flag45 As Boolean = False '確認學員是否為 中高齡
        Dim flag65 As Boolean = False '確認學員是否為 65歲(含)以上
        flag45 = TIMS.Check_YearsOld45(Birthday.Text, Convert.ToString(ViewState(vs_STDate)))
        flag65 = TIMS.Check_YearsOld65(Birthday.Text, Convert.ToString(ViewState(vs_STDate)))
        labIdentity.Text = ""
        '程序 GetOpenDate2 要先取得 'ViewState(vs_STDate)
        If flag45 Then labIdentity.Text = cst_str45yearsOld
        If flag65 Then labIdentity.Text = cst_str65yearsOld

        'If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If dr("MIdentityID").ToString="05" Then
        '        NativeTr1.Style("display")=cst_inline1
        '        If IsDBNull(dr("Native")) Then NativeID.Items(0).Selected=True
        '    Else
        '        NativeTr1.Style("display")=cst_none1
        '    End If
        'End If

        TIMS.Tooltip(DegreeID, "")
        'TIMS.Tooltip(SubsidyID, "")
        'TIMS.Tooltip(SubsidyIdentity, "")
        TIMS.Tooltip(Name, "")
        TIMS.Tooltip(IDNO, "")
        TIMS.Tooltip(Birthday, "")
        TIMS.Tooltip(PriorWorkType1, "")
        TIMS.Tooltip(PriorWorkOrg1, "")
        TIMS.Tooltip(SOfficeYM1, "")
        TIMS.Tooltip(FOfficeYM1, "")
        TIMS.Tooltip(rdo_HighEduBg, "")

        '檢查是否有職訓生活津貼 補助申請資料
        'ViewState("MsgBox")=""
        If Chk_Sub_SubSidyApply(SOCIDStr) Then
            SubsidyID.Enabled = False '鎖定
            SubsidyIdentity.Enabled = False '鎖定
            'MIdentityID.Enabled=False '鎖定
            'rdo_HighEduBg.Enabled=False '鎖定
            rdo_HighEduBg.Attributes.Add("disabled", "disabled")  '專上畢業學歷失業者

            '有職訓生活津貼，不可修改姓名 by AMU 2009-09-14
            Name.Enabled = False '鎖定 'Name.ToolTip=Cst_Msg1
            TIMS.Tooltip(Name, cst_Msg1)

            '有職訓生活津貼，不可修改身分證號碼,生日資料
            'Birthday.ReadOnly=True
            Birthday.Attributes.Add("onkeydown", "this.blur()")
            Birthday.Attributes.Add("oncontextmenu", "return false;")
            Birthday.Enabled = False
            Img1.Style("display") = cst_none1 '出生日期選擇功能
            Img1.Disabled = True 'hidBirthBtn.Disabled=True '失效
            TIMS.Tooltip(Birthday, cst_Msg1)

            IDNO.Enabled = False
            DegreeID.Enabled = False
            'Birthday.ToolTip=Cst_Msg1
            'IDNO.ToolTip=Cst_Msg1
            'SubsidyID.ToolTip=Cst_Msg1
            'SubsidyIdentity.ToolTip=Cst_Msg1
            TIMS.Tooltip(Birthday, cst_Msg1)
            TIMS.Tooltip(IDNO, cst_Msg1)
            TIMS.Tooltip(DegreeID, cst_Msg1)
            TIMS.Tooltip(SubsidyID, cst_Msg1)
            TIMS.Tooltip(SubsidyIdentity, cst_Msg1)

            If sm.UserInfo.RoleID <= 1 Then '系統管理者開放修改
                DegreeID.Enabled = True
                TIMS.Tooltip(DegreeID, cst_Msg1b)
            End If
            'ViewState("MsgBox")=Cst_Msg1 & vbCrLf
        End If

        'BtnCheckBli.Visible=True
        '28:產業人才投資計劃 54:充電起飛計畫（在職）15:學習券 '學員資料維護於訓後21日鎖定(排除計畫:28.54.15)
        If Not TIMS.Cst_TPlanID14DayCanEditStud.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Const Cst_sStartDay As String="2012/08/01" '起動日期為 "2012/08/01"
            '學員資料維護於訓後14日鎖定'Cst_Msg14
            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21st, CDate(STDateHidden.Value)), Today) >= 0 Then
                '授權設定該班級有設定則開放
                Dim Arcflag As Boolean = False '沒有權限
                Dim vTitle As String = "授權設定該班級有開放"
                If Not TIMS.ChkIsEndDate(rqOCID, TIMS.cst_FunID_學員資料維護, dtArc) Then
                    Arcflag = True '有權限
                    '授權設定該班級有開放
                    TIMS.Tooltip(RejectSOCID, vTitle) '遞補者
                    TIMS.Tooltip(Name, vTitle)
                    TIMS.Tooltip(IDNO, vTitle)
                    TIMS.Tooltip(Birthday, vTitle)
                    TIMS.Tooltip(PriorWorkType1, vTitle)
                    TIMS.Tooltip(PriorWorkOrg1, vTitle)
                    TIMS.Tooltip(SOfficeYM1, vTitle)
                    TIMS.Tooltip(FOfficeYM1, vTitle)
                    TIMS.Tooltip(rdo_HighEduBg, vTitle)
                Else
                    '未授權設定該班級 (沒有開放)
                    RejectSOCID.Enabled = False '遞補者
                    Name.Enabled = False '姓名
                    IDNO.Enabled = False '身分證號
                    'Birthday.ReadOnly=True
                    Birthday.Attributes.Add("onkeydown", "this.blur()")
                    Birthday.Attributes.Add("oncontextmenu", "return false;")
                    Birthday.Enabled = False
                    Img1.Style("display") = cst_none1 '出生日期選擇功能
                    Img1.Disabled = True
                    'hidBirthBtn.Disabled=True '失效

                    'If gFlagEnv Then 'true:正式環境。(false:測試用) / TestStr
                    '    BtnCheckBli.Visible=False
                    '    PriorWorkType1.Enabled=False
                    '    PriorWorkOrg1.Enabled=False
                    '    SOfficeYM1.Enabled=False
                    '    IMG2.Style("display")=cst_none1 '受訓前任職起日選擇功能
                    '    IMG2.Disabled=True
                    '    FOfficeYM1.Enabled=False
                    '    IMG3.Style("display")=cst_none1 '受訓前任職迄日選擇功能
                    '    IMG3.Disabled=True
                    'End If
                    'If Not gFlagEnv Then BtnCheckBli.Visible=True 'true:正式環境。(false:測試用) / TestStr

                    'rdo_HighEduBg.Enabled=False '鎖定
                    rdo_HighEduBg.Attributes.Add("disabled", "disabled")  '專上畢業學歷失業者
                End If

                If Not Arcflag Then '沒有權限 
                    Dim sMsg1 As String = cst_Msg21 '(21日鎖定)
                    TIMS.Tooltip(RejectSOCID, sMsg1) '遞補者
                    TIMS.Tooltip(Name, sMsg1) '沒有權限
                    TIMS.Tooltip(IDNO, sMsg1)
                    TIMS.Tooltip(Birthday, sMsg1)
                    TIMS.Tooltip(PriorWorkType1, sMsg1)
                    TIMS.Tooltip(PriorWorkOrg1, sMsg1)
                    TIMS.Tooltip(SOfficeYM1, sMsg1)
                    TIMS.Tooltip(FOfficeYM1, sMsg1)
                    TIMS.Tooltip(rdo_HighEduBg, sMsg1)
                End If
            End If
        End If

        'If Not IsDBNull(dr("Native")) Then Common.SetListItem(NativeID, dr("Native").ToString)

        '06:身心障礙者
        If Convert.ToString(dr("IdentityIDEX")) <> "" Then
            TIMS.Tooltip(HandTypeID, "", True)
            TIMS.Tooltip(HandLevelID, "", True)
            TIMS.Tooltip(HandTypeID2, "", True)
            TIMS.Tooltip(HandLevelID2, "", True)
            If InStr(Convert.ToString(dr("IdentityIDEX")), "06", CompareMethod.Binary) = 0 Then
                HandTypeID.Enabled = False
                HandLevelID.Enabled = False
                'tdHandTypeID2.Disabled=True
                tdHandTypeID2.Attributes.Add("disabled", "disabled")
                HandLevelID2.Enabled = False
                TIMS.Tooltip(HandTypeID, "身分別不含身心障礙者", True)
                TIMS.Tooltip(HandLevelID, "身分別不含身心障礙者", True)
                TIMS.Tooltip(HandTypeID2, "身分別不含身心障礙者", True)
                TIMS.Tooltip(HandLevelID2, "身分別不含身心障礙者", True)
            Else
                HandTypeID.Enabled = True
                HandLevelID.Enabled = True
                'tdHandTypeID2.Disabled=False
                tdHandTypeID2.Attributes.Remove("disabled")
                HandLevelID2.Enabled = True
            End If

            ViewState(vs_IdentityID) = Nothing
            '取得
            'Dim all() As String=Split(Convert.ToString(dr("IdentityIDEX")), ",", , CompareMethod.Text)
            '12:屆退官兵(須單位將級以上長官薦送函)
            'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '    Const cst_id12 As String="12"
            '    Dim flag_ChkID12 As Boolean=False '未勾選
            '    If Convert.ToString(dr("IdentityIDEX")).IndexOf(cst_id12) > -1 Then
            '        flag_ChkID12=True '已勾選 屆退官兵
            '    End If
            '    If Not flag_ChkID12 Then '未勾選 屆退官兵
            '        Dim flag_SRSOLDIERS As Boolean=False '是否為屆退官兵
            '        flag_SRSOLDIERS=TIMS.CheckRESOLDER(objconn, dr("IDNO"), sm.UserInfo.DistID,ViewState(vs_STDate))
            '        Dim MyValue As String=Convert.ToString(dr("IdentityIDEX"))
            '        If flag_SRSOLDIERS Then
            '            '為屆退官兵 且資料並無勾選 屆退官兵
            '            If MyValue <> "" Then MyValue &= ","
            '            MyValue &= cst_id12
            '            dr("IdentityIDEX")=MyValue '重新填入身分別
            '        End If
            '    End If
            'End If

            TIMS.SetCblValue(IdentityID, Convert.ToString(dr("IdentityIDEX")))
            Dim sIdentityIDEX As String = "" '集合目前所勾選的身分別存入 ViewState(vs_IdentityID)
            For i As Integer = 0 To IdentityID.Items.Count - 1
                If IdentityID.Items(i).Value <> "" AndAlso IdentityID.Items(i).Selected Then
                    If Not sIdentityIDEX <> "" Then sIdentityIDEX &= ","
                    sIdentityIDEX &= IdentityID.Items(i).Value
                End If
            Next
            If sIdentityIDEX <> "" Then ViewState(vs_IdentityID) = sIdentityIDEX
        End If

        '身心障礙者
        If Convert.ToString(dr("HandTypeID")) <> "" Then
            ' flag_HandType=1 '1:舊制
            Common.SetListItem(HandTypeID, dr("HandTypeID").ToString)
        End If
        If Convert.ToString(dr("HandLevelID")) <> "" Then
            'flag_HandType=1 '1:舊制
            Common.SetListItem(HandLevelID, dr("HandLevelID").ToString)
        End If
        If Convert.ToString(dr("HandTypeID2")) <> "" Then
            'flag_HandType=2 '2:新制
            Call TIMS.SetCblValue(HandTypeID2, Convert.ToString(dr("HandTypeID2")))
        End If
        If Convert.ToString(dr("HandLevelID2")) <> "" Then
            'flag_HandType=2 '2:新制
            Common.SetListItem(HandLevelID2, dr("HandLevelID2").ToString)
        End If

        '身心障礙者
        Dim ifg_HandType As Integer = 0 '0:未選 1:舊制 2:新制
        If Convert.ToString(dr("HandTypeID")) <> "" AndAlso Convert.ToString(dr("HandLevelID")) <> "" Then ifg_HandType = 1 '1:舊制
        If Convert.ToString(dr("HandTypeID2")) <> "" AndAlso Convert.ToString(dr("HandLevelID2")) <> "" Then ifg_HandType = 2 '2:新制

        trHandTypeID2.Style("display") = cst_none1 '新制
        trHandTypeID.Style("display") = cst_none1 '舊制
        Select Case ifg_HandType
            Case 1 '1:舊制
                trHandTypeID.Style("display") = cst_inline1 '舊制
                Common.SetListItem(rblHandType, "1")
            Case Else '0:未選 2:新制
                trHandTypeID2.Style("display") = cst_inline1 '新制
                Common.SetListItem(rblHandType, "2")
        End Select

        If Convert.ToString(dr("RejectTDate1")) <> "" Then RejectTDate1.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate1")))
        If Convert.ToString(dr("RejectTDate2")) <> "" Then RejectTDate2.Text = Common.FormatDate(Convert.ToString(dr("RejectTDate2")))
        EmergencyContact.Text = Convert.ToString(dr("EmergencyContact"))
        EmergencyPhone.Text = Convert.ToString(dr("EmergencyPhone"))
        EmergencyRelation.Text = Convert.ToString(dr("EmergencyRelation"))

        'ZipCode3 緊急聯絡人的郵遞區號
        ZipCode3.Value = Convert.ToString(dr("ZipCode3"))
        hidZipCode3_6W.Value = Convert.ToString(dr("ZipCode3_6W"))
        ZipCode3_B3.Value = TIMS.GetZIPCODEB3(hidZipCode3_6W.Value)
        ZipCode3_N.Value = Convert.ToString(dr("ZipCode3_N"))
        City3.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode3")), Convert.ToString(dr("ZipCode3_N")), objconn)
        EmergencyAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("EmergencyAddress")))
        Hid_JnZipCode3.Value = TIMS.GetZipCodeJn(ZipCode3.Value, ZipCode3_B3.Value, hidZipCode3_6W.Value, City3.Text, EmergencyAddress.Text)

        'CtID5.Value=TIMS.Get_Ctid(Convert.ToString(dr("ZipCode3")), objconn)

        ''------ sart受訓前任職清單2011/04/27 先讀CLASS_STUDENTSOFCLASS 若讀不到讀STUD_SUBDATA-------
        '若有存取過 EnterChannelSave
        'Dim flagCanRead As Boolean=False
        'If Convert.ToString(dr("EnterChannelSave"))="Y" Then flagCanRead=True

        '職前專用
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID)=-1 Then
        '    'Select Case Convert.ToString(dr("EnterChannel"))
        '    '    Case "1" '1.網;2.現;3.通;4.推
        '    '    Case Else
        '    '        '受訓前任職清單
        '    '        Call Create_PriorWorkOrg1(dr)
        '    'End Select
        '    '受訓前任職清單
        '    Call Create_PriorWorkOrg1(dr)
        'End If
        ' ------end受訓前任職清單2011/04/27 先讀CLASS_STUDENTSOFCLASS 若讀不到讀STUD_SUBDATA--------

        'Dim v_BudID As String=TIMS.GetListValue(BudID) '.SelectedValue
        v_BudID = TIMS.GetListValue(BudID) '.SelectedValue
        If v_BudID = "" AndAlso Convert.ToString(dr("BudgetID")) = "" Then
            '尚未選擇(且空白)帶預設
            '" & cst_ECFA & "預設測試  
            '=該計畫是否使用ECFA
            If blnTPlanUseEcfa Then
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If DateDiff(DateInterval.Day, CDate(cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
                        If TIMS.CheckIsECFA(Me, ActNo2.Text, "", STDateHidden.Value, objconn) = True Then Common.SetListItem(BudID, "97")  '2011/05/20 新增ECFA判斷
                    End If
                Else
                    If TIMS.CheckIsECFA(Me, ActNo2.Text, FOfficeYM1.Text, "", objconn) = True Then Common.SetListItem(BudID, "97")  '2011/05/20 新增ECFA判斷
                End If
            End If
            '該計畫是否使用ECFA
        Else
            '有值
            If Convert.ToString(dr("BudgetID")) <> "" Then Common.SetListItem(BudID, Convert.ToString(dr("BudgetID")))
        End If

        Common.SetListItem(Traffic, Convert.ToString(dr("Traffic")))
        Common.SetListItem(ShowDetail, dr("ShowDetail").ToString)
        'Common.SetListItem(BudID, dr("BudgetID").ToString)
        'If BudID.Items.Count=1 Then
        '    BudID.SelectedIndex=-1
        '    BudID.Items(0).Selected=True
        'End If
        If trPMode.Visible Then Common.SetListItem(PMode, dr("PMode").ToString)

        '應該是 0000 ~ 1111
        For i As Integer = 0 To dr("RelClass_Unit").ToString.Length - 1
            If dr("RelClass_Unit").ToString.Chars(i) = "1" Then RelClass_Unit.Items(i).Selected = True
        Next

        Unit1Hour.Text = dr("Unit1Hour").ToString
        Unit2Hour.Text = dr("Unit2Hour").ToString
        Unit3Hour.Text = dr("Unit3Hour").ToString
        Unit4Hour.Text = dr("Unit4Hour").ToString

        Unit1Score.Text = dr("Unit1Score").ToString
        Unit2Score.Text = dr("Unit2Score").ToString
        Unit3Score.Text = dr("Unit3Score").ToString
        Unit4Score.Text = dr("Unit4Score").ToString

        ActNo.Text = TIMS.ChangeIDNO(dr("ActNo").ToString)
        'Me.ViewState("IDNO")=Convert.ToString(dr("IDNO"))
        'Me.ViewState("OCID")=Convert.ToString(dr("OCID"))
        'Dim strIDNO As String=Convert.ToString(dr("IDNO"))

        Call Create1_Stud_2(dr, drCC)
    End Sub

    Sub Create1_Stud_2(ByRef dr As DataRow, ByRef drCC As DataRow)
        WSITR.Visible = False
        WSITR2.Visible = False
        If Hid_show_actno_budid.Value <> "Y" Then Return '(若不顯示預算別) (後續停止)

        Dim SOCIDStr As String = CStr(dr("SOCID"))
        Dim strIDNO As String = CStr(dr("IDNO"))
        Dim rOCID As String = CStr(dr("OCID"))

        '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
        Dim pms_1 As New Hashtable From {{"IDNO", $"{dr("IDNO")}"}, {"OCID1", TIMS.CINT1(dr("OCID"))}}
        Dim sql As String = ""
        sql &= " SELECT c.SEID" & vbCrLf '/*PK*/
        sql &= " ,c.ESERNUM,c.MIDENTITYID,c.HANDTYPEID,c.HANDLEVELID,c.PRIORWORKORG1,c.TITLE1,c.PRIORWORKORG2,c.TITLE2,c.SOFFICEYM1,c.FOFFICEYM1,c.SOFFICEYM2,c.FOFFICEYM2" & vbCrLf
        sql &= " ,c.PRIORWORKPAY,c.REALJOBLESS,c.JOBLESSID,c.TRAFFIC,c.SHOWDETAIL,c.ACCTMODE,c.POSTNO,c.ACCTHEADNO,c.BANKNAME,c.ACCTEXNO,c.EXBANKNAME,c.ACCTNO,c.FIRDATE,c.UNAME,c.INTAXNO" & vbCrLf
        'sql &= " ,c.ACTNO,c.ACTNAME" & vbCrLf
        sql &= " ,null ACTNO,null ACTNAME" & vbCrLf
        sql &= " ,c.SERVDEPT,c.JOBTITLE,c.ZIP,c.ZIP6W,c.ZIP_N,c.ADDR,c.TEL,c.FAX,c.SDATE,c.SJDATE,c.SPDATE,c.Q1,c.Q2_1,c.Q2_2,c.Q2_3,c.Q2_4,c.Q3,c.Q3_OTHER,c.Q4,c.Q5,c.Q61,c.Q62,c.Q63,c.Q64,c.ISEMAIL" & vbCrLf
        'sql &= " ,c.MODIFYACCT ,c.MODIFYDATE" & vbCrLf
        sql &= " ,c.ACTTYPE,c.SCALE,c.ZIPCODE2,c.ZIPCODE2_6W,c.ZIPCODE2_N,c.HOUSEHOLDADDRESS,c.ACTTEL,c.ZIPCODE3,c.ZIPCODE3_6W,c.ZIPCODE3_N,c.ACTADDRESS,c.INSURED,c.SERVDEPTID,c.JOBTITLEID" & vbCrLf
        sql &= " ,c.ZipCode3 ActZipCode" & vbCrLf
        sql &= " ,c.ZipCode3_6W ActZipCode_6W" & vbCrLf
        sql &= " ,c.ZipCode3_N ActZipCode_N" & vbCrLf
        sql &= " FROM dbo.STUD_ENTERTEMP2 a" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTYPE2 b ON a.esetid=b.esetid" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTRAIN2 c ON c.eSerNum=b.eSerNum" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO AND b.OCID1=@OCID1" & vbCrLf
        'Dim drT2 As DataRow '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
        Dim drT2 As DataRow = DbAccess.GetOneRow(sql, objconn, pms_1) 'STUD_ENTERTRAIN2

        '學員服務單位(產學訓)
        Dim pms_SP As New Hashtable From {{"SOCID", TIMS.CINT1(SOCIDStr)}}
        Dim sql_SP As String = "SELECT * FROM STUD_SERVICEPLACE WHERE SOCID=@SOCID"
        dr = DbAccess.GetOneRow(sql_SP, objconn, pms_SP) 'STUD_SERVICEPLACE
        '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
        If dr Is Nothing Then dr = drT2

        If dr IsNot Nothing Then
            '為勞工團體時，會多一個訓練單位代轉現金的選項，所以增加=2的Flag判斷
            'AcctMode 0:郵政1:金融(銀行)2:訓練單位代轉現金
            Select Case Convert.ToString(dr("AcctMode"))
                Case "0"
                    'AcctMode.SelectedIndex=0
                    Common.SetListItem(AcctMode, "0")
                    Dim sPostNo As String = TIMS.ClearSQM(dr("PostNo"))
                    If sPostNo <> "" AndAlso sPostNo.IndexOf("-") <> -1 Then sPostNo = Replace(sPostNo, "-", "")
                    PostNo_1.Text = sPostNo

                    Dim sAcctNo As String = TIMS.ClearSQM(dr("AcctNo"))
                    If sAcctNo <> "" AndAlso sAcctNo.IndexOf("-") <> -1 Then sAcctNo = Replace(sAcctNo, "-", "")
                    AcctNo1_1.Text = sAcctNo

                    PortTR.Style("display") = cst_inline1
                    BankTR1.Style("display") = cst_none1
                    BankTR2.Style("display") = cst_none1
                    BankTR3.Style("display") = cst_none1
                Case "2"
                    '**by Milor 20080509--由訓練單位代轉現金時，所有轉帳資料都不填入值----start
                    '當取出的資料是選擇訓練單位代轉現金時，所有的帳號填入欄位都不顯示
                    'AcctMode.SelectedIndex=2
                    Common.SetListItem(AcctMode, "2")

                    PortTR.Style("display") = cst_none1
                    BankTR1.Style("display") = cst_none1
                    BankTR2.Style("display") = cst_none1
                    BankTR3.Style("display") = cst_none1
                    '**by Milor 20080509----end
                Case Else '"1"
                    'AcctMode.SelectedIndex=1
                    Common.SetListItem(AcctMode, "1")
                    BankName.Text = dr("BankName").ToString
                    AcctheadNo.Text = dr("AcctHeadNo").ToString
                    'amu 20061225 kevin同意再次加入
                    ExBankName.Text = dr("ExBankName").ToString
                    AcctExNo.Text = dr("AcctExNo").ToString
                    AcctNo2.Text = dr("AcctNo").ToString

                    PortTR.Style("display") = cst_none1
                    BankTR1.Style("display") = cst_inline1
                    BankTR2.Style("display") = cst_inline1
                    BankTR3.Style("display") = cst_inline1
            End Select

            FirDate.Text = TIMS.Cdate3(dr("FirDate"))
            'If IsDate(dr("FirDate")) Then FirDate.Text=Common.FormatDate(dr("FirDate"))
            Uname.Text = dr("Uname").ToString
            Intaxno.Text = dr("Intaxno").ToString

            ActName.Text = Convert.ToString(dr("ActName")) 'STUD_SERVICEPLACE
            ActNo1.Text = Convert.ToString(dr("ActNo")) 'STUD_SERVICEPLACE
            'OJT-24102401：<系統> 產投、充飛 - 學員資料維護：調整欄位位置與相關輸入卡控
            If sm.UserInfo.LID <> 0 Then
                'ActNo1.Enabled=False
                TIMS.INPUT_ReadOnly(ActName, cst_Msg2c)
                TIMS.INPUT_ReadOnly(ActNo1, cst_Msg2c)
                'TIMS.Tooltip(ActNo1, Cst_Msg2c, True)
                'ActNo1.CssClass="in-read-only-a"
                'ActNo1.Attributes.Add("onfocus", "this.blur()")
                'ActNo1.Attributes.Add("onkeydown", "this.blur()")
                'ActNo1.Attributes.Add("oncontextmenu", "return false;")
            Else
                TIMS.Tooltip(ActName, $"{cst_Msg2c},權限開放署", True)
                TIMS.Tooltip(ActNo1, $"{cst_Msg2c},權限開放署", True)
            End If

            Dim s_ACTNAME_bli As String = ""
            Dim s_ACTNO_bli As String = ""
            Dim flag_can_use_default_actno As Boolean = False
            'BEGINCLASS,Y,(過開訓日)
            If Convert.ToString(drCC("BEGINCLASS")).Equals("Y") Then
                '過開訓日才可使用勾稽資料
                's_ACTNAME_bli=TIMS.GET_BLIGATEDATA28E(strIDNO, rOCID, objconn, "ACTNAME") 'STUD_BLIGATEDATA28E
                's_ACTNO_bli=TIMS.GET_BLIGATEDATA28E(strIDNO, rOCID, objconn, "ACTNO") 'STUD_BLIGATEDATA28E
                s_ACTNAME_bli = TIMS.GET_BLIGATEDATA28(SOCIDStr, strIDNO, objconn, "ACTNAME") 'STUD_BLIGATEDATA28
                s_ACTNO_bli = TIMS.GET_BLIGATEDATA28(SOCIDStr, strIDNO, objconn, "ACTNO") 'STUD_BLIGATEDATA28
                flag_can_use_default_actno = True
            End If
            '預設使用勾稽資料 '如果有儲存資料就使用儲存資料
            If flag_can_use_default_actno Then
                'STUD_BLIGATEDATA28E / STUD_SERVICEPLACE
                If ActName.Text = "" Then ActName.Text = s_ACTNAME_bli 'STUD_SERVICEPLACE
                If ActNo1.Text = "" Then ActNo1.Text = s_ACTNO_bli 'STUD_SERVICEPLACE
            End If

            '最後存取值 'If s_ActNo <> "" AndAlso ActNo1.Text="" Then ActNo1.Text=s_ActNo 'ActType: 投保類別1.勞2.農3.漁
            'If Convert.ToString(dr("ActType")) <> "" Then Common.SetListItem(ActType, dr("ActType").ToString)
            ServDept.Text = $"{dr("ServDept")}"
            JobTitle.Text = $"{dr("JobTitle")}"
            If $"{dr("SERVDEPTID")}" <> "" Then Common.SetListItem(ddlSERVDEPTID, dr("SERVDEPTID"))
            If $"{dr("JOBTITLEID")}" <> "" Then Common.SetListItem(ddlJOBTITLEID, dr("JOBTITLEID"))

            Zip.Value = "" 'TIMS.GetValue1(dr("Zip"))
            hidZIP6W.Value = ""
            ZIPB3.Value = "" 'TIMS.GetValue1(dr("ZIP"))
            Zip_N.Value = "" 'TIMS.GetValue1(dr("ZIP_N"))
            City5.Text = ""
            Addr.Text = ""
            '-1表示來自 drT2 沒有 ZIP6W 的值
            If Convert.ToString(dr("Zip")) <> "-1" AndAlso Convert.ToString(dr("Zip")) <> "" Then
                'Zip 通訊地址
                Zip.Value = Convert.ToString(dr("Zip"))
                hidZIP6W.Value = Convert.ToString(dr("ZIP6W"))
                ZIPB3.Value = TIMS.GetZIPCODEB3(hidZIP6W.Value)
                Zip_N.Value = Convert.ToString(dr("ZIP_N"))
                City5.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("Zip")), Convert.ToString(dr("ZIP_N")), objconn)
                Addr.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("Addr")))
                Hid_JnZip.Value = TIMS.GetZipCodeJn(Zip.Value, ZIPB3.Value, hidZIP6W.Value, City5.Text, Addr.Text)
            End If

            Tel.Text = $"{dr("Tel")}"
            Fax.Text = $"{dr("Fax")}"

            If IsDate(dr("SDate")) Then SDate.Text = Common.FormatDate(dr("SDate"))
            If IsDate(dr("SJDate")) Then SJDate.Text = Common.FormatDate(dr("SJDate"))
            If IsDate(dr("SPDate")) Then SPDate.Text = Common.FormatDate(dr("SPDate"))

            txt_ActPhone.Text = Convert.ToString(dr("ActTel"))  '加入投保單位電話、地址

            'ActZipCode  
            txt_ActZip.Value = Convert.ToString(dr("ActZipCode"))
            hid_ActZIP6W.Value = Convert.ToString(dr("ActZipCode_6W"))
            txt_ActZIPB3.Value = TIMS.GetZIPCODEB3(hid_ActZIP6W.Value)
            hidActZip_N.Value = Convert.ToString(dr("ActZipCode_N"))
            txt_ActCity.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ActZipCode")), Convert.ToString(dr("ActZipCode_N")), objconn)
            txt_ActAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ActAddress")))
            Hid_JnActZip.Value = TIMS.GetZipCodeJn(txt_ActZip.Value, txt_ActZIPB3.Value, hid_ActZIP6W.Value, txt_ActCity.Text, txt_ActAddress.Text)
        End If

        '相同時 Checkbox4 打勾
        Checkbox4.Checked = If(txt_ActZip.Value <> "" AndAlso txt_ActZip.Value = Zip.Value AndAlso hid_ActZIP6W.Value = hidZIP6W.Value AndAlso txt_ActAddress.Text = Addr.Text, True, False)

        '學員參訓背景(產學訓)
        Dim pms_BG As New Hashtable From {{"SOCID", TIMS.CINT1(SOCIDStr)}}
        Dim sql_BG As String = " SELECT * FROM STUD_TRAINBG WHERE SOCID=@SOCID"
        dr = DbAccess.GetOneRow(sql_BG, objconn, pms_BG)
        '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
        If dr Is Nothing Then dr = drT2

        If dr IsNot Nothing Then
            '是否由公司推薦參訓 'STUD_ENTERTRAIN2
            If Convert.ToString(dr("Q1")) <> "" Then Common.SetListItem(Q1, dr("Q1")) 'numeric
            If Convert.ToString(dr("Q3")) <> "" Then Common.SetListItem(Q3, dr("Q3")) 'numeric
            If Convert.ToString(dr("Q4")) <> "" Then Common.SetListItem(Q4, dr("Q4")) 'varchar
            If Convert.ToString(dr("Q5")) <> "" Then Common.SetListItem(Q5, dr("Q5")) 'numeric

            Q61.Text = Convert.ToString(dr("Q61"))
            Q62.Text = Convert.ToString(dr("Q62"))
            Q63.Text = Convert.ToString(dr("Q63"))
            Q64.Text = Convert.ToString(dr("Q64"))
        End If

        Dim pms_BG2 As New Hashtable From {{"SOCID", TIMS.CINT1(SOCIDStr)}}
        Dim sql_BG2 As String = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID=@SOCID"
        Dim dt_BG2 As DataTable = DbAccess.GetDataTable(sql_BG2, objconn, pms_BG2)
        If TIMS.dtHaveDATA(dt_BG2) Then
            For Each dr In dt_BG2.Rows
                For Each item As ListItem In Q2.Items
                    If Convert.ToString(dr("Q2")) = item.Value Then
                        item.Selected = True
                        Exit For
                    End If
                Next
            Next
        Else
            '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
            dr = drT2
            If dr IsNot Nothing Then
                Dim s_q2_1 As String = $"{dr("q2_1")}"
                Dim s_q2_2 As String = $"{dr("q2_2")}"
                Dim s_q2_3 As String = $"{dr("q2_3")}"
                Dim s_q2_4 As String = $"{dr("q2_4")}"
                For Each item As ListItem In Q2.Items
                    Select Case item.Value
                        Case "1"
                            If s_q2_1 = "1" Then item.Selected = True
                        Case "2"
                            If s_q2_2 = "1" Then item.Selected = True
                        Case "3"
                            If s_q2_3 = "1" Then item.Selected = True
                        Case "4"
                            If s_q2_4 = "1" Then item.Selected = True
                    End Select
                Next
            End If
        End If

        '20090330(Milor)專上畢業學歷失業者
        Dim v_HighEduBg As String = TIMS.GetListValue(rdo_HighEduBg) '.SelectedValue
        If ViewState(vs_HighEduBg) = True Then
            '追加特別預算
            Const Cst_特別預算 As String = "特別預算" '98
            With BudID
                'If .Items.IndexOf(.Items.FindByText(Cst_特別預算))=-1 Then  .Items.Insert(.Items.Count, New ListItem(Cst_特別預算, "98"))
                If .Items.FindByValue("98") Is Nothing Then .Items.Insert(.Items.Count, New ListItem(Cst_特別預算, "98"))
            End With
            'If BudID.SelectedValue="" AndAlso Convert.ToString(dr("BudgetID"))="" Then '尚未選擇
            'End If
            If v_HighEduBg = "Y" Then
                Common.SetListItem(BudID, "98") ' 專上畢業學歷失業者 必使用 特別預算
                'BudID.SelectedIndex=BudID.Items.Count - 1 '最後值
                BudID.Attributes.Add("disabled", "disabled") '鎖定
            End If
        End If

        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '含職前webservice
        '    SubsidyCost.Text=TIMS.Get_SubsidyCost(IDNO.Text, STDateHidden.Value, "", "Y", objconn)
        '    WSITR.Visible=True
        '    WSITR2.Visible=True
        'End If
    End Sub

    '若有此SOCID 則為真， 其它情況則為 否
    Function Chk_Sub_SubSidyApply(ByVal SOCID As String) As Boolean
        Dim rst As Boolean = False
        SOCID = TIMS.ClearSQM(SOCID)
        If SOCID = "" Then Return rst

        Dim hPMS As New Hashtable From {{"SOCID", TIMS.CINT1(SOCID)}}
        Dim sql As String = " SELECT 'x' FROM SUB_SUBSIDYAPPLY WHERE SOCID=@SOCID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, hPMS)
        If dr IsNot Nothing Then rst = True
        Return rst
    End Function

    '設定 RIDValue.Value & DistValue.Value  
    Sub FindRIDValue()
        '設定 RIDValue.Value & DistValue.Value  
        DistValue.Value = ""
        RIDValue.Value = ""
        Dim vRIDValue As String = TIMS.GetMyValue(ViewState(vs_SearchStr), "RIDValue")
        If vRIDValue = "" Then Return
        vRIDValue = TIMS.ClearSQM(vRIDValue)
        If vRIDValue = "" Then Return
        RIDValue.Value = vRIDValue

        Dim hPMS As New Hashtable From {{"RID", vRIDValue}}
        Dim sql As String = " SELECT DISTID FROM AUTH_RELSHIP WHERE RID=@RID"
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn, hPMS)
        If dr1 Is Nothing Then Return
        DistValue.Value = Convert.ToString(dr1("DistID"))
    End Sub

    '產學訓專用 學員限制 (Enabled /Disabled) true:鎖定 false:開放
    Sub FacLimit(ByVal fg_lock_input As Boolean, ByVal sTip As String)
        'fg_lock_input :True: '鎖定 / False :解除
        If fg_lock_input Then
            Dim v_RejectSOCID As String = TIMS.GetListValue(RejectSOCID) '.SelectedValue
            If v_RejectSOCID <> "" Then RejectSOCID.Enabled = False '遞補者
            If Name.Text <> "" Then Name.Enabled = False
            If RMPNAME.Text <> "" Then RMPNAME.Enabled = False
            If LName.Text <> "" Then LName.Enabled = False
            If FName.Text <> "" Then FName.Enabled = False
            If IDNO.Text <> "" Then IDNO.Enabled = False
            If Birthday.Text <> "" Then
                'Birthday.ReadOnly=True
                Birthday.Attributes.Add("onkeydown", "this.blur()")
                Birthday.Attributes.Add("oncontextmenu", "return false;")
                Birthday.Enabled = False
                Img1.Style("display") = cst_none1 '出生日期選擇功能
                Img1.Disabled = True
                'hidBirthBtn.Disabled=True '失效
            End If
            If BudID.SelectedIndex <> -1 Then BudID.Enabled = False Else BudID.Enabled = True
            If MIdentityID.SelectedIndex <> -1 Then MIdentityID.Enabled = False Else MIdentityID.Enabled = True
            If IdentityID.SelectedIndex <> -1 Then IdentityID.Enabled = False Else IdentityID.Enabled = True
            If SupplyID.SelectedIndex <> -1 Then SupplyID.Enabled = False Else SupplyID.Enabled = True

            Dim s_tit12 As String = cst_Msg2
            TIMS.Tooltip(Name, s_tit12)
            TIMS.Tooltip(LName, s_tit12)
            TIMS.Tooltip(FName, s_tit12)
            TIMS.Tooltip(IDNO, s_tit12)
            TIMS.Tooltip(Birthday, s_tit12)
            TIMS.Tooltip(MIdentityID, s_tit12, True)
            TIMS.Tooltip(IdentityID, s_tit12, True)
            TIMS.Tooltip(BudID, s_tit12, True)
            TIMS.Tooltip(SupplyID, s_tit12, True)
        Else
            RejectSOCID.Enabled = True '遞補者
            'OJT-24102401：<系統> 產投、充飛 - 學員資料維護：調整欄位位置與相關輸入卡控
            'Name.Enabled=(sm.UserInfo.LID=0)
            'If Not Name.Enabled Then TIMS.Tooltip(Name, Cst_Msg2b, True)
            RMPNAME.Enabled = True
            LName.Enabled = True
            FName.Enabled = True
            IDNO.Enabled = True

            '產學訓停用修改出生年月日(TIMS.Tooltip(Birthday, cst_msgBirth)) 'Birthday.ReadOnly=False
            Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                Case "2" '委訓單位 產學訓停用修改出生年月日 BY AMU 20151103
                    'Birthday.ReadOnly=True
                    Birthday.Attributes.Add("onkeydown", "this.blur()")
                    Birthday.Attributes.Add("oncontextmenu", "return false;")
                    Birthday.Enabled = False
                    Img1.Style("display") = cst_none1 '出生日期選擇功能
                    Img1.Disabled = True '失效
                    'hidBirthBtn.Disabled=True '失效
                    'TIMS.Tooltip(Birthday, sTip)
                Case Else
                    Birthday.Attributes.Remove("onkeydown")
                    Birthday.Attributes.Remove("oncontextmenu")
                    Birthday.Enabled = True
                    Img1.Style("display") = cst_inline1 '出生日期選擇功能
                    Img1.Disabled = False '有效
                    'Img1.Attributes("onclick")="callCalendar('" & Birthday.ClientID & "','" & hidBirthBtn.ClientID & "');"
                    'hidBirthBtn.Disabled=False '有效
                    TIMS.Tooltip(Birthday, sTip)
            End Select

            BudID.Enabled = True
            MIdentityID.Enabled = True
            IdentityID.Enabled = True
            SupplyID.Enabled = True
            'Name.ToolTip=String.Empty
            RMPNAME.ToolTip = String.Empty
            LName.ToolTip = String.Empty 'LName.ToolTip.Empty
            FName.ToolTip = String.Empty 'FName.ToolTip.Empty
            IDNO.ToolTip = String.Empty 'IDNO.ToolTip.Empty
            Birthday.ToolTip = String.Empty 'Birthday.ToolTip.Empty
            BudID.ToolTip = String.Empty 'BudID.ToolTip.Empty
            MIdentityID.ToolTip = String.Empty 'MIdentityID.ToolTip.Empty
            IdentityID.ToolTip = String.Empty 'IdentityID.ToolTip.Empty
            'SupplyID.ToolTip=String.Empty 'SupplyID.ToolTip.Empty

            TIMS.Tooltip(RejectSOCID, sTip) '遞補者
            'TIMS.Tooltip(Name, sTip)
            TIMS.Tooltip(RMPNAME, sTip)
            TIMS.Tooltip(LName, sTip)
            TIMS.Tooltip(FName, sTip)
            TIMS.Tooltip(IDNO, sTip)
            'TIMS.Tooltip(Birthday, sTip)
            TIMS.Tooltip(BudID, sTip)
            TIMS.Tooltip(MIdentityID, sTip)
            TIMS.Tooltip(IdentityID, sTip)
            TIMS.Tooltip(SupplyID, sTip)
        End If

        '產投相關計畫執行此功能。
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If Hid_show_actno_budid.Value = "Y" Then
            flag_BudIDNoLock = False '如果是分署(中心)承辦人，預算別不鎖定。by AMU 20140328 (本功能是每次執行)
            'iBudFlag - OUT: 'iBudFlag :0,1,2: 'iFlag :0:未開放 1:21天內修改 2:開放被登功能
            flag_BudIDNoLock = Chk_CanEditBudgetID(sm.UserInfo.LID, CDate(STDateHidden.Value), iBudFlag) '什麼時候可以修改預算別。
            If flag_BudIDNoLock Then
                If Not BudID.Enabled Then
                    BudID.Enabled = True
                    TIMS.Tooltip(BudID, cst_Msg21ok)
                End If
            End If
        End If
    End Sub

    '學員資料審核功能的欄位鎖住
    Sub GetScript(ByVal SOCID As String)
        Dim dt As DataTable = Nothing
        '------ start -----學員資料審核功能的欄位鎖住,若按確認為鎖住
        Dim hPMS As New Hashtable From {{"SOCID", TIMS.CINT1(SOCID)}}
        Dim sql As String = " SELECT STUDENTID,ISAPPRPAPER,APPLIEDRESULT FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID"
        dt = DbAccess.GetDataTable(sql, objconn, hPMS)
        If TIMS.dtHaveDATA(dt) Then
            Dim dr As DataRow = dt.Rows(0)
            Dim fg_Stud_IsApprPaper As Boolean = (dr("IsApprPaper").ToString = "Y") '1.假如學員資料確定就鎖住某些欄位
            Dim fg_Stud_AppliedResult As Boolean = (dr("AppliedResult").ToString = "Y") '2.假如學員資料審核通過就鎖住某些欄位
            '如果是系統管理者開啟功能。
            Dim fg_acct_IsSuperUser_1 As Boolean = TIMS.IsSuperUser(Me, 1)
            '1.假如學員資料確定就鎖住某些欄位 , '2.假如學員資料審核通過就鎖住某些欄位
            If (fg_Stud_IsApprPaper AndAlso fg_Stud_AppliedResult) Then
                If fg_acct_IsSuperUser_1 Then
                    Call FacLimit(False, cst_Msg1b) 'ROLEID=0 LID=0
                Else
                    Call FacLimit(True, cst_Msg2) '其他使用者鎖定。
                End If
            End If
        End If
        '------ End -----學員資料審核功能的欄位鎖住,若按確認為鎖住

        Dim sStudIDtmps As String = ""
        For i As Integer = 0 To dt.Rows.Count - 1
            If Len(dt.Rows(i).Item("StudentID")) = 12 Then
                sStudIDtmps &= String.Concat(If(sStudIDtmps <> "", ",", ""), "'", Right(dt.Rows(i).Item("StudentID"), 3), "'")
            Else
                sStudIDtmps &= String.Concat(If(sStudIDtmps <> "", ",", ""), "'", Right(dt.Rows(i).Item("StudentID"), 2), "'")
            End If
        Next

        Dim javascript As String = ""
        javascript = "<script language='javascript'>" & vbCrLf
        javascript &= " function chk_studentID(num,obj){" & vbCrLf
        javascript &= String.Concat("   var all=new Array(", sStudIDtmps, ");", vbCrLf)
        javascript &= "   for(var i=0;i<all.length;i++){" & vbCrLf
        javascript &= "     if(document.form1.StudentID.value==all[i] && all[i]!=document.form1.StudentIDstring.value){" & vbCrLf
        javascript &= "       alert('學號重複');" & vbCrLf
        javascript &= "       obj.focus();" & vbCrLf
        javascript &= "     }" & vbCrLf
        javascript &= "   }" & vbCrLf
        javascript &= " }" & vbCrLf
        javascript &= "</script>"

        Page.RegisterStartupScript("chk_studentID", javascript)
        'chgPriorWorkType1_disabled();
        Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")
        'Me.ViewState("script")=javascript
    End Sub

    '取出補助費用歷史頁 補助金額 學員輔助金撥款檔
    Sub GetHistorySumOfMoney(ByVal iSOCID As Integer)
        Const Cst_審核補助金額 As Integer = 5 '審核補助金額
        Const Cst_撥款補助金額 As Integer = 6 '撥款補助金額
        'Dim   sqlstr As String
        Dim dt As DataTable = Nothing

        Dim hPMS As New Hashtable From {{"SOCID", iSOCID}}
        Dim sql As String = ""
        sql &= " SELECT a.sid ,a.IDNO ,a.Name ,b.SOCID,c.ClassCName,c.CyclType" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,c.STDate ,c.FTDate ,c.Years" & vbCrLf
        sql &= " ,d.PlanYear ,d.PlanID ,d.ComIDNO ,d.SeqNo" & vbCrLf
        sql &= " ,ISNULL(e.SumOfMoney,0) SumOfMoney" & vbCrLf '審核
        sql &= " ,ISNULL(e2.SumOfMoney,0) SumOfMoney2" & vbCrLf '撥款
        sql &= " ,e.BUDID ,bb.budname" & vbCrLf '預算別
        sql &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS b ON a.SID=b.SID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO c ON b.OCID=c.OCID AND c.NotOpen='N'" & vbCrLf
        sql &= " JOIN Plan_Planinfo d ON c.PlanID=d.PlanID AND c.ComIDNO=d.ComIDNO AND c.SeqNO=d.SeqNO" & vbCrLf
        '審核結果 Y通過
        sql &= "  AND d.AppliedResult='Y' AND d.DefStdCost > 0" & vbCrLf
        '學員經費審核狀態-申請 Y成功
        sql &= " LEFT JOIN STUD_SUBSIDYCOST e ON b.SOCID=e.SOCID AND e.AppliedStatusM='Y'" & vbCrLf
        sql &= " LEFT JOIN VIEW_BUDGET bb ON bb.budid=e.budid" & vbCrLf
        'OR 學員經費撥款狀態 1通過 0失敗
        sql &= " LEFT JOIN STUD_SUBSIDYCOST e2 ON b.SOCID=e2.SOCID AND e2.AppliedStatus=1" & vbCrLf
        sql &= " WHERE EXISTS ( SELECT 'x'" & vbCrLf
        sql &= "  FROM STUD_STUDENTINFO ca" & vbCrLf
        sql &= "  JOIN CLASS_STUDENTSOFCLASS cb ON ca.SID=cb.SID" & vbCrLf
        sql &= "  WHERE ca.IDNO=a.IDNO AND cb.SOCID=@SOCID )" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn, hPMS)

        'Session("TC_table")=dt
        'hide_SumOfMoney.Value=0
        Panel.Visible = False
        DataGrid1.Visible = False
        msg.Text = "查無資料!!"
        'Me.bt_save.Visible=False
        If dt.Rows.Count > 0 Then
            'Me.bt_save.Visible=True
            Panel.Visible = True
            msg.Text = ""
            DataGrid1.Visible = True
            DataGrid1.Columns(Cst_審核補助金額).Visible = False
            DataGrid1.Columns(Cst_撥款補助金額).Visible = True
            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                DataGrid1.Columns(Cst_審核補助金額).Visible = True
                DataGrid1.Columns(Cst_撥款補助金額).Visible = False
            End If
            dt.DefaultView.Sort = "PlanYear,STDate"
            dt = TIMS.dv2dt(dt.DefaultView)
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
            'PageControler1.PageDataTable=dt '.SqlString=sqlstr
            'PageControler1.PrimaryKey="PlanYear"
            'PageControler1.Sort="PlanYear,STDate"
            'PageControler1.ControlerLoad()
            'Me.bt_save.Visible=False
        End If
        'dr=DbAccess.GetOneRow(sql)
        'DGIdentValue.Text=dr("Share_Name")
    End Sub

    ''' <summary>
    ''' 清理資料 (增加欄位規則性停用或可用)
    ''' </summary>
    Sub Clear_data()
        EnterChannel.Enabled = True
        TRNDMode.Enabled = True
        'TRNDType.Enabled=True
        SubsidyID.Enabled = True '有效
        SubsidyIdentity.Enabled = True '有效
        MIdentityID.Enabled = True '有效

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                Case "2"
                    '委訓單位 產學訓停用修改出生年月日 BY AMU 20151103
                    'Birthday.ReadOnly=True
                    Birthday.Attributes.Add("onkeydown", "this.blur()")
                    Birthday.Attributes.Add("oncontextmenu", "return false;")
                    Birthday.Enabled = False
                    Img1.Style("display") = cst_none1 '出生日期選擇功能
                    Img1.Disabled = True '失效
                    'hidBirthBtn.Disabled=True '失效
                    'TIMS.Tooltip(Birthday, sTip)
            End Select
        Else
            '非產投還可以修改出生年月日
            Birthday.ReadOnly = False
            Birthday.Attributes.Remove("onkeydown")
            Birthday.Attributes.Remove("oncontextmenu")
            Birthday.Enabled = True
            Img1.Style("display") = cst_inline1 '出生日期選擇功能
            Img1.Disabled = False '有效
            'hidBirthBtn.Disabled=False '有效
        End If

        IDNO.Enabled = True
        '身心障礙者
        HandTypeID.Enabled = True
        HandLevelID.Enabled = True
        'tdHandTypeID2.Disabled=False
        tdHandTypeID2.Attributes.Add("disabled", "disabled")
        HandLevelID2.Enabled = True
        RejectSOCID.Enabled = True '遞補者
        'Name.Enabled=True
        RMPNAME.Enabled = True
        LName.Enabled = True
        FName.Enabled = True
        'IDNO.Enabled=True
        'Birthday.Enabled=True
        BudID.Enabled = True
        'MIdentityID.Enabled=True
        IdentityID.Enabled = True
        SupplyID.Enabled = True
        LevelNo.Enabled = True

        Name.Text = ""
        RMPNAME.Text = ""
        StudentID.Text = ""
        LName.Text = ""
        FName.Text = ""
        IDNO.Text = ""
        Birthday.Text = ""
        OpenDate.Text = ""
        CloseDate.Text = ""
        EnterDate.Text = ""
        School.Text = ""
        Department.Text = ""
        ServiceID.Text = ""
        MilitaryAppointment.Text = ""
        MilitaryRank.Text = ""
        ServiceOrg.Text = ""
        ChiefRankName.Text = ""
        ServicePhone.Text = ""
        SServiceDate.Text = ""
        FServiceDate.Text = ""
        ZipCode4.Value = ""
        hidZipCode4_6W.Value = ""
        ZipCode4_B3.Value = ""
        ZipCode4_N.Value = ""
        City4.Text = ""
        ServiceAddress.Text = ""

        PhoneD.Text = ""
        PhoneN.Text = ""
        CellPhone.Text = ""
        rblMobil.SelectedIndex = -1
        'Common.SetListItem(rblMobil, "Y")

        ZipCode1.Value = ""
        hidZipCode1_6W.Value = ""
        ZipCode1_B3.Value = ""
        ZipCode1_N.Value = ""
        City1.Text = ""
        Address.Text = ""

        ZipCode2.Value = ""
        hidZipCode2_6W.Value = ""
        ZipCode2_B3.Value = ""
        ZipCode2_N.Value = ""
        City2.Text = ""
        HouseholdAddress.Text = ""

        Email.Text = ""
        RejectTDate1.Text = ""
        RejectTDate2.Text = ""
        EmergencyContact.Text = ""
        EmergencyPhone.Text = ""
        EmergencyRelation.Text = ""

        ZipCode3.Value = ""
        hidZipCode3_6W.Value = ""
        ZipCode3_B3.Value = ""
        ZipCode3_N.Value = ""
        City3.Text = ""
        EmergencyAddress.Text = ""

        PriorWorkType1.SelectedIndex = -1
        PriorWorkOrg1.Text = ""
        PriorWorkOrg2.Text = ""
        SOfficeYM1.Text = ""
        FOfficeYM1.Text = ""
        SOfficeYM2.Text = ""
        FOfficeYM2.Text = ""
        ActNo2.Text = ""
        PriorWorkPay.Text = ""
        Title1.Text = ""
        Title2.Text = ""
        'RealJobless.Text=""
        'JoblessID.SelectedIndex=-1

        PostNo_1.Text = ""
        'PostNo_2.Text=""
        AcctNo1_1.Text = ""
        'AcctNo1_2.Text=""
        BankName.Text = ""

        AcctheadNo.Text = ""
        ExBankName.Text = ""

        AcctExNo.Text = ""
        AcctNo2.Text = ""
        FirDate.Text = ""
        Uname.Text = ""
        Intaxno.Text = ""
        ActName.Text = ""
        ActNo1.Text = ""

        Zip.Value = ""
        hidZIP6W.Value = ""
        ZIPB3.Value = ""
        Zip_N.Value = ""
        City5.Text = ""
        Addr.Text = ""

        Tel.Text = ""
        Fax.Text = ""

        ServDept.Text = ""
        JobTitle.Text = ""
        ddlSERVDEPTID.SelectedIndex = -1
        ddlJOBTITLEID.SelectedIndex = -1
        SDate.Text = ""
        SJDate.Text = ""
        SPDate.Text = ""

        '加入投保單位電話、地址
        txt_ActPhone.Text = ""
        'ActZipCode  
        txt_ActZip.Value = ""
        hid_ActZIP6W.Value = ""
        txt_ActZIPB3.Value = ""
        hidActZip_N.Value = ""
        txt_ActCity.Text = ""
        txt_ActAddress.Text = ""

        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = False
        Checkbox4.Checked = False
        'Checkbox2.Checked=False

        'If ViewState("ADD") <> 1 Then
        'End If
        If Not PassPortNO.SelectedItem Is Nothing Then PassPortNO.SelectedItem.Selected = False
        If Not Sex.SelectedItem Is Nothing Then Sex.SelectedItem.Selected = False
        If Not MaritalStatus.SelectedItem Is Nothing Then MaritalStatus.SelectedItem.Selected = False
        If Not DegreeID.SelectedItem Is Nothing Then DegreeID.SelectedItem.Selected = False
        If Not GraduateStatus.SelectedItem Is Nothing Then GraduateStatus.SelectedItem.Selected = False
        If Not graduatey.SelectedItem Is Nothing Then graduatey.SelectedItem.Selected = False
        If Not MilitaryID.SelectedItem Is Nothing Then MilitaryID.SelectedItem.Selected = False
        'If Not NativeID.SelectedItem Is Nothing Then NativeID.SelectedItem.Selected=False
        If Not SubsidyID.SelectedItem Is Nothing Then SubsidyID.SelectedItem.Selected = False
        If Not SubsidyIdentity.SelectedItem Is Nothing Then SubsidyIdentity.SelectedItem.Selected = False

        '身心障礙者
        trHandTypeID2.Style("display") = cst_inline1 '新制
        trHandTypeID.Style("display") = cst_none1    '舊制
        Common.SetListItem(rblHandType, "2")

        If Not HandTypeID.SelectedItem Is Nothing Then HandTypeID.SelectedItem.Selected = False
        If Not HandLevelID.SelectedItem Is Nothing Then HandLevelID.SelectedItem.Selected = False
        Call TIMS.SetCblValue(HandTypeID2, "")
        If Not HandLevelID2.SelectedItem Is Nothing Then HandLevelID2.SelectedItem.Selected = False
        'If Not JoblessID.SelectedItem Is Nothing Then JoblessID.SelectedItem.Selected=False
        If Not Traffic.SelectedItem Is Nothing Then Traffic.SelectedItem.Selected = False
        If Not ShowDetail.SelectedItem Is Nothing Then ShowDetail.SelectedItem.Selected = False
        If Not SupplyID.SelectedItem Is Nothing Then SupplyID.SelectedItem.Selected = False
        If Not TRNDMode.SelectedItem Is Nothing Then TRNDMode.SelectedItem.Selected = False
        'If Not TRNDType.SelectedItem Is Nothing Then TRNDType.SelectedItem.Selected=False
        If Not EnterChannel.SelectedItem Is Nothing Then EnterChannel.SelectedItem.Selected = False
        If Not JobStateType.SelectedItem Is Nothing Then JobStateType.SelectedItem.Selected = False '就職狀況'0:失業 1:在職 

        For i As Integer = 0 To IdentityID.Items.Count - 1
            IdentityID.Items(i).Selected = False
        Next
        If Not BudID.SelectedItem Is Nothing Then BudID.SelectedItem.Selected = False
        For i As Integer = 0 To RelClass_Unit.Items.Count - 1
            RelClass_Unit.Items(i).Selected = False
        Next
        'For i As Integer=0 To BudID.Items.Count - 1
        '    BudID.Items(i).Selected=False
        'Next
        AcctMode.SelectedIndex = -1
        PMode.SelectedIndex = -1
        Q1.SelectedIndex = -1
        For Each item As ListItem In Q2.Items
            item.Selected = False
        Next
        Q3.SelectedIndex = -1
        Q3_Other.Text = ""
        Q4.SelectedIndex = -1
        Q5.SelectedIndex = -1

        Q61.Text = ""
        Q62.Text = ""
        Q63.Text = ""
        Q64.Text = ""

        SolTR.Style.Item("display") = cst_none1
        TRNDTR.Style.Item("display") = cst_none1
        '專上學歷失業者
        rdo_HighEduBg.ClearSelection()
        rdo_HighEduBg.Attributes.Clear() '清除
        'rdo_HighEduBg.Attributes.Add("disabled", "disabled")  '專上畢業學歷失業者

        '移除 特別預算
        If BudID.Items.Count > 0 Then
            If BudID.Items.FindByValue("98") IsNot Nothing Then
                BudID.Items.Remove(BudID.Items.FindByValue("98"))
            End If
        End If

        '是否為在職者補助身分
        rblWorkSuppIdent.Enabled = True
        rblWorkSuppIdent.SelectedIndex = -1
    End Sub

    ''' <summary>取得 參訓身分別</summary>
    ''' <returns></returns>
    Function Get_All_Identity2() As String
        Dim rst As String = ""
        For i As Integer = 0 To IdentityID.Items.Count - 1
            If IdentityID.Items(i).Value <> "" AndAlso IdentityID.Items(i).Selected = True Then
                rst &= String.Concat(If(rst <> "", ",", ""), IdentityID.Items(i).Value)
            End If
        Next
        Return rst
    End Function

    Dim rqOCID As String = ""
    Dim dtArc As DataTable '暫時權限Table
    Dim dtArc2 As DataTable '暫時權限Table
    'Dim blnCanAdds As Boolean=False '新增
    'Dim blnCanMod As Boolean=False '修改
    'Dim blnCanDel As Boolean=False '刪除
    'Dim blnCanSech As Boolean=False '查詢
    'Dim blnCanPrnt As Boolean=False '列印
    Dim objconn As SqlConnection
    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub
    '載入資料
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) 'TIMS.Get_TitleLab(Request("ID"), TitleLab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call SUtl_PageInit1()
        MenuTable_td_4.Visible = TIMS.GFG_BuildinAIimage
        tr_DDL_DISASTER.Visible = TIMS.GFG_20250710A

        Call SUtl_Create0()
        If Not IsPostBack Then
            Call SUtl_Create1()
        End If
        Call SUtl_Create0bk()
    End Sub
    '載入資料(每次)
    Sub SUtl_Create0()
        'PageControler1.PageDataGrid=DataGrid1
        'ECFA 005 勞動力發展署雲嘉南分署 2024 暫不限定ECFA
        'ECFA 003 勞動力發展署桃竹苗分署 2024 暫不限定ECFA (stop 20241213)
        '(sm.UserInfo.DistID="005" OrElse sm.UserInfo.DistID="003")
        'flag_BudID_ECFA_NoLock=((sm.UserInfo.DistID="005" OrElse sm.UserInfo.DistID="003") AndAlso sm.UserInfo.Years="2024")
        '(桃分署)114年度產投公務預算不足，暫時解除不卡控預算別 (stop 20250515) (卡 20251003) 
        'flag_BudID_ECFA_NoLock = (sm.UserInfo.DistID = "003" AndAlso sm.UserInfo.Years = "2025")

        Dim flag_SHOW_2020x70 As Boolean = TIMS.SHOW_2020x70(sm)
        Dim flag_SHOW_2020x06 As Boolean = TIMS.SHOW_2020x06(sm)
        'Dim flag_show_actno_budid As Boolean=False '保險證號/預算別代碼 false:不顯示 true:顯示
        flag_show_actno_budid = False '保險證號/預算別代碼 false:不顯示 true:顯示
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_actno_budid = True
        If flag_SHOW_2020x70 Then flag_show_actno_budid = True
        If flag_SHOW_2020x06 Then flag_show_actno_budid = True
        Hid_show_actno_budid.Value = ""
        If (flag_show_actno_budid) Then Hid_show_actno_budid.Value = "Y"

        'trTPlanid28_1 '頁籤控制
        trTPlanid28_1.Visible = False '頁籤控制 false:不顯示 true:顯示
        If Hid_show_actno_budid.Value = "Y" Then
            trTPlanid28_1.Visible = True
        End If

        '暫時權限Table------Start
        ''Dim dtArc As DataTable '暫時權限Table
        dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
        dtArc2 = TIMS.Get_Auth_REndClass2(Me, objconn)
        '暫時權限Table------End

        If TIMS.sUtl_ChkTest Then gFlagEnv = False '測試用。
        Button5.Visible = False '(回上一頁)

        'rqOCID=Request("OCID")
        'If Not gFlagEnv Then rqOCID="95945"'"C120191434"
        rqOCID = TIMS.ClearSQM(Request("OCID"))
        If rqOCID = "" Then Exit Sub

        '該民眾不具失、待業身分，不得參加失業者職前訓練。STUD_SELRESULTBLIDET / STUD_SELRESULTBLI
        'dtBLIDET1=TIMS.Get_dtBLIDET1(rqOCID, objconn)

        '屆退官兵者 (依系統日期判斷)
        'flagTPlanID02Plan2=False '判斷計畫為自辦職前。
        'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagTPlanID02Plan2=True '判斷計畫為自辦職前。

        trShowDetail.Visible = True '提供基本資料供求才廠商查詢
        If TIMS.Cst_TPlanID06AppPlan1.IndexOf(sm.UserInfo.TPlanID) > -1 Then trShowDetail.Visible = False

        '該計畫是否使用ECFA True:使用 False:不使用
        blnTPlanUseEcfa = TIMS.CheckTPlanUseEcfa(sm.UserInfo.TPlanID)

        '(在職進修必填)
        'SELECT * FROM Sys_GlobalVar where gvid='22' and trim(itemvar1) is not null
        sTPlan06_G22 = TIMS.GetGlobalVar(Me, "22", "1", objconn)

        rdo_HighEduBg.Attributes.Add("onClick", "Change_BudID();")
        BudID.Attributes.Add("onClick", "if(!document.getElementById('BudID').disabled){tmpBudID='none';Change_BudID();}")

        SOCID.Attributes("onchange") = "if(this.selectedIndex==0) return false;"
        IDNO.Attributes("onblur") = "if(this.value.length==10){document.getElementById('IDNO').value=this.value.toUpperCase();}"

        '同意 本署將學員個人資料提供社家署做就業媒合之用
        trHouseMatch.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            fontGraduateY.Style.Add("color", "black") '產投非必填
            MIdentityID.Attributes.Add("onchange", "MIdentityChg(this.value);ChkMIdentityID();")
        Else
            fontGraduateY.Style.Add("color", "red") '維持紅色。
            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '同意 本署將學員個人資料提供社家署做就業媒合之用
                trHouseMatch.Visible = True
                'Dim attValue1 As String="checkNativeID();ChkMIdentityID();ChangeWorkSuppIdent();"
                Dim attValue1 As String = "MIdentityChg(this.value);ChkMIdentityID();ChangeWorkSuppIdent();"
                MIdentityID.Attributes.Add("onchange", attValue1)
            Else
                'Dim attValue1 As String="checkNativeID();ChkMIdentityID();"
                Dim attValue1 As String = "MIdentityChg(this.value);ChkMIdentityID();"
                MIdentityID.Attributes.Add("onchange", attValue1)
            End If
        End If

        Button1.Attributes("onclick") = "return chkdata();"
        Button2.Attributes("onclick") = "return chkdata();"
        EnterChannel.Attributes("onchange") = "EnterChannelChange();"
        'TRNDMode.Attributes("onchange")="TRNDModeChange();"
        If StudentID.Text <> "" Then StudentID.Attributes("onblur") = "chk_studentID(this.value,this);"
        Button4.Attributes("onclick") = "if(document.getElementById('IDNO').value==''){alert('請輸入身分證號碼');return false;}"
        'Button7.Attributes("onclick")="if(document.getElementById('ActNo1').value==''){alert('請輸入投保單位保險證號碼');return false;}"
        'Button8.Attributes("onclick")="if(document.getElementById('ActNo2').value==''){alert('請輸入投保單位保險證號碼');return false;}"

        '該計畫是否使用ECFA
        'Button7.Visible=False '檢查 ECFA鈕 消失
        ''Button8.Visible=False '檢查 ECFA鈕 消失
        'If blnTPlanUseEcfa Then '該計畫是否使用ECFA 
        '    Button7.Visible=True
        '    'Button8.Visible=True

        '    '產投與在職判斷方式
        '    If Not TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '        Dim vsStrTitle As String=""
        '        vsStrTitle=""
        '        vsStrTitle &= cst_ECFA & "基金產業被認定日" & vbCrLf
        '        vsStrTitle &= "依第1筆【受訓前任職起迄資料】結束日期" & vbCrLf
        '        TIMS.Tooltip(Button7, vsStrTitle, True)
        '        'TIMS.Tooltip(Button8, vsStrTitle, True)
        '    End If
        'End If

        PassPortNO.Attributes("onclick") = "ChangePassPort();"
        ChinaOrNot.Attributes("onclick") = "if(getRadioValue(document.form1.ChinaOrNot)==1){document.getElementById('Nationality').value='中國';}else{document.getElementById('Nationality').value='';}"
        AcctMode.Attributes("onclick") = "ChangeBank();"

        '就職狀況改變是否為在職者補助
        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            JobStateType.Attributes("onclick") = "ChangeWorkSuppIdent();"
            rblWorkSuppIdent.Attributes("onclick") = "ChangeWorkSuppIdent(2);"
        End If
        'DGTR.Style.Item("display")=cst_none1
        GovTR.Style.Item("display") = cst_none1

        '英文姓名要能自動轉為大寫
        LName.Attributes.Add("onBlur", "LName.value=LName.value.toUpperCase();")
        FName.Attributes.Add("onBlur", "FName.value=FName.value.toUpperCase();")

        '民族別欄位隱藏
        'NativeTr1.Style("display")=cst_none1

        '緊急聯絡人地址勾選檢驗
        'CheckBox2.Attributes.Add("onClick", "if(document.getElementById('CheckBox2').checked==true){ document.getElementById('CheckBox3').checked=false; }")
        'CheckBox3.Attributes.Add("onClick", "if(document.getElementById('CheckBox3').checked==true){ document.getElementById('CheckBox2').checked=false; }")
        '同通訊地址-戶籍地址
        CheckBox1.Attributes.Add("onClick", "return ock_CheckBox1();")
        '同通訊地址-緊急通知人地址
        CheckBox2.Attributes.Add("onClick", "return ock_CheckBox2();")
        '同戶籍地址-緊急通知人地址
        CheckBox3.Attributes.Add("onClick", "return ock_CheckBox3();")
        '同投保單位地址-公司地址
        Checkbox4.Attributes.Add("onClick", "return ock_CheckBox4();")

        hide_Years.Value = CInt(sm.UserInfo.Years)
        '可用補助額
        hide_3Y_SupplyMoney.Value = TIMS.Get_3Y_SupplyMoney()

        Dim flag_OnTheJob_Display As Boolean = False 'false:職前班顯示／'在職班-有些職前的選項不顯示 
        '06:自辦在職
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_OnTheJob_Display = True 'true:不顯示 
        '07:接受企業委託訓練
        If TIMS.Cst_TPlanID07.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_OnTheJob_Display = True 'true:不顯示 
        '70:區域產業據點職業訓練計畫(在職)
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_OnTheJob_Display = True 'true:不顯示 
        '28,54 產業人才投資計劃與充電起飛計畫-在職
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_OnTheJob_Display = True 'true:不顯示 

        'false:職前班顯示／'在職班-有些職前的選項不顯示 
        If flag_OnTheJob_Display Then
            '產投/在職
            star1.Visible = False
            star2.Visible = False
            star3.Visible = False
            star4.Visible = False
            star5.Visible = False
            'star6.Visible=False
            star7.Visible = False

            '若為產投下列資訊不顯示：
            '津貼類別、津貼身分別、受訓前服務單位、受訓前任職起迄日期、受訓前薪資、交通方式
            trSubsidyID.Visible = False '津貼類別
            trSubsidyIdentity.Visible = False '津貼身分別

            trPriorWorkOrg1.Visible = False '受訓前服務單位1
            trPriorWorkOrg2.Visible = False '受訓前服務單位2
            trTable6.Visible = False '受訓前任職起迄日期
            'trPriorWorkPay.Visible=False '受訓前薪資
            ActNo2_TR.Visible = False '投保單位保險證號 投保薪資級距
            trTitle1.Visible = False '職稱 受訓前薪資
            trTitle1b.Visible = False '職稱 受訓前薪資
            trTraffic.Visible = False '交通方式

            '受訓前失業週數
            'tdRealJobless.Style("display")=cst_none1
            'RealJobless.Visible=False
            'JoblessID.Visible=False
            'PModeTD.Visible=False
            'PModeTD.Style("display")=cst_none1
            trPMode.Visible = False
            PMode.Visible = False

            '就職狀況 
            '就職狀況'0:失業 1:在職 
            jobstatetd.Style("display") = cst_none1
            JobStateType.Visible = False
            '津貼類別
            SubsidyLabel.Visible = False

            '津貼 身分別
            LabSubsidy.Visible = False
            CDateTR.Visible = False
            star8.Visible = False
            star9.Visible = False
            'star10.Visible=False'MilitaryID '兵役狀況非必填
            star11.Visible = False

            '受訓前任職狀況
            PWType_TR.Style("display") = cst_none1
            ActNo2_TR.Style("display") = cst_none1

            ServDept.Visible = True
            JobTitle.Visible = True
            ddlSERVDEPTID.Visible = False
            ddlJOBTITLEID.Visible = False
            'If TestStr="AmuTest" Then gFlagEnv=False '測試用。
            If sm.UserInfo.Years >= "2016" OrElse Not gFlagEnv Then
                ServDept.Visible = False
                JobTitle.Visible = False
                ddlSERVDEPTID.Visible = True
                ddlJOBTITLEID.Visible = True
            End If
            'Sql="SELECT SERVDEPTID,SDNAME FROM KEY_SERVDEPT ORDER BY SERVDEPTID"
            'dtSERVDEPT=DbAccess.GetDataTable(Sql, objConn)
            'Sql="SELECT JOBTITLEID,JTNAME FROM KEY_JOBTITLE ORDER BY JOBTITLEID"
            'dtJOBTITLE=DbAccess.GetDataTable(Sql, objConn)

            Select Case sm.UserInfo.LID
                Case "2" '委訓單位 產學訓停用修改出生年月日 BY AMU 20151103
                    '產投 出生年月日停止修改
                    Birthday.Attributes.Add("onkeydown", "this.blur()")
                    Birthday.Attributes.Add("oncontextmenu", "return false;")
                    Birthday.Enabled = False
                    Img1.Style("display") = cst_none1 '出生日期選擇功能
                    Img1.Disabled = True
                    'hidBirthBtn.Disabled=True '失效
                    TIMS.Tooltip(Birthday, cst_msgBirth)
            End Select
        End If

        '補助比例'產投必填
        Dim str_display_NX As String = If(Hid_show_actno_budid.Value = "Y", cst_inline1, cst_none1)
        SupplyTD.Style("display") = str_display_NX
        SupplyID.Style("display") = str_display_NX

        '顯示介面--職前顯示如下
        If Not flag_OnTheJob_Display Then
            Select Case sm.UserInfo.TPlanID
                Case TIMS.Cst_TPlanID06Plan1 '"06" '06:在職進修訓練
                    '受訓前任職狀況
                    PWType_TR.Style("display") = cst_none1
                    ActNo2_TR.Style("display") = cst_none1
                    If sTPlan06_G22 <> "" Then '取消必填
                        '就職狀況 
                        If sTPlan06_G22.IndexOf("JobStateType") > -1 Then
                            jobstatetd.Style("display") = cst_none1
                            JobStateType.Visible = False
                        End If
                        '主要參訓身分別
                        If sTPlan06_G22.IndexOf("MIdentityID") > -1 Then StarMIdentityID.Visible = False
                        If sTPlan06_G22.IndexOf("SubsidyID") > -1 Then SubsidyLabel.Visible = False '津貼類別
                        If sTPlan06_G22.IndexOf("SubsidyIdentity") > -1 Then LabSubsidy.Visible = False '津貼 身分別
                        ''受訓前任職 起迄日期
                        'If sTPlan06_G22.IndexOf("SOfficeYM1") > -1 Then
                        'End If
                        ''受訓前薪資
                        'If sTPlan06_G22.IndexOf("PriorWorkPay") > -1 Then
                        'End If
                        '受訓前失業週數
                        'If sTPlan06_G22.IndexOf("RealJobless") > -1 Then
                        '    tdRealJobless.Style("display")=cst_none1
                        '    RealJobless.Visible=False
                        '    JoblessID.Visible=False
                        'End If
                    End If
                Case Else '其他(職前)
                    HistoryTable.Style("display") = cst_none1
                    Panel.Visible = False
                    msg.Visible = False
                    'star6.Visible=False
                    PWType_TR.Style("display") = cst_inline1
                    ActNo2_TR.Style("display") = cst_inline1
            End Select
        End If

        ViewState(vs_HighEduBg) = TIMS.Check_OptOptions("專上畢業學歷失業者", Convert.ToString(sm.UserInfo.TPlanID), objconn)

        'Dim iPYNum17 As Integer=1 'iPYNum17=TIMS.sUtl_GetPYNum17(Me)
        iPYNum17 = TIMS.sUtl_GetPYNum17(Me)  '若是登入年度為 2017年以後，則傳回2，其餘為1

        '2017年後 '2017職前 使用勞保明細檢查鈕
        BtnCheckBli.Visible = False
        labMsg2017a.Visible = False '經網路報名者，需重新勾選確認以下欄位
        If iPYNum17 = 2 AndAlso TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            HidPreUseLimited17f.Value = TIMS.cst_YES
            'BtnCheckBli.Visible=False
            'labMsg2017a.Visible=False '經網路報名者，需重新勾選確認以下欄位
            Button8.Visible = False
            lab2017_1.Text = "" '"任職<br />單位名稱"
            lab2017_2.Text = "" '"投保單位<br />加退保日期"
            lab2017_3.Text = "" '"投保單位<br />保險證號"
            lab2017_4.Text = "" '"投保薪資級距"
            BtnCheckBli.Visible = True
            labMsg2017a.Visible = True
            lab2017_1.Text = "任職<br />單位名稱"
            lab2017_2.Text = "投保單位<br />加退保日期"
            lab2017_3.Text = "投保單位<br />保險證號"
            lab2017_4.Text = "投保薪資級距"
            'Select Case iPYNum17
            '    Case 2 '2017年後
            '    Case 1 '2016年前
            'End Select
        Else
            '2016年前
            Button8.Visible = True
            lab2017_1.Text = "受訓前<br />服務單位"
            lab2017_2.Text = "受訓前任職<br />起迄日期"
            lab2017_3.Text = "最後投保單<br />位保險證號"
            lab2017_4.Text = "受訓前薪資"
        End If

        '2017職前 使用勞保明細檢查鈕
        'BtnCheckBli.Visible=False
        'labMsg2017a.Visible=False '經網路報名者，需重新勾選確認以下欄位
        'If TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    HidPreUseLimited17f.Value=TIMS.cst_YES
        '    BtnCheckBli.Visible=False
        '    labMsg2017a.Visible=False '經網路報名者，需重新勾選確認以下欄位
        '    Button8.Visible=False
        '    lab2017_1.Text="" '"任職<br />單位名稱"
        '    lab2017_2.Text="" '"投保單位<br />加退保日期"
        '    lab2017_3.Text="" '"投保單位<br />保險證號"
        '    lab2017_4.Text="" '"投保薪資級距"
        '    Select Case iPYNum17
        '        Case 2 '2017年後
        '            BtnCheckBli.Visible=True
        '            labMsg2017a.Visible=True
        '            lab2017_1.Text="任職<br />單位名稱"
        '            lab2017_2.Text="投保單位<br />加退保日期"
        '            lab2017_3.Text="投保單位<br />保險證號"
        '            lab2017_4.Text="投保薪資級距"
        '        Case 1 '2016年前
        '            Button8.Visible=True
        '            lab2017_1.Text="受訓前<br />服務單位"
        '            lab2017_2.Text="受訓前任職<br />起迄日期"
        '            lab2017_3.Text="最後投保單<br />位保險證號"
        '            lab2017_4.Text="受訓前薪資"
        '    End Select
        'End If

        '非職前，也非產投，應該是在職班
        'If TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID)=-1 _
        '    AndAlso TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID)=-1 Then
        '    BtnCheckBli.Visible=False
        '    labMsg2017a.Visible=False
        '    Button8.Visible=True
        '    lab2017_1.Text="受訓前<br />服務單位"
        '    lab2017_2.Text="受訓前任職<br />起迄日期"
        '    lab2017_3.Text="最後投保單<br />位保險證號"
        '    lab2017_4.Text="受訓前薪資"
        'End If
    End Sub
    '載入資料(首頁載入)(執行1次) If Not IsPostBack Then Call sUtl_Create1()
    Sub SUtl_Create1()
        TRNDMode = TIMS.Get_TRNDMode(TRNDMode)

        '專上畢業學歷失業者
        If ViewState(vs_HighEduBg) = False Then HGTR.Visible = False

        If Session(vs_SearchStr) IsNot Nothing Then
            ViewState(vs_SearchStr) = Session(vs_SearchStr)
            Session(vs_SearchStr) = ViewState(vs_SearchStr)
            'Session(vs_SearchStr)=Nothing
        End If

        '職前專用
        '2017年後 '2017職前 使用勞保明細檢查鈕
        If iPYNum17 = 2 AndAlso TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            HidPreUseLimited17f.Value = TIMS.cst_YES
            '勞保及3合1就業資料查詢 '勞保及三合一就業資料查詢(MDate)
            BtnCheckBli.Attributes("onclick") = "open_SD01001sch();return false;"
            PriorWorkType1.Attributes("onclick") = "chgPriorWorkType1();"
        End If

        '產投專用
        'BtnAutoInputActno.Visible=False
        BtnCheckBli2.Visible = False
        If Hid_show_actno_budid.Value = "Y" Then
            'BtnAutoInputActno.Visible=True
            BtnCheckBli2.Visible = True
            BtnCheckBli2.Attributes("onclick") = "open_SD01001sch2();return false;"
        End If

        Call FindRIDValue()
        TPlanID.Value = sm.UserInfo.TPlanID
        RoleID.Value = sm.UserInfo.RoleID
        SolTR.Style.Item("display") = cst_none1
        TRNDTR.Style.Item("display") = cst_none1
        Call Add_Items() '補列各資料選項 (排第1順位) '依班級資料
        '塞入班級資料 (排第2順位) '重要 '新增資料時取得開、結訓日期 依 Request("OCID")
        Call GetOpenDate2()   '塞入班級資料 (排第2順位) '重要不要移動該順位，應在 create1()之前

        Dim Req_SD_03_002 As String = TIMS.ClearSQM(Request("SD_03_002"))
        Dim Req_SD_03_002_classver As String = TIMS.ClearSQM(Request("SD_03_002_classver"))
        If Req_SD_03_002 = "VIEW" Or Req_SD_03_002_classver = "VIEW" Then
            Dim rqSOCID As String = TIMS.ClearSQM(Request("SOCID"))
            Session(cst_SearchSOCID) = rqSOCID 'Request("SOCID")
        End If

        If Session(cst_SearchSOCID) IsNot Nothing Then '處裡狀態
            Process.Value = "edit"
            Session(cst_SearchSOCID) = TIMS.ClearSQM(Session(cst_SearchSOCID))
            Call Create1_Stud(Session(cst_SearchSOCID)) '塞入學員資料 [SQL]
            Call GetScript(Session(cst_SearchSOCID)) '學員資料審核功能的欄位鎖住
            Session(cst_SearchSOCID) = Nothing
            StdTr.Visible = True
            Button2.Visible = True '(儲存2)
            Button4.Visible = False '依身分證號檢查資料
        Else
            ChinaOrNotTable.Style("display") = cst_none1
            PPNO.Style("display") = cst_none1
            ForeTr1.Style("display") = cst_none1
            ForeTr2.Style("display") = cst_none1
            ForeTr3.Style("display") = cst_none1
            ForeTr4.Style("display") = cst_none1
            ForeTr5.Style("display") = cst_none1
            PortTR.Style("display") = cst_none1
            BankTR1.Style("display") = cst_none1
            BankTR2.Style("display") = cst_none1
            BankTR3.Style("display") = cst_none1
            Process.Value = "add"
            StdTr.Visible = False
            Button2.Visible = False '(儲存2)
            Button4.Visible = True '依身分證號檢查資料
            'Dim TICKET_NO As String=TIMS.ClearSQM(Request("TICKET_NO"))
            'If TICKET_NO <> "" Then Call createDG(TICKET_NO) '取得 Adp_DGTRNData
        End If

        If Req_SD_03_002 = "VIEW" OrElse Req_SD_03_002_classver = "VIEW" Then
            Button1.Visible = False '(儲存1)
            Button2.Visible = False '(儲存2)
            Button3.Visible = False '(不儲存回上一頁)
            Button5.Visible = True '(回上一頁)
        End If

        If Req_SD_03_002 = "VIEW" Then Button5.Text = "回系統功能" '(回上一頁)
    End Sub
    '載入資料(每次)(第2段)
    Sub SUtl_Create0bk()
        '加入時數限制
        'Dim sql As String
        'Dim dtDGHR As DataTable=Get_DGTHourDT()
        ''sql="SELECT * FROM Key_DGTHour ORDER BY DGID"
        'sql="SELECT * FROM Key_DGTHour ORDER BY DGhour" 'chk RelClass_Unit
        'dtDGHR=DbAccess.GetDataTable(sql, objConn) '目前系統有4筆資料。
        Dim dtDGHR As DataTable = TIMS.Get_DGTHourDT(objconn)
        Label1.Text = dtDGHR.Rows(0)("DGHour")
        Label2.Text = dtDGHR.Rows(1)("DGHour")
        Label3.Text = dtDGHR.Rows(2)("DGHour")
        Label4.Text = dtDGHR.Rows(3)("DGHour")

        LearnTR1.Style("display") = cst_none1
        LearnTR2.Style("display") = cst_none1
        LearnTR3.Style("display") = cst_none1
        LearnTR4.Style("display") = cst_none1
        LearnTR5.Style("display") = cst_none1

        TPlan23TR.Visible = False
        MenuTable.Visible = False ''菜單
        BackTable.Visible = False ''參訓背景

        If Hid_show_actno_budid.Value = "Y" Then
            '產學訓計畫
            MenuTable.Visible = True
            BackTable.Visible = True
            If Not IsPostBack Then
                Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")

                'Page.ClientScript.RegisterStartupScript(GetType(), "11111", "ChangeMode(1);", True)
                'Page.RegisterStartupScript("11111", "<script>ChangeMode(1)</script>")
                'Page.ClientScript.RegisterClientScriptBlock(GetType(), "11111", "ChangeMode(1)", True)
                'ScriptManager.RegisterStartupScript(Page, Page.GetType(), "11111", "ChangeMode(1);", True)
                'Page.RegisterStartupScript(TIMS.GetGUID, "<script>ChangeMode(1)</script>")
                'ClientScript.RegisterStartupScript(GetType, "chgmode", "ChangeMode(1);", True)
                'Page.Common.RespWrite(Me, "<script>ChangeMode(1);</script>")

                'strScript1=""
                'strScript1 &= "<script language=""javascript"">" + vbCrLf
                'strScript1 &= "alert('系統準備中請洽系統管理者!!!!');" + vbCrLf
                'strScript1 &= "</script>"
                'Page.RegisterStartupScript(TIMS.GetGUID, strScript1)
                'DetailTable.Style.Add("display", cst_inline1)
                'BackTable.Style.Add("display", "none")
                'HistoryTable.Style.Add("display", "none")
                'ButtonTable4.Style.Add("display", cst_inline1)

                'document.getElementById('DetailTable').style.display='inline';
                '    if (document.getElementById('BackTable')) document.getElementById('BackTable').style.display='none';
                '    document.getElementById('HistoryTable').style.display='none';
                '    document.getElementById('ButtonTable4').style.display='inline';
            End If
        Else
            Select Case sm.UserInfo.TPlanID
                Case "15"  '學習卷計畫
                    LearnTR1.Style("display") = cst_inline1
                    LearnTR2.Style("display") = cst_inline1
                    LearnTR3.Style("display") = cst_inline1
                    LearnTR4.Style("display") = cst_inline1
                    LearnTR5.Style("display") = cst_inline1
                Case "23", "34", "41" '23:訓用合一, 34:與企業合作辦理職前訓練, 41:推動營造業事業單位辦理職前培訓計畫
                    TPlan23TR.Visible = True
                Case Else
                    '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                    If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        MenuTable.Visible = True
                        BackTable.Visible = True
                        If Not IsPostBack Then
                            Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")
                            DetailTable.Style.Add("display", cst_inline1)
                            BackTable.Style.Add("display", cst_none1)
                            HistoryTable.Style.Add("display", cst_none1)
                            ButtonTable4.Style.Add("display", cst_inline1)
                        End If
                    End If
            End Select
        End If

        'If OpenDate.Text <> "" AndAlso FOfficeYM1.Text <> "" Then Page.RegisterStartupScript("getDiffDate", "<script language='javascript'>getDiffDate();</script>")
        'OpenDate.Attributes("onpropertychange")="javascript:getDiffDate();"  '提示失業週數 
        'FOfficeYM1.Attributes("onpropertychange")="javascript:getDiffDate();" '受訓前任職 起迄日期

        '\TIMS.NET40o\js\OpenWin\openwin.js
        'openCalendar(Birthday, '1911/01/01', '2099/12/31', Date(), '', ButtonID);
        Img1.Attributes("onclick") = "callCalendar('" & Birthday.ClientID & "','" & hidBirthBtn.ClientID & "');"

        IMG2.Attributes("onclick") = "javascript:show_calendar('" & SOfficeYM1.ClientID & "','','','CY/MM/DD');"
        IMG3.Attributes("onclick") = "javascript:show_calendar('" & FOfficeYM1.ClientID & "','','','CY/MM/DD');"

        hidBirthBtn.Attributes("onclick") = "javascript:ChkBirthday();" '檢查 ChkBirthday 中高齡者(45歲)
        Birthday.Attributes("onchange") = "javascript:ChkBirthday();"
        Birthday.Attributes("onblur") = "javascript:ChkBirthday();"
        IDNO.Attributes("onchange") = "javascript:chkidnosex();"
        IDNO.Attributes("onblur") = "javascript:chkidnosex();"

        'Dim vSOCID_Sel As String=TIMS.ClearSQM(SOCID.SelectedValue)
        Dim vSOCID_Sel As String = TIMS.GetListValue(SOCID)
        '勞工團體顯示項目-訓練單位代轉現金；非勞工團體則不用顯示
        Dim OrgKind2 As String = ""
        '產學訓的編輯狀態下，才去搜尋機構別
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If Process.Value = "edit" Then OrgKind2 = TIMS.Get_OrgKind2(vSOCID_Sel, TIMS.c_SOCID, objconn)
        End If

        '只要不是產學訓且非勞工團體時，都不顯示訓練單位代轉現金這個項目
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If OrgKind2 <> "W" Then AcctMode.Items.Remove(AcctMode.Items.FindByValue(2)) 'W:提升在職勞工自主學習計畫 'G:產業人才投資計畫
        Else
            AcctMode.Items.Remove(AcctMode.Items.FindByValue(2))
        End If

        '3+2郵遞區號查詢 link
        LitZipCode4.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCode1.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCode2.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCode3.Text = TIMS.Get_WorkZIPB3Link2()
        LitForeZip.Text = TIMS.Get_WorkZIPB3Link2()
        LitActZip.Text = TIMS.Get_WorkZIPB3Link2()
        LitZip.Text = TIMS.Get_WorkZIPB3Link2()

        '輸入郵遞區號3碼判斷 & 代入 CityName 
        ZipCode4.Attributes.Add("onblur", "getZipName('City4',this,this.value);")
        ZipCode1.Attributes.Add("onblur", "getZipName('City1',this,this.value);")
        ZipCode2.Attributes.Add("onblur", "getZipName('City2',this,this.value);")
        ZipCode3.Attributes.Add("onblur", "getZipName('City3',this,this.value);")
        ForeZip.Attributes.Add("onblur", "getZipName('City6',this,this.value);")
        txt_ActZip.Attributes.Add("onblur", "getZipName('txt_ActCity',this,this.value);")
        Zip.Attributes.Add("onblur", "getZipName('City5',this,this.value);")

        '查詢郵遞區號button控制
        Dim bt1_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode4, ZipCode4_B3, hidZipCode4_6W, City4, hidCityName4, hidAREA4, ZipCode4_N, ServiceAddress)
        bt_openZip1.Attributes.Add("onclick", bt1_Attr_VAL)
        Dim bt2_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode1, ZipCode1_B3, hidZipCode1_6W, City1, hidCityName1, hidAREA1, ZipCode1_N, Address)
        bt_openZip2.Attributes.Add("onclick", bt2_Attr_VAL)
        Dim bt3_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode2, ZipCode2_B3, hidZipCode2_6W, City2, hidCityName2, hidAREA2, ZipCode2_N, HouseholdAddress)
        bt_openZip3.Attributes.Add("onclick", bt3_Attr_VAL)
        Dim bt4_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ZipCode3, ZipCode3_B3, hidZipCode3_6W, City3, hidCityName3, hidAREA3, ZipCode3_N, EmergencyAddress)
        bt_openZip4.Attributes.Add("onclick", bt4_Attr_VAL)
        Dim bt5_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(ForeZip, ForeZIPB3, hidForeZIP6W, City6, hidCityNameFore, hidAREAFore, ForeZip_N, ForeAddr)
        bt_openZip5.Attributes.Add("onclick", bt5_Attr_VAL)
        Dim bt6_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(txt_ActZip, txt_ActZIPB3, hid_ActZIP6W, txt_ActCity, hidCityNameAct, hidAREAAct, hidActZip_N, txt_ActAddress)
        bt_openZip6.Attributes.Add("onclick", bt6_Attr_VAL)
        Dim bt7_Attr_VAL As String = TIMS.GET_ZipCodeWOAspxUrl(Zip, ZIPB3, hidZIP6W, City5, hidCityName7, hidAREA7, Zip_N, Addr)
        bt_openZip7.Attributes.Add("onclick", bt7_Attr_VAL)

        '-----for 個資保護------
        Dim Req_SM As String = TIMS.ClearSQM(Request("SM"))
        Call Std_Data_Mask(Req_SM)
    End Sub

    ''' <summary>增修需求 OJT-121201_系統_產投_學員資料維護_外籍補助80%+預算別與補助比例連動調整</summary>
    ''' <param name="in_parms"></param>
    ''' <param name="sErrmsg1"></param>
    ''' <returns></returns>
    Public Shared Function SUtl_CheckData2(ByRef in_parms As Hashtable, ByRef sErrmsg1 As String) As Boolean
        Dim rst As Boolean = True '於儲存 系統檢核-true:OK / false:有錯誤資訊 

        Const cst_MID_因應貿易自由化協助勞工_30 As String = "30"
        Const cst_MID_經公告之重大災害受災者_40 As String = "40"
        Const cst_BUDID_97_公務ECFA As String = "97"
        '身分別 PassPortNO 1:本國 /2:外籍(含大陸人士)
        Dim v_PassPortNO As String = TIMS.GetMyValue2(in_parms, "PassPortNO")
        '主要參訓身分別 SELECT * FROM KEY_IDENTITY WHERE IdentityID IN ('30','40') 
        Dim v_MIdentityID As String = TIMS.GetMyValue2(in_parms, "MIdentityID")
        '預算別 01:公務;02:就安;03:就保;04:再出發;97:公務(ECFA);98:特別預算;99:不補助--SELECT ';'+BUDID+':'+BUDNAME FROM VIEW_BUDGET
        Dim v_BudID As String = TIMS.GetMyValue2(in_parms, "BudID")
        '補助比例 1:80%;2:100%;9:不補助 --SELECT ';'+SUPPLYID+':'+SNAME FROM VIEW_SUPPLYID
        Dim v_SupplyID As String = TIMS.GetMyValue2(in_parms, "SupplyID")
        Dim v_TPlanID As String = TIMS.GetMyValue2(in_parms, "TPlanID") 'v_TPlanID

        '主要參訓身分別為「因應貿易自由化協助勞工」，預算別只能為 不補助/ECFA
        '70:區域產業據點，只有1種預算別(就安)，沒有補助比例欄位,要移除掉身分別與預算別、補助比例之間的卡控
        Dim flag_PLAN_1 As Boolean = ((v_MIdentityID = cst_MID_因應貿易自由化協助勞工_30) AndAlso (TIMS.Cst_TPlanID70.IndexOf(v_TPlanID) = -1)) '案例1
        Dim flag_PLAN_2 As Boolean = ((v_MIdentityID = cst_MID_經公告之重大災害受災者_40) AndAlso (TIMS.Cst_TPlanID70.IndexOf(v_TPlanID) = -1)) '案例2
        Dim flag_PLAN_3 As Boolean = ((v_BudID = cst_BUDID_97_公務ECFA) AndAlso (TIMS.Cst_TPlanID70.IndexOf(v_TPlanID) = -1)) '案例3

        'Dim sErrmsg1 As String=""
        If sErrmsg1 <> "" Then Return False

        If v_PassPortNO = "" Then sErrmsg1 &= "請選擇 身分別!" & vbCrLf
        If v_MIdentityID = "" Then sErrmsg1 &= "請選擇 主要參訓身分別!" & vbCrLf
        If v_BudID = "" Then sErrmsg1 &= "請選擇 預算別!" & vbCrLf
        If v_SupplyID = "" OrElse v_SupplyID = "0" Then sErrmsg1 &= "請選擇 補助比例!" & vbCrLf
        If sErrmsg1 <> "" Then Return False

        If v_BudID.Equals("99") AndAlso Not v_SupplyID.Equals("9") Then
            sErrmsg1 &= "預算別為不補助，補助比例有誤!(應為不補助)" & vbCrLf
        Else
            If v_SupplyID.Equals("9") AndAlso Not v_BudID.Equals("99") Then
                sErrmsg1 &= "補助比例為不補助，預算別有誤!(應為不補助)" & vbCrLf
            End If
        End If
        If sErrmsg1 <> "" Then Return False

        If v_PassPortNO = "2" AndAlso Not flag_PLAN_1 AndAlso Not flag_PLAN_2 Then
            If v_SupplyID.Equals("2") Then '1:80%;2:100%;9:不補助 
                sErrmsg1 &= "身分別為「外籍(含大陸人士)」，補助比例有誤!(不可為100%)" & vbCrLf
            End If
        End If
        If sErrmsg1 <> "" Then Return False

        If flag_PLAN_1 Then
            Dim flag_PLAN_1_OK As Boolean = True
            If v_BudID.Equals("02") OrElse v_BudID.Equals("03") Then flag_PLAN_1_OK = False
            If Not flag_PLAN_1_OK Then
                '只能為 97:公務(ECFA) /99:不補助
                sErrmsg1 &= "主要參訓身分別為「因應貿易自由化協助勞工」，預算別有誤!(不可為就安或就保)" & vbCrLf
            Else
                If v_SupplyID.Equals("1") Then '1:80% (不可為80%)
                    sErrmsg1 &= "主要參訓身分別為「因應貿易自由化協助勞工」，補助比例有誤!(不可為80%)" & vbCrLf
                End If
            End If
            If sErrmsg1 <> "" Then Return False
        End If

        If flag_PLAN_2 Then
            Dim flag_PLAN_2_OK As Boolean = True
            If v_BudID.Equals("97") Then flag_PLAN_2_OK = False
            If Not flag_PLAN_2_OK Then
                '只能為 02:就安/ 03:就保/ 99:不補助
                sErrmsg1 &= "主要參訓身分別為「經公告之重大災害受災者」，預算別有誤!(不可為公務(ECFA))" & vbCrLf
            Else
                If v_SupplyID.Equals("1") Then '1:80% (不可為80%)
                    sErrmsg1 &= "主要參訓身分別為「經公告之重大災害受災者」，補助比例有誤!(不可為80%)" & vbCrLf
                End If
            End If
            If sErrmsg1 <> "" Then Return False
        End If

        If flag_PLAN_3 AndAlso v_MIdentityID <> cst_MID_因應貿易自由化協助勞工_30 Then
            sErrmsg1 &= "預算別為 公務(ECFA)，主要參訓身份別 須為「因應貿易自由化協助勞工」，主要參訓身分別有誤!" & vbCrLf
            If sErrmsg1 <> "" Then Return False
        End If

        Return rst
    End Function

    'sUtl_CheckData1 儲存前先檢查輸入資料的正確性
    ''' <summary>
    ''' 儲存前先檢查輸入資料的正確性
    ''' </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_CheckData1(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = True '沒有異常為True
        'flgROLEIDx0xLIDx0  : false: 非 ROLEID=0 LID=0/ true: ROLEID=0 LID=0 IsSuperUser
        Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = If(TIMS.IsSuperUser(Me, 1), True, False) '判斷登入者的權限。

        Name.Text = TIMS.ClearSQM(Name.Text)
        RMPNAME.Text = TIMS.ClearSQM(RMPNAME.Text)
        LName.Text = TIMS.ClearSQM(LName.Text)
        FName.Text = TIMS.ClearSQM(FName.Text)

        ZipCode1.Value = TIMS.ClearSQM(ZipCode1.Value)
        ZipCode1_B3.Value = TIMS.ClearSQM(ZipCode1_B3.Value)
        ZipCode1_N.Value = TIMS.ClearSQM(ZipCode1_N.Value)
        Address.Text = TIMS.ClearSQM(Address.Text)

        ZipCode2.Value = TIMS.ClearSQM(ZipCode2.Value)
        ZipCode2_B3.Value = TIMS.ClearSQM(ZipCode2_B3.Value)
        ZipCode2_N.Value = TIMS.ClearSQM(ZipCode2_N.Value)
        HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)

        ZipCode3.Value = TIMS.ClearSQM(ZipCode3.Value)
        ZipCode3_B3.Value = TIMS.ClearSQM(ZipCode3_B3.Value)
        ZipCode3_N.Value = TIMS.ClearSQM(ZipCode3_N.Value)
        EmergencyAddress.Text = TIMS.ClearSQM(EmergencyAddress.Text)

        '身分證驗証
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        Dim aIDNO As String = IDNO.Text

        '通訊地址前3碼郵遞區號
        'If ZipCode1.Value="" Then'End If
        Dim sql As String = "SELECT * FROM dbo.VIEW_ZIPNAME ORDER BY ZIPCODE"
        Dim dtZipCode As DataTable = DbAccess.GetDataTable(sql, objconn)
        TIMS.CheckValeuErr(ZipCode1.Value, "通訊地址前3碼郵遞區號", True, "ZipCode", dtZipCode, Errmsg)
        TIMS.CheckZipCODEB3(ZipCode1_B3.Value, "通訊地址郵遞區號後2碼或後3碼", True, Errmsg)
        TIMS.CheckVal(Address.Text, "通訊地址", Errmsg)

        '同通訊地址。
        If Not CheckBox1.Checked Then
            TIMS.CheckValeuErr(ZipCode2.Value, "戶籍地址前3碼郵遞區號", True, "ZipCode", dtZipCode, Errmsg)
            TIMS.CheckZipCODEB3(ZipCode2_B3.Value, "戶籍地址郵遞區號後2碼或後3碼", True, Errmsg)
            TIMS.CheckVal(HouseholdAddress.Text, "戶籍地址", Errmsg)
        End If

        If SOfficeYM1.Text <> "" Then SOfficeYM1.Text = TIMS.ClearSQM(SOfficeYM1.Text)
        If FOfficeYM1.Text <> "" Then FOfficeYM1.Text = TIMS.ClearSQM(FOfficeYM1.Text)
        Call TIMS.CheckDateErr(SOfficeYM1.Text, "受訓前任職起迄日期的起日", False, Errmsg)
        Call TIMS.CheckDateErr(FOfficeYM1.Text, "受訓前任職起迄日期的迄日", False, Errmsg)

        If SOfficeYM2.Text <> "" Then SOfficeYM2.Text = TIMS.ClearSQM(SOfficeYM2.Text)
        If FOfficeYM2.Text <> "" Then FOfficeYM2.Text = TIMS.ClearSQM(FOfficeYM2.Text)
        Call TIMS.CheckDateErr(SOfficeYM2.Text, "第2筆 受訓前任職起迄日期的起日", False, Errmsg)
        Call TIMS.CheckDateErr(FOfficeYM2.Text, "第2筆 受訓前任職起迄日期的迄日", False, Errmsg)

        '(受訓前任職狀況)有顯示，才執行檢核
        If Convert.ToString(PWType_TR.Style("display")) = cst_inline1 Then
            Select Case sm.UserInfo.TPlanID
                Case "06" '排除 06:在職進修訓練
                Case Else
                    'rblWorkSuppIdent 
                    Dim v_PriorWorkType1 As String = TIMS.GetListValue(PriorWorkType1)
                    Dim v_rblWorkSuppIdent As String = TIMS.GetListValue(rblWorkSuppIdent) 'rblWorkSuppIdent.SelectedValue 
                    Select Case v_PriorWorkType1'PriorWorkType1.SelectedValue
                        Case "1" '曾工作過
                            PriorWorkOrg1.Text = TIMS.ClearSQM(PriorWorkOrg1.Text)
                            If PriorWorkOrg1.Text = "" Then Errmsg &= "請輸入受訓服務單位！" & vbCrLf
                            If SOfficeYM1.Text = "" Then Errmsg &= "請輸入受訓前任職起迄日期的起日！" & vbCrLf
                            If v_rblWorkSuppIdent <> "Y" Then
                                If FOfficeYM1.Text = "" Then Errmsg &= "請輸入受訓前任職起迄日期的迄日！" & vbCrLf
                            End If

                            'If ErrMessage="" Then
                            '    Try
                            '        SOfficeYM1.Text=CDate(SOfficeYM1.Text).ToString("yyyy/MM/dd")
                            '    Catch ex As Exception
                            '        ErrMessage &= "受訓前任職起迄日期的起日 格式有誤(yyyy/MM/dd)!" & vbCrLf
                            '    End Try
                            '    Try
                            '        FOfficeYM1.Text=CDate(FOfficeYM1.Text).ToString("yyyy/MM/dd")
                            '    Catch ex As Exception
                            '        ErrMessage &= "受訓前任職起迄日期的迄日 格式有誤(yyyy/MM/dd)!" & vbCrLf
                            '    End Try
                            'End If

                        Case "2" '未曾工作過
                        Case "3" '先前從事為非勞保性質工作
                        Case "4"
                        Case Else
                            Errmsg &= "請選擇受訓前任職狀況！" & vbCrLf
                    End Select
            End Select
        End If
        'RealJobless'受訓前失業週數
        'RealJobless.Text=TIMS.ClearSQM(RealJobless.Text)
        'If RealJobless.Text <> "" AndAlso RealJobless.Text <> "0" Then
        '    If Not TIMS.IsNumeric2(RealJobless.Text) Then Errmsg &= "受訓前失業週數，請輸入 正整數數字！" & RealJobless.Text & vbCrLf
        'End If

        Dim v_DDL_DISASTER As String = TIMS.GetListValue(DDL_DISASTER) 'ADID 重大災害選項
        Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
        Dim v_BudID As String = TIMS.GetListValue(BudID) 'BudID.SelectedValue
        Dim v_SupplyID As String = TIMS.GetListValue(SupplyID) 'SupplyID.SelectedValue

        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If flag_show_actno_budid Then
            '產投
            ActNo1.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(ActNo1.Text))
            If ActNo1.Text <> "" Then
                '投保單位保險證號為09開頭者，為訓字保，亦不可報名 '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
                '01、04、05、15及08者其補助經費來源歸屬為 03:就保基金 '02、03、06、07者其經費來源歸屬為 02:就安基金 '09與無法辨視者為 99:不予補助對象
                Select Case Left(ActNo1.Text, 2)
                    Case "09"
                        Errmsg &= "學員資格 投保單位保險證號 為09開頭者為訓字保 不符合可參訓條件！" & vbCrLf
                End Select
            End If

            '檢測此學員是否 可參訓 產業人才投資方案 (大於15歲者)
            If Not TIMS.Check_YearsOld15(Birthday.Text, Convert.ToString(ViewState(vs_STDate))) Then Errmsg &= "學員資格 年齡不滿15歲 不符合可參訓條件！" & vbCrLf

            'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
            If v_SupplyID = "" OrElse v_SupplyID = "0" Then Errmsg &= "請選擇補助比例!" & vbCrLf

            '產投 儲存 預算別(配含補助比例顯示)
            Select Case v_BudID'BudID.SelectedValue
                Case "99"
                    '預算別為不補助99
                    Select Case v_SupplyID'SupplyID.SelectedValue
                        Case "9" '不補助 0%
                        Case Else '"1", "2"'補助比例大於0
                            Errmsg &= cst_Msg9 & vbCrLf
                    End Select
                Case "97" 'ECFA
                    If v_SupplyID <> "2" Then Errmsg &= cst_Msg3b & vbCrLf '特定對象 100%"
                Case "04" '再出發
                    If v_SupplyID <> "2" Then Errmsg &= cst_Msg8 & vbCrLf '特定對象 100%"
                Case Else
                    '其他預算別，不可為一般身分。
                    'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
                    Dim flag_IdentityID_01 As Boolean = False
                    For i As Integer = 0 To IdentityID.Items.Count - 1
                        If IdentityID.Items(i).Selected AndAlso IdentityID.Items(i).Value = "01" Then
                            flag_IdentityID_01 = True
                            Exit For
                        End If
                    Next
                    If v_MIdentityID = "01" OrElse flag_IdentityID_01 Then
                        '身分別為一般身分，補助比例 不可為特定100%!
                        If v_SupplyID = "2" Then Errmsg &= cst_Msg3c & vbCrLf
                    End If
                    '2.主要參訓身分別屬「一般身分者」，補助比例欄位可儲存「一般對象80%」或不補助」
                    '，若非屬前述邏輯，不可被儲存，儲存時，請顯示告警訊息「主要參訓身分別屬「一般身分」
                    '，補助比例應為80%。」，直到選對才可儲存。
                    '3.主要參訓身分別非屬「一般身分者」，補助比例欄位可儲存「特定對象100%」或「不補助」
                    '，若非屬前述邏輯，不可被儲存，儲存時，請顯示告警訊息「主要參訓身分別屬「特定對象」
                    '，補助比例應為100%。」，直到選對才可儲存。
                    Select Case v_MIdentityID'MIdentityID.SelectedValue
                        Case "01"
                            'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 不補助0%
                            Select Case v_SupplyID'.SelectedValue
                                Case "1", "9"
                                Case Else
                                    Errmsg &= cst_Msg4a & vbCrLf
                            End Select
                        Case Else
                            'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 不補助0%
                            Select Case v_SupplyID'.SelectedValue
                                Case "2", "9"
                                Case Else
                                    Errmsg &= cst_Msg4b & vbCrLf
                            End Select
                    End Select
            End Select
        Else
            'TIMS 儲存 驗證 (注意：沒有 SupplyID 補助比例選項)
            If IDNO.Text <> "" AndAlso rqOCID <> "" Then
                '排除學習卷(15)，採自動匯入方式 2009-1-6 penny 
                Select Case sm.UserInfo.TPlanID
                    Case "15"
                    Case Else
                        'If TestStr="AmuTest" Then gFlagEnv=False '測試用。
                        If gFlagEnv Then '正式環境。(測試用) / TestStr
                            'stella add 2007/11/02判斷是否已有報名資料
                            IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
                            'Dim rqOCID As String=Request("OCID") 'rqOCID=TIMS.ClearSQM(rqOCID)
                            Dim SingUp As Boolean = TIMS.CheckIfSingUp(IDNO.Text, rqOCID, 1, objconn)
                            If Not SingUp Then Errmsg &= "無報名(轉班)資料，請先輸入此學員之報名(轉班)資料！" & vbCrLf
                        End If
                End Select
            Else
                Errmsg &= "基本資料不全，請重新操作學員資料維護作業！" & vbCrLf
            End If
        End If

        '更新MakeSOCID
        'Dim vRejectSOCID As String=TIMS.ClearSQM(RejectSOCID.SelectedValue)
        'Dim vSOCID_Sel As String=TIMS.ClearSQM(SOCID.SelectedValue)
        Dim vRejectSOCID As String = TIMS.GetListValue(RejectSOCID)
        Dim vSOCID_Sel As String = TIMS.GetListValue(SOCID)
        hide_MakeSOCID.Value = TIMS.ClearSQM(hide_MakeSOCID.Value)
        If hide_MakeSOCID.Value <> "" AndAlso vRejectSOCID <> "" Then
            If hide_MakeSOCID.Value = vRejectSOCID Then Errmsg &= "遞補者與被遞補者不可相同！" & vbCrLf
        End If

        If vRejectSOCID <> "" Then
            Dim vMakeSOCID As String = TIMS.GetMakeSOCID(vRejectSOCID, objconn)
            If vMakeSOCID <> "" Then
                If vMakeSOCID <> vSOCID_Sel Then Errmsg &= "該遞補者已被其他學員使用！" & vbCrLf
            End If
            If vRejectSOCID = vSOCID_Sel Then Errmsg &= "被遞補者與學員不可相同！" & vbCrLf
        End If

        v_BudID = TIMS.GetListValue(BudID)
        If v_BudID = "" Then Errmsg &= "預算別(經費來源別) 不可為空未選擇！" & vbCrLf

        Dim vsFOfficeYM1 As String = ""
        If TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投與在職判斷方式
            vsFOfficeYM1 = ""
        Else
            If v_BudID = "97" Then
                '第1筆受訓前任職起迄資料 迄止日期未填寫
                FOfficeYM1.Text = TIMS.ClearSQM(FOfficeYM1.Text)
                If FOfficeYM1.Text <> "" Then
                    Try
                        FOfficeYM1.Text = CDate(FOfficeYM1.Text).ToString("yyyy/MM/dd")
                    Catch ex As Exception
                        Errmsg &= "預算別選" & cst_ECFA & "："
                        Errmsg &= "第1筆受訓前任職起迄資料 迄止日期為必填 格式有誤(yyyy/MM/dd)!" & vbCrLf
                    End Try
                Else
                    Errmsg &= "預算別選" & cst_ECFA & "："
                    Errmsg &= "第1筆受訓前任職起迄資料 迄止日期為必填 未填寫!" & vbCrLf
                End If
                vsFOfficeYM1 = FOfficeYM1.Text
            End If
        End If

        If Errmsg = "" Then
            '投保證號更新
            ActNo1.Text = TIMS.ClearSQM(ActNo1.Text)
            ActNo2.Text = TIMS.ClearSQM(ActNo2.Text)
            If ActNo1.Text <> "" Then ActNo1.Text = TIMS.ChangeIDNO(ActNo1.Text)
            If ActNo2.Text <> "" Then ActNo2.Text = TIMS.ChangeIDNO(ActNo2.Text)
            'If ActNo1.Text = "" Then ActNo1.Text = ActNo2.Text
            If ActNo1.Text = "" AndAlso v_BudID = "97" Then
                Errmsg &= String.Concat("預算別選", cst_ECFA, "：", cst_Msg6, vbCrLf)
            End If

            '檢查保險證號是否為ECFA
            If ActNo1.Text <> "" Then
                STDateHidden.Value = TIMS.ClearSQM(STDateHidden.Value)
                STDateHidden.Value = If(STDateHidden.Value <> "", STDateHidden.Value, TIMS.Cdate3(Now.ToString("yyyy/MM/dd")))
                '2011/4/15日'產投、充電起飛計畫公告為 4/15日後才可使用" & cst_ECFA & "基金 
                '該計畫是否使用ECFA True:使用 False:不使用 BY AMU 20110929
                If blnTPlanUseEcfa Then
                    '產投相關計畫執行此功能。
                    If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                        '如果是分署(中心)承辦人，預算別不鎖定。by AMU 20140328 (本功能是每次執行)
                        flag_BudIDNoLock = False 'true: nolock false: lock
                        '什麼時候可以修改預算別。'iBudFlag - OUT: 'iBudFlag :0,1,2: 'iFlag :0:未開放 1:21天內修改 2:開放被登功能
                        flag_BudIDNoLock = Chk_CanEditBudgetID(sm.UserInfo.LID, CDate(STDateHidden.Value), iBudFlag)

                        If DateDiff(DateInterval.Day, CDate(cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
                            Select Case v_BudID 'BudID.SelectedValue
                                Case "99", "04" '(預算別) 排除使用 ECFA
                                    'BudID 99:不予補助對象  'BudID 04:再出發對象  
                                Case Else
                                    'ECFA  005	勞動力發展署雲嘉南分署	2024 暫不限定ECFA
                                    If Not flag_BudID_ECFA_NoLock Then
                                        '檢驗是否可用ECFA，並限定使用ECFA
                                        If TIMS.CheckIsECFA(Me, ActNo1.Text, vsFOfficeYM1, STDateHidden.Value, objconn) = True Then
                                            '若是鎖定，則做檢查ECFA 若 開放 則不做判斷
                                            If Not flag_BudIDNoLock Then
                                                If v_BudID <> "97" OrElse v_SupplyID <> "2" Then Errmsg &= cst_Msg3 & vbCrLf
                                            End If
                                        Else
                                            If v_BudID = "97" Then Errmsg &= cst_Msg5 & vbCrLf
                                        End If
                                    End If
                            End Select
                        Else
                            If v_BudID = "97" Then Errmsg &= cst_Msg7 & vbCrLf
                        End If
                    Else
                        If v_BudID = "97" Then
                            If Not TIMS.CheckIsECFA(Me, ActNo1.Text, vsFOfficeYM1, STDateHidden.Value, objconn) Then Errmsg &= cst_Msg5 & vbCrLf
                        End If
                    End If
                End If
                ' 該計畫是否使用ECFA 
            End If
        End If

        Dim v_SubsidyIdentity As String = TIMS.GetListValue(SubsidyIdentity)
        If Errmsg = "" Then
            If Not TIMS.Check_YearsOld45(Birthday.Text, Convert.ToString(ViewState(vs_STDate))) Then
                '04	中高齡者
                If v_MIdentityID = "04" Then Errmsg &= cst_errMsg3 & vbCrLf
                For i As Integer = 0 To IdentityID.Items.Count - 1
                    If IdentityID.Items(i).Selected = True Then
                        If IdentityID.Items(i).Value = "04" Then
                            Errmsg &= cst_errMsg4 & vbCrLf
                            Exit For
                        End If
                    End If
                Next
                If v_SubsidyIdentity = "04" Then Errmsg &= cst_errMsg5 & vbCrLf
            End If
        End If

        '檢測此學員是否 屬於六十五歲以上者資格 65歲以上 BY AMU 20121212
        If Errmsg = "" Then
            If Not TIMS.Check_YearsOld65(Birthday.Text, Convert.ToString(ViewState(vs_STDate))) Then
                If v_MIdentityID = "37" Then Errmsg &= cst_errMsg6 & vbCrLf '37:六十五歲以上者資格
                For i As Integer = 0 To IdentityID.Items.Count - 1
                    If IdentityID.Items(i).Selected = True Then
                        If IdentityID.Items(i).Value = "37" Then
                            Errmsg &= cst_errMsg7 & vbCrLf
                            Exit For
                        End If
                    End If
                Next
                If v_SubsidyIdentity = "37" Then Errmsg &= cst_errMsg8 & vbCrLf
            End If
        End If

        'If Errmsg="" Then
        '    If flagTPlanID02Plan2 Then
        '        '屆退官兵者 (依開訓日期(系統日期)判斷)
        '        If Not TIMS.CheckRESOLDER(objconn, IDNO.Text, sm.UserInfo.DistID, ViewState(vs_STDate)) Then
        '            If v_MIdentityID="12" Then Errmsg &= cst_errMsg9 & vbCrLf
        '        End If
        '    End If
        'End If

        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)
        'PassPortNO 1:本國 /2:外籍(含大陸人士)
        Select Case v_PassPortNO
            Case "1", "2"
            Case Else
                Errmsg &= "請選擇身分別" & vbCrLf 'v_PassPortNO="2"
        End Select

        If aIDNO = "" Then
            Errmsg &= "必須填寫身分證號碼" & vbCrLf
        Else
            'PassPortNO 1:本國 /2:外籍(含大陸人士)
            Select Case v_PassPortNO'aPassPortNO
                Case "2" '身分別為外籍 
                    Dim v_PPNO As String = TIMS.GetListValue(PPNO) '1:護照號碼 2:居留(工作)證號
                    Select Case v_PPNO
                        Case "1" '護照號碼
                            If Not TIMS.CheckIDNO2(aIDNO, 3) Then Errmsg &= "護照號碼錯誤!請聯絡系統管理員" & vbCrLf '一般驗証
                        Case "2" '2:居留(工作)證號
                            '2:居留證 4:居留證2021
                            Dim flag2 As Boolean = TIMS.CheckIDNO2(aIDNO, 2)
                            Dim flag4 As Boolean = TIMS.CheckIDNO2(aIDNO, 4)
                            If Not flag2 AndAlso Not flag4 Then Errmsg &= "居留證號碼錯誤!請聯絡系統管理員" & vbCrLf '一般驗証
                        Case Else
                            Errmsg &= "請選擇護照或居留(工作)證號" & vbCrLf
                    End Select
                Case "1"
                    If TIMS.CheckIDNO(aIDNO) Then '一般驗証
                        'If sm.UserInfo.RoleID=1 Then '角色代碼為1 可執行安全性規則確認
                        'End If
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
                        If Not IDNOFlag Then Errmsg &= "身分證號碼錯誤!請聯絡系統管理員" & vbCrLf
                    Else
                        Errmsg &= "身分證號碼錯誤!請聯絡系統管理員" & vbCrLf
                    End If
                Case Else
                    'Errmsg &= "請選擇身分別" & vbCrLf
            End Select
        End If

        'OK_IDNO()
        Dim v_Sex As String = TIMS.GetListValue(Sex)
        Select Case v_Sex'v_Sex
            Case "F", "M"
            Case Else
                Errmsg &= "請選擇性別" & vbCrLf
        End Select

        If Errmsg = "" Then
            If Not TIMS.CheckMemberSex(IDNO.Text, v_Sex) Then Errmsg &= "依身分證號判斷 性別選項 不正確！" & vbCrLf
        End If

        If CloseDate.Text <> "" Then CloseDate.Text = TIMS.Cdate3(CloseDate.Text)
        If EnterDate.Text <> "" Then EnterDate.Text = TIMS.Cdate3(EnterDate.Text)
        If EnterDate.Text <> "" AndAlso CloseDate.Text <> "" Then
            If DateDiff(DateInterval.Day, CDate(CloseDate.Text), CDate(EnterDate.Text)) > 0 Then
                Errmsg &= cst_errMsg10 & vbCrLf
                Return False
            End If
        End If
        If EnterDate.Text <> "" AndAlso FTDateHidden.Value <> "" Then
            If DateDiff(DateInterval.Day, CDate(FTDateHidden.Value), CDate(EnterDate.Text)) > 0 Then
                Errmsg &= cst_errMsg11 & vbCrLf
                Return False
            End If
        End If

        Dim all_Identity2 As String = "" '判斷使用
        '參訓身分別 先儲存位置
        all_Identity2 = Get_All_Identity2()

        'If flagTPlanID02Plan2 Then
        '    '屆退官兵者 (依開訓日期(系統日期)判斷)
        '    If TIMS.CheckRESOLDER(objconn, IDNO.Text, sm.UserInfo.DistID, ViewState(vs_STDate)) Then
        '        If all_Identity2="" Then Errmsg &= "此訓練學員為屆退官兵，請於參訓身分別勾選！" & vbCrLf
        '        If all_Identity2 <> "" Then
        '            If all_Identity2.IndexOf("12")=-1 Then Errmsg &= "此訓練學員為屆退官兵，請於參訓身分別勾選！" & vbCrLf
        '        End If
        '    Else
        '        If all_Identity2 <> "" Then
        '            If all_Identity2.IndexOf("12") > -1 Then Errmsg &= "此訓練學員不為屆退官兵，參訓身分別不可勾選！" & vbCrLf
        '        End If
        '    End If
        'End If

        If all_Identity2 <> "" Then
            '06:身心障礙者
            Dim v_rblHandType As String = TIMS.GetListValue(rblHandType)
            Dim v_HandLevelID2 As String = TIMS.GetListValue(HandLevelID2)
            Dim v_HandTypeID As String = TIMS.GetListValue(HandTypeID)
            Dim v_HandLevelID As String = TIMS.GetListValue(HandLevelID)

            If all_Identity2.IndexOf("06") > -1 Then
                Select Case v_rblHandType'rblHandType.SelectedValue
                    Case "2" '新制
                        '障礙類別2 障礙等級2
                        Dim sMyValue As String = TIMS.GetCblValue(HandTypeID2)
                        If sMyValue = "" Then Errmsg &= "參訓身分別 有身心障礙者 障礙類別2 為必填！" & vbCrLf
                        If HandLevelID2.SelectedIndex = 0 OrElse v_HandLevelID2 = "" Then Errmsg &= "參訓身分別 有身心障礙者 障礙等級2 為必填！" & vbCrLf
                    Case "1" '舊制
                        '障礙類別 障礙等級
                        If HandTypeID.SelectedIndex = 0 OrElse v_HandTypeID = "" Then Errmsg &= "參訓身分別 有身心障礙者 障礙類別 為必填！" & vbCrLf
                        If HandLevelID.SelectedIndex = 0 OrElse v_HandLevelID = "" Then Errmsg &= "參訓身分別 有身心障礙者 障礙等級 為必填！" & vbCrLf
                    Case Else '新舊制未選
                        '身心障礙者 新舊制未選擇
                        Errmsg &= "參訓身分別 有身心障礙者 障礙種類 新／舊制 為必填！" & vbCrLf
                End Select
            End If
        End If

        'If TestStr="AmuTest" Then gFlagEnv=False '測試用。
        rqOCID = TIMS.ClearSQM(rqOCID)
        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        FTDateHidden.Value = TIMS.Cdate3(FTDateHidden.Value)
        If gFlagEnv Then '正式環境。('測試用。)
            'Dim vTitle As String=""
            'vTitle="授權設定該班級有開放"
            If FTDateHidden.Value = "" OrElse Hid_OCID.Value = "" OrElse rqOCID = "" Then
                Errmsg &= cst_errMsg12 & vbCrLf 'Exit Sub
            End If
            If FTDateHidden.Value <> "" AndAlso TIMS.ChkIsEndDate(rqOCID, TIMS.cst_FunID_學員資料維護, dtArc) Then
                ''學員資料維護於訓後30日不能修改'Cst_Msg30
                'If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 30, CDate(FTDateHidden.Value)), Today) >= 0 Then ErrMessage &= Cst_Msg30 & vbCrLf
                If TIMS.Cst_TPlanID14DayCanEditStud.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID <= 1 Then
                    If Not flgROLEIDx0xLIDx0 Then '判斷登入者的權限。
                        '非 ROLEID=0 LID=0 '學員資料維護於訓後30日不能修改'Cst_Msg30
                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, 3, CDate(FTDateHidden.Value)), Today) >= 0 Then Errmsg &= cst_Msg30x28 & vbCrLf
                    End If
                Else
                    '針對委外職前訓練計畫，系統權限 限制
                    Dim flgElseEvent1 As Boolean = True '其它狀況 預設為True
                    'If cst_TPlanIDCanEditStud_id37.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '    Select Case sm.UserInfo.LID
                    '        Case "1" '分署(中心)
                    '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                    '                flgElseEvent1 = False '已經設定狀況 其它@False
                    '                Errmsg &= cst_MsgTPlanID37a & vbCrLf
                    '            End If
                    '        Case "2" '委外
                    '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                    '                flgElseEvent1 = False '已經設定狀況 其它@False
                    '                Errmsg &= cst_MsgTPlanID37b & vbCrLf
                    '            End If
                    '    End Select
                    'End If

                    '其它狀況 預設為True
                    If flgElseEvent1 Then
                        If FTDateHidden.Value = "" Then
                            '學員資料維護於訓後30日不能修改'Cst_Msg30
                            If Not flgROLEIDx0xLIDx0 Then Errmsg &= cst_Msg30 & vbCrLf '判斷登入者的權限。
                        Else
                            '學員資料維護於訓後30日不能修改'Cst_Msg30
                            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 30, CDate(FTDateHidden.Value)), Today) >= 0 Then
                                If Not flgROLEIDx0xLIDx0 Then Errmsg &= cst_Msg30 & vbCrLf '判斷登入者的權限。
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If SServiceDate.Text <> "" AndAlso Not TIMS.IsDate1(SServiceDate.Text) Then
            SServiceDate.Text = ""
            Errmsg &= "服役日期 起始日期 格式有誤!" & vbCrLf
        End If
        If FServiceDate.Text <> "" AndAlso Not TIMS.IsDate1(FServiceDate.Text) Then
            FServiceDate.Text = ""
            Errmsg &= "服役日期 結束日期 格式有誤!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False
        SServiceDate.Text = TIMS.Cdate3(SServiceDate.Text)
        FServiceDate.Text = TIMS.Cdate3(FServiceDate.Text)

        Dim dr As DataRow = Nothing
        If Not StdTr.Visible Then
            '新增狀態，檢查有沒有個人資料存在
            Dim pms2 As New Hashtable From {{"OCID", TIMS.CINT1(rqOCID)}, {"IDNO", IDNO.Text}}
            sql = "" & vbCrLf
            sql &= " SELECT 'x' x" & vbCrLf
            sql &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
            sql &= " JOIN STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
            sql &= " WHERE a.OCID=@OCID AND b.IDNO=@IDNO" & vbCrLf
            dr = DbAccess.GetOneRow(sql, objconn, pms2)
            If Not dr Is Nothing Then
                Errmsg &= "此班級已經有相同的身分證號碼!" & vbCrLf
                'Common.MessageBox(Me, "此班級已經有相同的身分證號碼!")
                'Page.RegisterStartupScript("hard", "<script>hard();</script>")
                'Exit Function '離開 
            End If
        End If

        '產投檢查。
        Dim flagTPlanID28a As Boolean = False '(產投 28.54) 在職用
        Dim flagTIMSNot28a As Boolean = True '(TIMS) 職前用(非在職)
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If Hid_show_actno_budid.Value = "Y" Then
            flagTPlanID28a = True
            flagTIMSNot28a = False
        End If

        School.Text = TIMS.ClearSQM(School.Text)
        Department.Text = TIMS.ClearSQM(Department.Text)
        If School.Text = "" Then Errmsg &= "請輸入 個人基本資料-學校名稱" & vbCrLf
        If Department.Text = "" Then Errmsg &= "請輸入 個人基本資料-科系" & vbCrLf

        '產投檢查。
        If flagTPlanID28a Then
            Dim TestVal As String = ""
            Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
            TestVal = TIMS.Get_GraduateStatusValue(v_GraduateStatus)
            If TestVal = "" Then Errmsg &= "請選擇 個人基本資料-畢業狀況" & vbCrLf
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) = -1 Then
                Dim v_graduatey As String = TIMS.GetListValue(graduatey)
                If v_graduatey = "" Then Errmsg &= "請選擇 個人基本資料-畢業狀況-畢業年份" & vbCrLf
            End If
            If Uname.Text <> "" Then Uname.Text = TIMS.ClearSQM(Uname.Text)
            If ServDept.Text <> "" Then ServDept.Text = TIMS.ClearSQM(ServDept.Text)
            If Uname.Text = "" Then Errmsg &= "請輸入服務單位資料-目前任職公司" & vbCrLf '服務部門(目前任職公司) Uname
            If ServDept.Visible AndAlso ServDept.Text = "" Then Errmsg &= "請輸入服務單位資料-目前任職部門" & vbCrLf '服務部門(目前任職部門) ServDept

            Dim v_ddlSERVDEPTID As String = TIMS.GetListValue(ddlSERVDEPTID)
            Dim v_ddlJOBTITLEID As String = TIMS.GetListValue(ddlJOBTITLEID)
            Dim v_Q1 As String = TIMS.GetListValue(Q1)
            'Dim v_MIdentityID As String=TIMS.GetListValue(MIdentityID)
            If ddlSERVDEPTID.Visible AndAlso v_ddlSERVDEPTID = "" Then Errmsg &= "請選擇 服務單位資料-目前任職部門" & vbCrLf
            If v_MIdentityID <> "02" Then
                If JobTitle.Visible AndAlso JobTitle.Text = "" Then Errmsg &= "請輸入服務單位資料-職務" & vbCrLf
                If ddlJOBTITLEID.Visible AndAlso v_ddlJOBTITLEID = "" Then Errmsg &= "請選擇服務單位資料-職務" & vbCrLf
            End If

            Select Case v_Q1'Q1.SelectedValue
                Case "1", "0"
                Case Else
                    Errmsg &= "請選擇參訓背景資料-是否由公司推薦參訓" & vbCrLf
            End Select

            TestVal = ""
            For Each item As ListItem In Q2.Items
                If item.Selected Then
                    TestVal = "1"
                    Exit For
                End If
            Next
            If TestVal = "" Then Errmsg &= "請選擇參訓背景資料-參訓動機(至少選擇1項)" & vbCrLf

            Dim v_Q5 As String = TIMS.GetListValue(Q5)
            If Q5.SelectedIndex = -1 OrElse v_Q5 = "" Then Errmsg &= "請選擇參訓背景資料-服務單位是否屬於中小企業" & vbCrLf

            'If Trim(Q61.Text) <> "" Then Q61.Text=TIMS.ClearSQM(Q61.Text) Else Q61.Text=""
            'If Trim(Q62.Text) <> "" Then Q62.Text=TIMS.ClearSQM(Q62.Text) Else Q62.Text=""
            'If Trim(Q63.Text) <> "" Then Q63.Text=TIMS.ClearSQM(Q63.Text) Else Q63.Text=""
            'If Trim(Q64.Text) <> "" Then Q64.Text=TIMS.ClearSQM(Q64.Text) Else Q64.Text=""
            Q61.Text = TIMS.ClearSQM(Q61.Text)
            Q62.Text = TIMS.ClearSQM(Q62.Text)
            Q63.Text = TIMS.ClearSQM(Q63.Text)
            Q64.Text = TIMS.ClearSQM(Q64.Text)

            If Q61.Text <> "" Then
                If Not IsNumeric(Q61.Text) Then
                    Errmsg &= "參訓背景資料-個人工作年資必須為數字" & vbCrLf
                Else
                    Q61.Text = Val(Q61.Text)
                End If
            End If

            If Q62.Text <> "" Then
                If Not IsNumeric(Q62.Text) Then
                    Errmsg &= "參訓背景資料-在這家公司的年資必須為數字" & vbCrLf
                Else
                    Q62.Text = Val(Q62.Text)
                End If
            End If

            If Q63.Text <> "" Then
                If Not IsNumeric(Q63.Text) Then
                    Errmsg &= "參訓背景資料-在這職位的年資必須為數字" & vbCrLf
                Else
                    Q63.Text = Val(Q63.Text)
                End If
            End If

            If Q64.Text <> "" Then
                If Not IsNumeric(Q64.Text) Then
                    Errmsg &= "參訓背景資料-最近升遷離本職幾年必須為數字" & vbCrLf
                Else
                    Q64.Text = Val(Q64.Text)
                End If
            End If

            If Errmsg = "" Then
                Const cst_x05 As Double = 0.5
                If Q61.Text <> "" Then
                    If Not Val(Q61.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-個人工作年資 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If

                If Q62.Text <> "" Then
                    If Not Val(Q62.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-在這家公司的年資 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If

                If Q63.Text <> "" Then
                    If Not Val(Q63.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-在這職位的年資 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If

                If Q64.Text <> "" Then
                    If Not Val(Q64.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-最近升遷離本職幾年 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If
            End If
        End If

        If flagTIMSNot28a Then
            '非產投檢查(有可能是其他在職班)
            'If Trim(Q61.Text) <> "" Then Q61.Text=Trim(Q61.Text) Else Q61.Text=""
            'If Trim(Q62.Text) <> "" Then Q62.Text=Trim(Q62.Text) Else Q62.Text=""
            'If Trim(Q63.Text) <> "" Then Q63.Text=Trim(Q63.Text) Else Q63.Text=""
            'If Trim(Q64.Text) <> "" Then Q64.Text=Trim(Q64.Text) Else Q64.Text=""
            Q61.Text = TIMS.ClearSQM(Q61.Text)
            Q62.Text = TIMS.ClearSQM(Q62.Text)
            Q63.Text = TIMS.ClearSQM(Q63.Text)
            Q64.Text = TIMS.ClearSQM(Q64.Text)

            If Q61.Text <> "" Then
                If Not IsNumeric(Q61.Text) Then
                    Errmsg &= "參訓背景資料-個人工作年資必須為數字" & vbCrLf
                Else
                    Q61.Text = Val(Q61.Text)
                End If
            End If

            If Q62.Text <> "" Then
                If Not IsNumeric(Q62.Text) Then
                    Errmsg &= "參訓背景資料-在這家公司的年資必須為數字" & vbCrLf
                Else
                    Q62.Text = Val(Q62.Text)
                End If
            End If

            If Q63.Text <> "" Then
                If Not IsNumeric(Q63.Text) Then
                    Errmsg &= "參訓背景資料-在這職位的年資必須為數字" & vbCrLf
                Else
                    Q63.Text = Val(Q63.Text)
                End If
            End If

            If Q64.Text <> "" Then
                If Not IsNumeric(Q64.Text) Then
                    Errmsg &= "參訓背景資料-最近升遷離本職幾年必須為數字" & vbCrLf
                Else
                    Q64.Text = Val(Q64.Text)
                End If
            End If

            If Errmsg = "" Then
                Const cst_x05 As Double = 0.5
                If Q61.Text <> "" Then
                    If Not Val(Q61.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-個人工作年資 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If

                If Q62.Text <> "" Then
                    If Not Val(Q62.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-在這家公司的年資 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If

                If Q63.Text <> "" Then
                    If Not Val(Q63.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-在這職位的年資 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If

                If Q64.Text <> "" Then
                    If Not Val(Q64.Text) Mod cst_x05 = 0 Then
                        '年資開放小數點填寫，但必須以0.5為最小單位。
                        Errmsg &= "參訓背景資料-最近升遷離本職幾年 開放小數點填寫，但必須以0.5為最小單位。" & vbCrLf
                    End If
                End If
            End If
        End If

        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        If v_DegreeID = "" Then Errmsg &= "請選擇最高學歷。" & vbCrLf

        If tr_DDL_DISASTER.Visible AndAlso v_MIdentityID = TIMS.cst_Identity_40 AndAlso v_DDL_DISASTER = "" Then 'ADID 重大災害選項
            Errmsg &= "主要參訓身分別選擇「經公告之重大災害受災者」，須選擇「重大災害選項」不可為空！" & vbCrLf
        End If

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>(儲存)Button1:儲存回查詢頁面 ／Button2:維護下一位學員</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click
        '重載數據(JS無法讀取有效值)
        'Call ReLoad_SB4IDx()

        Dim Errmsg As String = ""
        Dim js_Errmsg As String = ""
        Dim chkFlag As Boolean = False '有錯誤為False
        '儲存前先檢查輸入資料的正確性
        chkFlag = SUtl_CheckData1(Errmsg)

        If Errmsg = "" Then
            If Hid_show_actno_budid.Value = "Y" Then
                Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)
                Dim v_MIdentityID As String = TIMS.GetListValue(MIdentityID)
                Dim v_BudID As String = TIMS.GetListValue(BudID)
                Dim v_SupplyID As String = TIMS.GetListValue(SupplyID)

                'in_parms.Clear()
                Dim in_parms As New Hashtable From {
                    {"PassPortNO", v_PassPortNO},
                    {"MIdentityID", v_MIdentityID},
                    {"BudID", v_BudID},
                    {"SupplyID", v_SupplyID},
                    {"TPlanID", sm.UserInfo.TPlanID}
                }
                chkFlag = SUtl_CheckData2(in_parms, Errmsg)
            End If
        End If

        '= 錯誤訊息，顯示並離開 =
        If Errmsg <> "" Then
            js_Errmsg = Common.GetJsString(Errmsg)
            If Hid_show_actno_budid.Value = "Y" Then
                Page.RegisterStartupScript("11111", "<script>blockAlert('" & js_Errmsg & "','錯誤',function(){ChangeMode(1);});</script>")
                'Common.MessageBox(Me, Errmsg)
                'Exit Sub
            Else
                '46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
                If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Page.RegisterStartupScript("11111", "<script>blockAlert('" & js_Errmsg & "','錯誤',function(){Change_MBTable(2);});</script>")
                Else
                    Page.RegisterStartupScript("11111", "<script>blockAlert('" & js_Errmsg & "','錯誤',function(){ChangeMode(1);});</script>")
                End If
            End If
            Exit Sub
            'Return False
        End If
        '= 錯誤訊息，顯示並離開 =

        If Not chkFlag AndAlso Errmsg = "" Then
            Errmsg = "有異常的錯誤，請洽系統管理員!"
            Common.MessageBox(Me, Errmsg)

            Dim strErrmsg As String = "Not chkFlag AndAlso Errmsg Empty!!"
            strErrmsg &= "Errmsg:" & vbCrLf & Errmsg & vbCrLf
            strErrmsg &= "IDNO:" & vbCrLf & IDNO.Text & vbCrLf
            strErrmsg &= "rqOCID:" & vbCrLf & rqOCID & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
            'Return False
        End If

        '更新MakeSOCID 'Dim vRejectSOCID As String=TIMS.ClearSQM(RejectSOCID.SelectedValue)
        'Dim vSOCID_Sel As String=TIMS.ClearSQM(v_SOCID) 'SOCID.SelectedValue)
        Dim v_SOCID As String = TIMS.GetListValue(SOCID)
        Call TIMS.OpenDbConn(objconn)
        Dim fg_NG_updateCS As Boolean = False
        'Dim sql As String=""
        If StdTr.Visible Then
            '修改
            fg_NG_updateCS = (v_SOCID = "") '應該只有值，其他答案為錯誤
            If fg_NG_updateCS Then
                Common.MessageBox(Me, "修改動作，查無可修改學員資料，請重新查詢操作!")
                Exit Sub
            End If

            '修改
            Dim hPMS As New Hashtable From {{"SOCID", TIMS.CINT1(v_SOCID)}}
            Dim sql As String = "SELECT 'x' FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID" '" & v_SOCID & "'"
            Dim dtRR As DataTable = DbAccess.GetDataTable(sql, objconn, hPMS)
            '應該只有一筆，其他數字者為錯誤
            fg_NG_updateCS = (dtRR Is Nothing OrElse dtRR.Rows.Count <> 1)
            If fg_NG_updateCS Then
                Common.MessageBox(Me, "修改動作，查無可修改學員資料，請重新查詢操作!")
                Exit Sub
            End If
        End If

        Dim ref_SID As String = "" '回傳SID
        Try
            Call SUtl_SaveData1(ref_SID) 'STUD_STUDENTINFO 傳出 SID
            Hid_SID_C1.Value = ref_SID
            If ref_SID = "" Then
                Common.MessageBox(Me, "儲存學員資料有誤，請重新操作!")
                Return 'Exit Sub
            End If
            Call SUtl_SaveData2() 'STUD_SUBDATA
            Call SUtl_SaveData3() 'STUD_ENTERTEMP'STUD_ENTERTEMP2
            Call SUtl_SaveData4() 'CLASS_STUDENTSOFCLASS '傳入 SID/ Hid_SID_C1.Value/ref_SID
        Catch ex As Exception
            Dim sErrmsg As String = ""
            sErrmsg = String.Concat("儲存學員資料有誤，請重新操作!!", ex.Message)
            Common.MessageBox(Me, sErrmsg)

            Dim strErrmsg As String = ""
            strErrmsg = String.Concat("##WDAIIP.SD_03_002_add.Button1_Click.儲存學員資料 : ", vbCrLf, " ,SID:", ref_SID, " ,IDNO:", IDNO.Text, " ,rqOCID:", rqOCID, vbCrLf)
            strErrmsg &= String.Concat(",ex.Message: ", ex.Message, vbCrLf)
            If gstr_ROWVAL_1 <> "" Then strErrmsg &= String.Concat(",gstr_ROWVAL_1:", vbCrLf, gstr_ROWVAL_1, vbCrLf)
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入 'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
            Exit Sub
        End Try

        'Common.RespWrite(Me, "<script>alert('儲存成功');</script>")
        'Common.RespWrite(Me, "<script>blockAlert('儲存成功!','提示訊息!',function(){ChangeMode(1);});</script>")
        'Common.MessageBox(Me, "儲存成功!")
        If Session(vs_SearchStr) IsNot Nothing Then
            ViewState(vs_SearchStr) = Session(vs_SearchStr)
            Session(vs_SearchStr) = ViewState(vs_SearchStr)
            'Session(vs_SearchStr)=Nothing
        End If

        If sender Is Button1 Then '(儲存1)
            '儲存回查詢頁面
            'Session(vs_SearchStr)=ViewState(vs_SearchStr)
            Common.RespWrite(Me, "<script>location.href='SD_03_002.aspx?ID=" & Request("ID") & "'</script>")

        ElseIf sender Is Button2 Then '(儲存2)
            '維護下一位學員
            Dim Index As Integer = SOCID.SelectedIndex
            If Index < SOCID.Items.Count - 1 Then
                SOCID.SelectedItem.Selected = False
                SOCID.Items(Index + 1).Selected = True
                v_SOCID = TIMS.GetListValue(SOCID) 'SOCID.SelectedValue
                Call Clear_data()   '清理資料
                Call GetOpenDate2()   '塞入班級資料 (排第2順位) '重要
                Call Create1_Stud(v_SOCID) '塞入學員資料
                Call GetScript(v_SOCID) '學員資料審核功能的欄位鎖住
                'Call GetOpenDate2()   '塞入班級資料 (排第2順位) '重要
                'Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")
                Page.RegisterStartupScript("11111", "<script>blockAlert('儲存成功!',null,function(){ChangeMode(1);});</script>")
            Else
                Page.RegisterStartupScript("11111", "<script>blockAlert('已經到最後一位學員!','錯誤',function(){ChangeMode(1);});</script>")
            End If
        End If
    End Sub

    ''' <summary>UDPATE STUD_STUDENTINFO</summary>
    ''' <param name="ref_SID"></param>
    ''' <returns>SID</returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData1(ByRef ref_SID As String) As Boolean
        Dim rst As Boolean = False '儲存成功為True
        'Dim SID As String="" '回傳SID
        'ByRef SID As String

        '無身分證號，無法作業
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text = "" OrElse IDNO.Text.Length < 10 Then Return False '長度-小於10-異常

        ForeIDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(ForeIDNO.Text))

        Dim s_rqFUNID As String = TIMS.ClearSQM(Request("ID"))
        '先檢查是否有資料
        If StdTr.Visible = False Then
            '新增狀態，檢查有沒有個人資料存在
            Dim sMemo As String = ""
            sMemo &= "&NAME=" & Name.Text
            '寫入Log查詢 SubInsAccountLog1(Auth_Accountlog)
            Call TIMS.SubInsAccountLog1(Me, s_rqFUNID, TIMS.cst_wm新增, Session(TIMS.gcst_rblWorkMode), rqOCID, sMemo, objconn)
            'Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm新增, Session(TIMS.gcst_rblWorkMode), rqOCID, "")
        Else
            '修改狀態，檢查有沒有個人資料存在
            Dim sMemo As String = ""
            sMemo &= "&NAME=" & Name.Text
            '寫入Log查詢 SubInsAccountLog1(Auth_Accountlog)
            Call TIMS.SubInsAccountLog1(Me, s_rqFUNID, TIMS.cst_wm修改, Session(TIMS.gcst_rblWorkMode), rqOCID, sMemo, objconn)
            'Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm修改, Session(TIMS.gcst_rblWorkMode), rqOCID, "")
        End If
        Call UPDATE_STUDENTINFO(objconn, IDNO.Text, ref_SID)

        'Try '測試。
        'Catch ex As Exception
        '    If Not objTrans Is Nothing Then DbAccess.RollbackTrans(objTrans)
        '    'Me.Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
        '    'Common.MessageBox(Me, "!!發生錯誤,儲存失敗!!")
        '    'Common.MessageBox(Me, ex.ToString)
        '    Dim jsStr As String=""
        '    jsStr=Common.GetJsString("!!發生錯誤,儲存失敗!!")
        '    Common.RespWrite(Me, "<script>alert('" & jsStr & "');</script>")
        '    jsStr=Common.GetJsString(ex.ToString)
        '    Common.RespWrite(Me, "<script>alert('" & jsStr & "');</script>")
        '    'Exit Sub
        '    Throw
        'End Try
        Return rst
    End Function
    'UPDATE STUD_SUBDATA
    ''' <summary>UPDATE STUD_SUBDATA</summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData2() As Boolean
        Dim rst As Boolean = False '儲存成功為True
        Dim all_Identity2 As String = ""
        all_Identity2 = Get_All_Identity2()

        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)

        'STUD_STUDENTINFO
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Dim hPMS9 As New Hashtable From {{"IDNO", IDNO.Text}}
        Dim sql9 As String = "SELECT SID FROM STUD_STUDENTINFO WHERE IDNO=@IDNO" '找出SID
        Dim dt9 As DataTable = DbAccess.GetDataTable(sql9, objconn, hPMS9)
        If TIMS.dtNODATA(dt9) Then Return rst

        For z As Integer = 0 To dt9.Rows.Count - 1
            Dim Myda2 As SqlDataAdapter = Nothing
            Dim Mydt2 As DataTable = Nothing 'STUD_SUBDATA
            Dim Mydr2 As DataRow = Nothing 'STUD_SUBDATA
            Dim SID As String = Convert.ToString(dt9.Rows(z)("SID"))
            Dim sql2 As String = " SELECT * FROM STUD_SUBDATA WHERE SID='" & SID & "'"
            Mydt2 = DbAccess.GetDataTable(sql2, Myda2, objconn)
            If Mydt2.Rows.Count = 0 Then
                Mydr2 = Mydt2.NewRow
                Mydt2.Rows.Add(Mydr2)
                Mydr2("SID") = SID
            Else
                Mydr2 = Mydt2.Rows(0)
            End If

            '更新學員資料副檔---   Start
            Mydr2("Name") = Name.Text  '姓名
            Mydr2("School") = School.Text '學校
            Mydr2("Department") = Department.Text

            ZipCode1.Value = TIMS.ClearSQM(ZipCode1.Value)
            ZipCode1_B3.Value = TIMS.ClearSQM(ZipCode1_B3.Value)
            hidZipCode1_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
            ZipCode1_N.Value = TIMS.ClearSQM(ZipCode1_N.Value)
            Address.Text = TIMS.ClearSQM(Address.Text)
            Mydr2("ZipCode1") = TIMS.GetValue1(ZipCode1.Value)
            Mydr2("ZipCode1_6W") = TIMS.GetValue1(hidZipCode1_6W.Value)
            Mydr2("ZipCode1_N") = TIMS.GetValue1(ZipCode1_N.Value)
            'Mydr2("Address")=Address.Text
            'for 個資保護,有**表示沒有修改
            If InStr(Address.Text, "*", CompareMethod.Text) = 0 Then Mydr2("Address") = TIMS.GetSpace(Address.Text)

            '同通訊地址
            If CheckBox1.Checked Then
                Mydr2("ZipCode2") = TIMS.GetValue1(ZipCode1.Value)
                Mydr2("ZipCode2_6W") = TIMS.GetValue1(hidZipCode1_6W.Value)
                Mydr2("ZipCode2_N") = TIMS.GetValue1(ZipCode1_N.Value)
                'Mydr2("HouseholdAddress")=Address.Text
                'for 個資保護,有**表示沒有修改
                If InStr(Address.Text, "*", CompareMethod.Text) = 0 Then Mydr2("HouseholdAddress") = TIMS.GetSpace(Address.Text)
            Else
                '相異…自行儲存
                hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value)
                Mydr2("ZipCode2") = TIMS.GetValue1(ZipCode2.Value)
                Mydr2("ZipCode2_6W") = TIMS.GetValue1(hidZipCode2_6W.Value)
                Mydr2("ZipCode2_N") = TIMS.GetValue1(ZipCode2_N.Value)
                'Mydr2("HouseholdAddress")=HouseholdAddress.Text
                'for 個資保護,有**表示沒有修改
                If InStr(HouseholdAddress.Text, "*", CompareMethod.Text) = 0 Then Mydr2("HouseholdAddress") = TIMS.GetSpace(HouseholdAddress.Text)
            End If

            Email.Text = TIMS.ClearSQM(Email.Text)
            PhoneD.Text = TIMS.ClearSQM(PhoneD.Text)
            PhoneN.Text = TIMS.ClearSQM(PhoneN.Text)
            CellPhone.Text = TIMS.ClearSQM(CellPhone.Text)
            Mydr2("Email") = Email.Text
            Mydr2("PhoneD") = PhoneD.Text
            Mydr2("PhoneN") = PhoneN.Text
            Dim v_rblMobil As String = TIMS.GetListValue(rblMobil)
            Mydr2("CellPhone") = If(v_rblMobil = "Y", CellPhone.Text, "")

            EmergencyContact.Text = TIMS.ClearSQM(EmergencyContact.Text)
            EmergencyRelation.Text = TIMS.ClearSQM(EmergencyRelation.Text)
            EmergencyPhone.Text = TIMS.ClearSQM(EmergencyPhone.Text)
            Mydr2("EmergencyContact") = EmergencyContact.Text
            Mydr2("EmergencyRelation") = EmergencyRelation.Text
            Mydr2("EmergencyPhone") = EmergencyPhone.Text
            'By Milor 20080904----start
            Dim iTypeSame As Integer = 0 '0:相異 2:同通訊地址 3:同戶籍地址
            '2.緊急聯絡人地址同通訊地址/ '3.緊急聯絡人地址同戶籍地址
            iTypeSame = If(CheckBox2.Checked, 2, If(CheckBox3.Checked, 3, 0))

            Select Case iTypeSame
                Case 2 '0:相異 2:同通訊地址 3:同戶籍地址
                    ZipCode1.Value = TIMS.ClearSQM(ZipCode1.Value)
                    ZipCode1_B3.Value = TIMS.ClearSQM(ZipCode1_B3.Value)
                    hidZipCode1_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
                    ZipCode1_N.Value = TIMS.ClearSQM(ZipCode1_N.Value)
                    Address.Text = TIMS.ClearSQM(Address.Text)
                    Mydr2("ZipCode3") = TIMS.GetValue1(ZipCode1.Value)
                    Mydr2("ZipCode3_6W") = TIMS.GetValue1(hidZipCode1_6W.Value)
                    Mydr2("ZipCode3_N") = TIMS.GetValue1(ZipCode1_N.Value)
                    'Mydr2("EmergencyAddress")=Address.Text
                    'Mydr2("EmergencyAddress")=Address.Text '20090520 fix(Jimmy)
                    'for 個資保護,有**表示沒有修改
                    If InStr(Address.Text, "*", CompareMethod.Text) = 0 Then Mydr2("EmergencyAddress") = TIMS.GetSpace(Address.Text) '20090520 fix(Jimmy)
                Case 3 '0:相異 2:同通訊地址 3:同戶籍地址
                    ZipCode2.Value = TIMS.ClearSQM(ZipCode2.Value)
                    ZipCode2_B3.Value = TIMS.ClearSQM(ZipCode2_B3.Value)
                    hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value)
                    ZipCode2_N.Value = TIMS.ClearSQM(ZipCode2_N.Value)
                    HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)

                    Mydr2("ZipCode3") = TIMS.GetValue1(ZipCode2.Value)
                    Mydr2("ZipCode3_6W") = TIMS.GetValue1(hidZipCode2_6W.Value)
                    Mydr2("ZipCode3_N") = TIMS.GetValue1(ZipCode2_N.Value)
                    'Mydr2("EmergencyAddress")=HouseholdAddress.Text
                    'Mydr2("EmergencyAddress")=HouseholdAddress.Text '20090520 fix(Jimmy)
                    'for 個資保護,有**表示沒有修改
                    If InStr(HouseholdAddress.Text, "*", CompareMethod.Text) = 0 Then Mydr2("EmergencyAddress") = TIMS.GetSpace(HouseholdAddress.Text) '20090520 fix(Jimmy)
                Case Else '0:'0:相異 2:同通訊地址 3:同戶籍地址
                    '相異…自行儲存
                    hidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value)
                    EmergencyAddress.Text = TIMS.ClearSQM(EmergencyAddress.Text)
                    Mydr2("ZipCode3") = TIMS.GetValue1(ZipCode3.Value)
                    Mydr2("ZipCode3_6W") = TIMS.GetValue1(hidZipCode3_6W.Value)
                    Mydr2("ZipCode3_N") = TIMS.GetValue1(ZipCode3_N.Value)
                    'Mydr2("EmergencyAddress")=EmergencyAddress.Text
                    'Mydr2("EmergencyAddress")=EmergencyAddress.Text '20090520 fix (Jimmy)
                    'for 個資保護,有**表示沒有修改
                    If InStr(EmergencyAddress.Text, "*", CompareMethod.Text) = 0 Then Mydr2("EmergencyAddress") = TIMS.GetSpace(EmergencyAddress.Text) '20090520 fix(Jimmy)
            End Select

            'By Milor 20080904----end
            Title1.Text = TIMS.ClearSQM(Title1.Text)
            Title2.Text = TIMS.ClearSQM(Title2.Text)
            PriorWorkOrg1.Text = TIMS.ClearSQM(PriorWorkOrg1.Text)
            PriorWorkOrg2.Text = TIMS.ClearSQM(PriorWorkOrg2.Text)
            If trPriorWorkOrg1.Visible AndAlso trPriorWorkOrg2.Visible Then
                'If PriorWorkOrg1.Text <> "" Then PriorWorkOrg1.Text=Trim(PriorWorkOrg1.Text)
                If PriorWorkOrg1.Text <> "" Then PriorWorkOrg1.Text = Mid(PriorWorkOrg1.Text, 1, 30)
                If PriorWorkOrg2.Text <> "" Then PriorWorkOrg2.Text = Mid(PriorWorkOrg2.Text, 1, 30)
                Mydr2("PriorWorkOrg1") = Convert.DBNull
                If PriorWorkOrg1.Text <> "" Then Mydr2("PriorWorkOrg1") = PriorWorkOrg1.Text
                Mydr2("PriorWorkOrg2") = Convert.DBNull
                If PriorWorkOrg2.Text <> "" Then Mydr2("PriorWorkOrg2") = PriorWorkOrg2.Text
                Mydr2("Title1") = Convert.DBNull
                If Title1.Text <> "" Then Mydr2("Title1") = Title1.Text
                Mydr2("Title2") = Convert.DBNull
                If Title2.Text <> "" Then Mydr2("Title2") = Title2.Text
            End If
            If trTable6.Visible Then
                Mydr2("SOfficeYM1") = Convert.DBNull
                If SOfficeYM1.Text <> "" Then Mydr2("SOfficeYM1") = SOfficeYM1.Text
                Mydr2("FOfficeYM1") = Convert.DBNull
                If FOfficeYM1.Text <> "" Then Mydr2("FOfficeYM1") = FOfficeYM1.Text
                Mydr2("SOfficeYM2") = Convert.DBNull
                If SOfficeYM2.Text <> "" Then Mydr2("SOfficeYM2") = SOfficeYM2.Text
                Mydr2("FOfficeYM2") = Convert.DBNull
                If FOfficeYM2.Text <> "" Then Mydr2("FOfficeYM2") = FOfficeYM2.Text
            End If

            'trPriorWorkPay.Visible
            'ActNo2_TR.Visible=False '投保單位保險證號 投保薪資級距
            'trTitle1.Visible=False '職稱 受訓前薪資
            PriorWorkPay.Text = TIMS.ClearSQM(PriorWorkPay.Text)
            If ActNo2_TR.Visible Then
                'Mydr2("PriorWorkPay")=Convert.DBNull
                'If PriorWorkPay.Text <> "" Then Mydr2("PriorWorkPay")=PriorWorkPay.Text
                Mydr2("PriorWorkPay") = If(PriorWorkPay.Text <> "", PriorWorkPay.Text, Convert.DBNull)
            End If
            If trTraffic.Visible Then
                'If Traffic.SelectedValue <> "0" AndAlso Traffic.SelectedValue <> "" Then Mydr2("Traffic")=Traffic.SelectedValue
                Dim v_Traffic As String = TIMS.GetListValue(Traffic)
                Mydr2("Traffic") = If(v_Traffic <> "0" AndAlso v_Traffic <> "", v_Traffic, Convert.DBNull)
            End If

            'DropDownList
            Dim v_ShowDetail As String = TIMS.GetListValue(ShowDetail)
            Mydr2("ShowDetail") = If(v_ShowDetail = "Y", "Y", "N")
            Mydr2("ServiceID") = ServiceID.Text
            Mydr2("MilitaryAppointment") = MilitaryAppointment.Text
            Mydr2("MilitaryRank") = MilitaryRank.Text
            Mydr2("SServiceDate") = If(SServiceDate.Text <> "", TIMS.Cdate2(SServiceDate.Text), Convert.DBNull)
            Mydr2("FServiceDate") = If(FServiceDate.Text <> "", TIMS.Cdate2(FServiceDate.Text), Convert.DBNull)
            Mydr2("ServiceOrg") = ServiceOrg.Text
            Mydr2("ChiefRankName") = ChiefRankName.Text

            ZipCode4.Value = TIMS.ClearSQM(ZipCode4.Value)
            ZipCode4_B3.Value = TIMS.ClearSQM(ZipCode4_B3.Value)
            hidZipCode4_6W.Value = TIMS.GetZIPCODE6W(ZipCode4.Value, ZipCode4_B3.Value)
            ZipCode4_N.Value = TIMS.ClearSQM(ZipCode4_N.Value)
            ServiceAddress.Text = TIMS.ClearSQM(ServiceAddress.Text)
            Mydr2("ZipCode4") = TIMS.GetValue1(ZipCode4.Value)
            Mydr2("ZipCode4_6W") = TIMS.GetValue1(hidZipCode4_6W.Value)
            Mydr2("ZipCode4_N") = TIMS.GetValue1(ZipCode4_N.Value)
            Mydr2("ServiceAddress") = TIMS.GetSpace(ServiceAddress.Text)
            Mydr2("ServicePhone") = TIMS.ClearSQM(ServicePhone.Text)

            'Dim v_rblHandType As String=TIMS.GetListValue(rblHandType)
            '取得 CheckBoxList 的值 從0開始
            Dim v_HandTypeID2 As String = TIMS.GetCblValue(HandTypeID2)
            Dim v_HandLevelID2 As String = TIMS.GetListValue(HandLevelID2)
            Dim v_HandTypeID As String = TIMS.GetListValue(HandTypeID)
            Dim v_HandLevelID As String = TIMS.GetListValue(HandLevelID)

            '06:身心障礙者 '只有選 身心障礙者 才會做異動。
            If all_Identity2 <> "" AndAlso all_Identity2.IndexOf("06") > -1 Then
                '06:身心障礙者
                Mydr2("HandTypeID") = If(v_HandTypeID <> "", v_HandTypeID, Convert.DBNull)
                Mydr2("HandLevelID") = If(v_HandLevelID <> "", v_HandLevelID, Convert.DBNull)
                Mydr2("HandTypeID2") = If(v_HandTypeID2 <> "", v_HandTypeID2, Convert.DBNull)
                Mydr2("HandLevelID2") = If(v_HandLevelID2 <> "", v_HandLevelID2, Convert.DBNull)
            Else
                '其它狀況為原值
                Mydr2("HandTypeID") = If(v_HandTypeID <> "", v_HandTypeID, Mydr2("HandTypeID"))
                Mydr2("HandLevelID") = If(v_HandLevelID <> "", v_HandLevelID, Mydr2("HandLevelID"))
                Mydr2("HandTypeID2") = If(v_HandTypeID2 <> "", v_HandTypeID2, Mydr2("HandTypeID2"))
                Mydr2("HandLevelID2") = If(v_HandLevelID2 <> "", v_HandLevelID2, Mydr2("HandLevelID2"))
            End If

            '外國籍新增部分2005/12/16------ Start
            Dim v_ForeSex As String = TIMS.GetListValue(ForeSex)
            If v_PassPortNO = "1" Then
                Mydr2("ForeName") = Convert.DBNull
                Mydr2("ForeTitle") = Convert.DBNull
                Mydr2("ForeSex") = Convert.DBNull
                Mydr2("ForeBirth") = Convert.DBNull
                Mydr2("ForeIDNO") = Convert.DBNull
                Mydr2("ForeZip") = Convert.DBNull
                Mydr2("ForeZIP6W") = Convert.DBNull
                Mydr2("ForeZip_N") = Convert.DBNull
                Mydr2("ForeAddr") = Convert.DBNull
            Else
                Mydr2("ForeName") = If(ForeName.Text = "", Convert.DBNull, ForeName.Text)
                Mydr2("ForeTitle") = If(ForeTitle.Text = "", Convert.DBNull, ForeTitle.Text)
                Mydr2("ForeSex") = If(v_ForeSex <> "", v_ForeSex, Convert.DBNull)
                Mydr2("ForeBirth") = If(ForeBirth.Text = "", Convert.DBNull, TIMS.Cdate2(ForeBirth.Text))
                Mydr2("ForeIDNO") = If(ForeIDNO.Text = "", Convert.DBNull, ForeIDNO.Text)
                Mydr2("ForeZip") = TIMS.GetValue1(ForeZip.Value)
                Mydr2("ForeZIP6W") = TIMS.GetValue1(hidForeZIP6W.Value)
                Mydr2("ForeZip_N") = TIMS.GetValue1(ForeZip_N.Value)
                Mydr2("ForeAddr") = TIMS.GetSpace(ForeAddr.Text)
            End If
            '外國籍新增部分2005/12/16------ End
            Mydr2("ModifyAcct") = sm.UserInfo.UserID
            Mydr2("ModifyDate") = Now()

            gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(Mydt2)
            gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, Mydr2)
            DbAccess.UpdateDataTable(Mydt2, Myda2)
            '更新學員資料副檔---   End
        Next

        rst = True
        Return rst
    End Function

    ''' <summary>UPDATE STUD_ENTERTEMP,STUD_ENTERTEMP2</summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData3() As Boolean
        Dim rst As Boolean = False '儲存成功為True
        Dim sql As String = ""
        Dim da4 As SqlDataAdapter = Nothing
        Dim MyTable4 As DataTable = Nothing 'STUD_ENTERTEMP
        Dim mydr4 As DataRow = Nothing

        'IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Dim v_Sex As String = TIMS.GetListValue(Sex)
        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        Dim v_MilitaryID As String = TIMS.GetListValue(MilitaryID)
        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)
        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)

        'PassPortNO 1:本國 /2:外籍(含大陸人士)
        'mydr4("PassPortNO")=v_PassPortNO 'PassPortNO.SelectedValue
        Select Case v_PassPortNO'.SelectedValue
            Case "1", "2"
            Case Else
                v_PassPortNO = "2"
        End Select
        'mydr4("MaritalStatus")=MaritalStatus.SelectedValue
        Select Case v_MaritalStatus'MaritalStatus.SelectedValue
            Case "1", "2"
            Case Else
                v_MaritalStatus = ""
        End Select
        Dim v_rblMobil As String = TIMS.GetListValue(rblMobil)
        ' statr更新Stud_EnterTemp
        sql = " SELECT * FROM STUD_ENTERTEMP WHERE IDNO='" & IDNO.Text & "' "
        MyTable4 = DbAccess.GetDataTable(sql, da4, objconn)
        If MyTable4.Rows.Count <> 0 Then
            For x As Integer = 0 To MyTable4.Rows.Count - 1
                mydr4 = MyTable4.Rows(x)
                mydr4("Name") = Name.Text
                mydr4("Sex") = v_Sex
                mydr4("Birthday") = CDate(TIMS.Cdate2(Birthday.Text))
                mydr4("PassPortNO") = If(v_PassPortNO <> "", v_PassPortNO, Convert.DBNull) '"2"
                mydr4("MaritalStatus") = If(v_MaritalStatus <> "", v_MaritalStatus, Convert.DBNull)
                mydr4("DegreeID") = v_DegreeID
                mydr4("GradID") = TIMS.Get_GraduateStatusValue(GraduateStatus.SelectedValue) 'GraduateStatus.SelectedValue
                mydr4("School") = School.Text
                mydr4("Department") = Department.Text
                mydr4("MilitaryID") = If(v_MilitaryID = "", Convert.DBNull, v_MilitaryID)
                If ZipCode1.Value <> "" AndAlso hidZipCode1_6W.Value <> "" AndAlso Address.Text <> "" Then
                    '若有填寫資料則修改報名資料。
                    mydr4("zipcode") = TIMS.GetValue1(ZipCode1.Value)
                    mydr4("zipCODE6W") = TIMS.GetValue1(hidZipCode1_6W.Value)
                    mydr4("ZIPCODE_N") = TIMS.GetValue1(ZipCode1_N.Value)
                    mydr4("Address") = TIMS.GetSpace(Address.Text)
                End If
                mydr4("Phone1") = TIMS.ClearSQM(PhoneD.Text)
                mydr4("Phone2") = TIMS.ClearSQM(PhoneN.Text)
                mydr4("CellPhone") = If(v_rblMobil = "Y", CellPhone.Text, "")
                mydr4("Email") = TIMS.ChangeEmail(Email.Text)
                mydr4("IsAgree") = "Y"

                gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(MyTable4)
                gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, mydr4)
                DbAccess.UpdateDataTable(MyTable4, da4)
            Next
        End If
        '------ End 

        Dim MyTable5 As DataTable
        Dim da5 As SqlDataAdapter = Nothing
        Dim mydr5 As DataRow
        '------ statr更新STUD_ENTERTEMP2
        sql = " SELECT * FROM STUD_ENTERTEMP2 WHERE IDNO='" & IDNO.Text & "' "
        MyTable5 = DbAccess.GetDataTable(sql, da5, objconn)
        If MyTable5.Rows.Count <> 0 Then
            For y As Integer = 0 To MyTable5.Rows.Count - 1
                mydr5 = MyTable5.Rows(y)
                mydr5("Name") = Name.Text
                mydr5("Sex") = v_Sex
                mydr5("Birthday") = CDate(TIMS.Cdate2(Birthday.Text))
                mydr5("PassPortNO") = If(v_PassPortNO <> "", v_PassPortNO, Convert.DBNull) '"2"
                mydr5("MaritalStatus") = If(v_MaritalStatus <> "", v_MaritalStatus, Convert.DBNull)
                mydr5("DegreeID") = v_DegreeID
                mydr5("GradID") = TIMS.Get_GraduateStatusValue(GraduateStatus.SelectedValue) ' GraduateStatus.SelectedValue
                mydr5("School") = School.Text
                mydr5("Department") = Department.Text
                mydr5("MilitaryID") = If(v_MilitaryID = "", Convert.DBNull, v_MilitaryID)
                If ZipCode1.Value <> "" AndAlso hidZipCode1_6W.Value <> "" AndAlso Address.Text <> "" Then
                    '若有填寫資料則修改報名資料。
                    mydr5("ZIPCODE") = TIMS.GetValue1(ZipCode1.Value)
                    mydr5("ZIPCODE6W") = TIMS.GetValue1(hidZipCode1_6W.Value)
                    mydr5("ZIPCODE_N") = TIMS.GetValue1(ZipCode1_N.Value)
                    mydr5("Address") = TIMS.GetSpace(Address.Text)
                End If
                mydr5("Phone1") = TIMS.ClearSQM(PhoneD.Text)
                mydr5("Phone2") = TIMS.ClearSQM(PhoneN.Text)
                mydr5("CellPhone") = If(v_rblMobil = "Y", CellPhone.Text, "")
                mydr5("Email") = TIMS.ChangeEmail(Email.Text)
                mydr5("IsAgree") = "Y"

                gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(MyTable5)
                gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, mydr5)
                DbAccess.UpdateDataTable(MyTable5, da5)
            Next
        End If
        '------ End 
        rst = True
        Return rst
    End Function

    ''' <summary>津貼類別 不等於 03:就業促進津貼實施辦法 刪除請領津貼的資料。</summary>
    ''' <param name="vSOCID_Sel"></param>
    Sub UPDATE_SUBSIDYRESULT(ByRef vSOCID_Sel As String)
        '津貼類別  不等於 03:就業促進津貼實施辦法 刪除請領津貼的資料。
        If SubsidyID.SelectedValue <> "03" Then
            If vSOCID_Sel <> "" AndAlso StdTr.Visible = True Then
                'Parms.Clear()
                Dim Parms As New Hashtable From {{"SOCID", TIMS.CINT1(vSOCID_Sel)}}
                Dim sql As String = "SELECT 'X' FROM dbo.STUD_SUBSIDYRESULT WHERE SOCID=@SOCID"
                Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn, Parms)
                If dt1.Rows.Count = 0 Then Return '無資料離開
                '有資料刪除
                sql = " DELETE STUD_SUBSIDYRESULT WHERE SOCID=@SOCID"
                DbAccess.ExecuteNonQuery(sql, objconn, Parms)
            End If
        End If

    End Sub

    ''' <summary>
    ''' 更新MakeSOCID
    ''' </summary>
    Function UPDATE_MakeSOCID() As String
        'Dim vRejectSOCID As String=TIMS.ClearSQM(RejectSOCID.SelectedValue)
        Dim vRejectSOCID As String = TIMS.GetListValue(RejectSOCID)

        'UPDATE CLASS_STUDENTSOFCLASS (被遞補學員)
        '清空MakeSOCID
        hide_RejectSOCID.Value = TIMS.ClearSQM(hide_RejectSOCID.Value)

        If hide_RejectSOCID.Value <> "" Then
            If hide_RejectSOCID.Value <> vRejectSOCID Then
                'Parms.Clear()
                Dim Parms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"SOCID", TIMS.CINT1(hide_RejectSOCID.Value)}}
                Dim sql As String = ""
                sql &= " UPDATE CLASS_STUDENTSOFCLASS"
                sql &= " SET MakeSOCID=NULL ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()"
                sql &= " WHERE SOCID=@SOCID "
                DbAccess.ExecuteNonQuery(sql, objconn, Parms)
            End If
        End If

        Return vRejectSOCID
    End Function

    ''' <summary>UPDATE STUD_SERVICEPLACE </summary>
    ''' <param name="iSOCID"></param>
    Sub UPDATE_SERVICEPLACE(ByRef iSOCID As Integer)
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = $" SELECT * FROM STUD_SERVICEPLACE WHERE SOCID={iSOCID}"
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SOCID") = iSOCID 'SOCIDValue
        Else
            dr = dt.Rows(0)
        End If

        PostNo_1.Text = TIMS.ClearSQM(PostNo_1.Text)
        AcctNo1_1.Text = TIMS.ClearSQM(AcctNo1_1.Text)
        BankName.Text = TIMS.ClearSQM(BankName.Text)
        AcctheadNo.Text = TIMS.ClearSQM(AcctheadNo.Text)
        ExBankName.Text = TIMS.ClearSQM(ExBankName.Text)
        AcctExNo.Text = TIMS.ClearSQM(AcctExNo.Text)
        AcctNo2.Text = TIMS.ClearSQM(AcctNo2.Text)
        Dim v_AcctMode As String = TIMS.GetListValue(AcctMode)
        Select Case v_AcctMode'.SelectedValue
            Case "0"
                dr("AcctMode") = 0
                dr("PostNo") = PostNo_1.Text '& "-" & PostNo_2.Text
                dr("AcctNo") = AcctNo1_1.Text '& "-" & AcctNo1_2.Text
                dr("BankName") = Convert.DBNull
                dr("AcctHeadNo") = Convert.DBNull
                dr("ExBankName") = Convert.DBNull
                dr("AcctExNo") = Convert.DBNull
            Case "1"
                dr("AcctMode") = 1
                dr("PostNo") = Convert.DBNull
                dr("BankName") = BankName.Text
                dr("AcctHeadNo") = AcctheadNo.Text
                dr("ExBankName") = ExBankName.Text
                dr("AcctExNo") = AcctExNo.Text
                dr("AcctNo") = AcctNo2.Text
            Case "2"
                '**by Milor 20080509--由訓練單位代轉現金時，所有轉帳資料都不填入值----start
                dr("AcctMode") = 2
                dr("PostNo") = Convert.DBNull
                dr("AcctNo") = Convert.DBNull
                dr("BankName") = Convert.DBNull
                dr("AcctHeadNo") = Convert.DBNull
                dr("ExBankName") = Convert.DBNull
                dr("AcctExNo") = Convert.DBNull
                '**by Milor 20080509----end
        End Select

        FirDate.Text = TIMS.ClearSQM(FirDate.Text)
        Uname.Text = TIMS.ClearSQM(Uname.Text)
        Intaxno.Text = TIMS.ClearSQM(Intaxno.Text)
        ActName.Text = TIMS.ClearSQM(ActName.Text)
        ActNo1.Text = TIMS.ClearSQM(ActNo1.Text)
        dr("FirDate") = If(FirDate.Text <> "", TIMS.Cdate2(FirDate.Text), Convert.DBNull)
        dr("Uname") = If(Uname.Text <> "", Uname.Text, Convert.DBNull)
        dr("Intaxno") = If(Intaxno.Text <> "", Intaxno.Text, Convert.DBNull)

        dr("ActName") = If(ActName.Text <> "", ActName.Text, Convert.DBNull)
        dr("ActNo") = If(ActNo1.Text <> "", ActNo1.Text, Convert.DBNull)
        'ActType: 投保類別1.勞2.農3.漁
        dr("ActType") = Convert.DBNull '"1"
        ServDept.Text = TIMS.ClearSQM(ServDept.Text)
        JobTitle.Text = TIMS.ClearSQM(JobTitle.Text)
        dr("ServDept") = TIMS.GetValue1(ServDept.Text)
        dr("JobTitle") = TIMS.GetValue1(JobTitle.Text)
        Dim t_ddlSERVDEPTID As String = TIMS.GetListText(ddlSERVDEPTID)
        Dim t_ddlJOBTITLEID As String = TIMS.GetListText(ddlJOBTITLEID)
        Dim v_ddlSERVDEPTID As String = TIMS.GetListValue(ddlSERVDEPTID)
        Dim v_ddlJOBTITLEID As String = TIMS.GetListValue(ddlJOBTITLEID)
        dr("ServDept") = If(t_ddlSERVDEPTID <> "", t_ddlSERVDEPTID, Convert.DBNull)
        dr("JobTitle") = If(t_ddlJOBTITLEID <> "", t_ddlJOBTITLEID, Convert.DBNull)
        dr("SERVDEPTID") = If(v_ddlSERVDEPTID <> "", v_ddlSERVDEPTID, Convert.DBNull)
        dr("JOBTITLEID") = If(v_ddlJOBTITLEID <> "", v_ddlJOBTITLEID, Convert.DBNull)
        Tel.Text = TIMS.ClearSQM(Tel.Text)
        Fax.Text = TIMS.ClearSQM(Fax.Text)
        dr("Tel") = If(Tel.Text <> "", Tel.Text, Convert.DBNull) 'Tel.Text
        dr("Fax") = If(Fax.Text <> "", Fax.Text, Convert.DBNull)
        SDate.Text = TIMS.ClearSQM(SDate.Text)
        SJDate.Text = TIMS.ClearSQM(SJDate.Text)
        SPDate.Text = TIMS.ClearSQM(SPDate.Text)
        dr("SDate") = TIMS.Cdate2(SDate.Text)
        dr("SJDate") = TIMS.Cdate2(SJDate.Text)
        dr("SPDate") = TIMS.Cdate2(SPDate.Text)

        '**by Milor 20081017--增加投保單位電話、地址
        txt_ActPhone.Text = TIMS.ClearSQM(txt_ActPhone.Text)
        dr("ActTel") = If(txt_ActPhone.Text <> "", txt_ActPhone.Text, Convert.DBNull)

        '投保單位地址
        txt_ActZip.Value = TIMS.ClearSQM(txt_ActZip.Value)
        txt_ActZIPB3.Value = TIMS.ClearSQM(txt_ActZIPB3.Value)
        hid_ActZIP6W.Value = TIMS.GetZIPCODE6W(txt_ActZip.Value, txt_ActZIPB3.Value)
        hidActZip_N.Value = TIMS.ClearSQM(hidActZip_N.Value)
        txt_ActAddress.Text = TIMS.ClearSQM(txt_ActAddress.Text)
        dr("ActZipCode") = TIMS.GetValue1(txt_ActZip.Value)
        dr("ActZipCode_6W") = TIMS.GetValue1(hid_ActZIP6W.Value)
        dr("ActZipCode_N") = TIMS.GetValue1(hidActZip_N.Value)
        dr("ActAddress") = TIMS.GetSpace(txt_ActAddress.Text)

        If Checkbox4.Checked Then '公司地址 同投保單位地址
            dr("Zip") = dr("ActZipCode") '= dr("Zip")
            dr("ZIP6W") = dr("ActZipCode_6W")
            dr("Zip_N") = dr("ActZipCode_N")
            dr("Addr") = TIMS.GetSpace(dr("ActAddress")) '=Addr.Text 'dr("Addr")
        Else
            '公司地址 'dr("Zip")=Zip.Value '20090209 andy edit
            Zip.Value = TIMS.ClearSQM(Zip.Value)
            ZIPB3.Value = TIMS.ClearSQM(ZIPB3.Value)
            hidZIP6W.Value = TIMS.GetZIPCODE6W(Zip.Value, ZIPB3.Value)
            Zip_N.Value = TIMS.ClearSQM(Zip_N.Value)
            Addr.Text = TIMS.ClearSQM(Addr.Text)
            If Zip.Value = "" Then Zip.Value = "-1" '空值加入-1
            dr("Zip") = TIMS.GetValue1(Zip.Value)
            dr("ZIP6W") = TIMS.GetValue1(hidZIP6W.Value)
            dr("Zip_N") = TIMS.GetValue1(Zip_N.Value)
            dr("Addr") = TIMS.GetSpace(Addr.Text)
        End If
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(dt)
        gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, dr)
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    ''' <summary>UPDATE STUD_TRAINBG</summary>
    ''' <param name="iSOCID"></param>
    Sub UPDATE_TRAINBG(ByRef iSOCID As Integer)
        Dim v_Q1 As String = TIMS.GetListValue(Q1)
        Dim v_Q3 As String = TIMS.GetListValue(Q3)
        Q3_Other.Text = TIMS.ClearSQM(Q3_Other.Text)
        Dim v_Q4 As String = TIMS.GetListValue(Q4)
        Dim v_Q5 As String = TIMS.GetListValue(Q5)

        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim sql As String = $" SELECT * FROM dbo.STUD_TRAINBG WHERE SOCID={iSOCID}"
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SOCID") = iSOCID
        Else
            dr = dt.Rows(0)
        End If

        dr("Q1") = Val(If(v_Q1 <> "", v_Q1, "0")) 'Val(Q1.SelectedValue)'numeric
        dr("Q3") = If(v_Q3 = "", Convert.DBNull, Val(v_Q3)) 'numeric
        dr("Q3_Other") = If(Q3_Other.Text = "", Convert.DBNull, Q3_Other.Text)
        dr("Q4") = If(v_Q4 = "", Convert.DBNull, v_Q4) 'varchar
        dr("Q5") = If(v_Q5 = "", Convert.DBNull, Val(v_Q5)) 'numeric

        dr("Q61") = If(Q61.Text = "", Convert.DBNull, Val(Q61.Text))
        dr("Q62") = If(Q62.Text = "", Convert.DBNull, Val(Q62.Text))
        dr("Q63") = If(Q63.Text = "", Convert.DBNull, Val(Q63.Text))
        dr("Q64") = If(Q64.Text = "", Convert.DBNull, Val(Q64.Text))
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(dt)
        gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, dr)
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    ''' <summary>UPDATE STUD_TRAINBGQ2</summary>
    ''' <param name="iSOCID"></param>
    Sub UPDATE_TRAINBGQ2(ByRef iSOCID As Integer)
        If iSOCID <= 0 Then Return

        Dim hPMS As New Hashtable From {{"SOCID", iSOCID}}
        Dim sql As String = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID=@SOCID"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, hPMS)
        If TIMS.dtHaveDATA(dt) Then
            Dim hPMS_d As New Hashtable From {{"SOCID", iSOCID}}
            Dim sql_d As String = " DELETE STUD_TRAINBGQ2 WHERE SOCID=@SOCID"
            DbAccess.ExecuteNonQuery(sql_d, objconn, hPMS_d)
        End If

        Dim v_Q2 As String = TIMS.GetCblValue(Q2)
        If v_Q2 = "" Then Return '無值離開

        Try
            For Each item As ListItem In Q2.Items
                If item.Value <> "" AndAlso item.Selected Then
                    v_Q2 = TIMS.ClearSQM(item.Value)
                    Dim hPMS2 As New Hashtable From {{"SOCID", iSOCID}, {"Q2", v_Q2}}
                    Dim sql2 As String = " SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID=@SOCID AND Q2=@Q2"
                    Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, hPMS2)
                    If TIMS.dtNODATA(dt2) Then
                        Dim hPMS_i As New Hashtable From {{"SOCID", iSOCID}, {"Q2", v_Q2}}
                        Dim sql_i As String = " INSERT INTO STUD_TRAINBGQ2 (SOCID,Q2) VALUES (@SOCID,@Q2)"
                        DbAccess.ExecuteNonQuery(sql_i, objconn, hPMS_i)
                    End If
                End If
            Next
        Catch ex As Exception
            '取得錯誤資訊寫入 'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Dim strErrmsg As String = $"#SD_03_002_add.aspx UPDATE_TRAINBGQ2:{vbCrLf},iSOCID:{iSOCID} ,v_Q2:{v_Q2}{vbCrLf},ex.Message: {ex.Message}{vbCrLf}{TIMS.GetErrorMsg(Me)}"
            Call TIMS.WriteTraceLog(strErrmsg, ex)
        End Try
    End Sub

    Sub UPDATE_ENTERTEMP12(ByRef iSOCID As Integer, vIDNO As String)
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing

        vIDNO = TIMS.ClearSQM(vIDNO)
        Dim v_Sex As String = TIMS.GetListValue(Sex)
        '**by Milor 20080527--產學訓當姓名、性別、生日、身分證號改變時，要回填報名資料----start
        '資料驗證的流程:
        '1.Class_StudentOfClass取得SETID、ETEnterDate、SerNum來對應Stud_EnterType，
        '  以驗證存在有報名職類檔，才去回填報名資料，避免回填到有問題資料。
        '2.E網報名資料STUD_ENTERTEMP2的SETID，因為某些User操作錯誤的因素等，
        '  導致有重複的SETID、IDNO、Name等資料，所以必須符合SETID、IDNO、Name、Sex、Birthday時，
        '  才將修改的資料回填。
        Dim hPMS As New Hashtable From {{"SOCID", iSOCID}}
        Dim sql As String = ""
        sql &= " SELECT SETID ,CONVERT(VARCHAR, ETEnterDate, 111) ETEnterDate ,SerNum" & vbCrLf
        sql &= " FROM CLASS_STUDENTSOFCLASS" & vbCrLf
        sql &= " WHERE SOCID=@SOCID AND SETID IS NOT NULL AND ETEnterDate IS NOT NULL AND SerNum IS NOT NULL"
        dt = DbAccess.GetDataTable(sql, objconn, hPMS)
        If dt.Rows.Count = 0 Then Return

        dr = dt.Rows(0)
        'Common.FormatDate()
        Dim uSETID As String = Convert.ToString(dr("SETID"))
        Dim uETED As String = Convert.ToString(dr("ETEnterDate"))
        Dim uSerNum As String = Convert.ToString(dr("SerNum"))
        '先驗證有報名職類檔，才進行檢查報名資料
        sql = $" SELECT SETID FROM STUD_ENTERTYPE WHERE SETID={uSETID} AND EnterDate={TIMS.To_date(uETED)} AND SerNum={uSerNum}"
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then Return

        da = Nothing
        '取出報名資料時，同時取出身分證號、姓名、性別、生日，以備後續判斷E網報名資料暫存檔
        sql = $" SELECT * FROM STUD_ENTERTEMP WHERE SETID={uSETID} AND IDNO='{vIDNO}'"
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then Return

        dr = dt.Rows(0)
        IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
        Birthday.Text = Common.FormatDate(Birthday.Text)

        Dim uIDNO As String = TIMS.ChangeIDNO(dr("IDNO"))
        Dim uName As String = TIMS.ClearSQM(dr("Name"))
        Dim uSex As String = TIMS.ClearSQM(dr("Sex"))
        Dim uBirthday As Date = Common.FormatDate(dr("Birthday"))
        Name.Text = TIMS.ClearSQM(Name.Text)

        'dr("IDNO")=IDNO.Text
        dr("Name") = Name.Text
        dr("Sex") = v_Sex
        dr("Birthday") = CDate(TIMS.Cdate2(Birthday.Text))
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now()
        DbAccess.UpdateDataTable(dt, da)

        '透過SETID取出E網報名資料暫存檔，比對身分證號、姓名、性別、生日與報名資料未更動前的資料一致時才變更
        'sql="select * from STUD_ENTERTEMP2 where SETID='" & uSETID & "' and IDNO='" & uIDNO & "' and Name='" & uName & "' and Sex='" & uSex & "' and Birthday='" & Convert.ToDateTime(uBirthday) & "'"
        'sql="select * from STUD_ENTERTEMP2 where SETID='" & uSETID & "' and upper(IDNO)='" & uIDNO & "' and Name='" & uName & "' and Sex='" & uSex & "' and Birthday='" & uBirthday & "'"
        uSETID = TIMS.ClearSQM(uSETID)
        uIDNO = TIMS.ClearSQM(uIDNO)
        uName = TIMS.ClearSQM(uName)
        uSex = TIMS.ClearSQM(uSex)
        uBirthday = TIMS.ClearSQM(uBirthday)
        sql = ""
        sql &= " SELECT * FROM STUD_ENTERTEMP2"
        sql &= $" WHERE SETID={uSETID} AND IDNO='{uIDNO}' AND Name='{uName}' AND Sex='{uSex}' AND Birthday={TIMS.To_date(uBirthday)}" 'fix ORA-01861
        da = Nothing
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then Return

        '萬一還是很不幸有多筆符合的資料，就只好全部都修改
        For i As Integer = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            'dr("IDNO")=IDNO.Text
            dr("Name") = Name.Text
            dr("Sex") = v_Sex
            dr("Birthday") = CDate(TIMS.Cdate2(Birthday.Text))
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now()
            DbAccess.UpdateDataTable(dt, da)
        Next
        '**by Milor 20080527----end
    End Sub

    ''' <summary> UPDATE CLASS_STUDENTSOFCLASS </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData4() As Boolean
        Dim rst As Boolean = False '儲存成功為True

        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        StudentIDValue.Value = TIMS.ClearSQM(StudentIDValue.Value)
        StudentID.Text = TIMS.ClearSQM(StudentID.Text)
        Hid_SID_C1.Value = TIMS.ClearSQM(Hid_SID_C1.Value)
        Dim vSID As String = Hid_SID_C1.Value
        If IDNO.Text = "" OrElse vSID = "" OrElse StudentIDValue.Value = "" OrElse StudentID.Text = "" Then Return False
        '(檢核SID與IDNO)
        Dim fg_CHK_SID_IDNO As Boolean = CHECK_SID_IDNO(objconn, vSID, IDNO.Text)
        '(檢核有誤，查無資料離開儲存功能)
        If Not fg_CHK_SID_IDNO Then Return rst

        Dim all_Identity2 As String = Get_All_Identity2()

        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        Dim v_MilitaryID As String = TIMS.GetListValue(MilitaryID)
        Dim vSOCID_Sel As String = TIMS.GetListValue(SOCID)
        If vSOCID_Sel = "" Then Return False

        '津貼類別  不等於 03:就業促進津貼實施辦法 刪除請領津貼的資料。
        Call UPDATE_SUBSIDYRESULT(vSOCID_Sel)
        '更新MakeSOCID
        Dim vRejectSOCID As String = UPDATE_MakeSOCID()

        Dim sql As String = ""
        'UPDATE CLASS_STUDENTSOFCLASS
        Dim iSOCID As Integer = If(vSOCID_Sel <> "" AndAlso vSOCID_Sel <> "0", Val(vSOCID_Sel), 0)  'String="" 0: 新增1筆學員資料 / >0: 修改學員資料
        Dim bflagNew1 As Boolean = If(Not StdTr.Visible OrElse iSOCID = 0, True, False) 'TRUE: 新增一筆 /FALSE: 修改動作1筆
        If Not bflagNew1 Then '修改檢核
            sql = $"SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID={iSOCID}"
            Dim dtCS As DataTable = DbAccess.GetDataTable(sql, objconn)
            If dtCS.Rows.Count <> 1 Then bflagNew1 = True '新增一筆 '資料不等於1筆(異常新增)
        End If

        Dim v_EnterChannel As String = TIMS.GetListValue(EnterChannel)
        Dim v_TRNDMode As String = TIMS.GetListValue(TRNDMode)
        Dim v_StudentID As String = $"{StudentIDValue.Value}{StudentID.Text}"

        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        Dim Mydr3 As DataRow = Nothing
        Dim MyTable3 As DataTable = Nothing
        Dim da3 As SqlDataAdapter = Nothing
        If bflagNew1 Then
            iSOCID = DbAccess.GetNewId(objconn, "CLASS_STUDENTSOFCLASS_SOCID_SE,CLASS_STUDENTSOFCLASS,SOCID")
            '新增
            sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE 1<>1 "
            'CLASS_STUDENTSOFCLASS_SOCID_SE
            MyTable3 = DbAccess.GetDataTable(sql, da3, objconn)
            '查無資料新增1筆
            Mydr3 = MyTable3.NewRow
            MyTable3.Rows.Add(Mydr3)
            Mydr3("SOCID") = iSOCID
            Mydr3("SID") = vSID '取得目前最新的SID
            Mydr3("StudStatus") = 1
            Mydr3("OCID") = TIMS.CINT1(rqOCID)
        Else
            '修改 'UPDATE CLASS_STUDENTSOFCLASS
            sql = $"SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID={iSOCID}"
            'iSOCIDValue=SOCID.SelectedValue
            MyTable3 = DbAccess.GetDataTable(sql, da3, objconn)
            Mydr3 = MyTable3.Rows(0)
            Mydr3("SID") = vSID '取得目前最新的SID
        End If

        '更新班級學員檔-----   Start
        Mydr3("RejectSOCID") = If(vRejectSOCID <> "", vRejectSOCID, Convert.DBNull) '被遞補者 學員
        Mydr3("StudentID") = v_StudentID 'StudentIDValue.Value & StudentID.Text
        '報名階段:無效塞0
        Dim vLevelNo As String = TIMS.GetListValue(LevelNo) '.ClearSQM(LevelNo.SelectedValue)
        vLevelNo = If(vLevelNo <> "", If(LevelNo.Enabled, vLevelNo, "0"), "0")
        Mydr3("LevelNo") = vLevelNo
        Mydr3("EnterDate") = If(EnterDate.Text <> "", TIMS.Cdate2(EnterDate.Text), Convert.DBNull)
        Mydr3("OpenDate") = If(OpenDate.Text <> "", TIMS.Cdate2(OpenDate.Text), Convert.DBNull)
        Mydr3("CloseDate") = If(CloseDate.Text <> "", TIMS.Cdate2(CloseDate.Text), Convert.DBNull)
        Mydr3("RejectTDate1") = If(RejectTDate1.Text <> "", TIMS.Cdate2(RejectTDate1.Text), Convert.DBNull)
        Mydr3("RejectTDate2") = If(RejectTDate2.Text <> "", TIMS.Cdate2(RejectTDate2.Text), Convert.DBNull)
        Mydr3("TRNDMode") = If(v_TRNDMode <> "", v_TRNDMode, Convert.DBNull)

        'Dim MyVal As String=""
        'MyVal=TIMS.ClearSQM(EnterChannel.SelectedValue)
        'EnterChannel: '1.網;2.現;3.通;4.推
        Dim MyVal As String = ""
        MyVal = v_EnterChannel
        If v_EnterChannel = "" Then
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '產投預設為網路報名
                MyVal = "1" 'Convert.DBNull'預設:1.網 
            Else
                MyVal = "2" 'Convert.DBNull'預設:2.現
                If hide_TrainMode.Value <> "" Then MyVal = "4" 'Convert.DBNull'4.推
            End If
        End If
        Common.SetListItem(EnterChannel, MyVal)
        v_EnterChannel = TIMS.GetListValue(EnterChannel)
        Mydr3("EnterChannel") = MyVal
        Dim vMIdentityID As String = TIMS.GetListValue(MIdentityID)
        Dim v_DDL_DISASTER As String = TIMS.GetListValue(DDL_DISASTER) 'ADID 重大災害選項
        If tr_DDL_DISASTER.Visible AndAlso vMIdentityID <> TIMS.cst_Identity_40 AndAlso v_DDL_DISASTER <> "" Then v_DDL_DISASTER = "" '(非)「經公告之重大災害受災者」清除「重大災害選項」

        'Dim vMIdentityID As String=TIMS.ClearSQM(MIdentityID.SelectedValue)
        '預設 一般身分者:01 (在職進修訓練:06)
        If vMIdentityID = "" AndAlso sm.UserInfo.TPlanID = "06" Then vMIdentityID = "01"
        Mydr3("MIdentityID") = vMIdentityID 'MIdentityID.SelectedValue
        Mydr3("ADID") = If(v_DDL_DISASTER <> "", v_DDL_DISASTER, Convert.DBNull) 'ADID 重大災害選項

        'Dim all_Identity As String=""
        If IdentityID.Enabled Then 'false時，不寫入
            '不管如何都要得知 all_Identity
            Mydr3("IdentityID") = all_Identity2 'all_Identity
        End If
        If Not IdentityID.Enabled Then
            '原呼叫的資料
            'Enabled 為 False 不做儲存
            If ViewState(vs_IdentityID) IsNot Nothing Then all_Identity2 = ViewState(vs_IdentityID)
        End If

        'Mydr3("Native")=Convert.DBNull
        '原住民
        'If InStr(all_Identity2, "05") > 0 OrElse MIdentityID.SelectedValue="05" Then
        '    'by Kevin 95.10.27 原住民才寫入民族別
        '    'by Vicient
        '    If NativeID.SelectedValue <> "" AndAlso NativeID.SelectedIndex <> 0 Then Mydr3("Native")=NativeID.SelectedValue
        'End If

        Dim v_SubsidyID As String = "" 'TIMS.GetListValue(SubsidyID)
        Dim v_SubsidyIdentity As String = "" 'TIMS.GetListValue(SubsidyIdentity)
        If trSubsidyID.Visible Then
            v_SubsidyID = TIMS.GetListValue(SubsidyID)
            '預設 未申請:01 (在職進修訓練:06)
            If v_SubsidyID = "" AndAlso sm.UserInfo.TPlanID = "06" Then v_SubsidyID = "01"
        End If
        If trSubsidyIdentity.Visible Then v_SubsidyIdentity = TIMS.GetListValue(SubsidyIdentity)
        Mydr3("SubsidyID") = If(v_SubsidyID <> "", v_SubsidyID, Convert.DBNull)
        Mydr3("SubsidyIdentity") = If(v_SubsidyIdentity <> "", v_SubsidyIdentity, Convert.DBNull)

        '------ 受訓前任職資料start 2010/04/27 開始存class_classinfo 
        If trPriorWorkOrg1.Visible AndAlso trTable6.Visible Then
            PriorWorkOrg1.Text = TIMS.ClearSQM(PriorWorkOrg1.Text)
            If Len(PriorWorkOrg1.Text) > 30 Then PriorWorkOrg1.Text = Mid(PriorWorkOrg1.Text, 1, 30)
            Dim v_PriorWorkType1 As String = TIMS.GetListValue(PriorWorkType1)
            Mydr3("PWType1") = If(v_PriorWorkType1 <> "", v_PriorWorkType1, Convert.DBNull)
            Mydr3("PWOrg1") = If(PriorWorkOrg1.Text <> "", PriorWorkOrg1.Text, Convert.DBNull)
            Mydr3("SOfficeYM1") = If(SOfficeYM1.Text <> "", SOfficeYM1.Text, Convert.DBNull)
            Mydr3("FOfficeYM1") = If(FOfficeYM1.Text <> "", FOfficeYM1.Text, Convert.DBNull)
        End If
        '------ 受訓前任職資料end 

        '學習券要判斷上課單元 Start
        If sm.UserInfo.TPlanID = "15" Then
            If Unit1Hour.Text <> "" Then Unit1Hour.Text = Trim(Unit1Hour.Text)
            If Unit2Hour.Text <> "" Then Unit2Hour.Text = Trim(Unit2Hour.Text)
            If Unit3Hour.Text <> "" Then Unit3Hour.Text = Trim(Unit3Hour.Text)
            If Unit4Hour.Text <> "" Then Unit4Hour.Text = Trim(Unit4Hour.Text)

            If Unit1Hour.Text <> "" Then Unit1Hour.Text = Val(Unit1Hour.Text)
            If Unit2Hour.Text <> "" Then Unit2Hour.Text = Val(Unit2Hour.Text)
            If Unit3Hour.Text <> "" Then Unit3Hour.Text = Val(Unit3Hour.Text)
            If Unit4Hour.Text <> "" Then Unit4Hour.Text = Val(Unit4Hour.Text)

            Mydr3("RelClass_Unit") = ""
            Mydr3("RelClass_Hour") = ""
            Mydr3("Unit1Hour") = If(Unit1Hour.Text = "", 0, Val(Unit1Hour.Text))
            Mydr3("Unit2Hour") = If(Unit2Hour.Text = "", 0, Val(Unit2Hour.Text))
            Mydr3("Unit3Hour") = If(Unit3Hour.Text = "", 0, Val(Unit3Hour.Text))
            Mydr3("Unit4Hour") = If(Unit4Hour.Text = "", 0, Val(Unit4Hour.Text))
            'add by nick 060316
            Mydr3("Unit1Score") = If(Unit1Score.Text = "", 0, Val(Unit1Score.Text))
            Mydr3("Unit2Score") = If(Unit2Score.Text = "", 0, Val(Unit2Score.Text))
            Mydr3("Unit3Score") = If(Unit3Score.Text = "", 0, Val(Unit3Score.Text))
            Mydr3("Unit4Score") = If(Unit4Score.Text = "", 0, Val(Unit4Score.Text))

            Dim RelClassUnitV As String = "" '0000~1111
            'RelClassUnitV=""
            For i As Integer = 0 To RelClass_Unit.Items.Count - 1 '0~3
                RelClassUnitV &= If(RelClass_Unit.Items(i).Selected, "1", "0")
            Next
            Mydr3("RelClass_Unit") = RelClassUnitV

            '若為空補0
            If Unit1Hour.Text = "" Then Unit1Hour.Text = "0"
            If Unit2Hour.Text = "" Then Unit2Hour.Text = "0"
            If Unit3Hour.Text = "" Then Unit3Hour.Text = "0"
            If Unit4Hour.Text = "" Then Unit4Hour.Text = "0"

            Dim RelClassHourV As String = ""
            Call SUtl_CboRelClassHourV(RelClassHourV, CInt(Unit1Hour.Text))
            Call SUtl_CboRelClassHourV(RelClassHourV, CInt(Unit2Hour.Text))
            Call SUtl_CboRelClassHourV(RelClassHourV, CInt(Unit3Hour.Text))
            Call SUtl_CboRelClassHourV(RelClassHourV, CInt(Unit4Hour.Text))
            Mydr3("RelClass_Hour") = RelClassHourV
        Else
            Mydr3("RelClass_Unit") = Convert.DBNull
            Mydr3("RelClass_Hour") = Convert.DBNull
            Mydr3("Unit1Hour") = 0
            Mydr3("Unit2Hour") = 0
            Mydr3("Unit3Hour") = 0
            Mydr3("Unit4Hour") = 0

            Mydr3("Unit1Score") = 0
            Mydr3("Unit2Score") = 0
            Mydr3("Unit3Score") = 0
            Mydr3("Unit4Score") = 0
        End If
        '學習券要判斷上課單元 End

        Dim v_BudID As String = TIMS.GetListValue(BudID)
        Mydr3("BudgetID") = If(v_BudID <> "", v_BudID, Convert.DBNull)
        'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
        Dim v_SupplyID As String = TIMS.GetListValue(SupplyID)
        If v_SupplyID = "0" Then v_SupplyID = "" '等於0，等同未選擇
        If v_BudID = "99" Then v_SupplyID = "9" '不補助 0%
        Mydr3("SupplyID") = If(v_SupplyID <> "", v_SupplyID, Convert.DBNull)

        Dim v_AppliedResult As String = Convert.ToString(Mydr3("AppliedResult"))
        If v_BudID = "99" AndAlso v_SupplyID = "9" Then v_AppliedResult = "N"  '學員資料複審狀態=N
        Mydr3("AppliedResult") = If(v_AppliedResult <> "", v_AppliedResult, Convert.DBNull) '學員資料複審狀態=N

        ''選錯為 2: 特定100% , 且為一般身分, 且為 就安 就保。
        'If SupplyID.SelectedValue="2" AndAlso MIdentityID.SelectedValue="01" AndAlso all_Identity2="01" Then Mydr3("SupplyID")="1"
        Dim v_PMode As String = "" 'TIMS.GetListValue(PMode)
        If trPMode.Visible Then v_PMode = TIMS.GetListValue(PMode)
        Mydr3("PMode") = If(v_PMode <> "", v_PMode, Convert.DBNull)

        Dim v_ActNoXX As String = ""
        ActNo.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(ActNo.Text))
        ActNo1.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(ActNo1.Text))
        ActNo2.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(ActNo2.Text))
        If TPlan23TR.Visible Then
            v_ActNoXX = If(ActNo.Text <> "", ActNo.Text, "")
        Else
            'ActNo1 產投 ActNo2 '其他一般計畫
            v_ActNoXX = If(ActNo1.Text <> "", ActNo1.Text, If(ActNo2.Text <> "", ActNo2.Text, ""))
        End If
        Mydr3("ActNo") = If(v_ActNoXX <> "", v_ActNoXX, Convert.DBNull)

        Mydr3("ModifyAcct") = sm.UserInfo.UserID
        Mydr3("ModifyDate") = Now
        '20090330(Milot)專上畢業學歷失業者
        Dim v_HighEduBg As String = TIMS.GetListValue(rdo_HighEduBg)
        Mydr3("HighEduBg") = If(v_HighEduBg <> "", v_HighEduBg, Convert.DBNull)

        '是否為在職者補助身分
        Dim v_WorkSuppIdent As String = TIMS.GetListValue(rblWorkSuppIdent)
        Select Case v_WorkSuppIdent
            Case "Y", "N"
            Case Else
                v_WorkSuppIdent = ""
        End Select
        Mydr3("WorkSuppIdent") = If(v_WorkSuppIdent <> "", v_WorkSuppIdent, Convert.DBNull)

        '同意 本署將學員個人資料提供社家署做就業媒合之用
        '46,47 補助辦理保母職業訓練、補助辦理照顧服務員職業訓練
        '58:補助辦理托育人員職業訓練()
        Dim v_rblHouseMatch As String = TIMS.GetListValue(rblHouseMatch)
        'Dim vHouseMatch As String=""
        If trHouseMatch.Visible Then
            Select Case v_rblHouseMatch
                Case "Y", "N" 'rblHouseMatch
                    'vHouseMatch=v_rblHouseMatch
            End Select
        End If
        Mydr3("HouseMatch") = If(v_rblHouseMatch <> "", v_rblHouseMatch, Convert.DBNull)

        Dim vPriorWorkPay As String = ""
        PriorWorkPay.Text = TIMS.ClearSQM(PriorWorkPay.Text)
        vPriorWorkPay = If(ActNo2_TR.Visible, PriorWorkPay.Text, "")
        Mydr3("PriorWorkPay") = If(vPriorWorkPay <> "", vPriorWorkPay, Convert.DBNull)

        If trPriorWorkOrg1.Visible AndAlso trPriorWorkOrg2.Visible Then
            Title1.Text = TIMS.ClearSQM(Title1.Text)
            Title2.Text = TIMS.ClearSQM(Title2.Text)
            Mydr3("Title1") = If(Title1.Text <> "", Title1.Text, Convert.DBNull)
            Mydr3("Title2") = If(Title2.Text <> "", Title2.Text, Convert.DBNull)
        End If

        Mydr3("JoblessID") = Convert.DBNull 'JoblessID.SelectedValue
        Mydr3("RealJobless") = Convert.DBNull 'RealJobless.Text
        Try
            gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(MyTable3)
            gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, Mydr3)
            DbAccess.UpdateDataTable(MyTable3, da3)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("SID:{0},IDNO:{1},rqOCID:{2},iSOCID:{3}", vSID, IDNO.Text, TIMS.CINT1(rqOCID), iSOCID) & vbCrLf
            strErrmsg &= String.Format("StudentIDValue:{0},StudentID:{1}", StudentIDValue.Value, StudentID.Text) & vbCrLf
            strErrmsg &= String.Format("bflagNew1:{0}", bflagNew1.ToString()) & vbCrLf
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Common.MessageBox(Me, "儲存學員資料有誤，請重新操作!")
            Return rst
            'Throw ex
        End Try
        '更新班級學員檔-----   End

        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        '修改預算別(同步修改補助申請的預算別)'https://cm.turbotech.com.tw/browse/TIMS-2150
        v_BudID = TIMS.GetListValue(BudID)
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso v_BudID <> "" Then
            Dim u_Parms As New Hashtable From {{"BUDID", v_BudID}, {"MODIFYACCT", sm.UserInfo.UserID}, {"SOCID", iSOCID}}
            Dim u_sql As String = ""
            u_sql &= " UPDATE STUD_SUBSIDYCOST"
            u_sql &= " SET BUDID=@BUDID ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()"
            u_sql &= " WHERE SOCID=@SOCID"
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_Parms)  'edit，by:20181017
        End If

        '企訓專用 Start
        '46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        '是否為在職者補助身分(rblWorkSuppIdent) 選 Y 可存取
        'If sm.UserInfo.TPlanID="28" Or (sm.UserInfo.TPlanID="46" And rblWorkSuppIdent.SelectedValue="Y") Or (sm.UserInfo.TPlanID="47" And rblWorkSuppIdent.SelectedValue="Y") Then
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If Hid_show_actno_budid.Value = "Y" Then
            Call UPDATE_SERVICEPLACE(iSOCID) 'UPDATE STUD_SERVICEPLACE
            Call UPDATE_TRAINBG(iSOCID) 'UPDATE STUD_TRAINBG
            Call UPDATE_TRAINBGQ2(iSOCID) 'UPDATE STUD_TRAINBGQ2
            Call UPDATE_ENTERTEMP12(iSOCID, IDNO.Text)
        End If
        '企訓專用 End

        '結訓學員資料卡更新 Start
        If StdTr.Visible = True Then
            Dim iSex As Integer = (Sex.SelectedIndex + 1)
            da = Nothing
            sql = $" SELECT * FROM STUD_RESULTSTUDDATA WHERE SOCID={vSOCID_Sel}"
            dt = DbAccess.GetDataTable(sql, da, objconn)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                Dim DLID As Integer = dr("DLID")
                Dim SubNo As Integer = dr("SubNo")
                dr("StdName") = Name.Text
                dr("StudentID") = StudentID.Text
                dr("StdPID") = IDNO.Text 'TIMS.ChangeIDNO(IDNO.Text)
                dr("Sex") = If(iSex = 1, iSex, If(iSex = 2, iSex, Convert.DBNull)) 'Sex.SelectedIndex + 1 '1.男2.女
                dr("BirthYear") = CDate(Birthday.Text).Year
                dr("BirthMonth") = CDate(Birthday.Text).Month
                dr("BirthDate") = CDate(Birthday.Text).Day
                dr("DegreeID") = v_DegreeID
                dr("MilitaryID") = If(v_MilitaryID = "", Convert.DBNull, v_MilitaryID)
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                DbAccess.UpdateDataTable(dt, da)

                'Stud_ResultIdentData不在匯入此 TABLE 改用 CLASS_STUDENTSOFCLASS.IdentityID	參訓身分別代碼
                'BY AMU 2009-07-30
                '非署(局)屬狀況加入 Stud_ResultIdentData  BY AMU 2009-08-25
                da = Nothing
                sql = "SELECT * FROM STUD_RESULTIDENTDATA WHERE DLID='" & DLID & "' and SubNo='" & SubNo & "'"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                For Each item As ListItem In IdentityID.Items
                    If item.Selected = True Then
                        If dt.Select("IdentityID='" & item.Value & "'").Length = 0 Then
                            dr = dt.NewRow
                            dt.Rows.Add(dr)
                            dr("DLID") = DLID
                            dr("SubNo") = SubNo
                            dr("IdentityID") = item.Value
                        End If
                    Else
                        If dt.Select("IdentityID='" & item.Value & "'").Length <> 0 Then dt.Select("IdentityID='" & item.Value & "'")(0).Delete()
                    End If
                Next
                DbAccess.UpdateDataTable(dt, da)
            End If
        End If
        '結訓學員資料卡更新 End

        '20100415 andy 報名管道有異動時要更新 Stud_EnterType(因為報名管道太多，不一定全部從CLASS_STUDENTSOFCLASS有辦法回推，暫時只能這樣)
        sql = ""
        sql &= " SELECT SETID,CONVERT(VARCHAR, ETEnterDate, 111) ETEnterDate ,SerNum "
        sql &= " FROM CLASS_STUDENTSOFCLASS WITH(NOLOCK)"
        sql &= " WHERE SOCID='" & iSOCID & "' "
        sql &= " AND SETID IS NOT NULL "
        sql &= " AND ETEnterDate IS NOT NULL "
        sql &= " AND SerNum IS NOT NULL "
        dt = DbAccess.GetDataTable(sql, objconn)

        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
            Dim uSETID As String = dr("SETID").ToString
            Dim uETED As String = dr("ETEnterDate").ToString
            Dim uSerNum As String = dr("SerNum").ToString
            sql = ""
            sql &= " SELECT * FROM STUD_ENTERTYPE "
            sql &= " WHERE SETID='" & uSETID & "'"
            sql &= " AND EnterDate=" & TIMS.To_date(uETED)
            sql &= " AND SerNum='" & uSerNum & "'"
            dt = DbAccess.GetDataTable(sql, da, objconn)

            If dt.Rows.Count > 0 Then
                Dim vTRNDMode As String = ""
                Select Case v_EnterChannel'EnterChannel.SelectedValue
                    Case "4" '4.推
                        Select Case v_TRNDMode'.SelectedValue
                            Case "1", "2", "3" '1.職2.學3.推
                                vTRNDMode = v_TRNDMode
                        End Select
                        'Select Case TRNDType.SelectedValue
                        '    Case "1", "2" '1.甲式2.乙式
                        '        dr("TRNDType")=TRNDType.SelectedValue
                        'End Select
                End Select
                dr = dt.Rows(0)
                'Dim v_TRNDMode As String=TIMS.GetListValue(TRNDMode)
                '1.網;2.現;3.通;4.推--'1.網(自行預設網路)
                dr("EnterChannel") = If(v_EnterChannel <> "", v_EnterChannel, "1") 'EnterChannel.SelectedValue
                dr("TRNDMode") = Convert.DBNull
                dr("TRNDType") = Convert.DBNull
                dr("TRNDMode") = If(vTRNDMode <> "", vTRNDMode, Convert.DBNull) 'TRNDMode.SelectedValue '1.職2.學3.推
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now()
                DbAccess.UpdateDataTable(dt, da)
            End If
        End If
        'DbAccess.CommitTrans(objTrans)
        'Call TIMS.CloseDbConn(tConn)

        Call SUtl_SaveData5(iSOCID)
        rst = True
        Return rst
    End Function

    ''' <summary>'檢核SID與IDNO是否正確</summary>
    ''' <param name="oConn"></param>
    ''' <param name="vSID"></param>
    ''' <param name="tIDNO"></param>
    ''' <returns></returns>
    Private Function CHECK_SID_IDNO(ByRef oConn As SqlConnection, vSID As String, tIDNO As String) As Boolean
        Dim hPMS As New Hashtable From {{"SID", vSID}, {"IDNO", tIDNO}}
        Dim sql_s As String = "SELECT 1 FROM STUD_STUDENTINFO WITH(NOLOCK) WHERE SID=@SID AND IDNO=@IDNO" '找出SID
        Dim dt As DataTable = DbAccess.GetDataTable(sql_s, oConn, hPMS)
        If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then Return True '(查有資料，正常True)
        Return False '(異常False)
    End Function

    ''' <summary>UPDATE STUD_BLIGATEDATA28 </summary>
    ''' <param name="str_SOCID"></param>
    ''' <returns></returns>
    Function SUtl_SaveData5(ByVal str_SOCID As String) As Boolean
        '首頁>>學員動態管理>>表單列印>>參訓學員投保狀況檢核表
        '依專案執行檢討會議決議（附件1），學員資料維護之「保險證號」與「投保單位名稱」
        '將與參訓學員投保狀況檢核表勾稽到的「投保單位」及「保險證號」連動，若單位於學員資料維護勾選其他「保險證號」
        '即會一併連動修改參訓學員投保狀況檢核表之「投保單位」及「保險證號」。
        '目前有參訓學員共有2個投保證號， 於學員資料維護勾選第2個投保證號後（附件2）
        '至參訓學員投保狀況檢核表列印時並未將該學員投保證號連動修正（附件3）， 請儘速修正。  
        'End If
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        If Hid_show_actno_budid.Value = "Y" Then
            If hidSB4ID.Value <> "" AndAlso ActNo1.Text <> "" Then
                'hssT.Add("ActName", ActName.Text)
                Dim hssT As New Hashtable From {
                    {"ACTNO1", ActNo1.Text},
                    {"SB4ID", hidSB4ID.Value},
                    {"IDNO", IDNO.Text},
                    {"OCID1", Hid_OCID.Value},
                    {"SOCID", str_SOCID},
                    {"USERID", sm.UserInfo.UserID}
                }
                Call TIMS.UPDATE_STUD_BLIGATEDATA28(hssT, objconn)
            End If
        End If
    End Function
    '選擇了該班級另一位學員資料
    Private Sub SOCID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SOCID.SelectedIndexChanged
        If SOCID.SelectedIndex <> 0 Then
            Call Clear_data()   '清理資料
            Call GetOpenDate2()   '塞入班級資料 (排第2順位) '重要
            Dim vSOCID_Sel As String = TIMS.ClearSQM(SOCID.SelectedValue)
            Call Create1_Stud(vSOCID_Sel) '塞入學員資料
            Call GetScript(vSOCID_Sel) '學員資料審核功能的欄位鎖住
            'Call GetOpenDate2()   '塞入班級資料 (排第2順位) '重要
        End If
    End Sub

    ''' <summary>取得原資料</summary>
    ''' <returns></returns>
    Function Get_STUDINFOdr() As DataRow
        Dim dr As DataRow = Nothing
        IDNO.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(IDNO.Text))
        If IDNO.Text = "" Then Return dr

        Dim hPMS As New Hashtable From {{"IDNO", IDNO.Text}}
        Dim sql As String = ""
        sql &= " SELECT a.SID" & vbCrLf '/*PK*/
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " ,a.NAME" & vbCrLf
        sql &= " ,a.RMPNAME" & vbCrLf
        sql &= " ,a.ENGNAME" & vbCrLf
        sql &= " ,a.PASSPORTNO" & vbCrLf
        sql &= " ,a.SEX" & vbCrLf
        sql &= " ,a.BIRTHDAY" & vbCrLf
        sql &= " ,a.MARITALSTATUS" & vbCrLf
        sql &= " ,a.DEGREEID" & vbCrLf
        sql &= " ,a.GRADUATESTATUS" & vbCrLf
        sql &= " ,a.MILITARYID" & vbCrLf
        sql &= " ,a.IDENTITYID" & vbCrLf
        'sql &= " ,a.SUBSIDYID" & vbCrLf
        'sql &= " ,a.JOBLESSID" & vbCrLf
        'sql &= " ,a.REALJOBLESS" & vbCrLf
        'sql &= " ,a.GETCERTIFICATE" & vbCrLf
        'sql &= " ,a.GETSUBSIDY" & vbCrLf
        'sql &= " ,a.ISAGREE" & vbCrLf
        'sql &= " ,a.CHINAORNOT" & vbCrLf
        'sql &= " ,a.NATIONALITY" & vbCrLf
        'sql &= " ,a.PPNO" & vbCrLf
        sql &= " ,a.JOBSTATE" & vbCrLf
        'sql &= " ,a.FTYPE" & vbCrLf
        'sql &= " ,a.ACTNO" & vbCrLf
        'sql &= " ,a.MDATE" & vbCrLf
        'sql &= " ,a.SALID" & vbCrLf
        'sql &= " ,a.FIXID" & vbCrLf
        'sql &= " ,a.JOBLESSID_99" & vbCrLf
        sql &= " ,a.GRADUATEY" & vbCrLf
        sql &= " ,b.SCHOOL" & vbCrLf
        sql &= " ,b.DEPARTMENT" & vbCrLf
        sql &= " ,b.ZIPCODE1" & vbCrLf
        sql &= " ,b.ZIPCODE1_6W" & vbCrLf
        sql &= " ,b.ZIPCODE1_N" & vbCrLf
        sql &= " ,b.ADDRESS" & vbCrLf
        sql &= " ,b.ZIPCODE2" & vbCrLf
        sql &= " ,b.ZIPCODE2_6W" & vbCrLf
        sql &= " ,b.ZIPCODE2_N" & vbCrLf
        sql &= " ,b.HOUSEHOLDADDRESS" & vbCrLf
        sql &= " ,b.EMAIL" & vbCrLf
        sql &= " ,b.PHONED" & vbCrLf
        sql &= " ,b.PHONEN" & vbCrLf
        sql &= " ,b.CELLPHONE" & vbCrLf
        sql &= " ,b.EMERGENCYCONTACT" & vbCrLf
        sql &= " ,b.EMERGENCYRELATION" & vbCrLf
        sql &= " ,b.EMERGENCYPHONE" & vbCrLf
        sql &= " ,b.ZIPCODE3" & vbCrLf
        sql &= " ,b.ZIPCODE3_6W" & vbCrLf
        sql &= " ,b.ZIPCODE3_N" & vbCrLf
        sql &= " ,b.EMERGENCYADDRESS" & vbCrLf
        sql &= " ,b.PRIORWORKORG1" & vbCrLf
        sql &= " ,b.TITLE1" & vbCrLf
        sql &= " ,b.TITLE2" & vbCrLf
        sql &= " ,b.SOFFICEYM1" & vbCrLf
        sql &= " ,b.FOFFICEYM1" & vbCrLf
        sql &= " ,b.SOFFICEYM2" & vbCrLf
        sql &= " ,b.FOFFICEYM2" & vbCrLf
        'sql &= " ,b.PRIORWORKORG2" & vbCrLf
        'sql &= " ,b.PRIORWORKPAY" & vbCrLf
        sql &= " ,b.TRAFFIC" & vbCrLf
        sql &= " ,b.SHOWDETAIL" & vbCrLf
        sql &= " ,b.SERVICEID" & vbCrLf
        sql &= " ,b.MILITARYAPPOINTMENT" & vbCrLf
        sql &= " ,b.MILITARYRANK" & vbCrLf
        sql &= " ,b.SSERVICEDATE" & vbCrLf
        sql &= " ,b.FSERVICEDATE" & vbCrLf
        sql &= " ,b.SERVICEORG" & vbCrLf
        'sql &= " ,b.CHIEFRANKNAME" & vbCrLf
        sql &= " ,b.ZIPCODE4" & vbCrLf
        sql &= " ,b.ZIPCODE4_6W" & vbCrLf
        sql &= " ,b.ZIPCODE4_N" & vbCrLf
        sql &= " ,b.SERVICEADDRESS" & vbCrLf
        sql &= " ,b.SERVICEPHONE" & vbCrLf
        sql &= " ,b.HANDTYPEID" & vbCrLf
        sql &= " ,b.HANDLEVELID" & vbCrLf
        sql &= " ,b.HANDTYPEID2" & vbCrLf
        sql &= " ,b.HANDLEVELID2" & vbCrLf
        'sql &= " ,CASE WHEN b.HANDTYPEID2 IS NOT NULL AND b.HANDLEVELID2 IS NOT NULL THEN '2' ELSE '1' END HANDTYPE" & vbCrLf

        'sql &= " ,b.FORENAME" & vbCrLf
        'sql &= " ,b.FORETITLE" & vbCrLf
        'sql &= " ,b.FORESEX" & vbCrLf
        'sql &= " ,b.FOREBIRTH" & vbCrLf
        'sql &= " ,b.FOREIDNO" & vbCrLf
        'sql &= " ,b.FOREZIP" & vbCrLf
        'sql &= " ,b.FOREADDR" & vbCrLf
        'sql &= " ,b.FOREZIP6W" & vbCrLf
        'sql &= " SELECT a.* ,b.*" & vbCrLf

        sql &= " ,cs.PWType1" & vbCrLf
        sql &= " ,cs.PWOrg1" & vbCrLf
        sql &= " ,cs.SOfficeYM1 SOfficeYM3" & vbCrLf
        sql &= " ,cs.FOfficeYM1 FOfficeYM3" & vbCrLf
        sql &= " ,cs.ActNO ACTNO2" & vbCrLf
        sql &= " FROM STUD_STUDENTINFO a" & vbCrLf
        sql &= " JOIN STUD_SUBDATA b ON a.SID=b.SID" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs ON cs.SID=b.SID" & vbCrLf
        sql &= " WHERE a.IDNO=@IDNO" & vbCrLf
        sql &= " ORDER BY cs.SOCID DESC" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn, hPMS)

        Return dr
    End Function
    ''' <summary>取得原資料-顯示</summary>
    Sub SHOW_STUDINFO(ByRef dr As DataRow)
        If dr Is Nothing Then Return

        Name.Text = Convert.ToString(dr("Name")) '.ToString
        RMPNAME.Text = Convert.ToString(dr("RMPNAME")) '.ToString
        If dr("EngName").ToString.IndexOf(" ") = -1 Then
            '沒有空格。
            LName.Text = dr("EngName").ToString
        Else
            '有空格。
            LName.Text = Trim(Left(dr("EngName").ToString, dr("EngName").ToString.IndexOf(" ")))
            FName.Text = Trim(Right(dr("EngName").ToString, dr("EngName").ToString.Length - 1 - dr("EngName").ToString.IndexOf(" ")))
        End If
        Common.SetListItem(PassPortNO, dr("PassPortNO").ToString)
        IDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)
        Common.SetListItem(Sex, dr("Sex").ToString)
        If dr("Birthday").ToString <> "" Then Birthday.Text = Common.FormatDate(dr("Birthday"))

        Select Case Convert.ToString(dr("MaritalStatus"))
            Case "1", "2"
                Common.SetListItem(MaritalStatus, dr("MaritalStatus").ToString)
            Case Else
                Common.SetListItem(MaritalStatus, "3")
        End Select

        'Common.SetListItem(DegreeID, dr("DegreeID").ToString)
        Dim DegreeIDValTmp As String = ""
        '修正學歷代碼 (學員資料維護)
        DegreeIDValTmp = TIMS.Fix_DegreeValue(Convert.ToString(dr("DegreeID")))
        Common.SetListItem(DegreeID, DegreeIDValTmp)

        School.Text = dr("School").ToString
        Department.Text = dr("Department").ToString
        If Convert.ToString(dr("GraduateStatus")) <> "" Then Common.SetListItem(GraduateStatus, dr("GraduateStatus"))
        If Convert.ToString(dr("GraduateY")) <> "" Then Common.SetListItem(graduatey, Convert.ToString(dr("GraduateY")))

        SolTR.Style.Item("display") = cst_none1
        'MilitaryID.SelectedIndex=-1
        If Convert.ToString(dr("MilitaryID")) <> "" Then
            Common.SetListItem(MilitaryID, dr("MilitaryID"))
            If Convert.ToString(dr("MilitaryID")) = "04" Then SolTR.Style.Item("display") = cst_inline1
        End If

        ServiceID.Text = dr("ServiceID").ToString
        MilitaryAppointment.Text = dr("MilitaryAppointment").ToString
        MilitaryRank.Text = dr("MilitaryRank").ToString
        ServiceOrg.Text = dr("ServiceOrg").ToString
        If dr("SServiceDate").ToString <> "" Then SServiceDate.Text = Common.FormatDate(dr("SServiceDate"))
        If dr("FServiceDate").ToString <> "" Then FServiceDate.Text = Common.FormatDate(dr("FServiceDate"))

        'ZipCode4  服役單位地址
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode4")), objconn)
        ZipCode4.Value = Convert.ToString(dr("ZipCode4"))
        hidZipCode4_6W.Value = Convert.ToString(dr("ZipCode4_6W"))
        ZipCode4_B3.Value = TIMS.GetZIPCODEB3(hidZipCode4_6W.Value)
        ZipCode4_N.Value = Convert.ToString(dr("ZipCode4_N"))
        City4.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode4")), Convert.ToString(dr("ZipCode4_N")), objconn)
        City4.Text &= If(tZipLName <> "", "[" & tZipLName & "]", "")
        ServiceAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ServiceAddress")))
        Hid_JnZipCode4.Value = TIMS.GetZipCodeJn(ZipCode4.Value, ZipCode4_B3.Value, hidZipCode4_6W.Value, City4.Text, ServiceAddress.Text)

        PhoneD.Text = $"{dr("PhoneD")}"
        PhoneN.Text = $"{dr("PhoneN")}"
        CellPhone.Text = $"{dr("CellPhone")}" 'Convert.ToString(dr("CellPhone"))
        If CellPhone.Text <> "" Then CellPhone.Text = Trim(CellPhone.Text)

        Dim vMobil As String = TIMS.cst_NO
        If CellPhone.Text <> "" Then vMobil = TIMS.cst_YES
        Common.SetListItem(rblMobil, vMobil)

        ' ZipCode1  通訊地址
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode1")), objconn)
        ZipCode1.Value = Convert.ToString(dr("ZipCode1"))
        hidZipCode1_6W.Value = Convert.ToString(dr("ZipCode1_6W"))
        ZipCode1_B3.Value = TIMS.GetZIPCODEB3(hidZipCode1_6W.Value)
        ZipCode1_N.Value = Convert.ToString(dr("ZipCode1_N"))
        City1.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode1")), Convert.ToString(dr("ZipCode1_N")), objconn)
        City1.Text &= If(tZipLName <> "", "[" & tZipLName & "]", "")
        Address.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("Address")))
        Hid_JnZipCode1.Value = TIMS.GetZipCodeJn(ZipCode1.Value, ZipCode1_B3.Value, hidZipCode1_6W.Value, City1.Text, Address.Text)

        ' ZipCode2  戶籍地址
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode2")), objconn)
        ZipCode2.Value = Convert.ToString(dr("ZipCode2"))
        hidZipCode2_6W.Value = Convert.ToString(dr("ZipCode2_6W"))
        ZipCode2_B3.Value = TIMS.GetZIPCODEB3(hidZipCode2_6W.Value)
        ZipCode2_N.Value = Convert.ToString(dr("ZipCode2_N"))
        City2.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode2")), Convert.ToString(dr("ZipCode2_N")), objconn)
        City2.Text &= If(tZipLName <> "", "[" & tZipLName & "]", "")
        HouseholdAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("HouseholdAddress")))
        Hid_JnZipCode2.Value = TIMS.GetZipCodeJn(ZipCode2.Value, ZipCode2_B3.Value, hidZipCode2_6W.Value, City2.Text, HouseholdAddress.Text)

        Email.Text = dr("Email").ToString
        'Page.RegisterStartupScript("hard", "<script>hard();</script>")

        '身心障礙者
        Common.SetListItem(HandTypeID, Convert.ToString(dr("HandTypeID")))
        Common.SetListItem(HandLevelID, Convert.ToString(dr("HandLevelID")))
        Call TIMS.SetCblValue(HandTypeID2, Convert.ToString(dr("HandTypeID2")))
        Common.SetListItem(HandLevelID2, Convert.ToString(dr("HandLevelID2")))

        Dim flag_HandType As Integer = 0 '0:未選 1:舊制 2:新制
        If Convert.ToString(dr("HandTypeID")) <> "" AndAlso Convert.ToString(dr("HandLevelID")) <> "" Then flag_HandType = 1 '1:舊制
        If Convert.ToString(dr("HandTypeID2")) <> "" AndAlso Convert.ToString(dr("HandLevelID2")) <> "" Then flag_HandType = 2 '2:新制

        trHandTypeID2.Style("display") = cst_none1 '新制
        trHandTypeID.Style("display") = cst_none1 '舊制
        Select Case flag_HandType
            Case 1 '1:舊制
                trHandTypeID.Style("display") = cst_inline1 '舊制
                Common.SetListItem(rblHandType, "1")
            Case Else '0:未選 2:新制
                trHandTypeID2.Style("display") = cst_inline1 '新制
                Common.SetListItem(rblHandType, "2")
        End Select

        EmergencyContact.Text = dr("EmergencyContact").ToString
        EmergencyPhone.Text = dr("EmergencyPhone").ToString
        EmergencyRelation.Text = dr("EmergencyRelation").ToString

        ' ZipCode3  緊急通知人地址 
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode3")), objconn)
        ZipCode3.Value = Convert.ToString(dr("ZipCode3"))
        hidZipCode3_6W.Value = Convert.ToString(dr("ZipCode3_6W"))
        ZipCode3_B3.Value = TIMS.GetZIPCODEB3(hidZipCode3_6W.Value)
        ZipCode3_N.Value = Convert.ToString(dr("ZipCode3_N"))
        City3.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode3")), Convert.ToString(dr("ZipCode3_N")), objconn)
        City3.Text &= If(tZipLName <> "", "[" & tZipLName & "]", "")
        EmergencyAddress.Text = Convert.ToString(dr("EmergencyAddress"))
        Hid_JnZipCode3.Value = TIMS.GetZipCodeJn(ZipCode3.Value, ZipCode3_B3.Value, hidZipCode3_6W.Value, City3.Text, EmergencyAddress.Text)
        'CtID4.Value=TIMS.Get_Ctid(Convert.ToString(dr("ZipCode3")), objconn)

        '- 受訓前任職資料 
        If dr("PWType1").ToString <> "" Then Common.SetListItem(PriorWorkType1, dr("PWType1").ToString)
        PriorWorkOrg1.Text = dr("PriorWorkOrg1").ToString
        Title1.Text = dr("Title1").ToString
        Title2.Text = dr("Title2").ToString

        If dr("ActNo2").ToString <> "" Then ActNo2.Text = TIMS.ChangeIDNO(dr("ActNo2").ToString)

        'If TIMS.CheckIsECFA(Me, dr("ActNo2"), dr("FOfficeYM1"))=True Then Common.SetListItem(BudID, "97") '判斷是否為ECFA
        '該計畫是否使用ECFA
        If blnTPlanUseEcfa Then
            STDateHidden.Value = TIMS.ClearSQM(STDateHidden.Value)
            STDateHidden.Value = If(STDateHidden.Value <> "", STDateHidden.Value, TIMS.Cdate3(Now.ToString("yyyy/MM/dd")))
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                If DateDiff(DateInterval.Day, CDate(cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
                    'ECFA  005	勞動力發展署雲嘉南分署	2024 暫不限定ECFA
                    If Not flag_BudID_ECFA_NoLock Then
                        If TIMS.CheckIsECFA(Me, dr("ActNo2"), "", STDateHidden.Value, objconn) = True Then Common.SetListItem(BudID, "97") '2011/05/20 新增ECFA判斷
                    End If
                End If
            Else
                If TIMS.CheckIsECFA(Me, dr("ActNo2"), dr("FOfficeYM1"), "", objconn) = True Then Common.SetListItem(BudID, "97") '2011/05/20 新增ECFA判斷
            End If
        End If

        SOfficeYM1.Text = ""
        FOfficeYM1.Text = ""
        SOfficeYM2.Text = ""
        FOfficeYM2.Text = ""
        If dr("SOfficeYM1").ToString <> "" Then SOfficeYM1.Text = Format(dr("SOfficeYM1"), "yyyy/MM/dd")
        If dr("FOfficeYM1").ToString <> "" Then FOfficeYM1.Text = Format(dr("FOfficeYM1"), "yyyy/MM/dd")
        If dr("SOfficeYM2").ToString <> "" Then SOfficeYM2.Text = Format(dr("SOfficeYM2"), "yyyy/MM/dd")
        If dr("FOfficeYM2").ToString <> "" Then FOfficeYM2.Text = Format(dr("FOfficeYM2"), "yyyy/MM/dd")
        '- 受訓前任職資料 

        'PriorWorkPay.Text=dr("PriorWorkPay").ToString
        'RealJobless.Text=dr("RealJobless").ToString
        ''20100927 andy
        'RealJobless_msg.Text=""
        'RealJobless.Style.Add("background-color", "fffff")
        'If IsInt(Trim(Convert.ToString(dr("RealJobless")))) Then
        '    If chkJobless(Convert.ToString(dr("RealJobless")), Convert.ToString(dr("JoblessID")))=False Then
        '        RealJobless_msg.Text="*所填寫之受訓前失業週數與<br/>所選擇下拉式選單選項不一致!"
        '        RealJobless.Style.Add("background-color", "LightPink")
        '    End If
        'End If
        'Common.SetListItem(JoblessID, dr("JoblessID").ToString)

        Common.SetListItem(Traffic, dr("Traffic").ToString)
        Common.SetListItem(ShowDetail, dr("ShowDetail").ToString)
        'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
        Common.SetListItem(SupplyID, "0")
        If BudID.Items.Count = 1 Then BudID.Items(0).Selected = True '選項為1時選1
        'Common.SetListItem(IsAgree, dr("IsAgree").ToString) ' 2009/07/01 改成一律同意

        '就職狀況'0:失業 1:在職  (JobStateType)
        Call SUtl_AutoJobStateType(Me, JobStateType, Convert.ToString(dr("JobState")))
    End Sub
    '檢查(取得原資料)
    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        'Dim sql As String
        'Dim dr As DataRow
        'If Not gFlagEnv Then IDNO.Text="T122702064"        '(測試)
        'IDNO.Text=TIMS.ClearSQM(TIMS.ChangeIDNO(IDNO.Text))
        'If IDNO.Text="" Then
        '    Common.MessageBox(Me, "查無相關參訓個人資料!")
        '    Exit Sub
        'End If

        Dim dr As DataRow = Get_STUDINFOdr()
        If dr Is Nothing Then
            Common.MessageBox(Me, "查無相關參訓個人資料!")
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then Page.RegisterStartupScript("11111", "<script>ChangeMode(1);</script>") '企訓專用
            Exit Sub
        End If

        Call SHOW_STUDINFO(dr)

        '頁籤控制
        Dim flag_show_MenuTable As Boolean = False '頁籤控制 false:不顯示 true:顯示
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_MenuTable = True
        If TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_MenuTable = True
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_MenuTable = True
        If flag_show_MenuTable Then
            '企訓專用
            Page.RegisterStartupScript("11111", "<script>ChangeMode(1);</script>")
        End If
    End Sub
    '不儲存回上一頁
    Private Sub Button3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button3.Click
        If Not Session(vs_SearchStr) Is Nothing Then
            ViewState(vs_SearchStr) = Session(vs_SearchStr)
            Session(vs_SearchStr) = ViewState(vs_SearchStr)
            'Session(vs_SearchStr)=Nothing
        End If

        Dim RetrunUrl As String = ""
        If Convert.ToString(Session("RetrunUrl")) <> "" Then
            RetrunUrl = Convert.ToString(Session("RetrunUrl"))
            Session("RetrunUrl") = Nothing
        End If
        If Request("TICKET_NO") <> "" Then
            Call TIMS.Utl_Redirect(Me, objconn, "SD_03_002_3in1.aspx?ID=" & Request("ID"))
        Else
            If RetrunUrl <> "" Then
                Call TIMS.Utl_Redirect(Me, objconn, RetrunUrl & "?ID=" & Request("ID"))
            Else
                Call TIMS.Utl_Redirect(Me, objconn, "SD_03_002.aspx?ID=" & Request("ID"))
            End If
        End If
    End Sub
    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                '政府已補助經費   Start
                If drv("SumOfMoney").ToString <> "" Then
                    Dim SumOfMoney As Label = e.Item.FindControl("SumOfMoney")
                    SumOfMoney.Text = drv("SumOfMoney").ToString
                    If CInt(SumOfMoney.Text) >= CST_SumOfMoneyMax Then
                        SumOfMoney.ForeColor = Color.Red 'SumOfMoney.ForeColor.Red '超過4萬的提示，將字變為紅色的
                        TIMS.Tooltip(SumOfMoney, CST_ToolTipSumOfMoneyMaxMSG1, True)
                        SumOfMoney.Font.Bold = True
                    End If
                    'hide_SumOfMoney.Value &= CInt(SumOfMoney.Text)
                End If
                '政府已補助經費   End
        End Select
        'If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
        'End If
    End Sub
    '回上一頁
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'If ViewState("search") <> "" Then Session("search")=ViewState("search")
        If Session(vs_SearchStr) IsNot Nothing Then
            ViewState(vs_SearchStr) = Session(vs_SearchStr)
            Session(vs_SearchStr) = ViewState(vs_SearchStr) 'Session(vs_SearchStr)=Nothing
        End If

        Select Case TIMS.ClearSQM(Request("todo"))
            Case "1"
                Dim s_rqFUNID As String = TIMS.ClearSQM(Request("ID"))
                Dim s_act As String = TIMS.ClearSQM(Request("act"))
                If rqOCID = "" Then rqOCID = TIMS.ClearSQM(Request("OCID"))
                Dim s_redirect1 As String = String.Concat("../03/SD_03_002_classver.aspx?ID=", s_rqFUNID, "&todo=1", "&OCID=", rqOCID, If(s_act <> "", String.Concat("&act=", s_act), ""))
                Call TIMS.Utl_Redirect(Me, objconn, s_redirect1)
            Case "2"
                'If Not ViewState(vs_SearchStr) Is Nothing Then Session(vs_SearchStr)=ViewState(vs_SearchStr)
                Dim s_rqFUNID As String = TIMS.ClearSQM(Request("ID"))
                If rqOCID = "" Then rqOCID = TIMS.ClearSQM(Request("OCID"))
                Dim s_redirect2 As String = String.Concat("../03/SD_03_002.aspx?ID=", s_rqFUNID, "&todo=2", "&OCID=", rqOCID)
                Call TIMS.Utl_Redirect(Me, objconn, s_redirect2)
            Case Else
                Dim s_rqFUNID As String = TIMS.ClearSQM(Request("ID"))
                Dim s_redirect_else As String = String.Concat(TIMS.GetFunIDUrl(s_rqFUNID, 0, objconn), "?ID=", s_rqFUNID)
                Call TIMS.Utl_Redirect(Me, objconn, s_redirect_else)
        End Select
    End Sub
    '個資保護的顯示
    Private Sub Std_Data_Mask(ByVal dataMode As String)
        If dataMode = "1" Then
            txtShowIDNO.Text = TIMS.strMask(txtShowIDNO.Text, 1)
            'txtShowBirthday.Text=TIMS.strMask(txtShowBirthday.Text, 2)
            Address.Text = TIMS.strMask(Address.Text, 3)
            HouseholdAddress.Text = TIMS.strMask(HouseholdAddress.Text, 3)
            EmergencyAddress.Text = TIMS.strMask(EmergencyAddress.Text, 3)
        End If
    End Sub

    ''' <summary>異動學員基本資料。(STUD_STUDENTINFO)</summary>
    ''' <param name="oConn"></param>
    ''' <param name="vIDNO"></param>
    ''' <param name="ref_SID"></param>
    Sub UPDATE_STUDENTINFO(ByRef oConn As SqlConnection, ByVal vIDNO As String, ByRef ref_SID As String)
        'Dim SID As String="" '回傳SID
        'ByRef SID As String
        vIDNO = TIMS.ClearSQM(vIDNO)
        If vIDNO = "" Then Return

        Dim sql As String = ""
        Dim da1 As SqlDataAdapter = Nothing
        Dim MyTable1 As DataTable = Nothing 'STUD_STUDENTINFO
        Dim Mydr1 As DataRow = Nothing 'STUD_STUDENTINFO

        '2009/07/17 改成只判斷身分證字號 (多筆。)
        sql = String.Concat("SELECT * FROM STUD_STUDENTINFO WHERE IDNO='", vIDNO, "'")
        MyTable1 = DbAccess.GetDataTable(sql, da1, oConn)
        If MyTable1.Rows.Count = 0 Then
            'Common.MessageBox(Me, "此學員個人資料不存在，將要把您輸入的資料新增存入")
            ref_SID = String.Concat(TIMS.Get_DateNo, "01")
            Mydr1 = MyTable1.NewRow
            MyTable1.Rows.Add(Mydr1)
            Mydr1("SID") = ref_SID
            Call SUtl_AddUpdateRow1(Mydr1)
            gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(MyTable1)
            gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, Mydr1)
            '更新學員基本資料檔
            DbAccess.UpdateDataTable(MyTable1, da1)
        Else
            'Common.MessageBox(Me, "此學員個人資料已存在，將更新您所輸入的資料")
            'Dim j As Integer=MyTable1.Rows.Count - 1
            For p As Integer = 0 To (MyTable1.Rows.Count - 1)
                Mydr1 = MyTable1.Rows(p)
                ref_SID = Mydr1("SID") '隨意一組 SID '可能有多組。(合理情況下，應該只有1筆)
                Call SUtl_AddUpdateRow1(Mydr1)
                gstr_COLUMN_1 = TIMS.Get_DataTableCOLUMN2(MyTable1)
                gstr_ROWVAL_1 = TIMS.Get_DataRowValues(gstr_COLUMN_1, Mydr1)
                '更新學員基本資料檔
                DbAccess.UpdateDataTable(MyTable1, da1)
            Next
        End If
    End Sub

    ''' <summary>'異動學員基本資料。(STUD_STUDENTINFO)</summary>
    ''' <param name="Mydr1"></param>
    Sub SUtl_AddUpdateRow1(ByRef Mydr1 As DataRow)
        'Mydr1("IDNO")=TIMS.ChangeIDNO(IDNO.Text)
        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)
        Select Case v_PassPortNO
            Case "1", "2"
            Case Else
                v_PassPortNO = "2"
        End Select
        Dim v_ChinaOrNot As String = TIMS.GetListValue(ChinaOrNot)
        Nationality.Text = TIMS.ClearSQM(Nationality.Text)
        Dim v_PPNO As String = TIMS.GetListValue(PPNO)
        Select Case v_PassPortNO
            Case "1"
                v_ChinaOrNot = ""
                Nationality.Text = ""
                v_PPNO = ""
        End Select
        Dim v_Sex As String = TIMS.GetListValue(Sex)
        Dim v_MaritalStatus As String = TIMS.GetListValue(MaritalStatus)
        Select Case v_MaritalStatus
            Case "1", "2"
            Case Else
                v_MaritalStatus = ""
        End Select
        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
        Dim v_graduatey As String = TIMS.GetListValue(graduatey)
        Dim v_MilitaryID As String = TIMS.GetListValue(MilitaryID)
        Dim v_JobStateType As String = TIMS.GetListValue(JobStateType)
        'If Not TIMS.IsNumeric1(v_JobStateType) Then v_JobStateType="" '數字，改為INT

        Mydr1("IDNO") = IDNO.Text
        Mydr1("Name") = Name.Text
        Mydr1("RMPNAME") = If(RMPNAME.Text <> "", RMPNAME.Text, Convert.DBNull)
        Mydr1("EngName") = LName.Text & " " & FName.Text
        Mydr1("PassPortNO") = If(v_PassPortNO <> "", v_PassPortNO, "2")
        Mydr1("ChinaOrNot") = If(v_ChinaOrNot <> "", v_ChinaOrNot, Convert.DBNull)
        Mydr1("Nationality") = If(Nationality.Text <> "", Nationality.Text, Convert.DBNull)
        Mydr1("PPNO") = If(v_PPNO <> "", v_PPNO, Convert.DBNull)

        Mydr1("Sex") = If(v_Sex <> "", v_Sex, Convert.DBNull)
        Mydr1("Birthday") = TIMS.Cdate2(Birthday.Text)
        Mydr1("MaritalStatus") = If(v_MaritalStatus <> "", v_MaritalStatus, Convert.DBNull)
        Mydr1("DegreeID") = If(v_DegreeID <> "", v_DegreeID, Convert.DBNull)
        Mydr1("GraduateStatus") = TIMS.Get_GraduateStatusValue(v_GraduateStatus)
        Mydr1("GraduateY") = If(v_graduatey <> "", v_graduatey, Convert.DBNull)
        Mydr1("MilitaryID") = If(v_MilitaryID <> "", v_MilitaryID, Convert.DBNull) 'v_MilitaryID

        Mydr1("JoblessID") = Convert.DBNull
        Mydr1("RealJobless") = Convert.DBNull
        '就職狀況'0:失業 1:在職  (JobStateType)
        Mydr1("JobState") = If(v_JobStateType <> "", v_JobStateType, Convert.DBNull) 'JobStateType.SelectedValue

        Mydr1("IsAgree") = "Y"  '2009/07/01 改成一律同意
        Mydr1("ModifyAcct") = sm.UserInfo.UserID
        Mydr1("ModifyDate") = Now()
    End Sub

    ''' <summary>
    ''' 什麼時候可以修改預算別。'如果是分署(中心)承辦人，預算別不鎖定。by AMU 20140328 (本功能是每次執行)
    ''' true: nolock false: lock
    ''' </summary>
    ''' <param name="sLID"></param>
    ''' <param name="STDate"></param>
    ''' <param name="iBudFlag"></param>
    ''' <returns></returns>
    Function Chk_CanEditBudgetID(ByVal sLID As String, ByVal STDate As Date, ByRef iBudFlag As Integer) As Boolean
        Dim rst As Boolean = False
        iBudFlag = 0 'OUTPUT傳出值。
        'iBudFlag :0,1,2: 'iFlag :0:未開放 1:21天內修改 2:開放被登功能

        '14天後可修改預算別，不限定" & cst_ECFA & "助。
        'If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 14, CDate(STDateHidden.Value)), Today) >= 0 Then
        '    If sm.UserInfo.LID <= "1" Then rst=True '分署(中心)以上階層開放
        'End If

        '21天內可修改預算別，不限定" & cst_ECFA & "協助。
        If DateDiff(DateInterval.Day, Today, DateAdd(DateInterval.Day, cst_limitDay21st, STDate)) >= 0 Then
            If sLID <= "1" Then
                rst = True '分署(中心)以上階層開放 (NoLock)
                iBudFlag = 1
            End If
        End If

        '若已結束使用 但登入者為 署或分署承辦人，再CHECK一次可使用權限。dtArc2
        If Not rst AndAlso sm.UserInfo.LID <= "1" Then
            '未到結束時間可開放使用。
            If Not TIMS.ChkIsEndDate(rqOCID, TIMS.cst_FunID_學員資料維護, dtArc2) Then
                rst = True '分署(中心)以上階層開放 (NoLock)
                iBudFlag = 2
            End If
        End If
        Return rst
    End Function

    '將字串組合到 RelClassHourV 數字長度應該為 00~99
    Sub SUtl_CboRelClassHourV(ByRef RelClassHourV As String, ByVal iUnit1Hour As Integer)
        Dim Val1 As String = "" '組合字串
        Val1 = Right("0" & Convert.ToString(iUnit1Hour), 2)
        RelClassHourV &= Val1
    End Sub

    '受訓前任職清單 (CLASS_STUDENTSOFCLASS)
    Sub Create_PriorWorkOrg1(ByVal dr As DataRow)
        If Convert.ToString(dr("PWType1")) <> "" Then
            'CLASS_STUDENTSOFCLASS.PWType1
            Common.SetListItem(PriorWorkType1, Convert.ToString(dr("PWType1")))
        End If

        'STUD_SUBDATA.PriorWorkOrg1
        'PriorWorkOrg1.Text=Convert.ToString(dr("PriorWorkOrg1"))
        If Convert.ToString(dr("PWOrg1")) <> "" Then
            'CLASS_STUDENTSOFCLASS.PWOrg1
            PriorWorkOrg1.Text = Convert.ToString(dr("PWOrg1"))
        End If

        '先塞3?? cs.SOfficeYM1 SOfficeYM3
        '先塞3?? cs.FOfficeYM1 FOfficeYM3
        SOfficeYM1.Text = ""
        FOfficeYM1.Text = ""
        If Convert.ToString(dr("SOfficeYM3")) <> "" Then
            'CLASS_STUDENTSOFCLASS.SOfficeYM1
            SOfficeYM1.Text = Common.FormatDate(Convert.ToString(dr("SOfficeYM3")))
            'SOfficeYM1.Text=Format(dr("SOfficeYM3"), "yyyy/MM/dd")
        End If

        If Convert.ToString(dr("FOfficeYM3")) <> "" Then
            'CLASS_STUDENTSOFCLASS.FOfficeYM1
            FOfficeYM1.Text = Common.FormatDate(Convert.ToString(dr("FOfficeYM3")))
            'FOfficeYM1.Text=Format(dr("FOfficeYM3"), "yyyy/MM/dd")
        End If

        If Convert.ToString(dr("ActNo2")) <> "" Then
            '先讀CLASS_STUDENTSOFCLASS.ActNo 若讀不到讀STUD_STUDENTINFO.ActNo
            ActNo2.Text = TIMS.ChangeIDNO(dr("ActNo2").ToString)
        End If

        'STUD_SUBDATA.PriorWorkOrg2
        'STUD_SUBDATA.SOfficeYM2
        'STUD_SUBDATA.FOfficeYM2
        PriorWorkOrg2.Text = "" 'Convert.ToString(dr("PriorWorkOrg2"))
        'SOfficeYM2.Text=Convert.ToString(dr("SOfficeYM2"))
        SOfficeYM2.Text = ""
        FOfficeYM2.Text = ""
        PriorWorkOrg2.Text = Convert.ToString(dr("PriorWorkOrg2"))
        If Convert.ToString(dr("SOfficeYM2")) <> "" Then SOfficeYM2.Text = Common.FormatDate(Convert.ToString(dr("SOfficeYM2")))
        If Convert.ToString(dr("FOfficeYM2")) <> "" Then FOfficeYM2.Text = Common.FormatDate(Convert.ToString(dr("FOfficeYM2")))

        '--STUD_SUBDATA.PriorWorkPay
        'CLASS_STUDENTSOFCLASS.PriorWorkPay
        'CLASS_STUDENTSOFCLASS.Title1
        'CLASS_STUDENTSOFCLASS.Title2
        PriorWorkPay.Text = Convert.ToString(dr("PriorWorkPay"))
        Title1.Text = Convert.ToString(dr("Title1"))
        Title2.Text = Convert.ToString(dr("Title2"))
        'CLASS_STUDENTSOFCLASS.RealJobless
        'RealJobless.Text=Convert.ToString(dr("RealJobless"))
        'CLASS_STUDENTSOFCLASS.JoblessID
        'Common.SetListItem(JoblessID, dr("JoblessID").ToString)

        'STUD_STUDENTINFO.RealJobless
        'STUD_STUDENTINFO.JoblessID
        '20100927 andy
        'RealJobless_msg.Text=""
        'RealJobless.Style.Add("background-color", "fffff")
        'If IsInt(Trim(Convert.ToString(dr("RealJobless")))) Then
        '    If chkJobless(Convert.ToString(dr("RealJobless")), Convert.ToString(dr("JoblessID")))=False Then
        '        RealJobless_msg.Text="*所填寫之受訓前失業週數與<br/>所選擇下拉式選單選項不一致!"
        '        RealJobless.Style.Add("background-color", "LightPink")
        '    End If
        'End If
    End Sub

    '直接更動為在職者 (顯示時修正資料)
    Sub UPDATE_WorkSuppIdent(ByVal SOCID As String, ByVal OCID As String, ByVal flag_WSI As Boolean, ByVal tConn As SqlConnection)
        Dim v_WorkSuppIdent As String = TIMS.cst_NO
        If flag_WSI Then v_WorkSuppIdent = TIMS.cst_YES
        Dim u_Parms As New Hashtable From {{"WorkSuppIdent", v_WorkSuppIdent}, {"SOCID", SOCID}, {"OCID", OCID}}
        Dim Usql As String = ""
        Usql &= " UPDATE CLASS_STUDENTSOFCLASS" & vbCrLf
        Usql &= " SET WorkSuppIdent=@WorkSuppIdent" & vbCrLf
        Usql &= " WHERE SOCID=@SOCID AND OCID=@OCID" & vbCrLf
        DbAccess.ExecuteNonQuery(Usql, tConn, u_Parms)  'edit，by:20181017
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        '<asp:Button ID="Button6F" runat="server" Text="正面" CssClass="asp_button_M" CommandName="SF1" />
        '<asp:Button ID="Button7B" runat="server" Text="反面" CssClass="asp_button_M" CommandName="SB2" />
        '<asp:Button ID="Button8PB" runat="server" Text="存摺" CssClass="asp_button_M" CommandName="SPB"/>
        If e Is Nothing OrElse Convert.ToString(e.CommandArgument) = "" Then Return
        Dim StrCmdArg1 As String = e.CommandArgument

        Hid_OCID.Value = TIMS.ClearSQM(Hid_OCID.Value)
        Dim ETYPE As String = TIMS.GetMyValue(StrCmdArg1, "ETYPE")
        Dim EMID1 As String = TIMS.GetMyValue(StrCmdArg1, "EMID1")
        Dim FILENAME As String = TIMS.GetMyValue(StrCmdArg1, "FILENAME")
        Dim v_SOCID As String = TIMS.GetListValue(SOCID)
        Dim FuncID As String = TIMS.Get_MRqID(Me)
        Dim Url_1 As String = String.Concat("SD_03_002_IMG?ID=", FuncID, "&ECMD=", e.CommandName, "&ETYPE=", ETYPE, "&EMID1=", EMID1, "&FILENAME=", FILENAME, "&OCID=", Hid_OCID.Value, "&SOCID=", v_SOCID)
        Select Case e.CommandName
            Case "SF1"
                Call TIMS.Utl_Redirect(Me, objconn, Url_1)
            Case "SB2"
                Call TIMS.Utl_Redirect(Me, objconn, Url_1)
            Case "SPB"
                Call TIMS.Utl_Redirect(Me, objconn, Url_1)
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Button6F As Button = e.Item.FindControl("Button6F") ' 正面
                Dim Button7B As Button = e.Item.FindControl("Button7B") ' 反面
                Dim Button8PB As Button = e.Item.FindControl("Button8PB") ' 存摺
                Button6F.Visible = False
                Button7B.Visible = False
                Button8PB.Visible = False

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Select Case Convert.ToString(drv("ETYPE"))
                    Case "1" '存摺
                        If Convert.ToString(drv("ISUSE")) = "Y" Then
                            Button8PB.Visible = True
                            Button8PB.CommandName = "SPB"
                            Dim StrCmdArg1 As String = ""
                            'TIMS.SetMyValue(StrCmdArg1, "ECMD", "SPB")
                            TIMS.SetMyValue(StrCmdArg1, "ETYPE", drv("ETYPE"))
                            TIMS.SetMyValue(StrCmdArg1, "EMID1", drv("EMID1"))
                            TIMS.SetMyValue(StrCmdArg1, "FILENAME", drv("FILENAME1W"))
                            Button8PB.CommandArgument = StrCmdArg1
                        End If
                    Case "2" '身分證
                        If Convert.ToString(drv("ISUSE")) = "Y" Then
                            Button6F.Visible = True
                            Button6F.CommandName = "SF1"
                            Button7B.Visible = True
                            Button7B.CommandName = "SB2"

                            Dim StrCmdArg1 As String = ""
                            'TIMS.SetMyValue(StrCmdArg1, "ECMD", "SF1")
                            TIMS.SetMyValue(StrCmdArg1, "ETYPE", drv("ETYPE"))
                            TIMS.SetMyValue(StrCmdArg1, "EMID1", drv("EMID1"))
                            TIMS.SetMyValue(StrCmdArg1, "FILENAME", drv("FILENAME1W"))
                            Button6F.CommandArgument = StrCmdArg1

                            Dim StrCmdArg2 As String = ""
                            'TIMS.SetMyValue(StrCmdArg2, "ECMD", "SB2")
                            TIMS.SetMyValue(StrCmdArg2, "ETYPE", drv("ETYPE"))
                            TIMS.SetMyValue(StrCmdArg2, "EMID1", drv("EMID1"))
                            TIMS.SetMyValue(StrCmdArg2, "FILENAME", drv("FILENAME2W"))
                            Button7B.CommandArgument = StrCmdArg2
                        End If
                End Select

        End Select
    End Sub

    ''' <summary>
    ''' 114年確定性需求8：網站+系統 建置AI影像辯識功能
    ''' </summary>
    ''' <param name="IDNOStr"></param>
    Private Sub SHOW_IMG12(IDNOStr As String)
        If Not TIMS.GFG_BuildinAIimage Then Return

        PassPtTableMsg1.Text = "查無資料!!"
        Panel_PassPt.Visible = False
        DataGrid2.Visible = False

        If IDNOStr = "" Then Return

        Dim PMS_S1 As New Hashtable From {{"IDNO", IDNOStr}}
        Dim SQL_S1 As String = ""
        SQL_S1 &= " SELECT E.ETYPE,E.EMID1,E.IDNO,E.CREATEDATE,E.FILENAME1,E.FILENAME1W,E.SRCFILENAME1,E.FILEPATH1" & vbCrLf
        SQL_S1 &= " ,E.FILENAME2,E.FILENAME2W,E.SRCFILENAME2,E.FILEPATH2" & vbCrLf
        SQL_S1 &= " ,E.ISUSE,E.ISDEL,E.MODIFYACCT,E.MODIFYDATE,E.CATEGORY1,E.ACTION1" & vbCrLf
        SQL_S1 &= " FROM V_EIMG12 E" & vbCrLf
        SQL_S1 &= " WHERE E.MODIFYDATE IS NOT NULL AND LEN(E.FILENAME1)>0 AND E.IDNO=@IDNO" & vbCrLf
        SQL_S1 &= " ORDER BY E.MODIFYDATE DESC" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(SQL_S1, objconn, PMS_S1)

        If TIMS.dtNODATA(dt) Then Return

        PassPtTableMsg1.Text = ""
        Panel_PassPt.Visible = True
        DataGrid2.Visible = True

        DataGrid2.DataSource = dt
        DataGrid2.DataBind()
    End Sub

    'INTO CLASS_STUDENTSOFCLASS_SOCID_SE,CLASS_STUDENTSOFCLASS,SOCID </summary>
    'E05290DA76997FBCE4589CE5CFA48542C8BD080CF8C9A9843FD168D64CE6EE002F8186001BE890E4E4CE25F917F96E9E4DAD502AE7FDE2CC7F820ED9
    '03809089FED49FAE54BCD0A8F6A9D862AAB2063EF896F4429CFB741BE07D278330A9B8AC434BE6F562F3021B7722F2FE005AFB6368A196F63746F50E
    '0A8341510D85064FFC9676BACC654165E9966702FF7DBAC46C6536936949E5D67B1628D1E111F9199CC7890EE8E964B32395B3324D15D796D877F997
    '369DDC1CB85F3B426BF9D4DFEF644833ABC0EF6E49B10B805A89726F4454A9F8DFC523D2265415B01093678B822D94F9768A9748332F4425DB7B162E
    '29AF757E942B6F6B6FE428210EFC9FD1FC6887CB28FDE173C683D4697D83EE43D468CC7237DB30388DCBC3D65B243C8D008EA987C2C364C1C15E11C6
    '091218AE405D9FC396293458CDDC1987FDCFEABD587D95A40837006C91458018938CF55EBEDFB9E5629BCBC94AC6FE1FC5B91285F3BFBBE5D53CF06D
    '4C03DCE64F675150E2BBD42685A29320DCA9F9203CB0FE5A62356D16EBA1397E96B149AA9A701077F3730E88F22450C83699D7542A93CC566D9CF86A
    'D4EACD9D0129C110F924BDF432E182A5547E8CC09F00AAA12CE65802309B52C98A5F8008D9DE33F8B16834484CD5048DB85A3A77C54842B1B2165D99
    'AC6C0E62C0B7DF4557280B705F6502F66328B55048827990FC84DB52E90B278C52A3545616764361D720B26E06FA39978C0A1B7EC6F941632A5219E6
    'F46C160A9618E819B80AC201F77956B01C7AF52C7D15F96561CBF017B2E85779C70E76095DC7412CF035DD7D0576772EEF88E56E88776C644E8F784E
    '9E3D7C4338C6F388CE41B350835E1A6321C0534C1167FA332E99C94152A931343FDA18FCFB64AB3F8353544C05F7320B662D367F3B9B9A479B0EDD52
    '960BB4D7D188AF17074383368C424DE402EF04B46D4AD5F92834A58704915EA6AD9A03B91516770664AE81052561B610483C117C856DF2C69504C5F9
    '507C68939823625868EFE3B168B45A476E2A3D526E23111022CF1A8257F99535216A007618E7FDEE676C3784620B32A08A9C5CE9C972CA030D2E3775
    '46F5EE3A16F7AD3154BE2C3BE24829ED0B39C06C3E02D50A615A81598D01867B012AB42AC863A229D0F6B6FA7DE588000E9B4ED72C228BD5C57ABD70
    '18874AED776F57F4A03D2DF99F3E1B34B9C22114597786FA61BCDFF4B2A548CE87427B446949852058EBD8A08DC2D9DD76042607E3D472D706FABC9D
    '9520DBB8151A3C39EF17E630D3EC5283A741507C27816BB781F1DAAAB1395A6372DD64B93D3FADEC48FF5C8734E673A00516E322E0E08F62D6D25F8D
    '342555BE43A69C5E1EC8DC529CF9969CB48DDA1CF45CB72A8AB77384A5C1355C876B710655D9F59BB14EDD692E6B3189558B03D18A4523F8E65151C0
    '819198F83B22D06BF80968B636582450D67E42F51BCC70509CC5BD73BD023C48641D6EA50335709B61DA4A165BE6957151D57526169C2F1013DA551B
    '53CC9F74112CF13C5F31889B7772AE11D94E6D9AB0F5E37B521D1C1FB68C1B66312CC4072B44E8F9680E37F9F0DC06E140AE4EAA0E30CB3E9907E25E
    '0319D451E56033CE500DA9C7A41683B14BB712AF8B55AD14EFF54AAB9420AAE507E81AF7096EBBCC76A66ACF1B35ACBEDE5296D01A43C293BED4A5A2
    '22109F179C057D9DACA9EB6DFCD5896980CD01E51FC62AFF8D7557D1D9454287753EED9F38BCF5B028799F34E3C3643F636BFB15FE31C518162FA028
    'DFF85E23694E0579D2D2A162C17E327ABE9A188B9B5D17E032E9CF5EDDF04A7F36A60A83F30FBB997524CB0EEF516F09D435B2AB1B083A1D604F323A
    '8672962D6F437D5953163B1F055CAE47657E94991522FD67B802D3EAC78E73591EA6B2A44DD01E1D27EAF847744C7CABB9B61BB50F3C69A99F837F18
    '826471341A28BA0EEE6E1ECCD886F92361D212A650B097103F5B49D93752D3880F099B8F82A279C6C160AF090E35225AB5EA1B40CC453A92C96367D4
    '51EC22D53049EE55C777BD617819AA72EBBCBB11ECCE621A4E5CD0DFE25089125481AACD3854FCC45E701B452E1725E89689C3D5278E1ACD5B19AB1D
    'C9152A7EC7EF2C52AB5919E94ACC053EE8282BE8A00A84A21A034C8DE7D0759D8438D6436F3BD34ADF002DDDFDECC1E457DD08A9F1E712A66FF02400
    '3544CACAF05176FD089AC82ACB9F02FC075190348CBE57009ADA6B24952FD9AA237F29103C624DF754A18D658BD3A728F4C18CAAE408BEFB52D4ED73
    'A336752CC9D940AC8C2227FA1A64CBAAFE1FFACA6813F73D8C3FC5A35CCAF222228EEB3465FA1A0CC155A0D346A292E4CA231E1BA607B455A0EC5E28
    '26D5CCD184549665416AC68C12379634ECD38C21B0E913BED6CB25DC76E32CAFECA46FF06B63D290832C17F0F03BB1C4057DB284B870E38004AA23CA
    '400CBE65C416C8E3A5162CC0FC8C7AD3B446E5899D2F75B2884C402680924DD0B6BA79B8D46F7A2438D6E1AF6A9F9BFFD9255993ABE620B2FFE51E27
    'E287E3536DDAB9F0C00EF5947446C978CF5C363F33DC1122D0D7B02DE69CB554E6D6ED71DED4FE9A520EDCFE3B43B0C5B1B9CF4E881F632A019E36BD
    'E9D810E9580195F95EE152C0DC3085C718706C0BF2BCEDDD78F167DC94B3384376BA9FBE480F3DEBE8554378B04A6B5D5618D35AD1CBEE1D603181CE
    '85ECAEB97FDC28EACD8C6B27FE5B2C389053FA52526ABD38335646EDE6C2BD39537A6199C843DB5FCE8E90DF882CEE3FE4472063AE148FC45E1C8D51
    '821B15A10BFD1227F82F62BF80C6974AD5CD7902A257F9396DB5E1305BF2ACCB367C172B8C2D358C31538CBAE41B1AEF2EAE734F33BF57904C0B4BD8
    'D966296CAB8B39229CE9F27480578E7C670079A088E9BC8343D5D10A09FCCEEBE239428D4BF5755C65CED9DC869475732AD4262F712F2739A6CFB67F
    '97B500A29BF6814AC188704ECBE9A35F0FDE5337C7FD3A410F0F318E3E7F93D83A296E3F827294C43BDD57BFAAEDCE085695F668B40A6DB541E67AF0
    '48C8495F28768F8AA13CDEE2D14B03F12018D486A67738DF61C409705B30A2856099258F466F52FFAE859A50AE71A22AEFC8B74C6564EACC19A4D34E
    '0017A2B01876C892617E7364E9B3C012B2EFBD9B2888D9BA01051ADE04F57880539BD2554E9E38C28E96C70AAB1C85CE6B391777BA4FEC85CA3E996D
    '305DCE50CDBF50EC4440CBA7D9EDD2E610CF29D4DD7B218BDCC187E8FFDECF80393D4893F9E4EAAC33CE044209F854632B280AD5948B48F3349A73E3
    '8823C5B2F83347E42A34BC9DB2AD52A7C2E303712759A82397DCE4E8DA53DAEFE0DF5CD3C52D20FE18078D727BFE0D4F45A7DC784845A4C2D2E9DB9F
    '370CD75A0AF8D6516980AAC9DFC9639CF95F83A1080E807D1CF8494BB6BC0A091DDF128AB457782983F187FC205C61CEB9A2F9A9272C6569579CAFAE
    'FBDC3932998565AE2DEA52B0DECFA4D6B417AB82FD30FB692418583730F2431212E41A8E5AFC39DC7CCF1E7937906BD364F318D09CC6843047E9649F
    'A5F101332AEAD46E927136041D75EA09602E537AC8A52FEE8D6698091B86C0AB5D2472FE843689919621A2C1C35B4B059C5D9F75B552A1E70BBE2274
    '88CF25EFD4B6AEB1B997336ECC476ABB7F0909B34922F499D8F06674382C95AE1ABF8FB57D08F058B9F86ABFC3E6CA6408B105B2548C1354C389B129
    'DD51700836C6FDBC420C18487935FDFFC1098AE5037DBB7EC84A543AF06F29ADB443C19B19E452F01F40C263088E6D42534395EAE4196191D1C01866
    'EF5D08E4C80C03B7E5FE5EE7A4729273D269C58B8AEF27D929CC25E67A3CBA8754CADE64FDCA99C3A25BBCD79D189299B3E22AD4C965B9590FB61F80
    '02A72E5F4F50A01ECF85762B60F143062430199F6E034813701C29A302FE0A756A40285ED3133C75D8C8905CF79D7A6BFB279E7449D5BFE7DF7BE1DF
    '9CF7AE82CEB9316FC7FFD696CF4C76A097CAB5A163571D8802856AB7F449D88154A3F7F5D86A764D24B64803DD7FC88536EFEA60E02FAFEC6C39FE87
    'AFE37F9DC3B942A70E1CF90A973306472C462B1F0801EB6D3A0A303A27E5040B129C43A16F5A44E46ADDB70DBDC4BA55D24F777C5D5D412CE7CF1BC1
    'E309CCE61779CE3F13A3CDA54BFB7ABBC8A396566207D200AFE86364F83DC12AB926DD5F8D8B5E56C23ADB74D5FD145A8D37E1CEF734D61B87CAA95F
    'FB5934A8A408A3B219A8F4A30588F88BC831C100D58EF9433DD7955083B91CD3A493C6AD67895527D68E734F69CD0D7896968EC3988C77801C1781AC
    '4C20B41C02DE8167BC5C17D0E727A6B5842F676FB3FE8A2121344744D5BD5184A7F27AF3F85D804974CB88727948E387B8C598D71CF6C7B32C9EDF29
    'D4939E0B2CCDAF2969CCEA4030C36128F573AAFAC561719D392116CF4B59F979317BCA60A3F9EC0E0B425A40A27A429059270553327BE48BFC404557
    '59C602A4D54F7802535B90A50005F60D42259D5AD874D5C27478E0500F7C62F706A169FC070E7013CEEE4A4D5BBC26A19EDC1E4E3F690D8DEC31B1E2
    '1E3E1A90C1D3EC1E9DC7ED684087E6391BB5C93C9D354E19E77FEF7D08E55AD82BC0BAB4F4CFB47B3066FF4831C4335E6FB374F921E841B7F998B868
    '7B605AE9CF7ED84662BCC474D88964D4262A6A1F1DFF57CF1BE2F43922508CCDBD92A914C5C69DB696674A091D87EA711B74F41F4C629A7C3B83A003
    '44D397A6E21A2990C14631FA7FE49D0BC19EFDF20549A758CE58D93BEFDEC4376E5A040D42364B22BC7170268034F0DE7496561013F4ACCECC200D0F
    '55EC90DE6143AEC8F7A58F0207782989EB2FF3D29E2CA0A6B4E970D6FAE96185801FCCA11D539607DFFC515D7D5F55213EE80CA1A9E0BEA716FEF276
    '104A2262396B834908AEB9A4FBE173FB77B262E4CB4EE59A75FBC71A01B4D26D15F8D82FFB596AA07F942314E0C43971F8331B53E311DD90A7526AA4
    '51D686384ED4C3B1939FA08EE30431ABDD44FF66BD78CE18048CFAA749E8182AD08B764ABAF78404A525383D7E520AEAED3BBAC10A2004AA057F2BD5
    '8083728F70CAFD11709474CF5D1111841730B02F494E8B3C86C2094B2B54178C9EDFD82DC89160537FEBE54139C842AEA7EB27D0FD16CA43C10C8A09
    '62D410226CCD2263566481C1D8C5D750B8225776F28289305A787D903ED224A6334746A090FA1EE8AC087C9C099288D36A4D060C69B2C9A54D8C25E9
    '6DC0699D379AA1ECA8ECD9E6F6EAC00D650D12B0CFF1552EB01B0194FC7944F8986A09F716543441E90D7A531BE458F2A15CCECAC271BAF6C14B6128
    '9F91593A487B571A6FC9D56C98E7D43FC5EA654D9B33A2CB8686319E237874D8E9BFE6984A3204104167999F043233770CE665800A008C1B5174824A
    '2AE54C9247633D2D9B362D227AFB9461CFFD77B14FEB8917F43C697B9C1D3032F4A2CD055727ED525B5E204A4598B70718B2478BB863E8A2749634E4
    '39FC7A05A42D2D1D732420D90E2986D3213728F351B75409B7648165EA094F6F2BEBF7E03DD6BB36CF3199DF398EDF4304567CC697813FBD98AF1E5B
    'F29F86EC48DB4421D9082D3ABAE38D54995A40A695F507BA7E293FBD5245D4FDDB25F07F14C5F1BC1E3692597D47B4096D84F322EE4E25DCB86CD3E0
    '0B7D16C65E3162B45A0047EB82D80094155024805100D734E4D39A40F642F9F6CE9D93D8A696F203F7EADB74AB93ABBBB8F498439A0E024C41EA009B
    'CA940D930F8B142BF540255F01700DE4099C253253FB356198A94BC4003EB930179A7DB0E367707E97670EC8BCC5C89C24E1EBBE89DE650B416CC01E
    '1830B4997A1FF9971A778AEFFBA7B3BCBAAA9A1B8591BEFC8243B617DADBB5A630FDAAB096A1DE1983BE7FE221926300EDE3CD5A3DB28C842E7CB27A
    '2F9CA7F190155B6515C1507F5E072D14F08762964E83721851C4D0EA793109811B08C681FBF7B6A8626D20BCCAEA6C8055329FE0BE433390AD90A102
    'D09A518A3E9F2BEE95DCFE545E406FB7DF3C636B1EA9BA537ED467AB45A22125B75A9EEF09C9683980B743C4D2041A60E98483A04B6335100C8A267B
    '06AFC1B70F3107D085418A9E05A4A9878FFC3F4EE468DD8DB11EC4169EEA2F226C00BED244ED1E6DE9955CDF7837F8BEF256E1F019A52BF085A685E1
    'C36791ABFF0D59737E3B515780C5A05D5C41F8DF179055C61C17B682E2DC41A8D017D4B4A867756ABEF5C37D2538FF4ECCFC12BB8206716205A935A4
    '6528FFCEAD96BC45DDFDE53F3F7E09E1802BF7739A92FDF99EC4FF193EF2D4D139180316F56B1418E0195DC8559ACDB10761C671333ED7A731EA68B5
    '57CDC2CFABCD1FBA4217F1ED453E706AC520F4BF57AFF978FE58AF91DD80C3B95D2B6C1880639B6B84DB4ABC711C84E2CE0E4F007FD3DD400F99A726
    '3A49C525B51E412BDC4D01E6DC42CF929623DD830E2E2C66C2739CB65F679F89FD484826C3C71B779E7C06989DB1D503B641E2CEFDAD24A1308B987C
    'DA9046CE595909D2579C4FF0E90DCBC6D632B3DBA162DAF612FA68FF4A499F04970864001F88828B6ADA7644CCC82E4839F68B3D5AAF45527C090C99
    '1D96616A6B0AC01ECAFE7A6EB5D589998EAD4323975EF4559981EB2EAD08688B0FF382CC8947F081E61B9418AF34FF1FAE294B7490314B933B331CA9
    '260D5BF8D283CC18CF3C3C5F47FF63DB5A6283924EFDA5738522D595B9C67F83DF8D5FDA8EE0A88965B870B39A541D186EF7196220F6B4FD59E82271
    '9143BA9C69CB12BB6225D3B9BF929855F3A07C1FEFE55946216CFD83E3D7A091519F13770D4EEA58E37D0FC790E6F3ACB10941F275CC6B4062E89CA7
    'B62CCD8A750733CC7D475AC9B18E5D247EC4ADB273E0E7582A47E6ADF9B549CF7D7307AE2FAA61B8153B9157F7A12EA3ACD654AEE433A78DEFCF5727
    'EA2E4CBCDD1B8A31319F6C667F92544478B9A892B7E5E234570A536E761CFFF2C6DCD720F9D6808198DFDEA2BF3A0D00FC37A39A88274DE4A3DFD897
    '835319694F3A77594100293F3653DDB210B6862AEEF74C740418A3239ECEA3129C48D9CD24A47BCBE550739E85A581DE8B3DA0897CC9604255ABCC30
    'A727C556B6657E9B7981A174AACF97699D1DC5DB72850E0E2B3043B4D5AE65126AB9BC098479DA6DFC69708F6FD022255C6FB5E9A00B3DBED63E2EF2
    '69AE7EF33DE2747E61C5091938812EDBD75E96B4547C7742F27F92050336516036D1EBC89A92B532105B0C30081A52BA6D19F818644408F01306B6A2
    'CA4CB15458842EC61C8CF61DA3DA901265616B815A338EC4BE61A49D4B9CB2B9B4E84BA3170ED217BED2993D38AC4B3A9990045604F9D63065F0E3BC
    'F406E7055EBC4E53BEA8C95225B6CA7D27A513B44B36A02C99C0C4DE3EF223AED455CF76AAD66FFBAE57890D88A326BD4E1A509F476F0E61A073F28A
    'E3175C67318A00EED6D8AA69B3AA173F54098B69AACFF4DF5614F18958BDEFB1F178008B0F88BC3A651D4A7BBD8CC35BC663931566E7711ED193D5B2
    '5DA3542846B4A4C41769201A10E11DF7DCF3B6D75CE5F0FB7163CCF4E136A3C52477CD95EDFCBE8741CC017732ED9B7E70B78A4C6A0F321F87AF8A0E
    '9ECF44725D77C66F812E44D6F258ED862D69A7685DBAD8869D202191AC9DAC988CFC45E8DA78CC167FAD9FAF13753E05AFE3DFA29CACFC96DF1807BC
    'A78AE17EE5662F73C95F96D5DA2E7E0103CB3CA4E932FEC6E19D7C15EAC5BC959DD34CF7B1AAB4D93EF6ACB3D7B68A11CB1CD3A242D36B4C108772B1
    '4F6EF6C782BAE53FD51E203380A07CE6C8142C18F47F0C6918AC33D060A6ED1AAB165A4CF268E901D9253747F01DE3D64D8C01E4EE1BA6AEF36E9E1C
    '4BDC7C48686E41622B08F69427E4F83770FBDA1F0AE811AF44033353B4A517970FEEAB98C3670832FE5160BCC1260015BFA6967B9BC5321AE95CE40D
    '2BA78C4F4828EB39B70834FA3174E6CB0FB73F281F90E17710E1F87CA1EA75710560D6762D5026E2AADE99DFEAAFE6F30392CBD33A69BCE2D4823FDE
    '0B6613448A201158D55AB1D8A9C60BD255D9639C38FF945C1631041B679D1229FFDE58C22D3EECCEEA3954615B39616C45500B003AB3C05DB8B807BB
    '089DFEFA5C98B24E6FD9EAE162F5E17D093C42F6CA4C9709D85DAD0E822099ACD4F15EE50B466108E1E8EE81226288A7E1BBE44E2A1894397C321AEF
    'A4FBA9110AC3CF3E98752AD5734C43246838A75AC027DB1AD3CF62B5C2DA7F056AB554D3DF0B16CBAC6296E63A26671912FD173104C0E26524E62292
    '194A1AF4238D67BC456F7574653F7726E10B40FB1C3887293AFE3FAE17ABE83DF95879C001DAC43D8EBB6EB12DFF384B61281E38988BC18F572FB129
    '0673A2E7DBDBE16BC412C56B25F59457F01E5175A9404C02D09D19082523AB7D22305E744DE935DC38FB10124CA25C5D64C7E88F1E4C5C8241F10821
    '03E9A656D72810F7090E4056CECAB5D10FB3C2D111AAE33FF6C61A1B458F74701A53A2E85FDEF48B66624B0003519640BC8ECD72B7B4533FEBB9C439
    '662228E022CA6E6B3C6E9DFCF1867B7C9A1D66F70C8DA8E8493FA02D367BAEA4A6C957A65136FA75FA703CE1E5D0A827FFA5D2857539CF69A74CBCF9
    '9D8130228C658EE9644E9CAD18FA7B7E6962040A7CBD2A9C2214024E026EBA7AA0ED9E521BFFE598C1E1BB1057AA5564134788FC27BA2C60CA576E58
    '687D999815E5B60D8642CC434EF9DF684D7FE89207EE989F4C93214C2DF04363292A6116B412E7B0DF1876024637EE6841616AEA5475D712F9DCA268
    'B88BCDAF8DBBBF38CC39ECF1569FCB6178A3EF4EB1491F25D0C84D145A71AB37021E1942D0712120CA0473F1569CB1C1462AEB42CF5E554529796742
    'F82DCBAFCC89BDE0F0FB61E1F311B1756378C41CB538E55C22BD99E20E3E097310805F2C6D7347AE10D3574343370C60DA1C0F2FC60883DA32A3C894
    '5A693E8893C58D0E86B2AAA40698FA558D1A53AA1C9EF40072CFBAA9B73B299AAAC6B2B3CF42C6FC37B0F1E48BBC447188009C051AB22A0130CBFCF6
    'C99D1141BAA1AC046C73DEFF10967D09A317FAFB0C2B99CD669AACCA3E87CEFEEA38405E6B1FF70AB305D38F3C20096C3E7F5AAE2D7D31AE8D401BE8
    '9C20F64348A187EBFA7DECC35965ECCEBC19360DE87861D14CD138EE9F4849B66A5C8A19E8AC2EC472E0F2DAAB5789E6F5D69565D097B080982D59D3
    '775D22932675A9F4527078CEB8B93A68591BE54911E446AB67E991A1315DBCF8277E3E1A1EF59919A1C91C9FF6355C5329A33A6BA75A77878954ACCA
    'F7199FFEF6B5B0AB49BC7EB2C3990D73191667F50030A41912E711712A81B7142EC0EDB53EC5B74B8F62B782E82CA562DE602F7F9DFE98A34A78C42A
    'AFC32AF5B26901CD8A75CC48A36A7A4275602BAB5AAA56AEAAC2477AC3838014A930BBAE11533A2DFF96F10151D0B8711A182AF878F3DD03FC0D59A0
    '655EF7571643DB455F13C26EE549877660AC8D5C4419255EE5361B9F81A7DB1F4823CDB231643AF7F457093FDA7C39C622D39ECF7E1D054FC3C2626C
    'A5825FBFBB8704428B73E3867B1C9CBDAE7B0935A1848267CCE857AE3C6BA1A81F5BC6DBD2C6DF5F7D499802E1306AE7046AED9670A6A1F8800199C7
    'E73BFBB3890A4105A3747F15E53E6C3F936A00C492E6576CACA8B9222D1175D1D36864302C8F1B0C02874476C5456E7AA0098BD4F80FDD1166BA5E2F
    '6A0438B1485A5892A69F26E00C1DCFCC09F525928DCD399A25924BD531AF3619F8D1D9843592B205722EC4E2579506363CD4CA09C09AC81E5FEFC4AC
    '4C29623F711DE67C3E1E9DFC220581D9B7A578696D19D17DDFF155CD3ED853E0F3FFD856822329B5D4039EF69691C15621C04F2DEE0AABBE9F193288
    'CDC8902A5F1FDF88E3BAAF64CA205F83DB82A8B64FE76D7A968C44068C873012FADB7AF2B0233BBF55D453EF5B077A826CD24DA30FAE531E609EAB36
    '211D683A61CD4C73C5C6D1160D322B7C23C3522EE45445B7995A86F2D6DF25CD643E937715CA6CEFAC535F7B0A40D6EFED62214ED575A99DF1C836F9
    '9EBB2A9470125E11908D82E2BA1EF5E878DEDE4856AC246B0A6D69ED18B612F6822E1CC1411AA2992912A38BD765F77ECE705235DC711266A703B83A
    '05FCF32DCA7E620E3455FECD33D4AC6C9F7B546DEA4F14D8559691B2079118447ADFEDA4F4AC07FC099F1F0073FF6ACECB1B867EB0064CF47774D978
    '226FEA26136CF2C9931CE294D501725658D6A776BD202BD35E8898653FC2B682B2AD8D4DCEF7C0BBCE025B7C204A1AA000D088DD439653D985E06856
    '4935DAE9C4D0A087B81BB9321FB42FC96F46F6D4371A6DE545D51E6F640C0D257A222569DBE89F4FCF1FECBA07C21A9CB9E2D544FCDD1EC4623E948F
    'CE166FA47FCFECED75EDEB187BB5338DF62807D4BD6FD514C06547E4F6AA5A0E3F7A957AD8563A6DB8B658D5407B842C1181CAB832D5CE4F8E0BF760
    'E28B84D213489547A123BA6EB958FD8F96A122CA4A0DAC69831B2724E14FBB98E9C0610A6554BC45D549AD46CF1F3E78B3D99F39E8EF1D8021B8B755
    'C25387EC041BD04257779DB13A620450FE8E334BE117DC069319F19AAE8A1F1F9A142DF3967C24FEA4F222BFDB2301BBAF68C0ACCB5601D054749C0E
    '4A30DDCF316EC54613F96F89484F64B5C6BD3AB7E27629CC380B02F7FC81DDEE529FA057B2C758A321091DAD2DB9AB25BB79F890613E50867BDB304F
    '86266CC6C5EE67B816EEA405320615BA1C9C5D3CA8F8E8D8CE93EA9C11ABC053C0A222A2ACC2F9E3CBCCE1294A971C0FB895EAD59148A0A6EB97A26F
    'D263D3F9FFBA8B068147051D9AE5CC33AE00FA95F77FDAD2CCA1DEE380E7D2FB0501A38A663C203A42A5DACDC00A264921E2ECA2F348FD16C9CF1B13
    '15AE589F0016EEF7EA56C1C5486F2DB753BC77F20B367B3757A091B5D5BF1F9240F0963A03CBFBF647C8949BE870D5D5E9A37F93803A2A3D522C1EBD
    'A6058846DCC8485BC8EA7D992EBBE1D1AEA1DB87925099FCE1C603292C276C0BDE64452D9B61FE6D168230A501E5381C5F6117A19AC247FF4103B198
    '124FE25A0E6EAB03809321D4EDAB106CA07B750EC380C8B4AFE27E8ADF38C87117359A6E2554D8BA68387F23DE7AFE680881A18F90AB916929F2C56B
    '7C49A0EBBD9E35744D6472BF92EB510E9DD0927F949C5116F65771F981EF6C4EA5DAB4A3A8831A8D65D489059EEC2D646BD5C2628FC36DD965EF3DE2
    '0928A08DC99814A3A35A684C9ADE765ED2B612CB0DD520B4E774EB819622C7D7037EC1A8DFD785942AE05D2440074ADFB0077C9AE11CB1DCA53F7651
    'BBD0D0BD1EEF65BD8BEE22BEF2D30E8B4453F5A44CC1B1130637CFDF4CF85D6103F8515D2B6FD5A555D77B5D58193C145DBDBDFC82318B4B5A562EB3
    '8D58A9E675120761DF677C8A9F60A09D1EB573835B428C86F1A3D9B43EAA3E8C1DC7EC9F92FAC9311C20D80C6F329F123338BEEA8528D9AE4706F83B
    '81EADE8E6260BFCBB960D74726806E041006BB159A5222801AD7351D5DB7B21D19E7D95AD327D1EEA8D3ABA66ED21A66EE45222BE039E3296F449FA7
    'C3A2643625A6A0AA559833B86B39B29DD8F5096485AB9BFAC7FE63B496AD8CC2D4DE02187F93E9F1E7E98B789AFBEFD9A785990AB8309511A4B37D0D
    'FF3046F3DAF4CA50D350603C7DF7CEE7161AAF60A2845C1602625DF211457315708C58B59C35FF675275372F4E48AF523D6B334F2C9A31417F977A3B
    '6166BBE84588286BE053D03630005F540126F0FC660C63FC81ED0D5E6A1AE51579C80B2AEFA1CB8705BF8253F511E4102C7563E56305610E673130F2
    'C5CF3397079EE5F1843C50EF14EC47D18F93377EAF0D6BFF6C04863969EA8E450DD7D357D1B954042FBA90804E938A4BC50B196D7933EBB629B706E9
    'AFD86071A83F97CA334E62CFABA7F38AE2AD4937199A8DD108EE187916AA67780F1922001C0D65F4F1E569C6E8ADB049E8D8517EB1F4543C055C4168
    'CAC05D219BD9992DD79E39F84E4D75C94D447457E9B60FBFB2D278018A60C9D0F3704A5339120EC10AAF2FFD5AD043E63DB13B4D5229ADF1543E185B
    '23D015F1A9D7A1AA9F093280CDAC49CE69ED6FAAB3280AAE4B628376DCCC359BD4A61D8C74D655D24CA4E93A2FA526EC8E7828097EF8CC4384701ECA
    '9733B464891BA2561E583011326DC1166FA7C8222091F7E49B0A1E4B369F3544DC32E151372539358C57AFD680856CD308A3001EDE0AB51C6B9DD17C
    'C9C6FE1144388D54E56735461897B91929D833A920B28CD36CA8F3CFCEC4D73955FAFDCDBBB32913131F57482368D2AD93C8C201021236977B96CE80
    'CB335A69442C89830C2360E5FFDB5BA6B7DF35ED61013858EC244B5B59A87DC9A64D47916387FB0EBC8604195E1FF4F4007856149B5F9636542E5C70
    '4A95E394BEE055C709BB6C86F6BD4E2B17189E64D348713D4E0D6C736B993E4257B538FC7C0F13E594FB5CBA460C56576BF820533B376B877BEFC5BE
    'DF41DD72FC3AB8F82CE61B404AB5E5FA8FE9B7842DCC87438BBED2510CFE11787B8D780A11C458408A8F8E060FB13EF92AE5AA13E01A5464EAA7BFCB
    'C9D72BF5DE421DB4BB6D58D1FEC6B773B9E99ABFFEDABC87F87A52DF6CE7CA72E4915FF51F32AE6B7F0CD2716B4A5B9087B62BF23DC8C6EC192B5FE6
    '0CF2BD37FF4CFAEFA94B1634EC1D0FDED1869924BA2F516F4F7BEFA090C01D64D4D5F427641E76F530E6FE2705F0B71CFC1D9B6FE8D45B8CBB2B64CE
    'CA53F0911114A4E7E7F2FB6B1D9CDAC113E3BE89C86E73EFAB3548D7EABC4496713DDC416EA38EE12D5CA3AC374AC7BF93F48BC6AB2E943EE7DFE65F
    'E3F28E1E248E95CB2B5A25A742BD6AB974D8351E9E0B44031552F0FD04B3176E40FF6C27ECD08EC7D8C5D430BD66270A900A2A989D3996CD10FCB6FA
    '99A51D3E7A35AEA75378D18D6B556F748F55E9C7578D515AF077B58D6220994A2C97D8084CA868F572BC590E54F272EFFAC034A607428282A382E13F
    '8ED048736CA1C0D9C7C36A43550757F3F9110A6C8400EE8B7F2FD8894D41433D33E36FDB60B984DDAFD8F7C750C7E3764AB833479312E0873A856B22
    '7B801E523AD36128D96AC6294071938E32EF2290DBF53B5784A45BCA2CB3EDF539B0A5F9A2F8C4B244FE13F7D50303D0CACF2435503B5ED4A9C484DC
    '0514B773CC68D09C1118591CC8D8A97614C7A3EC64CDEA944C94552A1AA50D7BCACD42F3D86FA74432A197786922DC22AB546884082A92C67079F638
    'B2801323C24E4EDD3545B8AF6737703F95C4CCC5074FAB96E198306F2A02CA122599B36C2D65D3BF73A57A6A8F2E6F11B4A478BE4604D3A5217CF8C1
    'C0873A96126C500DA74066F44BD7760795B39209BAA08B31316562E6527DB5AF5C2E61620BAE3145A9173176809BAF9C593F7877582F2384E8F22220
    'AFEF493C3F52A0B1FBE3C941F78AEB1DDEB31F058602BA0E9A10DEF6DBE28AC23F22FF48B409EFECE7D17AB6C597D731376E789BC721920F6C91C081
    '4AAFD9165759E0A04692A49954B90D194E1D19527271A5A031162256D7C2B610B1F020D85412580A6A382A3A2C7FF758B0EEC934CD62F178778FF2BC
    '79C17A645F00872926DDC1CD6E7AB1CEFA67ACFA612729AFF1B1C9C085DD2AE25EB0DEE6253EA86F8EC4EFAA2A0AA63BDE92C9EDE92B621C664A27C9
    'A3BC507A99A7DC765AC5E73EBB587F798CBA18ED27CF1D984E2FE0CA9637AF6D9159110C71E63CABBCEFB2EC1944197B4C7008C137576942884A6FD1
    'CAD6A513EA8330F29AD61FD75B9279CE48B0FA602500096D3488608705256F53B29DD81850E94E75396E3AAA80D2EB66D7819FBD8CD32F24A0CE5B1F
    '3528F50F2DD2B2DD969AEEC0DC1A8D12EAFAD522FB9792DE4225D3A3868D2E128BD2F67DF17CB7F35CD7E6A893D6483FAEF1230572B797F1FB549594
    '122AD130082D6D2F3BEF03B7C9327D390893070E16A28EE58633F4D267561EADCF5537A9DD9052F4115500F6A0B4FEF85FBD76A265AE339C99F29D35
    '702F1DC8A9F308D273FA275CD8189035ADECF5DEBD5AB8F96D16C8CACE24B1BCCB1E1CEF571CF354C10AA0C935E5824566586E1918FD1EE53C1E2853
    'BFEA72B90DCCE3B5402BA25638BD68A90AADFA7CFD2D2C956116004D3296642D04E8B40968AD8784E18DE4815C6FC12528EB1AB86F04F3129C44B6CD
    '2BEE83ACE7CCC2047F5B1C5E69CC163F9A9E5E625353AA606A19D6EC5D3DE6AF559A6274FDB22CF54DFD179C1F692178ADFA82E155D02F5BCB70B438
    '89F32CAFBF07A9021268BDF247D90781816E314FF5A0032C85B1E1301C5FD764648D9726AAB90F116D835823C3D4B59D2F05FA060B929844B1504074
    '3EEC4DFC3ADB368CC2E64476C622800295DB0A54615B3EACD9E1B8BD0FDA33634E5AA9B8A7B21FEEA2F2C7AE93A92D709DD3E88890FCDA34835939CF
    '18C7FAEA51C65D0970AC54B8B7DE08BADED9E6547BF358FA061363C0C9346DD74082095AE22DFC6F429B5F9D6FC9759D6909751C33AD0DAAE8344FC4
    '4F531FF2E97FBD9EB43BBBEA90AFA32B8DB4F11241D159D362408C194D713B86E960919ABC358EB8B528A9D132E2334AAA3124664C6B8E12A758F120
    '22D38D7320BBFCC9DF8A381A2A48F300C4C3A89D0CCDE1A876C13B558E60DA4F3DB07AB212103FEDFEAC65D737F5BD483915574AAE11B9C608D5EF86
    'ED01D3A866835ECE536FF224F2904968363B1BDDF9FFE1A51945B659614D438D450C8AB62F9398CAF333D9E2CEC689C3CB0548ACA5621362E8370DED
    'FF7574A449422FF84BBAB6FCB8626633DCB9EE78BBB5414B17715F9514F452874988753355AB312903A3ED12CD64A5EEC58CDDA1A340F49322740BA9
    '7FD0261754B7D4F9AE089436938D16897733D68EB5C387A3ABD3A408D7AB52AFBCD927170A8265C4816C09CAFC317B317B63972EA61B97F53108ED7E
    'F6AB81B3EAA98D3027261370DBC779C0914D77ABB39A8C4F81C753E3635EEF905F85B06290A91629ACE923819D2F1F9C77B27AFE71B715FCE31284F7
    '4F03CD76870420835ED794AE4F4410930D39D54C3C1424870178DC220141DDA2423832DDDAE2F6A54539D8D11F9C763FDF246495ADD35AE546163EBD
    'A8BC179C756F01BCE61700A32FE9730C10B30465906CBF606C7A505E06ECF05CD06E7C23AA4F0758D37AEC5FB42A28B7212C26F2AD736A2CBA5755FB
    'CE3F1913D21AD792548AE9DC0DC01D10CA901B244808FDA520C4AC934084EF6F5ECEAD62A4CBF8746633104EFB32D621F3A1229F368CA8B8B5A2A250
    '73C1F732287EDAE82AF809DCF8B2AD32C411A7E8A9D33ED3D7B3E4B701CB22620D81C90EA3C502105869768BA16985F346CE40C42E3AEE085B35C066
    'D7EA9830897311BC038A6EA0FBDCB1FC675FEF3BDBAF8870507B979540D3918779AA4C380E19DE5120C0986E1D161E58C16BED73204B940D9AF39CA9
    'A50682093220D04AA3DBB0696C85B028EE5822B265FE1EA2D44A6385E24EF6B624F64A352E01388BC241EDED85FCDDD84378AAC9C50E194C30ECCBCB
    'EF2BC0DEC6C21C6AC4586DF51F7D0C33398755D1C60B9AAB0BC9AC54DACD015A97E46022EBD0F180E075E27519D2B2E7525FF3A212A3A42314960B4C
    '00E5FE47AC51B697565207C9F72B64358D2C2E794D46A0C0E4F8A5DAA0CD9F2D25A93C7BBCE789A2BFFE4DC57420AC8078CB3320840E2C344654581C
    'B7872BBEEC1C074C471C8713FDE8AA923EA68169583A2B9F96BE0AB4A810B1E779BA547E50E61CA2AF8E9E612DF48E6D971F7CB0DBFD27471495E0AA
    '4BBE71E19E611A1BD946D9F17B6F123435E02099633FC915CA5EBF2856B33D374395582DB736985A28D5EF7CDED91258E0F18A7EE4C1692422F575C4
    '3E0A4A9A64AC92982A6D77BFA82979F3B0D769ABA52C3E917DD62F2CDFF04E0FB9D032206BA13D11BB0A4F7EC25850AF1E0F5E8D573BF5218FD461FD
    'D39263D01204F2F92D11F7C63A099D4EFEDA46A578D9A03DC5A73DAFEAB25A108A988A471A7A90888B961492E1EC1549308E8F97A995715F687F354E
    'D9F79C93F781CA5413918F5C9742F9C493E6694DD9444A2F1F82318297836E611DEFC7E3B3B57E68BB1979BA847FB1E7C44733DEF09B377E44B03981
    '7E9975B5F494E0F0DA382A8EC80B79F6DBF24002679C45ADC60A30AD085E283E5408839A7E1C0E68532FAADF41FBA18800A112931737D63FF9AB80AC
    '2DBA3EFCC7E883E25BC0E5B6EC82241C5B9099ACE89330545437BEC2978F16A6178C9A057F768ECDE345F22EE7184AC07D70CE675A62010A62AB3017
    'C95B845DE83027F1CC2639A07A4C709F06B8684E5E1B5C608AAAC032690735223CB0A9B28C1D0EC386A7FAF6B0122FC3A7F33F126A7B30E5EB65A4F8
    '79E1DDD797FC90CD3EB88AA85262A45F20DB9A412CA2327ABB8AB77971CA5080AD248801B48A8F3546D4650AF2164560016FA78A8403BB20D5E334A8
    'D1D1AA6853F1AA6071BF673093EEECD7B88048FBE1B4AD903CE119043D3B4A17660C18559A7F05052B950C435178311EFC96C56D89BD96293E204181
    '0A267E28079461B6A95A79F6387EE0F64A6FC30CC2129FCBBAF5577E902DD5A470B609F54EF70F761ACB34BE6ECE4CDEFB6717151E6548440BF80CCB
    'FB0EFF1C5A6D966F91755F6DE7300E1810DF664FDC01380051BAFBDC648C294BFBD2190EEDAA6234FD6E79BC238167C77A81256B7766BF38056E9A66
    '942F5638DF9E8819

#Region "NO USE"
    ' 取得ACTNO
    'Public Function Get_StudACTNO(ByVal OCID1 As String, ByVal IDNO As String, ByVal oConn As SqlConnection) As String
    '    Dim rst As String=""
    '    Dim sql As String=""
    '    sql=" SELECT ACTNO FROM STUD_BLIGATEDATA28e WHERE IDNO=@IDNO AND OCID1=@OCID1 "
    '    Dim parms As New Hashtable
    '    parms.Clear()
    '    parms.Add("IDNO", IDNO)
    '    parms.Add("OCID1", OCID1)
    '    Dim dt As DataTable=DbAccess.GetDataTable(sql, oConn, parms)
    '    If dt.Rows.Count > 0 Then rst=Convert.ToString(dt.Rows(0)("ACTNO"))
    '    Return rst
    'End Function

    'BtnCheckBli2.Attributes("onclick")="open_SD01001sch2();return false;"
    'BtnCheckBli2
    'Protected Sub BtnCheckBli2_Click(sender As Object, e As EventArgs) Handles BtnCheckBli2.Click
    'End Sub

    ' 自動帶入投保證號
    'Protected Sub BtnAutoInputActno_Click(sender As Object, e As EventArgs) Handles BtnAutoInputActno.Click
    '    ActNo1.Text=""
    '    Dim StudACTNO As String=Get_StudACTNO(rqOCID, IDNO.Text, objconn)
    '    ActNo1.Text=StudACTNO
    '    Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(2);</script>")
    'End Sub
    'Dim StudACTNO As String=Get_StudACTNO(rqOCID, IDNO.Text, objconn)
    'If ActNo1.Text="" Then ActNo1.Text=StudACTNO
    ' ------ 2009/06/01 以後除了己選擇負擔家計婦女者外,不顯示負擔家計者婦女選項
    'sqlstr="select STdate from class_classinfo where OCID =" & Request("OCID") & ""
    'STdate=DbAccess.ExecuteScalar(sqlstr)
    'sqlstr2="SELECT case when IdentityID like '%03%' or SubsidyIdentity ='03' then 'Y' else 'N'  end as IdentityID  FROM CLASS_STUDENTSOFCLASS WHERE  SOCID='" & Request("SOCID") & "'"
    'sqlstr3=DbAccess.ExecuteScalar(sqlstr2)

    'If sm.UserInfo.TPlanID="28" Then
    '    '產業人才投資方 學員資料維護
    '    '20100325 AMU 取消 29天然災害受災民眾
    '    '20100325 AMU 改回 其他(就服法24條)="更生保護人"(鍵值代碼10)
    '    'sql="SELECT * FROM Key_Identity WHERE IdentityID IN ('01','03','04','05','06','07','10','26')"
    '    '如果開訓日期小於2009/06/01 就顯示 03:負擔家計婦女
    '    If STdate < CDate("2009/06/01") Then
    '        '20090922  Andy edit 為因應88水災受災參訓身分別增列「天然災害受災民眾」
    '        hide_IdentityIDType.Value=2 '為了判斷身心障礙使用
    '        'sql=" SELECT IdentityID,case when IdentityID='10' then '其他(就服法24條)' else Name end as Name"
    '        sql=" SELECT IdentityID, Name"
    '        sql &= " FROM Key_Identity WHERE IdentityID IN ('01','03','04','05','06','07','26','10','28')"
    '    Else
    '        If sqlstr3="Y" Then
    '            '20090922  Andy edit 為因應88水災受災參訓身分別增列「天然災害受災民眾」
    '            hide_IdentityIDType.Value=2 '為了判斷身心障礙使用
    '            sql=" SELECT IdentityID, Name"
    '            sql &= " FROM Key_Identity WHERE IdentityID IN ('01','03','04','05','06','07','26','10','28')"
    '        Else
    '            '2009/06/24增加獨力負擔家計者,2009/07/01拿掉負擔家計婦女代碼03  '20090922  Andy edit 為因應88水災受災參訓身分別增列「天然災害受災民眾」
    '            hide_IdentityIDType.Value=1 '為了判斷身心障礙使用
    '            sql=" SELECT IdentityID, Name"
    '            sql &= " FROM Key_Identity WHERE IdentityID IN ('01','04','05','06','07','26','10','28')"
    '        End If
    '    End If
    '    '    '20090123 Andy   2009年新增一個「非自願離職者」身分別
    '    '    If CInt(sm.UserInfo.Years) > 2008 Then sql &= "when IdentityID='02' then '非自願離職者'"
    '    '    sql &= " else Name end as Name"
    '    '    If CInt(sm.UserInfo.Years) > 2008 Then
    '    '        sql &= " FROM Key_Identity WHERE IdentityID IN ('01','03','04','05','06','07','26','10','02')"
    '    '    Else
    '    '        sql &= " FROM Key_Identity WHERE IdentityID IN ('01','03','04','05','06','07','26','10')"
    '    '    End If
    '    dt=DbAccess.GetDataTable(sql)
    '#Region "Function 1"

    '取出學習卷的學員身分資料
    'Sub GetDGIdent(ByVal SOCID As Integer)
    '    Dim sql As String
    '    Dim dr As DataRow
    '    sql=""
    '    sql &= " SELECT b.Share_Name "
    '    sql &= " FROM Adp_DGTRNData a "
    '    sql &= " JOIN Adp_ShareSource b ON a.OBJECT_TYPE=b.Share_ID AND b.Share_Type='301'"
    '    sql &= " WHERE a.SOCID='" & SOCID & "'"
    '    dr=DbAccess.GetOneRow(sql, objconn)
    '    If dr Is Nothing Then
    '        DGIdentValue.Text="學習卷的學員身分不明"
    '        DGIdentValue.ForeColor=Color.Red ' DGIdentValue.ForeColor.Red
    '    Else
    '        DGIdentValue.Text=dr("Share_Name")
    '    End If
    'End Sub

    '取出學習卷的學員身分資料
    'Sub GetGovIdent(ByVal SOCID As Integer)
    '    Dim sql As String=""
    '    Dim dr As DataRow
    '    'SELECT * FROM ADP_SHARESOURCE WHERE Share_Type='528'
    '    sql=""
    '    sql &= " SELECT b.Share_Name "
    '    sql &= " FROM Adp_GOVTRNData a "
    '    sql &= " JOIN Adp_ShareSource b ON a.SPECIAL_TYPE=b.Share_ID AND b.Share_Type='528' "
    '    sql &= " WHERE 1=1 "
    '    sql &= " AND a.SOCID='" & SOCID & "' "
    '    dr=DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then GovSpecial_Type.Text=Convert.ToString(dr("Share_Name")) '推介單個案區分
    '    'SELECT * FROM ADP_SHARESOURCE WHERE Share_Type='527'
    '    sql=""
    '    sql &= " SELECT b.Share_Name "
    '    sql &= " FROM Adp_GOVTRNData a "
    '    sql &= " JOIN Adp_ShareSource b ON a.OBJECT_TYPE=b.Share_ID AND b.Share_Type='527' "
    '    sql &= " WHERE 1=1 "
    '    sql &= " AND a.SOCID='" & SOCID & "' "
    '    dr=DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then GovObject_Type.Text=Convert.ToString(dr("Share_Name")) '推介單身分別 
    'End Sub
    '查看學號是否存在於3合1(就服)
    'Function Get_AdpTotal(ByVal tmpSOCID As String) As Integer
    '    Dim rst As Integer=0
    '    If tmpSOCID="" Then Return rst
    '    tmpSOCID=TIMS.ClearSQM(tmpSOCID)
    '    If tmpSOCID="" Then Return rst
    '    'https://jira.turbotech.com.tw/browse/TIMSC-272
    '    Dim sqlStr As String=""
    '    sqlStr=" SELECT COUNT(1) CNT FROM dbo.ADP_GOVTRNDATA WITH(NOLOCK) WHERE SOCID='" & tmpSOCID & "'"
    '    Dim dr1 As DataRow=DbAccess.GetOneRow(sqlStr, objconn)
    '    If dr1 Is Nothing Then Return rst
    '    rst=Val(dr1("CNT"))
    '    'Catch ex As Exception
    '    Return rst
    'End Function
    '重載數據(JS無法讀取有效值)
    'Sub ReLoad_SB4IDx()
    '    IDNO.Text=TIMS.ClearSQM(IDNO.Text)
    '    hidSB4ID.Value=TIMS.ClearSQM(hidSB4ID.Value)
    '    If hidSB4ID.Value="" Then Exit Sub '為空離開
    '    Select Case PriorWorkType1.SelectedValue
    '        Case "1" '曾工作過
    '        Case Else
    '            Exit Sub '未選擇 (曾工作過) 離開
    '    End Select
    '    Dim drSB4ID As DataRow=TIMS.Get_BLIGATEDATA4(hidSB4ID.Value, IDNO.Text, objconn)
    '    If drSB4ID Is Nothing Then
    '        'Common.MessageBox(Me, "查無資料，無法回傳值")
    '        Exit Sub
    '    End If
    '    '任職單位名稱
    '    PriorWorkOrg1.Text=Convert.ToString(drSB4ID("COMNAME"))
    '    '投保單位保險證號
    '    ActNo2.Text=Convert.ToString(drSB4ID("actno"))
    '    '投保薪資級距
    '    PriorWorkPay.Text=Convert.ToString(drSB4ID("SALARY"))
    'End Sub
    ''產投檢查ECFA按鈕
    'Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
    '    Dim vsMsg As String=""
    '    Dim vsFOfficeYM1 As String=""
    '    Dim flagFOfficeYM1 As Boolean=False ' 第1筆受訓前任職起迄資料 迄止日期 不須填寫

    '    ActNo1.Text=TIMS.ClearSQM(ActNo1.Text)
    '    If Not TIMS.IsDate1((STDateHidden.Value)) Then vsMsg &= Cst_Msg7 & "!!\n"

    '    If vsMsg="" AndAlso DateDiff(DateInterval.Day, CDate(Cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
    '        If TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '            '產投與在職判斷方式
    '            vsFOfficeYM1=""
    '        Else
    '            '第1筆受訓前任職起迄資料 迄止日期未填寫
    '            If FOfficeYM1.Text <> "" Then FOfficeYM1.Text=Trim(FOfficeYM1.Text)
    '            If FOfficeYM1.Text <> "" Then
    '                Try
    '                    FOfficeYM1.Text=CDate(FOfficeYM1.Text).ToString("yyyy/MM/dd")
    '                Catch ex As Exception
    '                    vsMsg &= "帶入" & cst_ECFA & "基金產業被認定日\n第1筆【受訓前任職起迄資料】結束日期 格式有誤(yyyy/MM/dd)!\n"
    '                End Try
    '            Else
    '                vsMsg &= "帶入" & cst_ECFA & "基金產業被認定日\n第1筆【受訓前任職起迄資料】結束日期未填寫!\n"
    '            End If
    '            vsFOfficeYM1=FOfficeYM1.Text
    '            If vsFOfficeYM1 <> "" Then flagFOfficeYM1=True
    '        End If

    '        ActNo1.Text=TIMS.ChangeIDNO(ActNo1.Text)
    '        If vsMsg="" AndAlso ActNo1.Text <> "" Then
    '            If flagFOfficeYM1=True Then vsMsg &= "依 第1筆受訓前任職起迄資料 迄止日期" & vsFOfficeYM1 & "\n"
    '            Try
    '                If TIMS.CheckIsECFA(Me, ActNo1.Text, "", STDateHidden.Value, objconn)=True Then
    '                    vsMsg &= "---此學員為" & cst_ECFA & "基金補助對象---\n"
    '                    'Common.SetListItem(BudID, "")
    '                    If BudID.SelectedValue <> "97" Then Common.SetListItem(BudID, "97") '自動帶入 " & cst_ECFA & "基金預算別
    '                Else
    '                    vsMsg &= "!!!此學員不是" & cst_ECFA & "基金補助對象!!!\n"
    '                End If
    '            Catch ex As Exception
    '                vsMsg="!!查詢時發生錯誤!!"
    '                Page.RegisterStartupScript("ChangeMode2", "<script>ChangeMode(2);alert('" & vsMsg & "');</script>")
    '                Common.MessageBox(Me, ex.ToString)
    '                Exit Sub
    '            End Try
    '        Else
    '            vsMsg &= "請填寫 投保單位保險證號!!!\n"
    '        End If
    '    Else
    '        vsMsg &= Cst_Msg7 & "!!\n"
    '    End If
    '    '企訓專用
    '    Page.RegisterStartupScript("ChangeMode2a", "<script>ChangeMode(2);alert('" & vsMsg & "');</script>")
    'End Sub

    '職前檢查ECFA按鈕
    'Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
    '    Dim vsMsg As String=""
    '    Dim flagFOfficeYM1 As Boolean=False ' 第1筆受訓前任職起迄資料 迄止日期 不須填寫
    '    '第1筆受訓前任職 起迄資料 迄止日期未填寫
    '    Dim vsFOfficeYM1 As String=""
    '    If TIMS.Cst_TPlanID28AppPlan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then
    '        '產投與在職判斷方式
    '        vsFOfficeYM1=""
    '    Else
    '        '第1筆受訓前任職起迄資料 迄止日期未填寫
    '        If FOfficeYM1.Text <> "" Then FOfficeYM1.Text=Trim(FOfficeYM1.Text)
    '        If FOfficeYM1.Text <> "" Then
    '            Try
    '                FOfficeYM1.Text=CDate(FOfficeYM1.Text).ToString("yyyy/MM/dd")
    '            Catch ex As Exception
    '                vsMsg &= "帶入" & cst_ECFA & "基金產業被認定日\n第1筆【受訓前任職起迄資料】結束日期 格式有誤(yyyy/MM/dd)!\n"
    '            End Try
    '        Else
    '            vsMsg &= "帶入" & cst_ECFA & "基金產業被認定日\n第1筆【受訓前任職起迄資料】結束日期未填寫!\n"
    '        End If
    '        vsFOfficeYM1=FOfficeYM1.Text
    '        If vsFOfficeYM1 <> "" Then flagFOfficeYM1=True
    '    End If
    '    If ActNo2.Text <> "" Then ActNo2.Text=TIMS.ChangeIDNO(ActNo2.Text)
    '    If vsMsg="" AndAlso ActNo2.Text <> "" Then
    '        If flagFOfficeYM1=True Then vsMsg &= "依 第1筆受訓前任職起迄資料 迄止日期" & vsFOfficeYM1 & "\n"

    '        Try
    '            If TIMS.CheckIsECFA(Me, ActNo2.Text.ToString, vsFOfficeYM1, "", objconn)=True Then
    '                vsMsg &= "---此學員為" & cst_ECFA & "基金補助對象---\n"
    '                If BudID.SelectedValue <> "97" Then Common.SetListItem(BudID, "97") '自動帶入 " & cst_ECFA & "基金預算別
    '            Else
    '                vsMsg &= "!!!此學員不是" & cst_ECFA & "基金補助對象!!!\n"
    '            End If
    '        Catch ex As Exception
    '            vsMsg="!!查詢時發生錯誤!!"
    '            Page.RegisterStartupScript("alert2", "<script>alert('" & vsMsg & "');</script>")
    '            Common.MessageBox(Me, ex.ToString)
    '            Exit Sub
    '        End Try
    '    Else
    '        vsMsg &= "請填寫 最後投保單位保險證號!!!\n"
    '    End If
    '    Page.RegisterStartupScript("alert2", "<script>alert('" & vsMsg & "');</script>")
    'End Sub
#End Region
End Class
