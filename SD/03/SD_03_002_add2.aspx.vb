Partial Class SD_03_002_add2
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
    '否則，不能儲存!"，並不予儲存，除非參訓身分別已勾選"屆退官兵(須單位將級以上長官薦送函)"選項'，才可儲存。
    'SELECT * FROM KEY_IDENTITY WHERE 1=1 AND NAME LIKE '%屆退%'--12	 屆退官兵(須單位將級以上長官薦送函)

#End Region

    '保險證號(ActNo)前二碼意義：
    '01:工廠'02:公會(工會)'03:漁會'04:政府機關'05:公司'06:農會'07:自由業'08:自由業'09:職訓保

    '02:公會/'03:漁會/'06:農會
    '若是登入年度為 2017年以後，則傳回2，其餘為1
    Dim iPYNum17 As Integer = 1 'iPYNum17=TIMS.sUtl_GetPYNum17(Me)
    Dim flag_show_actno_budid As Boolean = False '保險證號/預算別代碼 false:不顯示 true:顯示

    Dim tZipLName As String = "" '暫存資訊
    Dim tZipNameN As String = "" '暫存資訊
    'Const cst_inline1 As String="inline"
    Const cst_inline1 As String = ""
    Const cst_none1 As String = "none"

    '該民眾不具失、待業身分，不得參加失業者職前訓練。
    Dim dtBLIDET1 As DataTable = Nothing
    'dtBLIDET1=TIMS.Get_dtBLIDET1(rqOCID, objconn)
    Dim strScript1 As String = ""
    Const cst_SearchSOCID As String = "SearchSOCID"
    Const Cst_Msg1 As String = "此學員現在有職訓生活津貼資料，不能修改 "
    Const Cst_Msg1b As String = "(系統管理者開放修改) "
    Const Cst_Msg2 As String = "學員資料確定，不可修改"
    'Const Cst_Msg14 As String="學員資料維護於訓後14日鎖定" '學員資料維護於訓後14日鎖定
    'Const Cst_Msg14ok As String="中心承辦(預算別)於訓後14日開放修改" '學員資料維護於訓後14日鎖定 (產投)
    '28:產業人才投資計劃 54:充電起飛計畫（在職）15:學習券 '學員資料維護於訓後21日鎖定(排除計畫:28.54.15)
    Const cst_limitDay21st As Integer = 21 '委外、中心開訓後可修改期限(開訓 產投、TIMS限定天數)
    Const cst_limitDay21ft As Integer = 21 '委外、中心結訓後可修改期限(產投、TIMS限定天數)
    Const Cst_Msg21 As String = "學員資料維護於開訓後21日鎖定" '學員資料維護於訓後21日鎖定(委訓)
    'Const Cst_Msg21ok As String="中心承辦(預算別)於訓後21日開放修改" '學員資料維護於訓後21日鎖定 (產投)
    Const Cst_Msg21ok As String = "分署承辦(預算別)於訓後21日開放修改" '學員資料維護於訓後21日鎖定 (產投)
    Const Cst_Msg30 As String = "學員資料維護於訓後超過30日不能修改" '學員資料維護於訓後30日不能修改
    '委外職前訓練
    '針對委外職前訓練計畫，系統權限 限制 (Page Load/儲存鈕做限制)
    '針對委外職前訓練，系統開放各分署(中心)修改權限為21日，委外訓練單位仍保持開訓後14日後，限制資料修改之邏輯。
    'Const Cst_TPlanIDCanEditStud_id37 As String = "37" '委外職前訓練
    '--Cst_MsgTPlanID37a="針對委外職前訓練，系統權限 限制各中心修改權限為21日內之後不能修改"
    'Const Cst_MsgTPlanID37a As String = "針對委外職前訓練，系統權限 限制各分署修改權限為21日內之後不能修改"
    'Const Cst_MsgTPlanID37b As String = "針對委外職前訓練，系統權限 限制委外訓練單位修改權限為21日內之後不能修改"
    Const Cst_Msg30x28 As String = "學員資料維護於訓後超過3個月不能修改" '學員資料維護於訓後3個月不能修改 (限中心)

    'Const Cst_Msg3 As String="投保單位保險證號為 受ECFA影響之單位,預算別必須為協助，補助比例必須為特定100%!" '產投
    ''BUDID 97
    'Const Cst_Msg3b As String="預算別為協助，補助比例須為特定100%!" '產投
    'Const Cst_Msg4 As String="預算別為協助，補助比例須為特定100%!" '產投
    'Const Cst_Msg5 As String="投保單位保險證號為 非受ECFA影響之單位,預算別不可為協助!" '產投、TIMS
    'Const Cst_Msg6 As String="投保單位保險證號為必填資料" '產投、TIMS
    'Const Cst_Msg7 As String="該計畫開訓日期為 2011/4/15 日後才可使用 協助基金補助對象!" '產投

    '協助 -> 公務(ECFA) BudID: SELECT * FROM Key_Budget 
    Const cst_ECFA As String = "公務(ECFA)"
    Const Cst_Msg3 As String = "投保單位保險證號為 受ECFA影響之單位,預算別必須為 公務(ECFA)，補助比例必須為特定100%!" '產投
    'BUDID 97
    Const Cst_Msg3b As String = "預算別為公務(ECFA)，補助比例須為特定100%!" '產投
    Const Cst_Msg4 As String = "預算別為公務(ECFA)，補助比例須為特定100%!" '產投
    Const Cst_Msg5 As String = "投保單位保險證號為 非受ECFA影響之單位,預算別不可為公務(ECFA)!" '產投、TIMS
    Const Cst_Msg6 As String = "投保單位保險證號為必填資料" '產投、TIMS
    Const Cst_Msg7 As String = "該計畫開訓日期為 2011/4/15 日後才可使用 公務(ECFA)基金補助對象!" '產投
    'BUDID 04
    Const Cst_Msg8 As String = "預算別為再出發，補助比例須為特定100%!" '產投
    'BUDID 99
    Const Cst_Msg9 As String = "預算別為不補助，補助比例大於0，有誤!"  '產投
    'BUDID 02
    Const Cst_Msg65 As String = "參訓學員為逾65歲者(滿65歲生日隔天), 其預算別一律運用就安預算!!預算別，(非就安)有誤!"  '產投

    Const cst_str45yearsOld As String = "(此學員為中高齡)" '45歲~65歲
    Const cst_str65yearsOld As String = "此學員為逾65歲者(滿65歲生日隔天)" '"(此學員為65歲(含)以上)" '65歲(含)以上

    Const Cst_Msg3c As String = "身分別為一般身分，補助比例 不可為特定100%!" '產投
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

    'Const Cst_errMsg2 As String="資料異常，無法修改生日與姓名，如需修改請將資料提供給中心承辦人。" '產投、TIMS
    Const Cst_errMsg2 As String = "資料異常，無法修改生日與姓名，如需修改請將資料提供給分署承辦人。" '產投、TIMS
    Const Cst_errMsg2b As String = "(學員基本資料有重複，造成系統無法儲存，請提供相關資料聯繫OJT 窗口。)" '產投、TIMS
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
    Const Cst_20110415 As String = "2011/04/15"

    'Dim blnCanAdds As Boolean=False
    'Dim blnCanMod As Boolean=False
    'Dim blnCanDel As Boolean=False
    'Dim blnCanSech As Boolean=False
    Dim blnTPlanUseEcfa As Boolean = False '該計畫是否使用ECFA True:使用 False:不使用
    '在職進修 取消必填，學員資料維護 (SD_03_002_add.aspx)
    Dim sTPlan06_G22 As String = ""
    Dim flag_BudIDNoLock As Boolean = False 'flag_BudIDNoLock BudID NoLock'如果是中心承辦人，預算別不鎖定。by AMU 20140328

    Dim iBudFlag As Integer = 0
    '屆退官兵者 (依系統日期判斷)
    'Dim flagTPlanID02Plan2 As Boolean=False '判斷計畫為自辦職前。

    'OJT-21020401：在職進修訓練(自辦) - 學員資料維護：判斷學員為現役軍人時於投保保險證號顯示「在役軍人」、預算別判斷為「就安」
    Dim flagTPlanID06Plan3 As Boolean = False
    Const cst_Serviceman As String = "在役軍人" '在役軍人

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
        'Call TIMS.sUtl_SetMaxLen(dt, "ACTNO", ActNo)
        Call TIMS.sUtl_SetMaxLen(dt, "ACTNAME", ActName)
        'Call TIMS.sUtl_SetMaxLen(dt, "ACCTNO", AcctNo2)
        Call TIMS.sUtl_SetMaxLen(dt, "SERVDEPT", ServDept)
        Call TIMS.sUtl_SetMaxLen(dt, "JOBTITLE", JobTitle)
        'Call TIMS.sUtl_SetMaxLen(dt, "ADDR", Addr)
        'Call TIMS.sUtl_SetMaxLen(dt, "TEL", Tel)
        'Call TIMS.sUtl_SetMaxLen(dt, "FAX", Fax)
        'Call TIMS.sUtl_SetMaxLen(dt, "ACCTHEADNO", AcctheadNo)
        'Call TIMS.sUtl_SetMaxLen(dt, "ACCTEXNO", AcctExNo)
        'Call TIMS.sUtl_SetMaxLen(dt, "BANKNAME", BankName)
        'Call TIMS.sUtl_SetMaxLen(dt, "EXBANKNAME", ExBankName)
        'Call TIMS.sUtl_SetMaxLen(dt, "Q3_OTHER", Q3_Other)
    End Sub

    '補列各資料選項 (排第1順位)(含本班所有學員下拉)
    Sub Add_Items()
        rqOCID = TIMS.ClearSQM(rqOCID)
        If rqOCID = "" Then Exit Sub
        'Dim sql As String
        'Dim dr As DataRow
        'Dim dt As DataTable
        'Dim sqlstr As String '取得該班開訓日期 sql
        'Dim sqlstr2 As String '顯示 03:負擔家計婦女 sql  
        'Dim sqlstr3 As String '顯示 03:負擔家計婦女 判斷
        'Dim STdate As Date

        '增加民族別選項
        'NativeID=TIMS.Get_KeyNative(NativeID)
        DegreeID = TIMS.Get_Degree(DegreeID, 1, objconn)
        GraduateStatus = TIMS.Get_GradState(GraduateStatus, objconn)
        graduatey = TIMS.GetSyear(graduatey, Year(Now) - 110, Year(Now), True) '畢業年份

        '列出兵役下拉選單資料-by Vicient
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
            If IdentityID.Items.Item(i).Value = "04" Then hide_IdentityID_04.Value = "IdentityID_" & i
            '06:身心障礙者
            If IdentityID.Items.Item(i).Value = "06" Then hide_IdentityID_06.Value = "IdentityID_" & i
            '都-有值-離開
            If hide_IdentityID_04.Value <> "" AndAlso hide_IdentityID_06.Value <> "" Then Exit For
        Next

        '生活津貼身分別 (SD_03_002_add.aspx)
        'SubsidyIdentity=TIMS.Get_SubsidyIdentity(SubsidyIdentity, 1, objconn)
        '參訓身分別
        'IdentityID.Attributes("onclick")="hard(" & IdentityID.ClientID & ")"
        IdentityID.Attributes.Add("onclick", "hard();")
        '身心障礙者
        rblHandType.Attributes.Add("onclick", "hard();")
        '---End
        '津貼類別 
        'SubsidyID=TIMS.Get_SubsidyID(SubsidyID)
        '身心障礙者
        HandTypeID = TIMS.Get_HandicatType(HandTypeID)
        HandLevelID = TIMS.Get_HandicatLevel(HandLevelID)
        HandTypeID2 = TIMS.Get_HandicatType2(HandTypeID2)
        HandLevelID2 = TIMS.Get_HandicatLevel2(HandLevelID2)
        'JoblessID=TIMS.Get_JoblessID(JoblessID, Nothing, sm.UserInfo.Years)

        'Dim dtDGHR As DataTable=Get_DGTHourDT() '目前系統有4筆資料。
        'RelClass_Unit=TIMS.Get_DGTHour(RelClass_Unit, dtDGHR)
        'Call TIMS.Get_Trade(Q4)

        '班別學號基本碼 
        StudentIDValue.Value = TIMS.Get_ClassStudentID(Me, rqOCID, objconn)
        If StudentIDValue.Value = "" Then
            Common.MessageBox(Me, "沒有班別學號基本碼")
        End If
        '== (兩週內)離退訓 可供遞補
        Dim pms1 As New Hashtable From {{"OCID", TIMS.CINT1(rqOCID)}}
        Dim sql1 As String = ""
        sql1 &= " SELECT a.StudentID ,b.Name+'('+dbo.FN_CSTUDID2(a.StudentID)+')' Name" & vbCrLf
        'sql1 &= " ,CASE WHEN LEN(a.StudentID)=12 THEN b.Name + '(' + RIGHT(a.StudentID,3) + ')' ELSE b.Name + '(' + RIGHT(a.StudentID,2) + ')' END Name" & vbCrLf
        sql1 &= " ,a.SOCID ,a.RejectDayIn14" & vbCrLf
        sql1 &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql1 &= " JOIN STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
        sql1 &= " WHERE a.OCID=@OCID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql1, objconn, pms1)
        '(兩週內)離退訓 可供遞補
        dt.DefaultView.RowFilter = "RejectDayIn14='Y'"
        dt.DefaultView.Sort = "StudentID"
        With RejectSOCID
            .DataSource = dt.DefaultView
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With
        '== (兩週內)離退訓 可供遞補

        Dim pms2 As New Hashtable From {{"OCID", TIMS.CINT1(rqOCID)}}
        Dim sql2 As String = ""
        sql2 &= " SELECT a.StudentID ,b.Name+'('+dbo.FN_CSTUDID2(a.StudentID)+')' Name" & vbCrLf
        sql2 &= " ,a.SOCID ,a.RejectDayIn14" & vbCrLf
        sql2 &= " FROM CLASS_STUDENTSOFCLASS a" & vbCrLf
        sql2 &= " JOIN STUD_STUDENTINFO b ON a.SID=b.SID" & vbCrLf
        sql2 &= " WHERE a.OCID=@OCID" & vbCrLf
        dt = DbAccess.GetDataTable(sql2, objconn, pms2)
        'dt.DefaultView.RowFilter="RejectDayIn14 IS NULL"
        dt.DefaultView.RowFilter = ""
        dt.DefaultView.Sort = "StudentID"

        With SOCID
            .DataSource = dt.DefaultView
            .DataTextField = "Name"
            .DataValueField = "SOCID"
            .DataBind()
            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End With

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

        'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0% '請選擇變為空。
        'SupplyID=TIMS.Get_SupplyID(SupplyID)

        Dim Sqlt As String = " SELECT SERVDEPTID ,SDNAME FROM KEY_SERVDEPT ORDER BY SERVDEPTID"
        Dim dtSERVDEPT As DataTable = DbAccess.GetDataTable(Sqlt, objconn)
        Dim Sqlj As String = " SELECT JOBTITLEID ,JTNAME FROM KEY_JOBTITLE ORDER BY JOBTITLEID"
        Dim dtJOBTITLE As DataTable = DbAccess.GetDataTable(Sqlj, objconn)
        ddlSERVDEPTID = TIMS.Get_SERVDEPTID(ddlSERVDEPTID, dtSERVDEPT)
        ddlJOBTITLEID = TIMS.Get_JOBTITLEID(ddlJOBTITLEID, dtJOBTITLE)
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
        'Dim sql As String = ""
        'sql &= " SELECT b.TPlanID ,d.ActNo" & vbCrLf
        'sql &= " FROM CLASS_CLASSINFO a" & vbCrLf
        'sql &= " JOIN ID_Plan b ON a.PlanID=b.PlanID" & vbCrLf
        'sql &= " JOIN Auth_Relship c ON a.RID=c.RID" & vbCrLf
        'sql &= " JOIN Org_OrgPlanInfo d ON c.RSID=d.RSID" & vbCrLf
        'sql &= $" WHERE a.OCID={rqOCID}" & vbCrLf
        'Dim drOP As DataRow = DbAccess.GetOneRow(sql, objconn)
        'If drOP Is Nothing Then
        '    Common.MessageBox(Me, cst_errMsg12)
        '    Exit Sub
        'End If

        ViewState(vs_STDate) = Common.FormatDate(drCC("STDate"))
        ClassName.Text = Convert.ToString(drCC("CLASSCNAME2")) 'TIMS.ClearSQM()
        'ClassName.Text=Convert.ToString(drCC("ClassCName")) 'TIMS.ClearSQM()
        'If Convert.ToString(dr("CyclType")) <> "" Then ClassName.Text &= "第" & Convert.ToString(dr("CyclType")) & "期"

        LevelNo.Items.Clear()
        If Convert.ToString(drCC("LevelCount")) <> "" Then
            If Int(drCC("LevelCount")) <> 0 Then
                For i As Integer = 1 To Int(drCC("LevelCount"))
                    LevelNo.Items.Add(New ListItem("第" & i & "階段", i))
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
        'If ActNo.Text="" Then ActNo.Text=TIMS.ChangeIDNO(dr("ActNo").ToString)
        Hid_OCID.Value = $"{drCC("OCID")}" 'a.OCID
        STDateHidden.Value = Common.FormatDate(drCC("STDate")) 'yyyy/MM/dd
        FTDateHidden.Value = Common.FormatDate(drCC("FTDate")) 'yyyy/MM/dd
        hide_THours.Value = $"{drCC("THours")}"

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

    ''' <summary>
    ''' 設定-預算別-序設值 
    ''' </summary>
    ''' <param name="dr"></param>
    Sub SETDEF_BudgetID(ByRef dr As DataRow, ByRef s_ActNo As String)
        Dim v_def_BudgetID As String = ""
        Dim t_msg As String = ""
        If Convert.ToString(dr("BudgetID")) <> "" Then
            v_def_BudgetID = Convert.ToString(dr("BudgetID"))
        Else
            'Convert.ToString(dr("ActNo")) 'STUD_STUDENTINFO
            'Convert.ToString(dr("ActNo2")) 'CLASS_STUDENTSOFCLASS
            If Convert.ToString(dr("ActNo2")) <> "" Then s_ActNo = Convert.ToString(dr("ActNo2"))

            'BudID (預算別) by AMU 20080602
            '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
            '01、04、05、15、08 其補助經費來源歸屬為 03:就保基金
            '02、03、06、07 其經費來源歸屬為 02:就安基金
            '09與無法辨視者為 99:不予補助對象
            '2.開頭數字為075、175（裁減續保）、076、176（職災續保）、09（訓）皆為不予補助對象，並設定阻擋。
            '03:就保基金'02:就安基金'99:不予補助對象
            Select Case Left(s_ActNo, 2) 'CLASS_STUDENTSOFCLASS
                Case "01", "04", "05", "15", "08"
                    v_def_BudgetID = "03"'03:就保基金
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
            t_msg = "(未儲存)"
        End If
        Common.SetListItem(BudID, v_def_BudgetID)
        If t_msg <> "" Then TIMS.Tooltip(BudID, t_msg, True)
    End Sub

    '塞入學員資料 [SQL]
    Sub Create1_Stud(ByVal SOCIDStr As String)
        'If SOCIDStr="" Then Exit Sub
        SOCIDStr = TIMS.ClearSQM(SOCIDStr)
        If SOCIDStr = "" Then Exit Sub

        'Dim dr As DataRow = Nothing 'STUD_STUDENTINFO,STUD_SUBDATA,CLASS_STUDENTSOFCLASS
        Dim drT2 As DataRow '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)


        Dim parms As New Hashtable() From {{"SOCID", TIMS.CINT1(SOCIDStr)}}
        'STUD_STUDENTINFO a
        Dim sql As String = ""
        sql &= " SELECT a.SID ,a.IDNO ,a.Name,a.EngName,a.RMPNAME,a.PassPortNO ,a.Sex ,a.Birthday" & vbCrLf
        sql &= " ,a.MaritalStatus ,a.DegreeID ,a.GraduateStatus ,a.MilitaryID ,a.IdentityID ,a.SubsidyID ,a.IsAgree" & vbCrLf
        sql &= " ,a.ChinaOrNot ,a.Nationality ,a.PPNO ,a.JobState ,a.ActNo ,a.GraduateY" & vbCrLf
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
            Page.RegisterStartupScript("SD_03_002", "<script>location.herf='SD_03_002.aspx?ID=" & rqMID & "';</script>")
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
        Dim sZIPCODE1 As String = Convert.ToString(dr("ZipCode1"))
        Dim iADID1 As Integer = 0
        Dim sZIPCODE2 As String = Convert.ToString(dr("ZipCode2"))
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
                    strMSG3 = "主要參訓身分別選擇「經公告之重大災害受災者」，須選擇「重大災害選項」不可為空！"
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

        '檢驗同一身分證號是否有兩筆以上的STUD_STUDENTINFO。
        Dim chkStud As Boolean = TIMS.Check_StudStudentInfo($"{dr("IDNO")}", objconn)
        LabErrMsg.Text = ""
        If Not chkStud Then
            'show laberrMsg
            Dim strErrmsg As String = ""
            strErrmsg &= Cst_errMsg2b & vbCrLf
            strErrmsg &= "班級OCID： " & Convert.ToString(dr("OCID")) & vbCrLf
            strErrmsg &= "班級名稱： " & ClassName.Text & vbCrLf
            strErrmsg &= "學員姓名： " & Convert.ToString(dr("Name")) & vbCrLf
            strErrmsg &= "學員出生年月日： " & Convert.ToString(dr("Birthday")) & vbCrLf
            strErrmsg &= "學員身分證號： " & Convert.ToString(dr("IDNO")) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            LabErrMsg.Text = Cst_errMsg2b
            Button1.Enabled = False '(儲存1)
            Button2.Enabled = False '(儲存2)
            TIMS.Tooltip(Button1, Cst_errMsg2, True)
            TIMS.Tooltip(Button2, Cst_errMsg2, True)
            Common.MessageBox(Me, Cst_errMsg2)
            Return
        End If

        'Dim rOCID As String = $"{dr("OCID")}"
        rqOCID = TIMS.ClearSQM(rqOCID)
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
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{dr("OCID")}", objconn)
        If rqOCID = "" OrElse drCC Is Nothing Then
            Button1.Enabled = False '(儲存1)
            Button2.Enabled = False '(儲存2)
            TIMS.Tooltip(Button1, cst_errMsg12, True)
            TIMS.Tooltip(Button2, cst_errMsg12, True)
            Common.MessageBox(Me, cst_errMsg12)
            Return
        End If

        'Dim vTitle As String="" 'vTitle="授權設定該班級有開放"
        If FTDateHidden.Value <> "" AndAlso TIMS.ChkIsEndDate(rqOCID, TIMS.cst_FunID_學員資料維護, dtArc) Then
            '28:產業人才投資計劃 54:充電起飛計畫（在職） '限分署(中心)
            If TIMS.Cst_TPlanID14DayCanEditStud.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID <= 1 Then
                '學員資料維護於訓後3個月不能修改 'Cst_Msg30x28
                If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, 3, CDate(FTDateHidden.Value)), Today) >= 0 Then
                    Button1.Enabled = False '(儲存1)
                    Button2.Enabled = False '(儲存2)
                    TIMS.Tooltip(Button1, Cst_Msg30x28, True)
                    TIMS.Tooltip(Button2, Cst_Msg30x28, True)
                    'Common.MessageBox(Me, Cst_Msg30x28)
                End If
            Else
                '針對委外職前訓練計畫，系統權限 限制
                Dim flgElseEvent1 As Boolean = True '其它狀況 預設為True
                'If Cst_TPlanIDCanEditStud_id37.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '    Select Case sm.UserInfo.LID
                '        Case "1" '分署(中心)
                '            flgElseEvent1 = False '已經設定狀況 其它@False
                '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                '                Button1.Enabled = False '(儲存1)
                '                Button2.Enabled = False '(儲存2)
                '                TIMS.Tooltip(Button1, Cst_MsgTPlanID37a, True)
                '                TIMS.Tooltip(Button2, Cst_MsgTPlanID37a, True)
                '            End If
                '        Case Else '"2" '委外
                '            flgElseEvent1 = False '已經設定狀況 其它@False
                '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                '                Button1.Enabled = False '(儲存1)
                '                Button2.Enabled = False '(儲存2)
                '                TIMS.Tooltip(Button1, Cst_MsgTPlanID37b, True)
                '                TIMS.Tooltip(Button2, Cst_MsgTPlanID37b, True)
                '            End If
                '    End Select
                'End If

                '其它狀況 預設為True
                If flgElseEvent1 Then
                    '學員資料維護於訓後30日不能修改'Cst_Msg30
                    If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 30, CDate(FTDateHidden.Value)), Today) >= 0 Then
                        Button1.Enabled = False '(儲存1)
                        Button2.Enabled = False '(儲存2)
                        TIMS.Tooltip(Button1, Cst_Msg30, True)
                        TIMS.Tooltip(Button2, Cst_Msg30, True)
                        'Turbo.Common.MessageBox(Me, Cst_Msg30)
                    End If
                End If
            End If
        End If

        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            Button1.Enabled = True '(儲存1)
            Button2.Enabled = True '(儲存2)
            TIMS.Tooltip(Button1, Cst_Msg1b)
            TIMS.Tooltip(Button2, Cst_Msg1b)
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
        If Convert.ToString(dr("MakeSOCID")) <> "" Then
            hide_MakeSOCID.Value = Convert.ToString(dr("MakeSOCID"))
            labmakesocid.Text = "遞補學員：" & TIMS.GetSOCIDName(Convert.ToString(dr("MakeSOCID")), objconn)
        End If

        '遞補學員(被遞補學員)
        'RejectSOCID.Enabled=True
        hide_RejectSOCID.Value = ""
        If Convert.ToString(dr("RejectSOCID")) <> "" Then
            hide_RejectSOCID.Value = Convert.ToString(dr("RejectSOCID"))
            Common.SetListItem(RejectSOCID, dr("RejectSOCID"))
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
        If Convert.ToString(dr("WorkSuppIdent")) <> "" Then Common.SetListItem(rblWorkSuppIdent, Convert.ToString(dr("WorkSuppIdent")))

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

        Dim ss As String = $"IDNO='{dr("IDNO")}'"
        '為在職者補助身份
        '該民眾不具失、待業身分，不得參加失業者職前訓練。'限定計畫執行
        'https://jira.turbotech.com.tw/browse/TIMSB-1247
        '僅涉托育人員及照顧服務員2支計畫 而非所有職前計畫
        LabWSImsg.Text = ""
        Dim flag_WSI As Boolean = False '(非)勾稽為在職者
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If Not dtBLIDET1 Is Nothing Then
                If dtBLIDET1.Select(ss).Length > 0 Then
                    '是否為在職者補助身分
                    flag_WSI = True '勾稽為在職者
                    Common.SetListItem(rblWorkSuppIdent, TIMS.cst_YES)
                    rblWorkSuppIdent.Enabled = False '鎖定
                    TIMS.Tooltip(rblWorkSuppIdent, cst_workman1)
                End If

                'NGACTNO:02:公會 x '03:漁會 x '06:農會 x (排除保險證號)
                ss = $"IDNO='{dr("IDNO")}' AND NGACTNO='Y'"
                If Not rblWorkSuppIdent.Enabled _
                    AndAlso dtBLIDET1.Select(ss).Length > 0 Then
                    flag_WSI = False '勾稽為非在職者
                    rblWorkSuppIdent.Enabled = True '解鎖
                    'rblWorkSuppIdent.Enabled=False '鎖定
                    Common.SetListItem(rblWorkSuppIdent, TIMS.cst_NO)
                    TIMS.Tooltip(rblWorkSuppIdent, cst_workman2, True)
                    'LabWSImsg.Text=cst_workman2
                End If

                '直接更動為在職者 (顯示時修正資料)
                'Call UPDATE_WorkSuppIdent(SOCIDStr, rOCID, flag_WSI, objconn)

                '勾稽為在職者
                If flag_WSI Then
                    Select Case Convert.ToString(dr("WorkSuppIdent"))
                        Case TIMS.cst_YES
                            'LabWSImsg.Text &= "目前是在職者"
                        Case Else
                            LabWSImsg.Text = "資料庫是「非在職者」"
                    End Select
                End If
            End If
        End If

        '是否為在職者補助身分
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            Select Case Convert.ToString(dr("WorkSuppIdent"))
                Case "Y"
                    Page.RegisterStartupScript("Change_MBTable1", "<script>Change_MBTable(1);</script>")
                Case "N"
                    Page.RegisterStartupScript("Change_MBTable2", "<script>Change_MBTable(2);</script>")
                Case Else
                    Page.RegisterStartupScript("Change_MBTable2", "<script>Change_MBTable(2);</script>")
            End Select
        End If

        Call GetHistorySumOfMoney(SOCIDStr)

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
        RMPNAME.Text = Convert.ToString(dr("RMPNAME"))
        If chkStud = False Then
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
        'for 個資保護
        txtShowIDNO.Text = TIMS.ChangeIDNO(dr("IDNO").ToString)

        Common.SetListItem(Sex, dr("Sex").ToString)

        If Convert.ToString(dr("Birthday")) <> "" Then
            Birthday.Text = Common.FormatDate(dr("Birthday"))
            'Birthday.Text=Common.FormatDate(Convert.ToString(dr("Birthday")), DateFormat.ShortDate)
            'ViewState("Birthday")=Common.FormatDate(dr("Birthday"))
            'for 個資保護
            'txtShowBirthday.Text=Common.FormatDate(dr("Birthday"))
        End If

        '有兩筆以上的STUD_STUDENTINFO,身分證號唯讀。
        If chkStud = False Then
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

        '判斷報名資料來源
        'Dim intTotal As Integer=Get_AdpTotal(SOCIDStr)
        'intTotal=Get_AdpTotal(SOCIDStr)

        '大於0=>表示從三合一報名, 等於0=>本系統報名
        'If intTotal > 0 Then
        '    '將推介單號存在hide給script備用
        '    hide_TrainMode.Value=TIMS.Get_GOVTRNData(Convert.ToString(dr("OCID")), IDNO.Text, objconn)
        '    If hide_TrainMode.Value <> "" Then
        '        '1.網;2.現;3.通;4.推
        '        Common.SetListItem(EnterChannel, "4")
        '        EnterChannel.Enabled=False
        '        TRNDMode.Enabled=False
        '        'TRNDType.Enabled=False
        '        TIMS.Tooltip(EnterChannel, "三合一報名資料回傳")
        '        TIMS.Tooltip(TRNDMode, "三合一報名資料回傳")
        '        'TIMS.Tooltip(TRNDType, "三合一報名資料回傳")
        '    End If
        '    Dim v_EnterChannel As String=TIMS.GetListValue(EnterChannel)
        '    If v_EnterChannel="4" Then
        '        EnterChannel.Enabled=False
        '        TRNDMode.Enabled=False
        '        TIMS.Tooltip(EnterChannel, "該學員無法從推介更改為其他報名管道")
        '        TIMS.Tooltip(TRNDMode, "該學員無法從推介更改為其他報名管道")
        '    End If

        '    TRNDTR.Style.Item("display")=cst_none1 'cst_inline1
        '    GovTR.Style.Item("display")=cst_none1 'cst_inline1
        '    'https://jira.turbotech.com.tw/browse/TIMSC-272
        '    Select Case Convert.ToString(dr("TRNDMode")) '1.職2.學3.推
        '        Case "1", "2"
        '            '資料異常修改。
        '            TRNDTR.Style.Item("display")=cst_inline1
        '            GovTR.Style.Item("display")=cst_inline1
        '            'Call GetGovIdent(SOCIDStr)
        '            'Page.RegisterStartupScript("TRNDModeChange", "<script>TRNDModeChange();</script>")
        '    End Select

        '    Select Case Convert.ToString(dr("TRNDMode")) '1.職2.學3.推
        '        Case "1"
        '            'https://jira.turbotech.com.tw/browse/TIMSC-272
        '            'Common.SetListItem(TRNDType, dr("TRNDType").ToString)
        '        Case "2"
        '            'https://jira.turbotech.com.tw/browse/TIMSC-272
        '            'DGTR.Style.Item("display")=cst_inline1
        '            'Call GetDGIdent(SOCIDStr)
        '            'Page.RegisterStartupScript("TRNDModeChange", "<script>TRNDModeChange();</script>")
        '        Case "3"
        '            TRNDTR.Style.Item("display")=cst_inline1
        '            GovTR.Style.Item("display")=cst_inline1
        '            'Call GetGovIdent(SOCIDStr)
        '            'Page.RegisterStartupScript("TRNDModeChange", "<script>TRNDModeChange();</script>")
        '    End Select
        'End If
        'TRNDMode推介種類(1.職2.學3.推)
        'Select Case Convert.ToString(dr("TRNDMode"))
        '    Case "1"
        '        Common.SetListItem(TRNDType, dr("TRNDType").ToString)
        '    Case Else
        '        Page.RegisterStartupScript("TRNDModeChange", "<script>TRNDModeChange();</script>")
        'End Select

        If $"{dr("OpenDate")}" <> "" Then OpenDate.Text = Common.FormatDate($"{dr("OpenDate")}")
        If $"{dr("CloseDate")}" <> "" Then CloseDate.Text = Common.FormatDate($"{dr("CloseDate")}")
        If $"{dr("EnterDate")}" <> "" Then EnterDate.Text = Common.FormatDate($"{dr("EnterDate")}")

        'Common.SetListItem(DegreeID, dr("DegreeID").ToString)
        Dim DegreeIDValTmp As String = ""
        '修正學歷代碼 (學員資料維護)
        DegreeIDValTmp = TIMS.Fix_DegreeValue($"{dr("DegreeID")}")
        Common.SetListItem(DegreeID, DegreeIDValTmp)

        'Dim v_BudID As String=""
        ''Dim v_SupplyID As String=""
        ''BudID (預算別) 
        'If Hid_show_actno_budid.Value="Y" Then
        '    '產投 設定預算別 create
        'End If

        Dim strIDNO As String = TIMS.ClearSQM($"{dr("IDNO")}")
        Dim s_ACTNAME As String = ""
        Dim s_ACTNO As String = ""
        'BEGINCLASS,Y,(過開訓日)
        If Convert.ToString(drCC("BEGINCLASS")).Equals("Y") Then
            '過開訓日才可使用勾稽資料
            's_ACTNAME=TIMS.GET_BLIGATEDATA28E(strIDNO, rOCID, objconn, "ACTNAME") 'STUD_BLIGATEDATA28E
            's_ACTNO=TIMS.GET_BLIGATEDATA28E(strIDNO, rOCID, objconn, "ACTNO") 'STUD_BLIGATEDATA28E
            's_ACTNAME=TIMS.GET_BLIGATEDATA28(SOCIDStr, strIDNO, objconn, "ACTNAME") 'STUD_BLIGATEDATA28
            's_ACTNO=TIMS.GET_BLIGATEDATA28(SOCIDStr, strIDNO, objconn, "ACTNO") 'STUD_BLIGATEDATA28
            s_ACTNAME = TIMS.GET_SELRESULTBLI(strIDNO, rqOCID, objconn, "ACTNAME") 'STUD_SELRESULTBLI
            s_ACTNO = TIMS.GET_SELRESULTBLI(strIDNO, rqOCID, objconn, "ACTNO") 'STUD_SELRESULTBLI
        End If
        Call SETDEF_BudgetID(dr, s_ACTNO)

        School.Text = $"{dr("School")}"
        Department.Text = $"{dr("Department")}"
        If $"{dr("GraduateStatus")}" <> "" Then Common.SetListItem(GraduateStatus, dr("GraduateStatus"))
        If $"{dr("GraduateY")}" <> "" Then Common.SetListItem(graduatey, dr("GraduateY"))

        SolTR.Style.Item("display") = cst_none1
        'MilitaryID.SelectedIndex=-1
        If $"{dr("MilitaryID")}" <> "" Then
            Common.SetListItem(MilitaryID, $"{dr("MilitaryID")}")
            If $"{dr("MilitaryID")}" = "04" Then SolTR.Style.Item("display") = cst_inline1
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
        If Convert.ToString(dr("ZipCode1")) = Convert.ToString(dr("ZipCode2")) And Convert.ToString(dr("Address")) = Convert.ToString(dr("HouseholdAddress")) And Convert.ToString(dr("ZipCode1_6W")) = Convert.ToString(dr("ZipCode2_6W")) Then CheckBox1.Checked = True

        '判斷緊急聯絡人地址是否有與通訊地址或戶籍地址相同
        If Convert.ToString(dr("ZipCode1")) = Convert.ToString(dr("ZipCode3")) And Convert.ToString(dr("Address")) = Convert.ToString(dr("EmergencyAddress")) And Convert.ToString(dr("ZipCode1_6W")) = Convert.ToString(dr("ZipCode3_6W")) Then
            CheckBox2.Checked = True
        Else
            If Convert.ToString(dr("ZipCode2")) = Convert.ToString(dr("ZipCode3")) And Convert.ToString(dr("HouseholdAddress")) = Convert.ToString(dr("EmergencyAddress")) And Convert.ToString(dr("ZipCode2_6W")) = Convert.ToString(dr("ZipCode3_6W")) Then CheckBox3.Checked = True
        End If

        Email.Text = Convert.ToString(dr("Email"))

        'Common.SetListItem(SubsidyID, dr("SubsidyIDEX").ToString)
        'If Convert.ToString(dr("SubsidyIdentity")) <> "" Then Common.SetListItem(SubsidyIdentity, dr("SubsidyIdentity").ToString)
        'If Convert.ToString(dr("SubsidyIDEX"))="03" Then
        '    SubsidyHidden.Value="1"
        'Else
        '    SubsidyHidden.Value="0"
        'End If
        'SubsidyID.Attributes("onchange")="ChangeSubsidy();"

        Common.SetListItem(MIdentityID, Convert.ToString(dr("MIdentityID")))
        hide_MIdentityID.Value = TIMS.GetListValue(MIdentityID) '.SelectedValue

        If tr_DDL_DISASTER.Visible Then
            If Convert.ToString(dr("ADID")) <> "" Then
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
        'If Chk_Sub_SubSidyApply(SOCIDStr) Then
        '    'SubsidyID.Enabled=False '鎖定
        '    SubsidyIdentity.Enabled=False '鎖定
        '    'MIdentityID.Enabled=False '鎖定
        '    'rdo_HighEduBg.Enabled=False '鎖定
        '    rdo_HighEduBg.Attributes.Add("disabled", "disabled")  '專上畢業學歷失業者

        '    '有職訓生活津貼，不可修改姓名 by AMU 2009-09-14
        '    Name.Enabled=False '鎖定
        '    'Name.ToolTip=Cst_Msg1
        '    TIMS.Tooltip(Name, Cst_Msg1)

        '    '有職訓生活津貼，不可修改身分證號碼,生日資料
        '    'Birthday.ReadOnly=True
        '    Birthday.Attributes.Add("onkeydown", "this.blur()")
        '    Birthday.Attributes.Add("oncontextmenu", "return false;")
        '    Birthday.Enabled=False
        '    Img1.Style("display")=cst_none1 '出生日期選擇功能
        '    Img1.Disabled=True
        '    'hidBirthBtn.Disabled=True '失效
        '    TIMS.Tooltip(Birthday, Cst_Msg1)

        '    IDNO.Enabled=False
        '    DegreeID.Enabled=False
        '    'Birthday.ToolTip=Cst_Msg1
        '    'IDNO.ToolTip=Cst_Msg1
        '    'SubsidyID.ToolTip=Cst_Msg1
        '    'SubsidyIdentity.ToolTip=Cst_Msg1
        '    TIMS.Tooltip(Birthday, Cst_Msg1)
        '    TIMS.Tooltip(IDNO, Cst_Msg1)
        '    TIMS.Tooltip(DegreeID, Cst_Msg1)
        '    'TIMS.Tooltip(SubsidyID, Cst_Msg1)
        '    'TIMS.Tooltip(SubsidyIdentity, Cst_Msg1)

        '    If sm.UserInfo.RoleID <= 1 Then '系統管理者開放修改
        '        DegreeID.Enabled=True
        '        TIMS.Tooltip(DegreeID, Cst_Msg1b)
        '    End If
        '    'ViewState("MsgBox")=Cst_Msg1 & vbCrLf
        'End If

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
                    Dim sMsg1 As String = Cst_Msg21 '(21日鎖定)
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
            '    '已勾選 屆退官兵
            '    If Convert.ToString(dr("IdentityIDEX")).IndexOf(cst_id12) > -1 Then flag_ChkID12=True '已勾選 屆退官兵

            '    If Not flag_ChkID12 Then '未勾選 屆退官兵
            '        Dim flag_SRSOLDIERS As Boolean=False '是否為屆退官兵
            '        flag_SRSOLDIERS=TIMS.CheckRESOLDER(objconn, dr("IDNO"), sm.UserInfo.DistID, ViewState(vs_STDate))
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
            Dim sIdentityIDEX As String = "" '集合目前所勾選的身分別存入  ViewState(vs_IdentityID)
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
        Dim flag_HandType As Integer = 0 '0:未選 1:舊制 2:新制
        If Convert.ToString(dr("HandTypeID")) <> "" AndAlso Convert.ToString(dr("HandLevelID")) <> "" Then
            flag_HandType = 1 '1:舊制
        End If
        If Convert.ToString(dr("HandTypeID2")) <> "" AndAlso Convert.ToString(dr("HandLevelID2")) <> "" Then
            flag_HandType = 2 '2:新制
        End If

        Select Case flag_HandType
            Case 1 '1:舊制
                trHandTypeID2.Style("display") = cst_none1 '新制
                trHandTypeID.Style("display") = cst_inline1 '舊制
                Common.SetListItem(rblHandType, "1")
            Case Else '0:未選 2:新制
                trHandTypeID2.Style("display") = cst_inline1 '新制
                trHandTypeID.Style("display") = cst_none1 '舊制
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

        ' sart受訓前任職清單2011/04/27 先讀CLASS_STUDENTSOFCLASS 若讀不到讀STUD_SUBDATA 
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
        ' end受訓前任職清單2011/04/27 先讀CLASS_STUDENTSOFCLASS 若讀不到讀STUD_SUBDATA 

        Dim v_BudID As String = TIMS.GetListValue(BudID) '.SelectedValue
        'v_BudID=TIMS.GetListValue(BudID) '.SelectedValue
        If v_BudID = "" AndAlso Convert.ToString(dr("BudgetID")) = "" Then
            '尚未選擇(且空白)帶預設
            '" & cst_ECFA & "預設測試  
            '=該計畫是否使用ECFA
            If blnTPlanUseEcfa Then
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If DateDiff(DateInterval.Day, CDate(Cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
                        If TIMS.CheckIsECFA(Me, ActNo2.Text, "", STDateHidden.Value, objconn) = True Then Common.SetListItem(BudID, "97")  '2011/05/20 新增ECFA判斷
                    End If
                Else
                    If TIMS.CheckIsECFA(Me, ActNo2.Text, FOfficeYM1.Text, "", objconn) = True Then Common.SetListItem(BudID, "97")  '2011/05/20 新增ECFA判斷
                End If
            End If
        Else
            '有值
            If Convert.ToString(dr("BudgetID")) <> "" Then Common.SetListItem(BudID, Convert.ToString(dr("BudgetID")))
        End If

        Common.SetListItem(Traffic, Convert.ToString(dr("Traffic")))
        Common.SetListItem(ShowDetail, dr("ShowDetail").ToString)

        If Hid_show_actno_budid.Value = "Y" Then
            '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
            'sql="" & vbCrLf 'sql &= " select c.*" & vbCrLf 'Dim sql As String=""

            Dim pmsT2 As New Hashtable From {{"IDNO", $"{dr("IDNO")}"}, {"OCID1", TIMS.CINT1(dr("OCID"))}}
            sql = ""
            sql &= " SELECT c.SEID,c.ESERNUM,c.MIDENTITYID,c.HANDTYPEID,c.HANDLEVELID" & vbCrLf
            sql &= " ,c.PRIORWORKORG1,c.TITLE1,c.PRIORWORKORG2,c.TITLE2,c.SOFFICEYM1,c.FOFFICEYM1,c.SOFFICEYM2,c.FOFFICEYM2,c.PRIORWORKPAY,c.REALJOBLESS" & vbCrLf
            sql &= " ,c.JOBLESSID,c.TRAFFIC,c.SHOWDETAIL,c.ACCTMODE,c.POSTNO,c.ACCTHEADNO,c.BANKNAME,c.ACCTEXNO,c.EXBANKNAME,c.ACCTNO,c.FIRDATE,c.UNAME,c.INTAXNO" & vbCrLf
            sql &= " ,c.ACTNO,c.ACTNAME" & vbCrLf
            sql &= " ,c.SERVDEPT ,c.JOBTITLE" & vbCrLf
            sql &= " ,c.ZIP,c.ZIP6W,c.ZIP_N" & vbCrLf
            sql &= " ,c.ADDR,c.TEL,c.FAX,c.SDATE,c.SJDATE,c.SPDATE,c.Q1,c.Q2_1,c.Q2_2,c.Q2_3,c.Q2_4,c.Q3,c.Q3_OTHER,c.Q4,c.Q5,c.Q61,c.Q62,c.Q63,c.Q64,c.ISEMAIL" & vbCrLf
            'sql &= " ,c.MODIFYACCT,c.MODIFYDATE" & vbCrLf
            sql &= " ,c.ACTTYPE,c.SCALE,c.ZIPCODE2,c.ZIPCODE2_6W,c.ZIPCODE2_N,c.HOUSEHOLDADDRESS,c.ACTTEL,c.ZIPCODE3,c.ZIPCODE3_6W,c.ZIPCODE3_N,c.ACTADDRESS,c.INSURED,c.SERVDEPTID,c.JOBTITLEID" & vbCrLf
            sql &= " ,c.ZipCode3 ActZipCode" & vbCrLf
            sql &= " ,c.ZipCode3_6W ActZipCode_6W" & vbCrLf
            sql &= " ,c.ZipCode3_N ActZipCode_N" & vbCrLf
            sql &= " FROM dbo.STUD_ENTERTEMP2 a" & vbCrLf
            sql &= " JOIN dbo.STUD_ENTERTYPE2 b ON a.esetid=b.esetid" & vbCrLf
            sql &= " LEFT JOIN dbo.STUD_ENTERTRAIN2 c ON c.eSerNum=b.eSerNum" & vbCrLf
            sql &= " WHERE a.IDNO=@IDNO AND b.OCID1=@OCID1" & vbCrLf
            drT2 = DbAccess.GetOneRow(sql, objconn, pmsT2) 'STUD_ENTERTRAIN2

            '學員服務單位(產學訓)
            Dim pmsSP As New Hashtable From {{"SOCID", TIMS.CINT1(SOCIDStr)}}
            Dim sql_sp As String = ""
            sql_sp &= " SELECT uname ,Intaxno ,ActName ,ActNo ,ServDept ,JobTitle" & vbCrLf
            sql_sp &= " ,SERVDEPTID ,JOBTITLEID ,ActTel ,ActZipCode ,ActZipCode_6W,ActZipCode_N,ActAddress" & vbCrLf
            sql_sp &= " FROM dbo.STUD_SERVICEPLACE WHERE SOCID=@SOCID "
            dr = DbAccess.GetOneRow(sql_sp, objconn, pmsSP) 'STUD_SERVICEPLACE

            If dr IsNot Nothing Then
                'STUD_SERVICEPLACE
                '為勞工團體時，會多一個訓練單位代轉現金的選項，所以增加=2的Flag判斷
                'AcctMode 0:郵政1:金融(銀行)2:訓練單位代轉現金
                Uname.Text = dr("Uname").ToString
                Intaxno.Text = dr("Intaxno").ToString
                '如果有儲存資料就使用儲存資料 '預設使用勾稽資料
                ActName.Text = $"{dr("ActName")}" 'STUD_SERVICEPLACE,,null
                ActNo1.Text = $"{dr("ActNo")}" 'STUD_SERVICEPLACE,,null
                '最後存取值
                'ActType: 投保類別1.勞2.農3.漁
                'If Convert.ToString(dr("ActType")) <> "" Then Common.SetListItem(ActType, dr("ActType").ToString)
                ServDept.Text = Convert.ToString(dr("ServDept"))
                JobTitle.Text = Convert.ToString(dr("JobTitle"))
                If Convert.ToString(dr("SERVDEPTID")) <> "" Then Common.SetListItem(ddlSERVDEPTID, dr("SERVDEPTID"))
                If Convert.ToString(dr("JOBTITLEID")) <> "" Then Common.SetListItem(ddlJOBTITLEID, dr("JOBTITLEID"))
                txt_ActPhone.Text = Convert.ToString(dr("ActTel"))  '加入投保單位電話、地址
                'ActZipCode  
                txt_ActZip.Value = Convert.ToString(dr("ActZipCode"))
                hid_ActZIP6W.Value = Convert.ToString(dr("ActZipCode_6W"))
                txt_ActZIPB3.Value = TIMS.GetZIPCODEB3(hid_ActZIP6W.Value)
                hidActZip_N.Value = Convert.ToString(dr("ActZipCode_N"))
                txt_ActCity.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ActZipCode")), Convert.ToString(dr("ActZipCode_N")), objconn)
                txt_ActAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ActAddress")))
                Hid_JnActZip.Value = TIMS.GetZipCodeJn(txt_ActZip.Value, txt_ActZIPB3.Value, hid_ActZIP6W.Value, txt_ActCity.Text, txt_ActAddress.Text)

            ElseIf drT2 IsNot Nothing Then
                '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
                dr = drT2
                'If dr Is Nothing Then dr=drT2
                Uname.Text = dr("Uname").ToString
                Intaxno.Text = dr("Intaxno").ToString
                '如果有儲存資料就使用儲存資料 '預設使用勾稽資料
                'ActName.Text=Convert.ToString(dr("ActName")) 'STUD_ENTERTRAIN2
                'ActNo1.Text=Convert.ToString(dr("ActNo")) 'STUD_ENTERTRAIN2
                '如果有儲存資料就使用儲存資料 '預設使用勾稽資料
                ActName.Text = s_ACTNAME 'STUD_SELRESULTBLI
                ActNo1.Text = s_ACTNO 'STUD_SELRESULTBLI
                '最後存取值
                'ActType: 投保類別1.勞2.農3.漁
                'If Convert.ToString(dr("ActType")) <> "" Then Common.SetListItem(ActType, dr("ActType").ToString)
                ServDept.Text = Convert.ToString(dr("ServDept"))
                JobTitle.Text = Convert.ToString(dr("JobTitle"))
                If Convert.ToString(dr("SERVDEPTID")) <> "" Then Common.SetListItem(ddlSERVDEPTID, dr("SERVDEPTID"))
                If Convert.ToString(dr("JOBTITLEID")) <> "" Then Common.SetListItem(ddlJOBTITLEID, dr("JOBTITLEID"))
                '加入投保單位電話、地址
                txt_ActPhone.Text = Convert.ToString(dr("ActTel"))

                '加入投保單位電話、地址'ActZipCode  
                txt_ActZip.Value = Convert.ToString(dr("ActZipCode"))
                hid_ActZIP6W.Value = Convert.ToString(dr("ActZipCode_6W"))
                txt_ActZIPB3.Value = TIMS.GetZIPCODEB3(hid_ActZIP6W.Value)
                hidActZip_N.Value = Convert.ToString(dr("ActZipCode_N"))
                txt_ActCity.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ActZipCode")), Convert.ToString(dr("ActZipCode_N")), objconn)
                txt_ActAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ActAddress")))
                Hid_JnActZip.Value = TIMS.GetZipCodeJn(txt_ActZip.Value, txt_ActZIPB3.Value, hid_ActZIP6W.Value, txt_ActCity.Text, txt_ActAddress.Text)

            End If

            '相同時 Checkbox4 打勾
            'If Convert.ToString(txt_ActZip.Value)=Convert.ToString(Zip.Value) And Convert.ToString(txt_ActZIP6W.Value)=Convert.ToString(ZIP6W.Value) And Convert.ToString(txt_ActAddress.Text)=Convert.ToString(Addr.Text) Then Checkbox4.Checked=True
            '學員參訓背景(產學訓)
            'sql=" SELECT * FROM STUD_TRAINBG WHERE SOCID=@SOCID "
            'parms.Clear()
            'parms.Add("SOCID", SOCIDStr)
            'dr=DbAccess.GetOneRow(sql, objconn, parms)
            '試著取得 STUD_ENTERTRAIN2 :線上報名資料(產學訓)
            'If dr Is Nothing Then dr=drT2
        End If

        '系統先去比對送訓官兵名冊，(比照參訓學員投保狀況檢核表)
        '如該學員為現役軍人：
        '1.【投保單位保險證號】預帶「在役軍人」
        '2.【預算別】預帶「就安」
        '3.【投保單位名稱】預帶送訓官兵名冊之「任職單位」
        '4. 儲存時，增加檢核【投保單位保險證號】為「在役軍人」，【預算別】應為「就安」
        Hid_out_POSITION.Value = ""
        If flagTPlanID06Plan3 Then
            Dim out_POSITION As String = ""
            Dim flag_SRSOLDIERS As Boolean = False '是否為屆退官兵
            'sm.UserInfo.DistID / Convert.ToString(drCC("DISTID"))
            If (ActNo1.Text = "") Then flag_SRSOLDIERS = TIMS.CheckRESOLDER(objconn, IDNO.Text, Convert.ToString(drCC("DISTID")), STDateHidden.Value, out_POSITION)
            If (flag_SRSOLDIERS) Then
                Hid_out_POSITION.Value = out_POSITION
                ActName.Text = out_POSITION '【投保單位名稱】預帶送訓官兵名冊之「任職單位」
                ActNo1.Text = cst_Serviceman '在役軍人
                '03:就保基金'02:就安基金'99:不予補助對象
                Common.SetListItem(BudID, "02")
                'ActName.Enabled=False
                ActNo1.Enabled = False
                'BudID.Enabled=False
                'TIMS.Tooltip(ActName, cst_Serviceman)
                TIMS.Tooltip(ActNo1, cst_Serviceman)
                'TIMS.Tooltip(BudID, cst_Serviceman)
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

        WSITR.Visible = False
        WSITR2.Visible = False
        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '含職前webservice
        '    SubsidyCost.Text=TIMS.Get_SubsidyCost(IDNO.Text, STDateHidden.Value, "", "Y", objconn)
        '    WSITR.Visible=True
        '    WSITR2.Visible=True
        'End If
    End Sub

#Region "NO USE"
    '取得 Adp_DGTRNData
    'Sub createDG(ByVal TICKET_NO As String)
    '    If TICKET_NO="" Then Return
    '    TICKET_NO=TIMS.ClearSQM(TICKET_NO)
    '    If TICKET_NO="" Then Return

    '    Dim sql As String=""
    '    sql &= " SELECT a.IDNO" & vbCrLf
    '    sql &= " ,b.Name ,b.Sex ,b.Birth ,b.Marri ,b.Edgr ,b.Gradu ,b.School ,b.DeptName ,b.Solder" & vbCrLf
    '    sql &= " ,b.Addr_Zip" & vbCrLf
    '    sql &= " ,b.Addr" & vbCrLf
    '    sql &= " ,b.Tel ,b.Mobile ,b.Email ,c.Share_Name" & vbCrLf
    '    sql &= " FROM ADP_DGTRNDATA a" & vbCrLf
    '    sql &= " LEFT JOIN ADP_STDDATA b ON a.IDNO=b.IDNO" & vbCrLf
    '    sql &= " LEFT JOIN ADP_SHARESOURCE c ON a.OBJECT_TYPE=c.Share_ID AND c.Share_Type='301'" & vbCrLf
    '    sql &= " WHERE a.TICKET_NO='" & TICKET_NO & "'" & vbCrLf
    '    Dim dr As DataRow=DbAccess.GetOneRow(sql, objconn)
    '    If dr Is Nothing Then Return

    '    IDNO.Text=TIMS.ChangeIDNO(dr("IDNO").ToString)
    '    Name.Text=dr("Name").ToString
    '    Select Case Convert.ToString(dr("Sex"))
    '        Case "1", "M"
    '            Common.SetListItem(Sex, "M")
    '        Case "2", "F"
    '            Common.SetListItem(Sex, "F")
    '    End Select
    '    If Convert.ToString(dr("Birth")) <> "" Then Birthday.Text=Common.FormatDate(dr("Birth"))

    '    Select Case Convert.ToString(dr("Marri"))
    '        Case "1", "2"
    '            Common.SetListItem(MaritalStatus, dr("Marri").ToString)
    '        Case Else
    '            Common.SetListItem(MaritalStatus, "3")
    '    End Select

    '    If Convert.ToString(dr("Edgr")) <> "" Then
    '        Dim DegreeIDValTmp As String=""
    '        '修正學歷代碼 (學員資料維護)
    '        DegreeIDValTmp=TIMS.Fix_DegreeValue(Convert.ToString(dr("Edgr")))
    '        Common.SetListItem(DegreeID, DegreeIDValTmp)
    '    End If

    '    If Convert.ToString(dr("Gradu")) <> "" Then Common.SetListItem(GraduateStatus, dr("Gradu"))
    '    School.Text=dr("School").ToString
    '    Department.Text=dr("DeptName").ToString
    '    'MilitaryID@Solder
    '    If Convert.ToString(dr("Solder")) <> "" Then Common.SetListItem(MilitaryID, dr("Solder").ToString)

    '    ' Addr_Zip (三合一學員基本資料檔)  郵遞區號 
    '    tZipLName=TIMS.Get_ZipLName(Convert.ToString(dr("Addr_Zip")), objconn)
    '    ZipCode1.Value=Convert.ToString(dr("Addr_Zip"))
    '    ZipCode1_6W.Value=""
    '    ZipCode1_N.Value=""
    '    City1.Text=TIMS.Get_ZipNameN(Convert.ToString(dr("Addr_Zip")), "", objconn)
    '    If tZipLName <> "" Then City1.Text &= "[" & tZipLName & "]"
    '    Address.Text=HttpUtility.HtmlDecode(Convert.ToString(dr("Addr")))
    '    PhoneD.Text=dr("Tel").ToString
    '    CellPhone.Text=Convert.ToString(dr("Mobile"))
    '    If CellPhone.Text <> "" Then CellPhone.Text=Trim(CellPhone.Text)
    '    If CellPhone.Text <> "" Then
    '        Common.SetListItem(rblMobil, "Y")
    '    Else
    '        Common.SetListItem(rblMobil, "N")
    '    End If
    '    Email.Text=dr("Email").ToString
    '    'DGIdentValue.Text=dr("Share_Name").ToString
    '    'DGTR.Style.Item("display")=cst_inline1
    '    'MilitaryID
    '    'Page.RegisterStartupScript("sol", "<script>sol(" & v_MilitaryID & ");</script>")
    '    Dim v_MilitaryID As String=TIMS.GetListValue(MilitaryID) '.SelectedValue
    '    If v_MilitaryID="04" Then SolTR.Style.Item("display")=cst_inline1
    'End Sub
#End Region

    '若有此SOCID 則為真， 其它情況則為 否
    'Function Chk_Sub_SubSidyApply(ByVal SOCID As String) As Boolean
    '    Dim rst As Boolean=False
    '    SOCID=TIMS.ClearSQM(SOCID)
    '    If SOCID="" Then Return rst

    '    Dim dr As DataRow
    '    Dim sql As String
    '    sql=" SELECT 'x' FROM SUB_SUBSIDYAPPLY WHERE SOCID='" & SOCID & "' "
    '    dr=DbAccess.GetOneRow(sql, objconn)
    '    If dr IsNot Nothing Then rst=True
    '    Return rst
    'End Function

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

        Dim sql As String
        sql = " SELECT DISTID FROM AUTH_RELSHIP WHERE RID='" & vRIDValue & "'"
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr1 Is Nothing Then Return
        DistValue.Value = Convert.ToString(dr1("DistID"))
    End Sub

    '產學訓專用 學員限制 (Enabled /Disabled) true:鎖定 false:開放
    Sub FacLimit(ByVal Flag As Boolean, ByVal sTip As String)
        If Flag Then
            '鎖定
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
            'If SupplyID.SelectedIndex <> -1 Then SupplyID.Enabled=False Else SupplyID.Enabled=True

            'Name.ToolTip=Cst_Msg2 'BudID.ToolTip.Empty
            'LName.ToolTip=Cst_Msg2 ' LName.ToolTip.Empty
            'FName.ToolTip=Cst_Msg2 'FName.ToolTip.Empty
            'IDNO.ToolTip=Cst_Msg2 'IDNO.ToolTip.Empty
            'Birthday.ToolTip=Cst_Msg2 'Birthday.ToolTip.Empty
            'BudID.Attributes("title")=Cst_Msg2
            'MIdentityID.Attributes("title")=Cst_Msg2
            'IdentityID.Attributes("title")=Cst_Msg2
            'SupplyID.Attributes("title")=Cst_Msg2
            'SupplyID.Attributes.Remove("title")
            'BudID.ToolTip=Cst_Msg2
            'MIdentityID.ToolTip=Cst_Msg2
            'IdentityID.ToolTip=Cst_Msg2
            'SupplyID.ToolTip=Cst_Msg2

            TIMS.Tooltip(Name, Cst_Msg2)
            TIMS.Tooltip(LName, Cst_Msg2)
            TIMS.Tooltip(FName, Cst_Msg2)
            TIMS.Tooltip(IDNO, Cst_Msg2)
            TIMS.Tooltip(Birthday, Cst_Msg2)
            TIMS.Tooltip(MIdentityID, Cst_Msg2, True)
            TIMS.Tooltip(IdentityID, Cst_Msg2, True)
            TIMS.Tooltip(BudID, Cst_Msg2, True)
            'TIMS.Tooltip(SupplyID, Cst_Msg2, True)
        Else
            RejectSOCID.Enabled = True '遞補者
            Name.Enabled = True
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
            'SupplyID.Enabled=True
            Name.ToolTip = String.Empty
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
            TIMS.Tooltip(Name, sTip)
            TIMS.Tooltip(RMPNAME, sTip)
            TIMS.Tooltip(LName, sTip)
            TIMS.Tooltip(FName, sTip)
            TIMS.Tooltip(IDNO, sTip)
            'TIMS.Tooltip(Birthday, sTip)
            TIMS.Tooltip(BudID, sTip)
            TIMS.Tooltip(MIdentityID, sTip)
            TIMS.Tooltip(IdentityID, sTip)
            'TIMS.Tooltip(SupplyID, sTip)
        End If

        '產投相關計畫執行此功能。
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If Hid_show_actno_budid.Value = "Y" Then
            flag_BudIDNoLock = False '如果是分署(中心)承辦人，預算別不鎖定。by AMU 20140328 (本功能是每次執行)
            flag_BudIDNoLock = Chk_CanEditBudgetID(sm.UserInfo.LID, CDate(STDateHidden.Value), iBudFlag) '什麼時候可以修改預算別。
            If flag_BudIDNoLock Then
                If Not BudID.Enabled Then
                    BudID.Enabled = True
                    TIMS.Tooltip(BudID, Cst_Msg21ok)
                End If
            End If
        End If
    End Sub

    '學員資料審核功能的欄位鎖住
    Sub GetScript(ByVal SOCID As String)
        Dim sql As String = ""
        Dim dt As DataTable = Nothing

        'start 學員資料審核功能的欄位鎖住,若按確認為鎖住
        sql = " SELECT STUDENTID ,ISAPPRPAPER ,APPLIEDRESULT FROM CLASS_STUDENTSOFCLASS WHERE SOCID='" & SOCID & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            Dim dr As DataRow = Nothing
            dr = dt.Rows(0)
            If dr("IsApprPaper").ToString = "Y" Then '1.假如學員資料確定就鎖住某些欄位
                If dr("AppliedResult").ToString = "Y" Then '2.假如學員資料審核通過就鎖住某些欄位
                    '如果是系統管理者開啟功能。
                    If TIMS.IsSuperUser(Me, 1) Then
                        'ROLEID=0 LID=0
                        Call FacLimit(False, Cst_Msg1b)
                    Else
                        '其他使用者鎖定。
                        Call FacLimit(True, Cst_Msg2)
                    End If
                End If
            End If
        End If
        'End 學員資料審核功能的欄位鎖住,若按確認為鎖住

        Dim sStudIDtmps As String = ""
        For i As Integer = 0 To dt.Rows.Count - 1
            If Len(dt.Rows(i).Item("StudentID")) = 12 Then
                If sStudIDtmps <> "" Then sStudIDtmps &= ","
                sStudIDtmps &= "'" & Right(dt.Rows(i).Item("StudentID"), 3) & "'"
            Else
                If sStudIDtmps <> "" Then sStudIDtmps &= ","
                sStudIDtmps &= "'" & Right(dt.Rows(i).Item("StudentID"), 2) & "'"
            End If
        Next

        Dim javascript As String = ""
        javascript = "<script language='javascript'>" & vbCrLf
        javascript &= "   function chk_studentID(num,obj){" & vbCrLf
        javascript &= "      var all=new Array("
        javascript &= sStudIDtmps
        javascript &= ");" & vbCrLf
        javascript &= "      for(var i=0;i<all.length;i++){" & vbCrLf
        javascript &= "         if(document.form1.StudentID.value==all[i] && all[i]!=document.form1.StudentIDstring.value){" & vbCrLf
        javascript &= "            alert('學號重複');" & vbCrLf
        javascript &= "            obj.focus();" & vbCrLf
        javascript &= "         }" & vbCrLf
        javascript &= "      }" & vbCrLf
        javascript &= "   }" & vbCrLf
        javascript &= "</script>"

        Page.RegisterStartupScript("chk_studentID", javascript)
        'chgPriorWorkType1_disabled();
        Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")
        'ViewState("script")=javascript
    End Sub

#Region "NO USE"
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
#End Region

    '取出補助費用歷史頁 補助金額 學員輔助金撥款檔
    Sub GetHistorySumOfMoney(ByVal iSOCID As Integer)
        Const Cst_審核補助金額 As Integer = 5 '審核補助金額
        Const Cst_撥款補助金額 As Integer = 6 '撥款補助金額
        Dim dt As DataTable

        Dim pms2 As New Hashtable From {{"SOCID", iSOCID}}
        Dim sql As String = ""
        sql &= " SELECT a.sid ,a.idno ,a.Name ,b.SOCID ,c.ClassCName ,c.CyclType" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(c.CLASSCNAME,c.CYCLTYPE) CLASSNAME" & vbCrLf
        sql &= " ,c.STDate ,c.FTDate ,c.Years" & vbCrLf
        sql &= " ,d.PlanYear ,d.PlanID ,d.ComIDNO ,d.SeqNo" & vbCrLf
        sql &= " ,ISNULL(e.SumOfMoney,0) SumOfMoney" & vbCrLf '審核
        sql &= " ,ISNULL(e2.SumOfMoney,0) SumOfMoney2" & vbCrLf '撥款
        sql &= " ,e.BUDID ,bb.BUDNAME" & vbCrLf '預算別
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
        sql &= "  WHERE ca.IDNO=a.idno AND cb.SOCID=@SOCID )" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn, pms2)

        'Session("TC_table")=dt
        'hide_SumOfMoney.Value=0
        Panel.Visible = False
        DataGrid1.Visible = False
        msg.Text = "查無資料!!"
        'bt_save.Visible=False
        If dt.Rows.Count > 0 Then
            'bt_save.Visible=True
            Panel.Visible = True
            msg.Text = ""
            DataGrid1.Visible = True
            DataGrid1.Columns(Cst_審核補助金額).Visible = False
            DataGrid1.Columns(Cst_撥款補助金額).Visible = False
            '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            'If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '    DataGrid1.Columns(Cst_審核補助金額).Visible=True
            '    DataGrid1.Columns(Cst_撥款補助金額).Visible=False
            'End If
            dt.DefaultView.Sort = "PlanYear,STDate"
            dt = TIMS.dv2dt(dt.DefaultView)
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
            'PageControler1.PageDataTable=dt '.SqlString=sqlstr
            'PageControler1.PrimaryKey="PlanYear"
            'PageControler1.Sort="PlanYear,STDate"
            'PageControler1.ControlerLoad()
            'bt_save.Visible=False
        End If
        'dr=DbAccess.GetOneRow(sql)
        'DGIdentValue.Text=dr("Share_Name")
    End Sub

    '取出學習卷的學員身分資料
    'Sub GetGovIdent(ByVal SOCID As Integer)
    '    Dim sql As String=""
    '    Dim dr As DataRow
    '    'SELECT * FROM ADP_SHARESOURCE WHERE Share_Type='528'
    '    sql=""
    '    sql &= " SELECT b.Share_Name "
    '    sql &= " FROM Adp_GOVTRNData a "
    '    sql &= " JOIN Adp_ShareSource b ON a.SPECIAL_TYPE=b.Share_ID AND b.Share_Type='528' "
    '    sql &= " WHERE a.SOCID='" & SOCID & "' "
    '    dr=DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then GovSpecial_Type.Text=Convert.ToString(dr("Share_Name")) '推介單個案區分
    '    'SELECT * FROM ADP_SHARESOURCE WHERE Share_Type='527'
    '    sql=""
    '    sql &= " SELECT b.Share_Name "
    '    sql &= " FROM Adp_GOVTRNData a "
    '    sql &= " JOIN Adp_ShareSource b ON a.OBJECT_TYPE=b.Share_ID AND b.Share_Type='527' "
    '    sql &= " WHERE a.SOCID='" & SOCID & "' "
    '    dr=DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then GovObject_Type.Text=Convert.ToString(dr("Share_Name")) '推介單身分別 
    'End Sub

    '清理資料 (增加欄位規則性停用或可用)
    Sub Clear_data()
        EnterChannel.Enabled = True
        TRNDMode.Enabled = True
        'TRNDType.Enabled=True
        'SubsidyID.Enabled=True '有效
        'SubsidyIdentity.Enabled=True '有效
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
        Name.Enabled = True
        RMPNAME.Enabled = True
        LName.Enabled = True
        FName.Enabled = True
        'IDNO.Enabled=True
        'Birthday.Enabled=True
        BudID.Enabled = True
        'MIdentityID.Enabled=True
        IdentityID.Enabled = True
        'SupplyID.Enabled=True
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

        'PostNo_1.Text=""
        ''PostNo_2.Text=""
        'AcctNo1_1.Text=""
        ''AcctNo1_2.Text=""
        'BankName.Text=""

        'AcctheadNo.Text=""
        'ExBankName.Text=""

        'AcctExNo.Text=""
        'AcctNo2.Text=""
        'FirDate.Text=""
        Uname.Text = ""
        Intaxno.Text = ""
        ActName.Text = ""
        ActNo1.Text = ""

        'Zip.Value=""
        'ZIP6W.Value=""
        'Zip_N.Value=""
        'City5.Text=""
        'Addr.Text=""
        'Tel.Text=""
        'Fax.Text=""

        ServDept.Text = ""
        JobTitle.Text = ""
        ddlSERVDEPTID.SelectedIndex = -1
        ddlJOBTITLEID.SelectedIndex = -1
        'SDate.Text=""
        'SJDate.Text=""
        'SPDate.Text=""

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
        'Checkbox4.Checked=False
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
        'If Not SubsidyID.SelectedItem Is Nothing Then SubsidyID.SelectedItem.Selected=False
        'If Not SubsidyIdentity.SelectedItem Is Nothing Then SubsidyIdentity.SelectedItem.Selected=False

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
        'If Not SupplyID.SelectedItem Is Nothing Then SupplyID.SelectedItem.Selected=False
        If Not TRNDMode.SelectedItem Is Nothing Then TRNDMode.SelectedItem.Selected = False
        'If Not TRNDType.SelectedItem Is Nothing Then TRNDType.SelectedItem.Selected=False
        If Not EnterChannel.SelectedItem Is Nothing Then EnterChannel.SelectedItem.Selected = False
        If Not JobStateType.SelectedItem Is Nothing Then JobStateType.SelectedItem.Selected = False '就職狀況'0:失業 1:在職 

        For i As Integer = 0 To IdentityID.Items.Count - 1
            IdentityID.Items(i).Selected = False
        Next
        If Not BudID.SelectedItem Is Nothing Then BudID.SelectedItem.Selected = False
        'For i As Integer=0 To RelClass_Unit.Items.Count - 1
        '    RelClass_Unit.Items(i).Selected=False
        'Next
        'For i As Integer=0 To BudID.Items.Count - 1
        '    BudID.Items(i).Selected=False
        'Next
        'AcctMode.SelectedIndex=-1
        'PMode.SelectedIndex=-1
        'Q1.SelectedIndex=-1
        'For Each item As ListItem In Q2.Items
        '    item.Selected=False
        'Next
        'Q3.SelectedIndex=-1
        'Q3_Other.Text=""
        'Q4.SelectedIndex=-1
        'Q5.SelectedIndex=-1

        'Q61.Text=""
        'Q62.Text=""
        'Q63.Text=""
        'Q64.Text=""

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

    ''' <summary>
    ''' 取得 參訓身分別
    ''' </summary>
    ''' <returns></returns>
    Function Get_All_Identity2() As String
        Dim rst As String = ""
        For i As Integer = 0 To IdentityID.Items.Count - 1
            If IdentityID.Items(i).Value <> "" AndAlso IdentityID.Items(i).Selected = True Then
                If rst <> "" Then rst &= ","
                rst &= IdentityID.Items(i).Value
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
        tr_DDL_DISASTER.Visible = TIMS.GFG_20250710A

        Call SUtl_Create0()
        If Not IsPostBack Then Call SUtl_Create1()
        Call SUtl_Create0bk()
    End Sub

    '載入資料(每1次)
    Sub SUtl_Create0()
        ''分頁設定 Start
        'PageControler1.PageDataGrid=DataGrid1
        ''分頁設定 End
        Dim flag_SHOW_2020x70 As Boolean = TIMS.SHOW_2020x70(sm)
        Dim flag_SHOW_2020x06 As Boolean = TIMS.SHOW_2020x06(sm)
        'Dim flag_show_actno_budid As Boolean=False '保險證號/預算別代碼 false:不顯示 true:顯示
        flag_show_actno_budid = False '保險證號/預算別代碼 false:不顯示 true:顯示
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then flag_show_actno_budid = True
        If flag_SHOW_2020x70 Then flag_show_actno_budid = True
        If flag_SHOW_2020x06 Then flag_show_actno_budid = True

        Hid_show_actno_budid.Value = ""
        If (flag_show_actno_budid) Then Hid_show_actno_budid.Value = "Y"

        Hid_nouse_SupplyID.Value = ""
        If flag_SHOW_2020x70 Then Hid_nouse_SupplyID.Value = "Y"
        If flag_SHOW_2020x06 Then Hid_nouse_SupplyID.Value = "Y"

        'Dim flag_use_supplyID As Boolean=True
        'If Hid_nouse_SupplyID.Value.Equals("Y") Then flag_use_supplyID=False '不可使用

        'trTPlanid28_1 '頁籤控制
        trTPlanid28_1.Visible = False '頁籤控制 false:不顯示 true:顯示
        If Hid_show_actno_budid.Value = "Y" Then trTPlanid28_1.Visible = True

        '暫時權限Table Start
        ''Dim dtArc As DataTable '暫時權限Table
        dtArc = TIMS.Get_Auth_REndClass(Me, objconn)
        dtArc2 = TIMS.Get_Auth_REndClass2(Me, objconn)
        '暫時權限Table End

        If TIMS.sUtl_ChkTest Then gFlagEnv = False '測試用。
        Button5.Visible = False '(回上一頁)

        'rqOCID=Request("OCID")
        'If Not gFlagEnv Then rqOCID="95945"'"C120191434"
        rqOCID = TIMS.ClearSQM(Request("OCID"))
        If rqOCID = "" Then Exit Sub

        '該民眾不具失、待業身分，不得參加失業者職前訓練。STUD_SELRESULTBLIDET / STUD_SELRESULTBLI
        dtBLIDET1 = TIMS.Get_dtBLIDET1(rqOCID, objconn)

        '屆退官兵者 (依系統日期判斷)
        'flagTPlanID02Plan2=False '判斷計畫為自辦職前。
        'If TIMS.Cst_TPlanID02Plan2.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagTPlanID02Plan2=True '判斷計畫為自辦職前。

        'OJT-21020401：在職進修訓練(自辦) - 學員資料維護：判斷學員為現役軍人時於投保保險證號顯示「在役軍人」、預算別判斷為「就安」
        flagTPlanID06Plan3 = False '判斷計畫為 在職進修訓練(自辦)。
        If TIMS.Cst_TPlanID06Plan3.IndexOf(sm.UserInfo.TPlanID) > -1 Then flagTPlanID06Plan3 = True '在職進修訓練(自辦)。

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
        'fontGraduateY.Style.Add("color", "black") '畢業年份-非必填
        MIdentityID.Attributes.Add("onchange", "MIdentityChg(this.value);ChkMIdentityID();")

        Button1.Attributes("onclick") = "return chkdata();"
        Button2.Attributes("onclick") = "return chkdata();"
        EnterChannel.Attributes("onchange") = "EnterChannelChange();"
        'TRNDMode.Attributes("onchange")="TRNDModeChange();"
        If StudentID.Text <> "" Then StudentID.Attributes("onblur") = "chk_studentID(this.value,this);"
        Button4.Attributes("onclick") = "if(document.getElementById('IDNO').value==''){alert('請輸入身分證號碼');return false;}"
        'Button7.Attributes("onclick")="if(document.getElementById('ActNo1').value==''){alert('請輸入投保單位保險證號碼');return false;}"
        'Button8.Attributes("onclick")="if(document.getElementById('ActNo2').value==''){alert('請輸入投保單位保險證號碼');return false;}"
        '=該計畫是否使用ECFA
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
        'AcctMode.Attributes("onclick")="ChangeBank();"

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
        '同通訊地址-緊急通知人
        CheckBox2.Attributes.Add("onClick", "return ock_CheckBox2();")
        '同戶籍地址-緊急通知人
        CheckBox3.Attributes.Add("onClick", "return ock_CheckBox3();")

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
            'trSubsidyID.Visible=False '津貼類別
            'trSubsidyIdentity.Visible=False '津貼身分別

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
            'trPMode.Visible=False
            'PMode.Visible=False

            '就職狀況 
            '就職狀況'0:失業 1:在職 
            jobstatetd.Style("display") = cst_none1
            JobStateType.Visible = False
            '津貼類別
            'SubsidyLabel.Visible=False
            '津貼 身分別
            'LabSubsidy.Visible=False
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

        '補助比例
        'SupplyTD.Style("display")=cst_none1
        'SupplyID.Style("display")=cst_none1
        'If Hid_show_actno_budid.Value="Y" Then
        '    '補助比例
        '    SupplyTD.Style("display")=cst_inline1 '產投必填
        '    SupplyID.Style("display")=cst_inline1 '產投必填
        'End If
        'If Hid_nouse_SupplyID.Value="Y" Then
        '    SupplyID.Style("display")=cst_none1
        'End If

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
                        'If sTPlan06_G22.IndexOf("SubsidyID") > -1 Then SubsidyLabel.Visible=False '津貼類別
                        'If sTPlan06_G22.IndexOf("SubsidyIdentity") > -1 Then LabSubsidy.Visible=False '津貼 身分別
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

    '載入資料(首頁載入1次)
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
        'If iPYNum17=2 AndAlso TIMS.Cst_TPlanID_PreUseLimited17f.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    HidPreUseLimited17f.Value=TIMS.cst_YES
        '    '勞保及3合1就業資料查詢 '勞保及三合一就業資料查詢(MDate)
        '    BtnCheckBli.Attributes("onclick")="open_SD01001sch();return false;"
        '    PriorWorkType1.Attributes("onclick")="chgPriorWorkType1();"
        'End If

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
            'PortTR.Style("display")=cst_none1
            'BankTR1.Style("display")=cst_none1
            'BankTR2.Style("display")=cst_none1
            'BankTR3.Style("display")=cst_none1
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
        'Dim dtDGHR As DataTable=Get_DGTHourDT()
        'Label1.Text=dtDGHR.Rows(0)("DGHour")
        'Label2.Text=dtDGHR.Rows(1)("DGHour")
        'Label3.Text=dtDGHR.Rows(2)("DGHour")
        'Label4.Text=dtDGHR.Rows(3)("DGHour")

        'LearnTR1.Style("display")=cst_none1
        'LearnTR2.Style("display")=cst_none1
        'LearnTR3.Style("display")=cst_none1
        'LearnTR4.Style("display")=cst_none1
        'LearnTR5.Style("display")=cst_none1

        'TPlan23TR.Visible=False
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
                'Table4.Style.Add("display", cst_inline1)

                'document.getElementById('DetailTable').style.display='inline';
                '    if (document.getElementById('BackTable')) document.getElementById('BackTable').style.display='none';
                '    document.getElementById('HistoryTable').style.display='none';
                '    document.getElementById('Table4').style.display='inline';
            End If
        Else
            'Select Case sm.UserInfo.TPlanID
            '    Case "15"  '學習卷計畫
            '        LearnTR1.Style("display")=cst_inline1
            '        LearnTR2.Style("display")=cst_inline1
            '        LearnTR3.Style("display")=cst_inline1
            '        LearnTR4.Style("display")=cst_inline1
            '        LearnTR5.Style("display")=cst_inline1
            '    Case "23", "34", "41" '23:訓用合一, 34:與企業合作辦理職前訓練, 41:推動營造業事業單位辦理職前培訓計畫
            '        TPlan23TR.Visible=True
            '    Case Else
            '        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
            '        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '            MenuTable.Visible=True
            '            BackTable.Visible=True
            '            If Not IsPostBack Then
            '                Page.RegisterStartupScript("ChangeMode1", "<script>ChangeMode(1);</script>")
            '                DetailTable.Style.Add("display", cst_inline1)
            '                BackTable.Style.Add("display", cst_none1)
            '                HistoryTable.Style.Add("display", cst_none1)
            '                Table4.Style.Add("display", cst_inline1)
            '            End If
            '        End If
            'End Select
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
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If OrgKind2 <> "W" Then AcctMode.Items.Remove(AcctMode.Items.FindByValue(2)) 'W:提升在職勞工自主學習計畫 'G:產業人才投資計畫
        'Else
        '    AcctMode.Items.Remove(AcctMode.Items.FindByValue(2))
        'End If

        '3+2郵遞區號查詢 link
        LitZipCode4.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCode1.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCode2.Text = TIMS.Get_WorkZIPB3Link2()
        LitZipCode3.Text = TIMS.Get_WorkZIPB3Link2()
        LitForeZip.Text = TIMS.Get_WorkZIPB3Link2()
        LitActZip.Text = TIMS.Get_WorkZIPB3Link2()

        '輸入郵遞區號3碼判斷 & 代入 CityName 
        ZipCode4.Attributes.Add("onblur", "getZipName('City4',this,this.value)")
        ZipCode1.Attributes.Add("onblur", "getZipName('City1',this,this.value);")
        ZipCode2.Attributes.Add("onblur", "getZipName('City2',this,this.value);")
        ZipCode3.Attributes.Add("onblur", "getZipName('City3',this,this.value);")
        ForeZip.Attributes.Add("onblur", "getZipName('City6',this,this.value);")
        txt_ActZip.Attributes.Add("onblur", "getZipName('txt_ActCity',this,this.value);")
        'Zip.Attributes.Add("onblur", "getZipName('City5',this,this.value);")

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
        'Dim bt7_Attr_VAL As String=TIMS.GET_ZipCodeWOAspxUrl(Zip, ZIPB3, hidZIP6W, City5, hidCityName7, hidAREA7, Zip_N, Addr)
        'bt_openZip7.Attributes.Add("onclick", bt7_Attr_VAL)

        'for 個資保護
        Dim Req_SM As String = TIMS.ClearSQM(Request("SM"))
        Call Std_Data_Mask(Req_SM)
    End Sub

    ''' <summary>
    ''' 增修需求 OJT-121201_系統_產投_學員資料維護_外籍補助80%+預算別與補助比例連動調整
    ''' </summary>
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
        '主要參訓身分別 SELECT * FROM KEY_IDENTITY WHERE 1=1 AND IdentityID IN ('30','40') 
        Dim v_MIdentityID As String = TIMS.GetMyValue2(in_parms, "MIdentityID")
        '預算別 01:公務;02:就安;03:就保;04:再出發;97:公務(ECFA);98:特別預算;99:不補助--SELECT ';'+BUDID+':'+BUDNAME FROM VIEW_BUDGET
        Dim v_BudID As String = TIMS.GetMyValue2(in_parms, "BudID")
        '補助比例 1:80%;2:100%;9:不補助 --SELECT ';'+SUPPLYID+':'+SNAME FROM VIEW_SUPPLYID
        Dim v_SupplyID As String = TIMS.GetMyValue2(in_parms, "SupplyID")
        Dim v_TPlanID As String = TIMS.GetMyValue2(in_parms, "TPlanID")

        Dim STR_NOUSE_SUPPLYID As String = TIMS.GetMyValue2(in_parms, "STR_NOUSE_SUPPLYID")
        Dim flag_use_supplyID As Boolean = True
        If STR_NOUSE_SUPPLYID.Equals("Y") Then flag_use_supplyID = False '不可使用

        '主要參訓身分別為「因應貿易自由化協助勞工」，預算別只能為 不補助/ECFA
        Dim flag_PLAN_1 As Boolean = (v_MIdentityID = cst_MID_因應貿易自由化協助勞工_30 AndAlso (TIMS.Cst_TPlanID70.IndexOf(v_TPlanID) = -1)) '案例1
        Dim flag_PLAN_2 As Boolean = (v_MIdentityID = cst_MID_經公告之重大災害受災者_40 AndAlso (TIMS.Cst_TPlanID70.IndexOf(v_TPlanID) = -1)) '案例2
        Dim flag_PLAN_3 As Boolean = (v_BudID = cst_BUDID_97_公務ECFA AndAlso (TIMS.Cst_TPlanID70.IndexOf(v_TPlanID) = -1)) '案例3

        'Dim sErrmsg1 As String=""
        If sErrmsg1 <> "" Then Return False

        If v_PassPortNO = "" Then sErrmsg1 &= "請選擇 身分別!" & vbCrLf
        If v_MIdentityID = "" Then sErrmsg1 &= "請選擇 主要參訓身分別!" & vbCrLf
        If v_BudID = "" Then sErrmsg1 &= "請選擇 預算別!" & vbCrLf
        If v_SupplyID = "" OrElse v_SupplyID = "0" Then If (flag_use_supplyID) Then sErrmsg1 &= "請選擇 補助比例!" & vbCrLf
        If sErrmsg1 <> "" Then Return False

        If v_BudID.Equals("99") AndAlso Not v_SupplyID.Equals("9") Then
            If (flag_use_supplyID) Then sErrmsg1 &= "預算別為不補助，補助比例有誤!(應為不補助)" & vbCrLf
        ElseIf v_SupplyID.Equals("9") AndAlso Not v_BudID.Equals("99") Then
            If (flag_use_supplyID) Then sErrmsg1 &= "補助比例為不補助，預算別有誤!(應為不補助)" & vbCrLf
        End If
        If sErrmsg1 <> "" Then Return False

        If v_PassPortNO = "2" AndAlso Not flag_PLAN_1 AndAlso Not flag_PLAN_2 Then
            If v_SupplyID.Equals("2") Then '1:80%;2:100%;9:不補助 
                If (flag_use_supplyID) Then sErrmsg1 &= "身分別為「外籍(含大陸人士)」，補助比例有誤!(不可為100%)" & vbCrLf
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
        End If
        If sErrmsg1 <> "" Then Return False

        If flag_PLAN_2 Then
            Dim flag_PLAN_2_OK As Boolean = True
            If v_BudID.Equals("97") Then flag_PLAN_2_OK = False
            If Not flag_PLAN_2_OK Then
                '只能為 02:就安/ 03:就保/ 99:不補助
                sErrmsg1 &= "主要參訓身分別為「經公告之重大災害受災者」，預算別有誤!(不可為公務(ECFA))" & vbCrLf
            Else
                If v_SupplyID.Equals("1") Then '1:80% (不可為80%)
                    If (flag_use_supplyID) Then sErrmsg1 &= "主要參訓身分別為「經公告之重大災害受災者」，補助比例有誤!(不可為80%)" & vbCrLf
                End If
            End If
        End If
        If sErrmsg1 <> "" Then Return False

        If flag_PLAN_3 AndAlso v_MIdentityID <> cst_MID_因應貿易自由化協助勞工_30 Then
            sErrmsg1 &= "預算別為 公務(ECFA)，主要參訓身份別 須為「因應貿易自由化協助勞工」，主要參訓身分別有誤!" & vbCrLf
        End If
        If sErrmsg1 <> "" Then Return False

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
        '非 ROLEID=0 LID=0
        Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        'ROLEID=0 LID=0
        If TIMS.IsSuperUser(Me, 1) Then flgROLEIDx0xLIDx0 = True '判斷登入者的權限。

        rqOCID = TIMS.ClearSQM(rqOCID)
        Dim drCC As DataRow = TIMS.GetOCIDDate(rqOCID, objconn) ' DbAccess.GetOneRow(sql, objconn)
        If rqOCID = "" OrElse drCC Is Nothing Then
            Errmsg &= cst_errMsg12 & vbCrLf
            Return False
        End If

        Dim v_MilitaryID As String = TIMS.GetListValue(MilitaryID) '兵役狀況 
        Dim flag_use_supplyID As Boolean = True
        If Hid_nouse_SupplyID.Value.Equals("Y") Then flag_use_supplyID = False '不可使用

        Name.Text = TIMS.ClearSQM(Name.Text)
        RMPNAME.Text = TIMS.ClearSQM(RMPNAME.Text)
        LName.Text = TIMS.ClearSQM(LName.Text)
        FName.Text = TIMS.ClearSQM(FName.Text)

        ZipCode1.Value = TIMS.ClearSQM(ZipCode1.Value)
        ZipCode1_B3.Value = TIMS.ClearSQM(ZipCode1_B3.Value)
        hidZipCode1_6W.Value = TIMS.GetZIPCODE6W(ZipCode1.Value, ZipCode1_B3.Value)
        ZipCode1_N.Value = TIMS.ClearSQM(ZipCode1_N.Value)
        Address.Text = TIMS.ClearSQM(Address.Text)

        ZipCode2.Value = TIMS.ClearSQM(ZipCode2.Value)
        ZipCode2_B3.Value = TIMS.ClearSQM(ZipCode2_B3.Value)
        hidZipCode2_6W.Value = TIMS.GetZIPCODE6W(ZipCode2.Value, ZipCode2_B3.Value)
        ZipCode2_N.Value = TIMS.ClearSQM(ZipCode2_N.Value)
        HouseholdAddress.Text = TIMS.ClearSQM(HouseholdAddress.Text)

        ZipCode3.Value = TIMS.ClearSQM(ZipCode3.Value)
        ZipCode3_B3.Value = TIMS.ClearSQM(ZipCode3_B3.Value)
        hidZipCode3_6W.Value = TIMS.GetZIPCODE6W(ZipCode3.Value, ZipCode3_B3.Value)
        ZipCode3_N.Value = TIMS.ClearSQM(ZipCode3_N.Value)
        EmergencyAddress.Text = TIMS.ClearSQM(EmergencyAddress.Text)

        '身分證驗証
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        Dim aIDNO As String = IDNO.Text

        '通訊地址前3碼郵遞區號
        'If ZipCode1.Value="" Then'End If
        Dim sql As String = ""
        Dim dtZipCode As DataTable
        sql = "SELECT * FROM dbo.VIEW_ZIPNAME ORDER BY ZIPCODE"
        dtZipCode = DbAccess.GetDataTable(sql, objconn)
        TIMS.CheckValeuErr(ZipCode1.Value, "通訊地址前3碼郵遞區號", True, "ZipCode", dtZipCode, Errmsg)
        TIMS.CheckZipCODEB3(ZipCode1_B3.Value, "通訊地址郵遞區號後2碼", True, Errmsg)
        TIMS.CheckVal(Address.Text, "通訊地址", Errmsg)

        '同通訊地址。
        If Not CheckBox1.Checked Then
            TIMS.CheckValeuErr(ZipCode2.Value, "戶籍地址前3碼郵遞區號", True, "ZipCode", dtZipCode, Errmsg)
            TIMS.CheckZipCODEB3(ZipCode2_B3.Value, "戶籍地址郵遞區號後2碼", True, Errmsg)
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
        'Dim v_SupplyID As String=TIMS.GetListValue(SupplyID) 'SupplyID.SelectedValue

        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If flag_show_actno_budid Then
            '產投
            ActNo1.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(ActNo1.Text))
            If ActNo1.Text <> "" Then
                '投保單位保險證號為09開頭者，為訓字保，亦不可報名 '根據參訓學員於e網所填列之保險證號前2碼判讀, 前2碼為
                '01、04、05、15及08者其補助經費來源歸屬為 03:就保基金
                '02、03、06、07者其經費來源歸屬為 02:就安基金
                '09與無法辨視者為 99:不予補助對象
                Select Case Left(ActNo1.Text, 2)
                    Case "09"
                        Errmsg &= "學員資格 投保單位保險證號 為09開頭者為訓字保 不符合可參訓條件！" & vbCrLf
                End Select
            End If

            '檢測此學員是否 可參訓 產業人才投資方案 (大於15歲者)
            If Not TIMS.Check_YearsOld15(Birthday.Text, Convert.ToString(ViewState(vs_STDate))) Then Errmsg &= "學員資格 年齡不滿15歲 不符合可參訓條件！" & vbCrLf

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
            If ActNo1.Text = "" Then ActNo1.Text = ActNo2.Text
            If ActNo1.Text = "" Then
                If v_BudID = "97" Then
                    Errmsg &= "預算別選" & cst_ECFA & "："
                    Errmsg &= Cst_Msg6 & vbCrLf
                End If
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
                        '什麼時候可以修改預算別。
                        flag_BudIDNoLock = Chk_CanEditBudgetID(sm.UserInfo.LID, CDate(STDateHidden.Value), iBudFlag)

                        If DateDiff(DateInterval.Day, CDate(Cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
                            Select Case v_BudID 'BudID.SelectedValue
                                Case "99", "04" '(預算別) 排除使用 ECFA 'BudID 99:不予補助對象  'BudID 04:再出發對象  
                                Case Else
                                    '檢驗是否可用ECFA ，並限定使用ECFA
                                    If TIMS.CheckIsECFA(Me, ActNo1.Text, vsFOfficeYM1, STDateHidden.Value, objconn) = True Then
                                        '若是鎖定，則做檢查ECFA 若 開放 則不做判斷
                                        If Not flag_BudIDNoLock Then
                                            'If v_BudID <> "97" OrElse v_SupplyID <> "2" Then Errmsg &= Cst_Msg3 & vbCrLf
                                            If v_BudID <> "97" Then Errmsg &= Cst_Msg3 & vbCrLf
                                        End If
                                    Else
                                        If v_BudID = "97" Then Errmsg &= Cst_Msg5 & vbCrLf
                                    End If
                            End Select
                        Else
                            If v_BudID = "97" Then Errmsg &= Cst_Msg7 & vbCrLf
                        End If
                    Else
                        If v_BudID = "97" Then
                            If Not TIMS.CheckIsECFA(Me, ActNo1.Text, vsFOfficeYM1, STDateHidden.Value, objconn) Then Errmsg &= Cst_Msg5 & vbCrLf
                        End If
                    End If
                End If
                '=該計畫是否使用ECFA
            End If
        End If

        'Dim v_SubsidyIdentity As String=TIMS.GetListValue(SubsidyIdentity)
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
                'If v_SubsidyIdentity="04" Then Errmsg &= cst_errMsg5 & vbCrLf
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
                'If v_SubsidyIdentity="37" Then Errmsg &= cst_errMsg8 & vbCrLf
            End If
        End If

        'If Errmsg="" Then
        '    If flagTPlanID02Plan2 Then
        '        '屆退官兵者 (依開訓日期(系統日期)判斷)
        '        Dim flag_is_RESOLDER As Boolean=TIMS.CheckRESOLDER(objconn, IDNO.Text, sm.UserInfo.DistID, ViewState(vs_STDate))
        '        If Not flag_is_RESOLDER Then If v_MIdentityID="12" Then Errmsg &= cst_errMsg9 & vbCrLf
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
                        Case "1"
                            If Not TIMS.CheckIDNO2(aIDNO, 3) Then Errmsg &= "護照號碼錯誤!請聯絡系統管理員" & vbCrLf '一般驗証
                        Case "2"
                            '2:居留證 4:居留證2021
                            Dim flag2 As Boolean = TIMS.CheckIDNO2(aIDNO, 2)
                            Dim flag4 As Boolean = TIMS.CheckIDNO2(aIDNO, 4)
                            If Not flag2 And Not flag4 Then Errmsg &= "居留證號碼錯誤!請聯絡系統管理員" & vbCrLf '一般驗証
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
        If FTDateHidden.Value = "" OrElse rqOCID = "" Then
            Errmsg &= cst_errMsg12 & vbCrLf 'Exit Sub
            Return False
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
        '    '屆退官兵者 (依開訓日期(系統日期)判斷) KEY_IDENTITY
        '    Dim flag_is_RESOLDER As Boolean=TIMS.CheckRESOLDER(objconn, IDNO.Text, sm.UserInfo.DistID, ViewState(vs_STDate))
        '    If flag_is_RESOLDER Then
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
        If gFlagEnv Then '正式環境。('測試用。)
            'Dim vTitle As String=""
            'vTitle="授權設定該班級有開放"
            If FTDateHidden.Value <> "" AndAlso TIMS.ChkIsEndDate(rqOCID, TIMS.cst_FunID_學員資料維護, dtArc) Then
                ''學員資料維護於訓後30日不能修改'Cst_Msg30
                'If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 30, CDate(FTDateHidden.Value)), Today) >= 0 Then ErrMessage &= Cst_Msg30 & vbCrLf
                If TIMS.Cst_TPlanID14DayCanEditStud.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID <= 1 Then
                    If Not flgROLEIDx0xLIDx0 Then '判斷登入者的權限。
                        '非 ROLEID=0 LID=0 '學員資料維護於訓後30日不能修改'Cst_Msg30
                        If DateDiff(DateInterval.Day, DateAdd(DateInterval.Month, 3, CDate(FTDateHidden.Value)), Today) >= 0 Then Errmsg &= Cst_Msg30x28 & vbCrLf
                    End If
                Else
                    '針對委外職前訓練計畫，系統權限 限制
                    Dim flgElseEvent1 As Boolean = True '其它狀況 預設為True
                    'If Cst_TPlanIDCanEditStud_id37.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '    Select Case sm.UserInfo.LID
                    '        Case "1" '分署(中心)
                    '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                    '                flgElseEvent1 = False '已經設定狀況 其它@False
                    '                Errmsg &= Cst_MsgTPlanID37a & vbCrLf
                    '            End If
                    '        Case "2" '委外
                    '            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, cst_limitDay21ft, CDate(FTDateHidden.Value)), Today) >= 0 Then
                    '                flgElseEvent1 = False '已經設定狀況 其它@False
                    '                Errmsg &= Cst_MsgTPlanID37b & vbCrLf
                    '            End If
                    '    End Select
                    'End If

                    '其它狀況 預設為True
                    If flgElseEvent1 Then
                        If FTDateHidden.Value = "" Then
                            '學員資料維護於訓後30日不能修改'Cst_Msg30
                            If Not flgROLEIDx0xLIDx0 Then Errmsg &= Cst_Msg30 & vbCrLf '判斷登入者的權限。
                        Else
                            '學員資料維護於訓後30日不能修改'Cst_Msg30
                            If DateDiff(DateInterval.Day, DateAdd(DateInterval.Day, 30, CDate(FTDateHidden.Value)), Today) >= 0 Then
                                If Not flgROLEIDx0xLIDx0 Then Errmsg &= Cst_Msg30 & vbCrLf '判斷登入者的權限。
                            End If
                        End If
                    End If
                End If
            End If
        End If


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
            If dr IsNot Nothing Then
                Errmsg &= "此班級已經有相同的身分證號碼!" & vbCrLf
                'Common.MessageBox(Me, "此班級已經有相同的身分證號碼!")
                'Page.RegisterStartupScript("hard", "<script>hard();</script>")
                'Exit Function '離開 
            End If
        End If

        '產投檢查。
        Dim flagTPlanID28a As Boolean = False '(產投 28.54) 在職用
        'Dim flagTIMSNot28a As Boolean=True '(TIMS) 職前用(非在職)
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If Hid_show_actno_budid.Value = "Y" Then flagTPlanID28a = True    'flagTIMSNot28a=False

        School.Text = TIMS.ClearSQM(School.Text)
        Department.Text = TIMS.ClearSQM(Department.Text)
        If School.Text = "" Then Errmsg &= "請輸入 個人基本資料-學校名稱" & vbCrLf
        If Department.Text = "" Then Errmsg &= "請輸入 個人基本資料-科系" & vbCrLf

        '產投檢查。
        If flagTPlanID28a Then
            Dim v_GraduateStatus As String = TIMS.GetListValue(GraduateStatus)
            Dim TestVal As String = TIMS.Get_GraduateStatusValue(v_GraduateStatus)
            If TestVal = "" Then Errmsg &= "請選擇 個人基本資料-畢業狀況" & vbCrLf
            'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID)=-1 Then
            '    Dim v_graduatey As String=TIMS.GetListValue(graduatey)
            '    If v_graduatey="" Then Errmsg &= "請選擇 個人基本資料-畢業狀況-畢業年份" & vbCrLf
            'End If
            If Uname.Text <> "" Then Uname.Text = TIMS.ClearSQM(Uname.Text)
            If ServDept.Text <> "" Then ServDept.Text = TIMS.ClearSQM(ServDept.Text)
            If Uname.Text = "" Then Errmsg &= "請輸入服務單位資料-目前任職公司" & vbCrLf '服務部門(目前任職公司) Uname
            If ServDept.Visible AndAlso ServDept.Text = "" Then Errmsg &= "請輸入服務單位資料-目前任職部門" & vbCrLf '服務部門(目前任職部門) ServDept

            Dim v_ddlSERVDEPTID As String = TIMS.GetListValue(ddlSERVDEPTID)
            Dim v_ddlJOBTITLEID As String = TIMS.GetListValue(ddlJOBTITLEID)

            'Dim v_MIdentityID As String=TIMS.GetListValue(MIdentityID)
            If ddlSERVDEPTID.Visible AndAlso v_ddlSERVDEPTID = "" Then Errmsg &= "請選擇 服務單位資料-目前任職部門" & vbCrLf
            If v_MIdentityID <> "02" Then
                If JobTitle.Visible AndAlso JobTitle.Text = "" Then Errmsg &= "請輸入服務單位資料-職務" & vbCrLf
                If ddlJOBTITLEID.Visible AndAlso v_ddlJOBTITLEID = "" Then Errmsg &= "請選擇服務單位資料-職務" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Return False

        ServiceID.Text = TIMS.ClearSQM(ServiceID.Text) '軍種
        MilitaryRank.Text = TIMS.ClearSQM(MilitaryRank.Text) '階級
        ServiceOrg.Text = TIMS.ClearSQM(ServiceOrg.Text) '服務單位名稱
        ServicePhone.Text = TIMS.ClearSQM(ServicePhone.Text) '單位電話
        SServiceDate.Text = TIMS.ClearSQM(SServiceDate.Text) '服役日期-開始
        FServiceDate.Text = TIMS.ClearSQM(FServiceDate.Text) '服役日期-結束

        MilitaryAppointment.Text = TIMS.ClearSQM(MilitaryAppointment.Text) '職務(兵役)
        ChiefRankName.Text = TIMS.ClearSQM(ChiefRankName.Text) '主管階級 姓名
        ServiceAddress.Text = TIMS.ClearSQM(ServiceAddress.Text) '服役單位 地址

        Const cst_isReq1 As String = " (兵役狀況 「在役中」 為必填)"
        '兵役狀況 -在役中
        If v_MilitaryID = "04" Then
            'SolTR.Attributes
            SolTR.Style("display") = cst_inline1
            If ServiceID.Text = "" Then Errmsg &= String.Format("軍種 不完整{0}", cst_isReq1) & vbCrLf
            If MilitaryRank.Text = "" Then Errmsg &= String.Format("階級 不完整{0}", cst_isReq1) & vbCrLf
            If ServiceOrg.Text = "" Then Errmsg &= String.Format("服務單位名稱 不完整{0}", cst_isReq1) & vbCrLf
            If ServicePhone.Text = "" Then Errmsg &= String.Format("單位電話  不完整{0}", cst_isReq1) & vbCrLf
            If SServiceDate.Text = "" Then Errmsg &= String.Format("服役日期 起始日期 不完整{0}", cst_isReq1) & vbCrLf
            If FServiceDate.Text = "" Then Errmsg &= String.Format("服役日期 結束日期 不完整{0}", cst_isReq1) & vbCrLf
        End If

        If SServiceDate.Text <> "" AndAlso Not TIMS.IsDate1(SServiceDate.Text) Then
            'SServiceDate.Text=""
            Errmsg &= "服役日期 起始日期 格式有誤!" & vbCrLf
        End If
        If FServiceDate.Text <> "" AndAlso Not TIMS.IsDate1(FServiceDate.Text) Then
            'FServiceDate.Text=""
            Errmsg &= "服役日期 結束日期 格式有誤!" & vbCrLf
        End If
        SServiceDate.Text = TIMS.Cdate3(SServiceDate.Text)
        FServiceDate.Text = TIMS.Cdate3(FServiceDate.Text)

        If rst AndAlso Errmsg = "" Then
            If Hid_show_actno_budid.Value = "Y" Then
                'Dim v_PassPortNO As String=TIMS.GetListValue(PassPortNO)
                'Dim v_MIdentityID As String=TIMS.GetListValue(MIdentityID)
                'Dim v_BudID As String=TIMS.GetListValue(BudID)
                'Dim v_SupplyID As String=TIMS.GetListValue(SupplyID)
                Dim in_parms As New Hashtable
                in_parms.Clear()
                in_parms.Add("PassPortNO", v_PassPortNO)
                in_parms.Add("MIdentityID", v_MIdentityID)
                in_parms.Add("BudID", v_BudID)
                'in_parms.Add("SupplyID", v_SupplyID)
                'STR_NOUSE_SUPPLYID
                in_parms.Add("STR_NOUSE_SUPPLYID", Hid_nouse_SupplyID.Value)
                in_parms.Add("TPlanID", sm.UserInfo.TPlanID)
                'chkFlag=sUtl_CheckData2(in_parms, Errmsg)
                rst = SUtl_CheckData2(in_parms, Errmsg)
            End If
        End If

        '系統先去比對送訓官兵名冊，(比照參訓學員投保狀況檢核表)
        '如該學員為現役軍人：
        '1.【投保單位保險證號】預帶「在役軍人」
        '2.【預算別】預帶「就安」
        '3.【投保單位名稱】預帶送訓官兵名冊之「任職單位」
        '4. 儲存時，增加檢核【投保單位保險證號】為「在役軍人」，【預算別】應為「就安」
        'Hid_out_POSITION.Value=""
        If flagTPlanID06Plan3 Then
            Dim out_POSITION As String = ""
            Dim flag_SRSOLDIERS As Boolean = False '是否為屆退官兵
            'sm.UserInfo.DistID / Convert.ToString(drCC("DISTID"))
            flag_SRSOLDIERS = TIMS.CheckRESOLDER(objconn, IDNO.Text, Convert.ToString(drCC("DISTID")), STDateHidden.Value, out_POSITION)
            If (flag_SRSOLDIERS) Then
                'Hid_out_POSITION.Value=out_POSITION
                'ActName.Text=out_POSITION '【投保單位名稱】預帶送訓官兵名冊之「任職單位」
                '在役軍人
                If (ActNo1.Text <> cst_Serviceman) Then Errmsg &= "學員為現役軍人 【投保單位保險證號】 應為「在役軍人」" & vbCrLf
                '03:就保基金'02:就安基金'99:不予補助對象
                If (v_BudID <> "02") Then Errmsg &= "學員為現役軍人 【預算別】應為「就安」" & vbCrLf
                '兵役狀況 ===請選擇===
                If (v_MilitaryID <> "04") Then Errmsg &= "學員為現役軍人 【兵役狀況】應為「在役中」" & vbCrLf
                'Common.SetListItem(BudID, "02")
                'ActName.Enabled=False
                'ActNo1.Enabled=False
                'BudID.Enabled=False
                'TIMS.Tooltip(ActName, cst_Serviceman)
                'TIMS.Tooltip(ActNo1, cst_Serviceman)
                'TIMS.Tooltip(BudID, cst_Serviceman)
            End If
        End If

        If tr_DDL_DISASTER.Visible AndAlso v_MIdentityID = TIMS.cst_Identity_40 AndAlso v_DDL_DISASTER = "" Then 'ADID 重大災害選項
            Errmsg &= "主要參訓身分別選擇「經公告之重大災害受災者」，須選擇「重大災害選項」不可為空！" & vbCrLf
        End If

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

#Region "(No Use)"
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
#End Region

    '(儲存)Button1:儲存回查詢頁面 ／Button2:維護下一位學員
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click, Button2.Click
        '重載數據(JS無法讀取有效值)
        'Call ReLoad_SB4IDx()

        Dim Errmsg As String = ""
        Dim js_Errmsg As String = ""
        Dim chkFlag As Boolean = False '有錯誤為False
        '儲存前先檢查輸入資料的正確性
        chkFlag = SUtl_CheckData1(Errmsg)

        '= 錯誤訊息，顯示並離開
        If Errmsg <> "" Then
            If Hid_show_actno_budid.Value = "Y" Then
                js_Errmsg = Common.GetJsString(Errmsg)
                'Page.RegisterStartupScript("sol", "<script>sol(" & v_MilitaryID & ");</script>")
                Page.RegisterStartupScript("11111", "<script>blockAlert('" & js_Errmsg & "','錯誤',function(){ChangeMode(1);});</script>")
                'Common.MessageBox(Me, Errmsg)
                'Exit Sub
            Else
                js_Errmsg = Common.GetJsString(Errmsg)
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
        '= 錯誤訊息，顯示並離開

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

        '更新MakeSOCID
        'Dim vRejectSOCID As String=TIMS.ClearSQM(RejectSOCID.SelectedValue)
        Dim v_SOCID As String = TIMS.GetListValue(SOCID)
        'Dim vSOCID_Sel As String=TIMS.ClearSQM(v_SOCID) 'SOCID.SelectedValue)
        Call TIMS.OpenDbConn(objconn)
        Dim updateCS As Boolean = True
        Dim sql As String = ""
        If StdTr.Visible Then
            '修改
            If v_SOCID = "" Then updateCS = False '應該只有值，其他答案為錯誤
            If Not updateCS Then
                Common.MessageBox(Me, "修改動作，查無可修改學員資料，請重新查詢操作!")
                Exit Sub
            End If

            '修改
            Dim pms2 As New Hashtable From {{"SOCID", TIMS.CINT1(v_SOCID)}}
            sql = "SELECT 'x' FROM CLASS_STUDENTSOFCLASS WHERE SOCID=@SOCID"
            Dim dtRR As DataTable = DbAccess.GetDataTable(sql, objconn, pms2)
            If dtRR.Rows.Count <> 1 Then
                updateCS = False '應該只有一筆，其他數字者為錯誤
            End If
            If Not updateCS Then
                Common.MessageBox(Me, "修改動作，查無可修改學員資料，請重新查詢操作!")
                Exit Sub
            End If
        End If

        Dim ref_SID As String = "" '回傳SID
        Try
            Call SUtl_SaveData1(ref_SID) 'STUD_STUDENTINFO 傳出 SID
        Catch ex As Exception
            Dim sErrmsg As String = ""
            sErrmsg = "儲存學員資料有誤，請重新操作!!" & ex.Message
            Common.MessageBox(Me, sErrmsg)

            Dim strErrmsg As String = ""
            strErrmsg = "##WDAIIP.SD_03_002_add.sUtl_SaveData1"
            strErrmsg &= "SID:" & vbCrLf & ref_SID & vbCrLf
            strErrmsg &= "IDNO:" & vbCrLf & IDNO.Text & vbCrLf
            strErrmsg &= "rqOCID:" & vbCrLf & rqOCID & vbCrLf
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try
        Hid_SID_C1.Value = ref_SID
        If ref_SID = "" Then
            Common.MessageBox(Me, "儲存學員資料有誤，請重新操作!")
            Exit Sub
        End If
        Call SUtl_SaveData2() 'STUD_SUBDATA
        Call SUtl_SaveData3() 'STUD_ENTERTEMP'STUD_ENTERTEMP2
        Call SUtl_SaveData4() 'CLASS_STUDENTSOFCLASS '傳入 SID/ Hid_SID_C1.Value/ref_SID
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

    'UDPATE STUD_STUDENTINFO
    ''' <summary>
    ''' UDPATE STUD_STUDENTINFO
    ''' </summary>
    ''' <param name="ref_SID"></param>
    ''' <returns>SID</returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData1(ByRef ref_SID As String) As Boolean
        Dim rst As Boolean = False '儲存成功為True
        'Dim SID As String="" '回傳SID
        'ByRef SID As String

        Dim sql As String = ""
        Dim da1 As SqlDataAdapter = Nothing
        Dim MyTable1 As DataTable = Nothing 'STUD_STUDENTINFO
        Dim Mydr1 As DataRow = Nothing 'STUD_STUDENTINFO

        '無身分證號，無法作業
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        If IDNO.Text = "" Then Return False
        If IDNO.Text.Length < 10 Then Return False '長度-小於10-異常
        ForeIDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(ForeIDNO.Text))
        '先檢查是否有資料
        If StdTr.Visible = False Then
            '新增狀態，檢查有沒有個人資料存在
            Dim sMemo As String = $"&NAME={Name.Text}"
            '寫入Log查詢 SubInsAccountLog1(Auth_Accountlog)
            Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm新增, Session(TIMS.gcst_rblWorkMode), rqOCID, "")

            sql = $" SELECT * FROM STUD_STUDENTINFO WHERE IDNO='{IDNO.Text}'" '2009/7/17 不檢查生日
            MyTable1 = DbAccess.GetDataTable(sql, da1, objconn)
            If MyTable1.Rows.Count = 0 Then
                ref_SID = TIMS.Get_DateNo & "01"
                Mydr1 = MyTable1.NewRow
                MyTable1.Rows.Add(Mydr1)
                Mydr1("SID") = ref_SID
                Call SUtl_AddUpdateRow1(Mydr1)

                Dim htPP As New Hashtable From {
                    {"TransType", TIMS.cst_TRANS_LOG_Insert},
                    {"TargetTable", "STUD_STUDENTINFO"},
                    {"FuncPath", "/SD/03/SD_03_002_add2"},
                    {"s_WHERE", String.Format("IDNO='{0}' AND SID='{1}'", IDNO.Text, ref_SID)}
                }
                TIMS.SaveTRANSLOG(sm, objconn, Mydr1, htPP)
                '更新學員基本資料檔
                DbAccess.UpdateDataTable(MyTable1, da1)
            Else
                Common.MessageBox(Me, "此學員個人資料已存在，將更新您所輸入的資料")
                For Each Mydr2 As DataRow In MyTable1.Rows
                    ref_SID = Mydr2("SID") '隨意一組 SID '可能有多組。(合理情況下，應該只有1筆)
                    Call SUtl_AddUpdateRow1(Mydr2)

                    'htPP.Clear()
                    Dim htPP As New Hashtable From {
                        {"TransType", TIMS.cst_TRANS_LOG_Update},
                        {"TargetTable", "STUD_STUDENTINFO"},
                        {"FuncPath", "/SD/03/SD_03_002_add2"},
                        {"s_WHERE", String.Format("IDNO='{0}' AND SID='{1}'", IDNO.Text, ref_SID)}
                    }
                    TIMS.SaveTRANSLOG(sm, objconn, Mydr2, htPP)
                Next
                '更新學員基本資料檔
                DbAccess.UpdateDataTable(MyTable1, da1)
            End If
        Else
            '修改狀態，檢查有沒有個人資料存在
            Dim sMemo As String = ""
            sMemo &= "&NAME=" & Name.Text
            '寫入Log查詢 SubInsAccountLog1(Auth_Accountlog)
            Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm修改, Session(TIMS.gcst_rblWorkMode), rqOCID, "")

            '2009/07/17 改成只判斷身分證字號 (多筆。)
            sql = $" SELECT * FROM STUD_STUDENTINFO WHERE IDNO='{IDNO.Text}' "
            MyTable1 = DbAccess.GetDataTable(sql, da1, objconn)
            If MyTable1.Rows.Count = 0 Then
                Common.MessageBox(Me, "此學員個人資料不存在，將要把您輸入的資料新增存入")
                ref_SID = TIMS.Get_DateNo & "01"
                Mydr1 = MyTable1.NewRow
                MyTable1.Rows.Add(Mydr1)
                Mydr1("SID") = ref_SID
                Call SUtl_AddUpdateRow1(Mydr1)

                Dim htPP As New Hashtable
                htPP.Clear()
                htPP.Add("TransType", TIMS.cst_TRANS_LOG_Insert)
                htPP.Add("TargetTable", "STUD_STUDENTINFO")
                htPP.Add("FuncPath", "/SD/03/SD_03_002_add2")
                htPP.Add("s_WHERE", String.Format("IDNO='{0}' AND SID='{1}'", IDNO.Text, ref_SID))
                TIMS.SaveTRANSLOG(sm, objconn, Mydr1, htPP)

                DbAccess.UpdateDataTable(MyTable1, da1) '更新學員基本資料檔
            Else
                Common.MessageBox(Me, "此學員個人資料已存在，將更新您所輸入的資料")

                Dim j As Integer = MyTable1.Rows.Count - 1
                For p As Integer = 0 To j
                    Mydr1 = MyTable1.Rows(p)
                    ref_SID = Mydr1("SID") '隨意一組 SID '可能有多組。(合理情況下，應該只有1筆)
                    Call SUtl_AddUpdateRow1(Mydr1)

                    Dim htPP As New Hashtable
                    htPP.Clear()
                    htPP.Add("TransType", TIMS.cst_TRANS_LOG_Update)
                    htPP.Add("TargetTable", "STUD_STUDENTINFO")
                    htPP.Add("FuncPath", "/SD/03/SD_03_002_add2")
                    htPP.Add("s_WHERE", String.Format("IDNO='{0}' AND SID='{1}'", IDNO.Text, ref_SID))
                    TIMS.SaveTRANSLOG(sm, objconn, Mydr1, htPP)
                Next

                '更新學員基本資料檔
                DbAccess.UpdateDataTable(MyTable1, da1)
            End If
        End If

        'Try '測試。
        'Catch ex As Exception
        '    If Not objTrans Is Nothing Then DbAccess.RollbackTrans(objTrans)
        '    'Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
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
    ''' <summary>
    ''' UPDATE STUD_SUBDATA
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData2() As Boolean
        Dim rst As Boolean = False '儲存成功為True
        Dim all_Identity2 As String = ""
        all_Identity2 = Get_All_Identity2()

        Dim v_PassPortNO As String = TIMS.GetListValue(PassPortNO)

        'IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        'Dim dt9 As DataTable 'STUD_STUDENTINFO
        Dim hPMS9 As New Hashtable From {{"IDNO", IDNO.Text}}
        Dim sql9 As String = "SELECT SID FROM STUD_STUDENTINFO WHERE IDNO=@IDNO" '找出SID
        Dim dt9 As DataTable = DbAccess.GetDataTable(sql9, objconn, hPMS9)

        For z As Integer = 0 To dt9.Rows.Count - 1
            Dim Myda2 As SqlDataAdapter = Nothing
            Dim Mydt2 As DataTable = Nothing 'STUD_SUBDATA
            Dim Mydr2 As DataRow = Nothing 'STUD_SUBDATA
            Dim SID As String = Convert.ToString(dt9.Rows(z)("SID"))
            Dim sql2 As String = " SELECT * FROM STUD_SUBDATA WHERE SID='" & SID & "' "
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
            If CheckBox2.Checked = True Then '緊急聯絡人地址同通訊地址
                iTypeSame = 2
            ElseIf CheckBox3.Checked = True Then '緊急聯絡人地址同戶籍地址
                iTypeSame = 3
            End If

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
            Dim v_ShowDetail As String = TIMS.GetListValue(ShowDetail) '求才廠商查詢
            Mydr2("ShowDetail") = If(v_ShowDetail = "Y", "Y", "N")
            Mydr2("ServiceID") = TIMS.ClearSQM(ServiceID.Text) '軍種
            Mydr2("MilitaryAppointment") = TIMS.ClearSQM(MilitaryAppointment.Text) '職務(兵役) 
            Mydr2("MilitaryRank") = TIMS.ClearSQM(MilitaryRank.Text) '階級
            '服役日期 
            Mydr2("SServiceDate") = If(SServiceDate.Text <> "", TIMS.Cdate2(SServiceDate.Text), Convert.DBNull)
            Mydr2("FServiceDate") = If(FServiceDate.Text <> "", TIMS.Cdate2(FServiceDate.Text), Convert.DBNull)
            Mydr2("ServiceOrg") = TIMS.ClearSQM(ServiceOrg.Text) '服務單位名稱
            Mydr2("ChiefRankName") = TIMS.ClearSQM(ChiefRankName.Text) '主管階級 姓名

            ZipCode4.Value = TIMS.ClearSQM(ZipCode4.Value)
            ZipCode4_B3.Value = TIMS.ClearSQM(ZipCode4_B3.Value)
            hidZipCode4_6W.Value = TIMS.GetZIPCODE6W(ZipCode4.Value, ZipCode4_B3.Value)
            ZipCode4_N.Value = TIMS.ClearSQM(ZipCode4_N.Value)
            ServiceAddress.Text = TIMS.ClearSQM(ServiceAddress.Text) '服役單位 地址
            Mydr2("ZipCode4") = TIMS.GetValue1(ZipCode4.Value)
            Mydr2("ZipCode4_6W") = TIMS.GetValue1(hidZipCode4_6W.Value)
            Mydr2("ZipCode4_N") = TIMS.GetValue1(ZipCode4_N.Value)
            Mydr2("ServiceAddress") = TIMS.GetSpace(ServiceAddress.Text) '服役單位 地址
            Mydr2("ServicePhone") = TIMS.ClearSQM(ServicePhone.Text) '單位電話 

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

            '外國籍新增部分2005/12/16 Start
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
            '外國籍新增部分2005/12/16 End
            Mydr2("ModifyAcct") = sm.UserInfo.UserID
            Mydr2("ModifyDate") = Now()

            DbAccess.UpdateDataTable(Mydt2, Myda2)
            '更新學員資料副檔 End
        Next
        rst = True
        Return rst
    End Function

    ''' <summary>
    ''' UPDATE STUD_ENTERTEMP,STUD_ENTERTEMP2
    ''' </summary>
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
        ' statr更新Stud_EnterTemp
        sql = " SELECT * FROM STUD_ENTERTEMP WHERE IDNO='" & IDNO.Text & "' "
        MyTable4 = DbAccess.GetDataTable(sql, da4, objconn)
        If MyTable4.Rows.Count <> 0 Then
            For x As Integer = 0 To MyTable4.Rows.Count - 1
                mydr4 = MyTable4.Rows(x)
                mydr4("Name") = Name.Text
                mydr4("Sex") = v_Sex
                mydr4("Birthday") = CDate(TIMS.Cdate2(Birthday.Text))

                Select Case v_PassPortNO'.SelectedValue
                    Case "1", "2"
                        'PassPortNO 1:本國 /2:外籍(含大陸人士)
                        mydr4("PassPortNO") = v_PassPortNO 'PassPortNO.SelectedValue
                    Case Else
                        mydr4("PassPortNO") = "2"
                End Select

                Select Case MaritalStatus.SelectedValue
                    Case "1", "2"
                        mydr4("MaritalStatus") = MaritalStatus.SelectedValue
                    Case Else
                        mydr4("MaritalStatus") = Convert.DBNull
                End Select

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
                Dim v_rblMobil As String = TIMS.GetListValue(rblMobil)
                mydr4("CellPhone") = If(v_rblMobil = "Y", CellPhone.Text, "")

                mydr4("Email") = Email.Text
                mydr4("IsAgree") = "Y"
                DbAccess.UpdateDataTable(MyTable4, da4)
            Next
        End If

        'IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Dim MyTable5 As DataTable
        Dim da5 As SqlDataAdapter = Nothing
        Dim mydr5 As DataRow
        ' statr更新STUD_ENTERTEMP2
        sql = " SELECT * FROM STUD_ENTERTEMP2 WHERE IDNO='" & IDNO.Text & "' "
        MyTable5 = DbAccess.GetDataTable(sql, da5, objconn)
        If MyTable5.Rows.Count <> 0 Then
            For y As Integer = 0 To MyTable5.Rows.Count - 1
                mydr5 = MyTable5.Rows(y)
                mydr5("Name") = Name.Text
                mydr5("Sex") = v_Sex
                mydr5("Birthday") = CDate(TIMS.Cdate2(Birthday.Text))

                Select Case v_PassPortNO'.SelectedValue
                    Case "1", "2"
                        mydr5("PassPortNO") = v_PassPortNO '.SelectedValue
                    Case Else
                        mydr5("PassPortNO") = "2"
                End Select

                Select Case MaritalStatus.SelectedValue
                    Case "1", "2"
                        mydr5("MaritalStatus") = MaritalStatus.SelectedValue
                    Case Else
                        mydr5("MaritalStatus") = Convert.DBNull
                End Select

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

                mydr5("Phone1") = PhoneD.Text
                mydr5("Phone2") = PhoneN.Text
                mydr5("CellPhone") = ""

                Select Case rblMobil.SelectedValue
                    Case "Y"
                        mydr5("CellPhone") = CellPhone.Text
                End Select

                mydr5("Email") = Email.Text
                mydr5("IsAgree") = "Y"
                DbAccess.UpdateDataTable(MyTable5, da5)
            Next
        End If
        rst = True
        Return rst
    End Function

    '津貼類別  不等於 03:就業促進津貼實施辦法 刪除請領津貼的資料。
    'Sub UPDATE_SUBSIDYRESULT(ByRef vSOCID_Sel As String)
    '    '津貼類別  不等於 03:就業促進津貼實施辦法 刪除請領津貼的資料。
    '    If SubsidyID.SelectedValue <> "03" Then
    '        If vSOCID_Sel <> "" AndAlso StdTr.Visible=True Then
    '            Dim Parms As Hashtable=New Hashtable
    '            Parms.Clear()
    '            Parms.Add("SOCID", vSOCID_Sel)

    '            Dim sql As String=""
    '            sql="SELECT 'X' FROM dbo.STUD_SUBSIDYRESULT WHERE SOCID =@SOCID"
    '            Dim dt1 As DataTable=DbAccess.GetDataTable(sql, objconn, Parms)
    '            If dt1.Rows.Count=0 Then Return '無資料離開
    '            '有資料刪除
    '            sql=" DELETE STUD_SUBSIDYRESULT WHERE SOCID =@SOCID"
    '            DbAccess.ExecuteNonQuery(sql, objconn, Parms)
    '        End If
    '    End If
    'End Sub

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
                Dim Parms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"SOCID", hide_RejectSOCID.Value}}
                Dim sql As String = ""
                sql &= " UPDATE CLASS_STUDENTSOFCLASS"
                sql &= " SET MakeSOCID=NULL ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()"
                sql &= " WHERE SOCID=@SOCID "
                DbAccess.ExecuteNonQuery(sql, objconn, Parms)
            End If
        End If

        Return vRejectSOCID
    End Function

    Sub UPDATE_SERVICEPLACE(ByRef iSOCID As Integer)
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        sql = $" SELECT * FROM STUD_SERVICEPLACE WHERE SOCID={iSOCID} "
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count = 0 Then
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("SOCID") = iSOCID 'SOCIDValue
        Else
            dr = dt.Rows(0)
        End If

        'PostNo_1.Text=TIMS.ClearSQM(PostNo_1.Text)
        'AcctNo1_1.Text=TIMS.ClearSQM(AcctNo1_1.Text)
        'BankName.Text=TIMS.ClearSQM(BankName.Text)
        'AcctheadNo.Text=TIMS.ClearSQM(AcctheadNo.Text)
        'ExBankName.Text=TIMS.ClearSQM(ExBankName.Text)
        'AcctExNo.Text=TIMS.ClearSQM(AcctExNo.Text)
        'AcctNo2.Text=TIMS.ClearSQM(AcctNo2.Text)
        'Dim v_AcctMode As String=TIMS.GetListValue(AcctMode)
        'Select Case v_AcctMode'.SelectedValue
        '    Case "0"
        '        dr("AcctMode")=0
        '        dr("PostNo")=PostNo_1.Text '& "-" & PostNo_2.Text
        '        dr("AcctNo")=AcctNo1_1.Text '& "-" & AcctNo1_2.Text
        '        dr("BankName")=Convert.DBNull
        '        dr("AcctHeadNo")=Convert.DBNull
        '        dr("ExBankName")=Convert.DBNull
        '        dr("AcctExNo")=Convert.DBNull
        '    Case "1"
        '        dr("AcctMode")=1
        '        dr("PostNo")=Convert.DBNull
        '        dr("BankName")=BankName.Text
        '        dr("AcctHeadNo")=AcctheadNo.Text
        '        dr("ExBankName")=ExBankName.Text
        '        dr("AcctExNo")=AcctExNo.Text
        '        dr("AcctNo")=AcctNo2.Text
        '    Case "2"
        '        '**by Milor 20080509--由訓練單位代轉現金時，所有轉帳資料都不填入值----start
        '        dr("AcctMode")=2
        '        dr("PostNo")=Convert.DBNull
        '        dr("AcctNo")=Convert.DBNull
        '        dr("BankName")=Convert.DBNull
        '        dr("AcctHeadNo")=Convert.DBNull
        '        dr("ExBankName")=Convert.DBNull
        '        dr("AcctExNo")=Convert.DBNull
        '        '**by Milor 20080509----end
        'End Select
        'FirDate.Text=TIMS.ClearSQM(FirDate.Text)
        'dr("FirDate")=If(FirDate.Text <> "", TIMS.cdate2(FirDate.Text), Convert.DBNull)

        Uname.Text = TIMS.ClearSQM(Uname.Text) '目前任職公司名稱
        Intaxno.Text = TIMS.ClearSQM(Intaxno.Text) '公司統一編號
        ActName.Text = TIMS.ClearSQM(ActName.Text)
        ActNo1.Text = TIMS.ClearSQM(ActNo1.Text)

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
        'Tel.Text=TIMS.ClearSQM(Tel.Text)
        'Fax.Text=TIMS.ClearSQM(Fax.Text)
        'dr("Tel")=If(Tel.Text <> "", Tel.Text, Convert.DBNull) 'Tel.Text
        'dr("Fax")=If(Fax.Text <> "", Fax.Text, Convert.DBNull)
        'SDate.Text=TIMS.ClearSQM(SDate.Text)
        'SJDate.Text=TIMS.ClearSQM(SJDate.Text)
        'SPDate.Text=TIMS.ClearSQM(SPDate.Text)
        'dr("SDate")=TIMS.cdate2(SDate.Text)
        'dr("SJDate")=TIMS.cdate2(SJDate.Text)
        'dr("SPDate")=TIMS.cdate2(SPDate.Text)

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

        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    'Sub UPDATE_TRAINBG(ByRef iSOCID As Integer)
    '    Dim v_Q1 As String=TIMS.GetListValue(Q1)
    '    Dim v_Q3 As String=TIMS.GetListValue(Q3)
    '    Q3_Other.Text=TIMS.ClearSQM(Q3_Other.Text)
    '    Dim v_Q4 As String=TIMS.GetListValue(Q4)
    '    Dim v_Q5 As String=TIMS.GetListValue(Q5)

    '    Dim sql As String=""
    '    Dim dr As DataRow=Nothing
    '    Dim dt As DataTable=Nothing
    '    Dim da As SqlDataAdapter=Nothing
    '    sql=" SELECT * FROM dbo.STUD_TRAINBG WHERE SOCID='" & iSOCID & "' "
    '    dt=DbAccess.GetDataTable(sql, da, objconn)
    '    If dt.Rows.Count=0 Then
    '        dr=dt.NewRow
    '        dt.Rows.Add(dr)
    '        dr("SOCID")=iSOCID
    '    Else
    '        dr=dt.Rows(0)
    '    End If

    '    dr("Q1")=Val(If(v_Q1 <> "", v_Q1, "0")) 'Val(Q1.SelectedValue)'numeric
    '    dr("Q3")=If(v_Q3="", Convert.DBNull, Val(v_Q3)) 'numeric
    '    dr("Q3_Other")=If(Q3_Other.Text="", Convert.DBNull, Q3_Other.Text)
    '    dr("Q4")=If(v_Q4="", Convert.DBNull, v_Q4) 'varchar
    '    dr("Q5")=If(v_Q5="", Convert.DBNull, Val(v_Q5)) 'numeric

    '    dr("Q61")=If(Q61.Text="", Convert.DBNull, Val(Q61.Text))
    '    dr("Q62")=If(Q62.Text="", Convert.DBNull, Val(Q62.Text))
    '    dr("Q63")=If(Q63.Text="", Convert.DBNull, Val(Q63.Text))
    '    dr("Q64")=If(Q64.Text="", Convert.DBNull, Val(Q64.Text))
    '    dr("ModifyAcct")=sm.UserInfo.UserID
    '    dr("ModifyDate")=Now
    '    DbAccess.UpdateDataTable(dt, da)
    'End Sub

    'Sub UPDATE_TRAINBGQ2(ByRef iSOCID As Integer)
    '    Dim sql As String=""
    '    Dim dr As DataRow=Nothing
    '    Dim dt As DataTable=Nothing
    '    Dim da As SqlDataAdapter=Nothing

    '    If iSOCID=0 Then Return
    '    sql=" SELECT * FROM STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' "
    '    dt=DbAccess.GetDataTable(sql, objconn)
    '    If dt.Rows.Count > 0 Then
    '        sql=" DELETE STUD_TRAINBGQ2 WHERE SOCID='" & iSOCID & "' "
    '        DbAccess.ExecuteNonQuery(sql, objconn)
    '    End If
    '    Dim v_Q2 As String=TIMS.GetCblValue(Q2)
    '    If v_Q2="" Then Return '無值離開

    '    sql=" INSERT INTO STUD_TRAINBGQ2 (SOCID,Q2) VALUES (@SOCID,@Q2)" & vbCrLf
    '    For Each item As ListItem In Q2.Items
    '        If item.Value <> "" AndAlso item.Selected Then
    '            Dim Parms As New Hashtable
    '            Parms.Clear()
    '            Parms.Add("SOCID", iSOCID)
    '            Parms.Add("Q2", item.Value)
    '            DbAccess.ExecuteNonQuery(sql, objconn, Parms)
    '        End If
    '    Next
    'End Sub

    ''' <summary> 基本資料回填報名資料 </summary>
    ''' <param name="iSOCID"></param>
    Sub UPDATE_ENTERTEMP12(ByRef iSOCID As Integer)
        'Dim sql As String="" Dim dt As DataTable=Nothing
        'Dim dr As DataRow=Nothing
        'Dim da As SqlDataAdapter=Nothing

        Dim v_Sex As String = TIMS.GetListValue(Sex)
        '**by Milor 20080527--產學訓當姓名、性別、生日、身分證號改變時，要回填報名資料----start
        '資料驗證的流程:
        '1.Class_StudentOfClass取得SETID、ETEnterDate、SerNum來對應Stud_EnterType，
        '  以驗證存在有報名職類檔，才去回填報名資料，避免回填到有問題資料。
        '2.E網報名資料STUD_ENTERTEMP2的SETID，因為某些User操作錯誤的因素等，
        '  導致有重複的SETID、IDNO、Name等資料，所以必須符合SETID、IDNO、Name、Sex、Birthday時，
        '  才將修改的資料回填。
        Dim pms3 As New Hashtable From {{"SOCID", iSOCID}}
        Dim sql3 As String = ""
        sql3 &= " SELECT SETID ,CONVERT(VARCHAR, ETEnterDate, 111) ETEnterDate ,SerNum" & vbCrLf
        sql3 &= " FROM CLASS_STUDENTSOFCLASS" & vbCrLf
        sql3 &= " WHERE SOCID=@SOCID AND SETID IS NOT NULL AND ETEnterDate IS NOT NULL AND SerNum IS NOT NULL "
        Dim dt As DataTable = DbAccess.GetDataTable(sql3, objconn, pms3)
        If dt.Rows.Count = 0 Then Return

        Dim dr As DataRow = dt.Rows(0)
        'Common.FormatDate()
        Dim uSETID As String = Convert.ToString(dr("SETID"))
        Dim uETED As String = Convert.ToString(dr("ETEnterDate"))
        Dim uSerNum As String = Convert.ToString(dr("SerNum"))
        '先驗證有報名職類檔，才進行檢查報名資料
        Dim pms4 As New Hashtable From {{"SETID", uSETID}, {"EnterDate", uETED}, {"SerNum", uSerNum}}
        Dim sql4 As String = " SELECT SETID FROM STUD_ENTERTYPE WHERE SETID=@SETID AND EnterDate=@EnterDate AND SerNum=@SerNum"
        Dim dt4 As DataTable = DbAccess.GetDataTable(sql4, objconn, pms4)
        If dt4.Rows.Count = 0 Then Return

        Dim da As SqlDataAdapter = Nothing
        '取出報名資料時，同時取出身分證號、姓名、性別、生日，以備後續判斷E網報名資料暫存檔
        Dim SqlS As String = " SELECT * FROM STUD_ENTERTEMP WHERE SETID='" & uSETID & "' "
        dt = DbAccess.GetDataTable(SqlS, da, objconn)
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
        DbAccess.UpdateDataTable(dt, da) 'STUD_ENTERTEMP

        '透過SETID取出E網報名資料暫存檔，比對身分證號、姓名、性別、生日與報名資料未更動前的資料一致時才變更
        'sql="select * from STUD_ENTERTEMP2 where SETID='" & uSETID & "' and IDNO='" & uIDNO & "' and Name='" & uName & "' and Sex='" & uSex & "' and Birthday='" & Convert.ToDateTime(uBirthday) & "'"
        'sql="select * from STUD_ENTERTEMP2 where SETID='" & uSETID & "' and upper(IDNO)='" & uIDNO & "' and Name='" & uName & "' and Sex='" & uSex & "' and Birthday='" & uBirthday & "'"
        uSETID = TIMS.ClearSQM(uSETID)
        uIDNO = TIMS.ClearSQM(uIDNO)
        uName = TIMS.ClearSQM(uName)
        uSex = TIMS.ClearSQM(uSex)
        uBirthday = TIMS.ClearSQM(uBirthday)
        Dim SqlT2 As String = ""
        SqlT2 &= " SELECT * FROM STUD_ENTERTEMP2 "
        SqlT2 &= " WHERE SETID='" & uSETID & "' AND IDNO='" & uIDNO & "'"
        SqlT2 &= " AND Name='" & uName & "' AND Sex='" & uSex & "' AND Birthday=" & TIMS.To_date(uBirthday) 'fix ORA-01861
        da = Nothing
        dt = DbAccess.GetDataTable(SqlT2, da, objconn)
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

    ''' <summary>
    ''' UPDATE CLASS_STUDENTSOFCLASS
    ''' </summary>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function SUtl_SaveData4() As Boolean
        Dim rst As Boolean = False '儲存成功為True

        Dim SID As String = TIMS.ClearSQM(Hid_SID_C1.Value)
        If SID = "" Then Return False

        Dim all_Identity2 As String = Get_All_Identity2()

        Dim v_DegreeID As String = TIMS.GetListValue(DegreeID)
        Dim v_MilitaryID As String = TIMS.GetListValue(MilitaryID)
        Dim vSOCID_Sel As String = TIMS.GetListValue(SOCID)
        If vSOCID_Sel = "" Then Return False

        '津貼類別  不等於 03:就業促進津貼實施辦法 刪除請領津貼的資料。
        'Call UPDATE_SUBSIDYRESULT(vSOCID_Sel)
        '更新MakeSOCID
        Dim vRejectSOCID As String = UPDATE_MakeSOCID()

        Dim sql As String = ""
        'UPDATE CLASS_STUDENTSOFCLASS
        Dim iSOCID As Integer = 0 'String=""
        If vSOCID_Sel <> "" AndAlso vSOCID_Sel <> "0" Then iSOCID = Val(vSOCID_Sel)
        Dim bflagNew1 As Boolean = False '修改動作
        If StdTr.Visible = False OrElse iSOCID = 0 Then bflagNew1 = True '新增一筆
        If Not bflagNew1 Then '修改檢核
            sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID='" & iSOCID & "' "
            Dim dtCS As DataTable = DbAccess.GetDataTable(sql, objconn)
            If dtCS.Rows.Count <> 1 Then bflagNew1 = True '新增一筆 '資料不等於1筆(異常新增)
        End If

        Dim v_EnterChannel As String = TIMS.GetListValue(EnterChannel)
        Dim v_TRNDMode As String = TIMS.GetListValue(TRNDMode)

        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        Dim Mydr3 As DataRow = Nothing
        Dim MyTable3 As DataTable = Nothing
        Dim da3 As SqlDataAdapter = Nothing
        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update
        If bflagNew1 Then
            s_TransType = TIMS.cst_TRANS_LOG_Insert
            iSOCID = DbAccess.GetNewId(objconn, "CLASS_STUDENTSOFCLASS_SOCID_SE,CLASS_STUDENTSOFCLASS,SOCID")
            '新增
            sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE 1<>1 "
            'CLASS_STUDENTSOFCLASS_SOCID_SE
            MyTable3 = DbAccess.GetDataTable(sql, da3, objconn)
            '查無資料新增1筆
            Mydr3 = MyTable3.NewRow
            MyTable3.Rows.Add(Mydr3)
            Mydr3("SOCID") = iSOCID
            Mydr3("SID") = SID '取得目前最新的SID
            Mydr3("StudStatus") = 1
            Mydr3("OCID") = Val(rqOCID)
        Else
            '修改
            'UPDATE CLASS_STUDENTSOFCLASS
            sql = " SELECT * FROM CLASS_STUDENTSOFCLASS WHERE SOCID='" & iSOCID & "'"
            'iSOCIDValue=SOCID.SelectedValue
            MyTable3 = DbAccess.GetDataTable(sql, da3, objconn)
            Mydr3 = MyTable3.Rows(0)
            Mydr3("SID") = SID '取得目前最新的SID
        End If

        '更新班級學員檔-----   Start
        Mydr3("RejectSOCID") = If(vRejectSOCID <> "", vRejectSOCID, Convert.DBNull) '被遞補者 學員
        Mydr3("StudentID") = StudentIDValue.Value & StudentID.Text
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

        'Dim v_SubsidyID As String="" 'TIMS.GetListValue(SubsidyID)
        'Dim v_SubsidyIdentity As String="" 'TIMS.GetListValue(SubsidyIdentity)
        'If trSubsidyID.Visible Then
        '    v_SubsidyID=TIMS.GetListValue(SubsidyID)
        '    '預設 未申請:01 (在職進修訓練:06)
        '    If v_SubsidyID="" AndAlso sm.UserInfo.TPlanID="06" Then v_SubsidyID="01"
        'End If
        'If trSubsidyIdentity.Visible Then v_SubsidyIdentity=TIMS.GetListValue(SubsidyIdentity)
        'Mydr3("SubsidyID")=If(v_SubsidyID <> "", v_SubsidyID, Convert.DBNull)
        'Mydr3("SubsidyIdentity")=If(v_SubsidyIdentity <> "", v_SubsidyIdentity, Convert.DBNull)

        ' 受訓前任職資料start 2010/04/27 開始存class_classinfo 
        If trPriorWorkOrg1.Visible AndAlso trTable6.Visible Then
            PriorWorkOrg1.Text = TIMS.ClearSQM(PriorWorkOrg1.Text)
            If Len(PriorWorkOrg1.Text) > 30 Then PriorWorkOrg1.Text = Mid(PriorWorkOrg1.Text, 1, 30)
            Dim v_PriorWorkType1 As String = TIMS.GetListValue(PriorWorkType1)
            Mydr3("PWType1") = If(v_PriorWorkType1 <> "", v_PriorWorkType1, Convert.DBNull)
            Mydr3("PWOrg1") = If(PriorWorkOrg1.Text <> "", PriorWorkOrg1.Text, Convert.DBNull)
            Mydr3("SOfficeYM1") = If(SOfficeYM1.Text <> "", SOfficeYM1.Text, Convert.DBNull)
            Mydr3("FOfficeYM1") = If(FOfficeYM1.Text <> "", FOfficeYM1.Text, Convert.DBNull)
        End If
        ' 受訓前任職資料end 

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

        Dim v_BudID As String = TIMS.GetListValue(BudID)
        Mydr3("BudgetID") = If(v_BudID <> "", v_BudID, Convert.DBNull)
        'SupplyID 0: 請選擇 ,1: 一般80% ,2: 特定100% ,9: 0%
        'Dim v_SupplyID As String=TIMS.GetListValue(SupplyID)
        'If v_SupplyID="0" Then v_SupplyID="" '等於0，等同未選擇
        'If v_BudID="99" Then v_SupplyID="9" '不補助 0%
        'Mydr3("SupplyID")=If(v_SupplyID <> "", v_SupplyID, Convert.DBNull)
        Mydr3("SupplyID") = Convert.DBNull
        Dim v_AppliedResult As String = Convert.ToString(Mydr3("AppliedResult"))
        'If v_BudID="99" AndAlso v_SupplyID="9" Then v_AppliedResult="N"  '學員資料複審狀態=N
        If v_BudID = "99" Then v_AppliedResult = "N"  '學員資料複審狀態=N
        Mydr3("AppliedResult") = If(v_AppliedResult <> "", v_AppliedResult, Convert.DBNull) '學員資料複審狀態=N

        ''選錯為 2: 特定100% , 且為一般身分, 且為 就安 就保。
        'If SupplyID.SelectedValue="2" AndAlso MIdentityID.SelectedValue="01" AndAlso all_Identity2="01" Then Mydr3("SupplyID")="1"
        'Dim v_PMode As String="" 'TIMS.GetListValue(PMode)
        'If trPMode.Visible Then v_PMode=TIMS.GetListValue(PMode)
        'Mydr3("PMode")=If(v_PMode <> "", v_PMode, Convert.DBNull)

        Dim v_ActNo As String = ""
        'ActNo.Text=TIMS.ClearSQM(TIMS.ChangeIDNO(ActNo.Text))
        ActNo1.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(ActNo1.Text))
        ActNo2.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(ActNo2.Text))
        'If TPlan23TR.Visible Then
        '    v_ActNo=If(ActNo.Text <> "", ActNo.Text, "")
        'Else
        '    'ActNo1 產投 ActNo2 '其他一般計畫
        '    v_ActNo=If(ActNo1.Text <> "", ActNo1.Text, If(ActNo2.Text <> "", ActNo2.Text, ""))
        'End If
        'ActNo1 產投 ActNo2 '其他一般計畫
        v_ActNo = If(ActNo1.Text <> "", ActNo1.Text, If(ActNo2.Text <> "", ActNo2.Text, ""))
        Mydr3("ActNo") = If(v_ActNo <> "", v_ActNo, Convert.DBNull)

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
            Dim htPP As New Hashtable
            htPP.Clear()
            htPP.Add("TransType", s_TransType)
            htPP.Add("TargetTable", "CLASS_STUDENTSOFCLASS")
            htPP.Add("FuncPath", "/SD/03/SD_03_002_add2")
            htPP.Add("s_WHERE", String.Format("SOCID={0}", iSOCID))
            TIMS.SaveTRANSLOG(sm, objconn, Mydr3, htPP)

            DbAccess.UpdateDataTable(MyTable3, da3)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("SID:{0},IDNO:{1},rqOCID:{2},iSOCID:{3}", SID, IDNO.Text, Val(rqOCID), iSOCID) & vbCrLf
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
            Dim u_sql As String = ""
            u_sql = ""
            u_sql &= " UPDATE STUD_SUBSIDYCOST "
            u_sql &= " SET BUDID=@BUDID ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE() "
            u_sql &= " WHERE SOCID=@SOCID "
            Dim u_Parms As New Hashtable
            u_Parms.Clear()
            u_Parms.Add("BUDID", v_BudID)
            u_Parms.Add("MODIFYACCT", sm.UserInfo.UserID)
            u_Parms.Add("SOCID", iSOCID)
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_Parms)  'edit，by:20181017
        End If

        '企訓專用 Start
        '46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        '是否為在職者補助身分(rblWorkSuppIdent) 選 Y 可存取
        'If sm.UserInfo.TPlanID="28" Or (sm.UserInfo.TPlanID="46" And rblWorkSuppIdent.SelectedValue="Y") Or (sm.UserInfo.TPlanID="47" And rblWorkSuppIdent.SelectedValue="Y") Then
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        If Hid_show_actno_budid.Value = "Y" Then
            Call UPDATE_SERVICEPLACE(iSOCID)
            'Call UPDATE_TRAINBG(iSOCID)
            'Call UPDATE_TRAINBGQ2(iSOCID)
            Call UPDATE_ENTERTEMP12(iSOCID)
        End If
        '企訓專用 End

        '結訓學員資料卡更新 Start
        If StdTr.Visible = True Then
            da = Nothing
            sql = " SELECT * FROM STUD_RESULTSTUDDATA WHERE SOCID='" & vSOCID_Sel & "' "
            dt = DbAccess.GetDataTable(sql, da, objconn)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                Dim DLID As Integer = dr("DLID")
                Dim SubNo As Integer = dr("SubNo")
                dr("StdName") = Name.Text
                dr("StudentID") = StudentID.Text
                dr("StdPID") = IDNO.Text 'TIMS.ChangeIDNO(IDNO.Text)
                dr("Sex") = Sex.SelectedIndex + 1 '1.男2.女
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
                sql = "SELECT * FROM Stud_ResultIdentData WHERE DLID='" & DLID & "' and SubNo='" & SubNo & "'"
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
            'UPDATE STUD_SELRESULTBLI BY AMU 2022
            Dim uSETID As String = dr("SETID").ToString
            Dim uETED As String = dr("ETEnterDate").ToString 'yyyy/MM/dd
            Dim uSerNum As String = dr("SerNum").ToString
            Hid_SETID.Value = uSETID
            Hid_ETENTERDATE.Value = uETED 'yyyy/MM/dd
            Hid_SERNUM.Value = uSerNum
        End If
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

    Function SUtl_SaveData5(ByVal str_SOCID As String) As Boolean
        '首頁>>學員動態管理>>表單列印>>參訓學員投保狀況檢核表
        '依專案執行檢討會議決議（附件1），學員資料維護之「保險證號」與「投保單位名稱」
        '將與參訓學員投保狀況檢核表勾稽到的「投保單位」及「保險證號」連動，若單位於學員資料維護勾選其他「保險證號」
        '即會一併連動修改參訓學員投保狀況檢核表之「投保單位」及「保險證號」。
        '目前有參訓學員共有2個投保證號， 於學員資料維護勾選第2個投保證號後（附件2）
        '至參訓學員投保狀況檢核表列印時並未將該學員投保證號連動修正（附件3）， 請儘速修正。  
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
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
            If hidSB4ID.Value <> "" AndAlso ActNo1.Text <> "" Then
                Dim hssT As New Hashtable From {
                    {"ACTNO1", ActNo1.Text},
                    {"SB4ID", hidSB4ID.Value},
                    {"IDNO", IDNO.Text},
                    {"OCID1", Hid_OCID.Value},
                    {"SOCID", str_SOCID},
                    {"SETID", Hid_SETID.Value},
                    {"ETENTERDATE", Hid_ETENTERDATE.Value},
                    {"SERNUM", Hid_SERNUM.Value},
                    {"USERID", sm.UserInfo.UserID}
                }
                Call TIMS.UPDATE_STUD_SELRESULTBLI(hssT, objconn)
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

    ''' <summary>
    ''' 取得原資料
    ''' </summary>
    ''' <returns></returns>
    Function Get_STUDINFOdr() As DataRow
        Dim dr As DataRow = Nothing
        IDNO.Text = TIMS.ClearSQM(TIMS.ChangeIDNO(IDNO.Text))
        If IDNO.Text = "" Then Return dr

        Dim sql As String = ""
        sql = "" & vbCrLf
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
        sql &= " WHERE a.IDNO='" & IDNO.Text & "'" & vbCrLf
        sql &= " ORDER BY cs.SOCID DESC" & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)

        Return dr
    End Function

    ''' <summary>
    ''' 取得原資料-顯示
    ''' </summary>
    Sub SHOW_STUDINFO(ByRef dr As DataRow)
        If dr Is Nothing Then Return

        Name.Text = dr("Name").ToString
        RMPNAME.Text = dr("RMPNAME").ToString
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

        ' ZipCode4  服役單位地址
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode4")), objconn)
        ZipCode4.Value = Convert.ToString(dr("ZipCode4"))
        hidZipCode4_6W.Value = Convert.ToString(dr("ZipCode4_6W"))
        ZipCode4_B3.Value = TIMS.GetZIPCODEB3(hidZipCode4_6W.Value)
        ZipCode4_N.Value = Convert.ToString(dr("ZipCode4_N"))
        City4.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode4")), Convert.ToString(dr("ZipCode4_N")), objconn)
        City4.Text &= If(tZipLName <> "", "[" & tZipLName & "]", "")
        ServiceAddress.Text = HttpUtility.HtmlDecode(Convert.ToString(dr("ServiceAddress")))
        Hid_JnZipCode4.Value = TIMS.GetZipCodeJn(ZipCode4.Value, ZipCode4_B3.Value, hidZipCode4_6W.Value, City4.Text, ServiceAddress.Text)

        PhoneD.Text = dr("PhoneD").ToString
        PhoneN.Text = dr("PhoneN").ToString
        CellPhone.Text = Convert.ToString(dr("CellPhone"))
        If CellPhone.Text <> "" Then CellPhone.Text = Trim(CellPhone.Text)

        Dim vMobil As String = TIMS.cst_NO
        If CellPhone.Text <> "" Then vMobil = TIMS.cst_YES
        Common.SetListItem(rblMobil, vMobil)

        ' ZipCode1  通訊地址
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode1")), objconn)
        ZipCode1.Value = Convert.ToString(dr("ZipCode1"))
        hidZipCode1_6W.Value = Convert.ToString(dr("ZipCode1_6W"))
        ZipCode1_B3.Value = TIMS.GetZIPCODEB3(hidZipCode4_6W.Value)
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
        If Convert.ToString(dr("HandTypeID")) <> "" AndAlso Convert.ToString(dr("HandLevelID")) <> "" Then
            flag_HandType = 1 '1:舊制
        End If
        If Convert.ToString(dr("HandTypeID2")) <> "" AndAlso Convert.ToString(dr("HandLevelID2")) <> "" Then
            flag_HandType = 2 '2:新制
        End If

        Select Case flag_HandType
            Case 1 '1:舊制
                trHandTypeID2.Style("display") = cst_none1 '新制
                trHandTypeID.Style("display") = cst_inline1 '舊制
                Common.SetListItem(rblHandType, "1")
            Case Else '0:未選 2:新制
                trHandTypeID2.Style("display") = cst_inline1 '新制
                trHandTypeID.Style("display") = cst_none1 '舊制
                Common.SetListItem(rblHandType, "2")
        End Select

        EmergencyContact.Text = dr("EmergencyContact").ToString
        EmergencyPhone.Text = dr("EmergencyPhone").ToString
        EmergencyRelation.Text = dr("EmergencyRelation").ToString

        ' ZipCode3  緊急通知人地址 
        tZipLName = TIMS.Get_ZipLName(Convert.ToString(dr("ZipCode3")), objconn)
        ZipCode3.Value = Convert.ToString(dr("ZipCode3"))
        hidZipCode3_6W.Value = Convert.ToString(dr("ZipCode3_6W"))
        ZipCode3_B3.Value = TIMS.GetZIPCODEB3(hidZipCode2_6W.Value)
        ZipCode3_N.Value = Convert.ToString(dr("ZipCode3_N"))
        City3.Text = TIMS.Get_ZipNameN(Convert.ToString(dr("ZipCode3")), Convert.ToString(dr("ZipCode3_N")), objconn)
        City3.Text &= If(tZipLName <> "", "[" & tZipLName & "]", "")
        EmergencyAddress.Text = Convert.ToString(dr("EmergencyAddress"))
        Hid_JnZipCode3.Value = TIMS.GetZipCodeJn(ZipCode3.Value, ZipCode3_B3.Value, hidZipCode3_6W.Value, City3.Text, EmergencyAddress.Text)
        'CtID4.Value=TIMS.Get_Ctid(Convert.ToString(dr("ZipCode3")), objconn)

        '受訓前任職資料 
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
                If DateDiff(DateInterval.Day, CDate(Cst_20110415), CDate(STDateHidden.Value)) >= 0 Then
                    If TIMS.CheckIsECFA(Me, dr("ActNo2"), "", STDateHidden.Value, objconn) = True Then Common.SetListItem(BudID, "97") '2011/05/20 新增ECFA判斷
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
        ' 受訓前任職資料 

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
        'Common.SetListItem(SupplyID, "0")
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
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            '政府已補助經費   Start
            If drv("SumOfMoney").ToString <> "" Then
                Dim SumOfMoney As Label = e.Item.FindControl("SumOfMoney")
                SumOfMoney.Text = drv("SumOfMoney").ToString
                If CInt(SumOfMoney.Text) >= 40000 Then
                    SumOfMoney.ForeColor = Color.Red 'SumOfMoney.ForeColor.Red '超過4萬的提示，將字變為紅色的
                    TIMS.Tooltip(SumOfMoney, "紅色為超過４萬之提醒", True)
                    SumOfMoney.Font.Bold = True
                End If
                'hide_SumOfMoney.Value &= CInt(SumOfMoney.Text)
            End If
            '政府已補助經費   End
        End If
    End Sub

    '回上一頁
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        'If ViewState("search") <> "" Then
        '    Session("search")=ViewState("search")
        'End If
        If Not Session(vs_SearchStr) Is Nothing Then
            ViewState(vs_SearchStr) = Session(vs_SearchStr)
            Session(vs_SearchStr) = ViewState(vs_SearchStr)
            'Session(vs_SearchStr)=Nothing
        End If

        Select Case Request("todo")
            Case 1
                Call TIMS.Utl_Redirect(Me, objconn, "../03/SD_03_002_classver.aspx?ID=" & Request("ID") & "&todo=1" & "&OCID=" & rqOCID)
            Case 2
                'If Not ViewState(vs_SearchStr) Is Nothing Then
                '    Session(vs_SearchStr)=ViewState(vs_SearchStr)
                'End If
                Call TIMS.Utl_Redirect(Me, objconn, "../03/SD_03_002.aspx?ID=" & Request("ID") & "&todo=2" & "&OCID=" & rqOCID)
            Case Else
                Dim url As String = TIMS.GetFunIDUrl(Request("ID"), 0, objconn)
                Call TIMS.Utl_Redirect(Me, objconn, url & "?ID=" & Request("ID"))
        End Select
    End Sub

#Region "NO USE"
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

    '異動學員基本資料。(STUD_STUDENTINFO)
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

        ''14天後可修改預算別，不限定" & cst_ECFA & "助。
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

    '學習卷時數排序label
    'Function Get_DGTHourDT() As DataTable
    '    Dim rst As DataTable
    '    'Dim sql As String
    '    'Dim dtDGHR As DataTable=Get_DGTHourDT()
    '    ''sql="SELECT * FROM Key_DGTHour ORDER BY DGID"
    '    'sql="SELECT * FROM Key_DGTHour ORDER BY DGhour" 'chk RelClass_Unit
    '    'dtDGHR=DbAccess.GetDataTable(sql, objConn) '目前系統有4筆資料。
    '    Dim sql As String="SELECT * FROM KEY_DGTHOUR ORDER BY DGHOUR" 'chk RelClass_Unit
    '    rst=DbAccess.GetDataTable(sql, objconn) '目前系統有4筆資料。
    '    Return rst
    'End Function

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
        Dim Usql As String = ""
        Usql &= " UPDATE CLASS_STUDENTSOFCLASS" & vbCrLf
        Usql &= " SET WorkSuppIdent=@WorkSuppIdent" & vbCrLf
        Usql &= " WHERE SOCID=@SOCID AND OCID=@OCID" & vbCrLf
        'u_Parms.Clear()
        Dim u_Parms As New Hashtable From {
            {"WorkSuppIdent", v_WorkSuppIdent},
            {"SOCID", SOCID},
            {"OCID", OCID}
        }
        DbAccess.ExecuteNonQuery(Usql, tConn, u_Parms)  'edit，by:20181017
    End Sub

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


End Class