Imports Microsoft.Security.Application

Partial Class TC_03_006
    Inherits AuthBasePage

    'SELECT * FROM PLAN_PLANINFO WHERE RID ='B5703' AND SEQNO=17 
    'PLAN_DEPOT,PLAN_ABILITY / PLAN_ONCLASS /PLAN_VERREPORT /PLAN_TEACHER /PLAN_TRAINDESC /PLAN_PERSONCOST /PLAN_COMMONCOST /PLAN_SHEETCOST /PLAN_OTHERCOST 
    'SELECT * FROM TEACH_TEACHERINFO WHERE RID ='B5703'
    'PLAN_TRAINPLACE /ORG_REMOTER

    '上課時間／時間內容，長度超過限制範圍{0}文字長度
    'Const cst_i_Times_c_max_length As Integer = 100 'by AMU 20191205
    Const cst_i_Times_c_max_length As Integer = 200 'by AMU 20220310
    Const cst_i_Times_c_min_length As Integer = 5

    '技檢訓練時數 '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    Const cst_EHour_t1 As String = "技檢訓練時數,必須大於0，若為0則毋須填寫,目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時可儲存，若不符合上述條件，不會存入資料庫。"
    '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_Use_TMID As String = "672"

    '三、基本儲存、正式儲存卡控充電起飛計畫不用這些卡控
    '1、「AI加值應用、職場續航」僅限申請階段為「政策性產業」，當單位有勾選，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」，若未填寫須出現提示訊息，且不能儲存。
    '2、當勾選「AI加值應用」時，須卡控AI應用時數之總數須等於或大於12小時，且不可超過總訓練時數1/2，若不符須出現提示訊息，且不能儲存。
    '3、當勾選「職場續航」時，須卡控職場續航時數之總數須等於6小時，不可小於或大於，若不符須出現提示訊息，且不能儲存。
    '4、因單位可能同時勾選「AI加值應用」及「職場續航」，且同一堂課可能會有AI應用時數及職場續航時數，各自時數仍須分開計算。
    '5、因政策性產業新增「職場續航」項目，此類政策性產業有規定特定訓練業別才能申請（詳附件檔案），須卡控班級申請之訓練業別代碼倘非規定之代碼，則跳出提示訊息【申請之訓練業別非「職場續航」課程職類】且不可儲存。
    Const cst_D25_8_CapAll_MSG As String = "班級招生以「工作15年以上年滿55歲者」、「工作25年以上中高齡者及高齡者」、「工作10年以上年滿60歲者」及「符合勞動基準法第54條第1項第1款所定得強制退休年齡前2年內之63-64歲在職中高齡者」等之在職勞工為優先錄訓，如未能招足額時，得開放最多30%名額給予其他一般在職身分者參訓。"

    Const cst_AIAHOUR_t1 As String = "「AI應用時數」,必須大於0，若為0則毋須填寫，AI應用時數之總數須等於或大於12小時，且不可超過總訓練時數1/2，若不符上述條件，不能儲存。"
    Const cst_AIAHOUR_msg1 As String = "政府政策性產業-勾選「AI加值應用」，AI應用時數之總數須等於或大於12小時，且不可超過總訓練時數1/2!"

    Const cst_WNLHOUR_t1 As String = "「職場續航時數」,必須大於0，若為0則毋須填寫，職場續航時數之總數須等於6小時，不可小於或大於，若不符條件，不能儲存。"
    Const cst_WNLHOUR_msg1 As String = "政府政策性產業-勾選「職場續航」，職場續航時數之總數須等於6小時，不可小於或大於，若不符條件，不能儲存!"

    '5、政策性產業「職場續航」項目，此類政策性產業有規定特定訓練業別才能申請（詳附件檔案），卡控班級申請之訓練業別代碼倘非規定之代碼，則跳出提示訊息【申請之訓練業別非「職場續航」課程職類】且不可儲存。
    'Const cst_WNLHOUR_t5 As String = "政策性產業「職場續航」項目，此類政策性產業有規定特定訓練業別才能申請-符合時可儲存，若不符條件，不可儲存。"
    'Const cst_WNLHOUR_msg5 As String = "政府政策性產業-勾選「職場續航」，【申請之訓練業別非「職場續航」課程職類】!"

    'TMIDCORRECT:若訓練業別有誤是否同意協助重新歸類
    Dim str_TMIDCORRECT_t As String = String.Concat("訓練業別同意協助", "<br/>", "重新歸類")
    Const cst_TMIDCORRECT_t As String = "訓練業別同意協助重新歸類"
    Const cst_TMIDCORRECT_c As String = "貴單位於研提課程時請務必確認所選之訓練業別正確性，如經審查小組審查所選之訓練業別有誤，是否同意協助重新歸類，如不同意，則將依貴單位所選之業別逕行審查。"
    '材料明細表 
    Const cst_printFN1 As String = "SD_14_020"
    '固定費用人時成本上限為140元  'ACTHUMCOST > 140  '充電起飛140 '產投160 by 20210929 AMU
    Const cst_iMAX_ACTHUMCOST_28_old As Integer = 160
    Const cst_iMAX_ACTHUMCOST_28_old2 As Integer = 165
    Const cst_iMAX_ACTHUMCOST_28 As Integer = 170

    '(固定費用人時成本上限) 產投／充電起飛計畫 班級申請 固定費用人時成本上限為140元配合111年計畫修正，修正為160元。
    Const cst_iMAX_ACTHUMCOST_54_old As Integer = 160
    Const cst_iMAX_ACTHUMCOST_54_old2 As Integer = 165
    Const cst_iMAX_ACTHUMCOST_54 As Integer = 170

    '辦理方式: '產投使用／遠距教學 暫不啟用 true:不啟用 false:啟用-遠距教學 
    Dim flag_StopDISTANCE2 As Boolean = False 'TIMS.Utl_GetConfigSet("STOP_DISTANCE")
    '辦理方式: '署的權限可以修改遠距教學  '改為全部都有修改權限 by 20230925  true:可修改遠距教學 false:不可修改
    Dim gflag_DISTANCE_can_updata As Boolean = True 'False 'TIMS.ChkUserLID(sm, 0):true

    Dim s_Msg1 As String = ""
    Dim gflag_can_save As Boolean = True
    Dim i_gSeqno As Integer = 0 '共用序號使用
    Dim sWOScript1 As String = "" '共用JS OPEN語法
    'Const cst_session_saveok As String = "saveok"
    Const cst_ptInsert As String = "Insert"
    Const cst_ptUpdate As String = "Update"
    Const cst_ptView As String = "View"
    Const cst_msg_memo8a As String = "本課程非屬於「職業安全衛生教育訓練規則」所訂定之訓練課程，無法作為時數認列。"
    Const cst_iMaxLen_TeacherDesc As Integer = 500
    'SaveType1
    Const cst_SaveDef As String = "草稿儲存"
    Const cst_SaveBasic As String = "基本儲存"
    Const cst_SaveRcc As String = "正式儲存" '計畫-正式儲存-正式送出
    'Request("PlanID") /TIMS.ClearSQM(Request("ComIDNO") /TIMS.ClearSQM(Request("SeqNO") 有空值或異常:true
    Dim g_flagNG As Boolean = False '有空值或異常
    '2024電話要有規則的輸入
    Dim fg_phone_2024 As Boolean = True
    '2026年啟用 work2026x02 :2026 政府政策性產業 (產投)
    Dim fg_Work2026x02 As Boolean = False
    '2026年啟用 work2026x02 :2026 政府政策性產業 (產投) -舊資料問題調整
    Const cst_VS_trKID25 As String = "VS_trKID25"

    Const cst_NNN As String = "NNN"

    Dim iPlanKind As Integer = 0
    Dim TPlanID As String = ""
    Dim iCostMode As Integer = 0
    Dim PlanID_value As String = ""
    Dim ComIDNO_value As String = ""
    Dim SeqNO_value As String = ""
    Dim dbld2TempTotal As Double = 0
    Dim tmpNoteDt As DataTable '暫存使用(有多個表格連續使用所以使用共用參數)
    Dim tmpPCS As String = "" '有儲存資料過了
    Const cst_flag1 As String = ","
    Const cst_titlemsg1 As String = "早上：7:00-13:00、下午：13:00-18:00、晚上：18:00-22:00"
    Const cst_errmsg1 As String = "程式出現例外狀況，請聯絡 系統駐點人員!"
    Const cst_errmsg2 As String = "產生 課程申請流水號 有誤!!(請確認訓練機構及訓練起日)"
    Const cst_errmsg2b As String = "取得 課程申請流水號 有誤!!"
    Const cst_errmsg3 As String = "傳入參數異常，請重新查詢!!"
    Const cst_errmsg4 As String = "查詢時發生錯誤，請重新輸入查詢值!!"
    Const cst_errmsg5 As String = "查無資料，請重新確認查詢值!!"
    Const cst_errmsg6 As String = "儲存資料有誤!!!"
    Const cst_errmsg7 As String = "課程大綱 儲存資料有誤!!!"
    Const cst_errmsg8 As String = "計畫經費項目檔 儲存資料有誤!!!"
    Const cst_errmsg9 As String = "計畫材料品名項目檔 儲存資料有誤!!!"
    Const cst_errmsg10 As String = "一人份材料明細 儲存資料有誤!!!"
    Const cst_errmsg11 As String = "共同材料明細 儲存資料有誤!!!"
    Const cst_errmsg12 As String = "教材明細 儲存資料有誤!!!"
    Const cst_errmsg13 As String = "其他明細 儲存資料有誤!!!"
    Const cst_errmsg14 As String = "上課時間 儲存資料有誤!!!"
    Const cst_errmsg15 As String = "計畫包班事業單位 儲存資料有誤!!!"
    Const cst_errmsg16 As String = "傳入表格資訊有誤，刪除失敗!!"
    Const cst_errmsg17 As String = "找不到對應的場地代碼"
    Const cst_errmsg18 As String = "請勿嘗試在頁面輸入具有危險性的字元!"
    Const cst_errmsg19 As String = "未建立基本儲存資料，不可按「匯出EXCEL」!!!"

    Const cst_errmsg20 As String = "檢驗有誤!!"
    Const cst_errmsg21 As String = "該功能，不提供該登入計畫使用，若有需要，請先與系統管理者聯繫!!謝謝!!"
    Const cst_errmsg22 As String = "登入者無正確的業務權限，不提供儲存服務!!(請勿在同一瀏覽器開不同視窗，同時登入不同計畫進行資料處理)"
    Const cst_errmsg23 As String = "尚未基本儲存，開班計劃表資料維護無法儲存或輸入!"
    Const cst_errmsg24 As String = "課程大綱，為必填資料"
    Const cst_errmsg25 As String = "課程大綱內容資料，請重新確認!!"

    Const cst_errmsg26 As String = "儲存資料有誤!(請洽系統管理者)!!"
    Const cst_errmsg32 As String = "促進學習機制-是否為iCAP課程 請選擇「是」或「否」"
    Const cst_errmsg33 As String = "促進學習機制-是否為iCAP課程 選擇「是」，請填寫「課程相關說明」"
    Const cst_errmsg33Y As String = "促進學習機制-是否為iCAP課程 選擇「是」，須填寫 班別資料-【iCap標章證號】" '-正式
    Const cst_errmsg33N As String = "促進學習機制-是否為iCAP課程 選擇「否」，不可填寫 班別資料-【iCap標章證號】" '-正式
    Const cst_errmsg34 As String = "儲存資料檢驗有誤，儲存失敗!請再試一次!(若持續發生，請洽系統管理者)"
    Const cst_errmsg35 As String = "班別資料「優先排序」欄位 同一個[申請階段]內，不可重複填寫相同數字!!"
    Const cst_errmsg36 As String = "[開班計劃表資料維護] 專長能力至少須填寫1個,請填 「專長能力標籤」1.名稱!"

    Const cst_errmsg37 As String = "[班別資料]【辦理方式】選擇『混成課程』，選基本儲存、正式儲存時，【遠距課程環境1】必須設定。" '-基本
    Const cst_errmsg38 As String = "[班別資料]「遠距教學總時數不得超過本班總訓練時數2/3」。" '-正式
    Const cst_errmsg39 As String = "[訓練費用] 填寫 材料費用總額 大於0，須填寫「一人份材料明細」或「共同材料明細」資料。" '-正式

    '二、於「基本儲存」及「正式儲存」時增加卡控，未符合以下條件不能儲存，並跳出提示訊息：
    '1、 若是【辦理方式】選擇混成課程，點選基本儲存、正式儲存時，【遠距課程環境】欄位必須填寫。
    '2、 遠距教學總時數不得超過該班訓練時數之1/3。例如總訓練時數為20小時，則遠距教學時數至多僅能6小時， 跳出提示訊息：「遠距教學總時數不得超過本班訓練時數1/3」。
    '3、 倘單位有填寫材料費總額，須填寫「一人材料明細」或「共同材料明細」。
    '4、 專長能力至少須填寫1個。
    '5、 倘單位於班別資料頁籤有填寫iCAP標章編碼，於開班計劃表資料維護頁籤之「是否為iCAP課程」須選「是」，並填寫課程相關說明。

    Const Cst_msgother1 As String = "( 學員資格* 請到 開班計畫表資料維護作業 )"
    'Const Cst_msgother3 As String = "※請先確認有【一人份材料明細】或【共同材料明細】資料後，先按「基本儲存」，再按「匯出EXCEL」!! <br>　更新資料訓練費用編列說明，請按「匯出EXCEL」!!"
    Const Cst_msgother3 As String = "※請先確認有【教材明細】資料後，先按「基本儲存」，再按「匯出EXCEL」!! <br>　更新資料訓練費用編列說明，請按「匯出EXCEL」!!"
    Const Cst_msgother3b As String = "※更新資料訓練費用編列說明，請按其他說明「修改」!!"
    'Const Cst_msgother7 As String = "※請先確認前面7個頁籤資料的資料的填寫動作，若已完成請按下「基本儲存」，再按「正式儲存」!! <br>更新後，請按「正式儲存」!!"
    Const Cst_msgother7 As String = "※步驟說明：<br />前面5個頁籤基本資料填寫完成，請先按下「2.基本儲存」。<br />待【開班計劃表資料維護】頁籤填寫完成或有資料更新，請按「3.正式儲存」。<br />欲送審，請至【班級查詢】清單頁點「送出」。"
    Const Cst_msgother8 As String = "※如要修改「訓練人數」或「訓練時數」，請先至訓練費用頁籤將「固定費用總額」刪除即可修改。"
    'LabMsg8.Text = Cst_msgother8

    Const cst_msg_save1 As String = "基本儲存成功!!"
    Const cst_msg_save2 As String = "正式儲存成功!!"
    Const cst_msg_save3 As String = "草稿儲存成功!!"
    Const cst_msg_save4 As String = "同一位師資授課時數已超過54小時， 請確認是否為特殊情況。"

    'SELECT * FROM PLAN_PLANINFO WHERE RID ='B5703' AND SEQNO=17 
    'PLAN_ONCLASS/PLAN_BUSPACKAGE/PLAN_VERREPORT/PLAN_TEACHER
    'SELECT * FROM TEACH_TEACHERINFO WHERE RID ='B5703'
    'Const cst_TrainDescTable As String = "TrainDescTable" '產學訓課程大綱 PLAN_TRAINDESC
    'Const cst_PLAN_ONCLASS As String = "PLAN_ONCLASS"
    'Const cst_PLAN_BUSPACKAGE As String = "PLAN_BUSPACKAGE"
    'Const Cst_PersonCostTable As String = "PersonCostTable" 'PersonCost: PLAN_PERSONCOST–一人份材料明細
    Const Cst_PersonCostpkName As String = "ppcID"
    'Const Cst_CommonCostTable As String = "CommonCostTable" 'CommonCost: PLAN_COMMONCOST–共同材料明細
    Const Cst_CommonCostpkName As String = "pcmID"
    'CommonCost
    'Const Cst_SheetCostTable As String = "SheetCostTable" 'SHEETCOST: PLAN_SHEETCOST–教材明細 (產學訓)
    Const Cst_SheetCostpkName As String = "pshID"
    'OtherCost
    'Const Cst_OtherCostTable As String = "OtherCostTable" 'OTHERCOST: PLAN_OTHERCOST–其他明細 (產學訓)
    Const Cst_OtherCostpkName As String = "potID"

    Const cst_學分班 As String = "Y"
    Const cst_非學分班 As String = "N"

    '不管什麼都是「年滿15歲以上」。
    Const cst_AgeOtherDef As Integer = 16 'other Years Start
    '目前為複製模式
    Dim gflag_ccopy As Boolean = False 'Request(cst_ccopy) true:copy /false: not copy 'gflag_ccopy
    'COPYSUB 1:課程表 /2:材料明細
    Dim gflag_can_copy1 As Boolean = False
    Dim gflag_can_copy2 As Boolean = False
    '產投 - 班級複製【訓練業別同意協助重新歸類】改為不複製
    Dim gflag_can_copy3 As Boolean = False

    '目前為複製模式-但是有編輯模式啟動
    Dim gflag_TrainDesc_edit1 As Boolean = False
    Const cst_ccopy As String = "ccopy" 'Request(cst_ccopy)'gflag_ccopy

    Const cst_Years_2007 As String = "2007"
    Const cst_Years_2019 As String = "2019"
    Dim strYears As String = "" '2014 / 2015 /2018'(經費分類代碼。)
    Const cst_strYears_2014 As String = "2014"
    Const cst_strYears_2015 As String = "2015"
    Const cst_strYears_2018 As String = "2018"

    'Dim flag_TIMS_Test_1 As Boolean = TIMS.sUtl_ChkTest()
    'Dim flag_TIMS_Test_1 As Boolean = False
    Dim flag_IsSuperUser_1 As Boolean = False

    '政策性產業課程可辦理班數-PLAN_PRECLASS
    'Dim flag_SHOW_2019_3 As Boolean = False

    'addkey/"OJT22071401"/value="Y"/ 'OJT-22071401系統-產投-班級申請：新增「訓練業別同意重新歸類」選項 +「與政策性產業課程之關聯性概述」欄位 
    'Dim flag_OJT22071401 As Boolean = False
    'work2013x01/材料明細
    Dim flag_work2013x01 As Boolean = False
    '產業別(管考) true:使用/false:不可使用
    Dim fg_USE_CBLKID60_TP28 As Boolean = False
    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '初始化，欄位長度設定
        Call SUtl_PageInit1()
        '每次執行頁面
        Call CCreate1_Every()

        '放置畫面上的Dropdonwlist
        If Not IsPostBack Then
            '產生新的GUID 避免記憶體相同 而異常
            Call CREATE_NEW_GUID21()

            'Call GetNewtable()  '建立空白表格
            ViewState("dtTaddress") = GetNewtable()  '建立空白表格

            Session(hid_TrainDescTable_guid1.Value) = Nothing 'PLAN_TRAINDESC
            Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
            Session(hid_PersonCostTable_guid1.Value) = Nothing
            Session(hid_CommonCostTable_guid1.Value) = Nothing
            Session(hid_SheetCostTable_guid1.Value) = Nothing
            Session(hid_OtherCostTable_guid1.Value) = Nothing
            Session(hid_planONCLASS_guid1.Value) = Nothing
            If Session("search") IsNot Nothing Then ViewState("search") = Session("search")

            '建立物件 顯示預設值
            Call CCreateItem()

            ViewState("GUID1") = TIMS.GetGUID() : Session("GUID1") = ViewState("GUID1")
            Page.RegisterStartupScript("window_onload", "<script language=""javascript"">Layer_change(1);</script>")
        End If

        If Not SUtl_PageLoad2() Then Exit Sub

        Dim v_State As String = TIMS.ClearSQM(Request("State"))
        If TIMS.ClearSQM(Request("State")) <> "" Then
            Button8.Enabled = False '草稿儲存
            TIMS.Tooltip(Button8, "狀態不可儲存。")
            btnAdd.Enabled = False '基本儲存
            TIMS.Tooltip(btnAdd, "狀態不可儲存。")
            BtnSAVE2.Enabled = False '正式儲存
            TIMS.Tooltip(BtnSAVE2, "狀態不可儲存。")
            gflag_can_save = False '不可儲存
            Button1.Enabled = False
            'TIMS.Tooltip(Button1, v_State)
            Button29.Enabled = False
            'TIMS.Tooltip(Button29, v_State)

            btnAddCost6.Enabled = False
            btnAddCost7.Enabled = False
            btnAddCost8.Enabled = False
            btnAddCost9.Enabled = False
            btnUptNote2.Enabled = False
            BtnImport1.Enabled = False
            BtnImport2.Enabled = False
            BtnImport8.Enabled = False
            BtnImport9.Enabled = False

            btu_sel.Disabled = True
            TIMS.Tooltip(btu_sel, v_State)
            btu_sel2.Disabled = True
            TIMS.Tooltip(btu_sel2, v_State)
        End If
        '檢查功能權限 End

        'Request("PlanID") /TIMS.ClearSQM(Request("ComIDNO") /TIMS.ClearSQM(Request("SeqNO") 有空值或異常:true
        g_flagNG = Get_GflagNG1()

        If Not IsPostBack Then
            Label3.Text = sm.UserInfo.Years '登入年度顯示

            'tr_PolicyPreVal.Style.Item("display") = "none"
            If TIMS.ClearSQM(Request("PlanID")) = "" Then
                '如果是自辦計劃，或者是委外並且是委訓登入，則帶入預設值
                If iPlanKind = 1 OrElse sm.UserInfo.LID = 2 Then
                    Dim dr As DataRow = GET_RID_ORGPLANINFO()
                    If dr IsNot Nothing Then
                        RIDValue.Value = sm.UserInfo.RID
                        ComidValue.Value = $"{dr("ComIDNO")}"
                        center.Text = $"{dr("orgname")}"
                        EMail.Text = TIMS.ChangeEmail($"{dr("ContactEmail")}")
                        EnterSupplyStyle.Enabled = False
                        Common.SetListItem(EnterSupplyStyle, "1")
                        '1.報名時應先繳全額訓練費用，待結訓審核通過後核撥補助款
                        '2.報名時應先繳50%訓練費用，待結訓審核通過後核撥補助款
                        Select Case $"{dr("ORGKIND2")}"
                            Case "G" '非勞工團體
                            Case "W" '勞工團體
                                EnterSupplyStyle.Enabled = True
                                Common.SetListItem(EnterSupplyStyle, "2")
                        End Select
                    End If
                End If

                Org.Disabled = False '訓練機構(可選)
                Button24.Visible = False '回上一頁
                '新增狀態、帶入預設值
                'Call SHOW_DG33_PRECLASSCNT(0)
            Else
                Org.Disabled = True '訓練機構(不可選)
                Button24.Visible = True '回上一頁

                Call SHOW_PLANPLANINFO() '顯示該計畫資料，應該是修改或檢視
                Call SHOW_PLAN_VERREPORT()
                Call SHOW_PLAN_TEACHER12()
                Call SHOW_PLAN_TRAINPLACE()
                Call SHOW_PLAN_ABILITYS() '專長能力標籤-ABILITY
                '儲存-政策性產業課程可辦理班數-PLAN_PRECLASS
                'If flag_SHOW_2019_3 Then Call SHOW_DG33_PRECLASSCNT(1)
                Call CreateClassTime()
                Call CreateTrainDesc() 'PLAN_TRAINDESC
                Call CreateBusPackage()
                Call CreatePersonCost()
                Call CreateCommonCost()
                Call CreateSheetCost()
                Call CreateOtherCost()
            End If
        End If

        '2004/12/7 前端增加javascript屬性 Start
        Hid_D25_8_CapAll_MSG.Value = cst_D25_8_CapAll_MSG
        If Not IsPostBack Then
            SciPlaceID.Attributes("onchange") = "javascript:{doGetPTID('SciPlaceID','hid_AddressSciPTID','SciPlaceID2');}"
            SciPlaceID2.Attributes("onchange") = "javascript:{doGetPTID('SciPlaceID2','hid_AddressSciPTID2','SciPlaceID');}"
            TechPlaceID.Attributes("onchange") = "javascript:{doGetPTID('TechPlaceID','hid_AddressTechPTID','TechPlaceID2');}"
            TechPlaceID2.Attributes("onchange") = "javascript:{doGetPTID('TechPlaceID2','hid_AddressTechPTID2','TechPlaceID');}"

            FIXSUMCOST.Attributes("onclick") = "chg_FIXSUMCOST();"
            FIXSUMCOST.Attributes("onblur") = "chg_FIXSUMCOST();"
            METSUMCOST.Attributes("onclick") = "chg_FIXSUMCOST();"
            METSUMCOST.Attributes("onblur") = "chg_FIXSUMCOST();"
            CBLKID25_8.Attributes("onclick") = "chg_CBLKID25_8();" 'CapAll,Hid_D25_8_CapAll_MSG,Hid_CapAll

            'Hid_PDF20171226.Value = TIMS.pdf_g20171226
            Hid_PDF20171226.Value = TIMS.Get_PDF_GovTraining(sm)
            btnPDF20171226.Attributes("onclick") = "return openPDF20171226();"

            date1.Attributes("onclick") = "javascript:show_calendar('STDate','','','CY/MM/DD');"
            date2.Attributes("onclick") = "javascript:show_calendar('FDDate','','','CY/MM/DD');"
            date3.Attributes("onclick") = "return chkTrainDate('STrainDate');"
        End If

        'Dim s_bicapoC2 As String = ""
        's_bicapoC2 = String.Format("{0}?CIGD={1}", "TC_03_ICAP.aspx", TIMS.GetGUID())
        BtnICAPonlineC2.Attributes("onclick") = String.Concat("open_ICAP1('", TIMS.xBlockName(), "');")

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?btnName=Button28');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        '增加快速點選機構清單
        If Not Org.Disabled Then
            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "Button28")
            If HistoryRID.Rows.Count <> 0 Then
                center.Attributes("onclick") = "showObj('HistoryList2');"
                center.Style("CURSOR") = "hand"
            End If
        End If

        '自辦申請計畫
        If Not IsPostBack Then
            Button8.Attributes("onclick") = "return Check_Temp();" '草稿儲存
            Button1.Attributes("onclick") = "return CheckTrainDescTable();" '課程大綱檢查
            Button29.Attributes("onclick") = "return CheckAddTime();" ''上課時間檢查
            btnAddBusPackage.Attributes("onclick") = "return CheckAddBusPackage();" '包班事業單位資料檢查
        End If

        '計算經費來源的加總
        If Not IsPostBack Then
            TNum.Attributes("onblur") = "CountCostSource();"
            box6.Attributes("onclick") += "javascript:CountCostSource();"
            box7.Attributes("onclick") += "javascript:CountCostSource();"
            DefGovCost.Attributes("onblur") = "CountCostSource();"
            DefStdCost.Attributes("onblur") = "CountCostSource();"
            If gflag_can_save Then
                GCIDName.Attributes.Add("onDblClick", "javascript:Get_GovClass('GCIDName');")
                GCIDName.Style("CURSOR") = "hand"
                'btn_GCID.Attributes.Add("onclick", "javascript:Get_GovClass('GCIDName');")
            End If
            If Not gflag_can_save Then
                btn_GCID.Disabled = True
                btn_GCID.Style.Item("display") = "none"
            End If
        End If

        IsBusiness.Enabled = False
        'EnterpriseName.Enabled = False
        IsBusiness.ToolTip = "暫不開放此功能" '本年度暫不開放此功能"
        'EnterpriseName.ToolTip = "暫不開放此功能" '"本年度暫不開放此功能"

        '112下半年產投方案預計於3/24~4/26受理提案，「申請階段」固定預設為「下半年」，減少訓練單位誤選其他階段
        'Call Utl_ojt23021601()
        Dim str_CCS As String = "CountCostSource();"
        Dim str_SPT As String = "showPTID('Classification1','PTID1','PTID2');"
        Dim str_SCT As String = $"showCostType('{Radiobuttonlist1.ClientID}');"
        Dim str_CEZ As String = $"CHK_EnvZeroTrain();"
        Dim str_script As String = $"<script>{str_CCS}{str_SPT}{str_SCT}{str_CEZ}</script>"
        Page.RegisterStartupScript(TIMS.xBlockName(), str_script)

        '確認機構是否為黑名單 (處份/處分)
        Dim vsMsg2 As String = "" ' vsMsg2 = ""
        If Chk_OrgBlackList(vsMsg2) Then
            Button8.Visible = False
            TIMS.Tooltip(Button8, vsMsg2)
            btnAdd.Visible = False
            TIMS.Tooltip(btnAdd, vsMsg2)
            BtnSAVE2.Visible = False
            TIMS.Tooltip(BtnSAVE2, vsMsg2)
            gflag_can_save = False '不可儲存

            Dim vsStrScript As String = $"<script>alert('{vsMsg2}');</script>"
            Page.RegisterStartupScript("", vsStrScript)
        End If
        PointType.Attributes("onclick") = "return GetPointName();"
        rbl_AppStage.Attributes("onclick") = "return GetAppStageMSG1();"

        '包班種類 '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            TIMS.Tooltip(PackageType, "充電起飛計畫不可選擇非包班!!")
            PackageType.Attributes("onclick") = "GetPackageName54();"
        Else
            PackageType.Attributes("onclick") = "GetPackageName();"
        End If

        '2011 功能按鈕權限控管--Start
        Dim strSechObjID As String = "" '查詢按鈕物件ID
        Dim strAddsObjID As String = "" '維護按鈕物件ID
        Dim strPrntObjID As String = "" '列印按鈕物件ID
        strAddsObjID = Button1.ClientID & "," & Button29.ClientID & "," & btnAddBusPackage.ClientID & "," & Button8.ClientID & "," & btnAdd.ClientID
        'Call TIMS.CheckBtnAuth(Me, strSechObjID, strAddsObjID, strPrntObjID)
        '2011 功能按鈕權限控管--End
    End Sub
    ''' <summary>'依RID 取得機構部分資訊</summary>
    ''' <returns></returns>
    Function GET_RID_ORGPLANINFO() As DataRow
        Dim sParms As New Hashtable From {{"RID", sm.UserInfo.RID}}
        Dim sql As String = ""
        sql &= " SELECT b.ORGNAME, b.ComIDNO, c.ContactEmail, c.ZipCode, c.Address, b.OrgKind2"
        sql &= " FROM AUTH_RELSHIP a"
        sql &= " JOIN ORG_ORGINFO b ON a.ORGID=b.ORGID"
        sql &= " JOIN ORG_ORGPLANINFO c ON a.RSID=c.RSID"
        sql &= " WHERE a.RID=@RID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, sParms)
        Return dr
    End Function

    'rbl_AppStage : 112下半年產投方案預計於3/24~4/26受理提案，「申請階段」固定預設為「下半年」，減少訓練單位誤選其他階段
    'Sub Utl_ojt23021601()
    '    '<add key = "ojt23021601_SDATE" value="2023/03/24"  />
    '    '<add key = "ojt23021601_EDATE" value="2023/04/26"  />
    '    '<add key = "ojt23021601_rblAppStage" value="2"  />
    '    Dim ojt23021601_SDATE As String = TIMS.Utl_GetConfigSet("ojt23021601_SDATE")
    '    Dim ojt23021601_EDATE As String = TIMS.Utl_GetConfigSet("ojt23021601_EDATE")
    '    Dim oRblAppStageVal As String = TIMS.Utl_GetConfigSet("ojt23021601_rblAppStage")
    '    If ojt23021601_SDATE = "" OrElse ojt23021601_EDATE = "" OrElse oRblAppStageVal = "" Then Return

    '    ojt23021601_SDATE = TIMS.cdate3(ojt23021601_SDATE)
    '    ojt23021601_EDATE = TIMS.cdate3(ojt23021601_EDATE)
    '    If ojt23021601_SDATE = "" OrElse ojt23021601_EDATE = "" Then Return

    '    Dim iS1 As Long = DateDiff(DateInterval.Day, CDate(ojt23021601_SDATE), Date.Now)
    '    Dim iS2 As Long = DateDiff(DateInterval.Day, Date.Now, CDate(ojt23021601_EDATE))
    '    Dim flagCanUse As Boolean = (iS1 >= 0 AndAlso iS2 >= 0)
    '    If Not flagCanUse Then Return
    '    '2023年後
    '    If sm.UserInfo.Years < 2023 Then Return
    '    If TIMS.ClearSQM(Request("PlanID")) = "" Then
    '        Common.SetListItem(rbl_AppStage, oRblAppStageVal)
    '        rbl_AppStage.Enabled = False
    '        TIMS.Tooltip(rbl_AppStage, "申請階段欄位鎖定中")
    '    ElseIf gflag_ccopy Then
    '        Common.SetListItem(rbl_AppStage, oRblAppStageVal)
    '        rbl_AppStage.Enabled = False
    '        TIMS.Tooltip(rbl_AppStage, "申請階段欄位鎖定中")
    '    End If
    '    '限定委訓單位／分署
    '    'If sm.UserInfo.LID = 2 OrElse sm.UserInfo.LID = 1 Then
    '    'End If
    'End Sub

    ''' <summary>每次執行頁面</summary>
    Private Sub CCreate1_Every()
        hfScrollToAnchor.Value = ""

        '產業別(管考) true:使用/false:不可使用
        If Hid_USE_CBLKID60_TP28.Value = "" Then Hid_USE_CBLKID60_TP28.Value = TIMS.Utl_GetConfigVAL(objconn, "USE_CBLKID60_TP28")
        fg_USE_CBLKID60_TP28 = (Hid_USE_CBLKID60_TP28.Value = "Y")
        trCBLKID60.Visible = fg_USE_CBLKID60_TP28

        'OJT-22071401 系統-產投-班級申請：新增「訓練業別同意重新歸類」選項 +「與政策性產業課程之關聯性概述」欄位 
        'If Hid_OJT22071401.Value = "" Then Hid_OJT22071401.Value = TIMS.Utl_GetConfigVAL(objconn, "OJT22071401")
        'flag_OJT22071401 = (Hid_OJT22071401.Value = "Y")
        trTMIDCORRECT.Visible = (TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1)
        trPOLICYREL_t.Visible = True 'flag_OJT22071401
        trPOLICYREL_c.Visible = True 'flag_OJT22071401
        '若訓練業別有誤是否同意協助重新歸類
        'If flag_OJT22071401 Then labtTMIDCORRECT.Text = str_TMIDCORRECT_t 'If flag_OJT22071401 Then LabmsgTMIDCORRECT.Text = cst_TMIDCORRECT_c
        labtTMIDCORRECT.Text = str_TMIDCORRECT_t
        LabmsgTMIDCORRECT.Text = cst_TMIDCORRECT_c

        'Dim flag_TIMS_Test_1 As Boolean = TIMS.sUtl_ChkTest()
        'flag_TIMS_Test_1 = TIMS.sUtl_ChkTest()
        flag_IsSuperUser_1 = TIMS.IsSuperUser(Me, 1)
        '政策性產業課程可辦理班數-PLAN_PRECLASS
        'Dim flag_SHOW_2019_3 As Boolean = False
        'flag_SHOW_2019_3 = TIMS.SHOW_2019_3() 'work2019x03

        iPYNum = TIMS.sUtl_GetPYNum(Me)
        '(經費分類代碼。)
        strYears = cst_strYears_2014 '2014年  顯示層級。
        If sm.UserInfo.Years >= TIMS.CINT1(cst_strYears_2015) Then strYears = cst_strYears_2015 '2015年 不顯示層級。
        If iPYNum >= 3 Then strYears = cst_strYears_2018 '2018年  

        'gflag_ccopy
        gflag_ccopy = If(Convert.ToString(Request(cst_ccopy)) = "1", True, False)
        'COPYSUB 1:課程表 /2:材料明細
        gflag_can_copy1 = If(gflag_ccopy AndAlso Convert.ToString(Request("COPYSUB1")) = "Y", True, False) 'gflag_ccopy
        gflag_can_copy2 = If(gflag_ccopy AndAlso Convert.ToString(Request("COPYSUB2")) = "Y", True, False) 'gflag_ccopy

        TableCost6.Visible = False '一人份材料明細
        TableCost7.Visible = False '共同材料明細
        '材料明細
        flag_work2013x01 = If(TIMS.Utl_GetConfigSet("work2013x01") = "Y", True, False)
        If flag_work2013x01 Then
            Labmsg3.Text = Cst_msgother3
            Note.ReadOnly = True
            Note.Style.Item("background-color") = "#BDBDBD"
            TableCost6.Visible = True
            TableCost7.Visible = True
        End If

        LabMsg8.Text = Cst_msgother8
        LabMsg7.Text = Cst_msgother7
        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            hTPlanID54.Value = cst_errmsg21
            Dim sScript1 As String = ""
            sScript1 &= " <script>alert('" & hTPlanID54.Value & "');</script> "
            sScript1 &= " <script>location.href='../../main2.aspx';</script> "
            Call TIMS.Utl_RespWriteEnd(Me, objconn, "")
            Exit Sub
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then LabTMID.Text = "訓練業別"

        '(限定產投) OJT-21012202 2021/02/24 
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            IsBusiness.Visible = False '(企業包班)
            IsBusiness_chk1.Visible = False '(企業包班)
            'labEnterpriseName.Visible = False '(企業包班名稱)
            'EnterpriseName.Visible = False '(企業包班名稱)
            trPackageType.Visible = False '包班種類 '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
            PackageName.Visible = False '班別名稱'包班種類
        End If

        hTPlanID54.Value = If(TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, "1", "")  '充電起飛計畫'js client使用
        'tr_AppStage.Visible = If(TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, False, True)  '充電起飛計畫 '申請階段檢核 / '不檢核[申請階段]
        Datagrid4headTable.Visible = If(TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, True, False) '充電起飛計畫
        Datagrid4Table.Visible = If(TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, True, False) '充電起飛計畫

        '技檢訓練時數 產業人才投資方案(充飛不用)
        '1.【技檢訓練時數】需<=該堂課【時數】，可允許小數點一位
        '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-1才存欄位，否清空。
        '3.權限：跟其他欄位一樣。'訓練單位可填寫，但送審後鎖住不可修改。送審後分署可修改。'3.訓練班節計畫表加上【技檢時數】欄位顯示。
        '4.註記：'調整班級變更申請、班級變更審核、結訓證書等功能 '結訓證書上顯示： 符合申請技檢訓練時數
        td_EHour_h1.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        td_EHour_h2.Visible = td_EHour_h1.Visible
        TIMS.Tooltip(th_EHOUR, cst_EHour_t1, True)
        TIMS.Tooltip(EHOUR, cst_EHour_t1, True)

        trKID19.Visible = False 'If(fg_SHOW_2019_1, False, True)
        trKID18.Visible = False 'If(fg_SHOW_2019_1, False, True)
        '2019年啟用 work2019x01:2019 政府政策性產業 'trKID20.Visible = False
        'Dim fg_SHOW_2019_1 As Boolean = TIMS.SHOW_2019_1(sm)
        Dim fg_SHOW_2025_1 As Boolean = TIMS.SHOW_2025_1(sm)
        trKID20.Visible = If(fg_SHOW_2025_1, False, True)
        trKID25.Visible = If(fg_SHOW_2025_1, True, False)

        '2026年啟用 work2026x02 :2026 政府政策性產業 (產投),'Dim fg_Work2026x02 As Boolean = TIMS.SHOW_Work2026x02(sm)
        fg_Work2026x02 = TIMS.SHOW_W2026x02(sm)
        trKID26.Visible = fg_Work2026x02 '(未啟動暫不顯示)
        If fg_Work2026x02 Then
            If $"{ViewState(cst_VS_trKID25)}" = "Y" Then
                TIMS.Display_Inline(trKID25) '(顯示)
                TIMS.Display_None(trKID26) '(不顯示)
            Else
                TIMS.Display_None(trKID25) '(不顯示)
                TIMS.Display_Inline(trKID26) '(顯示)
            End If
        End If

        '2、於政府政策性產業增加「AI加值應用」、「職場續航」之勾選欄位 (下圖)。充電起飛計畫不用新增
        'AI加值應用
        trCBLKID25_7.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        '職場續航
        trCBLKID25_8.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        'AI應用時數
        th_AIAHOUR.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        '職場續航時數
        th_WNLHOUR.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        Sp_th_AIAHOUR_WNLHOUR.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        Sp_AIAHOUR_WNLHOUR.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)

        'If tr_AppStage_TP28.Visible Then
        '    增加 申請階段 說明
        '    Dim str_labAppStageMsg As String = ""
        '    Select Case sm.UserInfo.TPlanID
        '        Case "28"
        '            str_labAppStageMsg = ""
        '            str_labAppStageMsg += "申請上半年課程：1<br>" & vbCrLf
        '            str_labAppStageMsg += "申請下半年課程：2<br>" & vbCrLf
        '        Case "54"
        '            str_labAppStageMsg = ""
        '            str_labAppStageMsg += " 1月1~15日：1、 1月16~31日：2  <br>" & vbCrLf
        '            str_labAppStageMsg += " 2月1~15日：3、 2月16~29日：4  <br>" & vbCrLf
        '            str_labAppStageMsg += " 3月1~15日：5、 3月16~31日：6  <br>" & vbCrLf
        '            str_labAppStageMsg += " 4月1~15日：7、 4月16~30日：8  <br>" & vbCrLf
        '            str_labAppStageMsg += " 5月1~15日：9、 5月16~31日：10 <br>" & vbCrLf
        '            str_labAppStageMsg += " 6月1~15日：11、6月16~30日：12 <br>" & vbCrLf
        '            str_labAppStageMsg += " 7月1~15日：13、7月16~31日：14 <br>" & vbCrLf
        '            str_labAppStageMsg += " 8月1~15日：15、8月16~31日：16 <br>" & vbCrLf
        '            str_labAppStageMsg += " 9月1~15日：17、9月16~30日：18 <br>" & vbCrLf
        '            str_labAppStageMsg += "10月1~15日：19、10月16~31日：20<br>" & vbCrLf
        '            str_labAppStageMsg += "11月1~15日：21、11月16~30日：22<br>" & vbCrLf
        '            str_labAppStageMsg += "12月1~15日：23、12月16~31日：24<br>" & vbCrLf
        '    End Select
        '    labAppStageMsg.Text = str_labAppStageMsg
        'End If

        FactModeOther.ReadOnly = True
        FactMode.Enabled = False
        FactMode.Visible = False
        FactModeOther.Enabled = False
        FactModeOther.Visible = False
        FactModeTR.Visible = False
        trRoomName.Visible = False
        RoomName.ReadOnly = True
        RoomName.Enabled = False
        RoomName.Visible = False

        'TIMS.Display_None(tr_Taddress2)
        'TIMS.Display_None(trOtherDesc23)

        ContentTR.Visible = False
        'PHour.Attributes.Add("onBlur", "if(this.value !='') {var msg=''; if(!isInt(this.value)){msg+='時數只能輸入整數。\n';} if(this.value <= 0){msg+='時數必須大於0\n';} if(msg !=''){alert(msg);this.focus();}}")
        Page.RegisterStartupScript("ZipCcript1", TIMS.Get_ZipNameJScript(objconn))
        TMScience.Enabled = False
        TMScience.Text = "(不提供輸入)"

        '--是否要載入系統自動暫存功能
        '必免重複執行
        If Not gflag_ccopy AndAlso IsPostBack Then Hid_sisyphus.Value = "N"
        '--是否要載入系統自動暫存功能
        'COPY 功能不啟用
        If gflag_ccopy Then Hid_sisyphus.Value = "N"
        '辦理方式: '產投使用／遠距教學 暫不啟用
        'BY 2023/9/23 啟用遠距教學
        flag_StopDISTANCE2 = False 'If(TIMS.Utl_GetConfigSet("STOP_DISTANCE_2").Equals("Y"), True, False)

        '辦理方式: '署的權限可以修改遠距教學 '(非署單位完全不可修改)
        gflag_DISTANCE_can_updata = True 'TIMS.ChkUserLID(sm, 0)

        If gflag_DISTANCE_can_updata Then flag_StopDISTANCE2 = False '(非署單位完全不可修改)
    End Sub

    ''' <summary>判斷計畫種類，選擇要顯示的經費項目</summary>
    ''' <returns></returns>
    Function SUtl_PageLoad2() As Boolean
        Dim rst As Boolean = True
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID")) '外部傳入 copy可能
        If rqPlanID = "" Then rqPlanID = sm.UserInfo.PlanID

        '判斷計畫種類，選擇要顯示的經費項目
        Dim sql As String = " SELECT TPLANID, PLANKIND, YEARS FROM ID_PLAN WHERE PlanID=@PlanID "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.Int).Value = TIMS.CINT1(rqPlanID)
            dt.Load(.ExecuteReader())
        End With

        If TIMS.dtHaveDATA(dt) Then
            Dim dr As DataRow = dt.Rows(0)
            iPlanKind = TIMS.CINT1(dr("PlanKind"))
            TPlanID = $"{dr("TPlanID")}"
        Else
            sm.LastErrorMessage = cst_errmsg1
            rst = False 'Exit Sub
        End If

        '顯示E-Mail欄位給予填寫（自辦）
        Table1_Email.Visible = If(iPlanKind = 1, False, True) '自辦
        Return rst
    End Function

    '如果是複製狀態, 則RID還是為原登入計畫之RID by nick
    ''' <summary>儲存按鈕啟動關鍵</summary>
    ''' <returns></returns>
    Function SUtl_GetRIDn() As String
        Dim rst As String = ""
        Dim sql As String = $"
SELECT RID AS RIDN FROM AUTH_RELSHIP WHERE PlanID='{sm.UserInfo.PlanID}' AND DistID='{sm.UserInfo.DistID}'
AND ORGID IN (SELECT ORGID FROM ORG_ORGINFO WHERE COMIDNO='{TIMS.ClearSQM(Request("ComIDNO"))}')
"
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If sm.UserInfo.LID = 0 Then
                Button8.Visible = False
                btnAdd.Visible = True
                BtnSAVE2.Visible = True '顯示儲存鈕
                Button24.ToolTip = "審核通過 or 審核後修正者,檢視班級"
                sql = $" SELECT RID AS RIDN FROM AUTH_RELSHIP WHERE RID='{sm.UserInfo.RID}' AND DistID='{sm.UserInfo.DistID}' "
            End If
        End If
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        If dr IsNot Nothing Then rst = $"{dr("RIDN")}"
        Return rst
    End Function

    'Request("PlanID") /TIMS.ClearSQM(Request("ComIDNO") /TIMS.ClearSQM(Request("SeqNO") 有空值或異常:true
    Function Get_GflagNG1() As Boolean
        Dim rst As Boolean = False
        If TIMS.ClearSQM(Request("PlanID")) = "" Then rst = True
        If TIMS.ClearSQM(Request("ComIDNO")) = "" Then rst = True 'Exit Sub
        If TIMS.ClearSQM(Request("SeqNO")) = "" Then rst = True 'Exit Sub
        If TIMS.CINT1(Request("PlanID")) = 0 Then rst = True
        If TIMS.CINT1(TIMS.Get_Substr1($"{Request("ComIDNO")}", 8)) = 0 Then rst = True 'Exit Sub
        If TIMS.CINT1(Request("SeqNO")) = 0 Then rst = True 'Exit Sub
        Return rst
    End Function

    ''' <summary>SHOW, PLAN_DEPOT</summary>
    ''' <param name="PMS_PD1"></param>
    Sub SHOW_PLAN_DEPOT(PMS_PD1 As Hashtable)
        Dim hrPLANID As String = TIMS.GetMyValue2(PMS_PD1, "PLANID")
        Dim hrCOMIDNO As String = TIMS.GetMyValue2(PMS_PD1, "COMIDNO")
        Dim hrSEQNO As String = TIMS.GetMyValue2(PMS_PD1, "SEQNO")

        Dim drPLAN As DataRow = TIMS.GetPlanID1(hrPLANID, objconn)
        Dim s_DISTID As String = If(drPLAN IsNot Nothing, $"{drPLAN("DISTID")}", "")
        If sm.UserInfo.LID = 0 AndAlso s_DISTID <> "" Then '(可重新調整選項)
            Call TIMS.CHG_ddlDEPOT15(ddlDEPOT15, s_DISTID, sm.UserInfo.TPlanID, objconn)
        End If

        'Dim fg_Work2026x02 As Boolean = TIMS.SHOW_Work2026x02(sm)
        Dim dr2 As DataRow = TIMS.GET_PLANDEPOT(hrPLANID, hrCOMIDNO, hrSEQNO, objconn)
        If dr2 IsNot Nothing Then
            Dim s_SEQNOD15 As String = $"{dr2("SEQNOD15")}"
            If s_SEQNOD15 <> "" AndAlso Not gflag_ccopy Then
                '轄區重點產業(排除停用)(轄區) 15/21 搜尋舊值顯示
                'OJT-21082501、OJT-2108250２：<系統> 產投 - 班級申請-班別資料、重點產業審核確認：轄區重點產業項目調整（南分署）
                TIMS.CHG_ddlDEPOT15_21(ddlDEPOT15, sm.UserInfo.DistID, objconn, s_SEQNOD15)
            End If
            If s_SEQNOD15 <> "" Then Common.SetListItem(ddlDEPOT15, s_SEQNOD15)

            Dim KID06 As String = $"{dr2("KID06")}"
            Dim KID10 As String = $"{dr2("KID10")}"
            Dim KID19 As String = $"{dr2("KID19")}"
            Dim KID18 As String = $"{dr2("KID18")}"
            Dim cvKID60 As String = $"{dr2("KID60")}"
            '2019年啟用 work2019x01:2019 政府政策性產業
            Dim cvKID20 As String = $"{dr2("KID20")}"
            Dim cvKID25 As String = $"{dr2("KID25")}"
            Dim cvKID26 As String = $"{dr2("KID26")}"
            If gflag_ccopy Then cvKID20 = "" '(複制狀態清空)
            If gflag_ccopy Then cvKID25 = "" '(複制狀態清空)
            If gflag_ccopy Then cvKID26 = "" '(複制狀態清空)

            '2026年啟用 work2026x02 :2026 政府政策性產業 (產投) 'trKID26.Visible = fg_Work2026x02
            If fg_Work2026x02 Then
                Dim v_IR_AppStage As Integer = TIMS.CINT1(TIMS.GetListValue(rbl_AppStage))
                If (v_IR_AppStage = 0) Then v_IR_AppStage = 3 'NULL(強制轉為政策性)
                Dim fg_USE_trKID25 As Boolean = $"{sm.UserInfo.Years}.{v_IR_AppStage}" <= "2026.1" OrElse cvKID25 <> "" '(2026上半年)強制使用trKID25 或有值
                ViewState(cst_VS_trKID25) = ""
                If fg_USE_trKID25 Then
                    ViewState(cst_VS_trKID25) = "Y" '(顯示)
                    trKID25.Visible = True '(顯示)
                    TIMS.Display_Inline(trKID25) '(顯示)
                    TIMS.Display_None(trKID26) '(不顯示)
                Else
                    '(其它狀況使用新-trKID26)
                    ViewState(cst_VS_trKID25) = "N" '(不顯示)
                    trKID25.Visible = False '(不顯示)
                    TIMS.Display_None(trKID25) '(不顯示)
                    TIMS.Display_Inline(trKID26) '(顯示)
                End If
            End If

            'If KID06 <> "" Then Common.SetListItem(ddlKID06, KID06)
            'If KID10 <> "" Then Common.SetListItem(ddlKID10, KID10)
            If Not (trKID19.Visible) Then KID19 = ""
            If KID19 <> "" Then Common.SetListItem(ddlKID19, KID19)
            If KID18 <> "" Then Common.SetListItem(ddlKID18, KID18)
            '2019年啟用 work2019x01:2019 政府政策性產業
            Call TIMS.SetCblValue(CBLKID20_1, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_2, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_3, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_4, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_5, cvKID20)
            Call TIMS.SetCblValue(CBLKID20_6, cvKID20)

            Call TIMS.SetCblValue(CBLKID60, cvKID60)

            Call TIMS.SetCblValue(CBLKID25_1, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_2, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_3, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_4, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_5, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_6, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_7, cvKID25)
            Call TIMS.SetCblValue(CBLKID25_8, cvKID25)

            '進階政策性產業類別
            If trKID25.Visible Then
                Dim v_KID22 As String = $"{dr2("KID22")}"
                Call TIMS.SetCblValue(CBLKID22B, v_KID22)
            Else
                Dim v_KID22 As String = $"{dr2("KID22")}"
                Call TIMS.SetCblValue(CBLKID22, v_KID22)
            End If

            Call TIMS.SetCblValue(CBLKID26_1, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_2, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_3, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_4, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_5, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_6, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_7, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_8, cvKID26)
            Call TIMS.SetCblValue(CBLKID26_9, cvKID26)
        End If
    End Sub
    ''' <summary> 顯示該計畫資料 PLAN_PLANINFO </summary>
    Sub SHOW_PLANPLANINFO()
        If g_flagNG Then
            sm.LastErrorMessage = cst_errmsg3
            Exit Sub
        End If

        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        Dim sRIDn As String = SUtl_GetRIDn()

        Dim hParms As New Hashtable From {{"PlanID", rqPlanID}, {"ComIDNO", rqComIDNO}, {"SeqNO", rqSeqNO}}
        Dim sql As String = ""
        sql &= " SELECT a.*" & vbCrLf
        sql &= " ,b.OrgName,c.RID RIDValue, b.OrgKind2" & vbCrLf
        sql &= " ,ISNULL(d.JobID, d.TrainID) JobID" & vbCrLf
        sql &= " ,ISNULL(d.JobName, d.TrainName) JobName" & vbCrLf
        sql &= " ,ISNULL(d.JobID, d.TrainID) TrainID" & vbCrLf
        sql &= " ,ISNULL(d.JobName, d.TrainName) TrainName" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN ORG_ORGINFO b ON a.ComIDNO=b.ComIDNO" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP c ON c.RID=a.RID AND c.OrgID=b.OrgID AND c.PlanID=a.PlanID" & vbCrLf
        sql &= " LEFT JOIN KEY_TRAINTYPE d ON a.TMID = d.TMID" & vbCrLf
        sql &= " LEFT JOIN SHARE_CJOB s ON s.CJOB_UNKEY=a.CJOB_UNKEY" & vbCrLf
        sql &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO" & vbCrLf
        Dim dr As DataRow = Nothing
        Try
            dr = DbAccess.GetOneRow(sql, objconn, hParms)
        Catch ex As Exception
            sm.LastErrorMessage = cst_errmsg4
            Dim strErrmsg As String = ""
            strErrmsg &= "/* sql: */" & vbCrLf
            strErrmsg &= sql & vbCrLf
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rqPlanID, rqComIDNO, rqSeqNO) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try

        Dim pms_PV As New Hashtable From {{"PlanID", rqPlanID}, {"ComIDNO", rqComIDNO}, {"SeqNO", rqSeqNO}}
        Dim sql_PV As String = " SELECT a.* FROM PLAN_VERREPORT a WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNO=@SeqNO"
        Dim drPV As DataRow = DbAccess.GetOneRow(sql_PV, objconn, pms_PV)
        If dr Is Nothing Then
            sm.LastErrorMessage = cst_errmsg5
            Exit Sub
        End If

        '辦理方式: '複製時清空 '辦理方式'遠距教學 'OJT-21102201：產投-班級複製作業：不要複製遠距教學相關欄位
        If gflag_ccopy Then dr("DISTANCE") = Convert.DBNull
        '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
        '2.功能：首頁>> 訓練機構管理 >> 班級複製作業 【包班種類】欄位不要複製，以避免將過去已填非包班的值copy過來。
        If gflag_ccopy AndAlso TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then dr("PackageType") = Convert.DBNull

        '辦理方式'預設值-遠距教學 DISTANCE／rbl_DISTANCE
        Common.SetListItem(rbl_DISTANCE, Convert.ToString(dr("DISTANCE")))
        'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        'If (Convert.ToString(dr("DISTANCE")).Equals("2")) Then rbl_DISTANCE.Enabled = False '(值不可更動)
        'rbl_DISTANCE.Visible = False '(不顯示選項)
        lab_DISTANCE.Visible = (Not rbl_DISTANCE.Visible)

        lab_DISTANCE.Text = TIMS.GET_DISTANCE_N(1, Convert.ToString(dr("DISTANCE")))
        'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        Hid_DISTANCE.Value = Convert.ToString(dr("DISTANCE"))

        FIXSUMCOST.Text = Convert.ToString(dr("FIXSUMCOST"))
        ACTHUMCOST.Text = Convert.ToString(dr("ACTHUMCOST"))
        '超出人時成本原因說明
        FIXExceeDesc.Text = Convert.ToString(dr("FIXExceeDesc"))
        METSUMCOST.Text = Convert.ToString(dr("METSUMCOST"))
        METCOSTPER.Text = Convert.ToString(dr("METCOSTPER")) '材料費占比
        '超出材料費比率上限原因說明
        METExceeDesc.Text = Convert.ToString(dr("METExceeDesc"))

        Dim PMS_PD1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}}
        Call SHOW_PLAN_DEPOT(PMS_PD1)

        Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, objconn)

        If $"{dr("PointYN")}" <> "" Then '是否為學分班
            Common.SetListItem(Radiobuttonlist1, dr("PointYN").ToString)
            Select Case Convert.ToString(dr("PointYN"))
                Case cst_學分班
                    Labmsg3.Text = Cst_msgother3b
                Case cst_非學分班
                    Labmsg3.Text = Cst_msgother3
            End Select
        End If
        tNote2.Text = Convert.ToString(dr("Note2")) 'cst_非學分班

        'Call GetNewtable() '建立空白表格Taddress2下拉選單用
        ViewState("dtTaddress") = GetNewtable()  '建立空白表格

        Dim sTMP_SciPlaceID As String = TIMS.ClearSQM(Convert.ToString(dr("SciPlaceID")))
        If sTMP_SciPlaceID <> "" Then
            Common.SetListItem(SciPlaceID, sTMP_SciPlaceID)
            If SciPlaceID.SelectedIndex = 0 Then
                SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, dr("ComIDNO"), 4, sTMP_SciPlaceID, objconn)
                Common.SetListItem(SciPlaceID, sTMP_SciPlaceID)
            End If
            TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, sTMP_SciPlaceID, 1, 1, objconn)
        End If

        Dim sTMP_TechPlaceID As String = TIMS.ClearSQM(Convert.ToString(dr("TechPlaceID")))
        If sTMP_TechPlaceID <> "" Then
            Common.SetListItem(TechPlaceID, sTMP_TechPlaceID)
            If TechPlaceID.SelectedIndex = 0 Then
                TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, dr("ComIDNO"), 4, sTMP_TechPlaceID, objconn)
                Common.SetListItem(TechPlaceID, sTMP_TechPlaceID)
            End If
            TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, sTMP_TechPlaceID, 2, 2, objconn)
        End If

        Dim sTMP_SciPlaceID2 As String = TIMS.ClearSQM(Convert.ToString(dr("SciPlaceID2")))
        If sTMP_SciPlaceID2 <> "" Then
            Common.SetListItem(SciPlaceID2, sTMP_SciPlaceID2)
            If SciPlaceID2.SelectedIndex = 0 Then
                SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, dr("ComIDNO"), 4, sTMP_SciPlaceID2, objconn)
                Common.SetListItem(SciPlaceID2, sTMP_SciPlaceID2)
            End If
            TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, sTMP_SciPlaceID2, 3, 1, objconn)
        End If

        Dim sTMP_TechPlaceID2 As String = TIMS.ClearSQM(Convert.ToString(dr("TechPlaceID2")))
        If sTMP_TechPlaceID2 <> "" Then
            Common.SetListItem(TechPlaceID2, sTMP_TechPlaceID2)
            If TechPlaceID2.SelectedIndex = 0 Then
                TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, dr("ComIDNO"), 4, sTMP_TechPlaceID2, objconn)
                Common.SetListItem(TechPlaceID2, sTMP_TechPlaceID2)
            End If
            TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, sTMP_TechPlaceID2, 4, 2, objconn)
        End If
        '遠距課程環境1/2
        Common.SetListItem(ddl_REMOTEID1, Convert.ToString(dr("RMTID")))
        'Hid_RMTID1.Value = Convert.ToString(dr("RMTID"))
        Common.SetListItem(ddl_REMOTEID2, Convert.ToString(dr("RMTID2")))
        'Hid_RMTID2.Value = Convert.ToString(dr("RMTID2"))

        hid_AddressSciPTID.Value = TIMS.ClearSQM(dr("AddressSciPTID"))
        hid_AddressSciPTID2.Value = TIMS.ClearSQM(dr("AddressSciPTID2"))
        hid_AddressTechPTID.Value = TIMS.ClearSQM(dr("AddressTechPTID"))
        hid_AddressTechPTID2.Value = TIMS.ClearSQM(dr("AddressTechPTID2"))

        'If Convert.ToString(dr("AddressSciPTID")) <> "" Then Common.SetListItem(Taddress2, dr("AddressSciPTID"))
        'If Convert.ToString(dr("AddressTechPTID")) <> "" Then Common.SetListItem(Taddress3, dr("AddressTechPTID"))
        'If dr("SciPlaceID").ToString <> "" OrElse dr("TechPlaceID").ToString <> "" OrElse dr("SciPlaceID2").ToString <> "" OrElse dr("TechPlaceID2").ToString <> "" Then
        '    RoomName.Enabled = False
        '    FactMode.Enabled = False
        '    FactModeOther.Enabled = False
        'Else
        '    RoomName.Enabled = False
        '    FactMode.Enabled = False
        '    FactModeOther.Enabled = False
        'End If

        RIDValue.Value = dr("RID").ToString
        ComidValue.Value = dr("ComIDNO").ToString
        center.Text = dr("orgname").ToString

        'PackageType 欄位是2011/5/12才加進去的,如果是舊資料才需帶IsBusiness
        IsBusiness.Checked = If(Convert.ToString(dr("IsBusiness")) <> "", If(Convert.ToString(dr("IsBusiness")) = "N", False, True), False)
        'IsBusiness.Checked = False
        'If Not IsDBNull(dr("IsBusiness")) Then
        '    IsBusiness.Checked = True
        '    If Convert.ToString(dr("IsBusiness")).ToString = "N" Then IsBusiness.Checked = False
        'End If
        'EnterpriseName.Text = dr("EnterpriseName").ToString
        FirstSort.Text = Convert.ToString(dr("FirstSort"))
        iCAPNUM.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(dr("iCAPNUM")))
        iCAPMARKDATE.Text = TIMS.Cdate3(dr("iCAPMARKDATE"))

        If dr("EnterSupplyStyle").ToString <> "" Then
            Common.SetListItem(EnterSupplyStyle, dr("EnterSupplyStyle").ToString)
            '1.報名時應先繳全額訓練費用，待結訓審核通過後核撥補助款
            '2.報名時應先繳50%訓練費用，待結訓審核通過後核撥補助款
            Select Case Convert.ToString(dr("OrgKind2"))
                Case "G" '非勞工團體
                    EnterSupplyStyle.Enabled = False
                Case "W" '勞工團體
                    EnterSupplyStyle.Enabled = True
            End Select
        End If

        '設定 GCID1Value.Value  '取得要比對的業別資料。
        Select Case strYears
            Case cst_strYears_2014 '"2014"
                If dr("GCID").ToString <> "" Then
                    GCIDValue.Value = dr("GCID").ToString
                    GCIDName.Text = TIMS.Get_GCIDName(dr("GCID").ToString, strYears, objconn)

                    Dim hPMS99 As New Hashtable From {{"GCID", Convert.ToString(dr("GCID"))}}
                    Dim sql99 As String = " SELECT GCODE1 FROM ID_GOVCLASSCAST WHERE GCID =@GCID"
                    Dim dr99 As DataRow = DbAccess.GetOneRow(sql99, objconn, hPMS99)
                    If dr99 IsNot Nothing Then GCID1Value.Value = Convert.ToString(dr99("GCode1"))
                End If
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    jobValue.Value = dr("TMID").ToString
                    TB_career_id.Text = String.Concat("[", dr("jobID"), "]", dr("jobName"))
                Else
                    trainValue.Value = dr("TMID").ToString
                    TB_career_id.Text = String.Concat("[", dr("TrainID"), "]", dr("TrainName"))
                End If

            Case cst_strYears_2015 '"2015"
                If dr("GCID2").ToString <> "" Then
                    GCIDValue.Value = dr("GCID2").ToString
                    GCIDName.Text = TIMS.Get_GCIDName(dr("GCID2").ToString, strYears, objconn)
                    Dim sql99 As String = " SELECT GCODE1 FROM V_GOVCLASSCAST2 WHERE GCID2 = " & Convert.ToString(dr("GCID2"))
                    Dim dr99 As DataRow = DbAccess.GetOneRow(sql99, objconn)
                    If Not dr99 Is Nothing Then GCID1Value.Value = Convert.ToString(dr99("GCODE1"))
                End If
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    jobValue.Value = dr("TMID").ToString
                    TB_career_id.Text = String.Concat("[", dr("jobID"), "]", dr("jobName"))
                Else
                    trainValue.Value = dr("TMID").ToString
                    TB_career_id.Text = String.Concat("[", dr("TrainID"), "]", dr("TrainName"))
                End If

            Case cst_strYears_2018 '"2018"
                'jobValue.Value = Convert.ToString(dr("TMID"))
                'jobValue.Value = ""
                If Convert.ToString(dr("TMID")) <> "" Then
                    'Dim sql99 As String = "SELECT '[' + c.GCODE2 + ']' + c.CNAME CNAME,GCID3 FROM V_GOVCLASSCAST3 c WHERE TMID=" & Convert.ToString(dr("TMID"))
                    Dim sql8 As String = ""
                    sql8 &= " SELECT '[' + c.GCODE2 + ']' + c.CNAME CNAME" & vbCrLf
                    sql8 &= " ,c.GCID3,c.TMID,c.PERC100,c.GCODE31" & vbCrLf
                    sql8 &= " FROM V_GOVCLASSCAST3 c" & vbCrLf
                    sql8 &= " WHERE TMID=" & Convert.ToString(dr("TMID"))
                    Dim dr8 As DataRow = DbAccess.GetOneRow(sql8, objconn)
                    If dr8 IsNot Nothing Then
                        TB_career_id.Text = Convert.ToString(dr8("CNAME"))
                        If Convert.ToString(dr8("GCID3")) <> "" Then
                            GCIDValue.Value = dr8("GCID3").ToString
                            GCIDName.Text = TIMS.Get_GCIDName(dr8("GCID3").ToString, strYears, objconn)
                            Hid_PERC100.Value = Convert.ToString(If(Convert.ToString(dr8("PERC100")) <> "", dr8("PERC100") * 100, 0))
                            GCID1Value.Value = Convert.ToString(dr8("GCODE31"))
                            jobValue.Value = ""
                            trainValue.Value = dr8("TMID")
                            TB_career_id.Text = TIMS.Get_GCIDName(GCIDValue.Value, strYears, objconn)
                        End If
                    End If
                End If
                If Convert.ToString(dr("GCID3")) <> "" Then
                    GCIDValue.Value = dr("GCID3").ToString
                    GCIDName.Text = TIMS.Get_GCIDName(dr("GCID3").ToString, strYears, objconn)
                    Dim sql99 As String = "SELECT TMID,PERC100,GCODE31 FROM V_GOVCLASSCAST3 WHERE GCID3=" & Convert.ToString(dr("GCID3"))
                    Dim dr99 As DataRow = DbAccess.GetOneRow(sql99, objconn)
                    If dr99 IsNot Nothing Then
                        Hid_PERC100.Value = Convert.ToString(If(Convert.ToString(dr99("PERC100")) <> "", dr99("PERC100") * 100, 0))
                        GCID1Value.Value = Convert.ToString(dr99("GCODE31"))
                        jobValue.Value = ""
                        trainValue.Value = dr99("TMID")
                        TB_career_id.Text = TIMS.Get_GCIDName(GCIDValue.Value, strYears, objconn)
                    End If
                End If
        End Select

        cjobValue.Value = Convert.ToString(dr("CJOB_UNKEY")) '.ToString
        txtCJOB_NAME.Text = TIMS.Get_CJOBNAME(dtSCJOB, cjobValue.Value)

        '訓練業別正確性:
        '產投 - 班級複製【訓練業別同意協助重新歸類】改為不複製
        Dim vTMIDCORRECT As String = If(gflag_ccopy AndAlso (Not gflag_can_copy3), "", Convert.ToString(dr("TMIDCORRECT")))
        Common.SetListItem(rblTMIDCORRECT, vTMIDCORRECT)

        '年度複製時要變動為登入年度
        Label3.Text = If(gflag_ccopy, sm.UserInfo.Years.ToString(), Convert.ToString(dr("PlanYear")))
        'If gflag_ccopy Then Label3.Text = sm.UserInfo.Years

        TIMS.PL_settextbox1(PlanCause, dr("PlanCause"))
        TIMS.PL_settextbox1(PurScience, dr("PurScience"))
        TIMS.PL_settextbox1(PurTech, dr("PurTech"))
        TIMS.PL_settextbox1(PurMoral, dr("PurMoral"))
        Common.SetListItem(Degree, Convert.ToString(dr("CapDegree")))
        'Common.SetListItem(AppStage, Convert.ToString(dr("AppStage")))

        '申請階段 '複製時清空(AppStage)
        Dim v_AppStage As String = ""
        If tr_AppStage_TP28.Visible Then
            '申請階段 '複製時清空(AppStage)
            v_AppStage = If(gflag_ccopy, "", $"{dr("AppStage")}")
            Common.SetListItem(rbl_AppStage, v_AppStage)
        End If

        TIMS.Display_None(sp_AppStage_1)
        TIMS.Display_None(sp_AppStage_2)
        Select Case v_AppStage
            Case "1"
                TIMS.Display_Inline(sp_AppStage_1)
            Case "2"
                TIMS.Display_Inline(sp_AppStage_2)
        End Select

        'tr_PolicyPreVal.Style.Item("display") = "none"
        'Select Case v_AppStage
        '    Case "3"
        '        tr_PolicyPreVal.Style.Remove("display")
        'End Select

        '不管什麼都是「年滿15歲以上」。
        rdoAge1.Checked = True
        rdoAge2.Checked = False
        txtAge1.Text = "" 'cst_AgeOtherDef
        If Convert.ToString(dr("CapAge1")) <> "" AndAlso TIMS.CINT1(dr("CapAge1")) >= cst_AgeOtherDef Then
            '若不是 年滿15歲以上 選擇顯示 目前所輸入的年齡。
            txtAge1.Text = Convert.ToString(dr("CapAge1"))
            rdoAge1.Checked = False
            rdoAge2.Checked = True
        End If

        '該欄位2015年後 暫不使用。
        Dim flag_NO_USE_OTHER As Boolean = (sm.UserInfo.Years >= TIMS.CINT1(cst_strYears_2015) OrElse gflag_ccopy)
        'If sm.UserInfo.Years >= Val(cst_strYears_2015) OrElse gflag_ccopy Then flag_NO_USE_OTHER = True
        Other1.Text = If(flag_NO_USE_OTHER, "", dr("CapOther1").ToString)  ' dr("CapOther1").ToString
        Other2.Text = If(flag_NO_USE_OTHER, "", dr("CapOther2").ToString)  ' dr("CapOther1").ToString
        Other3.Text = If(flag_NO_USE_OTHER, "", dr("CapOther3").ToString)  ' dr("CapOther1").ToString
        If Other1.Text = "" Then Other1.Text = Cst_msgother1
        If Other2.Text = "" Then Other2.Text = Cst_msgother1
        If Other3.Text = "" Then Other3.Text = Cst_msgother1
        TIMS.Tooltip(Other1, Cst_msgother1)
        TIMS.Tooltip(Other2, Cst_msgother1)
        TIMS.Tooltip(Other3, Cst_msgother1)

        Other1.Enabled = False
        Other2.Enabled = False
        Other3.Enabled = False
        GenSciHours.Text = dr("GenSciHours").ToString
        ProSciHours.Text = dr("ProSciHours").ToString
        SciHours.Text = Int(If(dr("GenSciHours").ToString = "", 0, Val(dr("GenSciHours")))) + TIMS.CINT1(If(dr("ProSciHours").ToString = "", 0, Val(dr("ProSciHours"))))
        ProTechHours.Text = dr("ProTechHours").ToString
        OtherHours.Text = dr("OtherHours").ToString
        TotalHours.Text = dr("TotalHours").ToString

        EMail.Text = dr("PlanEMail").ToString
        Hid_CredPoint.Value = Convert.ToString(dr("CredPoint"))
        CredPoint.Text = Convert.ToString(dr("CredPoint")) '.ToString
        RoomName.Text = dr("RoomName").ToString
        Common.SetListItem(FactMode, dr("FactMode").ToString)
        FactModeOther.Text = dr("FactModeOther").ToString
        '課程內容有室外教學 室外教學課程
        Common.SetListItem(rbl_OUTDOOR, Convert.ToString(dr("OUTDOOR")))
        ConNum.Text = dr("ConNum").ToString '容納人數必須為數字
        ContactName.Text = dr("ContactName").ToString '聯絡人 ContactName

        '2023/2024/ContactPhone
        ContactPhone.Text = Convert.ToString(dr("ContactPhone")) '電話
        Dim hCtPhone As New Hashtable
        TIMS.CHK_ContactPhoneFMT(Convert.ToString(dr("ContactPhone")), hCtPhone)
        ContactPhone_1.Text = hCtPhone("ContactPhone_1")
        ContactPhone_2.Text = hCtPhone("ContactPhone_2")
        ContactPhone_3.Text = hCtPhone("ContactPhone_3")
        Dim hCtMobile As New Hashtable
        TIMS.CHK_ContactMobileFMT(Convert.ToString(dr("ContactMobile")), hCtMobile)
        ContactMobile_1.Text = hCtMobile("ContactMobile_1")
        ContactMobile_2.Text = hCtMobile("ContactMobile_2")

        ContactEmail.Text = dr("ContactEmail").ToString
        ContactFax.Text = dr("ContactFax").ToString
        Common.SetListItem(ClassCate, Convert.ToString(dr("ClassCate")))
        Content.Text = dr("Content").ToString

        '複製狀態下,有些資料不複製 Start
        'Dim sRIDn As String = sUtl_GetRIDn()
        '如果是複製狀態, 則RID還是為原登入計畫之RID by nick
        RIDValue.Value = If(gflag_ccopy, sRIDn, Convert.ToString(dr("RID")))
        'If gflag_ccopy Then RIDValue.Value = sRIDn '如果是複製狀態, 則RID還是為原登入計畫之RID by nick

        PointType.SelectedIndex = -1
        PointName.Text = ""
        Dim v_PointType As String = TIMS.ClearSQM(dr("PointType"))
        Select Case v_PointType
            Case "1", "2", "3"
                Common.SetListItem(PointType, Convert.ToString(dr("PointType")))
                'Common.SetListItem(PointType, Convert.ToString(dr("PointType")).ToString)
                If PointType.SelectedItem IsNot Nothing Then PointName.Text = PointType.SelectedItem.Text  '學分種類名稱
            Case Else
                PointType.SelectedIndex = -1
        End Select

        '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
        PackageName.Text = ""
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        If Convert.ToString(dr("PackageType")) <> "" Then '包班種類名稱
            Common.SetListItem(PackageType, dr("PackageType").ToString)
            v_PackageType = TIMS.GetListValue(PackageType)
            Dim v_PackageTypeTxt As String = TIMS.GetListText(PackageType)
            If v_PackageType <> "" AndAlso v_PackageTypeTxt <> "" Then
                Select Case v_PackageType 'PackageType.SelectedValue
                    Case "1" '非包班
                    Case Else
                        PackageName.Text = "(" & v_PackageTypeTxt & ")" '包班種類名稱
                End Select
            End If
        End If

        ClassName.Text = TIMS.ClearSQM(dr("ClassName"))
        PointName.Text = TIMS.ClearSQM(PointName.Text)
        PackageName.Text = TIMS.ClearSQM(PackageName.Text)
        '取得班級名稱去掉學士學分班,碩士學分班,博士學分班
        If PointName.Text <> "" Then ClassName.Text = Replace(ClassName.Text, PointName.Text, "") '學分班種類
        '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
        If PackageName.Text <> "" Then ClassName.Text = Replace(ClassName.Text, PackageName.Text, "") '企業包班種類

        Class_Unit.Value = dr("Class_Unit").ToString
        TNum.Text = Convert.ToString(dr("TNum"))
        THours.Text = Convert.ToString(dr("THours"))
        CyclType.Text = TIMS.FmtCyclType(dr("CyclType"))
        ClassCount.Text = TIMS.ClearSQM(dr("ClassCount"))
        If ClassCount.Text = "" Then ClassCount.Text = "1"

        If Not gflag_ccopy Then
            STDate.Text = TIMS.Cdate3(dr("STDate"))
            FDDate.Text = TIMS.Cdate3(dr("FDDate"))

            DefGovCost.Text = dr("DefGovCost").ToString
            DefStdCost.Text = dr("DefStdCost").ToString
            Hid_TotalCost1.Value = dr("TotalCost").ToString

            '已存為正式資料，而且不是要複製計畫，草稿儲存功能不啟用
            If Convert.ToString(dr("IsApprPaper")) = "Y" Then
                Button8.Visible = False
                TIMS.Tooltip(Button8, "已存為正式資料，草稿儲存功能不啟用!")
            End If

            '2007以前的資料只可查詢，不可儲存 by AMU 2008-01-14
            If CStr(sm.UserInfo.Years) <= cst_Years_2007 Then
                If Convert.ToString(dr("IsApprPaper")) = "Y" Then
                    '已存為正式資料
                    'Button8.Visible = False '草稿儲存
                    'btnAdd.Visible = False '基本儲存
                    'BtnSAVE1.Visible = False '正式儲存
                    'btnAdd.Visible = False '正式儲存
                    'btnAdd.ToolTip += "本班已正式儲存，再次儲存請小心謹慎"
                    TIMS.Tooltip(btnAdd, "本班已正式儲存，再次儲存請小心謹慎")
                End If
                'If Not dr("IsApprPaper") = "Y" Then
                '    'Button8.Visible = True '草稿儲存
                '    'btnAdd.Visible = True '正式儲存
                'End If
            End If
            Button24.Visible = True '回上一頁

            '審核狀況。
            Select Case Convert.ToString(dr("AppliedResult"))
                Case "N"
                Case "M", ""
                    If Convert.ToString(dr("TransFlag")) = "Y" Then
                        Const cst_t_msg1 As String = "已轉班,不可修改"
                        center.Enabled = False
                        Org.Disabled = True
                        TIMS.Tooltip(center, cst_t_msg1)
                        TIMS.Tooltip(Org, cst_t_msg1)
                    End If
                Case Else '"Y" 審核狀況。通過
                    '2005/6/20--Melody審核通過or審核後修正者,不可修改班級名稱,期別,開結訓日,課程時數
                    '審核通過者,不可再修改班級名稱,期別,開結訓日,課程時數
                    Const cst_t_msg1 As String = "審核通過者,不可修改"
                    Dim fgIsApprPaper As Boolean = False
                    If (drPV IsNot Nothing) Then
                        fgIsApprPaper = (Convert.ToString(dr("IsApprPaper")) = "Y") AndAlso ((Convert.ToString(drPV("IsApprPaper")) = "Y"))
                    Else
                        fgIsApprPaper = (Convert.ToString(dr("IsApprPaper")) = "Y")
                    End If
                    If fgIsApprPaper Then
                        ClassName.ReadOnly = True
                        CyclType.ReadOnly = True
                        TIMS.Tooltip(ClassName, cst_t_msg1)
                        TIMS.Tooltip(CyclType, cst_t_msg1)
                        CustomValidator4.Enabled = False
                        STDate.ReadOnly = True
                        FDDate.ReadOnly = True
                        date1.Visible = False
                        date2.Visible = False
                        TIMS.Tooltip(STDate, cst_t_msg1)
                        TIMS.Tooltip(FDDate, cst_t_msg1)
                        SciHours.ReadOnly = True
                        GenSciHours.ReadOnly = True
                        ProSciHours.ReadOnly = True
                        ProTechHours.ReadOnly = True
                        OtherHours.ReadOnly = True
                        TotalHours.ReadOnly = True
                        THours.ReadOnly = True
                        TIMS.Tooltip(SciHours, cst_t_msg1)
                        TIMS.Tooltip(GenSciHours, cst_t_msg1)
                        TIMS.Tooltip(ProSciHours, cst_t_msg1)
                        TIMS.Tooltip(ProTechHours, cst_t_msg1)
                        TIMS.Tooltip(OtherHours, cst_t_msg1)
                        TIMS.Tooltip(TotalHours, cst_t_msg1)
                        TIMS.Tooltip(THours, cst_t_msg1)
                    End If

                    If Convert.ToString(dr("AppliedResult")) = "Y" Then
                        If iPlanKind = 1 AndAlso sm.UserInfo.LID = 2 AndAlso Convert.ToString(dr("TransFlag")) = "N" Then Disabled_Items("委訓單位限制") '委訓單位限制
                        If iPlanKind = 2 AndAlso sm.UserInfo.LID = 2 Then Disabled_Items("計畫種類為委辦") '計畫種類為委外者
                        '---Gloria 2007/8/30同意分署(中心)可在審核後，再次修改---
                        Dim s_RoleTxt As String = ""
                        Dim flag_can_update1 As Boolean = False
                        If flag_IsSuperUser_1 Then s_RoleTxt = "系統管理者"
                        If sm.UserInfo.LID = 1 Then s_RoleTxt = "分署"
                        If flag_IsSuperUser_1 OrElse sm.UserInfo.LID = 1 Then flag_can_update1 = True
                        If flag_can_update1 Then
                            Dim str_t_msg2 As String = String.Format("(使用者權限:{0})可修改", s_RoleTxt)
                            ClassName.ReadOnly = False '班別名稱
                            STDate.ReadOnly = False '訓練起日
                            FDDate.ReadOnly = False '訓練迄日
                            date1.Visible = True
                            date2.Visible = True
                            CyclType.ReadOnly = False '期別
                            ClassCount.ReadOnly = False '班數
                            TIMS.Tooltip(ClassName, str_t_msg2)
                            TIMS.Tooltip(STDate, str_t_msg2)
                            TIMS.Tooltip(FDDate, str_t_msg2)
                            TIMS.Tooltip(CyclType, str_t_msg2)
                            TIMS.Tooltip(ClassCount, str_t_msg2)
                        End If
                    End If

            End Select

        End If
        '複製狀態下,有些資料不複製 End
        'If Request(cst_ccopy) = "1" Then'gflag_ccopy
        'Else
        'End If

        If TIMS.ClearSQM(Request("todo")) = 1 Then Disabled_Items("僅顯示") '按鈕狀態控制
    End Sub

    ''' <summary>機構黑名單內容(訓練單位處分功能) block black</summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function Chk_OrgBlackList(ByRef Errmsg As String) As Boolean
        Dim rst As Boolean = False
        Errmsg = ""
        Dim vsComIDNO As String = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        If TIMS.Check_OrgBlackList(Me, vsComIDNO, objconn) Then
            rst = True
            Errmsg = $"{sm.UserInfo.OrgName}，已列入處分名單!!"
            isBlack.Value = "Y"
            Blackorgname.Value = sm.UserInfo.OrgName
            Return rst
        ElseIf Hid_ComIDNO.Value <> "" AndAlso TIMS.Check_OrgBlackList(Me, Hid_ComIDNO.Value, objconn) Then
            Dim drO2 As DataRow = TIMS.Get_ORGINFOdr2(Hid_ComIDNO.Value, objconn)
            Dim v_ORGNAME As String = If(drO2 IsNot Nothing, $"{drO2("ORGNAME")}", Hid_ComIDNO.Value)
            rst = True
            Errmsg = $"{v_ORGNAME}，已列入處分名單!!"
            isBlack.Value = "Y"
            Blackorgname.Value = v_ORGNAME 'sm.UserInfo.OrgName
            Return rst
        End If
        Return rst
    End Function

    ''' <summary>建立下拉選單物件</summary>
    Sub CCreateItem()
        'Const Cst_EmptySelValue As String = "==無==" 'TIMS.cst_ddl_PleaseChoose3
        TIMS.PL_placeholder(PCont)
        TIMS.PL_placeholder(PlanCause)
        TIMS.PL_placeholder(PurScience)
        TIMS.PL_placeholder(PurTech)
        TIMS.PL_placeholder(PurMoral)

        Call TIMS.CHG_ddlDEPOT15(ddlDEPOT15, sm.UserInfo.DistID, sm.UserInfo.TPlanID, objconn) '轄區重點產業(排除停用)(轄區)

        Call TIMS.Tooltip(EHOUR, cst_EHour_t1, True)
        Call TIMS.Tooltip(AIAHOUR, cst_AIAHOUR_t1, True)
        Call TIMS.Tooltip(WNLHOUR, cst_WNLHOUR_t1, True)

        '有啟動才檢核/ 塞值
        'If tr_AppStage_TP28.Visible Then
        '    AppStage = If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(AppStage), TIMS.Get_AppStage(AppStage))
        'End If

        '申請階段／'申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇) 
        If tr_AppStage_TP28.Visible Then
            Dim htParms As New Hashtable From {{"YEARS", CStr(sm.UserInfo.Years)}}
            Dim dtAppStage As DataTable = TIMS.GET_PLAN_APPSTAGE_PTYPE01_dt(objconn, htParms)
            rbl_AppStage = TIMS.Get_APPSTAGE2(rbl_AppStage, dtAppStage) 'If(sm.UserInfo.Years >= 2018, TIMS.Get_APPSTAGE2(rbl_AppStage), TIMS.Get_AppStage(rbl_AppStage))
        End If

        '1.為讓訓練單位了解審查計分表計算區間，於「班別資料」頁籤增加說明文字，顯示位置如圖一。
        '並請依不同申請階段， 顯示不同之說明文字。僅上、下半年， 政策性產業 / 進階政策性產業： 不用顯示。
        TIMS.Display_None(sp_AppStage_1)
        TIMS.Display_None(sp_AppStage_2)

        '訓練單位反映於班級申請裡面有一欄「優先排序」 (下圖)，此欄可由單位自行決定此班於後續分署核班時的優先順序。
        '惟如單位於申請階段「上半年」已提案3班，優先排序各為1、2、3，但申請階段於「政策性產業」或「下半年」時，又提案3班，這時優先排序卻不能再打1、2、3，
        '爰請調整邏輯， 每個申請階段之班級優先排序， 都可從1開始。
        Call TIMS.Tooltip(FirstSort, cst_errmsg35, True)

        Dim sql As String = ""
        'ddlKID06.Items.Clear() '2015 六大新興產業 (DEPID='10')
        'sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID='10' ORDER BY KID "
        'DbAccess.MakeListItem(ddlKID06, sql, objconn)
        'ddlKID06.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        '2017 (重點服務業) SELECT * FROM KEY_DEPOT WHERE DEPID='16' 
        'ddlKID10.Items.Clear()
        'sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID='16' ORDER BY KID "
        'DbAccess.MakeListItem(ddlKID10, sql, objconn)
        'ddlKID10.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        'https://jira.turbotech.com.tw/browse/TIMSC-276
        '2018 (政府政策性產業)
        ddlKID19.Items.Clear()
        If (trKID19.Visible) Then
            sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID='19' ORDER BY KID "
            DbAccess.MakeListItem(ddlKID19, sql, objconn)
            ddlKID19.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))
        End If
        '2017 (新南向政策) SELECT * FROM KEY_DEPOT WHERE DEPID='18' 
        ddlKID18.Items.Clear()
        sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID='18' ORDER BY KID "
        DbAccess.MakeListItem(ddlKID18, sql, objconn)
        ddlKID18.Items.Insert(0, New ListItem(TIMS.cst_EmptySelValue, ""))

        'trKID20.Visible
        '2018 (政府政策性產業)
        'sql = " SELECT KID, KNAME FROM KEY_BUSINESS WHERE DEPID='20' ORDER BY KID"
        'Dim dtKID_N20 As DataTable = DbAccess.GetDataTable(sql, objconn)
        Dim dtKID_N20 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "20")
        Call TIMS.GET_CBL_KID20(CBLKID20_1, dtKID_N20, 1)
        Call TIMS.GET_CBL_KID20(CBLKID20_2, dtKID_N20, 2)
        Call TIMS.GET_CBL_KID20(CBLKID20_3, dtKID_N20, 3)
        Call TIMS.GET_CBL_KID20(CBLKID20_4, dtKID_N20, 4)
        Call TIMS.GET_CBL_KID20(CBLKID20_5, dtKID_N20, 5)
        Call TIMS.GET_CBL_KID20(CBLKID20_6, dtKID_N20, 6)

        '進階政策性產業類別
        Dim dtKID_N22 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "22")
        Call TIMS.GET_CBL_KID22(CBLKID22, dtKID_N22)
        Call TIMS.GET_CBL_KID22(CBLKID22B, dtKID_N22)

        'CheckBoxList 選項設定-政府政策性產業 2025
        Dim dtKID_N25 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "25")
        Call TIMS.GET_CBL_KID25(CBLKID25_1, dtKID_N25, 1)
        Call TIMS.GET_CBL_KID25(CBLKID25_2, dtKID_N25, 2)
        Call TIMS.GET_CBL_KID25(CBLKID25_3, dtKID_N25, 3)
        Call TIMS.GET_CBL_KID25(CBLKID25_4, dtKID_N25, 4)
        Call TIMS.GET_CBL_KID25(CBLKID25_5, dtKID_N25, 5)
        Call TIMS.GET_CBL_KID25(CBLKID25_6, dtKID_N25, 6)
        Call TIMS.GET_CBL_KID25(CBLKID25_7, dtKID_N25, 7)
        Call TIMS.GET_CBL_KID25(CBLKID25_8, dtKID_N25, 8)

        'CheckBoxList 選項設定-政府政策性產業 2026,FN_GET_KID
        Dim dtKID_N26 As DataTable = TIMS.Get_BUSINESS_KID_dt(objconn, "26")
        Call TIMS.GET_CBL_KID26(CBLKID26_1, dtKID_N26, 1)
        Call TIMS.GET_CBL_KID26(CBLKID26_2, dtKID_N26, 2)
        Call TIMS.GET_CBL_KID26(CBLKID26_3, dtKID_N26, 3)
        Call TIMS.GET_CBL_KID26(CBLKID26_4, dtKID_N26, 4)
        Call TIMS.GET_CBL_KID26(CBLKID26_5, dtKID_N26, 5)
        Call TIMS.GET_CBL_KID26(CBLKID26_6, dtKID_N26, 6)
        Call TIMS.GET_CBL_KID26(CBLKID26_7, dtKID_N26, 7)
        Call TIMS.GET_CBL_KID26(CBLKID26_8, dtKID_N26, 8)
        Call TIMS.GET_CBL_KID26(CBLKID26_9, dtKID_N26, 9)

        '產業別(管考)
        If fg_USE_CBLKID60_TP28 Then
            CBLKID60.Items.Clear()
            sql = "SELECT KID,KNAME FROM VIEW_DEPOT60 ORDER BY KID"
            DbAccess.MakeListItem(CBLKID60, sql, objconn)
        End If

        '六大職能別查詢清單
        Call TIMS.Get_ClassCatelog(ClassCate, objconn)
        Weeks = TIMS.Get_ddlWeeks(Weeks)

        Call CreateTimesItem(ddlpnH1, ddlpnH2, ddlpnM1, ddlpnM2) '設定時間物件值（DropDownList）

        Degree = TIMS.Get_Degree(Degree, 2, objconn)

        '將ComidValue.Value 塞入有效值
        If ComidValue.Value = "" Then
            ComidValue.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", ComidValue.Value)
            If ComidValue.Value = "" Then ComidValue.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        End If

        '班級-辦理方式-遠距教學 'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        rbl_DISTANCE = TIMS.GET_DISTANCE(rbl_DISTANCE, 2)
        'rbl_DISTANCE.Visible = False
        lab_DISTANCE.Visible = (Not rbl_DISTANCE.Visible)

        '空值，請顯示 整班為實體教學
        Dim s_DISTANCE_N_def As String = TIMS.GET_DISTANCE_N(1, "3")
        lab_DISTANCE.Text = s_DISTANCE_N_def 'lab_DISTANCE.Text = "(不提供該選項)"
        TIMS.Tooltip(lab_DISTANCE, "(不提供該選項)", True)

        '班級-辦理方式-'產投使用／遠距教學 暫不啟用
        If flag_StopDISTANCE2 Then tr_rbl_DISTANCE.Visible = False
        If flag_StopDISTANCE2 Then td_cbFARLEARN_h.Visible = False
        td_cbFARLEARN_d.Visible = td_cbFARLEARN_h.Visible

        '班級-辦理方式-'署的權限可以修改遠距教學
        cbFARLEARN.Visible = If(gflag_DISTANCE_can_updata, True, False) 'False"(不提供該選項)"
        lab_cbFARLEARN.Text = If(gflag_DISTANCE_can_updata, "", "(無)")
        If Not gflag_DISTANCE_can_updata Then TIMS.Tooltip(lab_cbFARLEARN, "(不提供該選項)", True)
        If gflag_DISTANCE_can_updata Then TIMS.Tooltip(cbFARLEARN, "業務權限可修改辦理方式", True)

        SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, ComidValue.Value, 2, "", objconn)
        SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, ComidValue.Value, 2, "", objconn)

        TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, ComidValue.Value, 2, "", objconn)
        TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, ComidValue.Value, 2, "", objconn)

        TIMS.Tooltip(SciPlaceID, "學科場地以登入者的機構為準")
        TIMS.Tooltip(TechPlaceID, "術科場地以登入者的機構為準")
        TIMS.Tooltip(SciPlaceID2, "學科場地以登入者的機構為準")
        TIMS.Tooltip(TechPlaceID2, "術科場地以登入者的機構為準")
        '遠距課程環境1/2 ORG_REMOTER
        ddl_REMOTEID1 = TIMS.GET_REMOTEID(ddl_REMOTEID1, ComidValue.Value, objconn)
        ddl_REMOTEID2 = TIMS.GET_REMOTEID(ddl_REMOTEID2, ComidValue.Value, objconn)

        '建立教師Script
        If RIDValue.Value <> "" Then
            'exErrmsg &= "RIDValue.Value : " & RIDValue.Value & vbCrLf
            TIMS.CreateTeacherScript(Me, RIDValue.Value, objconn)
        Else
            If sm.IsLogin Then
                'exErrmsg &= "sm.(""RID"") : " & Convert.ToString(sm.UserInfo.RID) & vbCrLf
                TIMS.CreateTeacherScript(Me, sm.UserInfo.RID, objconn)
            End If
        End If

        ' Classification1.Attributes.Add("onchange", "javascript:showPTID('Classification1','PTID1','PTID2');Layer_change(5);")
        Classification1.Attributes.Add("onchange", "javascript:showPTID('Classification1','PTID1','PTID2');")
        Radiobuttonlist1.Attributes.Add("onclick", "javascript:Layer_change(5);showCostType('" & Radiobuttonlist1.ClientID & "');")

        '任課教師
        OLessonTeah1.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addx','OLessonTeah1','OLessonTeah1Value');") 'SD/04/LessonTeah1.aspx
        OLessonTeah1.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah1Value','OLessonTeah1');"
        OLessonTeah1.Style.Item("CURSOR") = "hand"

        '助教
        OLessonTeah2.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addy','OLessonTeah2','OLessonTeah2Value');") 'SD/04/LessonTeah1.aspx
        OLessonTeah2.Attributes("onchange") = "GetTeacherId(this.value,'OLessonTeah2Value','OLessonTeah2');"
        OLessonTeah2.Style.Item("CURSOR") = "hand"

        'Dim selsql As String = ""
        'selsql = "SELECT COSTID+','+ITEMCOSTNAME COSTID,COSTNAME FROM KEY_COSTITEM2 ORDER BY SORT"
        'btnAdd.Attributes("onclick") = "javascript:notshow_button(); init(); window.setTimeout('show_secs()',1);"
        'Button8.Attributes("onclick") = "javascript:notshow_button(); init(); window.setTimeout('show_secs()',1);"
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'Dim rqRID As String = Convert.ToString(sm.UserInfo.RID)
        'If RIDValue.Value <> "" Then rqRID = RIDValue.Value
        'Dim sWOScript1 As String = ""
        'sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & rqRID & "&TB1=" & TeacherDesc_A.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
        'btn_TCTYPEA.Attributes("onclick") = sWOScript1
        'sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & rqRID & "&TB1=" & TeacherDesc_B.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
        'btn_TCTYPEB.Attributes("onclick") = sWOScript1
    End Sub
    ''' <summary>設定時間物件值（DropDownList）</summary>
    ''' <param name="oddlTimesH1"></param>
    ''' <param name="oddlTimesH2"></param>
    ''' <param name="oddlTimesM1"></param>
    ''' <param name="oddlTimesM2"></param>
    Sub CreateTimesItem(ByRef oddlTimesH1 As DropDownList, ByRef oddlTimesH2 As DropDownList, ByRef oddlTimesM1 As DropDownList, ByRef oddlTimesM2 As DropDownList)
        oddlTimesH1.Items.Clear()
        oddlTimesH2.Items.Clear()
        oddlTimesM1.Items.Clear()
        oddlTimesM2.Items.Clear()

        For intTimeHM As Integer = 0 To 22
            If intTimeHM >= 0 AndAlso intTimeHM <= 5 Then
                If intTimeHM = 0 Then
                    oddlTimesM1.Items.Add(New ListItem("00", "00"))
                    oddlTimesM2.Items.Add(New ListItem("00", "00"))
                Else
                    oddlTimesM1.Items.Add(New ListItem(CStr(intTimeHM * 10), CStr(intTimeHM * 10)))
                    oddlTimesM2.Items.Add(New ListItem(CStr(intTimeHM * 10), CStr(intTimeHM * 10)))
                End If
            End If
            If intTimeHM >= 8 AndAlso intTimeHM <= 22 Then
                If CStr(intTimeHM).Length < 2 Then
                    oddlTimesH1.Items.Add(New ListItem("0" & CStr(intTimeHM), "0" & CStr(intTimeHM)))
                    oddlTimesH2.Items.Add(New ListItem("0" & CStr(intTimeHM), "0" & CStr(intTimeHM)))
                Else
                    oddlTimesH1.Items.Add(New ListItem("" & CStr(intTimeHM), "" & CStr(intTimeHM)))
                    oddlTimesH2.Items.Add(New ListItem("" & CStr(intTimeHM), "" & CStr(intTimeHM)))
                End If
            End If
        Next
    End Sub
    ''' <summary>'建立計畫 企業包班事業單位</summary>
    Sub CreateBusPackage()
        Datagrid4headTable.Visible = False '若選擇非包班，則不能再選企業包班喔

        If Not TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '非充電計畫者 不可做企業包班新增
            Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
            Exit Sub
        End If

        If hTPlanID54.Value = "" Then '非  '充電起飛計畫  (hTPlanID54.Value = "1")
            If Session(hid_PLAN_BUSPACKAGE_guid1.Value) IsNot Nothing Then Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
            Exit Sub
        End If

        '充電起飛計畫 '非 聯合企業包班
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        Select Case v_PackageType 'PackageType.SelectedValue
            Case "3"  '充電起飛計畫' 聯合企業包班
                Datagrid4headTable.Visible = True
                btnAddBusPackage.Style.Item("display") = ""
            Case "2"  '充電起飛計畫' 企業包班
                Datagrid4headTable.Visible = True
                btnAddBusPackage.Style.Item("display") = "none"
            Case Else
                If Session(hid_PLAN_BUSPACKAGE_guid1.Value) IsNot Nothing Then Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
                Exit Sub
        End Select

        Const Cst_PKName As String = "BPID"
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        If Not TIMS.IS_DataTable(Session(hid_PLAN_BUSPACKAGE_guid1.Value)) Then
            Dim dt1 As DataTable = Nothing
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If g_flagNG Then sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE 1<>1"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE 1<>1 "
                Else
                    sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If g_flagNG Then sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
            dt.Columns(Cst_PKName).AutoIncrement = True
            dt.Columns(Cst_PKName).AutoIncrementSeed = -1
            dt.Columns(Cst_PKName).AutoIncrementStep = -1
            If gflag_ccopy Then Call TIMS.CopyDATATABLE(dt, dt1, Cst_PKName)
        Else
            dt = Session(hid_PLAN_BUSPACKAGE_guid1.Value)
        End If
        If dt Is Nothing Then Return

        Session(hid_PLAN_BUSPACKAGE_guid1.Value) = dt
        Datagrid4Table.Visible = False

        If dt.Rows.Count > 0 Then
            txtUname.Text = ""
            txtIntaxno.Text = ""
            txtUbno.Text = ""
            'dr1 As DataRow 
            For Each dr In dt.Rows
                If Not dr.RowState = DataRowState.Deleted Then
                    'Dim v_PackageType As String = TIMS.GetListValue(PackageType)
                    Select Case v_PackageType 'PackageType.SelectedValue
                        Case "2"  '充電起飛計畫' 企業包班
                            'dr = dt.Rows(0)
                            txtUname.Text = Convert.ToString(dr("UName"))
                            txtIntaxno.Text = Convert.ToString(dr("Intaxno"))
                            txtUbno.Text = Convert.ToString(dr("Ubno"))
                            Exit For
                    End Select
                End If
            Next

            Datagrid4Table.Visible = True
            Datagrid4.DataSource = dt
            Datagrid4.DataBind()
        End If
    End Sub
    ''' <summary>'建立上課時間</summary>
    Sub CreateClassTime()
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing
        If Not TIMS.IS_DataTable(Session(hid_planONCLASS_guid1.Value)) Then
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If (Not gflag_can_copy1) Then sql &= " AND 1<>1"
                    If g_flagNG Then sql = " SELECT * FROM PLAN_ONCLASS WHERE 1<>1"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT * FROM PLAN_ONCLASS WHERE 1<>1 "
                Else
                    sql = " SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If g_flagNG Then sql = " SELECT * FROM PLAN_ONCLASS WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(hid_planONCLASS_guid1.Value)
        End If
        If dt Is Nothing Then Return

        dt.Columns("POCID").AutoIncrement = True
        dt.Columns("POCID").AutoIncrementSeed = -1
        dt.Columns("POCID").AutoIncrementStep = -1
        If gflag_ccopy Then Call TIMS.CopyDATATABLE(dt, dt1, "POCID")

        Session(hid_planONCLASS_guid1.Value) = dt
        DataGrid1Table.Visible = False
        If dt.Rows.Count > 0 Then
            DataGrid1Table.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If

        Dim strCheckAddTime As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim flag_canUse As Boolean = True
            Dim Weeks1 As Label = eItem.FindControl("Weeks1")
            Dim Times1 As Label = eItem.FindControl("Times1")
            If Weeks1 Is Nothing Then flag_canUse = False
            If Times1 Is Nothing Then flag_canUse = False
            If (flag_canUse) Then
                Dim WeeksSedIdx As String = Get_WeekSedIdx(Weeks, Weeks1.Text)
                Dim xAddValue As String = "Ws:" & WeeksSedIdx & "/Ts:" & Times1.Text
                If (strCheckAddTime <> "") Then strCheckAddTime &= ","
                strCheckAddTime &= xAddValue
            End If
        Next
        If (strCheckAddTime <> "") Then Hid_CheckAddTime.Value = strCheckAddTime
    End Sub
    ''' <summary>重建 計畫訓練內容簡介'PLAN_TRAINDESC 依SESSION</summary>
    Sub CreateTrainDesc()
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing
        If Not TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If (Not gflag_can_copy1) Then sql &= " AND 1<>1"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    '辦理方式 '遠距教學 'OJT-21102201：產投-班級複製作業：不要複製遠距教學相關欄位
                    For Each dr1 As DataRow In dt1.Rows : dr1("FARLEARN") = Convert.DBNull : Next
                    sql = " SELECT * FROM PLAN_TRAINDESC WHERE 1<>1" & vbCrLf
                Else
                    sql = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If g_flagNG Then sql = " SELECT * FROM PLAN_TRAINDESC WHERE 1<>1"
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(hid_TrainDescTable_guid1.Value)
        End If
        If dt Is Nothing Then Return
        dt.Columns("PTDID").AutoIncrement = True
        dt.Columns("PTDID").AutoIncrementSeed = -1
        dt.Columns("PTDID").AutoIncrementStep = -1
        '20090401 andy edit 加入排序
        dt.DefaultView.Sort = "STrainDate asc,PName asc,PTDID asc"
        dt = dt.DefaultView.Table

        'Session(hid_TrainDescTable_guid1.Value)
        If dt1 IsNot Nothing AndAlso dt1.Rows.Count > 0 AndAlso gflag_ccopy AndAlso Not gflag_TrainDesc_edit1 Then
            For Each dr1 As DataRow In dt1.Rows
                If Not dr1.RowState = DataRowState.Deleted Then
                    Dim dr As DataRow = dt.NewRow
                    dt.Rows.Add(dr)
                    'dr(Cst_OtherCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_OtherCostpkName)
                    For i As Integer = 0 To dr1.ItemArray.Length - 1
                        dr(dr.Table.Columns(i).ColumnName) = dr1(dr.Table.Columns(i).ColumnName)
                    Next
                    '清除日期/師資-複製的資料
                    dr("PTDID") = 0 - TIMS.CINT1(dr("PTDID")) '變負數
                    dr("STRAINDATE") = Convert.DBNull
                    dr("ETRAINDATE") = Convert.DBNull
                    dr("TECHID") = Convert.DBNull
                    dr("TECHID2") = Convert.DBNull
                End If
            Next
            'dt.AcceptChanges()
        End If
        'If dt.Rows.Count > 0 AndAlso gflag_ccopy AndAlso Not gflag_TrainDesc_edit1 Then
        '    '清除日期/師資-複製的資料
        '    For Each dr As DataRow In dt.Rows
        '        dr("PTDID") = 0 - Val(dr("PTDID")) '變負數
        '        dr("STRAINDATE") = Convert.DBNull
        '        dr("ETRAINDATE") = Convert.DBNull
        '        dr("TECHID") = Convert.DBNull
        '        dr("TECHID2") = Convert.DBNull
        '    Next
        '    dt.AcceptChanges()
        'End If
        Session(hid_TrainDescTable_guid1.Value) = dt

        '0:日期  '1:授課時段	'2:授課時間	'3:時數	'4:技檢訓練時數	'5:課程進度/內容 
        '6:學/術科	'7:上課地點	'8:遠距教學	'9:室外教學	'10:授課師資 '11:助教 '12:功能
        '技檢訓練時數 產業人才投資方案 使用(充飛不使用)
        Const cst_dg3_col_i技檢訓練時數 As Integer = 4
        '產投使用／遠距教學 暫不啟用
        Const cst_dg3_col_i遠距教學 As Integer = 8

        '技檢訓練時數 產業人才投資方案 使用(充飛不使用)
        Datagrid3.Columns(cst_dg3_col_i技檢訓練時數).Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        '產投使用／遠距教學 暫不啟用
        Datagrid3.Columns(cst_dg3_col_i遠距教學).Visible = If(flag_StopDISTANCE2, False, True)

        Datagrid3Table.Visible = False
        Datagrid3Table.Style.Item("display") = "none"

        If dt.Rows.Count > 0 Then
            Datagrid3Table.Visible = True
            Datagrid3Table.Style.Item("display") = ""
            With Datagrid3
                .DataSource = dt
                .DataKeyField = "PTDID"
                .DataBind()
            End With
        End If

        Dim xAddValue As String = ""
        Dim StrChkTDescH1 As String = ""
        Dim StrChkTDescH2 As String = ""
        'Dim iSciHours As Integer = 0 'SciHours-學科 
        Dim iGenSciHours As Integer = 0 'GenSciHours-一般學科 
        Dim iProSciHours As Integer = 0 'ProSciHours-專業學科(0)
        Dim iProTechHours As Integer = 0 'ProTechHours-術科 
        Dim iOtherHours As Integer = 0 'OtherHours-其他時數(0)
        Dim iTotalHours As Integer = 0 'TotalHours-總計 

        For Each eItem As DataGridItem In Datagrid3.Items
            Dim PHourLabel As Label = eItem.FindControl("PHourLabel") '時數
            Dim PHourTxt As TextBox = eItem.FindControl("PHourTxt") '時數(edit)
            'Dim EHourLabel As Label = eItem.FindControl("EHourLabel") '技檢訓練時數
            'Dim EHourTxt As TextBox = eItem.FindControl("EHourTxt") '技檢訓練時數(edit)
            Dim drpClassification1 As DropDownList = eItem.FindControl("drpClassification1") '1:學/2:術
            Dim drpClassEdit As DropDownList = eItem.FindControl("drpClassEdit") '1:學/2:術(edit)
            Dim iPHour As Integer = 0
            If iPHour = 0 AndAlso PHourLabel IsNot Nothing Then iPHour = If(TIMS.IsNumeric1(PHourLabel.Text), TIMS.CINT1(PHourLabel.Text), 0)
            If iPHour = 0 AndAlso PHourTxt IsNot Nothing Then iPHour = If(TIMS.IsNumeric1(PHourTxt.Text), TIMS.CINT1(PHourTxt.Text), 0)

            Dim v_drpClassification1 As String = TIMS.GetListValue(drpClassification1)
            Dim v_drpClassEdit As String = TIMS.GetListValue(drpClassEdit)
            If v_drpClassification1 = "" AndAlso v_drpClassEdit <> "" Then v_drpClassification1 = v_drpClassEdit '若為空且不為空，才執行
            iTotalHours += iPHour
            Select Case v_drpClassification1
                Case "1" '1:學/2:術
                    'iSciHours += iPHour '一般學科 
                    iGenSciHours += iPHour'一般學科 
                Case "2" '1:學/2:術
                    iProTechHours += iPHour
            End Select

            Dim flag_canUse As Boolean = True
            Dim STrainDateLabel As Label = eItem.FindControl("STrainDateLabel")
            Dim PNameLabel As Label = eItem.FindControl("PNameLabel")
            If STrainDateLabel Is Nothing Then flag_canUse = False
            If PNameLabel Is Nothing Then flag_canUse = False
            If (flag_canUse) Then
                xAddValue = "SD:" + STrainDateLabel.Text + "/H1:" + Mid(PNameLabel.Text, 1, 2)
                If (StrChkTDescH1 <> "") Then StrChkTDescH1 &= ","
                StrChkTDescH1 &= xAddValue
                xAddValue = "SD:" + STrainDateLabel.Text + "/H2:" + Mid(PNameLabel.Text, 7, 2)
                If (StrChkTDescH2 <> "") Then StrChkTDescH2 &= ","
                StrChkTDescH2 &= xAddValue
            End If
        Next

        'SciHours.Text = CStr(iSciHours) 'SciHours-學科 
        GenSciHours.Text = CStr(iGenSciHours) 'GenSciHours-一般學科 
        ProSciHours.Text = "0" 'CStr(iProSciHours) 'ProSciHours-專業學科(0)
        ProTechHours.Text = CStr(iProTechHours) 'ProTechHours-術科 
        OtherHours.Text = "0" 'CStr(iOtherHours) 'OtherHours-其他時數(0)
        TotalHours.Text = CStr(iTotalHours) 'TotalHours-總計 

        'If (StrChkTDescH1 <> "") Then Hid_ChkTDescH1.Value = StrChkTDescH1 '開始時間比對值
        'If (StrChkTDescH2 <> "") Then Hid_ChkTDescH2.Value = StrChkTDescH2 '結束時間比對值
        Hid_ChkTDescH1.Value = StrChkTDescH1 '開始時間比對值
        Hid_ChkTDescH2.Value = StrChkTDescH2 '結束時間比對值
        Dim StrChkTDescH3 As String = Get_TDescH3()
        'If (StrChkTDescH3 <> "") Then Hid_ChkTDescH3.Value = StrChkTDescH3 '組合每日課節資訊
        Hid_ChkTDescH3.Value = StrChkTDescH3 '組合每日課節資訊 Chk_doubleDAY

        'If flag_TIMS_Test_1 Then Hid_ChkTEACHHOURS1.Value = "Y"
        '課程大綱-師資-師資授課時數-超過54小時 為 true
        Dim flag_TEACHHOURS1 As Boolean = TIMS.CHK_TEACHHOURS1(dt)
        Hid_ChkTEACHHOURS1.Value = If(flag_TEACHHOURS1, "Y", "")
    End Sub
    ''' <summary>'組合每日課節資訊</summary>
    Function Get_TDescH3() As String
        Dim rst As String = ""
        For Each eItem As DataGridItem In Datagrid3.Items
            Dim flag_canUse As Boolean = True
            Dim STrainDateLabel As Label = eItem.FindControl("STrainDateLabel")
            Dim PNameLabel As Label = eItem.FindControl("PNameLabel")
            If STrainDateLabel Is Nothing Then flag_canUse = False
            If PNameLabel Is Nothing Then flag_canUse = False
            If (flag_canUse) Then
                'rst 連續設定
                Dim tmpX As String = Get_TDescH3_Lesson(STrainDateLabel.Text, PNameLabel.Text, rst)
                If tmpX <> "" Then '組合每日課節資訊 (以這筆日期資訊為最新)
                    If (rst <> "") Then rst &= ";"
                    rst &= tmpX
                End If
            End If
        Next
        Return rst
    End Function
    ''' <summary>'同日1組，不可n組 (針對單1日期，重組節數資訊)，其它日期不關我事 ,回傳當日組合</summary>
    Function Get_TDescH3_Lesson(ByVal STrainDate As String, ByVal sPNameLabel As String, ByRef sAll As String) As String
        'sAll
        Dim rst As String = ""
        If STrainDate.Length = 0 Then Return rst '日期資訊有誤
        If sPNameLabel.Length = 0 Then Return rst '時間資訊有誤

        Dim tmpH1H2 As String = "" '舊節數暫存
        Dim iH1 As Integer = TIMS.CINT1(Mid(sPNameLabel, 1, 2)) '起始時間(時)
        Dim iH2 As Integer = TIMS.CINT1(Mid(sPNameLabel, 7, 2)) '結束時間(時)
        'Dim iM2 As Integer = Val(Mid(sPNameLabel, 10, 2)) '結束時間(分)
        iH2 = iH2 - 1 '00減1
        For i As Integer = iH1 To iH2
            tmpH1H2 &= $"{If(tmpH1H2 <> "", ",", "")}{i}"
        Next

        If sAll.IndexOf(STrainDate) = -1 Then
            '找不到有相同的日期
            rst = $"{STrainDate}:{tmpH1H2}"  '日期1:H1,H2,H3,H4;日期2:H1,H2,H3,H4
        Else
            '有相同的日期(只能有1個，不能2個)
            Dim tmpH1H2_new As String = "" '新的節數暫存
            Dim sAll2 As String = ""
            Dim aryAll As String() = sAll.Split(";") '日期1:H1,H2,H3,H4;日期2:H1,H2,H3,H4
            For i As Integer = 0 To aryAll.Length - 1
                If aryAll(i).ToString().Length = 0 Then Return rst '分段資訊有誤

                If (aryAll(i).IndexOf(STrainDate) <> -1) Then
                    '搜尋到相同日期
                    Dim aryLesson2 As String() = aryAll(i).Split(":")(1).Split(",")
                    For j As Integer = 1 To 20 '(預估最大節數 20) '自走計算重複節數 1-20
                        Dim flag_ADDJ As Boolean = False
                        If Array.IndexOf(aryLesson2, j) <> -1 Then flag_ADDJ = True '搜尋到相同節數
                        If flag_ADDJ Then '可加入(節數)
                            tmpH1H2_new &= $"{If(tmpH1H2_new <> "", ",", "")}{j}"
                        End If
                    Next
                Else
                    '重組 sAll '未搜尋到相同日期  
                    If (aryAll(i)) <> "" Then
                        If (sAll2 <> "") Then sAll2 &= ";"
                        sAll2 &= aryAll(i) '補充新日期
                    End If
                End If
            Next
            sAll = sAll2 '全部日期塞入-sAll

            Dim aryLesson1 As String() = tmpH1H2.Split(",") '舊節數暫存
            Dim aryLesson3 As String() = tmpH1H2_new.Split(",") '新的節數暫存
            tmpH1H2_new = "" '(重整)
            For j As Integer = 1 To 20 '(預估最大節數 20) '自走計算重複節數 1-20
                Dim flag_ADDJ As Boolean = False
                If Array.IndexOf(aryLesson1, j) <> -1 Then flag_ADDJ = True
                If Array.IndexOf(aryLesson3, j) <> -1 Then flag_ADDJ = True
                If flag_ADDJ Then
                    tmpH1H2_new &= $"{If(tmpH1H2_new <> "", ",", "")}{j}"
                End If
            Next
            rst = $"{STrainDate}:{tmpH1H2_new}"  '回傳新資訊
        End If
        Return rst '回傳當日組合
    End Function

    'PLAN_PERSONCOST–一人份材料明細
    Function CreatePersonCost() As DataTable
        Dim dt1 As DataTable = Nothing
        Dim dt As DataTable = Nothing
        Dim DGobj As DataGrid = DataGrid6
        Const cst_sSupFd As String = ",0 Total,0 subtotal" '補充欄位
        Dim sql As String = ""
        If Not TIMS.IS_DataTable(Session(hid_PersonCostTable_guid1.Value)) Then
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT PLAN_PERSONCOST.* " & cst_sSupFd & " FROM PLAN_PERSONCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "  'Copy機制
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT PLAN_PERSONCOST.* " & cst_sSupFd & " FROM PLAN_PERSONCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If (Not gflag_can_copy2) Then sql &= " AND 1<>1"
                    If g_flagNG Then sql = " SELECT PLAN_PERSONCOST.* " & cst_sSupFd & " FROM PLAN_PERSONCOST WHERE 1<>1 "  'Copy機制
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT PLAN_PERSONCOST.* " & cst_sSupFd & " FROM PLAN_PERSONCOST WHERE 1<>1 "  'Copy機制
                Else
                    '修改資料取得
                    sql = ""
                    sql &= " SELECT PLAN_PERSONCOST.* " & cst_sSupFd & " FROM PLAN_PERSONCOST" & vbCrLf
                    sql &= " WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT PLAN_PERSONCOST.* " & cst_sSupFd & " FROM PLAN_PERSONCOST WHERE 1<>1 "
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(hid_PersonCostTable_guid1.Value)  '有資料
        End If
        If dt Is Nothing Then Return dt
        dt.Columns(Cst_PersonCostpkName).AutoIncrement = True
        dt.Columns(Cst_PersonCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_PersonCostpkName).AutoIncrementStep = -1
        If gflag_ccopy Then TIMS.CopyDATATABLE(dt, dt1, Cst_PersonCostpkName)

        Session(hid_PersonCostTable_guid1.Value) = dt
        With DGobj
            .Style.Item("display") = If(dt.Rows.Count > 0, "", "none")
            .DataSource = dt
            .DataKeyField = Cst_PersonCostpkName
            .DataBind()
        End With
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    'PLAN_COMMONCOST–共同材料明細
    Function CreateCommonCost() As DataTable
        Dim dt1 As DataTable = Nothing
        Dim dt As DataTable = Nothing
        Dim DGobj As DataGrid = DataGrid7
        Const cst_sSupFd As String = ",0 subtotal, 0 eachCost" '補充欄位
        Dim sql As String = ""
        If Not TIMS.IS_DataTable(Session(hid_CommonCostTable_guid1.Value)) Then
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT PLAN_COMMONCOST.* " & cst_sSupFd & " FROM PLAN_COMMONCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT PLAN_COMMONCOST.* " & cst_sSupFd & " FROM PLAN_COMMONCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If (Not gflag_can_copy2) Then sql &= " AND 1<>1"
                    If g_flagNG Then sql = " SELECT PLAN_COMMONCOST.* " & cst_sSupFd & " FROM PLAN_COMMONCOST WHERE 1<>1 "  'Copy機制
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT PLAN_COMMONCOST.* " & cst_sSupFd & " FROM PLAN_COMMONCOST WHERE 1<>1 "  'Copy機制
                Else
                    '修改資料取得
                    sql = ""
                    sql &= " SELECT PLAN_COMMONCOST.* " & cst_sSupFd & " FROM PLAN_COMMONCOST" & vbCrLf
                    sql &= " WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT PLAN_COMMONCOST.* " & cst_sSupFd & " FROM PLAN_COMMONCOST WHERE 1<>1 "
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(hid_CommonCostTable_guid1.Value) '有資料
        End If
        If dt Is Nothing Then Return dt
        dt.Columns(Cst_CommonCostpkName).AutoIncrement = True
        dt.Columns(Cst_CommonCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_CommonCostpkName).AutoIncrementStep = -1
        If gflag_ccopy Then Call TIMS.CopyDATATABLE(dt, dt1, Cst_CommonCostpkName)

        Session(hid_CommonCostTable_guid1.Value) = dt
        With DGobj
            .Style.Item("display") = If(dt.Rows.Count > 0, "", "none")
            .DataSource = dt
            .DataKeyField = Cst_CommonCostpkName
            .DataBind()
        End With
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function
    '更新 (訓練計劃開班總表(產學訓)) PLAN_VERREPORT
    Sub UPDATE_PLAN_VERREPORT(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByRef conn As SqlConnection)
        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Return 'rst Exit Sub
        If Not TIMS.OpenDbConn(conn) Then Return

        Dim hParms As New Hashtable From {{"PlanID", TIMS.CINT1(PlanID)}, {"ComIDNO", ComIDNO}, {"SeqNo", TIMS.CINT1(SeqNo)}}
        Dim sql As String = ""
        sql = " SELECT * FROM PLAN_VERREPORT WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, conn, hParms)
        If dt.Rows.Count = 0 Then Exit Sub

        Dim U_PARMS As New Hashtable From {{"PlanID", Val(PlanID)}, {"ComIDNO", ComIDNO}, {"SeqNo", Val(SeqNo)},
            {"Content", Content.Text}, {"MODIFYACCT", sm.UserInfo.UserID}}
        Dim U_sql As String = ""
        U_sql &= " UPDATE PLAN_VERREPORT" & vbCrLf
        U_sql &= " SET Content=@Content,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        U_sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf
        DbAccess.ExecuteNonQuery(U_sql, conn, U_PARMS)
    End Sub

    ''' <summary> 更新 CLASS_CLASSINFO </summary>
    ''' <param name="PlanID"></param>
    ''' <param name="ComIDNO"></param>
    ''' <param name="SeqNo"></param>
    ''' <param name="oConn"></param>
    Sub UPDATE_CLASS_CLASSINFO(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByRef oConn As SqlConnection)
        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Return 'rst Exit Sub
        Dim s_COLUMN_P As String = "TaddressZip,TaddressZIP6W,TAddress"
        Dim drP As DataRow = Get_PP_INFO_DR(oConn, PlanID, ComIDNO, SeqNo, s_COLUMN_P)
        If drP Is Nothing Then Return

        Dim s_COLUMN_C As String = "OCID"
        Dim drC As DataRow = Get_CC_INFO_DR(oConn, PlanID, ComIDNO, SeqNo, s_COLUMN_C)
        If drC Is Nothing Then Return

        Dim s_OCID As String = Convert.ToString(drC("OCID"))
        PointName.Text = TIMS.ClearSQM(PointName.Text)
        PackageName.Text = TIMS.ClearSQM(PackageName.Text)
        ClassName.Text = Replace(ClassName.Text, "&", "＆")
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        ClassName.Text = Replace(ClassName.Text, PointName.Text, "") '學分班種類
        ClassName.Text = Replace(ClassName.Text, PackageName.Text, "") '企業包班種類
        Dim v_Radiobuttonlist1 As String = TIMS.GetListValue(Radiobuttonlist1)
        Dim v_CLASSCNAME As String = ""
        Select Case v_Radiobuttonlist1'Radiobuttonlist1.SelectedValue
            Case cst_學分班 ' "Y"
                v_CLASSCNAME = ClassName.Text & PointName.Text & PackageName.Text
            Case Else 'cst_非學分班
                v_CLASSCNAME = ClassName.Text & PackageName.Text
        End Select
        'ClassCount: 預設值為01
        ClassCount.Text = TIMS.ClearSQM(ClassCount.Text)
        Dim vClassCount As String = ClassCount.Text
        vClassCount = TIMS.FmtCyclType(If(vClassCount <> "", vClassCount, "01"))
        'CyclType
        CyclType.Text = TIMS.FmtCyclType(CyclType.Text)

        Dim U_sql As String = ""
        U_sql &= " UPDATE CLASS_CLASSINFO" & vbCrLf
        U_sql &= " SET CLASSCNAME=@CLASSCNAME ,CYCLTYPE=@CYCLTYPE" & vbCrLf
        U_sql &= " ,TNUM=@TNUM ,THOURS=@THOURS ,TMID=@TMID" & vbCrLf
        U_sql &= " ,STDATE=@STDATE ,FTDATE=@FTDATE" & vbCrLf
        U_sql &= " ,TaddressZip=@TaddressZip" & vbCrLf
        U_sql &= " ,TaddressZIP6W=@TaddressZIP6W" & vbCrLf
        U_sql &= " ,TAddress=@TAddress" & vbCrLf
        U_sql &= " ,ClassNum=@ClassNum" & vbCrLf
        U_sql &= " ,LastState='M'" & vbCrLf 'M: 修改(最後異動狀態)  
        U_sql &= " ,ModifyAcct=@ModifyAcct" & vbCrLf
        U_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf 'NOW
        U_sql &= " WHERE OCID=@OCID" & vbCrLf
        U_sql &= " AND PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf

        Dim U_PARMS As New Hashtable From {
            {"CLASSCNAME", v_CLASSCNAME},
            {"CYCLTYPE", If(CyclType.Text <> "", CyclType.Text, Convert.DBNull)},
            {"TNUM", If(TNum.Text <> "", Val(TNum.Text), Convert.DBNull)},
            {"THOURS", If(THours.Text <> "", Val(THours.Text), Convert.DBNull)},
            {"TMID", If(trainValue.Value <> "", Val(trainValue.Value), Convert.DBNull)},
            {"STDATE", TIMS.Cdate2(STDate.Text)},
            {"FTDATE", TIMS.Cdate2(FDDate.Text)},
            {"TaddressZip", If(Convert.ToString(drP("TaddressZip")) <> "", drP("TaddressZip"), Convert.DBNull)},
            {"TaddressZIP6W", If(Convert.ToString(drP("TaddressZIP6W")) <> "", drP("TaddressZIP6W"), Convert.DBNull)},
            {"TAddress", If(Convert.ToString(drP("TAddress")) <> "", drP("TAddress"), Convert.DBNull)},
            {"ClassNum", vClassCount},
            {"ModifyAcct", sm.UserInfo.UserID},
            {"OCID", s_OCID},
            {"PlanID", PlanID},
            {"ComIDNO", ComIDNO},
            {"SeqNo", SeqNo}
        }
        DbAccess.ExecuteNonQuery(U_sql, oConn, U_PARMS)
    End Sub

    ''' <summary> '檢核是否有 CLASSINFO 有true 沒有false </summary>
    ''' <returns></returns>
    Public Shared Function CHK_CLASSINFO(ByRef MyPage As Page, ByRef gflag_ccopy As Boolean, ByRef oConn As SqlConnection) As Boolean
        Dim rst As Boolean = False
        If (MyPage Is Nothing) Then Return rst
        If (Convert.ToString(MyPage.Request("PlanID")) = "" OrElse gflag_ccopy) Then Return rst
        '修改
        Dim PlanID_value As String = TIMS.ClearSQM(MyPage.Request("PlanID"))
        Dim ComIDNO_value As String = TIMS.ClearSQM(MyPage.Request("ComIDNO"))
        Dim SeqNO_value As String = TIMS.ClearSQM(MyPage.Request("SeqNO"))
        If (PlanID_value = "" OrElse ComIDNO_value = "" OrElse SeqNO_value = "") Then Return rst

        Dim s_PCS As String = String.Format("{0}x{1}x{2}", PlanID_value, ComIDNO_value, SeqNO_value)
        Dim drPC2 As DataRow = TIMS.GetPCSDate2(s_PCS, oConn)
        rst = (drPC2 IsNot Nothing) '有資料
        Return rst
    End Function

    ''' <summary> GET PLAN_VERREPORT ReportCount </summary>
    ''' <returns></returns>
    Function GET_PLAN_VERREPORT_CNT() As String
        Dim rst As String = ""
        'rst = "0"  '外部copy而來
        If gflag_ccopy Then Return If(rst = "", "0", rst)

        Dim ReqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        If ReqSeqNo = "" Then Return If(rst = "", "0", rst) '資料空白

        Dim hParms As New Hashtable From {{"PlanID", TIMS.CINT1(sm.UserInfo.PlanID)}, {"ComIDNO", ComidValue.Value}, {"SeqNo", TIMS.CINT1(ReqSeqNo)}}
        Dim sql As String = "SELECT COUNT(1) ReportCount FROM PLAN_VERREPORT WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, hParms)
        If dt.Rows.Count = 0 Then Return If(rst = "", "0", rst)

        rst = $"{dt.Rows(0)("ReportCount")}"
        Return If(rst = "", "0", rst)
    End Function

    Function Get_PP_INFO_DR(ByRef oConn As SqlConnection, ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByVal s_COLUMN_P As String) As DataRow
        Dim drP As DataRow = Nothing

        Dim parms_pp As New Hashtable From {{"PlanID", PlanID}, {"ComIDNO", ComIDNO}, {"SeqNo", SeqNo}}
        Dim sql_pp As String = ""
        sql_pp = String.Concat(" SELECT ", If(s_COLUMN_P <> "", s_COLUMN_P, "*"), " FROM PLAN_PLANINFO")
        sql_pp &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf
        Dim dtP As DataTable = DbAccess.GetDataTable(sql_pp, oConn, parms_pp)
        If dtP.Rows.Count = 0 Then Return drP 'Exit Sub
        drP = dtP.Rows(0)
        Return drP
    End Function

    Function Get_CC_INFO_DR(ByRef oConn As SqlConnection, ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNo As String, ByVal s_COLUMN_C As String) As DataRow
        Dim drC As DataRow = Nothing

        Dim parms_c As New Hashtable From {{"PlanID", PlanID}, {"ComIDNO", ComIDNO}, {"SeqNo", SeqNo}}
        Dim sql As String = ""
        sql = String.Concat(" SELECT ", If(s_COLUMN_C <> "", s_COLUMN_C, "*"), " FROM CLASS_CLASSINFO")
        sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf
        Dim dtC As DataTable = DbAccess.GetDataTable(sql, oConn, parms_c)
        If dtC.Rows.Count = 0 Then Return drC 'Exit Sub
        drC = dtC.Rows(0)
        Return drC
    End Function

    ''' <summary> 事前整理</summary>
    Sub UTL_PREARRANGEMENT()
        STDate.Text = TIMS.ClearSQM(STDate.Text)
        FDDate.Text = TIMS.ClearSQM(FDDate.Text)
        STDate.Text = TIMS.Cdate3(STDate.Text)
        FDDate.Text = TIMS.Cdate3(FDDate.Text)
    End Sub

    ''' <summary>'草稿儲存--檢核 </summary>
    ''' <param name="Errmsg"></param>
    Sub CheckSaveDef(ByRef Errmsg As String)
        'Dim Errmsg As String = ""
        TNum.Text = TIMS.ClearSQM(TNum.Text)
        If TNum.Text = "" Then Errmsg &= "請輸入訓練人數" & vbCrLf
        If Errmsg <> "" Then Return

        Dim dtTemp As DataTable
        'Errmsg = ""
        If TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then 'PLAN_TRAINDESC
            dtTemp = Session(hid_TrainDescTable_guid1.Value)
            For Each drv As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                If Not drv.RowState = DataRowState.Deleted Then
                    If Convert.ToString(drv("STRAINDATE")) = "" Then
                        Errmsg &= "課程大網日期不可為空(必填)" & vbCrLf
                    End If
                    If Convert.ToString(drv("TECHID")) = "" Then
                        Errmsg &= "課程大網的任課教師(必填)" & vbCrLf
                    End If
                End If
                If Errmsg <> "" Then Exit For
            Next
        End If
        If Errmsg <> "" Then Return

        Dim i_Times_c_max_length As Integer = cst_i_Times_c_max_length
        Dim i_Times_c_min_length As Integer = cst_i_Times_c_min_length
        Dim s_err_msg1 As String = String.Format("上課時間／時間內容，長度超過限制範圍{0}文字長度", i_Times_c_max_length)
        Dim s_err_msg2 As String = String.Format("上課時間／時間內容，長度小於限制範圍{0}文字長度", i_Times_c_min_length)
        If TIMS.IS_DataTable(Session(hid_planONCLASS_guid1.Value)) Then
            dtTemp = Session(hid_planONCLASS_guid1.Value)
            For Each drv As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                If Not drv.RowState = DataRowState.Deleted Then
                    'drv("Times") = ClearTimesFMT(drv("Times"))
                    If Convert.ToString(drv("Times")) <> "" Then
                        If Convert.ToString(drv("Times")).ToString.Length > i_Times_c_max_length Then Errmsg &= s_err_msg1 & vbCrLf
                        If Convert.ToString(drv("Times")).ToString.Length < i_Times_c_min_length Then Errmsg &= s_err_msg2 & vbCrLf
                    Else
                        Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
                    End If
                End If
            Next
        End If
        '其他說明(欄位字數為1000)，超過欄位字數
        If tNote2.Text <> "" AndAlso tNote2.Text.Length > 1000 Then Errmsg &= "其他說明(欄位字數為1000)，超過欄位字數" & vbCrLf

        '檢核是否有業務權限
        Dim flag_CAN_NOCHECK_RIDPLAN As Boolean = False 'true:(可以不檢查業務權限)
        Dim flag_RIDPLAN As Boolean = Chk_RIDPLAN(RIDValue.Value, sm.UserInfo.PlanID)
        If flag_IsSuperUser_1 Then flag_CAN_NOCHECK_RIDPLAN = True
        'If flag_TIMS_Test_1 Then flag_CAN_NOCHECK_RIDPLAN = True
        '檢核是否有業務權限 ／ 登入者無正確的業務權限，不提供儲存服務!!" & vbCrLf
        If Not flag_CAN_NOCHECK_RIDPLAN AndAlso Not flag_RIDPLAN Then
            Errmsg &= $"{cst_errmsg22},{RIDValue.Value},{sm.UserInfo.PlanID},{sm.UserInfo.Years}{vbCrLf}" '"登入者無正確的業務權限，不提供儲存服務!!" & vbCrLf
        End If
        iCAPMARKDATE.Text = TIMS.ClearSQM(iCAPMARKDATE.Text)
        If iCAPMARKDATE.Text <> "" AndAlso Not TIMS.IsDate1(iCAPMARKDATE.Text) Then
            Errmsg &= "班別資料-【iCAP標章有效期限】有填寫, 日期格式有誤!" & vbCrLf
        ElseIf iCAPMARKDATE.Text <> "" Then
            iCAPMARKDATE.Text = TIMS.Cdate3(iCAPMARKDATE.Text)
        End If

        If Errmsg <> "" Then Return
    End Sub

    ''' <summary>'草稿儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        Hid_sender1.Value = cst_SaveDef 'sender.text

        Call UTL_PREARRANGEMENT()

        Dim Errmsg As String = ""
        Call CheckSaveDef(Errmsg)
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "5"
            Dim s_js11 As String = String.Concat("<script>document.getElementById('Button8').style.display="""";Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            'Page.RegisterStartupScript("Londing", "<script>document.getElementById('Button8').style.display="""";Layer_change(5);</script>") 'window.scroll(0,document.body.scrollHeight);
            sm.LastErrorMessage = Errmsg 'cst_errmsg15
            Return 'Exit Sub
        End If

        '假設處理某段程序需花費n毫秒 (避免機器不同步)
        If Session("GUID1") <> ViewState("GUID1") Then Threading.Thread.Sleep(1)
        ViewState("GUID1") = TIMS.GetGUID() : Session("GUID1") = ViewState("GUID1")

        '儲存 開班計畫/開班計畫表資料維護 '(草稿儲存)
        Call INSERT_PLAN_TABLE(cst_SaveDef)

        If (LayerState.Value = "") Then LayerState.Value = "1"
        Dim s_js1 As String = String.Concat("<script>document.getElementById('Button8').style.display="""";Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("_onload", s_js1)
    End Sub

    '檢核是否有業務權限
    Function Chk_RIDPLAN(ByVal RID As String, ByVal PlanID As String) As Boolean
        Dim rst As Boolean = False
        If RID = "" OrElse PlanID = "" Then Return rst
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = "SELECT * FROM AUTH_RELSHIP WHERE RID=@RID AND PLANID=@PLANID"
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dtAR As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RID
            .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(PlanID)
            dtAR.Load(.ExecuteReader())
        End With
        If TIMS.dtHaveDATA(dtAR) AndAlso dtAR.Rows.Count = 1 Then rst = True '有業務權限!!
        Return rst
    End Function

    ''' <summary> 檢核-iCAP標章證號 true:(可使用)OK false:(不可使用)已被其它單位使用</summary>
    ''' <param name="s_iCAPNUM"></param>
    ''' <param name="s_ComIDNO"></param>
    ''' <returns></returns>
    Function ChkiCAPNUM(ByVal s_iCAPNUM As String, ByVal s_ComIDNO As String) As Boolean
        Dim rst As Boolean = True
        s_iCAPNUM = TIMS.ClearSQM(s_iCAPNUM)
        If s_iCAPNUM = "" Then Return rst
        'SELECT iCAPNUM,count(1) c FROM PLAN_PLANINFO WHERE iCAPNUM is not null group by iCAPNUM order by 1 desc
        Dim parms As New Hashtable From {{"iCAPNUM", s_iCAPNUM}} 'parms.Clear()
        Dim sql As String = "SELECT TOP 1 COMIDNO FROM PLAN_PLANINFO WITH(NOLOCK) WHERE iCAPNUM=@iCAPNUM"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then Return rst

        Dim dr As DataRow = dt.Rows(0)
        If dr("COMIDNO") = s_ComIDNO Then Return rst
        If dr("COMIDNO") <> s_ComIDNO Then Return False
        If dr("COMIDNO") <> s_ComIDNO Then rst = False
        Return rst
    End Function

    Function GET_WNLDATA(oConn As SqlConnection) As DataTable
        TIMS.OpenDbConn(oConn)
        Dim sSql As String = ""
        sSql &= " SELECT a.TMID,a.BUSID,a.BUSNAME,a.JOBID,a.JOBNAME,a.TRAINID,a.TRAINNAME,a.GCID3" & vbCrLf
        sSql &= " ,g.GCODE31,g.GCODE2,a.WNL" & vbCrLf
        sSql &= " ,case when a.JobName is not null then concat('[',g.GCODE31,']',a.JobName) end NewJobName" & vbCrLf
        sSql &= " ,case when a.TrainName is not null then concat('[',g.GCODE2,']',a.TrainName) end NewTrainName" & vbCrLf
        sSql &= " FROM VIEW_TRAINTYPE a" & vbCrLf
        sSql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 g on g.GCID3=a.GCID3" & vbCrLf
        sSql &= " WHERE a.WNL=1" & vbCrLf
        Dim sCmd As New SqlCommand(sSql, oConn)
        Dim dt1 As New DataTable
        dt1.Load(sCmd.ExecuteReader())
        Return dt1
    End Function

    ''' <summary>檢核 -政策性產業 「職場新續航」項目，有規定特定訓練業別才能申請 </summary>
    ''' <param name="dtWNL"></param>
    ''' <param name="o_TMID"></param>
    ''' <returns></returns>
    Private Function CHK_WNLDATA(dtWNL As DataTable, o_TMID As Object) As Boolean
        '有資料為正常-True，異常：False
        If TIMS.dtNODATA(dtWNL) Then Return False
        Dim v_TMID As String = Convert.ToString(o_TMID)
        If v_TMID = "" Then Return False
        Dim fff As String = String.Concat("TMID='", v_TMID, "'")
        Dim rst As Boolean = dtWNL.Select(fff).Length > 0
        Return rst
    End Function

    ''' <summary> 基本儲存-資料檢核確認!! </summary>
    ''' <param name="ErrMsg"></param>
    Sub CheckAddData(ByRef ErrMsg As String)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        ErrMsg = ""
        Dim i_MaxTxtLen1 As Integer = 0

        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
            Select Case iBlackType
                Case 1, 2, 3
                    ErrMsg &= "於處分日期起的期間，班級申請資料建檔不可正式儲存。" & vbCrLf
            End Select
        End If
        If ErrMsg <> "" Then Exit Sub 'Return False '有錯誤訊息,不可儲存

        Hid_ComIDNO.Value = TIMS.ClearSQM(Hid_ComIDNO.Value)
        If ComidValue.Value = "" Then ComidValue.Value = Hid_ComIDNO.Value

        '基本儲存。
        Dim v_rblTMIDCORRECT As String = TIMS.GetListValue(rblTMIDCORRECT)
        'If flag_OJT22071401 Then End If
        If trTMIDCORRECT.Visible AndAlso v_rblTMIDCORRECT = "" Then
            ErrMsg &= String.Concat("「", cst_TMIDCORRECT_t, "」為必選欄位，請勾選同意或不同意") & vbCrLf
        End If

        'Dim flag_error As Boolean = False
        Dim v_Degree As String = TIMS.GetListValue(Degree)

        'If ErrMsg <> "" Then Exit Sub 'Return False '有錯誤訊息,不可儲存
        FirstSort.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(FirstSort.Text))
        If FirstSort.Text = "" Then
            ErrMsg &= "班別資料「優先排序」為必填欄位(為正整數數字且大於0)" & vbCrLf
        ElseIf (FirstSort.Text <> "" AndAlso Not TIMS.IsNumeric1(FirstSort.Text)) Then
            ErrMsg &= "班別資料「優先排序」請輸入數字(必須為正整數數字且大於0)!" & vbCrLf
        ElseIf (FirstSort.Text <> "" AndAlso Not TIMS.IsNumeric2(FirstSort.Text)) Then
            ErrMsg &= "班別資料「優先排序」請輸入數字(必須為正整數數字且大於0)!!" & vbCrLf
        End If

        iCAPNUM.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(iCAPNUM.Text))
        iCAPMARKDATE.Text = TIMS.Cdate3(TIMS.ClearSQM(iCAPMARKDATE.Text))
        '正式儲存的時候，系統檢核若是【iCap標章證號】有填寫，【iCAP標章有效期限】不可為空，跳出提示訊息
        Dim fg_CAN_CHECK_iCAPNUM_R1 As Boolean = (iCAPNUM.Text <> "" AndAlso iCAPMARKDATE.Text <> "" AndAlso FDDate.Text <> "" AndAlso TIMS.IsDate1(iCAPMARKDATE.Text) AndAlso TIMS.IsDate1(FDDate.Text))
        If (iCAPNUM.Text <> "") AndAlso iCAPMARKDATE.Text = "" Then
            ErrMsg &= String.Concat("班別資料-【iCap標章證號】有填寫，【iCAP標章有效期限】不可為空!") & vbCrLf
        ElseIf iCAPMARKDATE.Text <> "" AndAlso iCAPNUM.Text = "" Then
            ErrMsg &= String.Concat("班別資料-【iCAP標章有效期限】有填寫，【iCap標章證號】不可為空!") & vbCrLf
        ElseIf fg_CAN_CHECK_iCAPNUM_R1 AndAlso DateDiff(DateInterval.Day, CDate(iCAPMARKDATE.Text), CDate(FDDate.Text)) > 0 Then
            ErrMsg &= String.Concat("班別資料-iCAP標章有效期限須涵蓋完整訓練期間!") & vbCrLf
        End If
        '檢核-iCAP標章證號 true:OK false:已被其它單位使用
        If iCAPNUM.Text <> "" AndAlso Not ChkiCAPNUM(iCAPNUM.Text, ComidValue.Value) Then
            ErrMsg &= "「iCAP標章證號」已被其它單位使用!!(單位不同，不可使用相同的編碼), " & iCAPNUM.Text & vbCrLf
        End If

        ConNum.Text = TIMS.ClearSQM(ConNum.Text) '容納人數必須為數字
        If ConNum.Text <> "" AndAlso Not TIMS.IsNumeric2(ConNum.Text) Then
            ErrMsg &= "[班別資料]的[容納人數]有誤，必須為正整數數字且大於0!" & vbCrLf
        End If
        If ErrMsg <> "" Then Exit Sub 'Return False '有錯誤訊息,不可儲存

        'PHour
        'Dim iA_PHour As Integer = 0 '學科總時數(課程大綱)
        'Dim iT_PHour As Integer = 0 '術科總時數(課程大綱)
        'Dim iThours2 As Integer = 0 '訓練小時數
        'Dim iALL_PHour As Integer = 0 '總時數(課程大綱)
        'Dim iALL_PHour2 As Integer = 0 '總時數(課程大綱)

        'Dim rowi As Integer = 0
        Dim cdr4 As DataRow = Nothing '材料費
        Dim cost04 As Integer = 0 '材料費
        Dim gvid19 As String = TIMS.GetGlobalVar(Me, "19", "1", objconn)
        If gvid19 = "" Then
            ErrMsg &= "請至首頁>>系統管理>>系統參數管理>>參數設定-裡設定訓練人數(上限)"
            Exit Sub
        End If

        '申請階段／'申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇) 
        Dim v_rbl_AppStage As String = TIMS.GetListValue(rbl_AppStage)
        STDate.Text = TIMS.Cdate3(TIMS.ClearSQM(STDate.Text))
        FDDate.Text = TIMS.Cdate3(TIMS.ClearSQM(FDDate.Text))
        If STDate.Text = "" Then ErrMsg &= "請輸入訓練起日" & vbCrLf
        If FDDate.Text = "" Then ErrMsg &= "請輸入訓練迄日" & vbCrLf
        TNum.Text = TIMS.ClearSQM(TNum.Text)
        If TNum.Text = "" Then ErrMsg &= "請輸入訓練人數" & vbCrLf

        'https://jira.turbotech.com.tw/browse/TIMSC-235
        '修改說明：班級申請的正式儲存時，請依據使用者登入的年度，
        '判斷該班級的訓練起日的年度，相同年度，才可儲存，若不是相同年度，則不可儲存。
        Dim smYearMin As String = Convert.ToString(sm.UserInfo.Years)
        Dim smYearMax As String = Convert.ToString(sm.UserInfo.Years + 1)
        If STDate.Text <> "" Then
            Dim STDateYearS As String = CDate(STDate.Text).ToString("yyyy")
            '增修需求 OJT-21012701：產投 - 班級申請：【政府政策性產業】課程訓練起迄檢核邏輯調整
            ' 申請階段2 '表示 (1：上半年、2：下半年、3：政策性產業)
            If v_rbl_AppStage = "3" Then
                If (STDateYearS < smYearMin) OrElse (STDateYearS > smYearMax) Then
                    ErrMsg &= "「政策性產業」課程訓練起日的年度 需為使用者登入的年度或次年度，不可儲存!" & vbCrLf
                End If
            Else
                If Not smYearMin.Equals(STDateYearS) Then ErrMsg &= "訓練起日的年度 與 使用者登入的年度，不是相同年度，不可儲存!" & vbCrLf
            End If
        End If
        '增修需求 OJT-21012701：產投 - 班級申請：【政府政策性產業】課程訓練起迄檢核邏輯調整
        ' 申請階段2 '表示 (1：上半年、2：下半年、3：政策性產業)
        If v_rbl_AppStage = "3" AndAlso FDDate.Text <> "" Then
            Dim ffdate1 As Date = CDate(String.Format("{0}/4/30", smYearMax))
            Dim ffdate2 As Date = CDate(FDDate.Text)
            If DateDiff(DateInterval.Day, ffdate1, ffdate2) > 0 Then
                ErrMsg &= "「政策性產業」課程訓練迄日不可超過使用者登入的年度之次年度4月30日，不可儲存!" & vbCrLf
            End If
        End If

        '訓練人數上限
        TNum.Text = TIMS.ClearSQM(TNum.Text)
        If TNum.Text = "" Then
            ErrMsg &= "訓練人數-必須輸入為數字!" & vbCrLf
        ElseIf TNum.Text <> "" AndAlso Not TIMS.IsNumberStr(TNum.Text) Then
            ErrMsg &= String.Format("訓練人數-應為正整數數字格式(須大於0)!{0}", TNum.Text) & vbCrLf
        ElseIf TNum.Text <> "" AndAlso Not TIMS.IsNumeric2(TNum.Text) Then
            ErrMsg &= String.Format("訓練人數-應為正整數數字格式(須大於0)!!{0}", TNum.Text) & vbCrLf
        ElseIf TNum.Text <> "" AndAlso TIMS.IsNumberStr(TNum.Text) AndAlso gvid19 <> "" AndAlso TIMS.IsNumberStr(gvid19) _
            AndAlso TIMS.VAL1(TNum.Text) > TIMS.VAL1(gvid19) Then
            ErrMsg &= String.Format("訓練人數-系統限制為{0}人, 輸入{1}人超過系統限制", Val(gvid19), TNum.Text) & vbCrLf
        End If

        '申請階段 'v_rbl_AppStage/v_APPSTAGE  '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
        If tr_AppStage_TP28.Visible Then
            '有啟動才檢核
            'Dim VAL_rbl_AppStage As String = ""
            'If rbl_AppStage.SelectedIndex <> -1 Then VAL_rbl_AppStage = rbl_AppStage.SelectedValue
            'VAL_rbl_AppStage = TIMS.ClearSQM(VAL_rbl_AppStage)
            'If VAL_rbl_AppStage = "" Then ErrMsg &= "請選擇申請階段,申請階段為必填" & vbCrLf
            '申請階段／'申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇) 
            'Dim v_rbl_AppStage As String = TIMS.GetListValue(rbl_AppStage)
            If rbl_AppStage.SelectedIndex = -1 Then
                ErrMsg &= "請選擇申請階段,申請階段為必填" & vbCrLf
            Else
                If v_rbl_AppStage = "" Then ErrMsg &= "請選擇申請階段,申請階段為必填" & vbCrLf
                If v_rbl_AppStage = "0" Then ErrMsg &= "請選擇申請階段,申請階段為必填大於0" & vbCrLf
            End If
            'If AppStage.SelectedIndex = 0 OrElse AppStage.SelectedValue = "" Then ErrMsg &= "請選擇申請階段,申請階段為必填" & vbCrLf
        End If

        'If v_Degree = "" Then ErrMsg &= "受訓資格「學歷」為必選,請選擇" & vbCrLf
        ''不管什麼都是「年滿15歲以上」。   'Const cst_ageoDef As Integer = 16 'other Years Start
        'txtAge1.Text = TIMS.ClearSQM(txtAge1.Text)
        'If Not rdoAge1.Checked AndAlso Not rdoAge2.Checked Then ErrMsg &= "請選擇 受訓資格 年齡選項為必選" & vbCrLf
        'If ErrMsg <> "" Then Exit Sub 'Return False '有錯誤訊息,不可儲存
        'If rdoAge2.Checked Then
        '    If txtAge1.Text = "" Then
        '        ErrMsg &= "受訓資格 年齡選項2 未輸入有效年齡" & vbCrLf
        '    Else
        '        If Not TIMS.IsNumeric2(txtAge1.Text) Then ErrMsg &= "請檢查 受訓資格 年齡選項2 未輸入有效年齡:" & txtAge1.Text & vbCrLf
        '        If ErrMsg = "" Then
        '            If Val(txtAge1.Text) < cst_AgeOtherDef Then ErrMsg &= "請檢查 受訓資格 年齡選項2 有效年齡(須大於15歲(不含)以上)" & vbCrLf
        '            If Val(txtAge1.Text) > 99 Then ErrMsg &= "請檢查 受訓資格 年齡選項2 有效年齡(須小於99歲(含)以下)" & vbCrLf
        '        End If
        '    End If
        'End If

        'Dim v_Taddress2 As String = TIMS.GetListValue(Taddress2)
        'Dim v_Taddress3 As String = TIMS.GetListValue(Taddress3)
        Dim v_SciPlaceID As String = TIMS.GetListValue(SciPlaceID)
        Dim v_SciPlaceID2 As String = TIMS.GetListValue(SciPlaceID2)
        Dim v_TechPlaceID As String = TIMS.GetListValue(TechPlaceID)
        Dim v_TechPlaceID2 As String = TIMS.GetListValue(TechPlaceID2)

        Dim flag_OK_SciPlace_R As Boolean = If(v_SciPlaceID <> "", TIMS.Check_SciPlaceID(ComidValue.Value, v_SciPlaceID, objconn), True)
        Dim flag_OK_SciPlace_R2 As Boolean = If(v_SciPlaceID2 <> "", TIMS.Check_SciPlaceID(ComidValue.Value, v_SciPlaceID2, objconn), True)
        Dim flag_OK_TechPlace_R As Boolean = If(v_TechPlaceID <> "", TIMS.Check_TechPlaceID(ComidValue.Value, v_TechPlaceID, objconn), True)
        Dim flag_OK_TechPlace_R2 As Boolean = If(v_TechPlaceID2 <> "", TIMS.Check_TechPlaceID(ComidValue.Value, v_TechPlaceID2, objconn), True)

        If v_SciPlaceID <> "" AndAlso Not flag_OK_SciPlace_R Then ErrMsg &= "學科場地1已被刪除，請重新選擇" & vbCrLf
        If v_SciPlaceID2 <> "" AndAlso Not flag_OK_SciPlace_R2 Then ErrMsg &= "學科場地2已被刪除，請重新選擇" & vbCrLf
        If v_TechPlaceID <> "" AndAlso Not flag_OK_TechPlace_R Then ErrMsg &= "術科場地1已被刪除，請重新選擇" & vbCrLf
        If v_TechPlaceID2 <> "" AndAlso Not flag_OK_TechPlace_R2 Then ErrMsg &= "術科場地2已被刪除，請重新選擇" & vbCrLf

        Dim flag_NG_Place_R1 As Boolean = If((SciPlaceID.SelectedIndex = 0 AndAlso SciPlaceID2.SelectedIndex = 0 AndAlso TechPlaceID.SelectedIndex = 0 AndAlso TechPlaceID2.SelectedIndex = 0), True, False)
        If flag_NG_Place_R1 Then ErrMsg &= "【學科場地1】、【學科場地2】、【術科場地1】、【術科場地2】至少要設定其中一項" & vbCrLf

        If v_SciPlaceID = "" AndAlso v_SciPlaceID2 <> "" Then ErrMsg &= "請先設定【學科場地1】再設定【學科場地2】" & vbCrLf
        If v_TechPlaceID = "" AndAlso v_TechPlaceID2 <> "" Then ErrMsg &= "請先設定【術科場地1】再設定【術科場地2】" & vbCrLf
        If v_SciPlaceID = "" AndAlso v_TechPlaceID = "" Then ErrMsg &= "【學科場地1上課地址】與【術科場地1上課地址】至少要設定其中一項" & vbCrLf

        '辦理方式:null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        Dim vrbl_DISTANCE As String = TIMS.GetListValue(rbl_DISTANCE)
        '遠距課程環境1/2
        Dim v_REMOTEID1 As String = TIMS.GetListValue(ddl_REMOTEID1)
        Dim v_REMOTEID2 As String = TIMS.GetListValue(ddl_REMOTEID2)
        '遠距課程環境1/2
        If vrbl_DISTANCE = "2" AndAlso v_REMOTEID1 = "" Then
            '【辦理方式】選擇混成課程，點選基本儲存、正式儲存時，【遠距課程環境】欄位必須填寫
            ErrMsg &= cst_errmsg37 & vbCrLf
        ElseIf v_REMOTEID2 <> "" AndAlso v_REMOTEID1 = "" Then
            ErrMsg &= "【遠距課程環境2】有設定時，【遠距課程環境1】必須設定" & vbCrLf
        ElseIf v_REMOTEID2 <> "" AndAlso v_REMOTEID1 <> "" AndAlso v_REMOTEID1 = v_REMOTEID2 Then
            ErrMsg &= "【遠距課程環境】1.2 有設定時，【遠距課程環境】選項不可相同" & vbCrLf
        End If

        '課程內容有室外教學
        Dim v_rbl_OUTDOOR As String = TIMS.GetListValue(rbl_OUTDOOR)
        If v_rbl_OUTDOOR = "" Then ErrMsg &= "請選擇 班別資料-課程內容有室外教學 為必須選擇" & vbCrLf

        '【辦公室電話】、【行動電話】至少須擇一填寫
        '2023/2024/ContactPhone
        ContactPhone.Text = TIMS.ClearSQM(ContactPhone.Text)
        ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
        ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
        ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
        ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
        ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
        Dim s_ContactPhone As String = If(fg_phone_2024, TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text), ContactPhone.Text)
        If fg_phone_2024 Then
            Dim s_ContactMobile As String = TIMS.ChangContactMobile(ContactMobile_1.Text, ContactMobile_2.Text)
            If s_ContactPhone = "" AndAlso s_ContactMobile = "" Then ErrMsg &= "請輸入 班別資料-【辦公室電話】、【行動電話】至少須擇一填寫" & vbCrLf

            If s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_1.Text) Then
                ErrMsg &= "班別資料-【辦公室電話】區碼，僅能為數字" & vbCrLf
            ElseIf s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso Not ContactPhone_1.Text.StartsWith("0") Then
                ErrMsg &= "班別資料-【辦公室電話】區碼，第1碼應該為0" & vbCrLf
            ElseIf s_ContactPhone <> "" AndAlso ContactPhone_1.Text <> "" AndAlso ContactPhone_1.Text.Length < 2 Then
                ErrMsg &= "班別資料-【辦公室電話】區碼，長度須大於1" & vbCrLf
            ElseIf s_ContactPhone <> "" AndAlso ContactPhone_2.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_2.Text) Then
                ErrMsg &= "班別資料-【辦公室電話】電話(8碼)，僅能為數字" & vbCrLf
            ElseIf s_ContactPhone <> "" AndAlso ContactPhone_3.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactPhone_3.Text) Then
                ErrMsg &= "班別資料-【辦公室電話】分機，僅能為數字" & vbCrLf
            ElseIf s_ContactPhone <> "" AndAlso (ContactPhone_1.Text = "" OrElse ContactPhone_2.Text = "") Then
                ErrMsg &= "班別資料-【辦公室電話】不為空，請填寫完整(區碼與電話)為必填" & vbCrLf
            End If

            If s_ContactMobile <> "" AndAlso ContactMobile_1.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactMobile_1.Text) Then
                ErrMsg &= "班別資料-【行動電話】手機前4碼，僅能為數字" & vbCrLf
            ElseIf s_ContactMobile <> "" AndAlso ContactMobile_1.Text <> "" AndAlso Not ContactMobile_1.Text.StartsWith("0") Then
                ErrMsg &= "班別資料-【行動電話】手機前4碼，第1碼應該為0" & vbCrLf
            ElseIf s_ContactMobile <> "" AndAlso ContactMobile_2.Text <> "" AndAlso Not TIMS.IsNumberStr(ContactMobile_2.Text) Then
                ErrMsg &= "班別資料-【行動電話】手機後6碼，僅能為數字" & vbCrLf
            ElseIf s_ContactMobile <> "" AndAlso (ContactMobile_1.Text = "" OrElse ContactMobile_2.Text = "") Then
                ErrMsg &= "班別資料-【行動電話】不為空，請填寫完整(前4碼與後6碼)為必填" & vbCrLf
            End If
        Else
            '2023 old
            If s_ContactPhone = "" Then ErrMsg &= "請輸入 班別資料-【電話】" & vbCrLf
        End If

        jobValue.Value = TIMS.ClearSQM(jobValue.Value)
        trainValue.Value = TIMS.ClearSQM(trainValue.Value)
        Select Case strYears
            Case cst_strYears_2014 '"2014"
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    'dr("TMID") = If( jobValue.Value <> "", jobValue.Value, Convert.DBNull)
                    If jobValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
                Else
                    'dr("TMID") = If( trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                    If trainValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
                End If
            Case cst_strYears_2015 '"2015"
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    'dr("TMID") = If( jobValue.Value <> "", jobValue.Value, Convert.DBNull)
                    If jobValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
                Else
                    'dr("TMID") = If( trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                    If trainValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
                End If
            Case cst_strYears_2018 '"2018"
                'dr("TMID") = If(trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                If trainValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
            Case Else '(最新)
                If trainValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
        End Select
        'If jobValue.Value = "" Then ErrMsg &= "請選擇訓練業別，訓練業別為必須選擇" & vbCrLf
        'cjobValue
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)
        If cjobValue.Value = "" Then ErrMsg &= "請選擇通俗職類，通俗職類為必須選擇" & vbCrLf

        Dim iThours2 As Integer = 0 '訓練小時數
        THours.Text = TIMS.ClearSQM(THours.Text)
        If THours.Text <> "" Then
            Dim flagOk As Boolean = True
            Try
                THours.Text = CInt(Val(THours.Text))
            Catch ex As Exception
                ErrMsg &= "訓練時數必須為數字" & vbCrLf
                flagOk = False
            End Try
            If flagOk Then
                iThours2 = Val(THours.Text) '訓練小時數
                If Not IsNumeric(THours.Text) Then ErrMsg &= "訓練時數必須為數字" & vbCrLf
                If CInt(THours.Text) < 16 Then ErrMsg &= "訓練時數必須大於等於16" & vbCrLf
            End If
        Else
            iThours2 = 0 '訓練小時數
            ErrMsg &= "訓練時數必須填寫" & vbCrLf
        End If
        If ErrMsg <> "" Then Exit Sub

        If FIXSUMCOST.Text = "" Then ErrMsg &= " 必須填寫 固定費用總額" & vbCrLf
        If ACTHUMCOST.Text = "" Then ErrMsg &= " 必須計算 固定費用總額單一人時成本" & vbCrLf
        If METSUMCOST.Text = "" Then ErrMsg &= " 必須填寫 材料費用總額" & vbCrLf
        'If METCOSTPER.Text = "" Then ErrMsg &= " 必須計算 材料費用總額" & vbCrLf'材料費占比
        If METCOSTPER.Text = "" Then ErrMsg &= " 必須計算 材料費占比" & vbCrLf '材料費占比
        If ErrMsg <> "" Then Exit Sub

        Dim iPERC100 As Double = Val(Hid_PERC100.Value)
        Dim iACTHUMCOST As Double = CDbl(FIXSUMCOST.Text) / CDbl(TNum.Text) / CDbl(THours.Text)
        ACTHUMCOST.Text = TIMS.ROUND(iACTHUMCOST, 2)

        '材料費占比
        Dim iMETCOSTPER As Double = CDbl(METSUMCOST.Text) / CDbl(FIXSUMCOST.Text) * 100
        METCOSTPER.Text = TIMS.ROUND(iMETCOSTPER, 2)

        If iACTHUMCOST > Get_iMAX_ACTHUMCOST() AndAlso FIXExceeDesc.Text = "" Then ErrMsg &= "超出人時成本 必須填寫 超出人時成本原因說明" & vbCrLf
        If iMETCOSTPER > iPERC100 AndAlso METExceeDesc.Text = "" Then ErrMsg &= "超出材料費編列比上限 必須填寫 超出材料費比率上限原因說明" & vbCrLf
        If iMETCOSTPER > iPERC100 AndAlso METExceeDesc.Text = "" Then
            Dim msgtip_PERC100 As String = String.Format("材料費編列比上限:{0}%", iPERC100)
            TIMS.Tooltip(METExceeDesc, msgtip_PERC100, True)
        End If
        '*系統增加以下檢核
        '當計算結果個位數非 = 0時，須跳出提示訊息「每人總訓練費用，個位數須為0，請修改固定費用總額或材料費用總額」，並不可儲存。
        '個位數非 = 0 - -> 【每人總訓練費用】/ 10 取 餘數 <> 0
        '每人總訓練費用 ： 新臺幣
        If tPerPersonCost.Text <> "" AndAlso (Val(tPerPersonCost.Text) Mod 10) <> 0 Then
            ErrMsg &= "每人總訓練費用，個位數及小數點之後須為0，請修改固定費用總額或材料費用總額!" & vbCrLf
        ElseIf tPerPersonCost.Text <> "" AndAlso (Val(tPerPersonCost.Text) * 100 Mod 10) <> 0 Then
            '每人總訓練費用：訓練費用總額/訓練人數，調整計算四捨五入至小數點第 2 位
            '當此欄位個位數及小數點第1、2位非0時，須跳出提示訊息「每人總訓練費用，個位數及小數點之後須為0，
            '請修改固定費用總額或材料費用總額」，並不能儲存
            ErrMsg &= "每人總訓練費用，個位數及小數點之後須為0，請修改固定費用總額或材料費用總額!!" & vbCrLf
        End If

        If DefGovCost.Text = "" OrElse DefGovCost.Text = "0" Then ErrMsg &= "政府補助金額必須大於 0" & vbCrLf
        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        Dim v_Radiobuttonlist1 As String = TIMS.GetListValue(Radiobuttonlist1)
        Dim v_PointType As String = TIMS.GetListValue(PointType) '.SelectedValue
        CredPoint.Text = TIMS.ClearSQM(CredPoint.Text)
        If CredPoint.Text <> "" Then
            If Not IsNumeric(CredPoint.Text) Then ErrMsg &= "學分數必須為數字" & vbCrLf
        Else
            'Dim v_Radiobuttonlist1 As String = TIMS.GetListValue(Radiobuttonlist1)
            Select Case v_Radiobuttonlist1'Radiobuttonlist1.SelectedValue
                Case cst_學分班 ' "Y"
                    ErrMsg &= "選擇學分班，學分數為必填數字" & vbCrLf
                Case Else 'cst_非學分班
            End Select
        End If

        '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
        If hTPlanID54.Value = "1" Then
            Select Case v_PackageType 'PackageType.SelectedValue
                Case "1" '非包班
                Case "2", "3" '2:企業包班,3:聯合企業包班
                Case Else
                    ErrMsg &= "選擇包班種類，包班種類為必填!!" & vbCrLf
            End Select
        End If
        If ErrMsg <> "" Then Exit Sub
        Dim I_TGovExam As Integer = 0
        If (TGovExamCY.Checked) Then I_TGovExam += 1
        If (TGovExamCN.Checked) Then I_TGovExam += 1
        If (TGovExamCG.Checked) Then I_TGovExam += 1
        If I_TGovExam > 1 Then ErrMsg &= "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定，請擇一勾選！" & vbCrLf
        If TGovExamCY.Checked AndAlso GOVAGENAME.Text = "" Then
            ErrMsg &= "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定，若為是。政府機關名稱不可為空" & vbCrLf
        End If
        If TGovExamCY.Checked AndAlso TGovExamName.Text = "" Then
            ErrMsg &= "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定，若為是。證照或檢定名稱不可為空" & vbCrLf
        End If

        If hTPlanID54.Value = "1" Then   '充電起飛計畫 (hTPlanID54.Value = "1")
            txtUname.Text = TIMS.ClearSQM(txtUname.Text)
            txtIntaxno.Text = TIMS.ClearSQM(txtIntaxno.Text)
            Select Case v_PackageType 'PackageType.SelectedValue
                Case "2" '充電起飛計畫 '企業包班
                    If Session(hid_PLAN_BUSPACKAGE_guid1.Value) IsNot Nothing Then Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
                    If txtUname.Text = "" Then
                        txtUname.Text = ""
                        ErrMsg &= "包班事業單位 企業名稱，不可為空" & vbCrLf
                    Else
                        If txtUname.Text.ToString.Length > 50 Then ErrMsg &= "包班事業單位 企業名稱，長度超過限制範圍50文字長度" & vbCrLf  '錯誤檢查
                    End If
                    If txtIntaxno.Text <> "" Then
                        If Not TIMS.CheckIsECFA(TIMS.ChangeIDNO(txtIntaxno.Text), objconn) Then ErrMsg &= "「" & Convert.ToString(txtUname.Text) & "」該包班事業單位 企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf  '未填寫 ECFA包班事業單位資料
                    Else
                        txtIntaxno.Text = ""
                        ErrMsg &= "包班事業單位 服務單位統一編號，不可為空" & vbCrLf
                    End If
                Case "3" '充電起飛計畫 '聯合企業包班
                    If TIMS.IS_DataTable(Session(hid_PLAN_BUSPACKAGE_guid1.Value)) Then
                        Dim j As Integer = 0
                        Dim dt As DataTable = Session(hid_PLAN_BUSPACKAGE_guid1.Value)
                        If dt.Rows.Count > 0 Then
                            For i As Integer = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then
                                    Dim dr1 As DataRow = dt.Rows(i)
                                    If Not TIMS.CheckIsECFA(Convert.ToString(dr1("Intaxno")), objconn) Then ErrMsg &= "「" & Convert.ToString(dr1("Uname")) & "」該包班事業單位 企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf  '未填寫 包班事業單位資料
                                    j += 1
                                End If
                            Next
                            If j = 0 Then ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf '未填寫 包班事業單位資料
                        Else
                            ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf '未填寫 包班事業單位資料
                        End If
                    Else
                        ErrMsg &= "充電起飛計畫(聯合企業包班)，包班事業單位資料 至少要填1筆!!" & vbCrLf '未填寫 包班事業單位資料
                    End If
                Case Else
                    If Session(hid_PLAN_BUSPACKAGE_guid1.Value) IsNot Nothing Then Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
            End Select
        End If

        '訓練職能 ClassCate 六大職能別查詢清單
        Dim v_ClassCate As String = TIMS.GetListValue(ClassCate)
        'Dim vsClassCate As String = TIMS.ClearSQM(ClassCate.SelectedValue)
        If v_ClassCate = "" Then ErrMsg &= "請選擇訓練職能!" & vbCrLf
        If ErrMsg <> "" Then Exit Sub

        Dim vsErrMsg2 As String = ""

        'PLAN_TRAINDESC
        Dim iA_PHour As Integer = 0 '學科總時數(課程大綱)
        Dim iT_PHour As Integer = 0 '術科總時數(課程大綱)
        Dim iALL_PHour As Double = 0 '總時數(課程大綱)
        Dim iALL_PHour2 As Integer = 0 '總時數(課程大綱)(驗證)
        Dim i_DESCROWS As Integer = 0 '資料筆數
        Dim i_FARLEARN As Integer = 0 '遠距教學-筆數
        Dim i_OUTLEARN As Integer = 0 '室外教學-筆數
        Dim iALL_AIAHOUR As Double = 0 'AI應用時數
        Dim iALL_WNLHOUR As Double = 0 '職場續航時數

        'Dim i_FARLEARN_PHours As Double = 0 '遠距教學時數
        Dim rowi As Integer = 0
        If Not TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then ErrMsg &= cst_errmsg24 & "!!" & vbCrLf '"課程大綱為必填資料" & vbCrLf
        If TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then
            Dim dt As DataTable = Session(hid_TrainDescTable_guid1.Value)
            If dt.Rows.Count > 0 Then
                Dim bolSDateFlag As Boolean = False
                Dim bolEDateFlag As Boolean = False
                rowi = 1
                For i As Int16 = 0 To dt.Rows.Count - 1
                    If Not dt.Rows(i).RowState = DataRowState.Deleted Then
                        iALL_AIAHOUR += If(Convert.ToString(dt.Rows(i)("AIAHOUR")) <> "", Val(dt.Rows(i)("AIAHOUR")), 0)
                        iALL_WNLHOUR += If(Convert.ToString(dt.Rows(i)("WNLHOUR")) <> "", Val(dt.Rows(i)("WNLHOUR")), 0)
                        iALL_PHour += Val(dt.Rows(i)("PHour"))
                        If dt.Rows(i)("Classification1") = "1" Then '學科總時數  '1:學/2:術
                            iA_PHour += Val(dt.Rows(i)("PHour"))
                        ElseIf dt.Rows(i)("Classification1") = "2" Then '術科總時數  '1:學/2:術
                            iT_PHour += Val(dt.Rows(i)("PHour"))
                        End If

                        Dim flag_STrainDate_NULL As Boolean = False
                        If Convert.ToString(dt.Rows(i)("STrainDate")) = "" Then flag_STrainDate_NULL = True
                        If flag_STrainDate_NULL Then
                            vsErrMsg2 &= String.Concat("第", rowi, "筆：課程大網日期不能為空") & vbCrLf
                        End If
                        If Not flag_STrainDate_NULL Then
                            Select Case True
                                Case DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(dt.Rows(i)("STrainDate"))) < 0
                                    vsErrMsg2 &= String.Concat("第", rowi, "筆：課程大網日期不能超過訓練起日", STDate.Text) & vbCrLf
                                Case DateDiff(DateInterval.Day, CDate(dt.Rows(i)("STrainDate")), CDate(FDDate.Text)) < 0
                                    vsErrMsg2 &= String.Concat("第", rowi, "筆：課程大網日期不能超過訓練迄日", FDDate.Text) & vbCrLf
                            End Select
                        End If
                        If IsDBNull(dt.Rows(i)("PTID")) Then vsErrMsg2 &= String.Concat("第", rowi, "筆：請選擇課程大網的上課地點(必填)") & vbCrLf
                        If IsDBNull(dt.Rows(i)("TechID")) Then vsErrMsg2 &= String.Concat("第", rowi, "筆：請選擇課程大網的任課教師(必填)") & vbCrLf
                        '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
                        '(產業人才投資方案) 增加設定課程大綱時,新增內容需有固定二筆資料日期為訓練起日及訓練迄日
                        If Not flag_STrainDate_NULL Then
                            If CDate(STDate.Text) = CDate(dt.Rows(i)("STrainDate")) Then bolSDateFlag = True
                            If CDate(FDDate.Text) = CDate(dt.Rows(i)("STrainDate")) Then bolEDateFlag = True
                        End If
                        rowi += 1
                        i_DESCROWS += 1 '資料筆數
                        '遠距教學-筆數
                        If Convert.ToString(dt.Rows(i)("FARLEARN")).Equals("Y") Then i_FARLEARN += 1
                        '室外教學-筆數
                        If Convert.ToString(dt.Rows(i)("OUTLEARN")).Equals("Y") Then i_OUTLEARN += 1
                    End If
                    If vsErrMsg2 <> "" Then Exit For
                Next
                If (Not bolSDateFlag OrElse Not bolEDateFlag) Then vsErrMsg2 &= "課程大綱內容資料，需固定二筆資料日期為訓練起日及訓練迄日" & vbCrLf

                '辦理方式'flag_StopDISTANCE2: 產投使用／遠距教學 暫不啟用 False:啟用中 ( i_DESCROWS > 0) 資料筆數大於0
                If Not flag_StopDISTANCE2 AndAlso i_DESCROWS > 0 Then
                    '辦理方式'啟用才檢核 null"無遠距教學", 1.申請整班為遠距教學, 2.申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
                    Dim s_DISTANCE_N As String = TIMS.GET_DISTANCE_N(0, vrbl_DISTANCE)
                    If vrbl_DISTANCE = "" Then
                        vsErrMsg2 &= "請選擇【辦理方式】為必填!" & vbCrLf
                    ElseIf vrbl_DISTANCE.Equals("2") AndAlso i_DESCROWS.Equals(i_FARLEARN) Then
                        vsErrMsg2 &= String.Format("選擇「{0}」時，課程大網不能將所有課程都勾選【辦理方式】", s_DISTANCE_N) & vbCrLf
                    ElseIf vrbl_DISTANCE.Equals("2") AndAlso i_FARLEARN = 0 Then
                        vsErrMsg2 &= String.Format("選擇「{0}」時，課程大網部分課程須勾選【辦理方式】", s_DISTANCE_N) & vbCrLf
                    ElseIf vrbl_DISTANCE.Equals("1") AndAlso Not i_DESCROWS.Equals(i_FARLEARN) Then
                        vsErrMsg2 &= String.Format("選擇「{0}」時，課程大綱所有課程須勾選【辦理方式】", s_DISTANCE_N) & vbCrLf
                    ElseIf vrbl_DISTANCE.Equals("3") AndAlso i_FARLEARN > 0 Then
                        vsErrMsg2 &= String.Format("選擇「{0}」時，課程大綱所有課程不能勾選【辦理方式】", s_DISTANCE_N) & vbCrLf
                    End If
                End If
                'TIMS.LOG.Debug(String.Concat("", ",#flag_StopDISTANCE2:", flag_StopDISTANCE2, ",#i_DESCROWS:", i_DESCROWS, ",#i_FARLEARN:", i_FARLEARN))

                If i_DESCROWS > 0 Then
                    '如第3點【課程內容有室外教學】勾「是」，課程大綱之「室外教學」必須至少有一個勾選。否則提示："勾選【課程內容有室外教學】為「是」時，請至少勾選一堂課為「室外教學」"
                    '如第3點【課程內容有室外教學】勾「否」，課程大綱之「室外教學」不可有勾選。否則提示："勾選【課程內容有室外教學】為「否」時，不可勾選課程為「室外教學」"
                    If v_rbl_OUTDOOR.Equals("Y") AndAlso i_OUTLEARN = 0 Then vsErrMsg2 &= "勾選【課程內容有室外教學】為「是」時，請至少勾選一堂課為「室外教學」" & vbCrLf
                    If v_rbl_OUTDOOR.Equals("N") AndAlso i_OUTLEARN > 0 Then vsErrMsg2 &= "勾選【課程內容有室外教學】為「否」時，不可勾選課程為「室外教學」" & vbCrLf
                End If
            End If

            If iALL_PHour > 0 AndAlso ErrMsg = "" Then
                rowi = 1
                Try
                    For Each eItem As DataGridItem In Datagrid3.Items
                        Dim PHourLabel As Label = eItem.FindControl("PHourLabel")
                        PHourLabel.Text = TIMS.ClearSQM(PHourLabel.Text)
                        '總時數(課程大綱)
                        If PHourLabel.Text <> "" Then iALL_PHour2 += Val(PHourLabel.Text)
                        rowi += 1
                    Next
                Catch ex As Exception
                    ErrMsg &= String.Format("第{0}筆：{1}", rowi, cst_errmsg25) & vbCrLf '"課程大綱內容資料，請重新確認!!" & vbCrLf
                    Exit Sub
                End Try
            End If
        End If

        'Dim flag_TIMS_Test_1 As Boolean = TIMS.sUtl_ChkTest() 'If Not flag_TIMS_Test_1 Then Exit Sub
        If vsErrMsg2 <> "" Then ErrMsg &= vsErrMsg2
        If ErrMsg <> "" Then Exit Sub

        '2019年啟用 work2019x01:2019 政府政策性產業
        Dim sErrMsg1 As String = CHK_KID20_VAL_OTH()
        If trKID20.Visible AndAlso sErrMsg1 <> "" Then
            ErrMsg &= sErrMsg1 ' Common.MessageBox(Me, sErrMsg1)
            Return  'Exit Sub
        End If
        Dim sErrMsg2 As String = TIMS.CHK_KID60_VAL(CBLKID60)
        If fg_USE_CBLKID60_TP28 AndAlso sErrMsg2 <> "" Then
            ErrMsg &= sErrMsg2 ' Common.MessageBox(Me, sErrMsg2)
            Return  'Exit Sub
        End If
        Dim sErrMsg3 As String = CHK_KID25_VAL_OTH() '2025 政府政策性產業
        If trKID25.Visible AndAlso sErrMsg3 <> "" Then
            ErrMsg &= sErrMsg3 ' Common.MessageBox(Me, sErrMsg1)
            Return  'Exit Sub
        End If
        Dim sErrMsg4 As String = CHK_KID26_VAL_OTH() '2026 政府政策性產業
        If trKID26.Visible AndAlso sErrMsg4 <> "" Then
            ErrMsg &= sErrMsg4 ' Common.MessageBox(Me, sErrMsg1)
            Return  'Exit Sub
        End If

        If i_DESCROWS = 0 Then
            ErrMsg &= $"{cst_errmsg24}!!!{vbCrLf}"   '"產學訓課程大綱，為必填資料" & vbCrLf '97產學訓課程大綱，為必填資料
        Else
            If iALL_PHour <= 0 Then ErrMsg &= "課程大綱的時數，為必填資料，請修正!!" & vbCrLf
            If ErrMsg = "" AndAlso iALL_PHour2 <= 0 Then ErrMsg &= "課程大綱的時數，為必填資料，請修正!!" & vbCrLf
            If Val(THours.Text) <= 0 Then ErrMsg &= "訓練時數，為必填資料，請修正!!" & vbCrLf
            If ErrMsg = "" Then
                '沒有錯誤整理
                THours.Text = CInt(Val(THours.Text)) '課程時數為整數
                If iALL_PHour <> Val(THours.Text) Then ErrMsg &= "產業人才投資方案課程大綱時數加總需等於訓練時數" & vbCrLf
                If ErrMsg = "" AndAlso iALL_PHour2 <> Val(THours.Text) Then ErrMsg &= "產業人才投資方案課程大綱時數加總需等於訓練時數" & vbCrLf
            End If
        End If

        '2024'確認單位所選「課程大綱日期」與「上課時間之星期」是否正確，於基本儲存、正式儲存時
        Dim dtDesc As DataTable = If(Session(hid_TrainDescTable_guid1.Value) Is Nothing, Nothing, CType(Session(hid_TrainDescTable_guid1.Value), DataTable))
        Dim dtOnClass As DataTable = If(Session(hid_planONCLASS_guid1.Value) Is Nothing, Nothing, CType(Session(hid_planONCLASS_guid1.Value), DataTable))
        Dim hMsg1 As New Hashtable
        If Not TIMS.CHK_WEEKDAY1(dtDesc, dtOnClass, hMsg1) Then
            Dim s_STRAIN As String = TIMS.GetMyValue2(hMsg1, "STRAINDATE")
            Dim s_WEEKS As String = TIMS.GetMyValue2(hMsg1, "WEEKS")
            ErrMsg &= $"課程大綱檢核：所選的「課程大綱日期」({s_STRAIN})({s_WEEKS})與「上課時段之星期」不符合!"
            Return
        End If

        '2022進階政策性產業類別/2025/B
        'Dim v_CBLKID22 As String = TIMS.GetCblValue(CBLKID22)
        'Dim v_CBLKID22B As String = TIMS.GetCblValue(CBLKID22B)
        'If v_CBLKID22B <> "" Then v_CBLKID22 = v_CBLKID22B
        '進階政策性產業類別/B
        Dim v_CBLKID22 As String = If(trKID25.Visible, TIMS.GetCblValue(CBLKID22B), TIMS.GetCblValue(CBLKID22)) 'v_CBLKID22
        'v_rbl_AppStage/v_APPSTAGE  '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
        If v_rbl_AppStage = "4" AndAlso v_CBLKID22 = "" Then
            ErrMsg &= "【申請階段】為「4:進階政策性產業」時，進階政策性產業類別，為必選!" & vbCrLf
        ElseIf v_rbl_AppStage <> "4" AndAlso v_CBLKID22 <> "" Then
            ErrMsg &= "【申請階段】非「4:進階政策性產業」時，進階政策性產業類別，不可勾選!" & vbCrLf
        End If
        '當【申請階段】點選「4:進階政策性產業」時，結訓日期 ( 訓練迄日 ) 必須卡控在當年度12/31。 (如今年，必須<=2022/12/31)
        '影響的功能包括： 班級申請、班級變更申請，請增加【訓練迄日】之系統檢核
        If v_rbl_AppStage = "4" AndAlso FDDate.Text <> "" AndAlso IsDate(FDDate.Text) Then
            If (CDate(FDDate.Text).Year <> Now.Year) Then
                ErrMsg &= "【申請階段】為「4:進階政策性產業」時，結訓日期(訓練迄日), 必須在當年度!!" & vbCrLf
            End If
        End If

        Dim v_CBLKID25 As String = GET_KID25_VAL() '2025 政府政策性產業
        Dim v_CBLKID26 As String = GET_KID26_VAL() '2026 政府政策性產業
        If v_rbl_AppStage = "3" AndAlso $"{v_CBLKID25}{v_CBLKID26}" = "" Then
            ErrMsg &= "申請階段為「政策性產業」必須勾選任一政府政策性產業!" & vbCrLf
        End If
        '三、基本儲存、正式儲存卡控：充電起飛計畫不用這些卡控
        '1、「AI加值應用、職場續航」僅限申請階段為「政策性產業」，當單位有勾選，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」，若未填寫須出現提示訊息，且不能儲存。
        '2、當勾選「AI加值應用」時，須卡控AI應用時數之總數須等於或大於12小時，且不可超過總訓練時數1/2，若不符須出現提示訊息，且不能儲存。
        '3、當勾選「職場續航」時，須卡控職場新續航時數之總數須等於6小時，不可小於或大於，若不符須出現提示訊息，且不能儲存。
        '4、因單位可能同時勾選「AI加值應用」及「職場新續航」，且同一堂課可能會有AI應用時數及職場新續航時數，各自時數仍須分開計算。
        '5、因政策性產業新增「職場新續航」項目，此類政策性產業有規定特定訓練業別才能申請（詳附件檔案），須卡控班級申請之訓練業別代碼倘非規定之代碼，則跳出提示訊息【申請之訓練業別非「職場新續航」課程職類】且不可儲存。
        Dim v_CBLKID25_7 As String = TIMS.GetCblValue(CBLKID25_7) 'AI加值應用
        Dim v_CBLKID25_8 As String = TIMS.GetCblValue(CBLKID25_8) '職場續航
        If v_rbl_AppStage <> "3" AndAlso v_CBLKID25_7 <> "" Then ErrMsg &= "「AI加值應用」僅限申請階段為「政策性產業」!" & vbCrLf
        If v_CBLKID25_7 <> "" AndAlso iALL_AIAHOUR < 12 Then
            ErrMsg &= String.Concat("勾選「AI加值應用」，AI應用時數之總數須等於或大於12小時!.", iALL_AIAHOUR, vbCrLf)
        ElseIf v_CBLKID25_7 <> "" AndAlso iALL_AIAHOUR > (iALL_PHour / 2) Then
            ErrMsg &= String.Concat("勾選「AI加值應用」，AI應用時數之總數，不可超過總訓練時數1/2!.", iALL_AIAHOUR, ">", (iALL_PHour / 2), vbCrLf)
        ElseIf v_CBLKID25_7 = "" AndAlso iALL_AIAHOUR > 0 Then
            ErrMsg &= String.Concat("有輸入AI應用時數，請勾選「AI加值應用」!.", iALL_AIAHOUR, vbCrLf)
        End If
        If v_rbl_AppStage <> "3" AndAlso v_CBLKID25_8 <> "" Then ErrMsg &= "「職場續航」僅限申請階段為「政策性產業」!" & vbCrLf
        If v_CBLKID25_8 <> "" AndAlso iALL_WNLHOUR <> 6 Then
            ErrMsg &= String.Concat("勾選「職場續航」，職場續航時數之總數須等於6小時!.", iALL_WNLHOUR, vbCrLf)
        ElseIf v_CBLKID25_8 = "" AndAlso iALL_WNLHOUR > 0 Then
            ErrMsg &= String.Concat("有輸入職場續航時數，請勾選「職場續航」!.", iALL_WNLHOUR, vbCrLf)
        End If
        If v_CBLKID25_8 <> "" Then
            Dim o_TMID As Object = GET_o_TMID_VAL() '取得-訓練業別
            Dim dtWNL As DataTable = GET_WNLDATA(objconn) 'TMID
            Dim fg_WNLDATA As Boolean = CHK_WNLDATA(dtWNL, o_TMID)
            If Not fg_WNLDATA Then ErrMsg &= "勾選「職場續航」，「訓練業別」非「職場續航」申請課程職類!" & vbCrLf
        End If
        Dim V_tPOLICYREL As String = TIMS.Get_placeholder_TXTVAL(tPOLICYREL)
        If v_CBLKID25_7 <> "" AndAlso V_tPOLICYREL = "" Then
            ErrMsg &= "勾選「AI加值應用」，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」!" & vbCrLf
        End If
        If v_CBLKID25_8 <> "" AndAlso V_tPOLICYREL = "" Then
            ErrMsg &= "勾選「職場續航」，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」!" & vbCrLf
        End If

        If STDate.Text <> "" AndAlso FDDate.Text <> "" Then
            If IsDate(STDate.Text) AndAlso IsDate(FDDate.Text) Then
                If DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text)) < 0 Then ErrMsg &= "訓練日期起迄，訓練迄日需大於訓練起日" & vbCrLf
                If DateDiff(DateInterval.Day, CDate(STDate.Text), CDate(FDDate.Text)) = 0 Then ErrMsg &= "訓練日期起迄，訓練起日不能和訓練迄日同一天" & vbCrLf
            Else
                ErrMsg &= "訓練日期起迄，格式有誤" & vbCrLf
            End If
        End If
        If Convert.ToString(GCIDValue.Value) = "156" Then ErrMsg &= "經費分類代碼有誤(其他 停用),請重新選擇!!" & vbCrLf
        If Convert.ToString(GCIDValue.Value) = "157" Then ErrMsg &= "經費分類代碼有誤(學分班依教育部規定辦理 停用),請重新選擇!!" & vbCrLf
        If Convert.ToString(GCIDValue.Value) = "158" Then ErrMsg &= "經費分類代碼有誤(3C共通核心職能課程 停用),請重新選擇!!" & vbCrLf

        '須要檢核申請階段 (false:不檢核/true:檢核)
        Dim flag_AppStage_1 As Boolean = (tr_AppStage_TP28.Visible AndAlso v_rbl_AppStage <> "") 'False '須要檢核申請階段 (false:不檢核)
        'If tr_AppStage.Visible AndAlso AppStage.SelectedValue <> "" Then flag_AppStage_1 = True '須要檢核申請階段
        'If tr_AppStage_TP28.Visible AndAlso rbl_AppStage.SelectedValue <> "" Then flag_AppStage_1 = True '須要檢核申請階段
        'If tr_AppStage_TP28.Visible AndAlso v_rbl_AppStage <> "" Then flag_AppStage_1 = True '須要檢核申請階段
        If ErrMsg = "" AndAlso flag_AppStage_1 AndAlso FirstSort.Text <> "" Then
            'https://jira.turbotech.com.tw/browse/TIMSC-138
            '修改說明：班別資料之「優先排序」欄位，如有重複植入之序號者，即無法儲存並跳出出提醒文字。
            '此外，107年度啟用，區隔上、下年度，以6/30做為區隔，開訓日為1/1~6/30為上半年度，7/1~12/31為下半年度，上、下半年的數字要區分，
            '同一個半年度內，不可重複填寫相同數字。
            If ComidValue.Value = "" Then ComidValue.Value = Hid_ComIDNO.Value
            Dim ss As String = ""
            TIMS.SetMyValue(ss, "PlanID", Convert.ToString(sm.UserInfo.PlanID))
            TIMS.SetMyValue(ss, "ComIDNO", Convert.ToString(ComidValue.Value))
            TIMS.SetMyValue(ss, "FirstSort", FirstSort.Text)
            'TIMS.SetMyValue(ss, "AppStage", AppStage.SelectedValue)
            TIMS.SetMyValue(ss, "AppStage", v_rbl_AppStage) 'rbl_AppStage.SelectedValue)

            Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
            Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
            Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
            Dim PCSVALUE As String = If(rqPlanID <> "" AndAlso rqComIDNO <> "" AndAlso rqSeqNO <> "", String.Concat(rqPlanID, "x", rqComIDNO, "x", rqSeqNO), "")
            TIMS.SetMyValue(ss, "PCSVALUE", PCSVALUE)
            Dim o_OthClassName1 As String = ""
            Dim flagFSort As Boolean = TIMS.Chk_FirstSort1(ss, objconn, o_OthClassName1)

            If flagFSort AndAlso o_OthClassName1 <> "" Then ErrMsg &= String.Concat(cst_errmsg35, "重複：", o_OthClassName1) & vbCrLf
        End If

        'Dim v_Radiobuttonlist1 As String = TIMS.GetListValue(Radiobuttonlist1)
        Select Case v_Radiobuttonlist1'Radiobuttonlist1.SelectedValue
            Case cst_學分班 ' "Y"
                Dim flag_no_PointType As Boolean = False
                If PointType.SelectedIndex = -1 Then flag_no_PointType = True 'ErrMsg &= "學分班種類為必填欄位" & vbCrLf
                Select Case v_PointType'PointType.SelectedValue
                    Case "1", "2", "3"
                    Case Else
                        flag_no_PointType = True
                        'ErrMsg &= "學分班種類為必填欄位" & vbCrLf
                End Select
                If flag_no_PointType Then ErrMsg &= "學分班種類為必填欄位" & vbCrLf

            Case cst_非學分班 'cst_非學分班 '非學分班
                Dim gvid20x1 As String = TIMS.GetGlobalVar(Me, "20", "1", objconn)
                Dim gvid20x2 As String = TIMS.GetGlobalVar(Me, "20", "2", objconn)
                gvid20x1 = TIMS.ClearSQM(gvid20x1)
                gvid20x2 = TIMS.ClearSQM(gvid20x2)
                If gvid20x1 = "" OrElse gvid20x2 = "" Then
                    ErrMsg &= "請至首頁>>系統管理>>系統參數管理>>參數設定裡設定訓練時數" '(未設定) '訓練時數設定
                    Exit Sub
                End If
                '訓練時數設定
                If THours.Text <> "" And gvid20x1 <> "" Then
                    If CInt(Val(THours.Text)) > CInt(gvid20x1) Then ErrMsg &= "若為【非學分班】，訓練時數不得大於" & CInt(gvid20x1) & vbCrLf
                End If
                If THours.Text <> "" And gvid20x2 <> "" Then
                    If CInt(Val(THours.Text)) < CInt(gvid20x2) Then ErrMsg &= "若為【非學分班】，訓練時數必須大於等於" & CInt(gvid20x2) & vbCrLf
                End If

                Dim sqls As String = ""
                Dim drs As DataRow = Nothing
                Select Case strYears
                    Case cst_strYears_2014
                        If jobValue.Value <> "" Then
                            Dim pms1 As New Hashtable From {{"TMID", jobValue.Value}}
                            sqls = " SELECT GCID,GCID2 FROM KEY_TRAINTYPE WHERE TMID=@TMID"
                            drs = DbAccess.GetOneRow(sqls, objconn, pms1)
                        End If
                    Case cst_strYears_2015 '"2014", "2015"
                        If jobValue.Value <> "" Then
                            Dim pms1 As New Hashtable From {{"TMID", jobValue.Value}}
                            sqls = " SELECT GCID,GCID2 FROM KEY_TRAINTYPE WHERE TMID=@TMID"
                            drs = DbAccess.GetOneRow(sqls, objconn, pms1)
                        End If
                    Case cst_strYears_2018 '"2018"
                        If trainValue.Value <> "" Then
                            Dim pms1 As New Hashtable From {{"TMID", trainValue.Value}}
                            sqls = " SELECT JOBID,TRAINID,GCID3 FROM VIEW_TRAINTYPE WHERE TMID=@TMID"
                            drs = DbAccess.GetOneRow(sqls, objconn, pms1)
                        End If
                End Select
                If drs Is Nothing Then
                    ErrMsg &= "[訓練業別]資料異常,請更正" & vbCrLf
                    Exit Sub
                End If

                Select Case strYears
                    Case cst_strYears_2014 '"2014"
                        If Convert.ToString(drs("GCID")) <> "" Then
                            Dim pms2 As New Hashtable From {{"GCID", drs("GCID")}}
                            Dim sqls2 As String = "SELECT GCODE1 FROM ID_GOVCLASSCAST WHERE GCID=@GCID"
                            Dim drs2 As DataRow = DbAccess.GetOneRow(sqls2, objconn, pms2)
                            If Convert.ToString(GCID1Value.Value) <> Convert.ToString(drs2("GCode1")) Then
                                Dim msgG1 As String = "(G1:" & GCID1Value.Value & "/G2:" & drs2("GCode1") & "/GC:" & drs("GCID") & "/J:" & jobValue.Value & ")"
                                ErrMsg &= "訓練費用編列說明的[經費分類代碼]與[訓練業別]不符,請更正" & msgG1 & vbCrLf
                            End If
                        Else
                            ErrMsg &= "[訓練業別]查無[經費分類代碼]資料異常,請更正" & vbCrLf
                        End If
                    Case cst_strYears_2015 '"2015"
                        If Convert.ToString(drs("GCID2")) <> "" Then
                            Dim pms2 As New Hashtable From {{"GCID2", drs("GCID2")}}
                            Dim sqls2 As String = "SELECT GCODE1 FROM V_GOVCLASSCAST2 WHERE GCID2=@GCID2"
                            Dim drs2 As DataRow = DbAccess.GetOneRow(sqls2, objconn, pms2)
                            If Convert.ToString(GCID1Value.Value) <> Convert.ToString(drs2("GCODE1")) Then
                                Dim msgG1 As String = "(G1:" & GCID1Value.Value & "/G2:" & drs2("GCODE1") & "/GC:" & drs("GCID2") & "/J:" & jobValue.Value & ")"
                                ErrMsg &= "訓練費用編列說明的[經費分類代碼2]與[訓練業別]不符,請更正" & msgG1 & vbCrLf
                            End If
                        Else
                            ErrMsg &= "[訓練業別]查無[經費分類代碼2]資料異常,請更正" & vbCrLf
                        End If
                    Case cst_strYears_2018 '"2018"
                        If Convert.ToString(drs("GCID3")) <> "" Then
                            Dim pms2 As New Hashtable From {{"GCID3", drs("GCID3")}}
                            Dim sqls2 As String = "SELECT TMID,GCODE31,GCID3 FROM V_GOVCLASSCAST3 WHERE GCID3=@GCID3"
                            Dim drs2 As DataRow = DbAccess.GetOneRow(sqls2, objconn, pms2)
                            Dim flagErr8 As Boolean = False
                            If Convert.ToString(GCIDValue.Value) <> Convert.ToString(drs2("GCID3")) Then flagErr8 = True
                            If Convert.ToString(GCID1Value.Value) <> Convert.ToString(drs2("GCODE31")) Then flagErr8 = True
                            If flagErr8 Then
                                Dim msgG1 As String = String.Concat("(G1:", GCID1Value.Value, "/G2:", drs2("GCODE31"), "/GC:", drs("GCID3"), "/J:", drs2("TMID"), ")")
                                ErrMsg &= String.Concat("訓練費用編列說明的[經費分類代碼]與[訓練業別]不符,請更正", msgG1, vbCrLf)
                            End If
                        Else
                            ErrMsg &= "[訓練業別]查無[經費分類代碼]資料異常,請更正" & vbCrLf
                        End If
                End Select
                'ItemVar1(,ItemVar2)
                'https://jira.turbotech.com.tw/browse/TWJOBS-154
                'https://jira.turbotech.com.tw/browse/TIMSC-306
                '產投-班級申請，暫時解除開結訓日期不得超過4個月之卡控 by 20210820 (疫情有太多不確定因素)
                '加卡 所有班級目前暫時結訓日最遲只能到 次年 04/30 (防手殘)， 包含政策性課程 (原邏輯不變)
                If STDate.Text <> "" AndAlso FDDate.Text <> "" AndAlso IsDate(STDate.Text) AndAlso IsDate(FDDate.Text) Then
                    Dim tmpDate_NY As Date = DateAdd(DateInterval.Year, 1, CDate(STDate.Text))
                    Dim tempDate As Date = CDate(String.Format("{0}/04/30", tmpDate_NY.Year.ToString()))
                    If DateDiff(DateInterval.Day, tempDate, CDate(FDDate.Text)) > 0 Then ErrMsg &= "【非學分班】，訓練起迄日期區間，迄日不得超過次年04/30" & vbCrLf

                    Dim tempDate4 As Date = DateAdd(DateInterval.Month, 4, CDate(STDate.Text))
                    If DateDiff(DateInterval.Day, tempDate4, CDate(FDDate.Text)) > 0 Then ErrMsg &= "【非學分班】，訓練起迄日期區間，不得超過4個月" & vbCrLf
                    'v_rbl_AppStage/v_APPSTAGE  '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
                    '(2) 申請階段在「上半年」之課程，結訓日期最遲須在當年度8月底前。
                    Dim tempDate2 As Date = CDate(String.Concat(CDate(STDate.Text).Year, "/8/31"))
                    '(3) 申請階段在「下半年」之課程，結訓日期最遲須在翌年2月底前。
                    Dim tempDate3 As Date = CDate(String.Concat((CDate(STDate.Text).Year + 1), "/3/1"))
                    If v_rbl_AppStage = "1" AndAlso (DateDiff(DateInterval.Day, tempDate2, CDate(FDDate.Text)) > 0) Then
                        ErrMsg &= "非學分班,申請階段在「上半年」之課程，結訓日期最遲須在當年度8月底前。" & vbCrLf
                        'Common.MessageBox(Me, "非學分班,申請階段在「上半年」之課程，結訓日期最遲須在當年度8月底前。") Return
                    ElseIf v_rbl_AppStage = "2" AndAlso (DateDiff(DateInterval.Day, tempDate3, CDate(FDDate.Text)) >= 0) Then
                        ErrMsg &= "非學分班,申請階段在「下半年」之課程，結訓日期最遲須在翌年2月底前。" & vbCrLf
                        'Common.MessageBox(Me, "非學分班,申請階段在「下半年」之課程，結訓日期最遲須在翌年2月底前。") Return
                    End If
                End If

        End Select
        If ErrMsg <> "" Then Exit Sub

        '檢查所選的學術科場地是否有被選中
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_SciPlaceID, 1, 1, objconn)
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_TechPlaceID, 2, 2, objconn)
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_SciPlaceID2, 3, 1, objconn)
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_TechPlaceID2, 4, 2, objconn)
        Dim dtTaddress As DataTable = CType(ViewState("dtTaddress"), DataTable)

        Dim TrainDesc_DataTable As DataTable = If(Not TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)), Nothing, CType(Session(hid_TrainDescTable_guid1.Value), DataTable))
        'Dim x As String = If("a", "b")
        If TrainDesc_DataTable Is Nothing Then
            ErrMsg &= cst_errmsg24 & "!" & vbCrLf '"產學訓課程大綱，為必填資料" & vbCrLf
            Exit Sub
        End If

        Dim iRow As Integer = 0
        Try
            If TrainDesc_DataTable IsNot Nothing Then
                For Each drTrainDesc As DataRow In TrainDesc_DataTable.Rows
                    If Not drTrainDesc.RowState = DataRowState.Deleted Then
                        iRow += 1
                        Dim fg_mach1 As Boolean = False
                        For Each drTaddress As DataRow In dtTaddress.Rows
                            If Not drTaddress.RowState = DataRowState.Deleted Then
                                fg_mach1 = (Convert.ToString(drTrainDesc("PTID")) = Convert.ToString(drTaddress("PTID")))
                                If fg_mach1 Then Exit For
                            End If
                        Next
                        If Not fg_mach1 Then
                            ErrMsg &= String.Concat("[課程大綱]的[上課地點]不在所選的[學術科場地]範圍內,請修改!!(第", iRow, "筆)", vbCrLf)
                            Exit For
                        End If
                        'htTrainDesc("sClassification1")  '1:學/2:術
                        Dim s_TECHID2 As String = Convert.ToString(drTrainDesc("TECHID2"))
                        Dim s_CLASSIFICATION1 As String = Convert.ToString(drTrainDesc("Classification1"))
                        If s_CLASSIFICATION1 = "1" AndAlso s_TECHID2 <> "" Then
                            ErrMsg &= String.Concat("[課程大綱]的學科課程不得規劃助教共同授課!(第", iRow, "筆)", vbCrLf)
                            Exit For
                        End If
                    End If
                Next
            End If
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            ErrMsg &= String.Concat("[課程大綱]的[上課地點]不在所選的[學術科場地]範圍內,請修改!!!(第", iRow, "筆)", vbCrLf)
        End Try
        'end 檢查所選的學術科場地是否有被選中

        'Dim i_MaxTxtLen1 As Integer = 0
        i_MaxTxtLen1 = 1000
        If i_MaxTxtLen1 > 0 AndAlso (tNote2.Text.Length > i_MaxTxtLen1) Then ErrMsg &= String.Format("其他說明(欄位字數為{0})，超過欄位字數", i_MaxTxtLen1) & vbCrLf

        '專長能力標籤-ABILITY-PLAN_ABILITY
        For i_SEQ As Integer = 1 To 4
            Dim s_SEQ As String = Convert.ToString(i_SEQ)
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            otxtA1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA1.Text), 30)
            otxtA2.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA2.Text), 200)
            If otxtA1.Text = "" Then
                ErrMsg &= cst_errmsg36 & vbCrLf
                Exit For
            End If
            Exit For
        Next

        iCAPMARKDATE.Text = TIMS.ClearSQM(iCAPMARKDATE.Text)
        If iCAPMARKDATE.Text <> "" AndAlso Not TIMS.IsDate1(iCAPMARKDATE.Text) Then
            ErrMsg &= "班別資料-【iCAP標章有效期限】有填寫, 日期格式有誤!" & vbCrLf
        ElseIf iCAPMARKDATE.Text <> "" Then
            iCAPMARKDATE.Text = TIMS.Cdate3(iCAPMARKDATE.Text)
        End If

        If ErrMsg <> "" Then Exit Sub '(檢核)結束
    End Sub

    ''' <summary> 基本儲存/--正式儲存--(檢核) </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        Hid_sender1.Value = cst_SaveBasic 'sender.text

        Call UTL_PREARRANGEMENT()
        'Dim flag_TIMS_Test_1 As Boolean = TIMS.sUtl_ChkTest()
        Dim RegKey1 As String = "_onload"
        'Dim RegScript1 As String = "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";Layer_change(5);</script>"
        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim RegScript1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")

        Dim sErrMsg As String = ""
        '基本儲存-資料檢核確認
        Call CheckAddData(sErrMsg)
        If sErrMsg <> "" Then
            '有錯誤訊息
            Page.RegisterStartupScript(RegKey1, RegScript1)
            sm.LastErrorMessage = sErrMsg
            Exit Sub 'Return False '不可儲存
        End If

        '檢核是否有業務權限
        Dim flag_CAN_NOCHECK_RIDPLAN As Boolean = False 'true:(可以不檢查業務權限) false:必須檢查
        Dim flag_RIDPLAN As Boolean = Chk_RIDPLAN(RIDValue.Value, sm.UserInfo.PlanID)
        If flag_IsSuperUser_1 Then flag_CAN_NOCHECK_RIDPLAN = True
        'If flag_TIMS_Test_1 Then flag_CAN_NOCHECK_RIDPLAN = True '檢核是否有業務權限
        If Not flag_CAN_NOCHECK_RIDPLAN AndAlso Not flag_RIDPLAN Then
            sErrMsg &= $"{cst_errmsg22},{RIDValue.Value},{sm.UserInfo.PlanID},{sm.UserInfo.Years}{vbCrLf}" '"登入者無正確的業務權限，不提供儲存服務!!" & vbCrLf
        End If
        If sErrMsg <> "" Then
            '有錯誤訊息
            Page.RegisterStartupScript(RegKey1, RegScript1)
            sm.LastErrorMessage = sErrMsg
            Exit Sub 'Return False '不可儲存
        End If

        '儲存點
        '假設處理某段程序需花費n毫秒 (避免機器不同步)
        If Session("GUID1") <> ViewState("GUID1") Then Threading.Thread.Sleep(1)
        ViewState("GUID1") = TIMS.GetGUID() : Session("GUID1") = ViewState("GUID1")

        '儲存 開班計畫/開班計畫表資料維護 '(基本儲存)
        'Try
        '    Call INSERT_PLAN_TABLE(cst_SaveBasic) '若儲存成功，則下列不執行，直接跳頁 ../01/TC_01_014_add.aspx
        'Catch ex As Exception
        '    Dim strErrmsg As String = ex.Message
        '    strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
        '    strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
        '    'strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
        '    Call TIMS.WriteTraceLog(strErrmsg, ex)
        '    sm.LastErrorMessage = String.Concat(cst_errmsg6, vbCrLf, ex.Message)
        '    Return ' Exit Sub
        'End Try
        '儲存 開班計畫/開班計畫表資料維護 '(基本儲存)
        Call INSERT_PLAN_TABLE(cst_SaveBasic) '若儲存成功，則下列不執行，直接跳頁 ../01/TC_01_014_add.aspx
        'ViewState("dtTaddress") = Nothing
        Page.RegisterStartupScript(RegKey1, RegScript1)
    End Sub

    ''' <summary> 鎖定輸入項。</summary>
    ''' <param name="sTitle"></param>
    Sub Disabled_Items(ByVal sTitle As String)
        EMail.ReadOnly = True
        trainValue.Disabled = True
        cjobValue.Disabled = True
        jobValue.Disabled = True
        PlanCause.ReadOnly = True
        PurScience.ReadOnly = True
        PurTech.ReadOnly = True
        PurMoral.ReadOnly = True
        Degree.Enabled = False
        Other1.ReadOnly = True
        Other2.ReadOnly = True
        Other3.ReadOnly = True
        TIMS.Tooltip(Other1, Cst_msgother1)
        TIMS.Tooltip(Other2, Cst_msgother1)
        TIMS.Tooltip(Other3, Cst_msgother1)

        TMScience.ReadOnly = True
        SciHours.ReadOnly = True
        GenSciHours.ReadOnly = True
        ProSciHours.ReadOnly = True
        ProTechHours.ReadOnly = True
        TotalHours.ReadOnly = True
        ClassName.ReadOnly = True
        TNum.ReadOnly = True
        STDate.ReadOnly = True
        FDDate.ReadOnly = True
        CyclType.ReadOnly = True
        CustomValidator4.Enabled = False
        ClassCount.ReadOnly = True
        DefGovCost.ReadOnly = True
        DefStdCost.ReadOnly = True
        TIMS.Tooltip(DefGovCost, sTitle)
        TIMS.Tooltip(DefStdCost, sTitle)

        Note.ReadOnly = True
        CredPoint.ReadOnly = True
        RoomName.ReadOnly = True
        FactMode.Enabled = False
        FactModeOther.ReadOnly = True
        rbl_OUTDOOR.Enabled = False
        ConNum.ReadOnly = True
        ContactName.ReadOnly = True
        '2023/2024/ContactPhone
        ContactPhone.ReadOnly = True
        ContactPhone_1.ReadOnly = True
        ContactPhone_2.ReadOnly = True
        ContactPhone_3.ReadOnly = True
        ContactMobile_1.ReadOnly = True
        ContactMobile_2.ReadOnly = True
        ContactEmail.ReadOnly = True
        ContactFax.ReadOnly = True
        ClassCate.Enabled = False
        Content.ReadOnly = True

        Button29.Enabled = False
        btnAddBusPackage.Enabled = False
        center.Enabled = False
        Org.Disabled = True

        Button8.Visible = False '草稿儲存
        btnAdd.Visible = False '基本儲存
        BtnSAVE2.Visible = False '正式儲存

        btu_sel.Disabled = True
        TIMS.Tooltip(btu_sel, sTitle)
        btu_sel2.Disabled = True
        TIMS.Tooltip(btu_sel2, sTitle)
    End Sub

    '回上一頁
    Private Sub Button24_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button24.Click
        If Convert.ToString(ViewState("search")) <> "" AndAlso Session("search") Is Nothing Then Session("search") = ViewState("search")
        Dim url1 As String = ""
        If TIMS.ClearSQM(Request("todo")) = "1" Then
            url1 = "../04/TC_04_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        ElseIf gflag_ccopy Then
            url1 = "../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        Else
            url1 = "../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        End If
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    ''' <summary>機構資訊(隱藏) 點選機構時重新載入</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button28_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button28.Click

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim s_RID As String = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim parms As New Hashtable From {{"RID", s_RID}}
        Dim dr As DataRow = Nothing
        Dim sql As String = ""
        sql &= " SELECT b.ComIDNO,c.ContactEmail, c.ZipCode, c.Address, c.ContactName, c.Phone, c.ContactEmail, c.ContactFax"
        sql &= " FROM AUTH_RELSHIP a WITH(NOLOCK)"
        sql &= " JOIN ORG_ORGINFO b WITH(NOLOCK) ON a.OrgID=b.OrgID"
        sql &= " JOIN ORG_ORGPLANINFO c WITH(NOLOCK) ON a.RSID=c.RSID"
        sql &= " WHERE a.RID=@RID"
        dr = DbAccess.GetOneRow(sql, objconn, parms)
        If dr Is Nothing Then
            Page.RegisterStartupScript("Londing", "<script>Layer_change('');</script>")
            sm.LastErrorMessage = cst_errmsg1
            Exit Sub
        End If
        ComidValue.Value = Convert.ToString(dr("ComIDNO"))
        If Table1_Email.Visible = True Then EMail.Text = Convert.ToString(dr("ContactEmail"))
        ContactName.Text = dr("ContactName").ToString
        '2023/2024/ContactPhone
        If ContactPhone.Text <> "" Then ContactPhone.Text = Convert.ToString(dr("Phone"))
        Dim hCtPhone As New Hashtable
        Call TIMS.CHK_ContactPhoneFMT(Convert.ToString(dr("Phone")), hCtPhone)
        If ContactPhone_1.Text <> "" Then ContactPhone_1.Text = hCtPhone("ContactPhone_1")
        If ContactPhone_2.Text <> "" Then ContactPhone_2.Text = hCtPhone("ContactPhone_2")
        If ContactPhone_3.Text <> "" Then ContactPhone_3.Text = hCtPhone("ContactPhone_3")
        Dim hCtMobile As New Hashtable
        Call TIMS.CHK_ContactMobileFMT(Convert.ToString(dr("Phone")), hCtMobile)
        If ContactMobile_1.Text <> "" Then ContactMobile_1.Text = hCtMobile("ContactMobile_1")
        If ContactMobile_2.Text <> "" Then ContactMobile_2.Text = hCtMobile("ContactMobile_2")

        ContactEmail.Text = dr("ContactEmail").ToString
        ContactFax.Text = dr("ContactFax").ToString

        SciPlaceID = TIMS.Get_SciPlaceID(SciPlaceID, ComidValue.Value, 2, "", objconn)
        TechPlaceID = TIMS.Get_TechPlaceID(TechPlaceID, ComidValue.Value, 2, "", objconn)

        SciPlaceID2 = TIMS.Get_SciPlaceID(SciPlaceID2, ComidValue.Value, 2, "", objconn)
        TechPlaceID2 = TIMS.Get_TechPlaceID(TechPlaceID2, ComidValue.Value, 2, "", objconn)
        Page.RegisterStartupScript("Londing", "<script>Layer_change('');</script>")
    End Sub

    '清理時間格式 '18:00~21:00，多筆以 ;  false:上課時間／時間內容，格式有誤
    Function CheckTimesFMT(ByRef s_Times As String) As Boolean
        Dim rst As Boolean = False '異常
        's_Times = TIMS.ClearSQM(s_Times)
        'If (s_Times.IndexOf("；") > -1) Then s_Times = Replace(s_Times, "；", ";")
        'If (s_Times.IndexOf("：") > -1) Then s_Times = Replace(s_Times, "：", ":")
        'If (s_Times.IndexOf("～") > -1) Then s_Times = Replace(s_Times, "～", "~")
        If s_Times = "" Then Return True '沒有值直接退出

        If (s_Times.IndexOf(";") = -1) Then
            If s_Times.Length <> 11 Then
                Return rst '異常 退出
            Else
                Dim as_Timesf2 As String() = s_Times.Split("~") '00:33~11:22
                If as_Timesf2.Length <> 2 Then Return rst
                Dim as_Timesf20 As String() = as_Timesf2(0).Split(":") '00:33
                If as_Timesf20.Length <> 2 Then Return rst
                If as_Timesf20(0).Length <> 2 Then Return rst '00
                If as_Timesf20(1).Length <> 2 Then Return rst '33
                Dim as_Timesf21 As String() = as_Timesf2(1).Split(":") '11:22
                If as_Timesf21.Length <> 2 Then Return rst
                If as_Timesf21(0).Length <> 2 Then Return rst '11
                If as_Timesf21(1).Length <> 2 Then Return rst '22
                '異常 退出
            End If
            Return True 'ok
        End If

        Dim as_Times As String() = s_Times.Split(";")
        For Each sTimeV1 As String In as_Times
            Dim flag_TimeV1 As Boolean = CheckTimesFMT(sTimeV1) '有值再檢核
            If Not flag_TimeV1 Then Return flag_TimeV1 '異常 退出
        Next
        Return True 'ok
    End Function

    '新增 上課時間／內容
    Private Sub Button29_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button29.Click
        Dim Errmsg As String = ""
        Dim v_Weeks As String = TIMS.GetListValue(Weeks)
        txtTimes.Text = TIMS.ClearTimesFMT(txtTimes.Text)
        Dim i_Times_c_max_length As Integer = cst_i_Times_c_max_length
        Dim i_Times_c_min_length As Integer = cst_i_Times_c_min_length
        Dim s_err_msg1 As String = String.Format("上課時間／時間內容，長度超過限制範圍{0}文字長度", i_Times_c_max_length)
        Dim s_err_msg2 As String = String.Format("上課時間／時間內容，長度小於限制範圍{0}文字長度", i_Times_c_min_length)
        If txtTimes.Text <> "" Then
            If txtTimes.Text.ToString.Length > i_Times_c_max_length Then Errmsg &= s_err_msg1 & vbCrLf
            If txtTimes.Text.ToString.Length < i_Times_c_min_length Then Errmsg &= s_err_msg2 & vbCrLf
        Else
            Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
        End If
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "5"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11) 'window.scroll(0,document.body.scrollHeight);
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        If Session(hid_planONCLASS_guid1.Value) Is Nothing Then Call CreateClassTime()
        dt = Session(hid_planONCLASS_guid1.Value)
        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("POCID") = TIMS.GET_NEWPK_INT(Me, "POCID")
        dr("Weeks") = v_Weeks 'vsWeeks
        dr("Times") = txtTimes.Text
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        Session(hid_planONCLASS_guid1.Value) = dt
        Call CreateClassTime()

        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1) 'window.scroll(0,document.body.scrollHeight);
    End Sub

    ''' <summary>再次檢核 (server check)</summary>
    ''' <param name="Errmsg"></param>
    ''' <param name="htTrainDesc"></param>
    ''' <param name="dt"></param>
    ''' <returns></returns>
    Function Chk_TrainDescInput(ByRef Errmsg As String, ByRef htTrainDesc As Hashtable, ByRef dt As DataTable) As Boolean
        Dim rst As Boolean = True
        Dim sTPERIOD28 As String = htTrainDesc("sTPERIOD28")
        Dim sSTDate As String = htTrainDesc("sSTDate")
        Dim sFDDate As String = htTrainDesc("sFDDate")
        Dim sPName As String = htTrainDesc("sPName")
        Dim sddlpnH1 As String = htTrainDesc("sddlpnH1")
        Dim sddlpnM1 As String = htTrainDesc("sddlpnM1")
        Dim sddlpnH2 As String = htTrainDesc("sddlpnH2")
        Dim sddlpnM2 As String = htTrainDesc("sddlpnM2")
        Dim sPHour As String = htTrainDesc("sPHour") '時數
        Dim sEHour As String = htTrainDesc("sEHour") '技檢訓練時數
        Dim sAIAHour As String = htTrainDesc("sAIAHour") 'AI應用時數
        Dim sWNLHour As String = htTrainDesc("sWNLHour") '職場續航時數
        Dim sPCont As String = htTrainDesc("sPCont")
        Dim sSTrainDate As String = htTrainDesc("sSTrainDate")
        Dim sClassification1 As String = htTrainDesc("sClassification1")  '1:學/2:術
        Dim sComidValue As String = htTrainDesc("sComidValue")
        Dim sPTID1 As String = htTrainDesc("sPTID1")
        Dim sPTID2 As String = htTrainDesc("sPTID2")
        Dim sOLessonTeah1Value As String = htTrainDesc("sOLessonTeah1Value")
        Dim sOLessonTeah2Value As String = htTrainDesc("sOLessonTeah2Value")
        Dim sPTDID As String = htTrainDesc("sPTDID")

        Errmsg = ""
        If sTPERIOD28 = cst_NNN Then
            Errmsg &= "授課時段:早上、下午、晚上 至少要設定其中一項" & vbCrLf
        ElseIf TIMS.CHK_STR_CNT(sTPERIOD28, "Y") > 1 Then
            Errmsg &= "授課時段:早上、下午、晚上為單選"
        End If
        sSTDate = TIMS.ClearSQM(sSTDate)
        If sSTDate = "" Then
            Errmsg &= "請先輸入訓練起日" & vbCrLf
        ElseIf Not TIMS.IsDate1(sSTDate) Then
            Errmsg &= "訓練起日請填日期格式" & vbCrLf
            sSTDate = TIMS.Cdate3(sSTDate)
        End If

        sFDDate = TIMS.ClearSQM(sFDDate)
        If sFDDate = "" Then
            Errmsg &= "請先輸入訓練迄日" & vbCrLf
        ElseIf Not TIMS.IsDate1(sFDDate) Then
            Errmsg &= "訓練迄日請填日期格式" & vbCrLf
            sFDDate = TIMS.Cdate3(sFDDate)
        End If

        If Errmsg = "" Then
            Select Case True
                Case DateDiff(DateInterval.Day, CDate(sSTDate), CDate(sFDDate)) < 0
                    Errmsg &= "訓練起日不能超過訓練迄日" & vbCrLf
                Case DateDiff(DateInterval.Day, CDate(sSTDate), CDate(sFDDate)) = 0
                    Errmsg &= "訓練起日不能和訓練迄日同一天" & vbCrLf
            End Select
        End If
        sPName = sddlpnH1 & ":" & sddlpnM1 & "~" & sddlpnH2 & ":" & sddlpnM2

        '20090318--原程式判斷是否整數是mark起來的，將其mark取消。
        sPHour = TIMS.ClearSQM(sPHour)
        If sPHour = "" Then '上課時數
            Errmsg &= "【時數】未填寫，請填數字" & vbCrLf
        ElseIf Not IsNumeric(sPHour) Then '上課時數
            Errmsg &= "【時數】請填數字格式" & vbCrLf
        ElseIf Not TIMS.IsNumeric2(sPHour) Then
            Errmsg &= "【時數】格式 內容需為整數，請修正!!" & vbCrLf
        ElseIf Val(sPHour) <= 0 Then
            Errmsg &= "【時數】必須大於0" & vbCrLf
        ElseIf Val(sPHour) > 4 Then
            Errmsg &= "【時數】必須小於等於4" & vbCrLf
        End If
        '僅提供填寫小數點後為5，或是整數的數字
        'If Not TIMS.IsNumeric5(sPHour) Then
        '    Errmsg &= "【時數】格式僅提供填寫小數點後為5，或是整數的數字，且不得為負數" & vbCrLf
        '    flag_CHECK_sPHour_OK = False
        'End If

        '技檢訓練時數 產業人才投資方案(充飛不用)
        '1.【技檢訓練時數】需<=該堂課【時數】，可允許小數點一位
        '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-1才存欄位，否清空。
        '3.權限：跟其他欄位一樣。'訓練單位可填寫，但送審後鎖住不可修改。送審後分署可修改。'3.訓練班節計畫表加上【技檢時數】欄位顯示。
        '4.註記：'調整班級變更申請、班級變更審核、結訓證書等功能 '結訓證書上顯示： 符合申請技檢訓練時數
        sEHour = TIMS.ClearSQM(sEHour) '技檢訓練時數
        If Errmsg = "" AndAlso sPHour <> "" AndAlso sEHour <> "" Then
            If Not IsNumeric(sEHour) Then '上課時數
                Errmsg &= "【技檢訓練時數】請填數字，必須大於0，若為0則毋須填寫" & vbCrLf
            ElseIf Not TIMS.IsNumeric2(sEHour) AndAlso Not TIMS.CheckFloat1(sEHour) Then
                Errmsg &= "【技檢訓練時數】格式 內容需為整數數字 或 小數點1位數字!" & vbCrLf
            ElseIf Val(sEHour) <= 0 Then
                Errmsg &= "【技檢訓練時數】必須大於0，若為0則毋須填寫" & vbCrLf
            ElseIf Val(sEHour) > Val(sPHour) Then
                Errmsg &= "【技檢訓練時數】必須小於等於該堂課【時數】!" & vbCrLf
            End If
        End If
        sAIAHour = TIMS.ClearSQM(sAIAHour) 'AI應用時數
        If Errmsg = "" AndAlso sPHour <> "" AndAlso sAIAHour <> "" Then
            If Not IsNumeric(sAIAHour) Then '上課時數
                Errmsg &= "【AI應用時數】請填數字，必須大於0，若為0則毋須填寫" & vbCrLf
            ElseIf Not TIMS.IsNumeric2(sAIAHour) AndAlso Not TIMS.CheckFloat1(sAIAHour) Then
                Errmsg &= "【AI應用時數】格式 內容需為整數數字 或 小數點1位數字!" & vbCrLf
            ElseIf Val(sAIAHour) <= 0 Then
                Errmsg &= "【AI應用時數】必須大於0，若為0則毋須填寫" & vbCrLf
            ElseIf Val(sAIAHour) > Val(sPHour) Then
                Errmsg &= "【AI應用時數】必須小於等於該堂課【時數】!" & vbCrLf
            End If
        End If
        sWNLHour = TIMS.ClearSQM(sWNLHour) '職場續航時數
        If Errmsg = "" AndAlso sPHour <> "" AndAlso sWNLHour <> "" Then
            If Not IsNumeric(sWNLHour) Then '上課時數
                Errmsg &= "【職場續航時數】請填數字，必須大於0，若為0則毋須填寫" & vbCrLf
            ElseIf Not TIMS.IsNumeric2(sWNLHour) AndAlso Not TIMS.CheckFloat1(sWNLHour) Then
                Errmsg &= "【職場續航時數】格式 內容需為整數數字 或 小數點1位數字!" & vbCrLf
            ElseIf Val(sWNLHour) <= 0 Then
                Errmsg &= "【職場續航時數】必須大於0，若為0則毋須填寫" & vbCrLf
            ElseIf Val(sWNLHour) > Val(sPHour) Then
                Errmsg &= "【職場續航時數】必須小於等於該堂課【時數】!" & vbCrLf
            End If
        End If

        'Dim flag_CHECK_sPHour_OK As Boolean = True '避免有錯誤的判斷使用
        'If flag_CHECK_sPHour_OK AndAlso sPHour = "" Then '上課時數
        '    Errmsg &= "時數未填寫，請填數字" & vbCrLf
        '    flag_CHECK_sPHour_OK = False
        'End If
        'If flag_CHECK_sPHour_OK AndAlso Not IsNumeric(sPHour) Then '上課時數
        '    Errmsg &= "時數請填數字格式" & vbCrLf
        '    flag_CHECK_sPHour_OK = False
        'End If
        ''僅提供填寫小數點後為5，或是整數的數字
        ''If flag_CHECK_sPHour_OK AndAlso Not TIMS.IsNumeric5(sPHour) Then
        ''    Errmsg &= "時數格式僅提供填寫小數點後為5，或是整數的數字，且不得為負數" & vbCrLf
        ''    flag_CHECK_sPHour_OK = False
        ''End If
        'If flag_CHECK_sPHour_OK AndAlso Not TIMS.IsNumeric2(sPHour) Then
        '    Errmsg &= "時數格式 內容需為整數，請修正!!" & vbCrLf
        '    flag_CHECK_sPHour_OK = False
        'End If
        ''課程大綱
        'If flag_CHECK_sPHour_OK AndAlso Val(sPHour) <= 0 Then
        '    Errmsg &= "時數必須大於0" & vbCrLf
        '    flag_CHECK_sPHour_OK = False
        'End If
        'If flag_CHECK_sPHour_OK AndAlso Val(sPHour) > 4 Then
        '    Errmsg &= "時數必須小於等於4" & vbCrLf
        '    flag_CHECK_sPHour_OK = False
        'End If

        sPCont = TIMS.ClearSQM(sPCont)
        '課程進度／內容
        If sPCont = "" Then
            Errmsg &= "課程進度／內容未填寫" & vbCrLf
        Else
            If sPCont.Length > 250 Then Errmsg &= "課程進度／內容，長度超過限制範圍250文字長度" & vbCrLf
        End If

        If sSTrainDate = "" Then
            Errmsg &= "請輸入上課日期" & vbCrLf
        Else
            If Not TIMS.IsDate1(sSTrainDate) Then
                Errmsg &= "上課日期請填日期格式" & vbCrLf
            Else
                '計算當日時數不可大於8小時
                Dim ixHour As Double = 0 '計算當日時數不可大於8小時
                If Errmsg = "" Then sSTrainDate = TIMS.Cdate3(sSTrainDate) '.ToString("yyyy/MM/dd")
                If Not dt Is Nothing Then
                    If sPTDID <> "" Then
                        '修改
                        If dt.Rows.Count > 0 AndAlso IsNumeric(sPTDID) Then
                            '計算當日時數不可大於8小時
                            'Dim ixHour As Double = 0 '計算當日時數不可大於8小時
                            ixHour = 0
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted _
                                    AndAlso dt.Select("PTDID='" & sPTDID & "'").Length = 0 Then '已刪除者不可做更動
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 _
                                        AndAlso dt.Rows(i).Item("PName").ToString = sPName Then
                                        Errmsg &= "此日期+授課時間已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 Then
                                        Try
                                            ixHour += Val(dt.Rows(i).Item("PHour"))
                                        Catch ex As Exception
                                            Errmsg &= "上課時數異常，請重新載入計算" & vbCrLf
                                            Exit For
                                        End Try
                                    End If
                                End If
                            Next
                            If Errmsg = "" Then
                                ixHour += Val(sPHour)
                                If ixHour > 8 Then Errmsg &= "該日上課時數超過8小時，請重新填寫" & vbCrLf
                            End If
                        End If
                    Else
                        '新增
                        If dt.Rows.Count > 0 Then
                            'Dim ixHour As Double = 0 '計算當日時數不可大於8小時
                            ixHour = 0
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Convert.ToString(dt.Rows(i).Item("STrainDate")) <> "" AndAlso sSTrainDate <> "" Then
                                        If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 _
                                        AndAlso dt.Rows(i).Item("PName").ToString = sPName Then
                                            Errmsg &= "此日期+授課時間已在表格中" & vbCrLf
                                            Exit For
                                        End If
                                    End If
                                    If Convert.ToString(dt.Rows(i).Item("STrainDate")) = "" Then
                                        Errmsg &= "課程大綱-日期資料 不可為空" & vbCrLf
                                        Exit For
                                    End If
                                    If DateDiff(DateInterval.Day, CDate(dt.Rows(i).Item("STrainDate").ToString), CDate(sSTrainDate)) = 0 Then
                                        Try
                                            ixHour += Val(dt.Rows(i).Item("PHour"))
                                        Catch ex As Exception
                                            Errmsg &= "上課時數異常，請重新載入計算" & vbCrLf
                                            Exit For
                                        End Try
                                    End If
                                End If
                            Next
                            If Errmsg = "" Then
                                ixHour += Val(sPHour)
                                If ixHour > 8 Then Errmsg &= "該日上課時數超過8小時，請重新填寫" & vbCrLf
                            End If
                        End If
                    End If
                End If
            End If
        End If

        If Errmsg = "" Then
            Select Case True
                Case DateDiff(DateInterval.Day, CDate(sSTDate), CDate(sSTrainDate)) < 0
                    Errmsg &= "課程大網日期不能超過訓練起日" & vbCrLf
                Case DateDiff(DateInterval.Day, CDate(sSTrainDate), CDate(sFDDate)) < 0
                    Errmsg &= "課程大網日期不能超過訓練迄日" & vbCrLf
            End Select
        End If

        '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
        Select Case CInt(sClassification1)  '1:學/2:術
            Case 0
                Errmsg &= "請選擇課程大網的學/術科(必填)" & vbCrLf
            Case 1 '1:學
                If sPTID1 = "" Then
                    Errmsg &= "請選擇課程大網的上課地點(必填)" & vbCrLf
                ElseIf Not TIMS.Check_SciPTID(sComidValue, sPTID1, objconn) Then
                    Errmsg &= "課程大網的上課地點學科場地已被刪除，請重新選擇" & vbCrLf
                End If
            Case 2 '2:術
                If sPTID2 = "" Then
                    Errmsg &= "請選擇課程大網的上課地點(必填)" & vbCrLf
                ElseIf Not TIMS.Check_TechPTID(sComidValue, sPTID2, objconn) Then
                    Errmsg &= "課程大網的上課地點術科場地已被刪除，請重新選擇" & vbCrLf
                End If
        End Select
        If sOLessonTeah1Value = "" Then
            Errmsg &= "請選擇課程大網的任課教師(必填)" & vbCrLf
        ElseIf sOLessonTeah2Value = sOLessonTeah1Value Then
            Errmsg &= "任課教師與助教為同一人錯誤" & vbCrLf
        End If
        'htTrainDesc("sClassification1")  '1:學/2:術
        If CInt(sClassification1) = 1 AndAlso sOLessonTeah2Value <> "" Then
            Errmsg &= "學科課程不得規劃助教共同授課" & vbCrLf
        End If

        htTrainDesc("sTPERIOD28") = sTPERIOD28
        htTrainDesc("sSTDate") = sSTDate
        htTrainDesc("sFDDate") = sFDDate
        htTrainDesc("sPName") = sPName
        htTrainDesc("sddlpnH1") = sddlpnH1
        htTrainDesc("sddlpnM1") = sddlpnM1
        htTrainDesc("sddlpnH2") = sddlpnH2
        htTrainDesc("sddlpnM2") = sddlpnM2
        htTrainDesc("sPHour") = sPHour '時數
        htTrainDesc("sEHour") = sEHour '技檢訓練時數
        htTrainDesc("sAIAHour") = sAIAHour 'AI應用時數
        htTrainDesc("sWNLHour") = sWNLHour '職場續航時數
        htTrainDesc("sPCont") = sPCont
        htTrainDesc("sSTrainDate") = sSTrainDate
        htTrainDesc("sClassification1") = sClassification1  '1:學/2:術
        htTrainDesc("sComidValue") = sComidValue
        htTrainDesc("sPTID1") = sPTID1
        htTrainDesc("sPTID2") = sPTID2
        htTrainDesc("sOLessonTeah1Value") = sOLessonTeah1Value
        htTrainDesc("sOLessonTeah2Value") = sOLessonTeah2Value

        If Errmsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>新增 課程大網 SESSION (PLAN_TRAINDESC)</summary>
    Sub INSERT_TRAINDESC_SESS()
        'If Session(hid_TrainDescTable_guid1.Value) Is Nothing Then Call CreateTrainDesc() 'PLAN_TRAINDESC
        If Not TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then Call CreateTrainDesc() 'PLAN_TRAINDESC 
        Dim dt As DataTable = Session(hid_TrainDescTable_guid1.Value) 'PLAN_TRAINDESC 
        Dim dr As DataRow = Nothing
        Dim Errmsg As String = ""

        Dim sTPERIOD28 As String = $"{If(TPERIOD28_1.Checked, "Y", "N")}{If(TPERIOD28_2.Checked, "Y", "N")}{If(TPERIOD28_3.Checked, "Y", "N")}" '授課時段'早上'下午'晚上
        STDate.Text = TIMS.ClearSQM(STDate.Text)
        FDDate.Text = TIMS.ClearSQM(FDDate.Text)
        PName.Text = TIMS.ClearSQM(PName.Text)
        Dim v_ddlpnH1 As String = TIMS.GetListValue(ddlpnH1)
        Dim v_ddlpnM1 As String = TIMS.GetListValue(ddlpnM1)
        Dim v_ddlpnH2 As String = TIMS.GetListValue(ddlpnH2)
        Dim v_ddlpnM2 As String = TIMS.GetListValue(ddlpnM2)
        Dim v_Classification1 As String = TIMS.GetListValue(Classification1)
        Dim v_PTID1 As String = TIMS.GetListValue(PTID1)
        Dim v_PTID2 As String = TIMS.GetListValue(PTID2)
        'Dim vsddlpnH1 As String = TIMS.ClearSQM(ddlpnH1.SelectedValue)
        'Dim vsddlpnM1 As String = TIMS.ClearSQM(ddlpnM1.SelectedValue)
        'Dim vsddlpnH2 As String = TIMS.ClearSQM(ddlpnH2.SelectedValue)
        'Dim vsddlpnM2 As String = TIMS.ClearSQM(ddlpnM2.SelectedValue)
        PHour.Text = TIMS.ClearSQM(PHour.Text) '時數
        EHOUR.Text = TIMS.ClearSQM(EHOUR.Text) '技檢訓練時數
        AIAHOUR.Text = TIMS.ClearSQM(AIAHOUR.Text) 'AI應用時數
        WNLHOUR.Text = TIMS.ClearSQM(WNLHOUR.Text) '職場續航時數

        PCont.Text = TIMS.ClearSQM(PCont.Text)
        STrainDate.Text = TIMS.ClearSQM(STrainDate.Text)
        'Dim vsClassification1 As String = TIMS.ClearSQM(Classification1.SelectedValue)
        ComidValue.Value = TIMS.ClearSQM(ComidValue.Value)
        'Dim vsPTID1 As String = TIMS.ClearSQM(PTID1.SelectedValue)
        'Dim vsPTID2 As String = TIMS.ClearSQM(PTID2.SelectedValue)
        Dim vsOLessonTeah1Value As String = TIMS.ClearSQM(OLessonTeah1Value.Value)
        Dim vsOLessonTeah2Value As String = TIMS.ClearSQM(OLessonTeah2Value.Value)

        TIMS.Chk_placeholder(PCont)
        Dim htTrainDesc As New Hashtable From {
            {"sTPERIOD28", sTPERIOD28}, '授課時段
            {"sSTDate", STDate.Text},
            {"sFDDate", FDDate.Text},
            {"sPName", PName.Text},
            {"sddlpnH1", v_ddlpnH1},
            {"sddlpnM1", v_ddlpnM1},
            {"sddlpnH2", v_ddlpnH2},
            {"sddlpnM2", v_ddlpnM2},
            {"sPHour", PHour.Text}, '時數
            {"sEHour", EHOUR.Text}, '技檢訓練時數
            {"sAIAHour", AIAHOUR.Text}, 'AI應用時數
            {"sWNLHour", WNLHOUR.Text}, '職場續航時數
            {"sPCont", PCont.Text},
            {"sSTrainDate", STrainDate.Text},
            {"sClassification1", v_Classification1}, '1:學/2:術
            {"sComidValue", ComidValue.Value},
            {"sPTID1", v_PTID1},
            {"sPTID2", v_PTID2},
            {"sOLessonTeah1Value", vsOLessonTeah1Value},
            {"sOLessonTeah2Value", vsOLessonTeah2Value},
            {"sPTDID", ""}
        }
        Call Chk_TrainDescInput(Errmsg, htTrainDesc, dt)
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "5"
            Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js1)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If
        PName.Text = TIMS.ClearSQM(htTrainDesc("sPName"))
        PHour.Text = TIMS.ClearSQM(htTrainDesc("sPHour")) '時數
        EHOUR.Text = TIMS.ClearSQM(htTrainDesc("sEHour")) '技檢訓練時數
        AIAHOUR.Text = TIMS.ClearSQM(htTrainDesc("sAIAHour")) 'AI應用時數
        WNLHOUR.Text = TIMS.ClearSQM(htTrainDesc("sWNLHour")) '職場續航時數
        PCont.Text = TIMS.ClearSQM(htTrainDesc("sPCont"))
        STrainDate.Text = TIMS.ClearSQM(htTrainDesc("sSTrainDate"))
        v_Classification1 = TIMS.ClearSQM(htTrainDesc("sClassification1"))
        v_PTID1 = TIMS.ClearSQM(htTrainDesc("sPTID1"))
        v_PTID2 = TIMS.ClearSQM(htTrainDesc("sPTID2"))
        vsOLessonTeah1Value = TIMS.ClearSQM(htTrainDesc("sOLessonTeah1Value"))
        vsOLessonTeah2Value = TIMS.ClearSQM(htTrainDesc("sOLessonTeah2Value"))

        dr = dt.NewRow
        dt.Rows.Add(dr) '產業人才投資方案--
        dr("PTDID") = TIMS.GET_NEWPK_INT(Me, "PTDID")
        dr("TPERIOD28") = sTPERIOD28 '授課時段-7
        dr("STrainDate") = CDate(STrainDate.Text) '97年產學訓課程大綱-日期
        dr("ETrainDate") = CDate(STrainDate.Text) '97年產學訓課程大綱-日期
        dr("PName") = PName.Text '97年產學訓授課時間
        dr("PHour") = Val(PHour.Text) '時數
        dr("EHOUR") = If(EHOUR.Text <> "", Val(EHOUR.Text), Convert.DBNull) '技檢訓練時數
        dr("AIAHOUR") = If(AIAHOUR.Text <> "", Val(AIAHOUR.Text), Convert.DBNull) 'AI應用時數
        dr("WNLHOUR") = If(WNLHOUR.Text <> "", Val(WNLHOUR.Text), Convert.DBNull) '職場續航時數
        dr("PCont") = TIMS.ClearSQM(PCont.Text)
        dr("Classification1") = CInt(v_Classification1) 'CInt(Classification1.SelectedValue) '學科術科 '1:學/2:術
        '上課地點
        dr("PTID") = If(CInt(v_Classification1).Equals(1), If(v_PTID1 <> "", v_PTID1, Convert.DBNull), If(CInt(v_Classification1).Equals(2), If(v_PTID2 <> "", v_PTID2, Convert.DBNull), Convert.DBNull))
        '遠距教學 'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        dr("FARLEARN") = If(cbFARLEARN.Checked, "Y", If(Hid_DISTANCE.Value = "1", "Y", Convert.DBNull))
        '室外教學
        dr("OUTLEARN") = If(cbOUTLEARN.Checked, "Y", Convert.DBNull)
        dr("TechID") = If(vsOLessonTeah1Value <> "", vsOLessonTeah1Value, Convert.DBNull) '任課教師
        dr("TechID2") = If(vsOLessonTeah2Value <> "", vsOLessonTeah2Value, Convert.DBNull) '助教
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        Session(hid_TrainDescTable_guid1.Value) = dt
        gflag_TrainDesc_edit1 = True '啟動編輯模式
    End Sub

    ''' <summary>新增 課程大網</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '新增 課程大網 SESSION (PLAN_TRAINDESC)
        Call INSERT_TRAINDESC_SESS()
        '重建 計畫訓練內容簡介'PLAN_TRAINDESC 依SESSION
        Call CreateTrainDesc() 'PLAN_TRAINDESC

        hfScrollToAnchor.Value = "Y"
        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1) 'window.scroll(0,document.body.scrollHeight);
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "edit"
                DataGrid1.EditItemIndex = e.Item.ItemIndex
            Case "del"
                Dim dt As DataTable = Session(hid_planONCLASS_guid1.Value)
                Dim DGobj As DataGrid = DataGrid1
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                If dt.Select("POCID='" & e.CommandArgument & "'").Length <> 0 Then dt.Select("POCID='" & e.CommandArgument & "'")(0).Delete()
                Session(hid_planONCLASS_guid1.Value) = dt
                If dt.Rows.Count = 0 Then
                    DataGrid1Table.Visible = False
                Else
                    DataGrid1Table.Visible = True
                    DataGrid1.DataSource = dt
                End If
                DataGrid1.EditItemIndex = -1
            Case "save"
                Dim dt As DataTable = Nothing
                Dim dr As DataRow = Nothing
                Dim Weeks2 As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times2 As TextBox = e.Item.FindControl("Times2")
                Dim Errmsg As String = ""
                Dim v_Weeks2 As String = TIMS.GetListValue(Weeks2)
                'Dim vsWeeks2 As String = TIMS.ClearSQM(Weeks2.SelectedValue)
                Times2.Text = TIMS.ClearTimesFMT(Times2.Text)
                Dim i_Times_c_max_length As Integer = cst_i_Times_c_max_length
                Dim i_Times_c_min_length As Integer = cst_i_Times_c_min_length
                Dim s_err_msg1 As String = String.Format("上課時間／時間內容，長度超過限制範圍{0}文字長度", i_Times_c_max_length)
                Dim s_err_msg2 As String = String.Format("上課時間／時間內容，長度小於限制範圍{0}文字長度", i_Times_c_min_length)
                If Times2.Text <> "" Then
                    If Times2.Text.ToString().Length > i_Times_c_max_length Then Errmsg &= s_err_msg1 & vbCrLf
                    If Times2.Text.ToString().Length < i_Times_c_min_length Then Errmsg &= s_err_msg2 & vbCrLf
                Else
                    Errmsg &= "上課時間／時間內容，不可為空字串" & vbCrLf
                End If
                If Errmsg <> "" Then
                    If (LayerState.Value = "") Then LayerState.Value = "5"
                    Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
                    Page.RegisterStartupScript("Londing", s_js11) 'window.scroll(0,document.body.scrollHeight);
                    sm.LastErrorMessage = Errmsg
                    Exit Sub
                End If

                'PLAN_ONCLASS
                If Not Session(hid_planONCLASS_guid1.Value) Is Nothing Then
                    dt = Session(hid_planONCLASS_guid1.Value)
                    If dt.Select("POCID='" & e.CommandArgument & "'").Length <> 0 Then
                        dr = dt.Select("POCID='" & e.CommandArgument & "'")(0)
                        dr("Weeks") = v_Weeks2 'vsWeeks2 'Weeks2.SelectedValue
                        dr("Times") = Times2.Text
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                    End If
                    Session(hid_planONCLASS_guid1.Value) = dt
                    DataGrid1.EditItemIndex = -1
                End If
            Case "cancel"
                DataGrid1.EditItemIndex = -1
        End Select
        Call CreateClassTime()

        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1) 'window.scroll(0,document.body.scrollHeight);
    End Sub

    Public Shared Function Get_WeekSedIdx(ByRef oWeeks As DropDownList, ByVal value As String) As Integer
        Dim rst As Integer = 0
        For i As Integer = 0 To oWeeks.Items.Count - 1
            If (oWeeks.Items(i).Value = value) Then Return i
        Next
        Return rst
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim Flag_AddEnabled As Boolean = Button29.Enabled
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Weeks1 As Label = e.Item.FindControl("Weeks1")
                Dim Times1 As Label = e.Item.FindControl("Times1")
                'Dim labmsgTimes1 As Label = e.Item.FindControl("labmsgTimes1")
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn1 As Button = e.Item.FindControl("Button2")
                Dim btn2 As Button = e.Item.FindControl("Button3")

                btn1.Enabled = Button29.Enabled 'Flag_AddEnabled
                btn2.Enabled = Button29.Enabled 'Flag_AddEnabled
                Weeks1.Text = drv("Weeks").ToString
                Times1.Text = drv("Times").ToString
                'labmsgTimes1.Text = If(Not CheckTimesFMT(Times1.Text), "(上課時間／上課時段，格式不符合範例)", "")
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn2.CommandArgument = drv("POCID")

            Case ListItemType.EditItem
                Dim Weeks2 As DropDownList = e.Item.FindControl("Weeks2")
                Dim Times2 As TextBox = e.Item.FindControl("Times2")
                Times2.MaxLength = txtTimes.MaxLength
                Dim btn1 As Button = e.Item.FindControl("Button4")
                Dim btn2 As Button = e.Item.FindControl("Button5")

                Dim drv As DataRowView = e.Item.DataItem
                Weeks2 = TIMS.Get_ddlWeeks(Weeks2)

                Common.SetListItem(Weeks2, drv("Weeks").ToString)
                Times2.Text = drv("Times").ToString
                btn1.CommandArgument = drv("POCID")
                btn1.Enabled = Button29.Enabled 'Flag_AddEnabled
                'btn2.Enabled = Button29.Enabled 'Flag_AddEnabled

        End Select
    End Sub

    ''' <summary>修改 課程大網</summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub Datagrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid3.ItemCommand
        'If Session(hid_TrainDescTable_guid1.Value) Is Nothing Then Exit Sub
        If Not TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then Return
        gflag_TrainDesc_edit1 = True '啟動編輯模式
        Dim dt As DataTable = Session(hid_TrainDescTable_guid1.Value) '取得SESSION到 dt

        Select Case e.CommandName
            Case "edit" '編輯列 開啟
                Datagrid3.EditItemIndex = e.Item.ItemIndex
                Session(hid_TrainDescTable_guid1.Value) = dt
            Case "del" '刪除
                Dim DGobj As DataGrid = Datagrid3
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                If Convert.ToString(Datagrid3.DataKeys(e.Item.ItemIndex)) <> "" Then
                    If dt.Select("PTDID='" & Datagrid3.DataKeys(e.Item.ItemIndex) & "'").Length <> 0 Then
                        Dim dr As DataRow = dt.Select("PTDID='" & Datagrid3.DataKeys(e.Item.ItemIndex) & "'")(0)
                        dr.Delete()
                    End If
                End If
                Session(hid_TrainDescTable_guid1.Value) = dt
                Datagrid3Table.Visible = False
                If dt.Rows.Count > 0 Then
                    Datagrid3Table.Visible = True
                    Datagrid3.DataSource = dt
                End If
                Datagrid3.EditItemIndex = -1 '關閉編輯列
            Case "save" '存檔
                Dim DGobj As DataGrid = Datagrid3
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                Dim Errmsg As String = ""
                Dim TPERIOD28_1e As CheckBox = e.Item.FindControl("TPERIOD28_1e")
                Dim TPERIOD28_2e As CheckBox = e.Item.FindControl("TPERIOD28_2e")
                Dim TPERIOD28_3e As CheckBox = e.Item.FindControl("TPERIOD28_3e")
                Dim STrainDateTxt As TextBox = e.Item.FindControl("STrainDateTxt")
                Dim Img1 As HtmlImage = e.Item.FindControl("Img2")
                Dim Eddlh1 As DropDownList = e.Item.FindControl("Eddlh1")
                Dim Eddlm1 As DropDownList = e.Item.FindControl("Eddlm1")
                Dim Eddlh2 As DropDownList = e.Item.FindControl("Eddlh2")
                Dim Eddlm2 As DropDownList = e.Item.FindControl("Eddlm2")
                Dim PNameTxt As TextBox = e.Item.FindControl("PNameTxt")
                Dim PHourTxt As TextBox = e.Item.FindControl("PHourTxt") '時數
                Dim EHourTxt As TextBox = e.Item.FindControl("EHourTxt") '技檢訓練時數
                Dim AIAHOURTxt As TextBox = e.Item.FindControl("AIAHOURTxt") 'AI應用時數
                Dim WNLHOURTxt As TextBox = e.Item.FindControl("WNLHOURTxt") '職場續航時數

                Dim PContEdit As TextBox = e.Item.FindControl("PContEdit")
                Dim drpClassEdit As DropDownList = e.Item.FindControl("drpClassEdit")
                Dim drpPTIDEdit1 As DropDownList = e.Item.FindControl("drpPTIDEdit1")
                Dim drpPTIDEdit2 As DropDownList = e.Item.FindControl("drpPTIDEdit2")
                Dim cb_FARLEARNe As CheckBox = e.Item.FindControl("cb_FARLEARNe")
                Dim cb_OUTLEARNe As CheckBox = e.Item.FindControl("cb_OUTLEARNe")
                Dim Tech1ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech1ValueEdit")
                Dim Tech1Edit As TextBox = e.Item.FindControl("Tech1Edit")
                Dim Tech2ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech2ValueEdit")
                Dim Tech2Edit As TextBox = e.Item.FindControl("Tech2Edit")

                Dim sTPERIOD28 As String = ""
                sTPERIOD28 &= If(TPERIOD28_1e.Checked, "Y", "N")
                sTPERIOD28 &= If(TPERIOD28_2e.Checked, "Y", "N")
                sTPERIOD28 &= If(TPERIOD28_3e.Checked, "Y", "N")
                STDate.Text = TIMS.ClearSQM(STDate.Text)
                FDDate.Text = TIMS.ClearSQM(FDDate.Text)
                PNameTxt.Text = TIMS.ClearSQM(PNameTxt.Text)
                Dim vsddlpnH1 As String = TIMS.GetListValue(Eddlh1) '.SelectedValue
                Dim vsddlpnM1 As String = TIMS.GetListValue(Eddlm1) '.SelectedValue
                Dim vsddlpnH2 As String = TIMS.GetListValue(Eddlh2) '.SelectedValue
                Dim vsddlpnM2 As String = TIMS.GetListValue(Eddlm2) '.SelectedValue
                PHourTxt.Text = TIMS.ClearSQM(PHourTxt.Text) '時數
                EHourTxt.Text = TIMS.ClearSQM(EHourTxt.Text) '技檢訓練時數
                AIAHOURTxt.Text = TIMS.ClearSQM(AIAHOURTxt.Text) 'AI應用時數
                WNLHOURTxt.Text = TIMS.ClearSQM(WNLHOURTxt.Text) '職場續航時數
                PContEdit.Text = TIMS.ClearSQM(PContEdit.Text)
                STrainDateTxt.Text = TIMS.ClearSQM(STrainDateTxt.Text)
                Dim vsClassification1 As String = TIMS.GetListValue(drpClassEdit) '.SelectedValue
                ComidValue.Value = TIMS.ClearSQM(ComidValue.Value)
                Dim vsPTID1 As String = TIMS.GetListValue(drpPTIDEdit1) '.SelectedValue
                Dim vsPTID2 As String = TIMS.GetListValue(drpPTIDEdit2) '.SelectedValue
                Dim vsOLessonTeah1Value As String = TIMS.ClearSQM(Tech1ValueEdit.Value)
                Dim vsOLessonTeah2Value As String = TIMS.ClearSQM(Tech2ValueEdit.Value)
                Dim vsPTDID As String = TIMS.ClearSQM(e.CommandArgument)

                Dim htTrainDesc As New Hashtable From {
                    {"sTPERIOD28", sTPERIOD28},
                    {"sSTDate", STDate.Text},
                    {"sFDDate", FDDate.Text},
                    {"sPName", PNameTxt.Text},
                    {"sddlpnH1", vsddlpnH1},
                    {"sddlpnM1", vsddlpnM1},
                    {"sddlpnH2", vsddlpnH2},
                    {"sddlpnM2", vsddlpnM2},
                    {"sPHour", PHourTxt.Text}, '時數
                    {"sEHour", EHourTxt.Text}, '技檢訓練時數
                    {"sAIAHour", AIAHOURTxt.Text}, 'AI應用時數
                    {"sWNLHour", WNLHOURTxt.Text}, '職場續航時數
                    {"sPCont", PContEdit.Text},
                    {"sSTrainDate", STrainDateTxt.Text},
                    {"sClassification1", vsClassification1}, '1:學/2:術
                    {"sComidValue", ComidValue.Value},
                    {"sPTID1", vsPTID1},
                    {"sPTID2", vsPTID2},
                    {"sOLessonTeah1Value", vsOLessonTeah1Value},
                    {"sOLessonTeah2Value", vsOLessonTeah2Value},
                    {"sPTDID", vsPTDID}
                }
                Call Chk_TrainDescInput(Errmsg, htTrainDesc, dt)
                '(產業人才投資方案) 增加設定課程大綱時,場地與師資為必填  by AMU 20090901
                If Errmsg <> "" Then
                    'Page.RegisterStartupScript("Londing", "<script>Layer_change(5);</script>")
                    'Page.RegisterStartupScript("Londing3", "<script>showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');</script>")
                    'Dim prss1 As String = "Layer_change(5);showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');"
                    Dim prss1 As String = "showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');"
                    Page.RegisterStartupScript("Londing", "<script>" & prss1 & "</script>")
                    sm.LastErrorMessage = Errmsg
                    Exit Sub
                End If

                PNameTxt.Text = TIMS.ClearSQM(htTrainDesc("sPName"))
                PHourTxt.Text = TIMS.ClearSQM(htTrainDesc("sPHour")) '時數
                EHourTxt.Text = TIMS.ClearSQM(htTrainDesc("sEHour")) '技檢訓練時數
                AIAHOURTxt.Text = TIMS.ClearSQM(htTrainDesc("sAIAHour")) 'AI應用時數
                WNLHOURTxt.Text = TIMS.ClearSQM(htTrainDesc("sWNLHour")) '職場續航時數
                PContEdit.Text = TIMS.ClearSQM(htTrainDesc("sPCont"))
                STrainDateTxt.Text = TIMS.ClearSQM(htTrainDesc("sSTrainDate"))
                vsClassification1 = TIMS.ClearSQM(htTrainDesc("sClassification1"))
                vsOLessonTeah1Value = TIMS.ClearSQM(htTrainDesc("sOLessonTeah1Value"))
                vsOLessonTeah2Value = TIMS.ClearSQM(htTrainDesc("sOLessonTeah2Value"))

                Dim oPTDID As Object = e.CommandArgument
                Dim fffPTDID As String = If(IsNumeric(oPTDID), String.Concat("PTDID='", oPTDID, "'"), "")
                If IsNumeric(oPTDID) AndAlso fffPTDID <> "" AndAlso dt.Select(fffPTDID).Length <> 0 Then
                    Dim dr As DataRow = dt.Select(fffPTDID)(0)
                    dr("TPERIOD28") = sTPERIOD28
                    dr("STrainDate") = STrainDateTxt.Text
                    dr("ETrainDate") = STrainDateTxt.Text
                    dr("PName") = PNameTxt.Text
                    dr("PHour") = Val(PHourTxt.Text) '時數
                    dr("EHour") = If(EHourTxt.Text <> "", Val(EHourTxt.Text), Convert.DBNull) '技檢訓練時數
                    dr("AIAHOUR") = If(AIAHOURTxt.Text <> "", Val(AIAHOURTxt.Text), Convert.DBNull) 'AI應用時數
                    dr("WNLHOUR") = If(WNLHOURTxt.Text <> "", Val(WNLHOURTxt.Text), Convert.DBNull) '職場續航時數
                    dr("PCont") = TIMS.ClearSQM(PContEdit.Text)
                    dr("Classification1") = CInt(vsClassification1) '學科術科
                    '上課地點
                    dr("PTID") = If(CInt(vsClassification1).Equals(1), If(vsPTID1 <> "", vsPTID1, Convert.DBNull), If(CInt(vsClassification1).Equals(2), If(vsPTID2 <> "", vsPTID2, Convert.DBNull), Convert.DBNull))
                    dr("FARLEARN") = If(cb_FARLEARNe.Checked, "Y", Convert.DBNull) '遠距教學
                    dr("OUTLEARN") = If(cb_OUTLEARNe.Checked, "Y", Convert.DBNull) '室外教學
                    dr("TechID") = If(vsOLessonTeah1Value <> "", vsOLessonTeah1Value, Convert.DBNull) '任課教師
                    dr("TechID2") = If(vsOLessonTeah2Value <> "", vsOLessonTeah2Value, Convert.DBNull) '助教
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                Session(hid_TrainDescTable_guid1.Value) = dt
                Datagrid3.EditItemIndex = -1
            Case "cancel"
                Datagrid3.EditItemIndex = -1
        End Select
        'gflag_TrainDesc_edit1 = True '啟動編輯模式

        Call CreateTrainDesc() 'PLAN_TRAINDESC

        hfScrollToAnchor.Value = "Y"
        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Dim Flag_AddEnabled As Boolean = Button1.Enabled
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        ''--System.Web.UI.HtmlControls.HtmlGenericControl
        'Sp_lb_AIAHOUR_WNLHOUR.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        'Sp_AIAHOUR_WNLHOURLabel.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
        'Sp_AIAHOUR_WNLHOURTxt.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)


        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Sp_lb_AIAHOUR_WNLHOUR As HtmlGenericControl = e.Item.FindControl("Sp_lb_AIAHOUR_WNLHOUR")
                Sp_lb_AIAHOUR_WNLHOUR.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Sp_AIAHOUR_WNLHOURLabel As HtmlGenericControl = e.Item.FindControl("Sp_AIAHOUR_WNLHOURLabel")
                Sp_AIAHOUR_WNLHOURLabel.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
                Dim TPERIOD28_1t As CheckBox = e.Item.FindControl("TPERIOD28_1t")
                Dim TPERIOD28_2t As CheckBox = e.Item.FindControl("TPERIOD28_2t")
                Dim TPERIOD28_3t As CheckBox = e.Item.FindControl("TPERIOD28_3t")
                Dim STrainDateLabel As Label = e.Item.FindControl("STrainDateLabel")
                Dim PNameLabel As Label = e.Item.FindControl("PNameLabel")
                Dim PHourLabel As Label = e.Item.FindControl("PHourLabel") '時數
                Dim EHourLabel As Label = e.Item.FindControl("EHourLabel") '技檢訓練時數
                Dim AIAHOURLabel As Label = e.Item.FindControl("AIAHOURLabel") 'AI應用時數
                Dim WNLHOURLabel As Label = e.Item.FindControl("WNLHOURLabel") '職場續航時數
                Dim PContText As TextBox = e.Item.FindControl("PContText")
                Dim drpClassification1 As DropDownList = e.Item.FindControl("drpClassification1") '1:學/2:術
                Dim drpPTID As DropDownList = e.Item.FindControl("drpPTID")
                Dim cb_FARLEARNi As CheckBox = e.Item.FindControl("cb_FARLEARNi")
                Dim cb_OUTLEARNi As CheckBox = e.Item.FindControl("cb_OUTLEARNi") '室外教學
                Dim Tech1Value As HtmlInputHidden = e.Item.FindControl("Tech1Value")
                Dim Tech1Text As TextBox = e.Item.FindControl("Tech1Text")
                Dim Tech2Value As HtmlInputHidden = e.Item.FindControl("Tech2Value")
                Dim Tech2Text As TextBox = e.Item.FindControl("Tech2Text")
                Dim btn1 As Button = e.Item.FindControl("Button6") 'edit
                Dim btn2 As Button = e.Item.FindControl("Button7") 'del
                Dim t_PTDID As String = "PTDID:" & Convert.ToString(drv("PTDID"))
                TIMS.Tooltip(btn1, t_PTDID, True)
                TIMS.Tooltip(btn2, t_PTDID, True)

                Dim str_v_TPERIOD28 As String = cst_NNN '"NNN"
                If $"{drv("TPERIOD28")}" <> "" AndAlso $"{drv("TPERIOD28")}".Length >= 3 Then str_v_TPERIOD28 = $"{drv("TPERIOD28")}"
                TPERIOD28_1t.Checked = If(str_v_TPERIOD28.Substring(0, 1) = "Y", True, False)
                TPERIOD28_2t.Checked = If(str_v_TPERIOD28.Substring(1, 1) = "Y", True, False)
                TPERIOD28_3t.Checked = If(str_v_TPERIOD28.Substring(2, 1) = "Y", True, False)

                If $"{drv("STrainDate")}" <> "" Then STrainDateLabel.Text = TIMS.Cdate3(drv("STrainDate"))
                PNameLabel.Text = $"{drv("PName")}" '時間
                PHourLabel.Text = $"{drv("PHour")}" '時數
                EHourLabel.Text = $"{drv("EHour")}" '技檢訓練時數
                AIAHOURLabel.Text = $"{drv("AIAHOUR")}" 'AI應用時數
                WNLHOURLabel.Text = $"{drv("WNLHOUR")}" '職場續航時數
                TIMS.Tooltip(EHourLabel, cst_EHour_t1, True) '技檢訓練時數
                TIMS.Tooltip(AIAHOURLabel, cst_AIAHOUR_t1, True) 'AI應用時數
                TIMS.Tooltip(WNLHOURLabel, cst_WNLHOUR_t1, True) '職場續航時數

                PContText.Text = $"{drv("PCont")}" '內容
                PContText.Text = TIMS.HtmlDecode1(PContText.Text)
                If drv("Classification1").ToString <> "" Then '1:學/2:術
                    Common.SetListItem(drpClassification1, drv("Classification1").ToString)
                    Dim v_drpClassification1 As String = TIMS.GetListValue(drpClassification1)
                    Select Case v_drpClassification1'drpClassification1.SelectedValue
                        Case "1"  '學科
                            '將Hid_ComIDNO.Value 塞入有效值
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" AndAlso sm.UserInfo.LID = 0 Then Hid_ComIDNO.Value = ComidValue.Value
                            If Hid_ComIDNO.Value = "" AndAlso sm.UserInfo.LID = 1 Then Hid_ComIDNO.Value = ComidValue.Value
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                            drpPTID = TIMS.Get_SciPTID(drpPTID, Hid_ComIDNO.Value, 1, objconn)
                        Case "2"  '術科
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "ComIDNO", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" AndAlso sm.UserInfo.LID = 0 Then Hid_ComIDNO.Value = ComidValue.Value
                            If Hid_ComIDNO.Value = "" AndAlso sm.UserInfo.LID = 1 Then Hid_ComIDNO.Value = ComidValue.Value
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                            drpPTID = TIMS.Get_TechPTID(drpPTID, Hid_ComIDNO.Value, 1, objconn)
                    End Select
                    If drv("PTID").ToString <> "" Then Common.SetListItem(drpPTID, drv("PTID").ToString)
                End If
                '遠距教學 '署的權限可以修改遠距教學
                cb_FARLEARNi.Enabled = False
                If Not gflag_DISTANCE_can_updata Then TIMS.Tooltip(cb_FARLEARNi, "(不提供該選項)", True)
                If gflag_DISTANCE_can_updata Then TIMS.Tooltip(cb_FARLEARNi, "業務權限可修改遠距教學", True)
                '產投使用／遠距教學 暫不啟用
                If cb_FARLEARNi IsNot Nothing Then
                    If flag_StopDISTANCE2 Then cb_FARLEARNi.Visible = False
                    cb_FARLEARNi.Checked = If(Convert.ToString(drv("FARLEARN")).Equals("Y"), True, False)
                End If
                '室外教學
                cb_OUTLEARNi.Checked = If(Convert.ToString(drv("OUTLEARN")).Equals("Y"), True, False)

                If Convert.ToString(drv("TechID")) <> "" Then
                    Tech1Value.Value = drv("TechID").ToString
                    Tech1Text.Text = TIMS.Get_TeachCName(Tech1Value.Value, objconn) '
                End If
                If Convert.ToString(drv("TechID2")) <> "" Then
                    Tech2Value.Value = drv("TechID2").ToString
                    Tech2Text.Text = TIMS.Get_TeachCName(Tech2Value.Value, objconn) '
                End If
                btn2.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btn2.CommandArgument = drv("PTDID").ToString
                btn1.Enabled = Flag_AddEnabled
                btn2.Enabled = Flag_AddEnabled

            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Sp_AIAHOUR_WNLHOURTxt As HtmlGenericControl = e.Item.FindControl("Sp_AIAHOUR_WNLHOURTxt")
                Sp_AIAHOUR_WNLHOURTxt.Visible = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1, True, False)
                Dim TPERIOD28_1e As CheckBox = e.Item.FindControl("TPERIOD28_1e")
                Dim TPERIOD28_2e As CheckBox = e.Item.FindControl("TPERIOD28_2e")
                Dim TPERIOD28_3e As CheckBox = e.Item.FindControl("TPERIOD28_3e")
                Dim STrainDateTxt As TextBox = e.Item.FindControl("STrainDateTxt")
                Dim Img1 As HtmlImage = e.Item.FindControl("Img2")
                Dim Eddlh1 As DropDownList = e.Item.FindControl("Eddlh1")
                Dim Eddlm1 As DropDownList = e.Item.FindControl("Eddlm1")
                Dim Eddlh2 As DropDownList = e.Item.FindControl("Eddlh2")
                Dim Eddlm2 As DropDownList = e.Item.FindControl("Eddlm2")
                Dim PNameTxt As TextBox = e.Item.FindControl("PNameTxt")
                Dim PHourTxt As TextBox = e.Item.FindControl("PHourTxt") '時數
                Dim EHourTxt As TextBox = e.Item.FindControl("EHourTxt") '技檢訓練時數
                Dim AIAHOURTxt As TextBox = e.Item.FindControl("AIAHOURTxt") 'AI應用時數
                Dim WNLHOURTxt As TextBox = e.Item.FindControl("WNLHOURTxt") '職場續航時數
                Dim PContEdit As TextBox = e.Item.FindControl("PContEdit")
                Dim drpClassEdit As DropDownList = e.Item.FindControl("drpClassEdit")
                Dim drpPTIDEdit1 As DropDownList = e.Item.FindControl("drpPTIDEdit1")
                Dim drpPTIDEdit2 As DropDownList = e.Item.FindControl("drpPTIDEdit2")
                Dim cb_FARLEARNe As CheckBox = e.Item.FindControl("cb_FARLEARNe")
                Dim cb_OUTLEARNe As CheckBox = e.Item.FindControl("cb_OUTLEARNe") '室外教學
                Dim Tech1ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech1ValueEdit")
                Dim Tech1Edit As TextBox = e.Item.FindControl("Tech1Edit")
                Dim Tech2ValueEdit As HtmlInputHidden = e.Item.FindControl("Tech2ValueEdit")
                Dim Tech2Edit As TextBox = e.Item.FindControl("Tech2Edit")
                Dim btn3 As Button = e.Item.FindControl("Button10") 'save
                Dim btn4 As Button = e.Item.FindControl("Button11") 'cancel
                Dim t_PTDID As String = "PTDID:" & Convert.ToString(drv("PTDID"))
                TIMS.Tooltip(btn3, t_PTDID)
                TIMS.Tooltip(btn4, t_PTDID)

                Dim str_v_TPERIOD28 As String = cst_NNN '"NNN"
                If Convert.ToString(drv("TPERIOD28")) <> "" AndAlso Convert.ToString(drv("TPERIOD28")).Length >= 3 Then str_v_TPERIOD28 = Convert.ToString(drv("TPERIOD28"))
                TPERIOD28_1e.Checked = If(str_v_TPERIOD28.Substring(0, 1) = "Y", True, False)
                TPERIOD28_2e.Checked = If(str_v_TPERIOD28.Substring(1, 1) = "Y", True, False)
                TPERIOD28_3e.Checked = If(str_v_TPERIOD28.Substring(2, 1) = "Y", True, False)

                Call CreateTimesItem(Eddlh1, Eddlh2, Eddlm1, Eddlm2)
                Img1.Attributes("onclick") = "return chkTrainDate('" & STrainDateTxt.ClientID & "');"

                '任課教師
                Tech1Edit.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addx','" & Tech1Edit.ClientID & "','" & Tech1ValueEdit.ClientID & "');")
                Tech1Edit.Attributes("onchange") = "GetTeacherId(this.value,'" & Tech1ValueEdit.ClientID & "','" & Tech1Edit.ClientID & "');"
                Tech1Edit.Style.Item("CURSOR") = "hand"

                '助教
                Tech2Edit.Attributes.Add("onDblClick", "javascript:LessonTeah1('Addy','" & Tech2Edit.ClientID & "','" & Tech2ValueEdit.ClientID & "');")
                Tech2Edit.Attributes("onchange") = "GetTeacherId(this.value,'" & Tech2ValueEdit.ClientID & "','" & Tech2Edit.ClientID & "');"
                Tech2Edit.Style.Item("CURSOR") = "hand"

                If drv("STrainDate").ToString <> "" Then STrainDateTxt.Text = TIMS.Cdate3(drv("STrainDate"))

                PNameTxt.Text = drv("PName").ToString
                If PNameTxt.Text <> "" Then
                    Try
                        PNameTxt.Text = Replace(PNameTxt.Text, "：", ":")
                        PNameTxt.Text = Replace(PNameTxt.Text, "-", "~")
                        PNameTxt.Text = TIMS.ChangeIDNO(PNameTxt.Text)
                        Dim hm1hm2 As String() = Convert.ToString(PNameTxt.Text).Split("~")
                        Dim hm1 As String()
                        Dim hm2 As String()

                        If hm1hm2.Length > 1 Then
                            If hm1hm2(0).IndexOf(":") > -1 Then
                                hm1 = hm1hm2(0).Split(":")
                                hm2 = hm1hm2(1).Split(":")
                                If hm1.Length > 1 Then
                                    Common.SetListItem(Eddlh1, Convert.ToString(hm1(0)))
                                    Common.SetListItem(Eddlm1, Convert.ToString(hm1(1)))
                                End If
                                If hm2.Length > 1 Then
                                    Common.SetListItem(Eddlh2, Convert.ToString(hm2(0)))
                                    Common.SetListItem(Eddlm2, Convert.ToString(hm2(1)))
                                End If
                            Else
                                If Convert.ToString(hm1hm2(0)).Length = 4 AndAlso IsNumeric(hm1hm2(0)) Then
                                    Common.SetListItem(Eddlh1, Convert.ToString(hm1hm2(0).Substring(0, 2)))
                                    Common.SetListItem(Eddlm1, Convert.ToString(hm1hm2(0).Substring(2, 2)))
                                End If
                                If Convert.ToString(hm1hm2(0)).Length = 4 AndAlso IsNumeric(hm1hm2(1)) Then
                                    Common.SetListItem(Eddlh2, Convert.ToString(hm1hm2(1).Substring(0, 2)))
                                    Common.SetListItem(Eddlm2, Convert.ToString(hm1hm2(1).Substring(2, 2)))
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                        Dim strErrmsg As String = ""
                        strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                        strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                        strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                        strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                        Call TIMS.WriteTraceLog(strErrmsg)
                    End Try
                End If

                PHourTxt.Text = TIMS.ClearSQM(drv("PHour")) '時數
                EHourTxt.Text = TIMS.ClearSQM(drv("EHour")) '技檢訓練時數
                AIAHOURTxt.Text = TIMS.ClearSQM(drv("AIAHOUR")) 'AI應用時數
                WNLHOURTxt.Text = TIMS.ClearSQM(drv("WNLHOUR")) '職場續航時數
                TIMS.Tooltip(EHourTxt, cst_EHour_t1, True) '技檢訓練時數
                TIMS.Tooltip(AIAHOURTxt, cst_AIAHOUR_t1, True) 'AI應用時數
                TIMS.Tooltip(WNLHOURTxt, cst_WNLHOUR_t1, True) '職場續航時數

                PContEdit.Text = Convert.ToString(drv("PCont"))
                PContEdit.Text = TIMS.ClearSQM(PContEdit.Text)
                'drpClassEdit.Attributes.Add("onchange", "javascript:showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');Layer_change(5);")
                drpClassEdit.Attributes.Add("onchange", "javascript:showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');")
                drpPTIDEdit1 = GetPTID(drpPTIDEdit1, 1)
                drpPTIDEdit2 = GetPTID(drpPTIDEdit2, 2)
                Common.SetListItem(drpClassEdit, drv("Classification1").ToString) '1:學/2:術
                Dim v_drpClassEdit As String = TIMS.GetListValue(drpClassEdit)
                Select Case v_drpClassEdit'drpClassEdit.SelectedValue
                    Case "1"
                        If drv("PTID").ToString <> "" Then
                            Common.SetListItem(drpPTIDEdit1, drv("PTID").ToString)
                        End If
                    Case "2"
                        If drv("PTID").ToString <> "" Then
                            Common.SetListItem(drpPTIDEdit2, drv("PTID").ToString)
                        End If
                End Select
                Page.RegisterStartupScript("Londing3", "<script>showPTID('" & drpClassEdit.ClientID & "','" & drpPTIDEdit1.ClientID & "','" & drpPTIDEdit2.ClientID & "');</script>")

                '遠距教學 '署的權限可以修改遠距教學
                cb_FARLEARNe.Enabled = If(gflag_DISTANCE_can_updata, True, False) 'False
                If Not gflag_DISTANCE_can_updata Then TIMS.Tooltip(cb_FARLEARNe, "(不提供該選項)", True)
                If gflag_DISTANCE_can_updata Then TIMS.Tooltip(cb_FARLEARNe, "業務權限可修改遠距教學", True)
                '產投使用／遠距教學 暫不啟用
                If cb_FARLEARNe IsNot Nothing Then
                    If flag_StopDISTANCE2 Then cb_FARLEARNe.Visible = False
                    cb_FARLEARNe.Checked = If(Convert.ToString(drv("FARLEARN")).Equals("Y"), True, False)
                End If
                '室外教學
                cb_OUTLEARNe.Checked = If(Convert.ToString(drv("OUTLEARN")).Equals("Y"), True, False)

                Tech1ValueEdit.Value = ""
                Tech1Edit.Text = ""
                If Convert.ToString(drv("TechID")) <> "" Then
                    Tech1ValueEdit.Value = drv("TechID").ToString
                    Tech1Edit.Text = TIMS.Get_TeachCName(Tech1ValueEdit.Value, objconn) '
                End If
                Tech2ValueEdit.Value = ""
                Tech2Edit.Text = ""
                If Convert.ToString(drv("TechID2")) <> "" Then
                    Tech2ValueEdit.Value = drv("TechID2").ToString
                    Tech2Edit.Text = TIMS.Get_TeachCName(Tech2ValueEdit.Value, objconn) '
                End If

                btn3.CommandArgument = drv("PTDID").ToString
                btn3.Enabled = Flag_AddEnabled
                'btn4.Enabled = Flag_AddEnabled
        End Select
    End Sub

    Private Sub Classification1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Classification1.SelectedIndexChanged
        PTID1.Items.Clear()
        PTID2.Items.Clear()

        If center.Text = "" OrElse ComidValue.Value = "" Then
            Common.RespWrite(Me, "<Script>alert('請先選擇【訓練機構】');</Script>")
            If (LayerState.Value = "") Then LayerState.Value = "5"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            Return ' Exit Sub
        End If

        Dim v_Classification1 As String = TIMS.GetListValue(Classification1)
        Select Case v_Classification1'Classification1.SelectedValue '1:學/2:術
            Case "1"
                PTID1 = GetPTID(PTID1, 1)
            Case "2"
                PTID2 = GetPTID(PTID2, 2)
            Case Else
                Dim sPlaceNAME As String = "(請先選擇學／術科)"
                PTID1.Items.Add(New ListItem(sPlaceNAME, ""))
                PTID2.Items.Add(New ListItem(sPlaceNAME, ""))
        End Select

        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    ''' <summary>getPTID 訓練地點 取得</summary>
    ''' <param name="obj"></param>
    ''' <param name="iType "></param>
    ''' <returns></returns>
    Function GetPTID(ByVal obj As ListControl, ByVal iType As Integer) As ListControl
        'Dim tempdt As DataTable
        'Dim i As Integer
        'iType :1:學科/2:術科
        Dim drAry() As DataRow = Nothing
        Dim v_SciPlaceID As String = TIMS.GetListValue(SciPlaceID)
        Dim v_SciPlaceID2 As String = TIMS.GetListValue(SciPlaceID2)
        Dim v_TechPlaceID As String = TIMS.GetListValue(TechPlaceID)
        Dim v_TechPlaceID2 As String = TIMS.GetListValue(TechPlaceID2)

        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_SciPlaceID, 1, 1, objconn)
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_TechPlaceID, 2, 2, objconn)
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_SciPlaceID2, 3, 1, objconn)
        TIMS.GetTaddresstable(sm, ViewState("dtTaddress"), ComidValue.Value, v_TechPlaceID2, 4, 2, objconn)
        Dim tempdt As DataTable = ViewState("dtTaddress")
        obj.Items.Clear()
        If tempdt Is Nothing Then Return obj
        If iType = 1 Then     '學科
            If tempdt.Select("PID IN (1,3)").Length > 0 Then drAry = tempdt.Select("PID IN (1,3)")
        ElseIf iType = 2 Then '術科
            If tempdt.Select("PID IN (2,4)").Length > 0 Then drAry = tempdt.Select("PID IN (2,4)")
        End If
        If drAry Is Nothing Then Return obj
        Dim i_Addrow As Integer = 0
        For i As Integer = 0 To drAry.Length - 1
            'obj.Items.Insert(i, New ListItem(dr(i)("PlaceNAME"), dr(i)("PTID")))
            Dim sPlaceNAME As String = TIMS.ClearSQM(drAry(i)("PlaceNAME"))
            Dim sPTID As String = TIMS.ClearSQM(drAry(i)("PTID"))
            If sPlaceNAME <> "" AndAlso sPTID <> "" Then
                i_Addrow += 1
                obj.Items.Add(New ListItem(sPlaceNAME, sPTID))
            End If
        Next
        If i_Addrow = 0 Then
            Dim sPlaceNAME As String = String.Format("(請先選擇{0}場地)", If(iType = 1, "學科", "術科"))
            obj.Items.Add(New ListItem(sPlaceNAME, ""))
        End If
        Return obj
    End Function

    Public Shared Function GetNewtable() As DataTable
        Dim dtSpace As New DataTable
        Dim dr As DataRow = Nothing
        'ViewState("dtTaddress") = Nothing
        dtSpace.Columns.Add("PID")
        dtSpace.Columns.Add("PlaceID")
        dtSpace.Columns.Add("Name")
        dtSpace.Columns.Add("classification")
        dtSpace.Columns.Add("PTID")
        dtSpace.Columns.Add("PlaceNAME")
        For i As Integer = 0 To 4
            dr = dtSpace.NewRow()
            dr("PID") = i
            dr("PlaceID") = ""
            dr("Name") = If(i = 0, TIMS.cst_ddl_PlsChos6, "")
            dr("classification") = ""
            dr("PTID") = ""
            dr("PlaceNAME") = ""
            dtSpace.Rows.Add(dr)
        Next
        'ViewState("dtTaddress") = dtSpace
        Return dtSpace
    End Function

    '(新增)PLAN_BUSPACKAGE-計畫包班事業單位
    Private Sub BtnAddBusPackage_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddBusPackage.Click
        Const Cst_PKName As String = "BPID"
        '錯誤檢查
        Dim Errmsg As String = ""
        txtUname.Text = TIMS.ClearSQM(txtUname.Text)
        txtIntaxno.Text = TIMS.ClearSQM(txtIntaxno.Text)
        If txtUname.Text = "" Then
            txtUname.Text = ""
            Errmsg &= "企業名稱，不可為空" & vbCrLf
        Else
            If txtUname.Text.ToString.Length > 50 Then Errmsg &= "企業名稱，長度超過限制範圍50文字長度" & vbCrLf  '錯誤檢查
        End If
        If txtIntaxno.Text <> "" Then
            If Not TIMS.CheckIsECFA(TIMS.ChangeIDNO(txtIntaxno.Text), objconn) Then Errmsg &= "「" & Convert.ToString(txtUname.Text) & "」該企業單位統一編號 不屬於ECFA名單之企業，請重新填寫!!" & vbCrLf '未填寫 ECFA包班事業單位資料
        Else
            txtIntaxno.Text = ""
            Errmsg &= "服務單位統一編號，不可為空" & vbCrLf
        End If

        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "5"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11) 'window.scroll(0,document.body.scrollHeight);
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        '錯誤檢查
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        If Session(hid_PLAN_BUSPACKAGE_guid1.Value) Is Nothing Then Call CreateBusPackage()
        dt = Session(hid_PLAN_BUSPACKAGE_guid1.Value)
        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr(Cst_PKName) = TIMS.GET_NEWPK_INT(Me, Cst_PKName)
        dr("Uname") = Convert.ToString(txtUname.Text)
        dr("Intaxno") = TIMS.ChangeIDNO(txtIntaxno.Text)
        dr("Ubno") = TIMS.ChangeIDNO(txtUbno.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(hid_PLAN_BUSPACKAGE_guid1.Value) = dt
        Call CreateBusPackage()
        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1) 'window.scroll(0,document.body.scrollHeight);
    End Sub

    Private Sub Datagrid4_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles Datagrid4.ItemCommand
        If e.CommandArgument = "" Then Exit Sub
        Const Cst_PKName As String = "BPID"
        Dim objTable As HtmlTable = CType(Datagrid4Table, HtmlTable)
        Select Case e.CommandName
            Case "xedit"
                source.EditItemIndex = e.Item.ItemIndex
            Case "xdel"
                Dim dt As DataTable = Session(hid_PLAN_BUSPACKAGE_guid1.Value)
                Dim DGobj As DataGrid = Datagrid4
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                If dt.Select(Cst_PKName & "='" & e.CommandArgument & "'").Length <> 0 Then dt.Select(Cst_PKName & "='" & e.CommandArgument & "'")(0).Delete()
                Session(hid_PLAN_BUSPACKAGE_guid1.Value) = dt
                objTable.Visible = False
                If dt.Rows.Count > 0 Then
                    objTable.Visible = True
                    source.DataSource = dt
                End If
                source.EditItemIndex = -1
            Case "xsave"
                Dim okflag As Boolean = True
                Dim tUName As TextBox = e.Item.FindControl("ttxtUName")
                Dim tIntaxno As TextBox = e.Item.FindControl("ttxtIntaxno")
                Dim tUbno As TextBox = e.Item.FindControl("ttxtUbno")
                If Session(hid_PLAN_BUSPACKAGE_guid1.Value) Is Nothing Then okflag = False
                If tUName Is Nothing Then okflag = False
                If tIntaxno Is Nothing Then okflag = False
                If tUbno Is Nothing Then okflag = False
                If Not okflag Then Exit Sub '異常離開
                Dim dt As DataTable = Session(hid_PLAN_BUSPACKAGE_guid1.Value)
                If dt Is Nothing Then okflag = False
                If Not okflag Then Exit Sub '異常離開
                If dt.Select(Cst_PKName & "='" & e.CommandArgument & "'").Length <> 0 Then
                    tUName.Text = TIMS.ClearSQM(tUName.Text)
                    tIntaxno.Text = TIMS.ClearSQM(tIntaxno.Text)
                    tUbno.Text = TIMS.ClearSQM(tUbno.Text)
                    Dim dr As DataRow = dt.Select(Cst_PKName & "='" & e.CommandArgument & "'")(0)
                    dr("Uname") = Convert.ToString(tUName.Text)
                    dr("Intaxno") = TIMS.ChangeIDNO(tIntaxno.Text)
                    dr("Ubno") = TIMS.ChangeIDNO(tUbno.Text)
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
                Session(hid_PLAN_BUSPACKAGE_guid1.Value) = dt
                source.EditItemIndex = -1
            Case "xcancel"
                source.EditItemIndex = -1
        End Select
        Call CreateBusPackage()
        If (LayerState.Value = "") Then LayerState.Value = "5"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1) 'window.scroll(0,document.body.scrollHeight);
    End Sub

    Private Sub Datagrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        Const Cst_PKName As String = "BPID"
        Dim strSechObjID As String = "" '查詢按鈕ID
        Dim strAddsObjID As String = "" '維護按鈕ID
        Dim strPrntObjID As String = "" '列印按鈕ID

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim slsbUname As Label = e.Item.FindControl("slsbUname")
                Dim slabIntaxno As Label = e.Item.FindControl("slabIntaxno")
                Dim slabUbno As Label = e.Item.FindControl("slabUbno")
                Dim Button17 As Button = e.Item.FindControl("Button17") '修改
                Dim Button18 As Button = e.Item.FindControl("Button18") '刪除
                Button17.Enabled = btnAddBusPackage.Enabled
                Button18.Enabled = btnAddBusPackage.Enabled
                slsbUname.Text = drv("Uname").ToString
                slabIntaxno.Text = drv("Intaxno").ToString
                slabUbno.Text = drv("Ubno").ToString
                Button17.CommandArgument = drv(Cst_PKName)
                Button18.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                Button18.CommandArgument = drv(Cst_PKName)

            Case ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim ttxtUname As TextBox = e.Item.FindControl("ttxtUname")
                Dim ttxtIntaxno As TextBox = e.Item.FindControl("ttxtIntaxno")
                Dim ttxtUbno As TextBox = e.Item.FindControl("ttxtUbno")
                Dim Button19 As Button = e.Item.FindControl("Button19") '儲存
                Dim Button20 As Button = e.Item.FindControl("Button20") '取消
                ttxtUname.Text = Convert.ToString(drv("Uname"))
                ttxtIntaxno.Text = Convert.ToString(drv("Intaxno"))
                ttxtUbno.Text = Convert.ToString(drv("Ubno"))
                Button19.Enabled = btnAddBusPackage.Enabled
                Button20.Enabled = btnAddBusPackage.Enabled
                Button19.CommandArgument = drv(Cst_PKName)
                Button20.CommandArgument = drv(Cst_PKName)

        End Select
    End Sub

    '匯出文字 一人份材料明細
    Function InputCost6(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【一人份材料明細】"

        If TIMS.IS_DataTable(Session(hid_PersonCostTable_guid1.Value)) Then
            Dim dt As DataTable = Session(hid_PersonCostTable_guid1.Value)
            Dim subtotal As Integer = 0
            'subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then i += 1
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If

                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    'PIXOT原子筆（0.5mm藍）：單價10元╳1支╳30人＝300元
                    Dim iPerCount As Integer = If(Convert.ToString(dr("PerCount")) <> "", Val(dr("PerCount")), 0)
                    Dim iTNum As Integer = If(Convert.ToString(dr("TNum")) <> "", Val(dr("TNum")), 0)   '取得外部資料
                    Dim tmpStr As String = ""
                    tmpStr = ""
                    tmpStr &= Convert.ToString(dr("CName"))
                    tmpStr &= "(" & Convert.ToString(dr("Standard")) & ")："
                    tmpStr &= Convert.ToString(iPerCount) & " " & Convert.ToString(dr("Unit"))
                    tmpStr &= "╳" & Convert.ToString(iTNum) & "人"
                    rNote += tmpStr & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = tmpStr
                Next
            End If
        End If
        Return rst
    End Function

    '匯出文字 共同材料明細
    Function InputCost7(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【共同材料明細】"
        If TIMS.IS_DataTable(Session(hid_CommonCostTable_guid1.Value)) Then
            Dim dt As DataTable = Session(hid_CommonCostTable_guid1.Value)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then i += 1
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If

                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    If Not dr.RowState = DataRowState.Deleted Then
                        Dim iAllCount As Integer = If(Convert.ToString(dr("AllCount")) <> "", Val(dr("AllCount")), 0)
                        Dim tmpStr As String = ""
                        'tmpStr = ""
                        tmpStr &= Convert.ToString(dr("CName"))
                        tmpStr &= "(" & Convert.ToString(dr("Standard")) & ")："
                        tmpStr &= Convert.ToString(iAllCount) & " " & Convert.ToString(dr("Unit"))
                        rNote += tmpStr & vbCrLf
                        dr3 = dt3.NewRow
                        dt3.Rows.Add(dr3)
                        dr3("str1") = tmpStr
                    End If
                Next
            End If
        End If
        Return rst
    End Function

    '匯出文字 教材明細
    Function InputCost8(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【教材明細】"
        If TIMS.IS_DataTable(Session(hid_SheetCostTable_guid1.Value)) Then
            Dim dt As DataTable = Session(hid_SheetCostTable_guid1.Value)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then i += 1
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If
                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    If Not dr.RowState = DataRowState.Deleted Then
                        Dim iAllCount As Integer = If(Convert.ToString(dr("AllCount")) <> "", Val(dr("AllCount")), 0)
                        Dim tmpStr As String = ""
                        tmpStr = ""
                        tmpStr &= Convert.ToString(dr("CName"))
                        tmpStr &= "(" & Convert.ToString(dr("Standards")) & ")："
                        tmpStr &= Convert.ToString(iAllCount) & " " & Convert.ToString(dr("Unit"))
                        rNote += tmpStr & vbCrLf
                        dr3 = dt3.NewRow
                        dt3.Rows.Add(dr3)
                        dr3("str1") = tmpStr
                    End If
                Next
            End If
        End If
        Return rst
    End Function

    '匯出文字 其他明細
    Function InputCost9(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        Const cst_t1 As String = "【其他明細】"
        If TIMS.IS_DataTable(Session(hid_OtherCostTable_guid1.Value)) Then
            Dim dt As DataTable = Session(hid_OtherCostTable_guid1.Value)
            Dim subtotal As Integer = 0
            subtotal = 0
            If dt.Rows.Count > 0 Then
                Dim i As Integer = 0
                For Each dr As DataRow In dt.Rows
                    If Not dr.RowState = DataRowState.Deleted Then i += 1
                Next
                If i > 0 Then
                    rNote = cst_t1 & vbCrLf
                    dr3 = dt3.NewRow
                    dt3.Rows.Add(dr3)
                    dr3("str1") = cst_t1
                End If
                For Each dr As DataRow In dt.Select(Nothing, "ItemNo", DataViewRowState.CurrentRows)
                    If Not dr.RowState = DataRowState.Deleted Then
                        Dim iAllCount As Integer = If(Convert.ToString(dr("AllCount")) <> "", Val(dr("AllCount")), 0)
                        Dim tmpStr As String = ""
                        tmpStr = ""
                        tmpStr &= Convert.ToString(dr("CName"))
                        tmpStr &= "(" & Convert.ToString(dr("Standards")) & ")："
                        tmpStr &= Convert.ToString(iAllCount) & " " & Convert.ToString(dr("Unit"))
                        rNote += tmpStr & vbCrLf
                        dr3 = dt3.NewRow
                        dt3.Rows.Add(dr3)
                        dr3("str1") = tmpStr
                    End If
                Next
            End If
        End If
        Return rst
    End Function

    '匯出文字 其他說明
    Function InputNote2(ByRef rNote As String, ByRef dt3 As DataTable) As Boolean
        'dt3為資料主軸
        Dim rst As Boolean = True 'ok為True
        rNote = ""
        Dim dr3 As DataRow = Nothing
        '加入抬頭
        Const cst_t1 As String = "【其他說明】"
        If Trim(tNote2.Text) <> "" Then
            rNote = cst_t1 & vbCrLf
            dr3 = dt3.NewRow
            dt3.Rows.Add(dr3)
            dr3("str1") = cst_t1

            Dim tmpStr As String = ""
            tmpStr = ""
            tmpStr &= tNote2.Text
            rNote += tmpStr & vbCrLf
            dr3 = dt3.NewRow
            dt3.Rows.Add(dr3)
            dr3("str1") = tmpStr
        End If
        Return rst
    End Function

    '修正 Note中的文字 (匯出文字)
    Function ChangNoteText(ByRef tmpNoteDt As DataTable) As Boolean
        Dim rst As Boolean = True '正常/false:異常
        'tmpNoteDt = Nothing
        tmpNoteDt = New DataTable
        tmpNoteDt.Columns.Add(New DataColumn("str1"))

        Dim tmpNote As String = ""
        Note.Text = ""
        Labmsg3.Text = Cst_msgother3
        If FIXSUMCOST.Text <> "" Then Note.Text &= String.Format("固定費用總額：{0}元", FIXSUMCOST.Text) & vbCrLf
        If ACTHUMCOST.Text <> "" Then Note.Text &= String.Format("固定費用單一人時成本：{0}元", ACTHUMCOST.Text) & vbCrLf
        If FIXExceeDesc.Text <> "" Then Note.Text &= "超出人時成本原因說明：" & FIXExceeDesc.Text & vbCrLf
        If METSUMCOST.Text <> "" Then Note.Text &= String.Format("材料費用總額：{0}元", METSUMCOST.Text) & vbCrLf
        If METCOSTPER.Text <> "" Then Note.Text &= String.Format("材料費占比：{0}%", METCOSTPER.Text) & vbCrLf '材料費占比
        If METExceeDesc.Text <> "" Then Note.Text &= "超出材料費比率上限原因說明：" & METExceeDesc.Text & vbCrLf

        If rst Then
            rst = InputCost6(tmpNote, tmpNoteDt)
            Note.Text &= tmpNote
        End If
        If rst Then
            rst = InputCost7(tmpNote, tmpNoteDt)
            Note.Text &= tmpNote
        End If
        If rst Then
            rst = InputCost8(tmpNote, tmpNoteDt)
            Note.Text &= tmpNote
        End If
        If rst Then
            rst = InputCost9(tmpNote, tmpNoteDt)
            Note.Text &= tmpNote
        End If
        If rst Then
            rst = InputNote2(tmpNote, tmpNoteDt)
            Note.Text &= tmpNote
        End If
        Return rst
    End Function

    '新增 一人份材料明細
    Sub AddPersonCost()
        Dim dt As DataTable = Nothing
        Dim Errmsg As String = ""
        Dim iItemNo As Integer = Val(tItemNo6.Text)
        Dim sCName As String = "" & TIMS.ClearSQM(tCName6.Text) 'Trim( tCName6.Text)
        Dim sStandard As String = "" & TIMS.ClearSQM(tStandard6.Text) 'Trim( tStandard6.Text)
        Dim sUnit As String = "" & TIMS.ClearSQM(tUnit6.Text) 'Trim( tUnit6.Text)
        Dim iPerCount As Integer = Val(tPerCount6.Text)
        Dim iTNum As Integer = Val(TNum.Text) '取得外部資料
        Dim sPurpose As String = "" & TIMS.ClearSQM(tPurpose6.Text) 'Trim( tPurpose6.Text)
        If Not TIMS.IS_DataTable(Session(hid_PersonCostTable_guid1.Value)) Then
            dt = CreatePersonCost()
        Else
            '有資料
            dt = Session(hid_PersonCostTable_guid1.Value)
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11) 'window.scroll(0,document.body.scrollHeight);
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        If iTNum = 0 Then Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf

        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入品名" & vbCrLf
            If sStandard = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iPerCount = 0 Then Errmsg &= "請輸入每人數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "品名" & " 長度不可超過30" & vbCrLf
            If Len(sStandard) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        'If Convert.ToString(iItemNo).Length > 30 Then Errmsg &= "項次長度，超過系統範圍，請修正" & vbCrLf
        If sCName.Length > 30 Then Errmsg &= "品名長度，超過系統範圍，請修正" & vbCrLf
        If sStandard.Length > 300 Then Errmsg &= "規格長度，超過系統範圍，請修正" & vbCrLf
        If sUnit.Length > 30 Then Errmsg &= "單位長度，超過系統範圍，請修正" & vbCrLf
        'If Convert.ToString(iPerCount).Length > 30 Then Errmsg &= "每人數量長度，超過系統範圍，請修正" & vbCrLf 'int
        If sPurpose.Length > 300 Then Errmsg &= "用途說明長度，超過系統範圍，請修正" & vbCrLf
        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr(Cst_PersonCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_PersonCostpkName)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standard") = sStandard '300
        dr("Unit") = sUnit '30
        dr("PerCount") = iPerCount 'int
        dr("TNum") = iTNum 'int
        dr("Total") = (iPerCount * iTNum) '顯示重算
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(hid_PersonCostTable_guid1.Value) = dt
        Call CreatePersonCost()
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '新增 共同材料明細
    Sub AddCommonCost()
        Dim dt As DataTable = Nothing
        Dim iItemNo As Integer = Val(tItemNo7.Text)
        Dim sCName As String = "" & Trim(tCName7.Text)
        Dim sStandard As String = "" & Trim(tStandard7.Text)
        Dim sUnit As String = "" & Trim(tUnit7.Text)
        Dim iAllCount As Integer = Val(tAllCount7.Text)
        Dim iTNum As Integer = Val(TNum.Text)   '取得外部資料
        Dim sPurpose As String = "" & Trim(tPurPose7.Text)

        If Session(hid_CommonCostTable_guid1.Value) Is Nothing Then
            dt = CreateCommonCost()
        Else
            '有資料
            dt = Session(hid_CommonCostTable_guid1.Value)
        End If

        '有錯誤離開
        Dim Errmsg As String = ""
        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        If iTNum = 0 Then Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入品名" & vbCrLf
            If sStandard = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iAllCount = 0 Then Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "品名" & " 長度不可超過30" & vbCrLf
            If Len(sStandard) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr(Cst_CommonCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_CommonCostpkName)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standard") = sStandard '300
        dr("Unit") = sUnit '30
        dr("AllCount") = iAllCount 'int
        dr("TNum") = iTNum 'int
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(hid_CommonCostTable_guid1.Value) = dt
        Call CreateCommonCost()
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '修改 一人份材料明細
    Function Chkdg6(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo6")
        Dim eCName As TextBox = e.Item.FindControl("eCName6")
        Dim eStandard As TextBox = e.Item.FindControl("eStandard6")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit6")
        Dim ePerCount As TextBox = e.Item.FindControl("ePerCount6")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum6") '訓練人數
        Dim eTotal As TextBox = e.Item.FindControl("eTotal6") '總數量 = val(ePerCount6.text)* val(eTNum6.text)
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose6")

        eCName.Text = TIMS.ClearSQM(eCName.Text)
        eStandard.Text = TIMS.ClearSQM(eStandard.Text)
        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & eCName.Text
        Dim sStandard As String = "" & eStandard.Text
        Dim sUnit As String = "" & eUnit.Text
        Dim iPerCount As Integer = Val(ePerCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)
        Dim sPurpose As String = "" & ePurPose.Text
        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料

        Dim Errmsg As String = ""
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入品名" & vbCrLf
            If sStandard = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iPerCount = 0 Then Errmsg &= "請輸入每人數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
            If Not TIMS.IsNumeric2(eItemNo.Text) Then Errmsg &= "項次格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Not TIMS.IsNumeric2(ePerCount.Text) Then Errmsg &= "每人數量格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid6
                Dim dt As DataTable = Session(hid_PersonCostTable_guid1.Value)
                Dim sfilter As String = String.Concat(Cst_PersonCostpkName, "<>'", DGobj.DataKeys(e.Item.ItemIndex), "' AND ItemNo='", eItemNo.Text, "'")
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then Errmsg &= "[" & eItemNo.Text & "]該項次 已在表格中" & vbCrLf
                dt = Nothing
            End If
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "品名" & " 長度不可超過30" & vbCrLf
            If Len(sStandard) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js1)
            sm.LastErrorMessage = Errmsg
            rst = False 'Exit Function
        End If
        Return rst
    End Function

    '修改 共同材料明細
    Function Chkdg7(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo7")
        Dim eCName As TextBox = e.Item.FindControl("eCName7")
        Dim eStandard As TextBox = e.Item.FindControl("eStandard7")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit7")
        Dim eAllCount As TextBox = e.Item.FindControl("eAllCount7")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum7") '訓練人數
        Dim esubtotal As TextBox = e.Item.FindControl("esubtotal7") '小計
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose7")

        eCName.Text = TIMS.ClearSQM(eCName.Text)
        eStandard.Text = TIMS.ClearSQM(eStandard.Text)
        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & eCName.Text
        Dim sStandard As String = "" & eStandard.Text
        Dim sUnit As String = "" & eUnit.Text
        Dim iAllCount As Integer = Val(eAllCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
        Dim sPurpose As String = "" & ePurPose.Text
        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料

        Dim Errmsg As String = ""
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入品名" & vbCrLf
            If sStandard = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iAllCount = 0 Then Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
            If Not TIMS.IsNumeric2(eItemNo.Text) Then Errmsg &= "項次格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Not TIMS.IsNumeric2(eAllCount.Text) Then Errmsg &= "使用數量格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid7
                Dim dt As DataTable = Session(hid_CommonCostTable_guid1.Value)
                Dim sfilter As String = String.Concat(Cst_CommonCostpkName, "<>'", DGobj.DataKeys(e.Item.ItemIndex), "' AND ItemNo='", eItemNo.Text, "'")
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then Errmsg &= "[" & eItemNo.Text & "]該項次已在表格中" & vbCrLf
                dt = Nothing
            End If
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "品名" & " 長度不可超過30" & vbCrLf
            If Len(sStandard) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js1)
            sm.LastErrorMessage = Errmsg
            rst = False  'Exit Function
        End If
        Return rst
    End Function

    Private Sub DataGrid6_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid6.ItemCommand
        Dim Errmsg As String = ""
        If Session(hid_PersonCostTable_guid1.Value) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid6
        Dim dt As DataTable = Session(hid_PersonCostTable_guid1.Value)
        'Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case "EDT6" '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case "DEL6" '刪除
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                Dim sfilter As String = "" & Cst_PersonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr As DataRow In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            Case "UPD6" '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo6")
                Dim eCName As TextBox = e.Item.FindControl("eCName6")
                Dim eStandard As TextBox = e.Item.FindControl("eStandard6")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit6")
                Dim ePerCount As TextBox = e.Item.FindControl("ePerCount6")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum6") '訓練人數
                Dim eTotal As TextBox = e.Item.FindControl("eTotal6") '總數量 = val(ePerCount6.text)* val(eTNum6.text)
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose6")
                If Chkdg6(e) Then
                    Dim sfilter As String = "" & Cst_PersonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                        eItemNo.Text = TIMS.ClearSQM(eItemNo.Text)
                        eCName.Text = TIMS.ClearSQM(eCName.Text)
                        eStandard.Text = TIMS.ClearSQM(eStandard.Text)
                        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
                        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
                        Dim iPerCount As Integer = Val(ePerCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)
                        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                        Dim dr As DataRow = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standard") = eStandard.Text
                        dr("Unit") = eUnit.Text
                        dr("PerCount") = iPerCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("Total") = iPerCount * iTNum '顯示重算
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If
            Case "CLS6" '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(hid_PersonCostTable_guid1.Value) = dt  '要新  
        CreatePersonCost() '建立
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid6_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid6.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost6.Enabled
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                'Dim CheckBoxDG6 As CheckBox = e.Item.FindControl("CheckBoxDG6")
                Dim lItemNo6 As Label = e.Item.FindControl("lItemNo6")
                Dim lCName6 As Label = e.Item.FindControl("lCName6")
                Dim lStandard6 As Label = e.Item.FindControl("lStandard6")
                Dim lUnit6 As Label = e.Item.FindControl("lUnit6")
                Dim lPerCount6 As Label = e.Item.FindControl("lPerCount6")
                Dim lTNum6 As Label = e.Item.FindControl("lTNum6")
                Dim lTotal6 As Label = e.Item.FindControl("lTotal6")
                Dim lPurPose6 As Label = e.Item.FindControl("lPurPose6")
                Dim btnDEL6 As Button = e.Item.FindControl("btnDEL6") '刪除
                Dim btnEDT6 As Button = e.Item.FindControl("btnEDT6") '修改
                Dim Hid_DataKey As HiddenField = e.Item.FindControl("Hid_DataKey")
                Hid_DataKey.Value = TIMS.EncryptAes(drv(Cst_PersonCostpkName))

                lItemNo6.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName6.Text = "" & Convert.ToString(drv("CName"))
                lStandard6.Text = "" & Convert.ToString(drv("Standard"))
                lUnit6.Text = "" & Convert.ToString(drv("Unit"))
                lPerCount6.Text = "" & Convert.ToString(drv("PerCount"))
                lTNum6.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                lTotal6.Text = "" & (Val(drv("PerCount")) * Val(drv("TNum"))) '顯示重算
                lPurPose6.Text = "" & Convert.ToString(drv("PurPose"))
                btnDEL6.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDEL6.Enabled = Flag_AddEnabled
                btnEDT6.Enabled = Flag_AddEnabled
            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim tlItemNo6 As TextBox = e.Item.FindControl("eItemNo6")
                Dim tlCName6 As TextBox = e.Item.FindControl("eCName6")
                Dim tlStandard6 As TextBox = e.Item.FindControl("eStandard6")
                Dim tlUnit6 As TextBox = e.Item.FindControl("eUnit6")
                Dim tlPerCount6 As TextBox = e.Item.FindControl("ePerCount6")
                Dim tlTNum6 As TextBox = e.Item.FindControl("eTNum6")
                Dim tlTotal6 As TextBox = e.Item.FindControl("eTotal6")
                Dim tlPurPose6 As TextBox = e.Item.FindControl("ePurPose6")
                Dim btnUPD6 As Button = e.Item.FindControl("btnUPD6") '更新
                Dim btnCLS6 As Button = e.Item.FindControl("btnCLS6") '取消
                tlItemNo6.Text = "" & Convert.ToString(drv("ItemNo"))
                tlCName6.Text = "" & Convert.ToString(drv("CName"))
                tlStandard6.Text = "" & Convert.ToString(drv("Standard"))
                tlUnit6.Text = "" & Convert.ToString(drv("Unit"))
                tlPerCount6.Text = "" & Convert.ToString(drv("PerCount"))
                tlTNum6.Text = "" & Convert.ToString(drv("TNum")) '顯示原資料
                tlTotal6.Text = "" & (Val(drv("PerCount")) * Val(drv("TNum"))) '顯示重算
                tlPurPose6.Text = "" & Convert.ToString(drv("PurPose"))
                tlTNum6.ReadOnly = True
                tlTotal6.ReadOnly = True
                tlTNum6.Style.Item("background-color") = "#BDBDBD"
                tlTotal6.Style.Item("background-color") = "#BDBDBD"
                btnUPD6.Enabled = Flag_AddEnabled
                btnCLS6.Enabled = True
        End Select
    End Sub

    '新增 'PLAN_PERSONCOST–一人份材料明細 
    Private Sub BtnAddCost6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCost6.Click
        Call AddPersonCost()
    End Sub

    '匯入明細 'PLAN_PERSONCOST–一人份材料明細
    Protected Sub BtnImport1_Click(sender As Object, e As EventArgs) Handles BtnImport1.Click
        Dim Errmsg As String = ""
        Dim rst As Boolean = True
        Try
            rst = File1_test(Errmsg)
        Catch ex As Exception
            rst = False
            Errmsg = ex.Message
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        If rst Then
            Call CreatePersonCost() '顯示 內容
        Else
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '新增 'PLAN_COMMONCOST–共同材料明細
    Private Sub BtnAddCost7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddCost7.Click
        Call AddCommonCost()
    End Sub

    '匯入明細  'PLAN_COMMONCOST–共同材料明細
    Protected Sub BtnImport2_Click(sender As Object, e As EventArgs) Handles BtnImport2.Click
        Dim Errmsg As String = ""
        Dim rst As Boolean = True
        Try
            rst = File2_test(Errmsg)
        Catch ex As Exception
            rst = False
            Errmsg = ex.Message
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        If rst Then
            Call CreateCommonCost() '顯示 內容
        Else
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid7_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid7.ItemCommand
        Dim Errmsg As String = ""
        If Session(hid_CommonCostTable_guid1.Value) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid7
        Dim dt As DataTable = Session(hid_CommonCostTable_guid1.Value)
        Dim dr As DataRow = Nothing
        Errmsg = ""

        Select Case e.CommandName
            Case "EDT7" '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case "DEL7" '刪除
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                Dim sfilter As String = "" & Cst_CommonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            Case "UPD7" '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo7")
                Dim eCName As TextBox = e.Item.FindControl("eCName7")
                Dim eStandard As TextBox = e.Item.FindControl("eStandard7")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit7")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount7")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum7") '訓練人數
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose7")
                If Chkdg7(e) Then
                    Dim sfilter As String = "" & Cst_CommonCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                        eItemNo.Text = TIMS.ClearSQM(eItemNo.Text)
                        eCName.Text = TIMS.ClearSQM(eCName.Text)
                        eStandard.Text = TIMS.ClearSQM(eStandard.Text)
                        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
                        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
                        Dim iAllCount As Integer = Val(eAllCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
                        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standard") = eStandard.Text
                        dr("Unit") = eUnit.Text
                        dr("AllCount") = iAllCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If
            Case "CLS7" '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(Cst_CommonCostpkName) = dt  '要新  
        CreateCommonCost() '建立  
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid7_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid7.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost7.Enabled

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo As Label = e.Item.FindControl("lItemNo7")
                Dim lCName As Label = e.Item.FindControl("lCName7")
                Dim lStandard As Label = e.Item.FindControl("lStandard7")
                Dim lUnit As Label = e.Item.FindControl("lUnit7")
                Dim lAllCount As Label = e.Item.FindControl("lAllCount7")
                Dim lTNum As Label = e.Item.FindControl("lTNum7")
                Dim lPurPose As Label = e.Item.FindControl("lPurPose7")
                Dim btnDEL As Button = e.Item.FindControl("btnDEL7") '刪除
                Dim btnEDT As Button = e.Item.FindControl("btnEDT7") '修改
                Dim Hid_DataKey As HiddenField = e.Item.FindControl("Hid_DataKey")
                Hid_DataKey.Value = TIMS.EncryptAes(drv(Cst_CommonCostpkName))

                lItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName.Text = "" & Convert.ToString(drv("CName"))
                lStandard.Text = "" & Convert.ToString(drv("Standard"))
                lUnit.Text = "" & Convert.ToString(drv("Unit"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                lAllCount.Text = iAllCount
                lTNum.Text = iTNum '顯示原資料
                lPurPose.Text = "" & Convert.ToString(drv("PurPose"))
                btnDEL.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDEL.Enabled = Flag_AddEnabled
                btnEDT.Enabled = Flag_AddEnabled
            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim tlItemNo As TextBox = e.Item.FindControl("eItemNo7")
                Dim tlCName As TextBox = e.Item.FindControl("eCName7")
                Dim tlStandard As TextBox = e.Item.FindControl("eStandard7")
                Dim tlUnit As TextBox = e.Item.FindControl("eUnit7")
                Dim tlAllCount As TextBox = e.Item.FindControl("eAllCount7")
                Dim tlTNum As TextBox = e.Item.FindControl("eTNum7")
                Dim tlPurPose As TextBox = e.Item.FindControl("ePurPose7")
                Dim btnUPD As Button = e.Item.FindControl("btnUPD7") '更新
                Dim btnCLS As Button = e.Item.FindControl("btnCLS7") '取消
                tlItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                tlCName.Text = "" & Convert.ToString(drv("CName"))
                tlStandard.Text = "" & Convert.ToString(drv("Standard"))
                tlUnit.Text = "" & Convert.ToString(drv("Unit"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                tlAllCount.Text = iAllCount
                tlTNum.Text = iTNum '顯示原資料
                tlPurPose.Text = "" & Convert.ToString(drv("PurPose"))
                tlTNum.ReadOnly = True
                tlTNum.Style.Item("background-color") = "#BDBDBD"
                btnUPD.Enabled = Flag_AddEnabled
                btnCLS.Enabled = True
        End Select
    End Sub

    Function SUtl_UpdateNote(ByRef tConn As SqlConnection) As Boolean
        Dim rst As Boolean = True '正常/false:異常
        Dim s_NoteText As String = Note.Text 'TIMS.ClearSQM(Note.Text)
        If s_NoteText = "" Then Return rst

        Call TIMS.OpenDbConn(tConn)
        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        If upt_PlanX.Value <> "" Then '有儲存資料過了
            '有儲存資料過了,準備儲存資料
            tmpPCS = upt_PlanX.Value  '有儲存資料過了
            PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
            ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
            SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
            sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "'"
            dt = DbAccess.GetDataTable(sql, da, tConn)
            dr = dt.Rows(0)
        Else
            If (Convert.ToString(Request("PlanID")) = "" OrElse gflag_ccopy) Then
                '新增資料 、copy=1 、草稿新增 而來
                sm.LastErrorMessage = cst_errmsg19
                rst = False
                Return False
            Else
                '修改
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "'"
                dt = DbAccess.GetDataTable(sql, da, tConn)
                dr = dt.Rows(0)
            End If
            tmpPCS = ""
            TIMS.SetMyValue(tmpPCS, "PlanID", PlanID_value)
            TIMS.SetMyValue(tmpPCS, "ComIDNO", ComIDNO_value)
            TIMS.SetMyValue(tmpPCS, "SeqNO", SeqNO_value)
            upt_PlanX.Value = tmpPCS
        End If
        dr("Note") = Note.Text
        DbAccess.UpdateDataTable(dt, da)

        Return rst
    End Function

    '匯出EXCEL
    Private Sub Button21b_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button21b.Click
        Dim rst As Boolean = True '正常/false:異常
        rst = ChangNoteText(tmpNoteDt) '將資料寫入Note
        '匯出EXCEL
        If (LayerState.Value = "") Then LayerState.Value = "8"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("window_onload", s_js1)

        If rst Then
            '只修改 Note儲存格
            rst = SUtl_UpdateNote(objconn) '<==內部有顯示錯誤訊息
            If Not rst Then Exit Sub '異常離開
        End If
        If Not rst Then
            sm.LastErrorMessage = cst_errmsg20
            Exit Sub
        End If

        rst = False '異常 True:正常
        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing
        Dim dt2 As DataTable = Nothing
        Dim sql As String = ""
        PlanID_value = TIMS.ClearSQM(Request("PlanID"))
        ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
        SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
        Dim flag_error_1 As Boolean = (PlanID_value = "" OrElse ComIDNO_value = "" OrElse SeqNO_value = "")
        If flag_error_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim pms1 As New Hashtable From {{"PlanID", PlanID_value}, {"ComIDNO", ComIDNO_value}, {"SeqNO", SeqNO_value}}
        sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
        dt = DbAccess.GetDataTable(sql, objconn, pms1)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim pms1p As New Hashtable From {{"PlanID", PlanID_value}, {"ComIDNO", ComIDNO_value}, {"SeqNO", SeqNO_value}}
        sql = " SELECT * FROM PLAN_PERSONCOST WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
        dt1 = DbAccess.GetDataTable(sql, objconn, pms1p)

        Dim pms1c As New Hashtable From {{"PlanID", PlanID_value}, {"ComIDNO", ComIDNO_value}, {"SeqNO", SeqNO_value}}
        sql = " SELECT * FROM PLAN_COMMONCOST WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO"
        dt2 = DbAccess.GetDataTable(sql, objconn, pms1c)
        If dt.Rows.Count = 1 Then
            If dt1.Rows.Count > 0 OrElse dt2.Rows.Count > 0 Then rst = True
        End If
        If Not rst Then
            sm.LastErrorMessage = Cst_msgother3
            Exit Sub
        End If

        Dim MyValue As String = ""
        MyValue = "YEARS=" & Convert.ToString(sm.UserInfo.Years - 1911)
        MyValue += "&PLANID=" & PlanID_value
        MyValue += "&ComIDNO=" & ComIDNO_value
        MyValue += "&SEQNO=" & SeqNO_value
        MyValue += "&PCSValue=" & String.Concat(PlanID_value, "x", ComIDNO_value, "x", SeqNO_value)
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue) '材料明細表
    End Sub

    Private Sub BtnUptNote2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnUptNote2.Click
        Call ChangNoteText(tmpNoteDt)
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '修改 教材明細
    Function Chkdg8(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo8")
        Dim eCName As TextBox = e.Item.FindControl("eCName8")
        Dim eStandards As TextBox = e.Item.FindControl("eStandards8")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit8")
        Dim eAllCount As TextBox = e.Item.FindControl("eAllCount8")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum8") '訓練人數
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose8")

        eCName.Text = TIMS.ClearSQM(eCName.Text)
        eStandards.Text = TIMS.ClearSQM(eStandards.Text)
        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & eCName.Text
        Dim sStandards As String = "" & eStandards.Text
        Dim sUnit As String = "" & eUnit.Text
        Dim iAllCount As Integer = Val(eAllCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
        Dim sPurpose As String = "" & ePurPose.Text
        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料

        Dim Errmsg As String = ""
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入品名" & vbCrLf
            If sStandards = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iAllCount = 0 Then Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
            If Not TIMS.IsNumeric2(eItemNo.Text) Then Errmsg &= "項次格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Not TIMS.IsNumeric2(eAllCount.Text) Then Errmsg &= "使用數量格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid8
                Dim dt As DataTable = Session(hid_SheetCostTable_guid1.Value)
                Dim sfilter As String = String.Concat(Cst_SheetCostpkName, "<>'", DGobj.DataKeys(e.Item.ItemIndex), "' AND ItemNo='", eItemNo.Text, "'")
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then Errmsg &= "[" & eItemNo.Text & "]該項次已在表格中" & vbCrLf
                dt = Nothing
            End If
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "品名" & " 長度不可超過30" & vbCrLf
            If Len(sStandards) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js1)
            sm.LastErrorMessage = Errmsg
            rst = False 'Exit Function
        End If
        Return rst
    End Function

    '修改 其他明細
    Function Chkdg9(ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) As Boolean
        Dim rst As Boolean = True '檢查ok 
        Dim eItemNo As TextBox = e.Item.FindControl("eItemNo9")
        Dim eCName As TextBox = e.Item.FindControl("eCName9")
        Dim eStandards As TextBox = e.Item.FindControl("eStandards9")
        Dim eUnit As TextBox = e.Item.FindControl("eUnit9")
        Dim eAllCount As TextBox = e.Item.FindControl("eAllCount9")
        Dim eTNum As TextBox = e.Item.FindControl("eTNum9") '訓練人數
        Dim ePurPose As TextBox = e.Item.FindControl("ePurPose9")

        eCName.Text = TIMS.ClearSQM(eCName.Text)
        eStandards.Text = TIMS.ClearSQM(eStandards.Text)
        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
        Dim iItemNo As Integer = Val(eItemNo.Text)
        Dim sCName As String = "" & eCName.Text
        Dim sStandards As String = "" & eStandards.Text
        Dim sUnit As String = "" & eUnit.Text
        Dim iAllCount As Integer = Val(eAllCount.Text)
        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
        Dim sPurpose As String = "" & ePurPose.Text
        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料

        Dim Errmsg As String = ""
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入項目" & vbCrLf
            If sStandards = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iAllCount = 0 Then Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
            If Not TIMS.IsNumeric2(eItemNo.Text) Then Errmsg &= "項次格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Not TIMS.IsNumeric2(eAllCount.Text) Then Errmsg &= "使用數量格式有誤，應為正整數數字格式" & vbCrLf 'int
            If Errmsg = "" AndAlso eItemNo.Text <> "" Then
                Dim DGobj As DataGrid = DataGrid9
                Dim dt As DataTable = Session(hid_OtherCostTable_guid1.Value)
                Dim sfilter As String = String.Concat(Cst_OtherCostpkName, "<>'", DGobj.DataKeys(e.Item.ItemIndex), "' AND ItemNo='", eItemNo.Text, "'")
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then Errmsg &= "[" & eItemNo.Text & "]該項次 已在表格中" & vbCrLf
                dt = Nothing
            End If
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "項目" & " 長度不可超過30" & vbCrLf
            If Len(sStandards) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js1)
            sm.LastErrorMessage = Errmsg
            rst = False 'Exit Function
        End If
        Return rst
    End Function

    '建立 教材明細
    Function CreateSheetCost() As DataTable
        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing 'copy
        Dim DGobj As DataGrid = DataGrid8
        Const cst_sSupFd As String = ",0 subtotal, 0 eachCost" '補充欄位
        Dim sql As String = ""
        If Not TIMS.IS_DataTable(Session(hid_SheetCostTable_guid1.Value)) Then
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT PLAN_SHEETCOST.* " & cst_sSupFd & " FROM PLAN_SHEETCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT * FROM PLAN_SHEETCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If (Not gflag_can_copy2) Then sql &= " AND 1<>1"
                    If g_flagNG Then sql = " SELECT *  FROM PLAN_SHEETCOST WHERE 1<>1"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT PLAN_SHEETCOST.* " & cst_sSupFd & " FROM PLAN_SHEETCOST where 1<>1 " 'Copy機制
                Else
                    '修改資料取得
                    sql = ""
                    sql &= " SELECT PLAN_SHEETCOST.* " & cst_sSupFd & " FROM PLAN_SHEETCOST" & vbCrLf
                    sql &= " WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT PLAN_SHEETCOST.* " & cst_sSupFd & " FROM PLAN_SHEETCOST where 1<>1 "
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(hid_SheetCostTable_guid1.Value) '有資料
        End If
        If dt Is Nothing Then Return dt

        dt.Columns(Cst_SheetCostpkName).AutoIncrement = True
        dt.Columns(Cst_SheetCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_SheetCostpkName).AutoIncrementStep = -1
        If gflag_ccopy Then TIMS.CopyDATATABLE(dt, dt1, Cst_SheetCostpkName)

        Session(hid_SheetCostTable_guid1.Value) = dt
        With DGobj
            .Style.Item("display") = If(dt.Rows.Count > 0, "", "none")
            .DataSource = dt
            .DataKeyField = Cst_SheetCostpkName
            .DataBind()
        End With
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    '建立 其他明細
    Function CreateOtherCost() As DataTable
        Dim dt As DataTable = Nothing
        Dim dt1 As DataTable = Nothing 'copy
        Dim DGobj As DataGrid = DataGrid9
        Const cst_sSupFd As String = ",0 subtotal, 0 eachCost" '補充欄位
        Dim sql As String = ""
        If Not TIMS.IS_DataTable(Session(hid_OtherCostTable_guid1.Value)) Then
            If upt_PlanX.Value <> "" Then
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
                sql = " SELECT PLAN_OTHERCOST.* " & cst_sSupFd & " FROM PLAN_OTHERCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' " 'Copy機制
            Else
                PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                If gflag_ccopy Then
                    sql = " SELECT *  FROM PLAN_OTHERCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    If (Not gflag_can_copy2) Then sql &= " AND 1<>1"
                    If g_flagNG Then sql = " SELECT *  FROM PLAN_OTHERCOST WHERE 1<>1"
                    dt1 = DbAccess.GetDataTable(sql, objconn)
                    sql = " SELECT PLAN_OTHERCOST.* " & cst_sSupFd & " FROM PLAN_OTHERCOST WHERE 1<>1 " 'Copy機制
                Else
                    '修改資料取得
                    sql = ""
                    sql &= " SELECT PLAN_OTHERCOST.* " & cst_sSupFd & " FROM PLAN_OTHERCOST" & vbCrLf
                    sql &= " WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "'"
                    sql &= " ORDER BY ItemNo" & vbCrLf
                    If g_flagNG Then sql = " SELECT PLAN_OTHERCOST.* " & cst_sSupFd & " FROM PLAN_OTHERCOST WHERE 1<>1 "
                End If
            End If
            dt = DbAccess.GetDataTable(sql, objconn)
        Else
            dt = Session(hid_OtherCostTable_guid1.Value) '有資料
        End If
        If dt Is Nothing Then Return dt

        dt.Columns(Cst_OtherCostpkName).AutoIncrement = True
        dt.Columns(Cst_OtherCostpkName).AutoIncrementSeed = -1
        dt.Columns(Cst_OtherCostpkName).AutoIncrementStep = -1
        If gflag_ccopy Then TIMS.CopyDATATABLE(dt, dt1, Cst_OtherCostpkName)

        Session(hid_OtherCostTable_guid1.Value) = dt
        With DGobj
            .Style.Item("display") = If(dt.Rows.Count > 0, "", "none")
            .DataSource = dt
            .DataKeyField = Cst_OtherCostpkName
            .DataBind()
        End With
        Call ChangNoteText(tmpNoteDt)
        Return dt
    End Function

    '新增 教材明細
    Sub AddSheetCost()
        Dim dt As DataTable = Nothing
        Dim Errmsg As String = ""
        Dim iItemNo As Integer = Val(tItemNo8.Text)
        Dim sCName As String = "" & Trim(tCName8.Text)
        Dim sStandards As String = "" & Trim(tStandards8.Text)
        Dim sUnit As String = "" & Trim(tUnit8.Text)
        Dim iAllCount As Integer = Val(tAllCount8.Text)
        Dim iTNum As Integer = Val(TNum.Text)   '取得外部資料
        Dim sPurpose As String = "" & Trim(tPurPose8.Text)

        If Session(hid_SheetCostTable_guid1.Value) Is Nothing Then
            dt = CreateSheetCost()
        Else
            '有資料
            dt = Session(hid_SheetCostTable_guid1.Value)
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        If iTNum = 0 Then Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf 'CHECK:1
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入品名" & vbCrLf
            If sStandards = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iAllCount = 0 Then Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "品名" & " 長度不可超過30" & vbCrLf
            If Len(sStandards) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr(Cst_SheetCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_SheetCostpkName)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standards") = sStandards '300
        dr("Unit") = sUnit '30
        dr("AllCount") = iAllCount 'int
        dr("TNum") = iTNum 'int
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(hid_SheetCostTable_guid1.Value) = dt
        Call CreateSheetCost()
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '新增 其他明細
    Sub AddOtherCost()
        Dim dt As DataTable = Nothing
        Dim Errmsg As String = ""
        Dim iItemNo As Integer = Val(tItemNo9.Text)
        Dim sCName As String = "" & Trim(tCName9.Text)
        Dim sStandards As String = "" & Trim(tStandards9.Text)
        Dim sUnit As String = "" & Trim(tUnit9.Text)
        Dim iAllCount As Integer = Val(tAllCount9.Text)
        Dim iTNum As Integer = Val(TNum.Text)   '取得外部資料
        Dim sPurpose As String = "" & Trim(tPurpose9.Text)

        If Session(hid_OtherCostTable_guid1.Value) Is Nothing Then
            dt = CreateOtherCost()
        Else
            '有資料
            dt = Session(hid_OtherCostTable_guid1.Value)
        End If

        If dt.Rows.Count > 0 Then
            For i As Int16 = 0 To dt.Rows.Count - 1
                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                        Errmsg &= "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                        Exit For
                    End If
                End If
            Next
        End If

        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        If iTNum = 0 Then Errmsg &= "請先輸入訓練人數，不可為0" & vbCrLf 'CHECK:1
        If Errmsg = "" Then
            If iItemNo = 0 Then Errmsg &= "請輸入項次，不可為0" & vbCrLf 'int
            If sCName = "" Then Errmsg &= "請輸入項目" & vbCrLf
            If sStandards = "" Then Errmsg &= "請輸入規格" & vbCrLf
            If sUnit = "" Then Errmsg &= "請輸入單位" & vbCrLf
            If iAllCount = 0 Then Errmsg &= "請輸入使用數量，不可為0" & vbCrLf 'int
            If sPurpose = "" Then Errmsg &= "請輸入用途說明" & vbCrLf
        End If
        If Errmsg = "" Then
            If Len(sCName) > 30 Then Errmsg &= "項目" & " 長度不可超過30" & vbCrLf
            If Len(sStandards) > 300 Then Errmsg &= "規格" & " 長度不可超過300" & vbCrLf
            If Len(sUnit) > 30 Then Errmsg &= "單位" & " 長度不可超過30" & vbCrLf
            If Len(sPurpose) > 300 Then Errmsg &= "用途說明" & " 長度不可超過300" & vbCrLf
        End If
        '有錯誤離開
        If Errmsg <> "" Then
            If (LayerState.Value = "") Then LayerState.Value = "6"
            Dim s_js11 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
            Page.RegisterStartupScript("Londing", s_js11)
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If

        '產業人才投資方案專用
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr(Cst_OtherCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_OtherCostpkName)
        dr("ItemNo") = iItemNo 'int
        dr("CName") = sCName '30
        dr("Standards") = sStandards '300
        dr("Unit") = sUnit '30
        dr("AllCount") = iAllCount 'int
        dr("TNum") = iTNum 'int
        dr("PurPose") = sPurpose '300
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now

        Session(hid_OtherCostTable_guid1.Value) = dt
        Call CreateOtherCost()
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '匯入'PLAN_SHEETCOST–教材明細
    Protected Sub BtnImport8_Click(sender As Object, e As EventArgs) Handles BtnImport8.Click
        Dim Errmsg As String = ""
        Dim rst As Boolean = True
        Try
            rst = File3_test(Errmsg)
        Catch ex As Exception
            rst = False
            Errmsg = ex.Message
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        If rst Then
            Call CreateSheetCost() '顯示 內容
        Else
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '新增 'PLAN_SHEETCOST–教材明細
    Protected Sub BtnAddCost8_Click(sender As Object, e As EventArgs) Handles btnAddCost8.Click
        Call AddSheetCost()
    End Sub

    Private Sub DataGrid8_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid8.ItemCommand
        Dim Errmsg As String = ""
        If Session(hid_SheetCostTable_guid1.Value) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid8
        Const cst_eCmdEDT As String = "EDT8"
        Const cst_eCmdDEL As String = "DEL8"
        Const cst_eCmdUPD As String = "UPD8"
        Const cst_eCmdCLS As String = "CLS8"
        Dim dt As DataTable = Session(hid_SheetCostTable_guid1.Value)
        Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case cst_eCmdEDT '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case cst_eCmdDEL '刪除
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                Dim sfilter As String = "" & Cst_SheetCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            Case cst_eCmdUPD '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo8")
                Dim eCName As TextBox = e.Item.FindControl("eCName8")
                Dim eStandards As TextBox = e.Item.FindControl("eStandards8")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit8")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount8")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum8") '訓練人數
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose8")
                If Chkdg8(e) Then
                    Dim sfilter As String = "" & Cst_SheetCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                        eItemNo.Text = TIMS.ClearSQM(eItemNo.Text)
                        eCName.Text = TIMS.ClearSQM(eCName.Text)
                        eStandards.Text = TIMS.ClearSQM(eStandards.Text)
                        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
                        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
                        Dim iAllCount As Integer = Val(eAllCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
                        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standards") = eStandards.Text
                        dr("Unit") = eUnit.Text
                        dr("AllCount") = iAllCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If
            Case cst_eCmdCLS '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select
        Session(Cst_SheetCostpkName) = dt  '要新  
        CreateSheetCost() '建立  
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid8_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid8.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost8.Enabled
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo As Label = e.Item.FindControl("lItemNo8")
                Dim lCName As Label = e.Item.FindControl("lCName8")
                Dim lStandards As Label = e.Item.FindControl("lStandards8")
                Dim lUnit As Label = e.Item.FindControl("lUnit8")
                Dim lAllCount As Label = e.Item.FindControl("lAllCount8")
                Dim lTNum As Label = e.Item.FindControl("lTNum8")
                Dim lPurPose As Label = e.Item.FindControl("lPurPose8")
                Dim btnDel8 As Button = e.Item.FindControl("btnDel8") '刪除
                Dim btnEdt8 As Button = e.Item.FindControl("btnEdt8") '修改
                Dim Hid_DataKey As HiddenField = e.Item.FindControl("Hid_DataKey")
                Hid_DataKey.Value = TIMS.EncryptAes(drv(Cst_SheetCostpkName))

                lItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName.Text = "" & Convert.ToString(drv("CName"))
                lStandards.Text = "" & Convert.ToString(drv("Standards"))
                lUnit.Text = "" & Convert.ToString(drv("Unit"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                lAllCount.Text = iAllCount
                lTNum.Text = iTNum '顯示原資料
                lPurPose.Text = "" & Convert.ToString(drv("PurPose"))
                btnDel8.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDel8.Enabled = Flag_AddEnabled
                btnEdt8.Enabled = Flag_AddEnabled
            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim tlItemNo As TextBox = e.Item.FindControl("eItemNo8")
                Dim tlCName As TextBox = e.Item.FindControl("eCName8")
                Dim tlStandards As TextBox = e.Item.FindControl("eStandards8")
                Dim tlUnit As TextBox = e.Item.FindControl("eUnit8")
                Dim tlAllCount As TextBox = e.Item.FindControl("eAllCount8")
                Dim tlTNum As TextBox = e.Item.FindControl("eTNum8")
                Dim tlPurPose As TextBox = e.Item.FindControl("ePurPose8")
                Dim btnUpd8 As Button = e.Item.FindControl("btnUpd8") '更新
                Dim btnCls8 As Button = e.Item.FindControl("btnCls8") '取消
                tlItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                tlCName.Text = "" & Convert.ToString(drv("CName"))
                tlStandards.Text = "" & Convert.ToString(drv("Standards"))
                tlUnit.Text = "" & Convert.ToString(drv("Unit"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                tlAllCount.Text = iAllCount
                tlTNum.Text = iTNum '顯示原資料
                tlPurPose.Text = "" & Convert.ToString(drv("PurPose"))
                tlTNum.ReadOnly = True
                tlTNum.Style.Item("background-color") = "#BDBDBD"
                btnUpd8.Enabled = Flag_AddEnabled
                btnCls8.Enabled = True
        End Select
    End Sub

    '匯入'PLAN_OTHERCOST–其他明細
    Protected Sub BtnImport9_Click(sender As Object, e As EventArgs) Handles BtnImport9.Click
        Dim Errmsg As String = ""
        Dim rst As Boolean = True
        Try
            rst = File4_test(Errmsg)
        Catch ex As Exception
            rst = False
            Errmsg = ex.Message
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        If rst Then
            Call CreateOtherCost() '顯示 內容
        Else
            sm.LastErrorMessage = Errmsg
            Exit Sub
        End If
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    '新增 'PLAN_OTHERCOST–其他明細
    Protected Sub BtnAddCost9_Click(sender As Object, e As EventArgs) Handles btnAddCost9.Click
        Call AddOtherCost()
    End Sub

    Private Sub DataGrid9_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid9.ItemCommand
        Dim Errmsg As String = ""
        If Session(hid_OtherCostTable_guid1.Value) Is Nothing Then Exit Sub
        Dim DGobj As DataGrid = DataGrid9
        Const cst_eCmdEDT As String = "EDT9"
        Const cst_eCmdDEL As String = "DEL9"
        Const cst_eCmdUPD As String = "UPD9"
        Const cst_eCmdCLS As String = "CLS9"

        Dim dt As DataTable = Session(hid_OtherCostTable_guid1.Value)
        Dim dr As DataRow = Nothing
        Errmsg = ""
        Select Case e.CommandName
            Case cst_eCmdEDT '修改
                DGobj.EditItemIndex = e.Item.ItemIndex '修改列數改變
            Case cst_eCmdDEL '刪除
                If DGobj Is Nothing OrElse dt Is Nothing Then
                    sm.LastErrorMessage = cst_errmsg16
                    Exit Sub
                End If
                Dim sfilter As String = "" & Cst_OtherCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                '搜尋刪除資料刪除
                If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length <> 0 Then
                    For Each dr In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            Case cst_eCmdUPD '更新
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo9")
                Dim eCName As TextBox = e.Item.FindControl("eCName9")
                Dim eStandards As TextBox = e.Item.FindControl("eStandards9")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit9")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount9")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum9") '訓練人數
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose9")
                If Chkdg9(e) Then
                    Dim sfilter As String = "" & Cst_OtherCostpkName & "='" & DGobj.DataKeys(e.Item.ItemIndex) & "'"
                    If Convert.ToString(DGobj.DataKeys(e.Item.ItemIndex)) <> "" AndAlso dt.Select(sfilter).Length > 0 Then
                        eItemNo.Text = TIMS.ClearSQM(eItemNo.Text)
                        eCName.Text = TIMS.ClearSQM(eCName.Text)
                        eStandards.Text = TIMS.ClearSQM(eStandards.Text)
                        eUnit.Text = TIMS.ClearSQM(eUnit.Text)
                        ePurPose.Text = TIMS.ClearSQM(ePurPose.Text)
                        Dim iAllCount As Integer = Val(eAllCount.Text)
                        Dim iTNum As Integer = Val(eTNum.Text)  '顯示原資料
                        If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                        dr = dt.Select(sfilter)(0)
                        dr("ItemNo") = eItemNo.Text
                        dr("CName") = eCName.Text
                        dr("Standards") = eStandards.Text
                        dr("Unit") = eUnit.Text
                        dr("AllCount") = iAllCount
                        dr("TNum") = iTNum '顯示原資料
                        dr("PurPose") = ePurPose.Text
                    End If
                    DGobj.EditItemIndex = -1 '還原修改列數
                End If
            Case cst_eCmdCLS '取消
                DGobj.EditItemIndex = -1 '還原修改列數
        End Select

        Session(hid_OtherCostTable_guid1.Value) = dt  '要新  
        CreateOtherCost() '建立  
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DataGrid9_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid9.ItemDataBound
        Dim Flag_AddEnabled As Boolean = btnAddCost9.Enabled 'btnAddCost6.Enabled
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                '顯示
                Dim drv As DataRowView = e.Item.DataItem
                Dim lItemNo As Label = e.Item.FindControl("lItemNo9")
                Dim lCName As Label = e.Item.FindControl("lCName9")
                Dim lStandards As Label = e.Item.FindControl("lStandards9")
                Dim lUnit As Label = e.Item.FindControl("lUnit9")
                Dim lAllCount As Label = e.Item.FindControl("lAllCount9")
                Dim lTNum As Label = e.Item.FindControl("lTNum9")
                Dim lPurPose As Label = e.Item.FindControl("lPurPose9")
                Dim btnDel9 As Button = e.Item.FindControl("btnDel9") '刪除
                Dim btnEdt9 As Button = e.Item.FindControl("btnEdt9") '修改
                Dim Hid_DataKey As HiddenField = e.Item.FindControl("Hid_DataKey")
                Hid_DataKey.Value = TIMS.EncryptAes(drv(Cst_OtherCostpkName))

                lItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                lCName.Text = "" & Convert.ToString(drv("CName"))
                lStandards.Text = "" & Convert.ToString(drv("Standards"))
                lUnit.Text = "" & Convert.ToString(drv("Unit"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                lAllCount.Text = iAllCount
                lTNum.Text = iTNum '顯示原資料
                lPurPose.Text = "" & Convert.ToString(drv("PurPose"))
                btnDel9.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                btnDel9.Enabled = Flag_AddEnabled
                btnDel9.Enabled = Flag_AddEnabled

            Case ListItemType.EditItem
                '編輯
                Dim drv As DataRowView = e.Item.DataItem
                Dim eItemNo As TextBox = e.Item.FindControl("eItemNo9")
                Dim eCName As TextBox = e.Item.FindControl("eCName9")
                Dim eStandards As TextBox = e.Item.FindControl("eStandards9")
                Dim eUnit As TextBox = e.Item.FindControl("eUnit9")
                Dim eAllCount As TextBox = e.Item.FindControl("eAllCount9")
                Dim eTNum As TextBox = e.Item.FindControl("eTNum9")
                Dim ePurPose As TextBox = e.Item.FindControl("ePurPose9")
                Dim btnUpd9 As Button = e.Item.FindControl("btnUpd9") '更新
                Dim btnCls9 As Button = e.Item.FindControl("btnCls9") '取消
                eItemNo.Text = "" & Convert.ToString(drv("ItemNo"))
                eCName.Text = "" & Convert.ToString(drv("CName"))
                eStandards.Text = "" & Convert.ToString(drv("Standards"))
                eUnit.Text = "" & Convert.ToString(drv("Unit"))
                Dim iAllCount As Integer = Val(drv("AllCount"))
                Dim iTNum As Integer = Val(drv("TNum"))  '取得外部資料
                If iTNum <> Val(TNum.Text) Then iTNum = Val(TNum.Text) '取得外部資料
                eAllCount.Text = iAllCount
                eTNum.Text = iTNum '顯示原資料
                ePurPose.Text = "" & Convert.ToString(drv("PurPose"))
                eTNum.ReadOnly = True
                eTNum.Style.Item("background-color") = "#BDBDBD"
                btnUpd9.Enabled = Flag_AddEnabled
                btnCls9.Enabled = True
        End Select
    End Sub

    ''' <summary> (固定費用人時成本上限) 產投／充電起飛計畫 班級申請 固定費用人時成本上限為140元配合111年計畫修正，修正為160元。 </summary>
    ''' <returns></returns>
    Function Get_iMAX_ACTHUMCOST() As Integer
        '「固定費用總額單一人時成本」之卡控金額調整為160元，當此欄位數字超過160.00元，須填寫「超出人時成本原因說明」，未填不能基本儲存。
        Dim i_rst As Integer = cst_iMAX_ACTHUMCOST_28_old
        '二、上開增修，基於產投及充飛受理時間不同，請依不同計畫各自上版：
        '1、產業人才投資方案先於9月26日上版。 2023/09/26
        '2、充電起飛計畫(54)於11月30日上版。 2023/11/30
        '2、充電起飛計畫(54)於11月30日上版。 2023/11/30
        Dim date_ACTHUMCOST_28_2023 As Date = CDate("2023/09/26")
        Dim date_ACTHUMCOST_54_2023 As Date = CDate("2023/11/30")
        Dim date_ACTHUMCOST_28_2025 As Date = CDate("2025/1/1")
        Dim date_ACTHUMCOST_54_2025 As Date = CDate("2025/1/1")

        If (TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            If DateDiff(DateInterval.Day, date_ACTHUMCOST_28_2025, Now) >= 0 Then
                i_rst = cst_iMAX_ACTHUMCOST_28
            ElseIf DateDiff(DateInterval.Day, date_ACTHUMCOST_28_2023, Now) >= 0 Then
                i_rst = cst_iMAX_ACTHUMCOST_28_old2
            Else
                i_rst = cst_iMAX_ACTHUMCOST_28_old
            End If
        ElseIf (TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then
            If DateDiff(DateInterval.Day, date_ACTHUMCOST_54_2025, Now) >= 0 Then
                i_rst = cst_iMAX_ACTHUMCOST_54
            ElseIf DateDiff(DateInterval.Day, date_ACTHUMCOST_54_2023, Now) >= 0 Then
                i_rst = cst_iMAX_ACTHUMCOST_54_old2
            Else
                i_rst = cst_iMAX_ACTHUMCOST_54_old
            End If
        Else
            i_rst = cst_iMAX_ACTHUMCOST_28_old2
        End If
        Return i_rst
    End Function

    ''' <summary>初始化，欄位長度設定</summary>
    Sub SUtl_PageInit1()
        '2023/2024/ContactPhone
        fg_phone_2024 = TIMS.CHK_PHONE_2024(sm)
        trContactPhone_2023_O1.Visible = (Not fg_phone_2024)
        trContactPhone_2024_N1.Visible = fg_phone_2024
        trContactPhone_2024_N2.Visible = fg_phone_2024

        '固定費用人時成本上限為140元  'ACTHUMCOST > 140  '充電起飛140 '產投160 by 20210929 AMU
        '填寫說明： <br />一、固定費用人時成本上限為140元，如人時成本編列需超出上限，請填寫超出原因。 <br />二、19大類材料費編列比上限
        hid_MAX_iACTHUMCOST.Value = Get_iMAX_ACTHUMCOST().ToString()
        Dim msg_Label4a As String = ""
        msg_Label4a = ""
        msg_Label4a &= "填寫說明： <br />" ' & vbCrLf
        msg_Label4a &= String.Format("一、固定費用人時成本上限為{0}元，如人時成本編列需超出上限，請填寫超出原因。 <br />", Get_iMAX_ACTHUMCOST()) '& vbCrLf
        msg_Label4a &= "二、19大類材料費編列比上限"
        Label4a.Text = msg_Label4a

        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("PLAN_PLANINFO,PLAN_ONCLASS", objconn) ' DbAccess.GetDataTable(sql)
        If dt.Rows.Count = 0 Then Exit Sub
        'Call TIMS.sUtl_SetMaxLen(dt, "ENTERPRISENAME", EnterpriseName) '企業包班名稱
        Call TIMS.sUtl_SetMaxLen(dt, "PLANEMAIL", EMail) 'EMAIL
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER1", Other1) '其他一
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER2", Other2) '其他二
        Call TIMS.sUtl_SetMaxLen(dt, "CAPOTHER3", Other3) '其他三
        'Call TIMS.sUtl_SetMaxLen(dt, "CLASSNAME", ClassName, -15) '班別名稱
        Call TIMS.sUtl_SetMaxLen(dt, "CYCLTYPE", CyclType) '期別
        Call TIMS.sUtl_SetMaxLen(dt, "ROOMNAME", RoomName) '上課教室名稱
        Call TIMS.sUtl_SetMaxLen(dt, "FACTMODEOTHER", FactModeOther) '場地類型其他說明
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTNAME", ContactName) '聯絡人
        '2023/2024/ContactPhone
        'Call TIMS.sUtl_SetMaxLen(dt, "CONTACTPHONE", ContactPhone) '電話
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTEMAIL", ContactEmail) '電子郵件
        Call TIMS.sUtl_SetMaxLen(dt, "CONTACTFAX", ContactFax) '傳真
        Call TIMS.sUtl_SetMaxLen(dt, "TIMES", txtTimes) '時間 Times
    End Sub

    ''' <summary> (SAVE) INSERT PLAN_DEPOT (SAVE) </summary>
    ''' <param name="sSearchW"></param>
    ''' <param name="oConn"></param>
    Sub SAVE_PLAN_DEPOT(ByVal sSearchW As String, ByVal oConn As SqlConnection)
        Dim sql As String = ""
        '確認
        Dim PlanID As String = TIMS.GetMyValue(sSearchW, "PlanID")
        Dim ComIDNO As String = TIMS.GetMyValue(sSearchW, "ComIDNO")
        Dim SeqNo As String = TIMS.GetMyValue(sSearchW, "SeqNo")
        Dim SEQNOD15 As String = TIMS.GetMyValue(sSearchW, "SEQNOD15")
        'Dim KID06 As String = TIMS.GetMyValue(sSearchW, "KID06") 'Dim KID10 As String = TIMS.GetMyValue(sSearchW, "KID10")
        Dim KID18 As String = TIMS.GetMyValue(sSearchW, "KID18")
        Dim KID19 As String = TIMS.GetMyValue(sSearchW, "KID19")
        Dim KID20 As String = TIMS.GetMyValue(sSearchW, "KID20")
        Dim KID22 As String = TIMS.GetMyValue(sSearchW, "KID22")
        Dim KID25 As String = TIMS.GetMyValue(sSearchW, "KID25")
        Dim KID26 As String = TIMS.GetMyValue(sSearchW, "KID26")
        Dim KID60 As String = TIMS.GetMyValue(sSearchW, "KID60")
        If PlanID = "" OrElse ComIDNO = "" OrElse SeqNo = "" Then Exit Sub
        If KID25 <> "" AndAlso KID20 <> "" Then KID20 = ""
        If KID26 <> "" AndAlso KID25 <> "" Then KID25 = ""
        If KID26 <> "" AndAlso KID20 <> "" Then KID20 = ""
        sql = " SELECT 1 FROM dbo.PLAN_DEPOT WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim sCmd As New SqlCommand(sql, oConn)
        TIMS.OpenDbConn(oConn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PLANID", SqlDbType.Int).Value = TIMS.CINT1(PlanID)
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = ComIDNO
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = TIMS.CINT1(SeqNo)
            dt1.Load(.ExecuteReader())
        End With

        If dt1.Rows.Count = 0 Then
            'INSERT
            Dim iCmd_Parms As New Hashtable From {
                {"PLANID", TIMS.CINT1(PlanID)}, {"COMIDNO", ComIDNO}, {"SEQNO", TIMS.CINT1(SeqNo)},
                {"SEQNOD15", If(SEQNOD15 <> "", SEQNOD15, Convert.DBNull)}, 'SEQNOD15
                {"KID19", If(KID19 <> "", KID19, Convert.DBNull)}, 'KID19
                {"KID18", If(KID18 <> "", KID18, Convert.DBNull)}, 'KID18
                {"KID20", If(KID20 <> "", KID20, Convert.DBNull)}, 'KID20
                {"KID22", If(KID22 <> "", KID22, Convert.DBNull)}, 'KID22
                {"KID25", If(KID25 <> "", KID25, Convert.DBNull)}, 'KID25
                {"KID26", If(KID26 <> "", KID26, Convert.DBNull)}, 'KID26
                {"KID60", If(KID60 <> "", KID60, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim i_sql As String = "
INSERT INTO PLAN_DEPOT(PLANID,COMIDNO,SEQNO,SEQNOD15,KID19,KID18,KID20,KID22,KID25,KID26,KID60,APPRESULT ,MODIFYACCT ,MODIFYDATE)
VALUES (@PLANID,@COMIDNO,@SEQNO ,@SEQNOD15,@KID19,@KID18,@KID20,@KID22,@KID25,@KID26,@KID60,'Y',@MODIFYACCT ,GETDATE())
"
            DbAccess.ExecuteNonQuery(i_sql, oConn, iCmd_Parms)
        Else
            'UPDATE 'SQL &= " ,KID06 = @KID06 ,KID10 = @KID10" & VBCRLF
            Dim uCmd_Parms As New Hashtable From {
                {"SEQNOD15", If(SEQNOD15 <> "", SEQNOD15, Convert.DBNull)}, 'SEQNOD15
                {"KID19", If(KID19 <> "", KID19, Convert.DBNull)},
                {"KID18", If(KID18 <> "", KID18, Convert.DBNull)}, 'KID18
                {"KID20", If(KID20 <> "", KID20, Convert.DBNull)}, 'KID20
                {"KID22", If(KID22 <> "", KID22, Convert.DBNull)}, 'KID22
                {"KID25", If(KID25 <> "", KID25, Convert.DBNull)}, 'KID25
                {"KID26", If(KID26 <> "", KID26, Convert.DBNull)}, 'KID26
                {"KID60", If(KID60 <> "", KID60, Convert.DBNull)},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"PLANID", TIMS.CINT1(PlanID)}, {"COMIDNO", ComIDNO}, {"SEQNO", TIMS.CINT1(SeqNo)}
            }
            Dim u_sql As String = "
UPDATE PLAN_DEPOT SET SEQNOD15 = @SEQNOD15 ,KID19=@KID19 ,KID18=@KID18,KID20=@KID20,KID22=@KID22,KID25=@KID25,KID26=@KID26,KID60=@KID60
,APPRESULT='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE() WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO
"
            DbAccess.ExecuteNonQuery(u_sql, oConn, uCmd_Parms)
        End If
    End Sub

    '匯入 一人份材料明細
    Public Function File1_test(ByRef rErrmsg As String) As Boolean
        Dim rst As Boolean = False '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim cst_maxColCount As Integer = 6
        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        If File1.Value = "" Then
            rErrmsg = "請選擇匯入的檔案!"
            Return rst 'Exit Function
        End If
        If File1.PostedFile.ContentLength = 0 Then
            rErrmsg = "檔案位置錯誤!"
            Return rst 'Exit Function
        End If

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        If MyFileName.IndexOf(".") = -1 Then
            rErrmsg = "檔案類型錯誤!"
            sm.LastErrorMessage = rErrmsg
            Return rst 'Exit Function
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If Not LCase(MyFileType) = "csv" Then
            rErrmsg = "檔案類型錯誤，必須為CSV檔!"
            sm.LastErrorMessage = rErrmsg
            Return rst 'Exit Function
        End If

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
        File1.PostedFile.SaveAs(filePath1) '上傳檔案

        '將檔案讀出放入記憶體
        Dim sr As System.IO.Stream = IO.File.OpenRead(filePath1)
        Dim srr As System.IO.StreamReader = New System.IO.StreamReader(sr, System.Text.Encoding.Default)
        Dim RowIndex As Integer = 0 '讀取行累計數
        Dim OneRow As String        'srr.ReadLine 一行一行的資料
        Dim colArray As Array
        Dim dt As DataTable = Nothing
        If Session(hid_PersonCostTable_guid1.Value) Is Nothing Then
            dt = CreatePersonCost()
        Else
            '有資料
            dt = Session(hid_PersonCostTable_guid1.Value)
        End If

        Dim Reason As String = ""
        Do While srr.Peek >= 0
            If Reason <> "" Then Exit Do
            OneRow = srr.ReadLine
            If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈
            If RowIndex <> 0 Then
                Reason = ""
                colArray = Split(OneRow, cst_flag1)
                If colArray.Length < cst_maxColCount Then
                    rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                    sm.LastErrorMessage = rErrmsg
                    rst = False
                    Exit Do
                End If
                Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                Dim sCName As String = "" & TIMS.ClearSQM(colArray(1).ToString) '品名
                Dim sStandard As String = "" & TIMS.ClearSQM(colArray(2).ToString) '規格
                Dim sUnit As String = "" & TIMS.ClearSQM(colArray(3).ToString) '單位
                Dim iPerCount As Integer = Val(colArray(4).ToString)  '每人數量
                Dim iTNum As Integer = Val(TNum.Text) '取得外部資料(不可為0)
                Dim sPurpose As String = "" & TIMS.ClearSQM(colArray(5).ToString) '用途說明

                '檢查資料正確性
                If Reason = "" Then
                    If dt.Rows.Count > 0 Then
                        For i As Int16 = 0 To dt.Rows.Count - 1
                            If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                    Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                    Exit For
                                End If
                            End If
                        Next
                    End If

                    If iTNum = 0 Then Reason += "請先輸入訓練人數，不可為0" & vbCrLf 'CHECK:1

                    If Reason = "" Then
                        If iItemNo = 0 Then Reason += "請輸入項次，不可為0" & vbCrLf 'int
                        If sCName = "" Then Reason += "請輸入品名" & vbCrLf
                        If sStandard = "" Then Reason += "請輸入規格" & vbCrLf
                        If sUnit = "" Then Reason += "請輸入單位" & vbCrLf
                        If iPerCount = 0 Then Reason += "請輸入每人數量，不可為0" & vbCrLf 'int
                        If sPurpose = "" Then Reason += "請輸入用途說明" & vbCrLf
                    End If

                    If Reason = "" Then
                        If Len(sCName) > 30 Then Reason += "品名" & " 長度不可超過30" & vbCrLf
                        If Len(sStandard) > 300 Then Reason += "規格" & " 長度不可超過300" & vbCrLf
                        If Len(sUnit) > 30 Then Reason += "單位" & " 長度不可超過30" & vbCrLf
                        If Len(sPurpose) > 300 Then Reason += "用途說明" & " 長度不可超過300" & vbCrLf
                    End If
                End If

                If Reason = "" Then
                    '產業人才投資方案專用
                    Dim dr As DataRow = dt.NewRow
                    dt.Rows.Add(dr)
                    dr(Cst_PersonCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_PersonCostpkName)
                    dr("ItemNo") = iItemNo 'int
                    dr("CName") = sCName '30
                    dr("Standard") = sStandard '300
                    dr("Unit") = sUnit '30
                    dr("PerCount") = iPerCount 'int
                    dr("TNum") = iTNum 'int
                    dr("Total") = (iPerCount * iTNum) '顯示重算
                    dr("PurPose") = sPurpose '300
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                    Session(hid_PersonCostTable_guid1.Value) = dt
                End If
            End If
            RowIndex += 1 '讀取行累計數
        Loop
        sr.Close()
        srr.Close()
        'IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
        '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
        TIMS.MyFileDelete(filePath1)

        If Reason <> "" Then
            rErrmsg = Reason
            Return False
        End If
        rst = True
        Return rst
    End Function

    '匯入 共同材料明細
    Function File2_test(ByRef rErrmsg As String) As Boolean
        Dim rst As Boolean = False '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = ""               '儲存錯誤的原因
        Dim MyFileName, MyFileType As String
        Dim flag As String = ","

        Dim dt As DataTable = Nothing
        If Session(hid_CommonCostTable_guid1.Value) Is Nothing Then
            dt = CreateCommonCost()
        Else
            '有資料
            dt = Session(hid_CommonCostTable_guid1.Value)
        End If

        Dim oFile As HtmlInputFile = File2
        If oFile.Value <> "" Then
            '檢查檔案格式與大小 Start
            If File2.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                sm.LastErrorMessage = "檔案位置錯誤!"
                Return rst
            Else
                MyFileName = Split(oFile.PostedFile.FileName, "\")((Split(oFile.PostedFile.FileName, "\")).Length - 1) '取出檔案名稱
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    sm.LastErrorMessage = "檔案類型錯誤!"
                    Return rst
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        sm.LastErrorMessage = "檔案類型錯誤，必須為CSV檔!"
                        Return rst
                    End If
                End If
            End If
            '檢查檔案格式與大小 End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
            oFile.PostedFile.SaveAs(filePath1) '上傳檔案

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(filePath1)
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, ",", "") = "" Then Exit Do '若資料為空白行，則離開回圈
                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, flag)
                    If colArray.Length < 6 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & TIMS.ClearSQM(colArray(1).ToString) '品名
                    Dim sStandard As String = "" & TIMS.ClearSQM(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & TIMS.ClearSQM(colArray(3).ToString) '單位
                    Dim iAllCount As Integer = Val(colArray(4).ToString) '使用數量
                    Dim iTNum As Integer = Val(TNum.Text)   '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & TIMS.ClearSQM(colArray(5).ToString) '用途說明

                    '檢查資料正確性
                    If Reason = "" Then
                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then Reason += "請先輸入訓練人數，不可為0" & vbCrLf 'CHECK:1

                        If Reason = "" Then
                            If iItemNo = 0 Then Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            If sCName = "" Then Reason += "請輸入品名" & vbCrLf
                            If sStandard = "" Then Reason += "請輸入規格" & vbCrLf
                            If sUnit = "" Then Reason += "請輸入單位" & vbCrLf
                            If iAllCount = 0 Then Reason += "請輸入使用數量，不可為0" & vbCrLf 'int
                            If sPurpose = "" Then Reason += "請輸入用途說明" & vbCrLf
                        End If
                        If Reason = "" Then
                            If Len(sCName) > 30 Then Reason += "品名" & " 長度不可超過30" & vbCrLf
                            If Len(sStandard) > 300 Then Reason += "規格" & " 長度不可超過300" & vbCrLf
                            If Len(sUnit) > 30 Then Reason += "單位" & " 長度不可超過30" & vbCrLf
                            If Len(sPurpose) > 300 Then Reason += "用途說明" & " 長度不可超過300" & vbCrLf
                        End If
                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr(Cst_CommonCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_CommonCostpkName)
                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standard") = sStandard '300
                        dr("Unit") = sUnit '30
                        dr("AllCount") = iAllCount 'int
                        dr("TNum") = iTNum 'int
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        Session(hid_CommonCostTable_guid1.Value) = dt
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            'IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
            '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
            TIMS.MyFileDelete(filePath1)
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    '匯入 教材明細
    Function File3_test(ByRef rErrmsg As String) As Boolean
        Dim rst As Boolean = False '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim oFile As HtmlInputFile = File3
        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = "" '儲存錯誤的原因
        Const cst_flag As String = ","
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        Dim dt As DataTable = Nothing
        If Session(hid_SheetCostTable_guid1.Value) Is Nothing Then
            dt = CreateSheetCost()
        Else
            '有資料
            dt = Session(hid_SheetCostTable_guid1.Value)
        End If

        If oFile.Value <> "" Then
            '檢查檔案格式與大小 Start
            If oFile.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                sm.LastErrorMessage = "檔案位置錯誤!"
                Return rst
            Else
                MyFileName = Split(oFile.PostedFile.FileName, "\")((Split(oFile.PostedFile.FileName, "\")).Length - 1) '取出檔案名稱
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    sm.LastErrorMessage = "檔案類型錯誤!"
                    Return rst 'Exit Function
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        sm.LastErrorMessage = "檔案類型錯誤，必須為CSV檔!"
                        Return rst 'Exit Function
                    End If
                End If
            End If
            '檢查檔案格式與大小 End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
            oFile.PostedFile.SaveAs(filePath1) '上傳檔案

            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(filePath1)
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, cst_flag, "") = "" Then Exit Do '若資料為空白行，則離開回圈
                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, cst_flag)
                    If colArray.Length < 6 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & TIMS.ClearSQM(colArray(1).ToString) '項目
                    Dim sStandards As String = "" & TIMS.ClearSQM(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & TIMS.ClearSQM(colArray(3).ToString) '單位
                    Dim iAllCount As Integer = Val(colArray(4).ToString) '使用數量
                    Dim iTNum As Integer = Val(TNum.Text)   '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & TIMS.ClearSQM(colArray(5).ToString) '用途說明

                    '檢查資料正確性
                    If Reason = "" Then
                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then Reason += "請先輸入訓練人數，不可為0" & vbCrLf 'CHECK:1

                        If Reason = "" Then
                            If iItemNo = 0 Then Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            If sCName = "" Then Reason += "請輸入品名" & vbCrLf
                            If sStandards = "" Then Reason += "請輸入規格" & vbCrLf
                            If sUnit = "" Then Reason += "請輸入單位" & vbCrLf
                            If iAllCount = 0 Then Reason += "請輸入使用數量，不可為0" & vbCrLf 'int
                            If sPurpose = "" Then Reason += "請輸入用途說明" & vbCrLf
                        End If
                        If Reason = "" Then
                            If Len(sCName) > 30 Then Reason += "品名" & " 長度不可超過30" & vbCrLf
                            If Len(sStandards) > 300 Then Reason += "規格" & " 長度不可超過300" & vbCrLf
                            If Len(sUnit) > 30 Then Reason += "單位" & " 長度不可超過30" & vbCrLf
                            If Len(sPurpose) > 300 Then Reason += "用途說明" & " 長度不可超過300" & vbCrLf
                        End If
                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr(Cst_SheetCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_SheetCostpkName)
                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standards") = sStandards '300
                        dr("Unit") = sUnit '30
                        dr("AllCount") = iAllCount 'int
                        dr("TNum") = iTNum 'int
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        Session(hid_SheetCostTable_guid1.Value) = dt
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            'IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
            '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
            TIMS.MyFileDelete(filePath1)
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    '匯入 其他費用明細
    Function File4_test(ByRef rErrmsg As String) As Boolean
        Dim rst As Boolean = False '得知檔案，初步確認資料內容是否無誤 true 有誤 false
        Dim oFile As HtmlInputFile = File4
        Dim Upload_Path As String = "~/TC/01/Temp/"
        Call TIMS.MyCreateDir(Me, Upload_Path)
        Dim Reason As String = "" '儲存錯誤的原因
        Const cst_flag As String = ","
        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        Dim dt As DataTable = Nothing
        If Session(hid_OtherCostTable_guid1.Value) Is Nothing Then
            dt = CreateOtherCost()
        Else
            '有資料
            dt = Session(hid_OtherCostTable_guid1.Value)
        End If

        If oFile.Value <> "" Then
            '檢查檔案格式與大小 Start
            If oFile.PostedFile.ContentLength = 0 Then
                Common.MessageBox(Me, "檔案位置錯誤!")
                sm.LastErrorMessage = "檔案位置錯誤!"
                Return rst 'Exit Function
            Else
                MyFileName = Split(oFile.PostedFile.FileName, "\")((Split(oFile.PostedFile.FileName, "\")).Length - 1) '取出檔案名稱
                '取出檔案類型
                If MyFileName.IndexOf(".") = -1 Then
                    Common.MessageBox(Me, "檔案類型錯誤!")
                    sm.LastErrorMessage = "檔案類型錯誤!"
                    Return rst 'Exit Function
                Else
                    MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
                    If Not LCase(MyFileType) = "csv" Then
                        Common.MessageBox(Me, "檔案類型錯誤，必須為CSV檔!")
                        sm.LastErrorMessage = "檔案類型錯誤，必須為CSV檔!"
                        Return rst 'Exit Function
                    End If
                End If
            End If
            '檢查檔案格式與大小 End

            '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
            Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile.PostedFile.FileName).ToLower()
            MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
            Dim filePath1 As String = Server.MapPath($"{Upload_Path}{MyFileName}")
            oFile.PostedFile.SaveAs(filePath1) '上傳檔案
            '將檔案讀出放入記憶體
            Dim sr As System.IO.Stream
            Dim srr As System.IO.StreamReader
            sr = IO.File.OpenRead(filePath1)
            srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

            Dim RowIndex As Integer = 0 '讀取行累計數
            Dim OneRow As String        'srr.ReadLine 一行一行的資料
            Dim colArray As Array

            Do While srr.Peek >= 0
                If Reason <> "" Then Exit Do
                OneRow = srr.ReadLine
                If Replace(OneRow, cst_flag, "") = "" Then Exit Do '若資料為空白行，則離開回圈
                If RowIndex <> 0 Then
                    Reason = ""
                    colArray = Split(OneRow, cst_flag)
                    If colArray.Length < 6 Then
                        rErrmsg &= "匯入的檔案欄位資料不足，請檢查匯入檔案正確性。" & vbCrLf
                        rst = False
                        Exit Do
                    End If
                    Dim iItemNo As Integer = Val(colArray(0).ToString) '項次
                    Dim sCName As String = "" & TIMS.ClearSQM(colArray(1).ToString) '品名
                    Dim sStandards As String = "" & TIMS.ClearSQM(colArray(2).ToString) '規格
                    Dim sUnit As String = "" & TIMS.ClearSQM(colArray(3).ToString) '單位
                    Dim iAllCount As Integer = Val(colArray(4).ToString) '使用數量
                    Dim iTNum As Integer = Val(TNum.Text)   '取得外部資料(不可為0)
                    Dim sPurpose As String = "" & TIMS.ClearSQM(colArray(5).ToString) '用途說明

                    '檢查資料正確性
                    If Reason = "" Then
                        If dt.Rows.Count > 0 Then
                            For i As Int16 = 0 To dt.Rows.Count - 1
                                If Not dt.Rows(i).RowState = DataRowState.Deleted Then '已刪除者不可做更動
                                    If Val(dt.Rows(i).Item("ItemNo")) = iItemNo Then
                                        Reason += "[" & iItemNo & "]該項次 已在表格中" & vbCrLf
                                        Exit For
                                    End If
                                End If
                            Next
                        End If

                        If iTNum = 0 Then Reason += "請先輸入訓練人數，不可為0" & vbCrLf 'CHECK:1

                        If Reason = "" Then
                            If iItemNo = 0 Then Reason += "請輸入項次，不可為0" & vbCrLf 'int
                            If sCName = "" Then Reason += "請輸入項目" & vbCrLf
                            If sStandards = "" Then Reason += "請輸入規格" & vbCrLf
                            If sUnit = "" Then Reason += "請輸入單位" & vbCrLf
                            If iAllCount = 0 Then Reason += "請輸入使用數量，不可為0" & vbCrLf 'int
                            If sPurpose = "" Then Reason += "請輸入用途說明" & vbCrLf
                        End If
                        If Reason = "" Then
                            If Len(sCName) > 30 Then Reason += "品名" & " 長度不可超過30" & vbCrLf
                            If Len(sStandards) > 300 Then Reason += "規格" & " 長度不可超過300" & vbCrLf
                            If Len(sUnit) > 30 Then Reason += "單位" & " 長度不可超過30" & vbCrLf
                            If Len(sPurpose) > 300 Then Reason += "用途說明" & " 長度不可超過300" & vbCrLf
                        End If

                    End If

                    If Reason = "" Then
                        '產業人才投資方案專用
                        Dim dr As DataRow = dt.NewRow
                        dt.Rows.Add(dr)
                        dr(Cst_OtherCostpkName) = TIMS.GET_NEWPK_INT(Me, Cst_OtherCostpkName)
                        dr("ItemNo") = iItemNo 'int
                        dr("CName") = sCName '30
                        dr("Standards") = sStandards '300
                        dr("Unit") = sUnit '30
                        dr("AllCount") = iAllCount 'int
                        dr("TNum") = iTNum 'int
                        dr("PurPose") = sPurpose '300
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyDate") = Now
                        Session(hid_OtherCostTable_guid1.Value) = dt
                    End If
                End If
                RowIndex += 1 '讀取行累計數
            Loop
            sr.Close()
            srr.Close()
            'IO.File.Delete(Server.MapPath(Upload_Path & MyFileName))
            '刪除檔案 IO.File.Delete(Server.MapPath(Upload_Path & MyFileName)),IO.File.Delete(filePath1)
            TIMS.MyFileDelete(filePath1)
        Else
            Reason += "請選擇匯入的檔案" & vbCrLf
        End If

        If Reason <> "" Then
            rErrmsg = Reason
            rst = False
        Else
            rst = True
        End If
        Return rst
    End Function

    ''' <summary>
    ''' 授課教師/助教 - 遴選辦法說明
    ''' </summary>
    ''' <param name="TECHTYPE"></param>
    ''' <returns></returns>
    Function GET_TeacherDesc_AB(ByVal TECHTYPE As String) As String
        Dim rst As String = ""
        If g_flagNG Then Return rst 'Exit Sub
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return rst 'Exit Sub

        'from PLAN_TEACHER
        Dim sql As String = ""
        Dim dr As DataRow = Nothing
        Select Case UCase(TECHTYPE)
            Case "A" '教師
                sql = ""
                sql &= " WITH WC1 AS (select max(TechID) TechID from PLAN_TRAINDESC where planid=" & rqPlanID & " and comidno='" & rqComIDNO & "' and seqno=" & rqSeqNO & " and TechID is not null)" & vbCrLf
                sql &= " select dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'A', a.TechID) TeacherDesc" & vbCrLf
                sql &= " from PLAN_PLANINFO b" & vbCrLf
                sql &= " CROSS JOIN WC1 a" & vbCrLf
                sql &= " WHERE planid=" & rqPlanID & " and comidno='" & rqComIDNO & "' and seqno=" & rqSeqNO & "" & vbCrLf
                dr = DbAccess.GetOneRow(sql, objconn)
            Case "B" '教師2/助教
                sql = ""
                sql &= " WITH WC1 AS (select max(TechID2) TechID from PLAN_TRAINDESC where planid=" & rqPlanID & " and comidno='" & rqComIDNO & "' and seqno=" & rqSeqNO & " and TechID2 is not null)" & vbCrLf
                sql &= " select dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'B', a.TechID) TeacherDesc" & vbCrLf
                sql &= " from PLAN_PLANINFO b" & vbCrLf
                sql &= " CROSS JOIN WC1 a" & vbCrLf
                sql &= " WHERE planid=" & rqPlanID & " and comidno='" & rqComIDNO & "' and seqno=" & rqSeqNO & "" & vbCrLf
                dr = DbAccess.GetOneRow(sql, objconn)
        End Select
        If dr IsNot Nothing Then rst = Convert.ToString(dr("TeacherDesc"))
        Return rst
    End Function

    ''' <summary> SHOW PLAN_VERREPORT </summary>
    Sub SHOW_PLAN_VERREPORT() 'ByVal htSS As Hashtable)
        If g_flagNG Then
            sm.LastErrorMessage = cst_errmsg3
            Exit Sub
        End If
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return 'rst 'Exit Sub
        'Dim sRIDn As String = sUtl_GetRIDn()

        BtnSAVE2.Visible = False '不顯示儲存鈕
        Dim flag_can_show_data24 As Boolean = False '不顯示儲存鈕
        'Dim flag_saveBasic_ok As Boolean = Chk_SAVEBASICOK()
        Dim drPP As DataRow = TIMS.GetPCSDate(rqPlanID, rqComIDNO, rqSeqNO, objconn)
        '顯示儲存鈕
        flag_can_show_data24 = If(drPP IsNot Nothing AndAlso drPP("ISAPPRPAPER") = "Y", True, False)
        If flag_can_show_data24 Then BtnSAVE2.Visible = True '顯示儲存鈕

        Dim hParms As New Hashtable From {{"PlanID", Val(rqPlanID)}, {"ComIDNO", rqComIDNO}, {"SeqNo", Val(rqSeqNO)}}
        Dim sql As String = "SELECT * FROM PLAN_VERREPORT WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, hParms)
        If TIMS.dtNODATA(dt) Then Return

        Dim dr As DataRow = dt.Rows(0)
        '本課程屬環境部淨零綠領人才培育課程
        CB_EnvZeroTrain.Checked = If($"{dr("EnvZeroTrain")}" = "Y", True, False)
        '其他設施說明
        txtOthFacDesc23.Text = TIMS.ClearSQM2(dr("OthFacDesc23"))
        '報請主管機關核備
        Dim v_REPORTE As String = Convert.ToString(dr("REPORTE"))
        rbl_REPORTE_Y.Checked = If(v_REPORTE.Equals("Y"), True, False)
        rbl_REPORTE_N.Checked = If(v_REPORTE.Equals("N"), True, False)

        Common.SetListItem(rblFuncLevel, Convert.ToString(dr("FuncLevel")))
        If Convert.ToString(dr("TMethod")) <> "" Then TIMS.SetCblValue(cblTMethod, Convert.ToString(dr("TMethod")))
        TMethodOth.Text = Convert.ToString(dr("TMethodOth"))
        'Common.SetListItem(ClassID, dr("ClassID"))
        TIMS.PL_settextbox1(tPOWERNEED1, dr("POWERNEED1"))
        TIMS.PL_settextbox1(tPOWERNEED2, dr("POWERNEED2"))
        TIMS.PL_settextbox1(tPOWERNEED3, dr("POWERNEED3"))
        '與政策性產業課程之關聯性概述
        TIMS.PL_settextbox1(tPOLICYREL, dr("POLICYREL"))

        cbPOWERNEED4.Checked = (Convert.ToString(dr("POWERNEED4CHK")) = TIMS.cst_YES)
        If cbPOWERNEED4.Checked AndAlso Convert.ToString(dr("POWERNEED4")) <> "" Then TIMS.PL_settextbox1(tPOWERNEED4, dr("POWERNEED4"))
        'If tPlanCause.Text = "" AndAlso Convert.ToString(dr("PlanCause")) <> "" Then tPlanCause.Text = Convert.ToString(dr("PlanCause"))
        'If tPurScience.Text = "" AndAlso Convert.ToString(dr("PurScience")) <> "" Then tPurScience.Text = Convert.ToString(dr("PurScience"))
        'If tPurTech.Text = "" AndAlso Convert.ToString(dr("PurTech")) <> "" Then tPurTech.Text = Convert.ToString(dr("PurTech"))
        'If tPurMoral.Text = "" AndAlso Convert.ToString(dr("PurMoral")) <> "" Then tPurMoral.Text = Convert.ToString(dr("PurMoral"))
        CapAll.Text = dr("CapAll").ToString
        Hid_CapAll.Value = dr("CapAll").ToString
        'If  CostDesc.Text = "" Then
        '    If dr("CostDesc").ToString <> "" Then  CostDesc.Text = dr("CostDesc").ToString
        'End If
        ' TrainMode.Enabled = False
        ' TrainMode.Text = "(請勾選教學方法)"
        RecDesc.Text = Convert.ToString(dr("RecDesc")) '.ToString
        LearnDesc.Text = Convert.ToString(dr("LearnDesc")) '.ToString
        ActDesc.Text = Convert.ToString(dr("ActDesc")) '.ToString
        ResultDesc.Text = Convert.ToString(dr("ResultDesc")) '.ToString
        OtherDesc.Text = Convert.ToString(dr("OtherDesc")) '.ToString

        chk_RecDesc.Checked = If(RecDesc.Text <> "", True, False)
        chk_LearnDesc.Checked = If(LearnDesc.Text <> "", True, False)
        chk_ActDesc.Checked = If(ActDesc.Text <> "", True, False)
        chk_ResultDesc.Checked = If(ResultDesc.Text <> "", True, False)
        chk_OtherDesc.Checked = If(OtherDesc.Text <> "", True, False)

        '是否為iCAP課程 / 是, 請填寫/否/ 課程相關說明
        Dim sISiCAPCOUR As String = Convert.ToString(dr("ISiCAPCOUR"))
        RB_ISiCAPCOUR_Y.Checked = If(sISiCAPCOUR = "Y", True, False)
        RB_ISiCAPCOUR_N.Checked = If(sISiCAPCOUR = "N", True, False)
        iCAPCOURDESC.Text = Convert.ToString(dr("iCAPCOURDESC")) '課程相關說明
        Recruit.Text = Convert.ToString(dr("Recruit")) '招訓方式
        Selmethod.Text = Convert.ToString(dr("Selmethod")) '遴選方式
        Inspire.Text = Convert.ToString(dr("Inspire")) '學員激勵辦法

        'TGovExamCY.Checked = False 'TGovExamCN.Checked = False
        Dim sTGovExam As String = $"{dr("TGovExam")}"
        TGovExamCY.Checked = If(sTGovExam = "Y", True, False)
        TGovExamCN.Checked = If(sTGovExam = "N", True, False)
        'TGovExamCG:本課程結訓後須參加環境部辦理之淨零綠領人才培育課程測驗；測驗成績達及格，即可申請本方案補助。
        TGovExamCG.Checked = If(sTGovExam = "G", True, False)
        Hid_TGovExam.Value = sTGovExam

        GOVAGENAME.Text = $"{dr("GOVAGENAME")}" '政府機關名稱
        TGovExamName.Text = $"{dr("TGovExamName")}" '證照或檢定名稱
        'chkMEMO8C1.Checked = False'chkMEMO8C2.Checked = False
        chkMEMO8C1.Checked = If(Convert.ToString(dr("memo8")) <> "", True, False)
        chkMEMO8C2.Checked = If(Convert.ToString(dr("memo82")) <> "", True, False)
        'txtMemo8.Text = ""
        txtMemo8.Text = If(Convert.ToString(dr("memo82")) <> "", Convert.ToString(dr("memo82")), "")

    End Sub

    '建立可選教師列表
    Sub SHOW_PLAN_TEACHER12()
        If g_flagNG Then
            sm.LastErrorMessage = cst_errmsg3
            Exit Sub
        End If
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return 'rst 'Exit Sub

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim rqRID As String = Convert.ToString(sm.UserInfo.RID)
        If RIDValue.Value <> "" Then rqRID = RIDValue.Value

        Dim pmsT1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"RID", rqRID}}
        Dim sqlT As String = ""
        sqlT &= " SELECT a.TechID" & vbCrLf '教師ID
        sqlT &= " ,a.TeachCName" & vbCrLf '教師姓名 
        sqlT &= " ,a.DegreeID" & vbCrLf '學歷
        sqlT &= " ,c.Name DegreeName" & vbCrLf '學歷
        '專業領域
        'sqlT &=" ,REPLACE(ISNULL(a.Specialty1, ' '),',',' ')" & vbCrLf
        'sqlT &=" + REPLACE(ISNULL(a.Specialty2, ' '),',',' ')" & vbCrLf
        'sqlT &=" + REPLACE(ISNULL(a.Specialty3, ' '),',',' ')" & vbCrLf
        'sqlT &=" + REPLACE(ISNULL(a.Specialty4, ' '),',',' ')" & vbCrLf
        'sqlT &=" + REPLACE(ISNULL(a.Specialty5, ' '),',',' ') major" & vbCrLf
        '專業領域 Specialty1
        sqlT &= " ,ISNULL(a.Specialty1, '') Specialty1" & vbCrLf
        '專業證照-相關證照
        sqlT &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2 ELSE a.ProLicense END ProLicense" & vbCrLf
        sqlT &= " ,dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'A', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
        sqlT &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sqlT &= " JOIN ( SELECT DISTINCT TechID, planid, comidno, seqno" & vbCrLf
        sqlT &= " FROM PLAN_TRAINDESC" & vbCrLf
        sqlT &= " WHERE TECHID IS NOT NULL AND PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO ) b ON a.TechID = b.TechID" & vbCrLf
        sqlT &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID = c.DegreeID" & vbCrLf
        'sqlT &=" WHERE a.WorkStatus='1'" & vbCrLf
        sqlT &= " WHERE a.RID=@RID" & vbCrLf
        sqlT &= " ORDER BY a.TechID" & vbCrLf
        Dim dtT As DataTable = DbAccess.GetDataTable(sqlT, objconn, pmsT1)
        i_gSeqno = 0
        tbDataGrid21.Visible = False
        If dtT.Rows.Count > 0 Then
            tbDataGrid21.Visible = True
            DataGrid21.DataSource = dtT
            DataGrid21.DataBind()
        End If

        Dim pmsT2 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"RID", rqRID}}
        Dim sqlT2 As String = ""
        sqlT2 &= " SELECT a.TechID" & vbCrLf '教師ID
        sqlT2 &= " ,a.TeachCName" & vbCrLf '教師姓名 
        sqlT2 &= " ,a.DegreeID" & vbCrLf '學歷
        sqlT2 &= " ,c.Name DegreeName" & vbCrLf '學歷
        'sqlT2 &= " ,REPLACE(ISNULL(a.Specialty1, ' '),',',' ')" & vbCrLf
        'sqlT2 &= " + REPLACE(ISNULL(a.Specialty2, ' '),',',' ')" & vbCrLf
        'sqlT2 &= " + REPLACE(ISNULL(a.Specialty3, ' '),',',' ')" & vbCrLf
        'sqlT2 &= " + REPLACE(ISNULL(a.Specialty4, ' '),',',' ')" & vbCrLf
        'sqlT2 &= " + REPLACE(ISNULL(a.Specialty5, ' '),',',' ') major" & vbCrLf
        '專業領域 Specialty1
        sqlT2 &= " ,ISNULL(a.Specialty1, '') Specialty1" & vbCrLf
        '專業證照-相關證照
        sqlT2 &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2 ELSE a.ProLicense END ProLicense" & vbCrLf
        sqlT2 &= " ,dbo.FN_GET_PLAN_TEACHER3(b.planid, b.comidno, b.seqno, 'B', a.TechID) TeacherDesc " 'TechTYPE: A:師資/B:助教
        sqlT2 &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sqlT2 &= " JOIN ( SELECT DISTINCT TECHID2 TechID, planid, comidno, seqno" & vbCrLf
        sqlT2 &= " FROM PLAN_TRAINDESC" & vbCrLf
        sqlT2 &= " WHERE TECHID2 IS NOT NULL AND PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO ) b ON a.TechID = b.TechID" & vbCrLf
        sqlT2 &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID = c.DegreeID" & vbCrLf 'sqlT2 &= " WHERE a.WorkStatus='1'" & vbCrLf
        sqlT2 &= " WHERE a.RID=@RID" & vbCrLf
        sqlT2 &= " ORDER BY a.TechID" & vbCrLf
        Dim dtT2 As DataTable = DbAccess.GetDataTable(sqlT2, objconn, pmsT2)
        i_gSeqno = 0
        tbDataGrid22.Visible = False
        If dtT2.Rows.Count > 0 Then
            tbDataGrid22.Visible = True
            DataGrid22.DataSource = dtT2
            DataGrid22.DataBind()
        End If
    End Sub

    '取得TrainPlace
    Sub SHOW_PLAN_TRAINPLACE() '(ByVal ComIDNO As String)
        '#Region "取得TrainPlace"
        If g_flagNG Then
            sm.LastErrorMessage = cst_errmsg3
            Exit Sub
        End If
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return 'rst 'Exit Sub

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim rqRID As String = Convert.ToString(sm.UserInfo.RID)
        If RIDValue.Value <> "" Then rqRID = RIDValue.Value

        Dim v_OtherMsg1 As String = ""
        Dim oParms As New Hashtable From {{"PTCOMIDNO", rqComIDNO}, {"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}}
        Dim objstr As String = ""
        objstr &= " SELECT b.connum ,b.hwdesc ,b.OtherDesc" & vbCrLf
        objstr &= " ,b.PTID ,b.PLACEID ,b.ClassIFICation" & vbCrLf
        objstr &= " FROM PLAN_PLANINFO a" & vbCrLf
        objstr &= " JOIN PLAN_TRAINPLACE b ON a.COMIDNO=b.COMIDNO AND a.SCIPLACEID=b.PLACEID AND b.COMIDNO=@PTCOMIDNO" & vbCrLf
        objstr &= " WHERE a.PLANID=@PLANID AND a.COMIDNO=@COMIDNO AND a.SEQNO=@SEQNO" & vbCrLf
        objstr &= " AND b.ClassIFICation IN (1,3)" & vbCrLf '學科共用。
        Dim dt1 As DataTable = DbAccess.GetDataTable(objstr, objconn, oParms)

        Dim v_OtherMsg2 As String = ""
        Dim oParms2 As New Hashtable From {{"PTCOMIDNO", rqComIDNO}, {"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}}
        Dim objstr2 As String = ""
        objstr2 &= " SELECT b.connum ,b.hwdesc ,b.OtherDesc" & vbCrLf
        objstr2 &= " ,b.PTID,b.PLACEID,b.ClassIFICation" & vbCrLf
        objstr2 &= " FROM PLAN_PLANINFO a" & vbCrLf
        objstr2 &= " JOIN PLAN_TRAINPLACE b ON a.COMIDNO=b.COMIDNO AND a.TECHPLACEID=b.PLACEID AND b.COMIDNO=@PTCOMIDNO" & vbCrLf
        objstr2 &= " WHERE a.PLANID=@PLANID AND a.COMIDNO=@COMIDNO AND a.SEQNO=@SEQNO" & vbCrLf
        objstr2 &= " AND b.ClassIFICation IN (2,3)" & vbCrLf '術科共用。
        Dim dt2 As DataTable = DbAccess.GetDataTable(objstr2, objconn, oParms2)

        T2Dnum2.Enabled = False
        HwDesc2.Enabled = False
        If dt1.Rows.Count > 0 Then
            Dim dr1 As DataRow = dt1.Rows(0)
            T2Dnum2.Text = TIMS.ClearSQM(dr1("connum"))
            HwDesc2.Text = If(T2Dnum2.Text <> "", Convert.ToString(dr1("hwdesc")), "")
            v_OtherMsg1 = Convert.ToString(dr1("OtherDesc"))
        End If
        'If v_OtherMsg <> "" Then v_OtherMsg &= vbCrLf '換行

        T2Dnum3.Enabled = False
        HwDesc3.Enabled = False
        If dt2.Rows.Count > 0 Then
            Dim dr2 As DataRow = dt2.Rows(0)
            T2Dnum3.Text = TIMS.ClearSQM(dr2("connum"))
            HwDesc3.Text = If(T2Dnum3.Text <> "", Convert.ToString(dr2("hwdesc")), "")
            v_OtherMsg2 = Convert.ToString(dr2("OtherDesc"))
        End If

        Dim v_OtherMsgA As String = ""
        If v_OtherMsg1 <> "" AndAlso v_OtherMsg2 <> "" Then
            v_OtherMsgA = String.Concat(v_OtherMsg1, If(v_OtherMsg1 <> v_OtherMsg2, String.Concat(vbCrLf, v_OtherMsg2), ""))
        ElseIf v_OtherMsg1 <> "" AndAlso v_OtherMsg1.Length > 1 Then
            v_OtherMsgA = v_OtherMsg1
        ElseIf v_OtherMsg2 <> "" AndAlso v_OtherMsg2.Length > 1 Then
            v_OtherMsgA = v_OtherMsg2
        End If
        '術科
        'OtherDesc23.Enabled = False'(欄位另存 PLAN_VERREPORT.OTHFACDESC23)
        If txtOthFacDesc23.Text = "" AndAlso v_OtherMsgA <> "" Then txtOthFacDesc23.Text = TIMS.ClearSQM2(v_OtherMsgA)
    End Sub

    '取得資訊-政策性產業課程可辦理班數-PLAN_PRECLASS
    Function GET_PRECLASS_PCNT1(ByVal rqPlanID As String, ByVal rqComIDNO As String, ByVal rqSeqNO As String, ByVal vYEARS As Integer, ByRef iPREID As Integer) As Integer
        Dim iRst As Integer = 0

        'pParms.Clear()
        Dim pParms As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"YEARS", vYEARS}}
        Dim sql As String = ""
        sql &= " SELECT PREID,PCNT1"
        sql &= " FROM PLAN_PRECLASS"
        sql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND YEARS=@YEARS"

        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, pParms)
        If dr Is Nothing Then Return iRst
        iPREID = TIMS.CINT1(dr("PREID"))
        iRst = TIMS.CINT1(dr("PCNT1"))
        Return iRst
    End Function

#Region "NOUSE"
    '新增／修改-政策性產業課程可辦理班數
    'Sub SHOW_DG33_PRECLASSCNT(ByVal iType As Integer)
    '    'iType 0:新增 1:修改
    '    Const cst_PREC_PREID As String = "PREID"
    '    Const cst_PREC_YEARS As String = "YEARS"
    '    Const cst_PREC_PCNT1 As String = "PCNT1"
    '    Dim dt As New DataTable
    '    dt.Columns.Add(New DataColumn("PREID", System.Type.GetType("System.Int32")))
    '    dt.Columns.Add(New DataColumn("YEARS"))
    '    dt.Columns.Add(New DataColumn("PCNT1", System.Type.GetType("System.Int32")))
    '    Dim i_Y1 As Integer = Val(sm.UserInfo.Years) '今年
    '    Dim i_Y3 As Integer = Val(sm.UserInfo.Years) + 2 '後年

    '    Select Case iType
    '        Case 0
    '            For iY As Integer = i_Y1 To i_Y3
    '                Dim dr As DataRow = dt.NewRow()
    '                dt.Rows.Add(dr)
    '                dr(cst_PREC_PREID) = 0
    '                dr(cst_PREC_YEARS) = CStr(iY)
    '                dr(cst_PREC_PCNT1) = Convert.DBNull
    '            Next

    '        Case Else '1
    '            Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
    '            Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
    '            Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))

    '            For iY As Integer = i_Y1 To i_Y3
    '                Dim iPREID As Integer = 0
    '                Dim vPCNT1 As Integer = GET_PRECLASS_PCNT1(rqPlanID, rqComIDNO, rqSeqNO, iY, iPREID) '年可辦理班數
    '                Dim dr As DataRow = dt.NewRow()
    '                dt.Rows.Add(dr)
    '                dr(cst_PREC_PREID) = iPREID
    '                dr(cst_PREC_YEARS) = CStr(iY) ' - 1911)
    '                dr(cst_PREC_PCNT1) = vPCNT1
    '            Next

    '    End Select
    '    dt.AcceptChanges()

    '    'Dim Sql As String = ""
    '    'Dim dt As DataTable = DbAccess.GetDataTable(Sql,)
    '    DataGrid33.DataSource = dt
    '    DataGrid33.DataBind()
    'End Sub
#End Region

    '檢查 班級申請老師
    Function CHK_PLAN_TEACHER12(ByRef errmsg As String) As Boolean
        '#Region "檢查 班級申請老師"

        Dim rst As Boolean = True
        Const Cst_授課教師限制數 As Integer = 0 '10 '0:無限制

        'Select Case rqProcessType 'ProcessType @Insert/Update/View
        '    Case cst_ptInsert, cst_ptUpdate
        'End Select

        Dim i As Integer = 0
        Dim errT As String = ""
        Dim errI2 As Integer = 0
        For Each eItem As DataGridItem In DataGrid21.Items
            'Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            Dim seqno As Label = eItem.FindControl("seqno")
            Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim Specialty1 As Label = eItem.FindControl("Specialty1")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA") 'TechTYPE: A:師資/B:助教
            i += 1
            If TeacherDesc.Text = "" Then
                errT = String.Concat(seqno.Text, ":", TeachCName.Text)
                errI2 += 1
                Exit For
            End If
        Next
        If i = 0 Then
            errmsg &= "至少選擇1筆授課教師" & vbCrLf
            Return False
        End If
        If errI2 > 0 Then
            errmsg &= String.Concat("授課教師-", errT, "-遴選辦法說明辦法為必填", vbCrLf)
            Return False
        End If
        If Cst_授課教師限制數 <> 0 Then '0:無限制
            If Not (i <= Cst_授課教師限制數) Then
                errmsg &= String.Concat("僅可選擇", Cst_授課教師限制數, "筆授課教師", vbCrLf)
                Return False
            End If
        End If

        Dim errTB As String = ""
        Dim errI2B As Integer = 0
        For Each eItem As DataGridItem In DataGrid22.Items
            'Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            Dim seqno As Label = eItem.FindControl("seqno")
            Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim Specialty1 As Label = eItem.FindControl("Specialty1")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEB As HtmlInputButton = eItem.FindControl("btn_TCTYPEB") 'TechTYPE: A:師資/B:助教
            If TeacherDesc.Text = "" Then
                errTB = String.Concat(seqno.Text, ":", TeachCName.Text)
                errI2B += 1
                Exit For
            End If
        Next
        If errI2B > 0 Then
            errmsg &= String.Concat("授課助教-", errTB, "-遴選辦法說明辦法為必填", vbCrLf)
            Return False
        End If

        Return rst
    End Function

    ''' <summary> 正式儲存檢核 </summary>
    ''' <param name="ErrMsg"></param>
    ''' <returns></returns>
    Function CheckData14(ByRef ErrMsg As String) As Boolean
        '#Region "CheckData1"
        Dim rst As Boolean = False

        RecDesc.Text = TIMS.ClearSQM(RecDesc.Text)
        LearnDesc.Text = TIMS.ClearSQM(LearnDesc.Text)
        ActDesc.Text = TIMS.ClearSQM(ActDesc.Text)
        ResultDesc.Text = TIMS.ClearSQM(ResultDesc.Text)
        OtherDesc.Text = TIMS.ClearSQM(OtherDesc.Text)

        Hid_ComIDNO.Value = TIMS.ClearSQM(Hid_ComIDNO.Value)
        If ComidValue.Value = "" Then ComidValue.Value = Hid_ComIDNO.Value

        Select Case Hid_sender1.Value '(檢核儲存階段)
            Case cst_SaveDef '草稿儲存(不檢查)
                Return True
        End Select

        '正式檢測。
        Dim iALL_PHour As Integer = 0 '總時數(課程大綱)
        Dim i_FARLEARN_PHours As Double = 0 '遠距教學時數
        If TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then
            Dim dt As DataTable = Session(hid_TrainDescTable_guid1.Value)
            If dt.Rows.Count > 0 Then
                For i As Int16 = 0 To dt.Rows.Count - 1
                    If Not dt.Rows(i).RowState = DataRowState.Deleted Then
                        iALL_PHour += TIMS.CINT1(dt.Rows(i)("PHour")) '總時數

                        '遠距教學
                        If Convert.ToString(dt.Rows(i)("FARLEARN")).Equals("Y") Then
                            i_FARLEARN_PHours += Val(dt.Rows(i)("PHour")) '遠距教學時數
                        End If
                    End If
                Next
            End If
        End If
        '辦理方式:null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        'Dim vrbl_DISTANCE As String = TIMS.GetListValue(rbl_DISTANCE)
        '"[班別資料]「遠距教學總時數不得超過本班總訓練時數2/3」。" '-正式
        If Val(i_FARLEARN_PHours) > 0 AndAlso Val(iALL_PHour) > 0 Then
            Dim d_FARLEARN_PHours As Double = CDbl(i_FARLEARN_PHours)
            Dim d_iALL_PHour23 As Double = CDbl(TIMS.ROUND(CDbl(iALL_PHour) * 2 / 3, 2))
            If d_FARLEARN_PHours > d_iALL_PHour23 Then
                ErrMsg &= String.Format(String.Concat(cst_errmsg38, ",,遠距教學總時數：{0}, 訓練總時數2/3：{1}"), d_FARLEARN_PHours, d_iALL_PHour23) & vbCrLf
                Return False
            End If
        End If

        iCAPNUM.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(iCAPNUM.Text))
        iCAPMARKDATE.Text = TIMS.ClearSQM(iCAPMARKDATE.Text)
        If iCAPMARKDATE.Text <> "" AndAlso Not TIMS.IsDate1(iCAPMARKDATE.Text) Then
            ErrMsg &= "班別資料-【iCAP標章有效期限】有填寫, 日期格式有誤!" & vbCrLf
        ElseIf iCAPMARKDATE.Text <> "" Then
            iCAPMARKDATE.Text = TIMS.Cdate3(iCAPMARKDATE.Text)
        End If
        '正式儲存的時候，系統檢核若是【iCap標章證號】有填寫，【iCAP標章有效期限】不可為空，跳出提示訊息
        Dim fg_CAN_CHECK_iCAPNUM_R1 As Boolean = (iCAPNUM.Text <> "" AndAlso iCAPMARKDATE.Text <> "" AndAlso FDDate.Text <> "" AndAlso TIMS.IsDate1(iCAPMARKDATE.Text) AndAlso TIMS.IsDate1(FDDate.Text))
        If (iCAPNUM.Text <> "") AndAlso iCAPMARKDATE.Text = "" Then
            ErrMsg &= String.Concat("班別資料-【iCap標章證號】有填寫，【iCAP標章有效期限】不可為空!") & vbCrLf
        ElseIf iCAPMARKDATE.Text <> "" AndAlso iCAPNUM.Text = "" Then
            ErrMsg &= String.Concat("班別資料-【iCAP標章有效期限】有填寫，【iCap標章證號】不可為空!") & vbCrLf
        ElseIf fg_CAN_CHECK_iCAPNUM_R1 AndAlso DateDiff(DateInterval.Day, CDate(iCAPMARKDATE.Text), CDate(FDDate.Text)) > 0 Then
            ErrMsg &= String.Concat("班別資料-iCAP標章有效期限須涵蓋完整訓練期間!") & vbCrLf
        End If
        '檢核-iCAP標章證號 true:OK false:已被其它單位使用
        If iCAPNUM.Text <> "" AndAlso Not ChkiCAPNUM(iCAPNUM.Text, ComidValue.Value) Then
            ErrMsg &= "「iCAP標章證號」已被其它單位使用!!(單位不同，不可使用相同的編碼), " & iCAPNUM.Text & vbCrLf
        End If

        Dim v_KID20 As String = GET_KID20_VAL() '2019 政府政策性產業
        Dim v_CBLKID25 As String = GET_KID25_VAL() '2025 政府政策性產業
        Dim v_CBLKID26 As String = GET_KID26_VAL() '2026 政府政策性產業
        '進階政策性產業類別/B
        Dim v_CBLKID22 As String = If(trKID25.Visible, TIMS.GetCblValue(CBLKID22B), TIMS.GetCblValue(CBLKID22)) 'v_CBLKID22
        Dim fg_KID20222526 As Boolean = ($"{v_KID20}{v_CBLKID22}{v_CBLKID25}{v_CBLKID26}" <> "") '有勾選政府政策性產業

        'Dim fg_KID222_HaveVAL As Boolean = (fg_KID2225 OrElse v_CBLKID22 <> "") '進階政策性產業類別/B
        '有勾選政府政策性產業，須填寫「與政策性產業課程之關聯性概述」
        tPOLICYREL.Text = TIMS.ClearSQM(tPOLICYREL.Text)
        Dim V_tPOLICYREL As String = TIMS.Get_placeholder_TXTVAL(tPOLICYREL)
        'If flag_OJT22071401 Then End If
        If fg_KID20222526 AndAlso V_tPOLICYREL = "" Then
            ErrMsg &= String.Concat("有勾選政府政策性產業，須填寫 訓練需求調查-「與政策性產業課程之關聯性概述」") & vbCrLf
        ElseIf Not fg_KID20222526 AndAlso V_tPOLICYREL <> "" Then
            '(有勾選政府政策性產業，須填寫「與政策性產業課程之關聯性概述」，若無勾選政府政策性產業，則不儲存此欄資料)
            If (V_tPOLICYREL <> "") Then Hid_tPOLICYREL.Value = V_tPOLICYREL
            tPOLICYREL.Text = ""
        End If

        '申請階段／'申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇) 
        Dim v_rbl_AppStage As String = TIMS.GetListValue(rbl_AppStage)
        If v_rbl_AppStage = "3" AndAlso $"{v_CBLKID25}{v_CBLKID26}" = "" Then
            ErrMsg &= "申請階段為「政策性產業」必須勾選任一政府政策性產業!" & vbCrLf
        End If

        '三、基本儲存、正式儲存卡控：充電起飛計畫不用這些卡控
        '1、「AI加值應用、職場續航」僅限申請階段為「政策性產業」，當單位有勾選，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」，若未填寫須出現提示訊息，且不能儲存。
        '2、當勾選「AI加值應用」時，須卡控AI應用時數之總數須等於或大於12小時，且不可超過總訓練時數1/2，若不符須出現提示訊息，且不能儲存。
        '3、當勾選「職場續航」時，須卡控職場新續航時數之總數須等於6小時，不可小於或大於，若不符須出現提示訊息，且不能儲存。
        '4、因單位可能同時勾選「AI加值應用」及「職場新續航」，且同一堂課可能會有AI應用時數及職場新續航時數，各自時數仍須分開計算。
        '5、因政策性產業新增「職場新續航」項目，此類政策性產業有規定特定訓練業別才能申請（詳附件檔案），須卡控班級申請之訓練業別代碼倘非規定之代碼，則跳出提示訊息【申請之訓練業別非「職場新續航」課程職類】且不可儲存。
        Dim v_CBLKID25_7 As String = TIMS.GetCblValue(CBLKID25_7)
        Dim v_CBLKID25_8 As String = TIMS.GetCblValue(CBLKID25_8)
        If v_CBLKID25_7 <> "" AndAlso V_tPOLICYREL = "" Then
            ErrMsg &= "勾選「AI加值應用」，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」!" & vbCrLf
        End If
        If v_CBLKID25_8 <> "" AndAlso V_tPOLICYREL = "" Then
            ErrMsg &= "勾選「職場續航」，須於開班計劃表資料維護頁籤填寫「與政策性產業課程之關聯性概述：」!" & vbCrLf
        End If

        If RecDesc.Text = "" AndAlso chk_RecDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 反應評估，請輸入內容" & vbCrLf
        If LearnDesc.Text = "" AndAlso chk_LearnDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 學習評估，請輸入內容" & vbCrLf
        If ActDesc.Text = "" AndAlso chk_ActDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 行為評估，請輸入內容" & vbCrLf
        If ResultDesc.Text = "" AndAlso chk_ResultDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 成果評估，請輸入內容" & vbCrLf
        If OtherDesc.Text = "" AndAlso chk_OtherDesc.Checked Then ErrMsg &= "四、訓練績效評估-勾選 其他機制，請輸入內容" & vbCrLf
        If ErrMsg <> "" Then Return False

        If RecDesc.Text <> "" AndAlso Not chk_RecDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 反應評估，請勿輸入內容" & vbCrLf
        If LearnDesc.Text <> "" AndAlso Not chk_LearnDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 學習評估，請勿輸入內容" & vbCrLf
        If ActDesc.Text <> "" AndAlso Not chk_ActDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 行為評估，請勿輸入內容" & vbCrLf
        If ResultDesc.Text <> "" AndAlso Not chk_ResultDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 成果評估，請勿輸入內容" & vbCrLf
        If OtherDesc.Text <> "" AndAlso Not chk_OtherDesc.Checked Then ErrMsg &= "四、訓練績效評估-未勾選 其他機制，請勿輸入內容" & vbCrLf
        If ErrMsg <> "" Then Return False

        Dim i_chk2 As Integer = 0
        If chk_RecDesc.Checked Then i_chk2 += 1
        If chk_LearnDesc.Checked Then i_chk2 += 1
        If chk_ActDesc.Checked Then i_chk2 += 1
        If chk_ResultDesc.Checked Then i_chk2 += 1
        If chk_OtherDesc.Checked Then i_chk2 += 1
        If i_chk2 = 0 Then ErrMsg &= "四、訓練績效評估-未勾選 (至少要勾選一項)" & vbCrLf
        If ErrMsg <> "" Then Return False

        Dim flagCPT As Boolean = CHK_PLAN_TEACHER12(ErrMsg)
        If Not flagCPT Then Return False

        Dim vTMethod As String = TIMS.GetCblValue(cblTMethod)
        If vTMethod.IndexOf("99") > -1 AndAlso TMethodOth.Text = "" Then ErrMsg &= "教學方法-若選「其他教學方法」，需填寫輸入其它說明，上限100個字" & vbCrLf
        '課程須符合目的事業主管機關相關規定
        If cbPOWERNEED4.Checked AndAlso tPOWERNEED4.Text = "" Then ErrMsg &= "訓練需求調查-若勾選「課程須符合目的事業主管機關相關規定」，需填寫它說明，上限200個字" & vbCrLf
        '本課程是否應報請主管機關核備 報請主管機關核備
        Dim v_rbl_REPORTE As String = If(rbl_REPORTE_Y.Checked, "Y", If(rbl_REPORTE_N.Checked, "N", If(Hid_REPORTE.Value <> "", Hid_REPORTE.Value, "")))
        If v_rbl_REPORTE = "" Then ErrMsg &= "訓練需求調查-本課程是否應報請主管機關核備 為必選,請選擇" & vbCrLf
        '職能級別
        Dim v_rblFuncLevel As String = TIMS.GetListValue(rblFuncLevel)
        If v_rblFuncLevel = "" Then ErrMsg &= "訓練目標-職能級別：(單選) 為必選,請選擇" & vbCrLf

        Dim v_Degree As String = TIMS.GetListValue(Degree)
        If v_Degree = "" Then ErrMsg &= "受訓資格 學歷 為必選,請選擇" & vbCrLf
        '不管什麼都是「年滿15歲以上」。   'Const cst_ageoDef As Integer = 16 'other Years Start
        txtAge1.Text = TIMS.ClearSQM(txtAge1.Text)
        If Not rdoAge1.Checked AndAlso Not rdoAge2.Checked Then ErrMsg &= "受訓資格 年齡 為必選,請選擇" & vbCrLf
        'If ErrMsg <> "" Then Exit Sub 'Return False '有錯誤訊息,不可儲存

        If rdoAge2.Checked Then
            If txtAge1.Text = "" Then
                ErrMsg &= "受訓資格 年齡選項2 未輸入有效年齡" & vbCrLf
            Else
                If Not TIMS.IsNumeric2(txtAge1.Text) Then ErrMsg &= "請檢查 受訓資格 年齡選項2 未輸入有效年齡:" & txtAge1.Text & vbCrLf
                If ErrMsg = "" Then
                    If TIMS.CINT1(txtAge1.Text) < cst_AgeOtherDef Then ErrMsg &= "請檢查 受訓資格 年齡選項2 有效年齡(須大於15歲(不含)以上)" & vbCrLf
                    If TIMS.CINT1(txtAge1.Text) > 99 Then ErrMsg &= "請檢查 受訓資格 年齡選項2 有效年齡(須小於99歲(含)以下)" & vbCrLf
                End If
            End If
        End If

        If CapAll.Text = "" Then ErrMsg &= "受訓資格 學員資格 未輸入，請確認填寫" & vbCrLf

        Dim i_MaxTxtLen1 As Integer = 0

        '【是否為iCAP課程】如果是、否都沒選，
        If Not RB_ISiCAPCOUR_Y.Checked AndAlso Not RB_ISiCAPCOUR_N.Checked Then ErrMsg &= cst_errmsg32 & vbCrLf
        iCAPCOURDESC.Text = TIMS.TrimValue(iCAPCOURDESC.Text) '課程相關說明
        If RB_ISiCAPCOUR_Y.Checked AndAlso iCAPCOURDESC.Text = "" Then
            ErrMsg &= cst_errmsg33 & vbCrLf '課程相關說明
        End If
        '超過欄位字數
        If iCAPCOURDESC.Text <> "" Then
            i_MaxTxtLen1 = 500
            If i_MaxTxtLen1 > 0 AndAlso (iCAPCOURDESC.Text.Length > i_MaxTxtLen1) Then ErrMsg &= String.Format("促進學習機制-是否為iCAP課程-課程相關說明(欄位字數為{0})，超過欄位字數", i_MaxTxtLen1) & vbCrLf
        End If
        '檢核-iCAP標章證號
        If RB_ISiCAPCOUR_Y.Checked AndAlso iCAPNUM.Text = "" Then
            ErrMsg &= cst_errmsg33Y & vbCrLf 'iCAP標章證號
        ElseIf RB_ISiCAPCOUR_N.Checked AndAlso iCAPNUM.Text <> "" Then
            ErrMsg &= cst_errmsg33N & vbCrLf 'iCAP標章證號
        End If

        Recruit.Text = TIMS.TrimValue(Recruit.Text)
        Selmethod.Text = TIMS.TrimValue(Selmethod.Text)
        Inspire.Text = TIMS.TrimValue(Inspire.Text)
        If Recruit.Text = "" Then ErrMsg &= "未輸入 招訓方式，請確認填寫" & vbCrLf
        If Selmethod.Text = "" Then ErrMsg &= "未輸入 遴選方式，請確認填寫" & vbCrLf
        If Inspire.Text = "" Then ErrMsg &= "未輸入 學員激勵辦法，請確認填寫" & vbCrLf

        i_MaxTxtLen1 = 2000
        If i_MaxTxtLen1 > 0 AndAlso (Recruit.Text.Length > i_MaxTxtLen1) Then ErrMsg &= String.Format("招訓方式(欄位字數為{0})，超過欄位字數", i_MaxTxtLen1) & vbCrLf
        i_MaxTxtLen1 = 2000
        If i_MaxTxtLen1 > 0 AndAlso (Selmethod.Text.Length > i_MaxTxtLen1) Then ErrMsg &= String.Format("遴選方式(欄位字數為{0})，超過欄位字數", i_MaxTxtLen1) & vbCrLf
        i_MaxTxtLen1 = 2000
        If i_MaxTxtLen1 > 0 AndAlso (Inspire.Text.Length > i_MaxTxtLen1) Then ErrMsg &= String.Format("學員激勵辦法(欄位字數為{0})，超過欄位字數", i_MaxTxtLen1) & vbCrLf
        Dim I_TGovExam As Integer = 0
        If (TGovExamCY.Checked) Then I_TGovExam += 1
        If (TGovExamCN.Checked) Then I_TGovExam += 1
        If (TGovExamCG.Checked) Then I_TGovExam += 1
        If I_TGovExam > 1 Then ErrMsg &= "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定，請擇一勾選！" & vbCrLf
        If TGovExamCY.Checked AndAlso GOVAGENAME.Text = "" Then
            ErrMsg &= "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定，若為是。政府機關名稱不可為空" & vbCrLf
        End If
        If TGovExamCY.Checked AndAlso TGovExamName.Text = "" Then
            ErrMsg &= "學員是否可依個人需求參加政府機關辦理相關證照考試或技能檢定，若為是。證照或檢定名稱不可為空" & vbCrLf
        End If

        '專長能力標籤-ABILITY
        For i_SEQ As Integer = 1 To 4
            Dim s_SEQ As String = Convert.ToString(i_SEQ)
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            otxtA1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA1.Text), 30)
            otxtA2.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA2.Text), 200)
            If otxtA1.Text = "" Then
                ErrMsg &= cst_errmsg36 & vbCrLf
                Exit For
            End If
            Exit For
        Next

        '一人份材料明細
        Dim iPersonCost_Rows As Integer = 0
        Dim dtPersonCost As DataTable = If(Session(hid_PersonCostTable_guid1.Value), Nothing)
        If dtPersonCost IsNot Nothing AndAlso dtPersonCost.Rows.Count > 0 Then iPersonCost_Rows = dtPersonCost.Rows.Count
        '共同材料明細
        Dim iCommonCost_Rows As Integer = 0
        Dim dtCommonCost As DataTable = If(Session(hid_CommonCostTable_guid1.Value), Nothing)
        If dtCommonCost IsNot Nothing AndAlso dtCommonCost.Rows.Count > 0 Then iCommonCost_Rows = dtCommonCost.Rows.Count
        '"[訓練費用] 填寫 材料費用總額 大於0，須填寫「一人份材料明細」或「共同材料明細」資料。" '-正式
        If (METSUMCOST.Text <> "" AndAlso Val(METSUMCOST.Text) > 0) AndAlso iPersonCost_Rows = 0 AndAlso iCommonCost_Rows = 0 Then
            ErrMsg &= cst_errmsg39 & vbCrLf
            Return rst
        End If

        If ErrMsg = "" Then rst = True
        Return rst
    End Function

    '儲存 班級申請老師-PLAN_TEACHER
    Sub SAVE_PLAN_TEACHER(ByVal tConn As SqlConnection)
        '#Region "儲存 班級申請老師"
        If upt_PlanX.Value = "" Then Return
        tmpPCS = upt_PlanX.Value  '有儲存資料過了
        PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
        ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
        SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
        'If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
        Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
        Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
        Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return '(有異常離開)

        'Dim DATA_YN As String = TIMS.Get_PLAN_VERREPORT(rqPlanID, rqComIDNO, rqSeqNO, TIMS.cst_PVR_DATA_YN, objconn)
        'Dim rqProcessType As String = cst_ptInsert  'ProcessType @Insert/Update/View
        'If DATA_YN = "Y" Then rqProcessType = cst_ptUpdate

        Dim dParms1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}}
        Dim dSql As String = "DELETE PLAN_TEACHER WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        DbAccess.ExecuteNonQuery(dSql, tConn, dParms1)

        Dim iSql As String = ""
        iSql &= " INSERT INTO PLAN_TEACHER (PLANID,COMIDNO,SEQNO,TECHID,TECHTYPE,TEACHERDESC,MODIFYACCT,MODIFYDATE)" & vbCrLf
        iSql &= " VALUES (@PLANID,@COMIDNO,@SEQNO,@TECHID,@TECHTYPE,@TEACHERDESC,@MODIFYACCT,GETDATE())" & vbCrLf

        Dim sSql1 As String = ""
        sSql1 = " SELECT 1 FROM PLAN_TEACHER WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND TECHID=@TECHID AND TECHTYPE=@TECHTYPE" & vbCrLf

        'Dim tTECHTYPE As String = "" 'TechTYPE: A:師資/B:助教
        '更新師資表 'TechTYPE: A:師資/B:助教
        Const cst_tTECHTYPE_A As String = "A"
        Const cst_tTECHTYPE_B As String = "B"

        For Each eItem As DataGridItem In DataGrid21.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label = eItem.FindControl("seqno")
            'Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim Specialty1 As Label = eItem.FindControl("Specialty1")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable From {{"PLANID", TIMS.CINT1(rqPlanID)}, {"COMIDNO", rqComIDNO}, {"SEQNO", TIMS.CINT1(rqSeqNO)}}
                sParms.Add("TECHID", TIMS.CINT1(HidTechID.Value)) 'dr("TECHID"))
                sParms.Add("TECHTYPE", cst_tTECHTYPE_A) 'TechTYPE: A:師資/B:助教
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iParms As New Hashtable From {{"PLANID", TIMS.CINT1(rqPlanID)}, {"COMIDNO", rqComIDNO}, {"SEQNO", TIMS.CINT1(rqSeqNO)}}
                    iParms.Add("TECHID", TIMS.CINT1(HidTechID.Value)) 'dr("TECHID"))
                    iParms.Add("TECHTYPE", cst_tTECHTYPE_A) 'TechTYPE: A:師資/B:助教
                    iParms.Add("TEACHERDESC", tTEACHERDESC)
                    iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(iSql, tConn, iParms)
                End If
            End If
        Next

        'tTECHTYPE = "B" 'TechTYPE: A:師資/B:助教
        For Each eItem As DataGridItem In DataGrid22.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label = eItem.FindControl("seqno")
            'Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim Specialty1 As Label = eItem.FindControl("Specialty1")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEB As HtmlInputButton = eItem.FindControl("btn_TCTYPEB")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable From {{"PLANID", TIMS.CINT1(rqPlanID)}, {"COMIDNO", rqComIDNO}, {"SEQNO", TIMS.CINT1(rqSeqNO)}}
                sParms.Add("TECHID", TIMS.CINT1(HidTechID.Value)) 'dr("TECHID"))
                sParms.Add("TECHTYPE", cst_tTECHTYPE_B) 'TechTYPE: A:師資/B:助教
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iParms As New Hashtable From {{"PLANID", TIMS.CINT1(rqPlanID)}, {"COMIDNO", rqComIDNO}, {"SEQNO", TIMS.CINT1(rqSeqNO)}}
                    iParms.Add("TECHID", TIMS.CINT1(HidTechID.Value)) 'dr("TECHID"))
                    iParms.Add("TECHTYPE", cst_tTECHTYPE_B) 'TechTYPE: A:師資/B:助教
                    iParms.Add("TEACHERDESC", tTEACHERDESC)
                    iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(iSql, tConn, iParms)
                End If
            End If
        Next
    End Sub

    '儲存 班級申請老師(CLASS_TEACHER)
    Sub SAVE_CLASS_TEACHER(ByVal iOCID As Integer, ByVal tConn As SqlConnection)
        '#Region "儲存 班級申請老師(CLASS_TEACHER)"
        If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
        'tmpPCS = upt_PlanX.Value  '有儲存資料過了
        'PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
        'ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
        'SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
        'Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
        'Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
        'Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")

        Dim dParms As New Hashtable From {{"OCID", iOCID}}
        Dim dSql As String = "DELETE CLASS_TEACHER WHERE OCID =@OCID"
        DbAccess.ExecuteNonQuery(dSql, tConn, dParms)

        Dim iSqlc As String = ""
        iSqlc &= " INSERT INTO CLASS_TEACHER (CTRID ,OCID,TECHID,MODIFYACCT,MODIFYDATE,TECHTYPE,TEACHERDESC)" & vbCrLf
        iSqlc &= " VALUES (@CTRID ,@OCID,@TECHID,@MODIFYACCT,GETDATE(),@TECHTYPE,@TEACHERDESC )" & vbCrLf

        Dim sSql1 As String = ""
        sSql1 = " SELECT 1 FROM CLASS_TEACHER WHERE OCID=@OCID AND TECHID=@TECHID AND TECHTYPE=@TECHTYPE" & vbCrLf

        '更新師資表 'TechTYPE: A:師資/B:助教
        Const cst_tTECHTYPE_A As String = "A"
        Const cst_tTECHTYPE_B As String = "B"
        For Each eItem As DataGridItem In DataGrid21.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label = eItem.FindControl("seqno")
            'Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim Specialty1 As Label = eItem.FindControl("Specialty1")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable From {
                    {"OCID", iOCID},
                    {"TECHID", TIMS.CINT1(HidTechID.Value)}, 'dr("TECHID"))
                    {"TECHTYPE", cst_tTECHTYPE_A} 'TechTYPE: A:師資/B:助教
                    }
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(tConn, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    Dim iParms As New Hashtable From {
                        {"CTRID", iCTRID},
                        {"OCID", iOCID},
                        {"TECHID", TIMS.CINT1(HidTechID.Value)}, 'dr("TECHID"))
                        {"MODIFYACCT", sm.UserInfo.UserID},
                        {"TECHTYPE", cst_tTECHTYPE_A},
                        {"TEACHERDESC", tTEACHERDESC}
                    }
                    DbAccess.ExecuteNonQuery(iSqlc, tConn, iParms)
                End If
            End If
        Next

        'TechTYPE: A:師資/B:助教 'tTECHTYPE = "B"
        For Each eItem As DataGridItem In DataGrid22.Items
            Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
            'Dim seqno As Label = eItem.FindControl("seqno")
            'Dim TeachCName As Label = eItem.FindControl("TeachCName")
            'Dim DegreeName As Label = eItem.FindControl("DegreeName")
            'Dim Specialty1 As Label = eItem.FindControl("Specialty1")
            Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
            'Dim btn_TCTYPEA As HtmlInputButton = eItem.FindControl("btn_TCTYPEA")
            Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
            If HidTechID.Value <> "" Then
                Dim sParms As New Hashtable From {
                    {"OCID", iOCID},
                    {"TECHID", TIMS.CINT1(HidTechID.Value)}, 'dr("TECHID"))
                    {"TECHTYPE", cst_tTECHTYPE_B} 'TechTYPE: A:師資/B:助教
                    }
                Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, objconn, sParms)
                If dr1 Is Nothing Then
                    Dim iCTRID As Integer = DbAccess.GetNewId(tConn, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                    Dim iParms As New Hashtable From {
                        {"CTRID", iCTRID},
                        {"OCID", iOCID},
                        {"TECHID", TIMS.CINT1(HidTechID.Value)}, 'dr("TECHID"))
                        {"MODIFYACCT", sm.UserInfo.UserID},
                        {"TECHTYPE", cst_tTECHTYPE_B},
                        {"TEACHERDESC", tTEACHERDESC}
                    }
                    DbAccess.ExecuteNonQuery(iSqlc, tConn, iParms)
                End If
            End If
        Next
        '更新師資表 -End
    End Sub

    '儲存 開班計畫表資料維護 INSERT PLAN_VERREPORT,SAVE PLAN_VERREPORT,UPDATE PLAN_VERREPORT
    Sub SAVE_PLAN_VERREPORT(ByVal SaveType1 As String) 'ByRef htSS As Hashtable)
        '#Region "儲存 開班計畫表資料維護"
        'Dim SaveType1 As String = TIMS.GetMyValue2(htSS, "SaveType1")
        If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
        tmpPCS = upt_PlanX.Value  '有儲存資料過了
        PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
        ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
        SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
        Dim flag_error_1 As Boolean = (PlanID_value = "" OrElse ComIDNO_value = "" OrElse SeqNO_value = "")
        If flag_error_1 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
        Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
        Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")
        Dim DATA_YN As String = TIMS.Get_Plan_VerReport(rqPlanID, rqComIDNO, rqSeqNO, TIMS.cst_PVR_DATA_YN, objconn)
        Dim rqProcessType As String = cst_ptInsert  'ProcessType @Insert/Update/View
        If DATA_YN = "Y" Then rqProcessType = cst_ptUpdate

        Dim sql As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing

        Dim vTMethod As String = TIMS.GetCblValue(cblTMethod)
        Dim v_CBLKID25_8 As String = TIMS.GetCblValue(CBLKID25_8) '職場續航
        If v_CBLKID25_8 <> "" Then CapAll.Text = cst_D25_8_CapAll_MSG

        Dim iPVID As Integer = 0
        Select Case rqProcessType 'ProcessType @Insert/Update/View
            Case cst_ptInsert
                sql = $" SELECT * FROM PLAN_VERREPORT WHERE PlanID={TIMS.CINT1(rqPlanID)} AND ComIDNO='{rqComIDNO}' AND SeqNo={TIMS.CINT1(rqSeqNO)}"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count = 0 Then
                    iPVID = DbAccess.GetNewId(objconn, "PLAN_VERREPORT_PVID_SEQ,PLAN_VERREPORT,PVID")
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("PVID") = iPVID
                    dr("PlanID") = rqPlanID
                    dr("ComIDNO") = rqComIDNO
                    dr("SeqNo") = rqSeqNO
                ElseIf dt.Rows.Count = 1 Then
                    '新增 卻有資料
                    dr = dt.Rows(0)
                    iPVID = dt.Rows(0)("PVID")
                ElseIf dt.Rows.Count > 1 Then
                    sm.LastErrorMessage = cst_errmsg26 '"儲存資料有誤!(請洽系統管理者)!!"
                    Exit Sub
                End If

            Case cst_ptUpdate
                sql = $" SELECT * FROM PLAN_VERREPORT WHERE PlanID={TIMS.CINT1(rqPlanID)} AND ComIDNO='{rqComIDNO}' AND SeqNo={TIMS.CINT1(rqSeqNO)}"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count <> 1 Then
                    sm.LastErrorMessage = cst_errmsg26 '"儲存資料有誤!(請洽系統管理者)!!"
                    Exit Sub
                End If
                dr = dt.Rows(0)
                iPVID = dt.Rows(0)("PVID")

            Case Else
                sm.LastResultMessage = TIMS.cst_NODATAMsg1
                Exit Sub

        End Select
        '本課程屬環境部淨零綠領人才培育課程
        dr("EnvZeroTrain") = If(CB_EnvZeroTrain.Checked, "Y", "N")
        txtOthFacDesc23.Text = TIMS.ClearSQM2(txtOthFacDesc23.Text)
        '其他設施說明
        dr("OthFacDesc23") = If(txtOthFacDesc23.Text <> "", txtOthFacDesc23.Text, Convert.DBNull)
        '報請主管機關核備
        Dim v_rbl_REPORTE As String = If(rbl_REPORTE_Y.Checked, "Y", If(rbl_REPORTE_N.Checked, "N", If(Hid_REPORTE.Value <> "", Hid_REPORTE.Value, "")))
        dr("REPORTE") = If(v_rbl_REPORTE <> "", v_rbl_REPORTE, Convert.DBNull)
        Dim v_rblFuncLevel As String = TIMS.GetListValue(rblFuncLevel)
        dr("FuncLevel") = If(v_rblFuncLevel <> "", v_rblFuncLevel, Convert.DBNull)
        dr("TMethod") = vTMethod
        dr("TMethodOth") = TIMS.Get_Substr1(TIMS.ClearSQM(TMethodOth.Text), 200)
        'dr("ClassID") = ClassID.SelectedValue
        TIMS.Chk_placeholder(tPOWERNEED1)
        TIMS.Chk_placeholder(tPOWERNEED2)
        TIMS.Chk_placeholder(tPOWERNEED4)

        dr("POWERNEED1") = If(tPOWERNEED1.Text <> "", TIMS.Get_Substr1(tPOWERNEED1.Text, 2000), Convert.DBNull)
        dr("POWERNEED2") = If(tPOWERNEED2.Text <> "", TIMS.Get_Substr1(tPOWERNEED2.Text, 2000), Convert.DBNull)
        dr("POWERNEED3") = If(tPOWERNEED3.Text <> "", TIMS.Get_Substr1(tPOWERNEED3.Text, 2000), Convert.DBNull)
        '與政策性產業課程之關聯性概述
        TIMS.Chk_placeholder(tPOLICYREL)
        dr("POLICYREL") = If(tPOLICYREL.Text <> "", TIMS.Get_Substr1(tPOLICYREL.Text, 2000), Convert.DBNull)

        Dim objD4CHK As Object = If(cbPOWERNEED4.Checked, TIMS.cst_YES, Convert.DBNull)
        dr("POWERNEED4CHK") = objD4CHK
        If Not cbPOWERNEED4.Checked Then tPOWERNEED4.Text = ""
        dr("POWERNEED4") = If(tPOWERNEED4.Text <> "", TIMS.Get_Substr1(tPOWERNEED4.Text, 2000), Convert.DBNull)
        dr("PlanCause") = If(PlanCause.Text <> "", PlanCause.Text, Convert.DBNull)
        dr("PurScience") = If(PurScience.Text <> "", PurScience.Text, Convert.DBNull)
        dr("PurTech") = If(PurTech.Text <> "", PurTech.Text, Convert.DBNull)
        dr("PurMoral") = If(PurMoral.Text <> "", PurMoral.Text, Convert.DBNull)

        dr("Domain") = Convert.DBNull ' Domain.Text
        CapAll.Text = TIMS.ClearSQM(CapAll.Text)
        dr("CapAll") = If(CapAll.Text <> "", CapAll.Text, Convert.DBNull) ' CapAll.Text
        'dr("CostDesc") = If( CostDesc.Text <> "",  CostDesc.Text, Convert.DBNull) ' CostDesc.Text 'Note
        dr("CostDesc") = If(Note.Text <> "", Note.Text, Convert.DBNull) ' CostDesc.Text
        dr("RecDesc") = If(RecDesc.Text <> "", RecDesc.Text, Convert.DBNull) ' RecDesc.Text
        dr("LearnDesc") = If(LearnDesc.Text <> "", LearnDesc.Text, Convert.DBNull) ' LearnDesc.Text
        dr("ActDesc") = If(ActDesc.Text <> "", ActDesc.Text, Convert.DBNull) '  ActDesc.Text
        dr("ResultDesc") = If(ResultDesc.Text <> "", ResultDesc.Text, Convert.DBNull) ' ResultDesc.Text
        dr("OtherDesc") = If(OtherDesc.Text <> "", OtherDesc.Text, Convert.DBNull) ' OtherDesc.Text

        '是否為iCAP課程 / 是, 請填寫 課程相關說明 /否/
        Dim sISiCAPCOUR As String = If(RB_ISiCAPCOUR_Y.Checked, "Y", If(RB_ISiCAPCOUR_N.Checked, "N", ""))
        dr("ISiCAPCOUR") = If(sISiCAPCOUR <> "", sISiCAPCOUR, Convert.DBNull)
        iCAPCOURDESC.Text = TIMS.Get_Substr1(TIMS.ClearSQM(iCAPCOURDESC.Text), 500) '課程相關說明
        dr("iCAPCOURDESC") = If(iCAPCOURDESC.Text <> "", iCAPCOURDESC.Text, Convert.DBNull) '(500)
        dr("Recruit") = If(Recruit.Text <> "", Recruit.Text, Convert.DBNull) '  Recruit.Text
        dr("Selmethod") = If(Selmethod.Text <> "", Selmethod.Text, Convert.DBNull)
        dr("Inspire") = If(Inspire.Text <> "", Inspire.Text, Convert.DBNull)

        'TGovExamCG:本課程結訓後須參加環境部辦理之淨零綠領人才培育課程測驗；測驗成績達及格，即可申請本方案補助。
        Dim sTGovExamC As String = If(TGovExamCY.Checked, "Y", If(TGovExamCN.Checked, "N", If(TGovExamCG.Checked, "G", If(Hid_TGovExam.Value <> "", Hid_TGovExam.Value, ""))))
        dr("TGovExam") = If(sTGovExamC <> "", sTGovExamC, Convert.DBNull)
        'If TGovExamName.Text <> "" Then TGovExamName.Text = Trim(TGovExamName.Text)
        GOVAGENAME.Text = TIMS.Get_Substr1(TIMS.ClearSQM(GOVAGENAME.Text), 50) '政府機關名稱
        dr("GOVAGENAME") = If(GOVAGENAME.Text <> "", GOVAGENAME.Text, Convert.DBNull)
        TGovExamName.Text = TIMS.Get_Substr1(TIMS.ClearSQM(TGovExamName.Text), 50) '證照或檢定名稱
        dr("TGovExamName") = If(TGovExamName.Text <> "", TGovExamName.Text, Convert.DBNull)

        dr("memo8") = If(chkMEMO8C1.Checked, cst_msg_memo8a, Convert.DBNull)
        txtMemo8.Text = TIMS.ClearSQM(txtMemo8.Text)
        Dim v_memo82 As String = ""
        If chkMEMO8C2.Checked Then
            If txtMemo8.Text = "" Then txtMemo8.Text = " " '若沒有值，輸入一個空白
            v_memo82 = TIMS.Get_Substr1(txtMemo8.Text, 500)
        End If
        dr("memo82") = If(v_memo82 <> "", v_memo82, Convert.DBNull)

        'PLAN_VERREPORT '未-正式送出-都存NULL
        'select ISAPPRPAPER,count(1) from PLAN_VERREPORT group by ISAPPRPAPER
        'dr("IsApprPaper") = If(Convert.ToString(dr("IsApprPaper")) <> "", Convert.ToString(dr("IsApprPaper")), "N")
        '計畫-正式儲存-正式送出
        Dim V_IsApprPaper As String = Convert.ToString(dr("IsApprPaper"))
        Select Case Hid_sender1.Value
            Case cst_SaveRcc
                V_IsApprPaper = "Y"
            Case Else
                If (V_IsApprPaper <> "Y") Then V_IsApprPaper = "N"
        End Select
        dr("IsApprPaper") = If(V_IsApprPaper <> "", V_IsApprPaper, Convert.DBNull)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
    End Sub

#Region "NOUSE"
    '儲存-政策性產業課程可辦理班數-PLAN_PRECLASS
    'Sub SAVE_PLAN_PRECLASS()
    '    If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
    '    tmpPCS = upt_PlanX.Value  '有儲存資料過了
    '    PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
    '    ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
    '    SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
    '    'Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
    '    'Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
    '    'Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")

    '    For Each eItem As DataGridItem In DataGrid33.Items
    '        'Dim lab_YEARS As Label = eItem.FindControl("lab_YEARS")
    '        Dim Hid_PREID As HiddenField = eItem.FindControl("Hid_PREID")
    '        Dim Hid_YEARS_V1 As HiddenField = eItem.FindControl("Hid_YEARS_V1")
    '        Dim txt_PRECLASSCNT As TextBox = eItem.FindControl("txt_PRECLASSCNT")
    '        Dim iPREID As Integer = Val(Hid_PREID.Value)
    '        Hid_YEARS_V1.Value = TIMS.ClearSQM(Hid_YEARS_V1.Value)
    '        'Dim sql As String = ""
    '        'sql = ""
    '        'sql &= " SELECT PREID,PCNT1"
    '        'sql &= " FROM PLAN_PRECLASS"
    '        'sql &= " WHERE 1=1"
    '        'sql &= " AND PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
    '        'sql &= " AND YEARS=@YEARS"
    '        Dim tmpV As String = GET_PRECLASS_PCNT1(PlanID_value, ComIDNO_value, SeqNO_value, Hid_YEARS_V1.Value, iPREID)
    '        If iPREID = 0 Then
    '            'insert
    '            iPREID = DbAccess.GetNewId(objconn, "PLAN_PRECLASS_PREID_SEQ,PLAN_PRECLASS,PREID")
    '            Dim iSql As String = ""
    '            iSql = "" & vbCrLf
    '            iSql &= " INSERT INTO PLAN_PRECLASS (PREID ,PLANID,COMIDNO,SEQNO,YEARS,PCNT1, CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf
    '            iSql &= " VALUES (@PREID ,@PLANID,@COMIDNO,@SEQNO,@YEARS,@PCNT1,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf
    '            Dim pParms As New Hashtable
    '            pParms.Clear()
    '            pParms.Add("PREID", iPREID)
    '            pParms.Add("PLANID", PlanID_value)
    '            pParms.Add("COMIDNO", ComIDNO_value)
    '            pParms.Add("SEQNO", SeqNO_value)
    '            pParms.Add("YEARS", Hid_YEARS_V1.Value)
    '            pParms.Add("PCNT1", Val(txt_PRECLASSCNT.Text))
    '            pParms.Add("CREATEACCT", sm.UserInfo.UserID)
    '            pParms.Add("MODIFYACCT", sm.UserInfo.UserID)
    '            DbAccess.ExecuteNonQuery(iSql, objconn, pParms)
    '        Else
    '            'update
    '            Dim uSql As String = ""
    '            uSql = ""
    '            uSql &= " UPDATE PLAN_PRECLASS"
    '            uSql &= " SET PCNT1=@PCNT1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()"
    '            uSql &= " WHERE 1=1"
    '            uSql &= " AND PREID=@PREID"
    '            Dim pParms As New Hashtable
    '            pParms.Clear()
    '            pParms.Add("PCNT1", Val(txt_PRECLASSCNT.Text))
    '            pParms.Add("MODIFYACCT", sm.UserInfo.UserID)
    '            pParms.Add("PREID", iPREID)
    '            DbAccess.ExecuteNonQuery(uSql, objconn, pParms)
    '        End If
    '    Next
    'End Sub
#End Region
    Private Sub DataGrid21_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid21.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim Specialty1 As Label = e.Item.FindControl("Specialty1")
                'Dim ProLicense As Label = e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEA As HtmlInputButton = e.Item.FindControl("btn_TCTYPEA")
                Dim rqRID As String = sm.UserInfo.RID
                If RIDValue.Value.Length > 1 Then rqRID = RIDValue.Value
                sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                btn_TCTYPEA.Attributes("onclick") = sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                Specialty1.Text = Convert.ToString(drv("Specialty1"))
                'ProLicense.Text = Convert.ToString(drv("ProLicense"))
                TIMS.sUtl_SetMaxLen(cst_iMaxLen_TeacherDesc, TeacherDesc)
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'TeacherDesc.Attributes.Add("disabled", "disabled")
                TeacherDesc.Style.Item("background-color") = "#BDBDBD"
                TeacherDesc.ReadOnly = False
                btn_TCTYPEA.Visible = True

                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly = True
                '        btn_TCTYPEA.Visible = False
                'End Select

                Dim flag_can_save As Boolean = True
                If RIDValue.Value = "" Then flag_can_save = False '不同單位 不提供儲存
                If sm.UserInfo.RID <> RIDValue.Value Then flag_can_save = False '不同單位 不提供儲存
                Select Case sm.UserInfo.LID
                    Case 2
                        '不同單位 不提供儲存
                        If Not flag_can_save Then
                            TeacherDesc.ReadOnly = True
                            btn_TCTYPEA.Visible = False
                        End If
                End Select
                If Not gflag_can_save Then
                    TeacherDesc.ReadOnly = True
                    btn_TCTYPEA.Visible = False
                End If
        End Select
    End Sub

    Private Sub DataGrid22_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid22.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim Specialty1 As Label = e.Item.FindControl("Specialty1")
                'Dim ProLicense As Label = e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEB As HtmlInputButton = e.Item.FindControl("btn_TCTYPEB")
                Dim rqRID As String = sm.UserInfo.RID '使用者單位
                If RIDValue.Value.Length > 1 Then rqRID = RIDValue.Value
                sWOScript1 = "wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                btn_TCTYPEB.Attributes("onclick") = sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                Specialty1.Text = Convert.ToString(drv("Specialty1"))
                'ProLicense.Text = Convert.ToString(drv("ProLicense"))
                TIMS.sUtl_SetMaxLen(cst_iMaxLen_TeacherDesc, TeacherDesc)
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'TeacherDesc.Attributes.Add("disabled", "disabled")
                TeacherDesc.Style.Item("background-color") = "#BDBDBD"
                TeacherDesc.ReadOnly = False
                btn_TCTYPEB.Visible = True

                Dim flag_can_save As Boolean = True
                If RIDValue.Value = "" Then flag_can_save = False '不同單位 不提供儲存
                If sm.UserInfo.RID <> RIDValue.Value Then flag_can_save = False '不同單位 不提供儲存
                Select Case sm.UserInfo.LID
                    Case 2
                        '不同單位 不提供儲存
                        If Not flag_can_save Then
                            TeacherDesc.ReadOnly = True
                            btn_TCTYPEB.Visible = False
                        End If
                End Select
                If Not gflag_can_save Then
                    TeacherDesc.ReadOnly = True
                    btn_TCTYPEB.Visible = False
                End If

                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly = True
                '        btn_TCTYPEB.Visible = False
                'End Select
        End Select

    End Sub

#Region "NO USE"
    '基本儲存若是ok 為 true 
    'Function Chk_SAVEBASICOK() As Boolean
    '    Dim flag_saveok As Boolean = True '儲存若是ok 為 true 
    '    'Const cst_session_saveok As String = "saveok"
    '    If Session(cst_session_saveok) Is Nothing Then flag_saveok = False
    '    If Session(cst_session_saveok) IsNot Nothing Then
    '        If Not Session(cst_session_saveok) = True Then flag_saveok = False
    '    End If
    '    If Not flag_saveok Then Session(cst_session_saveok) = Nothing
    '    Return flag_saveok
    'End Function
#End Region

    ''' <summary> 正式儲存 / 基本儲存之後 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSAVE2_Click(sender As Object, e As EventArgs) Handles BtnSAVE2.Click
        Hid_sender1.Value = cst_SaveRcc 'sender.text '計畫-正式儲存-正式送出

        'BtnSAVE2.Visible = False '不顯示儲存鈕

        'Dim flag_saveBasic_ok As Boolean = Chk_SAVEBASICOK()
        PlanID_value = TIMS.ClearSQM(Request("PlanID"))
        ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
        SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
        Dim drPP As DataRow = TIMS.GetPCSDate(PlanID_value, ComIDNO_value, SeqNO_value, objconn)

        '顯示儲存鈕 / 'FALSE:不顯示儲存鈕
        Dim flag_can_show_data24 As Boolean = If(drPP IsNot Nothing AndAlso Convert.ToString(drPP("ISAPPRPAPER")) = "Y", True, False)
        If Not flag_can_show_data24 Then
            BtnSAVE2.Visible = False '不顯示儲存鈕
            sm.LastErrorMessage = cst_errmsg23
            Exit Sub
        End If

        '(委訓單位增強檢核)
        If sm.UserInfo.LID = 2 Then
            Dim fg_APPLIEDRESULT_Y As Boolean = If(drPP IsNot Nothing AndAlso Convert.ToString(drPP("APPLIEDRESULT")) = "Y", True, False)
            If fg_APPLIEDRESULT_Y Then
                BtnSAVE2.Visible = False '不顯示儲存鈕
                sm.LastErrorMessage = "班級已審核通過，不可再次儲存！"
                Exit Sub
            End If
        End If

        'Dim flag_saveBasic_ok As Boolean = Chk_SAVEBASICOK()
        'If Not flag_saveBasic_ok Then
        '    'Session(cst_session_saveok) = Nothing
        '    Common.MessageBox(Me, cst_errmsg23)
        '    Exit Sub
        'End If

        Call UTL_PREARRANGEMENT()
        Dim RegKey1 As String = "_onload2"
        'Dim RegScript1 As String = "<script language=""javascript"">document.getElementById('btnAdd').style.display="""";Layer_change(9);</script>"
        If (LayerState.Value = "") Then LayerState.Value = "9"
        Dim RegScript1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")

        Dim sErrMsg As String = ""
        '基本儲存-資料檢核確認
        Call CheckAddData(sErrMsg)
        If sErrMsg <> "" Then
            '有錯誤訊息
            Page.RegisterStartupScript(RegKey1, RegScript1)
            sm.LastErrorMessage = sErrMsg
            Exit Sub 'Return False '不可儲存
        End If

        '檢核是否有業務權限
        Dim flag_CAN_NOCHECK_RIDPLAN As Boolean = False 'true:(可以不檢查業務權限)
        Dim flag_RIDPLAN As Boolean = Chk_RIDPLAN(RIDValue.Value, sm.UserInfo.PlanID)
        If flag_IsSuperUser_1 Then flag_CAN_NOCHECK_RIDPLAN = True
        'If flag_TIMS_Test_1 Then flag_CAN_NOCHECK_RIDPLAN = True '檢核是否有業務權限
        If Not flag_CAN_NOCHECK_RIDPLAN AndAlso Not flag_RIDPLAN Then
            sErrMsg &= $"{cst_errmsg22},{RIDValue.Value},{sm.UserInfo.PlanID},{sm.UserInfo.Years}{vbCrLf}" '"登入者無正確的業務權限，不提供儲存服務!!" & vbCrLf
        End If

        If sErrMsg <> "" Then
            '有錯誤訊息
            Page.RegisterStartupScript(RegKey1, RegScript1)
            sm.LastErrorMessage = sErrMsg
            Exit Sub 'Return False '不可儲存
        End If
        'Dim sErrMsg As String = ""
        '正式儲存檢核
        Dim rst As Boolean = CheckData14(sErrMsg)
        If sErrMsg <> "" Then
            Page.RegisterStartupScript(RegKey1, RegScript1)
            sm.LastErrorMessage = sErrMsg
            Exit Sub
        End If

        '儲存點
        '假設處理某段程序需花費n毫秒 (避免機器不同步)
        If Session("GUID1") <> ViewState("GUID1") Then Threading.Thread.Sleep(1)
        ViewState("GUID1") = TIMS.GetGUID() : Session("GUID1") = ViewState("GUID1")

        '儲存 開班計畫/開班計畫表資料維護 (正式儲存)
        'Try
        '    Call INSERT_PLAN_TABLE(cst_SaveRcc) '若儲存成功，則下列不執行，直接跳頁 ../01/TC_01_014_add.aspx
        'Catch ex As Exception
        '    Dim strErrmsg As String = ex.Message
        '    strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
        '    strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
        '    'strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
        '    Call TIMS.WriteTraceLog(strErrmsg, ex)
        '    sm.LastErrorMessage = String.Concat(cst_errmsg6, vbCrLf, ex.Message)
        '    Return ' Exit Sub
        'End Try
        '儲存 開班計畫/開班計畫表資料維護 (正式儲存)
        Call INSERT_PLAN_TABLE(cst_SaveRcc) '若儲存成功，則下列不執行，直接跳頁 ../01/TC_01_014_add.aspx
        'ViewState("dtTaddress") = Nothing
        'Page.RegisterStartupScript(RegKey1, RegScript1)
    End Sub

    ''' <summary> 檢核-2019年啟用 work2019x01:2019 政府政策性產業</summary>
    ''' <returns></returns>
    Function CHK_KID20_VAL_OTH() As String
        Dim Errmsg As String = ""
        Const CST_ERRM1 As String = "，不可複選(僅可單一勾選)"
        '「5+2」產業創新計畫 5+2產業,'【台灣AI行動計畫】 KID='08','【數位國家創新經濟發展方案】KID='09',
        '【國家資通安全發展方案】KID='10','【前瞻基礎建設計畫】,'【新南向政策】KID='19',
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"「5+2」產業創新計畫{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【台灣AI行動計畫】{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【數位國家創新經濟發展方案】{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【國家資通安全發展方案】{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【前瞻基礎建設計畫】{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【新南向政策】{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID22)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【進階政策性產業類別】{CST_ERRM1}{vbCrLf}"
        Return Errmsg
    End Function

    ''' <summary>取值-2019年啟用 work2019x01:2019 政府政策性產業</summary>
    ''' <returns></returns>
    Function GET_KID20_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID20_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID20_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function

    ''' <summary> 檢核-2025年啟用-2025 政府政策性產業</summary>
    ''' <returns></returns>
    Function CHK_KID25_VAL_OTH() As String
        Dim Errmsg As String = ""
        Const CST_ERRM1 As String = "，不可複選(僅可單一勾選)"
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID25_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"亞洲矽谷{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID25_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"重點產業{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID25_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"台灣AI行動計畫{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID25_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"智慧國家方案{CST_ERRM1}{vbCrLf}"

        tmp01 = TIMS.GetCblValue(CBLKID25_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"國家人才競爭力躍升方案{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID25_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"新南向政策{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID25_7)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"AI加值應用{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID25_8)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"職場續航{CST_ERRM1}{vbCrLf}"

        tmp01 = TIMS.GetCblValue(CBLKID22B)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"【進階政策性產業類別】{CST_ERRM1}{vbCrLf}"
        Return Errmsg
    End Function

    ''' <summary>取值-啟用-2025 政府政策性產業</summary>
    ''' <returns></returns>
    Function GET_KID25_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID25_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_7)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID25_8)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function

    ''' <summary> 檢核-2026年啟用-2026 政府政策性產業</summary>
    ''' <returns></returns>
    Function CHK_KID26_VAL_OTH() As String
        Dim Errmsg As String = ""
        Const CST_ERRM1 As String = "，不可複選(僅可單一勾選)"
        Dim KID26STR1 As String() = {"五大信賴產業推動方案", "六大區域產業及生活圈", "智慧國家2.0綱領", "新南向政策推動計畫", "國家人才競爭力躍升方案", "AI新十大建設推動方案", "台灣AI行動計畫2.0", "智慧機器人產業推動方案", "臺灣2050淨零轉型"}
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID26_1)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(0)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_2)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(1)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_3)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(2)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_4)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(3)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_5)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(4)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_6)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(5)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_7)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(6)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_8)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(7)}{CST_ERRM1}{vbCrLf}"
        tmp01 = TIMS.GetCblValue(CBLKID26_9)
        If tmp01 <> "" AndAlso tmp01.IndexOf(",") > -1 Then Errmsg &= $"{KID26STR1(8)}{CST_ERRM1}{vbCrLf}"

        Dim fg_show_trKID26 As Boolean = (trKID26.Visible AndAlso TIMS.Display_Chk_Inline(trKID26))
        If Not fg_show_trKID26 Then Return Errmsg '(未顯示此選項下列資訊不檢核)

        Dim vCBLKID26_9 As String = $",{TIMS.GetCblValue(CBLKID26_9)},"
        If vCBLKID26_9.IndexOf(",20,") > -1 AndAlso Not CB_EnvZeroTrain.Checked Then
            Errmsg &= $"政策性產業項目 勾選「{KID26STR1(8)}-環境部-【淨零綠領人才培育課程】」必須勾選「本課程屬環境部淨零綠領人才培育課程」{vbCrLf}"
        ElseIf CB_EnvZeroTrain.Checked AndAlso vCBLKID26_9.IndexOf(",20,") = -1 Then
            Errmsg &= $"勾選「本課程屬環境部淨零綠領人才培育課程」必須勾選 政策性產業項目-「{KID26STR1(8)}-環境部-【淨零綠領人才培育課程】」{vbCrLf}"
        End If
        Return Errmsg
    End Function

    ''' <summary>取值-啟用-2026 政府政策性產業</summary>
    ''' <returns></returns>
    Function GET_KID26_VAL() As String
        Dim rst As String = ""
        Dim tmp01 As String = ""
        tmp01 = TIMS.GetCblValue(CBLKID26_1)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_2)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_3)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_4)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_5)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_6)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_7)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_8)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        tmp01 = TIMS.GetCblValue(CBLKID26_9)
        If tmp01 <> "" Then rst &= String.Concat(If(rst <> "", ",", ""), tmp01)
        Return rst
    End Function

#Region "NOUSE"
    'DataGrid33-政策性產業課程可辦理班數
    'Private Sub DataGrid33_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid33.ItemDataBound
    '    Select Case e.Item.ItemType
    '        Case ListItemType.Item, ListItemType.AlternatingItem
    '            Dim lab_YEARS As Label = e.Item.FindControl("lab_YEARS")
    '            Dim Hid_PREID As HiddenField = e.Item.FindControl("Hid_PREID")
    '            Dim Hid_YEARS_V1 As HiddenField = e.Item.FindControl("Hid_YEARS_V1")
    '            Dim txt_PRECLASSCNT As TextBox = e.Item.FindControl("txt_PRECLASSCNT") '年可辦理班數
    '            Dim drv As DataRowView = e.Item.DataItem
    '            Dim vYears1_ROC As String = Convert.ToString(Val(drv("YEARS")) - 1911)
    '            Hid_PREID.Value = Convert.ToString(drv("PREID"))
    '            Hid_YEARS_V1.Value = Convert.ToString(drv("YEARS"))
    '            lab_YEARS.Text = vYears1_ROC 'Convert.ToString(drv("YEARS"))
    '            txt_PRECLASSCNT.Text = Convert.ToString(drv("PCNT1"))
    '    End Select
    'End Sub

    '檢核記憶體 '異常 (true:異常NG false:OK)
    'Public Shared Function CHK_SESSION_ERROR(ByRef dtTemp As DataTable, ByVal PCS_V2 As String) As Boolean
    '    'PCS_V2= (PlanID_value & ComIDNO_value & SeqNO_value)
    '    Dim flag_session_error As Boolean = False
    '    Dim flag_check_pcs As Boolean = False
    '    If dtTemp.Rows.Count > 0 Then
    '        For Each dr1 As DataRow In dtTemp.Rows
    '            If Not dr1.RowState = DataRowState.Deleted Then
    '                '檢核記憶體'異常
    '                'Dim dr1 As DataRow = dtTemp.Rows(0)
    '                Dim V_PlanID As String = Convert.ToString(dr1("PlanID"))
    '                Dim V_ComIDNO As String = Convert.ToString(dr1("ComIDNO"))
    '                Dim V_SeqNO As String = Convert.ToString(dr1("SeqNO"))
    '                If V_PlanID <> "" AndAlso V_ComIDNO <> "" AndAlso V_SeqNO <> "" Then flag_check_pcs = True
    '                If flag_check_pcs Then
    '                    If (V_PlanID & V_ComIDNO & V_SeqNO) <> PCS_V2 Then
    '                        flag_session_error = True
    '                    End If
    '                End If
    '                Exit For
    '            End If
    '        Next
    '    End If
    '    Return flag_session_error
    'End Function
#End Region

    '產生新的GUID 避免記憶體相同 而異常
    Sub CREATE_NEW_GUID21()
        Dim s_old_session_guid1 As String = "" '上個session guid清理

        hid_TrainDescTable_guid1.Value = TIMS.GetGUID()
        Session(hid_TrainDescTable_guid1.Value) = Nothing

        hid_PLAN_BUSPACKAGE_guid1.Value = TIMS.GetGUID()
        Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing

        hid_PersonCostTable_guid1.Value = TIMS.GetGUID()
        Session(hid_PersonCostTable_guid1.Value) = Nothing

        hid_CommonCostTable_guid1.Value = TIMS.GetGUID()
        Session(hid_CommonCostTable_guid1.Value) = Nothing

        hid_SheetCostTable_guid1.Value = TIMS.GetGUID()
        Session(hid_SheetCostTable_guid1.Value) = Nothing

        hid_OtherCostTable_guid1.Value = TIMS.GetGUID()
        Session(hid_OtherCostTable_guid1.Value) = Nothing

        hid_planONCLASS_guid1.Value = TIMS.GetGUID()
        Session(hid_planONCLASS_guid1.Value) = Nothing
    End Sub


#Region "SqlTransaction"

    ''' <summary> 取得序號(LOCK) </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="Trans"></param>
    ''' <param name="conn"></param>
    ''' <param name="PlanID_value"></param>
    ''' <param name="ComIDNO_value"></param>
    ''' <returns></returns>
    Public Shared Function GetMaxSeqNum(ByRef MyPage As Page, ByRef Trans As SqlTransaction, ByRef conn As SqlConnection,
                                        ByVal PlanID_value As String, ByVal ComIDNO_value As String) As Integer
        Dim Rst As Integer = 1 '由1開始
        Dim exErrmsg As String = ""
        Try
            '取得SeqNO
            Dim sql As String = $" SELECT Max(SeqNO) MaxSeqNO From PLAN_PLANINFO WHERE PlanID={PlanID_value} AND ComIDNO='{ComIDNO_value}'"
            Dim dr As DataRow = DbAccess.GetOneRow(sql, Trans)
            If dr IsNot Nothing Then '有取得
                If Not IsDBNull(dr("MaxSeqNO")) Then Rst = TIMS.CINT1(dr("MaxSeqNO")) + 1 '有值,最大值再加1
            End If
        Catch ex As Exception
            If Trans IsNot Nothing Then DbAccess.RollbackTrans(Trans)
            Call TIMS.CloseDbConn(conn)
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-Rst:{0}-{1}-{2}", PlanID_value, ComIDNO_value, Rst) & vbCrLf
            'strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(MyPage, ex) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)

            exErrmsg &= ex.ToString & vbCrLf
            Throw New Exception(exErrmsg)
        End Try

        Return Rst
    End Function

    ''' <summary> 取得正確的PLANID </summary>
    ''' <param name="sm_PlanID"></param>
    ''' <param name="sm_LID"></param>
    ''' <param name="RID_Val"></param>
    ''' <returns></returns>
    Function GET_CORRECT_PlanID(ByRef Trans As SqlTransaction, ByRef conn As SqlConnection,
                                ByRef sm_PlanID As String, ByRef sm_LID As Short, ByRef RID_Val As String) As String
        'sm.UserInfo.PlanID, sm.UserInfo.LID, RIDValue.Value
        Dim rst As String = sm_PlanID
        If RID_Val = "" Then Return rst
        If sm_LID <> 0 Then Return rst
        'sm_LID =0 依RID取得 PlanID\
        Dim rPlanID As String = TIMS.GET_RIDPLANID(Trans, conn, RID_Val)
        rst = If(rPlanID <> "", rPlanID, sm_PlanID)
        Return rst
    End Function

    ''' <summary> 取得有效的訓練計畫資料DR </summary>
    ''' <param name="Trans"></param>
    ''' <param name="conn"></param>
    ''' <param name="da"></param>
    ''' <param name="dt"></param>
    ''' <param name="dr"></param>
    ''' <param name="s_TransType"></param>
    ''' <returns></returns>
    Function GET_PPINFO_DR(ByRef Trans As SqlTransaction, ByRef conn As SqlConnection,
                           ByRef da As SqlDataAdapter, ByRef dt As DataTable, ByRef dr As DataRow,
                           ByRef s_TransType As String, ByRef s_WHERE As String) As Boolean
        'Dim rst As Boolean = False
        Const cst_fWHERE As String = "PLANID={0} AND COMIDNO='{1}' AND SEQNO={2}"
        Dim sql As String = ""
        Try
            If upt_PlanX.Value <> "" Then '有儲存資料過了
                '有儲存資料過了,準備儲存資料
                tmpPCS = upt_PlanX.Value  '有儲存資料過了
                PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
                ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
                SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")

                s_WHERE = String.Format(cst_fWHERE, PlanID_value, ComIDNO_value, SeqNO_value)
                sql = String.Format("SELECT * FROM PLAN_PLANINFO WHERE {0}", s_WHERE)
                dt = DbAccess.GetDataTable(sql, da, Trans)
                If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return False '(資料取得有誤)

                dr = dt.Rows(0)
            Else
                If (Convert.ToString(Request("PlanID")) = "" OrElse gflag_ccopy) Then
                    '新增資料 、ccopy=1 、草稿新增 而來
                    PlanID_value = GET_CORRECT_PlanID(Trans, conn, sm.UserInfo.PlanID, sm.UserInfo.LID, RIDValue.Value)
                    ComIDNO_value = ComidValue.Value
                    SeqNO_value = GetMaxSeqNum(Me, Trans, conn, PlanID_value, ComIDNO_value) '+1 ComidValue.Value  sm.UserInfo.PlanID
                    '準備儲存資料
                    sql = " SELECT * FROM PLAN_PLANINFO WHERE 1<>1 "
                    s_TransType = TIMS.cst_TRANS_LOG_Insert
                    s_WHERE = String.Format(cst_fWHERE, PlanID_value, ComidValue.Value, SeqNO_value)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("PlanID") = PlanID_value 'sm.UserInfo.PlanID
                    dr("ComIDNO") = ComidValue.Value
                    dr("SeqNO") = SeqNO_value

                    dr("RID") = RIDValue.Value '空的才存取
                    dr("PlanYear") = Label3.Text '空的才存取
                    dr("TPlanID") = TPlanID '空的才存取
                    '預防新增時選擇草稿儲存
                    '導致因為停留原畫面，再儲存時會第二次重複儲存。
                    Org.Disabled = True
                Else
                    '修改
                    PlanID_value = TIMS.ClearSQM(Request("PlanID"))
                    ComIDNO_value = TIMS.ClearSQM(Request("ComIDNO"))
                    SeqNO_value = TIMS.ClearSQM(Request("SeqNO"))
                    'sql = " SELECT * FROM PLAN_PLANINFO WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                    s_WHERE = String.Format(cst_fWHERE, PlanID_value, ComIDNO_value, SeqNO_value)
                    sql = String.Format("SELECT * FROM PLAN_PLANINFO WHERE {0}", s_WHERE)
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return False '(資料取得有誤)
                    dr = dt.Rows(0)
                End If
                tmpPCS = ""
                TIMS.SetMyValue(tmpPCS, "PlanID", PlanID_value)
                TIMS.SetMyValue(tmpPCS, "ComIDNO", ComIDNO_value)
                TIMS.SetMyValue(tmpPCS, "SeqNO", SeqNO_value)
                upt_PlanX.Value = tmpPCS
            End If
        Catch ex As Exception
            upt_PlanX.Value = ""
            DbAccess.RollbackTrans(Trans)
            Call TIMS.CloseDbConn(conn)

            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            sm.LastErrorMessage = cst_errmsg6
            Return False ' Exit Sub
        End Try
        Return True
    End Function

    ''' <summary>取得-訓練業別</summary>
    ''' <returns></returns>
    Function GET_o_TMID_VAL() As Object
        Dim o_TMID As Object = Convert.DBNull
        Select Case strYears
            Case cst_strYears_2014 '"2014"
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    o_TMID = If(jobValue.Value <> "", jobValue.Value, Convert.DBNull)
                Else
                    o_TMID = If(trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                End If
            Case cst_strYears_2015 '"2015"
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    o_TMID = If(jobValue.Value <> "", jobValue.Value, Convert.DBNull)
                Else
                    o_TMID = If(trainValue.Value <> "", trainValue.Value, Convert.DBNull)
                End If
            Case cst_strYears_2018 '"2018"
                o_TMID = If(trainValue.Value <> "", trainValue.Value, Convert.DBNull)
        End Select
        Return o_TMID
    End Function

    ''' <summary>(儲存) iNum:'1是正式 '2是草稿 PLAN_PLANINFO </summary>
    ''' <param name="SaveType1"></param>
    Private Sub INSERT_PLAN_TABLE(ByVal SaveType1 As String)
        'iNum:'1是正式 '2是草稿 'ByVal iNum As Integer 'SaveType1
        Call ChangNoteText(tmpNoteDt)

        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        '查無或新增為0 其餘數字大於0
        Dim strReportCount As String = GET_PLAN_VERREPORT_CNT()

        Dim re_update_flag As Boolean = False '重新UPDATE
        Dim AppliedResult1 As String = "" '審核狀態值
        'Dim str_CommandArgument As String = "" 'TC_01_014.aspx用

        '容納人數必須為數字
        ConNum.Text = TIMS.ChangeIDNO(ConNum.Text)

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DistID As String = If(sm.UserInfo.LID = 0, TIMS.Get_DistID_RID(RIDValue.Value, objconn), sm.UserInfo.DistID)

        '取得-訓練業別
        Dim o_TMID As Object = GET_o_TMID_VAL() '取得-訓練業別

        '計畫別：產業人才投資計畫呈現「A」、提升自主勞工計畫呈現「B」
        Dim vPLAN1 As String = TIMS.Get_PSNO28_PLAN1(ComidValue.Value, objconn)
        '取得產投流水號(前6碼)年度別(3)+上下年(1)+計畫別(1)+分署別(1) 課程申請流水號
        Dim vPSNO28_6 As String = TIMS.Get_PSNO28_6(sm.UserInfo.Years, STDate.Text, vPLAN1, s_DistID) 'sm.UserInfo.DistID)
        Select Case SaveType1
            Case cst_SaveBasic, cst_SaveRcc '計畫-正式儲存
                If Len(vPSNO28_6) <> 6 Then
                    '(正式才檢核)'取得產投流水號(前6碼) '長度應該為6 
                    sm.LastErrorMessage = cst_errmsg2
                    Return 'Exit Sub 'If flag_TIMS_Test_1 Then Exit Sub
                End If
        End Select

        If ClassCount.Text <> "" Then ClassCount.Text = TIMS.ChangeIDNO(ClassCount.Text)
        FirstSort.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(FirstSort.Text))
        If FirstSort.Text <> "" Then
            '異常數字修正
            FirstSort.Text = If(TIMS.IsNumeric1(FirstSort.Text), TIMS.CINT1(FirstSort.Text).ToString(), "1")
        End If

        Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        'Dim v_Taddress2 As String = TIMS.GetListValue(Taddress2)
        'Dim v_Taddress3 As String = TIMS.GetListValue(Taddress3)

        Dim v_SciPlaceID As String = TIMS.GetListValue(SciPlaceID)
        Dim v_SciPlaceID2 As String = TIMS.GetListValue(SciPlaceID2)
        Dim v_TechPlaceID As String = TIMS.GetListValue(TechPlaceID)
        Dim v_TechPlaceID2 As String = TIMS.GetListValue(TechPlaceID2)
        '辦理方式: null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
        Dim vrbl_DISTANCE As String = TIMS.GetListValue(rbl_DISTANCE)
        '遠距課程環境1/2
        Dim v_REMOTEID1 As String = TIMS.GetListValue(ddl_REMOTEID1)
        'If v_REMOTEID1 <> "" Then Hid_RMTID1.Value = v_REMOTEID1
        Dim v_REMOTEID2 As String = TIMS.GetListValue(ddl_REMOTEID2)
        'If v_REMOTEID2 <> "" Then Hid_RMTID2.Value = v_REMOTEID2

        Dim drSciPc1 As DataRow = TIMS.Get_SciTechDR(ComidValue.Value, v_SciPlaceID, 1, objconn) '取得場地的地址
        Dim drSciPc2 As DataRow = TIMS.Get_SciTechDR(ComidValue.Value, v_SciPlaceID2, 1, objconn)
        Dim drTechPc1 As DataRow = TIMS.Get_SciTechDR(ComidValue.Value, v_TechPlaceID, 2, objconn) '取得場地的地址
        Dim drTechPc2 As DataRow = TIMS.Get_SciTechDR(ComidValue.Value, v_TechPlaceID2, 2, objconn)

        If (drTechPc2 IsNot Nothing) Then hid_AddressTechPTID2.Value = drTechPc2("PTID").ToString()
        If (drTechPc1 IsNot Nothing) Then hid_AddressTechPTID.Value = drTechPc1("PTID").ToString()
        If (drSciPc2 IsNot Nothing) Then hid_AddressSciPTID2.Value = drSciPc2("PTID").ToString()
        If (drSciPc1 IsNot Nothing) Then hid_AddressSciPTID.Value = drSciPc1("PTID").ToString()

        Dim tmpPTID As String = ""
        If (drTechPc2 IsNot Nothing) Then tmpPTID = drTechPc2("PTID").ToString()
        If (drTechPc1 IsNot Nothing) Then tmpPTID = drTechPc1("PTID").ToString()
        If (drSciPc2 IsNot Nothing) Then tmpPTID = drSciPc2("PTID").ToString()
        If (drSciPc1 IsNot Nothing) Then tmpPTID = drSciPc1("PTID").ToString()
        'tmpPTID = If(v_Taddress2 <> "", v_Taddress2, If(v_Taddress3 <> "", v_Taddress3, ""))
        Dim vsTaddressZip As String = ""
        Dim vsTAddress As String = ""
        Dim vsTaddressZIP6W As String = ""
        Call TIMS.GetTaddressPTID(objconn, tmpPTID, vsTaddressZip, vsTAddress, vsTaddressZIP6W)

        '檢核 CLASSINFO NOT_Ready true:(沒有classinfo) // false:(有classinfo)
        Dim fg_NOT_Ready_CLASSINFO As Boolean = If(Not CHK_CLASSINFO(Page, gflag_ccopy, objconn), True, False)

        Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update
        Dim s_TargetTable As String = "PLAN_PLANINFO"
        Dim s_FuncPath As String = "/TC/03/TC_03_006"
        'Const cst_fWHERE As String = "PLANID={0} And COMIDNO='{1}' AND SEQNO={2}"
        Dim s_WHERE As String = ""

        '2006/03/28 add conn by matt
        Dim gDataRowString As String = ""
        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn) 'PLAN_PLANINFO
            Try
                Dim dr As DataRow = Nothing
                Dim flag_CAN_SAVE_OK As Boolean = GET_PPINFO_DR(Trans, TransConn, da, dt, dr, s_TransType, s_WHERE)
                If Not flag_CAN_SAVE_OK Then
                    sm.LastErrorMessage = cst_errmsg34
                    DbAccess.RollbackTrans(Trans)
                    Call TIMS.CloseDbConn(TransConn)
                    Return
                End If

                dr("FIXSUMCOST") = If(FIXSUMCOST.Text <> "", Val(FIXSUMCOST.Text), Convert.DBNull)
                dr("ACTHUMCOST") = If(ACTHUMCOST.Text <> "", Val(ACTHUMCOST.Text), Convert.DBNull)
                dr("FIXExceeDesc") = If(FIXExceeDesc.Text <> "", FIXExceeDesc.Text, Convert.DBNull)
                dr("METSUMCOST") = If(METSUMCOST.Text <> "", Val(METSUMCOST.Text), Convert.DBNull)
                dr("METCOSTPER") = If(METCOSTPER.Text <> "", Val(METCOSTPER.Text), Convert.DBNull)
                dr("METExceeDesc") = If(METExceeDesc.Text <> "", METExceeDesc.Text, Convert.DBNull)

                dr("TMID") = If(Convert.ToString(o_TMID) <> "", o_TMID, Convert.DBNull)
                dr("CJOB_UNKEY") = If(cjobValue.Value <> "", cjobValue.Value, Convert.DBNull)
                '訓練業別正確性: 'Common.SetListItem(rblTMIDCORRECT, Convert.ToString(dr("TMIDCORRECT")))
                Dim v_rblTMIDCORRECT As String = TIMS.GetListValue(rblTMIDCORRECT)
                dr("TMIDCORRECT") = If(v_rblTMIDCORRECT <> "", v_rblTMIDCORRECT, Convert.DBNull)
                TIMS.Chk_placeholder(PlanCause)
                TIMS.Chk_placeholder(PurScience)
                TIMS.Chk_placeholder(PurTech)
                TIMS.Chk_placeholder(PurMoral)
                dr("PlanCause") = If(PlanCause.Text <> "", PlanCause.Text, Convert.DBNull)
                dr("PurScience") = If(PurScience.Text <> "", PurScience.Text, Convert.DBNull)
                dr("PurTech") = If(PurTech.Text <> "", PurTech.Text, Convert.DBNull)
                dr("PurMoral") = If(PurMoral.Text <> "", PurMoral.Text, Convert.DBNull)

                'Dim tmpValue As String = ""
                '下拉視窗
                Dim v_Degree As String = TIMS.GetListValue(Degree)
                'If Degree.SelectedIndex <> 0 Then v_Degree = TIMS.ClearSQM( Degree.SelectedValue)
                dr("CapDegree") = If(v_Degree <> "", v_Degree, Convert.DBNull)
                Dim VAL_rbl_AppStage As String = ""
                If tr_AppStage_TP28.Visible Then
                    '有啟動才檢核/塞值 '申請階段／'申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇) 
                    Dim v_rbl_AppStage As String = TIMS.GetListValue(rbl_AppStage)
                    If rbl_AppStage.SelectedIndex <> -1 Then VAL_rbl_AppStage = v_rbl_AppStage 'TIMS.ClearSQM(rbl_AppStage.SelectedValue)
                End If
                dr("AppStage") = If(VAL_rbl_AppStage <> "", VAL_rbl_AppStage, Convert.DBNull)

                Dim v_CapAge1 As String = ""
                If rdoAge1.Checked Then v_CapAge1 = "15" '15歲以上
                If rdoAge2.Checked AndAlso txtAge1.Text <> "" Then v_CapAge1 = TIMS.CINT1(txtAge1.Text) '15歲以上
                dr("CapAge1") = If(v_CapAge1 <> "", TIMS.CINT1(v_CapAge1), Convert.DBNull)
                dr("CapAge2") = Convert.DBNull '無上限
                dr("CapSex") = Convert.DBNull '性別：不區分 Convert.DBNull
                dr("CapMilitary") = "00" '兵役：不限塞00

                Dim v_Other1 As String = ""
                Dim v_Other2 As String = ""
                Dim v_Other3 As String = ""
                If Other1.Text <> "" AndAlso Other1.Text <> Cst_msgother1 Then v_Other1 = Other1.Text
                If Other2.Text <> "" AndAlso Other2.Text <> Cst_msgother1 Then v_Other2 = Other2.Text
                If Other3.Text <> "" AndAlso Other3.Text <> Cst_msgother1 Then v_Other3 = Other3.Text
                dr("CapOther1") = If(v_Other1 <> "", v_Other1, Convert.DBNull)
                dr("CapOther2") = If(v_Other2 <> "", v_Other2, Convert.DBNull)
                dr("CapOther3") = If(v_Other3 <> "", v_Other3, Convert.DBNull)

                dr("GenSciHours") = If(GenSciHours.Text <> "", GenSciHours.Text, Convert.DBNull)
                dr("ProSciHours") = If(ProSciHours.Text <> "", ProSciHours.Text, Convert.DBNull)
                dr("ProTechHours") = If(ProTechHours.Text <> "", ProTechHours.Text, Convert.DBNull)
                dr("OtherHours") = If(OtherHours.Text <> "", OtherHours.Text, Convert.DBNull)
                dr("TotalHours") = If(TotalHours.Text <> "", TotalHours.Text, Convert.DBNull)

                dr("DefGovCost") = If(DefGovCost.Text <> "", DefGovCost.Text, Convert.DBNull) '政府負擔費用
                dr("DefStdCost") = If(DefStdCost.Text <> "", DefStdCost.Text, Convert.DBNull) '學員負擔費用
                dr("TotalCost") = If(Hid_TotalCost1.Value <> "", Hid_TotalCost1.Value, Convert.DBNull)

                'ProcID 2008 拿掉，因為完全沒有用到，寫了也是白寫  by amu 2008-01-14
                '優先排序(FirstSort)
                dr("FirstSort") = If(FirstSort.Text <> "", FirstSort.Text, Convert.DBNull)
                dr("iCAPNUM") = If(iCAPNUM.Text <> "", iCAPNUM.Text, Convert.DBNull)
                dr("iCAPMARKDATE") = If(iCAPMARKDATE.Text <> "", iCAPMARKDATE.Text, Convert.DBNull)

                Dim v_Radiobuttonlist1 As String = TIMS.GetListValue(Radiobuttonlist1)
                Select Case v_Radiobuttonlist1 'Radiobuttonlist1.SelectedValue
                    Case cst_學分班 ' "Y"
                        dr("PointYN") = v_Radiobuttonlist1 'Radiobuttonlist1.SelectedValue
                    Case Else
                        'Dim v_RblSel1 As String = TIMS.ClearSQM(Radiobuttonlist1.SelectedValue)
                        dr("PointYN") = If(v_Radiobuttonlist1 <> "", v_Radiobuttonlist1, Convert.DBNull)
                End Select

                '學分種類
                Dim v_PointType As String = TIMS.GetListValue(PointType)
                'If PointType IsNot Nothing Then v_PointType = TIMS.ClearSQM(PointType.SelectedValue)
                dr("PointType") = If(v_PointType <> "", v_PointType, Convert.DBNull)

                '(充飛使用)包班種類(PackageType) 1:非包班/2:企業包班/3:聯合企業包班 
                'Dim v_PackageType As String = TIMS.GetListValue(PackageType)
                'Dim v_PackageType As String = TIMS.ClearSQM(PackageType.SelectedValue)
                dr("PackageType") = If(v_PackageType <> "", v_PackageType, Convert.DBNull)
                dr("SciPlaceID") = If(v_SciPlaceID <> "", v_SciPlaceID, Convert.DBNull) ' SciPlaceID.SelectedValue
                dr("TechPlaceID") = If(v_TechPlaceID <> "", v_TechPlaceID, Convert.DBNull) 'TechPlaceID.SelectedValue
                dr("SciPlaceID2") = If(v_SciPlaceID2 <> "", v_SciPlaceID2, Convert.DBNull) 'SciPlaceID2.SelectedValue
                dr("TechPlaceID2") = If(v_TechPlaceID2 <> "", v_TechPlaceID2, Convert.DBNull) 'TechPlaceID2.SelectedValue
                '辦理方式: '(不顯示選項)'遠距教學 'null"無遠距教學", 1."申請整班為遠距教學", 2."申請部分課程為遠距教學,3.申請整班為實體教學/無遠距教學
                If rbl_DISTANCE.Visible Then
                    dr("DISTANCE") = If(vrbl_DISTANCE <> "", vrbl_DISTANCE, Convert.DBNull)
                End If
                '遠距課程環境1/2
                dr("RMTID") = If(v_REMOTEID1 <> "", v_REMOTEID1, Convert.DBNull)
                dr("RMTID2") = If(v_REMOTEID2 <> "", v_REMOTEID2, Convert.DBNull)
                'Dim v_Taddress2 As String = TIMS.GetListValue(Taddress2)
                'Dim v_Taddress3 As String = TIMS.GetListValue(Taddress3)
                'dr("AddressSciPTID") = If(v_Taddress2 <> "", v_Taddress2, Convert.DBNull) 'Taddress2.SelectedValue
                'dr("AddressTechPTID") = If(v_Taddress3 <> "", v_Taddress3, Convert.DBNull) 'Taddress3.SelectedValue
                dr("AddressSciPTID") = If(hid_AddressSciPTID.Value <> "", hid_AddressSciPTID.Value, Convert.DBNull) 'Taddress2.SelectedValue
                dr("AddressSciPTID2") = If(hid_AddressSciPTID2.Value <> "", hid_AddressSciPTID2.Value, Convert.DBNull) 'Taddress2.SelectedValue
                dr("AddressTechPTID") = If(hid_AddressTechPTID.Value <> "", hid_AddressTechPTID.Value, Convert.DBNull) 'Taddress3.SelectedValue
                dr("AddressTechPTID2") = If(hid_AddressTechPTID2.Value <> "", hid_AddressTechPTID2.Value, Convert.DBNull) 'Taddress3.SelectedValue

                dr("TaddressZip") = If(vsTaddressZip <> "", vsTaddressZip, Convert.DBNull)
                dr("TAddress") = If(vsTAddress <> "", vsTAddress, Convert.DBNull)
                dr("TaddressZIP6W") = If(vsTaddressZIP6W <> "", vsTaddressZIP6W, Convert.DBNull)

                PointName.Text = TIMS.ClearSQM(PointName.Text)
                PackageName.Text = TIMS.ClearSQM(PackageName.Text)
                ClassName.Text = Replace(ClassName.Text, "&", "＆")
                ClassName.Text = TIMS.ClearSQM(ClassName.Text)
                If PointName.Text <> "" Then ClassName.Text = Replace(ClassName.Text, PointName.Text, "") '學分班種類
                If PackageName.Text <> "" Then ClassName.Text = Replace(ClassName.Text, PackageName.Text, "") '企業包班種類

                Dim vsClassName As String = ""
                'Dim v_Radiobuttonlist1 As String = TIMS.GetListValue(Radiobuttonlist1)
                Select Case v_Radiobuttonlist1 'Radiobuttonlist1.SelectedValue
                    Case cst_學分班 ' "Y"
                        vsClassName = ClassName.Text & PointName.Text & PackageName.Text
                    Case Else 'cst_非學分班
                        vsClassName = ClassName.Text & PackageName.Text
                End Select
                dr("ClassName") = vsClassName

                dr("Class_Unit") = If(Class_Unit.Value <> "", Class_Unit.Value, Convert.DBNull)
                dr("TNum") = If(TNum.Text <> "", TIMS.CINT1(TNum.Text), Convert.DBNull)
                dr("THours") = If(THours.Text <> "", THours.Text, Convert.DBNull)
                dr("STDate") = If(STDate.Text <> "", STDate.Text, Convert.DBNull)
                dr("FDDate") = If(FDDate.Text <> "", FDDate.Text, Convert.DBNull)

                CyclType.Text = TIMS.FmtCyclType(CyclType.Text)
                dr("CyclType") = If(CyclType.Text <> "", CyclType.Text, Convert.DBNull)

                dr("ClassCount") = If(ClassCount.Text <> "", ClassCount.Text, Convert.DBNull)
                dr("CredPoint") = If(CredPoint.Text <> "", CredPoint.Text, Convert.DBNull)
                dr("RoomName") = If(RoomName.Text <> "", RoomName.Text, Convert.DBNull)

                Dim v_FactMode As String = TIMS.GetListValue(FactMode)
                dr("FactMode") = If(v_FactMode <> "", v_FactMode, Convert.DBNull)
                Dim v_FactModeOther As String = ""
                If v_FactMode = "99" Then v_FactModeOther = TIMS.ClearSQM(FactModeOther.Text)
                dr("FactModeOther") = If(v_FactModeOther <> "", v_FactModeOther, Convert.DBNull)

                '課程內容有室外教學
                Dim v_rbl_OUTDOOR As String = TIMS.GetListValue(rbl_OUTDOOR)
                dr("OUTDOOR") = If(v_rbl_OUTDOOR <> "", v_rbl_OUTDOOR, Convert.DBNull)

                dr("ConNum") = If(ConNum.Text <> "", ConNum.Text, Convert.DBNull)
                dr("ContactName") = If(ContactName.Text <> "", ContactName.Text, Convert.DBNull)

                '2023/2024/ContactPhone
                ContactPhone.Text = TIMS.ClearSQM(ContactPhone.Text)
                ContactPhone_1.Text = TIMS.ClearSQM(ContactPhone_1.Text)
                ContactPhone_2.Text = TIMS.ClearSQM(ContactPhone_2.Text)
                ContactPhone_3.Text = TIMS.ClearSQM(ContactPhone_3.Text)
                ContactMobile_1.Text = TIMS.ClearSQM(ContactMobile_1.Text)
                ContactMobile_2.Text = TIMS.ClearSQM(ContactMobile_2.Text)
                Dim s_ContactPhone As String = If(fg_phone_2024, TIMS.ChangContactPhone(ContactPhone_1.Text, ContactPhone_2.Text, ContactPhone_3.Text), ContactPhone.Text)
                dr("ContactPhone") = If(s_ContactPhone <> "", s_ContactPhone, Convert.DBNull)
                Dim s_ContactMobile As String = TIMS.ChangContactMobile(ContactMobile_1.Text, ContactMobile_2.Text)
                dr("ContactMobile") = If(s_ContactMobile <> "", s_ContactMobile, Convert.DBNull)

                ContactEmail.Text = TIMS.ChangeEmail(ContactEmail.Text)
                dr("ContactEmail") = If(ContactEmail.Text <> "", ContactEmail.Text, Convert.DBNull)
                dr("ContactFax") = If(ContactFax.Text <> "", ContactFax.Text, Convert.DBNull)
                Dim v_ClassCate As String = TIMS.GetListValue(ClassCate)
                'Dim vsClassCate As String = TIMS.ClearSQM(ClassCate.SelectedValue)
                dr("ClassCate") = If(v_ClassCate <> "", v_ClassCate, Convert.DBNull) '訓練職能 六大職能別查詢清單
                dr("Content") = If(Content.Text <> "", Content.Text, Convert.DBNull) '課程大綱

                Dim ivTotalCost As Integer = 0
                If TIMS.IsInt(FIXSUMCOST.Text) AndAlso TIMS.IsInt(METSUMCOST.Text) Then ivTotalCost = TIMS.CINT1(FIXSUMCOST.Text) + TIMS.CINT1(METSUMCOST.Text)
                dr("TotalCost") = ivTotalCost
                dr("Note2") = If(tNote2.Text <> "", tNote2.Text, Convert.DBNull)
                dr("Note") = If(Note.Text <> "", Note.Text, Convert.DBNull)
                Select Case SaveType1
                    Case cst_SaveRcc 'cst_SaveBasic '計畫-正式儲存-正式送出
                        If Convert.ToString(dr("AppliedDate")) = "" Then dr("AppliedDate") = Now.Date
                End Select
                dr("AppliedOrigin") = 1

                Select Case SaveType1
                    Case cst_SaveRcc 'cst_SaveBasic '計畫-正式儲存-正式送出
                        If (Convert.ToString(Request("PlanID")) = "" OrElse gflag_ccopy) Then
                            '新增的狀況
                            If iPlanKind = 1 Then '自辦
                                dr("AppliedResult") = "Y" '分署內訓計畫為審核通過
                            Else '委辦
                                '分署(中心)不動
                                If sm.UserInfo.LID > 1 Then dr("AppliedResult") = Convert.DBNull '委訓清空
                            End If
                        Else
                            '修改的狀況
                            If iPlanKind = 1 Then '自辦
                                dr("AppliedResult") = "Y" '分署(職訓中心)內訓計畫為審核通過
                            Else '委辦(取得 舊值 存入 AppliedResult1)
                                re_update_flag = True
                                AppliedResult1 = Convert.ToString(dr("AppliedResult"))
                                '分署(中心)不動
                                If sm.UserInfo.LID > 1 Then dr("AppliedResult") = Convert.DBNull '委訓清空
                            End If
                        End If
                End Select

                EMail.Text = TIMS.ChangeEmail(EMail.Text)
                dr("PlanEMail") = If(EMail.Text <> "", EMail.Text, Convert.DBNull)

                '不管如何儲存，空值都是未轉
                If Convert.ToString(dr("TransFlag")) = "" Then dr("TransFlag") = "N"
                Select Case SaveType1
                    Case cst_SaveBasic
                        dr("IsApprPaper") = "Y" '正式
                    Case cst_SaveRcc 'cst_SaveBasic '計畫-正式儲存-正式送出 'cst_SaveBasic, cst_SaveRcc
                        'If Convert.ToString(dr("TransFlag")) = "" Then dr("TransFlag") = "N"
                        dr("IsApprPaper") = "Y" '正式
                End Select

                Select Case v_PackageType 'PackageType.SelectedValue
                    Case "1" '非包班
                        dr("IsBusiness") = "N"
                    Case "2", "3"   '2:企業包班,3:聯合企業包班
                        dr("IsBusiness") = "Y"
                    Case Else
                        'PackageType 空：非包班 非空：包班
                        dr("IsBusiness") = If(v_PackageType <> "", "Y", "N")
                End Select
                'dr("EnterpriseName") = EnterpriseName.Text

                'G:非勞工團體 W:勞工團體
                '1.報名時應先繳全額訓練費用，待結訓審核通過後核撥補助款
                '2.報名時應先繳50%訓練費用，待結訓審核通過後核撥補助款
                Dim v_EnterSupplyStyle As String = TIMS.GetListValue(EnterSupplyStyle)
                Select Case v_EnterSupplyStyle 'EnterSupplyStyle.SelectedValue
                    Case "1", "2" '依狀況判斷儲存必要
                        dr("EnterSupplyStyle") = v_EnterSupplyStyle 'EnterSupplyStyle.SelectedValue
                    Case Else
                        dr("EnterSupplyStyle") = "1" 'Convert.DBNull (預設為全額)
                End Select

                dr("GCID") = Convert.DBNull
                dr("GCID2") = Convert.DBNull
                dr("GCID3") = Convert.DBNull
                Select Case strYears
                    Case cst_strYears_2014 '"2014"
                        If GCIDValue.Value <> "" Then dr("GCID") = GCIDValue.Value
                    Case cst_strYears_2015 '"2015"
                        If GCIDValue.Value <> "" Then dr("GCID2") = GCIDValue.Value
                    Case cst_strYears_2018 '"2018"
                        If GCIDValue.Value <> "" Then dr("GCID3") = GCIDValue.Value
                End Select

                'dr("ResultButton") = TIMS.cst_ResultButton_尚未送出 '被修改後尚未送出 (NULL:不可送出,R(還不行送出),Y可送出)
                Select Case SaveType1
                    Case cst_SaveRcc 'cst_SaveBasic '計畫-正式儲存-正式送出 '
                        '同意送出 (Y) 'Convert.DBNull '被修改後尚未送出 (NULL:不可送出,Y待送出) 'If(sm.UserInfo.LID = 2, TIMS.cst_ResultButton_尚未送出_待送審, dr("ResultButton"))
                        dr("ResultButton") = If(fg_NOT_Ready_CLASSINFO, TIMS.cst_ResultButton_尚未送出_待送審, Convert.DBNull)
                    Case Else
                        '不同意-送出 (R) 'If(sm.UserInfo.LID = 2, TIMS.cst_ResultButton_尚未送出_未送出, dr("ResultButton"))
                        If Convert.ToString(dr("ResultButton")) <> TIMS.cst_ResultButton_尚未送出_待送審 Then
                            dr("ResultButton") = If(fg_NOT_Ready_CLASSINFO, TIMS.cst_ResultButton_尚未送出_未送出, Convert.DBNull)
                        End If
                End Select
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                'If strReportCount = "" Then strReportCount = "0"
                Dim str_CommandArgument As String = String.Concat("&ProcessType=", If(strReportCount = "0", "Insert", "Update")) 'TC_01_014.aspx用
                'If strReportCount = "0" Then str_CommandArgument = "&ProcessType=Insert" Else str_CommandArgument = "&ProcessType=Update"  'ProcessType @Insert/Update/View
                Dim sCmdArg As String = ""
                'sCmdArg = ""
                sCmdArg += "&PlanYear=" & Convert.ToString(dr("PlanYear"))
                sCmdArg += "&PlanID=" & Convert.ToString(dr("PlanID"))
                sCmdArg += "&TPlanID=" & Convert.ToString(dr("TPlanID"))
                sCmdArg += "&TMID=" & Convert.ToString(dr("TMID"))
                sCmdArg += "&RID=" & Convert.ToString(dr("RID"))
                sCmdArg += "&ComIDNO=" & Convert.ToString(dr("ComIDNO"))
                sCmdArg += "&SeqNO=" & Convert.ToString(dr("SeqNO"))
                str_CommandArgument &= sCmdArg

                'htPP.Clear()
                Dim htPP As New Hashtable From {{"TransType", s_TransType}, {"TargetTable", s_TargetTable}, {"FuncPath", s_FuncPath}, {"s_WHERE", s_WHERE}}
                TIMS.SaveTRANSLOG(sm, TransConn, Trans, dr, htPP)

                gDataRowString = TIMS.GetDataRowString(dt, dr)

                DbAccess.UpdateDataTable(dt, da, Trans)
                DbAccess.CommitTrans(Trans)

            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("gDataRowString: {0}", gDataRowString) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg6
                Exit Sub
            End Try
            'Call TIMS.CloseDbConn(conn)
        End Using

        Dim v_ddlDEPOT15 As String = TIMS.GetListValue(ddlDEPOT15)
        'Dim v_ddlKID06 As String = TIMS.GetListValue(ddlKID06)
        'Dim v_ddlKID10 As String = TIMS.GetListValue(ddlKID10)
        Dim v_ddlKID19 As String = TIMS.GetListValue(ddlKID19)
        If Not (trKID19.Visible) Then v_ddlKID19 = ""
        Dim v_ddlKID18 As String = TIMS.GetListValue(ddlKID18)

        '2022進階政策性產業類別/2025/B
        'Dim v_CBLKID22 As String = TIMS.GetCblValue(CBLKID22)
        'Dim v_CBLKID22B As String = TIMS.GetCblValue(CBLKID22B)
        'If v_CBLKID22B <> "" Then v_CBLKID22 = v_CBLKID22B
        '進階政策性產業類別/B
        Dim v_CBLKID22 As String = If(trKID25.Visible, TIMS.GetCblValue(CBLKID22B), TIMS.GetCblValue(CBLKID22)) 'v_CBLKID22

        Dim cvKID20 As String = GET_KID20_VAL() '2019 政府政策性產業
        Dim cvKID25 As String = GET_KID25_VAL() '2025 政府政策性產業
        Dim cvKID26 As String = GET_KID26_VAL() '2026 政府政策性產業
        Dim cvKID60 As String = TIMS.GetCblValue(CBLKID60)
        Dim cmdArg As String = $"&PlanID={PlanID_value}&ComIDNO={ComIDNO_value}&SeqNo={SeqNO_value}"
        cmdArg &= $"&SEQNOD15={v_ddlDEPOT15}" 'ddlDEPOT15.SelectedValue
        'cmdArg &= "&KID06=" & v_ddlKID06 'ddlKID06.SelectedValue 'Convert.ToString(drv("KID06"))
        'cmdArg &= "&KID10=" & v_ddlKID10 'ddlKID10.SelectedValue 'Convert.ToString(drv("KID10"))
        cmdArg &= $"&KID19={v_ddlKID19}" 'ddlKID19.SelectedValue 'Convert.ToString(drv("KID19"))
        cmdArg &= $"&KID18={v_ddlKID18}" 'ddlKID18.SelectedValue 'Convert.ToString(drv("KID18"))
        '2019年啟用 work2019x01:2019 政府政策性產業
        cmdArg &= $"&KID20={cvKID20}"  'TIMS.EncryptAes(vKID20) 
        cmdArg &= $"&KID22={v_CBLKID22}" 'TIMS.GetCblValue(CBLKID22) 'v_CBLKID22
        cmdArg &= $"&KID25={cvKID25}"
        cmdArg &= $"&KID26={cvKID26}"
        cmdArg &= $"&KID60={cvKID60}"
        Try
            Call SAVE_PLAN_DEPOT(cmdArg, objconn)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
            strErrmsg &= String.Concat("cmdArg: ", cmdArg) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
        End Try

        Select Case SaveType1
            Case cst_SaveBasic, cst_SaveRcc '計畫-正式儲存-正式送出
                'Select Case SaveType1,'    Case cst_SaveBasic, cst_SaveRcc,'End Select,
                '若為正式儲存且為分署(中心)，一並更改 Class_ClassInfo 資料
                'If sm.UserInfo.LID = 1 Then Call UPDATE_CLASS_CLASSINFO(PlanID_value, ComIDNO_value, SeqNO_value, conn, Trans)
                Call UPDATE_CLASS_CLASSINFO(PlanID_value, ComIDNO_value, SeqNO_value, objconn) ', Trans

                '若有 訓練計劃開班總表(產學訓) 則一並更改課程大綱(PLAN_VERREPORT)
                If TIMS.CINT1(strReportCount) > 0 Then Call UPDATE_PLAN_VERREPORT(PlanID_value, ComIDNO_value, SeqNO_value, objconn) ', Trans

                '若為正式儲存'(更新PSNO28) 目前最大值(+1) 課程申請流水號
                Hid_PSNO28.Value = TIMS.UPDATE_PSNO28xPLANINFO(PlanID_value, ComIDNO_value, SeqNO_value, vPSNO28_6)
                If Hid_PSNO28.Value = "" Then
                    sm.LastErrorMessage = cst_errmsg2b
                    Exit Sub
                End If
        End Select

        '檢核記憶體'異常
        Dim sql As String = ""
        Dim flag_session_error As Boolean = False
        'Dim PCS_V2 As String = (PlanID_value & ComIDNO_value & SeqNO_value)
        Dim dtTemp As DataTable
        Dim s_logdtValues As String = ""

        '技檢時數限定訓練業別
        '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-1才存欄位，否清空。
        Dim fg_EHourCanSave1 As Boolean = (Convert.ToString(o_TMID) = TIMS.cst_EHour_Use_TMID)

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)

            Try
                'update (97產學訓課程大綱)(PLAN_TRAINDESC) 'Plan_Teacher 'PLAN_TRAINDESC
                If TIMS.IS_DataTable(Session(hid_TrainDescTable_guid1.Value)) Then
                    dtTemp = Session(hid_TrainDescTable_guid1.Value)
                    s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_TRAINDESC")

                    da = New SqlDataAdapter
                    sql = " SELECT * FROM PLAN_TRAINDESC WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "'"
                    dt = DbAccess.GetDataTable(sql, da, Trans)
                    For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        If Not dr.RowState = DataRowState.Deleted Then
                            If gflag_ccopy OrElse dr("PTDID") <= 0 Then
                                Dim iPTDID As Integer = DbAccess.GetNewId(Trans, "PLAN_TRAINDESC_PTDID_SEQ,PLAN_TRAINDESC,PTDID")
                                dr("PTDID") = iPTDID
                            End If
                            dr("PlanID") = PlanID_value
                            dr("ComIDNO") = ComIDNO_value
                            dr("SeqNO") = SeqNO_value
                        End If
                        '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-1才存欄位，否清空。
                        If Not fg_EHourCanSave1 Then dr("EHour") = Convert.DBNull
                    Next
                    dt = dtTemp.Copy
                    DbAccess.UpdateDataTable(dt, da, Trans)
                End If
                DbAccess.CommitTrans(Trans)
                Session(hid_TrainDescTable_guid1.Value) = Nothing
            Catch ex As Exception
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)

                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                sm.LastErrorMessage = cst_errmsg7
                Exit Sub
            End Try

        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If TIMS.IS_DataTable(Session(hid_PersonCostTable_guid1.Value)) Then
                    dtTemp = Session(hid_PersonCostTable_guid1.Value)
                    s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_PERSONCOST")

                    If dtTemp.Rows.Count > 0 Then
                        sql = " SELECT * FROM PLAN_PERSONCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If gflag_ccopy OrElse dr("PPCID") <= 0 Then
                                    Dim iPPCID As Integer = DbAccess.GetNewId(Trans, "PLAN_PERSONCOST_PPCID_SEQ,PLAN_PERSONCOST,PPCID")
                                    dr("PPCID") = iPPCID
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                                dr("TNUM") = If(TNum.Text <> "", TIMS.CINT1(TNum.Text), Convert.DBNull)
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)

                Session(hid_PersonCostTable_guid1.Value) = Nothing
            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg10
                Exit Sub
            End Try

        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                'CNAME nvarchar 30'STANDARD nvarchar 300 允許'Unit nvarchar 30'PRICE numeric   允許'ALLCOUNT numeric'PURPOSE nvarchar 300 允許'TNum numeric   
                If TIMS.IS_DataTable(Session(hid_CommonCostTable_guid1.Value)) Then
                    dtTemp = Session(hid_CommonCostTable_guid1.Value)
                    s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_COMMONCOST")

                    If dtTemp.Rows.Count > 0 Then
                        sql = " SELECT * FROM PLAN_COMMONCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If gflag_ccopy OrElse dr("PCMID") <= 0 Then
                                    Dim iPCMID As Integer = DbAccess.GetNewId(Trans, "PLAN_COMMONCOST_PCMID_SEQ,PLAN_COMMONCOST,PCMID")
                                    dr("PCMID") = iPCMID
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                                dr("TNUM") = If(TNum.Text <> "", TIMS.CINT1(TNum.Text), Convert.DBNull)
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(hid_CommonCostTable_guid1.Value) = Nothing
            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg11
                Exit Sub
            End Try
        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If TIMS.IS_DataTable(Session(hid_SheetCostTable_guid1.Value)) Then
                    dtTemp = Session(hid_SheetCostTable_guid1.Value)
                    s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_SHEETCOST")

                    If dtTemp.Rows.Count > 0 Then
                        '新增'或COPY
                        sql = " SELECT * FROM PLAN_SHEETCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If gflag_ccopy OrElse dr(Cst_SheetCostpkName) <= 0 Then
                                    Dim iPSHID As Integer = DbAccess.GetNewId(Trans, "PLAN_SHEETCOST_PSHID_SEQ,PLAN_SHEETCOST,PSHID")
                                    dr(Cst_SheetCostpkName) = iPSHID
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                                dr("TNUM") = If(TNum.Text <> "", TIMS.CINT1(TNum.Text), Convert.DBNull)
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(hid_SheetCostTable_guid1.Value) = Nothing
            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg12
                Exit Sub
            End Try

        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                If TIMS.IS_DataTable(Session(hid_OtherCostTable_guid1.Value)) Then
                    dtTemp = Session(hid_OtherCostTable_guid1.Value)
                    s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_OTHERCOST")

                    If dtTemp.Rows.Count > 0 Then
                        '新增'或COPY
                        sql = " SELECT * FROM PLAN_OTHERCOST WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If gflag_ccopy OrElse dr(Cst_OtherCostpkName) <= 0 Then
                                    Dim iPOTID As Integer = DbAccess.GetNewId(Trans, "PLAN_OTHERCOST_POTID_SEQ,PLAN_OTHERCOST,POTID")
                                    dr(Cst_OtherCostpkName) = iPOTID
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                                dr("TNUM") = If(TNum.Text <> "", TIMS.CINT1(TNum.Text), Convert.DBNull)
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(hid_OtherCostTable_guid1.Value) = Nothing
            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg13
                Exit Sub
            End Try

        End Using

        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                '上課時間
                If TIMS.IS_DataTable(Session(hid_planONCLASS_guid1.Value)) Then
                    dtTemp = Session(hid_planONCLASS_guid1.Value)
                    s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_ONCLASS")

                    If dtTemp.Rows.Count > 0 Then
                        sql = " SELECT * FROM PLAN_ONCLASS WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                        dt = DbAccess.GetDataTable(sql, da, Trans)
                        For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                            If Not dr.RowState = DataRowState.Deleted Then
                                If gflag_ccopy OrElse dr("POCID") <= 0 Then
                                    Dim iPOCID As Integer = DbAccess.GetNewId(Trans, "PLAN_ONCLASS_POCID_SEQ,PLAN_ONCLASS,POCID")
                                    dr("POCID") = iPOCID
                                End If
                                dr("PlanID") = PlanID_value
                                dr("ComIDNO") = ComIDNO_value
                                dr("SeqNO") = SeqNO_value
                            End If
                        Next
                        dt = dtTemp.Copy
                        DbAccess.UpdateDataTable(dt, da, Trans)
                    End If
                End If
                DbAccess.CommitTrans(Trans)
                Session(hid_planONCLASS_guid1.Value) = Nothing
            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg14
                Exit Sub
            End Try

        End Using

        'Dim v_PackageType As String = TIMS.GetListValue(PackageType)
        Using TransConn As SqlConnection = DbAccess.GetConnection()
            Dim Trans As SqlTransaction = DbAccess.BeginTrans(TransConn)
            Try
                s_logdtValues = ""
                '計畫包班事業單位
                Select Case v_PackageType 'PackageType.SelectedValue
                    Case "3" '充電起飛計畫 '聯合企業包班
                        If TIMS.IS_DataTable(Session(hid_PLAN_BUSPACKAGE_guid1.Value)) Then
                            dtTemp = Session(hid_PLAN_BUSPACKAGE_guid1.Value)
                            s_logdtValues = TIMS.GET_AllDataTableValues(dtTemp, "PLAN_BUSPACKAGE")

                            If dtTemp.Rows.Count > 0 Then
                                sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                                dt = DbAccess.GetDataTable(sql, da, Trans)
                                For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                                    If Not dr.RowState = DataRowState.Deleted Then
                                        If gflag_ccopy OrElse dr("BPID") <= 0 Then
                                            Dim iBPID As Integer = DbAccess.GetNewId(Trans, "PLAN_BUSPACKAGE_BPID_SEQ,PLAN_BUSPACKAGE,BPID")
                                            dr("BPID") = iBPID
                                        End If
                                        dr("PlanID") = PlanID_value
                                        dr("ComIDNO") = ComIDNO_value
                                        dr("SeqNO") = SeqNO_value
                                    End If
                                Next
                                dt = dtTemp.Copy
                                DbAccess.UpdateDataTable(dt, da, Trans)
                            End If
                        End If

                    Case "2" '充電起飛計畫' 企業包班(只有1筆)
                        txtUname.Text = TIMS.ClearSQM(txtUname.Text)
                        txtIntaxno.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txtIntaxno.Text))
                        txtUbno.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(txtUbno.Text))
                        If txtUname.Text <> "" AndAlso txtUname.Text.Length > 50 Then txtUname.Text = txtUname.Text.Substring(0, 50)
                        If txtIntaxno.Text <> "" AndAlso txtIntaxno.Text.Length > 10 Then txtIntaxno.Text = txtIntaxno.Text.Substring(0, 10)
                        If txtUbno.Text <> "" AndAlso txtUbno.Text.Length > 9 Then txtUbno.Text = txtUbno.Text.Substring(0, 9)
                        Dim s_Uname As String = If(txtUname.Text.Length > 50, TIMS.Get_Substr1(txtUname.Text, 50), txtUname.Text)
                        Dim s_Intaxno As String = If(txtIntaxno.Text.Length > 10, TIMS.Get_Substr1(txtIntaxno.Text, 10), txtIntaxno.Text)
                        Dim s_Ubno As String = If(txtUbno.Text.Length > 9, TIMS.Get_Substr1(txtUbno.Text, 9), txtUbno.Text)
                        Dim fg_can_save_1 As Boolean = ((s_Intaxno <> "" AndAlso TIMS.IsNumberStr(s_Intaxno)) OrElse (s_Intaxno = ""))
                        Dim fg_can_save_2 As Boolean = ((s_Ubno <> "" AndAlso TIMS.CheckABC123(s_Ubno)) OrElse (s_Ubno = ""))

                        If (fg_can_save_1 AndAlso fg_can_save_2) Then
                            sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                            Dim dtCB As DataTable = DbAccess.GetDataTable(sql, da, Trans)
                            If TIMS.dtHaveDATA(dtCB) Then
                                sql = " DELETE PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                                DbAccess.ExecuteNonQuery(sql, Trans)
                            End If

                            sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "' "
                            dt = DbAccess.GetDataTable(sql, da, Trans)
                            Dim dr As DataRow = dt.NewRow
                            dt.Rows.Add(dr)
                            Dim iBPID As Integer = DbAccess.GetNewId(Trans, "PLAN_BUSPACKAGE_BPID_SEQ,PLAN_BUSPACKAGE,BPID")
                            dr("BPID") = iBPID
                            If (Convert.ToString(Request("PlanID")) = "" OrElse gflag_ccopy) Then
                                dr("PlanID") = PlanID_value 'sm.UserInfo.PlanID
                                dr("ComIDNO") = ComidValue.Value
                                dr("SeqNO") = SeqNO_value ' ViewState("SeqNO")
                            Else
                                dr("PlanID") = TIMS.ClearSQM(Request("PlanID"))
                                dr("ComIDNO") = TIMS.ClearSQM(Request("ComIDNO"))
                                dr("SeqNO") = TIMS.ClearSQM(Request("SeqNO"))
                            End If
                            dr("Uname") = s_Uname 'txtUname.Text 'Convert.ToString(txtUname.Text.Trim)
                            dr("Intaxno") = If(s_Intaxno <> "", s_Intaxno, Convert.DBNull)
                            dr("Ubno") = If(s_Ubno <> "", s_Ubno, Convert.DBNull)
                            dr("ModifyAcct") = sm.UserInfo.UserID
                            dr("ModifyDate") = Now
                            s_logdtValues = TIMS.GET_AllDataTableValues(dt, "PLAN_BUSPACKAGE")
                            DbAccess.UpdateDataTable(dt, da, Trans)
                        End If

                    Case Else '清除
                        sql = " SELECT * FROM PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "'"
                        Dim dtCB As DataTable = DbAccess.GetDataTable(sql, da, Trans)
                        If TIMS.dtHaveDATA(dtCB) Then
                            sql = " DELETE PLAN_BUSPACKAGE WHERE PlanID='" & PlanID_value & "' AND ComIDNO='" & ComIDNO_value & "' AND SeqNO='" & SeqNO_value & "'"
                            DbAccess.ExecuteNonQuery(sql, Trans)
                        End If

                End Select
                DbAccess.CommitTrans(Trans)
                Session(hid_PLAN_BUSPACKAGE_guid1.Value) = Nothing
            Catch ex As Exception
                upt_PlanX.Value = ""
                DbAccess.RollbackTrans(Trans)
                Call TIMS.CloseDbConn(TransConn)
                'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Dim strErrmsg As String = ""
                strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", PlanID_value, ComIDNO_value, SeqNO_value) & vbCrLf
                strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                strErrmsg &= String.Format("/* ex.Message:{0} */", ex.Message) & vbCrLf
                strErrmsg &= String.Concat("logdtValues: ", s_logdtValues) & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Page, ex) '取得錯誤資訊寫入
                Call TIMS.WriteTraceLog(strErrmsg)
                sm.LastErrorMessage = cst_errmsg15
                Exit Sub
            End Try
        End Using
        'Call TIMS.CloseDbConn(TransConn) '結束資料連線

        '2009年產業人才投資方案班級審核改為分署(中心)直接複審 BY AMU
        If re_update_flag Then
            Select Case AppliedResult1
                Case "Y", "O", "N"
                Case Else '"R"
                    Call TIMS.Plan_VerRecord_Update(PlanID_value, ComIDNO_value, SeqNO_value, objconn)
                    Call TIMS.PLAN_VERREPROT_UPDATE(Me, PlanID_value, ComIDNO_value, SeqNO_value, "O", objconn)
            End Select
        End If

        '儲存 開班計畫表資料維護/開班計劃表資料維護
        Call SAVE_PLAN_VERREPORT(SaveType1) 'cst_SaveRcc)'計畫-正式儲存-正式送出
        '儲存 班級申請老師
        Call SAVE_PLAN_TEACHER(objconn)
        '專長能力標籤-ABILITY
        Call SAVE_PLAN_ABILITY()

        Dim drPP As DataRow = TIMS.GetPCSDate(PlanID_value, ComIDNO_value, SeqNO_value, objconn)
        If drPP IsNot Nothing AndAlso Convert.ToString(drPP("OCID")) <> "" Then Call SAVE_CLASS_TEACHER(drPP("OCID"), objconn) '修改班級師資資料

        '儲存-政策性產業課程可辦理班數
        'If flag_SHOW_2019_3 Then Call SAVE_PLAN_PRECLASS()
        '☆2011-12-11 for 28 ,54,56計畫，在正式儲存後要先提示訊息
        'If SaveType1 = cst_SaveBasic Then Common.RespWrite(Me, "<script language='javascript'>alert('請記得於班級查詢後，針對本班按下【送出】鍵！');</script>")

        If TIMS.ClearSQM(Request("PlanID")) = "" Then
            Select Case SaveType1
                Case cst_SaveBasic
                    'Session(cst_session_saveok) = True
                    s_Msg1 = cst_msg_save1
                    If Hid_ChkTEACHHOURS1.Value.Equals("Y") Then
                        '同一位師資授課時數已超過54小時， 請確認是否為特殊情況。
                        s_Msg1 &= vbCrLf & cst_msg_save4
                    End If
                    Common.MessageBox(Me, s_Msg1)
                    Dim url1 As String = "../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                    TIMS.Utl_Redirect(Me, objconn, url1)
                    'Common.RespWrite(Me, "<script>alert('基本儲存成功!!');location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "';</script>")

                Case cst_SaveRcc '計畫-正式儲存-正式送出
                    'Session(cst_session_saveok) = True
                    s_Msg1 = cst_msg_save2
                    If Hid_ChkTEACHHOURS1.Value.Equals("Y") Then
                        '同一位師資授課時數已超過54小時， 請確認是否為特殊情況。
                        s_Msg1 &= vbCrLf & cst_msg_save4
                    End If
                    Common.MessageBox(Me, s_Msg1)
                    Dim url1 As String = "../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                    TIMS.Utl_Redirect(Me, objconn, url1)
                    'Common.RespWrite(Me, "<script>alert('正式儲存成功!!');location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "';</script>")

                Case Else
                    '這一格應該是沒有，因為正式儲存後，應該不會有草稿儲存。  (新增的)草稿儲存
                    s_Msg1 = cst_msg_save3
                    Common.RespWrite(Me, "<script>alert('" & s_Msg1 & "');</script>")
                    Call CreateClassTime()
                    Call CreateTrainDesc() 'PLAN_TRAINDESC
                    Call CreateBusPackage()
                    Call CreatePersonCost()
                    Call CreateCommonCost()

            End Select

        Else
            'If ViewState("search") <> "" Then Session("search") = ViewState("search")
            If Convert.ToString(ViewState("search")) <> "" AndAlso Session("search") Is Nothing Then Session("search") = ViewState("search")

            Select Case SaveType1
                Case cst_SaveBasic
                    '正式送出 'Session(cst_session_saveok) = True 'Dim strScriptRcc1 As String = "" '班級查詢作業/'班級複製作業 '班級查詢作業
                    'strScriptRcc1 = "<script>alert('基本儲存成功!!');location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "';</script>"
                    s_Msg1 = cst_msg_save1
                    If Hid_ChkTEACHHOURS1.Value.Equals("Y") Then
                        '同一位師資授課時數已超過54小時， 請確認是否為特殊情況。
                        s_Msg1 &= vbCrLf & cst_msg_save4
                    End If
                    If strReportCount > "0" AndAlso gflag_ccopy Then
                        '班級複製作業
                        'strScriptRcc1 = "<script>alert('基本儲存成功!!');location.href='../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "';</script>"
                        Common.MessageBox(Me, s_Msg1)
                        Dim url1 As String = "../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                        TIMS.Utl_Redirect(Me, objconn, url1)
                    Else
                        Common.MessageBox(Me, s_Msg1)
                        Dim url1 As String = "../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                        TIMS.Utl_Redirect(Me, objconn, url1)
                    End If
                    'Common.RespWrite(Me, strScriptRcc1)

                Case cst_SaveRcc '計畫-正式儲存-正式送出
                    '正式送出 'Session(cst_session_saveok) = True 'Dim strScriptRcc1 As String = "" '班級查詢作業/'班級複製作業 '班級查詢作業
                    'strScriptRcc1 = "<script>alert('正式儲存成功!!');location.href='../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "';</script>"
                    s_Msg1 = cst_msg_save2
                    If Hid_ChkTEACHHOURS1.Value.Equals("Y") Then
                        '同一位師資授課時數已超過54小時， 請確認是否為特殊情況。
                        s_Msg1 &= vbCrLf & cst_msg_save4
                    End If
                    If strReportCount > "0" AndAlso gflag_ccopy Then
                        '班級複製作業
                        'strScriptRcc1 = "<script>alert('正式儲存成功!!');location.href='../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID")) & "';</script>"
                        Common.MessageBox(Me, s_Msg1)
                        Dim url1 As String = "../03/TC_03_002.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                        TIMS.Utl_Redirect(Me, objconn, url1)
                    Else
                        Common.MessageBox(Me, s_Msg1)
                        Dim url1 As String = "../02/TC_02_001.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
                        TIMS.Utl_Redirect(Me, objconn, url1)
                    End If
                    'Common.RespWrite(Me, strScriptRcc1)

                Case Else
                    '草稿儲存
                    s_Msg1 = cst_msg_save3
                    Common.RespWrite(Me, "<script>alert('" & s_Msg1 & "');</script>")
                    Call CreateClassTime()
                    Call CreateTrainDesc() 'PLAN_TRAINDESC
                    Call CreateBusPackage() 'xx
                    Call CreatePersonCost()
                    Call CreateCommonCost()
                    Call CreateSheetCost()
                    Call CreateOtherCost()
            End Select

        End If
    End Sub

#End Region

    Protected Sub BtnDelBatchDGX_Click(sender As Object, e As EventArgs) Handles BtnDelDG6.Click, BtnDelDG7.Click, BtnDelDG8.Click, BtnDelDG9.Click
        If sender Is Nothing Then Return
        Dim BtnDel As Button = If(sender Is Nothing, Nothing, CType(sender, Button))
        Select Case BtnDel.CommandName
            Case "BtnDelDG6" '批次刪除 一人份材料明細
                Call DelBatchXCost(Cst_PersonCostpkName)
            Case "BtnDelDG7" '批次刪除 共同材料明細
                Call DelBatchXCost(Cst_CommonCostpkName)
            Case "BtnDelDG8" '批次刪除 教材明細
                Call DelBatchXCost(Cst_SheetCostpkName)
            Case "BtnDelDG9" '批次刪除 其他明細
                Call DelBatchXCost(Cst_OtherCostpkName)
        End Select
    End Sub

    Sub DelBatchXCost(ByVal sCostpkName As String)
        'Dim dt As DataTable = Nothing
        Select Case sCostpkName
            Case Cst_PersonCostpkName '批次刪除 一人份材料明細
                Dim dt As DataTable = Nothing
                dt = If(Session(hid_PersonCostTable_guid1.Value), CreatePersonCost())
                If dt IsNot Nothing Then DelBatchPersonCost(dt, DataGrid6)
                Session(hid_PersonCostTable_guid1.Value) = dt
                Call CreatePersonCost()
            Case Cst_CommonCostpkName '批次刪除 共同材料明細
                Dim dt As DataTable = Nothing
                dt = If(Session(hid_CommonCostTable_guid1.Value), CreateCommonCost())
                If dt IsNot Nothing Then DelBatchCommonCost(dt, DataGrid7)
                Session(hid_CommonCostTable_guid1.Value) = dt
                Call CreateCommonCost()
            Case Cst_SheetCostpkName '批次刪除 教材明細
                Dim dt As DataTable = Nothing
                dt = If(Session(hid_SheetCostTable_guid1.Value), CreateSheetCost())
                If dt IsNot Nothing Then DelBatchSheetCost(dt, DataGrid8)
                Session(hid_SheetCostTable_guid1.Value) = dt
                Call CreateSheetCost()
            Case Cst_OtherCostpkName '批次刪除 其他明細
                Dim dt As DataTable = Nothing
                dt = If(Session(hid_OtherCostTable_guid1.Value), CreateOtherCost())
                If dt IsNot Nothing Then DelBatchOtherCost(dt, DataGrid9)
                Session(hid_OtherCostTable_guid1.Value) = dt
                Call CreateOtherCost()
        End Select
        If (LayerState.Value = "") Then LayerState.Value = "6"
        Dim s_js1 As String = String.Concat("<script language=""javascript"">Layer_change(", LayerState.Value, ");</script>")
        Page.RegisterStartupScript("Londing", s_js1)
    End Sub

    Private Sub DelBatchPersonCost(ByRef dt As DataTable, ByRef DGobj As DataGrid)
        'Dim DGobj As DataGrid = DataGrid6
        If DGobj Is Nothing OrElse dt Is Nothing Then
            sm.LastErrorMessage = cst_errmsg16
            Exit Sub
        End If
        Dim iCNT As Integer = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim CheckBoxDG As CheckBox = eItem.FindControl("CheckBoxDG6")
            Dim Hid_DataKey As HiddenField = eItem.FindControl("Hid_DataKey")
            Dim fg_continue As Boolean = (CheckBoxDG Is Nothing OrElse Hid_DataKey Is Nothing)
            If fg_continue Then Continue For
            Dim v_DataKeys As String = If(CheckBoxDG.Checked, TIMS.DecryptAes(Hid_DataKey.Value), "")
            If CheckBoxDG.Checked AndAlso v_DataKeys <> "" Then
                Dim sfilter As String = String.Concat(Cst_PersonCostpkName, "='", v_DataKeys, "'")
                '搜尋刪除資料刪除
                If dt.Select(sfilter).Length <> 0 Then
                    iCNT += 1
                    For Each dr As DataRow In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            End If
        Next
        If iCNT = 0 Then sm.LastErrorMessage = cst_errmsg10
    End Sub

    Private Sub DelBatchCommonCost(ByRef dt As DataTable, ByRef DGobj As DataGrid)
        'Dim DGobj As DataGrid = DataGrid7
        If DGobj Is Nothing OrElse dt Is Nothing Then
            sm.LastErrorMessage = cst_errmsg16
            Exit Sub
        End If
        Dim iCNT As Integer = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim CheckBoxDG As CheckBox = eItem.FindControl("CheckBoxDG7")
            Dim Hid_DataKey As HiddenField = eItem.FindControl("Hid_DataKey")
            Dim fg_continue As Boolean = (CheckBoxDG Is Nothing OrElse Hid_DataKey Is Nothing)
            If fg_continue Then Continue For
            Dim v_DataKeys As String = If(CheckBoxDG.Checked, TIMS.DecryptAes(Hid_DataKey.Value), "")
            If CheckBoxDG.Checked AndAlso v_DataKeys <> "" Then
                Dim sfilter As String = String.Concat(Cst_CommonCostpkName, "='", v_DataKeys, "'")
                '搜尋刪除資料刪除
                If dt.Select(sfilter).Length <> 0 Then
                    iCNT += 1
                    For Each dr As DataRow In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            End If
        Next
        If iCNT = 0 Then sm.LastErrorMessage = cst_errmsg11
    End Sub
    Private Sub DelBatchSheetCost(ByRef dt As DataTable, ByRef DGobj As DataGrid)
        'Dim DGobj As DataGrid = DataGrid8
        If DGobj Is Nothing OrElse dt Is Nothing Then
            sm.LastErrorMessage = cst_errmsg16
            Exit Sub
        End If
        Dim iCNT As Integer = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim CheckBoxDG As CheckBox = eItem.FindControl("CheckBoxDG8")
            Dim Hid_DataKey As HiddenField = eItem.FindControl("Hid_DataKey")
            Dim fg_continue As Boolean = (CheckBoxDG Is Nothing OrElse Hid_DataKey Is Nothing)
            If fg_continue Then Continue For
            Dim v_DataKeys As String = If(CheckBoxDG.Checked, TIMS.DecryptAes(Hid_DataKey.Value), "")
            If CheckBoxDG.Checked AndAlso v_DataKeys <> "" Then
                Dim sfilter As String = String.Concat(Cst_SheetCostpkName, "='", v_DataKeys, "'")
                '搜尋刪除資料刪除
                If dt.Select(sfilter).Length <> 0 Then
                    iCNT += 1
                    For Each dr As DataRow In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            End If
        Next
        If iCNT = 0 Then sm.LastErrorMessage = cst_errmsg12
    End Sub

    Private Sub DelBatchOtherCost(ByRef dt As DataTable, ByRef DGobj As DataGrid)
        'Dim DGobj As DataGrid = DataGrid9
        If DGobj Is Nothing OrElse dt Is Nothing Then
            sm.LastErrorMessage = cst_errmsg16
            Exit Sub
        End If
        Dim iCNT As Integer = 0
        For Each eItem As DataGridItem In DGobj.Items
            Dim CheckBoxDG As CheckBox = eItem.FindControl("CheckBoxDG9")
            Dim Hid_DataKey As HiddenField = eItem.FindControl("Hid_DataKey")
            Dim fg_continue As Boolean = (CheckBoxDG Is Nothing OrElse Hid_DataKey Is Nothing)
            If fg_continue Then Continue For
            Dim v_DataKeys As String = If(CheckBoxDG.Checked, TIMS.DecryptAes(Hid_DataKey.Value), "")
            If CheckBoxDG.Checked AndAlso v_DataKeys <> "" Then
                Dim sfilter As String = String.Concat(Cst_OtherCostpkName, "='", v_DataKeys, "'")
                '搜尋刪除資料刪除
                If dt.Select(sfilter).Length <> 0 Then
                    iCNT += 1
                    For Each dr As DataRow In dt.Select(sfilter)
                        If dr.RowState <> DataRowState.Deleted Then dr.Delete() '刪除
                    Next
                End If
            End If
        Next
        If iCNT = 0 Then sm.LastErrorMessage = cst_errmsg13
    End Sub

    ''' <summary>專長能力標籤-ABILITY</summary>
    Private Sub SHOW_PLAN_ABILITYS()
        If g_flagNG Then
            sm.LastErrorMessage = cst_errmsg3
            Exit Sub
        End If
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNO As String = TIMS.ClearSQM(Request("SeqNO"))
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return 'rst 'Exit Sub

        Dim oParms As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}}
        Dim sSql As String = ""
        sSql &= " SELECT PABID,PLANID,COMIDNO,SEQNO,SEQ_ID,ABILITY,ABILITY_DESC"
        sSql &= " FROM PLAN_ABILITY"
        sSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        'sSql &= " ORDER BY SEQ_ID DESC" & vbCrLf
        Dim tb_dt As DataTable = DbAccess.GetDataTable(sSql, objconn, oParms)

        For Each tb_dr As DataRow In tb_dt.Rows
            Dim s_SEQ As String = Convert.ToString(tb_dr("SEQ_ID"))
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            If otxtA1 IsNot Nothing Then otxtA1.Text = Convert.ToString(tb_dr("ABILITY"))
            If otxtA2 IsNot Nothing Then otxtA2.Text = Convert.ToString(tb_dr("ABILITY_DESC"))
        Next

    End Sub

    ''' <summary>專長能力標籤-ABILITY</summary>
    Sub SAVE_PLAN_ABILITY()
        If upt_PlanX.Value = "" Then Return
        tmpPCS = upt_PlanX.Value  '有儲存資料過了
        PlanID_value = TIMS.GetMyValue(tmpPCS, "PlanID")
        ComIDNO_value = TIMS.GetMyValue(tmpPCS, "ComIDNO")
        SeqNO_value = TIMS.GetMyValue(tmpPCS, "SeqNO")
        'If upt_PlanX.Value = "" Then Exit Sub '無有效值離開
        Dim rqPlanID As String = PlanID_value 'TIMS.GetMyValue2(htSS, "rqPlanID")
        Dim rqComIDNO As String = ComIDNO_value 'TIMS.GetMyValue2(htSS, "rqComIDNO")
        Dim rqSeqNO As String = SeqNO_value 'TIMS.GetMyValue2(htSS, "rqSeqNO")
        If rqPlanID = "" OrElse rqComIDNO = "" OrElse rqSeqNO = "" Then Return '(有異常離開)

        Dim iRst As Integer = 0
        For i_SEQ As Integer = 1 To 4
            Dim s_SEQ As String = Convert.ToString(i_SEQ)
            Dim otxtA1 As TextBox = If(s_SEQ = "1", txtABILITY1, If(s_SEQ = "2", txtABILITY2, If(s_SEQ = "3", txtABILITY3, If(s_SEQ = "4", txtABILITY4, Nothing))))
            Dim otxtA2 As TextBox = If(s_SEQ = "1", txtABILITY_DESC1, If(s_SEQ = "2", txtABILITY_DESC2, If(s_SEQ = "3", txtABILITY_DESC3, If(s_SEQ = "4", txtABILITY_DESC4, Nothing))))
            otxtA1.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA1.Text), 30)
            otxtA2.Text = TIMS.Get_Substr1(TIMS.ClearSQM(otxtA2.Text), 200)

            Dim pms_s1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ}}
            Dim sql_s1 As String = "SELECT PABID FROM PLAN_ABILITY WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SEQ_ID=@SEQ_ID"
            Dim dt As DataTable = DbAccess.GetDataTable(sql_s1, objconn, pms_s1)
            If otxtA1.Text <> "" Then
                If dt Is Nothing OrElse dt.Rows.Count = 0 Then
                    Dim iPABID As Integer = DbAccess.GetNewId(objconn, "PLAN_ABILITY_PABID_SEQ,PLAN_ABILITY,PABID")
                    Dim iParms As New Hashtable From {{"PABID", iPABID},
                        {"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ},
                        {"ABILITY", otxtA1.Text}, {"ABILITY_DESC", otxtA2.Text}, {"MODIFYACCT", sm.UserInfo.UserID}}
                    Dim isSql As String = ""
                    isSql &= " INSERT INTO PLAN_ABILITY(PABID, PLANID, COMIDNO, SEQNO, SEQ_ID, ABILITY, ABILITY_DESC, MODIFYACCT, MODIFYDATE)" & vbCrLf
                    isSql &= " VALUES(@PABID,@PLANID,@COMIDNO,@SEQNO,@SEQ_ID,@ABILITY,@ABILITY_DESC,@MODIFYACCT,GETDATE())" & vbCrLf
                    iRst = DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
                Else
                    Dim iPABID As Integer = CInt(dt.Rows(0)("PABID"))
                    Dim uParms As New Hashtable From {{"PABID", iPABID},
                        {"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ},
                        {"ABILITY", otxtA1.Text}, {"ABILITY_DESC", otxtA2.Text}, {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    Dim usSql As String = ""
                    usSql &= " UPDATE PLAN_ABILITY" & vbCrLf
                    usSql &= " SET ABILITY=@ABILITY,ABILITY_DESC=@ABILITY_DESC,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                    usSql &= " WHERE PABID=@PABID AND PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SEQ_ID=@SEQ_ID" & vbCrLf
                    iRst = DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
                End If
            Else
                If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
                    Dim pms_d1 As New Hashtable From {{"PLANID", rqPlanID}, {"COMIDNO", rqComIDNO}, {"SEQNO", rqSeqNO}, {"SEQ_ID", i_SEQ}}
                    Dim sql_d1 As String = "DELETE PLAN_ABILITY WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND SEQ_ID=@SEQ_ID"
                    iRst = DbAccess.ExecuteNonQuery(sql_d1, objconn, pms_d1)
                End If
            End If

        Next

    End Sub

End Class
