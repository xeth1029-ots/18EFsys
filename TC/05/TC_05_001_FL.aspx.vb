Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Partial Class TC_05_001_FL
    Inherits AuthBasePage

    'Private Shared ReadOnly FileUpload_lock As New Object
    'KEY_REVISESUB 'PLAN_REVISESUBFL 班級變更申請-上傳檔案
    'PLAN_REVISE Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
    Dim drPR As DataRow = Nothing
    'PLAN_PLANINFO / Dim drPP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
    Dim drPP As DataRow = Nothing
    'VIEW_RIDNAME / Dim drRR As DataRow = TIMS.Get_RID_DR(Convert.ToString(drPP("RID")), objconn)
    Dim drRR As DataRow = Nothing
    'AltDataID 變更項目 '於1:「開結訓日」、15:「上課時間」、14:「上課地點」、9:「停辦」等變更項目，新增「其他應備文件」欄位，放在公文項目前面，    

    '最近一次版本送件
    Const cst_MTYPE_LATEST_SEND1 As String = "MTYPE_LATEST_SEND1"
    '最近一次版本-下載
    Const cst_MTYPE_LATEST_DOWN1 As String = "MTYPE_LATEST_DOWN1"

    Dim fg_REVISESUBFL_VIEW1 As Boolean = False
    Const cst_tpmsg_enb1 As String = "(檢視功能)不能儲存"
    Const cst_tpmsg_enb2 As String = "(檢視功能)不能檔案上傳"
    Const cst_tpmsg_enb2b As String = "(項目參數)不能檔案上傳"
    Const cst_tpmsg_enb3 As String = "(檢視功能)不能送出"
    Const cst_tpmsg_enb9 As String = "尚未上傳文件"

    Const cst_labTitle_申請狀態 As String = "申請狀態"
    Const cst_PointYN_非學分班 As String = "非學分班"
    Const cst_PointYN_學分班 As String = "學分班"

    'Const cst_SearchMode_變更結果 As String = "變更結果"
    Const cst_SearchMode_線上送件 As String = "線上送件"
    Const cst_CheckMode_審核中 As String = "審核中"
    Const cst_CheckMode_審核通過 As String = "審核通過"
    Const cst_CheckMode_審核不通過 As String = "審核不通過"

    '10_師資助教基本資料表_WAIVED_TT
    Dim iDG10_ROWS As Integer = 0

    Const cst_txt_免附文件 As String = "(免附文件)"
    Const cst_txt_檔案下載 As String = "(檔案下載..)"
    'TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表,
    'TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本,
    'TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表
    Const cst_10_師資助教學經歷證書資料表_WAIVED_TT As String = "TT"

    Const cst_errMsg_1 As String = "資料有誤請重新查詢!"
    Const cst_errMsg_2 As String = "上傳檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)" 'Const cst_errMsg_2 As String = "上傳檔案壓縮時發生錯誤，請重新確認上傳檔案格式!"
    Const cst_errMsg_3 As String = "檔案位置錯誤!"
    Const cst_errMsg_4 As String = "檔案類型錯誤!"
    Const cst_errMsg_5 As String = "檔案類型錯誤，必須為PDF類型檔案!"
    Const cst_errMsg_6 As String = "(檔案上傳失敗／異常，請刪除後重新上傳)"
    Const cst_PostedFile_MAX_SIZE_10M As Integer = 10485760 '10*1024*1024=10,485,760  '2*1024*1024=2,097,152
    Const cst_PostedFile_MAX_SIZE_15M As Integer = 15728640 '1024*1024*15=15728640
    'Const cst_errMsg_7 As String = "檔案大小超過2MB!"
    Const cst_errMsg_7_10M As String = "檔案大小超過10MB!"
    Const cst_errMsg_7_15M As String = "檔案大小超過15MB!"
    Const cst_FileDescMsg_7_10M As String = "PDF(掃瞄畫面需清楚，檔案大小限制10MB以下)!"
    Const cst_FileDescMsg_7_15M As String = "PDF(掃瞄畫面需清楚，檔案大小限制15MB以下)!"

    Const cst_errMsg_8 As String = "請選擇上傳檔案(不可為空)!"
    'Const cst_errMsg_9 As String = "請選擇場地圖片--隸屬於教室1 或教室2!"
    'Const cst_errMsg_11 As String = "無效的檔案格式。"
    Const cst_errMsg_11_PDF As String = "無效的檔案格式(限PDF檔案)。"
    Const cst_errMsg_21 As String = "不可勾選免附文件又按上傳檔案。"

    ''' <summary>儲存(暫存)</summary>
    Const cst_ACTTYPE_BTN_SAVETMP1 As String = "BTN_SAVETMP1" '儲存(暫存)
    ''' <summary>'儲存後進下一步</summary>
    Const cst_ACTTYPE_BTN_SAVENEXT1 As String = "BTN_SAVENEXT1" '儲存後進下一步

    'Const cst_BTN_SENDCURRVER_SAVESEND1 As String = "BTN_SENDCURRVER_SAVESEND1" '以目前版本送出
    'Const cst_BTN_SENDCURRVER_DOWNLOAD1 As String = "BTN_SENDCURRVER_DOWNLOAD1" '以目前版本送出(下載)

    'vGW_ALT_RVID
    'Const cst_G_1_01_訓練計畫變更表 As String = "G_1_01"
    'Const cst_G_1_02_變更後之課程表 As String = "G_1_02"
    ''Const cst_G_1_03_TTQS評核展延證明文件 As String = "G_1_03"
    ''Const cst_G_1_04_公文 As String = "G_1_04"

    '師資／助教基本資料表
    'Const cst_G_9_01_訓練計畫變更表 As String = "G_9_01"
    ''Const cst_G_9_02_公文 As String = "G_9_02"
    'Const cst_G_11_01_訓練計畫變更表 As String = "G_11_01"
    'Const cst_G_11_02_變更後訓練計畫師資名冊 As String = "G_11_02"
    'Const cst_G_11_03_變更前訓練計畫師資名冊 As String = "G_11_03"
    'Const cst_G_11_04_變更後師資助教基本資料表 As String = "G_11_04"
    'Const cst_G_11_05_變更前師資助教基本資料表 As String = "G_11_05"
    ''Const cst_G_11_06_各授課師資學經歷證書影本 As String = "G_11_06"
    'Const cst_G_11_07_變更後之課程表 As String = "G_11_07"
    ''Const cst_G_11_08_公文 As String = "G_11_08"

    'Const cst_G_14_01_訓練計畫變更表 As String = "G_14_01"
    'Const cst_G_14_02_訓練計畫場地資料表 As String = "G_14_02"
    'Const cst_G_14_03_教學環境資料表 As String = "G_14_03"
    ''Const cst_G_14_04_消防安全設備檢修申報受理單 As String = "G_14_04"
    ''Const cst_G_14_05_建築物防火避難設施與設備安全檢查申報結果通知書 As String = "G_14_05"
    ''Const cst_G_14_06_檢附校方或機關同意租借證明文件 As String = "G_14_06"
    'Const cst_G_14_07_變更後之課程表 As String = "G_14_07"
    ''Const cst_G_14_08_公文 As String = "G_14_08"

    'Const cst_G_15_01_訓練計畫變更表 As String = "G_15_01"
    'Const cst_G_15_02_變更後之課程表 As String = "G_15_02"
    ''Const cst_G_15_03_公文 As String = "G_15_03"

    'Const cst_G_18_01_訓練計畫變更表 As String = "G_18_01"
    'Const cst_G_18_02_變更後之課程表 As String = "G_18_02"
    ''Const cst_G_18_03_公文 As String = "G_18_03"

    'Const cst_G_20_01_訓練計畫變更表 As String = "G_20_01"
    'Const cst_G_20_02_變更後師資助教基本資料表 As String = "G_20_02"
    'Const cst_G_20_03_變更後之課程表 As String = "G_20_03"
    ''Const cst_G_20_04_公文 As String = "G_20_04"

    'Const cst_W_1_01_訓練計畫變更表 As String = "W_1_01"
    'Const cst_W_1_02_變更後之課程表 As String = "W_1_02"
    ''Const cst_W_1_03_TTQS評核展延證明文件 As String = "W_1_03"
    ''Const cst_W_1_04_公文 As String = "W_1_04"
    'Const cst_W_9_01_訓練計畫變更表 As String = "W_9_01"
    ''Const cst_W_9_02_公文 As String = "W_9_02"

    'Const cst_W_11_01_訓練計畫變更表 As String = "W_11_01"
    'Const cst_W_11_02_變更後訓練計畫師資名冊 As String = "W_11_02"
    'Const cst_W_11_03_變更前訓練計畫師資名冊 As String = "W_11_03"
    'Const cst_W_11_04_變更後師資助教基本資料表 As String = "W_11_04"
    'Const cst_W_11_05_變更前師資助教基本資料表 As String = "W_11_05"
    ''Const cst_W_11_06_各授課師資學經歷證書影本 As String = "W_11_06"
    'Const cst_W_11_07_變更後之課程表 As String = "W_11_07"
    ''Const cst_W_11_08_公文 As String = "W_11_08"

    'Const cst_W_14_01_訓練計畫變更表 As String = "W_14_01"
    'Const cst_W_14_02_訓練計畫場地資料表 As String = "W_14_02"
    'Const cst_W_14_03_教學環境資料表 As String = "W_14_03"
    ''Const cst_W_14_04_消防安全設備檢修申報受理單 As String = "W_14_04"
    ''Const cst_W_14_05_建築物防火避難設施與設備安全檢查申報結果通知書 As String = "W_14_05"
    ''Const cst_W_14_06_檢附校方或機關同意租借證明文件 As String = "W_14_06"
    'Const cst_W_14_07_變更後之課程表 As String = "W_14_07"
    ''Const cst_W_14_08_公文 As String = "W_14_08"

    'Const cst_W_15_01_訓練計畫變更表 As String = "W_15_01"
    'Const cst_W_15_02_變更後之課程表 As String = "W_15_02"
    ''Const cst_W_15_03_公文 As String = "W_15_03"

    'Const cst_W_18_01_訓練計畫變更表 As String = "W_18_01"
    'Const cst_W_18_02_變更後之課程表 As String = "W_18_02"
    ''Const cst_W_18_03_公文 As String = "W_18_03"

    'Const cst_W_20_01_訓練計畫變更表 As String = "W_20_01"
    'Const cst_W_20_02_變更後師資助教基本資料表 As String = "W_20_02"
    'Const cst_W_20_03_變更後之課程表 As String = "W_20_03"
    ''Const cst_W_20_04_公文 As String = "W_20_04"

    Dim gsCmd As SqlCommand
    'Dim iPYNum14 As Integer = 1 'TIMS.sUtl_GetPYNum14(Me)
    Dim prtFilename As String = "" '列印表件名稱

    Dim rPlanID As String '計畫PK /PLANID
    Dim rComIDNO As String '計畫PK /CID
    Dim rSeqNo As String '計畫PK /NO
    Dim rSCDate As String '變更PK /CDate
    Dim iSubSeqNO As Integer = 0 '變更PK(INT) '此變數有重複宣告可能
    Dim rAltDataID As String = "" '變更 應為數字 chgState.Value =val(rAltDataID)
    Dim rORGKINDGW As String = ""

    Dim flag_TPlanID70_1 As Boolean = False
    Dim s_SPEC_PCSs1 As String = ""
    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '每次重載執行
        Call Utl_EveryCreate1()

        '下拉選項，首頁資料顯示
        If Not IsPostBack Then
            cCreate1()
        End If

    End Sub

    ''' <summary>每次重載執行</summary>
    Private Sub Utl_EveryCreate1()
        's_SPEC_PCSs1 = TIMS.Utl_GetConfigSet("spec_PCSs1") '某些班級可使用特殊規則。
        ROC_Years.Value = (sm.UserInfo.Years - 1911)

        hid_TPlanID28AppPlan.Value = TIMS.Cst_TPlanID28AppPlan
        '70:區域產業據點職業訓練計畫(在職)
        flag_TPlanID70_1 = (TIMS.Cst_TPlanID70.IndexOf(sm.UserInfo.TPlanID) > -1)

        If (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1) Then gsCmd = SD_14_010.Get_SEL_REVISE_SQLCMD1(objconn)

        'TIMS.SetMyValue(sCmdArg, "CDate", TIMS.cdate3(drv("CDate"))) '申請變更日
        'TIMS.SetMyValue(sCmdArg, "subno", Convert.ToString(drv("SubSeqNo")))
        'TIMS.SetMyValue(sCmdArg, "AltDataID", Convert.ToString(drv("AltDataID")))
        'TIMS.SetMyValue(sCmdArg, "PARTREDUC1", s_PARTREDUC1)
        'TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
        'TIMS.SetMyValue(sCmdArg, "cid", Convert.ToString(drv("ComIDNO")))
        'TIMS.SetMyValue(sCmdArg, "no", Convert.ToString(drv("SeqNO")))
        'TIMS.SetMyValue(sCmdArg, "check", v_SearchMode)
        Try
            rPlanID = "" & TIMS.ClearSQM(Request("PlanID")) 'Request("PlanID")
            rComIDNO = "" & TIMS.ClearSQM(Request("cid")) 'Request("cid")
            rSeqNo = "" & TIMS.ClearSQM(Request("no")) 'Request("no")
            'sPCS1 = rPlanID & "x" & rComIDNO & "x" & rSeqNo
            rSCDate = "" & TIMS.ClearSQM(Request("CDate")) '申請變更日
            If rSCDate <> "" AndAlso Not TIMS.IsDate1(rSCDate) Then rSCDate = ""
            Hid_rCDATE.Value = TIMS.Cdate3(rSCDate)

            iSubSeqNO = If(TIMS.ClearSQM(Request("subno")) <> "", Val(TIMS.ClearSQM(Request("subno"))), 0)

            rAltDataID = TIMS.ClearSQM(Request("AltDataID"))
            chgState.Value = TIMS.ClearSQM(Request("AltDataID")) 'rAltDataID

            hidReqPlanID.Value = rPlanID ' TIMS.ClearSQM(Request("PlanID")) 'Request("PlanID") planid
            hidReqcid.Value = rComIDNO ' TIMS.ClearSQM(Request("cid")) 'Request("cid") comidno
            hidReqno.Value = rSeqNo ' TIMS.ClearSQM(Request("no")) 'Request("no") seqno
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Dim vMsg As String = Common.GetJsString(ex.Message) '.ToString)
            Dim strScript1 As String = String.Concat("<script>", "alert('發生錯誤,請重新查詢選取!\n", vMsg, "');", vbCrLf)
            strScript1 &= String.Concat("location.href='TC_05_001.aspx?ID=", TIMS.Get_MRqID(Me), "';", vbCrLf, "</script>")
            Page.RegisterStartupScript("", strScript1)
            Exit Sub
        End Try
        If rORGKINDGW = "" AndAlso Hid_ORGKINDGW.Value <> "" Then rORGKINDGW = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
    End Sub

    Private Sub cCreate1()
        Try
            Call SHOW_PLANINFO()
            Call SHOW_PLANREVISE() '建立班級變更資料
            Call SHOW_Detail_REVISESUB() '重新查詢
        Catch ex As Exception
            Dim strErrmsg5 As String = ""
            strErrmsg5 &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNo) & vbCrLf
            strErrmsg5 &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg5 &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg5 &= "ex.ToString : " & ex.ToString & vbCrLf
            Call TIMS.WriteTraceLog(strErrmsg5)

            Dim vMsg As String = Common.GetJsString(ex.Message) '.ToString)
            Dim strScript1 As String = String.Concat("<script>", "alert('發生錯誤,請重新查詢選取!\n", vMsg, "');", vbCrLf)
            strScript1 &= String.Concat("location.href='TC_05_001.aspx?ID=", TIMS.Get_MRqID(Me), "';", vbCrLf, "</script>")
            Page.RegisterStartupScript("", strScript1)
            Exit Sub
        End Try

        Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, "")
        If dtFL IsNot Nothing AndAlso dtFL.Rows.Count > 0 Then
            Dim dr1 As DataRow = dtFL.Rows(dtFL.Rows.Count - 1) '取最後一排
            Dim vRVSID As String = Convert.ToString(dr1("RVSID"))
            Call SHOW_KEY_REVISESUB_RVSID(vRVSID, rORGKINDGW, rAltDataID)
        End If
    End Sub

    ''' <summary>建立基本資料 PLAN_PLANINFO</summary>
    Sub SHOW_PLANINFO()
        Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, objconn)
        'PLAN_PLANINFO / TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
        If drPP Is Nothing Then drPP = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
        If drPP Is Nothing Then Return

        rORGKINDGW = Convert.ToString(drPP("ORGKINDGW"))
        Hid_ORGKINDGW.Value = Convert.ToString(drPP("ORGKINDGW"))
        '年度
        YearList.Text = String.Concat(drPP("PlanYear"), "年度")
        Hid_PlanYear.Value = TIMS.ClearSQM(drPP("PlanYear"))
        '申請階段 '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
        Dim s_APPSTAGE2_NM2 As String = If(Convert.ToString(drPP("APPSTAGE")) <> "", TIMS.GET_APPSTAGE2_NM2(Convert.ToString(drPP("APPSTAGE"))), "")
        labAPPSTAGE.Text = If(s_APPSTAGE2_NM2 <> "", String.Concat("(", s_APPSTAGE2_NM2, ")"), "")

        '訓練業別／訓練職類
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then LabTMID.Text = "訓練業別"

        '訓練業別代碼文字
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '企訓專用 '產投 
            JobText.Text = String.Concat("[", drPP("JobID"), "]", drPP("JobName"))
        Else
            TrainText.Text = String.Concat("[", drPP("TrainID"), "]", drPP("TrainName")) 'TIMS
        End If
        hid_TMID.Value = Convert.ToString(drPP("TMID"))

        '通俗職類Labcjob
        CjobName.Text = TIMS.Get_CJOBNAME(dtSCJOB, Convert.ToString(drPP("CJOB_UNKEY")))
        '訓練機構
        OrgName.Text = drPP("OrgName").ToString
        'lab_REVISEACCT_Name.Text = TIMS.Get_ACCNAME(Convert.ToString(dr("REVISEACCT")), gobjconn)
        '業務代碼
        RIDValue.Value = Convert.ToString(drPP("RID"))
        '班別名稱
        ClassName.Text = TIMS.GET_CLASSNAME(Convert.ToString(drPP("ClassName")), Convert.ToString(drPP("CyclType")))
        '企訓專用 '產投 
        PointYN.Visible = (TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1)
        '學分班／是／否
        PointYN.Text = If(Convert.ToString(drPP("PointYN")) = "Y", cst_PointYN_學分班, cst_PointYN_非學分班)
        '訓練期間
        TRange.Text = String.Concat(Common.FormatDate(drPP("STDate")), "~", Common.FormatDate(drPP("FDDate")))
        '是否轉班
        ClassFlag.Text = If(Convert.ToString(drPP("OCID")) = "", "否", "是")
        'TransFlag='N'
        ClassFlag.Text &= If(drPP("TransFlag").ToString = "N", "<font color ='red'>(未轉班)</font>", "")
    End Sub

    ''' <summary>建立變更基本資料  PLAN_REVISE</summary>
    Sub SHOW_PLANREVISE()
        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", rSCDate) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", iSubSeqNO) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then Return ' Exit Sub

        Hid_REVISESTATUS.Value = Convert.ToString(drPR("ReviseStatus"))
        Hid_ONLINESENDSTATUS.Value = Convert.ToString(drPR("ONLINESENDSTATUS"))
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        fg_REVISESUBFL_VIEW1 = (Hid_REVISESTATUS.Value <> "" OrElse Hid_ONLINESENDSTATUS.Value <> "")

        lab_REVISEACCT_Name.Text = TIMS.Get_ACCNAME(Convert.ToString(drPR("REVISEACCT")), objconn)
        labTitle.Text = cst_labTitle_申請狀態
        '查詢模式
        SearchMode.Text = cst_SearchMode_線上送件 '"變更結果"
        Dim v_ReviseStatus As String = Convert.ToString(drPR("ReviseStatus"))
        Dim s_REVISEDATE As String = ""
        Dim s_CheckMode As String = cst_CheckMode_審核中  '"審核中"
        Select Case v_ReviseStatus
            Case "Y"
                s_CheckMode = cst_CheckMode_審核通過 '"審核通過"
                If Convert.ToString(drPR("REVISEDATE")) <> "" Then s_REVISEDATE = "(" & TIMS.Cdate3(drPR("REVISEDATE")) & ")"
            Case "N"
                s_CheckMode = cst_CheckMode_審核不通過 '"審核不通過"
            Case ""
                s_CheckMode = cst_CheckMode_審核中 '審核中"
            Case Else
                s_CheckMode = TIMS.ClearSQM(v_ReviseStatus)
        End Select
        CheckMode.Text = String.Format("{0}{1}", s_CheckMode, s_REVISEDATE)
        '變更項目
        chgState.Value = Convert.ToString(drPR("AltDataID")) 'rAltDataID
        Hid_ALTDATAID.Value = Convert.ToString(drPR("AltDataID")) 'rAltDataID
        Hid_SubSeqNO.Value = Convert.ToString(drPR("SubSeqNo"))
        '變更項目
        LabChgItem_N.Text = String.Concat(If(chgState.Value <> "", TIMS.GET_ChgItemName(sm, Val(chgState.Value)), ""), "(", iSubSeqNO, ")")
        '申請變更日
        labApplyDate_AD.Text = TIMS.Cdate3(drPR("CDATE"))

        Hid_PCS_PR.Value = String.Format("{0}x{1}x{2}x{3}x{4}x{5}x{6}", rPlanID, rComIDNO, rSeqNo, rSCDate, Hid_SubSeqNO.Value, chgState.Value, Hid_ORGKINDGW.Value)
    End Sub

    ''' <summary>重新查詢</summary>
    Sub SHOW_Detail_REVISESUB()
        'KEY_REVISESUB
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim vALTDATAID As String = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        If vORGKINDGW = "" OrElse vALTDATAID = "" Then Return

        Dim htPMS_kR As New Hashtable
        htPMS_kR.Add("ORGKINDGW", vORGKINDGW)
        htPMS_kR.Add("ALTDATAID", vALTDATAID)
        ddlSwitchTo = TIMS.GET_ddlREVISESUB(objconn, ddlSwitchTo, htPMS_kR)

        '最後RVID
        Hid_LastRVID.Value = GET_REVISESUB_LastRVID(objconn, htPMS_kR)
        '開始序號RVSID 
        Hid_FirstRVSID.Value = GET_REVISESUB_FirstRVSID(objconn, htPMS_kR)

        labProgress.Text = "0%"

        If Hid_RVSID.Value <> "" Then
            Call SHOW_KEY_REVISESUB_RVSID(Hid_RVSID.Value, vORGKINDGW, vALTDATAID)
        ElseIf Hid_FirstRVSID.Value <> "" Then
            Call SHOW_KEY_REVISESUB_RVSID(Hid_FirstRVSID.Value, vORGKINDGW, vALTDATAID)
        End If

        'Dim rPMS3 As New Hashtable
        'TIMS.SetMyValue2(rPMS3, "PCS_PR", Hid_PCS_PR.Value)
        'TIMS.SetMyValue2(rPMS3, "PLANID", hidReqPlanID.Value)
        'TIMS.SetMyValue2(rPMS3, "COMIDNO", hidReqcid.Value)
        'TIMS.SetMyValue2(rPMS3, "SEQNO", hidReqno.Value)
        'TIMS.SetMyValue2(rPMS3, "CDATE", Hid_rCDATE.Value)
        'TIMS.SetMyValue2(rPMS3, "SubSeqNO", Hid_SubSeqNO.Value)
        'TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        'TIMS.SetMyValue2(rPMS3, "ALTDATAID", Hid_ALTDATAID.Value)
        'TIMS.SetMyValue2(rPMS3, "RVSID", Hid_RVSID.Value)
        'TIMS.SetMyValue2(rPMS3, "RVID", Hid_RVID.Value)
        'TIMS.SetMyValue2(rPMS3, "BVFID", Hid_BVFID.Value)

        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then Return ' Exit Sub

        '顯示上傳檔案／細項
        'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVID,ALTDATAID,ORGKINDGW
        Call SHOW_REVISESUBFL_DG2()

        '===當該項班級變更申請【審核狀態】為「審核通過」、「審核不通過」時，僅顯示查看按鈕，無編輯、送出按鈕
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        ' ,z.ONLINESENDSTATUS,z.ONLINESENDACCT,z.ONLINESENDDATE
        fg_REVISESUBFL_VIEW1 = (Hid_REVISESTATUS.Value <> "" OrElse Hid_ONLINESENDSTATUS.Value <> "")
        If fg_REVISESUBFL_VIEW1 Then
            'Const cst_tpmsg_enb1 As String = "(檢視功能)不能儲存"
            '儲存(暫存)
            BTN_SAVETMP1.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
            TIMS.Tooltip(BTN_SAVETMP1, cst_tpmsg_enb1, True)
            '儲存後進下一步
            BTN_SAVENEXT1.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
            TIMS.Tooltip(BTN_SAVENEXT1, cst_tpmsg_enb1, True)
        End If
    End Sub

    ''' <summary>下拉資訊調整</summary>
    ''' <param name="vRVSID"></param>
    ''' <param name="vORGKINDGW"></param>
    ''' <param name="vALTDATAID"></param>
    Private Sub SHOW_KEY_REVISESUB_RVSID(vRVSID As String, vORGKINDGW As String, vALTDATAID As String)
        Dim v_ddlSwitchTo As String = TIMS.GetListValue(ddlSwitchTo)
        If (v_ddlSwitchTo <> vRVSID) Then
            Hid_RVSID.Value = vRVSID
            Common.SetListItem(ddlSwitchTo, vRVSID)
        End If

        'tb_PreFileUp.Visible = False
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, vRVSID, vORGKINDGW, vALTDATAID)
        If drKR Is Nothing Then Return

        'RVID,'.',RVNAME
        '(取得)RVID代號／非流水號
        Dim vRVID As String = Convert.ToString(drKR("RVID"))
        '取得文字說明
        Dim vRVDESC1 As String = Convert.ToString(drKR("RVDESC1"))
        Dim vNOTRVDESC1 As String = Convert.ToString(drKR("NOTRVDESC1")) '(Y:不使用KBDESC)
        Dim vNOTFLDESC1 As String = Convert.ToString(drKR("NOTFLDESC1")) '(Y:不使用FLDESC1)
        '必填資訊／免附文件(必填就不顯示)
        Dim vMUSTFILL As String = Convert.ToString(drKR("MUSTFILL"))
        'USELATESTVER : 以最近一次版本送件
        Dim vUSELATESTVER As String = Convert.ToString(drKR("USELATESTVER"))
        'DOWNLOADRPT '可下載報表
        Dim vDOWNLOADRPT As String = Convert.ToString(drKR("DOWNLOADRPT"))
        '(報表名稱)
        Dim vRPTNAME As String = Convert.ToString(drKR("RPTNAME"))
        '以目前版本批次送出:SENTBATVER
        Dim vSENTBATVER As String = Convert.ToString(drKR("SENTBATVER"))
        '以目前版本送出: SENDCURRVER
        Dim vSENDCURRVER As String = Convert.ToString(drKR("SENDCURRVER"))
        '檔案上傳:UPLOADFL1
        Dim vUPLOADFL1 As String = Convert.ToString(drKR("UPLOADFL1"))
        '備註說明:USEMEMO1
        Dim vUSEMEMO1 As String = Convert.ToString(drKR("USEMEMO1"))

        '檔案上傳大小調整 
        Dim fg_USE_FileDescMsg_7_15M As Boolean = False
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本
                fg_USE_FileDescMsg_7_15M = True
        End Select

        '檔案上傳大小調整  '檔案格式說明
        labFILEDESC1.Text = If(fg_USE_FileDescMsg_7_15M, cst_FileDescMsg_7_15M, cst_FileDescMsg_7_10M)
        Dim v_PostedFile_SIZE As String = If(fg_USE_FileDescMsg_7_15M, cst_PostedFile_MAX_SIZE_15M, cst_PostedFile_MAX_SIZE_10M)
        ' return checkFile1(sizeLimit);
        But1.Attributes.Remove("onclick")
        Dim str_rtn_checkFile1 As String = String.Concat("return checkFile1(", v_PostedFile_SIZE, ");")
        If vUPLOADFL1 = "Y" AndAlso str_rtn_checkFile1 <> "" Then
            But1.Attributes.Add("onclick", str_rtn_checkFile1)
        Else
            But1.Enabled = False '(項目參數)不能檔案上傳 (查無上傳參數)
            TIMS.Tooltip(But1, cst_tpmsg_enb2b, True)
        End If

        '取得文字說明
        LiteralSwitchTo.Text = If(vRVSID <> "", TIMS.HtmlDecode1(vRVDESC1), "(無)")
        tr_LiteralSwitchTo.Visible = If(vNOTRVDESC1 = "Y", False, True) '(Y:不使用DESC)
        '檔案格式說明
        tr_FILEDESC1.Visible = If(vNOTFLDESC1 = "Y", False, True) '(Y:不使用FLDESC1)
        '(使用)'(報表名稱)'DOWNLOADRPT '可下載報表
        BTN_DOWNLOADRPT1.Text = If(vRPTNAME <> "", vRPTNAME, "下載報表")
        tr_DOWNLOADRPT1.Visible = If(vDOWNLOADRPT = "Y", True, False)
        '以目前版本批次送出:SENTBATVER
        tr_SENTBATVER.Visible = If(vSENTBATVER = "Y", True, False)
        '以目前版本送出: SENDCURRVER
        tr_SENDCURRVER.Visible = If(vSENDCURRVER = "Y", True, False)
        '檔案上傳:UPLOADFL1
        tr_UPLOADFL1.Visible = If(vUPLOADFL1 = "Y", True, False)

        '取得KBID代號／非流水號
        Hid_RVID.Value = vRVID 'GET_KBID(vKBSID, vORGKINDGW) GET_KBDESC1(vKBSID, vORGKINDGW)
        LabSwitchTo.Text = If(vRVSID <> "", TIMS.GetListText(ddlSwitchTo), "")
        'USELATESTVER : 以最近一次版本送件
        tr_USELATESTVER.Visible = If(vUSELATESTVER = "Y", True, False)
        'MUSTFILL 必填資訊／WAIVED:免附文件(必填就不顯示)
        tr_WAIVED.Visible = If(vMUSTFILL = "Y", False, True)
        '備註說明:USEMEMO1
        tr_USEMEMO1.Visible = If(vUSEMEMO1 = "Y", True, False)
        '預設值-免附文件
        CHKB_WAIVED.Checked = False '(預設值不填寫)
        '預設值-(上傳檔案)
        Hid_BVFID.Value = ""
        '預設值-備註說明
        txtMEMO1.Text = ""

        '師資／助教基本資料表
        tr_DataGrid10.Visible = If(Convert.ToString(drKR("DataGrid10")) = "Y", True, False)
        If tr_DataGrid10.Visible Then Call SHOW_DATAGRID_10(vGW_ALT_RVID)

        '===當該項班級變更申請【審核狀態】為「審核通過」、「審核不通過」時，僅顯示查看按鈕，無編輯、送出按鈕
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        fg_REVISESUBFL_VIEW1 = (Hid_REVISESTATUS.Value <> "" OrElse Hid_ONLINESENDSTATUS.Value <> "")
        '(Enabled) begin
        File1.Disabled = If(fg_REVISESUBFL_VIEW1, True, False)
        But1.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
        TIMS.Tooltip(File1, If(But1.Enabled, "", cst_tpmsg_enb2), True)
        TIMS.Tooltip(But1, If(But1.Enabled, "", cst_tpmsg_enb2), True)

        BTN_SENTBATVER.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
        TIMS.Tooltip(BTN_SENTBATVER, If(BTN_SENTBATVER.Enabled, "", cst_tpmsg_enb3), True)
        BTN_SENDCURRVER.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
        TIMS.Tooltip(BTN_SENDCURRVER, If(BTN_SENDCURRVER.Enabled, "", cst_tpmsg_enb3), True)
        bt_latestSend1.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
        TIMS.Tooltip(bt_latestSend1, If(bt_latestSend1.Enabled, "", cst_tpmsg_enb3), True)
        CHKB_WAIVED.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
        TIMS.Tooltip(CHKB_WAIVED, If(CHKB_WAIVED.Enabled, "", cst_tpmsg_enb3), True)
        '備註說明
        txtMEMO1.Enabled = If(fg_REVISESUBFL_VIEW1, False, True)
        TIMS.Tooltip(txtMEMO1, If(txtMEMO1.Enabled, "", cst_tpmsg_enb3), True)
        ''(Enabled) close

        Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, vRVSID)
        Dim drFL As DataRow = If(dtFL.Rows.Count > 0, dtFL.Rows(0), Nothing)
        If drFL Is Nothing Then Return '資料不存在 以下省略

        Hid_BVFID.Value = Convert.ToString(drFL("BVFID"))
        '免附文件
        CHKB_WAIVED.Checked = If(Convert.ToString(drFL("WAIVED")) = "Y", True, False)

        txtMEMO1.Text = TIMS.ClearSQM(drFL("MEMO1"))
    End Sub

    ''' <summary>'師資／助教基本資料表 PLAN_REVISESUBFL_TT</summary>
    Private Sub SHOW_DATAGRID_10(vGW_ALT_RVID As String)
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        fg_REVISESUBFL_VIEW1 = (Hid_REVISESTATUS.Value <> "" OrElse Hid_ONLINESENDSTATUS.Value <> "")
        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then Return ' Exit Sub

        'Dim vNEWDATA11_1 As String = Convert.ToString(drPR("NEWDATA11_1"))
        'vNEWDATA11_1 = TIMS.CombiSQLIN(vNEWDATA11_1)
        'If vNEWDATA11_1 = "" Then Return
        Dim v_TECHIDS1 As String = ""
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表
                v_TECHIDS1 = Convert.ToString(drPR("NEWDATA11_1"))
            Case TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表
                v_TECHIDS1 = Convert.ToString(drPR("OLDDATA11_1"))
            Case TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本
                v_TECHIDS1 = Convert.ToString(drPR("NEWDATA11_1"))
            Case TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表
                v_TECHIDS1 = Convert.ToString(drPR("NEWDATA20_1"))
            Case Else
                v_TECHIDS1 = Convert.ToString(drPR("NEWDATA11_1"))
        End Select
        v_TECHIDS1 = TIMS.CombiSQLIN(v_TECHIDS1)
        If v_TECHIDS1 = "" Then Return

        'DataGrid10_ItemDataBound
        Hid_RVSID.Value = TIMS.ClearSQM(Hid_RVSID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim rParms2 As New Hashtable
        rParms2.Add("PLANID", Val(rPlanID))
        rParms2.Add("COMIDNO", rComIDNO)
        rParms2.Add("SEQNO", Val(rSeqNo))
        rParms2.Add("CDATE", TIMS.Cdate3(rSCDate))
        rParms2.Add("SUBSEQNO", iSubSeqNO)
        rParms2.Add("ALTDATAID", Val(rAltDataID))
        rParms2.Add("ORGKINDGW", rORGKINDGW)
        rParms2.Add("RVSID", Hid_RVSID.Value)
        rParms2.Add("RID", RIDValue.Value)

        Dim sSql2 As String = ""
        sSql2 &= " WITH WT1 AS (SELECT a.TECHID,a.BVFTID,a.BVFID,a.WAIVED" & vbCrLf
        sSql2 &= " ,a.SRCFILENAME1,a.FILENAME1,a.FILEPATH1,k.RVSID,k.RVID" & vbCrLf
        sSql2 &= " ,b.PLANID,b.COMIDNO,b.SEQNO,b.CDATE,b.SUBSEQNO,b.ALTDATAID,b.ORGKINDGW" & vbCrLf
        sSql2 &= " FROM dbo.PLAN_REVISESUBFL b" & vbCrLf
        sSql2 &= " JOIN dbo.KEY_REVISESUB k ON k.RVSID=b.RVSID" & vbCrLf
        sSql2 &= " JOIN dbo.PLAN_REVISESUBFL_TT a ON a.BVFID=b.BVFID" & vbCrLf
        sSql2 &= " WHERE b.PLANID=@PLANID AND b.COMIDNO=@COMIDNO AND b.SEQNO=@SEQNO" & vbCrLf
        sSql2 &= " AND b.CDATE=@CDATE AND b.SUBSEQNO=@SUBSEQNO" & vbCrLf
        sSql2 &= " AND b.ALTDATAID=@ALTDATAID AND b.ORGKINDGW=@ORGKINDGW AND b.RVSID=@RVSID)" & vbCrLf

        sSql2 &= " SELECT a.TechID,a.RID,a.TEACHCNAME,a.TEACHENAME,a.TEACHERID" & vbCrLf
        sSql2 &= " ,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql2 &= " ,a.KINDENGAGE,case a.KINDENGAGE when '1' then '內聘(專任)' else '外聘(兼任)' end KINDENGAGE_N" & vbCrLf
        sSql2 &= " ,a.WORKSTATUS,case a.WORKSTATUS when '1' then '是' else '否' end WORKSTATUS_N" & vbCrLf
        sSql2 &= " ,(SELECT x.KINDNAME FROM dbo.ID_KINDOFTEACHER x WHERE x.KINDID=a.KINDID) KINDNAME" & vbCrLf
        sSql2 &= " ,t1.BVFID,t1.BVFTID,t1.WAIVED" & vbCrLf
        sSql2 &= " ,t1.SRCFILENAME1,t1.FILENAME1,t1.FILEPATH1,t1.FILENAME1 OKFLAG" & vbCrLf
        sSql2 &= " ,t1.RVID,t1.RVSID" & vbCrLf
        sSql2 &= " FROM dbo.TEACH_TEACHERINFO a" & vbCrLf
        sSql2 &= " LEFT JOIN WT1 t1 on t1.TECHID=a.TECHID" & vbCrLf
        sSql2 &= " WHERE a.RID=@RID" & vbCrLf
        sSql2 &= String.Concat(" AND a.TechID IN (", v_TECHIDS1, ")")
        sSql2 &= " ORDER BY a.TEACHERID" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, rParms2)

        labmsg2.Text = ""
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG10_ROWS = dt2.Rows.Count

        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
        Dim download_Path As String = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO)
        Call TIMS.Check_dtREVISESUBFL(Me, dt2, download_Path)

        With DataGrid10
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    ''' <summary>KEY_REVISESUB 取得 FirstRVSID </summary>
    ''' <param name="oConn"></param>
    ''' <param name="htPMS_kR"></param>
    ''' <returns></returns>
    Public Shared Function GET_REVISESUB_FirstRVSID(oConn As SqlConnection, htPMS_kR As Hashtable) As String
        Dim vORGKINDGW As String = TIMS.GetMyValue2(htPMS_kR, "ORGKINDGW")
        Dim vALTDATAID As String = TIMS.GetMyValue2(htPMS_kR, "ALTDATAID")
        If vORGKINDGW = "" OrElse vALTDATAID = "" Then Return ""

        Dim xsPMS2 As New Hashtable
        xsPMS2.Add("ORGKINDGW", vORGKINDGW)
        xsPMS2.Add("ALTDATAID", vALTDATAID)
        Dim xSql2 As String = "SELECT TOP 1 RVSID FROM KEY_REVISESUB WHERE ORGKINDGW=@ORGKINDGW AND ALTDATAID=@ALTDATAID ORDER BY RSORT"
        Dim xRVSID As String = DbAccess.ExecuteScalar(xSql2, oConn, xsPMS2)
        Return xRVSID  'Hid_FirstRVSID.Value = xRVSID 'rFirstRVSID = xRVSID 
    End Function

    ''' <summary>KEY_REVISESUB 取得 LastRVID </summary>
    ''' <param name="oConn"></param>
    ''' <param name="htPMS_kR"></param>
    ''' <returns></returns>
    Public Shared Function GET_REVISESUB_LastRVID(oConn As SqlConnection, htPMS_kR As Hashtable) As String
        Dim vORGKINDGW As String = TIMS.GetMyValue2(htPMS_kR, "ORGKINDGW")
        Dim vALTDATAID As String = TIMS.GetMyValue2(htPMS_kR, "ALTDATAID")
        If vORGKINDGW = "" OrElse vALTDATAID = "" Then Return ""

        Dim xsPMS As New Hashtable
        xsPMS.Add("ORGKINDGW", vORGKINDGW)
        xsPMS.Add("ALTDATAID", vALTDATAID)
        Dim xSql As String = "SELECT TOP 1 RVID xKBID FROM KEY_REVISESUB WHERE ORGKINDGW=@ORGKINDGW AND ALTDATAID=@ALTDATAID ORDER BY RSORT DESC"
        Dim xRVID As String = DbAccess.ExecuteScalar(xSql, oConn, xsPMS)
        Return xRVID
    End Function

    ''' <summary>重新查詢 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH2_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH2.Click
        Call SHOW_Detail_REVISESUB()
    End Sub

    Protected Sub ddlSwitchTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSwitchTo.SelectedIndexChanged
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim vALTDATAID As String = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        If vORGKINDGW = "" OrElse vALTDATAID = "" Then Return

        Hid_RVSID.Value = TIMS.GetListValue(ddlSwitchTo)

        If Hid_RVSID.Value <> "" Then
            Call SHOW_KEY_REVISESUB_RVSID(Hid_RVSID.Value, vORGKINDGW, vALTDATAID)
        ElseIf Hid_FirstRVSID.Value <> "" Then
            Call SHOW_KEY_REVISESUB_RVSID(Hid_FirstRVSID.Value, vORGKINDGW, vALTDATAID)
        End If

        '顯示上傳檔案／細項
        'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVID,ALTDATAID,ORGKINDGW
        Call SHOW_REVISESUBFL_DG2()
    End Sub

    ''' <summary>確定檔案上傳</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But1_Click(sender As Object, e As EventArgs) Handles But1.Click
        'Dim vUploadPath As String = Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, Hid_RVSID.Value, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        If drKR Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!")
            Return
        End If

        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then Return ' Exit Sub

        'SyncLock FileUpload_lock End SyncLock
        'PLAN_REVISE
        Dim sAltDataID As String = Convert.ToString(drPR("AltDataID"))
        '取得 KEY_REVISESUB 代號／非流水號
        Dim vORGKINDGW As String = Convert.ToString(drKR("ORGKINDGW"))
        Dim vALTDATAID As String = Convert.ToString(drKR("ALTDATAID"))
        Dim vRVID As String = Convert.ToString(drKR("RVID"))
        Dim vRVNAME As String = Convert.ToString(drKR("RVNAME"))
        Dim vRVNAME2 As String = String.Concat(vORGKINDGW, vRVID, ".", vRVNAME)
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表
                '變更後師資助教基本資料表-師資檔案上傳
                Call FILE_UPLOAD_10(drPR)
                Call SHOW_DATAGRID_10(vGW_ALT_RVID)
            Case TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表
                '變更前師資助教基本資料表-師資檔案上傳
                Call FILE_UPLOAD_10(drPR)
                Call SHOW_DATAGRID_10(vGW_ALT_RVID)
            Case TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本
                '授課師資學經歷證書影本-師資檔案上傳
                Call FILE_UPLOAD_10(drPR)
                Call SHOW_DATAGRID_10(vGW_ALT_RVID)
            Case TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表
                '變更後師資助教基本資料表-師資檔案上傳
                Call FILE_UPLOAD_10(drPR)
                Call SHOW_DATAGRID_10(vGW_ALT_RVID)
            Case Else
                '(檢查儲存值)
                Dim flag_OK_PRFL As Boolean = CHK_PLAN_REVISESUBFL(Hid_RVSID.Value)
                'Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步")
                If flag_OK_PRFL Then
                    Common.MessageBox(Me, "已有上傳資料或檔案，重新上傳請先刪除!")
                    Return
                End If

                '檔案上傳／確定檔案上傳
                Call FILE_UPLOAD_1(drPR)

                '(檢查儲存值) Dim flag_OK_PRFL As Boolean 
                Dim flag_OK_PRFL2 As Boolean = CHK_PLAN_REVISESUBFL(Hid_RVSID.Value)
                'Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步")
                If Not flag_OK_PRFL2 Then
                    Common.MessageBox(Me, "查無上傳資料，請重新確認!")
                    Return
                End If

                '下一步
                Call MOVE_NEXT2()
        End Select

        '顯示上傳檔案／細項
        'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVID,ALTDATAID,ORGKINDGW
        Call SHOW_REVISESUBFL_DG2()
    End Sub

    ''' <summary>師資檔案上傳</summary>
    ''' <param name="drPR"></param>
    Private Sub FILE_UPLOAD_10(drPR As DataRow)
        Hid_TECHID.Value = ""
        For Each eItem As DataGridItem In DataGrid10.Items
            Dim chkItem1 As HtmlInputCheckBox = eItem.FindControl("chkItem1")
            Dim HDG10_TechID As HtmlInputHidden = eItem.FindControl("HDG10_TechID")
            'Dim HDG10_RID As HtmlInputHidden = eItem.FindControl("HDG10_RID") AndAlso HDG10_RID.Value <> ""
            If chkItem1.Checked AndAlso HDG10_TechID.Value <> "" Then
                Hid_TECHID.Value = HDG10_TechID.Value
                Exit For
            End If
        Next
        Hid_TECHID.Value = TIMS.ClearSQM(Hid_TECHID.Value)
        Dim vTECHID As String = Hid_TECHID.Value
        If vTECHID = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(未選擇老師)，請重新操作!!")
            Return
        End If
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Return ' Exit Sub
        End If

        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFilePdf(Me, File1, MyPostedFile) Then Return

        'Dim MyFileColl As HttpFileCollection = HttpContext.Current.Request.Files
        'Dim MyPostedFile As HttpPostedFile = MyFileColl.Item(0)
        'If Not MyPostedFile.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase) Then
        '    Common.MessageBox(Me, cst_errMsg_11_PDF)
        '    Return ' Exit Sub
        'ElseIf File1.Value = "" Then
        '    Common.MessageBox(Me, cst_errMsg_8)
        '    Return ' Exit Sub
        'ElseIf File1.PostedFile.ContentLength = 0 Then
        '    Common.MessageBox(Me, cst_errMsg_3)
        '    Return ' Exit Sub
        'End If

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Return ' Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "pdf"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                    Common.MessageBox(Me, cst_errMsg_7_10M)
                    Return ' Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Return ' Exit Sub
        End Select

        '取得代號／非流水號
        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)
        'Dim vALTDATAID As String = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        '上傳檔案 '計畫ID／機構ID
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, TIMS.Cdate3(rSCDate, "yyyyMMdd"), iSubSeqNO)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PR_T(vRID, vRVSID, rAltDataID, vTECHID, "pdf")
        Dim vSRCFILENAME1 As String = MyFileName
        '上傳檔案/存檔：檔名
        Try
            '上傳檔案
            TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Dim rPMS As New Hashtable
        rPMS.Add("RVSID", vRVSID)
        rPMS.Add("TECHID", vTECHID)
        rPMS.Add("FILENAME1", vFILENAME1)
        rPMS.Add("FILEPATH1", vUploadPath)
        rPMS.Add("SRCFILENAME1", vSRCFILENAME1)
        rPMS.Add("WAIVED", vWAIVED)
        rPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
        Call SAVE_PLAN_REVISESUBFL_TT(rPMS)
    End Sub

    Private Sub SAVE_PLAN_REVISESUBFL_TT(rPMS As Hashtable)
        Dim iBVFID As Integer = -1
        Dim vRVSID As String = TIMS.GetMyValue2(rPMS, "RVSID")
        Dim vTECHID As String = TIMS.GetMyValue2(rPMS, "TECHID")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS, "FILEPATH1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vWAIVED As String = TIMS.GetMyValue2(rPMS, "WAIVED")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "BCFID", If(vUploadPath <> "", iBCFID, -1)) '(可再次傳送)
            TIMS.SetMyValue2(rPMS2, "PLANID", rPlanID)
            TIMS.SetMyValue2(rPMS2, "COMIDNO", rComIDNO)
            TIMS.SetMyValue2(rPMS2, "SEQNO", rSeqNo)
            TIMS.SetMyValue2(rPMS2, "CDATE", rSCDate)
            TIMS.SetMyValue2(rPMS2, "SUBSEQNO", iSubSeqNO)
            TIMS.SetMyValue2(rPMS2, "RVSID", vRVSID)
            TIMS.SetMyValue2(rPMS2, "ALTDATAID", rAltDataID)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", rORGKINDGW)
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "FILEPATH1", vFILEPATH1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", cst_10_師資助教學經歷證書資料表_WAIVED_TT)
            iBVFID = SAVE_PLAN_REVISESUBFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.Message)

            Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub 'Throw ex
        End Try
        If (iBVFID = -1) Then Return

        Dim fPMST As New Hashtable
        fPMST.Add("BVFID", iBVFID)
        fPMST.Add("RVSID", vRVSID)
        fPMST.Add("TECHID", vTECHID)
        Dim fSqlT As String = "SELECT 1 FROM PLAN_REVISESUBFL_TT WHERE BVFID=@BVFID AND RVSID=@RVSID AND TECHID=@TECHID"
        Dim drFLT As DataRow = DbAccess.GetOneRow(fSqlT, objconn, fPMST)
        If drFLT IsNot Nothing Then
            TIMS.LOG.Warn(String.Concat("##PLAN_REVISESUBFL_TT EXISTS!!!", ",BVFID:", iBVFID, ",RVSID:", vRVSID, ",TECHID:", vTECHID))
            Common.MessageBox(Me, "已上傳或儲存過該文件，不可再次操作!!!")
            Return
        End If

        Dim iBVFTID As Integer = DbAccess.GetNewId(objconn, "PLAN_REVISESUBFL_TT_BVFTID_SEQ,PLAN_REVISESUBFL_TT,BVFTID")
        Dim iParms As New Hashtable
        iParms.Add("BVFTID", iBVFTID)
        iParms.Add("BVFID", iBVFID)
        iParms.Add("RVSID", vRVSID)
        iParms.Add("TECHID", vTECHID)
        iParms.Add("FILENAME1", vFILENAME1)
        iParms.Add("FILEPATH1", vFILEPATH1)
        iParms.Add("SRCFILENAME1", vSRCFILENAME1)
        'iParms.Add("PATTERN", PATTERN)
        'iParms.Add("MEMO1", MEMO1)
        iParms.Add("WAIVED", vWAIVED)
        iParms.Add("MODIFYACCT", sm.UserInfo.UserID)
        'iParms.Add("MODIFYDATE", MODIFYDATE)
        Dim isSql As String = ""
        isSql &= " INSERT INTO PLAN_REVISESUBFL_TT(BVFTID, BVFID, RVSID, TechID, FILENAME1,FILEPATH1, SRCFILENAME1, WAIVED, MODIFYACCT, MODIFYDATE)" & vbCrLf
        isSql &= " VALUES(@BVFTID,@BVFID,@RVSID,@TECHID,@FILENAME1,@FILEPATH1,@SRCFILENAME1 ,@WAIVED,@MODIFYACCT,GETDATE())" & vbCrLf
        DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
    End Sub

    Private Sub FILE_UPLOAD_1(drPR As DataRow)
        If drPR Is Nothing Then Return

        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Return ' Exit Sub
        End If

        Dim MyFileColl As HttpFileCollection = HttpContext.Current.Request.Files
        Dim MyPostedFile As HttpPostedFile = MyFileColl.Item(0)
        If Not MyPostedFile.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase) Then
            Common.MessageBox(Me, cst_errMsg_11_PDF)
            Return ' Exit Sub
        ElseIf File1.Value = "" Then
            Common.MessageBox(Me, cst_errMsg_8)
            Return ' Exit Sub
        ElseIf File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, cst_errMsg_3)
            Return ' Exit Sub
        End If

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Return ' Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "pdf"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                    Common.MessageBox(Me, cst_errMsg_7_10M)
                    Return ' Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Return ' Exit Sub
        End Select

        '取得代號／非流水號
        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)
        'Dim vALTDATAID As String = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
        'Dim vSubNO As String = TIMS.ClearSQM(Hid_SubSeqNO.Value)

        '上傳檔案 '計畫ID／機構ID
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PR(vRID, vRVSID, rAltDataID, "pdf")
        Dim vSRCFILENAME1 As String = MyFileName
        '上傳檔案/存檔：檔名
        Try
            '上傳檔案
            TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "BCFID", If(vUploadPath <> "", iBCFID, -1)) '(可再次傳送)
            TIMS.SetMyValue2(rPMS2, "PLANID", rPlanID)
            TIMS.SetMyValue2(rPMS2, "COMIDNO", rComIDNO)
            TIMS.SetMyValue2(rPMS2, "SEQNO", rSeqNo)
            TIMS.SetMyValue2(rPMS2, "CDATE", rSCDate)
            TIMS.SetMyValue2(rPMS2, "SUBSEQNO", iSubSeqNO)
            TIMS.SetMyValue2(rPMS2, "RVSID", Hid_RVSID.Value)
            TIMS.SetMyValue2(rPMS2, "ALTDATAID", rAltDataID)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", rORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            'TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED)
            Call SAVE_PLAN_REVISESUBFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.Message)

            Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub 'Throw ex
        End Try
    End Sub

    ''' <summary>上傳檔案 儲存 SAVE PLAN_REVISESUBFL</summary>
    ''' <param name="rPMS"></param>
    Function SAVE_PLAN_REVISESUBFL_UPLOAD(rPMS As Hashtable) As Integer
        Dim iBVFID As Integer = -1
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(rPMS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(rPMS, "SEQNO")
        Dim vCDATE As String = TIMS.GetMyValue2(rPMS, "CDATE")
        Dim vSUBSEQNO As String = TIMS.GetMyValue2(rPMS, "SUBSEQNO")
        Dim vRVSID As String = TIMS.GetMyValue2(rPMS, "RVSID")
        Dim vALTDATAID As String = TIMS.GetMyValue2(rPMS, "ALTDATAID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS, "FILEPATH1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        Dim vWAIVED As String = TIMS.GetMyValue2(rPMS, "WAIVED")

        '(任一錯誤不可儲存)
        If vPLANID = "" OrElse vCOMIDNO = "" OrElse vSEQNO = "" Then Return iBVFID '(異常)
        If vCDATE = "" OrElse vSUBSEQNO = "" OrElse vRVSID = "" Then Return iBVFID '(異常)
        If vALTDATAID = "" OrElse vORGKINDGW = "" Then Return iBVFID '(異常)
        If (vFILENAME1 = "" OrElse vFILEPATH1 = "" OrElse vSRCFILENAME1 = "") AndAlso vWAIVED = "" Then Return iBVFID '(異常)

        'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVSID,ALTDATAID,ORGKINDGW
        Dim sParms As New Hashtable
        sParms.Add("PLANID", vPLANID)
        sParms.Add("COMIDNO", vCOMIDNO)
        sParms.Add("SEQNO", vSEQNO)
        sParms.Add("CDATE", vCDATE)
        sParms.Add("SUBSEQNO", vSUBSEQNO)
        sParms.Add("RVSID", vRVSID)
        sParms.Add("ALTDATAID", vALTDATAID)
        sParms.Add("ORGKINDGW", vORGKINDGW)
        Dim sSql As String = ""
        sSql &= " SELECT BVFID FROM PLAN_REVISESUBFL" & vbCrLf
        sSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        sSql &= " AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO" & vbCrLf
        sSql &= " AND RVSID=@RVSID AND ALTDATAID=@ALTDATAID AND ORGKINDGW=@ORGKINDGW" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)
        If dt1.Rows.Count = 0 Then
            iBVFID = DbAccess.GetNewId(objconn, "PLAN_REVISESUBFL_BVFID_SEQ,PLAN_REVISESUBFL,BVFID")
            '/* IDENTITY(1,1) */
            Dim iParms As New Hashtable
            iParms.Add("BVFID", iBVFID)
            iParms.Add("PLANID", vPLANID)
            iParms.Add("COMIDNO", vCOMIDNO)
            iParms.Add("SEQNO", vSEQNO)
            iParms.Add("CDATE", vCDATE)
            iParms.Add("SUBSEQNO", vSUBSEQNO)
            iParms.Add("RVSID", vRVSID)
            iParms.Add("ALTDATAID", vALTDATAID)
            iParms.Add("ORGKINDGW", vORGKINDGW)
            iParms.Add("FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull))
            iParms.Add("FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull))
            iParms.Add("SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull))
            'iParms.Add("PATTERN", PATTERN)
            'iParms.Add("MEMO1", MEMO1)
            iParms.Add("MODIFYACCT", vMODIFYACCT)
            'iParms.Add("MODIFYDATE", MODIFYDATE)
            iParms.Add("WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull))

            Dim isSql As String = ""
            isSql &= " INSERT INTO PLAN_REVISESUBFL(BVFID,PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVSID,ALTDATAID,ORGKINDGW" & vbCrLf
            isSql &= " ,FILENAME1,FILEPATH1,SRCFILENAME1 ,MODIFYACCT,MODIFYDATE,WAIVED )" & vbCrLf
            isSql &= " VALUES (@BVFID,@PLANID,@COMIDNO,@SEQNO,@CDATE,@SUBSEQNO,@RVSID,@ALTDATAID,@ORGKINDGW" & vbCrLf
            isSql &= " ,@FILENAME1,@FILEPATH1,@SRCFILENAME1 ,@MODIFYACCT,GETDATE(),@WAIVED )" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            'Dim iBVFID As Integer = dt1.Rows(0)("BVFID") 'If vFILENAME1 <> "" AndAlso vSRCFILENAME1 <> "" Then
            iBVFID = dt1.Rows(0)("BVFID")
            Dim uParms2 As New Hashtable
            uParms2.Add("FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull))
            uParms2.Add("FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull))
            uParms2.Add("SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull))
            uParms2.Add("MODIFYACCT", vMODIFYACCT)
            uParms2.Add("WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull))

            uParms2.Add("PLANID", vPLANID)
            uParms2.Add("COMIDNO", vCOMIDNO)
            uParms2.Add("SEQNO", vSEQNO)
            uParms2.Add("CDATE", vCDATE)
            uParms2.Add("SUBSEQNO", vSUBSEQNO)
            uParms2.Add("RVSID", vRVSID)
            uParms2.Add("ALTDATAID", vALTDATAID)
            uParms2.Add("ORGKINDGW", vORGKINDGW)
            uParms2.Add("BVFID", iBVFID)
            Dim usSql2 As String = ""
            usSql2 &= " UPDATE PLAN_REVISESUBFL" & vbCrLf
            usSql2 &= " SET FILENAME1=@FILENAME1,FILEPATH1=@FILEPATH1,SRCFILENAME1=@SRCFILENAME1" & vbCrLf
            usSql2 &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql2 &= " ,WAIVED=@WAIVED" & vbCrLf
            usSql2 &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
            usSql2 &= " AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO" & vbCrLf
            usSql2 &= " AND RVSID=@RVSID AND ALTDATAID=@ALTDATAID AND ORGKINDGW=@ORGKINDGW" & vbCrLf
            usSql2 &= " AND BVFID=@BVFID" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        End If

        '(無檔案且免付文件，執行刪除檔案)
        If vFILENAME1 = "" AndAlso vSRCFILENAME1 = "" AndAlso vWAIVED = "Y" Then
            Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, vRVSID)
            Dim drFL As DataRow = If(dtFL.Rows.Count > 0, dtFL.Rows(0), Nothing)
            If drFL IsNot Nothing Then
                Dim oFILENAME1 As String = ""
                Dim oFILEPATH1 As String = ""
                Dim oUploadPath As String = ""
                Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
                Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
                Try
                    oFILENAME1 = Convert.ToString(drFL("FILENAME1"))
                    oFILEPATH1 = Convert.ToString(drFL("FILEPATH1"))
                    oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, CStr(iSubSeqNO)))
                    If oFILENAME1 <> "" Then TIMS.MyFileDelete(Server.MapPath(oUploadPath & oFILENAME1))
                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex)
                    Common.MessageBox(Me, ex.Message)

                    Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                    strErrmsg &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
                    strErrmsg &= String.Concat("oFILEPATH1: ", oFILEPATH1, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("Server.MapPath(oUploadPath & oFILENAME1): ", Server.MapPath(oUploadPath & oFILENAME1), vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                End Try
            End If
        End If

        Return iBVFID
    End Function

    ''' <summary>查詢SHOW PLAN_REVISESUBFL DG2</summary>
    Private Sub SHOW_REVISESUBFL_DG2()
        '===當該項班級變更申請【審核狀態】為「審核通過」、「審核不通過」時，僅顯示查看按鈕，無編輯、送出按鈕
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        fg_REVISESUBFL_VIEW1 = (Hid_REVISESTATUS.Value <> "" OrElse Hid_ONLINESENDSTATUS.Value <> "")

        Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, "")

        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
        Dim download_Path As String = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO)
        Call TIMS.Check_dtREVISESUBFL(Me, dtFL, download_Path)

        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

        Dim rPMS As New Hashtable
        TIMS.SetMyValue2(rPMS, "ORGKINDGW", rORGKINDGW)
        TIMS.SetMyValue2(rPMS, "ALTDATAID", rAltDataID)
        TIMS.SetMyValue2(rPMS, "PLANID", rPlanID)
        TIMS.SetMyValue2(rPMS, "COMIDNO", rComIDNO)
        TIMS.SetMyValue2(rPMS, "SEQNO", rSeqNo)
        TIMS.SetMyValue2(rPMS, "CDATE", TIMS.Cdate3(rSCDate))
        TIMS.SetMyValue2(rPMS, "SUBSEQNO", iSubSeqNO)
        Dim tmpMSG As String = ""
        '線上申辦進度 計算完成度百分比 (0-100)
        Dim iProgress As Integer = TIMS.GET_iPROGRESS_PR(objconn, tmpMSG, rPMS)
        labProgress.Text = String.Concat(iProgress, "%")

    End Sub

    ''' <summary>GET PLAN_REVISESUBFL dt 或 GET PLAN_REVISESUBFL 依 RVSID</summary>
    ''' <param name="oConn"></param>
    ''' <param name="rRVSID"></param>
    ''' <returns></returns>
    Function GET_PLAN_REVISESUBFL_TB(oConn As SqlConnection, rRVSID As String) As DataTable
        Dim rParms As New Hashtable
        rParms.Add("PLANID", Val(rPlanID))
        rParms.Add("COMIDNO", rComIDNO)
        rParms.Add("SEQNO", Val(rSeqNo))
        rParms.Add("CDATE", TIMS.Cdate3(rSCDate))
        rParms.Add("SUBSEQNO", iSubSeqNO)
        rParms.Add("ALTDATAID", Val(rAltDataID))
        rParms.Add("ORGKINDGW", rORGKINDGW)
        If rRVSID <> "" Then rParms.Add("RVSID", rRVSID)

        Dim rsSql As String = ""
        rsSql &= " SELECT a.BVFID,a.PLANID,a.COMIDNO,a.SEQNO,a.CDATE,a.SUBSEQNO" & vbCrLf
        rsSql &= " ,a.RVSID,a.ALTDATAID,a.ORGKINDGW" & vbCrLf
        rsSql &= " ,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE,a.RTUREASON,a.RTURESACCT,a.RTURESDATE" & vbCrLf
        rsSql &= " ,k.RVID,k.RVNAME,concat(k.RVID,'.',k.RVNAME) RVNAME2" & vbCrLf
        rsSql &= " ,concat(k.ORGKINDGW,k.RVID,'x',k.RVNAME) RVNAME3" & vbCrLf
        rsSql &= " ,concat(k.ORGKINDGW,'_',a.ALTDATAID,'_',k.RVID) GW_ALT_RVID" & vbCrLf
        rsSql &= " ,a.WAIVED,CASE WHEN a.WAIVED='Y' THEN '(免附文件)' ELSE a.SRCFILENAME1 END SRCFILENAME1" & vbCrLf
        rsSql &= " ,a.FILEPATH1,a.FILENAME1,a.FILENAME1 OKFLAG" & vbCrLf
        rsSql &= " FROM PLAN_REVISESUBFL a" & vbCrLf
        rsSql &= " JOIN KEY_REVISESUB k on k.RVSID=a.RVSID and k.ALTDATAID=a.ALTDATAID" & vbCrLf
        rsSql &= " JOIN PLAN_REVISE r on r.PLANID=a.PLANID AND r.COMIDNO=a.COMIDNO AND r.SEQNO=a.SEQNO AND r.CDATE=a.CDATE AND r.SUBSEQNO=a.SUBSEQNO" & vbCrLf
        rsSql &= " WHERE a.PLANID=@PLANID AND a.COMIDNO=@COMIDNO AND a.SEQNO=@SEQNO" & vbCrLf
        rsSql &= " AND a.CDATE=@CDATE AND a.SUBSEQNO=@SUBSEQNO AND k.ORGKINDGW=@ORGKINDGW AND k.ALTDATAID=@ALTDATAID" & vbCrLf
        If rRVSID <> "" Then rsSql &= " AND a.RVSID=@RVSID" & vbCrLf
        If rRVSID = "" Then rsSql &= " ORDER BY k.RSORT" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        Return dt
    End Function

#Region "NO USE"
    'Private Function GET_PLAN_REVISESUBFL(oConn As SqlConnection, rRVSID As String) As DataRow
    '    Dim rParms As New Hashtable
    '    rParms.Add("PLANID", Val(rPlanID))
    '    rParms.Add("COMIDNO", rComIDNO)
    '    rParms.Add("SEQNO", Val(rSeqNo))
    '    rParms.Add("CDATE", TIMS.cdate3(rSCDate))
    '    rParms.Add("SUBSEQNO", iSubSeqNO)
    '    rParms.Add("ALTDATAID", Val(rAltDataID))
    '    rParms.Add("ORGKINDGW", rORGKINDGW)
    '    rParms.Add("RVSID", rRVSID)

    '    Dim rsSql As String = ""
    '    rsSql &= " SELECT a.BVFID,a.PLANID,a.COMIDNO,a.SEQNO,a.CDATE,a.SUBSEQNO,a.RVSID,a.ALTDATAID,a.ORGKINDGW" & vbCrLf
    '    rsSql &= " ,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
    '    rsSql &= " ,a.RTUREASON,a.RTURESACCT,a.RTURESDATE" & vbCrLf
    '    rsSql &= " ,kb.RVID,kb.RVNAME,concat(kb.RVID,'.',kb.RVNAME) RVNAME2" & vbCrLf
    '    rsSql &= " ,a.RTUREASON" & vbCrLf
    '    rsSql &= " ,a.WAIVED,CASE WHEN a.WAIVED='Y' THEN '(免附文件)' ELSE a.SRCFILENAME1 END SRCFILENAME1" & vbCrLf
    '    rsSql &= " ,a.FILENAME1,a.FILENAME1 OKFLAG" & vbCrLf
    '    rsSql &= " FROM PLAN_REVISESUBFL a" & vbCrLf
    '    rsSql &= " JOIN KEY_REVISESUB kb on kb.RVSID=a.RVSID" & vbCrLf
    '    rsSql &= " JOIN PLAN_REVISE r on r.PLANID=a.PLANID AND r.COMIDNO=a.COMIDNO AND r.SEQNO=a.SEQNO AND r.CDATE=a.CDATE AND r.SUBSEQNO=a.SUBSEQNO" & vbCrLf
    '    rsSql &= " WHERE a.PLANID=@PLANID AND a.COMIDNO=@COMIDNO AND a.SEQNO=@SEQNO AND a.CDATE=@CDATE AND a.SUBSEQNO=@SUBSEQNO" & vbCrLf
    '    rsSql &= " AND kb.ORGKINDGW=@ORGKINDGW AND kb.ALTDATAID=@ALTDATAID AND a.RVSID=@RVSID" & vbCrLf
    '    'rsSql &= " ORDER BY kb.RSORT" & vbCrLf

    '    Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
    '    If dt.Rows.Count > 0 Then Return dt.Rows(0)
    '    Return Nothing
    'End Function
#End Region

    ''' <summary>回上一步</summary>
    Private Sub MOVE_PREV()
        If (Hid_RVID.Value = "" OrElse Hid_RVID.Value = "01" OrElse ddlSwitchTo.SelectedIndex - 1 = -1) Then
            Common.MessageBox(Me, "(目前沒有上一步)")
            Return
        End If

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Hid_RVSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex - 1).Value
        If Hid_RVSID.Value <> "" Then
            Call SHOW_KEY_REVISESUB_RVSID(Hid_RVSID.Value, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        ElseIf Hid_FirstRVSID.Value <> "" Then
            Call SHOW_KEY_REVISESUB_RVSID(Hid_FirstRVSID.Value, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        End If
    End Sub

    ''' <summary>儲存後進下一步</summary>
    ''' <param name="s_ACTTYPE"></param>
    Private Sub SAVEDATA2_BTN_ACTION1(s_ACTTYPE As String)
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If tr_WAIVED.Visible AndAlso vWAIVED = "" Then
            Common.MessageBox(Me, String.Concat("請勾選，免附文件，再按儲存!", vbCrLf, "要上傳檔案，請按「確定檔案上傳」!"))
            Return
        End If

        '顯示檔案資料表
        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!")
            Return ' Exit Sub
        End If
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, Hid_RVSID.Value, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        If drKR Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!!")
            Return ' Exit Sub
        End If
        If drPP Is Nothing Then drPP = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
        If drPP Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!!!")
            Return ' Exit Sub
        End If
        If drRR Is Nothing Then drRR = TIMS.Get_RID_DR(Convert.ToString(drPP("RID")), objconn)
        'Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
        If drRR Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!!!!")
            Return ' Exit Sub
        End If

        Dim sMyValue As String = ""
        'PLAN_REVISE
        Dim sAltDataID As String = Convert.ToString(drPR("AltDataID"))
        '取得 KEY_REVISESUB 代號／非流水號
        Dim vORGKINDGW As String = Convert.ToString(drKR("ORGKINDGW"))
        Dim vALTDATAID As String = Convert.ToString(drKR("ALTDATAID"))
        Dim vRVID As String = Convert.ToString(drKR("RVID"))
        Dim vRVNAME As String = Convert.ToString(drKR("RVNAME"))
        Dim vRVNAME2 As String = String.Concat(vORGKINDGW, vRVID, ".", vRVNAME)
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表,
                 TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本,
                 TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表
                'TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表,
                vWAIVED = cst_10_師資助教學經歷證書資料表_WAIVED_TT
        End Select

        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "BCFID", If(vUploadPath <> "", iBCFID, -1)) '(可再次傳送)
            TIMS.SetMyValue2(rPMS2, "PLANID", rPlanID)
            TIMS.SetMyValue2(rPMS2, "COMIDNO", rComIDNO)
            TIMS.SetMyValue2(rPMS2, "SEQNO", rSeqNo)
            TIMS.SetMyValue2(rPMS2, "CDATE", rSCDate)
            TIMS.SetMyValue2(rPMS2, "SUBSEQNO", iSubSeqNO)
            TIMS.SetMyValue2(rPMS2, "RVSID", Hid_RVSID.Value)
            TIMS.SetMyValue2(rPMS2, "ALTDATAID", rAltDataID)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", rORGKINDGW)
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "FILEPATH1", vFILEPATH1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED)
            Call SAVE_PLAN_REVISESUBFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.Message)

            Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Return 'Exit Sub 'Throw ex
        End Try

        Select Case s_ACTTYPE
            Case cst_ACTTYPE_BTN_SAVETMP1
                '儲存(暫存) 
                Dim flag_OK_PRFL As Boolean = CHK_PLAN_REVISESUBFL(Hid_RVSID.Value)
                If Not flag_OK_PRFL Then
                    Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行儲存")
                    Return
                End If
                Call SHOW_KEY_REVISESUB_RVSID(Hid_RVSID.Value, rORGKINDGW, rAltDataID) '項目(重跑1次)
            Case cst_ACTTYPE_BTN_SAVENEXT1
                '儲存後進下一步
                '(檢查儲存值)
                Dim flag_OK_PRFL As Boolean = CHK_PLAN_REVISESUBFL(Hid_RVSID.Value)
                If Not flag_OK_PRFL Then
                    Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步")
                    Return
                End If
                Call MOVE_NEXT() '下一步
        End Select

        '顯示上傳檔案／細項
        'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVID,ALTDATAID,ORGKINDGW
        Call SHOW_REVISESUBFL_DG2()
    End Sub

    ''' <summary>檢核資料存在與否 PLAN_REVISESUBFL 依 vRVSID</summary>
    ''' <param name="vRVSID"></param>
    ''' <returns></returns>
    Private Function CHK_PLAN_REVISESUBFL(vRVSID As String) As Boolean
        'Dim rst As Boolean = False 'False:異常／true:檢核正確
        '(外部參數)
        Dim fg_CANSAVE As Boolean = (rORGKINDGW = "G" OrElse rORGKINDGW = "W")
        '(任一錯誤不可儲存)
        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNo = "" Then Return False '(異常)
        If rSCDate = "" OrElse iSubSeqNO = 0 OrElse vRVSID = "" Then Return False '(異常)
        If rAltDataID = "" OrElse rORGKINDGW = "" Then Return False '(異常)
        'If vFILENAME1 = "" AndAlso vSRCFILENAME1 = "" AndAlso vWAIVED = "" Then Return False '(異常)
        If Not fg_CANSAVE Then Return False '(異常)

        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, vRVSID, rORGKINDGW, rAltDataID)
        If drKR Is Nothing Then Return False '(異常)

        Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, vRVSID)
        Dim drFL As DataRow = If(dtFL.Rows.Count > 0, dtFL.Rows(0), Nothing)
        If drFL Is Nothing Then Return False '(異常)
        Return True
    End Function

    ''' <summary>下一步(檢查若沒有下一步則提示訊息)</summary>
    Private Sub MOVE_NEXT()
        If Hid_RVID.Value <> "" AndAlso (Hid_RVID.Value = Hid_LastRVID.Value) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        ElseIf (ddlSwitchTo.SelectedIndex + 1 >= ddlSwitchTo.Items.Count) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        End If

        '下一步
        'Hid_KBSID.Value = If(Hid_KBID.Value = "" OrElse Hid_KBSID.Value = "", 1, Val(Hid_KBSID.Value) + 1)
        Hid_RVSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex + 1).Value
        Call SHOW_KEY_REVISESUB_RVSID(Hid_RVSID.Value, rORGKINDGW, rAltDataID)
    End Sub

    ''' <summary>下一步</summary>
    Private Sub MOVE_NEXT2()
        If Hid_RVID.Value <> "" AndAlso (Hid_RVID.Value = Hid_LastRVID.Value) Then
            'Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        ElseIf (ddlSwitchTo.SelectedIndex + 1 >= ddlSwitchTo.Items.Count) Then
            'Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        End If

        '下一步
        'Hid_KBSID.Value = If(Hid_KBID.Value = "" OrElse Hid_KBSID.Value = "", 1, Val(Hid_KBSID.Value) + 1)
        Hid_RVSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex + 1).Value
        Call SHOW_KEY_REVISESUB_RVSID(Hid_RVSID.Value, rORGKINDGW, rAltDataID)
    End Sub

    ''' <summary>回上一步</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PREV1_Click(sender As Object, e As EventArgs) Handles BTN_PREV1.Click
        MOVE_PREV()
    End Sub

    ''' <summary>'儲存(暫存)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SAVETMP1_Click(sender As Object, e As EventArgs) Handles BTN_SAVETMP1.Click
        Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVETMP1)
    End Sub

    ''' <summary>'儲存後進下一步</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SAVENEXT1_Click(sender As Object, e As EventArgs) Handles BTN_SAVENEXT1.Click
        Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVENEXT1)
    End Sub

    ''' <summary>'不儲存返回查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_BACK1_Click(sender As Object, e As EventArgs) Handles BTN_BACK1.Click
        Dim url1 As String = String.Concat("TC_05_001.aspx?ID=", TIMS.Get_MRqID(Me)) ' & TIMS.ClearSQM(Request("ID"))
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vBVFID As String = TIMS.GetMyValue(sCmdArg, "BVFID")
        Dim vRVID As String = TIMS.GetMyValue(sCmdArg, "RVID")
        Dim vRVSID As String = TIMS.GetMyValue(sCmdArg, "RVSID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")

        If e.CommandArgument = "" OrElse vBVFID = "" OrElse vRVID = "" OrElse vRVSID = "" Then Return '(異常)

        '顯示檔案資料表
        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then Return ' Exit Sub
        If drPP Is Nothing Then drPP = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
        If drPP Is Nothing Then Return
        If drRR Is Nothing Then drRR = TIMS.Get_RID_DR(Convert.ToString(drPP("RID")), objconn)
        If drRR Is Nothing Then Return ' Exit Sub

        'Hid_RVSID.Value = TIMS.ClearSQM(Hid_RVSID.Value)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, vRVSID, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        If drKR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!!")
            Return
        End If
        Dim vORGKINDGW As String = Convert.ToString(drKR("ORGKINDGW")) 'Hid_ORGKINDGW.Value
        Dim vRVNAME As String = Convert.ToString(drKR("RVNAME"))
        Dim vRVNAME2 As String = String.Concat(vORGKINDGW, vRVID, ".", vRVNAME)

        Select Case e.CommandName
            Case "DELFILE4"
                Dim sErrMsg1 As String = CHKDEL_PLAN_REVISESUBFL(vBVFID)
                If sErrMsg1 <> "" Then
                    Common.MessageBox(Me, sErrMsg1)
                    Return
                End If

                Dim dParms As New Hashtable
                dParms.Add("BVFID", vBVFID)
                Dim rdSql As String = "DELETE PLAN_REVISESUBFL WHERE BVFID=@BVFID"
                Dim iRst As Integer = DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)

                Dim oUploadPath As String = ""
                Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
                Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
                Try
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, CStr(iSubSeqNO)))
                    If vFILENAME1 <> "" Then TIMS.MyFileDelete(Server.MapPath(oUploadPath & vFILENAME1))
                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex)
                    Common.MessageBox(Me, ex.Message)

                    Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                    strErrmsg &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
                    strErrmsg &= String.Concat("vFILEPATH1: ", vFILEPATH1, vbCrLf)
                    strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("(Server.MapPath(oUploadPath & vFILENAME1): ", Server.MapPath(oUploadPath & vFILENAME1), vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                End Try
                'DataGrid1.EditItemIndex = -1

                '顯示上傳檔案／細項
                'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVID,ALTDATAID,ORGKINDGW
                Call SHOW_REVISESUBFL_DG2()

            Case "DOWNLOAD4" '下載
                'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)
                TIMS.SetMyValue2(rPMS4, "RVSID", vRVSID)
                TIMS.SetMyValue2(rPMS4, "BVFID", vBVFID)
                TIMS.SetMyValue2(rPMS4, "ORGNAME", Convert.ToString(drRR("ORGNAME")))
                'TIMS.SetMyValue2(rPMS4, "RVNAME2", vRVNAME2) '項目編號+項目名稱
                Call ResponseZIPFileC51(objconn, Me, rPMS4)

        End Select

        'Call SHOW_REVISESUBFL_DG2()

        'Call SHOW_Detail_BIDCASE(drRR, Hid_BCID.Value, Session(cst_ss_RqProcessType))
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim BTN_DELFILE4 As Button = e.Item.FindControl("BTN_DELFILE4") '刪除
                Dim BTN_DOWNLOAD4 As Button = e.Item.FindControl("BTN_DOWNLOAD4") '下載 
                'Dim labRTUREASON As Label = e.Item.FindControl("labRTUREASON") '退件原因
                'labRTUREASON.Text = Convert.ToString(drv("RTUREASON"))

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                    'HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                    titleMsg = Convert.ToString(drv("OKFLAG"))
                    BTN_DOWNLOAD4.Enabled = (Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")))
                ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                    'LabFileName1.Text = cst_txt_免附文件
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                ElseIf Convert.ToString(drv("WAIVED")) <> "" Then
                    titleMsg = cst_txt_檔案下載
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "BVFID", Convert.ToString(drv("BVFID")))
                TIMS.SetMyValue(sCmdArg, "RVID", Convert.ToString(drv("RVID")))
                TIMS.SetMyValue(sCmdArg, "RVSID", Convert.ToString(drv("RVSID")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                BTN_DELFILE4.CommandArgument = sCmdArg '刪除
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                BTN_DELFILE4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE4.Visible = If(fg_REVISESUBFL_VIEW1, False, True)
        End Select
    End Sub

    ''' <summary>下載報表</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_DOWNLOADRPT1_Click(sender As Object, e As EventArgs) Handles BTN_DOWNLOADRPT1.Click
        Hid_RVSID.Value = TIMS.ClearSQM(Hid_RVSID.Value)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, Hid_RVSID.Value, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        If drKR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        hidReqPlanID.Value = TIMS.ClearSQM(hidReqPlanID.Value)
        hidReqcid.Value = TIMS.ClearSQM(hidReqcid.Value)
        hidReqno.Value = TIMS.ClearSQM(hidReqno.Value)
        Hid_rCDATE.Value = TIMS.ClearSQM(Hid_rCDATE.Value)
        Hid_SubSeqNO.Value = TIMS.ClearSQM(Hid_SubSeqNO.Value)
        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無申請按件編號)，請重新操作!!")
            Return
        End If

        '列印
        Call UTL_PRINT1GW(drPR, drKR)
    End Sub

    ''' <summary>下載報表</summary>
    Private Sub UTL_PRINT1GW(drPR As DataRow, drKR As DataRow)
        If drPR Is Nothing OrElse drKR Is Nothing Then Return
        If drPP Is Nothing Then drPP = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
        If drPP Is Nothing Then Return
        If drRR Is Nothing Then drRR = TIMS.Get_RID_DR(Convert.ToString(drPP("RID")), objconn)
        'Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
        If drRR Is Nothing Then Return

        Dim sMyValue As String = ""
        'PLAN_REVISE
        Dim sAltDataID As String = Convert.ToString(drPR("AltDataID"))
        '取得 KEY_REVISESUB 代號／非流水號
        Dim vORGKINDGW As String = Convert.ToString(drKR("ORGKINDGW"))
        Dim vALTDATAID As String = Convert.ToString(drKR("ALTDATAID"))
        Dim vRVID As String = Convert.ToString(drKR("RVID"))
        Dim vRVNAME As String = Convert.ToString(drKR("RVNAME"))
        Dim vRVNAME2 As String = String.Concat(vORGKINDGW, vRVID, ".", vRVNAME)
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)

        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_1_01_訓練計畫變更表, TIMS.cst_RV_G_9_01_訓練計畫變更表, TIMS.cst_RV_G_11_01_訓練計畫變更表, TIMS.cst_RV_G_14_01_訓練計畫變更表, TIMS.cst_RV_G_15_01_訓練計畫變更表, TIMS.cst_RV_G_18_01_訓練計畫變更表, TIMS.cst_RV_G_20_01_訓練計畫變更表,
                 TIMS.cst_RV_W_1_01_訓練計畫變更表, TIMS.cst_RV_W_9_01_訓練計畫變更表, TIMS.cst_RV_W_11_01_訓練計畫變更表, TIMS.cst_RV_W_14_01_訓練計畫變更表, TIMS.cst_RV_W_15_01_訓練計畫變更表, TIMS.cst_RV_W_18_01_訓練計畫變更表, TIMS.cst_RV_W_20_01_訓練計畫變更表
                '變更申請表 '列印資料內容 SD_14_010_b
                prtFilename = SD_14_010.cst_printFN2 '"SD_14_010_b"
                sMyValue = String.Concat("&Years=", ROC_Years.Value, "&PlanID=", drPR("PlanID"), "&ComIDNO=", drPR("ComIDNO"), "&SeqNo=", drPR("SeqNo"), "&CDate=", TIMS.Cdate3(drPR("CDate")), "&SubSeqNO=", drPR("SubSeqNO"))
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)

            Case TIMS.cst_RV_G_1_02_變更後之課程表, TIMS.cst_RV_G_11_07_變更後之課程表, TIMS.cst_RV_G_14_07_變更後之課程表, TIMS.cst_RV_G_15_02_變更後之課程表, TIMS.cst_RV_G_18_02_變更後之課程表, TIMS.cst_RV_G_20_03_變更後之課程表,
                 TIMS.cst_RV_W_1_02_變更後之課程表, TIMS.cst_RV_W_11_07_變更後之課程表, TIMS.cst_RV_W_14_07_變更後之課程表, TIMS.cst_RV_W_15_02_變更後之課程表, TIMS.cst_RV_W_18_02_變更後之課程表, TIMS.cst_RV_W_20_03_變更後之課程表
                'Public Const TIMS.cst_RV_printFN1c As String = "SD_14_010_R1_c"  '列印變更課程表-課程進度/內容字多版本
                'Public Const TIMS.cst_RV_printFN1d As String = "SD_14_010_R1_d"  '列印變更課程表-課程進度/內容字多版本(增加 -技檢訓練時數)
                '變更後之課程表 SD_14_010_R1_c
                Dim iPTDRID As Integer = SD_14_010.Get_PTDRID(drPR("PlanID"), drPR("ComIDNO"), drPR("SeqNO"), TIMS.Cdate3(drPR("CDate")), drPR("SubSeqNO"), gsCmd)
                prtFilename = If(Convert.ToString(drPP("TMID")) = TIMS.cst_EHour_Use_TMID, SD_14_010.cst_printFN1d, SD_14_010.cst_printFN1c)
                sMyValue = String.Concat("&PTDRID=", iPTDRID, "&AltDataID=", sAltDataID)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)

            Case TIMS.cst_RV_G_11_02_變更後訓練計畫師資名冊, TIMS.cst_RV_W_11_02_變更後訓練計畫師資名冊
                'Common.MessageBox(Me, "變更後訓練計畫師資名冊,下載報表資訊有誤!") 'Return
                Dim vNEWDATA11_1 As String = TIMS.ClearSQM(drPR("NEWDATA11_1"))
                'If vNEWDATA11_1 <> "" Then vNEWDATA11_1 = TIMS.CombiSQLIN(vNEWDATA11_1)
                'concat(a.planid,'-',a.comidno,'-',a.seqno,'-',b.subseqno,'-',FORMAT(b.CDate,'yyyy-MM-dd')) PPIPK
                Dim selsqlstr As String = String.Concat(drPR("PlanID"), "-", drPR("ComIDNO"), "-", drPR("SeqNO"), "-", drPR("SubSeqNO"), "-", TIMS.Cdate3(drPR("CDate"), "yyyy-MM-dd"))
                Dim TechID_Value As String = TIMS.GetTechID(objconn, drPR("PlanID"), drPR("ComIDNO"), drPR("SeqNO"), drPR("SubSeqNO"), TIMS.Cdate3(drPR("CDate"), "yyyy-MM-dd"))
                Dim hPTDRID_Value As String = TIMS.GetPTDRID(objconn, drPR("PlanID"), drPR("ComIDNO"), drPR("SeqNO"), drPR("SubSeqNO"), TIMS.Cdate3(drPR("CDate"), "yyyy-MM-dd"))
                ''SD_14_007*.jrxml/0:未轉班/1:已轉班/2:變更待審 ''Const cst_reportFN0 As String = "SD_14_007"'1:已轉班 'Const cst_reportFN1 As String = "SD_14_007_1" '0:未轉班
                Const cst_reportFN2 As String = "SD_14_007_2" '2:變更待審
                prtFilename = cst_reportFN2
                sMyValue = ""
                TIMS.SetMyValue(sMyValue, "Years", ROC_Years.Value) 'sm.UserInfo.Years - 1911
                TIMS.SetMyValue(sMyValue, "OCID", Convert.ToString(drPP("OCID")))
                TIMS.SetMyValue(sMyValue, "TechID", TechID_Value)
                TIMS.SetMyValue(sMyValue, "selsqlstr", selsqlstr)
                TIMS.SetMyValue(sMyValue, "PTDRID", hPTDRID_Value)
                TIMS.SetMyValue(sMyValue, "PLANID", hidReqPlanID.Value)
                'TIMS.SetMyValue(sMyValue, "ComIDNO", hidReqcid.Value)
                'TIMS.SetMyValue(sMyValue, "SEQNO", hidReqno.Value)
                TIMS.SetMyValue(sMyValue, "Title", Convert.ToString(drRR("ORGPLANNAME")))
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)

            Case TIMS.cst_RV_G_11_03_變更前訓練計畫師資名冊, TIMS.cst_RV_W_11_03_變更前訓練計畫師資名冊
                'Common.MessageBox(Me, "變更前訓練計畫師資名冊,下載報表資訊有誤!") 'Return
                Dim vOLDDATA11_1 As String = TIMS.ClearSQM(drPR("OLDDATA11_1"))
                'If vOLDDATA11_1 <> "" Then vOLDDATA11_1 = TIMS.CombiSQLIN(vOLDDATA11_1)
                Dim vPCSVALUE As String = String.Concat(drPR("PlanID"), "x", drPR("ComIDNO"), "x", drPR("SeqNO"))
                Dim selsqlstr As String = Replace(vPCSVALUE, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                ''SD_14_007*.jrxml/0:未轉班/1:已轉班/2:變更待審
                Const cst_reportFN0 As String = "SD_14_007" '1:已轉班
                'Const cst_reportFN1 As String = "SD_14_007_1" '0:未轉班
                'Const cst_reportFN2 As String = "SD_14_007_2" '2:變更待審
                prtFilename = cst_reportFN0

                sMyValue = ""
                TIMS.SetMyValue(sMyValue, "Years", ROC_Years.Value) 'sm.UserInfo.Years - 1911
                TIMS.SetMyValue(sMyValue, "OCID", Convert.ToString(drPP("OCID")))
                TIMS.SetMyValue(sMyValue, "TechID", vOLDDATA11_1)
                TIMS.SetMyValue(sMyValue, "selsqlstr", selsqlstr)
                TIMS.SetMyValue(sMyValue, "PLANID", hidReqPlanID.Value)
                TIMS.SetMyValue(sMyValue, "ComIDNO", hidReqcid.Value)
                TIMS.SetMyValue(sMyValue, "SEQNO", hidReqno.Value)
                TIMS.SetMyValue(sMyValue, "Title", Convert.ToString(drRR("ORGPLANNAME")))
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)

            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表
                'TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表,
                'Common.MessageBox(Me, "(變更後)師資／助教基本資料表,下載報表資訊有誤!") Return
                Dim vNEWDATA11_1 As String = TIMS.ClearSQM(drPR("NEWDATA11_1"))
                If vNEWDATA11_1 <> "" Then vNEWDATA11_1 = TIMS.CombiSQLIN(vNEWDATA11_1)
                Const cst_printFN2 As String = "SD_14_004B" '產投 師資基本資料表
                prtFilename = cst_printFN2
                sMyValue = ""
                sMyValue = "kjk=fl"
                sMyValue &= String.Concat("&TechID=", vNEWDATA11_1)
                sMyValue &= String.Concat("&Years=", Hid_PlanYear.Value) '"&Years=" , Years.Value
                sMyValue &= String.Concat("&Title=", Convert.ToString(drRR("ORGPLANNAME"))) '"&Title=" & sTitle
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, sMyValue)

            Case TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表
                'TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表,
                'Common.MessageBox(Me, "(變更後)師資／助教基本資料表,下載報表資訊有誤!") Return
                Dim vNEWDATA20_1 As String = TIMS.ClearSQM(drPR("NEWDATA20_1"))
                If vNEWDATA20_1 <> "" Then vNEWDATA20_1 = TIMS.CombiSQLIN(vNEWDATA20_1)
                Const cst_printFN2 As String = "SD_14_004B" '產投 師資基本資料表
                prtFilename = cst_printFN2
                sMyValue = ""
                sMyValue = "kjk=fl"
                sMyValue &= String.Concat("&TechID=", vNEWDATA20_1)
                sMyValue &= String.Concat("&Years=", Hid_PlanYear.Value) '"&Years=" , Years.Value
                sMyValue &= String.Concat("&Title=", Convert.ToString(drRR("ORGPLANNAME"))) '"&Title=" & sTitle
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, sMyValue)

            Case TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表
                'Common.MessageBox(Me, "(變更前)師資／助教基本資料表,下載報表資訊有誤!") 'Return
                Dim vOLDDATA11_1 As String = TIMS.ClearSQM(drPR("OLDDATA11_1"))
                If vOLDDATA11_1 <> "" Then vOLDDATA11_1 = TIMS.CombiSQLIN(vOLDDATA11_1)
                Const cst_printFN2 As String = "SD_14_004B" '產投 師資基本資料表
                prtFilename = cst_printFN2
                sMyValue = "kjk=kjk"
                sMyValue &= String.Concat("&TechID=", vOLDDATA11_1)
                sMyValue &= String.Concat("&Years=", Hid_PlanYear.Value) '"&Years=" , Years.Value
                sMyValue &= String.Concat("&Title=", Convert.ToString(drRR("ORGPLANNAME"))) '"&Title=" & sTitle
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, sMyValue)

            Case TIMS.cst_RV_G_14_02_訓練計畫場地資料表, TIMS.cst_RV_W_14_02_訓練計畫場地資料表
                'Common.MessageBox(Me, "訓練計畫場地資料表,下載報表資訊有誤!") Return
                'Const cst_printFN1 As String = "SD_14_006_1" 'SD_14_006_1.jrxml (未轉班、已轉班)
                Const cst_printFN2 As String = "SD_14_006_2" 'SD_14_006_2.jrxml (變更待審)
                Dim vPCSVALUE As String = String.Concat(drPR("PlanID"), "x", drPR("ComIDNO"), "x", drPR("SeqNO"), "x", drPR("SubSeqNO"), "x", TIMS.Cdate3(drPR("CDATE"), "yyyyMMdd"))
                prtFilename = cst_printFN2
                sMyValue = ""
                sMyValue &= String.Concat("&Years=", ROC_Years.Value)
                sMyValue &= String.Concat("&PCSValue=", vPCSVALUE)
                sMyValue &= String.Concat("&PlanID=", drPR("PlanID"))
                'sValue1 &= String.Concat("&RID=", vRID) 
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)

            Case TIMS.cst_RV_G_14_03_教學環境資料表, TIMS.cst_RV_W_14_03_教學環境資料表
                '"13" '教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'Dim vPCSVALUE As String = String.Concat(drPR("PlanID"), "x", drPR("ComIDNO"), "x", drPR("SeqNO"))
                'Dim selsqlstr As String = Replace(vPCSVALUE, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                'Hid_PlanYear.Value = TIMS.ClearSQM(Hid_PlanYear.Value)
                'prtFilename = "SD_14_014" '0:未轉班' 1:已轉班
                'Dim TSTPRINT As String = If(TIMS.Utl_GetConfigSet("printtest") = "Y", "2", "1") '測試區2／'正式區1 
                'sMyValue = String.Concat("&Years=", ROC_Years.Value, "&selsqlstr=", selsqlstr, "&TPlanID=", drPP("TPlanID"), "&SYears=", Hid_PlanYear.Value)
                'sMyValue &= String.Concat("&TSTPRINT=", TSTPRINT) '正式區1 '測試區2

                'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_014_1&Years=2024&selsqlstr=5113-36647236-5-1-2024-08-29&TPlanID=28&SYears=2024&TSTPRINT=1&UserID=snoopy
                'sql &= " ,CONCAT(a.PLANID,'-',a.COMIDNO,'-',a.SEQNO,'-',c.SUBSEQNO,'-',FORMAT(c.CDATE,'yyyy-MM-dd')) PCS2" & vbCrLf
                Dim vPCS2 As String = String.Concat(drPR("PLANID"), "-", drPR("COMIDNO"), "-", drPR("SEQNO"), "-", drPR("SUBSEQNO"), "-", CDate(drPR("CDATE")).ToString("yyyy-MM-dd"))
                Hid_PlanYear.Value = TIMS.ClearSQM(Hid_PlanYear.Value)
                Dim TSTPRINT As String = If(TIMS.Utl_GetConfigSet("printtest") = "Y", "2", "1") '測試區2／'正式區1 
                prtFilename = "SD_14_014_1" '2:變更待審
                sMyValue = ""
                sMyValue &= "&Years=" & Hid_PlanYear.Value 'sm.UserInfo.Years
                sMyValue &= "&selsqlstr=" & vPCS2
                sMyValue &= "&TPlanID=" & drPP("TPlanID")
                sMyValue &= "&SYears=" & Hid_PlanYear.Value
                sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)
            Case Else
                Common.MessageBox(Me, String.Concat("下載報表資訊有誤!", vRVNAME2, "x", vGW_ALT_RVID))
                Return
        End Select
    End Sub

    ''' <summary>以目前版本批次送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SENTBATVER_Click(sender As Object, e As EventArgs) Handles BTN_SENTBATVER.Click
        Common.MessageBox(Me, "以目前版本批次送出,資訊有誤!")
        Return
    End Sub

    ''' <summary>以目前版本送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SENDCURRVER_Click(sender As Object, e As EventArgs) Handles BTN_SENDCURRVER.Click
        Hid_RVSID.Value = TIMS.ClearSQM(Hid_RVSID.Value)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, Hid_RVSID.Value, Hid_ORGKINDGW.Value, Hid_ALTDATAID.Value)
        If drKR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!!")
            Return
        End If
        hidReqPlanID.Value = TIMS.ClearSQM(hidReqPlanID.Value)
        hidReqcid.Value = TIMS.ClearSQM(hidReqcid.Value)
        hidReqno.Value = TIMS.ClearSQM(hidReqno.Value)
        Hid_rCDATE.Value = TIMS.ClearSQM(Hid_rCDATE.Value)
        Hid_SubSeqNO.Value = TIMS.ClearSQM(Hid_SubSeqNO.Value)
        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", hidReqPlanID.Value) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", hidReqcid.Value) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", hidReqno.Value) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", TIMS.Cdate3(Hid_rCDATE.Value)) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", Val(Hid_SubSeqNO.Value)) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        'Dim drPR As DataRow = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無申請按件編號)，請重新操作!!!")
            Return
        End If

        Call UTL_SENDCURRVER(drPR, drKR)
    End Sub

    ''' <summary>以目前版本送出</summary>
    ''' <param name="drPR"></param>
    ''' <param name="drKR"></param>
    Private Sub UTL_SENDCURRVER(drPR As DataRow, drKR As DataRow)
        If drPR Is Nothing OrElse drKR Is Nothing Then Return
        If drPP Is Nothing Then drPP = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNo, objconn)
        If drPP Is Nothing Then Return
        If drRR Is Nothing Then drRR = TIMS.Get_RID_DR(Convert.ToString(drPP("RID")), objconn)
        'Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
        If drRR Is Nothing Then Return

        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)

        Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, vRVSID)
        Dim drFL As DataRow = If(dtFL.Rows.Count > 0, dtFL.Rows(0), Nothing)
        If drFL IsNot Nothing Then
            Common.MessageBox(Me, "已儲存過該文件，不可再次操作!")
            Return
        End If

        Dim sMyValue As String = ""
        'PLAN_REVISE
        Dim sAltDataID As String = Convert.ToString(drPR("AltDataID"))
        '取得 KEY_REVISESUB 代號／非流水號
        Dim vORGKINDGW As String = Convert.ToString(drKR("ORGKINDGW"))
        Dim vALTDATAID As String = Convert.ToString(drKR("ALTDATAID"))
        Dim vRVID As String = Convert.ToString(drKR("RVID"))
        Dim vRVNAME As String = Convert.ToString(drKR("RVNAME"))
        Dim vRVNAME2 As String = String.Concat(vORGKINDGW, vRVID, ".", vRVNAME)
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)

        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_03_變更前訓練計畫師資名冊, TIMS.cst_RV_W_11_03_變更前訓練計畫師資名冊
                'Common.MessageBox(Me, "變更前訓練計畫師資名冊,以目前版本送出資訊有誤!") Return
                Dim vOLDDATA11_1 As String = TIMS.ClearSQM(drPR("OLDDATA11_1"))
                'If vOLDDATA11_1 <> "" Then vOLDDATA11_1 = TIMS.CombiSQLIN(vOLDDATA11_1)
                Dim vPCSVALUE As String = String.Concat(drPR("PlanID"), "x", drPR("ComIDNO"), "x", drPR("SeqNO"))
                Dim selsqlstr As String = Replace(vPCSVALUE, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                ''SD_14_007*.jrxml/0:未轉班/1:已轉班/2:變更待審
                Const cst_reportFN0 As String = "SD_14_007" '1:已轉班
                'Const cst_reportFN1 As String = "SD_14_007_1" '0:未轉班 'Const cst_reportFN2 As String = "SD_14_007_2" '2:變更待審
                prtFilename = cst_reportFN0

                sMyValue = ""
                TIMS.SetMyValue(sMyValue, "Years", ROC_Years.Value) 'sm.UserInfo.Years - 1911
                TIMS.SetMyValue(sMyValue, "OCID", Convert.ToString(drPP("OCID")))
                TIMS.SetMyValue(sMyValue, "TechID", vOLDDATA11_1)
                TIMS.SetMyValue(sMyValue, "selsqlstr", selsqlstr)
                TIMS.SetMyValue(sMyValue, "PLANID", hidReqPlanID.Value)
                TIMS.SetMyValue(sMyValue, "ComIDNO", hidReqcid.Value)
                TIMS.SetMyValue(sMyValue, "SEQNO", hidReqno.Value)
                TIMS.SetMyValue(sMyValue, "Title", Convert.ToString(drRR("ORGPLANNAME")))
                Dim s_RPTURL As String = ReportQuery.GetReportUrl2(Me, prtFilename, sMyValue)
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, sMyValue)
                Dim s_PDF_byte As Byte() = Nothing
                Try
                    Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
                Catch ex As Exception
                    Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
                    eErrmsg &= String.Concat(", s_RPTURL: ", s_RPTURL)
                    eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
                    eErrmsg &= String.Concat(", sMyValue: ", sMyValue)
                    TIMS.LOG.Error(eErrmsg, ex)
                    Common.MessageBox(Me, String.Concat(vRVNAME2, "下載檔案有誤，請確認檔案是否正確!"))
                    Return
                End Try
                '以目前版本送出 儲存 PDF_byte
                If s_PDF_byte IsNot Nothing Then Call SAVE_PLAN_REVISESUBFL_PDF_FILE(drPR, s_PDF_byte)

            Case Else
                Common.MessageBox(Me, String.Concat("以目前版本送出資訊有誤!!", vRVNAME2, "x", vGW_ALT_RVID))
                Return
        End Select

        '顯示上傳檔案／細項
        'PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO,RVID,ALTDATAID,ORGKINDGW
        Call SHOW_REVISESUBFL_DG2()
    End Sub

    ''' <summary>'以目前版本送出 儲存 PDF_byte</summary>
    ''' <param name="drPR"></param>
    ''' <param name="s_PDF_byte"></param>
    Private Sub SAVE_PLAN_REVISESUBFL_PDF_FILE(drPR As DataRow, s_PDF_byte() As Byte)
        If drPR Is Nothing Then Return
        '取得代號／非流水號
        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)
        'Dim vALTDATAID As String = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
        'Dim vSubNO As String = TIMS.ClearSQM(Hid_SubSeqNO.Value)

        Dim vUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = "" 'TIMS.GET_FILENAME1_EV(vBCID, vKBSID, vPCS, "pdf")
        'Dim vFILEPATH1 As String = ""
        Dim vSRCFILENAME1 As String = "" 'vFILENAME1 'Convert.ToString(oSRCFILENAME1)
        '上傳檔案/存檔：檔名
        Try
            '上傳檔案 '計畫ID／機構ID
            vUploadPath = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO)
            vFILENAME1 = TIMS.GET_FILENAME1_PR(vRID, vRVSID, rAltDataID, "pdf")
            vSRCFILENAME1 = vFILENAME1
            '上傳檔案/存檔：檔名
            Call TIMS.MyCreateDir(Me, vUploadPath)
            File.WriteAllBytes(Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), s_PDF_byte)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            'strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            'strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)): ", Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), vbCrLf)
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "BCFID", If(vUploadPath <> "", iBCFID, -1)) '(可再次傳送)
            TIMS.SetMyValue2(rPMS2, "PLANID", rPlanID)
            TIMS.SetMyValue2(rPMS2, "COMIDNO", rComIDNO)
            TIMS.SetMyValue2(rPMS2, "SEQNO", rSeqNo)
            TIMS.SetMyValue2(rPMS2, "CDATE", rSCDate)
            TIMS.SetMyValue2(rPMS2, "SUBSEQNO", iSubSeqNO)
            TIMS.SetMyValue2(rPMS2, "RVSID", Hid_RVSID.Value)
            TIMS.SetMyValue2(rPMS2, "ALTDATAID", rAltDataID)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", rORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            'TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED)
            Call SAVE_PLAN_REVISESUBFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub 'Throw ex
        End Try

    End Sub

    ''' <summary>下載檔案</summary>
    ''' <param name="objconn"></param>
    ''' <param name="MyPage"></param>
    ''' <param name="rPMS4"></param>
    Private Sub ResponseZIPFileC51(objconn As SqlConnection, MyPage As Page, ByRef rPMS4 As Hashtable)
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS4, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS4, "FILEPATH1")
        Dim vRVSID As String = TIMS.GetMyValue2(rPMS4, "RVSID")
        Dim vBVFID As String = TIMS.GetMyValue2(rPMS4, "BVFID")
        Dim vORGNAME As String = TIMS.GetMyValue2(rPMS4, "ORGNAME")
        'Dim vRVNAME2 As String = TIMS.GetMyValue2(rPMS4, "RVNAME2") '項目編號+項目名稱
        Dim vTECHID As String = TIMS.GetMyValue2(rPMS4, "TECHID")
        Dim vBVFTID As String = TIMS.GetMyValue2(rPMS4, "BVFTID")

        Dim dtFL As DataTable = GET_PLAN_REVISESUBFL_TB(objconn, vRVSID)
        Dim drFL As DataRow = If(dtFL.Rows.Count > 0, dtFL.Rows(0), Nothing)
        If drFL Is Nothing Then Return

        Dim vORGKINDGW As String = TIMS.ClearSQM(drFL("ORGKINDGW"))
        Dim vALTDATAID As String = Convert.ToString(drFL("ALTDATAID"))
        Dim vRVID As String = TIMS.ClearSQM(drFL("RVID"))
        Dim vRVNAME As String = TIMS.ClearSQM(drFL("RVNAME"))
        Dim vRVNAME2 As String = TIMS.ClearSQM(drFL("RVNAME2")) 'concat(k.RVID,'.',k.RVNAME) RVNAME2
        Dim vRVNAME3 As String = Convert.ToString(drFL("RVNAME3")) 'concat(k.ORGKINDGW,k.RVID,'x',k.RVNAME) RVNAME3
        Dim vGW_ALT_RVID As String = Convert.ToString(drFL("GW_ALT_RVID")) 'concat(k.ORGKINDGW,'_',a.ALTDATAID,'_',k.RVID) GW_ALT_RVID

        '取得代號／非流水號
        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
        Dim vYEARS_ROC As String = TIMS.ClearSQM(ROC_Years.Value)

        Dim Template_ZipPath2 As String = TIMS.GET_Template_ZipPath2(vBVFID)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath2)) Then
            Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath2))
        End If

        Dim iFILECNT As Integer = 0
        Dim tryFIND As String = ""
        'PLAN_REVISE
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表,
                 TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本,
                 TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表

                Dim dtFLTT As DataTable = TIMS.GET_PLAN_REVISESUBFL_TT(objconn, vBVFID)
                If dtFLTT Is Nothing OrElse dtFLTT.Rows.Count = 0 Then
                    Common.MessageBox(MyPage, String.Concat("(", vRVNAME3, ")查無資料!"))
                    Return
                End If
                tryFIND = If(vBVFTID <> "" AndAlso vFILENAME1 <> "", String.Concat("BVFTID=", vBVFTID, " AND FILENAME1='", vFILENAME1, "'"), "")
                If tryFIND <> "" Then
                    If dtFLTT.Select(tryFIND).Length = 0 Then
                        Common.MessageBox(MyPage, String.Concat("(", vRVNAME3, ")查無資料!"))
                        Return
                    End If
                    For Each drFLTT As DataRow In dtFLTT.Select(tryFIND)
                        Dim vTEACHCNAME As String = "" ' Convert.ToString(drFLTT("TEACHCNAME"))
                        Dim vTEACHERID As String = "" 'Convert.ToString(drFLTT("TEACHERID"))
                        Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                        Dim oFILEPATH1 As String = ""
                        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                        Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                        Dim t_FILENAME_TT As String = "" ' String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                        Dim t_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT))
                        Try
                            vTEACHCNAME = Convert.ToString(drFLTT("TEACHCNAME"))
                            vTEACHERID = Convert.ToString(drFLTT("TEACHERID"))
                            oFILENAME1 = Convert.ToString(drFLTT("FILENAME1"))
                            oFILEPATH1 = Convert.ToString(drFLTT("FILEPATH1"))
                            oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO))
                            s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                            t_FILENAME_TT = String.Concat(vYEARS_ROC, "_", vORGNAME, "_", vRVNAME3, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                            t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME_TT))
                            If IO.File.Exists(s_FilePath1) Then
                                iFILECNT += 1
                                Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                File.WriteAllBytes(t_FilePath1, dbyte)
                            End If
                        Catch ex As Exception
                            Dim strErrmsg As String = "/*Public Shared Sub ResponseZIPFileC51*/" & vbCrLf
                            strErrmsg &= String.Concat("vTEACHCNAME: ", vTEACHCNAME, vbCrLf, "vTEACHERID: ", vTEACHERID, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                            strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("t_FILENAME_TT: ", t_FILENAME_TT, vbCrLf)
                            strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                            Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                        End Try
                    Next
                Else
                    For Each drFLTT As DataRow In dtFLTT.Rows
                        Dim vTEACHCNAME As String = "" ' Convert.ToString(drFLTT("TEACHCNAME"))
                        Dim vTEACHERID As String = "" 'Convert.ToString(drFLTT("TEACHERID"))
                        Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                        Dim oFILEPATH1 As String = ""
                        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                        Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                        Dim t_FILENAME_TT As String = "" ' String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                        Dim t_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT))
                        Try
                            vTEACHCNAME = Convert.ToString(drFLTT("TEACHCNAME"))
                            vTEACHERID = Convert.ToString(drFLTT("TEACHERID"))
                            oFILENAME1 = Convert.ToString(drFLTT("FILENAME1"))
                            oFILEPATH1 = Convert.ToString(drFLTT("FILEPATH1"))
                            oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO))
                            s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                            t_FILENAME_TT = String.Concat(vYEARS_ROC, "_", vORGNAME, "_", vRVNAME3, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                            t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME_TT))
                            If IO.File.Exists(s_FilePath1) Then
                                iFILECNT += 1
                                Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                File.WriteAllBytes(t_FilePath1, dbyte)
                            End If
                        Catch ex As Exception
                            Dim strErrmsg As String = "/*Public Shared Sub ResponseZIPFileC51*/" & vbCrLf
                            strErrmsg &= String.Concat("vTEACHCNAME: ", vTEACHCNAME, vbCrLf, "vTEACHERID: ", vTEACHERID, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                            strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("t_FILENAME_TT: ", t_FILENAME_TT, vbCrLf)
                            strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                            Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                        End Try
                    Next
                End If

            Case Else
                '上傳檔案 '計畫ID／機構ID
                Dim oFILENAME1 As String = "" 'vFILENAME1
                Dim oFILEPATH1 As String = ""
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                '年度申請階段_單位名稱_項目編號+項目名稱
                Dim t_FILENAME As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf")
                Dim t_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))
                Try
                    oFILENAME1 = vFILENAME1
                    oFILEPATH1 = vFILEPATH1
                    oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, iSubSeqNO))
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度_單位名稱_項目編號名稱
                    t_FILENAME = String.Concat(vYEARS_ROC, "_", vORGNAME, "_", vRVNAME3, ".pdf")
                    t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME))
                    If oFILENAME1 <> "" AndAlso IO.File.Exists(s_FilePath1) Then
                        iFILECNT += 1
                        Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                        File.WriteAllBytes(t_FilePath1, dbyte)
                    End If
                Catch ex As Exception
                    Dim strErrmsg As String = "/*Public Shared Sub ResponseZIPFileC51*/" & vbCrLf
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                    strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("t_FILENAME: ", t_FILENAME, vbCrLf)
                    strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
        End Select

        If iFILECNT = 0 Then
            Common.MessageBox(Me, "下載數量為空,查無有效下載的檔案!")
            Return
        End If

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        'Dim zipFileName As String = String.Concat("p", strNOW, "_", vBCID, ".zip")
        Dim zipFileName As String = String.Concat("p", vYEARS_ROC, "_", vORGNAME, "_", vBVFID, "_", strNOW, ".zip")
        Dim filenames As String() = Directory.GetFiles(MyPage.Server.MapPath(String.Concat(Template_ZipPath2, "/")))
        Dim full_zipFileName As String = String.Concat(Template_ZipPath2, "/", zipFileName)
        Using zp As New ZipOutputStream(System.IO.File.Create(MyPage.Server.MapPath(full_zipFileName)))
            zp.SetLevel(6) ' 設定壓縮比
            ' 逐一將資料夾內的檔案抓出來壓縮，並寫入至目的檔(.ZIP)
            For Each filename As String In filenames
                Dim entry As New ZipEntry(Path.GetFileName(filename)) With {.IsUnicodeText = True}
                zp.PutNextEntry(entry) '建立下一個壓縮檔案或資料夾條目
                Try
                    Using fs As New FileStream(filename, FileMode.Open)
                        Dim buffer As Byte() = New Byte(fs.Length - 1) {}
                        Dim i_readLength As Integer
                        Do
                            i_readLength = fs.Read(buffer, 0, buffer.Length)
                            If i_readLength > 0 Then zp.Write(buffer, 0, i_readLength)
                        Loop While i_readLength > 0
                    End Using
                Catch ex As Exception
                    Dim strErrmsg As String = "/*ResponseZIPFileC51*/" & vbCrLf
                    strErrmsg &= String.Concat("full_zipFileName: ", full_zipFileName, vbCrLf)
                    strErrmsg &= String.Concat("filename: ", filename, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                    Common.MessageBox(MyPage, String.Concat("檔案下載有誤，請重新操作!", ex.Message))
                    Return
                End Try
                '假設處理某段程序需花費1毫秒 (避免機器不同步)
                Threading.Thread.Sleep(1)
                '刪除檔案
                Call TIMS.MyFileDelete(filename)
            Next
        End Using

        With MyPage
            Dim File As New FileInfo(.Server.MapPath(full_zipFileName))
            TIMS.SAVE_ADP_ZIPFILE(objconn, "-c51x2257", File)
            ' Clear the content of the response
            .Response.ClearContent()
            ' LINE1 Add the file name And attachment, which will force the open/cance/save dialog To show, to the header
            .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
            'Response.Headers["Content-Disposition"] = "attachment; filename=" + zipFileName;
            ' Add the file size into the response header
            .Response.AddHeader("Content-Length", File.Length.ToString())
            ' Set the ContentType
            .Response.ContentType = "application/zip"
            .Response.TransmitFile(File.FullName)
            ' End the response
            TIMS.Utl_RespWriteEnd(MyPage, objconn, "") '.Response.End()
        End With
    End Sub

    ''' <summary>(刪除資訊檢查) PLAN_REVISESUBFL 依 BVFID</summary>
    ''' <param name="vBVFID"></param>
    ''' <returns></returns>
    Private Function CHKDEL_PLAN_REVISESUBFL(vBVFID As String) As String
        Dim rst As String = ""

        Dim sParms1 As New Hashtable
        sParms1.Add("BVFID", vBVFID)
        Dim sSql As String = ""
        sSql &= " SELECT a.BVFID,a.RVSID" & vbCrLf
        sSql &= " ,kb.ORGKINDGW,kb.RVID,kb.RVNAME" & vbCrLf
        sSql &= " FROM PLAN_REVISESUBFL a" & vbCrLf
        sSql &= " JOIN KEY_REVISESUB kb on kb.RVSID=a.RVSID" & vbCrLf
        sSql &= " WHERE a.BVFID=@BVFID" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms1)
        If dr1 Is Nothing Then Return "查無資料!"

        Return rst
    End Function

    Private Sub DataGrid10_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid10.ItemCommand
        'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        If e.CommandArgument = "" Then Return
        Dim sCmdArg As String = e.CommandArgument

        Dim vTECHID As String = TIMS.GetMyValue(sCmdArg, "TECHID")
        Dim vBVFTID As String = TIMS.GetMyValue(sCmdArg, "BVFTID")
        Dim vBVFID As String = TIMS.GetMyValue(sCmdArg, "BVFID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")
        Dim vRVID As String = TIMS.GetMyValue(sCmdArg, "RVID")
        Dim vRVSID As String = TIMS.GetMyValue(sCmdArg, "RVSID")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_ALTDATAID.Value = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        Dim vALTDATAID As String = Hid_ALTDATAID.Value
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼),請確認機構正確!!")
            Return
        End If
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)

        Dim sMyValue As String = ""
        Select Case e.CommandName
            Case "REPORT10"
                If e.CommandArgument = "" OrElse vTECHID = "" Then Return
                'PLAN_REVISE
                Dim htSS As New Hashtable
                TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
                TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
                TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
                TIMS.SetMyValue2(htSS, "rCDate", rSCDate) 'Request("CDate")
                TIMS.SetMyValue2(htSS, "rSubNo", iSubSeqNO) 'Request("SubNo")
                'PLAN_REVISE
                If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
                '查無傳入資訊 '基本資料產生問題
                If drPR Is Nothing Then Return ' Exit Sub
                'Common.MessageBox(Me, "(變更後)師資／助教基本資料表,下載報表資訊有誤!") Return
                'Dim vNEWDATA11_1 As String = TIMS.ClearSQM(drPR("NEWDATA11_1"))
                'If vNEWDATA11_1 <> "" Then vNEWDATA11_1 = TIMS.CombiSQLIN(vNEWDATA11_1)
                Const cst_printFN2 As String = "SD_14_004B" '產投 師資基本資料表
                prtFilename = cst_printFN2
                sMyValue = ""
                sMyValue = "kjk=fl"
                sMyValue &= String.Concat("&TechID=", vTECHID) 'vNEWDATA11_1
                sMyValue &= String.Concat("&Years=", Hid_PlanYear.Value) '"&Years=" , Years.Value
                sMyValue &= String.Concat("&Title=", Convert.ToString(drRR("ORGPLANNAME"))) '"&Title=" & sTitle
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, sMyValue)

            Case "DOWNLOAD10"
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)
                TIMS.SetMyValue2(rPMS4, "RVSID", vRVSID)
                TIMS.SetMyValue2(rPMS4, "BVFID", vBVFID)
                TIMS.SetMyValue2(rPMS4, "ORGNAME", Convert.ToString(drRR("ORGNAME")))

                TIMS.SetMyValue2(rPMS4, "TECHID", vTECHID)
                TIMS.SetMyValue2(rPMS4, "BVFTID", vBVFTID)
                'TIMS.SetMyValue2(rPMS4, "RVID", vRVID)
                'TIMS.SetMyValue2(rPMS4, "RVNAME2", vRVNAME2) '項目編號+項目名稱
                Call ResponseZIPFileC51(objconn, Me, rPMS4)

            Case "DELFILE10"
                If e.CommandArgument = "" OrElse vBVFTID = "" Then Return
                '"ORG_BIDCASEFL"
                Dim dParms As New Hashtable
                dParms.Add("BVFTID", vBVFTID)
                Dim rdSql As String = "DELETE PLAN_REVISESUBFL_TT WHERE BVFTID=@BVFTID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)

                Dim oUploadPath As String = ""
                Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
                Dim rSCDateNT As String = TIMS.Cdate3(rSCDate, "yyyyMMdd")
                Try
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, rSCDateNT, CStr(iSubSeqNO)))
                    If vFILENAME1 <> "" Then TIMS.MyFileDelete(Server.MapPath(oUploadPath & vFILENAME1))
                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex)
                    Common.MessageBox(Me, ex.Message)

                    Dim strErrmsg As String = String.Concat("ex.Message:", ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                    strErrmsg &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
                    strErrmsg &= String.Concat("vFILEPATH1: ", vFILEPATH1, vbCrLf)
                    strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("(Server.MapPath(oUploadPath & vFILENAME1): ", Server.MapPath(oUploadPath & vFILENAME1), vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                End Try

                'DataGrid1.EditItemIndex = -1
                '師資／助教基本資料表
                Call SHOW_DATAGRID_10(vGW_ALT_RVID)

        End Select
    End Sub

    Private Sub DataGrid10_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid10.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim chkItem1 As HtmlInputCheckBox = e.Item.FindControl("chkItem1")
                Dim HDG10_TechID As HtmlInputHidden = e.Item.FindControl("HDG10_TechID")
                Dim HDG10_RID As HtmlInputHidden = e.Item.FindControl("HDG10_RID")
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                Dim BTN_REPORT10 As Button = e.Item.FindControl("BTN_REPORT10")
                Dim BTN_DOWNLOAD10 As Button = e.Item.FindControl("BTN_DOWNLOAD10")
                Dim BTN_DELFILE10 As Button = e.Item.FindControl("BTN_DELFILE10")

                e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e)
                chkItem1.Attributes("onclick") = String.Concat("selectOnlyThis('", chkItem1.ClientID, "',", iDG10_ROWS, ",'DataGrid10')")
                '0:未轉班,1:已轉班 '未轉班(依計畫查詢) 含列印
                HDG10_TechID.Value = Convert.ToString(drv("TechID"))
                HDG10_RID.Value = Convert.ToString(drv("RID"))

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                    'HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                    titleMsg = Convert.ToString(drv("OKFLAG"))
                    BTN_DOWNLOAD10.Enabled = (Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")))
                ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                    'LabFileName1.Text = cst_txt_免附文件
                    titleMsg = cst_txt_免附文件
                Else
                    titleMsg = cst_tpmsg_enb9
                    BTN_DOWNLOAD10.Enabled = False
                    BTN_DELFILE10.Enabled = False
                    Call TIMS.Tooltip(BTN_DELFILE10, cst_tpmsg_enb9, True)
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD10, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TECHID", drv("TECHID"))
                TIMS.SetMyValue(sCmdArg, "BVFTID", drv("BVFTID"))
                TIMS.SetMyValue(sCmdArg, "BVFID", drv("BVFID"))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", Convert.ToString(drv("FILEPATH1")))
                TIMS.SetMyValue(sCmdArg, "RVID", Convert.ToString(drv("RVID")))
                TIMS.SetMyValue(sCmdArg, "RVSID", Convert.ToString(drv("RVSID")))
                BTN_REPORT10.CommandArgument = sCmdArg '報表下載
                BTN_DOWNLOAD10.CommandArgument = sCmdArg '檔案下載
                BTN_DELFILE10.CommandArgument = sCmdArg '刪除檔案
                BTN_DELFILE10.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE10.Visible = If(fg_REVISESUBFL_VIEW1, False, True)

                Dim fg_BTN_REPORT10_Visible As Boolean = True
                Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
                Dim vALTDATAID As String = Hid_ALTDATAID.Value
                Dim vRVID As String = Hid_RVID.Value ' Convert.ToString(drv("RVID"))
                Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)
                Select Case vGW_ALT_RVID
                    Case TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本
                        fg_BTN_REPORT10_Visible = False '(沒有報表可供下載)
                End Select
                BTN_REPORT10.Visible = fg_BTN_REPORT10_Visible
        End Select
    End Sub

    ''' <summary>最近一次版本送件</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_latestSend1_Click(sender As Object, e As EventArgs) Handles bt_latestSend1.Click
        UTL_LATEST(cst_MTYPE_LATEST_SEND1) '最近一次版本送件
    End Sub

    ''' <summary>最近一次版本送件</summary>
    ''' <param name="vMTYPE"></param>
    Private Sub UTL_LATEST(vMTYPE As String)
        'cst_MTYPE_LATEST_SEND1: 最近一次版本送件 /cst_MTYPE_LATEST_DOWN1: 最近一次版本-下載
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        'PLAN_REVISE
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "rPlanID", rPlanID) 'Request("PlanID")
        TIMS.SetMyValue2(htSS, "rComIDNO", rComIDNO) 'Request("cid")
        TIMS.SetMyValue2(htSS, "rSeqNo", rSeqNo) 'Request("no")
        TIMS.SetMyValue2(htSS, "rCDate", rSCDate) 'Request("CDate")
        TIMS.SetMyValue2(htSS, "rSubNo", iSubSeqNO) 'Request("SubNo")
        'PLAN_REVISE
        If drPR Is Nothing Then drPR = TC_05_001_chg.Get_PlanReviseDataRow(htSS, objconn)
        '查無傳入資訊 '基本資料產生問題
        If drPR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無申請案件)，請重新選擇!")
            Return ' Exit Sub
        End If
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!")
            Return
        End If
        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim vALTDATAID As String = TIMS.ClearSQM(Hid_ALTDATAID.Value)
        Dim drKR As DataRow = TIMS.GET_KEY_REVISESUB(objconn, vRVSID, vORGKINDGW, vALTDATAID)
        If drKR Is Nothing Then Return

        Dim vRVID As String = Convert.ToString(drKR("RVID"))
        Dim vRVNAME As String = Convert.ToString(drKR("RVNAME"))
        Dim vRVNAME2 As String = String.Concat(vORGKINDGW, vRVID, ".", vRVNAME)
        Dim vGW_ALT_RVID As String = String.Concat(vORGKINDGW, "_", vALTDATAID, "_", vRVID)
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表,
                 TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表,
                 TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表,
                 TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本
                '師資助教基本資料表/師資學經歷證書影本
                Call FILE_COPY1_TT_TT2(drPR, drKR, vMTYPE, vGW_ALT_RVID)
            Case Else
                Common.MessageBox(Me, "(查無資料)")
                Return
        End Select
    End Sub

    ''' <summary> 師資助教基本資料表/師資學經歷證書影本 </summary>
    ''' <param name="drPR"></param>
    ''' <param name="drKR"></param>
    ''' <param name="vMTYPE"></param>
    Private Sub FILE_COPY1_TT_TT2(drPR As DataRow, drKR As DataRow, vMTYPE As String, vGW_ALT_RVID As String)
        Hid_TECHID.Value = ""
        For Each eItem As DataGridItem In DataGrid10.Items
            Dim chkItem1 As HtmlInputCheckBox = eItem.FindControl("chkItem1")
            Dim HDG10_TechID As HtmlInputHidden = eItem.FindControl("HDG10_TechID")
            'Dim HDG10_RID As HtmlInputHidden = eItem.FindControl("HDG10_RID") AndAlso HDG10_RID.Value <> ""
            If chkItem1.Checked AndAlso HDG10_TechID.Value <> "" Then
                Hid_TECHID.Value = HDG10_TechID.Value
                Exit For
            End If
        Next
        Hid_TECHID.Value = TIMS.ClearSQM(Hid_TECHID.Value)
        Dim vTECHID As String = Hid_TECHID.Value
        If vTECHID = "" Then
            Common.MessageBox(Me, "最近一次版本送件,資訊有誤(未選擇老師)，請重新操作!")
            Return
        End If
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Return ' Exit Sub
        End If

        Hid_BVFID.Value = TIMS.ClearSQM(Hid_BVFID.Value)
        Dim vBVFID As String = Hid_BVFID.Value
        Dim fg_OK As Boolean = True '(執行結果正常:TRUE 異常:FALSE)
        'PLAN_REVISESUBFL_TT
        Dim drOF1 As DataRow = Nothing
        Dim srPMS As New Hashtable
        srPMS.Add("COMIDNO", drPR("COMIDNO"))
        srPMS.Add("TECHID", vTECHID)
        srPMS.Add("RVSID", drKR("RVSID"))
        srPMS.Add("ALTDATAID", drKR("ALTDATAID"))
        srPMS.Add("ORGKINDGW", drKR("ORGKINDGW"))
        srPMS.Add("BVFID", vBVFID)
        drOF1 = GET_OLDFILE1_PLAN_REVISESUBFL_TT_row(srPMS)

        '師資助教基本資料表/師資學經歷證書影本
        Dim oUploadPath2 As String = ""
        Dim drOF2 As DataRow = Nothing
        Dim srPMS2 As New Hashtable
        Select Case vGW_ALT_RVID
            Case TIMS.cst_RV_G_11_04_變更後師資助教基本資料表, TIMS.cst_RV_W_11_04_變更後師資助教基本資料表,
                 TIMS.cst_RV_G_11_05_變更前師資助教基本資料表, TIMS.cst_RV_W_11_05_變更前師資助教基本資料表,
                 TIMS.cst_RV_G_20_02_變更後師資助教基本資料表, TIMS.cst_RV_W_20_02_變更後師資助教基本資料表
                srPMS2.Add("TPLANID", sm.UserInfo.TPlanID)
                srPMS2.Add("COMIDNO", drPR("COMIDNO"))
                srPMS2.Add("TECHID", vTECHID)
                srPMS2.Add("ORGKINDGW", drKR("ORGKINDGW"))
                '師資助教基本資料表 
                drOF2 = TIMS.GET_ORG_BIDCASEFL_TT_USELATESTVER_row(objconn, srPMS2)
                oUploadPath2 = TIMS.GetMyValue2(srPMS2, "UploadPath")
                If drOF1 Is Nothing AndAlso drOF2 Is Nothing Then
                    Common.MessageBox(Me, String.Concat("選擇老師(最近一次版本送件,查無資料)!", vTECHID))
                    Return
                End If
                If oUploadPath2 = "" Then
                    Common.MessageBox(Me, String.Concat("選擇老師(最近一次版本送件,查無檔案路徑資料)!", vTECHID))
                    Return
                End If

            Case TIMS.cst_RV_G_11_06_各授課師資學經歷證書影本, TIMS.cst_RV_W_11_06_各授課師資學經歷證書影本
                srPMS2.Add("TPLANID", sm.UserInfo.TPlanID)
                srPMS2.Add("COMIDNO", drPR("COMIDNO"))
                srPMS2.Add("TECHID", vTECHID)
                srPMS2.Add("ORGKINDGW", drKR("ORGKINDGW"))
                ' 師資學經歷證書影本
                drOF2 = TIMS.GET_ORG_BIDCASEFL_TT2_USELATESTVER_row(objconn, srPMS2)
                oUploadPath2 = TIMS.GetMyValue2(srPMS2, "UploadPath")
                If drOF1 Is Nothing AndAlso drOF2 Is Nothing Then
                    Common.MessageBox(Me, String.Concat("選擇老師(最近一次版本送件,查無資料)!!", vTECHID))
                    Return
                End If
                If oUploadPath2 = "" Then
                    Common.MessageBox(Me, String.Concat("選擇老師(最近一次版本送件,查無檔案路徑資料)!!", vTECHID))
                    Return
                End If

            Case Else
                Common.MessageBox(Me, "最近一次版本送件,資訊有誤(送件項目資訊有誤)，請連絡系統管理者!")
                Return
        End Select

        'If drOF1 Is Nothing AndAlso drOF2 Is Nothing Then
        '    Common.MessageBox(Me, String.Concat("選擇老師(最近一次版本送件,查無資料)!", vTECHID))
        '    Return
        'End If

        If drOF1 IsNot Nothing Then
            '班級變更申請-
            fg_OK = FILE_COPY1_TT_OF1(drOF1, vMTYPE, vTECHID)
            If Not fg_OK Then Return
        ElseIf drOF2 IsNot Nothing AndAlso oUploadPath2 <> "" Then
            '線上申辦-'師資助教基本資料表/師資學經歷證書影本
            fg_OK = FILE_COPY1_TT_OF2(drOF2, vMTYPE, vTECHID, oUploadPath2)
            If Not fg_OK Then Return
        End If
        '重新查詢
        If fg_OK Then SHOW_Detail_REVISESUB()
    End Sub

    ''' <summary>班級變更申請-最近一次版本送件／下載</summary>
    ''' <param name="drOF"></param>
    ''' <param name="vMTYPE"></param>
    ''' <param name="vTECHID"></param>
    ''' <returns></returns>
    Private Function FILE_COPY1_TT_OF1(drOF As DataRow, vMTYPE As String, vTECHID As String) As Boolean
        Dim oYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim oPLANID As String = Convert.ToString(drOF("PLANID"))
        Dim oCOMIDNO As String = Convert.ToString(drOF("COMIDNO"))
        Dim oSEQNO As String = Convert.ToString(drOF("SEQNO"))
        Dim oSCDateNT As String = TIMS.Cdate3(drOF("CDate"), "yyyyMMdd")
        Dim oSUBNO As String = Convert.ToString(drOF("SUBSEQNO"))
        Dim oFILENAME1 As String = Convert.ToString(drOF("FILENAME1"))
        Dim oFILEPATH1 As String = Convert.ToString(drOF("FILEPATH1"))
        Dim oSRCFILENAME1 As String = Convert.ToString(drOF("SRCFILENAME1"))
        Dim oUploadPath As String = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(oYEARS, oPLANID, oCOMIDNO, oSEQNO, oSCDateNT, oSUBNO))
        Dim s_FilePath1 As String = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))

        If vMTYPE = cst_MTYPE_LATEST_DOWN1 Then
            If Not IO.File.Exists(s_FilePath1) Then
                Common.MessageBox(Me, String.Concat("(班級變更申請)最近一次版本送件,資訊為空,查無有效下載的檔案!", oFILENAME1))
                Return False
            End If
            'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Call ResponsePDFFile1(Me, objconn, Path.Combine(oUploadPath, oFILENAME1))
            Return False
        End If

        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)
        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, TIMS.Cdate3(rSCDate, "yyyyMMdd"), iSubSeqNO)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PR_T(vRID, vRVSID, rAltDataID, vTECHID, "pdf")
        'Dim vFILEPATH1 As String = ""
        Dim vSRCFILENAME1 As String = oSRCFILENAME1
        '上傳檔案/存檔：檔名
        Try
            Call TIMS.MyCreateDir(Me, vUploadPath)
            Call TIMS.MyFileDelete(Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)))
            IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
            '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            strErrmsg &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            'strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            'strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Return False 'Exit Sub
        End Try

        Dim rPMS As New Hashtable
        rPMS.Add("RVSID", vRVSID)
        rPMS.Add("TECHID", vTECHID)
        rPMS.Add("FILENAME1", vFILENAME1)
        rPMS.Add("FILEPATH1", vUploadPath)
        rPMS.Add("SRCFILENAME1", vSRCFILENAME1)
        'rPMS.Add("WAIVED", vWAIVED)
        rPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
        Call SAVE_PLAN_REVISESUBFL_TT(rPMS)
        Return True
    End Function

    Private Function FILE_COPY1_TT_OF2(drOF2 As DataRow, vMTYPE As String, vTECHID As String, oUploadPath As String) As Boolean
        'Dim oYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        'Dim oPLANID As String = Convert.ToString(drOF("PLANID"))
        'Dim oCOMIDNO As String = Convert.ToString(drOF("COMIDNO"))
        'Dim oSEQNO As String = Convert.ToString(drOF("SEQNO"))
        'Dim oSCDateNT As String = TIMS.cdate3(drOF("CDate"), "yyyyMMdd")
        'Dim oSUBNO As String = Convert.ToString(drOF("SUBSEQNO"))
        Dim oFILENAME1 As String = Convert.ToString(drOF2("FILENAME1"))
        Dim oSRCFILENAME1 As String = Convert.ToString(drOF2("SRCFILENAME1"))
        'Dim oFILEPATH1 As String = Convert.ToString(drOF2("FILEPATH1"))
        'Dim oUploadPath As String = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(oYEARS, oPLANID, oCOMIDNO, oSEQNO, oSCDateNT, oSUBNO))
        Dim s_FilePath1 As String = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))

        If vMTYPE = cst_MTYPE_LATEST_DOWN1 Then
            If Not IO.File.Exists(s_FilePath1) Then
                Common.MessageBox(Me, String.Concat("(線上申辦)最近一次版本送件,資訊為空,查無有效下載的檔案!!", oFILENAME1))
                Return False
            End If
            'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Call ResponsePDFFile1(Me, objconn, Path.Combine(oUploadPath, oFILENAME1))
            Return True
        End If

        Dim vRVSID As String = TIMS.ClearSQM(Hid_RVSID.Value)
        Dim vRID As String = TIMS.ClearSQM(RIDValue.Value)
        Dim vYEARS As String = TIMS.ClearSQM(Hid_PlanYear.Value)
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNo, TIMS.Cdate3(rSCDate, "yyyyMMdd"), iSubSeqNO)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PR_T(vRID, vRVSID, rAltDataID, vTECHID, "pdf")
        'Dim vFILEPATH1 As String = ""
        Dim vSRCFILENAME1 As String = oSRCFILENAME1
        '上傳檔案/存檔：檔名
        Try
            Call TIMS.MyCreateDir(Me, vUploadPath)
            Call TIMS.MyFileDelete(Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)))
            IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
            '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            strErrmsg &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            'strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            'strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Return False 'Exit Sub
        End Try

        Dim rPMS As New Hashtable
        rPMS.Add("RVSID", vRVSID)
        rPMS.Add("TECHID", vTECHID)
        rPMS.Add("FILENAME1", vFILENAME1)
        rPMS.Add("FILEPATH1", vUploadPath)
        rPMS.Add("SRCFILENAME1", vSRCFILENAME1)
        'rPMS.Add("WAIVED", vWAIVED)
        rPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
        Call SAVE_PLAN_REVISESUBFL_TT(rPMS)
        Return True
    End Function

    ''' <summary>PLAN_REVISESUBFL_TT 尋找複製資料 </summary>
    ''' <param name="srPMS"></param>
    ''' <returns></returns>
    Private Function GET_OLDFILE1_PLAN_REVISESUBFL_TT_row(srPMS As Hashtable) As DataRow
        Dim vTECHID As String = TIMS.GetMyValue2(srPMS, "TECHID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(srPMS, "COMIDNO")
        Dim vRVSID As String = TIMS.GetMyValue2(srPMS, "RVSID")
        Dim vALTDATAID As String = TIMS.GetMyValue2(srPMS, "ALTDATAID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(srPMS, "ORGKINDGW")
        Dim vBVFID As String = TIMS.GetMyValue2(srPMS, "BVFID")

        Dim rPMS As New Hashtable
        rPMS.Add("TECHID", vTECHID)
        rPMS.Add("COMIDNO", vCOMIDNO)
        rPMS.Add("RVSID", vRVSID)
        rPMS.Add("ALTDATAID", vALTDATAID)
        rPMS.Add("ORGKINDGW", vORGKINDGW)
        Dim sSql As String = ""
        sSql &= " SELECT a.BVFTID,a.BVFID,a.RVSID" & vbCrLf
        sSql &= " ,a.TECHID,a.FILENAME1,a.SRCFILENAME1,a.FILEPATH1" & vbCrLf
        sSql &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= " ,b.PLANID,b.COMIDNO,b.SEQNO,b.CDATE,b.SUBSEQNO,b.ALTDATAID" & vbCrLf
        sSql &= " FROM PLAN_REVISESUBFL_TT a" & vbCrLf
        sSql &= " JOIN PLAN_REVISESUBFL b on b.BVFID=a.BVFID" & vbCrLf
        sSql &= " JOIN KEY_REVISESUB k on k.RVSID=a.RVSID" & vbCrLf
        sSql &= " WHERE a.TECHID=@TECHID AND b.COMIDNO=@COMIDNO" & vbCrLf
        sSql &= " AND b.RVSID=@RVSID AND b.ALTDATAID=@ALTDATAID AND b.ORGKINDGW=@ORGKINDGW" & vbCrLf
        If vBVFID <> "" Then
            sSql &= " AND b.BVFID!=@BVFID" & vbCrLf '(非同筆資料)
            rPMS.Add("BVFID", vBVFID)
        End If
        '(非今天的資料查詢)
        'sSql &= " AND DATEDIFF(DAY,MODIFYDATE,GETDATE())>0" & vbCrLf
        sSql &= " ORDER BY a.BVFTID DESC, a.MODIFYDATE DESC" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sSql, objconn, rPMS)
        Return dr
    End Function

    ''' <summary>最近一次版本送件-pdf下載</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="full_zipFileName"></param>
    Public Shared Sub ResponsePDFFile1(ByRef MyPage As Page, ByRef oConn As SqlConnection, ByVal full_zipFileName As String)
        If full_zipFileName = "" Then Return
        'Dim full_zipFileName As String = Path.Combine(oUploadPath, oFILENAME1)
        With MyPage
            Dim File As New FileInfo(.Server.MapPath(full_zipFileName))
            ' Clear the content of the response
            .Response.ClearContent()
            ' LINE1 Add the file name And attachment, which will force the open/cance/save dialog To show, to the header
            .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
            'Response.Headers["Content-Disposition"] = "attachment; filename=" + zipFileName;
            ' Add the file size into the response header
            .Response.AddHeader("Content-Length", File.Length.ToString())
            ' Set the ContentType
            .Response.ContentType = "application/pdf"
            .Response.TransmitFile(File.FullName)
            ' End the response
            TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
        End With
    End Sub

    ''' <summary>最近一次版本送件-下載</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_latestDown1_Click(sender As Object, e As EventArgs) Handles bt_latestDown1.Click
        UTL_LATEST(cst_MTYPE_LATEST_DOWN1) '最近一次版本送件-下載
    End Sub

End Class
