Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Partial Class TC_06_001_chk
    Inherits AuthBasePage

    'Private Shared ReadOnly download_lock As New Object
    '還原審核不通過的資料如下：
    'UPDATE PLAN_REVISE SET ReviseStatus=NULL,REVISEDATE=NULL,Reason=NULL,Verifier=NULL WHERE 1!=1
    'UPDATE PLAN_REVISE  'SET ReviseStatus=NULL,REVISEDATE=NULL,Reason=NULL,Verifier=NULL 
    'WHERE 1=1 And CONCAT(PLANID,'x',COMIDNO,'x',SEQNO)='5037x41173271x5' and altdataid='22' and ReviseStatus is not null
    Dim s_TransType As String = TIMS.cst_TRANS_LOG_Update 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
    Dim s_TargetTable As String = "CLASS_CLASSINFO"
    Dim s_FuncPath As String = "/TC/06/TC_06_001_chk"
    Dim s_WHERE As String = "" 'insert省略/update必要'String.Format(cst_fWHERE, pkVALUE)

    Const cst_DATA_IS_EMPTY_1 As String = "無"
    Const cst_now As String = "now" 'sType@now新版課表(產投)
    Const cst_old1 As String = "old1" 'sType@old1舊1課表(產投)

    '技檢訓練時數 '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_t1 As String="技檢訓練時數,目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時可儲存，若不符合上述條件，該資料不會存入資料庫。"
    '2.目前僅訓練業別為【[03-01]傳統及民俗復健整復課程】時需要填寫，但是當尚未儲存時應該還無法卡控。正式儲存時，檢核若為03-01才存欄位，否清空。
    'Const cst_EHour_Use_TMID As String="672"
    Const cst_DG3_EHour_技檢訓練時數_iCOL As Integer = 4
    Const cst_DG4_EHour_技檢訓練時數_iCOL As Integer = 4

    Dim vActType As String = ""
    Const cst_dgAct_View1 As String = "View1"
    Const cst_dgAct_Updat1 As String = "Updat1"
    Const cst_dgAct_Edit1 As String = "Edit1"
    'Const cst_dgAct_Delete1 AS String="Delete1"

    Dim dtR2G As DataTable = Nothing
    Dim rPlanID As String = ""
    Dim rComIDNO As String = ""
    Dim rSeqNO As String = ""
    Dim rSCDate As String = ""
    Dim iSubSeqNO As Integer = 0

    'Dim gsAltDataID AS String=""
    'Dim chgName AS Array
    Dim ChgItemName As String()

    Dim i_gSeqno As Integer = 0 '共用序號使用
    'Dim sWOScript1 As String="" '共用JS OPEN語法

    'Const vs_ClassID AS String="vsClassID"
    Const vs_OCID As String = "vsOCID"
    'Const vs_UpdateItem15 As String="UpdateItem15"
    Const vs_PTDRID As String = "PTDRID"
    Const cst_NNN As String = "NNN"

    'UPDATE PLAN_PLANINFO -(SAVE)
    'Const Cst_sChkmode As String="1,4,6,7,8,9,10,12,14,19,21,22" '功能 (計畫用功能)
    'Cst_sChkmode: (sChkmode_AltDataID)
    '1:'訓練期間(開、結訓日期),''2:訓練時段,''3:訓練課程地點,'4:課程編配,''5:訓練師資,'6:班別,'7:期別,
    '8:上課地址,'9:停辦,'10:上課時段,''11:師資,''20:助教,'12:招生人數,''13:增班,'14:上課地點 -學(術)科場地,''15:上課時間,'19:包班種類,
    Const Cst_i訓練期間 As Integer = 1
    Const Cst_i訓練時段 As Integer = 2
    Const Cst_i訓練地點 As Integer = 3
    Const Cst_i課程編配 As Integer = 4
    Const Cst_i訓練師資 As Integer = 5
    Const Cst_i班別名稱 As Integer = 6
    Const Cst_i期別 As Integer = 7
    Const Cst_i上課地址 As Integer = 8
    Const Cst_i停辦 As Integer = 9
    Const Cst_i上課時段 As Integer = 10
    Const Cst_i核定人數 As Integer = 12  'Cst_招生人數  as Integer=12
    Const Cst_i包班種類 As Integer = 19  '20111208 BY AMU 

    Const Cst_i師資 As Integer = 11
    Const Cst_i助教 As Integer = 20  '20120213 BY AMU (產投用助教)
    Const Cst_i科場地 As Integer = 14
    Const Cst_i上課時間 As Integer = 15
    Const Cst_i其他 As Integer = 16
    Const Cst_i報名日期 As Integer = 17  '20080825 andy  add 報名日期
    Const Cst_i課程表 As Integer = 18  '20080626 andy add 課程表
    Const Cst_i訓練費用 As Integer = 21  '20170908 (職前)
    Const Cst_i遠距教學 As Integer = 22  '2021/06/09'增修需求 OJT-21060201 產投 - 班級變更申請/審核：新增遠距教學變更 + 網站-顯示遠距教學資訊 DISTANCE learning /distance teaching
    Const Cst_變更項目總數 As Integer = 22

    'Dim oTest_flag AS Boolean=False '(正式)
    'If TIMS.sUtl_ChkTest() Then oTest_flag=True '測試
    'If Not oTest_flag Then '(正式)
    'If oTest_flag Then '(測試)(正式)

    '/**TIMS**/ CLASS_SCHEDULE , Course_CourseInfo, Teach_TeacherInfo
    '/**產投 無排課作業**/ CLASS_SCHEDULE , Course_CourseInfo
    '/**產投**/  Teach_TeacherInfo
    'Private Property objrow AS DataRow
    Dim ff33 As String = ""
    Dim strTMP1 As String = ""
    Dim dt_Key_CostItem As DataTable = Nothing

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        Call Utl_EveryCreate1()

        If Not IsPostBack Then
            Call CCreate1()
            '送出鍵-不顯示-(儲存)-SaveData1-But_Sub
            But_Sub.Attributes.Add("onclick", "return fnChkVerify();") 'add 送出檢核(必填)
            btnDelete.Attributes("onclick") = TIMS.cst_confirm_delmsg1
        End If

    End Sub

    ''' <summary>'每次重載執行 </summary>
    Sub Utl_EveryCreate1()
        'If TIMS.sUtl_ChkTest() Then oTest_flag=True '測試
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        'trPACKAGE_DOWNLOAD1.Visible=(hid_USE_PLAN_REVISESUB.Value="Y")

        '將變更項目的顯示字串，使用陣列管理，如果需要依不同條件套不同名稱的話，可以直接在這邊修改
        '產學訓套用的顯示字串  / '非產學訓套用的顯示字串
        ChgItemName = If(TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1, TIMS.TPlanID28ChgItemName, TIMS.TPlanIDChgItemName)

        vActType = TIMS.ClearSQM(Convert.ToString(Request("act")))
        dt_Key_CostItem = TIMS.GET_KEY_COSTITEMdt1(objconn)
    End Sub

    '第1次載入計畫變更資料。
    Sub CCreate1()
        ddlISPASS4 = TIMS.GET_ddlISPASSCNT_N(ddlISPASS4)
        'OJT-20231124:班級變更申請-線上送件 ONLINESENDSTATUS NULL/Y:已送出
        hid_USE_PLAN_REVISESUB.Value = TIMS.Utl_GetConfigVAL(objconn, "USE_PLAN_REVISESUB")
        trPACKAGE_DOWNLOAD1.Visible = (hid_USE_PLAN_REVISESUB.Value = "Y")

        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim uUrl1 As String = "TC/06/TC_06_001.aspx?ID=" & RqID
        'vActType=TIMS.ClearSQM(Convert.ToString(Request("act")))
        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(TIMS.ClearSQM(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Request("SubNo"))), 0)

        Dim drCls As DataRow = TIMS.Get_ClassDaRow(rPlanID, rComIDNO, rSeqNO, objconn)
        ViewState(vs_OCID) = If(drCls IsNot Nothing, Convert.ToString(drCls("OCID")), "")

        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNO = "" OrElse rSCDate = "" OrElse iSubSeqNO = 0 Then
            Dim sMsg As String = "查無該計畫變更資料，請重新查詢!!"
            Call TIMS.BlockAlert(Me, sMsg, uUrl1)
            Return 'Exit Sub
        End If
        Dim drPCS As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)
        If drPCS Is Nothing Then
            Dim sMsg As String = "查無該計畫資料，請重新查詢!!"
            Common.MessageBox(Page, sMsg)
            Return 'Exit Sub
        End If
        Dim objrow As DataRow = Get_REVISE1()
        If objrow Is Nothing Then
            Dim sMsg As String = "查無該計畫變更資料，請重新查詢!!"
            Call TIMS.BlockAlert(Me, sMsg, uUrl1)
            Return 'Exit Sub
        End If

        BTN_PACKAGE_DOWNLOAD1.Enabled = If(Convert.ToString(objrow("ONLINESENDSTATUS")) <> "Y", False, True)
        TIMS.Tooltip(BTN_PACKAGE_DOWNLOAD1, If(BTN_PACKAGE_DOWNLOAD1.Enabled, "", "上傳檔案未送出!"), True)

        'objdr=objtable.Rows(0)
        hid_TMID.Value = Convert.ToString(objrow("TMID"))
        hid_AltDataID.Value = Convert.ToString(objrow("AltDataID"))

        hid_OldData2_2.Value = TIMS.NullToStr(objrow("OldData2_2"))
        hid_OldData2_3.Value = TIMS.NullToStr(objrow("OldData2_3"))
        hid_NewData2_2.Value = TIMS.NullToStr(objrow("NewData2_2"))
        hid_NewData2_3.Value = TIMS.NullToStr(objrow("NewData2_3"))
        hid_OldData3_3.Value = TIMS.NullToStr(objrow("OldData3_3"))
        hid_NewData3_1.Value = TIMS.NullToStr(objrow("NewData3_1"))
        hid_OldData5_3.Value = TIMS.NullToStr(objrow("OldData5_3"))
        hid_NewData5_1.Value = TIMS.NullToStr(objrow("NewData5_1"))

        '依 vActType  顯示按鈕
        ChkMode.Enabled = True
        ReviseCont.Enabled = True
        But_Sub.Visible = True '送出鍵-不顯示-(儲存)-SaveData1-But_Sub
        btnDelete.Visible = True
        Btn_SAVE2.Visible = False '儲存-不顯示
        Select Case vActType
            Case cst_dgAct_Edit1
                'But_Sub.Visible=True
            Case cst_dgAct_Updat1
                Btn_SAVE2.Visible = True '啟動儲存
                Common.SetListItem(ChkMode, Convert.ToString(objrow("REVISESTATUS")))
                ChkMode.Enabled = False
                ReviseCont.Text = Convert.ToString(objrow("Reason"))
                ReviseCont.Enabled = False
                '送出鍵不顯示
                But_Sub.Visible = False
                '刪除鍵不顯示
                btnDelete.Visible = False
            Case cst_dgAct_View1
                Common.SetListItem(ChkMode, Convert.ToString(objrow("REVISESTATUS")))
                ChkMode.Enabled = False
                ReviseCont.Text = Convert.ToString(objrow("Reason"))
                ReviseCont.Enabled = False
                '送出鍵不顯示
                But_Sub.Visible = False
                '刪除鍵不顯示
                btnDelete.Visible = False
            Case Else
                Common.SetListItem(ChkMode, Convert.ToString(objrow("REVISESTATUS")))
                ChkMode.Enabled = False
                ReviseCont.Text = Convert.ToString(objrow("Reason"))
                ReviseCont.Enabled = False
                '送出鍵不顯示
                But_Sub.Visible = False
                '刪除鍵不顯示
                btnDelete.Visible = False
        End Select

        'ViewState(vs_OCID)=objrow("OCID")
        OrgName.Text = TIMS.ClearSQM(objrow("OrgName"))
        RIDValue.Value = TIMS.ClearSQM(objrow("RID"))
        ContactName.Text = TIMS.ClearSQM(objrow("ContactName"))
        YearList.Text = TIMS.ClearSQM(objrow("PlanYear"))
        '申請階段 '1：上半年、2：下半年、3：政策性產業 /4:進階政策性產業
        Dim s_APPSTAGE2_NM2 As String = If(Convert.ToString(objrow("APPSTAGE")) <> "", TIMS.GET_APPSTAGE2_NM2(Convert.ToString(objrow("APPSTAGE"))), "")
        labAPPSTAGE.Text = If(s_APPSTAGE2_NM2 <> "", String.Concat("(", s_APPSTAGE2_NM2, ")"), "")

        TrainText.Text = TIMS.ClearSQM(objrow("TrainName"))

        'If Convert.ToString(objrow("Cjob_NO")).Trim <> "" Then CjobNO.Text="[" & objrow("Cjob_NO") & "]"
        'If Convert.ToString(objrow("Cjob_Name")).Trim <> "" Then CjobName.Text=objrow("Cjob_Name")
        Dim dtSCJOB As DataTable = TIMS.Get_SHARECJOBdt(Me, objconn)
        CjobName.Text = TIMS.Get_CJOBNAME(dtSCJOB, Convert.ToString(objrow("CJOB_UNKEY")))
        'CjobNO.Text=objrow("Cjob_NO")
        'CjobName.Text=objrow("Cjob_Name")
        ApplyDate.Text = TIMS.Cdate3(objrow("CDate"))
        lbONLINESENDDATE.Text = Convert.ToString(objrow("lbONLINESENDDATE"))
        If (lbONLINESENDDATE.Text = "") Then lbONLINESENDDATE.Text = cst_DATA_IS_EMPTY_1

        Hid_stdate.Value = TIMS.Cdate3(objrow("STDate"))
        Hid_stdate_7.Value = TIMS.Cdate3(objrow("STDate_7"))
        Hid_ApplyDate.Value = TIMS.Cdate3(objrow("CDate"))

        ClassCName.Text = TIMS.GET_CLASSNAME(Convert.ToString(objrow("ClassName")), Convert.ToString(objrow("CyclType")))

        'OJT-22022301：產投 - 班級變更審核：增加【報名開始日期】、【報名結束日期】 分署提案 ‘委訓單位無法觀察
        'tr_SET_EnterDate.Visible=fg_can_SET_EnterDate
        Dim fg_can_SET_EnterDate As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID < 2 AndAlso Val(objrow("AltDataID")) = Cst_i訓練期間, True, False)
        tr_SET_EnterDate.Visible = fg_can_SET_EnterDate
        Dim fg_ENABLE_EnterDate As Boolean = If(vActType = cst_dgAct_Edit1, fg_can_SET_EnterDate, False)
        If fg_can_SET_EnterDate Then ENABLE_TR_SET_ENTERDATE(fg_ENABLE_EnterDate)

        '二、 當變更項目為「開結訓日期」，因開結訓日期的改變會連帶影響報名起迄時間，
        ' 故分署提案增加【報名開始日期】、【報名結束日期】兩欄位，位置如下圖二
        ' 系統依訓練單位申請變更後之開結訓日期，自動帶出新的報名開始日期/時間及報名結束日期/時間，並可讓分署手動修改。
        ' 報名起迄日期邏輯(依原邏輯)：
        ' 1、未開放報名 【報名開始日期】：開訓日前1個月(30天)。/【報名結束日期】：開訓日前3天。
        ' 2、已開放報名 【報名開始日期】：不變 / 【報名結束日期】：開訓日前3天。
        ' 當變更審核通過後，系統自動將儲存之設定回帶至開班資料查詢功能之【報名開始日期】、【報名結束日期】。
        'tr_SET_EnterDate'SEnterDate'sp_imgSEnterDate'SEnterDate_HR'SEnterDate_MI'FEnterDate'sp_imgFEnterDate'FEnterDate_HR'FEnterDate_MI

        Dim vSEnterDateOrg1 As String = "" '報名開始日期'
        '上架日期
        Hid_OnShellDate.Value = TIMS.Cdate3(objrow("OnShellDate"))
        Call TIMS.SUB_SET_HR_MI(SEnterDate_HR, SEnterDate_MI)
        If Not IsDBNull(objrow("SEnterDate")) Then
            SEnterDate.Text = TIMS.Cdate3(objrow("SEnterDate"))
            vSEnterDateOrg1 = SEnterDate.Text
            TIMS.SET_DateHM(CDate(objrow("SEnterDate")), SEnterDate_HR, SEnterDate_MI)
        End If
        Call TIMS.SUB_SET_HR_MI(FEnterDate_HR, FEnterDate_MI)
        If Not IsDBNull(objrow("FEnterDate")) Then
            FEnterDate.Text = TIMS.Cdate3(objrow("FEnterDate"))
            TIMS.SET_DateHM(CDate(objrow("FEnterDate")), FEnterDate_HR, FEnterDate_MI)
        End If

        'OJT-22022301：產投 - 班級變更審核：增加【報名開始日期】、【報名結束日期】 分署提案
        If fg_can_SET_EnterDate Then
            ' 1、未開放報名 【報名開始日期】：開訓日前1個月(30天)。/【報名結束日期】：開訓日前3天。
            ' 2、已開放報名 【報名開始日期】：不變 / 【報名結束日期】：開訓日前3天。
            ' 當變更審核通過後，系統自動將儲存之設定回帶至開班資料查詢功能之【報名開始日期】、【報名結束日期】。 'New_STDate
            Dim vNew_STDate As String = TIMS.Cdate3(objrow("NewData1_1")) '開訓日前1個月(30天)
            Dim vNew_SEnterDate As String = "" '開訓日前1個月(30天)'TIMS.GetMyValue2(htCC, "SEnterDate") 
            Dim vNew_FEnterDate As String = "" '開訓日前3天 'TIMS.GetMyValue2(htCC, "FEnterDate") 'Dim flag_chkSEnDate As Boolean=False 'false:異常
            Call TIMS.ChangeSEnterDate(vNew_STDate, vNew_SEnterDate, vNew_FEnterDate, vSEnterDateOrg1)
            '已開放報名 【報名開始日期】：不變
            Dim flag_EnterED As Boolean = If(Convert.ToString(objrow("EnterED")) = "Y", True, False)
            '報名開始日期
            If Not flag_EnterED AndAlso vNew_SEnterDate <> "" Then SEnterDate.Text = TIMS.Cdate3(vNew_SEnterDate)
            '報名結束日期
            If vNew_FEnterDate <> "" Then FEnterDate.Text = TIMS.Cdate3(vNew_FEnterDate)
        End If

        '顯示-產投-審查計分表 (含使用預設值)
        Call SHOW_STATUS4(objrow)

        '其他一般計畫
        '學分班資料不顯示
        PointYN.Visible = False
        '刪除鍵不顯示
        btnDelete.Visible = False

        trlbD20KNAME.Visible = False
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            trlbD20KNAME.Visible = True
            'dbo.DECODE(dd.AppResult,'Y',dbo.FN_GET_KID20NAME(pp.PlanID,pp.ComIDNO,pp.SeqNo) ,NULL) D20KNAME
            Dim dr2 As DataRow = TIMS.GET_PLANDEPOT(rPlanID, rComIDNO, rSeqNO, objconn)
            Dim s_lbD20KNAME As String = "" 'cst_DATA_IS_EMPTY_1
            Dim s_lbD25KNAME As String = "" 'cst_DATA_IS_EMPTY_1
            If dr2 IsNot Nothing AndAlso (Convert.ToString(dr2("APPRESULT")) = "Y") Then
                s_lbD20KNAME = Convert.ToString(dr2("D20KNAME"))
                s_lbD25KNAME = Convert.ToString(dr2("D25KNAME"))
            End If
            If s_lbD20KNAME <> "" Then lbD20KNAME.Text = s_lbD20KNAME
            If s_lbD25KNAME <> "" Then lbD25KNAME.Text = s_lbD25KNAME
            If lbD20KNAME.Text = "" AndAlso lbD25KNAME.Text = "" Then lbD25KNAME.Text = cst_DATA_IS_EMPTY_1

            '產學訓人才投資方案
            'NewData14_1=TIMS.Get_SciPlaceID(NewData14_1, rComIDNO, 5, "", objconn)
            'NewData14_2=TIMS.Get_TechPlaceID(NewData14_2, rComIDNO, 5, "", objconn)
            'NewData14_3=TIMS.Get_SciPlaceID(NewData14_3, rComIDNO, 5, "", objconn)
            'NewData14_4=TIMS.Get_TechPlaceID(NewData14_4, rComIDNO, 5, "", objconn)
            '20080811  andy  配合申請頁 學(術)科場地均有帶郵遞區號與地址，故新增SciPlaceID_2、TechPlaceID_2、NewData14_1_1、TechPlaceID_2 以使前後顯示資訊一致
            '---------------------------------------------------------------
            NewData14_1 = TIMS.Get_SciPlaceID(NewData14_1, rComIDNO, 4, "", objconn)
            NewData14_2 = TIMS.Get_TechPlaceID(NewData14_2, rComIDNO, 4, "", objconn)
            NewData14_3 = TIMS.Get_SciPlaceID(NewData14_3, rComIDNO, 4, "", objconn)
            NewData14_4 = TIMS.Get_TechPlaceID(NewData14_4, rComIDNO, 4, "", objconn)
            '14:上課地點
            'NewData8_5=TIMS.Get_TechPTID(NewData8_5, rComIDNO, 3, objconn)
            'NewData8_4=TIMS.Get_SciPTID(NewData8_4, rComIDNO, 3, objconn)
            '---------------------------------------------------------------
            '學分班資料-顯示
            PointYN.Visible = True
            PointYN.Text = If(Convert.ToString(objrow("PointYN")) = "Y", "學分班", "非學分班")
            btnDelete.Visible = False
            '判斷是否已審核，若審核則刪除鍵不顯示
            If Convert.ToString(objrow("ReviseStatusX")) = "X" Then btnDelete.Visible = True '(未審核)還沒審核時刪除鍵顯示
        End If

        '依資料
        iSubSeqNO = If(Convert.ToString(objrow("SubSeqNO")) <> "", Val(objrow("SubSeqNO")), 0)
        '依資料
        rSCDate = If(Convert.ToString(objrow("CDate")) <> "", Common.FormatDate(objrow("CDate")), "")

        '20080522 Andy 新增 程課大綱 
        '.Visible=False
        For i As Integer = 1 To Cst_變更項目總數 '(21)
            If i <> 13 Then
                Dim s_ItemN_N As String = ""
                For i2 As Integer = 1 To 3
                    s_ItemN_N = "Item" & i & "_" & i2
                    If FindControl(s_ItemN_N) IsNot Nothing Then FindControl(s_ItemN_N).Visible = False
                Next
            End If
        Next

        Try
            ChgItem.Text = ChgItemName(CInt(objrow("AltDataID")) - 1)
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
            ChgItem.Text = "未定義"
        End Try

        Select Case Val(objrow("AltDataID"))
            Case Cst_i訓練期間 '1
                Item1_1.Visible = True
                Item1_2.Visible = True
                Table_Sign1.Visible = False
                Table_Sign2.Visible = False
                'ChgItem.Text="訓練期間"
                BSDate.Text = TIMS.Cdate3(objrow("OldData1_1"))
                BEDate.Text = TIMS.Cdate3(objrow("OldData1_2"))
                ASDate.Text = TIMS.Cdate3(objrow("NewData1_1"))
                AEDate.Text = TIMS.Cdate3(objrow("NewData1_2"))

                '20081107 andy  add 報名日期
                '-------------------     
                If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '非產投
                    Table_Sign1.Visible = True
                    Table_Sign2.Visible = True

                    Dim obj_Old_SEnterDate2 As Object = If(Not TIMS.IsDBNull2(objrow("OldData17_1")), objrow("OldData17_1"), objrow("SEnterDate"))
                    Dim obj_Old_FEnterDate2 As Object = If(Not TIMS.IsDBNull2(objrow("OldData17_2")), objrow("OldData17_2"), objrow("FEnterDate"))
                    Dim obj_New_SEnterDate2 As Object = If(Not TIMS.IsDBNull2(objrow("NewData17_1")), objrow("NewData17_1"), objrow("SEnterDate"))
                    Dim obj_New_FEnterDate2 As Object = If(Not TIMS.IsDBNull2(objrow("NewData17_2")), objrow("NewData17_2"), objrow("FEnterDate"))

                    Old_SEnterDate2.Text = TIMS.GetDateTime1(obj_Old_SEnterDate2)
                    Old_FEnterDate2.Text = TIMS.GetDateTime1(obj_Old_FEnterDate2)
                    New_SEnterDate2.Text = TIMS.GetDateTime1(obj_New_SEnterDate2)
                    New_FEnterDate2.Text = TIMS.GetDateTime1(obj_New_FEnterDate2)

                    'New_ExamPeriod.SelectedIndex=-1
                    'OLDDATA3_1, NEWDATA3_1 '甄試日期
                    'OLDDATA10_1 NEWDATA10_1 '甄試日期時段[全天、上午、下午]
                    'Dim MyValue AS String=""
                    Old_Examdate.Text = "" '甄試日期
                    If Convert.ToString(objrow("OLDDATA3_1")) <> "" Then Old_Examdate.Text = TIMS.Cdate3(objrow("OLDDATA3_1"))
                    New_Examdate.Text = "" '甄試日期 'TIMS.cdate3(MyValue)
                    If Convert.ToString(objrow("NEWDATA3_1")) <> "" Then New_Examdate.Text = TIMS.Cdate3(objrow("NEWDATA3_1"))

                    HidOld_ExamPeriod.Value = ""
                    If Convert.ToString(objrow("OLDDATA10_1")) <> "" Then
                        HidOld_ExamPeriod.Value = Convert.ToString(objrow("OLDDATA10_1"))
                        Old_ExamPeriod.Text = TIMS.GetExamPeriod(HidOld_ExamPeriod.Value, objconn)
                    End If
                    HidNew_ExamPeriod.Value = ""
                    If Convert.ToString(objrow("NEWDATA10_1")) <> "" Then
                        HidNew_ExamPeriod.Value = Convert.ToString(objrow("NEWDATA10_1"))
                        New_ExamPeriod.Text = TIMS.GetExamPeriod(HidNew_ExamPeriod.Value, objconn)
                    End If

                    Old_CheckInDate.Text = ""
                    New_CheckInDate.Text = "" ' TIMS.cdate3(MyValue)
                    If Convert.ToString(objrow("OLDDATA2_1")) <> "" Then Old_CheckInDate.Text = TIMS.Cdate3(objrow("OLDDATA2_1"))
                    If Convert.ToString(objrow("NEWDATA2_1")) <> "" Then New_CheckInDate.Text = TIMS.Cdate3(objrow("NEWDATA2_1"))
                End If
            Case Cst_i訓練時段 '2
                '顯示loading
                Item2_1.Visible = True
                Item2_2.Visible = True
                'ChgItem.Text="訓練時段"
                TimeSDate.Text = objrow("OldData2_1")
                TimeEDate.Text = objrow("NewData2_1")
                '20080911 andy edit 
                Call ChgItem2(objrow("OldData2_1"), objrow("OldData2_3"), objrow("NewData2_1"), objrow("NewData2_3"))
                Call ChgItem2_1(objrow("OldData2_3"), objrow("OldData2_2"), objrow("NewData2_3"), objrow("NewData2_2"))
            Case Cst_i訓練地點 '3
                Item3_1.Visible = True
                Item3_2.Visible = True
                'ChgItem.Text="訓練課程地點"
                PlaceDate.Text = objrow("OldData3_1")
                Call ChgItem3(objrow("OldData3_1"), objrow("OldData3_3"))
                EPlace.Text = objrow("NewData3_1")
            Case Cst_i課程編配 '4
                Item4_1.Visible = True
                Item4_2.Visible = True
                'ChgItem.Text="課程編配"
                SSumSci.Text = objrow("OldData4_1").ToString
                SGenSci.Text = objrow("OldData4_2").ToString
                SProSci.Text = objrow("OldData4_3").ToString
                SProTech.Text = objrow("OldData4_4").ToString
                SOther.Text = If(IsNumeric(objrow("OldData4_5")), Val(objrow("OldData4_5")), 0)
                ESumSci.Text = objrow("NewData4_1").ToString
                EGenSci.Text = objrow("NewData4_2").ToString
                EProSci.Text = objrow("NewData4_3").ToString
                EProTech.Text = objrow("NewData4_4").ToString
                EOther.Text = If(IsNumeric(objrow("NewData4_5")), Val(objrow("NewData4_5")), 0)
            Case Cst_i訓練師資 '5
                Item5_1.Visible = True
                Item5_2.Visible = True
                'ChgItem.Text="訓練師資"
                TechDate.Text = objrow("OldData5_1")
                'OLessonTeah1.Text=Get_TeacherName(objrow("NewData5_1"))
                Dim techid As String() = Nothing
                If Convert.ToString(objrow("NewData5_1")) <> "" Then techid = Split(Convert.ToString(objrow("NewData5_1")), ",")
                OLessonTeah1.Text = ""
                OLessonTeah2.Text = ""
                OLessonTeah3.Text = ""
                If Not techid Is Nothing Then
                    If techid.Length > 0 Then OLessonTeah1.Text = TIMS.Get_TeachCName(techid(0), objconn) '1 
                    If techid.Length > 1 Then OLessonTeah2.Text = TIMS.Get_TeachCName(techid(1), objconn) '2
                    If techid.Length > 2 Then OLessonTeah3.Text = TIMS.Get_TeachCName(techid(2), objconn) '2
                End If
                ChgItem5(objrow("OldData5_1"), objrow("OldData5_3"))
            Case Cst_i班別名稱 '6
                Item6_1.Visible = True
                Item6_2.Visible = True
                'ChgItem.Text="班別名稱"
                OClassName.Text = objrow("OldData6_1")
                ChangeOClassName.Text = objrow("NewData6_1")
            Case Cst_i期別 '7
                Item7_1.Visible = True
                Item7_2.Visible = True
                'ChgItem.Text="期別"
                OClassName2.Text = objrow("OldData6_1")
                CyclType.Text = TIMS.FmtCyclType(objrow("OldData7_1"))
                ChangeCyclType.Text = objrow("NewData7_1")
            Case Cst_i上課地址 '8
                Item8_1.Visible = True
                Item8_2.Visible = True
                'ChgItem.Text="上課地址"
                OldData8_1.Value = objrow("OldData8_1").ToString
                OldData8_3.Value = objrow("OldData8_3").ToString
                hidOldData8_6W.Value = TIMS.GetZIPCODE6W(OldData8_1.Value, OldData8_3.Value)
                OldData8_2.Value = objrow("OldData8_2").ToString
                TAddress1.Text = String.Concat("(", OldData8_1.Value, "-", OldData8_3.Value, ")" & TIMS.Get_ZipName(OldData8_1.Value, objconn), OldData8_2.Value)

                NewData8_1.Value = objrow("NewData8_1").ToString
                NewData8_3.Value = objrow("NewData8_3").ToString
                hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
                NewData8_2.Value = objrow("NewData8_2").ToString
                TAddress2.Text = String.Concat("(", NewData8_1.Value, "-", NewData8_3.Value, ")" & TIMS.Get_ZipName(NewData8_1.Value, objconn), NewData8_2.Value)

            Case Cst_i停辦 '9 '停辦
                Item9_1.Visible = True
                Item9_2.Visible = True
                'ChgItem.Text="申請停辦"
                OldData9_1.Value = TIMS.ClearSQM(objrow("OldData9_1"))
                NewData9_1.Value = TIMS.ClearSQM(objrow("NewData9_1")) 'NOTOPEN
                NotOpen1.Text = If(objrow("OldData9_1").ToString = "Y", "停辦中", "開辦中")
                NotOpen2.Text = If(objrow("NewData9_1").ToString = "Y", "申請停辦", "申請開辦")

            Case Cst_i上課時段 '10
                'ChgItem.Text="上課時段"
                Item10_1.Visible = True
                Item10_2.Visible = True
                Dim sql As String = ""
                sql = " SELECT HourRanName FROM Key_HourRan WHERE HRID='" & objrow("OldData10_1") & "' "
                TrainTime1.Text = DbAccess.ExecuteScalar(sql, objconn)
                sql = " SELECT HourRanName FROM Key_HourRan WHERE HRID='" & objrow("NewData10_1") & "' "
                TrainTime2.Text = DbAccess.ExecuteScalar(sql, objconn)
                NewData10_1.Value = objrow("NewData10_1").ToString
            Case Cst_i師資 '11
                'ChgItem.Text="師資"
                Item11_1.Visible = True
                Item11_2.Visible = True
                Item11_3.Visible = True
                TeacherName1.Text = objrow("OldData11_2").ToString
                NewData11_1.Value = objrow("NewData11_1").ToString
                TeacherName1_2.Text = objrow("NewData11_2").ToString
                Hid_NewData11_3.Value = Convert.ToString(objrow("NewData11_3"))

                '產投用選項 'NewData11_1
                Dim htSS As New Hashtable From {{"RID", RIDValue.Value}, {"TECHIDs", NewData11_1.Value}, {"TechTYPE", "A"}}
                SHOW_REVISE_TEACHER3(htSS, objconn)

            Case Cst_i助教 '20
                'ChgItem.Text="助教"
                Item20_1.Visible = True
                Item20_2.Visible = True
                Item20_3.Visible = True
                TeacherName2.Text = objrow("OldData20_2").ToString
                NewData20_1.Value = objrow("NewData20_1").ToString
                TeacherName2_2.Text = objrow("NewData20_2").ToString
                Hid_NewData20_3.Value = Convert.ToString(objrow("NewData20_3"))

                '產投用選項 'NewData20_1
                Dim htSS As New Hashtable From {{"RID", RIDValue.Value}, {"TECHIDs", NewData20_1.Value}, {"TechTYPE", "B"}}
                SHOW_REVISE_TEACHER3(htSS, objconn)

            Case Cst_i核定人數 '12
                'ChgItem.Text="招生人數"
                Item12_1.Visible = True
                Item12_2.Visible = True
                OldData12_1.Text = objrow("OldData12_1").ToString
                NewData12_1.Text = objrow("NewData12_1").ToString

            Case Cst_i科場地 '14 '14:上課地點
                'ChgItem.Text="學(術)科場地"
                Item14_1.Visible = True
                Item14_2.Visible = True

                OldData14_1.Value = GET_OldDataVal(Convert.ToString(objrow("OldData14_1")), Convert.ToString(objrow("NewData14_1")), NewData14_1, SciPlaceID)
                OldData14_2.Value = GET_OldDataVal(Convert.ToString(objrow("OldData14_2")), Convert.ToString(objrow("NewData14_2")), NewData14_2, TechPlaceID)
                OldData14_3.Value = GET_OldDataVal(Convert.ToString(objrow("OldData14_3")), Convert.ToString(objrow("NewData14_3")), NewData14_3, SciPlaceID2)
                OldData14_4.Value = GET_OldDataVal(Convert.ToString(objrow("OldData14_4")), Convert.ToString(objrow("NewData14_4")), NewData14_4, TechPlaceID2)
                '學科場地地址
                Hid_OldData8_4.Value = Convert.ToString(objrow("OldData8_4"))
                '術科場地地址
                Hid_OldData8_5.Value = Convert.ToString(objrow("OldData8_5"))
                '學科場地地址2
                Hid_OldData8_6.Value = Convert.ToString(objrow("OldData8_6"))
                '術科場地地址2
                Hid_OldData8_7.Value = Convert.ToString(objrow("OldData8_7"))
                '學科場地地址
                Hid_NewData8_4.Value = Convert.ToString(objrow("NewData8_4"))
                '術科場地地址
                Hid_NewData8_5.Value = Convert.ToString(objrow("NewData8_5"))
                '學科場地地址2
                Hid_NewData8_6.Value = Convert.ToString(objrow("NewData8_6"))
                '術科場地地址2
                Hid_NewData8_7.Value = Convert.ToString(objrow("NewData8_7"))

                Dim vsTaddressZip As String = ""
                Dim vsTAddress As String = ""
                Dim vsTaddressZIP6W As String = ""
                '學科場地地址
                'If NewData8_4.Items.FindByValue(Convert.ToString(objrow("OldData8_4"))) IsNot Nothing Then AddressSciPTID.Text=NewData8_4.Items.FindByValue(Convert.ToString(objrow("OldData8_4"))).Text
                'Common.SetListItem(NewData8_4, Convert.ToString(objrow("NewData8_4")))
                If Hid_NewData8_4.Value <> "" Then
                    vsTaddressZip = "" : vsTAddress = "" : vsTaddressZIP6W = ""
                    TIMS.GetTaddressPTID(objconn, Hid_NewData8_4.Value, vsTaddressZip, vsTAddress, vsTaddressZIP6W)
                    hid_SP_ZIPCODE.Value = vsTaddressZip
                    hid_SP_ZIP6W.Value = vsTaddressZIP6W
                    hid_SP_ADDRESS.Value = vsTAddress
                End If
                '術科場地地址
                'If NewData8_5.Items.FindByValue(Convert.ToString(objrow("OldData8_5"))) IsNot Nothing Then AddressTechPTID.Text=NewData8_5.Items.FindByValue(Convert.ToString(objrow("OldData8_5"))).Text
                'Common.SetListItem(NewData8_5, Convert.ToString(objrow("NewData8_5")))
                If Hid_NewData8_5.Value <> "" Then
                    vsTaddressZip = "" : vsTAddress = "" : vsTaddressZIP6W = ""
                    TIMS.GetTaddressPTID(objconn, Hid_NewData8_5.Value, vsTaddressZip, vsTAddress, vsTaddressZIP6W)
                    hid_TP_ZIPCODE.Value = vsTaddressZip
                    hid_TP_ZIP6W.Value = vsTaddressZIP6W
                    hid_TP_ADDRESS.Value = vsTAddress
                End If

            Case Cst_i上課時間 '15 '15:上課時間
                'ChgItem.Text="上課時間"
                Item15_1.Visible = True
                Item15_2.Visible = True
                Session("Revise_OnClass") = Nothing
                PlanClassTime()
                ReviseClassTime()

            Case Cst_i其他 '16 '16:其他
                'ChgItem.Text="其他內容"
                Item16_1.Visible = True
                Item16_2.Visible = True
                'OldData16_1.Text=objrow("NewData5_1").ToString
                OldData16_1.Text = objrow("OldData15_1").ToString
                NewData16_1.Text = objrow("NewData15_1").ToString

            Case Cst_i包班種類 '19
                Item19_1.Visible = True
                Item19_2.Visible = True
                Session("Revise_BusPackage") = Nothing
                Call PlanBusPackage()
                Call ReviseBusPackage()
                Dim value1 As String = "未設定"
                Dim value2 As String = "未設定"
                value1 = If(Convert.ToString(objrow("OldData4_1")).Equals("2"), "企業包班", If(Convert.ToString(objrow("OldData4_1")).Equals("3"), "聯合企業包班", "未設定"))
                value2 = If(Convert.ToString(objrow("NewData4_1")).Equals("2"), "企業包班", If(Convert.ToString(objrow("NewData4_1")).Equals("3"), "聯合企業包班", "未設定"))
                hidPackageTypeOld.Value = Convert.ToString(objrow("OldData4_1"))
                hidPackageTypeNew.Value = Convert.ToString(objrow("NewData4_1"))
                PackageTypeOld.Text = value1
                PackageTypeNew.Text = value2

            Case Cst_i報名日期 '17
                Item17_1.Visible = True
                Item17_2.Visible = True
                '20080825 andy  add 報名日期
                Old_SEnterDate.Text = TIMS.GetDateTime1(If(Not TIMS.IsDBNull2(objrow("OldData17_1")), objrow("OldData17_1"), objrow("SEnterDate")))
                Old_FEnterDate.Text = TIMS.GetDateTime1(If(Not TIMS.IsDBNull2(objrow("OldData17_2")), objrow("OldData17_2"), objrow("FEnterDate")))
                New_SEnterDate.Text = TIMS.GetDateTime1(If(Not TIMS.IsDBNull2(objrow("NewData17_1")), objrow("NewData17_1"), objrow("SEnterDate")))
                New_FEnterDate.Text = TIMS.GetDateTime1(If(Not TIMS.IsDBNull2(objrow("NewData17_2")), objrow("NewData17_2"), objrow("FEnterDate")))


            Case Cst_i訓練費用 '21 'Cst_訓練費用
                Item21_1.Visible = True
                Item21_2.Visible = True
                'https://jira.turbotech.com.tw/browse/TIMSC-208
                Dim sPlanKind As String = TIMS.Get_PlanKind(Me, objconn)
                Dim iPlanKind As Integer = Val(sPlanKind)
                Dim iCostMode As Integer = TIMS.GetCostMode(Me, objconn)
                Dim drP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)

                Dim iAdmPercent As Integer = 0
                If Convert.ToString(drP("AdmPercent")) <> "" Then iAdmPercent = Val(drP("AdmPercent"))
                Dim iTaxPercent As Integer = 0
                If Convert.ToString(drP("TaxPercent")) <> "" Then iTaxPercent = Val(drP("TaxPercent"))
                Hid_PlanKind.Value = iPlanKind
                Hid_CostMode.Value = iCostMode
                Hid_AdmPercent.Value = iAdmPercent
                Hid_TaxPercent.Value = iTaxPercent

                'https://jira.turbotech.com.tw/browse/TIMSC-208
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "rPlanID", rPlanID)
                TIMS.SetMyValue(sCmdArg, "rComIDNO", rComIDNO)
                TIMS.SetMyValue(sCmdArg, "rSeqNo", rSeqNO)
                TIMS.SetMyValue(sCmdArg, "rSCDate", rSCDate)
                TIMS.SetMyValue(sCmdArg, "rSubSeqNO", iSubSeqNO)
                Dim iRCID1 As Integer = TIMS.Get_REVC_RCID(sCmdArg, 1, objconn)
                Dim iRCID2 As Integer = TIMS.Get_REVC_RCID(sCmdArg, 2, objconn)
                If iRCID1 = 0 AndAlso iRCID1 = 0 Then Exit Select '無資料離開
                Dim dtR1 As DataTable = TIMS.GET_REVISE_COSTITEMdt(sCmdArg, iRCID1, 1, objconn)
                Call SHOW_REVISE_COSTITEM(dtR1, iRCID1, 1)
                Dim dtR2 As DataTable = TIMS.GET_REVISE_COSTITEMdt(sCmdArg, iRCID2, 2, objconn)
                Call SHOW_REVISE_COSTITEM(dtR2, iRCID2, 2)

            Case Cst_i遠距教學 '22
                Item22_1.Visible = True
                Item22_2.Visible = True
                '遠距教學-old
                Hid_DISTANCE.Value = Convert.ToString(objrow("OldData22_1"))
                lab_DISTANCE.Text = TIMS.GET_DISTANCE_N(0, Convert.ToString(objrow("OldData22_1")))
                '遠距教學-new
                Hid_DISTANCE_new.Value = Convert.ToString(objrow("NewData22_1"))
                lab_DISTANCE_new.Text = TIMS.GET_DISTANCE_N(0, Convert.ToString(objrow("NewData22_1")))

        End Select

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '產投用選項
            LabTMID.Text = "訓練業別"
            '200800714 andy  
            'If CDate(objrow("CDate")) < CDate("2008-10-15") And CInt(objrow("AltDataID"))=15 Then  '2008-10-16前 上課時間-為舊資料暫
            '    Dim dg4_dt As HtmlTableCell=FindControl("dg4_dt")
            '    Dim dg3_dt As HtmlTableCell=FindControl("dg3_dt")
            '    dg4_dt.Visible=False
            '    dg3_dt.Visible=False
            '    ViewState(vs_UpdateItem15)="N"
            'Else
            '    Call CreateTrainDesc()
            '    ViewState(vs_UpdateItem15)="Y"
            'End If
            Call CreateTrainDesc()
            'ViewState(vs_UpdateItem15)="Y"
        Else
            Dim dg4_dt As HtmlTableCell = FindControl("dg4_dt")
            Dim dg3_dt As HtmlTableCell = FindControl("dg3_dt")
            dg4_dt.Visible = False
            dg3_dt.Visible = False
        End If

        changeReason.Text = ""
        Select Case Convert.ToString(objrow("changeReason"))
            Case "1"
                changeReason.Text = "天然災害或政策因素"
            Case "2"
                changeReason.Text = "其他"
        End Select

        If TIMS.IsDBNull2(objrow("ReviseCont")) Then ReviseReason.Text = "" Else ReviseReason.Text = objrow("ReviseCont").ToString 'Add by jack 04/12/30
        If TIMS.IsDBNull2(objrow("Reason")) Then ReviseCont.Text = "" Else ReviseCont.Text = objrow("Reason").ToString
    End Sub

    Private Sub ENABLE_TR_SET_ENTERDATE(ByRef fg_can_Use As Boolean)
        SEnterDate.Enabled = fg_can_Use
        sp_imgSEnterDate.Visible = fg_can_Use
        SEnterDate_HR.Enabled = fg_can_Use
        SEnterDate_MI.Enabled = fg_can_Use
        FEnterDate.Enabled = fg_can_Use
        sp_imgFEnterDate.Visible = fg_can_Use
        FEnterDate_HR.Enabled = fg_can_Use
        FEnterDate_MI.Enabled = fg_can_Use

        Const cst_tt1 As String = "僅供顯示"
        Const cst_tt1c As String = ""
        If Not fg_can_Use Then
            TIMS.Tooltip(SEnterDate, cst_tt1, True)
            TIMS.Tooltip(SEnterDate_HR, cst_tt1, True)
            TIMS.Tooltip(SEnterDate_MI, cst_tt1, True)
            TIMS.Tooltip(FEnterDate, cst_tt1, True)
            TIMS.Tooltip(FEnterDate_HR, cst_tt1, True)
            TIMS.Tooltip(FEnterDate_MI, cst_tt1, True)
        Else
            TIMS.Tooltip(SEnterDate, cst_tt1c, True)
            TIMS.Tooltip(SEnterDate_HR, cst_tt1c, True)
            TIMS.Tooltip(SEnterDate_MI, cst_tt1c, True)
            TIMS.Tooltip(FEnterDate, cst_tt1c, True)
            TIMS.Tooltip(FEnterDate_HR, cst_tt1c, True)
            TIMS.Tooltip(FEnterDate_MI, cst_tt1c, True)
        End If
    End Sub

    ''' <summary> (PLAN_REVISE)取得變更資料一筆 </summary>
    ''' <returns></returns>
    Function Get_REVISE1() As DataRow
        Dim rst As DataRow = Nothing

        Dim objstr As String = ""
        objstr &= " SELECT a.COMIDNO ,a.SEQNO ,a.PLANID" & vbCrLf '/*PK*/
        objstr &= " ,a.CDATE,a.SUBSEQNO,a.ALTDATAID" & vbCrLf
        objstr &= " ,a.ONLINESENDSTATUS,a.ONLINESENDACCT,a.ONLINESENDDATE" & vbCrLf
        objstr &= " ,format(a.ONLINESENDDATE,'yyyy/MM/dd tt hh:mm:ss') lbONLINESENDDATE" & vbCrLf
        objstr &= " ,a.OLDDATA1_1,a.OLDDATA1_2" & vbCrLf
        objstr &= " ,a.NEWDATA1_1,a.NEWDATA1_2" & vbCrLf
        objstr &= " ,a.OLDDATA2_1,a.OLDDATA2_2,a.OLDDATA2_3" & vbCrLf
        objstr &= " ,a.NEWDATA2_1,a.NEWDATA2_2,a.NEWDATA2_3" & vbCrLf
        objstr &= " ,a.OLDDATA3_1,a.OLDDATA3_2,a.OLDDATA3_3" & vbCrLf
        objstr &= " ,a.NEWDATA3_1" & vbCrLf
        objstr &= " ,a.OLDDATA4_1,a.OLDDATA4_2,a.OLDDATA4_3" & vbCrLf
        objstr &= " ,a.NEWDATA4_1,a.NEWDATA4_2,a.NEWDATA4_3,a.NEWDATA4_4" & vbCrLf
        objstr &= " ,a.OLDDATA5_1,a.OLDDATA5_2,a.OLDDATA5_3" & vbCrLf
        objstr &= " ,a.NEWDATA5_1" & vbCrLf

        objstr &= " ,a.REVISEACCT,a.REVISECONT,a.VERIFIER" & vbCrLf
        objstr &= " ,a.REASON,a.REVISESTATUS,a.REVISEDATE" & vbCrLf
        objstr &= " ,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf

        objstr &= " ,a.OLDDATA4_4" & vbCrLf
        objstr &= " ,a.OLDDATA6_1 ,a.NEWDATA6_1" & vbCrLf
        objstr &= " ,a.OLDDATA7_1 ,a.NEWDATA7_1" & vbCrLf
        objstr &= " ,a.OLDDATA8_1,a.OLDDATA8_2" & vbCrLf
        objstr &= " ,a.NEWDATA8_1,a.NEWDATA8_2" & vbCrLf
        objstr &= " ,a.OLDDATA9_1 ,a.NEWDATA9_1" & vbCrLf
        objstr &= " ,a.OLDDATA10_1 ,a.NEWDATA10_1" & vbCrLf

        objstr &= " ,a.OLDDATA12_1 ,a.NEWDATA12_1" & vbCrLf
        objstr &= " ,a.OLDDATA4_5 ,a.NEWDATA4_5" & vbCrLf
        objstr &= " ,a.OLDDATA13_1 ,a.NEWDATA13_1" & vbCrLf
        objstr &= " ,a.OLDDATA14_1 ,a.NEWDATA14_1" & vbCrLf
        objstr &= " ,a.OLDDATA14_2 ,a.NEWDATA14_2" & vbCrLf
        objstr &= " ,a.OLDDATA8_3 ,a.NEWDATA8_3" & vbCrLf
        objstr &= " ,a.OLDDATA15_1 ,a.NEWDATA15_1" & vbCrLf
        objstr &= " ,a.OLDDATA17_1,a.OLDDATA17_2" & vbCrLf
        objstr &= " ,a.NEWDATA17_1,a.NEWDATA17_2" & vbCrLf
        objstr &= " ,a.OLDDATA8_4,a.OLDDATA8_5,a.OLDDATA8_6,a.OLDDATA8_7" & vbCrLf
        objstr &= " ,a.NEWDATA8_4,a.NEWDATA8_5,a.NEWDATA8_6,a.NEWDATA8_7" & vbCrLf
        objstr &= " ,a.OLDDATA14_3,a.OLDDATA14_4" & vbCrLf
        objstr &= " ,a.NEWDATA14_3,a.NEWDATA14_4" & vbCrLf
        objstr &= " ,a.CHANGEREASON" & vbCrLf

        objstr &= " ,a.OLDDATA11_1,a.OLDDATA11_2" & vbCrLf
        objstr &= " ,a.NEWDATA11_1,a.NEWDATA11_2,a.NEWDATA11_3" & vbCrLf
        objstr &= " ,a.OLDDATA20_1,a.OLDDATA20_2" & vbCrLf
        objstr &= " ,a.NEWDATA20_1,a.NEWDATA20_2,a.NEWDATA20_3" & vbCrLf

        objstr &= " ,a.OLDDATA22_1 ,a.NEWDATA22_1" & vbCrLf

        '顯示-產投-審查計分表
        objstr &= " ,a.SENDACCT4,a.SENDDATE4,a.STATUS4" & vbCrLf
        objstr &= " ,a.ISPASS4,a.OVERWEEK4,a.NOINC4" & vbCrLf
        objstr &= " ,a.NODEDUC4" & vbCrLf '政策性課程不扣分

        objstr &= " ,ISNULL(b.PointYN,'N') PointYN" & vbCrLf 'PLAN_PLANINFO
        objstr &= " ,b.PlanYear" & vbCrLf 'PLAN_PLANINFO
        objstr &= " ,b.ClassName" & vbCrLf 'PLAN_PLANINFO
        objstr &= " ,b.CyclType" & vbCrLf 'PLAN_PLANINFO
        objstr &= " ,b.APPSTAGE" & vbCrLf 'PLAN_PLANINFO

        objstr &= " ,ISNULL(i.SEnterDate,b.SEnterDate) SEnterDate" & vbCrLf 'CLASS_CLASSINFO i/ PLAN_PLANINFO b
        objstr &= " ,ISNULL(i.FEnterDate,b.SEnterDate) FEnterDate" & vbCrLf 'CLASS_CLASSINFO i/ PLAN_PLANINFO b
        objstr &= " ,i.ONSHELLDATE" & vbCrLf '上架日期
        ' 報名起迄日期邏輯(依原邏輯)：
        ' 1、未開放報名 【報名開始日期】：開訓日前1個月(30天)。/【報名結束日期】：開訓日前3天。
        ' 2、已開放報名 【報名開始日期】：不變 / 【報名結束日期】：開訓日前3天。
        ' 當變更審核通過後，系統自動將儲存之設定回帶至開班資料查詢功能之【報名開始日期】、【報名結束日期】。
        '已開放報名
        objstr &= " ,CASE WHEN GETDATE()>=ISNULL(i.SEnterDate,b.SEnterDate) THEN 'Y' END EnterED" & vbCrLf

        objstr &= " ,CONVERT(varchar,ISNULL(i.STDate,b.STDate),111) STDate" & vbCrLf
        objstr &= " ,CONVERT(varchar,ISNULL(i.STDate,b.STDate)-7,111) STDate_7" & vbCrLf
        'objstr &=",CONVERT(varchar,ISNULL(i.FTDate,b.FDDate),111) FTDate" & vbCrLf
        objstr &= " ,c.TPlanID,c.PlanName,f.OrgName,b.RID,h.ContactName" & vbCrLf
        objstr &= " ,i.OCID" & vbCrLf 'CLASS_CLASSINFO
        objstr &= " ,i.IsSuccess" & vbCrLf 'CLASS_CLASSINFO
        objstr &= " ,b.CJOB_UNKEY" & vbCrLf
        'objstr &=",s.CJOB_NO" & vbCrLf
        'objstr &=",s.CJOB_NAME" & vbCrLf
        objstr &= " ,ISNULL(d.JobName,d.TrainName) TrainName" & vbCrLf
        objstr &= " ,ISNULL(d.JobName,d.TrainName) JobName" & vbCrLf
        objstr &= " ,b.TMID" & vbCrLf
        objstr &= " ,CASE WHEN a.ReviseStatus IS NULL THEN 'X' ELSE a.ReviseStatus END ReviseStatusX" & vbCrLf
        objstr &= " FROM dbo.PLAN_REVISE a" & vbCrLf
        objstr &= " JOIN dbo.PLAN_PLANINFO b ON a.PlanID=b.PlanID AND a.ComIDNO=b.ComIDNO AND a.SeqNo=b.SeqNo" & vbCrLf
        objstr &= " JOIN dbo.VIEW_PLAN c ON c.PlanID=b.PlanID" & vbCrLf
        objstr &= " JOIN dbo.KEY_TRAINTYPE d ON d.TMID=b.TMID" & vbCrLf
        objstr &= " JOIN dbo.ORG_ORGINFO f ON b.ComIDNO=f.ComIDNO" & vbCrLf
        objstr &= " JOIN dbo.AUTH_RELSHIP g ON g.RID=b.RID" & vbCrLf
        objstr &= " JOIN dbo.ORG_ORGPLANINFO h ON g.RSID=h.RSID" & vbCrLf
        objstr &= " LEFT JOIN dbo.SHARE_CJOB s ON s.CJOB_UNKEY=b.CJOB_UNKEY" & vbCrLf
        objstr &= " LEFT JOIN dbo.CLASS_CLASSINFO i ON a.PlanID=i.PlanID AND a.ComIDNO=i.ComIDNO AND a.SeqNo=i.SeqNo" & vbCrLf 'AND i.IsSuccess='Y'
        objstr &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNo=@SeqNo" & vbCrLf
        objstr &= " AND a.CDate=convert(date,@CDate) AND a.SubSeqNO=@SubSeqNO" & vbCrLf
        Dim htPMS As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}, {"CDate", rSCDate}, {"SubSeqNO", iSubSeqNO}}
        rst = DbAccess.GetOneRow(objstr, objconn, htPMS)
        Return rst
    End Function

    ''' <summary> (PLAN_REVISE)取得變更資料一筆 * </summary>
    ''' <returns></returns>
    Function Get_REVISE2(oConn As SqlConnection) As DataRow
        Dim objrow As DataRow = Nothing 'objTrans=DbAccess.BeginTrans(objconn)
        TIMS.OpenDbConn(oConn)
        Dim objstr As String = ""
        objstr &= " SELECT a.*" & vbCrLf
        objstr &= " FROM PLAN_REVISE a" & vbCrLf
        objstr &= " WHERE a.PlanID=@PlanID AND a.ComIDNO=@ComIDNO AND a.SeqNo=@SeqNo"
        objstr &= " AND a.CDate=@CDate AND a.SubSeqNO=@SubSeqNO" & vbCrLf
        Dim sCmd As New SqlCommand(objstr, oConn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.BigInt).Value = Val(rPlanID)
            .Parameters.Add("ComIDNO", SqlDbType.VarChar).Value = rComIDNO
            .Parameters.Add("SeqNo", SqlDbType.BigInt).Value = Val(rSeqNO)
            .Parameters.Add("CDate", SqlDbType.DateTime).Value = TIMS.Cdate2(rSCDate)
            .Parameters.Add("SubSeqNO", SqlDbType.BigInt).Value = iSubSeqNO
            dt1.Load(.ExecuteReader())
        End With
        'objtable.Load(DbAccess.GetReader(objstr, objconn))
        'Errmsg &="查無該計畫變更資料，請重新查詢" & vbCrLf
        If dt1.Rows.Count = 0 Then Return objrow 'Exit Function

        objrow = dt1.Rows(0)
        Return objrow 'Exit Function
    End Function

    '訓練時段(課程互換)
    Sub ChgItem2(ByVal OldData2_1 As String, ByVal OldData2_3 As String, ByVal NewData2_1 As String, ByVal NewData2_3 As String)
        'OldData2_1@SchoolDate
        'OldData2_3:'--'節次
        'select OldData2_3,COUNT(1) CNT FROM PLAN_REVISE GROUP BY OldData2_3
        'select NewData2_1,COUNT(1) CNT FROM PLAN_REVISE GROUP BY NewData2_1
        'select NewData2_3,COUNT(1) CNT FROM PLAN_REVISE GROUP BY NewData2_3
        'objstr=""
        'objstr &=" select Teacher1,Teacher2,Teacher3,Teacher4,Teacher5,Teacher6,Teacher7,Teacher8,Teacher9,Teacher10,Teacher11,Teacher12" & vbCrLf
        'objstr &=" ,Teacher13,Teacher14,Teacher15,Teacher16,Teacher17,Teacher18,Teacher19,Teacher20,Teacher21,Teacher22,Teacher23,Teacher24" & vbCrLf
        'objstr &=" ,Class1,Class2,Class3,Class4,Class5,Class6,Class7,Class8,Class9,Class10,Class11,Class12" & vbCrLf

        Dim objtable As DataTable = Nothing
        If ViewState(vs_OCID) <> "" Then
            Dim pms1 As New Hashtable From {{"OCID", ViewState(vs_OCID)}}
            Dim objstr As String = ""
            objstr &= " SELECT *" & vbCrLf
            objstr &= " FROM CLASS_SCHEDULE" & vbCrLf
            objstr &= " WHERE OCID=@OCID" & vbCrLf
            objstr &= " AND SchoolDate=" & TIMS.To_date(OldData2_1) & vbCrLf
            objtable = DbAccess.GetDataTable(objstr, objconn, pms1)
        End If

        If TIMS.dtHaveDATA(objtable) Then
            TimeSClass.Items.Clear()
            Dim objrow As DataRow = objtable.Rows(0)
            For i As Integer = 1 To 12
                If TIMS.IsDBNull2(objrow("Class" & i)) Then
                    TimeSClass.Items.Add(New ListItem("第" & i & "節----", i & "^"))
                Else
                    Dim sText As String = ""
                    Dim sValue As String = ""
                    Dim iType As Integer = 0
                    If Not TIMS.IsDBNull2(objrow("Teacher" & i)) Then iType = 1
                    If Not TIMS.IsDBNull2(objrow("Teacher" & i + 12)) Then iType = 2
                    If Not TIMS.IsDBNull2(objrow("Teacher" & i + 24)) Then iType = 3
                    Select Case iType
                        Case 3
                            sText = "第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & TIMS.Get_TeachCName(objrow("Teacher" & i), objconn) & "," & TIMS.Get_TeachCName(objrow("Teacher" & i + 12), objconn) & "," & TIMS.Get_TeachCName(objrow("Teacher" & i + 24), objconn)
                            sValue = i & "^" & objrow("Teacher" & i) & "," & objrow("Teacher" & i + 12) & "," & objrow("Teacher" & i + 24)
                        Case 2
                            sText = "第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & TIMS.Get_TeachCName(objrow("Teacher" & i), objconn) & "," & TIMS.Get_TeachCName(objrow("Teacher" & i + 12), objconn)
                            sValue = i & "^" & objrow("Teacher" & i) & "," & objrow("Teacher" & i + 12)
                        Case 1
                            sText = "第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & TIMS.Get_TeachCName(objrow("Teacher" & i), objconn)
                            sValue = i & "^" & objrow("Teacher" & i)
                    End Select
                    If iType <> 0 AndAlso sText <> "" AndAlso sValue <> "" Then
                        STeacher.Items.Add(New ListItem(sText, sValue))
                    End If
                    'TimeSClass.Items.Add(New ListItem("第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & Get_TeacherName(objrow("Teacher" & i)), i & "^" & objrow("Class" & i)))
                    '200811 andy edit
                    'TimeSClass.Items.Add(New ListItem("第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & Get_TeacherName(objrow("Teacher" & i)) & "," & Get_TeacherName(If(TIMS.IsDBNull2(objrow("Teacher" & i + 12)), "", objrow("Teacher" & i + 12))), i & "^" & objrow("Teacher" & i) & "," & objrow("Teacher" & i + 12)))
                End If
            Next
            TimeSClass.DataBind()

            For Each lItems As ListItem In STeacher.Items
                Dim sValue As String = "" '節次
                If lItems.Value <> "" Then
                    sValue = lItems.Value.Split("^")(0) '節次
                    If sValue <> "" Then
                        If OldData2_3.IndexOf(sValue) > -1 Then lItems.Selected = True
                    End If
                End If
            Next
            TimeSClass.Enabled = False
            'tmpary=Split(OldData2_3, ",")
            'For i=0 To tmpary.Length - 1
            '    TimeSClass.Items(tmpary(i) - 1).Selected=True
            'Next
        End If
    End Sub

    Sub ChgItem2_1(ByVal OldData2_3 As String, ByVal OldData2_2 As String, ByVal NewData2_3 As String, ByVal NewData2_2 As String)
        'select OldData2_3,COUNT(1) CNT FROM PLAN_REVISE GROUP BY OldData2_3 --'節次
        'select OldData2_2,COUNT(1) CNT FROM PLAN_REVISE GROUP BY OldData2_2 --'舊資料(課程)
        'select NewData2_3,COUNT(1) CNT FROM PLAN_REVISE GROUP BY NewData2_3 --'節次
        'select NewData2_2,COUNT(1) CNT FROM PLAN_REVISE GROUP BY NewData2_2 --'舊資料(課程)
        EditSClass.Text = ShowClassList(OldData2_3, OldData2_2, "name")
        EditEClass.Text = ShowClassList(NewData2_3, NewData2_2, "name")
        EditSClassItem.Text = ShowClassList(OldData2_3, OldData2_2, "item")
        EditEClassItem.Text = ShowClassList(NewData2_3, NewData2_2, "item")
    End Sub

    Function ShowClassList(ByVal orderlist As String, ByVal courlist As String, ByVal kind As String) As String
        Dim itemstr As String = ""
        Dim listary As Array = Split(orderlist, ",")
        Dim courary As Array = Split(courlist, ",")
        Select Case kind
            Case "name"
                itemstr = ""
                For i As Integer = 0 To listary.Length - 1
                    If (courary.Length <> listary.Length) And courary.Length = 1 Then '由於舊版程式帶出來的OldData2_2 課程只有一個值無法對應新程式
                        itemstr = itemstr & "(" & CStr(listary(i)) & ") " & CStr(TIMS.Get_CourseName(courary(0), Nothing, objconn)) & "  "
                    Else
                        If (CStr(courary(i)) = "x") Or (Trim(CStr(courary(i))) = "") Then
                            itemstr &= "(" & CStr(listary(i)) & ") " & "未排課" & "  "
                        Else
                            itemstr &= "(" & CStr(listary(i)) & ") " & CStr(TIMS.Get_CourseName(courary(i), Nothing, objconn)) & "  "
                        End If
                    End If
                Next
                itemstr = Left(itemstr, itemstr.Length - 2)       '課程名稱
            Case "item"
                itemstr = ""
                For i As Integer = 0 To listary.Length - 1
                    'itemstr=itemstr & "[第" & CStr(listary(i)) & "節]" & "  "
                    itemstr = itemstr & CStr(listary(i)) & ","
                Next
                itemstr = Left(itemstr, itemstr.Length - 1) '節次
        End Select
        Return itemstr
    End Function

    ''' <summary>'判斷是數字且小於20 (最大節次，只有12節課)</summary>
    ''' <param name="str"></param>
    ''' <returns></returns>
    Function sUtl_xNum(ByVal str As String) As Boolean
        Dim rst As Boolean = False
        'AndAlso sUtl_xNum(SPlace.Items(i).Value) 
        If Not IsNumeric(str) Then Return rst
        rst = (CInt(str) < 20)
        Return rst
    End Function

    Sub ChgItem3(ByVal OldData3_1 As String, ByVal OldData3_3 As String)

        Dim objtable As DataTable = Nothing
        If ViewState(vs_OCID) = "" Then Return

        Dim objstr As String = ""
        objstr &= " SELECT Room1,Room2,Room3,Room4,Room5,Room6,Room7,Room8,Room9,Room10,Room11,Room12"
        objstr &= " ,Class1,Class2,Class3,Class4,Class5,Class6,Class7,Class8,Class9,Class10,Class11,Class12" & vbCrLf
        objstr &= " FROM CLASS_SCHEDULE" & vbCrLf
        objstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
        objstr &= " AND SchoolDate=" & TIMS.To_date(OldData3_1) & vbCrLf
        objtable = DbAccess.GetDataTable(objstr, objconn)
        If TIMS.dtHaveDATA(objtable) Then
            SPlace.Items.Clear()
            Dim objrow As DataRow = objtable.Rows(0)
            For i As Integer = 1 To 12
                If TIMS.IsDBNull2(objrow("Class" & i)) Then
                    SPlace.Items.Add(New ListItem("第" & i & "節----", "^"))
                Else
                    SPlace.Items.Add(New ListItem("第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & objrow("Room" & i), i & "^" & objrow("Room" & i)))
                End If
            Next
            SPlace.DataBind()
            Dim tmpary As Array = Split(OldData3_3, ",")
            For i As Integer = 0 To tmpary.Length - 1
                If sUtl_xNum(tmpary(i)) Then SPlace.Items(tmpary(i) - 1).Selected = True
            Next
            SPlace.Enabled = False
        End If
    End Sub

    Sub ChgItem5(ByVal OldData5_1 As String, ByVal OldData5_3 As String)
        'OldData5_1 @SchoolDate
        'OldData5_3:'節次
        'select OldData5_1,COUNT(1) CNT FROM PLAN_REVISE GROUP BY OldData5_1
        'Dim tmpary AS Array
        'objstr=""
        'objstr &=" select Teacher1,Teacher2,Teacher3,Teacher4,Teacher5,Teacher6,Teacher7,Teacher8,Teacher9,Teacher10,Teacher11,Teacher12" & vbCrLf
        'objstr &=" ,Teacher13,Teacher14,Teacher15,Teacher16,Teacher17,Teacher18,Teacher19,Teacher20,Teacher21,Teacher22,Teacher23,Teacher24" & vbCrLf '2008/11/26  andy  edit
        'objstr &=" ,Class1,Class2,Class3,Class4,Class5,Class6,Class7,Class8,Class9,Class10,Class11,Class12" & vbCrLf
        If ViewState(vs_OCID) = "" Then Return
        Dim objstr As String = ""
        objstr &= " SELECT *" & vbCrLf
        objstr &= " FROM CLASS_SCHEDULE" & vbCrLf
        objstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
        objstr &= " AND SchoolDate=" & TIMS.To_date(OldData5_1) & vbCrLf
        Dim objtable As DataTable = DbAccess.GetDataTable(objstr, objconn)

        If TIMS.dtHaveDATA(objtable) Then
            STeacher.Items.Clear()
            Dim objrow As DataRow = objtable.Rows(0)
            For i As Integer = 1 To 12 '節次
                If TIMS.IsDBNull2(objrow("Class" & i)) Then
                    STeacher.Items.Add(New ListItem("第" & i & "節----", i & "^"))
                Else
                    Dim sText As String = ""
                    Dim sValue As String = ""
                    Dim iType As Integer = 0
                    If Not TIMS.IsDBNull2(objrow("Teacher" & i)) Then iType = 1
                    If Not TIMS.IsDBNull2(objrow("Teacher" & i + 12)) Then iType = 2
                    If Not TIMS.IsDBNull2(objrow("Teacher" & i + 24)) Then iType = 3
                    Select Case iType
                        Case 3
                            sText = "第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & TIMS.Get_TeachCName(objrow("Teacher" & i), objconn) & "," & TIMS.Get_TeachCName(objrow("Teacher" & i + 12), objconn) & "," & TIMS.Get_TeachCName(objrow("Teacher" & i + 24), objconn)
                            sValue = i & "^" & objrow("Teacher" & i) & "," & objrow("Teacher" & i + 12) & "," & objrow("Teacher" & i + 24)
                        Case 2
                            sText = "第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & TIMS.Get_TeachCName(objrow("Teacher" & i), objconn) & "," & TIMS.Get_TeachCName(objrow("Teacher" & i + 12), objconn)
                            sValue = i & "^" & objrow("Teacher" & i) & "," & objrow("Teacher" & i + 12)
                        Case 1
                            sText = "第" & i & "節--" & GetClassName(objrow("Class" & i)) & "--" & TIMS.Get_TeachCName(objrow("Teacher" & i), objconn)
                            sValue = i & "^" & objrow("Teacher" & i)
                    End Select
                    If iType <> 0 AndAlso sText <> "" AndAlso sValue <> "" Then STeacher.Items.Add(New ListItem(sText, sValue))
                End If
            Next
            STeacher.DataBind()

            '--'節次
            'SELECT OldData5_3,COUNT(1) CNT FROM PLAN_REVISE GROUP BY OldData5_3
            'tmpary=Split(OldData5_3, ",")
            For Each lItems As ListItem In STeacher.Items
                Dim sValue As String = "" '節次
                If lItems.Value <> "" Then
                    sValue = lItems.Value.Split("^")(0) '節次
                    If sValue <> "" Then
                        If OldData5_3.IndexOf(sValue) > -1 Then lItems.Selected = True
                    End If
                End If
            Next
            STeacher.Enabled = False
        End If
    End Sub

    Function GetClassName(ByVal ClassID As String) As String
        Dim rst As String = ""
        ClassID = TIMS.ClearSQM(ClassID)
        If ClassID = "" Then Return rst
        'If ClassID="" Then Exit Function
        Dim str As String = "SELECT COURSENAME FROM COURSE_COURSEINFO WHERE COURID=" & ClassID
        rst = Convert.ToString(DbAccess.ExecuteScalar(str, objconn))
        Return rst
    End Function

#Region "(No Use)"
    'Function Get_TeacherName(ByVal TeacherID AS String) AS String
    '    'Dim str AS String="select TeachCName from Teach_TeacherInfo where TechID='" & TeacherID & "'"
    '    TeacherID=Trim(TeacherID)
    '    If Convert.ToString(TeacherID)="" Then Exit Function
    '    Dim str AS String="select TeachCName from Teach_TeacherInfo where TechID in (" & TeacherID & ")"
    '    Return DbAccess.ExecuteScalar(str, objconn)
    'End Function

    'Function update_Plan_TrainDesc(ByVal ChkMode AS Integer)
    '    Dim objTrans AS SqlTransaction
    '    Dim dt1, dt2 AS DataTable
    '    Dim dr1, dr2 AS DataRow
    '    Dim objadapter AS SqlDataAdapter
    '    Dim sql AS String
    '    objTrans=DbAccess.BeginTrans(objConn)
    '    sql="SELECT * FROM Plan_TrainDesc WHERE  0=0" & vbCrLf
    '    sql &=" and  PlanID='" & Request("PlanID") & "'" & vbCrLf
    '    sql &=" and  ComIDNO='" & Request("cid") & "'" & vbCrLf
    '    sql &=" and  SeqNO='" & Request("no") & "'" & vbCrLf
    '    sql &=" ORDER  BY  PTDID" & vbCrLf
    '    dt1=DbAccess.GetDataTable(sql)
    '    If ChkMode=1 Then       '
    '        sql="SELECT  A.PTDID ,  (  case  when  b.NewData1_1  is  not  null  then   b.NewData1_1" & vbCrLf
    '        sql &=" else a.STrainDate  end ) AS   STrainDate  "
    '    ElseIf ChkMode=14 Then  '
    '        sql="SELECT  A.PTDID ,  (  case  when  b.NewData2_1  is  not  null  then   b.NewData2_1" & vbCrLf
    '        sql &=" else a.PTID  end ) AS   PTID  "
    '    ElseIf ChkMode=18 Then  '課程表
    '    End If
    '    sql &=" FROM Plan_TrainDesc  a "
    '    sql &=" LEFT JOIN Plan_TrainDesc_Revise b ON a.PlanID=b.PlanID and  a.PTDID=b.PTDID" & vbCrLf
    '    sql &=" WHERE 0=0" & vbCrLf
    '    sql &=" and  a.PlanID='" & Request("PlanID") & "'" & vbCrLf
    '    sql &=" and  a.ComIDNO='" & Request("cid") & "'" & vbCrLf
    '    sql &=" and  a.SeqNO='" & Request("no") & "'" & vbCrLf
    '    sql &=" and  b.CDate='" & Request("CDate") & "'" & vbCrLf
    '    sql &=" and  b.SubSeqNO='" & Request("SubNo") & "'" & vbCrLf
    '    If ChkMode=1 Then
    '        sql &=" and  b.AltPTDRDataID='1'" & vbCrLf
    '    ElseIf ChkMode=14 Then
    '        sql &=" and  b.AltPTDRDataID='14'" & vbCrLf
    '    End If
    '    sql &=" ORDER  BY  a.PTDID "
    '    dt2=DbAccess.GetDataTable(sql)
    '    Dim j AS Integer
    '    sql=""
    '    For j=0 To dt2.Rows.Count - 1
    '        dr2=dt2.Rows(j)
    '        sql &="  UPDATE PLAN_TRAINDESC" & vbCrLf
    '        If ChkMode=1 Then
    '            sql &=" set  STrainDate='" & dr2("STrainDate") & "' ,ETrainDate='" & dr2("STrainDate") & "'"
    '        ElseIf ChkMode=14 Then
    '            sql &=" set  PTID='" & dr2("PTID") & "'"
    '        End If
    '        sql &=" , ModifyAcct='" & sm.UserInfo.UserID & "',   ModifyDate='" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf
    '        sql &=" WHERE  0=0" & vbCrLf
    '        sql &=" and  PlanID='" & Request("PlanID") & "'" & vbCrLf
    '        sql &=" and  ComIDNO='" & Request("cid") & "'" & vbCrLf
    '        sql &=" and  SeqNO='" & Request("no") & "'" & vbCrLf
    '        sql &=" and  PTDID='" & dr2("PTDID") & "'" & vbCrLf
    '        sql &=" update  Plan_TrainDesc_Revise" & vbCrLf
    '        sql &=" set  ReviseStatus ='Y' , Verifier='" & sm.UserInfo.UserID & "'" & vbCrLf
    '        sql &=" WHERE  0=0" & vbCrLf
    '        sql &=" and  PlanID='" & Request("PlanID") & "'" & vbCrLf
    '        sql &=" and  ComIDNO='" & Request("cid") & "'" & vbCrLf
    '        sql &=" and  SeqNO='" & Request("no") & "'" & vbCrLf
    '        sql &=" and  CDate= '" & Request("CDate") & "'" & vbCrLf
    '        sql &=" and  SubSeqNO='" & Request("SubNo") & "'" & vbCrLf
    '        sql &=" and  PTDID='" & dr2("PTDID") & "'" & vbCrLf
    '        If ChkMode=1 Then
    '            sql &=" and   AltPTDRDataID='1'" & vbCrLf
    '        ElseIf ChkMode=14 Then
    '            sql &=" and   AltPTDRDataID='14'" & vbCrLf
    '        End If
    '    Next
    '    DbAccess.ExecuteNonQuery(sql, objTrans)
    '    DbAccess.UpdateDataTable(dt1, objadapter, objTrans)
    'End Function

    ''組合班級名稱依計畫。
    'Function Get_classname(ByVal oClassName AS String, ByVal oCyclType AS String) AS String
    '    Dim rst AS String=oClassName
    '    oCyclType=Trim(oCyclType)
    '    If oCyclType <> "" AndAlso IsNumeric(oCyclType) Then
    '        rst=oClassName & "第" & oCyclType & "期"
    '    End If
    '    Return rst
    'End Function

    ''取得班級資料 OCID
    'Function Get_ClassDaRow(ByVal PlanID AS String, ByVal ComIDNO AS String, ByVal SeqNO AS String, ByVal tConn AS SqlConnection) AS DataRow
    '    Dim rst AS DataRow=Nothing
    '    Dim objstr AS String=""
    '    objstr=""
    '    objstr &=" SELECT i.OCID"
    '    objstr &=" FROM CLASS_CLASSINFO i"
    '    objstr &=" WHERE 1=1"
    '    objstr &=" AND i.PlanID=" & PlanID
    '    objstr &=" and i.ComIDNO='" & ComIDNO & "'"
    '    objstr &=" and i.SeqNO=" & SeqNO
    '    Dim dtT AS DataTable=DbAccess.GetDataTable(objstr, tConn)
    '    If dtT.Rows.Count=1 Then rst=dtT.Rows(0)
    '    Return rst
    'End Function
#End Region

    ''' <summary> SAVE SERVER端 檢查 </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(Convert.ToString(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Convert.ToString(Request("SubNo")))), 0)

        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNO = "" OrElse rSCDate = "" Then
            Errmsg &= "查無該計畫變更資料，請重新查詢" & vbCrLf
            Return False 'Exit Function
        End If

        'Select Case Val(objrow("AltDataID"))
        'Case Cst_i停辦
        Dim v_ChkMode As String = TIMS.GetListValue(ChkMode)

        '檢核-產投-審查計分表
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '1. 判斷user key的[申請變更函送日期] - 該班第一次申請變更日 <=7 
            '--> 系統在[逾期週數]欄位點選 "無逾期"、[申請變更函送狀態]欄位點選 "依規定辦理"
            '判斷user key的[申請變更函送日期] - 該班第一次申請變更日 > 7  且 <=14 
            '--> 系統在[逾期週數]欄位點選 "1周以內"、[申請變更函送狀態]欄位點選 "逾期(扣分)"
            '判斷user key的[申請變更函送日期] - 該班第一次申請變更日 > 14 
            '--> 系統在[逾期週數]欄位點選 "1周以上"、[申請變更函送狀態]欄位點選 "逾期(扣分)"
            '2.若變更項目為停辦 --> 系統在[逾期週數]欄位點選 "停辦"
            '3.變更項目若為上課時間、其他，系統不用做檢核是否逾期
            '4.這只是在user key的[申請變更函送日期]當下觸發檢核預帶值，[逾期週數]欄位可讓user修改，最後儲存以user調整的為主。

            Dim vSTATUS4 As String = TIMS.GetListValue(STATUS4)
            Dim vddlISPASS4 As String = TIMS.GetListValue(ddlISPASS4)
            Dim vOVERWEEK4 As String = TIMS.GetListValue(OVERWEEK4)
            SENDDATE4.Text = TIMS.Cdate3(SENDDATE4.Text)
            hid_AltDataID.Value = TIMS.ClearSQM(hid_AltDataID.Value)
            Select Case v_ChkMode 'Convert.ToString(ChkMode.SelectedValue) '通過 ／不通過
                Case "Y"
                    '檢核'//9:停辦//15:上課時間//16:其他
                    Select Case hid_AltDataID.Value
                        Case "15", "16" '//9:停辦//15:上課時間//16:其他
                            If SENDDATE4.Text = "" Then
                                Errmsg &= "請輸入-申請變更函送日期" & vbCrLf
                            End If
                            If vSTATUS4 = "" Then
                                Errmsg &= "請選擇-申請變更函送狀態" & vbCrLf
                            End If
                            If vddlISPASS4 = "" Then
                                Errmsg &= "請選擇-資料內容不符合規定(2-1-2)" & vbCrLf
                            End If
                        Case Else
                            'Case "9" '//9:停辦//15:上課時間//16: 其他
                            If SENDDATE4.Text = "" Then
                                Errmsg &= "請輸入-申請變更函送日期" & vbCrLf
                            End If
                            If vSTATUS4 = "" Then
                                Errmsg &= "請選擇-申請變更函送狀態" & vbCrLf
                            End If
                            If vddlISPASS4 = "" Then
                                Errmsg &= "請選擇-資料內容不符合規定(2-1-2)" & vbCrLf
                            End If
                            If vOVERWEEK4 = "" Then '逾期週數
                                Errmsg &= "請選擇-逾期週數" & vbCrLf
                            End If
                    End Select
            End Select
            If Errmsg <> "" Then Return False 'Exit Function
        End If

        'Dim objstr As String=""
        'Dim i_AltDataID As Integer=0 '功能代碼  'Cst_sChkmode: (sChkmode_AltDataID)
        Select Case v_ChkMode 'Convert.ToString(ChkMode.SelectedValue) '通過 ／不通過
            Case "Y"
                Dim objrow As DataRow = Get_REVISE2(objconn)
                If objrow Is Nothing Then
                    Errmsg &= "查無該計畫變更資料，請重新查詢" & vbCrLf
                    Return False 'Exit Function
                End If

                hid_AltDataID.Value = Convert.ToString(objrow("AltDataID"))
                Dim i_AltDataID As Integer = If(Convert.ToString(hid_AltDataID.Value) <> "", Val(objrow("AltDataID")), 0)
                If i_AltDataID = 0 Then
                    'Common.MessageBox(Page, "查無該計畫變更資料，請重新查詢")
                    Errmsg &= "查無該計畫變更資料，請重新查詢" & vbCrLf
                    Return False 'Exit Function
                End If

                Dim objstr As String = ""
                Select Case i_AltDataID
                    Case Cst_i包班種類'19
                    Case Cst_i訓練時段 '2
                        If TimeSDate.Text = TimeEDate.Text Then
                            objstr = ""
                            objstr &= " SELECT * FROM CLASS_SCHEDULE "
                            objstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
                            objstr &= " AND (SchoolDate=" & TIMS.To_date(TimeSDate.Text) & " OR SchoolDate=" & TIMS.To_date(TimeEDate.Text) & ") "
                            Dim sCmd As New SqlCommand(objstr, objconn)
                            TIMS.OpenDbConn(objconn)
                            Dim dtSCH As New DataTable
                            With sCmd
                                .Parameters.Clear()
                                dtSCH.Load(.ExecuteReader())
                            End With
                            '(同日應有1筆資料)
                            If TIMS.dtNODATA(dtSCH) OrElse dtSCH.Rows.Count <> 1 Then
                                Errmsg &= "查詢資料有誤，請重新查詢!!" & vbCrLf
                                Return False 'Exit Function
                            End If
                        Else
                            objstr = ""
                            objstr &= " SELECT * FROM CLASS_SCHEDULE "
                            objstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
                            objstr &= " AND (SchoolDate=" & TIMS.To_date(TimeSDate.Text) & " OR SchoolDate=" & TIMS.To_date(TimeEDate.Text) & ")"
                            Dim sCmd As New SqlCommand(objstr, objconn)
                            TIMS.OpenDbConn(objconn)
                            Dim dtSCH As New DataTable
                            With sCmd
                                .Parameters.Clear()
                                dtSCH.Load(.ExecuteReader())
                            End With
                            '(不同日子應有2筆資料)
                            If TIMS.dtNODATA(dtSCH) OrElse dtSCH.Rows.Count <> 2 Then
                                Errmsg &= "查詢資料有誤，請重新查詢!!" & vbCrLf
                                Return False 'Exit Function
                            End If
                        End If
                End Select

                '(停辦不處分)
                Select Case i_AltDataID
                    Case Cst_i停辦
                    Case Else
                        '登入者檢查
                        Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                        Dim iBlackType As Integer = TIMS.Chk_OrgBlackType(Me, objconn)
                        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
                            Select Case iBlackType
                                Case 1, 2, 3
                                    'Errmsg &="於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                                    'Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
                                    'Exit Function '有錯誤訊息 'Return False '不可儲存
                                    Errmsg &= "於處分日期起的期間，已審核通過的班級若有進行變更申請，變更審核時，一律只能審核為失敗。" & vbCrLf
                                    Return False 'Exit Function
                            End Select
                        End If
                        'Hid_ComIDNO.Value=TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                        'Dim iBlackType AS Integer=TIMS.Chk_OrgBlackType(Me, objconn)

                        '審核機構者檢查
                        Hid_ComIDNO.Value = rComIDNO
                        If TIMS.Check_OrgBlackList2(Me, Hid_ComIDNO.Value, iBlackType, objconn) Then
                            Select Case iBlackType
                                Case 1, 2, 3
                                    'Errmsg &="於處分日期起的期間，已審核通過的班級不可進行轉班作業。"
                                    'Common.MessageBox(Me, "於處分日期起的期間，已審核通過的班級不可進行轉班作業。")
                                    'Exit Function '有錯誤訊息 'Return False '不可儲存
                                    Errmsg &= "於處分日期起的期間，已審核通過的班級若有進行變更申請，變更審核時，一律只能審核為失敗。" & vbCrLf
                                    Return False 'Exit Function
                            End Select
                        End If

                        'If oTest_flag Then '(測試)(正式)
                        '    Errmsg &="於處分日期起的期間，已審核通過的班級若有進行變更申請，變更審核時，一律只能審核為失敗。" & vbCrLf
                        '    Return False 'Exit Function
                        'End If
                End Select

        End Select

        'Errmsg +="電子信箱 EMail格式錯誤。" & vbCrLf
        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    ''' <summary> 檢核 SET_FENTERDATE </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function Chk_SET_FENTERDATE(ByRef Errmsg As String, ByVal iAltDataID As Integer) As Boolean
        Dim rst As Boolean = True
        'OJT-22022301：產投 - 班級變更審核：增加【報名開始日期】、【報名結束日期】 分署提案
        Dim flag_can_SET_EnterDate As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID < 2 AndAlso iAltDataID = Cst_i訓練期間, True, False)
        flag_can_SET_EnterDate = If(vActType = cst_dgAct_Edit1, flag_can_SET_EnterDate, False)
        If Not flag_can_SET_EnterDate Then Return rst

        '報名開始日期'(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
        Dim vsSEnterDate As String = TIMS.GET_YMDHM1(SEnterDate.Text, TIMS.GetListValue(SEnterDate_HR), TIMS.GetListValue(SEnterDate_MI))
        '報名結束日期'(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
        Dim vsFEnterDate As String = TIMS.GET_YMDHM1(FEnterDate.Text, TIMS.GetListValue(FEnterDate_HR), TIMS.GetListValue(FEnterDate_MI))
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "STDATE", ASDate.Text)
        TIMS.SetMyValue2(htSS, "FTDATE", AEDate.Text)
        TIMS.SetMyValue2(htSS, "SEnterDate", vsSEnterDate)
        TIMS.SetMyValue2(htSS, "FEnterDate", vsFEnterDate)
        TIMS.SetMyValue2(htSS, "OnShellDate", Hid_OnShellDate.Value) '上架日期
        Errmsg = TIMS.CHK_DateRange1(htSS)
        rst = If(Errmsg <> "", False, True)
        Return rst
    End Function

    ''' <summary> UPDATE CLASS_STUDENTSOFCLASS Cst_i停辦 </summary>
    ''' <param name="iAltDataID"></param>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    Sub UPDATE_CLASS_STUDENTSOFCLASS_ALT9(ByRef iAltDataID As Integer, ByRef tmpConn As SqlConnection, ByRef tmpTrans As SqlTransaction)
        If iAltDataID <> Cst_i停辦 Then Return
        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNO = "" Then Return
        Dim parms As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}}
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT ip.TPLANID,pp.PLANID,pp.COMIDNO,pp.SEQNO,cc.OCID" & vbCrLf
        sql &= " 	,concat(pp.planid,'x',pp.comidno,'x',pp.seqno) PCS" & vbCrLf
        sql &= " 	,cc.CLASSCNAME,cc.CYCLTYPE" & vbCrLf
        sql &= " 	,pp.ISAPPRPAPER,pp.APPLIEDRESULT,pp.TRANSFLAG" & vbCrLf
        sql &= " 	,cc.ISSUCCESS,cc.ISCLOSED,cc.NOTOPEN" & vbCrLf
        sql &= " 	FROM CLASS_CLASSINFO cc" & vbCrLf
        sql &= " 	JOIN PLAN_PLANINFO pp on pp.PLANID=cc.PLANID and pp.COMIDNO=cc.COMIDNO and pp.SEQNO=cc.SEQNO" & vbCrLf
        sql &= " 	JOIN ID_PLAN ip on ip.PLANID=pp.PLANID" & vbCrLf
        sql &= " 	WHERE pp.ISAPPRPAPER='Y' AND pp.APPLIEDRESULT='Y' AND pp.TRANSFLAG='Y'" & vbCrLf
        sql &= " 	AND cc.ISSUCCESS='Y' AND cc.NOTOPEN='Y'" & vbCrLf
        'sql &=" 	AND ip.YEARS>='2021' AND cc.STDATE<=GETDATE()" & vbCrLf
        sql &= " 	AND pp.PlanID=@PlanID AND pp.ComIDNO=@ComIDNO AND pp.SeqNo=@SeqNo )" & vbCrLf
        sql &= " ,WR1 AS ( SELECT c.OCID, MAX(r.REVISEDATE) REJECTTDATE1,COUNT(1) CNT1" & vbCrLf
        sql &= " 	FROM WC1 c" & vbCrLf
        sql &= " 	join PLAN_REVISE r on r.PLANID=c.PLANID and r.COMIDNO=c.COMIDNO and r.SEQNO=c.SEQNO" & vbCrLf
        sql &= " 	WHERE r.ALTDATAID=9" & vbCrLf ' /*Cst_i停辦:9*/" & vbCrLf
        sql &= " 	AND r.REVISESTATUS='Y'" & vbCrLf
        sql &= " 	GROUP BY c.OCID )" & vbCrLf
        sql &= " UPDATE CLASS_STUDENTSOFCLASS" & vbCrLf
        sql &= " SET REJECTTDATE1=R.REJECTTDATE1,STUDSTATUS=2" & vbCrLf
        sql &= " FROM WR1 r" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS cs on cs.OCID=r.OCID AND cs.STUDSTATUS=1 and cs.REJECTTDATE1 is null" & vbCrLf
        Dim rst As Integer = DbAccess.ExecuteNonQuery(sql, tmpTrans, parms)

        'Dim sCmd As New SqlCommand(sql, tmpConn, tmpTrans)

        'DbAccess.HashParmsChange(sCmd, parms)

        'sCmd.ExecuteNonQuery()
    End Sub

    ''' <summary> UPDATE CLASS_CLASSINFO -(SAVE) </summary>
    ''' <param name="iAltDataID"></param>
    ''' <param name="oTrans"></param>
    Sub UPDATE_CLASSCLASSINFO(ByRef iAltDataID As Integer, ByRef oTrans As SqlTransaction)
        ViewState(vs_OCID) = TIMS.ClearSQM(ViewState(vs_OCID))
        If ViewState(vs_OCID) = "" OrElse ViewState(vs_OCID) = "0" Then Return

        '2005/5/24修正開班資料中開結訓日期-CLASS_CLASSINFO
        'Dim objrow2 AS DataRow
        Dim class_table As DataTable = Nothing
        Dim objadapter2 As SqlDataAdapter = Nothing

        s_TransType = TIMS.cst_TRANS_LOG_Update 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
        s_TargetTable = "CLASS_CLASSINFO"
        s_FuncPath = "/TC/06/TC_06_001_chk"
        s_WHERE = String.Format("OCID={0} AND PlanID={1} AND ComIDNO='{2}' AND SeqNo={3}", ViewState(vs_OCID), rPlanID, rComIDNO, rSeqNO)

        Dim str_classinfo As String = ""
        str_classinfo = String.Concat("SELECT * FROM CLASS_CLASSINFO WHERE ", s_WHERE)
        class_table = DbAccess.GetDataTable(str_classinfo, objadapter2, oTrans)
        If TIMS.dtNODATA(class_table) Then Return '(查無資料)

        'OJT-22022301：產投 - 班級變更審核：增加【報名開始日期】、【報名結束日期】 分署提案
        Dim flag_can_SET_EnterDate As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID < 2 AndAlso iAltDataID = Cst_i訓練期間, True, False)
        flag_can_SET_EnterDate = If(vActType = cst_dgAct_Edit1, flag_can_SET_EnterDate, False)

        Dim objrow2 As DataRow = class_table.Rows(0)
        Select Case iAltDataID
            Case Cst_i訓練期間 '1
                'SELECT SEnterDate,FEnterDate,Examdate,ExamPeriod,FEnterDate2 FROM CLASS_CLASSINFO where rownum <=10
                objrow2("STDate") = TIMS.Cdate2(ASDate.Text)
                objrow2("FTDate") = TIMS.Cdate2(AEDate.Text)
                If flag_can_SET_EnterDate Then
                    '報名開始日期'(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
                    Dim vsSEnterDate As String = TIMS.GET_YMDHM1(SEnterDate.Text, TIMS.GetListValue(SEnterDate_HR), TIMS.GetListValue(SEnterDate_MI))
                    If vsSEnterDate <> "" Then objrow2("SEnterDate") = If(vsSEnterDate <> "", CDate(vsSEnterDate), Convert.DBNull)
                    '報名結束日期'(NEW)'(OLD)'(正確有值) '取得有效值(若有新值依新VAL，無則用舊VAL)
                    Dim vsFEnterDate As String = TIMS.GET_YMDHM1(FEnterDate.Text, TIMS.GetListValue(FEnterDate_HR), TIMS.GetListValue(FEnterDate_MI))
                    If vsFEnterDate <> "" Then objrow2("FEnterDate") = If(vsFEnterDate <> "", CDate(vsFEnterDate), Convert.DBNull)
                Else
                    '20081110 andy 一般計畫更新報名日期
                    If New_SEnterDate2.Text <> "" Then objrow2("SEnterDate") = CDate(New_SEnterDate2.Text)
                    If New_FEnterDate2.Text <> "" Then objrow2("FEnterDate") = CDate(New_FEnterDate2.Text)
                End If

                If New_Examdate.Text <> "" Then objrow2("Examdate") = CDate(New_Examdate.Text)
                If New_Examdate.Text <> "" AndAlso New_ExamPeriod.Text <> "" AndAlso HidNew_ExamPeriod.Value <> "" Then objrow2("ExamPeriod") = HidNew_ExamPeriod.Value
                If New_CheckInDate.Text <> "" Then objrow2("CheckInDate") = CDate(New_CheckInDate.Text)
                If Hid_sFENTERDATE2.Value <> "" Then objrow2("FEnterDate2") = CDate(Hid_sFENTERDATE2.Value)
            Case Cst_i課程編配 '4
                objrow2("THours") = Val(EGenSci.Text) + Val(EProSci.Text) + Val(EProTech.Text) + Val(EOther.Text)
            Case Cst_i班別名稱 '6
                objrow2("ClassCName") = ChangeOClassName.Text
            Case Cst_i期別 '7
                Dim vCyclType As String = TIMS.FmtCyclType(ChangeCyclType.Text)
                objrow2("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
            Case Cst_i上課地址 '8
                'CLASS_CLASSINFO
                hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
                objrow2("TaddressZip") = Val(NewData8_1.Value)
                objrow2("TaddressZip6W") = hidNewData8_6W.Value 'Val(NewData8_3.Value)
                objrow2("TAddress") = NewData8_2.Value
            Case Cst_i停辦 '9 '停辦-CLASS_CLASSINFO
                objrow2("NotOpen") = NewData9_1.Value
            Case Cst_i上課時段 '10
                objrow2("TPeriod") = NewData10_1.Value
            Case Cst_i核定人數 '12
                objrow2("TNum") = NewData12_1.Text
            Case Cst_i科場地 '14
                Dim vTADDRESSZIP As String = ""
                Dim vTADDRESSZIP6W As String = ""
                Dim vTADDRESS As String = ""
                '順序有變化 NewData8_5->NewData8_4 '術科場地地址
                If Hid_NewData8_5.Value <> "" Then
                    vTADDRESSZIP = hid_TP_ZIPCODE.Value
                    vTADDRESSZIP6W = hid_TP_ZIP6W.Value
                    vTADDRESS = hid_TP_ADDRESS.Value
                End If
                '順序有變化 NewData8_5->NewData8_4 '學科場地地址
                If Hid_NewData8_4.Value <> "" Then
                    vTADDRESSZIP = hid_SP_ZIPCODE.Value
                    vTADDRESSZIP6W = hid_SP_ZIP6W.Value
                    vTADDRESS = hid_SP_ADDRESS.Value
                End If
                objrow2("TADDRESSZIP") = If(vTADDRESSZIP <> "", vTADDRESSZIP, Convert.DBNull)
                objrow2("TADDRESSZIP6W") = If(vTADDRESSZIP6W <> "", vTADDRESSZIP6W, Convert.DBNull)
                objrow2("TADDRESS") = If(vTADDRESS <> "", vTADDRESS, Convert.DBNull)
        End Select
        objrow2("LastState") = "M" 'M: 修改(最後異動狀態)
        objrow2("ModifyAcct") = sm.UserInfo.UserID
        objrow2("ModifyDate") = Now()

        'ADD 記錄交易LOG (CLASS_CLASSINFO=> SYS_TRANS_LOG)
        Dim htPP As Hashtable = TIMS.Get_HashTablePP(s_TransType, s_TargetTable, s_FuncPath, s_WHERE)
        Call TIMS.SaveTRANSLOG(sm, oTrans.Connection, oTrans, objrow2, htPP)

        DbAccess.UpdateDataTable(class_table, objadapter2, oTrans)
    End Sub

    ''' <summary>  UPDATE PLAN_PLANINFO -(SAVE) </summary>
    ''' <param name="iAltDataID"></param>
    ''' <param name="oTrans"></param>
    Sub UPDATE_PLANPLANINFO(ByRef iAltDataID As Integer, ByRef oTrans As SqlTransaction)
        'UPDATE PLAN_PLANINFO -(SAVE)
        Const Cst_sChkmode As String = ",1,4,6,7,8,9,10,12,14,19,21,22," '功能 (計畫用功能)
        Dim ffChkmode As String = String.Format(",{0},", CStr(iAltDataID))
        If Cst_sChkmode.IndexOf(ffChkmode) = -1 Then Return

        Dim objtable As DataTable = Nothing
        Dim objadapter As SqlDataAdapter = Nothing

        s_TransType = TIMS.cst_TRANS_LOG_Update 'insert:cst_TRANS_LOG_Insert/update:cst_TRANS_LOG_Update
        s_TargetTable = "PLAN_PLANINFO"
        s_FuncPath = "/TC/06/TC_06_001_chk"
        s_WHERE = String.Format("PlanID={0} AND ComIDNO='{1}' AND SeqNo={2}", rPlanID, rComIDNO, rSeqNO)

        Dim objstr As String = ""
        objstr = String.Concat("SELECT * FROM PLAN_PLANINFO WHERE ", s_WHERE)
        objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
        'PLAN_PLANINFO
        If objtable.Rows.Count = 0 Then Return

        'OJT-22022301：產投 - 班級變更審核：增加【報名開始日期】、【報名結束日期】 分署提案
        Dim flag_can_SET_EnterDate As Boolean = If(TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso sm.UserInfo.LID < 2 AndAlso iAltDataID = Cst_i訓練期間, True, False)
        flag_can_SET_EnterDate = If(vActType = cst_dgAct_Edit1, flag_can_SET_EnterDate, False)

        Dim objrow As DataRow = objtable.Rows(0)
        Select Case iAltDataID
            Case Cst_i遠距教學
                objrow("DISTANCE") = If(Hid_DISTANCE_new.Value <> "", Hid_DISTANCE_new.Value, Convert.DBNull)
            Case Cst_i訓練費用 '21 'Cst_訓練費用
                TotalCost1New.Text = TIMS.ClearSQM(TotalCost1New.Text)
                If TotalCost1New.Text <> "" Then objrow("TotalCost") = Val(TotalCost1New.Text)
            Case Cst_i訓練期間 '1
                objrow("STDate") = ASDate.Text
                objrow("FDDate") = AEDate.Text
                If New_SEnterDate2.Text <> "" Then objrow("SEnterDate") = CDate(New_SEnterDate2.Text)
                If New_FEnterDate2.Text <> "" Then objrow("FEnterDate") = CDate(New_FEnterDate2.Text)
                If New_Examdate.Text <> "" Then objrow("Examdate") = CDate(New_Examdate.Text)
                If New_Examdate.Text <> "" AndAlso New_ExamPeriod.Text <> "" AndAlso HidNew_ExamPeriod.Value <> "" Then objrow("ExamPeriod") = HidNew_ExamPeriod.Value
                If New_CheckInDate.Text <> "" Then objrow("CheckInDate") = CDate(New_CheckInDate.Text)

            Case Cst_i課程編配 '4
                objrow("GenSciHours") = CInt(EGenSci.Text)
                objrow("ProSciHours") = CInt(EProSci.Text)
                objrow("ProTechHours") = CInt(EProTech.Text)
                objrow("OtherHours") = CInt(EOther.Text)
                objrow("TotalHours") = Val(EGenSci.Text) + Val(EProSci.Text) + Val(EProTech.Text) + Val(EOther.Text)
            Case Cst_i班別名稱 '6
                objrow("ClassName") = ChangeOClassName.Text
            Case Cst_i期別 '7
                Dim vCyclType As String = TIMS.FmtCyclType(ChangeCyclType.Text)
                objrow("CyclType") = If(vCyclType <> "", vCyclType, Convert.DBNull)
            Case Cst_i上課地址 '8
                'PLAN_PLANINFO
                hidNewData8_6W.Value = TIMS.GetZIPCODE6W(NewData8_1.Value, NewData8_3.Value)
                objrow("TaddressZip") = Val(NewData8_1.Value)
                objrow("TADDRESSZIP6W") = hidNewData8_6W.Value 'Val(NewData8_3.Value) 'objrow("TADDRESSZIP6W")=Val(NewData8_3.Value)
                objrow("TAddress") = NewData8_2.Value
            Case Cst_i核定人數 '12
                Dim i_TotalCost As Integer = 0
                Dim i_OldData12 As Integer = 0
                Dim i_NewData12 As Integer = 0
                objrow("TNum") = NewData12_1.Text
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If TIMS.IsDBNull2(objrow("TotalCost")) Then i_TotalCost = 0 Else i_TotalCost = Val(objrow("TotalCost"))
                    If OldData12_1.Text = "" Or OldData12_1.Text = "0" Then i_OldData12 = 1 Else i_OldData12 = Val(OldData12_1.Text)
                    If NewData12_1.Text = "" Or NewData12_1.Text = "0" Then i_NewData12 = 1 Else i_NewData12 = Val(NewData12_1.Text)
                    objrow("TotalCost") = (i_TotalCost / i_OldData12) * i_NewData12
                    objrow("DefStdCost") = objrow("TotalCost") * 0.2
                    objrow("DefGovCost") = objrow("TotalCost") * 0.8
                End If
                hid_NewData12.Value = i_NewData12

            Case Cst_i科場地 '14'學(術)科場地
                'Const cst_SciPlaceID As String="SciPlaceID"
                'Const cst_TechPlaceID As String="TechPlaceID"
                'Const cst_SciPlaceID2 As String="SciPlaceID2"
                'Const cst_TechPlaceID2 As String="TechPlaceID2"
                'Dim str_CHG_ADDRESS_TYPE AS String=""
                Dim v_NewData14_1 As String = TIMS.GetListValue(NewData14_1)
                Dim v_NewData14_2 As String = TIMS.GetListValue(NewData14_2)
                Dim v_NewData14_3 As String = TIMS.GetListValue(NewData14_3)
                Dim v_NewData14_4 As String = TIMS.GetListValue(NewData14_4)
                objrow("SciPlaceID") = If(v_NewData14_1 <> "", v_NewData14_1, Convert.DBNull)
                objrow("TechPlaceID") = If(v_NewData14_2 <> "", v_NewData14_2, Convert.DBNull)
                objrow("SciPlaceID2") = If(v_NewData14_3 <> "", v_NewData14_3, Convert.DBNull)
                objrow("TechPlaceID2") = If(v_NewData14_4 <> "", v_NewData14_4, Convert.DBNull)
                objrow("AddressTechPTID2") = If(Hid_NewData8_7.Value <> "", Hid_NewData8_7.Value, Convert.DBNull) '術科場地地址2
                objrow("AddressSciPTID2") = If(Hid_NewData8_6.Value <> "", Hid_NewData8_6.Value, Convert.DBNull)  '學科場地地址2
                objrow("AddressTechPTID") = If(Hid_NewData8_5.Value <> "", Hid_NewData8_5.Value, Convert.DBNull) '術科場地地址
                objrow("AddressSciPTID") = If(Hid_NewData8_4.Value <> "", Hid_NewData8_4.Value, Convert.DBNull)  '學科場地地址
                Dim vTADDRESSZIP As String = ""
                Dim vTADDRESSZIP6W As String = ""
                Dim vTADDRESS As String = ""
                '順序有變化 NewData8_5->NewData8_4 '術科場地地址
                If Hid_NewData8_5.Value <> "" Then
                    vTADDRESSZIP = hid_TP_ZIPCODE.Value
                    vTADDRESSZIP6W = hid_TP_ZIP6W.Value
                    vTADDRESS = hid_TP_ADDRESS.Value
                End If
                '順序有變化 NewData8_5->NewData8_4 '學科場地地址
                If Hid_NewData8_4.Value <> "" Then
                    vTADDRESSZIP = hid_SP_ZIPCODE.Value
                    vTADDRESSZIP6W = hid_SP_ZIP6W.Value
                    vTADDRESS = hid_SP_ADDRESS.Value
                End If
                objrow("TADDRESSZIP") = If(vTADDRESSZIP <> "", vTADDRESSZIP, Convert.DBNull)
                objrow("TADDRESSZIP6W") = If(vTADDRESSZIP6W <> "", vTADDRESSZIP6W, Convert.DBNull)
                objrow("TADDRESS") = If(vTADDRESS <> "", vTADDRESS, Convert.DBNull)
            Case Cst_i包班種類 '19
                objrow("PackageType") = If(hidPackageTypeNew.Value <> "", hidPackageTypeNew.Value, Convert.DBNull)
        End Select
        objrow("ModifyAcct") = sm.UserInfo.UserID
        objrow("ModifyDate") = Now()

        'ADD 記錄交易LOG (PLAN_PLANINFO=> SYS_TRANS_LOG)
        Dim htPP As Hashtable = TIMS.Get_HashTablePP(s_TransType, s_TargetTable, s_FuncPath, s_WHERE)
        Call TIMS.SaveTRANSLOG(sm, oTrans.Connection, oTrans, objrow, htPP)

        DbAccess.UpdateDataTable(objtable, objadapter, oTrans)
    End Sub

    ''' <summary> UPDATE PLAN_REVISE -(SAVE) </summary>
    ''' <param name="iAltDataID"></param>
    ''' <param name="oTrans"></param>
    Sub UPDATE_PLANREVISE(ByRef iAltDataID As Integer, ByRef oTrans As SqlTransaction)
        Dim objtable As DataTable = Nothing
        Dim objadapter As SqlDataAdapter = Nothing

        Dim objstr As String = ""
        objstr &= " SELECT *" & vbCrLf
        objstr &= " FROM PLAN_REVISE" & vbCrLf
        objstr &= " WHERE PlanID=" & rPlanID & " AND ComIDNO='" & rComIDNO & "' AND SeqNo='" & rSeqNO & "' "
        objstr &= " AND CDate=" & TIMS.To_date(rSCDate) & " AND SubSeqNO=" & iSubSeqNO & vbCrLf
        objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
        '"查無該計畫變更資料，請重新查詢
        If objtable.Rows.Count = 0 Then Return

        Dim objdr As DataRow = objtable.Rows(0)
        'Dim i_AltDataID As Integer=If(Convert.ToString(objdr("AltDataID")) <> "", Val(objdr("AltDataID")), 0)
        'If i_AltDataID <> iAltDataID Then
        '    'input 與查詢資料不同（異常離開）
        '    Dim strEx As String="input 與查詢資料不同（異常離開）,請重新查詢"
        '    Throw New Exception(strEx)
        'End If

        objdr("ReviseStatus") = "Y"
        objdr("REVISEDATE") = Now()
        objdr("Reason") = If(ReviseCont.Text <> "", ReviseCont.Text, Convert.DBNull) 'Add by jack 04/12/30
        objdr("Verifier") = sm.UserInfo.UserID     'Add by jack 04/12/30
        '------------------
        objdr("ModifyAcct") = sm.UserInfo.UserID   'Add by andy  20090324
        objdr("ModifyDate") = Now()                'Add by andy  20090324  最後審核異動時間需記錄
        '-----------------
        DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

        '資料-產投-審查計分表
        'Call save_divCo128()
    End Sub

    ''' <summary> GET PLAN_REVISE DataRow </summary>
    ''' <param name="oTrans"></param>
    ''' <returns></returns>
    Function GET_PLANREVISE_DataRow(ByRef oTrans As SqlTransaction) As DataRow
        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(Convert.ToString(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Convert.ToString(Request("SubNo")))), 0)

        Dim objDr As DataRow = Nothing
        Dim objtable As DataTable = Nothing
        Dim parms As New Hashtable From {{"PlanID", Val(rPlanID)}, {"ComIDNO", rComIDNO}, {"SeqNo", Val(rSeqNO)}, {"CDate", TIMS.Cdate2(rSCDate)}, {"SubSeqNO", iSubSeqNO}}
        Dim objstr As String = ""
        objstr &= " SELECT *" & vbCrLf
        objstr &= " FROM PLAN_REVISE" & vbCrLf
        objstr &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo AND CDate=@CDate AND SubSeqNO=@SubSeqNO"
        'objstr &=" WHERE PlanID=" & rPlanID & " AND ComIDNO='" & rComIDNO & "' AND SeqNo='" & rSeqNO & "' "
        'objstr &=" AND CDate=" & TIMS.to_date(rSCDate) & vbCrLf
        'objstr &=" AND SubSeqNO=" & iSubSeqNO & vbCrLf
        objtable = DbAccess.GetDataTable(objstr, oTrans, parms)
        If objtable.Rows.Count = 0 Then Return objDr 'Exit Sub

        objDr = objtable.Rows(0)
        Return objDr
    End Function

    '(通過應該變動相關資料) 儲存Y 'Public Shared Sub SaveData1Y(ByRef oConn AS SqlConnection, ByRef oTrans AS SqlTransaction)
    ''' <summary>
    ''' (通過應該變動相關資料) 儲存Y PLAN_REVISE
    ''' </summary>
    ''' <param name="oConn"></param>
    ''' <param name="oTrans"></param>
    Sub SaveData1Y(ByRef oConn As SqlConnection, ByRef oTrans As SqlTransaction)
        Dim objdr As DataRow = GET_PLANREVISE_DataRow(oTrans)
        If objdr Is Nothing Then
            Common.MessageBox(Me, "查無該計畫變更資料，請重新查詢")
            Exit Sub
        End If
        'objdr=objtable.Rows(0)
        hid_OldData2_2.Value = TIMS.NullToStr(objdr("OldData2_2"))
        hid_OldData2_3.Value = TIMS.NullToStr(objdr("OldData2_3"))
        hid_NewData2_2.Value = TIMS.NullToStr(objdr("NewData2_2"))
        hid_NewData2_3.Value = TIMS.NullToStr(objdr("NewData2_3"))
        hid_OldData3_3.Value = TIMS.NullToStr(objdr("OldData3_3"))
        hid_NewData3_1.Value = TIMS.NullToStr(objdr("NewData3_1"))
        hid_OldData5_3.Value = TIMS.NullToStr(objdr("OldData5_3"))
        hid_NewData5_1.Value = TIMS.NullToStr(objdr("NewData5_1"))
        '功能代碼  'Cst_sChkmode: (sChkmode_AltDataID)
        Dim i_AltDataID As Integer = If(Convert.ToString(objdr("AltDataID")) <> "", Val(objdr("AltDataID")), 0)
        If i_AltDataID = 0 Then
            Common.MessageBox(Me, "查無該計畫變更資料，請重新查詢")
            Exit Sub
        End If

        Call UPDATE_PLANREVISE(i_AltDataID, oTrans)
        'If sChkmode_AltDataID=0 Then
        '    Common.MessageBox(Me, "查無該計畫變更資料，請重新查詢")
        '    Exit Sub
        'End If

        'UPDATE PLAN_PLANINFO -(SAVE)
        Call UPDATE_PLANPLANINFO(i_AltDataID, oTrans)
        'Const Cst_sChkmode As String=",1,4,6,7,8,9,10,12,14,19,21,22," '功能 (計畫用功能)
        'Dim ffChkmode As String=String.Format(",{0},", CStr(i_AltDataID))
        'If Cst_sChkmode.IndexOf(ffChkmode) > -1 Then
        Call UPDATE_CLASSCLASSINFO(i_AltDataID, oTrans)

        Dim objstr As String = ""
        Dim objtable As DataTable = Nothing
        Dim objadapter As SqlDataAdapter = Nothing

        Select Case i_AltDataID
            Case Cst_i遠距教學
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                End If
            Case Cst_i停辦
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso NewData9_1.Value = "Y" Then
                    Call UPDATE_CLASS_STUDENTSOFCLASS_ALT9(i_AltDataID, oConn, oTrans)
                End If
            Case Cst_i核定人數
                '2007/04/15 修正產學訓計畫經費資料 Add by Kevin
                hid_NewData12.Value = TIMS.ClearSQM(hid_NewData12.Value)
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 AndAlso hid_NewData12.Value <> "" Then
                    objstr = ""
                    objstr &= " UPDATE PLAN_COSTITEM "
                    objstr &= " SET Itemage=" & hid_NewData12.Value & " "
                    objstr &= " WHERE PlanID=" & rPlanID & " AND ComIDNO='" & rComIDNO & "' AND SeqNo='" & rSeqNO & "' "
                    DbAccess.ExecuteNonQuery(objstr, oTrans)
                End If

            Case Cst_i上課時段
                'PLAN_VERREPORT
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    objstr = ""
                    objstr &= " UPDATE PLAN_VERREPORT "
                    objstr &= " SET TPeriod='" & NewData10_1.Value & "' "
                    objstr &= " WHERE PlanID=" & rPlanID & " AND ComIDNO='" & rComIDNO & "' AND SeqNo='" & rSeqNO & "' "
                    DbAccess.ExecuteNonQuery(objstr, oTrans)
                End If

            Case Cst_i科場地
                '20080709 andy 更新產學訓課程表
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                End If
            Case Cst_i訓練期間
                'Dim DateNow AS String=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                objstr = ""
                objstr &= " UPDATE CLASS_STUDENTSOFCLASS" & vbCrLf
                objstr &= " SET OpenDate=" & TIMS.To_date(ASDate.Text) & vbCrLf
                objstr &= " ,CloseDate=" & TIMS.To_date(AEDate.Text) & vbCrLf
                objstr &= " ,ModifyAcct='" & sm.UserInfo.UserID & "'" & vbCrLf
                objstr &= " ,ModifyDate=GETDATE()" & vbCrLf
                objstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'"
                DbAccess.ExecuteNonQuery(objstr, oTrans)
                '20080709 andy 更新產學訓課程表
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                End If
            Case Cst_i訓練費用 '21 'Cst_訓練費用
                If dtR2G.Rows.Count = 0 Then Exit Select '若為0不儲存

                Dim sql_i As String = ""
                sql_i &= " INSERT INTO PLAN_COSTITEM (PCID,PLANID,COMIDNO,SEQNO,COSTMODE,COSTID,ITEMOTHER,OPRICE,ITEMAGE,ITEMCOST,ADMFLAG,MODIFYACCT,MODIFYDATE,TAXFLAG)" & vbCrLf
                sql_i &= " VALUES (@PCID,@PLANID,@COMIDNO,@SEQNO,@COSTMODE,@COSTID,@ITEMOTHER,@OPRICE,@ITEMAGE,@ITEMCOST,@ADMFLAG,@MODIFYACCT,getdate(),@TAXFLAG)" & vbCrLf
                Dim iCmd As New SqlCommand(sql_i, oConn, oTrans)

                Dim sql_d As String = ""
                sql_d &= " DELETE PLAN_COSTITEM" & vbCrLf
                sql_d &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
                Dim dCmd As New SqlCommand(sql_d, oConn, oTrans)
                '執行刪除
                With dCmd
                    .Parameters.Clear()
                    .Parameters.Add("PLANID", SqlDbType.VarChar).Value = rPlanID
                    .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
                    .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = rSeqNO
                    '.ExecuteNonQuery()  'edit，by:20181012
                    DbAccess.ExecuteNonQuery(dCmd.CommandText, oTrans, dCmd.Parameters)  'edit，by:20181012
                End With

                '執行新增(多筆)
                'dt3=DbAccess.GetDataTable(sql, tConn)
                For Each drV As DataRow In dtR2G.Rows
                    Dim iPCID As Integer = DbAccess.GetNewId(oTrans, "PLAN_COSTITEM_PCID_SEQ,PLAN_COSTITEM,PCID")
                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("PCID", SqlDbType.Int).Value = iPCID
                        .Parameters.Add("PLANID", SqlDbType.VarChar).Value = rPlanID
                        .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = rComIDNO
                        .Parameters.Add("SEQNO", SqlDbType.VarChar).Value = rSeqNO
                        .Parameters.Add("COSTMODE", SqlDbType.Int).Value = Val(Hid_CostMode.Value)
                        .Parameters.Add("COSTID", SqlDbType.VarChar).Value = drV("COSTID")
                        .Parameters.Add("ITEMOTHER", SqlDbType.VarChar).Value = drV("ITEMOTHER")
                        .Parameters.Add("OPRICE", SqlDbType.Float).Value = drV("OPRICE")
                        .Parameters.Add("ITEMAGE", SqlDbType.Int).Value = drV("ITEMAGE")
                        .Parameters.Add("ITEMCOST", SqlDbType.Int).Value = drV("ITEMCOST")
                        .Parameters.Add("ADMFLAG", SqlDbType.VarChar).Value = drV("ADMFLAG")
                        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        .Parameters.Add("TAXFLAG", SqlDbType.VarChar).Value = drV("TAXFLAG")
                        '.ExecuteNonQuery()  'edit，by:20181012
                        DbAccess.ExecuteNonQuery(iCmd.CommandText, oTrans, iCmd.Parameters)  'edit，by:20181012
                    End With
                Next

            Case Cst_i包班種類 '19
                '包班種類 Plan_BusPackage
                'Dim dt1, dt2, dtTemp AS DataTable
                'Dim dr, dr1, dr2 AS DataRow
                'Dim DateNow AS String=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                'xSql="DELETE Plan_BusPackage where PlanID=" & Request("PlanID") & " and ComIDNO='" & Request("cid") & "' and SeqNo=" & Request("no")
                Dim xSql As String = ""
                xSql &= " DELETE PLAN_BUSPACKAGE "
                xSql &= " WHERE PlanID=" & rPlanID & " AND ComIDNO='" & rComIDNO & "' AND SeqNo='" & rSeqNO & "' "
                DbAccess.ExecuteNonQuery(xSql, oTrans)

                If Not Session("Revise_BusPackage") Is Nothing AndAlso DG_BusPackageNew.Items.Count > 0 Then
                    Dim dtTemp As DataTable
                    dtTemp = Session("Revise_BusPackage")
                    objstr = " SELECT * FROM PLAN_BUSPACKAGE WHERE 1<>1 "
                    objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                    For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        Dim objrow As DataRow = Nothing
                        objrow = objtable.NewRow
                        objtable.Rows.Add(objrow)
                        objrow("BPID") = DbAccess.GetNewId(oTrans, "PLAN_BUSPACKAGE_BPID_SEQ,PLAN_BUSPACKAGE,BPID")
                        objrow("PlanID") = dr("PlanID")
                        objrow("ComIDNO") = dr("ComIDNO")
                        objrow("SeqNo") = dr("SeqNo")
                        objrow("uname") = dr("uname")
                        objrow("intaxno") = dr("intaxno")
                        objrow("ubno") = dr("ubno")
                        objrow("ModifyAcct") = sm.UserInfo.UserID
                        objrow("ModifyDate") = Now
                    Next
                    DbAccess.UpdateDataTable(objtable, objadapter, oTrans)
                End If

            Case Cst_i訓練時段 '2
                objstr = ""
                objstr &= " SELECT * FROM CLASS_SCHEDULE"
                objstr &= " WHERE OCID='" & ViewState(vs_OCID) & "'" & vbCrLf
                objstr &= " AND (SchoolDate=" & TIMS.To_date(TimeSDate.Text) & " OR SchoolDate=" & TIMS.To_date(TimeEDate.Text) & ") "
                objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)

                Dim newrow As DataRow = Nothing
                Dim oldrow As DataRow = Nothing
                Dim Sqlold As String = "" '原始排課資料
                Dim Sqlnew As String = "" '變更排課資料
                Dim Sqlold2 As String = "" '原始排課資料
                Dim Sqlnew2 As String = "" '變更排課資料
                'Sqlnew=" UPDATE CLASS_SCHEDULE SET ModifyAcct='" & sm.UserInfo.UserID & "' ,ModifyDate='" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf
                'Sqlold=" UPDATE CLASS_SCHEDULE SET ModifyAcct='" & sm.UserInfo.UserID & "' ,ModifyDate='" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf
                Sqlold = " UPDATE CLASS_SCHEDULE SET ModifyAcct='" & sm.UserInfo.UserID & "' ,ModifyDate=GETDATE()," & vbCrLf
                Sqlnew = " UPDATE CLASS_SCHEDULE SET ModifyAcct='" & sm.UserInfo.UserID & "' ,ModifyDate=GETDATE()," & vbCrLf
                Dim oldary As String() '節次(變更前)
                Dim oclassary As String()  '課程(變更前)
                Dim newary As String() '節次(變更後)
                'Dim nclassary AS String()
                '一定會2筆才可對調咩 (可能是同1天所以只有1筆)
                If objtable.Rows.Count > 0 Then
                    '原資料/原日期
                    oldrow = objtable.Select("SchoolDate='" & TimeSDate.Text & "'")(0)
                    oldary = Split(hid_OldData2_3.Value, ",")     '節次(變更前)
                    oclassary = Split(hid_OldData2_2.Value, ",")  '課程(變更前)

                    '互換資料/互換日期
                    newrow = objtable.Select("SchoolDate='" & TimeEDate.Text & "'")(0)
                    newary = Split(hid_NewData2_3.Value, ",")      '節次(變更後)
                    'nclassary=Split(NewData2_2, ",")   '課程(變更後)
                    'newrow("class" & newary(i))'舊(原)資料 改為要互換的課程
                    'Convert.ToString(newrow("class" & newary(i))) <> "" 互換課程不為空
                    '舊(原)改互換(新)資料update
                    For i As Integer = 0 To oldary.Length - 1
                        If Convert.ToString(newrow("class" & newary(i))) <> "" Then                 '課程
                            Sqlnew2 = getStr(Sqlnew2, "class" & oldary(i) & "='" & newrow("class" & newary(i)) & "'")
                        Else
                            Sqlnew2 = getStr(Sqlnew2, "class" & oldary(i) & "=null")
                        End If
                        If Convert.ToString(newrow("Room" & newary(i))) <> "" Then                  '教室
                            Sqlnew2 = getStr(Sqlnew2, "Room" & oldary(i) & "='" & Trim(newrow("Room" & newary(i))) & "'")
                        Else
                            Sqlnew2 = getStr(Sqlnew2, "Room" & oldary(i) & "=null")
                        End If
                        If Convert.ToString(newrow("Teacher" & newary(i))) <> "" Then               '教師(一)
                            Sqlnew2 = getStr(Sqlnew2, "Teacher" & oldary(i) & "='" & Trim(newrow("Teacher" & newary(i))) & "'")
                        Else
                            Sqlnew2 = getStr(Sqlnew2, "Teacher" & oldary(i) & "=null")
                        End If
                        If Convert.ToString(newrow("Teacher" & CStr(CInt(newary(i)) + 12))) <> "" Then   '助教1'教師(二)
                            Sqlnew2 = getStr(Sqlnew2, "Teacher" & CStr(CInt(oldary(i)) + 12) & "='" & Trim(newrow("Teacher" & CStr(CInt(newary(i)) + 12))) & "'")
                        Else
                            Sqlnew2 = getStr(Sqlnew2, "Teacher" & CStr(CInt(oldary(i)) + 12) & "=null")
                        End If
                        If Convert.ToString(newrow("Teacher" & CStr(CInt(newary(i)) + 24))) <> "" Then   '助教2'教師(二)
                            Sqlnew2 = getStr(Sqlnew2, "Teacher" & CStr(CInt(oldary(i)) + 24) & "='" & Trim(newrow("Teacher" & CStr(CInt(newary(i)) + 24))) & "'")
                        Else
                            Sqlnew2 = getStr(Sqlnew2, "Teacher" & CStr(CInt(oldary(i)) + 24) & "=null")
                        End If
                        Sqlnew2 += vbCrLf
                    Next

                    'Trim(oldrow("class" & oldary(i))) 互換 改為舊(原)的課程
                    '互換(新)改舊(原)資料update
                    For i As Integer = 0 To newary.Length - 1
                        If Convert.ToString(oldrow("class" & oldary(i))).Trim <> "" Then                 '課程
                            Sqlold2 = getStr(Sqlold2, "class" & newary(i) & "='" & Trim(oldrow("class" & oldary(i))) & "'")
                        Else
                            Sqlold2 = getStr(Sqlold2, "class" & newary(i) & "=null")
                        End If
                        If Convert.ToString(oldrow("Room" & oldary(i))).Trim <> "" Then                '教室
                            Sqlold2 = getStr(Sqlold2, "Room" & newary(i) & "='" & Trim(oldrow("Room" & oldary(i))) & "'")
                        Else
                            Sqlold2 = getStr(Sqlold2, "Room" & newary(i) & "=null")
                        End If
                        If Convert.ToString(oldrow("Teacher" & oldary(i))).Trim <> "" Then             '教師(一)
                            Sqlold2 = getStr(Sqlold2, "Teacher" & newary(i) & "='" & Trim(oldrow("Teacher" & oldary(i))) & "'")
                        Else
                            Sqlold2 = getStr(Sqlold2, "Teacher" & newary(i) & "=null")
                        End If
                        If Convert.ToString(oldrow("Teacher" & CStr(CInt(oldary(i)) + 12))).Trim <> "" Then  '助教1'教師(二)
                            Sqlold2 = getStr(Sqlold2, "Teacher" & CStr(CInt(newary(i)) + 12) & "='" & Trim(oldrow("Teacher" & Trim(CStr(CInt(oldary(i)) + 12)))) & "'")
                        Else
                            Sqlold2 = getStr(Sqlold2, "Teacher" & CStr(CInt(newary(i)) + 12) & "=null")
                        End If
                        If Convert.ToString(oldrow("Teacher" & CStr(CInt(oldary(i)) + 24))).Trim <> "" Then  '助教2'教師(二)
                            Sqlold2 = getStr(Sqlold2, "Teacher" & CStr(CInt(newary(i)) + 24) & "='" & Trim(oldrow("Teacher" & Trim(CStr(CInt(oldary(i)) + 24)))) & "'")
                        Else
                            Sqlold2 = getStr(Sqlold2, "Teacher" & CStr(CInt(newary(i)) + 24) & "=null")
                        End If
                        Sqlold2 += vbCrLf
                    Next

                    '舊(原)改互換(新)資料 使用 舊(原)日期
                    Sqlnew = String.Concat(Trim(Sqlnew), Sqlnew2, " WHERE OCID='", ViewState(vs_OCID), "' AND SchoolDate=", TIMS.To_date(TimeSDate.Text))

                    '互換(新)改舊(原)資料 使用 互換(新)日期
                    Sqlold = String.Concat(Trim(Sqlold), Sqlold2, " WHERE OCID='", ViewState(vs_OCID), "' AND SchoolDate=", TIMS.To_date(TimeEDate.Text))

                    '本程式記錄SQL動作
                    TIMS.SaveSqlCommon("TC_06_001", String.Concat(TIMS.GetErrorMsg(Me), vbCrLf, "/**Sqlnew:**/", vbCrLf, Sqlnew, vbCrLf, "/**Sqlold:**/", vbCrLf, Sqlold))

                    'Call TIMS.OpenDbConn(objconn)
                    If TimeSDate.Text = TimeEDate.Text Then
                        '當日採覆蓋
                        Dim uCmd As New SqlCommand(Sqlnew, oConn, oTrans)
                        With uCmd
                            .Parameters.Clear()
                            .ExecuteNonQuery()
                        End With
                    Else
                        '不同日採交換
                        Dim uCmd As New SqlCommand(Sqlnew, oConn, oTrans)
                        With uCmd
                            .Parameters.Clear()
                            .ExecuteNonQuery()
                        End With
                        Dim uCmd2 As New SqlCommand(Sqlold, oConn, oTrans)
                        With uCmd2
                            .Parameters.Clear()
                            .ExecuteNonQuery()
                        End With
                    End If
                End If

            'Try
            'Catch ex AS Exception
            '    DbAccess.RollbackTrans(objTrans)
            '    Common.MessageBox(Me, ex.ToString)
            '    'Throw ex
            'End Try
            Case Cst_i訓練地點 '3
                If ViewState(vs_OCID) <> "" Then
                    objstr = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND schoolDate=" & TIMS.To_date(PlaceDate.Text) & vbCrLf
                    objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                    Dim oldary As String()
                    'Dim oclassary AS String()
                    'Dim newary AS String()
                    'Dim nclassary AS String()
                    If objtable.Rows.Count > 0 Then
                        Dim objrow As DataRow = Nothing
                        objrow = objtable.Rows(0)
                        oldary = Split(hid_OldData3_3.Value, ",")
                        For i As Integer = 0 To oldary.Length - 1
                            objrow("Room" & oldary(i)) = hid_NewData3_1.Value
                        Next
                    End If
                    DbAccess.UpdateDataTable(objtable, objadapter, oTrans)
                End If

            Case Cst_i訓練師資 '5
                objstr = " SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & ViewState(vs_OCID) & "' AND schoolDate=" & TIMS.To_date(TechDate.Text) & vbCrLf '" & TechDate.Text & "'"
                objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                Dim oldary As String()
                If objtable.Rows.Count > 0 Then
                    Dim objrow As DataRow = Nothing
                    objrow = objtable.Rows(0)
                    oldary = Split(hid_OldData5_3.Value, ",")
                    '--------------20081126 andy edit
                    Dim NewTeach() As String = Split(hid_NewData5_1.Value, ",")
                    Dim iType As Integer = If(NewTeach.Length > 1, 2, 1) '師資該堂課有2個 '師資該堂課有1個
                    If NewTeach.Length > 2 Then iType = 3 '師資該堂課有3個
                    For i As Integer = 0 To oldary.Length - 1
                        Select Case iType
                            Case 3
                                objrow("Teacher" & oldary(i)) = NewTeach(0)   '師資1
                                objrow("Teacher" & oldary(i) + 12) = NewTeach(1)  '依排課作業之規則 助教1 師資2=師資1 +12(欄位)
                                objrow("Teacher" & oldary(i) + 24) = NewTeach(2)  '依排課作業之規則 助教2 師資2=師資1 +24(欄位)
                            Case 2
                                objrow("Teacher" & oldary(i)) = NewTeach(0)   '師資1
                                objrow("Teacher" & oldary(i) + 12) = NewTeach(1)  '依排課作業之規則 助教1 師資2=師資1 +12(欄位)
                                objrow("Teacher" & oldary(i) + 24) = Convert.DBNull '
                            Case 1
                                objrow("Teacher" & oldary(i)) = NewTeach(0)   '師資1
                                objrow("Teacher" & oldary(i) + 12) = Convert.DBNull '
                                objrow("Teacher" & oldary(i) + 24) = Convert.DBNull '
                        End Select
                    Next
                End If
                DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

            Case Cst_i師資 '11 '產投-師資
                Dim s_TechTYPE As String = "A" 'TechTYPE: A:師資/B:助教
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '儲存 班級申請老師
                    If DataGrid21.Items.Count > 0 Then
                        Call SAVE_PLAN_TEACHER3(s_TechTYPE, oTrans) '修改 師資／助教-資料
                    Else
                        Call SAVE_OLD_PLAN_TEACHER2(s_TechTYPE, oTrans) '修改 師資／助教-資料
                    End If
                    Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                End If

            Case Cst_i助教 '20 '產投-'助教 
                Dim s_TechTYPE As String = "B" 'TechTYPE: A:師資/B:助教
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    '儲存 班級申請老師
                    If DataGrid22.Items.Count > 0 Then
                        Call SAVE_PLAN_TEACHER3(s_TechTYPE, oTrans) '修改 師資／助教-資料
                    Else
                        Call SAVE_OLD_PLAN_TEACHER2(s_TechTYPE, oTrans) '修改 師資／助教-資料
                    End If
                    Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                End If

            Case Cst_i報名日期 '17
                ' 20080825 andy  add 報名日期
                'Dim ocid AS String
                If ViewState(vs_OCID) <> "" Then
                    Dim dt As DataTable
                    objstr = "SELECT * FROM CLASS_CLASSINFO WHERE OCID='" & ViewState(vs_OCID) & "'"
                    dt = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                    If dt.Rows.Count <> 0 Then
                        Dim objrow As DataRow = dt.Rows(0) 'objrow=dt.Rows(0)
                        If New_SEnterDate.Text <> "" Then objrow("SEnterDate") = CDate(New_SEnterDate.Text)
                        If New_FEnterDate.Text <> "" Then objrow("FEnterDate") = CDate(New_FEnterDate.Text)
                        If Hid_sFENTERDATE2.Value <> "" Then objrow("FEnterDate2") = CDate(Hid_sFENTERDATE2.Value)

                        objrow("LastState") = "M" 'M: 修改(最後異動狀態)
                        objrow("ModifyAcct") = sm.UserInfo.UserID
                        objrow("ModifyDate") = Now()
                        DbAccess.UpdateDataTable(dt, objadapter, oTrans)
                    End If
                End If
            '20080522 Andy  "課程表"
            Case Cst_i課程表 '18
                '課程表
                Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
            'Dim sql AS String
            'Dim dt1, dt2 AS DataTable
            'Dim dr1, dr2 AS DataRow
            'Dim DateNow AS String=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
            '20080709 andy 更新產學訓課程表(課程表)--start
            '20080709 andy 更新產學訓課程表(課程表)  --end

            Case Cst_i上課時間 '15 '15:上課時間
                '上課時間
                'Dim sql AS String
                'Dim dt1, dt2, dtTemp AS DataTable
                'Dim dr, dr1, dr2 AS DataRow
                'Dim DateNow AS String=DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")
                'objstr="DELETE PLAN_ONCLASS WHERE PlanID=" & Request("PlanID") & " and ComIDNO='" & Request("cid") & "' and SeqNo=" & Request("no")
                objstr = ""
                objstr &= " DELETE PLAN_ONCLASS "
                objstr &= " WHERE PlanID=" & rPlanID & " AND ComIDNO='" & rComIDNO & "' AND SeqNo='" & rSeqNO & "'"
                DbAccess.ExecuteNonQuery(objstr, oTrans)
                If Session("Revise_OnClass") IsNot Nothing And DataGrid2.Items.Count > 0 Then
                    Dim dtTemp As DataTable = Session("Revise_OnClass")
                    objstr = " SELECT * FROM PLAN_ONCLASS WHERE 1<>1 "
                    objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                    For Each dr As DataRow In dtTemp.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        Dim iPOCID As Integer = DbAccess.GetNewId(oTrans, "PLAN_ONCLASS_POCID_SEQ,PLAN_ONCLASS,POCID")
                        Dim objrow As DataRow = Nothing
                        objrow = objtable.NewRow
                        objtable.Rows.Add(objrow)
                        objrow("POCID") = iPOCID
                        objrow("PlanID") = dr("PlanID")
                        objrow("ComIDNO") = dr("ComIDNO")
                        objrow("SeqNo") = dr("SeqNo")
                        objrow("Weeks") = dr("Weeks")
                        objrow("Times") = dr("Times")
                        objrow("ModifyAcct") = dr("ModifyAcct")
                        objrow("ModifyDate") = dr("ModifyDate")
                    Next
                    DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

                    'Call UPDATE_PLANONCLASS(oConn, oTrans)

                    'If (ViewState(vs_UpdateItem15)="Y") Then Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                    Call UPDATE_PLANTRAINDESC(ViewState(vs_PTDRID), oConn, oTrans)
                End If
                '----   Start
                '20081015 andy 更新產學訓課程表(上課時間) 
                '20081015 andy 更新產學訓課程表(上課時間)  
                '--   End
        End Select
        'DbAccess.CommitTrans(oTrans)
    End Sub

    ''' <summary>
    ''' 儲存1-SaveData1-SaveData1Y
    ''' </summary>
    Sub SaveData1()
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim uUrl1 As String = "TC/06/TC_06_001.aspx?ID=" & RqID

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(Convert.ToString(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Convert.ToString(Request("SubNo")))), 0)

        '通過 ／不通過(審核)
        Dim v_ChkMode As String = TIMS.GetListValue(ChkMode)

        Dim drPP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)
        ViewState(vs_OCID) = Convert.ToString(drPP("OCID"))

        Dim drCls As DataRow = Nothing 'TIMS.Get_ClassDaRow(rPlanID, rComIDNO, rSeqNO, objconn)
        If Convert.ToString(drPP("OCID")) <> "" Then
            drCls = TIMS.Get_ClassDaRow(rPlanID, rComIDNO, rSeqNO, objconn)
            ViewState(vs_OCID) = If(drCls IsNot Nothing, Convert.ToString(drCls("OCID")), "")
        End If

        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNO = "" OrElse rSCDate = "" OrElse iSubSeqNO = 0 Then
            Dim sMsg As String = "查無該計畫變更資料，請重新查詢!!"
            Call TIMS.BlockAlert(Me, sMsg, uUrl1)
            Return 'Exit Sub
        End If
        Dim drPCS As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)
        If drPCS Is Nothing Then
            Dim sMsg As String = "查無該計畫資料，請重新查詢!!"
            Common.MessageBox(Page, sMsg)
            Return 'Exit Sub
        End If
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Return 'Exit Sub
        End If
        '通過 ／不通過(審核)
        If v_ChkMode = "Y" Then
            Call Chk_SET_FENTERDATE(Errmsg, Val(hid_AltDataID.Value))
            If Errmsg <> "" Then
                Common.MessageBox(Page, Errmsg)
                Return 'Exit Sub
            End If
        End If

        Dim errFlag As Boolean = False
        Call TIMS.OpenDbConn(objconn)
        Hid_sFENTERDATE2.Value = ""
        If ViewState(vs_OCID) <> "" Then
            Select Case hid_AltDataID.Value
                Case Cst_i訓練期間.ToString() '"1"
                    If New_FEnterDate2.Text <> "" AndAlso New_Examdate.Text <> "" Then
                        Hid_RID1.Value = Convert.ToString(drCls("RID")).Substring(0, 1)
                        Dim sFENTERDATE As String = New_FEnterDate2.Text
                        Dim sEXAMDATE As String = New_Examdate.Text
                        Dim SS1 As String = ""
                        TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
                        Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
                        If sFENTERDATE2 <> "" Then Hid_sFENTERDATE2.Value = sFENTERDATE2
                    End If

                Case Cst_i報名日期.ToString() '"17"
                    Hid_RID1.Value = Convert.ToString(drCls("RID")).Substring(0, 1)
                    Dim sFENTERDATE As String = New_FEnterDate.Text
                    Dim sEXAMDATE As String = TIMS.Cdate3(drCls("EXAMDATE")) 'Examdate
                    Dim SS1 As String = ""
                    TIMS.SetMyValue(SS1, "RID1", Hid_RID1.Value) : TIMS.SetMyValue(SS1, "TPlanID", sm.UserInfo.TPlanID)
                    Dim sFENTERDATE2 As String = TIMS.GET_FENTERDATE2(SS1, sFENTERDATE, sEXAMDATE, objconn)
                    If sFENTERDATE2 <> "" Then Hid_sFENTERDATE2.Value = sFENTERDATE2
            End Select

        End If

        '通過(審核)
        Select Case v_ChkMode 'Convert.ToString(ChkMode.SelectedValue)
            Case "Y" '通過
                Select Case hid_AltDataID.Value
                    Case Cst_i訓練費用.ToString() '"21" 'Cst_訓練費用
                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "rPlanID", rPlanID)
                        TIMS.SetMyValue(sCmdArg, "rComIDNO", rComIDNO)
                        TIMS.SetMyValue(sCmdArg, "rSeqNo", rSeqNO)
                        TIMS.SetMyValue(sCmdArg, "rSCDate", rSCDate)
                        TIMS.SetMyValue(sCmdArg, "rSubSeqNO", iSubSeqNO)
                        Dim iRCID2 As Integer = TIMS.Get_REVC_RCID(sCmdArg, 2, objconn)
                        dtR2G = TIMS.GET_REVISE_COSTITEMdt(sCmdArg, iRCID2, 2, objconn)
                End Select
        End Select

        Dim pms_1 As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}, {"CDate3", TIMS.Cdate3(rSCDate)}, {"SubSeqNO", iSubSeqNO}}
        Dim sql_1 As String = ""
        sql_1 &= " SELECT *" & vbCrLf
        sql_1 &= " FROM PLAN_REVISE" & vbCrLf
        sql_1 &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
        sql_1 &= " AND CDate=@CDate3 AND SubSeqNO=@SubSeqNO"
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql_1, objconn, pms_1)
        Dim flag_data_ok As Boolean = If(dt1.Rows.Count = 0, False, If(dt1.Rows.Count <> 1, False, True))
        '有錯誤發生!'查無資料異常
        If Not flag_data_ok Then
            Common.MessageBox(Me, "審核失敗，請重新操作!!(查無資料)")
            Exit Sub
        End If

        'Dim v_ChkMode AS String=TIMS.GetListValue(ChkMode)
        Select Case v_ChkMode '通過 ／不通過
            Case "Y", "N"
            Case Else
                Common.MessageBox(Me, "審核失敗，請選擇通過或不通過!!")
                Exit Sub
        End Select

        Select Case v_ChkMode '通過 ／不通過 Y/N
            Case "Y" '通過

                Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
                Try
                    Call SaveData1Y(objconn, objTrans)
                    DbAccess.CommitTrans(objTrans)
                Catch ex As Exception
                    Dim strErrmsg As String = ""
                    strErrmsg &= String.Format("PlanID-ComIDNO-SeqNO:{0}-{1}-{2}", rPlanID, rComIDNO, rSeqNO) & vbCrLf
                    strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
                    strErrmsg &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
                    strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex)

                    errFlag = True
                    DbAccess.RollbackTrans(objTrans)
                    TIMS.CloseDbConn(objconn) 'Throw ex
                End Try
                '有錯誤發生!
                If errFlag Then
                    Common.MessageBox(Me, "審核失敗，請重新操作!!(發生異常)")
                    Exit Sub
                End If

                '資料-產投-審查計分表
                Call save_divCo128()

                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Call TIMS.OpenDbConn(objconn)
                    'rPlanID, rComIDNO, rSeqNO
                    Dim ssHash As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}}
                    'TIMS.SetMyValue2(ssHash, "PlanID", rPlanID) 'TIMS.SetMyValue2(ssHash, "ComIDNO", rComIDNO) 'TIMS.SetMyValue2(ssHash, "SeqNO", rSeqNO)
                    '修正- 'DISTANCE/FARLEARN 遠距教學 1.申請整班為遠距教學", 2."申請部分課程為遠距/實體教學", 3."申請整班為實體教學
                    Call UPDATE_PLANTRAINDESC_FARLEARN(ssHash, objconn)
                End If


            Case "N" '不通過
                ReviseCont.Text = TIMS.ClearSQM(ReviseCont.Text) '2字以上
                If ReviseCont.Text = "" OrElse ReviseCont.Text.Length < 2 Then
                    Common.MessageBox(Me, "審核不通過，請輸入審核說明!!")
                    Exit Sub
                End If
                Dim i_ReviseCont_maxlen As Integer = 250
                If ReviseCont.Text <> "" AndAlso ReviseCont.Text.Length > i_ReviseCont_maxlen Then
                    Common.MessageBox(Me, String.Format("審核不通過，輸入長度有誤!!({0})", i_ReviseCont_maxlen))
                    Exit Sub
                End If
                If ReviseCont.Text.Length > i_ReviseCont_maxlen Then ReviseCont.Text = ReviseCont.Text.Substring(0, i_ReviseCont_maxlen)

                'UPDATE PLAN_REVISE SET ReviseStatus=NULL,REVISEDATE=NULL,Reason=NULL,Verifier=NULL WHERE 1!=1
                'UPDATE PLAN_REVISE  'SET ReviseStatus=NULL,REVISEDATE=NULL,Reason=NULL,Verifier=NULL 
                'WHERE 1=1 And CONCAT(PLANID,'x',COMIDNO,'x',SEQNO)='5037x41173271x5' and altdataid='22' and ReviseStatus is not null

                Dim pms_2 As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}, {"CDate3", TIMS.Cdate3(rSCDate)}, {"SubSeqNO", iSubSeqNO}}
                Dim sql_2 As String = ""
                sql_2 &= " SELECT *" & vbCrLf
                sql_2 &= " FROM PLAN_REVISE" & vbCrLf
                sql_2 &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo"
                sql_2 &= " AND CDate=@CDate3 AND SubSeqNO=@SubSeqNO"
                'Dim objadapter As SqlDataAdapter=Nothing
                'Dim objTrans As SqlTransaction=DbAccess.BeginTrans(objconn)
                Dim objtable As DataTable = DbAccess.GetDataTable(sql_2, objconn, pms_2)
                Try
                    If objtable.Rows.Count > 0 Then
                        Dim pms_u2 As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}, {"CDate3", TIMS.Cdate3(rSCDate)}, {"SubSeqNO", iSubSeqNO}}
                        pms_u2.Add("Reason", ReviseCont.Text)
                        pms_u2.Add("Verifier", sm.UserInfo.UserID)
                        pms_u2.Add("ModifyAcct", sm.UserInfo.UserID)
                        Dim sql_u2 As String = ""
                        sql_u2 &= " UPDATE PLAN_REVISE" & vbCrLf
                        sql_u2 &= " SET ReviseStatus='N', REVISEDATE=GETDATE()" & vbCrLf
                        sql_u2 &= " ,Reason=@Reason ,Verifier=@Verifier,ModifyAcct=@ModifyAcct" & vbCrLf
                        sql_u2 &= " ,ModifyDate=GETDATE()" & vbCrLf '最後審核異動時間需記錄
                        sql_u2 &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf
                        sql_u2 &= " AND CDate=@CDate3 AND SubSeqNO=@SubSeqNO" & vbCrLf
                        DbAccess.ExecuteNonQuery(sql_u2, objconn, pms_u2)

                        '資料-產投-審查計分表
                        Call save_divCo128()
                    End If
                    'DbAccess.CommitTrans(objTrans)
                Catch ex As Exception
                    Dim strErrmsg As String = ""
                    strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex)
                    errFlag = True
                    'DbAccess.RollbackTrans(objTrans)
                    Common.MessageBox(Me, ex.ToString)
                    'Throw ex
                End Try
                '有錯誤發生!
                If errFlag Then
                    Common.MessageBox(Me, "審核失敗，請重新操作!!(發生異常)")
                    Exit Sub
                End If

        End Select

        'Add by jack 04/12/30
        'Session("_search")=ViewState("_search")
        If errFlag Then
            '有錯誤發生!
            Common.MessageBox(Me, "審核失敗，請重新操作!!(發生異常)")
            Exit Sub
        End If

        'Dim sChkmode_AltDataID AS Integer=Val(hid_AltDataID.Value)
        'Dim sScript1 AS String=""
        'If sChkmode_AltDataID=1 Then
        '    '訓練期間
        '    'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID)=-1 Then strMassage &="該班的排課審核已經還原，待訓練單位修改課表後，請記得於排課列表功能進行排課審核。\n" '非產投，職前班才顯示。
        '    sScript1="<script language=javascript>window.alert('" + strMassage + "');"
        '    sScript1 &="window.location.href='TC_06_001.aspx?ID=" & Request("ID") & "';</script>"
        'Else
        '    sScript1="<script language=javascript>window.alert('" + strMassage + "');"
        '    sScript1 &="window.location.href='TC_06_001.aspx?ID=" & Request("ID") & "';</script>"
        'End If
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(sScript1))
        'Common.MessageBox(Page, strMassage)
        'Exit Sub
        Dim strMassage As String = "審核成功!"
        Call TIMS.BlockAlert(Me, strMassage, uUrl1)
    End Sub

    ''' <summary> 儲存2-(顯示)-SaveData2 </summary>
    Sub SaveData2()
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim uUrl1 As String = "TC/06/TC_06_001.aspx?ID=" & RqID

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(Convert.ToString(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Convert.ToString(Request("SubNo")))), 0)

        Dim drCls As DataRow = TIMS.Get_ClassDaRow(rPlanID, rComIDNO, rSeqNO, objconn)
        ViewState(vs_OCID) = If(drCls IsNot Nothing, Convert.ToString(drCls("OCID")), "")
        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNO = "" OrElse rSCDate = "" OrElse iSubSeqNO = 0 Then
            'Dim sMsg AS String="查無該計畫變更資料，請重新查詢!!"
            'Common.RespWrite(Me, "<script language=javascript>window.alert('" & sMsg & "');")
            'Common.RespWrite(Me, "window.location.href='TC_06_001.aspx?ID=" & Request("ID") & "';</script>")
            'Common.MessageBox(Page, sMsg)
            'Exit Sub
            Dim sMsg As String = "查無該計畫變更資料，請重新查詢!!"
            Call TIMS.BlockAlert(Me, sMsg, uUrl1)
            Return 'Exit Sub
        End If
        Dim drPCS As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)
        If drPCS Is Nothing Then
            Dim sMsg As String = "查無該計畫資料，請重新查詢!!"
            Common.MessageBox(Page, sMsg)
            Return 'Exit Sub
        End If
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Return 'Exit Sub
        End If

        '儲存-產投-審查計分表
        Call SAVE_STATUS4()

        'Dim sChkmode_AltDataID AS Integer=Val(hid_AltDataID.Value)
        Common.MessageBox(Me, "儲存成功!")
        Dim s_Url1 As String = "TC_06_001.aspx?ID=" & RqID
        TIMS.Utl_Redirect(Me, objconn, s_Url1)

        'Const cst_SAVE_TXT1 AS String="儲存成功!"
        'Call TIMS.blockAlert(Me, cst_SAVE_TXT1, uUrl1)

        'Common.MessageBox(Page, cst_SAVE_TXT1)
        'sScript1="<script language=javascript>window.alert('" & cst_SAVE_TXT1 & "');"
        'sScript1 &="window.location.href='TC_06_001.aspx?ID=" & Request("ID") & "';</script>"
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(sScript1))

        'Dim sScript1 AS String=""
        'sScript1="<script language=javascript>"
        'sScript1 +="blockAlert('" & cst_SAVE_TXT1.Replace("\\n", "<br>") & "','',function(){ "
        'sScript1 +="location.href='TC/06/TC_06_001.aspx?ID=" & RqID & "';} "
        'sScript1 +=");"
        'sScript1 +="</script>"
        'Page.RegisterStartupScript(TIMS.xBlockName(), sScript1)
    End Sub

    '變更前上課時間
    Sub PlanClassTime()
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql &= " Select ROCID,PLANID,COMIDNO,SEQNO,Weeks,Times FROM REVISE_ONCLASS_OLD"
        sql &= " WHERE PlanID=@PlanID And ComIDNO=@ComIDNO And SeqNO=@SeqNO And SCDate=@SCDate And SubSeqNO=@SubSeqNO"
        sql &= " ORDER BY ROCID"
        Dim parms As New Hashtable
        TIMS.SetMyValue2(parms, "PlanID", Convert.ToString(rPlanID)) '計畫PK
        TIMS.SetMyValue2(parms, "ComIDNO", rComIDNO) '計畫PK
        TIMS.SetMyValue2(parms, "SeqNO", Convert.ToString(rSeqNO)) '計畫PK
        TIMS.SetMyValue2(parms, "SCDate", If(rSCDate <> "", rSCDate, Convert.DBNull)) '計畫PK
        TIMS.SetMyValue2(parms, "SubSeqNO", iSubSeqNO) '計畫PK
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then
            sql = ""
            sql &= " Select POCID,PLANID,COMIDNO,SEQNO,Weeks,Times FROM PLAN_ONCLASS"
            sql &= " WHERE PlanID=@PlanID And ComIDNO=@ComIDNO And SeqNO=@SeqNO"
            sql &= " ORDER BY POCID"
            Dim parms1 As New Hashtable
            TIMS.SetMyValue2(parms1, "PlanID", Convert.ToString(rPlanID)) '計畫PK
            TIMS.SetMyValue2(parms1, "ComIDNO", rComIDNO) '計畫PK
            TIMS.SetMyValue2(parms1, "SeqNO", Convert.ToString(rSeqNO)) '計畫PK
            dt = DbAccess.GetDataTable(sql, objconn, parms1)
        End If
        'sql=" SELECT * FROM Plan_OnClass WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNO & "' "
        'dt=DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    '變更後上課時間
    Sub ReviseClassTime()
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql = " SELECT * FROM REVISE_ONCLASS WHERE SCDate=" & TIMS.To_date(rSCDate) & " AND SubSeqNO=" & iSubSeqNO & " AND PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNO & "' "
        dt = DbAccess.GetDataTable(sql, objconn)
        Session("Revise_OnClass") = dt
        DataGrid2.DataSource = dt
        DataGrid2.DataBind()
    End Sub

    '變更前 包班種類 企業
    Private Sub PlanBusPackage()
        'Dim dt As DataTable = Nothing
        Dim pms1 As New Hashtable From {{"PlanID", TIMS.CINT1(rPlanID)}, {"ComIDNO", rComIDNO}, {"SeqNO", TIMS.CINT1(rSeqNO)}}
        Dim sql As String = ""
        sql &= " SELECT uname 企業名稱 ,intaxno 服務單位統一編號 ,ubno 保險證號" & vbCrLf
        sql &= " FROM PLAN_BUSPACKAGE" & vbCrLf
        sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNO=@SeqNO" & vbCrLf
        sql &= " ORDER BY BPID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, pms1)
        DG_BusPackageOld.DataSource = dt
        DG_BusPackageOld.DataBind()
    End Sub

    '變更後 包班種類 企業
    Sub ReviseBusPackage()
        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        sql &= " SELECT uname 企業名稱 ,intaxno 服務單位統一編號 ,ubno 保險證號" & vbCrLf
        sql &= " FROM Revise_BusPackage" & vbCrLf
        sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNO & "'" & vbCrLf
        sql &= " AND SCDate=" & TIMS.To_date(rSCDate) & vbCrLf
        sql &= " ORDER BY BPID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        DG_BusPackageNew.DataSource = dt
        DG_BusPackageNew.DataBind()

        'Session("Revise_BusPackage")=dt
        sql = "" & vbCrLf
        sql &= " SELECT *" & vbCrLf
        sql &= " FROM Revise_BusPackage" & vbCrLf
        sql &= " WHERE PlanID='" & rPlanID & "' AND ComIDNO='" & rComIDNO & "' AND SeqNO='" & rSeqNO & "'" & vbCrLf
        sql &= " AND SCDate=" & TIMS.To_date(rSCDate) & vbCrLf
        sql &= " ORDER BY BPID" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Session("Revise_BusPackage") = dt
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim OldWeeks1 As Label = e.Item.FindControl("OldWeeks1")
                Dim OldTimes1 As Label = e.Item.FindControl("OldTimes1")
                Dim drv As DataRowView = e.Item.DataItem
                OldWeeks1.Text = drv("Weeks").ToString
                OldTimes1.Text = drv("Times").ToString
            Case ListItemType.EditItem
                Dim OldWeeks2 As DropDownList = e.Item.FindControl("OldWeeks2")
                Dim OldTimes2 As TextBox = e.Item.FindControl("OldTimes2")
                Dim drv As DataRowView = e.Item.DataItem
                OldWeeks2 = TIMS.Get_ddlWeeks(OldWeeks2)
                Common.SetListItem(OldWeeks2, drv("Weeks").ToString)
                OldTimes2.Text = drv("Times").ToString
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim NewWeeks1 As Label = e.Item.FindControl("NewWeeks1")
                Dim NewTimes1 As Label = e.Item.FindControl("NewTimes1")
                Dim drv As DataRowView = e.Item.DataItem
                NewWeeks1.Text = drv("Weeks").ToString
                NewTimes1.Text = drv("Times").ToString
            Case ListItemType.EditItem
                Dim NewWeeks2 As DropDownList = e.Item.FindControl("NewWeeks2")
                Dim NewTimes2 As TextBox = e.Item.FindControl("NewTimes2")
                Dim drv As DataRowView = e.Item.DataItem
                NewWeeks2 = TIMS.Get_ddlWeeks(NewWeeks2)
                Common.SetListItem(NewWeeks2, drv("Weeks").ToString)
                NewTimes2.Text = drv("Times").ToString
        End Select
    End Sub

    '20080630 andy 產生課程表  start --
    Sub CreateTrainDesc()
        'Dim sql_1, sql_2, sql, AltPTDRDataID AS String
        'Dim dt, dt2 AS DataTable
        Dim dg4_dt As HtmlTableCell = FindControl("dg4_dt")
        Dim dg3_dt As HtmlTableCell = FindControl("dg3_dt")

        'rPlanID=TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        'rComIDNO=TIMS.ClearSQM(Convert.ToString(Request("cid")))
        'rSeqNO=TIMS.ClearSQM(Convert.ToString(Request("no")))
        'rSCDate=TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        'iSubSeqNO=If(Convert.ToString(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Convert.ToString(Request("SubNo")))), 0)

        hid_AltDataID.Value = TIMS.ClearSQM(hid_AltDataID.Value)
        If hid_AltDataID.Value = "" Then Return
        Dim AltPTDRDataID As String = hid_AltDataID.Value
        'Dim sql As String=""
        'sql=""
        'sql &=" SELECT AltDataID" & vbCrLf
        'sql &=" FROM PLAN_REVISE" & vbCrLf
        'sql &=" WHERE PlanID=" & rPlanID & vbCrLf
        'sql &=" AND ComIDNO='" & rComIDNO & "'" & vbCrLf
        'sql &=" AND SeqNo=" & rSeqNO & "" & vbCrLf
        'sql &=" AND CDate=" & TIMS.to_date(rSCDate) & vbCrLf
        'sql &=" AND SubSeqNO=" & iSubSeqNO & vbCrLf
        'AltPTDRDataID=DbAccess.ExecuteScalar(sql, objconn)

        Dim showlist As New ArrayList '產投用選項
        showlist.Add(Cst_i訓練期間)
        showlist.Add(Cst_i師資)
        showlist.Add(Cst_i助教)
        showlist.Add(Cst_i科場地)
        showlist.Add(Cst_i上課時間)
        showlist.Add(Cst_i課程表)
        showlist.Add(Cst_i遠距教學)
        Dim flag_showdg As Boolean = False '產投用選項
        For Each s_listVal As String In showlist
            If AltPTDRDataID = s_listVal Then flag_showdg = True
            If flag_showdg Then Exit For
        Next
        If Not flag_showdg Then
            dg4_dt.Style("display") = "none"
            dg3_dt.Style("display") = "none"
            Exit Sub
        End If

        '師資-'課程表 Then
        Dim iPTDRID As Integer = Get_PTDRID(rPlanID, rComIDNO, rSeqNO, rSCDate, iSubSeqNO, AltPTDRDataID)
        ViewState(vs_PTDRID) = iPTDRID
        Dim dt As DataTable = Nothing
        dt = Get_PlanTrainDescOldRevise(iPTDRID, AltPTDRDataID, cst_now) '課程表申請變更前
        If dt.Rows.Count = 0 Then dt = Get_PlanTrainDescOldRevise(iPTDRID, AltPTDRDataID, cst_old1) '課程表申請變更前

        Dim dt2 As DataTable = Nothing
        'iPTDRID=0 資訊異常，不提供變更後資訊
        If iPTDRID = 0 Then dt2 = dt Else dt2 = Get_PlanTrainDescNewRevise(iPTDRID) '課程表申請變更後

        If dt Is Nothing Then
            Datagrid4.Visible = False
            'Datagrid4.Style.Item("display")="none"
            dg4_dt.Style("display") = "none"
            Datagrid3.Visible = False
            'Datagrid3.Style.Item("display")="none"
            dg3_dt.Style("display") = "none"
        Else
            Dim fg_show_EHour As Boolean = If(hid_TMID.Value = TIMS.cst_EHour_Use_TMID, True, False)

            '課程表申請變更後
            Datagrid4.Visible = True
            Datagrid4.Columns(cst_DG4_EHour_技檢訓練時數_iCOL).Visible = fg_show_EHour
            Datagrid4.DataSource = dt2 '(變更後)
            Datagrid4.DataBind()

            '課程表申請變更前
            Datagrid3.Visible = True
            Datagrid3.Columns(cst_DG3_EHour_技檢訓練時數_iCOL).Visible = fg_show_EHour
            Datagrid3.DataSource = dt '(變更前)
            Datagrid3.DataBind()
        End If
    End Sub

    Private Sub Datagrid4_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drData As DataRowView = e.Item.DataItem '產投用選項
                Dim labSTrainDate As Label = e.Item.FindControl("lab_STrainDate")
                Dim OldTPERIOD28_1t As CheckBox = e.Item.FindControl("OldTPERIOD28_1t")
                Dim OldTPERIOD28_2t As CheckBox = e.Item.FindControl("OldTPERIOD28_2t")
                Dim OldTPERIOD28_3t As CheckBox = e.Item.FindControl("OldTPERIOD28_3t")
                Dim labPName As Label = e.Item.FindControl("lab_PName")
                Dim labPHour As Label = e.Item.FindControl("lab_PHour") '時數
                Dim labEHour As Label = e.Item.FindControl("lab_EHour") '技檢訓練時數
                Dim txtPCont As TextBox = e.Item.FindControl("txt_PCont")
                Dim listClassification As DropDownList = e.Item.FindControl("list_Classification")
                Dim labPTID As Label = e.Item.FindControl("lab_PTID")
                Dim newTechText As TextBox = e.Item.FindControl("newTechText")
                Dim newTech2Text As TextBox = e.Item.FindControl("newTech2Text")
                'Cst_i遠距教學
                Dim bx_FARLEARN As CheckBox = e.Item.FindControl("bx_FARLEARN")
                Dim hide_FARLEARN As HtmlInputHidden = e.Item.FindControl("hide_FARLEARN")
                bx_FARLEARN.Checked = If(Convert.ToString(drData("FARLEARN")).Equals("Y"), True, False)
                hide_FARLEARN.Value = Convert.ToString(drData("FARLEARN"))

                labSTrainDate.Text = Convert.ToString(drData("STrainDate"))
                Dim str_v_TPERIOD28 As String = If(Convert.ToString(drData("TPERIOD28")).Length >= 3, Convert.ToString(drData("TPERIOD28")), cst_NNN) '"NNN"
                OldTPERIOD28_1t.Checked = If(str_v_TPERIOD28.Substring(0, 1) = "Y", True, False)
                OldTPERIOD28_2t.Checked = If(str_v_TPERIOD28.Substring(1, 1) = "Y", True, False)
                OldTPERIOD28_3t.Checked = If(str_v_TPERIOD28.Substring(2, 1) = "Y", True, False)
                labPName.Text = Convert.ToString(drData("PName"))
                labPHour.Text = Convert.ToString(drData("PHour")) '時數
                labEHour.Text = Convert.ToString(drData("EHour")) '技檢訓練時數
                'txtPCont.Text=Convert.ToString(drData("PCont"))
                txtPCont.Text = TIMS.ClearSQM(Convert.ToString(drData("PCont")))
                'listClassification.SelectedValue=Convert.ToString(drData("Classification1"))
                If Convert.ToString(drData("Classification1")).Trim <> "" Then Common.SetListItem(listClassification, Convert.ToString(drData("Classification1")))
                If IsNumeric(Convert.ToString(drData("PTID"))) = True Then labPTID.Text = Get_TrainPlaceName(Convert.ToString(drData("PTID")), objconn)
                If IsNumeric(Convert.ToString(drData("TechID"))) = True Then newTechText.Text = TIMS.Get_TeachCName(drData("TechID"), objconn) 'TIMS.Get_TeacherName(Convert.ToString(drData("TechID")))
                If IsNumeric(Convert.ToString(drData("TechID2"))) = True Then newTech2Text.Text = TIMS.Get_TeachCName(drData("TechID2"), objconn) 'TIMS.Get_TeacherName(Convert.ToString(drData("TechID2")))
                'If IsNumeric(Convert.ToString(drData("TechID")))=True Then labTechID.Text=TIMS.Get_TeachCName(drData("TechID"), objconn) 'TIMS.Get_TeacherName(Convert.ToString(drData("TechID")))
                'If IsNumeric(Convert.ToString(drData("TechID2")))=True Then labTechID2.Text=TIMS.Get_TeachCName(drData("TechID2"), objconn) 'TIMS.Get_TeacherName(Convert.ToString(drData("TechID2")))
        End Select
    End Sub

    ''' <summary>
    ''' 課程表申請變更前
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Datagrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'Show
                Dim drv As DataRowView = e.Item.DataItem '產投用選項
                Dim OldSTrainDateLabel As Label = e.Item.FindControl("OldSTrainDateLabel")
                Dim OldTPERIOD28_1t As CheckBox = e.Item.FindControl("OldTPERIOD28_1t")
                Dim OldTPERIOD28_2t As CheckBox = e.Item.FindControl("OldTPERIOD28_2t")
                Dim OldTPERIOD28_3t As CheckBox = e.Item.FindControl("OldTPERIOD28_3t")
                Dim OldPNameLabel As Label = e.Item.FindControl("OldPNameLabel")
                Dim OldPHourLabel As Label = e.Item.FindControl("OldPHourLabel") '時數
                Dim OldEHourLabel As Label = e.Item.FindControl("OldEHourLabel") '技檢訓練時數
                Dim OldPContText As TextBox = e.Item.FindControl("OldPContText")
                Dim OlddrpClassification1 As DropDownList = e.Item.FindControl("OlddrpClassification1")
                Dim OlddrpPTID As DropDownList = e.Item.FindControl("OlddrpPTID")
                Dim OldTech1Value As HtmlInputHidden = e.Item.FindControl("OldTech1Value")
                Dim OldTech1Text As TextBox = e.Item.FindControl("OldTech1Text")
                Dim OldTech2Value As HtmlInputHidden = e.Item.FindControl("OldTech2Value")
                Dim OldTech2Text As TextBox = e.Item.FindControl("OldTech2Text")
                'Cst_i遠距教學
                Dim OldFARLEARN As CheckBox = e.Item.FindControl("OldFARLEARN")
                OldFARLEARN.Checked = If(Convert.ToString(drv("FARLEARN")).Equals("Y"), True, False)

                'If TIMS.ConvertStr(drv("STrainDate")) <> "" Then OldSTrainDateLabel.Text=TIMS.ConvertStr(drv("STrainDate"))
                If Convert.ToString(drv("STrainDate")) <> "" Then OldSTrainDateLabel.Text = TIMS.Cdate3(drv("STrainDate"))
                Dim str_v_TPERIOD28 As String = If(Convert.ToString(drv("TPERIOD28")).Length >= 3, Convert.ToString(drv("TPERIOD28")), cst_NNN) '"NNN"
                OldTPERIOD28_1t.Checked = If(str_v_TPERIOD28.Substring(0, 1) = "Y", True, False)
                OldTPERIOD28_2t.Checked = If(str_v_TPERIOD28.Substring(1, 1) = "Y", True, False)
                OldTPERIOD28_3t.Checked = If(str_v_TPERIOD28.Substring(2, 1) = "Y", True, False)

                OldPNameLabel.Text = drv("PName").ToString
                OldPHourLabel.Text = Convert.ToString(drv("PHour")) '時數
                OldEHourLabel.Text = Convert.ToString(drv("EHour")) '技檢訓練時數
                'OldPContText.Text=drv("PCont").ToString
                OldPContText.Text = TIMS.ClearSQM(Convert.ToString(drv("PCont")))
                If drv("Classification1").ToString <> "" Then
                    Common.SetListItem(OlddrpClassification1, drv("Classification1").ToString)
                    Dim v_OlddrpClassification1 As String = TIMS.GetListValue(OlddrpClassification1)
                    Select Case v_OlddrpClassification1 'OlddrpClassification1.SelectedValue
                        Case "1" '學科
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "cid", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                            OlddrpPTID = TIMS.Get_SciPTID(OlddrpPTID, Hid_ComIDNO.Value, 3, objconn)
                        Case "2" '術科
                            Hid_ComIDNO.Value = TIMS.sUtl_GetRqValue(Me, "cid", Hid_ComIDNO.Value)
                            If Hid_ComIDNO.Value = "" Then Hid_ComIDNO.Value = TIMS.Get_ComIDNOforOrgID(sm.UserInfo.OrgID, objconn)
                            OlddrpPTID = TIMS.Get_TechPTID(OlddrpPTID, Hid_ComIDNO.Value, 3, objconn)
                    End Select
                    If Convert.ToString(drv("PTID")) <> "" AndAlso OlddrpPTID.SelectedIndex <> -1 Then
                        Common.SetListItem(OlddrpPTID, Convert.ToString(drv("PTID")))
                        '整理下拉，只留選擇值
                        TIMS.GET_NewListItemVal(OlddrpPTID, Convert.ToString(drv("PTID")))
                    End If
                End If
                If TIMS.ConvertStr(drv("TechID")) <> "" Then
                    OldTech1Value.Value = Convert.ToString(drv("TechID")) 'drv("TechID").ToString
                    OldTech1Text.Text = TIMS.Get_TeachCName(drv("TechID"), objconn) ' TIMS.Get_TeacherName(drv("TechID").ToString)
                End If
                If TIMS.ConvertStr(drv("TechID2")) <> "" Then
                    OldTech2Value.Value = Convert.ToString(drv("TechID2")) 'drv("TechID2").ToString
                    OldTech2Text.Text = TIMS.Get_TeachCName(drv("TechID2"), objconn) 'TIMS.Get_TeacherName(drv("TechID2").ToString)
                End If
        End Select
    End Sub

    '補充增加字句
    Function getStr(ByVal oldstr As String, ByVal addstr As String) As String
        If oldstr <> "" Then oldstr &= ","
        oldstr &= addstr
        Return oldstr
    End Function

    Private Sub DataGrid21O_1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_1.ItemDataBound, DataGrid21New_1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim CostName As Label = e.Item.FindControl("CostName")
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim ItemCost As Label = e.Item.FindControl("ItemCost")
                Dim subtotal As Label = e.Item.FindControl("subtotal")
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg

                CostName.Text = ""
                Select Case Convert.ToString(drv("CostID"))
                    Case "99"
                        strTMP1 = "其他-" & Convert.ToString(drv("ItemOther")).ToString
                    Case Else
                        ff33 = "CostID='" & drv("CostID") & "'"
                        If dt_Key_CostItem.Select(ff33).Length > 0 Then strTMP1 = dt_Key_CostItem.Select(ff33)(0)("CostName")
                End Select
                CostName.Text = strTMP1
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                ItemCost.Text = Convert.ToString(drv("ItemCost"))
                'subtotal.Text=TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
        End Select
    End Sub

    Private Sub DataGrid21O_2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_2.ItemDataBound, DataGrid21New_2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim ItemCost As Label = e.Item.FindControl("ItemCost")
                Dim subtotal As Label = e.Item.FindControl("subtotal")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg

                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                ItemCost.Text = Convert.ToString(drv("ItemCost"))
                'subtotal.Text=TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
        End Select
    End Sub

    Private Sub DataGrid21O_3_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_3.ItemDataBound, DataGrid21New_3.ItemDataBound
        'e.Item.Cells(4).Style("display")="none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim subtotal As Label = e.Item.FindControl("subtotal")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg

                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                'subtotal.Text=TIMS.Round(CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost")))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
        End Select
    End Sub

    Private Sub DataGrid21O_4_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid21Old_4.ItemDataBound, DataGrid21New_4.ItemDataBound
        'e.Item.Cells(5).Style("display")="none"
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim CostName As Label = e.Item.FindControl("CostName")
                Dim OPrice As Label = e.Item.FindControl("OPrice")
                Dim Itemage As Label = e.Item.FindControl("Itemage")
                Dim subtotal As Label = e.Item.FindControl("subtotal")

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PCID", Convert.ToString(drv("PCID")))
                Dim btnDel1 As Button = e.Item.FindControl("btnDel1")
                If Not btnDel1 Is Nothing Then btnDel1.CommandArgument = sCmdArg

                CostName.Text = ""
                Select Case Convert.ToString(drv("CostID"))
                    Case "99"
                        strTMP1 = "其他-" & Convert.ToString(drv("ItemOther")).ToString
                    Case Else
                        ff33 = "CostID='" & drv("CostID") & "'"
                        If dt_Key_CostItem.Select(ff33).Length > 0 Then strTMP1 = dt_Key_CostItem.Select(ff33)(0)("CostName")
                End Select
                CostName.Text = strTMP1
                OPrice.Text = Convert.ToString(drv("OPrice"))
                Itemage.Text = Convert.ToString(drv("Itemage"))
                subtotal.Text = CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
        End Select
    End Sub

    '取得申請變更的資訊
    Sub SHOW_REVISE_COSTITEM(ByVal dt3 As DataTable, ByVal iRCID As Integer, ByVal iO1N2 As Integer)
        Dim iPlanKind As Integer = Val(Hid_PlanKind.Value) '
        Dim iCostMode As Integer = Val(Hid_CostMode.Value) '計價方案'計價種類
        Select Case iO1N2
            Case 1
                Dim objDG_Old As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_Old = DataGrid21Old_1
                Else
                    Select Case iCostMode
                        Case 2
                            objDG_Old = DataGrid21Old_2 '每人每時單價計價法
                        Case 3
                            objDG_Old = DataGrid21Old_3 '每人輔助單價計價法
                        Case 4
                            objDG_Old = DataGrid21Old_4 '個人單價計價法
                    End Select
                End If
                objDG_Old.Visible = True
                With objDG_Old
                    .DataSource = dt3
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                'Session(Hid_COSTITEM_GUID21.Value)=dt3
                Call SHOW_COSTITEM_GUID21(1, dt3)
            Case 2
                Dim objDG_New As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_New = DataGrid21New_1
                    objDG_New.Columns(5).Visible = False
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_New = DataGrid21New_2
                            objDG_New.Columns(4).Visible = False
                        Case 3
                            '每人輔助單價計價法
                            objDG_New = DataGrid21New_3
                            objDG_New.Columns(3).Visible = False
                        Case 4
                            '個人單價計價法
                            objDG_New = DataGrid21New_4
                            objDG_New.Columns(4).Visible = False
                    End Select
                End If
                objDG_New.Visible = True
                With objDG_New
                    .DataSource = dt3
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                'Session(Hid_COSTITEM_GUID21.Value)=dt3
                Call SHOW_COSTITEM_GUID21(2, dt3)
        End Select
    End Sub

    '計畫經費項目檔(PLAN_COSTITEM) 依SESSION 
    Sub SHOW_COSTITEM_GUID21(ByVal iO1N2 As Integer, ByRef dt1 As DataTable)
        'If Session(Hid_COSTITEM_GUID21.Value) Is Nothing Then Exit Sub
        'Dim dt1 AS DataTable=Session(Hid_COSTITEM_GUID21.Value)

        Dim iPlanKind As Integer = Val(Hid_PlanKind.Value)
        Dim iCostMode As Integer = Val(Hid_CostMode.Value) '計價方案'計價種類
        Dim iAdmPercent As Integer = Val(Hid_AdmPercent.Value) '行政管理費 '(行政管理費百分比)
        Dim iTaxPercent As Integer = Val(Hid_TaxPercent.Value) '營業稅 '(營業稅費用百分比)

        Dim ff As String = ""
        Dim diTotal As Double = 0 '總費用(浮點數)
        Dim diAdmTotal As Double = 0 '行政管理費 '(行政管理費百分比)(浮點數)
        Dim diTaxTotal As Double = 0 '營業稅 '(營業稅費用百分比)(浮點數)
        Dim AdmCostText As String = "" '顯示文字
        Dim TaxCostText As String = "" '顯示文字

        Const cst_Plankind1t As String = "費用列表"
        Const cst_CostMode2t As String = "每人每時計價"
        Const cst_CostMode3t As String = "每人輔助計價"
        Const cst_CostMode4t As String = "個人單價計價"

        Select Case iO1N2
            Case 1
                'PlanKind (2) 1:自辦 2:委外
                If iPlanKind = 1 Then
                    labcost21txt1Old.Text = cst_Plankind1t
                Else
                    Select Case iCostMode
                        Case 2 '每人每時單價計價法-2:委外
                            labcost21txt1Old.Text = cst_CostMode2t
                        Case 3 '每人輔助單價計價法-2:委外
                            labcost21txt1Old.Text = cst_CostMode3t
                        Case 4 '個人單價計價法-2:委外
                            labcost21txt1Old.Text = cst_CostMode4t
                    End Select
                End If

                DataGrid21Old_1.Visible = False
                DataGrid21Old_2.Visible = False
                DataGrid21Old_3.Visible = False
                DataGrid21Old_4.Visible = False
                Dim objDG_Old As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_Old = DataGrid21Old_1
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_Old = DataGrid21Old_2
                        Case 3
                            '每人輔助單價計價法
                            objDG_Old = DataGrid21Old_3
                        Case 4
                            '個人單價計價法
                            objDG_Old = DataGrid21Old_4
                    End Select
                End If
                objDG_Old.Visible = True
                With objDG_Old
                    .DataSource = dt1
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
            'Session(Hid_COSTITEM_GUID21.Value)=dt1
            Case 2
                'PlanKind (2) 1:自辦 2:委外
                If iPlanKind = 1 Then
                    labcost21txt1New.Text = cst_Plankind1t
                Else
                    Select Case iCostMode
                        Case 2 '每人每時單價計價法-2:委外
                            labcost21txt1New.Text = cst_CostMode2t
                        Case 3 '每人輔助單價計價法-2:委外
                            labcost21txt1New.Text = cst_CostMode3t
                        Case 4 '個人單價計價法-2:委外
                            labcost21txt1New.Text = cst_CostMode4t
                    End Select
                End If

                DataGrid21New_1.Visible = False
                DataGrid21New_2.Visible = False
                DataGrid21New_3.Visible = False
                DataGrid21New_4.Visible = False
                Dim objDG_New As DataGrid = Nothing
                If iPlanKind = 1 Then
                    objDG_New = DataGrid21New_1
                Else
                    Select Case iCostMode
                        Case 2
                            '每人每時單價計價法
                            objDG_New = DataGrid21New_2
                        Case 3
                            '每人輔助單價計價法
                            objDG_New = DataGrid21New_3
                        Case 4
                            '個人單價計價法
                            objDG_New = DataGrid21New_4
                    End Select
                End If
                objDG_New.Visible = True
                With objDG_New
                    .DataSource = dt1
                    .DataKeyField = "PCID"
                    .DataBind()
                End With
                'Session(Hid_COSTITEM_GUID21.Value)=dt1
        End Select

        Select Case iO1N2
            Case 1
                AdmGrantTROld.Visible = False
                TaxGrantTROld.Visible = False
            Case 2
                AdmGrantTRNew.Visible = False
                TaxGrantTRNew.Visible = False
        End Select

        If iPlanKind = 1 Then
            '1:自辦  'PlanKind (2) 1:自辦 2:委外
            '行政管理費 '(行政管理費百分比)
            If iAdmPercent > -1 Then
                Dim aFlag As Boolean = False 'false:未啟用行政管理費百分比
                ff = "AdmFlag='Y'"
                If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用行政管理費百分比
                If aFlag Then
                    Dim strTMP1 As String = ""
                    Dim fff As String = "AdmFlag='Y'"
                    For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                        diAdmTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                        If strTMP1 <> "" Then strTMP1 &= "+"
                        Select Case Convert.ToString(drv("CostID"))
                            Case "99"
                                strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                            Case Else
                                ff = "CostID='" & drv("CostID") & "'"
                                If dt_Key_CostItem.Select(ff).Length > 0 Then strTMP1 &= dt_Key_CostItem.Select(ff)(0)("CostName")
                        End Select
                    Next
                    AdmCostText = "(" & strTMP1 & ")*" & iAdmPercent & "%=" & TIMS.ROUND(diAdmTotal * iAdmPercent / 100)
                    Select Case iO1N2
                        Case 1
                            AdmGrantTROld.Visible = True 'true:顯示 
                            AdmCostOld.Text = AdmCostText
                        Case 2
                            AdmGrantTRNew.Visible = True 'true:顯示 
                            AdmCostNew.Text = AdmCostText
                    End Select
                End If
            End If

            '營業稅 '(營業稅費用百分比)
            If iTaxPercent > -1 Then
                Dim aFlag As Boolean = False 'false:未啟用 營業稅費用百分比
                ff = "TaxFlag='Y'"
                If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用 營業稅費用百分比
                If aFlag Then
                    Dim strTMP1 As String = ""
                    Dim fff As String = "TaxFlag='Y'"
                    For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                        diTaxTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                        If strTMP1 <> "" Then strTMP1 &= "+"
                        Select Case Convert.ToString(drv("CostID"))
                            Case "99"
                                strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                            Case Else
                                ff = "CostID='" & drv("CostID") & "'"
                                If dt_Key_CostItem.Select(ff).Length > 0 Then strTMP1 &= dt_Key_CostItem.Select(ff)(0)("CostName")
                        End Select
                    Next
                    TaxCostText = "(" & strTMP1 & ")*" & iTaxPercent & "%=" & TIMS.ROUND(diTaxTotal * iTaxPercent / 100)
                    Select Case iO1N2
                        Case 1
                            TaxGrantTROld.Visible = True 'true:顯示 
                            TaxCostOld.Text = TaxCostText
                        Case 2
                            TaxGrantTRNew.Visible = True 'true:顯示 
                            TaxCostNew.Text = TaxCostText
                    End Select
                End If
            End If

            diTotal = 0
            For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
            Next
            Select Case iO1N2
                Case 1
                    '行政管理費 '(行政管理費百分比)
                    If AdmGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                    '營業稅 '(營業稅費用百分比)
                    If TaxGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                    TotalCost1Old.Text = TIMS.ROUND(diTotal)
                Case 2
                    '行政管理費 '(行政管理費百分比)
                    If AdmGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                    '營業稅 '(營業稅費用百分比)
                    If TaxGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                    TotalCost1New.Text = TIMS.ROUND(diTotal)
            End Select

        Else
            'PlanKind (2) 1:自辦 2:委外
            Select Case iCostMode
                Case 2 '每人每時單價計價法-2:委外
                    diTotal = 0
                    For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                    Next
                    diTotal = CDbl(TIMS.ROUND(diTotal))
                Case 3 '每人輔助單價計價法-2:委外
                    diTotal = 0
                    For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
                    Next
                    diTotal = CDbl(TIMS.ROUND(diTotal))
                Case 4 '個人單價計價法-2:委外
                    '行政管理費 '(行政管理費百分比)
                    If iAdmPercent > -1 Then
                        Dim aFlag As Boolean = False 'false:未啟用行政管理費百分比
                        ff = "AdmFlag='Y'"
                        If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用行政管理費百分比
                        If aFlag Then
                            Dim strTMP1 As String = ""
                            Dim fff As String = "AdmFlag='Y'"
                            For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                                diAdmTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                                If strTMP1 <> "" Then strTMP1 &= "+"
                                Select Case Convert.ToString(drv("CostID"))
                                    Case "99"
                                        strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                                    Case Else
                                        ff = "CostID='" & drv("CostID") & "'"
                                        If dt_Key_CostItem.Select(ff).Length > 0 Then strTMP1 &= dt_Key_CostItem.Select(ff)(0)("CostName")
                                End Select
                            Next
                            AdmCostText = "(" & strTMP1 & ")*" & iAdmPercent & "%=" & TIMS.ROUND(diAdmTotal * iAdmPercent / 100)
                            Select Case iO1N2
                                Case 1
                                    AdmGrantTROld.Visible = True 'true:顯示 
                                    AdmCostOld.Text = TaxCostText
                                Case 2
                                    AdmGrantTRNew.Visible = True 'true:顯示 
                                    AdmCostNew.Text = AdmCostText
                            End Select
                        End If
                    End If

                    '營業稅 '(營業稅費用百分比)
                    If iTaxPercent > -1 Then
                        Dim aFlag As Boolean = False 'false:未啟用 營業稅費用百分比
                        ff = "TaxFlag='Y'"
                        If dt1.Select(ff, Nothing, DataViewRowState.CurrentRows).Length > 0 Then aFlag = True 'true:啟用 營業稅費用百分比
                        If aFlag Then
                            Dim strTMP1 As String = ""
                            Dim fff As String = "TaxFlag='Y'"
                            For Each drv As DataRow In dt1.Select(fff, Nothing, DataViewRowState.CurrentRows)
                                diTaxTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage")) * CDbl(drv("ItemCost"))
                                If strTMP1 <> "" Then strTMP1 &= "+"
                                Select Case Convert.ToString(drv("CostID"))
                                    Case "99"
                                        strTMP1 &= "其他-" & Convert.ToString(drv("ItemOther")).ToString
                                    Case Else
                                        ff = "CostID='" & drv("CostID") & "'"
                                        If dt_Key_CostItem.Select(ff).Length > 0 Then strTMP1 &= dt_Key_CostItem.Select(ff)(0)("CostName")
                                End Select
                            Next
                            TaxCostText = "(" & strTMP1 & ")*" & iTaxPercent & "%=" & TIMS.ROUND(diTaxTotal * iTaxPercent / 100)

                            Select Case iO1N2
                                Case 1
                                    TaxGrantTROld.Visible = True 'true:顯示 
                                    TaxCostOld.Text = TaxCostText
                                Case 2
                                    TaxGrantTRNew.Visible = True 'true:顯示 
                                    TaxCostNew.Text = TaxCostText
                            End Select

                        End If
                    End If

                    diTotal = 0
                    For Each drv As DataRow In dt1.Select(Nothing, Nothing, DataViewRowState.CurrentRows)
                        diTotal += CDbl(drv("OPrice")) * CDbl(drv("Itemage"))
                    Next
                    Select Case iO1N2
                        Case 1
                            '行政管理費 '(行政管理費百分比)
                            If AdmGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                            '營業稅 '(營業稅費用百分比)
                            If TaxGrantTROld.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                            TotalCost1Old.Text = TIMS.ROUND(diTotal)
                        Case 2
                            '行政管理費 '(行政管理費百分比)
                            If AdmGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diAdmTotal * iAdmPercent / 100))
                            '營業稅 '(營業稅費用百分比)
                            If TaxGrantTRNew.Visible Then diTotal += CDbl(TIMS.ROUND(diTaxTotal * iTaxPercent / 100))
                            TotalCost1New.Text = TIMS.ROUND(diTotal)
                    End Select
            End Select
        End If
    End Sub

#Region "(No Use)"
    '產投用選項
    'Private Function Get_PlanTrainDesc(ByVal planid AS Integer, ByVal comidno AS String, ByVal seqno AS Integer) AS DataTable
    '    Dim rst AS New DataTable '=Nothing
    '    Dim sqlStr AS String=""
    '    sqlStr=""
    '    sqlStr &=" select null AS PTDRID,PTDID" & vbCrLf
    '    sqlStr &=" ,null AS ID1, CONVERT(varchar, STrainDate, 111) STrainDate" & vbCrLf
    '    sqlStr &=" ,null AS ID2,PName" & vbCrLf
    '    sqlStr &=" ,null AS ID3,PHour" & vbCrLf
    '    sqlStr &=" ,PCont,Classification1" & vbCrLf
    '    sqlStr &=" ,null AS ID4,PTID" & vbCrLf
    '    sqlStr &=" ,null AS ID5,TechID" & vbCrLf
    '    sqlStr &=" ,null AS ID6,TechID2" & vbCrLf
    '    sqlStr &=" from Plan_TrainDesc" & vbCrLf
    '    sqlStr &=" where PlanID=@planid and ComIDNO=@comidno and SeqNO=@seqno" & vbCrLf
    '    sqlStr &=" ORDER BY STrainDate,PName asc "
    '    Dim sCmd AS New SqlCommand(sqlStr, objconn)
    '    Call TIMS.OpenDbConn(objconn)
    '    With sCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("planid", SqlDbType.Int).Value=planid
    '        .Parameters.Add("comidno", SqlDbType.VarChar).Value=comidno
    '        .Parameters.Add("seqno", SqlDbType.Int).Value=seqno
    '        rst.Load(.ExecuteReader())
    '    End With
    '    Return rst
    'End Function
#End Region

#Region "課程表Function"

    '課程表申請變更前
    Private Function Get_PlanTrainDescOldRevise(ByVal tPTDRid As Integer, ByVal sAltDataID As String, ByVal sType As String) As DataTable
        'sType: 'sType@now新版課表(產投) 'sType@old1舊1課表(產投)
        Dim rst As New DataTable '=Nothing
        'PLAN_TRAINDESC_REVISEITEM '(1.STrainDate 2.PName 3.PHour/9.EHour 4.PTID 5.TechID 6.TechID2,7.TPERIOD28,8.FARLEARN)
        'Dim da AS SqlDataAdapter=TIMS.GetOneDA(gobjconn)
        Dim sql As String = ""
        sql &= " select a.PTDRID,g.PTDID" & vbCrLf
        sql &= " ,b.PTDRIID AS ID1,convert(varchar, ISNULL(convert(datetime, b.OldData, 111),g.STrainDate), 111) STrainDate" & vbCrLf
        sql &= " ,c.PTDRIID AS ID2,ISNULL(c.OldData,g.PName) AS PName" & vbCrLf
        sql &= " ,d.PTDRIID AS ID3 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d.OldData),g.PHour)) AS PHour" & vbCrLf '時數
        sql &= " ,d9.PTDRIID AS ID9 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d9.OldData),g.EHour)) AS EHour" & vbCrLf '技檢訓練時數
        sql &= " ,d7.PTDRIID AS ID7 ,ISNULL(d7.OldData,g.TPERIOD28) TPERIOD28" & vbCrLf
        sql &= " ,d8.PTDRIID AS ID8 ,ISNULL(d8.OldData,g.FARLEARN) FARLEARN" & vbCrLf
        sql &= " ,g.PCont,g.Classification1" & vbCrLf
        sql &= " ,e.PTDRIID AS ID4,ISNULL(e.OldData,g.PTID) AS PTID" & vbCrLf
        sql &= " ,f.PTDRIID AS ID5,ISNULL(f.OldData,g.TechID) AS TechID" & vbCrLf
        sql &= " ,f2.PTDRIID AS ID6,ISNULL(f2.OldData,g.TechID2) AS TechID2" & vbCrLf
        sql &= " FROM PLAN_TRAINDESC_REVISE a" & vbCrLf '(1)
        Select Case sType
            Case cst_now
                sql &= " JOIN PLAN_TRAINDESC_RO g ON g.PlanID=a.PlanID AND g.ComIDNO=a.ComIDNO AND g.SeqNO=a.SeqNO AND g.PTDRID=a.PTDRID AND a.PTDRID=@PTDRid" & vbCrLf
            Case Else 'cst_old1
                sql &= " JOIN PLAN_TRAINDESC g ON g.PlanID=a.PlanID AND g.ComIDNO=a.ComIDNO AND g.SeqNO=a.SeqNO AND a.PTDRID=@PTDRid" & vbCrLf
        End Select
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM b on b.PTDID=g.PTDID and b.PTDRID=a.PTDRID and b.AltDataItem=1 and b.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM c on c.PTDID=g.PTDID and c.PTDRID=a.PTDRID and c.AltDataItem=2 and c.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d on d.PTDID=g.PTDID and d.PTDRID=a.PTDRID and d.AltDataItem=3 and d.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM e on e.PTDID=g.PTDID and e.PTDRID=a.PTDRID and e.AltDataItem=4 and e.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f on f.PTDID=g.PTDID and f.PTDRID=a.PTDRID and f.AltDataItem=5 and f.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f2 on f2.PTDID=g.PTDID and f2.PTDRID=a.PTDRID and f2.AltDataItem=6 and f2.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d7 ON d7.PTDID=g.PTDID AND d7.PTDRID=a.PTDRID AND d7.AltDataItem=7 AND d7.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d8 ON d8.PTDID=g.PTDID AND d8.PTDRID=a.PTDRID AND d8.AltDataItem=8 AND d8.AltDataID=@AltDataID" & vbCrLf
        sql &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d9 ON d9.PTDID=g.PTDID AND d9.PTDRID=a.PTDRID AND d9.AltDataItem=9 AND d9.AltDataID=@AltDataID" & vbCrLf
        sql &= " ORDER BY g.STrainDate,g.PName" & vbCrLf

        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRid", SqlDbType.Int).Value = tPTDRid
            .Parameters.Add("AltDataID", SqlDbType.VarChar).Value = sAltDataID
            rst.Load(.ExecuteReader())
        End With
        Return rst

    End Function

    '若無資料回傳nothing'產投用選項
    Private Function Get_PlanTrainDescNewRevise(ByVal iPTDRID As Int32) As DataTable
        Dim Rst As New DataTable
        Dim sqlStr As String = ""
        sqlStr &= " SELECT a.PTDRID,g.PTDID" & vbCrLf
        sqlStr &= " ,b.PTDRIID AS ID1,convert(varchar, ISNULL(convert(datetime, b.NewData, 111),g.STrainDate), 111) STrainDate" & vbCrLf
        sqlStr &= " ,c.PTDRIID AS ID2,ISNULL(c.NewData,g.PName) PName" & vbCrLf
        'sqlStr &=" ,d.PTDRIID AS ID3,ISNULL(CONVERT(NUMERIC, LTRIM(RTRIM(d.NewData))),g.PHour) as PHour" & vbCrLf 'ora-01722 add trim
        sqlStr &= " ,d.PTDRIID AS ID3 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d.NewData),g.PHour)) AS PHour" & vbCrLf '時數
        sqlStr &= " ,d9.PTDRIID AS ID9 ,CONVERT(NUMERIC(3,1),ISNULL(dbo.FN_VALUE1(d9.NewData),g.EHour)) AS EHour" & vbCrLf '技檢訓練時數
        sqlStr &= " ,d7.PTDRIID AS ID7,ISNULL(d7.NewData,g.TPERIOD28) TPERIOD28" & vbCrLf
        sqlStr &= " ,d8.PTDRIID AS ID8,ISNULL(d8.NewData,g.FARLEARN) FARLEARN" & vbCrLf
        sqlStr &= " ,g.PCont" & vbCrLf
        sqlStr &= " ,g.Classification1" & vbCrLf
        sqlStr &= " ,e.PTDRIID AS ID4,ISNULL(e.NewData,g.PTID) AS PTID" & vbCrLf
        sqlStr &= " ,f.PTDRIID AS ID5,ISNULL(LTRIM(RTRIM(f.NewData)),g.TechID) AS TechID" & vbCrLf 'ora-01722 add trim
        sqlStr &= " ,f2.PTDRIID AS ID6,ISNULL(LTRIM(RTRIM(f2.NewData)),g.TechID2) AS TechID2" & vbCrLf 'ora-01722 add trim

        sqlStr &= " FROM PLAN_TRAINDESC_REVISE a" & vbCrLf 'ora-01722 add trim
        sqlStr &= " JOIN PLAN_TRAINDESC g on g.PlanID=a.PlanID and g.ComIDNO=a.ComIDNO and g.SeqNO=a.SeqNO and a.PTDRID=@PTDRID" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM b on b.PTDID=g.PTDID and b.PTDRID=a.PTDRID and b.AltDataItem=1" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM c on c.PTDID=g.PTDID and c.PTDRID=a.PTDRID and c.AltDataItem=2" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d on d.PTDID=g.PTDID and d.PTDRID=a.PTDRID and d.AltDataItem=3" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM e on e.PTDID=g.PTDID and e.PTDRID=a.PTDRID and e.AltDataItem=4" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f on f.PTDID=g.PTDID and f.PTDRID=a.PTDRID and f.AltDataItem=5" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM f2 on f2.PTDID=g.PTDID and f2.PTDRID=a.PTDRID and f2.AltDataItem=6" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d7 ON d7.PTDID=g.PTDID AND d7.PTDRID=a.PTDRID AND d7.AltDataItem=7" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d8 ON d8.PTDID=g.PTDID AND d8.PTDRID=a.PTDRID AND d8.AltDataItem=8" & vbCrLf
        sqlStr &= " LEFT JOIN PLAN_TRAINDESC_REVISEITEM d9 ON d9.PTDID=g.PTDID AND d9.PTDRID=a.PTDRID AND d9.AltDataItem=9" & vbCrLf
        'sqlStr +=" ORDER BY g.STrainDate,g.PName "
        sqlStr &= " ORDER BY STrainDate,PName "
        Dim sCmd As New SqlCommand(sqlStr, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRID", SqlDbType.Int).Value = iPTDRID
            Rst.Load(.ExecuteReader())
        End With
        If Rst.Rows.Count = 0 Then Rst = Nothing
        Return Rst
    End Function

    ''' <summary>取得最可能性正常資料。</summary>
    ''' <param name="planid"></param>
    ''' <param name="comidno"></param>
    ''' <param name="seqno"></param>
    ''' <param name="adate"></param>
    ''' <param name="subseqno"></param>
    ''' <param name="sAltDataID"></param>
    ''' <returns></returns>
    Private Function Get_PTDRID(ByVal planid As Integer, ByVal comidno As String, ByVal seqno As Integer,
                                ByVal adate As String, ByVal subseqno As Integer, ByVal sAltDataID As String) As Integer
        Dim rst As Integer = 0

        Dim sqlStr As String = ""
        sqlStr &= " SELECT PTDRID" & vbCrLf
        sqlStr &= " FROM PLAN_TRAINDESC_REVISE" & vbCrLf
        sqlStr &= " WHERE PlanID=@planid and ComIDNO=@comidno and SeqNO=@seqno" & vbCrLf
        sqlStr &= " and CDate=@cdate and SubSeqNO=@subseqno" & vbCrLf
        Dim sCmd As New SqlCommand(sqlStr, objconn)
        Dim dtX As New DataTable
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            '.SelectCommand=New SqlCommand(sqlStr, objConn)
            .Parameters.Clear()
            .Parameters.Add("planid", SqlDbType.Int).Value = planid
            .Parameters.Add("comidno", SqlDbType.VarChar).Value = comidno
            .Parameters.Add("seqno", SqlDbType.Int).Value = seqno
            .Parameters.Add("cdate", SqlDbType.DateTime).Value = CDate(adate)
            .Parameters.Add("subseqno", SqlDbType.Int).Value = subseqno
            dtX.Load(.ExecuteReader())
        End With
        If dtX.Rows.Count > 0 AndAlso dtX.Rows.Count = 1 Then
            '若只有一筆為正常
            rst = dtX.Rows(0)("PTDRID")
            Return rst '結束
        End If

        If dtX.Rows.Count > 0 AndAlso sAltDataID <> "" Then
            '超過1筆 (異常) '試著取得可能正確的 PTDRID
            sqlStr = "" & vbCrLf
            sqlStr &= " SELECT MAX(a.PTDRID) PTDRID" & vbCrLf
            sqlStr &= " FROM PLAN_TRAINDESC_REVISE a" & vbCrLf
            sqlStr &= " WHERE a.PlanID=@planid and a.ComIDNO=@comidno and a.SeqNO=@seqno" & vbCrLf
            sqlStr &= " and a.CDate=@cdate and a.SubSeqNO=@subseqno" & vbCrLf
            sqlStr &= " AND EXISTS (" & vbCrLf
            sqlStr &= "   SELECT 'X' FROM PLAN_TRAINDESC_REVISEITEM x where x.PTDRID=a.PTDRID AND x.ALTDATAID=@ALTDATAID" & vbCrLf
            sqlStr &= " )" & vbCrLf
            Dim sCmd2 As New SqlCommand(sqlStr, objconn)
            Dim dtX2 As New DataTable
            With sCmd2
                '.SelectCommand=New SqlCommand(sqlStr, objConn)
                .Parameters.Clear()
                .Parameters.Add("planid", SqlDbType.Int).Value = planid
                .Parameters.Add("comidno", SqlDbType.VarChar).Value = comidno
                .Parameters.Add("seqno", SqlDbType.Int).Value = seqno
                .Parameters.Add("cdate", SqlDbType.DateTime).Value = CDate(adate)
                .Parameters.Add("subseqno", SqlDbType.Int).Value = subseqno
                .Parameters.Add("ALTDATAID", SqlDbType.VarChar).Value = sAltDataID
                dtX2.Load(.ExecuteReader())
            End With
            If dtX2.Rows.Count > 0 AndAlso Convert.ToString(dtX2.Rows(0)("PTDRID")) <> "" Then
                rst = dtX2.Rows(0)("PTDRID")
                Return rst '結束
            End If
        End If
        '無上述狀況，無法得知正確的 PTDRID 回傳為0(查無資料) 
        Return rst
    End Function

    Function Get_TrainPlaceName(ByVal i_PTID As Integer, ByRef oConn As SqlConnection) As String
        Dim rst As String = String.Empty
        'TIMS.OpenDbConn(objconn) 'TRAINPLACE
        Dim sqlStr As String = "SELECT PLACENAME FROM PLAN_TRAINPLACE WHERE PTID=@PTID"
        Dim sCmd As New SqlCommand(sqlStr, oConn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTID", SqlDbType.Int).Value = i_PTID
            rst = Convert.ToString(.ExecuteScalar())
        End With
        Return rst
    End Function

    ''' <summary>UPDATE PLAN_TRAINDESC</summary>
    ''' <param name="iPTDRID"></param>
    ''' <param name="tmpConn"></param>
    ''' <param name="tmpTrans"></param>
    Sub UPDATE_PLANTRAINDESC(ByVal iPTDRID As Integer, ByVal tmpConn As SqlConnection, ByVal tmpTrans As SqlTransaction)
        'Dim sqlAdp AS New SqlDataAdapter
        Dim sqlStr As String = "" 'String.Empty
        sqlStr &= " SELECT PTDID,AltDataItem,NewData"
        sqlStr &= " FROM PLAN_TRAINDESC_REVISEITEM "
        sqlStr &= " WHERE PTDRID=@PTDRID "
        sqlStr &= " ORDER BY PTDID,AltDataItem"
        Dim sCmd As New SqlCommand(sqlStr, tmpConn, tmpTrans)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PTDRID", SqlDbType.Int).Value = iPTDRID
            dt.Load(.ExecuteReader())
        End With

        If dt Is Nothing Then Return
        If dt.Rows.Count = 0 Then Return

        'SELECT ALTDATAITEM,COUNT(1) CC1 FROM PLAN_TRAINDESC_REVISEITEM GROUP BY ALTDATAITEM
        For Each dr As DataRow In dt.Rows
            Dim dbType As SqlDbType = SqlDbType.VarChar
            Select Case Convert.ToString(dr("AltDataItem"))
                Case "1"    '日期，對應Plan_TrainDesc.STrainDate
                    sqlStr = "UPDATE PLAN_TRAINDESC set STrainDate=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.DateTime
                Case "2"    '授課時間，對應Plan_TrainDesc.PName
                    sqlStr = "UPDATE PLAN_TRAINDESC set PName=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.NVarChar
                Case "3"    '時數，對應Plan_TrainDesc.PHour
                    sqlStr = "UPDATE PLAN_TRAINDESC set PHour=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = If(TIMS.IsNumeric2(dr("NewData")), SqlDbType.Int, SqlDbType.Float)
                Case "4"    '上課地點，對應Plan_TrainDesc.PTID
                    sqlStr = "UPDATE PLAN_TRAINDESC set PTID=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.Int
                Case "5"    '任課教師，對應Plan_TrainDesc.TechID
                    sqlStr = "UPDATE PLAN_TRAINDESC set TechID=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.Int
                Case "6"    '助教，對應Plan_TrainDesc.TechID2
                    sqlStr = "UPDATE PLAN_TRAINDESC set TechID2=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.Int
                Case "7"    '授課時間-授課時段-123，對應 Plan_TrainDesc.TPERIOD28
                    sqlStr = "UPDATE PLAN_TRAINDESC set TPERIOD28=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.VarChar
                Case "8"    '遠距教學
                    sqlStr = "UPDATE PLAN_TRAINDESC set FARLEARN=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = SqlDbType.VarChar
                Case "9"    '技檢訓練時數，對應Plan_TrainDesc.EHour
                    sqlStr = "UPDATE PLAN_TRAINDESC set EHour=@data1, modifydate=getdate() where PTDID=@id"
                    dbType = If(TIMS.IsNumeric2(dr("NewData")), SqlDbType.Int, SqlDbType.Float)
            End Select

            Dim o_data1Value As Object = Convert.DBNull
            Dim uCmd As New SqlCommand(sqlStr, tmpConn, tmpTrans)
            With uCmd
                .Parameters.Clear()
                If TIMS.ConvertStr(dr("NewData")) <> "" Then
                    Select Case dbType
                        Case SqlDbType.DateTime
                            o_data1Value = CDate(dr("NewData")) 'Value
                        Case SqlDbType.Int, SqlDbType.Float
                            o_data1Value = Val(dr("NewData")) 'Value
                        Case SqlDbType.NVarChar
                            o_data1Value = dr("NewData") 'Value
                        Case Else 'VarChar
                            o_data1Value = dr("NewData") 'Value
                    End Select
                End If
                .Parameters.Add("data1", dbType).Value = o_data1Value
                .Parameters.Add("id", SqlDbType.Int).Value = dr("PTDID")
                .ExecuteNonQuery()
            End With
        Next
        'sqlAdp.Dispose()
    End Sub

    '修正- 'DISTANCE/FARLEARN 遠距教學 1.申請整班為遠距教學", 2."申請部分課程為遠距/實體教學", 3."申請整班為實體教學
    Public Shared Sub UPDATE_PLANTRAINDESC_FARLEARN(ByRef ssHash As Hashtable, ByVal tmpConn As SqlConnection)
        Dim v_PlanID As String = TIMS.GetMyValue2(ssHash, "PlanID")
        Dim v_ComIDNO As String = TIMS.GetMyValue2(ssHash, "ComIDNO")
        Dim v_SeqNO As String = TIMS.GetMyValue2(ssHash, "SeqNO")
        'Dim v_FARLEARN As String=TIMS.GetMyValue2(ssHash, "FARLEARN")

        Dim drPP As DataRow = TIMS.GetPCSDate(v_PlanID, v_ComIDNO, v_SeqNO, tmpConn)
        If drPP Is Nothing Then Return '(無資料離開)
        Dim pDISTANCE As String = Convert.ToString(drPP("DISTANCE"))
        If pDISTANCE = "" Then Return '(無資料離開)

        Dim v_FARLEARN As String = If(pDISTANCE = "1", "Y", "")
        'Dim ssHash As New Hashtable
        'DISTANCE/FARLEARN 遠距教學 1.申請整班為遠距教學", 2."申請部分課程為遠距/實體教學", 3."申請整班為實體教學
        Dim flag_can_update_farlearn As Boolean = If(pDISTANCE = "1", True, If(pDISTANCE = "3", True, False))
        '不可修改 遠距教學 Return
        If Not flag_can_update_farlearn Then Return

        Dim sqlStr As String = "" 'String.Empty
        sqlStr &= " SELECT FARLEARN"
        sqlStr &= " FROM PLAN_TRAINDESC"
        sqlStr &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
        Dim sCmd As New SqlCommand(sqlStr, tmpConn)
        Dim dtDS As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PLANID", SqlDbType.Int).Value = v_PlanID ' dr("PLANID")
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = v_ComIDNO 'dr("COMIDNO")
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = v_SeqNO 'dr("SEQNO")
            dtDS.Load(.ExecuteReader())
        End With
        If dtDS Is Nothing Then Return '(無資料離開)
        If dtDS.Rows.Count = 0 Then Return '(無資料離開)

        Dim u_sqlStr As String = ""
        u_sqlStr &= " UPDATE PLAN_TRAINDESC"
        u_sqlStr &= " SET FARLEARN=@FARLEARN"
        u_sqlStr &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO "
        Dim u_Cmd As New SqlCommand(u_sqlStr, tmpConn)
        With u_Cmd
            .Parameters.Clear()
            .Parameters.Add("FARLEARN", SqlDbType.VarChar).Value = If(v_FARLEARN <> "", v_FARLEARN, Convert.DBNull)
            .Parameters.Add("PLANID", SqlDbType.Int).Value = v_PlanID ' dr("PLANID")
            .Parameters.Add("COMIDNO", SqlDbType.VarChar).Value = v_ComIDNO 'dr("COMIDNO")
            .Parameters.Add("SEQNO", SqlDbType.Int).Value = v_SeqNO 'dr("SEQNO")
            .ExecuteNonQuery()
        End With
    End Sub

#End Region


    ''' <summary>
    ''' 送出鍵-不顯示-(儲存)-SaveData1-But_Sub
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub But_Sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles But_Sub.Click
        Dim v_ChkMode As String = TIMS.GetListValue(ChkMode)
        If v_ChkMode = "" Then Exit Sub
        Try
            Call SaveData1()
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= "TC_06_001_chk.But_Sub_Click.SaveData1" & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
        End Try
    End Sub

    '回上一頁
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Session("_search")=ViewState("_search")
        'Response.Redirect("TC_06_001.aspx?ID=" & Request("ID"))
        Dim url1 As String = "TC_06_001.aspx?ID=" & TIMS.Get_MRqID(Me) 'Request("ID")
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    ''' <summary> '刪除鈕 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub BtnDelete_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnDelete.Click
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim uUrl1 As String = "TC/06/TC_06_001.aspx?ID=" & RqID

        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(Convert.ToString(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Convert.ToString(Request("SubNo")))), 0)

        Dim drCls As DataRow = TIMS.Get_ClassDaRow(rPlanID, rComIDNO, rSeqNO, objconn)
        ViewState(vs_OCID) = If(drCls IsNot Nothing, Convert.ToString(drCls("OCID")), "")

        If rPlanID = "" OrElse rComIDNO = "" OrElse rSeqNO = "" OrElse rSCDate = "" OrElse iSubSeqNO = 0 Then
            'Dim sMsg AS String="查無該計畫變更資料，請重新查詢!!"
            'Common.RespWrite(Me, "<script language=javascript>window.alert('" & sMsg & "');")
            'Common.RespWrite(Me, "window.location.href='TC_06_001.aspx?ID=" & Request("ID") & "';</script>")
            'Common.MessageBox(Page, sMsg)
            'Exit Sub
            Dim sMsg As String = "查無該計畫變更資料，請重新查詢!!"
            Call TIMS.BlockAlert(Me, sMsg, uUrl1)
            Exit Sub
        End If
        Dim drPCS As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)
        If drPCS Is Nothing Then
            Dim sMsg As String = "查無該計畫資料，請重新查詢!!"
            Common.MessageBox(Page, sMsg)
            Exit Sub
        End If

        Dim vsAltDataID As String = hid_AltDataID.Value '=TIMS.GetMyValue(e.CommandArgument, "AltDataID")
        Dim PTDRID As Integer = 0
        PTDRID = Get_PTDRID(Val(rPlanID), rComIDNO, Val(rSeqNO), rSCDate, iSubSeqNO, vsAltDataID)
        '刪除 DELETE PLAN_TRAINDESC_REVISEITEM. DELETE PLAN_TRAINDESC_REVISE.
        If PTDRID <> 0 Then Call TIMS.DEL_PLAN_TRAINDESC_REVISEITEM(sm, PTDRID, objconn)
        '刪除 DELETE PLAN_REVISE
        Call TIMS.DELETE_PLANREVISE(rPlanID, rComIDNO, rSeqNO, rSCDate, iSubSeqNO, vsAltDataID, objconn)

        'Common.RespWrite(Me, "<script language=javascript>window.alert('" & sMsg1 & "');")
        'Common.RespWrite(Me, "window.location.href='TC_06_001.aspx?ID=" & Request("ID") & "';</script>")
        'Common.MessageBox(Page, sMsg1)
        'Dim sMsg AS String="查無該計畫變更資料，請重新查詢!!"
        Dim sMsg1 As String = "刪除成功!"
        Call TIMS.BlockAlert(Me, sMsg1, uUrl1)
        'Exit Sub
    End Sub

    ''' <summary> 審核結果-儲存-SaveData2-Btn_SAVE2 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Btn_SAVE2_Click(sender As Object, e As EventArgs) Handles Btn_SAVE2.Click
        Dim v_ChkMode As String = TIMS.GetListValue(ChkMode)
        If v_ChkMode = "" Then Exit Sub
        Call SaveData2()

        'Try
        '    Call SaveData2()
        'Catch ex As Exception
        '    Dim strErrmsg As String=""
        '    strErrmsg &="TC_06_001_chk.Btn_SAVE2_Click.SaveData2" & vbCrLf
        '    strErrmsg &=TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg &="ex.ToString : " & ex.ToString & vbCrLf
        '    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    Call TIMS.WriteTraceLog(strErrmsg)
        'End Try
    End Sub

#Region "產投-審查計分表"

    ''' <summary> 儲存-產投-審查計分表 </summary>
    Sub SAVE_STATUS4()
        '(非產投)離開
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then Return

        Dim U_sql As String = ""
        U_sql &= " UPDATE PLAN_REVISE" & vbCrLf
        U_sql &= " SET SENDACCT4=@SENDACCT4 ,SENDDATE4=CONVERT(date,@SENDDATE4)" & vbCrLf
        U_sql &= " ,STATUS4=@STATUS4 ,ISPASS4=@ISPASS4 ,OVERWEEK4=@OVERWEEK4 ,NOINC4=@NOINC4" & vbCrLf
        'U_sql &=" ,NODEDUC4=@NODEDUC4" & vbCrLf
        U_sql &= " ,MODIFYACCT4=@MODIFYACCT4 ,MODIFYDATE4=GETDATE()" & vbCrLf
        U_sql &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo" & vbCrLf
        U_sql &= " AND CDate=CONVERT(date,@CDate) AND SubSeqNO=@SubSeqNO" & vbCrLf

        Dim vNOINC4 As String = If(chkbox_NOINC4.Checked, TIMS.cst_YES, TIMS.cst_NO) '不納入審查計分變更次數
        'Dim vNODEDUC4 As String=If(chkbox_NODEDUC4.Checked, TIMS.cst_YES, TIMS.cst_NO) '政策性課程不扣分
        Dim vSTATUS4 As String = TIMS.GetListValue(STATUS4)
        Dim vddlISPASS4 As String = TIMS.GetListValue(ddlISPASS4)
        Dim vOVERWEEK4 As String = TIMS.GetListValue(OVERWEEK4)

        'hPARMS.Clear()
        'hPARMS.Add("NODEDUC4", vNODEDUC4) '政策性課程不扣分
        Dim hPARMS As New Hashtable From {
            {"SENDACCT4", sm.UserInfo.UserID},
            {"SENDDATE4", TIMS.Cdate3(SENDDATE4.Text)},
            {"STATUS4", If(vSTATUS4 <> "", vSTATUS4, Convert.DBNull)},
            {"ISPASS4", If(vddlISPASS4 <> "", vddlISPASS4, Convert.DBNull)},
            {"OVERWEEK4", If(vOVERWEEK4 <> "", Val(vOVERWEEK4), Convert.DBNull)},
            {"MODIFYACCT4", sm.UserInfo.UserID},
            {"NOINC4", vNOINC4},
            {"PlanID", rPlanID},
            {"ComIDNO", rComIDNO},
            {"SeqNo", rSeqNO},
            {"CDate", TIMS.Cdate3(rSCDate)},
            {"SubSeqNO", iSubSeqNO}
        }
        DbAccess.ExecuteNonQuery(U_sql, objconn, hPARMS)
    End Sub

    ''' <summary>資料-產投-審查計分表</summary>
    Sub save_divCo128()
        '通過/不通過
        Dim vNOINC4 As String = If(chkbox_NOINC4.Checked, TIMS.cst_YES, TIMS.cst_NO) '不納入審查計分變更次數
        'Dim vNODEDUC4 As String=If(chkbox_NODEDUC4.Checked, TIMS.cst_YES, TIMS.cst_NO) '政策性課程不扣分
        Dim vSTATUS4 As String = TIMS.GetListValue(STATUS4)
        Dim vddlISPASS4 As String = TIMS.GetListValue(ddlISPASS4)
        Dim vOVERWEEK4 As String = TIMS.GetListValue(OVERWEEK4)

        Dim pms_u3 As New Hashtable From {{"PlanID", rPlanID}, {"ComIDNO", rComIDNO}, {"SeqNo", rSeqNO}, {"CDate3", TIMS.Cdate3(rSCDate)}, {"SubSeqNO", iSubSeqNO}}
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            pms_u3.Add("SENDACCT4", sm.UserInfo.UserID)
            pms_u3.Add("SENDDATE4", TIMS.Cdate2(SENDDATE4.Text))
            pms_u3.Add("STATUS4", If(vSTATUS4 <> "", vSTATUS4, Convert.DBNull))
            pms_u3.Add("ISPASS4", If(vddlISPASS4 <> "", vddlISPASS4, Convert.DBNull))
            pms_u3.Add("OVERWEEK4", If(vOVERWEEK4 <> "", Val(vOVERWEEK4), Convert.DBNull))
            pms_u3.Add("MODIFYACCT4", sm.UserInfo.UserID)
            pms_u3.Add("MODIFYDATE4", Now())
            pms_u3.Add("NOINC4", vNOINC4)
            'pms_u3.Add("NODEDUC4", vNODEDUC4)
        Else
            pms_u3.Add("SENDACCT4", Convert.DBNull)
            pms_u3.Add("SENDDATE4", Convert.DBNull)
            pms_u3.Add("STATUS4", Convert.DBNull)
            pms_u3.Add("ISPASS4", Convert.DBNull)
            pms_u3.Add("OVERWEEK4", Convert.DBNull)
            pms_u3.Add("MODIFYACCT4", Convert.DBNull)
            pms_u3.Add("MODIFYDATE4", Convert.DBNull)
            pms_u3.Add("NOINC4", Convert.DBNull)
            'pms_u3.Add("NODEDUC4", Convert.DBNull)
        End If

        Dim sql_u3 As String = ""
        sql_u3 &= " UPDATE PLAN_REVISE" & vbCrLf
        sql_u3 &= " SET SENDACCT4=@SENDACCT4 ,SENDDATE4=@SENDDATE4 ,STATUS4=@STATUS4 ,ISPASS4=@ISPASS4 ,OVERWEEK4=@OVERWEEK4" & vbCrLf
        sql_u3 &= " ,MODIFYACCT4=@MODIFYACCT4 ,MODIFYDATE4=@MODIFYDATE4 ,NOINC4=@NOINC4" & vbCrLf
        'sql_u3 &=" ,NODEDUC4=@NODEDUC4" & vbCrLf
        sql_u3 &= " WHERE PlanID=@PlanID AND ComIDNO=@ComIDNO AND SeqNo=@SeqNo AND CDate=@CDate3 AND SubSeqNO=@SubSeqNO" & vbCrLf
        DbAccess.ExecuteNonQuery(sql_u3, objconn, pms_u3)
    End Sub

    ''' <summary> 顯示-產投-審查計分表 </summary>
    Sub SHOW_STATUS4(ByRef objrow As DataRow)
        divCo128.Visible = False '(隱藏)
        chkbox_NOINC4.Checked = False '不納入審查計分變更次數 '(清空／不勾選)
        TIMS.Tooltip(chkbox_NOINC4, "")
        'chkbox_NODEDUC4.Checked=False '政策性課程不扣分 '(清空／不勾選)
        'TIMS.Tooltip(chkbox_NODEDUC4, "")
        '(查無資料)離開 ／ '(非產投)離開 
        If objrow Is Nothing OrElse TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) = -1 Then Return

        divCo128.Visible = True '(顯示)
        If Convert.ToString(objrow("SENDACCT4")) <> "" Then
            '有輸入值 顯示後離開
            SENDDATE4.Text = TIMS.Cdate3(objrow("SENDDATE4"))
            'STATUS4 -申請變更函送狀態 - 1.依規定辦理 2.逾期(扣分) 3.逾期(不扣分)
            Common.SetListItem(STATUS4, Convert.ToString(objrow("STATUS4")))
            Common.SetListItem(ddlISPASS4, Convert.ToString(objrow("ISPASS4")))
            '不納入審查計分變更次數
            chkbox_NOINC4.Checked = If(Convert.ToString(objrow("NOINC4")) = "Y", True, False) ' False
            'If Convert.ToString(objrow("NOINC4"))="Y" Then chkbox_NOINC4.Checked=True
            '政策性課程不扣分
            'chkbox_NODEDUC4.Checked=If(Convert.ToString(objrow("NODEDUC4"))="Y", True, False) ' False
            'If Convert.ToString(objrow("NODEDUC4"))="Y" Then chkbox_NODEDUC4.Checked=True
            'OVERWEEK4 -逾期週數 - 1.1週以內 2.1週以上 3.停辦 9.無逾期
            Common.SetListItem(OVERWEEK4, Convert.ToString(objrow("OVERWEEK4")))
            Return '有值後離開
        End If

        Const s_DEF_VAL1 As String = "政府政策性產業 預設勾選" '"(預設值)"
        Const s_DEF_VAL2 As String = "遠距教學 預設勾選" '"(預設值)"
        '(如果沒有值的執行位置) 預設動作如下：'預設'政策性課程不扣分
        Dim dr2 As DataRow = TIMS.GET_PLANDEPOT(rPlanID, rComIDNO, rSeqNO, objconn)
        If dr2 IsNot Nothing Then
            Dim KID19 As String = Convert.ToString(dr2("KID19"))
            Dim KID20 As String = Convert.ToString(dr2("KID20"))
            Dim KID25 As String = Convert.ToString(dr2("KID25"))
            '當班級為政府政策性產業之課程者，如下圖： (依原本判斷邏輯不變：有勾選【重點產業審核】者，非依【申請階段】)，
            '【不納入審查計分變更次數】調整為：預設 勾選「不納入」 '【政策性課程不扣分】調整為：預設不勾選 government policy industries
            Dim fg_is_gover As Boolean = (KID19 <> "" OrElse KID20 <> "" OrElse KID25 <> "")
            If fg_is_gover Then
                chkbox_NOINC4.Checked = True '不納入審查計分變更次數
                TIMS.Tooltip(chkbox_NOINC4, s_DEF_VAL1, True)
            End If
        End If
        'Cst_i遠距教學
        If Convert.ToString(objrow("AltDataID")) = CStr(Cst_i遠距教學) Then
            chkbox_NOINC4.Checked = True '不納入審查計分變更次數
            TIMS.Tooltip(chkbox_NOINC4, s_DEF_VAL2, True)
        End If
        'Return
    End Sub
#End Region

    '建立可選教師列表-遴選辦法說明
    Sub SHOW_REVISE_TEACHER3(ByRef htSS As Hashtable, ByRef oConn As SqlConnection)
        'Dim rPlanID As String=TIMS.GetMyValue2(htSS, "rPlanID") '計畫PK
        'Dim rComIDNO As String=TIMS.GetMyValue2(htSS, "rComIDNO") '計畫PK
        'Dim rSeqNo As String=TIMS.GetMyValue2(htSS, "rSeqNo") '計畫PK
        'Dim SCDate As String=TIMS.GetMyValue2(htSS, "SCDate") 'ApplyDate.Text
        'Dim SubSeqNo As String=TIMS.GetMyValue2(htSS, "SubSeqNo") 'iSubSeqNO
        'Dim vActCheck As String=TIMS.GetMyValue2(htSS, "ActCheck") 'RActCheck / Cst_cPlan '申請 /Cst_cRevise '審核查詢
        'Dim rSCDate As String=""
        'Dim iSubSeqNO As Integer=0

        Dim rqRID As String = TIMS.GetMyValue2(htSS, "RID")
        Dim rqTECHIDs As String = TIMS.GetMyValue2(htSS, "TECHIDs")
        Dim TechTYPE As String = TIMS.GetMyValue2(htSS, "TechTYPE") 'A/B
        If rqTECHIDs = "" Then Exit Sub
        Dim inTECHIDs As String = TIMS.CombiSQM2IN(rqTECHIDs)
        If inTECHIDs = "" Then Exit Sub

        'parms.Clear()
        Dim parms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}, {"CDATE", rSCDate}, {"SUBSEQNO", iSubSeqNO}}
        Dim sql As String = ""
        'AND PLANID='4818',AND COMIDNO='80592907',AND SEQNO='1',AND CDATE='2020-04-07',AND SUBSEQNO=1,
        sql &= " WITH WP1 AS ( SELECT PLANID,COMIDNO,SEQNO,CDATE,SUBSEQNO" & vbCrLf
        sql &= " 	FROM PLAN_REVISE" & vbCrLf
        sql &= " 	WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO" & vbCrLf
        sql &= " 	AND CDATE=@CDATE AND SUBSEQNO=@SUBSEQNO )" & vbCrLf
        sql &= " SELECT a.TechID" & vbCrLf '教師ID
        sql &= " ,a.TeachCName" & vbCrLf '教師姓名 
        sql &= " ,a.DegreeID" & vbCrLf '學歷
        sql &= " ,c.Name DegreeName" & vbCrLf '學歷
        '專業領域 Specialty1
        sql &= " ,ISNULL(a.Specialty1, '') Specialty1" & vbCrLf
        '專業證照-相關證照
        sql &= " ,CASE WHEN a.ProLicense1 IS NOT NULL AND a.ProLicense2 IS NOT NULL THEN a.ProLicense1 + '、' + a.ProLicense2" & vbCrLf
        sql &= " ELSE a.ProLicense END ProLicense" & vbCrLf
        '遴選辦法說明
        sql &= " ,dbo.FN_GET_REVISE_TEACHER3(b.PLANID, b.COMIDNO, b.SEQNO,b.CDATE, b.SUBSEQNO, '" & TechTYPE & "', a.TechID) TeacherDesc" & vbCrLf 'TechTYPE: A:師資/B:助教
        sql &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sql &= " LEFT JOIN KEY_DEGREE c ON a.DegreeID=c.DegreeID" & vbCrLf
        'CROSS JOIN
        sql &= " CROSS JOIN WP1 b" & vbCrLf
        sql &= " WHERE a.WorkStatus='1'" & vbCrLf
        sql &= " AND a.RID='" & rqRID & "'" & vbCrLf
        sql &= " AND a.TechID IN (" & inTECHIDs & ")" & vbCrLf
        'sql &=" AND a.RID='F3962'" & vbCrLf
        'sql &=" AND a.TechID IN (432456,441777,441706,441771,435528,441840,441696)" & vbCrLf
        sql &= " ORDER BY a.TechID" & vbCrLf

        Select Case TechTYPE
            Case "A" '師資
                Dim dtT As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
                i_gSeqno = 0
                tbDataGrid21.Visible = False
                If dtT.Rows.Count > 0 Then
                    tbDataGrid21.Visible = True
                    DataGrid21.DataSource = dtT
                    DataGrid21.DataBind()
                End If

            Case "B" '助教
                Dim dtT2 As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
                i_gSeqno = 0
                tbDataGrid22.Visible = False
                If dtT2.Rows.Count > 0 Then
                    tbDataGrid22.Visible = True
                    DataGrid22.DataSource = dtT2
                    DataGrid22.DataBind()
                End If
        End Select

    End Sub

    Private Sub DataGrid21_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid21.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim HidTechID As HtmlInputHidden = e.Item.FindControl("HidTechID")
                Dim seqno As Label = e.Item.FindControl("seqno")
                Dim TeachCName As Label = e.Item.FindControl("TeachCName")
                Dim DegreeName As Label = e.Item.FindControl("DegreeName")
                Dim Specialty1 As Label = e.Item.FindControl("Specialty1")
                'Dim ProLicense As Label=e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEA As HtmlInputButton = e.Item.FindControl("btn_TCTYPEA")
                Dim rqRID As String = sm.UserInfo.RID
                If RIDValue.Value.Length > 1 Then rqRID = RIDValue.Value
                'sWOScript1="wopen('../../Common/TeachDesc1.aspx?TCTYPE=A&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                'btn_TCTYPEA.Attributes("onclick")=sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                Specialty1.Text = Convert.ToString(drv("Specialty1"))
                'ProLicense.Text=Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'If Hid_NewData11_3.Value <> "" Then TeacherDesc.Text=Hid_NewData11_3.Value 'Convert.ToString(drv("TeacherDesc"))

                TeacherDesc.ReadOnly = False
                btn_TCTYPEA.Visible = True

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
                TeacherDesc.ReadOnly = True
                btn_TCTYPEA.Visible = False
                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly=True
                '        btn_TCTYPEA.Visible=False
                'End Select
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
                'Dim ProLicense As Label=e.Item.FindControl("ProLicense")
                Dim TeacherDesc As TextBox = e.Item.FindControl("TeacherDesc")
                Dim btn_TCTYPEB As HtmlInputButton = e.Item.FindControl("btn_TCTYPEB")
                Dim rqRID As String = sm.UserInfo.RID '使用者單位
                If RIDValue.Value.Length > 1 Then rqRID = RIDValue.Value
                'sWOScript1="wopen('../../Common/TeachDesc1.aspx?TCTYPE=B&RID=" & rqRID & "&TB1=" & TeacherDesc.ClientID & "','" & TIMS.xBlockName() & "',650,350,1);"
                'btn_TCTYPEB.Attributes("onclick")=sWOScript1

                HidTechID.Value = Convert.ToString(drv("TechID"))
                i_gSeqno += 1
                seqno.Text = i_gSeqno
                TeachCName.Text = Convert.ToString(drv("TeachCName"))
                DegreeName.Text = Convert.ToString(drv("DegreeName"))
                Specialty1.Text = Convert.ToString(drv("Specialty1"))
                'ProLicense.Text=Convert.ToString(drv("ProLicense"))
                TeacherDesc.Text = Convert.ToString(drv("TeacherDesc"))
                'If Hid_NewData20_3.Value <> "" Then TeacherDesc.Text=Hid_NewData20_3.Value 'Convert.ToString(drv("TeacherDesc"))

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
                TeacherDesc.ReadOnly = True
                btn_TCTYPEB.Visible = False
                'Select Case rqProcessType 'ProcessType @Insert/Update/View
                '    Case cst_ptView '查詢功能不提供儲存
                '        TeacherDesc.ReadOnly=True
                '        btn_TCTYPEB.Visible=False
                'End Select
        End Select
    End Sub

    ''' <summary>儲存- A:師資/B:助教</summary>
    ''' <param name="TechTYPE">TechTYPE: A:師資/B:助教</param>
    ''' <param name="oTrans"></param>
    Sub SAVE_PLAN_TEACHER3(ByVal TechTYPE As String, ByRef oTrans As SqlTransaction)
        Select Case TechTYPE 'TechTYPE: A:師資/B:助教
            Case "A", "B"
            Case Else
                Exit Sub
        End Select

        Dim v_OCID As String = Convert.ToString(ViewState(vs_OCID))
        If v_OCID <> "" Then SAVE_CLASS_TEACHER3(sm, DataGrid21, DataGrid22, v_OCID, TechTYPE, oTrans)

        Const cst_iMaxLen_TeacherDesc As Integer = 500
        '更新師資表 'TechTYPE: A:師資/B:助教
        Const cst_tTECHTYPE_A As String = "A"
        Const cst_tTECHTYPE_B As String = "B"

        Dim hdPMS As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}, {"TECHTYPE", TechTYPE}}
        Dim dSql2 As String = " DELETE PLAN_TEACHER WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND ISNULL(TECHTYPE,'A')=@TECHTYPE"
        DbAccess.ExecuteNonQuery(dSql2, oTrans, hdPMS)

        Dim iSql As String = ""
        iSql &= " INSERT INTO PLAN_TEACHER (PLANID,COMIDNO,SEQNO,TECHID,TECHTYPE,TEACHERDESC,MODIFYACCT,MODIFYDATE)" & vbCrLf
        iSql &= " VALUES (@PLANID,@COMIDNO,@SEQNO,@TECHID,@TECHTYPE,@TEACHERDESC,@MODIFYACCT,GETDATE())" & vbCrLf

        Dim sSql1 As String = "SELECT 1 FROM PLAN_TEACHER WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND TECHID=@TECHID AND ISNULL(TECHTYPE,'A')=@TECHTYPE"

        Select Case TechTYPE
            Case cst_tTECHTYPE_A '"A"
                For Each eItem As DataGridItem In DataGrid21.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    'Dim seqno As Label=eItem.FindControl("seqno")
                    'Dim TeachCName As Label=eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    'Dim btn_TCTYPEA As HtmlInputButton=eItem.FindControl("btn_TCTYPEA")
                    Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
                    If HidTechID.Value <> "" Then
                        Dim sParms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}}
                        sParms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                        sParms.Add("TECHTYPE", TechTYPE) 'TechTYPE: A:師資/B:助教
                        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, oTrans, sParms)
                        If dr1 Is Nothing Then
                            Dim parms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}}
                            parms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                            parms.Add("TECHTYPE", TechTYPE) 'TechTYPE: A:師資/B:助教
                            parms.Add("TEACHERDESC", tTEACHERDESC)
                            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                            DbAccess.ExecuteNonQuery(iSql, oTrans, parms)
                        End If
                    End If
                Next

            Case cst_tTECHTYPE_B '"B"
                For Each eItem As DataGridItem In DataGrid22.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    'Dim seqno As Label=eItem.FindControl("seqno")
                    'Dim TeachCName As Label=eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
                    'Dim btn_TCTYPEB As HtmlInputButton=eItem.FindControl("btn_TCTYPEB")
                    If HidTechID.Value <> "" Then
                        Dim sParms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}}
                        sParms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                        sParms.Add("TECHTYPE", TechTYPE) 'TechTYPE: A:師資/B:助教
                        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, oTrans, sParms)
                        If dr1 Is Nothing Then
                            Dim parms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}}
                            parms.Add("TECHID", Val(HidTechID.Value)) 'dr("TECHID"))
                            parms.Add("TECHTYPE", TechTYPE) 'TechTYPE: A:師資/B:助教
                            parms.Add("TEACHERDESC", tTEACHERDESC)
                            parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                            DbAccess.ExecuteNonQuery(iSql, oTrans, parms)
                        End If
                    End If
                Next
        End Select

        '有班級資料，且修改-學員資料維護-導師欄位
        If v_OCID <> "" AndAlso TechTYPE = cst_tTECHTYPE_A Then
            Dim s_parms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}}
            Dim s_sql As String = "SELECT dbo.FN_GET_PLAN_TEACHER(@PLANID,@COMIDNO,@SEQNO,'1') CTNAME1"
            Dim dr1 As DataRow = DbAccess.GetOneRow(s_sql, oTrans, s_parms)

            Dim s_CTNAME1 As String = ""
            If dr1 IsNot Nothing Then
                s_CTNAME1 = Convert.ToString(dr1("CTNAME1")) '查詢有資料
                If s_CTNAME1 <> "" Then s_CTNAME1 = Replace(s_CTNAME1, ";", ",") '有資料換符號
            End If
            s_CTNAME1 = TIMS.Get_CTNAME1(s_CTNAME1)

            Dim u_parms As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}}
            u_parms.Add("CTNAME1", s_CTNAME1)
            u_parms.Add("OCID", v_OCID)
            Dim uSql As String = ""
            uSql &= " UPDATE CLASS_CLASSINFO" & vbCrLf
            uSql &= " SET CTNAME=@CTNAME1" & vbCrLf 'CTNAME nvarchar 40 允許   
            uSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND OCID=@OCID" & vbCrLf
            DbAccess.ExecuteNonQuery(uSql, oTrans, u_parms)
        End If
    End Sub

    Public Shared Sub SAVE_CLASS_TEACHER3(sm As SessionModel, DataGrid21 As DataGrid, DataGrid22 As DataGrid,
                                          ByVal v_OCID As String, ByVal TechTYPE As String, ByRef oTrans As SqlTransaction)
        Dim drCC As DataRow = TIMS.GetOCIDDate(v_OCID, oTrans.Connection, oTrans)
        If drCC Is Nothing Then Return

        Const cst_iMaxLen_TeacherDesc As Integer = 500
        '更新師資表 'TechTYPE: A:師資/B:助教
        Const cst_tTECHTYPE_A As String = "A"
        Const cst_tTECHTYPE_B As String = "B"

        Dim dParms As New Hashtable From {{"OCID", v_OCID}, {"TECHTYPE", TechTYPE}}
        Dim dSql As String = " DELETE CLASS_TEACHER WHERE OCID=@OCID AND ISNULL(TECHTYPE,'A')=@TECHTYPE"
        DbAccess.ExecuteNonQuery(dSql, oTrans, dParms)

        Dim iSqlc As String = ""
        iSqlc &= " INSERT INTO CLASS_TEACHER ( CTRID ,OCID,TECHID,MODIFYACCT,MODIFYDATE,TECHTYPE,TEACHERDESC)" & vbCrLf
        iSqlc &= " VALUES ( @CTRID ,@OCID,@TECHID,@MODIFYACCT,GETDATE(),@TECHTYPE,@TEACHERDESC)" & vbCrLf

        Dim sSql1 As String = ""
        sSql1 = " SELECT 1 FROM CLASS_TEACHER WHERE OCID=@OCID AND TECHID=@TECHID AND TECHTYPE=@TECHTYPE" & vbCrLf

        Select Case TechTYPE
            Case cst_tTECHTYPE_A '"A"
                '更新師資表
                For Each eItem As DataGridItem In DataGrid21.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    'Dim seqno As Label=eItem.FindControl("seqno")
                    'Dim TeachCName As Label=eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
                    'Dim btn_TCTYPEA As HtmlInputButton=eItem.FindControl("btn_TCTYPEA") 'dr("TECHID")) 'TechTYPE: A:師資/B:助教
                    If HidTechID.Value <> "" Then
                        Dim sParms As New Hashtable From {{"OCID", v_OCID}, {"TECHID", Val(HidTechID.Value)}, {"TECHTYPE", cst_tTECHTYPE_A}}
                        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, oTrans, sParms)
                        If dr1 Is Nothing Then
                            Dim iCTRID As Integer = DbAccess.GetNewId(oTrans, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                            Dim parms As New Hashtable From {
                                {"CTRID", iCTRID},
                                {"OCID", v_OCID},
                                {"TECHID", Val(HidTechID.Value)}, 'dr("TECHID"))
                                {"MODIFYACCT", sm.UserInfo.UserID},
                                {"TECHTYPE", cst_tTECHTYPE_A},
                                {"TEACHERDESC", tTEACHERDESC}
                            }
                            DbAccess.ExecuteNonQuery(iSqlc, oTrans, parms)
                        End If
                    End If
                Next

            Case cst_tTECHTYPE_B '"B"
                '更新師資表
                For Each eItem As DataGridItem In DataGrid22.Items
                    Dim HidTechID As HtmlInputHidden = eItem.FindControl("HidTechID")
                    'Dim seqno As Label=eItem.FindControl("seqno")
                    'Dim TeachCName As Label=eItem.FindControl("TeachCName")
                    'Dim DegreeName As Label=eItem.FindControl("DegreeName")
                    'Dim Specialty1 As Label=eItem.FindControl("Specialty1")
                    Dim TeacherDesc As TextBox = eItem.FindControl("TeacherDesc")
                    Dim tTEACHERDESC As String = TIMS.Get_Substr1(TIMS.ClearSQM(TeacherDesc.Text), cst_iMaxLen_TeacherDesc)
                    'Dim btn_TCTYPEA As HtmlInputButton=eItem.FindControl("btn_TCTYPEA")'dr("TECHID"))'TechTYPE: A:師資/B:助教
                    If HidTechID.Value <> "" Then
                        Dim sParms As New Hashtable From {{"OCID", v_OCID}, {"TECHID", Val(HidTechID.Value)}, {"TECHTYPE", cst_tTECHTYPE_B}}
                        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql1, oTrans, sParms)
                        If dr1 Is Nothing Then
                            Dim iCTRID As Integer = DbAccess.GetNewId(oTrans, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                            Dim parms As New Hashtable From {
                                {"CTRID", iCTRID},
                                {"OCID", v_OCID},
                                {"TECHID", Val(HidTechID.Value)}, 'dr("TECHID"))
                                {"MODIFYACCT", sm.UserInfo.UserID},
                                {"TECHTYPE", cst_tTECHTYPE_B},
                                {"TEACHERDESC", tTEACHERDESC}
                            }
                            DbAccess.ExecuteNonQuery(iSqlc, oTrans, parms)
                        End If

                    End If
                Next

        End Select

    End Sub

    Sub SAVE_OLD_PLAN_TEACHER2(ByVal TechTYPE As String, ByRef oTrans As SqlTransaction)
        Select Case TechTYPE
            Case "A", "B"
            Case Else
                Exit Sub
        End Select
        If ViewState(vs_OCID) <> "" Then ViewState(vs_OCID) = TIMS.ClearSQM(ViewState(vs_OCID))

        Dim objstr As String = ""
        Dim objtable As DataTable = Nothing
        Dim objadapter As SqlDataAdapter = Nothing

        'Dim TechTYPE As String=TIMS.GetMyValue2(htSS, "TechTYPE") 'A/B
        Dim sq_TECHTYPE_AB As String = String.Concat(" AND ISNULL(TECHTYPE,'A')='", TechTYPE, "'")

        If ViewState(vs_OCID).ToString <> "" Then
            'TechTYPE: A:師資/B:助教
            Dim dSql As String = String.Concat(" DELETE CLASS_TEACHER WHERE OCID='", ViewState(vs_OCID), "'", sq_TECHTYPE_AB)
            DbAccess.ExecuteNonQuery(dSql, oTrans)

            Select Case TechTYPE
                Case "A"
                    objstr = " SELECT * FROM CLASS_TEACHER WHERE 1<>1" 'TechTYPE: A:師資/B:助教
                    objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                    For Each StrTechID As String In Split(NewData11_1.Value, ",")
                        StrTechID = TIMS.ClearSQM(StrTechID)
                        If StrTechID <> "" Then
                            Dim iCTRID As Integer = DbAccess.GetNewId(oTrans, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                            Dim objrow As DataRow = objtable.NewRow
                            objtable.Rows.Add(objrow)
                            objrow("CTRID") = iCTRID
                            objrow("OCID") = ViewState(vs_OCID)
                            objrow("TechID") = StrTechID 'Split(NewData20_1.Value, ",")(i)
                            objrow("TechTYPE") = TechTYPE 'TechTYPE: A:師資/B:助教
                            objrow("TEACHERDESC") = Hid_NewData20_3.Value
                            objrow("ModifyAcct") = sm.UserInfo.UserID
                            objrow("ModifyDate") = Now
                        End If
                    Next
                    DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

                Case "B"
                    objstr = " SELECT * FROM CLASS_TEACHER WHERE 1<>1" 'TechTYPE: A:師資/B:助教
                    objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                    For Each StrTechID As String In Split(NewData20_1.Value, ",")
                        StrTechID = TIMS.ClearSQM(StrTechID)
                        If StrTechID <> "" Then
                            Dim iCTRID As Integer = DbAccess.GetNewId(oTrans, "CLASS_TEACHER_CTRID_SEQ,CLASS_TEACHER,CTRID")
                            Dim objrow As DataRow = objtable.NewRow
                            objtable.Rows.Add(objrow)
                            objrow("CTRID") = iCTRID
                            objrow("OCID") = ViewState(vs_OCID)
                            objrow("TechID") = StrTechID 'Split(NewData20_1.Value, ",")(i)
                            objrow("TechTYPE") = TechTYPE 'TechTYPE: A:師資/B:助教
                            objrow("TEACHERDESC") = Hid_NewData20_3.Value
                            objrow("ModifyAcct") = sm.UserInfo.UserID
                            objrow("ModifyDate") = Now
                        End If
                    Next
                    DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

            End Select

        End If

        'AND ISNULL(TECHTYPE,'" & TechTYPE & "')='" & TechTYPE & "'" 'TechTYPE: A:師資/B:助教
        Dim dSql2 As String = String.Concat(" DELETE PLAN_TEACHER WHERE PlanID=", rPlanID, " AND ComIDNO='", rComIDNO, "' AND SeqNo='", rSeqNO, "'", sq_TECHTYPE_AB)
        DbAccess.ExecuteNonQuery(dSql2, oTrans)

        Select Case TechTYPE
            Case "A"
                objstr = " SELECT * FROM PLAN_TEACHER WHERE 1<>1" 'TechTYPE: A:師資/B:助教
                objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                For Each StrTechID As String In Split(NewData11_1.Value, ",")
                    StrTechID = TIMS.ClearSQM(StrTechID)
                    If StrTechID <> "" Then
                        Dim objrow As DataRow = Nothing
                        objrow = objtable.NewRow
                        objtable.Rows.Add(objrow)
                        objrow("PlanID") = rPlanID 'Request("PlanID")
                        objrow("ComIDNO") = rComIDNO 'Request("cid")
                        objrow("SeqNo") = rSeqNO 'Request("no")
                        objrow("TechID") = StrTechID
                        objrow("TechTYPE") = TechTYPE 'TechTYPE: A:師資/B:助教
                        objrow("TEACHERDESC") = Hid_NewData20_3.Value
                        objrow("ModifyAcct") = sm.UserInfo.UserID
                        objrow("ModifyDate") = Now
                    End If
                Next
                DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

            Case "B"
                objstr = " SELECT * FROM PLAN_TEACHER WHERE 1<>1" 'TechTYPE: A:師資/B:助教
                objtable = DbAccess.GetDataTable(objstr, objadapter, oTrans)
                For Each StrTechID As String In Split(NewData20_1.Value, ",")
                    StrTechID = TIMS.ClearSQM(StrTechID)
                    If StrTechID <> "" Then
                        Dim objrow As DataRow = Nothing
                        objrow = objtable.NewRow
                        objtable.Rows.Add(objrow)
                        objrow("PlanID") = rPlanID 'Request("PlanID")
                        objrow("ComIDNO") = rComIDNO 'Request("cid")
                        objrow("SeqNo") = rSeqNO 'Request("no")
                        objrow("TechID") = StrTechID
                        objrow("TechTYPE") = TechTYPE 'TechTYPE: A:師資/B:助教
                        objrow("TEACHERDESC") = Hid_NewData20_3.Value
                        objrow("ModifyAcct") = sm.UserInfo.UserID
                        objrow("ModifyDate") = Now
                    End If
                Next
                DbAccess.UpdateDataTable(objtable, objadapter, oTrans)

        End Select

    End Sub

    'GET_OldDataVal OldData14_1b.Value==dr("OldData14_1")/dr("NewData14_1")/NewData14_1b/SciPlaceIDb 
    Public Shared Function GET_OldDataVal(ByVal v_OldData As String, ByVal v_NewData As String, ByRef o_NewDDL1 As DropDownList, ByRef o_Label1 As Label) As String
        o_Label1.Text = If(o_NewDDL1.Items.FindByValue(v_OldData) IsNot Nothing, o_NewDDL1.Items.FindByValue(v_OldData).Text, "") '舊值顯示文字
        '整理下拉，只留選擇
        TIMS.GET_NewListItemVal(o_NewDDL1, v_NewData)
        o_NewDDL1.Enabled = False '鎖定 (已送出)
        Return v_OldData
    End Function

    ''' <summary>檔案打包下載</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_PACKAGE_DOWNLOAD1.Click
        'vActType=TIMS.ClearSQM(Convert.ToString(Request("act")))
        rPlanID = TIMS.ClearSQM(Convert.ToString(Request("PlanID")))
        rComIDNO = TIMS.ClearSQM(Convert.ToString(Request("cid")))
        rSeqNO = TIMS.ClearSQM(Convert.ToString(Request("no")))
        rSCDate = TIMS.ClearSQM(Convert.ToString(Request("CDate")))
        iSubSeqNO = If(TIMS.ClearSQM(Request("SubNo")) <> "", Val(TIMS.ClearSQM(Request("SubNo"))), 0)

        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        ' "DOWNLOAD4" '下載
        Call ResponseZIPFileALL_RVFL(Me)
        'SyncLock download_lock 'End SyncLock
    End Sub

    ''' <summary>檔案打包下載</summary>
    ''' <param name="MyPage"></param>
    Private Sub ResponseZIPFileALL_RVFL(MyPage As Page)
        Dim drPP As DataRow = TIMS.GetPCSDate(rPlanID, rComIDNO, rSeqNO, objconn)
        If drPP Is Nothing Then
            Common.MessageBox(Me, "計畫資訊有誤!")
            Return
        End If
        Dim drRV As DataRow = Get_REVISE2(objconn)
        If drRV Is Nothing Then
            Common.MessageBox(Me, "計畫變更資訊有誤!")
            Return
        End If

        Dim vAltDataID As String = Convert.ToString(drRV("AltDataID"))
        Dim vYEARS As String = Convert.ToString(drPP("PlanYear"))
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(sm.UserInfo.Years)
        Dim vDISTNAME3 As String = TIMS.GET_DISTNAME3(objconn, drPP("DISTID"))
        Dim vORGNAME As String = Convert.ToString(drPP("ORGNAME"))
        Dim rSCDateNT As String = TIMS.Cdate3(drRV("CDATE"), "yyyyMMdd")
        Dim Template_ZipPath2 As String = TIMS.GET_Template_ZipPath2(rSCDateNT)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath2)) Then Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath2))

        Dim hPMS As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}, {"CDATE", rSCDate}, {"SUBSEQNO", iSubSeqNO}, {"ALTDATAID", Val(vAltDataID)}}
        'PLAN_REVISESUBFL
        Dim dtFL As DataTable = TIMS.GET_dtPLAN_REVISESUBFL(objconn, hPMS)
        If dtFL IsNot Nothing AndAlso dtFL.Rows.Count > 0 Then
            For Each drFL As DataRow In dtFL.Rows
                Dim vRVNAME3 As String = ""
                Dim oFILENAME1 As String = "" 'Convert.ToString(drFL("FILENAME1"))
                Dim oFILEPATH1 As String = ""
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                '年度申請階段_單位名稱_項目編號+項目名稱
                Dim t_FILENAME As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf")
                Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))
                'Dim t_FilePath1 As String=Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
                Try
                    vRVNAME3 = Convert.ToString(drFL("RVNAME3"))
                    oFILENAME1 = Convert.ToString(drFL("FILENAME1"))
                    oFILEPATH1 = Convert.ToString(drFL("FILEPATH1"))
                    oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNO, rSCDateNT, CStr(iSubSeqNO)))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度_單位名稱_項目編號名稱
                    t_FILENAME = String.Concat(vYEARS_ROC, "_", vORGNAME, "_", vRVNAME3, ".pdf")
                    t_FilePath1 = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME))
                    If oFILENAME1 <> "" AndAlso IO.File.Exists(s_FilePath1) Then
                        Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                        File.WriteAllBytes(t_FilePath1, dbyte)
                    End If
                Catch ex As Exception
                    Dim strErrmsg As String = "/*Sub ResponseZIPFileALL(ByRef MyPage As Page)*/" & vbCrLf
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                    strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("t_FILENAME: ", t_FILENAME, vbCrLf)
                    strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        'PLAN_REVISESUBFL_TT
        Dim hPMS2 As New Hashtable From {{"PLANID", rPlanID}, {"COMIDNO", rComIDNO}, {"SEQNO", rSeqNO}, {"CDATE", rSCDate}, {"SUBSEQNO", iSubSeqNO}, {"ALTDATAID", Val(vAltDataID)}}
        Dim dtFLTT As DataTable = TIMS.GET_dtPLAN_REVISESUBFL_TT(objconn, hPMS2)
        If dtFLTT IsNot Nothing AndAlso dtFLTT.Rows.Count > 0 Then
            For Each drFLTT As DataRow In dtFLTT.Rows
                Dim vRVNAME3 As String = ""
                Dim vTEACHCNAME As String = "" 'Convert.ToString(drFLTT("TEACHCNAME"))
                Dim vTEACHERID As String = "" 'Convert.ToString(drFLTT("TEACHERID"))
                Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                Dim oFILEPATH1 As String = ""
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                Dim t_FILENAME_TT As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT))
                Try
                    vRVNAME3 = Convert.ToString(drFLTT("RVNAME3"))
                    vTEACHCNAME = Convert.ToString(drFLTT("TEACHCNAME"))
                    vTEACHERID = Convert.ToString(drFLTT("TEACHERID"))
                    oFILENAME1 = Convert.ToString(drFLTT("FILENAME1"))
                    oFILEPATH1 = Convert.ToString(drFLTT("FILEPATH1"))
                    oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH_PR2(vYEARS, rPlanID, rComIDNO, rSeqNO, rSCDateNT, CStr(iSubSeqNO)))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                    t_FILENAME_TT = String.Concat(vYEARS_ROC, "_", vORGNAME, "_", vRVNAME3, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                    t_FilePath1 = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME_TT))
                    If IO.File.Exists(s_FilePath1) Then
                        Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                        File.WriteAllBytes(t_FilePath1, dbyte)
                    End If
                Catch ex As Exception
                    Dim strErrmsg As String = "/*Sub ResponseZIPFileALL(ByRef MyPage As Page)*/" & vbCrLf
                    strErrmsg &= String.Concat("vTEACHCNAME: ", vTEACHCNAME, vbCrLf, "vTEACHERID: ", vTEACHERID, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("t_FILENAME_TT: ", t_FILENAME_TT, vbCrLf)
                    strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim zipFileName As String = String.Concat("r", vYEARS_ROC, "_", vDISTNAME3, "_", vORGNAME, "_", rSCDateNT, "_", strNOW, ".zip")
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath2)) Then
            Common.MessageBox(Me, String.Concat(Template_ZipPath2, "下載檔案資料夾有誤!"))
            Return
        End If
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
                    Dim strErrmsg As String = "/*ResponseZIPFileALL_RVFL*/" & vbCrLf
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
            TIMS.SAVE_ADP_ZIPFILE(objconn, "-rvfl4596", File)
            ' Clear the content of the response
            .Response.ClearContent()
            ' LINE1 Add the file name And attachment, which will force the open/cance/save dialog To show, to the header
            .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
            'Response.Headers["Content-Disposition"]="attachment; filename=" + zipFileName;
            ' Add the file size into the response header
            .Response.AddHeader("Content-Length", File.Length.ToString())
            ' Set the ContentType
            .Response.ContentType = "application/zip"
            .Response.TransmitFile(File.FullName)
            ' End the response
            TIMS.Utl_RespWriteEnd(MyPage, objconn, "") '.Response.End()
        End With

    End Sub

    'changSTATUS4
    'Protected Sub btnChkSend4_Click(sender As Object, e As EventArgs) Handles btnChkSend4.Click
    'End Sub
End Class
