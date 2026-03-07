Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Partial Class TC_11_002
    Inherits AuthBasePage

    '基本上分署單位使用
    '申請階段管理-受理期間設定 APPLISTAGE
    'Dim fg_can_applistage As Boolean = False
    Dim tryFIND As String = ""

    Const cst_MaxLen500_i As Integer = 500
    Const cst_MaxLen500_TPMSG1 As String = "限定500字元"

    '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
    'Dim vAPPLIEDRESULT As String = ""

    Dim iDG11_ROWS As Integer = 0
    Dim iDG10_ROWS As Integer = 0
    '以目前版本批次送出
    Const cst_txt_版本批次送出 As String = "(版本批次送出)"
    Const cst_txt_免附文件 As String = "(免附文件)"
    Const cst_08_訓練班別計畫表_WAIVED_PI As String = "PI"
    Const cst_08_1_iCap課程原始申請資料_WAIVED_PI3 As String = "PI3"
    Const cst_10_師資助教基本資料表_WAIVED_TT As String = "TT"
    Const cst_11_授課師資學經歷證書影本_WAIVED_TT2 As String = "TT2"
    Const cst_13_教學環境資料表_WAIVED_PI2 As String = "PI2" 'teaching ENVIRonment data sheet
    ''' <summary>W13-1混成課程教學環境資料表，遠距改混成，</summary>
    Const cst_13_1混成課程教學環境資料表_WAIVED_RT2 As String = "RT2"

    'add [TECHENV] [varchar] (1)  COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
    '[TECHENVACCT] [varchar](15)  COLLATE Chinese_Taiwan_Stroke_CS_AS NULL ,
    '[TECHENVDATE] [datetime]  NULL 

    Const cst_txt_師資助教基本資料表 As String = "(師資助教基本資料表)"
    Const cst_txt_授課師資學經歷證書影本 As String = "(師資學經歷證書影本)"
    Const cst_txt_教學環境資料表 As String = "(教學環境資料表)"
    Const cst_txt_混成課程教學環境資料表 As String = "(混成課程教學環境資料表)"
    Const cst_txt_iCap課程原始申請資料 As String = "(iCap課程原始申請資料)"

    'Const cst_printASPX_R As String = "../../SD/14/SD_14_002_R.aspx?ID="
    'Dim sPrintASPX1 As String = ""

    'outTYPE: CLSNM,PCSVAL
    'Const cst_outTYPE_CLSNM As String = "CLSNM"
    'Const cst_outTYPE_PCSVAL As String = "PCSVAL"

    Const cst_ss_RqProcessType As String = "RqProcessType"
    'Const cst_DG1CMDNM_VIEW1 As String = "VIEW1"
    Const cst_DG1CMDNM_EDIT1 As String = "EDIT1" '審核/審查/確認
    Const cst_DG1CMDNM_REVERT2 As String = "REVERT2" '還原"-確認

    Const cst_DG2CMDNM_RtuBACK1 As String = "RtuBACK1" '退回開放修改
    Const cst_DG2CMDNM_REVERT1 As String = "REVERT1" '還原"
    Const cst_DG2CMDNM_VIEWFILE4 As String = "VIEWFILE4" '查詢
    Const cst_DG2CMDNM_DOWNLOAD4 As String = "DOWNLOAD4" '下載

    'Dim G_UPDRV As String = "~/UPDRV"
    'Dim G_UPDRV_JS As String = "../../UPDRV"

    'Const cst_errMsg_1 As String = "資料有誤請重新查詢!"
    'Const cst_errMsg_2 As String = "上傳檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)" 'Const cst_errMsg_2 As String = "上傳檔案壓縮時發生錯誤，請重新確認上傳檔案格式!"
    'Const cst_errMsg_3 As String = "檔案位置錯誤!"
    'Const cst_errMsg_4 As String = "檔案類型錯誤!"
    'Const cst_errMsg_5 As String = "檔案類型錯誤，必須為PDF類型檔案!"
    Const cst_errMsg_6 As String = "(檔案上傳失敗／異常，請刪除後重新上傳)"
    'Const cst_PostedFile_MAX_SIZE As Integer = 2097152 '10*1024*1024 '2*1024*1024
    'Const cst_errMsg_7 As String = "檔案大小超過2MB!"
    'Const cst_errMsg_8 As String = "請選擇上傳檔案(不可為空)!"
    ''Const cst_errMsg_9 As String = "請選擇場地圖片--隸屬於教室1 或教室2!"
    'Const cst_errMsg_11 As String = "無效的檔案格式。"
    'Const cst_errMsg_21 As String = "不可勾選免附文件又按上傳檔案。"

    'Const cst_G01_TTQS評核證書影本 As String = "G01"
    'Const cst_G02_設立證明文件影本 As String = "G02"
    'Const cst_G03_組織章程影本 As String = "G03"
    'Const cst_G04_法人登記證書 As String = "G04"
    'Const cst_G05_辦理本計畫訓練課程之專職人員名冊 As String = "G05"
    'Const cst_G06_訓練單位基本資料表 As String = "G06"
    'Const cst_G07_訓練計畫總表 As String = "G07"
    'Const cst_G08_訓練班別計畫表 As String = "G08"
    'Const cst_G09_訓練計畫師資助教名冊 As String = "G09"
    'Const cst_G10_師資助教基本資料表 As String = "G10"
    'Const cst_G11_授課師資學經歷證書影本 As String = "G11"
    'Const cst_G12_訓練計畫場地資料表 As String = "G12"
    'Const cst_G13_教學環境資料表 As String = "G13"
    'Const cst_G14_消防安全設備檢修申報受理單影本 As String = "G14"
    'Const cst_G15_建築物防火避難設施與設備安全檢查申報結果通知書影本 As String = "G15"
    'Const cst_G16_機關構同意租借證明文件 As String = "G16"
    'Const cst_G17_目的事業主管機關核備開課之文件 As String = "G17"
    'Const cst_G18_報請教育部核准校外場地教學函 As String = "G18"
    'Const cst_G19_送件檢核表 As String = "G19"
    'Const cst_G20_公文 As String = "G20"

    'Const cst_W01_TTQS評核證書影本 As String = "W01"
    'Const cst_W02_設立登記影本 As String = "W02"
    'Const cst_W03_組織章程影本 As String = "W03"
    'Const cst_W04_辦理本計畫訓練課程之專職人員名冊 As String = "W04"
    'Const cst_W05_1_勞工團體組織名錄 As String = "W05-1"
    'Const cst_W05_2_無法代轉或不願代轉基層團體計畫文件 As String = "W05-2"
    'Const cst_W06_訓練單位基本資料表 As String = "W06"
    'Const cst_W07_訓練計畫總表 As String = "W07"
    'Const cst_W08_訓練班別計畫表 As String = "W08"
    'Const cst_W09_訓練計畫師資助教名冊 As String = "W09"
    'Const cst_W10_師資助教基本資料表 As String = "W10"
    'Const cst_W11_授課師資學經歷證書影本 As String = "W11"
    'Const cst_W12_訓練計畫場地資料表 As String = "W12"
    'Const cst_W13_教學環境資料表 As String = "W13"
    'Const cst_W14_消防安全設備檢修申報受理單影本 As String = "W14"
    'Const cst_W15_建築物防火避難設施與設備安全檢查申報結果通知書影本 As String = "W15"
    'Const cst_W16_機關構同意租借證明文件 As String = "W16"
    'Const cst_W17_目的事業主管機關核備開課之文件 As String = "W17"
    'Const cst_W18_送件檢核表 As String = "W18"
    'Const cst_W19_公文 As String = "W19"

    'KEY_BIDCASE
    'ORG_BIDCASE,ORG_BIDCASEPI,ORG_BIDCASEFL,ORG_BIDCASEFL_TT,ORG_BIDCASEFL_TT2
    'VIEW2B
    Dim ff3 As String = ""
    'Dim dtKEY_BIDCASE As DataTable
    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Dim url2 As String = "TC_11_002_54.aspx" '(54:充電起飛計畫（在職）)切換程式
        If TIMS.Cst_TPlanID54.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            TIMS.Utl_Redirect1(Me, url2)
            Return
        End If

        Call cCreate11() '每次執行

        If Not IsPostBack Then
            Call cCreate1(0)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button3.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            'HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If
    End Sub

    ''' <summary>每次執行</summary>
    Sub cCreate11()
        Call TIMS.OpenDbConn(objconn)
        'OJT-20231128:線上送件審核-還原按鈕 NULL/Y:可使用
        Hid_USE_ORG_BIDCASE_REVERT2.Value = TIMS.Utl_GetConfigVAL(objconn, "USE_ORG_BIDCASE_REVERT2")

        PageControler1.PageDataGrid = DataGrid1 '分頁設定
        'PageControler1.PageDataTable = dt
        'PageControler1.ControlerLoad()

        '申請階段管理-受理期間設定 APPLISTAGE
        'fg_can_applistage = TIMS.CAN_APPLISTAGE_1(objconn)

        'sPrintASPX1 = String.Concat(cst_printASPX_R, TIMS.Get_MRqID(Me))
        '<add key = "UPLOAD_OJT_Path" value="~/UPDRV" />
        '<add key = "DOWNLOAD_OJT_Path" value="../../UPDRV" />
        'Dim vUPLOAD_OJT_Path As String = TIMS.Utl_GetConfigSet("UPLOAD_OJT_Path")
        'Dim vDOWNLOAD_OJT_Path As String = TIMS.Utl_GetConfigSet("DOWNLOAD_OJT_Path")
        'Dim G_UPDRV As String = "~/UPDRV"
        'Dim G_UPDRV_JS As String = "../../UPDRV"
        'If (vUPLOAD_OJT_Path <> "") Then G_UPDRV = vUPLOAD_OJT_Path
        'If (vDOWNLOAD_OJT_Path <> "") Then G_UPDRV_JS = vDOWNLOAD_OJT_Path
    End Sub

    '設定 資料與顯示 狀況！
    Private Sub cCreate1(ByVal iNum As Integer)
        TableDataGrid1.Visible = False
        labmsg1.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        tr_LabSwitchTo.Visible = False
        tr_DataGrid10.Visible = False
        tr_DataGrid11.Visible = False
        tr_DataGrid13.Visible = False
        tr_DataGrid13B.Visible = False
        tr_DataGrid14.Visible = False

        Call SHOW_Frame1(0)
        '案件編號
        sch_txtBCASENO.Text = ""
        '計畫年度
        'Dim iEYearsN As Integer = Year(Now) - 1
        'Dim iEYears As Integer = If(iEYearsN > iSYears, iEYearsN, iSYears)
        Dim iSYears As Integer = 2023 '(起始年度)
        Dim iYearsLast As Integer = If((sm.UserInfo.Years + 1) > (Year(Now) + 1), sm.UserInfo.Years + 1, Year(Now) + 1)
        Dim iSYears2 As Integer = (iSYears + 2)
        Dim iEYears As Integer = If(iYearsLast > iSYears2, iYearsLast, iSYears2)

        sch_ddlYEARS = TIMS.GetSyear(sch_ddlYEARS, iSYears, iEYears, True)
        Common.SetListItem(sch_ddlYEARS, sm.UserInfo.Years)
        '1.【年度】鎖定登入年度，不可改 (避免登入錯誤年度操作)
        sch_ddlYEARS.Enabled = False
        TIMS.Tooltip(sch_ddlYEARS, "鎖定登入年度")

        '申請階段 'Dim v_APPSTAGE As String = If(Now.Month < 7, "1", "2")
        Dim v_APPSTAGE As String = TIMS.GET_CANUSE_APPSTAGE(objconn, CStr(sm.UserInfo.Years), TIMS.cst_APPSTAGE_PTYPE1_01)
        sch_ddlAPPSTAGE = TIMS.Get_APPSTAGE2(sch_ddlAPPSTAGE)
        Common.SetListItem(sch_ddlAPPSTAGE, v_APPSTAGE)

        '申辦人姓名
        sch_txtBINAME.Text = ""
        '申辦日期
        sch_txtBIDATE1.Text = ""
        sch_txtBIDATE2.Text = ""

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Get_TitleLab(objconn, MRqID, TitleLab1, TitleLab2)

        'If dtKEY_BIDCASE Is Nothing Then
        '    Dim rSql As String = "SELECT KBSID,KBID,KBNAME FROM KEY_BIDCASE"
        '    dtKEY_BIDCASE = DbAccess.GetDataTable(rSql, objconn)
        'End If
    End Sub

    ''' <summary>顯示調整</summary>
    ''' <param name="iNum"></param>
    Private Sub SHOW_Frame1(ByVal iNum As Integer)
        FrameTableSch1.Visible = False
        FrameTableEdt1.Visible = False
        If iNum = 0 Then
            FrameTableSch1.Visible = True
        ElseIf iNum = 1 Then
            FrameTableEdt1.Visible = True
        End If
    End Sub

    ''' <summary>查詢鈕1</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH1_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH1.Click
        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        '清理隱藏的參數
        Call sSearch1()
    End Sub

    ''' <summary>查詢1</summary>
    Private Sub sSearch1()
        '清理隱藏的參數
        Call ClearHidValue()

        Call SHOW_Frame1(0)
        labmsg1.Text = TIMS.cst_NODATAMsg1
        TableDataGrid1.Visible = False

        'RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If
        'orgid_value.VALUE = TIMS.ClearSQM(orgid_value.VALUE)
        sch_txtBCASENO.Text = TIMS.ClearSQM(sch_txtBCASENO.Text)
        sch_txtBINAME.Text = TIMS.ClearSQM(sch_txtBINAME.Text)
        sch_txtBIDATE1.Text = TIMS.Cdate3(sch_txtBIDATE1.Text)
        sch_txtBIDATE2.Text = TIMS.Cdate3(sch_txtBIDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(sch_txtBIDATE1.Text, sch_txtBIDATE2.Text) Then
            Dim T_DATE1 As String = sch_txtBIDATE1.Text
            sch_txtBIDATE1.Text = sch_txtBIDATE2.Text
            sch_txtBIDATE2.Text = T_DATE1
        End If

        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS) '計畫年度
        Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE) '申請階段

        '檢核查詢
        'Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        'Dim flag_CHECKOK As Boolean = CHK_Search1(drRR)
        'If Not flag_CHECKOK OrElse drRR Is Nothing Then Return
        'Dim v_ORGID As String = sm.UserInfo.OrgID
        'Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        'Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        'Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        'Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        'Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))

        Dim vDISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        Dim pParms As New Hashtable
        If sch_txtBCASENO.Text <> "" Then pParms.Add("BCASENO", sch_txtBCASENO.Text)
        If v_sch_ddlYEARS <> "" Then pParms.Add("YEARS", v_sch_ddlYEARS)
        If v_sch_ddlAPPSTAGE <> "" Then pParms.Add("APPSTAGE", v_sch_ddlAPPSTAGE)
        If sch_txtBINAME.Text <> "" Then pParms.Add("BINAME", sch_txtBINAME.Text) 'sSql &= " AND u.NAME LIKE '%'+@BINAME+'%'" & vbCrLf
        If sch_txtBIDATE1.Text <> "" Then pParms.Add("BIDDATE1", sch_txtBIDATE1.Text) 'sSql &= " AND a.BIDDATE >=@BIDDATE1" & vbCrLf
        If sch_txtBIDATE2.Text <> "" Then pParms.Add("BIDDATE2", sch_txtBIDATE2.Text) 'sSql &= " AND a.BIDDATE <=@BIDDATE2" & vbCrLf

        Dim sSql As String = ""
        sSql &= " SELECT a.BCID,a.BCASENO,a.YEARS,a.DISTID,a.ORGID,a.PLANID,a.RID,a.APPSTAGE" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        'APPSTAGE_N
        sSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        sSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上' WHEN 2 THEN '下' WHEN 3 THEN '政' WHEN 4 THEN '進' END APPSTAGE_S" & vbCrLf
        'DISTNAME
        sSql &= " ,dbo.FN_GET_DISTNAME(a.DISTID,3) DISTNAME" & vbCrLf
        'ORGNAME
        sSql &= " ,(SELECT ORGNAME FROM ORG_ORGINFO WHERE ORGID=a.ORGID) ORGNAME" & vbCrLf
        'BINAME
        sSql &= " ,a.BIDACCT,dbo.FN_GET_USERNAME(a.BIDACCT) BINAME" & vbCrLf
        'ORGKINDGW
        sSql &= " ,(SELECT x.ORGKINDGW FROM VIEW_RIDNAME x WHERE x.RID=a.RID) ORGKINDGW" & vbCrLf
        sSql &= " ,format(a.BIDDATE,'yyyy/MM/dd') BIDDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.BIDDATE) BIDDATE_ROC" & vbCrLf
        '申辦狀態：暫存/ 已送件
        sSql &= " ,a.BISTATUS" & vbCrLf
        sSql &= " , CASE WHEN a.BISTATUS IS NULL THEN '暫存'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='Y' THEN '通過'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END BISTATUS_N" & vbCrLf
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL"
        sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
        'sSql &= " ,a.CREATEACCT,a.CREATEDATE,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= " ,u.NAME BINAME" & vbCrLf
        sSql &= " FROM ORG_BIDCASE a" & vbCrLf
        sSql &= " JOIN AUTH_ACCOUNT u ON u.ACCOUNT=a.BIDACCT" & vbCrLf
        sSql &= " JOIN ID_PLAN ip ON ip.PLANID=a.PLANID" & vbCrLf
        'sSql &= " WHERE a.BISTATUS IS NOT NULL AND a.BISTATUS in ('B','R')" & vbCrLf
        sSql &= " WHERE a.BISTATUS IS NOT NULL AND a.DISTID=@DISTID AND ip.TPLANID=@TPLANID" & vbCrLf
        pParms.Add("DISTID", vDISTID)
        pParms.Add("TPLANID", sm.UserInfo.TPlanID)
        If RIDValue.Value.Length <> 1 Then
            sSql &= " AND a.RID=@RID" & vbCrLf
            pParms.Add("RID", RIDValue.Value)
        End If
        If sch_txtBCASENO.Text <> "" Then sSql &= " AND a.BCASENO=@BCASENO" & vbCrLf
        If v_sch_ddlYEARS <> "" Then sSql &= " AND a.YEARS=@YEARS" & vbCrLf
        If v_sch_ddlAPPSTAGE <> "" Then sSql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        If sch_txtBINAME.Text <> "" Then sSql &= " AND u.NAME LIKE '%'+@BINAME+'%'" & vbCrLf
        If sch_txtBIDATE1.Text <> "" Then sSql &= " AND a.BIDDATE>=@BIDDATE1" & vbCrLf
        If sch_txtBIDATE2.Text <> "" Then sSql &= " AND a.BIDDATE<=@BIDDATE2" & vbCrLf

        'a.RID=@RID AND  sSql &= " AND a.BISTATUS IS NOT NULL" & vbCrLf
        ' a.BISTATUS WHEN 'B' THEN '已送件' WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過
        'BISTATUS: 申辦狀態：NULL:暫存/B:已送件/R:退件修正
        'APPLIEDRESULT: 審查狀態：Y:申辦確認/R:申辦退件修正/N:申辦不通過
        Dim v_rbAPPLIEDRESULT As String = TIMS.GetListValue(rbAPPLIEDRESULT)
        If v_rbAPPLIEDRESULT = "A" Then
            sSql &= " AND a.BISTATUS IN ('R','B')" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "B" Then
            sSql &= " AND a.APPLIEDRESULT IS NULL" & vbCrLf
            sSql &= " AND a.BISTATUS='B'" & vbCrLf
            'sSql &= " AND (a.APPLIEDRESULT='R' OR a.APPLIEDRESULT IS NULL)" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "R" Then
            sSql &= " AND a.APPLIEDRESULT='R'" & vbCrLf
            sSql &= " AND (a.BISTATUS='R' OR a.BISTATUS='B')" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "Y" OrElse v_rbAPPLIEDRESULT = "N" Then 'Y:申辦確認 / N:申辦不通過
            pParms.Add("APPLIEDRESULT", v_rbAPPLIEDRESULT)
            sSql &= " AND a.APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        End If
        sSql &= " ORDER BY a.BCID DESC" & vbCrLf

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(pParms), vbCrLf, "--##TC_11_002:", vbCrLf, sSql))
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            labmsg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        labmsg1.Text = ""
        TableDataGrid1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        'DataGrid1.DataSource = dt
        'DataGrid1.DataBind()
    End Sub

    ''' <summary>清理隱藏的參數</summary>
    Sub ClearHidValue()
        Hid_KBSID.Value = ""
        Hid_KBID.Value = ""
        Hid_TECHID.Value = ""
        Hid_BCID.Value = ""
        Hid_RID.Value = ""
        Hid_BCFID.Value = ""
        Hid_BCASENO.Value = ""
        Hid_ORGKINDGW.Value = ""
        Hid_PCS.Value = ""
        Hid_APPLIEDRESULT.Value = ""
    End Sub

    ''' <summary>按下 審查</summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vBCID As String = TIMS.GetMyValue(sCmdArg, "BCID")
        Dim vRID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim vBCASENO As String = TIMS.GetMyValue(sCmdArg, "BCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        If sCmdArg = "" OrElse vBCID = "" OrElse vRID = "" Then Return

        Dim drRR As DataRow = TIMS.Get_RID_DR(vRID, objconn)
        If drRR Is Nothing Then
            Common.MessageBox(Me, "業務資訊有誤，請重新選擇訓練機構!")
            Return
        End If
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "業務資訊有誤，請重新選擇訓練機構!!")
            Return
        End If

        'Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        'Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        ''申請階段管理-受理期間設定 APPLISTAGE
        'Dim aParms As New Hashtable
        'aParms.Add("YEARS", vYEARS)
        'aParms.Add("APPSTAGE", vAPPSTAGE)
        'Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_1(objconn, aParms)

        Select Case e.CommandName
            Case cst_DG1CMDNM_EDIT1 '"EDIT1"審核
                'If Not fg_can_applistage Then
                '    Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                '    Return
                'End If
                If Convert.ToString(drOB("BISTATUS")) = "R" Then
                    Common.MessageBox(Me, "申辦狀態 退件待修正，待修正後送審，再行審核!")
                    Return
                ElseIf Convert.ToString(drOB("BISTATUS")) = "" Then
                    Common.MessageBox(Me, "申辦狀態 未填寫，待送審後，再行審核!")
                    Return
                End If
                '查詢使用資料顯示 依 ORG_BIDCASE-BCID
                Call SHOW_Detail_BIDCASE(drRR, vBCID, cst_DG1CMDNM_EDIT1) '"EDIT1"審核

            Case cst_DG1CMDNM_REVERT2 '"REVERT2" '還原"-確認
                '當【審核狀態】：已申辦確認、已申辦不通過，按下【還原】鈕即清空【審核狀態】且【申辦狀態】：已送件， 如圖二狀態
                '當【審核狀態】：退件修正，【還原】鈕反灰不可按
                If Convert.ToString(drOB("APPLIEDRESULT")) = "R" Then
                    Common.MessageBox(Me, "審核狀態 退件待修正，不可還原!")
                    Return
                ElseIf Convert.ToString(drOB("APPLIEDRESULT")) = "" Then
                    Common.MessageBox(Me, "審核狀態 未填寫，不可還原!")
                    Return
                End If

                Const cst_REVERT2_N As String = "還原"
                Dim s_HISREVIEW As String = Convert.ToString(drOB("HISREVIEW"))
                s_HISREVIEW &= String.Concat(If(s_HISREVIEW <> "", "，", ""), String.Concat(TIMS.Cdate3t(Now), "-", cst_REVERT2_N))
                '線上送件審核，新增還原按鈕 UPDATE ORG_BIDCASE / BISTATUS='B',APPLIEDRESULT=NULL
                Call UPDATE_REVERT2(s_HISREVIEW, vRID, vBCID, vBCASENO)
                '查詢1
                Call sSearch1()

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                Dim drv As DataRowView = e.Item.DataItem
                Dim lBTN_EDIT1 As LinkButton = e.Item.FindControl("lBTN_EDIT1") '審核/審查/確認
                Dim lBTN_REVERT2 As LinkButton = e.Item.FindControl("lBTN_REVERT2") '還原
                'OJT-20231128:線上送件審核-還原按鈕 NULL/Y:可使用
                lBTN_REVERT2.Visible = (Hid_USE_ORG_BIDCASE_REVERT2.Value = "Y")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "BCID", drv("BCID"))
                TIMS.SetMyValue(sCmdArg, "RID", drv("RID"))
                TIMS.SetMyValue(sCmdArg, "BCASENO", drv("BCASENO"))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", drv("ORGKINDGW"))

                Dim s_APPLIEDRESULT As String = Convert.ToString(drv("APPLIEDRESULT"))
                Dim tipMsg1 As String = Convert.ToString(drv("APPLIEDRESULT_N"))
                lBTN_EDIT1.CommandArgument = sCmdArg '審核
                TIMS.Tooltip(lBTN_EDIT1, tipMsg1, True)

                If lBTN_REVERT2.Visible Then
                    '當【審核狀態】：已申辦確認、已申辦不通過，按下【還原】鈕即清空【審核狀態】且【申辦狀態】：已送件，如圖二狀態
                    '當【審核狀態】：退件修正，【還原】鈕反灰不可按
                    Dim fg_CAN_REVERT1 As Boolean = (s_APPLIEDRESULT = "Y" OrElse s_APPLIEDRESULT = "N")
                    lBTN_REVERT2.Enabled = fg_CAN_REVERT1
                    lBTN_REVERT2.CommandArgument = sCmdArg '審核
                    If lBTN_REVERT2.Enabled AndAlso tipMsg1 <> "" Then
                        TIMS.Tooltip(lBTN_REVERT2, String.Concat(tipMsg1, "，【還原】可使用"), True)
                    ElseIf Not lBTN_REVERT2.Enabled AndAlso s_APPLIEDRESULT = "R" Then
                        tipMsg1 = "退件修正，【還原】不可使用"
                        TIMS.Tooltip(lBTN_REVERT2, tipMsg1, True)
                    ElseIf Not lBTN_REVERT2.Enabled AndAlso s_APPLIEDRESULT = "" Then
                        tipMsg1 = "尚未審核，【還原】不可使用"
                        TIMS.Tooltip(lBTN_REVERT2, tipMsg1, True)
                    End If
                End If
        End Select
    End Sub

    ''' <summary>新增使用資料顯示／查詢使用資料顯示 依 ORG_BIDCASE-BCID</summary>
    Private Sub SHOW_Detail_BIDCASE(ByRef drRR As DataRow, ByVal vBCID As String, ByVal vCmdName As String)
        '訓練機構有誤
        If drRR Is Nothing Then Return
        Call SHOW_Frame1(1)
        If vBCID = "" Then
            Common.MessageBox(Me, "傳入參數為空，異常!")
            Return
        End If

        Session(cst_ss_RqProcessType) = vCmdName
        Dim vRID As String = Convert.ToString(drRR("RID"))
        Dim vPLANID As String = Convert.ToString(drRR("PLANID"))
        Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))

        tr_HISREVIEW.Visible = False '歷程資訊

        Dim dtB1 As DataTable = Nothing
        '查詢資料
        'NULL 待送審、B 審核中、Y 審核通過、R 退件修正、N 審核不通過。
        Dim pParms As New Hashtable From {{"BCID", vBCID}, {"RID", vRID}, {"PLANID", vPLANID}}
        'pParms.Add("YEARS", v_sch_ddlYEARS)
        'pParms.Add("APPSTAGE", v_sch_ddlAPPSTAG)
        Dim sSql As String = ""
        sSql &= " SELECT a.BCID,a.BCASENO,a.YEARS,a.PLANID,a.DISTID,a.ORGID,a.RID,a.APPSTAGE" & vbCrLf
        'APPSTAGE_N
        sSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        'DISTNAME
        sSql &= " ,dbo.FN_GET_DISTNAME(a.DISTID,3) DISTNAME" & vbCrLf
        'ORGNAME
        sSql &= " ,(SELECT ORGNAME FROM ORG_ORGINFO WHERE ORGID=a.ORGID) ORGNAME" & vbCrLf
        'COMIDNO
        sSql &= " ,(SELECT COMIDNO FROM ORG_ORGINFO WHERE ORGID=a.ORGID) COMIDNO" & vbCrLf
        'BINAME
        sSql &= " ,a.BIDACCT,dbo.FN_GET_USERNAME(a.BIDACCT) BINAME" & vbCrLf
        'ORGKINDGW
        sSql &= " ,(SELECT x.ORGKINDGW FROM VIEW_RIDNAME x WHERE x.RID=a.RID) ORGKINDGW" & vbCrLf
        sSql &= " ,format(a.BIDDATE,'yyyy/MM/dd') BIDDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.BIDDATE) BIDDATE_ROC" & vbCrLf
        '申辦狀態：暫存/ 已送件
        sSql &= " ,a.BISTATUS" & vbCrLf
        sSql &= " , CASE WHEN a.BISTATUS IS NULL THEN '暫存'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='Y' THEN '通過'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END BISTATUS_N" & vbCrLf
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL"
        sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
        'tr_HISREVIEW.Visible = False '歷程資訊
        sSql &= " ,a.HISREVIEW" & vbCrLf
        '返回第1項
        sSql &= " ,(SELECT MIN(KBSID) FROM ORG_BIDCASEFL fl WHERE fl.BCID=a.BCID) CurrentKBSID" & vbCrLf
        'sSql &= " ,a.CREATEACCT,a.CREATEDATE,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= ",a.APPLIEDRESULT,a.REASONFORFAIL"
        sSql &= " FROM ORG_BIDCASE a" & vbCrLf
        sSql &= " WHERE BCID=@BCID AND a.RID=@RID AND a.PLANID=@PLANID" & vbCrLf
        dtB1 = DbAccess.GetDataTable(sSql, objconn, pParms)
        If dtB1 Is Nothing OrElse dtB1.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無有效資料，異常!")
            Return
        End If

        Dim drB1 As DataRow = dtB1.Rows(0)

        Dim vAPPSTAGE As String = Convert.ToString(drB1("APPSTAGE"))

        Common.SetListItem(ddlAPPLIEDRESULT, Convert.ToString(drB1("APPLIEDRESULT")))
        Reasonforfail.Text = Convert.ToString(drB1("REASONFORFAIL"))
        TIMS.Tooltip(Reasonforfail, cst_MaxLen500_TPMSG1, True)

        Hid_RID.Value = Convert.ToString(drB1("RID"))
        Hid_BCID.Value = Convert.ToString(drB1("BCID"))
        Hid_BCASENO.Value = Convert.ToString(drB1("BCASENO"))
        Hid_ORGKINDGW.Value = Convert.ToString(drB1("ORGKINDGW"))

        labOrgNAME.Text = Convert.ToString(drB1("ORGNAME"))
        labBCASENO.Text = Convert.ToString(drB1("BCASENO"))
        labBIYEARS.Text = TIMS.GET_YEARS_ROC(drB1("YEARS"))
        labAPPSTAGE.Text = Convert.ToString(drB1("APPSTAGE_N"))

        Hid_KBSID.Value = Convert.ToString(drB1("CurrentKBSID"))

        'Dim t_ddlAPPLIEDRESULT As String = TIMS.GetListText(ddlAPPLIEDRESULT)
        Hid_APPLIEDRESULT.Value = Convert.ToString(drB1("APPLIEDRESULT"))
        Dim fg_have_APPLIEDRESULT As Boolean = (Hid_APPLIEDRESULT.Value = "N" OrElse Hid_APPLIEDRESULT.Value = "Y")
        Dim t_ddlAPPLIEDRESULT As String = If(fg_have_APPLIEDRESULT, TIMS.GetListText(ddlAPPLIEDRESULT), "")
        Reasonforfail.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        ddlAPPLIEDRESULT.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        But_Sub.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        TIMS.Tooltip(Reasonforfail, t_ddlAPPLIEDRESULT, True)
        TIMS.Tooltip(ddlAPPLIEDRESULT, t_ddlAPPLIEDRESULT, True)
        TIMS.Tooltip(But_Sub, t_ddlAPPLIEDRESULT, True)

        '	於當年度/申請階段之開放班級申請期間，可新增申辦案件(非申辦期間無法新增)。
        '	同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能有一筆申辦案件。
        '	於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)。
        '	各項應備文件上傳會依項目依序分頁顯示，訓練單位也可跳頁填選。第一頁介面示意圖如下

        'sSql = " SELECT ri.RID,ri.ORGID,ri.ORGNAME" & vbCrLf
        'sSql &= " ,ri.YEARS,ri.DISTID,ri.TPLANID,ri.PLANID" & vbCrLf
        'sSql &= " ,ri.COMIDNO,ri.RELSHIP,ri.ORGKINDGW,ri.ORGLEVEL" & vbCrLf
        Dim S1Parms As New Hashtable
        S1Parms.Add("TPLANID", sm.UserInfo.TPlanID)
        S1Parms.Add("RID", vRID)
        S1Parms.Add("PLANID", vPLANID)
        S1Parms.Add("APPSTAGE", vAPPSTAGE)
        If vBCID <> "" Then S1Parms.Add("BCID", vBCID)
        Dim dtPP As DataTable = TIMS.GET_CLASS2S_BIdt(objconn, S1Parms)
        Dim strCLASSNAME2S As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_CLSNM)
        Dim v_PCS_Value As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_PCSVAL)
        labCLASSNAME2S.Text = TIMS.GetResponseWrite(strCLASSNAME2S)

        '歷程資訊
        If Convert.ToString(drB1("HISREVIEW")) <> "" AndAlso Convert.ToString(drB1("HISREVIEW")).Length > 1 Then
            tr_HISREVIEW.Visible = True '歷程資訊
            labHISREVIEW.Text = Convert.ToString(drB1("HISREVIEW"))
        End If

        Hid_PCS.Value = v_PCS_Value
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "RID", Hid_RID.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        TIMS.SetMyValue2(rPMS3, "BCASENO", Hid_BCASENO.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)
    End Sub

#Region "PRINT_1"
    ''' <summary>cst_W13_教學環境資料表 '"13" '教學環境資料表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_014(ByRef rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_014" '0:未轉班' 1:已轉班
        Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
        Dim TSTPRINT As String = If(sPrint_Test = "Y", "2", "1") '測試區2／'正式區1 

        'Const cst_printFN2 As String = "SD_14_014_1" '2:變更待審
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim selsqlstr As String = TIMS.GetMyValue2(rPMS, "selsqlstr") 'vPCS
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")

        Dim sfilename1 As String = "" 'cst_printFN1
        sfilename1 = cst_printFN1
        Dim sMyValue As String = ""
        sMyValue &= String.Concat("&Years=", vYEARS_ROC)
        sMyValue &= "&selsqlstr=" & selsqlstr
        sMyValue &= "&TPlanID=" & vTPlanID
        sMyValue &= "&SYears=" & vYEARS
        sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfilename1, sMyValue)
    End Sub
#End Region

#Region "NO USE"
    '教學環境資料表
    'Private Shared Function GET_ORG_BIDCASEFL_EV(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String = ""
    '    sSql &= " SELECT b.BCFEID,a.BCID,a.BCPID, a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,pp.PSNO28" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_EV b on b.BCPID=a.BCPID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
    '    sSql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID and pp.COMIDNO=a.COMIDNO and pp.SEQNO=a.SEQNO" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function

    'Private Shared Function GET_ORG_BIDCASEFL_TT(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String = ""
    '    sSql &= " SELECT a.BCID,f.BCFID" & vbCrLf
    '    sSql &= " ,b.BCFTID,b.TECHID" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,tt.TEACHCNAME,tt.TEACHERID" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASE a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCID=a.BCID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_TT b on b.BCFID=f.BCFID" & vbCrLf
    '    sSql &= " JOIN TEACH_TEACHERINFO tt on tt.TECHID=b.TECHID" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function

    'Private Shared Function GET_ORG_BIDCASEFL_TT2(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String = ""
    '    sSql = "" & vbCrLf
    '    sSql &= " SELECT a.BCID,f.BCFID" & vbCrLf
    '    sSql &= " ,b.BCFT2ID,b.TECHID" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,tt.TEACHCNAME,tt.TEACHERID" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASE a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCID=a.BCID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_TT2 b on b.BCFID=f.BCFID" & vbCrLf
    '    sSql &= " JOIN TEACH_TEACHERINFO tt on tt.TECHID=b.TECHID" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function

    '取得班級申請資料
    'Private Shared Function GET_ORG_BIDCASEFL_PI(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String = ""
    '    sSql &= " SELECT a.BCID,a.BCPID, a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,pp.PSNO28" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_PI b on b.BCPID=a.BCPID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
    '    sSql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID and pp.COMIDNO=a.COMIDNO and pp.SEQNO=a.SEQNO" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function
#End Region

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        '退回原因說明／ '退回開放修改／ '還原
        Dim txtRtuReason As TextBox = e.Item.FindControl("txtRtuReason") '退回原因說明
        Dim Btn_RtuBACK1 As Button = e.Item.FindControl("Btn_RtuBACK1") '退回開放修改
        Dim Btn_REVERT1 As Button = e.Item.FindControl("Btn_REVERT1") '還原
        Dim sCmdArg As String = e.CommandArgument
        Dim vBCID As String = TIMS.GetMyValue(sCmdArg, "BCID")
        Dim vBCFID As String = TIMS.GetMyValue(sCmdArg, "BCFID")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        If e.CommandArgument = "" OrElse vBCID = "" OrElse vBCFID = "" OrElse vKBID = "" OrElse vKBSID = "" Then Return

        If txtRtuReason IsNot Nothing Then txtRtuReason.Text = TIMS.ClearSQM(txtRtuReason.Text)

        Dim drFL As DataRow = TIMS.GET_ORG_BIDCASEFL(objconn, vBCID, vKBSID)
        If drFL Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!")
            Return
        End If

        Select Case e.CommandName
            Case cst_DG2CMDNM_RtuBACK1 '退回開放修改
                If txtRtuReason Is Nothing OrElse txtRtuReason.Text = "" Then
                    Common.MessageBox(Me, "請填寫 退回原因說明，重新操作!!")
                    Return
                End If

                Dim v_txtRtuReason As String = TIMS.Get_Substr1(txtRtuReason.Text, cst_MaxLen500_i)
                Dim uParms As New Hashtable
                uParms.Add("BCFID", TIMS.CINT1(vBCFID))
                uParms.Add("RTUREASON", If(v_txtRtuReason <> "", v_txtRtuReason, Convert.DBNull))
                uParms.Add("RTURESACCT", sm.UserInfo.UserID)
                Dim usSql As String = ""
                usSql &= " UPDATE ORG_BIDCASEFL" & vbCrLf
                usSql &= " SET RTUREASON=@RTUREASON,RTURESACCT=@RTURESACCT,RTURESDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE BCFID=@BCFID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Case cst_DG2CMDNM_REVERT1 '還原
                'txtRtuReason.Text = TIMS.ClearSQM(txtRtuReason.Text)
                Dim uParms As New Hashtable From {{"BCFID", Val(vBCFID)}}
                Dim usSql As String = ""
                usSql &= " UPDATE ORG_BIDCASEFL" & vbCrLf
                usSql &= " SET RTUREASON=NULL,RTURESACCT=NULL,RTURESDATE=NULL" & vbCrLf
                usSql &= " WHERE BCFID=@BCFID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Case cst_DG2CMDNM_VIEWFILE4
                '"VIEWFILE4" '查詢
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCFID", vBCFID)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                Call SHOW_BIDCASE_KBSID(rPMS4)

            Case cst_DG2CMDNM_DOWNLOAD4
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
                If drOB Is Nothing Then Return
                ' "DOWNLOAD4" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCFID", vBCFID)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return

        End Select
        If Not TIMS.OpenDbConn(objconn) Then Return
        '顯示檔案資料表
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "RID", Hid_RID.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        TIMS.SetMyValue2(rPMS3, "BCASENO", Hid_BCASENO.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)
        'Dim vBCFID As String = TIMS.GetMyValue(sCmdArg, "BCFID")
        'Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        'Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        'If e.CommandName = cst_DG2CMDNM_VIEWFILE4 Then
        'ElseIf e.CommandName = cst_DG2CMDNM_DOWNLOAD4 Then
        'End If
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim LabdepID As Label = e.Item.FindControl("LabdepID")
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                Dim txtRtuReason As TextBox = e.Item.FindControl("txtRtuReason") '退回原因說明
                Dim Btn_RtuBACK1 As Button = e.Item.FindControl("Btn_RtuBACK1") '退回開放修改
                Dim Btn_REVERT1 As Button = e.Item.FindControl("Btn_REVERT1") '還原
                Dim BTN_VIEWFILE4 As Button = e.Item.FindControl("BTN_VIEWFILE4")
                Dim BTN_DOWNLOAD4 As Button = e.Item.FindControl("BTN_DOWNLOAD4")
                BTN_VIEWFILE4.Visible = False
                'BTN_DOWNLOAD4.Visible = False

                txtRtuReason.Text = TIMS.ClearSQM(drv("RTUREASON"))
                Btn_RtuBACK1.Enabled = If(txtRtuReason.Text = "", True, False)
                Btn_REVERT1.Enabled = If(txtRtuReason.Text <> "", True, False)
                txtRtuReason.Enabled = If(txtRtuReason.Text <> "", False, True)
                TIMS.Tooltip(txtRtuReason, cst_MaxLen500_TPMSG1, True)
                TIMS.Tooltip(Btn_RtuBACK1, If(Not Btn_RtuBACK1.Enabled, "退回開放修改", ""), True)
                TIMS.Tooltip(Btn_REVERT1, If(Not Btn_REVERT1.Enabled, "退回開放修改", ""), True)

                If Hid_APPLIEDRESULT.Value = "N" OrElse Hid_APPLIEDRESULT.Value = "Y" Then
                    txtRtuReason.Enabled = False
                    Btn_RtuBACK1.Enabled = False
                    Btn_REVERT1.Enabled = False
                    TIMS.Tooltip(txtRtuReason, "不可修改", True)
                    TIMS.Tooltip(Btn_RtuBACK1, "不可修改", True)
                    TIMS.Tooltip(Btn_REVERT1, "不可修改", True)
                End If

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    titleMsg = Convert.ToString(drv("FILENAME1"))
                    'LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                    'HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                    'LabFileName1.Text = cst_txt_免附文件
                ElseIf Convert.ToString(drv("WAIVED")) = cst_08_訓練班別計畫表_WAIVED_PI Then
                    titleMsg = cst_txt_版本批次送出
                ElseIf Convert.ToString(drv("WAIVED")) = cst_08_1_iCap課程原始申請資料_WAIVED_PI3 Then
                    BTN_VIEWFILE4.Visible = True '查看(查詢細項)
                    titleMsg = cst_txt_iCap課程原始申請資料
                ElseIf Convert.ToString(drv("WAIVED")) = cst_10_師資助教基本資料表_WAIVED_TT Then
                    BTN_VIEWFILE4.Visible = True '查看(查詢細項)
                    titleMsg = cst_txt_師資助教基本資料表
                ElseIf Convert.ToString(drv("WAIVED")) = cst_11_授課師資學經歷證書影本_WAIVED_TT2 Then
                    BTN_VIEWFILE4.Visible = True '查看(查詢細項)
                    titleMsg = cst_txt_授課師資學經歷證書影本
                ElseIf Convert.ToString(drv("WAIVED")) = cst_13_教學環境資料表_WAIVED_PI2 Then
                    BTN_VIEWFILE4.Visible = True '查看(查詢細項)
                    titleMsg = cst_txt_教學環境資料表
                ElseIf Convert.ToString(drv("WAIVED")) = cst_13_1混成課程教學環境資料表_WAIVED_RT2 Then
                    BTN_VIEWFILE4.Visible = True '查看(查詢細項)
                    titleMsg = cst_txt_混成課程教學環境資料表
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "BCID", drv("BCID"))
                TIMS.SetMyValue(sCmdArg, "BCFID", drv("BCFID"))
                TIMS.SetMyValue(sCmdArg, "KBID", drv("KBID"))
                TIMS.SetMyValue(sCmdArg, "KBSID", drv("KBSID"))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", drv("FILENAME1"))

                BTN_VIEWFILE4.CommandArgument = sCmdArg '查看(查詢細項)
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                Btn_RtuBACK1.CommandArgument = sCmdArg '退回開放修改
                Btn_REVERT1.CommandArgument = sCmdArg '還原

                Btn_RtuBACK1.Attributes("onclick") = "return confirm('您確定要「退回開放修改」這一筆資料?');"
                Btn_REVERT1.Attributes("onclick") = "return confirm('您確定要「還原」這一筆資料?');"
                '檢視不能修改
                'BTN_DELFILE4.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        End Select
    End Sub

    ''' <summary>回上一頁</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But_BACK1_Click(sender As Object, e As EventArgs) Handles But_BACK1.Click
        '清理隱藏的參數
        Call sSearch1()
    End Sub

    ''' <summary> '儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But_Sub_Click(sender As Object, e As EventArgs) Handles But_Sub.Click
        '資料有退回，審核結果，僅可為退回修正
        '檢核 ORG_BIDCASEFL RTUREASON 有值:true 無值:false 
        Dim sErrMsg1 As String = ""
        'ddlAPPLIEDRESULT:／審核通過:Y／審核不通過:N／退件修正:R
        Dim v_ddlAPPLIEDRESULT As String = TIMS.GetListValue(ddlAPPLIEDRESULT)
        Reasonforfail.Text = TIMS.ClearSQM(Reasonforfail.Text)

        If v_ddlAPPLIEDRESULT = "" Then sErrMsg1 &= "請選擇，審查狀態結果!" & vbCrLf
        '(6) 下方審核結果系統檢核， 當上面文件項目有任一個是有點選退回開放修改，【審核結果】就必須要整份退件修正。 有值:true 無值:false
        Dim flag_NG1 As Boolean = CHK_ORG_BIDCASEFL_RTUREASON()

        If flag_NG1 AndAlso v_ddlAPPLIEDRESULT <> "" AndAlso v_ddlAPPLIEDRESULT <> "R" Then
            sErrMsg1 &= "文件項目有任1個點選「退回開放修改」,【審查狀態】必須整份 申辦退件修正" & vbCrLf
        ElseIf Not flag_NG1 AndAlso v_ddlAPPLIEDRESULT = "N" AndAlso Reasonforfail.Text = "" Then
            sErrMsg1 &= "審查狀態「申辦不通過」,請輸入【不通過原因】" & vbCrLf
        ElseIf v_ddlAPPLIEDRESULT = "R" AndAlso Not flag_NG1 Then
            sErrMsg1 &= "【審查狀態】選擇 申辦退件修正,文件項目至少要有1個點選「退回開放修改」" & vbCrLf
        End If

        If sErrMsg1 <> "" Then
            Common.MessageBox(Me, sErrMsg1) : Return
        End If

        Call SAVEDATA1()
    End Sub

#Region "Public Shared"
    ''' <summary>取得申請項目資訊</summary>
    ''' <param name="oConn"></param>
    ''' <param name="vBCID"></param>
    ''' <param name="vORGKINDGW"></param>
    ''' <returns></returns>
    Public Shared Function GET_ORG_BIDCASEFL(ByRef oConn As SqlConnection, vBCID As String, vORGKINDGW As String) As DataTable
        Dim rParms As New Hashtable
        rParms.Add("BCID", Val(vBCID))
        rParms.Add("ORGKINDGW", vORGKINDGW)
        Dim rsSql As String = ""
        rsSql &= " SELECT a.BCFID,a.YEARS,a.APPSTAGE,a.RID,a.BCID,a.KBSID,a.PATTERN,a.MEMO1" & vbCrLf
        rsSql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        'APPSTAGE_N
        rsSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        rsSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上' WHEN 2 THEN '下' WHEN 3 THEN '政' WHEN 4 THEN '進' END APPSTAGE_S" & vbCrLf
        'rsSql &= ",a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        rsSql &= " ,kb.KBID,concat(kb.KBID,'.',kb.KBNAME) KBNAME" & vbCrLf
        rsSql &= " ,kb.KBID,concat(kb.ORGKINDGW,kb.KBID,kb.KBNAME) KBNAME2" & vbCrLf
        rsSql &= " ,oo.ORGNAME" & vbCrLf
        'rsSql &= " ,CONCAT(a.YEARS,'/',rr.PLANID,'/',a.RID,'/',ob.BCASENO,'/') PATH1" & vbCrLf
        rsSql &= " ,a.WAIVED,a.SRCFILENAME1,a.FILENAME1,a.FILENAME1 OKFLAG" & vbCrLf
        '退回原因說明／ '退回開放修改／ '還原
        rsSql &= " ,a.RTUREASON,a.RTURESACCT,a.RTURESDATE" & vbCrLf
        rsSql &= " FROM ORG_BIDCASEFL a" & vbCrLf
        rsSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=a.KBSID" & vbCrLf
        rsSql &= " JOIN ORG_BIDCASE ob on ob.BCID=a.BCID" & vbCrLf
        rsSql &= " JOIN AUTH_RELSHIP rr on rr.RID=a.RID" & vbCrLf
        rsSql &= " JOIN ORG_ORGINFO oo ON oo.ORGID=ob.ORGID" & vbCrLf
        rsSql &= " WHERE a.BCID=@BCID AND kb.ORGKINDGW=@ORGKINDGW" & vbCrLf
        rsSql &= " ORDER BY kb.KSORT,a.BCFID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        Return dt
    End Function

    ''' <summary>檔案打包下載 (機構班級／機構案件)</summary>
    ''' <param name="MyPage"></param>
    Public Shared Sub ResponseZIPFileALL_BI(ByRef MyPage As Page, ByRef oConn As SqlConnection, ByRef rPMS As Hashtable)
        Const cst_UtlSubName As String = "/*ResponseZIPFileALL_BI(ByRef MyPage As Page, ByRef oConn As SqlConnection, ByRef rPMS As Hashtable)*/"
        'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
        'Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        'Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(oConn, vRID, vBCID, vBCASENO)
        If drOB Is Nothing Then Return

        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        'Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(drOB("YEARS"))
        Dim vDISTID As String = Convert.ToString(drOB("DISTID"))
        Dim vDISTNAME3 As String = TIMS.GET_DISTNAME3(oConn, vDISTID)
        Dim vAPPSTAGE_S As String = Get_APPSTAGE_S(Convert.ToString(drOB("APPSTAGE")))
        Dim vORGNAME As String = TIMS.GET_ORGNAME(Convert.ToString(drOB("ORGID")), oConn)

        Dim oYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim oAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim oPLANID As String = Convert.ToString(drOB("PLANID"))
        Dim oRID As String = Convert.ToString(drOB("RID"))
        Dim oBCASENO As String = Convert.ToString(drOB("BCASENO"))

        Dim Template_ZipPath1 As String = TIMS.GET_Template_ZipPath1(vBCID)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath1)) Then
            Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath1))
        End If

        Dim dtFL As DataTable = GET_ORG_BIDCASEFL(oConn, vBCID, vORGKINDGW)
        For Each drFL As DataRow In dtFL.Rows
            Dim vKBID As String = Convert.ToString(drFL("KBID"))
            Dim vKBSID As String = Convert.ToString(drFL("KBSID"))
            Dim vKBNAME2 As String = Convert.ToString(drFL("KBNAME2"))

            Select Case String.Concat(vORGKINDGW, vKBID)
                Case TIMS.cst_W08_訓練班別計畫表, TIMS.cst_G08_訓練班別計畫表
                    Dim dtFLPI As DataTable = TIMS.GET_ORG_BIDCASEFL_PI(oConn, vBCID)
                    If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 Then
                        '不同檔案項目分別群組，取得最新的一筆下載
                        'Dim grpfmlst = newfmlst.GroupBy(Function(x) New With {Key x.apy_main_key, Key x.apy_src_key}).ToList()
                        For Each drFLPI As DataRow In dtFLPI.Rows
                            Dim vPSNO28 As String = "" ' Convert.ToString(drFLPI("PSNO28"))
                            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLPI("FILENAME1"))
                            Dim oUploadPath As String = "" ' TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_班級課程流水號_項目編號+項目名稱
                            Dim t_FILENAME_PI As String = "" ' String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf")
                            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
                            Try
                                vPSNO28 = Convert.ToString(drFLPI("PSNO28"))
                                oFILENAME1 = Convert.ToString(drFLPI("FILENAME1"))
                                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                '年度申請階段_班級課程流水號_項目編號+項目名稱
                                t_FILENAME_PI = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf"))
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
                                'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                                strErrmsg &= String.Concat("vPSNO28: ", vPSNO28, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME_PI: ", t_FILENAME_PI, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

                Case TIMS.cst_W08_1_iCap課程原始申請資料, TIMS.cst_G08_1_iCap課程原始申請資料
                    Dim dtFLPI As DataTable = TIMS.GET_ORG_BIDCASEFL_PI3(oConn, vBCID)
                    If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 Then
                        '不同檔案項目分別群組，取得最新的一筆下載
                        'Dim grpfmlst = newfmlst.GroupBy(Function(x) New With {Key x.apy_main_key, Key x.apy_src_key}).ToList()
                        For Each drFLPI As DataRow In dtFLPI.Rows
                            Dim vPSNO28 As String = "" ' Convert.ToString(drFLPI("PSNO28"))
                            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLPI("FILENAME1"))
                            Dim oUploadPath As String = "" ' TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_班級課程流水號_項目編號+項目名稱
                            Dim t_FILENAME_PI As String = "" ' String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf")
                            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
                            Try
                                vPSNO28 = Convert.ToString(drFLPI("PSNO28"))
                                oFILENAME1 = Convert.ToString(drFLPI("FILENAME1"))
                                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                '年度申請階段_班級課程流水號_項目編號+項目名稱
                                t_FILENAME_PI = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf"))
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
                                'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                                strErrmsg &= String.Concat("vPSNO28: ", vPSNO28, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME_PI: ", t_FILENAME_PI, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

                Case TIMS.cst_W10_師資助教基本資料表, TIMS.cst_G10_師資助教基本資料表
                    Dim dtFLTT As DataTable = TIMS.GET_ORG_BIDCASEFL_TT(oConn, vBCID)
                    If dtFLTT IsNot Nothing AndAlso dtFLTT.Rows.Count > 0 Then
                        For Each drFLTT As DataRow In dtFLTT.Rows
                            Dim vTEACHCNAME As String = "" 'Convert.ToString(drFLTT("TEACHCNAME"))
                            Dim vTEACHERID As String = "" 'Convert.ToString(drFLTT("TEACHERID"))
                            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                            Dim t_FILENAME_TT As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT))
                            Try
                                vTEACHCNAME = Convert.ToString(drFLTT("TEACHCNAME"))
                                vTEACHERID = Convert.ToString(drFLTT("TEACHERID"))
                                oFILENAME1 = Convert.ToString(drFLTT("FILENAME1"))
                                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                                t_FILENAME_TT = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf"))
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT))
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                                strErrmsg &= String.Concat("vTEACHCNAME: ", vTEACHCNAME, vbCrLf, "vTEACHERID: ", vTEACHERID, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME_TT: ", t_FILENAME_TT, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

                Case TIMS.cst_G11_授課師資學經歷證書影本, TIMS.cst_W11_授課師資學經歷證書影本
                    Dim dtFLTT2 As DataTable = TIMS.GET_ORG_BIDCASEFL_TT2(oConn, vBCID)
                    If dtFLTT2 IsNot Nothing AndAlso dtFLTT2.Rows.Count > 0 Then
                        For Each drFLTT2 As DataRow In dtFLTT2.Rows
                            Dim vTEACHCNAME As String = "" ' Convert.ToString(drFLTT2("TEACHCNAME"))
                            Dim vTEACHERID As String = "" 'Convert.ToString(drFLTT2("TEACHERID"))
                            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT2("FILENAME1"))
                            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                            Dim t_FILENAME_TT2 As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf")
                            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT2))
                            'Dim t_FilePath1 As String = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
                            Try
                                vTEACHCNAME = Convert.ToString(drFLTT2("TEACHCNAME"))
                                vTEACHERID = Convert.ToString(drFLTT2("TEACHERID"))
                                oFILENAME1 = Convert.ToString(drFLTT2("FILENAME1"))
                                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                '年度申請階段_講師名稱_講師代碼_項目編號+項目名稱
                                t_FILENAME_TT2 = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vTEACHERID, "_", vTEACHCNAME, ".pdf"))
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_TT2))
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                                strErrmsg &= String.Concat("vTEACHCNAME: ", vTEACHCNAME, vbCrLf, "vTEACHERID: ", vTEACHERID, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME_TT2: ", t_FILENAME_TT2, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

                Case TIMS.cst_W13_教學環境資料表, TIMS.cst_G13_教學環境資料表
                    Dim dtFLEV As DataTable = TIMS.GET_ORG_BIDCASEFL_EV(oConn, vBCID)
                    If dtFLEV IsNot Nothing AndAlso dtFLEV.Rows.Count > 0 Then
                        '不同檔案項目分別群組，取得最新的一筆下載
                        'Dim grpfmlst = newfmlst.GroupBy(Function(x) New With {Key x.apy_main_key, Key x.apy_src_key}).ToList()
                        For Each drFLEV As DataRow In dtFLEV.Rows
                            Dim vPSNO28 As String = "" 'Convert.ToString(drFLPI2("PSNO28"))
                            Dim oFILENAME1 As String = "" ' Convert.ToString(drFLPI2("FILENAME1"))
                            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_班級課程流水號_項目編號+項目名稱
                            Dim t_FILENAME_PI2 As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf")
                            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI2))
                            'Dim t_FilePath1 As String = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
                            Try
                                vPSNO28 = Convert.ToString(drFLEV("PSNO28"))
                                oFILENAME1 = Convert.ToString(drFLEV("FILENAME1"))
                                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                '年度申請階段_班級課程流水號_項目編號+項目名稱
                                t_FILENAME_PI2 = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf"))
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI2))
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                                strErrmsg &= String.Concat("vPSNO28: ", vPSNO28, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME_PI2: ", t_FILENAME_PI2, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

                Case TIMS.cst_W13_1_混成課程教學環境資料表, TIMS.cst_G13_1_混成課程教學環境資料表
                    Dim dtFLRT As DataTable = TIMS.GET_ORG_BIDCASEFL_RT(oConn, vBCID)
                    If dtFLRT IsNot Nothing AndAlso dtFLRT.Rows.Count > 0 Then
                        '不同檔案項目分別群組，取得最新的一筆下載
                        'Dim grpfmlst = newfmlst.GroupBy(Function(x) New With {Key x.apy_main_key, Key x.apy_src_key}).ToList()
                        For Each drFLRT As DataRow In dtFLRT.Rows
                            Dim vPSNO28 As String = "" 'Convert.ToString(drFLPI2("PSNO28"))
                            Dim oFILENAME1 As String = "" ' Convert.ToString(drFLPI2("FILENAME1"))
                            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            '年度申請階段_班級課程流水號_項目編號+項目名稱
                            Dim t_FILENAME_PI2 As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf")
                            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI2))
                            'Dim t_FilePath1 As String = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
                            Try
                                vPSNO28 = Convert.ToString(drFLRT("PSNO28"))
                                oFILENAME1 = Convert.ToString(drFLRT("FILENAME1"))
                                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                '年度申請階段_班級課程流水號_項目編號+項目名稱
                                t_FILENAME_PI2 = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf"))
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI2))
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                                strErrmsg &= String.Concat("vPSNO28: ", vPSNO28, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME_PI2: ", t_FILENAME_PI2, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

                Case Else
                    Dim oFILENAME1 As String = "" 'Convert.ToString(drFL("FILENAME1"))
                    Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                    Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度申請階段_單位名稱_項目編號+項目名稱
                    Dim t_FILENAME As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf")
                    Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))
                    'Dim t_FilePath1 As String = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
                    Try
                        oFILENAME1 = Convert.ToString(drFL("FILENAME1"))
                        oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                        s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        '年度申請階段_單位名稱_項目編號+項目名稱
                        t_FILENAME = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf"))
                        t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))
                        If oFILENAME1 <> "" AndAlso IO.File.Exists(s_FilePath1) Then
                            Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                            File.WriteAllBytes(t_FilePath1, dbyte)
                        End If
                    Catch ex As Exception
                        Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                        strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                        strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                        strErrmsg &= String.Concat("t_FILENAME: ", t_FILENAME, vbCrLf)
                        strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                        strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                        Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                    End Try
            End Select
        Next

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim filenames As String() = Directory.GetFiles(MyPage.Server.MapPath(String.Concat(Template_ZipPath1, "/")))
        Dim zipFileName As String = TIMS.GetValidFileName(String.Concat("p", vYEARS_ROC, vAPPSTAGE_S, "_", vDISTNAME3, "_", vORGNAME, "_", vBCID, "_", strNOW, ".zip"))
        Dim full_zipFileName As String = String.Concat(Template_ZipPath1, "/", zipFileName)
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
                    Dim strErrmsg As String = "/*ResponseZIPFileALL_BI*/" & vbCrLf
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
            TIMS.SAVE_ADP_ZIPFILE(oConn, "-tc11002bi", File)
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
            TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
        End With
    End Sub

    ''' <summary>'【訓練班別計畫表】打包下載</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="oConn"></param>
    ''' <param name="rPMS"></param>
    Public Shared Sub ResponseZIPFilePI(MyPage As Page, ByRef oConn As SqlConnection, ByRef rPMS As Hashtable, ByRef dtOBWP As DataTable)
        Const cst_UtlSubName As String = "/*ResponseZIPFilePI*/"
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS, "DISTID")
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim vAPPSTAGE_S As String = Get_APPSTAGE_S(vAPPSTAGE)
        Dim vDISTNAME As String = TIMS.GET_DISTNAME(oConn, vDISTID)

        'Dim rPMS2 As New Hashtable
        'rPMS2.Add("DISTID", vDISTID)
        'rPMS2.Add("YEARS", vYEARS)
        'rPMS2.Add("APPSTAGE", vAPPSTAGE)
        ''打包【訓練班別計畫表】(單位有上傳就算)
        'Dim dtOBWP As DataTable = GET_ORG_BIDCASE_WAIVED_PI(oConn, rPMS2)
        If dtOBWP Is Nothing OrElse dtOBWP.Rows.Count = 0 Then Return

        Dim Template_ZipPath1 As String = TIMS.GET_Template_ZipPath1(vDISTID)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath1)) Then
            Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath1))
        End If

        For Each drOP As DataRow In dtOBWP.Rows
            Dim vBCID As String = Convert.ToString(drOP("BCID"))
            Dim vORGNAME As String = TIMS.GET_ORGNAME(Convert.ToString(drOP("ORGID")), oConn)
            Dim oYEARS As String = Convert.ToString(drOP("YEARS"))
            Dim oAPPSTAGE As String = Convert.ToString(drOP("APPSTAGE"))
            Dim oPLANID As String = Convert.ToString(drOP("PLANID"))
            Dim oRID As String = Convert.ToString(drOP("RID"))
            Dim oBCASENO As String = Convert.ToString(drOP("BCASENO"))
            Dim vKBSID As String = Convert.ToString(drOP("KBSID"))
            Dim vORGKINDGW As String = Convert.ToString(drOP("ORGKINDGW"))
            Dim vKBID As String = Convert.ToString(drOP("KBID"))
            Dim vKBNAME2 As String = String.Concat(vORGKINDGW, vKBID, drOP("KBNAME"))

            Dim dtFLPI As DataTable = TIMS.GET_ORG_BIDCASEFL_PI(oConn, vBCID)
            If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 Then
                '不同檔案項目分別群組，取得最新的一筆下載
                'Dim grpfmlst = newfmlst.GroupBy(Function(x) New With {Key x.apy_main_key, Key x.apy_src_key}).ToList()
                For Each drFLPI As DataRow In dtFLPI.Rows
                    Dim vPSNO28 As String = "" 'Convert.ToString(drFLPI("PSNO28"))
                    Dim oFILENAME1 As String = "" ' Convert.ToString(drFLPI("FILENAME1"))
                    Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                    Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度申請階段_班級課程流水號_項目編號+項目名稱
                    Dim t_FILENAME_PI As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf")
                    Dim t_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
                    Try
                        vPSNO28 = Convert.ToString(drFLPI("PSNO28"))
                        oFILENAME1 = Convert.ToString(drFLPI("FILENAME1"))
                        oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                        s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        '年度申請階段_班級課程流水號_項目編號+項目名稱
                        t_FILENAME_PI = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, "_", vPSNO28, ".pdf"))
                        t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
                        Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
                        If IO.File.Exists(s_FilePath1) Then
                            Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                            File.WriteAllBytes(t_FilePath1, dbyte)
                        End If
                    Catch ex As Exception
                        Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
                        strErrmsg &= String.Concat("vPSNO28: ", vPSNO28, vbCrLf)
                        strErrmsg &= String.Concat("t_FILENAME_PI: ", t_FILENAME_PI, vbCrLf)
                        strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                        strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                        strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                        Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                    End Try
                Next
            End If
        Next

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim zipFileName As String = String.Concat("p", vYEARS_ROC, vAPPSTAGE_S, "_", vDISTNAME, "_", strNOW, ".zip")
        Dim filenames As String() = Directory.GetFiles(MyPage.Server.MapPath(String.Concat(Template_ZipPath1, "/")))
        Dim full_zipFileName As String = String.Concat(Template_ZipPath1, "/", zipFileName)
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
                    Dim strErrmsg As String = "/*ResponseZIPFilePI*/" & vbCrLf
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
            TIMS.SAVE_ADP_ZIPFILE(oConn, "-tc11002pi", File)
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
            TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
        End With
    End Sub

#End Region

#Region "Private1"
    ''' <summary>檢核 ORG_BIDCASEFL RTUREASON 有值:true 無值:false </summary>
    ''' <returns></returns>
    Private Function CHK_ORG_BIDCASEFL_RTUREASON() As Boolean
        'Dim rst As Boolean = False
        Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim dtFL As DataTable = GET_ORG_BIDCASEFL(objconn, vBCID, vORGKINDGW)
        For Each drFL As DataRow In dtFL.Rows
            If Convert.ToString(drFL("RTUREASON")) <> "" Then Return True
        Next
        Return False
    End Function

    ''' <summary> '儲存 </summary>
    Sub SAVEDATA1()
        '申辦狀態：暫存/ 已送件
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!!")
            Return
        End If

        Dim vRID As String = Hid_RID.Value
        Dim vBCID As String = Hid_BCID.Value
        Dim vBCASENO As String = Hid_BCASENO.Value

        Dim v_ddlAPPLIEDRESULT As String = TIMS.GetListValue(ddlAPPLIEDRESULT)
        Reasonforfail.Text = TIMS.ClearSQM(Reasonforfail.Text)
        Dim v_Reasonforfail As String = TIMS.Get_Substr1(Reasonforfail.Text, cst_MaxLen500_i)
        'Reasonforfail.Text = TIMS.Get_Substr1(Reasonforfail.Text, cst_MaxLen500_i)

        'Dim s_HIS As String = String.Concat(TIMS.cdate3t(Now), "-", TIMS.GetListText(ddlAPPLIEDRESULT))
        Dim s_HISREVIEW As String = Convert.ToString(drOB("HISREVIEW"))
        s_HISREVIEW &= String.Concat(If(s_HISREVIEW <> "", "，", ""), String.Concat(TIMS.Cdate3t(Now), "-", TIMS.GetListText(ddlAPPLIEDRESULT)))

        Dim uParms As New Hashtable
        uParms.Add("HISREVIEW", s_HISREVIEW)
        uParms.Add("APPLIEDRESULT", v_ddlAPPLIEDRESULT)
        uParms.Add("REASONFORFAIL", If(v_Reasonforfail <> "", v_Reasonforfail, Convert.DBNull))
        uParms.Add("RESULTACCT", sm.UserInfo.UserID)
        'uParms.Add("RESULTDATE", RESULTDATE)
        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
        'uParms.Add("MODIFYDATE", MODIFYDATE)
        uParms.Add("RID", vRID)
        uParms.Add("BCID", vBCID)
        uParms.Add("BCASENO", vBCASENO)
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_BIDCASE" & vbCrLf
        usSql &= " SET HISREVIEW=@HISREVIEW,APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        usSql &= " ,REASONFORFAIL=@REASONFORFAIL,RESULTACCT=@RESULTACCT,RESULTDATE=GETDATE()" & vbCrLf
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE RID=@RID AND BCID=@BCID AND BCASENO=@BCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        If v_ddlAPPLIEDRESULT = "R" Then
            '退件修正 BISTATUS='R'
            Dim uParms2 As New Hashtable
            uParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
            uParms2.Add("RID", vRID)
            uParms2.Add("BCID", vBCID)
            uParms2.Add("BCASENO", vBCASENO)
            Dim usSql2 As String = ""
            usSql2 &= " UPDATE ORG_BIDCASE" & vbCrLf
            usSql2 &= " SET BISTATUS='R',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql2 &= " WHERE RID=@RID AND BCID=@BCID AND BCASENO=@BCASENO" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        End If

        Call sSearch1()
    End Sub

    ''' <summary>線上送件審核，還原按鈕 UPDATE ORG_BIDCASE / BISTATUS='B',APPLIEDRESULT=NULL </summary>
    Sub UPDATE_REVERT2(s_HISREVIEW As String, vRID As String, vBCID As String, vBCASENO As String)
        vRID = TIMS.ClearSQM(vRID)
        vBCID = TIMS.ClearSQM(vBCID)
        vBCASENO = TIMS.ClearSQM(vBCASENO)
        If vRID = "" OrElse vBCID = "" OrElse vBCASENO = "" Then Return
        '退件修正 BISTATUS
        Dim uParms2 As New Hashtable
        If s_HISREVIEW <> "" Then uParms2.Add("HISREVIEW", s_HISREVIEW)
        uParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
        uParms2.Add("RID", vRID)
        uParms2.Add("BCID", vBCID)
        uParms2.Add("BCASENO", vBCASENO)
        Dim usSql2 As String = ""
        usSql2 &= " UPDATE ORG_BIDCASE" & vbCrLf
        usSql2 &= " SET BISTATUS='B',APPLIEDRESULT=NULL" & vbCrLf
        If s_HISREVIEW <> "" Then usSql2 &= " ,HISREVIEW=@HISREVIEW" & vbCrLf
        usSql2 &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql2 &= " WHERE RID=@RID AND BCID=@BCID AND BCASENO=@BCASENO" & vbCrLf
        Dim iRst As Integer = DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        If iRst > 0 Then Common.MessageBox(Me, "審核狀態 已還原，尚未確認狀態!")
    End Sub


    ''' <summary>檢視目前上傳檔案</summary>
    ''' <param name="rPMS"></param>
    Sub SHOW_BIDCASEFL_DG2(ByRef rPMS As Hashtable)
        labmsg1.Text = ""
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        'objconn 因為有檔案輸出關閉的問題 所以要檢查
        If Not TIMS.OpenDbConn(objconn) OrElse Not fg_CANSAVE Then Return

        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Hid_RID.Value = vRID
        Hid_BCID.Value = vBCID 'TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = vBCASENO 'TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If

        Dim dtFL As DataTable = GET_ORG_BIDCASEFL(objconn, vBCID, vORGKINDGW)

        'If dt IsNot Nothing AndAlso dt.Rows.Count > 0 Then
        '    DataGrid3Table.Visible = True
        'End If
        labmsg1.Text = If(dtFL Is Nothing OrElse dtFL.Rows.Count = 0, "(查無文件項目)", "")

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        'Dim vRID As String = Convert.ToString(drOB("RID"))
        'Dim vBCASENO As String = Convert.ToString(drOB("BCASENO"))
        'Dim vKBSID As String = Convert.ToString(drKB("KBSID"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, "")
        'Call TIMS.Check_dtBIDCASEFL(Me, dtFL, download_Path)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()
    End Sub

    ''' <summary>'切換項目(預設)KEY_BIDCASE </summary>
    ''' <param name="rPMS4"></param>
    Private Sub SHOW_BIDCASE_KBSID(rPMS4 As Hashtable)
        tr_LabSwitchTo.Visible = False
        tr_DataGrid10.Visible = False
        tr_DataGrid11.Visible = False
        tr_DataGrid13.Visible = False
        tr_DataGrid13B.Visible = False
        tr_DataGrid14.Visible = False

        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS4, "ORGKINDGW") ', Hid_ORGKINDGW.Value)
        Dim vRID As String = TIMS.GetMyValue2(rPMS4, "RID") ', Hid_RID.Value)
        Dim vBCID As String = TIMS.GetMyValue2(rPMS4, "BCID") ', Hid_BCID.Value)
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS4, "BCASENO") ', Hid_BCASENO.Value)
        Dim vBCFID As String = TIMS.GetMyValue2(rPMS4, "BCFID") ', vBCFID)
        Dim vKBID As String = TIMS.GetMyValue2(rPMS4, "KBID") ', vKBID)
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS4, "KBSID") ', vKBSID)
        Hid_ORGKINDGW.Value = vORGKINDGW
        Hid_RID.Value = vRID
        Hid_BCID.Value = vBCID
        Hid_BCASENO.Value = vBCASENO
        Hid_BCFID.Value = vBCFID
        Hid_KBID.Value = vKBID
        Hid_KBSID.Value = vKBSID

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, vKBSID, vORGKINDGW)
        If drOB Is Nothing Then Return
        If drKB Is Nothing Then Return

        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_G08_1_iCap課程原始申請資料, TIMS.cst_W08_1_iCap課程原始申請資料
                tr_DataGrid14.Visible = True
                'iCap課程原始申請資料
                Call SHOW_DATAGRID_14(drOB, drKB)
            Case TIMS.cst_G10_師資助教基本資料表, TIMS.cst_W10_師資助教基本資料表
                tr_DataGrid10.Visible = True
                '師資／助教基本資料表
                Call SHOW_DATAGRID_10(drOB, drKB)
            Case TIMS.cst_G11_授課師資學經歷證書影本, TIMS.cst_W11_授課師資學經歷證書影本
                tr_DataGrid11.Visible = True
                '各授課師資學／經歷證書影本
                Call SHOW_DATAGRID_11(drOB, drKB)
            Case TIMS.cst_G13_教學環境資料表, TIMS.cst_W13_教學環境資料表
                tr_DataGrid13.Visible = True
                '教學環境資料表
                Call SHOW_DATAGRID_13(drOB, drKB)
            Case TIMS.cst_G13_1_混成課程教學環境資料表, TIMS.cst_W13_1_混成課程教學環境資料表
                tr_DataGrid13B.Visible = True
                '混成課程教學環境資料表
                Call SHOW_DATAGRID_13B(drOB, drKB)
        End Select
    End Sub

    Private Sub SHOW_DATAGRID_14(drOB As DataRow, drKB As DataRow)
        If drOB Is Nothing OrElse drKB Is Nothing Then Return
        Dim vBCID As String = Convert.ToString(drOB("BCID"))
        Dim vKBID As String = Convert.ToString(drKB("KBID"))
        Dim vKBNAME As String = Convert.ToString(drKB("KBNAME"))
        tr_LabSwitchTo.Visible = True
        LabSwitchTo.Text = String.Concat(vKBID, ".", vKBNAME)
        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim sParms1 As New Hashtable
        sParms1.Add("BCID", vBCID) 'ORG_BIDCASEPI
        sParms1.Add("RID", vRID)
        sParms1.Add("AppStage", vAPPSTAGE)
        Dim sSql As String = ""
        sSql &= " WITH WFPI3 AS ( SELECT b.BCFP3ID,b.BCFID,a.BCID,a.BCPID" & vbCrLf
        sSql &= " ,a.PLANID,a.COMIDNO,a.SEQNO,pp.PSNO28" & vbCrLf
        sSql &= " ,b.SRCFILENAME1,b.FILENAME1,b.WAIVED,kb.KBID,kb.KBSID,kb.ORGKINDGW" & vbCrLf
        sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL_PI3 b on b.BCPID=a.BCPID" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID" & vbCrLf
        sSql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID and pp.COMIDNO=a.COMIDNO and pp.SEQNO=a.SEQNO" & vbCrLf
        sSql &= " WHERE a.BCID=@BCID )" & vbCrLf

        sSql &= " SELECT ip.YEARS,a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID" & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sSql &= " ,b.OrgName ,a.RID" & vbCrLf
        sSql &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf
        sSql &= " ,p3.WAIVED,p3.SRCFILENAME1,p3.FILENAME1,p3.FILENAME1 OKFLAG" & vbCrLf
        sSql &= " ,p3.BCFP3ID,p3.BCFID,p3.BCID,p3.BCPID,p3.KBID,p3.KBSID,p3.ORGKINDGW" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID = b.RID" & vbCrLf
        sSql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = a.PlanID" & vbCrLf
        sSql &= " JOIN WFPI3 p3 ON p3.PLANID=a.PLANID AND p3.COMIDNO=a.COMIDNO AND p3.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        sSql &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        '使用登入者業務權限
        sSql &= " AND a.RID=@RID AND a.AppStage=@AppStage" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPlanID", sm.UserInfo.TPlanID)
            sParms1.Add("Years", sm.UserInfo.Years)
            sSql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            sSql &= " AND ip.Years=@Years" & vbCrLf
        Else
            sParms1.Add("PlanID", sm.UserInfo.PlanID)
            sSql &= " AND ip.PlanID=@PlanID" & vbCrLf
        End If
        sSql &= " ORDER BY a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        'Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        'Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vBCASENO As String = Convert.ToString(drOB("BCASENO"))
        Dim vKBSID As String = Convert.ToString(drKB("KBSID"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        labmsg2.Text = ""
        DataGrid14.Visible = True
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            DataGrid14.Visible = False
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If
        DataGrid14.DataSource = dt2
        DataGrid14.DataBind()
    End Sub

    Private Sub SHOW_DATAGRID_13(drOB As DataRow, drKB As DataRow)
        If drOB Is Nothing Then Return
        If drKB Is Nothing Then Return
        Dim vBCID As String = Convert.ToString(drOB("BCID"))
        Dim vKBID As String = Convert.ToString(drKB("KBID"))
        Dim vKBNAME As String = Convert.ToString(drKB("KBNAME"))
        tr_LabSwitchTo.Visible = True
        LabSwitchTo.Text = String.Concat(vKBID, ".", vKBNAME)

        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))

        Dim sParms1 As New Hashtable
        sParms1.Add("BCID", vBCID) 'ORG_BIDCASEPI
        sParms1.Add("RID", vRID)
        sParms1.Add("AppStage", vAPPSTAGE)

        Dim sSql As String = ""
        sSql &= " WITH WFPI2 AS (" & vbCrLf
        sSql &= "  SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= "  ,b.SRCFILENAME1,b.FILENAME1,b.WAIVED" & vbCrLf
        sSql &= "  ,ob.YEARS,b.BCFEID,a.BCPID,f.BCFID,kb.ORGKINDGW,kb.KBID,f.KBSID" & vbCrLf
        sSql &= "  FROM ORG_BIDCASEPI a" & vbCrLf
        sSql &= "  JOIN ORG_BIDCASEFL_EV b on b.BCPID=a.BCPID" & vbCrLf
        sSql &= "  JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql &= "  JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID" & vbCrLf
        sSql &= "  JOIN ORG_BIDCASE ob on ob.BCID=f.BCID" & vbCrLf
        sSql &= "  WHERE a.BCID=@BCID)" & vbCrLf

        sSql &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID " & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sSql &= " ,b.OrgName ,a.RID" & vbCrLf
        sSql &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf 'pp.MSD
        sSql &= " ,p2.WAIVED,p2.SRCFILENAME1,p2.FILENAME1,p2.FILENAME1 OKFLAG" & vbCrLf
        sSql &= " ,p2.YEARS,p2.BCFEID,p2.BCPID,p2.BCFID,p2.ORGKINDGW,p2.KBID,p2.KBSID" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
        sSql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID=a.PlanID" & vbCrLf
        sSql &= " JOIN WFPI2 p2 ON p2.PLANID=a.PLANID AND p2.COMIDNO=a.COMIDNO AND p2.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        sSql &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        '使用登入者業務權限
        sSql &= " AND a.RID=@RID AND a.AppStage=@AppStage" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPLANID", sm.UserInfo.TPlanID)
            sParms1.Add("YEARS", sm.UserInfo.Years)
            sSql &= " AND ip.TPLANID=@TPLANID" & vbCrLf
            sSql &= " AND ip.YEARS=@YEARS" & vbCrLf
        Else
            sParms1.Add("PLANID", sm.UserInfo.PlanID)
            sSql &= " AND ip.PLANID=@PLANID" & vbCrLf
        End If
        sSql &= " ORDER BY a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        'Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        'Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vBCASENO As String = Convert.ToString(drOB("BCASENO"))
        Dim vKBSID As String = Convert.ToString(drKB("KBSID"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        labmsg2.Text = ""
        DataGrid13.Visible = True
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            DataGrid13.Visible = False
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        With DataGrid13
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    ''' <summary>混成課程教學環境資料表／遠距</summary>
    Private Sub SHOW_DATAGRID_13B(drOB As DataRow, drKB As DataRow)
        'iDG13B_ROWS = 0
        Dim oDG1 As DataGrid = DataGrid13B
        If drOB Is Nothing OrElse drKB Is Nothing Then Return
        Dim vBCID As String = Convert.ToString(drOB("BCID"))
        Dim vKBID As String = Convert.ToString(drKB("KBID"))
        Dim vKBNAME As String = Convert.ToString(drKB("KBNAME"))
        tr_LabSwitchTo.Visible = True
        LabSwitchTo.Text = String.Concat(vKBID, ".", vKBNAME)

        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))

        Dim sParms1 As New Hashtable From {{"BCID", vBCID}, {"RID", vRID}, {"AppStage", vAPPSTAGE}}
        Dim sSql As String = ""
        sSql &= " WITH WFRT2 AS ( SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= "  ,b.SRCFILENAME1,b.FILENAME1,b.WAIVED" & vbCrLf
        sSql &= "  ,ob.YEARS,b.BCRTID,a.BCPID,f.BCFID,kb.ORGKINDGW,kb.KBID,f.KBSID" & vbCrLf
        sSql &= "  FROM ORG_BIDCASEPI a" & vbCrLf
        sSql &= "  JOIN ORG_BIDCASEFL_RT b on b.BCPID=a.BCPID" & vbCrLf
        sSql &= "  JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql &= "  JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID" & vbCrLf
        sSql &= "  JOIN ORG_BIDCASE ob on ob.BCID=f.BCID" & vbCrLf
        sSql &= "  WHERE a.BCID=@BCID )" & vbCrLf

        sSql &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID " & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sSql &= " ,b.OrgName ,a.RID" & vbCrLf
        sSql &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf 'pp.MSD
        sSql &= " ,p2.WAIVED,p2.SRCFILENAME1,p2.FILENAME1,p2.FILENAME1 OKFLAG" & vbCrLf
        sSql &= " ,p2.YEARS,p2.BCRTID,p2.BCPID,p2.BCFID,p2.ORGKINDGW,p2.KBID,p2.KBSID" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
        sSql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID=a.PlanID" & vbCrLf
        sSql &= " JOIN WFRT2 p2 ON p2.PLANID=a.PLANID AND p2.COMIDNO=a.COMIDNO AND p2.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        sSql &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        '使用登入者業務權限
        sSql &= " AND a.RID=@RID AND a.AppStage=@AppStage" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPLANID", sm.UserInfo.TPlanID)
            sParms1.Add("YEARS", sm.UserInfo.Years)
            sSql &= " AND ip.TPLANID=@TPLANID" & vbCrLf
            sSql &= " AND ip.YEARS=@YEARS" & vbCrLf
        Else
            sParms1.Add("PLANID", sm.UserInfo.PlanID)
            sSql &= " AND ip.PLANID=@PLANID" & vbCrLf
        End If
        sSql &= " ORDER BY a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        'Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        'Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vBCASENO As String = Convert.ToString(drOB("BCASENO"))
        Dim vKBSID As String = Convert.ToString(drKB("KBSID"))

        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        labmsg2.Text = TIMS.cst_NODATAMsg1
        oDG1.Visible = False

        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then Return

        labmsg2.Text = ""
        oDG1.Visible = True

        oDG1.DataSource = dt2
        oDG1.DataBind()
    End Sub

    Private Sub SHOW_DATAGRID_11(drOB As DataRow, drKB As DataRow)
        If drOB Is Nothing Then Return
        If drKB Is Nothing Then Return
        Dim vKBID As String = Convert.ToString(drKB("KBID"))
        Dim vKBNAME As String = Convert.ToString(drKB("KBNAME"))
        tr_LabSwitchTo.Visible = True
        LabSwitchTo.Text = String.Concat(vKBID, ".", vKBNAME)

        'DataGrid11_ItemDataBound
        Dim rParms2 As New Hashtable
        rParms2.Add("BCID", Convert.ToString(drOB("BCID")))
        rParms2.Add("RID", Convert.ToString(drOB("RID")))
        rParms2.Add("AppStage", Convert.ToString(drOB("AppStage")))
        Dim sSql2 As String = ""
        sSql2 &= " WITH WT1 AS (SELECT a.TECHID,a.BCFID,a.BCFT2ID,a.PATTERN,a.MEMO1,a.WAIVED" & vbCrLf
        sSql2 &= " ,a.SRCFILENAME1,a.FILENAME1" & vbCrLf
        sSql2 &= " ,ob.YEARS,rr.PLANID,rr.RID,ob.BCASENO,kb.ORGKINDGW,kb.KBID,bf.KBSID" & vbCrLf
        sSql2 &= " FROM VIEW_RIDNAME rr" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL bf ON bf.RID=rr.RID" & vbCrLf
        sSql2 &= " JOIN KEY_BIDCASE kb on kb.KBSID=bf.KBSID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASE ob ON ob.BCID=bf.BCID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL_TT2 a ON a.BCFID=bf.BCFID" & vbCrLf
        sSql2 &= " WHERE ob.BCID=@BCID AND rr.RID=@RID)" & vbCrLf

        sSql2 &= " ,WT2 AS (SELECT DISTINCT P2.TECHID" & vbCrLf
        sSql2 &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
        sSql2 &= " JOIN dbo.ORG_BIDCASEPI bp ON bp.PlanID=P1.PlanID AND bp.ComIDNO=P1.ComIDNO AND bp.SeqNo=P1.SeqNo" & vbCrLf
        sSql2 &= " JOIN dbo.V_PLAN_TEACHER1 P2 ON P1.PlanID=P2.PlanID AND P1.ComIDNO=P2.ComIDNO AND P1.SeqNo=P2.SeqNo" & vbCrLf
        sSql2 &= " WHERE P1.TransFlag='N' AND P1.IsApprPaper='Y' AND P1.AppliedResult IS NULL AND P1.RESULTBUTTON IS NULL" & vbCrLf
        sSql2 &= " AND bp.BCID=@BCID AND P1.RID=@RID AND P1.AppStage=@AppStage)" & vbCrLf
        'sSql2 &= " WHERE p1.RID=@RID)" & vbCrLf

        sSql2 &= " SELECT a.TechID,a.RID,a.TEACHCNAME,a.TEACHENAME,a.TEACHERID" & vbCrLf
        sSql2 &= " ,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql2 &= " ,a.KINDENGAGE,case a.KINDENGAGE when '1' then '內聘(專任)' else '外聘(兼任)' end KINDENGAGE_N" & vbCrLf
        sSql2 &= " ,a.WORKSTATUS,case a.WORKSTATUS when '1' then '是' else '否' end WORKSTATUS_N" & vbCrLf
        sSql2 &= " ,(SELECT x.KINDNAME FROM ID_KINDOFTEACHER x WHERE x.KINDID=a.KINDID) KINDNAME" & vbCrLf
        sSql2 &= " ,bt.BCFID,bt.BCFT2ID,bt.PATTERN,bt.MEMO1,bt.WAIVED" & vbCrLf
        'sSql2 &= " ,case when bt.TECHID IS NOT NULL THEN CONCAT(bt.YEARS,'/',bt.PLANID,'/',bt.RID,'/',bt.BCASENO,'/',bt.KBSID,'/') END PATH1" & vbCrLf
        sSql2 &= " ,bt.SRCFILENAME1,bt.FILENAME1,bt.FILENAME1 OKFLAG" & vbCrLf
        sSql2 &= " ,bt.YEARS,bt.ORGKINDGW,bt.KBID,bt.KBSID" & vbCrLf
        sSql2 &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sSql2 &= " JOIN WT2 t2 ON t2.TECHID=a.TECHID" & vbCrLf
        sSql2 &= " JOIN AUTH_RELSHIP b ON b.RID=a.RID" & vbCrLf
        sSql2 &= " JOIN WT1 bt on bt.TECHID=a.TECHID AND bt.RID=a.RID" & vbCrLf
        sSql2 &= " WHERE a.RID=@RID" & vbCrLf
        sSql2 &= " ORDER BY a.TEACHERID" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, rParms2)

        labmsg2.Text = ""
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG11_ROWS = dt2.Rows.Count

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vBCASENO As String = Convert.ToString(drOB("BCASENO"))
        Dim vKBSID As String = Convert.ToString(drKB("KBSID"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)
        With DataGrid11
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    Private Sub SHOW_DATAGRID_10(drOB As DataRow, drKB As DataRow)
        If drOB Is Nothing Then Return
        If drKB Is Nothing Then Return
        Dim vKBID As String = Convert.ToString(drKB("KBID"))
        Dim vKBNAME As String = Convert.ToString(drKB("KBNAME"))
        tr_LabSwitchTo.Visible = True
        LabSwitchTo.Text = String.Concat(vKBID, ".", vKBNAME)

        'DataGrid10_ItemDataBound
        Dim rParms2 As New Hashtable
        rParms2.Add("BCID", Convert.ToString(drOB("BCID")))
        rParms2.Add("RID", Convert.ToString(drOB("RID")))
        rParms2.Add("AppStage", Convert.ToString(drOB("AppStage")))
        Dim sSql2 As String = ""
        sSql2 &= " WITH WT1 AS (SELECT a.TECHID,a.BCFID,a.BCFTID,a.PATTERN,a.MEMO1,a.WAIVED" & vbCrLf
        sSql2 &= " ,a.SRCFILENAME1,a.FILENAME1" & vbCrLf
        sSql2 &= " ,ob.YEARS,rr.PLANID,rr.RID,ob.BCASENO,kb.ORGKINDGW,kb.KBID,bf.KBSID" & vbCrLf
        sSql2 &= " FROM VIEW_RIDNAME rr" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL bf ON bf.RID=rr.RID" & vbCrLf
        sSql2 &= " JOIN KEY_BIDCASE kb on kb.KBSID=bf.KBSID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASE ob ON ob.BCID=bf.BCID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL_TT a ON a.BCFID=bf.BCFID" & vbCrLf
        sSql2 &= " WHERE ob.BCID=@BCID AND rr.RID=@RID)" & vbCrLf

        sSql2 &= " ,WT2 AS (SELECT DISTINCT P2.TECHID" & vbCrLf
        sSql2 &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
        sSql2 &= " JOIN dbo.ORG_BIDCASEPI bp ON bp.PlanID=P1.PlanID AND bp.ComIDNO=P1.ComIDNO AND bp.SeqNo=P1.SeqNo" & vbCrLf
        sSql2 &= " JOIN dbo.V_PLAN_TEACHER1 P2 ON P1.PlanID=P2.PlanID AND P1.ComIDNO=P2.ComIDNO AND P1.SeqNo=P2.SeqNo" & vbCrLf
        sSql2 &= " WHERE P1.TransFlag='N' AND P1.IsApprPaper='Y' AND P1.AppliedResult IS NULL AND P1.RESULTBUTTON IS NULL" & vbCrLf
        sSql2 &= " AND bp.BCID=@BCID AND P1.RID=@RID AND P1.AppStage=@AppStage)" & vbCrLf
        'sSql2 &= " WHERE p1.RID=@RID)" & vbCrLf

        sSql2 &= " SELECT a.TECHID,a.RID,a.TEACHCNAME,a.TEACHENAME,a.TEACHERID" & vbCrLf
        sSql2 &= " ,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql2 &= " ,a.KINDENGAGE,case a.KINDENGAGE when '1' then '內聘(專任)' else '外聘(兼任)' end KINDENGAGE_N" & vbCrLf
        sSql2 &= " ,a.WORKSTATUS,case a.WORKSTATUS when '1' then '是' else '否' end WORKSTATUS_N" & vbCrLf
        sSql2 &= " ,(SELECT x.KINDNAME FROM ID_KINDOFTEACHER x WHERE x.KINDID=a.KINDID) KINDNAME" & vbCrLf
        sSql2 &= " ,bt.BCFID,bt.BCFTID,bt.PATTERN,bt.MEMO1,bt.WAIVED" & vbCrLf
        'sSql2 &= " ,case when bt.TECHID IS NOT NULL THEN CONCAT(bt.YEARS,'/',bt.PLANID,'/',bt.RID,'/',bt.BCASENO,'/',bt.KBSID,'/') END PATH1" & vbCrLf
        sSql2 &= " ,bt.SRCFILENAME1,bt.FILENAME1,bt.FILENAME1 OKFLAG" & vbCrLf
        sSql2 &= " ,bt.YEARS,bt.ORGKINDGW,bt.KBID,bt.KBSID" & vbCrLf
        sSql2 &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sSql2 &= " JOIN WT2 t2 ON t2.TECHID=a.TECHID" & vbCrLf
        sSql2 &= " JOIN AUTH_RELSHIP b ON b.RID=a.RID" & vbCrLf
        sSql2 &= " JOIN WT1 bt on bt.TECHID=a.TECHID AND bt.RID=a.RID" & vbCrLf
        sSql2 &= " WHERE a.RID=@RID" & vbCrLf
        sSql2 &= " ORDER BY a.TEACHERID" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, rParms2)

        labmsg2.Text = ""
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG10_ROWS = dt2.Rows.Count

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vBCASENO As String = Convert.ToString(drOB("BCASENO"))
        Dim vKBSID As String = Convert.ToString(drKB("KBSID"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)
        With DataGrid10
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    ''' <summary>打包【訓練班別計畫表】(單位有上傳就算)</summary>
    ''' <param name="oConn"></param>
    ''' <param name="rPMS2"></param>
    ''' <returns></returns>
    Private Shared Function GET_ORG_BIDCASE_WAIVED_PI(oConn As SqlConnection, rPMS2 As Hashtable) As DataTable
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS2, "DISTID")
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS2, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS2, "APPSTAGE")
        Dim sParms As New Hashtable
        sParms.Add("DISTID", vDISTID)
        sParms.Add("YEARS", vYEARS)
        sParms.Add("APPSTAGE", vAPPSTAGE)
        Dim sSql As String = ""
        sSql &= " SELECT a.BCID,a.BCASENO,a.YEARS,a.DISTID,a.ORGID" & vbCrLf
        sSql &= " ,a.PLANID,a.RID,a.APPSTAGE" & vbCrLf
        sSql &= " ,a.BIDACCT,a.BIDDATE,a.BISTATUS" & vbCrLf
        sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL" & vbCrLf
        sSql &= " ,a.RESULTACCT,a.RESULTDATE,a.HISREVIEW" & vbCrLf
        sSql &= " ,b.BCFID,b.KBSID,c.KBID,c.KBNAME,c.ORGKINDGW" & vbCrLf
        sSql &= " ,b.FILENAME1,b.SRCFILENAME1" & vbCrLf
        sSql &= " ,b.WAIVED,b.RTUREASON,b.RTURESACCT,b.RTURESDATE" & vbCrLf
        sSql &= " FROM ORG_BIDCASE a" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL b on b.BCID=a.BCID AND b.WAIVED='PI'" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE c on c.KBSID=b.KBSID" & vbCrLf
        sSql &= " WHERE a.YEARS=@YEARS AND a.DISTID=@DISTID AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        sSql &= " AND a.BISTATUS='B'" & vbCrLf
        'sSql &= " AND a.BISTATUS IS NOT NULL AND a.BISTATUS IN ('B')" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, oConn, sParms)
        Return dt
    End Function

#End Region

#Region "Friend"
    ''' <summary>取得 vAPPSTAGE  簡單1個字</summary>
    ''' <param name="vAPPSTAGE"></param>
    ''' <returns></returns>
    Friend Shared Function Get_APPSTAGE_S(ByVal vAPPSTAGE As Object) As String
        If vAPPSTAGE Is Nothing OrElse Convert.ToString(vAPPSTAGE) = "" Then Return ""
        Return If(vAPPSTAGE = "1", "上", If(vAPPSTAGE = "2", "下", If(vAPPSTAGE = "3", "政", If(vAPPSTAGE = "4", "進", ""))))
    End Function
#End Region

    Private Sub DataGrid13_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid13.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
        Dim vSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")

        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        Dim vBCFEID As String = TIMS.GetMyValue(sCmdArg, "BCFEID")
        Dim vBCPID As String = TIMS.GetMyValue(sCmdArg, "BCPID")
        Dim vBCFID As String = TIMS.GetMyValue(sCmdArg, "BCFID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        If e.CommandArgument = "" OrElse vPlanID = "" OrElse vComIDNO = "" OrElse vSeqNo = "" Then Return '(程式有誤中斷執行)

        Select Case e.CommandName
            Case "DOWNLOAD13" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", vORGKINDGW)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCFEID", vBCFEID) '*
                TIMS.SetMyValue2(rPMS4, "BCPID", vBCPID) '*
                TIMS.SetMyValue2(rPMS4, "BCFID", vBCFID) '*
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1) '*
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
            Case "REPORT13" '列印
                '"13"'教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'PCSVALUE
                'Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)
                'Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
                Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
                Dim rPMS As New Hashtable
                rPMS.Clear()
                rPMS.Add("YEARS", vYEARS) 'rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("selsqlstr", String.Concat(vPlanID, "-", vComIDNO, "-", vSeqNo))
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                'W13_教學環境資料表
                Call RPT_SD_14_014(rPMS)
        End Select
    End Sub

    Private Sub DataGrid13_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid13.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG_PlanID As HtmlInputHidden = e.Item.FindControl("HDG_PlanID")
                Dim HDG_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG_ComIDNO")
                Dim HDG_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG_SeqNo")
                Dim BTN_DOWNLOAD13 As Button = e.Item.FindControl("BTN_DOWNLOAD13") '下載
                Dim BTN_REPORT13 As Button = e.Item.FindControl("BTN_REPORT13") '列印
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                HDG_PlanID.Value = Convert.ToString(drv("PlanID"))
                HDG_ComIDNO.Value = Convert.ToString(drv("ComIDNO"))
                HDG_SeqNo.Value = Convert.ToString(drv("SeqNo"))
                BTN_DOWNLOAD13.Visible = (Not IsDBNull(drv("FILENAME1")))
                'If Not IsDBNull(drv("FILENAME1")) Then
                '    LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                '    HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                'ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                '    LabFileName1.Text = cst_txt_免附文件
                'End If

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "ComIDNO", Convert.ToString(drv("ComIDNO")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "WAIVED", Convert.ToString(drv("WAIVED")))

                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                TIMS.SetMyValue(sCmdArg, "KBID", Convert.ToString(drv("KBID")))
                TIMS.SetMyValue(sCmdArg, "KBSID", Convert.ToString(drv("KBSID")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                TIMS.SetMyValue(sCmdArg, "BCFEID", Convert.ToString(drv("BCFEID")))
                TIMS.SetMyValue(sCmdArg, "BCPID", Convert.ToString(drv("BCPID")))
                TIMS.SetMyValue(sCmdArg, "BCFID", Convert.ToString(drv("BCFID")))
                BTN_DOWNLOAD13.CommandArgument = sCmdArg '下載
                BTN_REPORT13.CommandArgument = sCmdArg '列印
        End Select
    End Sub

    Private Sub DataGrid11_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid11.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vTECHID As String = TIMS.GetMyValue(sCmdArg, "TECHID")
        Dim vBCFT2ID As String = TIMS.GetMyValue(sCmdArg, "BCFT2ID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        If e.CommandArgument = "" OrElse vTECHID = "" OrElse vBCFT2ID = "" Then Return '(程式有誤中斷執行)

        Select Case e.CommandName
            Case "DOWNLOAD11" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", vORGKINDGW)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCFT2ID", vBCFT2ID) '*
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1) '*
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select

    End Sub

    Private Sub DataGrid11_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid11.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim chkItem1 As HtmlInputCheckBox = e.Item.FindControl("chkItem1")
                'chkItem1.Attributes("onclick") = String.Concat("selectOnlyThis('", chkItem1.ClientID, "',", iDG11_ROWS, ",'DataGrid11')")
                Dim HDG11_TechID As HtmlInputHidden = e.Item.FindControl("HDG11_TechID")
                Dim HDG11_RID As HtmlInputHidden = e.Item.FindControl("HDG11_RID")
                Dim BTN_DOWNLOAD11 As Button = e.Item.FindControl("BTN_DOWNLOAD11")
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                '0:未轉班,1:已轉班 '未轉班(依計畫查詢) 含列印
                HDG11_TechID.Value = Convert.ToString(drv("TechID"))
                HDG11_RID.Value = Convert.ToString(drv("RID"))
                'Dim BTN_DELFILE11 As Button = e.Item.FindControl("BTN_DELFILE11")
                'If Not IsDBNull(drv("FILENAME1")) Then
                '    LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                '    HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                'End If
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TECHID", Convert.ToString(drv("TECHID")))
                TIMS.SetMyValue(sCmdArg, "BCFT2ID", Convert.ToString(drv("BCFT2ID")))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "KBID", Convert.ToString(drv("KBID")))
                TIMS.SetMyValue(sCmdArg, "KBSID", Convert.ToString(drv("KBSID")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                BTN_DOWNLOAD11.CommandArgument = sCmdArg '下載
        End Select
    End Sub

    Private Sub DataGrid10_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid10.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vTECHID As String = TIMS.GetMyValue(sCmdArg, "TECHID")
        Dim vBCFTID As String = TIMS.GetMyValue(sCmdArg, "BCFTID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        If e.CommandArgument = "" OrElse vTECHID = "" OrElse vBCFTID = "" Then Return '(程式有誤中斷執行)

        Select Case e.CommandName
            Case "DOWNLOAD10" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", vORGKINDGW)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCFTID", vBCFTID) '*
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1) '*
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select
    End Sub

    Private Sub DataGrid10_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid10.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG10_TechID As HtmlInputHidden = e.Item.FindControl("HDG10_TechID")
                Dim HDG10_RID As HtmlInputHidden = e.Item.FindControl("HDG10_RID")
                Dim BTN_DOWNLOAD10 As Button = e.Item.FindControl("BTN_DOWNLOAD10")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                HDG10_TechID.Value = Convert.ToString(drv("TechID"))
                HDG10_RID.Value = Convert.ToString(drv("RID"))
                'If Not IsDBNull(drv("FILENAME1")) Then
                '    LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                '    HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                'End If
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TECHID", drv("TECHID"))
                TIMS.SetMyValue(sCmdArg, "BCFTID", drv("BCFTID"))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "KBID", Convert.ToString(drv("KBID")))
                TIMS.SetMyValue(sCmdArg, "KBSID", Convert.ToString(drv("KBSID")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                BTN_DOWNLOAD10.CommandArgument = sCmdArg '下載
        End Select
    End Sub

    ''' <summary>檔案打包下載(機構班級／機構案件)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_PACKAGE_DOWNLOAD1.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        If drOB Is Nothing Then Return

        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        Dim rPMS As New Hashtable
        rPMS.Add("ORGKINDGW", Hid_ORGKINDGW.Value)
        rPMS.Add("RID", Hid_RID.Value)
        rPMS.Add("BCID", Hid_BCID.Value)
        rPMS.Add("BCASENO", Hid_BCASENO.Value)
        ' "DOWNLOAD4" '下載
        Call ResponseZIPFileALL_BI(Me, objconn, rPMS)
    End Sub

    ''' <summary>【訓練班別計畫表】打包下載</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_TRAIN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_TRAIN_PACKAGE_DOWNLOAD1.Click
        '清理隱藏的參數
        Call ClearHidValue()

        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS) '計畫年度
        Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE) '申請階段
        Hid_RID.Value = ""
        If Hid_RID.Value = "" AndAlso RIDValue.Value.Length > 0 Then Hid_RID.Value = RIDValue.Value.Substring(0, 1)
        If Hid_RID.Value = "" AndAlso sm.UserInfo.LID = 1 Then Hid_RID.Value = sm.UserInfo.RID
        If Hid_RID.Value.Length > 1 Then Hid_RID.Value = Hid_RID.Value.Substring(0, 1)
        Dim vDISTID As String = TIMS.Get_DistID_RID(Hid_RID.Value, objconn)

        If v_sch_ddlYEARS = "" OrElse v_sch_ddlAPPSTAGE = "" Then
            Common.MessageBox(Me, "計畫年度與申請階段為必選!")
            Return
        ElseIf Hid_RID.Value = "" OrElse vDISTID = "" Then
            Common.MessageBox(Me, "訓練機構為必選!")
            Return
        ElseIf vDISTID = "" Then
            Common.MessageBox(Me, "轄區分署有誤，請重新選擇!")
            Return
        End If

        'DOWNLOAD '下載 打包【訓練班別計畫表】(單位有上傳就算)
        Dim rPMS As New Hashtable
        rPMS.Add("DISTID", vDISTID)
        rPMS.Add("YEARS", v_sch_ddlYEARS)
        rPMS.Add("APPSTAGE", v_sch_ddlAPPSTAGE)
        '打包【訓練班別計畫表】(單位有上傳就算)
        Dim dtOBWP As DataTable = GET_ORG_BIDCASE_WAIVED_PI(objconn, rPMS)
        If dtOBWP Is Nothing OrElse dtOBWP.Rows.Count = 0 Then
            Common.MessageBox(Me, String.Concat("【訓練班別計畫表】", TIMS.cst_NODATAMsg1))
            Return
        End If

        Call ResponseZIPFilePI(Me, objconn, rPMS, dtOBWP)
    End Sub

    Private Sub DataGrid14_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid14.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG_PlanID As HtmlInputHidden = e.Item.FindControl("HDG_PlanID")
                Dim HDG_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG_ComIDNO")
                Dim HDG_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG_SeqNo")
                Dim BTN_DOWNLOAD14 As Button = e.Item.FindControl("BTN_DOWNLOAD14") '下載

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                HDG_PlanID.Value = Convert.ToString(drv("PlanID"))
                HDG_ComIDNO.Value = Convert.ToString(drv("ComIDNO"))
                HDG_SeqNo.Value = Convert.ToString(drv("SeqNo"))
                BTN_DOWNLOAD14.Visible = (Not IsDBNull(drv("FILENAME1")))

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "ComIDNO", Convert.ToString(drv("ComIDNO")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "WAIVED", Convert.ToString(drv("WAIVED")))

                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                TIMS.SetMyValue(sCmdArg, "KBID", Convert.ToString(drv("KBID")))
                TIMS.SetMyValue(sCmdArg, "KBSID", Convert.ToString(drv("KBSID")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                TIMS.SetMyValue(sCmdArg, "BCPID", Convert.ToString(drv("BCPID")))
                TIMS.SetMyValue(sCmdArg, "BCFID", Convert.ToString(drv("BCFID")))
                TIMS.SetMyValue(sCmdArg, "BCFP3ID", Convert.ToString(drv("BCFP3ID")))
                BTN_DOWNLOAD14.CommandArgument = sCmdArg '下載
        End Select
    End Sub

    Private Sub DataGrid14_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid14.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
        Dim vSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")
        Dim vBCFP3ID As String = TIMS.GetMyValue(sCmdArg, "BCFP3ID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        If e.CommandArgument = "" OrElse vPlanID = "" OrElse vComIDNO = "" OrElse vSeqNo = "" OrElse vBCFP3ID = "" Then Return '(程式有誤中斷執行)

        Select Case e.CommandName
            Case "DOWNLOAD14" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", vORGKINDGW)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCFP3ID", vBCFP3ID) '*
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1) '*
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select
    End Sub

    Private Sub DataGrid13B_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid13B.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG_PlanID As HtmlInputHidden = e.Item.FindControl("HDG_PlanID")
                Dim HDG_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG_ComIDNO")
                Dim HDG_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG_SeqNo")
                Dim BTN_DOWNLOAD13B As Button = e.Item.FindControl("BTN_DOWNLOAD13B") '下載
                Dim BTN_REPORT13B As Button = e.Item.FindControl("BTN_REPORT13B") '列印
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                HDG_PlanID.Value = Convert.ToString(drv("PlanID"))
                HDG_ComIDNO.Value = Convert.ToString(drv("ComIDNO"))
                HDG_SeqNo.Value = Convert.ToString(drv("SeqNo"))
                BTN_DOWNLOAD13B.Visible = (Not IsDBNull(drv("FILENAME1")))
                'If Not IsDBNull(drv("FILENAME1")) Then
                '    LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                '    HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                'ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                '    LabFileName1.Text = cst_txt_免附文件
                'End If

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", Convert.ToString(drv("PlanID")))
                TIMS.SetMyValue(sCmdArg, "ComIDNO", Convert.ToString(drv("ComIDNO")))
                TIMS.SetMyValue(sCmdArg, "SeqNo", Convert.ToString(drv("SeqNo")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "WAIVED", Convert.ToString(drv("WAIVED")))

                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", Convert.ToString(drv("ORGKINDGW")))
                TIMS.SetMyValue(sCmdArg, "KBID", Convert.ToString(drv("KBID")))
                TIMS.SetMyValue(sCmdArg, "KBSID", Convert.ToString(drv("KBSID")))
                TIMS.SetMyValue(sCmdArg, "YEARS", Convert.ToString(drv("YEARS")))
                TIMS.SetMyValue(sCmdArg, "BCRTID", Convert.ToString(drv("BCRTID")))
                TIMS.SetMyValue(sCmdArg, "BCPID", Convert.ToString(drv("BCPID")))
                TIMS.SetMyValue(sCmdArg, "BCFID", Convert.ToString(drv("BCFID")))
                BTN_DOWNLOAD13B.CommandArgument = sCmdArg '下載
                BTN_REPORT13B.CommandArgument = sCmdArg '列印
        End Select
    End Sub

    Private Sub DataGrid13B_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid13B.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
        Dim vSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")

        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vYEARS As String = TIMS.GetMyValue(sCmdArg, "YEARS")
        Dim vBCPID As String = TIMS.GetMyValue(sCmdArg, "BCPID")
        Dim vBCFID As String = TIMS.GetMyValue(sCmdArg, "BCFID")
        Dim vBCRTID As String = TIMS.GetMyValue(sCmdArg, "BCRTID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        If e.CommandArgument = "" OrElse vPlanID = "" OrElse vComIDNO = "" OrElse vSeqNo = "" Then Return '(程式有誤中斷執行)

        Select Case e.CommandName
            Case "DOWNLOAD13B" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", vORGKINDGW)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "RID", Hid_RID.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "BCPID", vBCPID) '*
                TIMS.SetMyValue2(rPMS4, "BCFID", vBCFID) '*
                TIMS.SetMyValue2(rPMS4, "BCRTID", vBCRTID) '*
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1) '*
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return

            Case "REPORT13B" '列印
                '"13"'W13-1混成課程教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'PCSVALUE
                'Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)
                'Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
                Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, Hid_RID.Value, Hid_BCID.Value, Hid_BCASENO.Value)
                Dim rPMS As New Hashtable
                rPMS.Clear()
                rPMS.Add("YEARS", vYEARS) 'rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("selsqlstr", String.Concat(vPlanID, "-", vComIDNO, "-", vSeqNo))
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                'W13-1混成課程教學環境資料表
                Call RPT_SD_14_014B(rPMS)

        End Select
    End Sub

    Private Sub RPT_SD_14_014B(rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_014B" '0:未轉班' 1:已轉班
        Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
        Dim TSTPRINT As String = If(sPrint_Test = "Y", "2", "1") '測試區2／'正式區1 

        'Const cst_printFN2 As String = "SD_14_014_1" '2:變更待審
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim selsqlstr As String = TIMS.GetMyValue2(rPMS, "selsqlstr") 'vPCS
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")

        Dim sfilename1 As String = "" 'cst_printFN1
        Dim sMyValue As String = ""
        sfilename1 = cst_printFN1
        sMyValue &= String.Concat("&Years=", vYEARS_ROC)
        sMyValue &= "&selsqlstr=" & selsqlstr
        sMyValue &= "&TPlanID=" & vTPlanID
        sMyValue &= "&SYears=" & vYEARS
        sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfilename1, sMyValue)
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
