Imports System.IO

Partial Class TC_11_001_T54
    Inherits AuthBasePage

    '基本上委訓單位使用
    '申請階段管理-受理期間設定 APPLISTAGE
    'Dim fg_can_applistage As Boolean = False
    Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
    'KEY_BIDCASE/ORG_BIDCASE/ORG_BIDCASEFL/ORG_BIDCASEFL_PI

    'SELECT COMIDNO,COUNT(1) C FROM VIEW2B pp
    'WHERE pp.TransFlag='N' AND pp.IsApprPaper='Y' AND pp.AppliedResult IS NULL AND pp.RESULTBUTTON IS NULL
    'AND pp.TPLANID='28' AND pp.DISTID='001' AND pp.YEARS='2023' 
    'GROUP BY COMIDNO ORDER BY COUNT(1) DESC

    '最近一次版本送件
    Const cst_MTYPE_LATEST_SEND1 As String = "MTYPE_LATEST_SEND1"
    '最近一次版本-下載
    Const cst_MTYPE_LATEST_DOWN1 As String = "MTYPE_LATEST_DOWN1"

    ''' <summary>儲存(暫存)</summary>
    Const cst_ACTTYPE_BTN_SAVETMP1 As String = "BTN_SAVETMP1" '儲存(暫存)
    ''' <summary>'儲存後進下一步</summary>
    Const cst_ACTTYPE_BTN_SAVENEXT1 As String = "BTN_SAVENEXT1" '儲存後進下一步

    Dim tryFIND As String = ""
    Dim iDG08_ROWS As Integer = 0
    Dim iDG10_ROWS As Integer = 0
    Dim iDG11_ROWS As Integer = 0
    Dim iDG13_ROWS As Integer = 0
    Dim iDG13B_ROWS As Integer = 0
    Dim iDG14_ROWS As Integer = 0
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

    Const cst_txt_師資助教基本資料表 As String = "(師資助教基本資料表)"
    Const cst_txt_授課師資學經歷證書影本 As String = "(師資學經歷證書影本)"
    Const cst_txt_教學環境資料表 As String = "(教學環境資料表)"
    Const cst_txt_混成課程教學環境資料表 As String = "(混成課程教學環境資料表)"
    Const cst_txt_iCap課程原始申請資料 As String = "(iCap課程原始申請資料)"

    Const cst_REUPLOADED_MSG As String = "(已重新上傳)"

    'sPrintASPX1 = $"{cst_printASPX_R}{TIMS.Get_MRqID(Me)}"
    Const cst_printASPX_R As String = "../../SD/14/SD_14_002_R.aspx?ID="
    Dim sPrintASPX1 As String = "" 'SD_14_002_R

    'outTYPE: CLSNM,PCSVAL
    'Const cst_outTYPE_CLSNM As String = "CLSNM"
    'Const cst_outTYPE_PCSVAL As String = "PCSVAL"

    Const cst_ss_RqProcessType As String = "RqProcessType" 'Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1
    Const cst_DG1CMDNM_DELETE1 As String = "DELETE1"
    Const cst_DG1CMDNM_VIEW1 As String = "VIEW1"
    Const cst_DG1CMDNM_EDIT1 As String = "EDIT1"
    Const cst_DG1CMDNM_SENDOUT1 As String = "SENDOUT1"

    Const cst_DG2_退件原因_iCOLUMN As Integer = 2

    Const cst_tpmsg_enb1 As String = "(檢視功能)不能儲存"
    Const cst_tpmsg_enb2 As String = "(檢視功能)不能檔案上傳"
    Const cst_tpmsg_enb3 As String = "(檢視功能)不能送出"

    Const cst_tpmsg_enb4 As String = "(已送出)審查中"
    Const cst_tpmsg_enb5 As String = "(退件修正)請先修改再送審"
    Const cst_tpmsg_enb6 As String = "(已送出)不可刪除"
    Const cst_tpmsg_enb7 As String = "(退件修正)不可刪除"
    Const cst_tpmsg_enb8 As String = "(退件修正)有退件原因,可重新上傳"
    Const cst_tpmsg_enb9 As String = "尚未上傳文件"
    Const cst_stopmsg_11 As String = "申請階段受理期間未開放，請確認後再操作!"

    'CASE a.BISTATUS WHEN 'B' THEN '已送件' WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過'" & vbCrLf

    'Dim G_UPDRV As String = "~/UPDRV"
    'Dim G_UPDRV_JS As String = "../../UPDRV"

    Const cst_errMsg_1 As String = "資料有誤請重新查詢!"
    Const cst_errMsg_2 As String = "上傳檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)" 'Const cst_errMsg_2 As String = "上傳檔案壓縮時發生錯誤，請重新確認上傳檔案格式!"
    Const cst_errMsg_3 As String = "檔案位置錯誤!"
    Const cst_errMsg_4 As String = "檔案類型錯誤!"
    Const cst_errMsg_5 As String = "檔案類型錯誤，必須為PDF類型檔案!"
    Const cst_errMsg_5b As String = "檔案類型錯誤，內容必須為PDF檔案!"
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

    'Add new application cases and add reminder messages
    'Const Cst_messages1 As String = "請務必確認此年度/申請階段之所有欲研提班級都已送審，【新增申辦案件】後才送審的班級，將無法納入此次線上申辦案件清單中!"
    'Const cst_ss_messages1 As String = "messages1"

    'KEY_BIDCASE    'ORG_BIDCASE,ORG_BIDCASEPI,ORG_BIDCASEFL,ORG_BIDCASEFL_TT,ORG_BIDCASEFL_TT2    'VIEW2B
    Dim tmpMSG As String = ""
    Dim ff3 As String = ""
    'Dim dtKEY_BIDCASE As DataTable
    Dim objconn As SqlConnection = Nothing

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call CCreate11() '每次執行

        If Not IsPostBack Then
            Call CCreate1(0)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button3.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            'center.Attributes("onclick") = "showObj('HistoryList2');ShowFrame();"
            'HistoryRID.Attributes("onclick") = "ShowFrame();"
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        'Dim jsScript1 As String = Common.GetJsString(String.Concat(Cst_messages1, vbCrLf, "確定是否要新增?"))
        'BTN_ADDNEW1.Attributes("onclick") = String.Concat("return confirm('", jsScript1, "');")

        '(顯示1次訊息)(有顯示過就存session)
        'If Session(cst_ss_messages1) Is Nothing OrElse Convert.ToString(Session(cst_ss_messages1)) <> Cst_messages1 Then
        '    Session(cst_ss_messages1) = Cst_messages1
        '    Common.MessageBox(Me, Cst_messages1)
        'End If

    End Sub

    ''' <summary>每次執行</summary>
    Sub CCreate11()
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        '申請階段管理-受理期間設定 APPLISTAGE
        'fg_can_applistage = TIMS.CAN_APPLISTAGE_1(objconn)
        sPrintASPX1 = $"{cst_printASPX_R}{TIMS.Get_MRqID(Me)}"

        Hid_TPlanID.Value = sm.UserInfo.TPlanID
        '<add key = "UPLOAD_OJT_Path" value="~/UPDRV" />
        '<add key = "DOWNLOAD_OJT_Path" value="../../UPDRV" />
        'Dim vUPLOAD_OJT_Path As String = TIMS.Utl_GetConfigSet("UPLOAD_OJT_Path")
        'Dim vDOWNLOAD_OJT_Path As String = TIMS.Utl_GetConfigSet("DOWNLOAD_OJT_Path")
        'Dim G_UPDRV As String = "~/UPDRV"  'Dim G_UPDRV_JS As String = "../../UPDRV"
        'If (vUPLOAD_OJT_Path <> "") Then G_UPDRV = vUPLOAD_OJT_Path
        'If (vDOWNLOAD_OJT_Path <> "") Then G_UPDRV_JS = vDOWNLOAD_OJT_Path
    End Sub

    '設定 資料與顯示 狀況！
    Private Sub CCreate1(ByVal iNum As Integer)
        labmsg1.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        'If Session(cst_ss_messages1) IsNot Nothing Then Session(cst_ss_messages1) = Nothing
        TableDataGrid1.Visible = False

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

        '申請階段
        'Dim v_APPSTAGE As String = If(Now.Month < 7, "1", "2")
        'Dim v_APPSTAGE As String = TIMS.GET_CANUSE_APPSTAGE(objconn, CStr(sm.UserInfo.Years), TIMS.cst_APPSTAGE_PTYPE1_01)
        'sch_ddlAPPSTAGE = TIMS.Get_APPSTAGE2(sch_ddlAPPSTAGE)
        'Common.SetListItem(sch_ddlAPPSTAGE, v_APPSTAGE)

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
        FrameTableEdt2.Visible = If(iNum = 2, True, False)
        FrameTableSch1.Visible = If(iNum = 0, True, False)
        FrameTableEdt1.Visible = If(iNum = 1, True, False)
    End Sub

    ''' <summary>新增申辦案件</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_ADDNEW1_Click(sender As Object, e As EventArgs) Handles BTN_ADDNEW1.Click
        '清理隱藏的參數
        Call ClearHidValue()

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Dim sERRMSG As String = ""
        Dim fg_CanADDNEW2 As Boolean = Utl_SHOW_DATA1(sERRMSG, drRR)
        If Not fg_CanADDNEW2 AndAlso sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Return
        ElseIf Not fg_CanADDNEW2 Then
            Common.MessageBox(Me, "查詢資料有誤，請檢查輸入參數!")
            Return
        End If

        SHOW_Frame1(2)

        'Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        'Call SHOW_Detail_BIDCASE(drRR, Hid_BCID.Value, "")
    End Sub

    ''' <summary>'檢核新增(查詢)調整</summary>
    ''' <param name="sERRMSG"></param>
    ''' <returns></returns>
    Private Function Utl_SHOW_DATA1(ByRef sERRMSG As String, ByRef drRR As DataRow) As Boolean
        '	於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)。
        '	各項應備文件上傳會依項目依序分頁顯示，訓練單位也可跳頁填選。第一頁介面示意圖如下
        If drRR Is Nothing Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構有誤")
            Return False
        End If

        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGNAME As String = Convert.ToString(drRR("ORGNAME"))
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))
        Dim v_RID As String = Convert.ToString(drRR("RID"))
        Dim fg_TY1 As Boolean = (v_TPLANID = sm.UserInfo.TPlanID AndAlso v_YEARS = v_sch_ddlYEARS)
        If sm.UserInfo.LID = 2 AndAlso v_ORGID <> sm.UserInfo.OrgID Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構有誤")
            Return False
        ElseIf sm.UserInfo.LID > 0 AndAlso v_PLANID <> sm.UserInfo.PlanID Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構計畫有誤")
            Return False
        ElseIf sm.UserInfo.LID = 0 AndAlso (Not fg_TY1) Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " (署)新增訓練機構有誤 計畫年度有誤")
            Return False
        ElseIf v_ORGLEVEL < 2 Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構層級有誤")
            Return False
        End If
        If v_sch_ddlYEARS = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增計畫年度不可為空")
            Return False
        End If

        '同一年度／轄區 (計畫)PLANID／申請階段APPSTAGE，每個(訓練單位)RID只能有一筆申辦案件
        'Dim sParms2 As New Hashtable From {{"PLANID", v_PLANID}, {"RID", v_RID}}
        'Dim sSql2 As String = "SELECT 1 FROM ORG_BIDCASE WHERE PLANID=@PLANID AND RID=@RID AND APPSTAGE IS NULL"
        'Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, sParms2)
        'If dt2.Rows.Count > 0 Then
        '    sERRMSG = "同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能有一筆申辦案件(或使用查詢修改)"
        '    Return False
        'End If

        lab_ORGNAME_1.Text = v_ORGNAME
        lab_YEARS_1.Text = TIMS.GET_YEARS_ROC(v_YEARS)

        Dim S1Parms As New Hashtable From {{"TPLANID", v_TPLANID}, {"RID", v_RID}, {"PLANID", v_PLANID}}
        Dim s_PCS_NOT_IN As String = GET_PCS_NOT_IN(sm, objconn, S1Parms)
        If s_PCS_NOT_IN <> "" Then S1Parms.Add("PCS_NOT_IN", TIMS.CombiSQLIN(s_PCS_NOT_IN))
        '於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)
        Dim dtPP As DataTable = TIMS.GET_CLASS2S_BIdt(objconn, S1Parms)
        Dim strCLASSNAME2S As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_CLSNM)
        Dim v_PCS_Value As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_PCSVAL)
        If strCLASSNAME2S = "" OrElse v_PCS_Value = "" Then
            sERRMSG = "查無已送審班級清單!"
            Return False
        End If

        With CBL_CLASSNAME1S
            .DataSource = dtPP
            .DataTextField = "CLASSCNAME3"
            .DataValueField = "PCS"
            .DataBind()
            .Items.Insert(0, New ListItem("全部", ""))
        End With

        CBL_CLASSNAME1S.Attributes("onclick") = "SelectAll('CBL_CLASSNAME1S','CBL_CLASSNAME1S_hid');"

        Return True
    End Function

    Private Function GET_PCS_NOT_IN(sm As SessionModel, oConn As SqlConnection, rPMS As Hashtable) As String
        Dim rst As String = ""
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")

        Dim dtPP As DataTable = Nothing
        Dim sParms1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"RID", vRID}, {"PLANID", vPLANID}}
        Dim sSql As String = ""
        sSql &= " SELECT pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.ORGNAME,pp.DISTNAME,pp.RID" & vbCrLf
        sSql &= " ,pp.DISTANCE,pp.STDATE,pp.FTDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(pp.STDATE) STDATE_ROC" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(pp.FTDATE) FTDATE_ROC" & vbCrLf
        sSql &= " ,pp.PSNO28,pp.PCS,pp.OCID,pp.DISTID,pp.APPSTAGE,pp.ORGID,pp.CLASSCNAME2" & vbCrLf
        sSql &= " ,pp.TransFlag,pp.IsApprPaper,pp.AppliedResult,pp.RESULTBUTTON" & vbCrLf
        sSql &= " FROM VIEW2B pp" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEPI bp ON bp.PLANID=pp.PLANID AND bp.COMIDNO=pp.COMIDNO AND bp.SEQNO=pp.SEQNO" & vbCrLf
        sSql &= " WHERE pp.TPLANID=@TPLANID AND pp.RID=@RID AND pp.PLANID=@PLANID" & vbCrLf
        dtPP = DbAccess.GetDataTable(sSql, oConn, sParms1)
        If dtPP Is Nothing OrElse dtPP.Rows.Count = 0 Then Return rst
        rst = ""
        For Each drPP As DataRow In dtPP.Rows
            rst &= String.Concat(If(rst <> "", ",", ""), drPP("PCS"))
        Next
        Return rst
    End Function

    Private Function Utl_ADDNEW_DATA2(ByRef sERRMSG As String, ByRef drRR As DataRow) As Boolean
        '	各項應備文件上傳會依項目依序分頁顯示，訓練單位也可跳頁填選。第一頁介面示意圖如下
        If drRR Is Nothing Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構有誤")
            Return False
        End If
        Dim v_CBL_CLASSNAME1S As String = TIMS.GetCblValue(CBL_CLASSNAME1S)
        If v_CBL_CLASSNAME1S = "" Then
            sERRMSG = "查無已送審班級清單 (請勾選已送審班級清單)!"
            Return False
        End If

        tr_HISREVIEW.Visible = False '歷程資訊
        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGNAME As String = Convert.ToString(drRR("ORGNAME"))
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))
        Dim v_RID As String = Convert.ToString(drRR("RID"))
        Dim fg_TY1 As Boolean = (v_TPLANID = sm.UserInfo.TPlanID AndAlso v_YEARS = v_sch_ddlYEARS)
        If sm.UserInfo.LID = 2 AndAlso v_ORGID <> sm.UserInfo.OrgID Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構有誤")
            Return False
        ElseIf sm.UserInfo.LID > 0 AndAlso v_PLANID <> sm.UserInfo.PlanID Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構計畫有誤")
            Return False
        ElseIf sm.UserInfo.LID = 0 AndAlso (Not fg_TY1) Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " (署)新增訓練機構有誤 計畫年度有誤")
            Return False
        ElseIf v_ORGLEVEL < 2 Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構層級有誤")
            Return False
        End If
        If v_sch_ddlYEARS = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增計畫年度不可為空")
            Return False
        End If

        '同一年度／轄區 (計畫)PLANID／申請階段APPSTAGE，每個(訓練單位)RID只能有一筆申辦案件
        'Dim sParms2 As New Hashtable From {{"PLANID", v_PLANID}, {"RID", v_RID}}
        'Dim sSql2 As String = "SELECT 1 FROM ORG_BIDCASE WHERE PLANID=@PLANID AND RID=@RID AND APPSTAGE IS NULL"
        'Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, sParms2)
        'If dt2.Rows.Count > 0 Then
        '    sERRMSG = "同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能有一筆申辦案件(或使用查詢修改)"
        '    Return False
        'End If

        '於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)
        Dim S1Parms As New Hashtable From {{"TPLANID", v_TPLANID}, {"RID", v_RID}, {"PLANID", v_PLANID}, {"PCS_IN", TIMS.CombiSQLIN(v_CBL_CLASSNAME1S)}}
        Dim dtPP As DataTable = TIMS.GET_CLASS2S_BIdt(objconn, S1Parms)
        Dim strCLASSNAME2S As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_CLSNM)
        Dim v_PCS_Value As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_PCSVAL)
        If strCLASSNAME2S = "" OrElse v_PCS_Value = "" Then
            sERRMSG = "查無已送審班級清單"
            Return False
        End If

        Hid_KBSID.Value = ""
        Hid_KBID.Value = ""

        Dim vBCASENO_NN As String = ""
        Dim iBCID As Integer = 0
        Try
            iBCID = DbAccess.GetNewId(objconn, "ORG_BIDCASE_BCID_SEQ,ORG_BIDCASE,BCID")
            Dim irParms As New Hashtable From {
                {"DISTID", v_DISTID},
                {"RID", v_RID}
            }
            vBCASENO_NN = GET_BCASENO_54NN(objconn, irParms)

            'iParms.Add("APPSTAGE", v_sch_ddlAPPSTAGE)
            'iParms.Add("BIDACCT", sm.UserInfo.UserID)
            'iParms.Add("BIDDATE", BIDDATE)
            'iParms.Add("BISTATUS", BISTATUS)
            'iParms.Add("CREATEDATE", CREATEDATE)
            Dim iParms As New Hashtable From {
                {"BCID", iBCID},
                {"BCASENO", vBCASENO_NN},
                {"YEARS", v_YEARS},
                {"DISTID", v_DISTID},
                {"ORGID", v_ORGID},
                {"PLANID", v_PLANID},
                {"RID", v_RID},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            'iParms.Add("MODIFYDATE", MODIFYDATE)
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_BIDCASE(BCID,BCASENO,YEARS,DISTID,ORGID" & vbCrLf
            isSql &= " ,PLANID,RID ,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)" & vbCrLf ',BIDACCT,BIDDATE,BISTATUS
            isSql &= " VALUES (@BCID,@BCASENO,@YEARS,@DISTID,@ORGID" & vbCrLf
            isSql &= " ,@PLANID,@RID ,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf ',@BIDACCT,@BIDDATE,@BISTATUS
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Catch ex As Exception
            sERRMSG = "資料庫序號有誤，請重新操作!"
            Return False
        End Try
        Hid_BCID.Value = iBCID
        Hid_BCASENO.Value = vBCASENO_NN

        Hid_ORGKINDGW.Value = $"{drRR("ORGKINDGW")}"
        labCLASSNAME2S.Text = TIMS.GetResponseWrite(strCLASSNAME2S)
        Hid_PCS.Value = v_PCS_Value

        '審核班級儲存
        Call SAVE_ORG_BIDCASEPI(sm, objconn, iBCID, v_PCS_Value)

        Return True
    End Function

    ''' <summary>審核班級儲存</summary>
    ''' <param name="iBCID"></param>
    Public Shared Sub SAVE_ORG_BIDCASEPI(sm As SessionModel, oConn As SqlConnection, ByVal iBCID As Integer, ByVal PCS_Value As String)
        If iBCID = 0 OrElse PCS_Value = "" Then Return
        Dim saPCSALL1 As String() = PCS_Value.Split(",")

        For Each sPCS1 As String In saPCSALL1
            Dim saPCS1 As String() = sPCS1.Split("x")
            If saPCS1.Length = 3 Then
                Dim pPLANID As String = saPCS1(0)
                Dim pCOMIDNO As String = saPCS1(1)
                Dim pSEQNO As String = saPCS1(2)
                Dim iBCPID As Integer = 0
                Dim sParms3 As New Hashtable From {{"PLANID", pPLANID}, {"COMIDNO", pCOMIDNO}, {"SEQNO", pSEQNO}}
                Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEPI WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, oConn, sParms3)
                If dt3.Rows.Count = 0 Then
                    iBCPID = DbAccess.GetNewId(oConn, "ORG_BIDCASEPI_BCPID_SEQ,ORG_BIDCASEPI,BCPID")
                    Dim iParms2 As New Hashtable From {
                        {"BCPID", iBCPID},
                        {"PLANID", TIMS.CINT1(pPLANID)},
                        {"COMIDNO", pCOMIDNO},
                        {"SEQNO", TIMS.CINT1(pSEQNO)},
                        {"BCID", iBCID},
                        {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    Dim isSql2 As String = ""
                    isSql2 &= " INSERT INTO ORG_BIDCASEPI(BCPID,PLANID,COMIDNO,SEQNO,BCID,MODIFYACCT,MODIFYDATE)" & vbCrLf
                    isSql2 &= " VALUES (@BCPID,@PLANID,@COMIDNO,@SEQNO,@BCID,@MODIFYACCT,GETDATE())" & vbCrLf
                    TIMS.ExecuteNonQuery(isSql2, oConn, iParms2)
                End If

            End If
        Next
    End Sub

    ''' <summary>儲存-iCap課程原始申請資料</summary>
    ''' <param name="drOB"></param>
    ''' <param name="drKB"></param>
    ''' <param name="drRR"></param>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_ORG_BIDCASEFL_PI3(drOB As DataRow, drKB As DataRow, drRR As DataRow, rPMS As Hashtable)
        Dim vPCS14 As String = TIMS.GetMyValue2(rPMS, "PCS14")
        Dim v_PlanID As String = TIMS.GetMyValue2(rPMS, "PlanID")
        Dim v_ComIDNO As String = TIMS.GetMyValue2(rPMS, "ComIDNO")
        Dim v_SeqNo As String = TIMS.GetMyValue2(rPMS, "SeqNo")
        'Dim vBCPID As String = TIMS.GetMyValue2(rPMS, "BCPID")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")

        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        'Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        '取得目前的序號找不到就不執行了
        Dim iBCPID As Integer = TIMS.GET_ORG_BIDCASEPI_iBCPID(sm, objconn, TIMS.CINT1(vBCID), v_PlanID, v_ComIDNO, v_SeqNo)
        If iBCPID <= 0 Then Return

        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        Dim iBCFID As Integer = -1
        'Const cst_WAIVED_TT As String = "TT"
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        'Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")
        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "YEARS", vYEARS)
            TIMS.SetMyValue2(rPMS2, "RID", vRID)
            TIMS.SetMyValue2(rPMS2, "BCID", vBCID)
            TIMS.SetMyValue2(rPMS2, "KBSID", vKBSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", cst_08_1_iCap課程原始申請資料_WAIVED_PI3)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            iBCFID = SAVE_ORG_BIDCASEFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try
        If iBCFID <= 0 Then
            Common.MessageBox(Me, "資料儲存有誤，請重新操作程式!")
            Exit Sub
        End If

        Dim fPMSP3 As New Hashtable From {{"BCFID", iBCFID}, {"BCID", TIMS.CINT1(vBCID)}, {"KBSID", TIMS.CINT1(vKBSID)}, {"BCPID", iBCPID}}
        Dim fSqlP3 As String = "SELECT * FROM ORG_BIDCASEFL_PI3 WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND BCPID=@BCPID"
        Dim drFLP3 As DataRow = DbAccess.GetOneRow(fSqlP3, objconn, fPMSP3)
        If drFLP3 IsNot Nothing AndAlso Hid_RTUREASON.Value = "" Then
            TIMS.LOG.Warn(String.Concat("##ORG_BIDCASEFL_PI3 EXISTS!!!", ",BCFID:", iBCFID, ",BCID:", vBCID, ",KBSID:", vKBSID, ",BCPID:", iBCPID))
            Common.MessageBox(Me, "已上傳或儲存過該文件，不可再次操作!!!")
            Return
        End If

        'SAVE_ORG_BIDCASEFL_PI3()( ORG_BIDCASEFL_PI3 沒有資料才進行新增儲存)
        If drFLP3 Is Nothing Then
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_BIDCASEFL_PI3(BCFP3ID,BCFID,BCID,KBSID,BCPID" & vbCrLf
            isSql &= " ,FILENAME1,SRCFILENAME1, MODIFYACCT,MODIFYDATE)" & vbCrLf
            isSql &= " VALUES (@BCFP3ID,@BCFID,@BCID,@KBSID,@BCPID" & vbCrLf
            isSql &= " ,@FILENAME1,@SRCFILENAME1, @MODIFYACCT,GETDATE())" & vbCrLf
            Dim iBCFP3ID As Integer = DbAccess.GetNewId(objconn, "ORG_BIDCASEFL_PI3_BCFP3ID_SEQ,ORG_BIDCASEFL_PI3,BCFP3ID")
            Dim iParmsP As New Hashtable From {
                {"BCFP3ID", iBCFP3ID},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"BCPID", iBCPID},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            DbAccess.ExecuteNonQuery(isSql, objconn, iParmsP)
        Else
            Dim oYEARS As String = $"{drOB("YEARS")}"
            Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
            Dim oPLANID As String = $"{drOB("PLANID")}"
            Dim oRID As String = $"{drOB("RID")}"
            Dim oBCASENO As String = $"{drOB("BCASENO")}"
            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
            Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Try
                oFILENAME1 = $"{drFLP3("FILENAME1")}"
                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Call TIMS.MyFileDelete(s_FilePath1)
            Catch ex As Exception
                Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
            End Try

            Dim iBCFP3ID As Integer = TIMS.CINT1(drFLP3("BCFP3ID"))
            Dim U_PMS As New Hashtable From {
                {"BCFP3ID", iBCFP3ID},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"BCPID", iBCPID},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim U_SQL As String = "UPDATE ORG_BIDCASEFL_PI3 
SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()
FROM ORG_BIDCASEFL_PI3 WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND BCPID=@BCPID AND BCFP3ID=@BCFP3ID"
            DbAccess.ExecuteNonQuery(U_SQL, objconn, U_PMS)

        End If
    End Sub

    ''' <summary>SAVE_ORG_BIDCASEFL_TT</summary>
    ''' <param name="drOB"></param>
    ''' <param name="drKB"></param>
    ''' <param name="drRR"></param>
    ''' <param name="rPMS"></param>
    Sub SAVE_ORG_BIDCASEFL_TT(drOB As DataRow, drKB As DataRow, drRR As DataRow, ByRef rPMS As Hashtable)
        Dim vTECHID As String = TIMS.GetMyValue2(rPMS, "TECHID")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")

        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        'Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        Dim iBCFID As Integer = -1
        'Const cst_WAIVED_TT As String = "TT"
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        'Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")
        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "YEARS", vYEARS)
            TIMS.SetMyValue2(rPMS2, "RID", vRID)
            TIMS.SetMyValue2(rPMS2, "BCID", vBCID)
            TIMS.SetMyValue2(rPMS2, "KBSID", vKBSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", cst_10_師資助教基本資料表_WAIVED_TT)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            iBCFID = SAVE_ORG_BIDCASEFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try
        If iBCFID <= 0 Then
            Common.MessageBox(Me, "資料儲存有誤，請重新操作程式!")
            Exit Sub
        End If

        Dim fPMST As New Hashtable From {
            {"BCFID", iBCFID},
            {"BCID", TIMS.CINT1(vBCID)},
            {"KBSID", TIMS.CINT1(vKBSID)},
            {"TECHID", TIMS.CINT1(vTECHID)}
        }
        Dim fSqlT As String = "SELECT * FROM ORG_BIDCASEFL_TT WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND TECHID=@TECHID"
        Dim drFLT As DataRow = DbAccess.GetOneRow(fSqlT, objconn, fPMST)
        If drFLT IsNot Nothing AndAlso Hid_RTUREASON.Value = "" Then
            TIMS.LOG.Warn(String.Concat("##ORG_BIDCASEFL_TT EXISTS!!!", ",BCFID:", iBCFID, ",BCID:", vBCID, ",KBSID:", vKBSID, ",TECHID:", vTECHID))
            Common.MessageBox(Me, "已上傳或儲存過該文件，不可再次操作!!!")
            Return
        End If

        'SAVE_ORG_BIDCASEFL_TT()(沒有資料才進行新增儲存)
        If drFLT Is Nothing Then
            Dim isSqlT As String = ""
            isSqlT &= " INSERT INTO ORG_BIDCASEFL_TT(BCFTID, BCFID, BCID, KBSID, TECHID" & vbCrLf
            isSqlT &= " ,FILENAME1,SRCFILENAME1,MODIFYACCT, MODIFYDATE)" & vbCrLf
            isSqlT &= " VALUES(@BCFTID,@BCFID,@BCID,@KBSID,@TECHID" & vbCrLf
            isSqlT &= " ,@FILENAME1,@SRCFILENAME1,@MODIFYACCT,GETDATE())" & vbCrLf
            Dim iBCFTID As Integer = DbAccess.GetNewId(objconn, "ORG_BIDCASEFL_TT_BCFTID_SEQ,ORG_BIDCASEFL_TT,BCFTID")
            Dim iParmsT As New Hashtable From {
                {"BCFTID", iBCFTID},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"TECHID", TIMS.CINT1(vTECHID)},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            DbAccess.ExecuteNonQuery(isSqlT, objconn, iParmsT)
        Else
            Dim oYEARS As String = $"{drOB("YEARS")}"
            Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
            Dim oPLANID As String = $"{drOB("PLANID")}"
            Dim oRID As String = $"{drOB("RID")}"
            Dim oBCASENO As String = $"{drOB("BCASENO")}"
            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
            Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Try
                oFILENAME1 = $"{drFLT("FILENAME1")}"
                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Call TIMS.MyFileDelete(s_FilePath1)
            Catch ex As Exception
                Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
            End Try

            Dim iBCFTID As Integer = TIMS.CINT1(drFLT("BCFTID"))
            Dim U_PMS As New Hashtable From {
                {"BCFTID", iBCFTID},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"TECHID", TIMS.CINT1(vTECHID)},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim U_SQL As String = "UPDATE ORG_BIDCASEFL_TT 
SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()
FROM ORG_BIDCASEFL_TT WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND TECHID=@TECHID AND BCFTID=@BCFTID"
            DbAccess.ExecuteNonQuery(U_SQL, objconn, U_PMS)
        End If
    End Sub

    ''' <summary>SAVE_ORG_BIDCASEFL_TT2</summary>
    ''' <param name="drOB"></param>
    ''' <param name="drKB"></param>
    ''' <param name="drRR"></param>
    ''' <param name="rPMS"></param>
    Sub SAVE_ORG_BIDCASEFL_TT2(drOB As DataRow, drKB As DataRow, drRR As DataRow, ByRef rPMS As Hashtable)
        Dim vTECHID As String = TIMS.GetMyValue2(rPMS, "TECHID")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")

        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        'Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        Dim iBCFID As Integer = -1
        'Const cst_WAIVED_TT As String = "TT"
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        'Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")
        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "YEARS", vYEARS)
            'TIMS.SetMyValue2(rPMS2, "APPSTAGE", vAPPSTAGE)
            TIMS.SetMyValue2(rPMS2, "RID", vRID)
            TIMS.SetMyValue2(rPMS2, "BCID", vBCID)
            TIMS.SetMyValue2(rPMS2, "KBSID", vKBSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", cst_11_授課師資學經歷證書影本_WAIVED_TT2)
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            iBCFID = SAVE_ORG_BIDCASEFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try
        If iBCFID <= 0 Then
            Common.MessageBox(Me, "資料儲存有誤，請重新操作程式!!")
            Exit Sub
        End If

        Dim fPMST2 As New Hashtable From {
            {"BCFID", iBCFID},
            {"BCID", TIMS.CINT1(vBCID)},
            {"KBSID", TIMS.CINT1(vKBSID)},
            {"TECHID", TIMS.CINT1(vTECHID)}
        }
        Dim fSqlT2 As String = "SELECT * FROM ORG_BIDCASEFL_TT2 WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND TECHID=@TECHID"
        Dim drFLT2 As DataRow = DbAccess.GetOneRow(fSqlT2, objconn, fPMST2)
        If drFLT2 IsNot Nothing AndAlso Hid_RTUREASON.Value = "" Then
            TIMS.LOG.Warn(String.Concat("##ORG_BIDCASEFL_TT2 EXISTS!!!", ",BCFID:", iBCFID, ",BCID:", vBCID, ",KBSID:", vKBSID, ",TECHID:", vTECHID))
            Common.MessageBox(Me, "已上傳或儲存過該文件，不可再次操作!!")
            Return
        End If

        'SAVE_ORG_BIDCASEFL_TT2()(沒有資料才進行新增儲存)
        If drFLT2 Is Nothing Then
            Dim isSqlT As String = ""
            isSqlT &= " INSERT INTO ORG_BIDCASEFL_TT2(BCFT2ID, BCFID, BCID, KBSID, TECHID" & vbCrLf
            isSqlT &= " ,FILENAME1,SRCFILENAME1,MODIFYACCT, MODIFYDATE)" & vbCrLf
            isSqlT &= " VALUES(@BCFT2ID,@BCFID,@BCID,@KBSID,@TECHID" & vbCrLf
            isSqlT &= " ,@FILENAME1,@SRCFILENAME1,@MODIFYACCT,GETDATE())" & vbCrLf
            Dim iBCFT2ID As Integer = DbAccess.GetNewId(objconn, "ORG_BIDCASEFL_TT2_BCFT2ID_SEQ,ORG_BIDCASEFL_TT2,BCFT2ID")
            Dim iParmsT As New Hashtable From {
                {"BCFT2ID", iBCFT2ID},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"TECHID", TIMS.CINT1(vTECHID)},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            DbAccess.ExecuteNonQuery(isSqlT, objconn, iParmsT)
        Else
            Dim oYEARS As String = $"{drOB("YEARS")}"
            Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
            Dim oPLANID As String = $"{drOB("PLANID")}"
            Dim oRID As String = $"{drOB("RID")}"
            Dim oBCASENO As String = $"{drOB("BCASENO")}"
            Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
            Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Try
                oFILENAME1 = $"{drFLT2("FILENAME1")}"
                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Call TIMS.MyFileDelete(s_FilePath1)
            Catch ex As Exception
                Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
            End Try

            Dim iBCFT2ID As Integer = TIMS.CINT1(drFLT2("BCFT2ID"))
            Dim U_PMS As New Hashtable From {
                {"BCFT2ID", iBCFT2ID},
                {"BCFID", iBCFID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"TECHID", TIMS.CINT1(vTECHID)},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim U_SQL As String = "UPDATE ORG_BIDCASEFL_TT2 
SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()
FROM ORG_BIDCASEFL_TT2 WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND TECHID=@TECHID AND BCFT2ID=@BCFT2ID"
            DbAccess.ExecuteNonQuery(U_SQL, objconn, U_PMS)
        End If
    End Sub

    ''' <summary>清理隱藏的參數</summary>
    Sub ClearHidValue()
        Hid_KBSID.Value = ""
        Hid_KBID.Value = ""
        Hid_LastKBID.Value = ""
        Hid_FirstKBSID.Value = ""

        Hid_TECHID.Value = ""
        Hid_BCID.Value = ""
        Hid_BCFID.Value = ""
        Hid_BCASENO.Value = ""
        Hid_ORGKINDGW.Value = ""
        Hid_PCS.Value = ""
        Hid_BISTATUS.Value = ""
        Hid_RTUREASON.Value = ""
    End Sub

    ''' <summary>查詢鈕1</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH1_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH1.Click
        Dim sERRMSG1 As String = ""
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        Dim flag_CHECKOK As Boolean = CHK_Search1(sm, drRR, sERRMSG1)
        If sERRMSG1 <> "" Then
            labmsg1.Text = sERRMSG1
            Common.MessageBox(Me, sERRMSG1)
            Return
        ElseIf Not flag_CHECKOK OrElse drRR Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Call SSearch1(drRR)
    End Sub

    ''' <summary>查詢1</summary>
    Private Sub SSearch1(ByRef drRR As DataRow)
        '清理隱藏的參數
        Call ClearHidValue()

        labmsg1.Text = TIMS.cst_NODATAMsg1
        TableDataGrid1.Visible = False

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
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
        'RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        'Dim v_sch_ddlAPPSTAG As String = TIMS.GetListValue(sch_ddlAPPSTAGE)

        '檢核查詢
        'Dim sERRMSG1 As String = ""
        'Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        'Dim flag_CHECKOK As Boolean = CHK_Search1(sm, drRR, sERRMSG1)
        'If sERRMSG1 <> "" Then
        '    labmsg1.Text = sERRMSG1
        '    Common.MessageBox(Me, sERRMSG1)
        '    Return
        'ElseIf Not flag_CHECKOK OrElse drRR Is Nothing Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
        '    Return
        'End If

        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        'Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        'Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))

        'NULL 待送審、B 審核中、Y 審核通過、R 退件修正、N 審核不通過。
        Dim pParms As New Hashtable
        If sch_txtBCASENO.Text <> "" Then pParms.Add("BCASENO", sch_txtBCASENO.Text)
        If v_sch_ddlYEARS <> "" Then pParms.Add("YEARS", v_sch_ddlYEARS)
        'If v_sch_ddlAPPSTAG <> "" Then pParms.Add("APPSTAGE", v_sch_ddlAPPSTAG)
        If sch_txtBINAME.Text <> "" Then pParms.Add("BINAME", sch_txtBINAME.Text) 'sSql &= " AND u.NAME LIKE '%'+@BINAME+'%'" & vbCrLf
        If sch_txtBIDATE1.Text <> "" Then pParms.Add("BIDDATE1", sch_txtBIDATE1.Text) 'sSql &= " AND a.BIDDATE >=@BIDDATE1" & vbCrLf
        If sch_txtBIDATE2.Text <> "" Then pParms.Add("BIDDATE2", sch_txtBIDATE2.Text) 'sSql &= " AND a.BIDDATE <=@BIDDATE2" & vbCrLf

        Dim sSql As String = ""
        sSql &= " SELECT a.BCID,a.BCASENO,a.YEARS,a.DISTID,a.ORGID,a.PLANID,a.RID" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        'APPSTAGE_N
        'sSql &= " ,a.APPSTAGE,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
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
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='Y' THEN '分署已收件'" & vbCrLf '通過
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
        sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END BISTATUS_N" & vbCrLf
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL"
        sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
        'sSql &= " ,a.CREATEACCT,a.CREATEDATE,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        sSql &= " FROM ORG_BIDCASE a" & vbCrLf
        sSql &= " JOIN VIEW_RIDNAME r on r.RID=a.RID" & vbCrLf
        sSql &= " LEFT JOIN AUTH_ACCOUNT u ON u.ACCOUNT=a.BIDACCT" & vbCrLf
        sSql &= " WHERE r.TPLANID=@TPLANID" & vbCrLf
        pParms.Add("TPLANID", sm.UserInfo.TPlanID)

        Select Case sm.UserInfo.LID
            Case 0
                If v_ORGLEVEL = 2 Then
                    sSql &= " AND r.RID=@RID" & vbCrLf
                    pParms.Add("RID", RIDValue.Value) 'sm.UserInfo.RID
                Else
                    sSql &= " AND r.DISTID=@DISTID" & vbCrLf
                    pParms.Add("DISTID", v_DISTID)
                    If v_sch_ddlYEARS = "" Then
                        sSql &= " AND r.YEARS=@YEARS" & vbCrLf
                        pParms.Add("YEARS", sm.UserInfo.Years)
                    End If
                End If
            Case 1
                If v_ORGLEVEL = 2 Then
                    sSql &= " AND r.RID=@RID" & vbCrLf
                    pParms.Add("RID", RIDValue.Value) 'sm.UserInfo.RID
                Else
                    sSql &= " AND r.DISTID=@DISTID" & vbCrLf
                    sSql &= " AND r.PLANID=@PLANID" & vbCrLf
                    pParms.Add("DISTID", sm.UserInfo.DistID)
                    pParms.Add("PLANID", sm.UserInfo.PlanID)
                    If v_sch_ddlYEARS = "" Then
                        sSql &= " AND r.YEARS=@YEARS" & vbCrLf
                        pParms.Add("YEARS", sm.UserInfo.Years)
                    End If
                End If
            Case 2
                sSql &= " AND r.RID=@RID" & vbCrLf
                sSql &= " AND r.ORGID=@ORGID" & vbCrLf
                sSql &= " AND r.ORGLEVEL=@ORGLEVEL" & vbCrLf
                sSql &= " AND r.DISTID=@DISTID" & vbCrLf
                pParms.Add("RID", RIDValue.Value) 'sm.UserInfo.RID
                pParms.Add("ORGID", sm.UserInfo.OrgID)
                pParms.Add("ORGLEVEL", sm.UserInfo.OrgLevel)
                pParms.Add("DISTID", v_DISTID)
            Case Else
                sSql &= " AND 1<>1" & vbCrLf
        End Select

        If sch_txtBCASENO.Text <> "" Then sSql &= " AND a.BCASENO=@BCASENO" & vbCrLf
        If v_sch_ddlYEARS <> "" Then sSql &= " AND a.YEARS=@YEARS" & vbCrLf
        'If v_sch_ddlAPPSTAG <> "" Then sSql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        If sch_txtBINAME.Text <> "" Then sSql &= " AND u.NAME LIKE '%'+@BINAME+'%'" & vbCrLf
        If sch_txtBIDATE1.Text <> "" Then sSql &= " AND a.BIDDATE >=@BIDDATE1" & vbCrLf
        If sch_txtBIDATE2.Text <> "" Then sSql &= " AND a.BIDDATE <=@BIDDATE2" & vbCrLf

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

    ''' <summary>檢核查詢／新增</summary>
    ''' <param name="drRR"></param>
    ''' <returns></returns>
    Function CHK_Search1(ByRef sm As SessionModel, ByRef drRR As DataRow, ByRef sERRMSG1 As String) As Boolean
        If drRR Is Nothing Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構有誤!")
            Return False
        End If

        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        'Dim v_sch_ddlAPPSTAG As String = TIMS.GetListValue(sch_ddlAPPSTAGE)
        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))

        Dim fg_TY1 As Boolean = (v_TPLANID = sm.UserInfo.TPlanID AndAlso v_YEARS = v_sch_ddlYEARS)
        If sm.UserInfo.LID = 2 AndAlso v_ORGID <> sm.UserInfo.OrgID Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (委訓單位)查詢訓練機構有誤!", v_ORGID))
            Return False
        ElseIf sm.UserInfo.LID > 0 AndAlso v_PLANID <> sm.UserInfo.PlanID Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (請選擇對應計畫年度機構)訓練機構計畫有誤!", v_PLANID))
            Return False
        ElseIf sm.UserInfo.LID = 0 AndAlso (Not fg_TY1) Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (署)(請選擇對應計畫年度機構)訓練機構計畫年度有誤!", v_YEARS))
            Return False
        ElseIf v_ORGLEVEL = "" OrElse TIMS.CINT1(v_ORGLEVEL) <> 2 Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" 訓練機構層級有誤!", v_ORGLEVEL))
            Return False
        End If

        If v_sch_ddlYEARS = "" Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 計畫年度不可為空!!")
            Return False
            'ElseIf v_sch_ddlAPPSTAG = "" Then
            '    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 申請階段不可為空!!")
            '    Return False
        End If
        Return True
        '(下面目前不會執行，但之後可能會用到)

        Select Case sm.UserInfo.LID
            Case 2
                Dim iPLANID As Integer = If(v_PLANID <> "", TIMS.CINT1(v_PLANID), 0)
                'If RIDValue.Value <> sm.UserInfo.RID Then
                '    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " (委訓單位)訓練機構有誤")
                '    Return False
                'End If
                If v_ORGID = "" OrElse v_ORGID <> sm.UserInfo.OrgID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (委訓單位)查詢訓練機構有誤!", v_ORGID))
                    Return False
                ElseIf iPLANID = 0 OrElse iPLANID <> sm.UserInfo.PlanID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (委訓單位)訓練機構計畫有誤!", iPLANID))
                    Return False
                ElseIf v_ORGLEVEL = "" OrElse TIMS.CINT1(v_ORGLEVEL) <> 2 Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (委訓單位)訓練機構層級有誤!", v_ORGLEVEL))
                    Return False
                End If
                If v_sch_ddlYEARS = "" Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 計畫年度不可為空!!!")
                    Return False
                    'ElseIf v_sch_ddlAPPSTAG = "" Then
                    '    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 申請階段不可為空!!!")
                    '    Return False
                End If
            Case 1
                'Dim iPLANID As Integer = If(v_PLANID <> "", TIMS.CINT1(v_PLANID), 0)
                'If iPLANID <> 0 AndAlso iPLANID <> sm.UserInfo.PlanID Then
                '    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構計畫有誤")
                '    Return False
                'End If
                If v_sch_ddlYEARS = "" Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 計畫年度不可為空!")
                    Return False
                    'ElseIf v_sch_ddlAPPSTAG = "" Then
                    '    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 申請階段不可為空!")
                    '    Return False
                End If
            Case 0
                'Dim fg_TY1 As Boolean = (v_TPLANID = sm.UserInfo.TPlanID AndAlso v_YEARS = v_sch_ddlYEARS)
                'If (Not fg_TY1) AndAlso v_ORGLEVEL = "0" Then
                '    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " (署)訓練機構計畫年度有誤")
                '    Return False
                'End If
        End Select
        Return True
    End Function

    ''' <summary>不儲存返回查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_BACK1_Click(sender As Object, e As EventArgs) Handles BTN_BACK1.Click
        '清理隱藏的參數
        Call ClearHidValue()

        Call SHOW_Frame1(0)
    End Sub

    ''' <summary>切換至(文件項目)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub DdlSwitchTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSwitchTo.SelectedIndexChanged
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_KBSID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_FirstKBSID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    ''' <summary>暫時儲存／正式儲存-UPDATE ORG_BIDCASE</summary>
    ''' <param name="iNum"></param>
    Private Sub SAVEDATE1(ByVal iNum As Integer)
        'iNum:0 暫時儲存/1 正式儲存
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim iBCID As Integer = If(Hid_BCID.Value <> "", TIMS.CINT1(Hid_BCID.Value), 0)
        If Hid_BCID.Value = "" OrElse iBCID <= 0 Then Return

        Dim uParms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"BCID", iBCID},
            {"RID", RIDValue.Value},
            {"BCASENO", Hid_BCASENO.Value}
        }
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_BIDCASE" & vbCrLf
        usSql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE BCID=@BCID AND RID=@RID AND BCASENO=@BCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        '審核班級儲存
        'Call SAVE_ORG_BIDCASEPI(iBCID)
    End Sub

    ''' <summary>按下 查看／修改／送出 </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vRID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim vBCID As String = TIMS.GetMyValue(sCmdArg, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue(sCmdArg, "BCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        If sCmdArg = "" OrElse vBCID = "" OrElse vRID = "" Then Return

        Dim drRR As DataRow = TIMS.Get_RID_DR(vRID, objconn) 'If drRR Is Nothing Then Return
        Call SHOW_RIDValue_DATA(drRR)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
        If drOB Is Nothing Then Return
        Dim vYEARS As String = $"{drOB("YEARS")}"

        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        '申請階段管理-受理期間設定 APPLISTAGE
        'Dim aParms As New Hashtable
        'aParms.Add("YEARS", vYEARS)
        'aParms.Add("APPSTAGE", vAPPSTAGE)
        'Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE01(objconn, aParms)

        Dim s_RESULTDATE_YMS2 As String = If($"{drOB("RESULTDATE")}" <> "", CDate(drOB("RESULTDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
        Dim s_MODIFYDATE_YMS2 As String = If($"{drOB("MODIFYDATE")}" <> "", CDate(drOB("MODIFYDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
        Dim fg_RESULTDATE_UPDATE As Boolean = (s_RESULTDATE_YMS2 <> "" AndAlso s_MODIFYDATE_YMS2 <> "" AndAlso DateDiff(DateInterval.Second, CDate(s_RESULTDATE_YMS2), CDate(s_MODIFYDATE_YMS2)) > 0)

        Select Case e.CommandName
            Case cst_DG1CMDNM_DELETE1 'DELETE1 (刪除)
                Call DELETE_Detail_BIDCASE(Me, objconn, drRR, drOB)
                Common.MessageBox(Me, TIMS.cst_DELETEOKMsg2)
                Call SSearch1(drRR)

            Case cst_DG1CMDNM_VIEW1 '"VIEW1 '查看
                Call SHOW_Detail_BIDCASE(drRR, vBCID, cst_DG1CMDNM_VIEW1)

            Case cst_DG1CMDNM_EDIT1 '"EDIT1 '修改
                'If Not fg_can_applistage Then
                '    Common.MessageBox(Me, cst_stopmsg_11) 'Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                '    Return
                'End If
                Call SHOW_Detail_BIDCASE(drRR, vBCID, cst_DG1CMDNM_EDIT1)

            Case cst_DG1CMDNM_SENDOUT1 'SENDOUT1 送出 
                'If Not fg_can_applistage Then
                '    Common.MessageBox(Me, cst_stopmsg_11) 'Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                '    Return
                'End If

                '線上申辦進度 計算完成度百分比 (0-100) 'Dim vORGKINDGW As String = $"{drOB("ORGKINDGW")}"
                Dim iProgress As Integer = TIMS.GET_iPROGRESS_BI(sm, objconn, tmpMSG, vBCID, vORGKINDGW)
                Dim EMSG As String = ""
                If iProgress < 100 Then
                    EMSG = $"線上申辦進度 未達100%，不可送出! ({iProgress}){vbCrLf}{If(tmpMSG <> "", $"請檢查：({tmpMSG})", "")}"
                    Common.MessageBox(Me, EMSG)
                    Return
                ElseIf $"{drOB("BISTATUS")}" = "R" AndAlso Not fg_RESULTDATE_UPDATE Then
                    EMSG = cst_tpmsg_enb5
                    Common.MessageBox(Me, EMSG)
                    Return
                End If

                Dim dtPI As DataTable = TIMS.GET_ORG_BIDCASEPI(objconn, vBCID)
                If dtPI Is Nothing OrElse dtPI.Rows.Count = 0 Then
                    EMSG = "線上申辦-查無班級資料無法送出!"
                    Common.MessageBox(Me, EMSG)
                    Return
                End If
                '"線上申辦中所列班級於「班級申請」都未按送出"
                If dtPI IsNot Nothing AndAlso dtPI.Rows.Count > 0 Then
                    For Each drPI As DataRow In dtPI.Rows
                        Dim oPLANID As String = Convert.ToString(drPI("PLANID"))
                        Dim oCOMIDNO As String = Convert.ToString(drPI("COMIDNO"))
                        Dim oSEQNO As String = Convert.ToString(drPI("SEQNO"))
                        Dim drPP As DataRow = TIMS.GetPCSDate(oPLANID, oCOMIDNO, oSEQNO, objconn)
                        If drPP Is Nothing Then
                            EMSG = String.Concat("請再次確認所有班級均已送審!查無班級資料無法送出!_", oPLANID, "x", oCOMIDNO, "x", oSEQNO)
                            Common.MessageBox(Me, EMSG)
                            Return
                        End If
                        'TransFlag,IsApprPaper,AppliedResult,RESULTBUTTON
                        Dim s_CLASSNAME As String = String.Concat(drPP("CLASSNAME"), "(", drPP("STDATE_ROC"), ")")
                        Dim fg_PP1_TransFlag As Boolean = (Convert.ToString(drPP("TransFlag")) = "N")
                        Dim fg_PP1_IsApprPaper As Boolean = (Convert.ToString(drPP("IsApprPaper")) = "Y")
                        Dim fg_PP1_AppliedResult As Boolean = Convert.IsDBNull(drPP("AppliedResult"))
                        Dim fg_PP1_RESULTBUTTON As Boolean = Convert.IsDBNull(drPP("RESULTBUTTON"))
                        Dim fg_PP1_ALL As Boolean = (fg_PP1_TransFlag AndAlso fg_PP1_IsApprPaper AndAlso fg_PP1_AppliedResult AndAlso fg_PP1_RESULTBUTTON)
                        If Not fg_PP1_ALL Then
                            EMSG = String.Concat("請再次確認所有班級均已送審!於「班級申請」未按送出!_", s_CLASSNAME)
                            Common.MessageBox(Me, EMSG)
                            Return
                        End If
                    Next
                End If

                ',CASE a.BISTATUS WHEN 'B' THEN '已送件' WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過'" & vbCrLf
                'uParms.Add("BIDDATE", BIDDATE)
                Dim uParms As New Hashtable From {
                    {"BIDACCT", sm.UserInfo.UserID},
                    {"BISTATUS", "B"},
                    {"MODIFYACCT", sm.UserInfo.UserID},
                    {"BCID", vBCID},
                    {"RID", vRID},
                    {"BCASENO", vBCASENO}
                }
                Dim usSql As String = ""
                usSql &= " UPDATE ORG_BIDCASE" & vbCrLf
                usSql &= " SET BIDACCT=@BIDACCT,BIDDATE=GETDATE(),BISTATUS=@BISTATUS" & vbCrLf
                usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE BCID=@BCID and RID=@RID and BCASENO=@BCASENO" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

                Call SSearch1(drRR)

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim lBTN_DELETE1 As LinkButton = e.Item.FindControl("lBTN_DELETE1") '(刪除)
                Dim lBTN_VIEW1 As LinkButton = e.Item.FindControl("lBTN_VIEW1") '查看
                Dim lBTN_EDIT1 As LinkButton = e.Item.FindControl("lBTN_EDIT1") '修改 
                Dim lBTN_SENDOUT1 As LinkButton = e.Item.FindControl("lBTN_SENDOUT1") '送出 

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "BCID", drv("BCID"))
                TIMS.SetMyValue(sCmdArg, "RID", drv("RID"))
                TIMS.SetMyValue(sCmdArg, "BCASENO", drv("BCASENO"))
                TIMS.SetMyValue(sCmdArg, "BISTATUS", drv("BISTATUS"))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", drv("ORGKINDGW"))

                Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                lBTN_DELETE1.Visible = If(flagS1, True, False)
                lBTN_DELETE1.Style.Item("display") = If(flagS1, "", "none")
                lBTN_DELETE1.CommandArgument = sCmdArg
                Dim vMsgB As String = "請注意：刪除此提案資料將無法還原，只能重新申請!是否確定要刪除?"
                lBTN_DELETE1.Attributes("onclick") = $"javascript:return confirm('{vMsgB}');"

                lBTN_VIEW1.CommandArgument = sCmdArg
                lBTN_EDIT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.Attributes("onclick") = "javascript:return confirm('此動作會送出審核資料且不可再次修改，是否確定?');"

                'BISTATUS 'lBTN_VIEW1.Enabled = If($"{drv("BISTATUS")}" <> "", False, True)
                If $"{drv("BISTATUS")}" = "R" Then
                    lBTN_EDIT1.Enabled = True
                    TIMS.Tooltip(lBTN_EDIT1, cst_tpmsg_enb5, True)
                    lBTN_SENDOUT1.Enabled = True
                    TIMS.Tooltip(lBTN_SENDOUT1, cst_tpmsg_enb5, True)
                Else
                    lBTN_EDIT1.Enabled = If($"{drv("BISTATUS")}" <> "", False, True)
                    TIMS.Tooltip(lBTN_EDIT1, If(lBTN_EDIT1.Enabled, "", cst_tpmsg_enb4), True)
                    lBTN_SENDOUT1.Enabled = If($"{drv("BISTATUS")}" <> "", False, True)
                    TIMS.Tooltip(lBTN_SENDOUT1, If(lBTN_SENDOUT1.Enabled, "", cst_tpmsg_enb4), True)
                End If

        End Select
    End Sub

    ''' <summary>新增使用資料顯示／查詢使用資料顯示 依 ORG_BIDCASE-BCID</summary>
    Private Sub SHOW_Detail_BIDCASE(ByRef drRR As DataRow, ByVal vBCID As String, ByVal vCmdName As String)
        '訓練機構有誤
        If drRR Is Nothing Then Return
        Call SHOW_Frame1(1)

        Dim rLastKBID As String = ""
        Dim rFirstKBSID As String = ""
        Session(cst_ss_RqProcessType) = vCmdName
        Dim vRID As String = Convert.ToString(drRR("RID"))
        Dim vPLANID As String = Convert.ToString(drRR("PLANID"))
        Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))
        ddlSwitchTo = TIMS.GET_ddlBIDCASE(sm, objconn, ddlSwitchTo, Hid_ORGKINDGW.Value)

        tr_HISREVIEW.Visible = False '歷程資訊

        Dim dtB1 As DataTable = Nothing
        If vBCID <> "" Then
            '查詢資料
            'NULL 待送審、B 審核中、Y 審核通過、R 退件修正、N 審核不通過。
            Dim pParms As New Hashtable From {{"BCID", vBCID}, {"RID", vRID}, {"PLANID", vPLANID}}
            Dim sSql As String = ""
            sSql &= " SELECT a.BCID,a.BCASENO,a.YEARS,a.PLANID,a.DISTID,a.ORGID,a.RID" & vbCrLf
            'APPSTAGE_N
            'sSql &= " ,a.APPSTAGE,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
            'DISTNAME
            sSql &= " ,dbo.FN_GET_DISTNAME(a.DISTID,3) DISTNAME" & vbCrLf
            'ORGNAME
            sSql &= " ,(SELECT ORGNAME FROM ORG_ORGINFO WHERE ORGID=a.ORGID) ORGNAME" & vbCrLf
            'COMIDNO
            sSql &= " ,(SELECT COMIDNO FROM ORG_ORGINFO WHERE ORGID=a.ORGID) COMIDNO" & vbCrLf
            'BINAME
            sSql &= " ,a.BIDACCT,dbo.FN_GET_USERNAME(a.BIDACCT) BINAME" & vbCrLf
            sSql &= " ,format(a.BIDDATE,'yyyy/MM/dd') BIDDATE" & vbCrLf
            sSql &= " ,dbo.FN_CDATE1B(a.BIDDATE) BIDDATE_ROC" & vbCrLf
            '申辦狀態：暫存/ 已送件
            sSql &= " ,a.BISTATUS ,CASE WHEN a.BISTATUS IS NULL THEN '暫存'" & vbCrLf
            sSql &= "  WHEN a.BISTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'" & vbCrLf
            sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'" & vbCrLf
            sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='Y' THEN '分署已收件'" & vbCrLf '通過
            sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
            sSql &= "  WHEN a.BISTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END BISTATUS_N" & vbCrLf
            '申辦狀態：暫存/ 已送件
            '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
            sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL"
            sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
            'tr_HISREVIEW.Visible = False '歷程資訊
            sSql &= " ,a.HISREVIEW" & vbCrLf

            sSql &= " ,(SELECT MAX(KBSID) FROM ORG_BIDCASEFL fl WHERE fl.BCID=a.BCID) CurrentKBSID" & vbCrLf
            sSql &= " ,(SELECT MIN(KBSID) FROM ORG_BIDCASEFL fl WHERE fl.BCID=a.BCID AND fl.RTUREASON IS NOT NULL) Curr2KBSID" & vbCrLf
            'sSql &= " ,a.CREATEACCT,a.CREATEDATE,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
            sSql &= " ,format(a.CREATEDATE,'yyyy-MM-dd HH:mm') CREATEDATE_F" & vbCrLf
            sSql &= " FROM ORG_BIDCASE a" & vbCrLf
            sSql &= " WHERE BCID=@BCID AND a.RID=@RID AND a.PLANID=@PLANID" & vbCrLf
            dtB1 = DbAccess.GetDataTable(sSql, objconn, pParms)
            If dtB1 Is Nothing OrElse dtB1.Rows.Count = 0 Then Return
            Dim drB1 As DataRow = dtB1.Rows(0)

            'Hid_KBSID.Value = ""
            'Hid_KBID.Value = ""
            'vAPPSTAGE = Convert.ToString(drB1("APPSTAGE"))
            Hid_BCID.Value = Convert.ToString(drB1("BCID"))
            Hid_BCASENO.Value = Convert.ToString(drB1("BCASENO"))

            labOrgNAME.Text = Convert.ToString(drB1("ORGNAME"))
            LabCREATEDATE.Text = Convert.ToString(drB1("CREATEDATE_F"))
            labBIYEARS.Text = TIMS.GET_YEARS_ROC(drB1("YEARS"))
            'labAPPSTAGE.Text = Convert.ToString(drB1("APPSTAGE_N"))

            '退件
            If Convert.ToString(drB1("BISTATUS")) = "R" Then
                Hid_KBSID.Value = Convert.ToString(drB1("Curr2KBSID"))
            Else
                Hid_KBSID.Value = Convert.ToString(drB1("CurrentKBSID"))
            End If

            '歷程資訊
            If Convert.ToString(drB1("HISREVIEW")) <> "" AndAlso Convert.ToString(drB1("HISREVIEW")).Length > 1 Then
                tr_HISREVIEW.Visible = True '歷程資訊
                labHISREVIEW.Text = Convert.ToString(drB1("HISREVIEW"))
            End If
            'labAPPSTAGE.Text = TIMS.GET_APPSTAGE2_NM2(dr1("APPSTAGE"))
            'labProgress.Text = "0%"
            'Call Utl_SwitchTo("1")
        Else
            Call TIMS.Utl_GET_SWITCHTO_VAL(sm, objconn, Hid_ORGKINDGW.Value, rLastKBID, rFirstKBSID)
            Hid_LastKBID.Value = rLastKBID
            Hid_FirstKBSID.Value = rFirstKBSID

            '(新增資料) '檢核查詢
            Dim sERRMSG1 As String = ""
            Dim flag_CHECKOK As Boolean = CHK_Search1(sm, drRR, sERRMSG1)
            If sERRMSG1 <> "" Then
                labmsg1.Text = sERRMSG1
                Common.MessageBox(Me, sERRMSG1)
                Return
            ElseIf Not flag_CHECKOK OrElse drRR Is Nothing Then
                Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                Return
            End If
            'orgid_value.Value = TIMS.ClearSQM(orgid_value.Value)
            labOrgNAME.Text = Convert.ToString(drRR("ORGNAME"))
            labBIYEARS.Text = TIMS.GET_YEARS_ROC(drRR("YEARS"))
            'labAPPSTAGE.Text = TIMS.GET_APPSTAGE2_NM2(v_sch_ddlAPPSTAGE)
            'Call Utl_SwitchTo(Hid_FirstKBSID.Value)
        End If

        '{"APPSTAGE", If(vAPPSTAGE <> "", vAPPSTAGE, v_sch_ddlAPPSTAGE)}
        Dim S1Parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"RID", vRID}, {"PLANID", vPLANID}}
        If vBCID <> "" Then S1Parms.Add("BCID", vBCID)
        Dim v_CBL_CLASSNAME1S As String = TIMS.GetCblValue(CBL_CLASSNAME1S)
        If v_CBL_CLASSNAME1S <> "" Then S1Parms.Add("PCS_IN", TIMS.CombiSQLIN(v_CBL_CLASSNAME1S))
        If vBCID = "" AndAlso v_CBL_CLASSNAME1S = "" Then
            Common.MessageBox(Me, "查無「線上申辦」序號與「已送審班級」清單!")
            Return
        End If

        Dim dtPP As DataTable = TIMS.GET_CLASS2S_BIdt(objconn, S1Parms)
        Dim strCLASSNAME2S As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_CLSNM)
        Dim v_PCS_Value As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_PCSVAL)
        Dim v_PP_DISTANCE As String = TIMS.GET_CLASSNAME2S_BI(dtPP, TIMS.cst_outTYPE_PP_DISTANCE)
        Hid_PP_DISTANCE.Value = TIMS.ClearSQM(v_PP_DISTANCE)

        '排除停用選項
        If Hid_PP_DISTANCE.Value <> "" Then Call RESET_SWITCHTO_VAL(ddlSwitchTo)

        Call TIMS.Utl_GET_SWITCHTO_VAL(sm, objconn, Hid_ORGKINDGW.Value, rLastKBID, rFirstKBSID)
        Hid_LastKBID.Value = rLastKBID
        Hid_FirstKBSID.Value = rFirstKBSID

        '	於當年度/申請階段之開放班級申請期間，可新增申辦案件(非申辦期間無法新增)。
        '	同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能有一筆申辦案件。
        '	於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)。
        '	各項應備文件上傳會依項目依序分頁顯示，訓練單位也可跳頁填選。第一頁介面示意圖如下
        '班級名稱
        labCLASSNAME2S.Text = TIMS.GetResponseWrite(strCLASSNAME2S)
        Hid_PCS.Value = v_PCS_Value

        labProgress.Text = "0%"

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_FirstKBSID.Value = TIMS.ClearSQM(Hid_FirstKBSID.Value)

        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)

        If Hid_KBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_FirstKBSID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    Private Sub RESET_SWITCHTO_VAL(ByRef ddlSwitchTo As DropDownList)
        Hid_PP_DISTANCE.Value = TIMS.ClearSQM(Hid_PP_DISTANCE.Value)

        If Hid_PP_DISTANCE.Value = "" Then Return

        'true: 可使用 false:不可使用
        Dim fg_Can_USE_13_1_混成課程教學環境資料表 As Boolean = (String.Format(",{0},", Hid_PP_DISTANCE.Value).IndexOf(",2,") > -1)

        If fg_Can_USE_13_1_混成課程教學環境資料表 Then Return '(可使用)不調整選項離開

        TIMS.OpenDbConn(objconn)
        Dim dtX13 As New DataTable
        Dim sql As String = "SELECT KBSID FROM KEY_BIDCASE WITH(NOLOCK) WHERE KBID='13-1' AND TPLANID=@TPLANID"
        Dim sCmd As New SqlCommand(sql, objconn)
        With sCmd
            .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
            dtX13.Load(.ExecuteReader())
        End With
        If dtX13 Is Nothing OrElse dtX13.Rows.Count = 0 Then Return
        For Each drX2 As DataRow In dtX13.Rows
            Dim litem As ListItem = ddlSwitchTo.Items.FindByValue(drX2("KBSID"))
            If litem Is Nothing Then Continue For
            ddlSwitchTo.Items.Remove(litem)
        Next
    End Sub

    ''' <summary>回上一步</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PREV1_Click(sender As Object, e As EventArgs) Handles BTN_PREV1.Click
        Call MOVE_PREV()
    End Sub

    ''' <summary>回上一步</summary>
    Private Sub MOVE_PREV()
        If (Hid_KBID.Value = "" OrElse Hid_KBID.Value = "01" OrElse ddlSwitchTo.SelectedIndex - 1 = -1) Then
            Common.MessageBox(Me, "(目前沒有上一步)")
            Return
        End If

        'TIMS.AddZero(Val(Hid_KBSID.Value) - 1, 2)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Hid_KBSID.Value = TIMS.CINT1(TIMS.GetListValue(ddlSwitchTo)) - 1
        Hid_KBSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex - 1).Value
        If Hid_KBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_FirstKBSID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    ''' <summary>暫時儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SAVETMP1_Click(sender As Object, e As EventArgs) Handles BTN_SAVETMP1.Click
        '儲存(暫存)
        Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVETMP1)
    End Sub

    ''' <summary>儲存後進下一步</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SAVENEXT1_Click(sender As Object, e As EventArgs) Handles BTN_SAVENEXT1.Click
        '儲存後進下一步
        Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVENEXT1)
    End Sub

    ''' <summary>暫時儲存／儲存後進下一步</summary>
    ''' <param name="s_ACTTYPE"></param>
    Sub SAVEDATA2_BTN_ACTION1(ByVal s_ACTTYPE As String)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_KBID.Value = TIMS.ClearSQM(Hid_KBID.Value)
        Hid_LastKBID.Value = TIMS.ClearSQM(Hid_LastKBID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)

        'SAVE_ORG_BIDCASEFL
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無案件編號)，請重新操作!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!")
            Return
        End If
        Dim vYEARS As String = $"{drOB("YEARS")}"

        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        ''申請階段管理-受理期間設定 APPLISTAGE
        'Dim aParms As New Hashtable
        'aParms.Add("YEARS", vYEARS)
        'aParms.Add("APPSTAGE", vAPPSTAGE)
        'Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE01(objconn, aParms)
        ''檢核查詢
        'If Not fg_can_applistage Then
        '    Common.MessageBox(Me, cst_stopmsg_11) ' "申請階段受理期間未開放，請確認後再操作!"
        '    Return
        'End If

        'Dim fg_FILE_MUSTBE_UPLOADED As Boolean = True '必須上傳檔案
        Dim fg_FILE_MUSTBE_UPLOADED As Boolean = True
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_W08_訓練班別計畫表, TIMS.cst_G08_訓練班別計畫表
                fg_FILE_MUSTBE_UPLOADED = False '(不必上傳檔案)
                vWAIVED = cst_08_訓練班別計畫表_WAIVED_PI
            Case TIMS.cst_W08_1_iCap課程原始申請資料, TIMS.cst_G08_1_iCap課程原始申請資料
                fg_FILE_MUSTBE_UPLOADED = False '(不必上傳檔案)
                vWAIVED = cst_08_1_iCap課程原始申請資料_WAIVED_PI3
            Case TIMS.cst_W10_師資助教基本資料表, TIMS.cst_G10_師資助教基本資料表
                fg_FILE_MUSTBE_UPLOADED = False '(不必上傳檔案)
                vWAIVED = cst_10_師資助教基本資料表_WAIVED_TT
            Case TIMS.cst_W11_授課師資學經歷證書影本, TIMS.cst_G11_授課師資學經歷證書影本
                fg_FILE_MUSTBE_UPLOADED = False '(不必上傳檔案)
                vWAIVED = cst_11_授課師資學經歷證書影本_WAIVED_TT2
            Case TIMS.cst_W13_教學環境資料表, TIMS.cst_G13_教學環境資料表
                fg_FILE_MUSTBE_UPLOADED = False '(不必上傳檔案)
                vWAIVED = cst_13_教學環境資料表_WAIVED_PI2
            Case TIMS.cst_W13_1_混成課程教學環境資料表, TIMS.cst_G13_1_混成課程教學環境資料表
                fg_FILE_MUSTBE_UPLOADED = False '(不必上傳檔案)
                vWAIVED = cst_13_1混成課程教學環境資料表_WAIVED_RT2
        End Select

        Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        'Dim vKBSID As String = TIMS.ClearSQM(Hid_KBSID.Value)
        Dim drFL As DataRow = TIMS.GET_ORG_BIDCASEFL(objconn, vBCID, vKBSID)
        '(退件修正)有退件原因,可重新儲存
        Dim fg_CanSaveAgain_1 As Boolean = (drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) <> "") '(有資料 且原因不為空 可再次傳送)
        Dim vFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
        'Dim vWAIVED As String = If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
        If fg_FILE_MUSTBE_UPLOADED AndAlso Not fg_CanSaveAgain_1 AndAlso drFL IsNot Nothing AndAlso Not CHKB_WAIVED.Checked AndAlso vFILENAME1 = "" Then
            Common.MessageBox(Me, "未上傳檔案也未勾選免附且儲存過該文件，不可再次操作!")
            Return
        End If

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vPLANID As String = $"{drOB("PLANID")}"
        Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, "A", vPLANID, vRID, vBCASENO, "")
        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            If (drFL IsNot Nothing AndAlso fg_CanSaveAgain_1) Then TIMS.SetMyValue2(rPMS2, "BCFID", drFL("BCFID"))

            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", Hid_ORGKINDGW.Value)
            TIMS.SetMyValue2(rPMS2, "YEARS", drOB("YEARS"))
            'TIMS.SetMyValue2(rPMS2, "APPSTAGE", drOB("APPSTAGE"))
            TIMS.SetMyValue2(rPMS2, "RID", RIDValue.Value)
            TIMS.SetMyValue2(rPMS2, "BCID", Hid_BCID.Value)
            TIMS.SetMyValue2(rPMS2, "KBSID", Hid_KBSID.Value)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED) ' If(CHKB_WAIVED.Checked, "Y", ""))
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            Select Case vWAIVED
                Case "Y", ""
                    Call SAVE_ORG_BIDCASEFL_UPLOAD(rPMS2)
            End Select
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub  'Throw ex
        End Try

        '暫時儲存／正式儲存-UPDATE ORG_BIDCASE
        Call SAVEDATE1(0)

        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)

        Select Case s_ACTTYPE
            Case cst_ACTTYPE_BTN_SAVETMP1
                '儲存(暫存) 

                '項目(重跑1次)
                Call SHOW_BIDCASE_KBSID(Hid_KBSID.Value, Hid_ORGKINDGW.Value)

            Case cst_ACTTYPE_BTN_SAVENEXT1
                '儲存後進下一步 

                '(檢查儲存值)
                Dim rPMS As New Hashtable From {
                    {"ORGKINDGW", Hid_ORGKINDGW.Value},
                    {"BCID", Hid_BCID.Value}
                }
                Dim flag_OK_OBFL As Boolean = CHK_ORG_BIDCASEFL(rPMS, Hid_KBSID.Value)
                If Not flag_OK_OBFL Then
                    Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步")
                    Return
                End If

                '下一步
                Call MOVE_NEXT()
        End Select

    End Sub

    ''' <summary>後進下一步</summary>
    Private Sub MOVE_NEXT()
        If Hid_KBID.Value <> "" AndAlso (Hid_KBID.Value = Hid_LastKBID.Value) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        ElseIf (ddlSwitchTo.SelectedIndex + 1 >= ddlSwitchTo.Items.Count) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        End If

        '下一步
        'Hid_KBSID.Value = If(Hid_KBID.Value = "" OrElse Hid_KBSID.Value = "", 1, TIMS.CINT1(Hid_KBSID.Value) + 1)
        Hid_KBSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex + 1).Value
        Call SHOW_BIDCASE_KBSID(Hid_KBSID.Value, Hid_ORGKINDGW.Value)
    End Sub

    ''' <summary>下載報表</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_DOWNLOADRPT1_Click(sender As Object, e As EventArgs) Handles BTN_DOWNLOADRPT1.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vKBSID As String = TIMS.GetListValue(ddlSwitchTo)
        If Hid_BCASENO.Value = "" OrElse Hid_BCID.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(案件號為空)，請重新操作!!")
            Return
        ElseIf RIDValue.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(業務代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_KBSID.Value = "" OrElse Hid_ORGKINDGW.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_KBSID.Value <> "" AndAlso Hid_KBSID.Value <> vKBSID Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目序號有誤)，請重新操作!!")
            Return
        End If

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        '首頁>>訓練機構管理>>表單列印>>訓練單位基本資料表 '訓練單位基本資料
        'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_001_18G&Years=112&RSID=47877&planid=5093&rid=%27E6762%27&AppStage=1&UserID=snoopy
        '列印
        Call UTL_PRINT1GW(drOB, drKB)
    End Sub

#Region "PRINT_1"

    ''' <summary>列印</summary>
    ''' <param name="drOB"></param>
    ''' <param name="drKB"></param>
    Sub UTL_PRINT1GW(ByRef drOB As DataRow, ByRef drKB As DataRow)
        If drOB Is Nothing OrElse drKB Is Nothing Then Return
        'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        'Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        'Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        ''(取得)
        'Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        ''取得文字說明
        'Dim vKBDESC1 As String = $"{drKB("KBDESC1")}"
        ''必填資訊／免附文件(必填就不顯示)
        'Dim vMUSTFILL As String = $"{drKB("MUSTFILL")}"
        ''USELATESTVER : 以最近一次版本送件
        'Dim vUSELATESTVER As String = $"{drKB("USELATESTVER")}"
        ''DOWNLOADRPT '可下載報表
        'Dim vDOWNLOADRPT As String = $"{drKB("DOWNLOADRPT")}"

        '取得KBID代號／非流水號
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Dim vKBNAME2 As String = String.Concat(vORGKINDGW, vKBID, drKB("KBNAME"))
        'Dim vRID As String = $"{drOB("RID")}"
        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim rPMS As New Hashtable
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_W04_辦理本計畫訓練課程之專職人員名冊, TIMS.cst_G05_辦理本計畫訓練課程之專職人員名冊
                '辦理本計畫訓練課程之專職人員名冊
                rPMS.Clear()
                rPMS.Add("Years", $"{drOB("YEARS")}")
                rPMS.Add("RID", $"{drOB("RID")}")
                rPMS.Add("FSQ1", String.Concat("0", drOB("APPSTAGE")))
                Call RPT_SD_14_018_2(rPMS)
            Case TIMS.cst_W06_訓練單位基本資料表, TIMS.cst_G06_訓練單位基本資料表
                '訓練單位基本資料表 / 首頁>>訓練機構管理>>表單列印>>訓練單位基本資料表
                Dim drRR As DataRow = TIMS.Get_RID_DR($"{drOB("RID")}", objconn)
                If drRR Is Nothing Then
                    Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
                    Return
                End If

                rPMS.Clear()
                rPMS.Add("ORGKINDGW", $"{drKB("ORGKINDGW")}")
                rPMS.Add("YEARS", $"{drOB("YEARS")}")
                rPMS.Add("RSID", Convert.ToString(drRR("RSID")))
                rPMS.Add("PLANID", $"{drOB("PLANID")}")
                rPMS.Add("RID", $"{drOB("RID")}")
                'rPMS.Add("APPSTAGE", $"{drOB("APPSTAGE")}")
                Call RPT_SD_14_001_18(rPMS)

            Case TIMS.cst_W07_訓練計畫總表, TIMS.cst_G07_訓練計畫總表
                '"07" '訓練計畫總表 / SD_14_003 '訓練機構管理>表單列印>訓練計畫總表 SD_14_003_2019G
                'Const cst_print_19W As String = "SD_14_003_19W" 'PP
                'Const cst_print_19G As String = "SD_14_003_19G" 'PP
                'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_003_2019W&OCID=150814&PlanID=5090&Years=112&AppStage=2&UserID=snoopy
                'PCSVALUE
                Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)

                rPMS.Clear()
                rPMS.Add("PCSVALUE", Hid_PCS.Value)
                rPMS.Add("ORGKINDGW", $"{drKB("ORGKINDGW")}")
                rPMS.Add("YEARS", $"{drOB("YEARS")}")
                'rPMS.Add("RSID", Convert.ToString(drRR("RSID")))
                rPMS.Add("PLANID", $"{drOB("PLANID")}")
                'rPMS.Add("RID", $"{drOB("RID")}")
                'rPMS.Add("APPSTAGE", $"{drOB("APPSTAGE")}")
                Call RPT_SD_14_003_19(rPMS)

                'Case cst_W08_訓練班別計畫表, cst_G08_訓練班別計畫表
                '"08" '訓練班別計畫表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_002?ID=273
                'rPMS.Clear()
                'rPMS.Add("ORGNAME", TIMS.GET_OrgName(drOB("ORGID"), objconn))
                'rPMS.Add("RID", $"{drOB("RID")}")
                'rPMS.Add("APPSTAGE", $"{drOB("APPSTAGE")}")
                'Call RPT_SD_14_002(rPMS)
            Case TIMS.cst_W09_訓練計畫師資助教名冊, TIMS.cst_G09_訓練計畫師資助教名冊
                '"09" '訓練計畫師資／助教名冊 SD_14_007_1 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_007?ID=278
                Dim drRR As DataRow = TIMS.Get_RID_DR($"{drOB("RID")}", objconn)
                If drRR Is Nothing Then
                    Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
                    Return
                End If

                'PCSVALUE
                Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)
                Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                rPMS.Clear()
                rPMS.Add("YEARS", $"{drOB("YEARS")}")
                rPMS.Add("Title", Convert.ToString(drRR("ORGPLANNAME")))
                rPMS.Add("selsqlstr", selsqlstr)
                Call RPT_SD_14_007_1(rPMS)

            Case TIMS.cst_W10_師資助教基本資料表, TIMS.cst_G10_師資助教基本資料表
                '"10" '師資／助教基本資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_004?ID=275
                Dim drRR As DataRow = TIMS.Get_RID_DR($"{drOB("RID")}", objconn)
                If drRR Is Nothing Then
                    Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
                    Return
                End If

                Dim rParms2 As New Hashtable From {
                    {"BCID", $"{drOB("BCID")}"},
                    {"RID", $"{drOB("RID")}"}
                }
                'rParms2.Add("AppStage", $"{drOB("AppStage")}")
                Dim sSql2 As String = ""
                sSql2 &= " SELECT DISTINCT P2.TECHID" & vbCrLf
                sSql2 &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
                sSql2 &= " JOIN dbo.ORG_BIDCASEPI bp ON bp.PlanID=P1.PlanID AND bp.ComIDNO=P1.ComIDNO AND bp.SeqNo=P1.SeqNo" & vbCrLf
                sSql2 &= " JOIN dbo.V_PLAN_TEACHER1 P2 ON P1.PlanID=P2.PlanID AND P1.ComIDNO=P2.ComIDNO AND P1.SeqNo=P2.SeqNo" & vbCrLf
                sSql2 &= " WHERE P1.TransFlag='N' AND P1.IsApprPaper='Y' AND P1.AppliedResult IS NULL AND P1.RESULTBUTTON IS NULL" & vbCrLf
                sSql2 &= " AND bp.BCID=@BCID AND P1.RID=@RID" & vbCrLf
                Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, rParms2)
                If dt2 Is Nothing Or dt2.Rows.Count = 0 Then
                    Common.MessageBox(Me, String.Concat("(", vKBNAME2, ")查無資料!"))
                    Return
                End If
                Dim vTechID As String = ""
                For Each dr2 As DataRow In dt2.Rows
                    vTechID &= String.Concat(If(vTechID <> "", ",", ""), "'", dr2("TechID"), "'")
                Next
                rPMS.Clear()
                rPMS.Add("TechID", vTechID)
                rPMS.Add("Years", drOB("YEARS"))
                rPMS.Add("Title", Convert.ToString(drRR("ORGPLANNAME")))
                Call RPT_SD_14_004(rPMS)

            Case TIMS.cst_W12_訓練計畫場地資料表, TIMS.cst_G12_訓練計畫場地資料表
                '"12" '訓練計畫場地資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_006?ID=277
                'PCSVALUE
                Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)

                rPMS.Clear()
                rPMS.Add("PCSValue", Hid_PCS.Value)
                rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("PLANID", $"{drOB("PLANID")}")
                'rPMS.Add("RID", $"{drOB("RID")}")
                Call RPT_SD_14_006(rPMS)

            Case TIMS.cst_W13_教學環境資料表, TIMS.cst_G13_教學環境資料表
                '"13" '教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'PCSVALUE
                Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)
                Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))

                rPMS.Clear()
                rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("selsqlstr", selsqlstr)
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                'W13_教學環境資料表
                Call RPT_SD_14_014(rPMS)

            Case TIMS.cst_W13_1_混成課程教學環境資料表, TIMS.cst_G13_1_混成課程教學環境資料表
                '"13B" '混成課程教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'PCSVALUE
                Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value)
                Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))

                rPMS.Clear()
                rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("selsqlstr", selsqlstr)
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                'cst_W13_1_混成課程教學環境資料表
                Call RPT_SD_14_014R(rPMS)

            Case TIMS.cst_G19_送件檢核表, TIMS.cst_W18_送件檢核表
                '送件檢核表
                Dim drRR As DataRow = TIMS.Get_RID_DR($"{drOB("RID")}", objconn)
                If drRR Is Nothing Then
                    Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
                    Return
                End If

                rPMS.Clear()
                rPMS.Add("ORGKINDGW", drKB("ORGKINDGW"))
                rPMS.Add("BCID", drOB("BCID"))
                rPMS.Add("TPLANID", sm.UserInfo.TPlanID)
                rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("DISTID", drOB("DISTID"))
                rPMS.Add("RID", drOB("RID"))
                Call RPT_SD_14_026_23(rPMS)
        End Select

    End Sub

    ''' <summary>列印報表 cst_W13_教學環境資料表 '"13" '教學環境資料表</summary>
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
        Dim sMyValue As String = ""
        sfilename1 = cst_printFN1
        sMyValue &= String.Concat("&Years=", vYEARS_ROC)
        sMyValue &= "&selsqlstr=" & selsqlstr
        sMyValue &= "&TPlanID=" & vTPlanID
        sMyValue &= "&SYears=" & vYEARS
        sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfilename1, sMyValue)
    End Sub

    ''' <summary>列印報表 cst_W13_1_混成課程教學環境資料表 '"13B" '混成課程教學環境資料表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_014R(ByRef rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_014R" '0:未轉班' 1:已轉班
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

    ''' <summary>教學環境資料表</summary>
    ''' <param name="rPMS4"></param>
    ''' <returns></returns>
    Private Function GET_RPTURL_SD_14_014(ByRef rPMS4 As Hashtable) As String
        Const cst_printFN1 As String = "SD_14_014" '0:未轉班' 1:已轉班
        Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
        Dim TSTPRINT As String = If(sPrint_Test = "Y", "2", "1") '測試區2／'正式區1 

        'Const cst_printFN2 As String = "SD_14_014_1" '2:變更待審
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS4, "YEARS")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim selsqlstr As String = TIMS.GetMyValue2(rPMS4, "selsqlstr") 'vPCS
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS4, "TPlanID")

        Dim sfilename1 As String = "" 'cst_printFN1
        Dim sMyValue As String = ""
        sfilename1 = cst_printFN1
        sMyValue &= String.Concat("&Years=", vYEARS_ROC)
        sMyValue &= "&selsqlstr=" & selsqlstr
        sMyValue &= "&TPlanID=" & vTPlanID
        sMyValue &= "&SYears=" & vYEARS
        sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
        Return ReportQuery.GetReportUrl2(Me, sfilename1, sMyValue)
    End Function

    ''' <summary>混成課程教學環境資料表</summary>
    ''' <param name="rPMS4"></param>
    ''' <returns></returns>
    Private Function GET_RPTURL_SD_14_014R(ByRef rPMS4 As Hashtable) As String
        Const cst_printFN1 As String = "SD_14_014R" '0:未轉班' 1:已轉班
        Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
        Dim TSTPRINT As String = If(sPrint_Test = "Y", "2", "1") '測試區2／'正式區1 

        'Const cst_printFN2 As String = "SD_14_014_1" '2:變更待審
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS4, "YEARS")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim selsqlstr As String = TIMS.GetMyValue2(rPMS4, "selsqlstr") 'vPCS
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS4, "TPlanID")

        Dim sfilename1 As String = "" 'cst_printFN1
        Dim sMyValue As String = ""
        sfilename1 = cst_printFN1
        sMyValue &= String.Concat("&Years=", vYEARS_ROC)
        sMyValue &= "&selsqlstr=" & selsqlstr
        sMyValue &= "&TPlanID=" & vTPlanID
        sMyValue &= "&SYears=" & vYEARS
        sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
        Return ReportQuery.GetReportUrl2(Me, sfilename1, sMyValue)
    End Function

    ''' <summary>訓練計畫場地資料表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_006(ByRef rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_006_1" 'SD_14_006_1.jrxml (未轉班、已轉班)
        'Const cst_printFN2 As String = "SD_14_006_2" 'SD_14_006_2.jrxml (變更待審)

        Dim sFilename1 As String = cst_printFN1
        Dim sValue1 As String = ""

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim vPCSValue As String = TIMS.GetMyValue2(rPMS, "PCSValue")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        'Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")

        sFilename1 = cst_printFN1
        sValue1 = ""
        sValue1 &= String.Concat("&Years=", vYEARS_ROC)
        sValue1 &= String.Concat("&PCSValue=", vPCSValue)
        sValue1 &= String.Concat("&PlanID=", vPLANID)
        'sValue1 &= String.Concat("&RID=", vRID) 
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sFilename1, sValue1)
    End Sub

    '師資助教基本資料表 '"10" '師資／助教基本資料表
    Private Sub RPT_SD_14_004(ByRef rPMS As Hashtable)
        Const cst_printFN2 As String = "SD_14_004B" '產投 師資基本資料表

        Dim TechIDStr As String = TIMS.GetMyValue2(rPMS, "TechID")
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "Years")
        Dim vTitle As String = TIMS.GetMyValue2(rPMS, "Title")

        'Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim MyValue As String = ""
        MyValue = "kjk=kjk"
        MyValue &= String.Concat("&TechID=", TechIDStr)
        MyValue &= String.Concat("&Years=", vYEARS) '"&Years=" , Years.Value
        MyValue &= String.Concat("&Title=", vTitle) '"&Title=" & sTitle
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
    End Sub

    '訓練計畫師資／助教名冊
    Private Sub RPT_SD_14_007_1(ByRef rPMS As Hashtable)
        'SD_14_007*.jrxml/0:未轉班/1:已轉班/2:變更待審
        'Const cst_reportFN0 As String = "SD_14_007"'1:已轉班
        Const cst_reportFN1 As String = "SD_14_007_1" '0:未轉班
        'Const cst_reportFN2 As String = "SD_14_007_2"'2:變更待審

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vTitle As String = TIMS.GetMyValue2(rPMS, "Title")
        Dim selsqlstr As String = TIMS.GetMyValue2(rPMS, "selsqlstr")
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)

        '28:產業人才投資方案
        Dim myvalue As String = ""
        TIMS.SetMyValue(myvalue, "Years", vYEARS_ROC) 'sm.UserInfo.Years - 1911
        TIMS.SetMyValue(myvalue, "Title", vTitle)
        TIMS.SetMyValue(myvalue, "selsqlstr", selsqlstr)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN1, myvalue)
    End Sub

    Private Sub RPT_SD_14_002(ByRef rPMS As Hashtable)
        'view-source:https://localhost:44383/SD/14/SD_14_002?ID=273
        Dim strSearch As String = ""
        Session("SCH_SD_14_002") = ""
        TIMS.SetMyValue(strSearch, "CALLFUNC", "TC11001")
        TIMS.SetMyValue(strSearch, "center", TIMS.GetMyValue2(rPMS, "ORGNAME"))
        TIMS.SetMyValue(strSearch, "RIDValue", TIMS.GetMyValue2(rPMS, "RID"))
        '班級狀態 0:未轉班
        TIMS.SetMyValue(strSearch, "Radio1", "0")
        TIMS.SetMyValue(strSearch, "AppStage", TIMS.GetMyValue2(rPMS, "APPSTAGE"))

        Session("SCH_SD_14_002") = strSearch
        Dim url1 As String = String.Concat("../../SD/14/SD_14_002.aspx?ID=", TIMS.Get_MRqID(Me))
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '訓練機構管理>表單列印>訓練計畫總表 SD_14_003_2019G
    Private Sub RPT_SD_14_003_19(ByRef rPMS As Hashtable)
        Const cst_print_19W As String = "SD_14_003_19W" 'PP
        Const cst_print_19G As String = "SD_14_003_19G" 'PP

        Dim vPCSVALUE As String = TIMS.GetMyValue2(rPMS, "PCSVALUE")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW") '$"{drKB("ORGKINDGW")}"
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(TIMS.GetMyValue2(rPMS, "YEARS"))
        'Dim vRSID As String = TIMS.GetMyValue2(rPMS, "RSID")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        'Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")

        Dim vsFileName1 As String = If(vORGKINDGW = "G", cst_print_19G, cst_print_19W)
        'SD_14_003_4_2009('未轉 'PP)
        Dim MyValue As String = ""
        '班級狀態 0:未轉班 
        MyValue &= "&PCSVALUE=" & vPCSVALUE
        MyValue &= "&PlanID=" & vPLANID
        MyValue &= "&Years=" & vYEARS_ROC
        MyValue &= "&AppStage=" & vAPPSTAGE
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vsFileName1, MyValue)
    End Sub

    '訓練單位基本資料表 / 首頁>>訓練機構管理>>表單列印>>訓練單位基本資料表
    Private Sub RPT_SD_14_001_18(ByRef rPMS As Hashtable)
        'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_001?ID=272
        Const cst_printFN2g4 As String = "SD_14_001_18G" '2018-AppStage
        Const cst_printFN2w4 As String = "SD_14_001_18W" '2018-AppStage

        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW") '$"{drKB("ORGKINDGW")}"
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(TIMS.GetMyValue2(rPMS, "YEARS"))
        Dim vRSID As String = TIMS.GetMyValue2(rPMS, "RSID")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")

        '產業人才投資計畫/提升勞工自主學習計畫
        Dim vsFileName1 As String = If(vORGKINDGW = "G", cst_printFN2g4, cst_printFN2w4)
        Dim prtstr As String = ""
        prtstr = ""
        prtstr &= String.Concat("&Years=", vYEARS_ROC)
        prtstr &= String.Concat("&RSID=", vRSID) '"&RSID=" & SelectValue.Value
        prtstr &= String.Concat("&planid=", vPLANID) '"&planid=" & TIMS.ClearSQM(hid_planid.Value)
        prtstr &= String.Concat("&rid=", vRID) '"&rid=" & RIDvalue1
        prtstr &= String.Concat("&AppStage=", vAPPSTAGE)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vsFileName1, prtstr)
    End Sub

    '辦理本計畫訓練課程之專職人員名冊
    Sub RPT_SD_14_018_2(ByRef rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_018_2"
        '首頁>>訓練機構管理>>表單列印>>訓練計畫專職/工作人員名冊
        'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_018_2&Years=2023&RID=B7370&FSQ1=01&SORT1=Y&UserID=snoopy
        'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_018_2&Years=2023&RID=B7169&FSQ1=02&SORT1=Y&UserID=snoopy
        'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_018_2&Years=2023&RID=B7169&FSQ1=02&SORT1=Y&UserID=snoopy

        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "Years", TIMS.GetMyValue2(rPMS, "Years"))
        TIMS.SetMyValue(MyValue, "RID", TIMS.GetMyValue2(rPMS, "RID"))
        TIMS.SetMyValue(MyValue, "FSQ1", TIMS.GetMyValue2(rPMS, "FSQ1"))
        TIMS.SetMyValue(MyValue, "SORT1", "Y")
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    ''' <summary>送件檢核表</summary>
    ''' <param name="rPMS"></param>
    Sub RPT_SD_14_026_23(ByRef rPMS As Hashtable)
        'SD_14_026_23G
        'SAVE_ORG_BIDCASEFL
        Const cst_print_23G As String = "SD_14_026_23G"
        Const cst_print_23W As String = "SD_14_026_23W"
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vsFileName1 As String = If(vORGKINDGW = "G", cst_print_23G, cst_print_23W)
        Dim MyValue As String = ""
        MyValue &= String.Concat("&BCID=", TIMS.GetMyValue2(rPMS, "BCID"))
        MyValue &= String.Concat("&TPLANID=", TIMS.GetMyValue2(rPMS, "TPLANID"))
        MyValue &= String.Concat("&YEARS=", TIMS.GetMyValue2(rPMS, "YEARS"))
        MyValue &= String.Concat("&DISTID=", TIMS.GetMyValue2(rPMS, "DISTID"))
        MyValue &= String.Concat("&RID=", TIMS.GetMyValue2(rPMS, "RID"))
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vsFileName1, MyValue)
    End Sub

#End Region

#Region "共用"
    ''' <summary>共用取號</summary>
    ''' <param name="oConn"></param>
    ''' <param name="rParms"></param>
    ''' <returns></returns>
    Public Shared Function GET_BCASENO_54NN(ByRef oConn As SqlConnection, ByRef rParms As Hashtable) As String
        Dim v_YEARS_ROC As String = TIMS.GET_YEARS_ROC(Now.Year)
        Dim v_NMD4 As String = Now.ToString("MMdd") 'TIMS.GET_YEARS_ROC()

        'Dim vDISTID As String = TIMS.GetMyValue2(rParms, "DISTID")
        Dim v_RID As String = TIMS.GetMyValue2(rParms, "RID")
        Dim vRID1 As String = If(v_RID.Length >= 1, Left(v_RID, 1), v_RID)
        Dim v_CASENO_T1 As String = String.Concat(v_YEARS_ROC, vRID1, "0", v_NMD4)
        'ROC_YEAR +RID +"0" +MMdd
        '3+1+1+4=9 

        'SELECT TOP 11 BCASENO,SUBSTRING(BCASENO,1,9) T1 
        ',SUBSTRING(BCASENO,10,9) T2,cast(SUBSTRING(BCASENO,10,9) as int) iCASENO 
        'FROM ORG_BIDCASE 

        Dim sParms As New Hashtable From {
            {"BCASENOT1", v_CASENO_T1}
        }
        '112B010270003' BCASENO_MAX_SEQ
        Dim sql As String = "SELECT MAX(CAST(SUBSTRING(BCASENO,10,9) AS INT)) BCASENO_MAX_SEQ FROM ORG_BIDCASE WITH(NOLOCK) WHERE SUBSTRING(BCASENO,1,9)=SUBSTRING(@BCASENOT1,1,9)"
        Dim drB1 As DataRow = DbAccess.GetOneRow(sql, oConn, sParms)
        Dim fg_NODATA As Boolean = (drB1 Is Nothing OrElse Convert.ToString(drB1("BCASENO_MAX_SEQ")) = "")
        Dim iBCASENO_LAST_SEQ As Integer = If(fg_NODATA, 1, TIMS.CINT1(drB1("BCASENO_MAX_SEQ")) + 1)
        Return String.Concat(v_CASENO_T1, TIMS.AddZero(iBCASENO_LAST_SEQ.ToString(), 4))
    End Function
#End Region

#Region "Private1"

    'GET ORG_BIDCASE
    'Function GET_ORG_BIDCASE(ByVal vRID As String, ByVal vBCID As String, ByVal vBCASENO As String) As DataRow
    '    Dim drOB As DataRow = Nothing
    '    vRID = TIMS.ClearSQM(vRID)
    '    vBCID = TIMS.ClearSQM(vBCID)
    '    vBCASENO = TIMS.ClearSQM(vBCASENO)
    '    If vRID = "" OrElse vBCID = "" OrElse vBCASENO = "" Then Return drOB

    '    Dim rParms As New Hashtable
    '    rParms.Add("RID", vRID)
    '    rParms.Add("BCID", TIMS.CINT1(vBCID))
    '    rParms.Add("BCASENO", vBCASENO)
    '    Dim rSql As String = "SELECT * FROM ORG_BIDCASE WHERE RID=@RID AND BCID=@BCID AND BCASENO=@BCASENO" 'Dim drOB As DataRow
    '    drOB = DbAccess.GetOneRow(rSql, objconn, rParms)
    '    Return drOB
    'End Function

    '取得 KEY_BIDCASE DataRow
    'Private Function GET_KEY_BIDCASE(ByVal vKBSID As String, ByVal vORGKINDGW As String) As DataRow
    '    Dim drKB As DataRow = Nothing
    '    If vKBSID = "" OrElse vORGKINDGW = "" Then Return drKB
    '    Dim rParms As New Hashtable
    '    rParms.Add("KBSID", vKBSID)
    '    rParms.Add("ORGKINDGW", vORGKINDGW)
    '    Dim rSql As String = ""
    '    'rSql &= " SELECT a.KBSID,a.KBID,a.KBNAME,a.KBDESC1,a.MUSTFILL"
    '    'rSql &= " ,a.ORGKINDGW,a.PARENT,a.KSORT,a.USELATESTVER,a.DOWNLOADRPT"
    '    'rSql &= " ,a.RPTNAME,a.SENTBATVER,a.UPLOADFL1,a.USEMEMO1"
    '    'rSql &= " FROM KEY_BIDCASE a WHERE a.KBSID=@KBSID AND a.ORGKINDGW=@ORGKINDGW"
    '    rSql &= " SELECT * FROM KEY_BIDCASE WHERE KBSID=@KBSID AND ORGKINDGW=@ORGKINDGW"

    '    Dim dtKB As DataTable = DbAccess.GetDataTable(rSql, objconn, rParms)

    '    If dtKB Is Nothing OrElse dtKB.Rows.Count = 0 Then Return drKB

    '    drKB = dtKB.Rows(0)
    '    Return drKB
    'End Function

    ''' <summary>檢核 ORG_BIDCASEFL - 正確為true ／異常為false</summary>
    ''' <param name="rPMS"></param>
    ''' <param name="vKBSID"></param>
    ''' <returns></returns>
    Private Function CHK_ORG_BIDCASEFL(rPMS As Hashtable, ByVal vKBSID As String) As Boolean
        '(外部參數)
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        If vBCID = "" OrElse vKBSID = "" OrElse Not fg_CANSAVE Then Return False '(異常)

        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, vKBSID, vORGKINDGW)
        If drKB Is Nothing Then Return False '(異常)

        Dim rParms As New Hashtable From {
            {"BCID", TIMS.CINT1(vBCID)},
            {"ORGKINDGW", vORGKINDGW},
            {"KBSID", vKBSID}
        }
        Dim rsSql As String = ""
        rsSql &= " SELECT a.BCFID,a.YEARS,a.APPSTAGE,a.RID,a.BCID,a.KBSID,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        rsSql &= " ,kb.KBID,concat(kb.KBID,'.',kb.KBNAME) KBNAME" & vbCrLf
        'rsSql &= " ,CONCAT(a.YEARS,'/',rr.PLANID,'/',a.RID,'/',ob.BCASENO,'/') PATH1" & vbCrLf
        rsSql &= " ,a.WAIVED,a.SRCFILENAME1,a.FILENAME1,a.FILENAME1 OKFLAG" & vbCrLf
        rsSql &= " FROM ORG_BIDCASEFL a" & vbCrLf
        rsSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=a.KBSID" & vbCrLf
        rsSql &= " JOIN ORG_BIDCASE ob on ob.BCID=a.BCID" & vbCrLf
        rsSql &= " JOIN AUTH_RELSHIP rr on rr.RID=a.RID" & vbCrLf
        rsSql &= " WHERE a.BCID=@BCID AND kb.ORGKINDGW=@ORGKINDGW AND a.KBSID=@KBSID" & vbCrLf
        '(若有多筆排序)rsSql &= " ORDER BY kb.KBID,a.KBSID,a.BCFID" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, objconn, rParms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return False '(異常)

        '(取得)
        '取得KBID代號／非流水號／KBID序號
        'Dim vKBID As String = $"{drKB("KBID")}"
        'USELATESTVER: 以最近一次版本送件
        'Dim vUSELATESTVER As String = $"{drKB("USELATESTVER")}"
        'DOWNLOADRPT '可下載報表
        'Dim vDOWNLOADRPT As String = $"{drKB("DOWNLOADRPT")}"
        '必填資訊/ 免附文件(必填就不顯示)
        'Dim vMUSTFILL As String = $"{drKB("MUSTFILL")}"
        'MUSTFILL 必填資訊／WAIVED:免附文件(必填就不顯示)
        'If (vMUSTFILL <> "Y" AndAlso CHKB_WAIVED.Checked) Then Return True
        Return True
    End Function

    ''' <summary>'切換項目(預設)KEY_BIDCASE</summary>
    ''' <param name="vKBSID"></param>
    ''' <param name="vORGKINDGW"></param>
    Sub SHOW_BIDCASE_KBSID(ByVal vKBSID As String, ByVal vORGKINDGW As String)
        'Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Dim v_ddlSwitchTo As String = TIMS.GetListValue(ddlSwitchTo)
        If (v_ddlSwitchTo <> vKBSID) Then
            Hid_KBSID.Value = vKBSID
            Common.SetListItem(ddlSwitchTo, vKBSID)
        End If

        'tb_PreFileUp.Visible = False
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, vKBSID, vORGKINDGW)
        If drKB Is Nothing Then Return
        'tb_PreFileUp.Visible = True

        '(取得)KBID代號／非流水號
        Dim vKBID As String = $"{drKB("KBID")}"
        '取得文字說明
        Dim vKBDESC1 As String = $"{drKB("KBDESC1")}"
        Dim vNOTKBDESC1 As String = $"{drKB("NOTKBDESC1")}" '(Y:不使用KBDESC)
        Dim vNOTFLDESC1 As String = $"{drKB("NOTFLDESC1")}" '(Y:不使用FLDESC1)
        '必填資訊／免附文件(必填就不顯示)
        Dim vMUSTFILL As String = $"{drKB("MUSTFILL")}"
        'USELATESTVER : 以最近一次版本送件
        Dim vUSELATESTVER As String = $"{drKB("USELATESTVER")}"
        'DOWNLOADRPT '可下載報表
        Dim vDOWNLOADRPT As String = $"{drKB("DOWNLOADRPT")}"
        '(報表名稱)
        Dim vRPTNAME As String = $"{drKB("RPTNAME")}"
        '以目前版本批次送出:SENTBATVER
        Dim vSENTBATVER As String = $"{drKB("SENTBATVER")}"
        '以目前版本送出: SENDCURRVER
        Dim vSENDCURRVER As String = $"{drKB("SENDCURRVER")}"
        '檔案上傳:UPLOADFL1
        Dim vUPLOADFL1 As String = $"{drKB("UPLOADFL1")}"
        '備註說明:USEMEMO1
        Dim vUSEMEMO1 As String = $"{drKB("USEMEMO1")}"
        '訓練班別計畫表'DataGrid08
        'Dim vDataGrid08 As String = $"{drKB("DataGrid08")}"
        '檔案格式說明
        labFILEDESC1.Text = cst_FileDescMsg_7_10M

        Dim drFL As DataRow = TIMS.GET_ORG_BIDCASEFL(objconn, TIMS.CINT1(Hid_BCID.Value), TIMS.CINT1(vKBSID))
        Hid_RTUREASON.Value = If(drFL IsNot Nothing, $"{drFL("RTUREASON")}", "")

        Dim str_rtn_checkFile1 As String = String.Concat("return checkFile1(", cst_PostedFile_MAX_SIZE_10M, ");")
        '訓練班別計畫表'DataGrid08
        tr_DataGrid08.Visible = If($"{drKB("DataGrid08")}" = "Y", True, False)
        tr_DataGrid10.Visible = If($"{drKB("DataGrid10")}" = "Y", True, False)
        tr_DataGrid11.Visible = If($"{drKB("DataGrid11")}" = "Y", True, False)
        tr_DataGrid13.Visible = If($"{drKB("DataGrid13")}" = "Y", True, False)
        tr_DataGrid13B.Visible = If($"{drKB("DataGrid13B")}" = "Y", True, False)
        tr_DataGrid14.Visible = If($"{drKB("DataGrid14")}" = "Y", True, False)
        If tr_DataGrid08.Visible Then
            '訓練班別計畫表'DataGrid08
            Call SHOW_DATAGRID_08()
        ElseIf tr_DataGrid10.Visible Then
            '師資／助教基本資料表
            Call SHOW_DATAGRID_10()
        ElseIf tr_DataGrid11.Visible Then
            '各授課師資學／經歷證書影本
            labFILEDESC1.Text = cst_FileDescMsg_7_15M
            Call SHOW_DATAGRID_11()
            str_rtn_checkFile1 = String.Concat("return checkFile1(", cst_PostedFile_MAX_SIZE_15M, ");")
        ElseIf tr_DataGrid13.Visible Then
            '教學環境資料表
            Call SHOW_DATAGRID_13()
        ElseIf tr_DataGrid13B.Visible Then
            '混成課程教學環境資料表
            Call SHOW_DATAGRID_13B()
        ElseIf tr_DataGrid14.Visible Then
            'iCap課程原始申請資料
            Call SHOW_DATAGRID_14()
        End If

        '取得文字說明
        LiteralSwitchTo.Text = If(vKBSID <> "", TIMS.HtmlDecode1(vKBDESC1), "(無)")
        tr_LiteralSwitchTo.Visible = If(vNOTKBDESC1 = "Y", False, True) '(Y:不使用KBDESC)
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
        ' return checkFile1(sizeLimit);
        But1.Attributes.Remove("onclick")
        If vUPLOADFL1 = "Y" AndAlso str_rtn_checkFile1 <> "" Then But1.Attributes.Add("onclick", str_rtn_checkFile1)

        '取得KBID代號／非流水號
        Hid_KBID.Value = vKBID 'GET_KBID(vKBSID, vORGKINDGW) GET_KBDESC1(vKBSID, vORGKINDGW)
        LabSwitchTo.Text = If(vKBSID <> "", TIMS.GetListText(ddlSwitchTo), "")
        'USELATESTVER : 以最近一次版本送件
        tr_USELATESTVER.Visible = If(vUSELATESTVER = "Y", True, False)
        'MUSTFILL 必填資訊／WAIVED:免附文件(必填就不顯示)
        tr_WAIVED.Visible = If(vMUSTFILL = "Y", False, True)
        '備註說明:USEMEMO1
        tr_USEMEMO1.Visible = If(vUSEMEMO1 = "Y", True, False)
        '預設值-免附文件
        CHKB_WAIVED.Checked = False '(預設值不填寫)
        '預設值-(上傳檔案)
        Hid_BCFID.Value = ""
        '預設值-備註說明
        txtMEMO1.Text = ""

        '(Enabled) begin
        File1.Disabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, True, False)
        But1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(File1, If(But1.Enabled, "", cst_tpmsg_enb2), True)
        TIMS.Tooltip(But1, If(But1.Enabled, "", cst_tpmsg_enb2), True)

        BTN_SENTBATVER.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SENTBATVER, If(BTN_SENTBATVER.Enabled, "", cst_tpmsg_enb3), True)
        BTN_SENDCURRVER.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SENDCURRVER, If(BTN_SENDCURRVER.Enabled, "", cst_tpmsg_enb3), True)
        bt_latestSend1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(bt_latestSend1, If(bt_latestSend1.Enabled, "", cst_tpmsg_enb3), True)
        CHKB_WAIVED.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(CHKB_WAIVED, If(CHKB_WAIVED.Enabled, "", cst_tpmsg_enb3), True)
        '備註說明
        txtMEMO1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(txtMEMO1, If(txtMEMO1.Enabled, "", cst_tpmsg_enb3), True)
        '(Enabled) close

        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        If drOB Is Nothing Then Return
        'Dim drFL As DataRow = TIMS.GET_ORG_BIDCASEFL(objconn, TIMS.CINT1(Hid_BCID.Value), TIMS.CINT1(vKBSID))
        If drFL Is Nothing Then Return

        Hid_BCFID.Value = Convert.ToString(drFL("BCFID"))
        '免附文件
        CHKB_WAIVED.Checked = If(Convert.ToString(drFL("WAIVED")) = "Y", True, False)

        txtMEMO1.Text = TIMS.ClearSQM(drFL("MEMO1"))

        '修改狀態，且為退件修正
        Dim fg_UPDATE_BISTATUS_R As Boolean = (Session(cst_ss_RqProcessType) = cst_DG1CMDNM_EDIT1 AndAlso Hid_BCFID.Value <> "" AndAlso $"{drOB("BISTATUS")}" = "R")
        If fg_UPDATE_BISTATUS_R Then
            '沒有退件原因就不能改
            Dim fg_LOCK_INPUT As Boolean = (Convert.ToString(drFL("RTUREASON")) = "")
            '(Enabled) begin
            File1.Disabled = If(fg_LOCK_INPUT, True, False)
            But1.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(File1, If(But1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            TIMS.Tooltip(But1, If(But1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)

            BTN_SENTBATVER.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(BTN_SENTBATVER, If(BTN_SENTBATVER.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            BTN_SENDCURRVER.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(BTN_SENDCURRVER, If(BTN_SENDCURRVER.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            bt_latestSend1.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(bt_latestSend1, If(bt_latestSend1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            CHKB_WAIVED.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(CHKB_WAIVED, If(CHKB_WAIVED.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            '備註說明
            txtMEMO1.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(txtMEMO1, If(txtMEMO1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            '(Enabled) close
            '儲存(暫存)
            BTN_SAVETMP1.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            '儲存後進下一步
            BTN_SAVENEXT1.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
        End If
    End Sub

    ''' <summary>iCap課程原始申請資料</summary>
    Private Sub SHOW_DATAGRID_14()
        iDG14_ROWS = 0
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vBCID As String = $"{drOB("BCID")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vRID As String = $"{drOB("RID")}"

        Dim sParms1 As New Hashtable From {
            {"BCID", vBCID}, 'ORG_BIDCASEPI/ORG_BIDCASEFL_PI3
            {"RID", vRID}
        }
        'sParms1.Add("AppStage", vAPPSTAGE)

        Dim sSql As String = ""
        sSql &= "WITH WPI1 AS ( SELECT a.BCID,a.BCPID,a.PLANID,a.COMIDNO,a.SEQNO,pp.PSNO28" & vbCrLf
        sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
        sSql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID and pp.COMIDNO=a.COMIDNO and pp.SEQNO=a.SEQNO" & vbCrLf
        sSql &= " WHERE a.BCID=@BCID )" & vbCrLf

        sSql &= ",WFPI3 AS ( SELECT b.BCFP3ID,a.BCID,a.BCPID" & vbCrLf
        sSql &= " ,a.PLANID,a.COMIDNO,a.SEQNO,a.PSNO28" & vbCrLf
        sSql &= " ,b.SRCFILENAME1,b.FILENAME1,b.WAIVED,kb.KBID,kb.KBSID,kb.ORGKINDGW" & vbCrLf
        sSql &= " FROM WPI1 a" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL_PI3 b on b.BCPID=a.BCPID and b.BCID=a.BCID" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID )" & vbCrLf

        sSql &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID" & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sSql &= " ,b.OrgName ,a.RID" & vbCrLf
        sSql &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf
        sSql &= " ,p3.WAIVED,p3.SRCFILENAME1,p3.FILENAME1,p3.FILENAME1 OKFLAG" & vbCrLf
        sSql &= " ,p3.BCFP3ID,p3.KBID,p3.KBSID,p3.ORGKINDGW" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID = b.RID" & vbCrLf
        sSql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = a.PlanID" & vbCrLf
        sSql &= " JOIN WPI1 p1 ON p1.PLANID=a.PLANID AND p1.COMIDNO=a.COMIDNO AND p1.SEQNO=a.SEQNO" & vbCrLf
        sSql &= " LEFT JOIN WFPI3 p3 ON p3.PLANID=a.PLANID AND p3.COMIDNO=a.COMIDNO AND p3.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        sSql &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        '使用登入者業務權限
        sSql &= " AND a.RID=@RID" & vbCrLf
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

        Dim vYEARS As String = $"{drOB("YEARS")}"
        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        'Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        labmsg2.Text = ""
        DataGrid14.Visible = True
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            DataGrid14.Visible = False
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG14_ROWS = dt2.Rows.Count

        DataGrid14.DataSource = dt2
        DataGrid14.DataBind()
    End Sub

    ''' <summary>教學環境資料表</summary>
    Private Sub SHOW_DATAGRID_13()
        iDG13_ROWS = 0
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vBCID As String = $"{drOB("BCID")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vRID As String = $"{drOB("RID")}"

        Dim sParms1 As New Hashtable From {
            {"BCID", vBCID}, 'ORG_BIDCASEPI/ORG_BIDCASEFL_EV
            {"RID", vRID}
        }
        'sParms1.Add("AppStage", vAPPSTAGE)

        Dim sSql As String = ""
        sSql &= "WITH WPI1 AS (" & vbCrLf
        sSql &= " SELECT a.BCID,a.BCPID,a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
        sSql &= " WHERE a.BCID=@BCID )" & vbCrLf

        sSql &= ",WFPI2 AS (" & vbCrLf
        sSql &= " SELECT b.BCFEID,a.BCPID,a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,b.SRCFILENAME1,b.FILENAME1,b.WAIVED" & vbCrLf
        sSql &= " ,kb.KBID,kb.KBSID" & vbCrLf
        sSql &= " FROM WPI1 a" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL_EV b on b.BCPID=a.BCPID AND b.BCID=a.BCID" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID )" & vbCrLf

        sSql &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID " & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql &= " ,CONVERT(varchar, a.STDate, 111) STDATE " & vbCrLf
        sSql &= " ,b.OrgName ,a.RID " & vbCrLf
        sSql &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf 'pp.MSD
        sSql &= " ,p2.WAIVED,p2.SRCFILENAME1,p2.FILENAME1,p2.FILENAME1 OKFLAG" & vbCrLf
        sSql &= " ,p2.BCFEID,p2.KBID,p2.KBSID" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO a " & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID = b.RID " & vbCrLf
        sSql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = a.PlanID " & vbCrLf
        sSql &= " JOIN WPI1 p1 ON p1.PLANID=a.PLANID AND p1.COMIDNO=a.COMIDNO AND p1.SEQNO=a.SEQNO" & vbCrLf
        sSql &= " LEFT JOIN WFPI2 p2 ON p2.PLANID=a.PLANID AND p2.COMIDNO=a.COMIDNO AND p2.SEQNO=a.SEQNO" & vbCrLf

        '0:未轉班,1:已轉班
        sSql &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        '使用登入者業務權限
        sSql &= " AND a.RID =@RID" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPlanID", sm.UserInfo.TPlanID)
            sParms1.Add("Years", sm.UserInfo.Years)
            sSql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            sSql &= " AND ip.Years =@Years" & vbCrLf
        Else
            sParms1.Add("PlanID", sm.UserInfo.PlanID)
            sSql &= " AND ip.PlanID=@PlanID" & vbCrLf
        End If
        sSql &= " ORDER BY a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)

        Dim vYEARS As String = $"{drOB("YEARS")}"
        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        'Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        labmsg2.Text = TIMS.cst_NODATAMsg1
        DataGrid13.Visible = False

        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then Return

        labmsg2.Text = ""
        DataGrid13.Visible = True

        iDG13_ROWS = dt2.Rows.Count

        DataGrid13.DataSource = dt2
        DataGrid13.DataBind()
    End Sub

    ''' <summary>混成課程教學環境資料表／遠距</summary>
    Private Sub SHOW_DATAGRID_13B()
        iDG13B_ROWS = 0
        Dim oDG1 As DataGrid = DataGrid13B
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then 'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then 'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        'Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
        Dim vBCID As String = $"{drOB("BCID")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vRID As String = $"{drOB("RID")}"

        'ORG_BIDCASEPI/ORG_BIDCASEFL_RT 
        Dim sParms1 As New Hashtable From {{"BCID", vBCID}, {"RID", vRID}}
        Dim sSql As String = ""
        sSql &= "WITH WPI1 AS (" & vbCrLf
        sSql &= " SELECT a.BCID,a.BCPID,a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
        sSql &= " WHERE a.BCID=@BCID )" & vbCrLf

        sSql &= ",WFRT2 AS (" & vbCrLf
        sSql &= " SELECT b.BCRTID,a.BCPID" & vbCrLf
        sSql &= " ,a.PLANID,a.COMIDNO,a.SEQNO,b.SRCFILENAME1,b.FILENAME1,b.WAIVED" & vbCrLf
        sSql &= " ,kb.KBID,kb.KBSID" & vbCrLf
        sSql &= " FROM WPI1 a" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL_RT b on b.BCPID=a.BCPID AND b.BCID=a.BCID" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID )" & vbCrLf

        sSql &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID " & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql &= " ,CONVERT(varchar, a.STDate, 111) STDATE " & vbCrLf
        sSql &= " ,b.OrgName ,a.RID " & vbCrLf
        sSql &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf 'pp.MSD
        sSql &= " ,p2.WAIVED,p2.SRCFILENAME1,p2.FILENAME1,p2.FILENAME1 OKFLAG" & vbCrLf
        sSql &= " ,p2.BCRTID,p2.KBID,p2.KBSID" & vbCrLf
        sSql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sSql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID = b.RID" & vbCrLf
        sSql &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = a.PlanID" & vbCrLf
        sSql &= " JOIN dbo.ORG_REMOTER ot ON ot.RMTID=a.RMTID AND a.DISTANCE ='2'" & vbCrLf
        sSql &= " JOIN WPI1 p1 ON p1.PLANID=a.PLANID AND p1.COMIDNO=a.COMIDNO AND p1.SEQNO=a.SEQNO" & vbCrLf
        sSql &= " LEFT JOIN WFRT2 p2 ON p2.PLANID=a.PLANID AND p2.COMIDNO=a.COMIDNO AND p2.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        sSql &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        '使用登入者業務權限
        sSql &= " AND a.RID=@RID" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPlanID", sm.UserInfo.TPlanID)
            sParms1.Add("Years", sm.UserInfo.Years)
            sSql &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            sSql &= " AND ip.Years =@Years" & vbCrLf
        Else
            sParms1.Add("PlanID", sm.UserInfo.PlanID)
            sSql &= " AND ip.PlanID=@PlanID" & vbCrLf
        End If
        sSql &= " ORDER BY a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)

        Dim vYEARS As String = $"{drOB("YEARS")}"
        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        'Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        labmsg2.Text = TIMS.cst_NODATAMsg1
        oDG1.Visible = False

        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then Return

        labmsg2.Text = ""
        oDG1.Visible = True

        iDG13B_ROWS = dt2.Rows.Count

        oDG1.DataSource = dt2
        oDG1.DataBind()
    End Sub


    ''' <summary>授課師資學經歷證書影本</summary>
    Private Sub SHOW_DATAGRID_11()
        iDG11_ROWS = 0
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        'DataGrid11_ItemDataBound
        Dim rParms2 As New Hashtable From {
            {"BCID", $"{drOB("BCID")}"},
            {"RID", $"{drOB("RID")}"}
        }
        'rParms2.Add("AppStage", $"{drOB("AppStage")}")
        Dim sSql2 As String = ""
        sSql2 &= " WITH WT1 AS (SELECT a.TECHID,a.BCFID,a.BCFT2ID,a.PATTERN,a.MEMO1,a.WAIVED" & vbCrLf
        sSql2 &= " ,a.SRCFILENAME1,a.FILENAME1" & vbCrLf
        sSql2 &= " ,ob.YEARS,rr.PLANID,rr.RID,ob.BCASENO,kb.KBID,kb.KBSID" & vbCrLf
        sSql2 &= " FROM VIEW_RIDNAME rr" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL bf ON bf.RID=rr.RID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASE ob ON ob.BCID=bf.BCID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL_TT2 a ON a.BCFID=bf.BCFID" & vbCrLf
        sSql2 &= " JOIN KEY_BIDCASE kb on kb.KBSID=bf.KBSID" & vbCrLf
        sSql2 &= " WHERE ob.BCID=@BCID AND rr.RID=@RID)" & vbCrLf

        sSql2 &= " ,WT2 AS (SELECT DISTINCT P2.TECHID" & vbCrLf
        sSql2 &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
        sSql2 &= " JOIN dbo.ORG_BIDCASEPI bp ON bp.PlanID=P1.PlanID AND bp.ComIDNO=P1.ComIDNO AND bp.SeqNo=P1.SeqNo" & vbCrLf
        sSql2 &= " JOIN dbo.V_PLAN_TEACHER1 P2 ON P1.PlanID=P2.PlanID AND P1.ComIDNO=P2.ComIDNO AND P1.SeqNo=P2.SeqNo" & vbCrLf
        sSql2 &= " WHERE P1.IsApprPaper='Y' AND P1.AppliedResult IS NULL AND P1.RESULTBUTTON IS NULL" & vbCrLf
        sSql2 &= " AND bp.BCID=@BCID AND P1.RID=@RID)" & vbCrLf
        'sSql2 &= " WHERE P1.TransFlag='N' AND P1.IsApprPaper='Y' AND P1.AppliedResult IS NULL AND P1.RESULTBUTTON IS NULL" & vbCrLf
        'sSql2 &= " AND P1.RID=@RID AND P1.AppStage=@AppStage)" & vbCrLf

        sSql2 &= " SELECT a.TechID,a.RID,a.TEACHCNAME,a.TEACHENAME,a.TEACHERID" & vbCrLf
        sSql2 &= " ,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql2 &= " ,a.KINDENGAGE,case a.KINDENGAGE when '1' then '內聘(專任)' else '外聘(兼任)' end KINDENGAGE_N" & vbCrLf
        sSql2 &= " ,a.WORKSTATUS,case a.WORKSTATUS when '1' then '是' else '否' end WORKSTATUS_N" & vbCrLf
        sSql2 &= " ,(SELECT x.KINDNAME FROM ID_KINDOFTEACHER x WHERE x.KINDID=a.KINDID) KINDNAME" & vbCrLf
        sSql2 &= " ,bt.BCFID,bt.BCFT2ID,bt.PATTERN,bt.MEMO1,bt.WAIVED" & vbCrLf
        'sSql2 &= " ,case when bt.TECHID IS NOT NULL THEN CONCAT(bt.YEARS,'/',bt.PLANID,'/',bt.RID,'/',bt.BCASENO,'/',bt.KBSID,'/') END PATH1" & vbCrLf
        sSql2 &= " ,bt.SRCFILENAME1,bt.FILENAME1,bt.FILENAME1 OKFLAG" & vbCrLf
        sSql2 &= " ,bt.KBID,bt.KBSID" & vbCrLf
        sSql2 &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sSql2 &= " JOIN WT2 t2 ON t2.TECHID=a.TECHID" & vbCrLf
        sSql2 &= " JOIN AUTH_RELSHIP b ON b.RID=a.RID" & vbCrLf
        sSql2 &= " LEFT JOIN WT1 bt on bt.TECHID=a.TECHID AND bt.RID=a.RID" & vbCrLf
        sSql2 &= " WHERE a.RID=@RID" & vbCrLf
        sSql2 &= " ORDER BY a.TEACHERID" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, rParms2)

        labmsg2.Text = ""
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If
        iDG11_ROWS = dt2.Rows.Count

        Dim vYEARS As String = $"{drOB("YEARS")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        With DataGrid11
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    ''' <summary>師資助教基本資料表</summary>
    Private Sub SHOW_DATAGRID_10()
        iDG10_ROWS = 0
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        'DataGrid10_ItemDataBound
        Dim rParms2 As New Hashtable From {
            {"BCID", $"{drOB("BCID")}"},
            {"RID", $"{drOB("RID")}"}
        }
        'rParms2.Add("AppStage", $"{drOB("AppStage")}")
        Dim sSql2 As String = ""
        sSql2 &= " WITH WT1 AS (SELECT a.TECHID,a.BCFID,a.BCFTID,a.PATTERN,a.MEMO1,a.WAIVED" & vbCrLf
        sSql2 &= " ,a.SRCFILENAME1,a.FILENAME1" & vbCrLf
        sSql2 &= " ,ob.YEARS,rr.PLANID,rr.RID,ob.BCASENO,kb.KBID,kb.KBSID" & vbCrLf
        sSql2 &= " FROM VIEW_RIDNAME rr" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL bf ON bf.RID=rr.RID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASE ob ON ob.BCID=bf.BCID" & vbCrLf
        sSql2 &= " JOIN ORG_BIDCASEFL_TT a ON a.BCFID=bf.BCFID" & vbCrLf
        sSql2 &= " JOIN KEY_BIDCASE kb on kb.KBSID=bf.KBSID" & vbCrLf
        sSql2 &= " WHERE ob.BCID=@BCID AND rr.RID=@RID)" & vbCrLf

        sSql2 &= " ,WT2 AS (SELECT DISTINCT P2.TECHID" & vbCrLf
        sSql2 &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
        sSql2 &= " JOIN dbo.ORG_BIDCASEPI bp ON bp.PlanID=P1.PlanID AND bp.ComIDNO=P1.ComIDNO AND bp.SeqNo=P1.SeqNo" & vbCrLf
        sSql2 &= " JOIN dbo.V_PLAN_TEACHER1 P2 ON P1.PlanID=P2.PlanID AND P1.ComIDNO=P2.ComIDNO AND P1.SeqNo=P2.SeqNo" & vbCrLf
        sSql2 &= " WHERE P1.TransFlag='N' AND P1.IsApprPaper='Y' AND P1.AppliedResult IS NULL AND P1.RESULTBUTTON IS NULL" & vbCrLf
        sSql2 &= " AND bp.BCID=@BCID AND P1.RID=@RID )" & vbCrLf

        sSql2 &= " SELECT a.TechID,a.RID,a.TEACHCNAME,a.TEACHENAME,a.TEACHERID" & vbCrLf
        sSql2 &= " ,a.IDNO,dbo.FN_GET_MASK1(a.IDNO) IDNO_MK" & vbCrLf
        sSql2 &= " ,a.KINDENGAGE,case a.KINDENGAGE when '1' then '內聘(專任)' else '外聘(兼任)' end KINDENGAGE_N" & vbCrLf
        sSql2 &= " ,a.WORKSTATUS,case a.WORKSTATUS when '1' then '是' else '否' end WORKSTATUS_N" & vbCrLf
        sSql2 &= " ,(SELECT x.KINDNAME FROM ID_KINDOFTEACHER x WHERE x.KINDID=a.KINDID) KINDNAME" & vbCrLf
        sSql2 &= " ,bt.BCFID,bt.BCFTID,bt.PATTERN,bt.MEMO1,bt.WAIVED" & vbCrLf
        sSql2 &= " ,bt.SRCFILENAME1,bt.FILENAME1,bt.FILENAME1 OKFLAG" & vbCrLf
        sSql2 &= " ,bt.KBID,bt.KBSID" & vbCrLf
        sSql2 &= " FROM TEACH_TEACHERINFO a" & vbCrLf
        sSql2 &= " JOIN WT2 t2 ON t2.TECHID=a.TECHID" & vbCrLf
        sSql2 &= " JOIN AUTH_RELSHIP b ON b.RID=a.RID" & vbCrLf
        sSql2 &= " LEFT JOIN WT1 bt on bt.TECHID=a.TECHID AND bt.RID=a.RID" & vbCrLf
        sSql2 &= " WHERE a.RID=@RID" & vbCrLf
        sSql2 &= " ORDER BY a.TEACHERID" & vbCrLf
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, rParms2)

        labmsg2.Text = ""
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG10_ROWS = dt2.Rows.Count

        Dim vYEARS As String = $"{drOB("YEARS")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        With DataGrid10
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    ''' <summary>訓練班別計畫表</summary>
    Private Sub SHOW_DATAGRID_08()
        iDG08_ROWS = 0
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim vBCID As String = Hid_BCID.Value 'TIMS.ClearSQM(Hid_BCID.Value)
        'Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vRID As String = $"{drOB("RID")}"

        '使用登入者業務權限
        Dim sParms1 As New Hashtable From {{"BCID", vBCID}, {"RID", vRID}}
        Dim sSql1 As String = ""
        sSql1 &= " WITH WBPI AS (SELECT a.BCID,a.BCPID,a.PLANID,a.COMIDNO,a.SEQNO FROM ORG_BIDCASEPI a WHERE a.BCID=@BCID)" & vbCrLf
        sSql1 &= " ,WFPI2 AS (SELECT a.BCPID,a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql1 &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
        sSql1 &= " ,b.BCFPID,b.BCFID,kb.KBID,kb.KBSID" & vbCrLf
        sSql1 &= " FROM WBPI a" & vbCrLf
        sSql1 &= " JOIN ORG_BIDCASEFL_PI b on b.BCPID=a.BCPID" & vbCrLf
        sSql1 &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
        sSql1 &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID)" & vbCrLf

        sSql1 &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql1 &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID" & vbCrLf
        sSql1 &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql1 &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql1 &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sSql1 &= " ,b.OrgName ,a.RID" & vbCrLf
        sSql1 &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf
        sSql1 &= " ,p2.SRCFILENAME1,p2.FILENAME1,p2.FILENAME1 OKFLAG" & vbCrLf
        sSql1 &= " ,p2.BCFPID,p2.BCFID,p2.KBID,p2.KBSID" & vbCrLf
        sSql1 &= " FROM WBPI p1" & vbCrLf
        sSql1 &= " JOIN dbo.PLAN_PLANINFO a ON a.PLANID=p1.PLANID AND a.COMIDNO=p1.COMIDNO AND a.SEQNO=p1.SEQNO" & vbCrLf
        sSql1 &= " JOIN dbo.VIEW_RIDNAME b ON a.RID = b.RID" & vbCrLf
        sSql1 &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = a.PlanID" & vbCrLf
        sSql1 &= " LEFT JOIN WFPI2 p2 ON p2.PLANID=a.PLANID AND p2.COMIDNO=a.COMIDNO AND p2.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        sSql1 &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        sSql1 &= " AND a.RID=@RID" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPlanID", sm.UserInfo.TPlanID)
            sParms1.Add("Years", sm.UserInfo.Years)
            sSql1 &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            sSql1 &= " AND ip.Years=@Years" & vbCrLf
        Else
            sParms1.Add("PlanID", sm.UserInfo.PlanID)
            sSql1 &= " AND ip.PlanID=@PlanID" & vbCrLf
        End If
        sSql1 &= " ORDER BY a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf

        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql1, objconn, sParms1)

        labmsg2.Text = ""
        If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG08_ROWS = dt2.Rows.Count

        Dim vYEARS As String = $"{drOB("YEARS")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        'Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)

        With DataGrid08
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

#Region "NOUSE"
    '設定下拉項目(產投G／自主W)
    'Sub Utl_SET_ddlSWITCHTO_VAL(ByVal vORGKINDGW As String)
    '    'RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
    '    'Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))
    '    If vORGKINDGW = "" Then Return 'rst

    '    Dim sPMS As New Hashtable
    '    sPMS.Add("ORGKINDGW", vORGKINDGW)
    '    Dim sSql As String = "SELECT KBSID,concat(KBID,'.',KBNAME) KBNAME,KSORT FROM KEY_BIDCASE WHERE ORGKINDGW=@ORGKINDGW AND PARENT IS NULL ORDER BY KSORT"
    '    Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    If dt Is Nothing OrElse dt.Rows.Count = 0 Then
    '        Common.MessageBox(Me, "項目資訊有誤(查無資料)，請連絡系統管理者!")
    '        Return
    '    End If
    '    ddlSwitchTo.Items.Clear()
    '    DbAccess.MakeListItem(ddlSwitchTo, sSql, objconn, sPMS)
    '    ddlSwitchTo.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose2, ""))

    '    Dim xsPMS As New Hashtable
    '    xsPMS.Add("ORGKINDGW", vORGKINDGW)
    '    Dim xSql As String = "SELECT MAX(KBID) xKBID FROM KEY_BIDCASE WHERE ORGKINDGW=@ORGKINDGW"
    '    Dim xKBID As String = DbAccess.ExecuteScalar(xSql, objconn, xsPMS)
    '    Hid_LastKBID.Value = xKBID

    '    Dim xsPMS2 As New Hashtable
    '    xsPMS2.Add("ORGKINDGW", vORGKINDGW)
    '    Dim xSql2 As String = "SELECT MIN(KBSID) KBSID FROM KEY_BIDCASE WHERE ORGKINDGW=@ORGKINDGW"
    '    Dim xKBSID As String = DbAccess.ExecuteScalar(xSql2, objconn, xsPMS2)
    '    Hid_FirstKBSID.Value = xKBSID
    'End Sub
#End Region

    ''' <summary>儲存上傳檔案(2)</summary>
    ''' <param name="rPMS"></param>
    Function SAVE_ORG_BIDCASEFL_UPLOAD(ByRef rPMS As Hashtable) As Integer
        Dim iBCFID As Integer = -1

        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        If Not fg_CANSAVE Then Return iBCFID

        '重審 vUploadPath / vBCFID 
        Dim vUploadPath As String = TIMS.GetMyValue2(rPMS, "UploadPath")
        Dim vBCFID As String = TIMS.GetMyValue2(rPMS, "BCFID")
        iBCFID = If(vBCFID <> "" AndAlso TIMS.CINT1(vBCFID) > 0, TIMS.CINT1(vBCFID), -1)

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        'Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")

        Dim vWAIVED As String = TIMS.GetMyValue2(rPMS, "WAIVED")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vPATTERN As String = TIMS.GetMyValue2(rPMS, "PATTERN")
        Dim vMEMO1 As String = TIMS.GetMyValue2(rPMS, "MEMO1")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        '免附文件或上傳檔案
        Dim fg_NG_SAVE As Boolean = (vWAIVED = "" AndAlso (vFILENAME1 = "" AndAlso vSRCFILENAME1 = ""))
        If fg_NG_SAVE Then Return iBCFID
        'WAIVED: 只能是Y/ ""
        'Dim fg_WAIVED_CAN_SAVE As Boolean = (vWAIVED = "" OrElse vWAIVED = "Y")
        'If Not fg_WAIVED_CAN_SAVE Then Return iBCFID

        Dim drFL As DataRow = TIMS.GET_ORG_BIDCASEFL(objconn, TIMS.CINT1(vBCID), TIMS.CINT1(vKBSID))
        If drFL IsNot Nothing Then
            Dim OldvFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
            Dim OldvSRCFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("SRCFILENAME1")), "")
            Dim vRTUREASON As String = If(drFL IsNot Nothing, Convert.ToString(drFL("RTUREASON")), "")

            '(重新儲存) (加入訊息) 'iBCFID > 0 
            Dim vREUPLOADED_MSG As String = ""
            If iBCFID > 0 Then
                If vWAIVED = "Y" Then
                    vFILENAME1 = ""
                    vSRCFILENAME1 = ""
                    If (vRTUREASON <> "" AndAlso vRTUREASON.EndsWith(cst_REUPLOADED_MSG)) Then vRTUREASON = vRTUREASON.Replace(cst_REUPLOADED_MSG, "")
                    vREUPLOADED_MSG = If(vRTUREASON <> "" AndAlso vRTUREASON.IndexOf(cst_txt_免附文件) = -1, String.Concat(vRTUREASON, cst_txt_免附文件), "")
                ElseIf (vFILENAME1 <> "" AndAlso vSRCFILENAME1 <> "") Then
                    If (vRTUREASON <> "" AndAlso vRTUREASON.EndsWith(cst_txt_免附文件)) Then vRTUREASON = vRTUREASON.Replace(cst_txt_免附文件, "")
                    vREUPLOADED_MSG = If(vRTUREASON <> "" AndAlso vRTUREASON.IndexOf(cst_REUPLOADED_MSG) = -1, String.Concat(vRTUREASON, cst_REUPLOADED_MSG), "")
                End If
            End If

            Dim uParms As New Hashtable From {
                {"WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull)}
            }
            If vWAIVED = "Y" OrElse vFILENAME1 <> "" Then uParms.Add("FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull)) 'vFILENAME1)
            If vWAIVED = "Y" OrElse vSRCFILENAME1 <> "" Then uParms.Add("SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull)) 'vSRCFILENAME1)
            If iBCFID > 0 Then uParms.Add("RTUREASON", If(vREUPLOADED_MSG <> "", vREUPLOADED_MSG, If(vRTUREASON <> "", vRTUREASON, Convert.DBNull)))
            uParms.Add("PATTERN", If(vPATTERN <> "", vPATTERN, Convert.DBNull))
            uParms.Add("MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull))
            uParms.Add("MODIFYACCT", vMODIFYACCT)
            uParms.Add("BCFID", TIMS.CINT1(drFL("BCFID")))

            Dim usSql As String = ""
            usSql &= " UPDATE ORG_BIDCASEFL" & vbCrLf
            usSql &= " SET WAIVED=@WAIVED" & vbCrLf
            If vWAIVED = "Y" OrElse vFILENAME1 <> "" Then usSql &= " ,FILENAME1=@FILENAME1" & vbCrLf
            If vWAIVED = "Y" OrElse vSRCFILENAME1 <> "" Then usSql &= " ,SRCFILENAME1=@SRCFILENAME1" & vbCrLf
            If iBCFID > 0 Then usSql &= " ,RTUREASON=@RTUREASON" & vbCrLf
            usSql &= " ,PATTERN=@PATTERN,MEMO1=@MEMO1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE BCFID=@BCFID" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            '刪除舊檔案(放在UPDATA 的後面表示已經處理完成DB了) OldvFILENAME1 <> ""
            Dim fg_CAN_DELETE_OLDFILE As Boolean = False
            If vWAIVED = "" Then
                '附文件置換檔案 / 刪除舊檔案
                fg_CAN_DELETE_OLDFILE = (vUploadPath <> "" AndAlso vFILENAME1 <> "" AndAlso OldvFILENAME1 <> "" AndAlso vFILENAME1 <> OldvFILENAME1)
            ElseIf vWAIVED = "Y" Then
                '免附文件 / 刪除舊檔案
                fg_CAN_DELETE_OLDFILE = (vUploadPath <> "" AndAlso OldvFILENAME1 <> "")
            End If
            If fg_CAN_DELETE_OLDFILE Then
                Try
                    '刪除舊檔案
                    TIMS.MyFileDelete(Server.MapPath(String.Concat(vUploadPath, OldvFILENAME1)))
                Catch ex As Exception
                    TIMS.LOG.Warn(ex.Message, ex)
                    'Common.MessageBox(Me, cst_errMsg_2) 'Common.MessageBox(Me, ex.ToString)
                    Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
                    strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
                    strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("OldvFILENAME1: ", OldvFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("Server.MapPath(String.Concat(vUploadPath, OldvFILENAME1)): ", Server.MapPath(String.Concat(vUploadPath, OldvFILENAME1)), vbCrLf)
                    'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                    'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                End Try
            End If

            iBCFID = TIMS.CINT1(drFL("BCFID"))
        Else
            iBCFID = DbAccess.GetNewId(objconn, "ORG_BIDCASEFL_BCFID_SEQ,ORG_BIDCASEFL,BCFID")
            'k_KBGWID = "KBGID"
            'k_BCFGWID = "BCFGID"
            'iParms.Add("APPSTAGE", Convert.DBNull)
            Dim iParms As New Hashtable From {
                {"BCFID", iBCFID},
                {"YEARS", vYEARS},
                {"RID", vRID},
                {"BCID", TIMS.CINT1(vBCID)},
                {"KBSID", TIMS.CINT1(vKBSID)},
                {"WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull)},
                {"FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull)}, 'vFILENAME1)
                {"SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull)}, 'vSRCFILENAME1)
                {"PATTERN", If(vPATTERN <> "", vPATTERN, Convert.DBNull)}, '
                {"MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull)},
                {"MODIFYACCT", vMODIFYACCT}
            }
            'iParms.Add("MODIFYDATE", MODIFYDATE)
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_BIDCASEFL(BCFID,YEARS, RID,BCID,KBSID,WAIVED"
            isSql &= " ,FILENAME1,SRCFILENAME1,PATTERN,MEMO1,MODIFYACCT,MODIFYDATE)" & vbCrLf
            isSql &= " VALUES (@BCFID,@YEARS, @RID,@BCID,@KBSID,@WAIVED"
            isSql &= " ,@FILENAME1,@SRCFILENAME1,@PATTERN,@MEMO1,@MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        End If

        '暫時儲存／正式儲存-UPDATE ORG_BIDCASE
        Call SAVEDATE1(0)

        Return iBCFID
    End Function

    ''' <summary>檢視目前上傳檔案</summary>
    ''' <param name="rPMS"></param>
    Sub SHOW_BIDCASEFL_DG2(ByRef rPMS As Hashtable)
        labmsg1.Text = ""
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        'objconn 因為有檔案輸出關閉的問題 所以要檢查
        If Not TIMS.OpenDbConn(objconn) OrElse Not fg_CANSAVE Then Return

        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If

        Dim dtFL As DataTable = GET_ORG_BIDCASEFL_TB(objconn, vBCID, vORGKINDGW)

        labmsg1.Text = If(dtFL Is Nothing OrElse dtFL.Rows.Count = 0, "(查無文件項目)", "")

        Hid_BISTATUS.Value = $"{drOB("BISTATUS")}"
        Dim vYEARS As String = $"{drOB("YEARS")}"
        Dim vAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim vPLANID As String = $"{drOB("PLANID")}"
        Dim vRID As String = $"{drOB("RID")}"
        Dim vBCASENO As String = $"{drOB("BCASENO")}"
        'Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, "")
        Call TIMS.Check_dtBIDCASEFL(Me, dtFL, download_Path)
        DataGrid2.Columns(cst_DG2_退件原因_iCOLUMN).Visible = If($"{drOB("APPLIEDRESULT")}" = "R", True, False)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

        'Dim iProgress As Integer = If(dtA.Rows.Count > 0, (dt.Rows.Count / dtA.Rows.Count * 100), 0)
        '線上申辦進度 計算完成度百分比 (0-100)
        Dim iProgress As Integer = TIMS.GET_iPROGRESS_BI(sm, objconn, tmpMSG, vBCID, vORGKINDGW)
        labProgress.Text = $"{iProgress}%"
        'BTN_SAVETMP1.Visible = (iProgress = 100)
        'BTN_SAVERC2.Visible = (iProgress = 100)
        '儲存(暫存)
        BTN_SAVETMP1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, "", cst_tpmsg_enb1), True)
        '儲存後進下一步
        BTN_SAVENEXT1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, "", cst_tpmsg_enb1), True)
    End Sub

    ''' <summary>顯示項目資訊／狀態</summary>
    ''' <param name="oConn"></param>
    ''' <param name="vBCID"></param>
    ''' <param name="vORGKINDGW"></param>
    ''' <returns></returns>
    Private Function GET_ORG_BIDCASEFL_TB(ByRef oConn As SqlConnection, vBCID As String, vORGKINDGW As String) As DataTable
        Dim rParms As New Hashtable From {
            {"BCID", TIMS.CINT1(vBCID)},
            {"ORGKINDGW", vORGKINDGW}
        }
        Dim rsSql As String = ""
        rsSql &= " SELECT a.BCFID,a.YEARS,a.APPSTAGE,a.RID,a.BCID,a.KBSID,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        rsSql &= " ,kb.KBID,concat(kb.KBID,'.',kb.KBNAME) KBNAME" & vbCrLf
        rsSql &= " ,a.RTUREASON" & vbCrLf
        rsSql &= " ,ob.BISTATUS,ob.APPLIEDRESULT" & vbCrLf
        'rsSql &= " ,CONCAT(a.YEARS,'/',rr.PLANID,'/',a.RID,'/',ob.BCASENO,'/') PATH1" & vbCrLf
        rsSql &= " ,a.WAIVED,a.SRCFILENAME1,a.FILENAME1,a.FILENAME1 OKFLAG" & vbCrLf
        rsSql &= " FROM ORG_BIDCASEFL a" & vbCrLf
        rsSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=a.KBSID" & vbCrLf
        rsSql &= " JOIN ORG_BIDCASE ob on ob.BCID=a.BCID" & vbCrLf
        rsSql &= " JOIN AUTH_RELSHIP rr on rr.RID=a.RID" & vbCrLf
        rsSql &= " WHERE a.BCID=@BCID AND kb.ORGKINDGW=@ORGKINDGW" & vbCrLf
        rsSql &= " ORDER BY kb.KSORT,a.BCFID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        Return dt
    End Function

    ''' <summary>教學環境資料表 SAVE REPORT</summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_ORG_BIDCASE_ALL_13(rPMS As Hashtable)
        ',ORG_BIDCASEPI,ORG_BIDCASEFL
        'Dim vPCSVALUE As String = TIMS.GetMyValue2(rPMS, "PCSVALUE")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW") '$"{drKB("ORGKINDGW")}"
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        'Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vBCFID As String = TIMS.GetMyValue2(rPMS, "BCFID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vMODIFYACCT = "" Then Return
        'Dim vBCASENO As String = $"{drOB("BCASENO")}"

        Dim dtFLEV As DataTable = TIMS.GET_ORG_BIDCASEFL_EV(objconn, vBCID)

        Dim s_TMPMSG1 As String = ""
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_BIDCASEPI" & vbCrLf
        usSql &= " SET TECHENV='Y',TECHENVACCT=@TECHENVACCT,TECHENVDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND BCID=@BCID" & vbCrLf
        Dim iRow As Integer = 0
        For Each eItem As DataGridItem In DataGrid13.Items
            iRow += 1
            Dim HDG_PlanID As HtmlInputHidden = eItem.FindControl("HDG_PlanID")
            Dim HDG_ComIDNO As HtmlInputHidden = eItem.FindControl("HDG_ComIDNO")
            Dim HDG_SeqNo As HtmlInputHidden = eItem.FindControl("HDG_SeqNo")
            'Dim BTN_DOWNLOAD13 As Button = eItem.FindControl("BTN_DOWNLOAD13")
            Dim vHDG_PlanID As String = TIMS.ClearSQM(HDG_PlanID.Value)
            Dim vHDG_ComIDNO As String = TIMS.ClearSQM(HDG_ComIDNO.Value)
            Dim vHDG_SeqNo As String = TIMS.ClearSQM(HDG_SeqNo.Value)
            Dim vPCS As String = String.Concat(vHDG_PlanID, "x", vHDG_ComIDNO, "x", vHDG_SeqNo)

            '取得目前的序號找不到就不執行了
            Dim iBCPID As Integer = TIMS.GET_ORG_BIDCASEPI_iBCPID(sm, objconn, TIMS.CINT1(vBCID), vHDG_PlanID, vHDG_ComIDNO, vHDG_SeqNo)
            If iBCPID <= 0 Then Return

            Dim uParms As New Hashtable From {
                {"PLANID", TIMS.CINT1(HDG_PlanID.Value)},
                {"COMIDNO", HDG_ComIDNO.Value},
                {"SEQNO", TIMS.CINT1(HDG_SeqNo.Value)},
                {"BCID", TIMS.CINT1(vBCID)},
                {"TECHENVACCT", vMODIFYACCT}
            }
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
            tryFIND = String.Concat("BCPID=", iBCPID, " AND PlanID=", vHDG_PlanID, " AND ComIDNO='", vHDG_ComIDNO, "' AND SeqNo=", vHDG_SeqNo)
            If dtFLEV IsNot Nothing AndAlso dtFLEV.Rows.Count > 0 AndAlso dtFLEV.Select(tryFIND).Length > 0 Then
                Dim drFLEV As DataRow = dtFLEV.Select(tryFIND)(0) 'MODIFY_DAY
                Dim vMODIFY_DAY As String = Convert.ToString(drFLEV("MODIFY_DAY")) 'MODIFY_DAY
                Dim vMODIFY_MI As String = Convert.ToString(drFLEV("MODIFY_MI")) 'MODIFY_MI
                fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
            End If
            If fg_RUN_REPORT_1 Then
                Dim rPMS4 As New Hashtable
                rPMS4.Clear()
                rPMS4.Add("YEARS", vYEARS)
                rPMS4.Add("selsqlstr", String.Concat(vHDG_PlanID, "-", vHDG_ComIDNO, "-", vHDG_SeqNo))
                rPMS4.Add("TPlanID", sm.UserInfo.TPlanID)
                Dim s_RPTURL As String = GET_RPTURL_SD_14_014(rPMS4)
                Dim s_PDF_byte As Byte() = Nothing
                Try
                    Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
                Catch ex As Exception
                    Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
                    eErrmsg &= String.Concat(", s_RPTURL: ", s_RPTURL)
                    eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
                    eErrmsg &= String.Concat(", rPMS4: ", TIMS.GetMyValue4(rPMS4))
                    TIMS.LOG.Error(eErrmsg, ex)
                    Common.MessageBox(Me, "教學環境資料表下載檔案有誤，請確認檔案是否正確!")
                    Return
                End Try
                If s_PDF_byte IsNot Nothing Then
                    Dim xPMS As New Hashtable
                    TIMS.SetMyValue2(xPMS, "PLANID", vHDG_PlanID)
                    TIMS.SetMyValue2(xPMS, "PCS", vPCS)
                    TIMS.SetMyValue2(xPMS, "BCFID", TIMS.CINT1(vBCFID))
                    TIMS.SetMyValue2(xPMS, "BCPID", iBCPID)
                    TIMS.SetMyValue2(xPMS, "MODIFYACCT", sm.UserInfo.UserID)
                    TIMS.SetMyValue2(xPMS, "YEARS", vYEARS)
                    TIMS.SetMyValue2(xPMS, "APPSTAGE", vAPPSTAGE)
                    TIMS.SetMyValue2(xPMS, "RID", vRID)
                    TIMS.SetMyValue2(xPMS, "BCASENO", vBCASENO)
                    TIMS.SetMyValue2(xPMS, "KBSID", vKBSID)
                    TIMS.SetMyValue2(xPMS, "BCID", vBCID)
                    Call SAVE_ORG_BIDCASEFL_EV_PDF_FILE(xPMS, s_PDF_byte)
                End If
            Else
                s_TMPMSG1 &= String.Concat(If(s_TMPMSG1 <> "", ", ", ""), iRow)
            End If
        Next
        If s_TMPMSG1 <> "" Then
            Common.MessageBox(Me, String.Concat("(部份) 教學環境資料表重複處理時間過短(3分鐘1次)，請等待3分鐘後再試!", vbCrLf, s_TMPMSG1))
            'Return
        End If
    End Sub

    ''' <summary>教學環境資料表 SAVE REPORT</summary>
    ''' <param name="rPMS"></param>
    ''' <param name="s_PDF_byte"></param>
    Private Sub SAVE_ORG_BIDCASEFL_EV_PDF_FILE(ByRef rPMS As Hashtable, ByRef s_PDF_byte As Byte())
        If rPMS Is Nothing Then Return

        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vPCS As String = TIMS.GetMyValue2(rPMS, "PCS")
        Dim vBCFID As String = TIMS.GetMyValue2(rPMS, "BCFID")
        Dim vBCPID As String = TIMS.GetMyValue2(rPMS, "BCPID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        'Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vPATTERN As String = ""

        Dim vUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = "" 'TIMS.GET_FILENAME1_EV(vBCID, vKBSID, vPCS, "pdf")
        Dim vSRCFILENAME1 As String = "" 'vFILENAME1 'Convert.ToString(oSRCFILENAME1)
        '上傳檔案/存檔：檔名
        Try
            vUploadPath = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
            vFILENAME1 = TIMS.GET_FILENAME1_EV(vBCID, vKBSID, vPCS, "pdf")
            vSRCFILENAME1 = vFILENAME1 'Convert.ToString(oSRCFILENAME1)
            Call TIMS.MyCreateDir(Me, vUploadPath)
            File.WriteAllBytes(Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), s_PDF_byte)
            'IO.File.WriteAllText(Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), s_PDF_contents)
            '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
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

        Dim sParms3 As New Hashtable From {
            {"BCFID", vBCFID},
            {"BCID", vBCID},
            {"KBSID", vKBSID},
            {"BCPID", vBCPID}
        }
        Dim sSql3 As String = "SELECT BCFEID FROM ORG_BIDCASEFL_EV WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND BCPID=@BCPID"
        Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
        Dim iBCFEID As Integer = 0
        If dt3.Rows.Count = 0 Then
            'iParms.Add("PATTERN", PATTERN) 'iParms.Add("MEMO1", MEMO1) 'iParms.Add("WAIVED", WAIVED)'iParms.Add("MODIFYDATE", MODIFYDATE)
            iBCFEID = DbAccess.GetNewId(objconn, "ORG_BIDCASEFL_EV_BCFEID_SEQ,ORG_BIDCASEFL_EV,BCFEID")
            Dim iParms As New Hashtable From {
                {"BCFEID", iBCFEID},
                {"BCFID", vBCFID},
                {"BCID", vBCID},
                {"KBSID", vKBSID},
                {"BCPID", vBCPID},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", vMODIFYACCT}
            }
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_BIDCASEFL_EV(BCFEID, BCFID, BCID, KBSID, BCPID, FILENAME1, SRCFILENAME1, MODIFYACCT, MODIFYDATE)" & vbCrLf
            isSql &= " VALUES(@BCFEID,@BCFID,@BCID,@KBSID,@BCPID,@FILENAME1,@SRCFILENAME1, @MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            'iParms.Add("PATTERN", PATTERN) 'iParms.Add("MEMO1", MEMO1) 'iParms.Add("WAIVED", WAIVED)
            ',PATTERN=@PATTERN,MEMO1=@MEMO1,WAIVED=@WAIVED
            iBCFEID = dt3.Rows(0)("BCFEID")
            Dim uParms As New Hashtable From {
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", vMODIFYACCT},
                {"BCFEID", iBCFEID}
            }
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_BIDCASEFL_EV" & vbCrLf
            usSql &= " SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE BCFEID=@BCFEID" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        End If

    End Sub

    ''' <summary>混成課程教學環境資料表 SAVE REPORT</summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_ORG_BIDCASE_ALL_13B(rPMS As Hashtable)
        ',ORG_BIDCASEPI,ORG_BIDCASEFL
        Dim oDG1 As DataGrid = DataGrid13B
        'Dim vPCSVALUE As String = TIMS.GetMyValue2(rPMS, "PCSVALUE")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW") '$"{drKB("ORGKINDGW")}"
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        'Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vBCFID As String = TIMS.GetMyValue2(rPMS, "BCFID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vMODIFYACCT = "" Then Return
        'Dim vBCASENO As String = $"{drOB("BCASENO")}"

        Dim dtFLRT As DataTable = TIMS.GET_ORG_BIDCASEFL_RT(objconn, vBCID)

        Dim s_TMPMSG1 As String = ""
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_BIDCASEPI" & vbCrLf
        usSql &= " SET TECHRMT='Y',TECHRMTACCT=@TECHRMTACCT,TECHRMTDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND BCID=@BCID" & vbCrLf
        Dim iRow As Integer = 0
        For Each eItem As DataGridItem In oDG1.Items
            iRow += 1
            Dim HDG_PlanID As HtmlInputHidden = eItem.FindControl("HDG_PlanID")
            Dim HDG_ComIDNO As HtmlInputHidden = eItem.FindControl("HDG_ComIDNO")
            Dim HDG_SeqNo As HtmlInputHidden = eItem.FindControl("HDG_SeqNo")
            'Dim BTN_DOWNLOAD13 As Button = eItem.FindControl("BTN_DOWNLOAD13")
            Dim vHDG_PlanID As String = TIMS.ClearSQM(HDG_PlanID.Value)
            Dim vHDG_ComIDNO As String = TIMS.ClearSQM(HDG_ComIDNO.Value)
            Dim vHDG_SeqNo As String = TIMS.ClearSQM(HDG_SeqNo.Value)
            Dim vPCS As String = String.Concat(vHDG_PlanID, "x", vHDG_ComIDNO, "x", vHDG_SeqNo)

            '取得目前的序號找不到就不執行了
            Dim iBCPID As Integer = TIMS.GET_ORG_BIDCASEPI_iBCPID(sm, objconn, TIMS.CINT1(vBCID), vHDG_PlanID, vHDG_ComIDNO, vHDG_SeqNo)
            If iBCPID <= 0 Then Return

            Dim uParms As New Hashtable From {{"PLANID", TIMS.CINT1(HDG_PlanID.Value)}, {"COMIDNO", HDG_ComIDNO.Value}, {"SEQNO", TIMS.CINT1(HDG_SeqNo.Value)}, {"BCID", TIMS.CINT1(vBCID)}}
            uParms.Add("TECHRMTACCT", vMODIFYACCT)
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
            tryFIND = String.Concat("BCPID=", iBCPID, " AND PlanID=", vHDG_PlanID, " AND ComIDNO='", vHDG_ComIDNO, "' AND SeqNo=", vHDG_SeqNo)
            If dtFLRT IsNot Nothing AndAlso dtFLRT.Rows.Count > 0 AndAlso dtFLRT.Select(tryFIND).Length > 0 Then
                Dim drFLEV As DataRow = dtFLRT.Select(tryFIND)(0) 'MODIFY_DAY
                Dim vMODIFY_DAY As String = Convert.ToString(drFLEV("MODIFY_DAY")) 'MODIFY_DAY
                Dim vMODIFY_MI As String = Convert.ToString(drFLEV("MODIFY_MI")) 'MODIFY_MI
                fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
            End If
            If fg_RUN_REPORT_1 Then
                Dim rPMS4 As New Hashtable From {
                    {"YEARS", vYEARS},
                    {"selsqlstr", String.Concat(vHDG_PlanID, "-", vHDG_ComIDNO, "-", vHDG_SeqNo)},
                    {"TPlanID", sm.UserInfo.TPlanID}
                } 'rPMS4.Clear()
                Dim s_RPTURL As String = GET_RPTURL_SD_14_014R(rPMS4)
                Dim s_PDF_byte As Byte() = Nothing
                Try
                    Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
                Catch ex As Exception
                    Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
                    eErrmsg &= String.Concat(",混成課程教學環境資料表 s_RPTURL: ", s_RPTURL)
                    eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
                    eErrmsg &= String.Concat(", rPMS4: ", TIMS.GetMyValue4(rPMS4))
                    TIMS.LOG.Error(eErrmsg, ex)
                    Common.MessageBox(Me, "混成課程教學環境資料表下載檔案有誤，請確認檔案是否正確!")
                    Return
                End Try
                If s_PDF_byte IsNot Nothing Then
                    Dim xPMS As New Hashtable
                    TIMS.SetMyValue2(xPMS, "PLANID", vHDG_PlanID)
                    TIMS.SetMyValue2(xPMS, "PCS", vPCS)
                    TIMS.SetMyValue2(xPMS, "BCFID", TIMS.CINT1(vBCFID))
                    TIMS.SetMyValue2(xPMS, "BCPID", iBCPID)
                    TIMS.SetMyValue2(xPMS, "MODIFYACCT", sm.UserInfo.UserID)
                    TIMS.SetMyValue2(xPMS, "YEARS", vYEARS)
                    TIMS.SetMyValue2(xPMS, "APPSTAGE", vAPPSTAGE)
                    TIMS.SetMyValue2(xPMS, "RID", vRID)
                    TIMS.SetMyValue2(xPMS, "BCASENO", vBCASENO)
                    TIMS.SetMyValue2(xPMS, "KBSID", vKBSID)
                    TIMS.SetMyValue2(xPMS, "BCID", vBCID)
                    Call SAVE_ORG_BIDCASEFL_RT_PDF_FILE(xPMS, s_PDF_byte)
                End If
            Else
                s_TMPMSG1 &= String.Concat(If(s_TMPMSG1 <> "", ", ", ""), iRow)
            End If
        Next
        If s_TMPMSG1 <> "" Then
            Common.MessageBox(Me, String.Concat("(部份) 混成課程教學環境資料表重複處理時間過短(3分鐘1次)，請等待3分鐘後再試!", vbCrLf, s_TMPMSG1))
            'Return
        End If
    End Sub

    ''' <summary>混成課程教學環境資料表 SAVE REPORT</summary>
    ''' <param name="rPMS"></param>
    ''' <param name="s_PDF_byte"></param>
    Private Sub SAVE_ORG_BIDCASEFL_RT_PDF_FILE(ByRef rPMS As Hashtable, ByRef s_PDF_byte As Byte())
        If rPMS Is Nothing Then Return

        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vPCS As String = TIMS.GetMyValue2(rPMS, "PCS")
        Dim vBCFID As String = TIMS.GetMyValue2(rPMS, "BCFID")
        Dim vBCPID As String = TIMS.GetMyValue2(rPMS, "BCPID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCASENO As String = TIMS.GetMyValue2(rPMS, "BCASENO")
        Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        'Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vPATTERN As String = ""

        Dim vUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = "" 'TIMS.GET_FILENAME1_EV(vBCID, vKBSID, vPCS, "pdf")
        Dim vSRCFILENAME1 As String = "" 'vFILENAME1 'Convert.ToString(oSRCFILENAME1)
        '上傳檔案/存檔：檔名
        Try
            vUploadPath = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
            vFILENAME1 = TIMS.GET_FILENAME1_RT(vBCID, vKBSID, vPCS, "pdf")
            vSRCFILENAME1 = vFILENAME1 'Convert.ToString(oSRCFILENAME1)
            Call TIMS.MyCreateDir(Me, vUploadPath)
            File.WriteAllBytes(Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), s_PDF_byte)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2) 'Common.MessageBox(Me, ex.ToString)

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

        Dim sParms3 As New Hashtable From {
            {"BCFID", vBCFID},
            {"BCID", vBCID},
            {"KBSID", vKBSID},
            {"BCPID", vBCPID}
        }
        Dim sSql3 As String = "SELECT BCRTID FROM ORG_BIDCASEFL_RT WHERE BCFID=@BCFID AND BCID=@BCID AND KBSID=@KBSID AND BCPID=@BCPID"
        Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
        Dim iBCRTID As Integer = 0
        If dt3.Rows.Count = 0 Then
            iBCRTID = DbAccess.GetNewId(objconn, "ORG_BIDCASEFL_RT_BCRTID_SEQ,ORG_BIDCASEFL_RT,BCRTID")
            Dim iParms As New Hashtable From {
                {"BCRTID", iBCRTID},
                {"BCFID", vBCFID},
                {"BCID", vBCID},
                {"KBSID", vKBSID},
                {"BCPID", vBCPID},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", vMODIFYACCT}
            }
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_BIDCASEFL_RT(BCRTID, BCFID, BCID, KBSID, BCPID, FILENAME1, SRCFILENAME1, MODIFYACCT, MODIFYDATE)" & vbCrLf
            isSql &= " VALUES(@BCRTID,@BCFID,@BCID,@KBSID,@BCPID,@FILENAME1,@SRCFILENAME1, @MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            'iParms.Add("PATTERN", PATTERN) 'iParms.Add("MEMO1", MEMO1) 'iParms.Add("WAIVED", WAIVED)
            ',PATTERN=@PATTERN,MEMO1=@MEMO1,WAIVED=@WAIVED
            iBCRTID = dt3.Rows(0)("BCRTID")
            Dim uParms As New Hashtable From {
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"MODIFYACCT", vMODIFYACCT},
                {"BCRTID", iBCRTID}
            }
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_BIDCASEFL_RT" & vbCrLf
            usSql &= " SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE BCRTID=@BCRTID" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        End If

    End Sub

    ''' <summary>SAVE_ORG_BIDCASE_ALL_W08/SAVE_ORG_BIDCASE_ALL_08 訓練班別計畫表 以目前版本批次送出</summary>
    Private Sub SAVE_ORG_BIDCASEPI_08(ByRef rPMS As Hashtable)
        ',ORG_BIDCASEPI,ORG_BIDCASEFL
        'Dim vPCSVALUE As String = TIMS.GetMyValue2(rPMS, "PCSVALUE")
        'Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW") '$"{drKB("ORGKINDGW")}"
        'Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        ''Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
        'Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        'Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vPlanID As String = TIMS.GetMyValue2(rPMS, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue2(rPMS, "ComIDNO")
        Dim vSeqNo As String = TIMS.GetMyValue2(rPMS, "SeqNo")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")

        'Dim vKBSID As String = TIMS.GetMyValue2(rPMS, "KBSID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vMODIFYACCT = "" Then Return

        Dim usSql As String = ""
        usSql &= " UPDATE ORG_BIDCASEPI" & vbCrLf
        usSql &= " SET SENTBATVER='Y',SENTBATACCT=@SENTBATACCT,SENTBATDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO AND BCID=@BCID" & vbCrLf
        Dim uParms As New Hashtable From {
            {"PLANID", TIMS.CINT1(vPlanID)},
            {"COMIDNO", vComIDNO},
            {"SEQNO", TIMS.CINT1(vSeqNo)},
            {"BCID", TIMS.CINT1(vBCID)},
            {"SENTBATACCT", vMODIFYACCT}
        }
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        'For Each eItem As DataGridItem In DataGrid08.Items
        '    'Dim HDG8_OCID As HtmlInputHidden = e.Item.FindControl("HDG8_OCID")
        '    Dim HDG8_PlanID As HtmlInputHidden = eItem.FindControl("HDG8_PlanID")
        '    Dim HDG8_ComIDNO As HtmlInputHidden = eItem.FindControl("HDG8_ComIDNO")
        '    Dim HDG8_SeqNo As HtmlInputHidden = eItem.FindControl("HDG8_SeqNo")
        '    'Dim HDG8_PrintRpt1 As HtmlInputButton = eItem.FindControl("HDG8_PrintRpt1")
        '    Dim uParms As New Hashtable
        '    uParms.Add("PLANID", TIMS.CINT1(HDG8_PlanID.Value))
        '    uParms.Add("COMIDNO", HDG8_ComIDNO.Value)
        '    uParms.Add("SEQNO", TIMS.CINT1(HDG8_SeqNo.Value))
        '    uParms.Add("BCID", TIMS.CINT1(vBCID))
        '    uParms.Add("SENTBATACCT", vMODIFYACCT)
        '    DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        'Next
    End Sub

    ''' <summary>最近一次版本送件</summary>
    ''' <param name="drOB">ORG_BIDCASE</param>
    ''' <param name="drKB">KEY_BIDCASE</param>
    ''' <param name="drRR"></param>
    Private Sub FILE_COPY1(drOB As DataRow, drKB As DataRow, drRR As DataRow, MTYPE As String)
        'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試
        '依資料找到機構資料
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))

        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        Dim drOF As DataRow = Nothing
        Dim srPMS As New Hashtable From {
            {"ORGKINDGW", drKB("ORGKINDGW")},
            {"KBID", drKB("KBID")},
            {"ORGID", drRR("ORGID")},
            {"DISTID", drRR("DISTID")},
            {"RID", drRR("RID")},
            {"APPSTAGE", $"{drOB("APPSTAGE")}"},
            {"BCID", drOB("BCID")}
        }
        drOF = GET_OLDFILE1(srPMS)
        If drOF Is Nothing Then
            'String.Concat("(查無資料)!", TIMS.GetMyValue3(srPMS))
            TIMS.LOG.Warn(String.Concat("(查無資料)!", TIMS.GetMyValue3(srPMS)))
            Common.MessageBox(Me, "(查無資料)!")
            Return
        End If

        Dim oYEARS As String = Convert.ToString(drOF("YEARS"))
        Dim oAPPSTAGE As String = Convert.ToString(drOF("APPSTAGE"))
        Dim oPLANID As String = Convert.ToString(drOF("PLANID"))
        Dim oRID As String = Convert.ToString(drOF("RID"))
        Dim oBCASENO As String = Convert.ToString(drOF("BCASENO"))
        '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim oUploadPath As String = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
        Dim oFILENAME1 As String = Convert.ToString(drOF("FILENAME1"))
        Dim oSRCFILENAME1 As String = Convert.ToString(drOF("SRCFILENAME1"))

        Dim s_LOGMSG1 As String = String.Concat("MTYPE: ", MTYPE, vbCrLf)
        s_LOGMSG1 &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
        s_LOGMSG1 &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
        s_LOGMSG1 &= String.Concat("oSRCFILENAME1: ", oSRCFILENAME1, vbCrLf)
        s_LOGMSG1 &= String.Concat("Path.Combine(oUploadPath, oFILENAME1): ", Path.Combine(oUploadPath, oFILENAME1), vbCrLf)
        TIMS.LOG.Debug(s_LOGMSG1)

        If MTYPE = cst_MTYPE_LATEST_DOWN1 Then
            'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Call ResponsePDFFile1(Me, objconn, Path.Combine(oUploadPath, oFILENAME1))
            Return
        End If

        '檔案複制  '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid/
        Dim vPATTERN As String = ""
        'Dim vPATTERN As String = If(flag_test, TIMS.GetDateNo2(4), "")
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, "") 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_B(vBCID, vKBSID, vPATTERN, "pdf")
        Dim vSRCFILENAME1 As String = Convert.ToString(oSRCFILENAME1)
        '上傳檔案/存檔：檔名
        Try
            Call TIMS.MyCreateDir(Me, vUploadPath)
            IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
            '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
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
            Exit Sub
        End Try

        '增加資料一筆
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", drKB("ORGKINDGW"))
            TIMS.SetMyValue2(rPMS2, "YEARS", drOB("YEARS"))
            'TIMS.SetMyValue2(rPMS2, "APPSTAGE", drOB("APPSTAGE"))
            TIMS.SetMyValue2(rPMS2, "RID", RIDValue.Value)
            TIMS.SetMyValue2(rPMS2, "BCID", Hid_BCID.Value)
            TIMS.SetMyValue2(rPMS2, "KBSID", Hid_KBSID.Value)
            'TIMS.SetMyValue2(rPMS2, "WAIVED", If(CHKB_WAIVED.Checked, "Y", ""))
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            Call SAVE_ORG_BIDCASEFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub  'Throw ex
        End Try

    End Sub

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

    ''' <summary>取得最後一筆資訊</summary>
    ''' <param name="rPMS"></param>
    ''' <returns></returns>
    Function GET_OLDFILE1(ByRef rPMS As Hashtable) As DataRow
        'Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
        Dim drRst As DataRow = Nothing
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue2(rPMS, "KBID")
        Dim vORGID As String = TIMS.GetMyValue2(rPMS, "ORGID")
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS, "DISTID")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        'Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")

        Dim sParms As New Hashtable From {
            {"ORGKINDGW", vORGKINDGW},
            {"KBID", vKBID},
            {"ORGID", vORGID},
            {"DISTID", vDISTID},
            {"BCID", vBCID}
        }
        'sParms.Add("RID", vRID)
        Dim sSql As String = ""
        sSql &= " SELECT f.BCFID,f.YEARS,f.APPSTAGE,a.PLANID,a.RID,a.BCASENO,f.FILENAME1,f.SRCFILENAME1" & vbCrLf
        sSql &= " FROM ORG_BIDCASE a" & vbCrLf
        sSql &= " JOIN ORG_BIDCASEFL f on f.BCID=a.BCID" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID" & vbCrLf
        sSql &= " WHERE kb.ORGKINDGW=@ORGKINDGW" & vbCrLf
        sSql &= " AND kb.KBID=@KBID" & vbCrLf
        sSql &= " AND a.ORGID=@ORGID" & vbCrLf
        sSql &= " AND a.DISTID=@DISTID" & vbCrLf
        sSql &= " AND a.BCID!=@BCID" & vbCrLf
        'sSql &= " AND a.RID !=@RID" & vbCrLf
        'If Not fg_test Then
        '    sParms.Add("RIDAPPSTAGE", String.Concat(vRID, "x", vAPPSTAGE))
        '    sSql &= " AND concat(a.RID,'x',a.APPSTAGE)!=@RIDAPPSTAGE" & vbCrLf
        'End If
        sSql &= " AND f.FILENAME1 IS NOT NULL" & vbCrLf
        sSql &= " ORDER BY a.BCID DESC" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return drRst
        drRst = dt.Rows(0)
        Return drRst
    End Function

    ''' <summary>最近一次版本送件-師資助教基本資料表</summary>
    ''' <param name="drOB"></param>
    ''' <param name="drKB"></param>
    ''' <param name="drRR"></param>
    Private Sub FILE_COPY1_TT(drOB As DataRow, drKB As DataRow, drRR As DataRow, MTYPE As String)
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
            Common.MessageBox(Me, "資訊有誤(未選擇老師)，請重新操作!!")
            Return
        End If

        'Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
        '依資料找到機構資料
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))

        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        'Select Case String.Concat(vORGKINDGW, vKBID) 'Case cst_G10_師資助教基本資料表, cst_W10_師資助教基本資料表
        Dim drOF As DataRow = Nothing
        Dim srPMS As New Hashtable From {
            {"ORGKINDGW", drKB("ORGKINDGW")},
            {"KBID", drKB("KBID")},
            {"ORGID", drRR("ORGID")},
            {"DISTID", drRR("DISTID")},
            {"TECHID", vTECHID},
            {"RID", drRR("RID")},
            {"APPSTAGE", drOB("APPSTAGE")},
            {"BCID", drOB("BCID")}
        }
        drOF = GET_OLDFILE1_TT(srPMS)
        If drOF Is Nothing Then
            TIMS.LOG.Warn(String.Concat("選擇老師(查無資料)!", TIMS.GetMyValue3(srPMS)))
            Common.MessageBox(Me, "選擇老師(查無資料)!")
            Return
        End If

        Dim oYEARS As String = Convert.ToString(drOF("YEARS"))
        Dim oAPPSTAGE As String = Convert.ToString(drOF("APPSTAGE"))
        Dim oPLANID As String = Convert.ToString(drOF("PLANID"))
        Dim oRID As String = Convert.ToString(drOF("RID"))
        Dim oBCASENO As String = Convert.ToString(drOF("BCASENO"))
        '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim oUploadPath As String = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
        Dim oFILENAME1 As String = Convert.ToString(drOF("FILENAME1"))
        Dim oSRCFILENAME1 As String = Convert.ToString(drOF("SRCFILENAME1"))

        Dim s_LOGMSG1 As String = String.Concat("MTYPE: ", MTYPE, vbCrLf)
        s_LOGMSG1 &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
        s_LOGMSG1 &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
        s_LOGMSG1 &= String.Concat("oSRCFILENAME1: ", oSRCFILENAME1, vbCrLf)
        s_LOGMSG1 &= String.Concat("Path.Combine(oUploadPath, oFILENAME1): ", Path.Combine(oUploadPath, oFILENAME1), vbCrLf)
        TIMS.LOG.Debug(s_LOGMSG1)

        If MTYPE = cst_MTYPE_LATEST_DOWN1 Then
            'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Call ResponsePDFFile1(Me, objconn, Path.Combine(oUploadPath, oFILENAME1))
            Return
        End If

        '檔案複制  '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim vPATTERN As String = ""
        'Dim vPATTERN As String = If(fg_test, TIMS.GetDateNo2(4), "")
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_T(vBCID, vKBSID, vTECHID, "pdf")
        Dim vSRCFILENAME1 As String = Convert.ToString(oSRCFILENAME1)

        '上傳檔案/存檔：檔名
        Try
            Call TIMS.MyCreateDir(Me, vUploadPath)
            IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
            '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
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
            Exit Sub
        End Try

        Dim rPMS As New Hashtable From {{"TECHID", Hid_TECHID.Value}, {"FILENAME1", vFILENAME1}, {"SRCFILENAME1", vSRCFILENAME1}}
        Call SAVE_ORG_BIDCASEFL_TT(drOB, drKB, drRR, rPMS)
    End Sub

    ''' <summary>取得最後一筆資訊-師資助教基本資料表</summary>
    ''' <param name="rPMS"></param>
    ''' <returns></returns>
    Function GET_OLDFILE1_TT(ByRef rPMS As Hashtable) As DataRow
        'Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
        Dim drRst As DataRow = Nothing
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue2(rPMS, "KBID")
        Dim vORGID As String = TIMS.GetMyValue2(rPMS, "ORGID")
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS, "DISTID")
        Dim vTECHID As String = TIMS.GetMyValue2(rPMS, "TECHID")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.GetMyValue2(rPMS, "BCID")
        'Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")

        Dim sParms1 As New Hashtable From {
            {"ORGKINDGW", vORGKINDGW},
            {"KBID", vKBID},
            {"ORGID", vORGID},
            {"DISTID", vDISTID},
            {"TECHID", vTECHID},
            {"BCID", vBCID}
        }
        'sParms1.Add("RIDAPPSTAGE", String.Concat(vRID, "x", vAPPSTAGE))
        Dim sSql1 As String = ""
        sSql1 &= " SELECT f.BCFID,f.YEARS,f.APPSTAGE,a.PLANID,a.RID,a.BCASENO,t.FILENAME1,t.SRCFILENAME1" & vbCrLf
        sSql1 &= " FROM ORG_BIDCASEFL_TT t" & vbCrLf
        sSql1 &= " JOIN ORG_BIDCASEFL f on f.BCFID=t.BCFID" & vbCrLf
        sSql1 &= " JOIN ORG_BIDCASE a on a.BCID=f.BCID" & vbCrLf
        sSql1 &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID" & vbCrLf
        sSql1 &= " WHERE kb.ORGKINDGW=@ORGKINDGW AND kb.KBID=@KBID" & vbCrLf
        sSql1 &= " AND a.ORGID=@ORGID AND a.DISTID=@DISTID" & vbCrLf
        sSql1 &= " AND t.TECHID=@TECHID" & vbCrLf
        sSql1 &= " AND a.BCID!=@BCID" & vbCrLf
        'sSql1 &= " AND a.RID !=@RID" & vbCrLf
        'sSql1 &= " AND concat(a.RID,'x',a.APPSTAGE)!=@RIDAPPSTAGE" & vbCrLf
        sSql1 &= " AND t.FILENAME1 is not null" & vbCrLf
        sSql1 &= " ORDER BY a.BCID DESC" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql1, objconn, sParms1)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return drRst
        drRst = dt.Rows(0)
        Return drRst
    End Function

    ''' <summary>最近一次版本送件-授課師資學經歷證書影本</summary>
    ''' <param name="drOB"></param>
    ''' <param name="drKB"></param>
    ''' <param name="drRR"></param>
    Private Sub FILE_COPY1_TT2(drOB As DataRow, drKB As DataRow, drRR As DataRow, MTYPE As String)
        Hid_TECHID.Value = ""
        For Each eItem As DataGridItem In DataGrid11.Items
            Dim chkItem1 As HtmlInputCheckBox = eItem.FindControl("chkItem1")
            Dim HDG11_TechID As HtmlInputHidden = eItem.FindControl("HDG11_TechID")
            'Dim HDG11_RID As HtmlInputHidden = eItem.FindControl("HDG11_RID")
            If chkItem1.Checked AndAlso HDG11_TechID.Value <> "" Then
                Hid_TECHID.Value = HDG11_TechID.Value
                Exit For
            End If
        Next
        Hid_TECHID.Value = TIMS.ClearSQM(Hid_TECHID.Value)
        Dim vTECHID As String = Hid_TECHID.Value
        If vTECHID = "" Then
            Common.MessageBox(Me, "資訊有誤(未選擇老師)，請重新操作!!")
            Return
        End If

        'Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
        '依資料找到機構資料
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        'Select Case String.Concat(vORGKINDGW, vKBID) 'Case cst_G10_師資助教基本資料表, cst_W10_師資助教基本資料表
        Dim drOF As DataRow = Nothing
        Dim srPMS As New Hashtable From {
            {"ORGKINDGW", drKB("ORGKINDGW")},
            {"KBID", drKB("KBID")},
            {"ORGID", drRR("ORGID")},
            {"DISTID", drRR("DISTID")},
            {"TECHID", vTECHID},
            {"RID", drRR("RID")},
            {"APPSTAGE", drOB("APPSTAGE")}
        }
        drOF = GET_OLDFILE1_TT2(srPMS)
        If drOF Is Nothing Then
            TIMS.LOG.Warn(String.Concat("選擇老師(查無資料)!!", TIMS.GetMyValue3(srPMS)))
            Common.MessageBox(Me, "選擇老師(查無資料)!!")
            Return
        End If

        Dim oYEARS As String = Convert.ToString(drOF("YEARS"))
        Dim oAPPSTAGE As String = Convert.ToString(drOF("APPSTAGE"))
        Dim oPLANID As String = Convert.ToString(drOF("PLANID"))
        Dim oRID As String = Convert.ToString(drOF("RID"))
        Dim oBCASENO As String = Convert.ToString(drOF("BCASENO"))
        '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim oUploadPath As String = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
        Dim oFILENAME1 As String = Convert.ToString(drOF("FILENAME1"))
        Dim oSRCFILENAME1 As String = Convert.ToString(drOF("SRCFILENAME1"))

        Dim s_LOGMSG1 As String = String.Concat("MTYPE: ", MTYPE, vbCrLf)
        s_LOGMSG1 &= String.Concat("oUploadPath: ", oUploadPath, vbCrLf)
        s_LOGMSG1 &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf)
        s_LOGMSG1 &= String.Concat("oSRCFILENAME1: ", oSRCFILENAME1, vbCrLf)
        s_LOGMSG1 &= String.Concat("Path.Combine(oUploadPath, oFILENAME1): ", Path.Combine(oUploadPath, oFILENAME1), vbCrLf)
        TIMS.LOG.Debug(s_LOGMSG1)

        If MTYPE = cst_MTYPE_LATEST_DOWN1 Then
            'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            Call ResponsePDFFile1(Me, objconn, Path.Combine(oUploadPath, oFILENAME1))
            Return
        End If

        '檔案複制  '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim vPATTERN As String = ""
        'Dim vPATTERN As String = If(fg_test, TIMS.GetDateNo2(4), "")
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_T(vBCID, vKBSID, vTECHID, "pdf")
        Dim vSRCFILENAME1 As String = Convert.ToString(oSRCFILENAME1)

        '上傳檔案/存檔：檔名
        Try
            Call TIMS.MyCreateDir(Me, vUploadPath)
            IO.File.Copy(Server.MapPath(Path.Combine(oUploadPath, oFILENAME1)), Server.MapPath(Path.Combine(vUploadPath, vFILENAME1)), True)
            '上傳檔案 'TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
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
            Exit Sub
        End Try

        Dim rPMS As New Hashtable From {
            {"TECHID", Hid_TECHID.Value},
            {"FILENAME1", vFILENAME1},
            {"SRCFILENAME1", vSRCFILENAME1}
        }
        Call SAVE_ORG_BIDCASEFL_TT2(drOB, drKB, drRR, rPMS)
    End Sub

    ''' <summary>取得最後一筆資訊-授課師資學經歷證書影本</summary>
    ''' <param name="rPMS"></param>
    ''' <returns></returns>
    Function GET_OLDFILE1_TT2(ByRef rPMS As Hashtable) As DataRow
        'Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
        Dim drRst As DataRow = Nothing
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vKBID As String = TIMS.GetMyValue2(rPMS, "KBID")
        Dim vORGID As String = TIMS.GetMyValue2(rPMS, "ORGID")
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS, "DISTID")
        Dim vTECHID As String = TIMS.GetMyValue2(rPMS, "TECHID")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")

        Dim sParms1 As New Hashtable From {
            {"ORGKINDGW", vORGKINDGW},
            {"KBID", vKBID},
            {"ORGID", vORGID},
            {"DISTID", vDISTID},
            {"TECHID", vTECHID},
            {"RIDAPPSTAGE", String.Concat(vRID, "x", vAPPSTAGE)}
        }
        Dim sSql1 As String = ""
        sSql1 &= " SELECT f.BCFID,f.YEARS,f.APPSTAGE,a.PLANID,a.RID,a.BCASENO,t.FILENAME1,t.SRCFILENAME1" & vbCrLf
        sSql1 &= " FROM ORG_BIDCASEFL_TT2 t" & vbCrLf
        sSql1 &= " JOIN ORG_BIDCASEFL f on f.BCFID=t.BCFID" & vbCrLf
        sSql1 &= " JOIN ORG_BIDCASE a on a.BCID=f.BCID" & vbCrLf
        sSql1 &= " JOIN KEY_BIDCASE kb on kb.KBSID=f.KBSID" & vbCrLf
        sSql1 &= " WHERE kb.ORGKINDGW=@ORGKINDGW" & vbCrLf
        sSql1 &= " AND kb.KBID=@KBID" & vbCrLf
        sSql1 &= " AND a.ORGID=@ORGID" & vbCrLf
        sSql1 &= " AND a.DISTID=@DISTID" & vbCrLf
        sSql1 &= " AND t.TECHID=@TECHID" & vbCrLf
        'sSql1 &= " AND a.RID !=@RID" & vbCrLf
        sSql1 &= " AND concat(a.RID,'x',a.APPSTAGE)!=@RIDAPPSTAGE" & vbCrLf
        sSql1 &= " AND t.FILENAME1 is not null" & vbCrLf
        sSql1 &= " ORDER BY a.BCID DESC" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sSql1, objconn, sParms1)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return drRst
        drRst = dt.Rows(0)
        Return drRst
    End Function

    ''' <summary>上傳 各授課師資學經歷證書影本 G/W</summary>
    ''' <param name="drOB"></param>
    Private Sub FILE_UPLOAD_11(drOB As DataRow)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        If drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If
        'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn,RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Hid_TECHID.Value = ""
        For Each eItem As DataGridItem In DataGrid11.Items
            Dim chkItem1 As HtmlInputCheckBox = eItem.FindControl("chkItem1")
            Dim HDG11_TechID As HtmlInputHidden = eItem.FindControl("HDG11_TechID")
            'Dim HDG11_RID As HtmlInputHidden = eItem.FindControl("HDG11_RID")
            If chkItem1.Checked AndAlso HDG11_TechID.Value <> "" Then
                Hid_TECHID.Value = HDG11_TechID.Value
                Exit For
            End If
        Next
        Hid_TECHID.Value = TIMS.ClearSQM(Hid_TECHID.Value)
        Dim vTECHID As String = Hid_TECHID.Value
        If vTECHID = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(未選擇老師)，請重新操作!!")
            Return
        End If
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Exit Sub
        End If

        'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        'Dim vKBSID As String = TIMS.GetListValue(ddlSwitchTo) 'TIMS.GetMyValue2(rPMS, "KBGWID")
        'Dim vPATTERN As String = "" 'TIMS.GetMyValue2(rPMS, "PATTERN")
        '年度／計畫ID／機構ID／caseno／1
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vBCASENO <> Hid_BCASENO.Value Then Return '(此狀況不太可能發生)

        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFilePdf(Me, File1, MyPostedFile) Then Return

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "pdf"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_15M Then
                    Common.MessageBox(Me, cst_errMsg_7_15M)
                    Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Exit Sub
        End Select

        '取得KBID代號／非流水號
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_T(vBCID, vKBSID, vTECHID, "pdf")
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
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Dim rPMS As New Hashtable From {
            {"TECHID", Hid_TECHID.Value},
            {"FILENAME1", vFILENAME1},
            {"SRCFILENAME1", vSRCFILENAME1}
        }
        Call SAVE_ORG_BIDCASEFL_TT2(drOB, drKB, drRR, rPMS)
    End Sub

    ''' <summary>上傳 10.師資／助教基本資料表 G/W</summary>
    ''' <param name="drOB"></param>
    Private Sub FILE_UPLOAD_10(drOB As DataRow)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        If drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If
        'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn,RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

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
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Exit Sub
        End If

        'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        'Dim vKBSID As String = TIMS.GetListValue(ddlSwitchTo) 'TIMS.GetMyValue2(rPMS, "KBGWID")
        'Dim vPATTERN As String = "" 'TIMS.GetMyValue2(rPMS, "PATTERN")

        '年度／計畫ID／機構ID／caseno／1
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vBCASENO <> Hid_BCASENO.Value Then Return '(此狀況不太可能發生)

        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFilePdf(Me, File1, MyPostedFile) Then Return

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "pdf"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                    Common.MessageBox(Me, cst_errMsg_7_10M)
                    Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Exit Sub
        End Select

        '取得KBID代號／非流水號
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_T(vBCID, vKBSID, vTECHID, "pdf")
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
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        'Dim vTECHID As String = TIMS.GetMyValue2(rPMS, "TECHID")
        'Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        'Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        'drOB As DataRow, drKB As DataRow, drRR As DataRow
        Dim rPMS As New Hashtable From {
            {"TECHID", Hid_TECHID.Value},
            {"FILENAME1", vFILENAME1},
            {"SRCFILENAME1", vSRCFILENAME1}
        }
        Call SAVE_ORG_BIDCASEFL_TT(drOB, drKB, drRR, rPMS)
    End Sub

    ''' <summary>檔案上傳</summary>
    ''' <param name="drOB"></param>
    Private Sub FILE_UPLOAD_1(ByRef drOB As DataRow, iBCFID As Integer)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        If drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If
        'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn,RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        'Hid_TECHID.Value = TIMS.ClearSQM(Hid_TECHID)
        'Dim vTECHID As String = Hid_TECHID.Value
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")

        'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        'Dim vKBSID As String = TIMS.GetListValue(ddlSwitchTo) 'TIMS.GetMyValue2(rPMS, "KBGWID")
        Dim vPATTERN As String = "" 'TIMS.GetMyValue2(rPMS, "PATTERN")

        '年度／計畫ID／機構ID／caseno／1
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vBCASENO <> Hid_BCASENO.Value Then Return '(此狀況不太可能發生)

        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Exit Sub
        End If

        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFilePdf(Me, File1, MyPostedFile) Then Return

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "pdf"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                    Common.MessageBox(Me, cst_errMsg_7_10M)
                    Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Exit Sub
        End Select

        '取得KBID代號／非流水號
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, "")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_B(vBCID, vKBSID, vPATTERN, "pdf") 'String.Concat("B", TIMS.GetDateNo2(4), "x", vBCID, "x", vKBSID, vPATTERN, ".pdf")
        Dim vSRCFILENAME1 As String = MyFileName
        '上傳檔案/存檔：檔名
        Try
            '上傳檔案
            TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename = GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = String.Concat(ex.Message, vbCrLf, "ex.ToString:", ex.ToString, vbCrLf)
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath(String.Concat(vUploadPath, vFILENAME1)), vbCrLf)
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "BCFID", If(vUploadPath <> "", iBCFID, -1)) '(可再次傳送)

            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "YEARS", vYEARS)
            'TIMS.SetMyValue2(rPMS2, "APPSTAGE", vAPPSTAGE)
            TIMS.SetMyValue2(rPMS2, "RID", vRID)
            TIMS.SetMyValue2(rPMS2, "BCID", vBCID)
            TIMS.SetMyValue2(rPMS2, "KBSID", vKBSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", "")
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", vMODIFYACCT)
            Call SAVE_ORG_BIDCASEFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try
    End Sub
#End Region

    ''' <summary>確定檔案上傳</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But1_Click(sender As Object, e As EventArgs) Handles But1.Click
        'Dim vUploadPath As String = Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If Hid_BCASENO.Value = "" OrElse Hid_BCID.Value = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(案件號為空)，請重新操作!!")
            Return
        End If

        'SAVE_ORG_BIDCASEFL
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        '取得KBID代號／非流水號
        'Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_G08_1_iCap課程原始申請資料, TIMS.cst_W08_1_iCap課程原始申請資料
                'iCap課程原始申請資料(上傳)
                Call FILE_UPLOAD_14(drOB)
                Call SHOW_DATAGRID_14()

            Case TIMS.cst_G10_師資助教基本資料表, TIMS.cst_W10_師資助教基本資料表
                'vFILENAME1 = String.Concat("T", TIMS.GetDateNo2(4), "x", vBCID, "x", vKBSID, vPATTERN, "x", vTECHID, ".pdf")  'TIMS.GetMyValue2(rPMS, "FILENAME1")
                '師資檔案上傳
                Call FILE_UPLOAD_10(drOB)
                Call SHOW_DATAGRID_10()

            Case TIMS.cst_G11_授課師資學經歷證書影本, TIMS.cst_W11_授課師資學經歷證書影本
                'vFILENAME1 = String.Concat("T", TIMS.GetDateNo2(4), "x", vBCID, "x", vKBSID, vPATTERN, "x", vTECHID, ".pdf")  'TIMS.GetMyValue2(rPMS, "FILENAME1")
                '師資檔案上傳
                Call FILE_UPLOAD_11(drOB)
                Call SHOW_DATAGRID_11()

            Case Else
                Dim vBCID As String = TIMS.ClearSQM(Hid_BCID.Value)
                Dim vKBSID As String = TIMS.ClearSQM(Hid_KBSID.Value)
                Dim drFL As DataRow = TIMS.GET_ORG_BIDCASEFL(objconn, vBCID, vKBSID)
                '(退件修正)有退件原因,可重新上傳
                'Dim flag_NG_UPLOAD_1 As Boolean = (drFL IsNot Nothing) '(有資料 不可再次傳送)
                Dim flag_NG_UPLOAD_2 As Boolean = (drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) = "") '(有資料不可傳送且原因為空 不可再次傳送)

                Dim vFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
                'Dim vWAIVED As String = If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
                If vFILENAME1 <> "" AndAlso flag_NG_UPLOAD_2 Then
                    '符合所有 不可再次傳送 'cst_tpmsg_enb8
                    Common.MessageBox(Me, "已上傳儲存過該文件，不可再次操作!")
                    Return
                End If

                '有錯誤原因 可再次傳送 並記錄 iBCFID
                Dim iBCFID As Integer = If(drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) <> "", TIMS.CINT1(drFL("BCFID")), -1)
                '檔案上傳／確定檔案上傳
                Call FILE_UPLOAD_1(drOB, iBCFID)
        End Select

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)
    End Sub

    Private Sub FILE_UPLOAD_14(drOB As DataRow)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        If drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!(14)")
            Return
        End If
        'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn,RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!(14)")
            Return
        End If
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!(14)")
            Return
        End If

        Dim v_PlanID As String = ""
        Dim v_ComIDNO As String = ""
        Dim v_SeqNo As String = ""
        Hid_PCS14.Value = ""
        For Each eItem As DataGridItem In DataGrid14.Items
            Dim chkItem1 As HtmlInputCheckBox = eItem.FindControl("chkItem1")
            Dim HDG_PCS As HtmlInputHidden = eItem.FindControl("HDG_PCS")
            Dim HDG_PlanID As HtmlInputHidden = eItem.FindControl("HDG_PlanID")
            Dim HDG_ComIDNO As HtmlInputHidden = eItem.FindControl("HDG_ComIDNO")
            Dim HDG_SeqNo As HtmlInputHidden = eItem.FindControl("HDG_SeqNo")
            If chkItem1.Checked AndAlso HDG_PCS.Value <> "" Then
                Hid_PCS14.Value = TIMS.ClearSQM(HDG_PCS.Value)
                v_PlanID = TIMS.ClearSQM(HDG_PlanID.Value)
                v_ComIDNO = TIMS.ClearSQM(HDG_ComIDNO.Value)
                v_SeqNo = TIMS.ClearSQM(HDG_SeqNo.Value)
                Exit For
            End If
        Next
        Hid_PCS14.Value = TIMS.ClearSQM(Hid_PCS14.Value)
        Dim vPCS14 As String = Hid_PCS14.Value
        If vPCS14 = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(未選擇班級)，請重新操作!(14)")
            Return
        ElseIf Hid_PCS14.Value <> String.Concat(v_PlanID, "x", v_ComIDNO, "x", v_SeqNo) Then
            Common.MessageBox(Me, "上傳資訊有誤(班級序號有誤)，請重新操作!(14)")
            Return
        End If
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Exit Sub
        End If

        '年度／計畫ID／機構ID／caseno／1
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_BCASENO.Value, "/")
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vBCID As String = TIMS.ClearSQM(drOB("BCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vBCASENO As String = TIMS.ClearSQM(drOB("BCASENO"))
        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vBCASENO <> Hid_BCASENO.Value Then Return '(此狀況不太可能發生)

        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFilePdf(Me, File1, MyPostedFile) Then Return

        '取出檔案名稱
        Dim MyFileName As String = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, cst_errMsg_4)
            Exit Sub
        End If
        Dim MyFileType As String = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        Select Case LCase(MyFileType)
            Case "pdf"
                If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                    Common.MessageBox(Me, cst_errMsg_7_10M)
                    Exit Sub
                End If
            Case Else
                Common.MessageBox(Me, cst_errMsg_5)
                Exit Sub
        End Select

        '取得KBID代號／非流水號
        Dim vKBSID As String = $"{drKB("KBSID")}"
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        '上傳檔案 '年度／計畫ID／機構ID／caseno／kbsid
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_BI(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_PI(vBCID, vKBSID, vPCS14, "pdf")
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
            'strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Dim rPMS As New Hashtable From {
            {"PCS14", Hid_PCS14.Value},
            {"PlanID", v_PlanID},
            {"ComIDNO", v_ComIDNO},
            {"SeqNo", v_SeqNo},
            {"FILENAME1", vFILENAME1},
            {"SRCFILENAME1", vSRCFILENAME1}
        }
        Call SAVE_ORG_BIDCASEFL_PI3(drOB, drKB, drRR, rPMS)
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vBCFID As String = TIMS.GetMyValue(sCmdArg, "BCFID")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        Dim vRID As String = RIDValue.Value
        Dim vBCID As String = Hid_BCID.Value
        Dim vBCASENO As String = Hid_BCASENO.Value
        If e.CommandArgument = "" OrElse vBCFID = "" Then Return

        Select Case e.CommandName
            Case "DELFILE4"
                Dim sErrMsg1 As String = CHKDEL_ORG_BIDCASEFL(vBCFID)
                If sErrMsg1 <> "" Then
                    Common.MessageBox(Me, sErrMsg1)
                    Return
                End If

                '刪除檔案 '"ORG_BIDCASEFL"
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
                If drOB Is Nothing Then Return
                Dim oYEARS As String = $"{drOB("YEARS")}"
                Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
                Dim oPLANID As String = $"{drOB("PLANID")}"
                Dim oRID As String = $"{drOB("RID")}"
                Dim oBCASENO As String = $"{drOB("BCASENO")}"
                Dim oFILENAME1 As String = ""
                Dim oUploadPath As String = ""
                Dim s_FilePath1 As String = ""
                Try
                    oFILENAME1 = vFILENAME1
                    oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try

                '"ORG_BIDCASEFL"
                Dim dParms As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim rdSql As String = "DELETE ORG_BIDCASEFL WHERE BCFID=@BCFID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)
                'DataGrid1.EditItemIndex = -1

            Case "DOWNLOAD4" '下載
                Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "BCFID", vBCFID)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select

        If Not TIMS.OpenDbConn(objconn) Then Return
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        '顯示檔案資料表
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)

        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Call SHOW_Detail_BIDCASE(drRR, Hid_BCID.Value, Session(cst_ss_RqProcessType))
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim LabdepID As Label = e.Item.FindControl("LabdepID")
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                Dim BTN_DELFILE4 As Button = e.Item.FindControl("BTN_DELFILE4") '刪除
                Dim BTN_DOWNLOAD4 As Button = e.Item.FindControl("BTN_DOWNLOAD4") '下載 
                Dim labRTUREASON As Label = e.Item.FindControl("labRTUREASON") '退件原因
                labRTUREASON.Text = $"{drv("RTUREASON")}"

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = $"{drv("OKFLAG")}"
                    BTN_DOWNLOAD4.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                ElseIf $"{drv("WAIVED")}" = cst_08_訓練班別計畫表_WAIVED_PI Then
                    titleMsg = cst_txt_版本批次送出
                ElseIf $"{drv("WAIVED")}" = cst_08_1_iCap課程原始申請資料_WAIVED_PI3 Then
                    titleMsg = cst_txt_iCap課程原始申請資料
                ElseIf $"{drv("WAIVED")}" = cst_10_師資助教基本資料表_WAIVED_TT Then
                    titleMsg = cst_txt_師資助教基本資料表
                ElseIf $"{drv("WAIVED")}" = cst_11_授課師資學經歷證書影本_WAIVED_TT2 Then
                    titleMsg = cst_txt_授課師資學經歷證書影本
                ElseIf $"{drv("WAIVED")}" = cst_13_教學環境資料表_WAIVED_PI2 Then
                    titleMsg = cst_txt_教學環境資料表
                ElseIf $"{drv("WAIVED")}" = cst_13_1混成課程教學環境資料表_WAIVED_RT2 Then
                    titleMsg = cst_txt_混成課程教學環境資料表
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "BCFID", $"{drv("BCFID")}")
                TIMS.SetMyValue(sCmdArg, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg, "KBSID", $"{drv("KBSID")}")
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                BTN_DELFILE4.CommandArgument = sCmdArg '刪除
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                BTN_DELFILE4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE4.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
                '(其他原因調整) '送件／退件修正，不提供刪除
                If $"{drv("BISTATUS")}" = "B" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb6, True)
                ElseIf $"{drv("BISTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" <> "" Then
                    BTN_DELFILE4.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb8, True)
                ElseIf $"{drv("BISTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" = "" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb7, True)
                End If
        End Select
    End Sub

    ''' <summary>刪除前檢核</summary>
    ''' <param name="vBCFID"></param>
    ''' <returns></returns>
    Private Function CHKDEL_ORG_BIDCASEFL(vBCFID As String) As String
        Dim rst As String = ""
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        If drOB Is Nothing Then
            rst &= "資訊有誤(查無案件編號)，請重新操作!!"
            Return rst
        End If

        Dim sParms1 As New Hashtable From {
            {"BCFID", vBCFID}
        }
        Dim sSql As String = ""
        sSql &= " SELECT a.BCFID,a.YEARS,a.APPSTAGE,a.RID,a.BCID" & vbCrLf
        sSql &= " ,a.KBSID,a.RTUREASON,a.RTURESACCT,a.RTURESDATE" & vbCrLf
        sSql &= " ,kb.ORGKINDGW,kb.KBID,kb.KBNAME" & vbCrLf
        sSql &= " FROM ORG_BIDCASEFL a" & vbCrLf
        sSql &= " JOIN KEY_BIDCASE kb on kb.KBSID=a.KBSID" & vbCrLf
        sSql &= " WHERE a.BCFID=@BCFID" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms1)
        If dr1 Is Nothing Then Return rst

        Dim vORGKINDGW As String = Convert.ToString(dr1("ORGKINDGW"))
        Dim vKBID As String = Convert.ToString(dr1("KBID"))
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_W08_訓練班別計畫表, TIMS.cst_G08_訓練班別計畫表
                'Dim sParms3 As New Hashtable
                'sParms3.Add("BCFID", vBCFID)
                'Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEFL_PI WHERE BCFID=@BCFID"
                'Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                'If dt3.Rows.Count > 0 Then rst &= "該項目，有檔案資訊(子項)，不可刪除"
                'Return rst
                Dim dParms As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim rdSql As String = "DELETE ORG_BIDCASEFL_PI WHERE BCFID=@BCFID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)
            Case TIMS.cst_W08_1_iCap課程原始申請資料, TIMS.cst_G08_1_iCap課程原始申請資料
                Dim sParms3 As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEFL_PI3 WHERE BCFID=@BCFID"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If dt3.Rows.Count > 0 Then rst &= "該項目，有檔案資訊(子項)，不可刪除"

            Case TIMS.cst_W10_師資助教基本資料表, TIMS.cst_G10_師資助教基本資料表
                Dim sParms3 As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEFL_TT WHERE BCFID=@BCFID"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If dt3.Rows.Count > 0 Then rst &= "該項目，有檔案資訊(子項)，不可刪除"

            Case TIMS.cst_W11_授課師資學經歷證書影本, TIMS.cst_G11_授課師資學經歷證書影本
                Dim sParms3 As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEFL_TT2 WHERE BCFID=@BCFID"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If dt3.Rows.Count > 0 Then rst &= "該項目，有檔案資訊(子項)，不可刪除"

            Case TIMS.cst_W13_教學環境資料表, TIMS.cst_G13_教學環境資料表
                Dim sParms3 As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEFL_EV WHERE BCFID=@BCFID"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If dt3.Rows.Count > 0 Then rst &= "該項目，有檔案資訊(子項)，不可刪除"

            Case TIMS.cst_W13_1_混成課程教學環境資料表, TIMS.cst_G13_1_混成課程教學環境資料表
                Dim sParms3 As New Hashtable From {
                    {"BCFID", vBCFID}
                }
                Dim sSql3 As String = "SELECT 1 FROM ORG_BIDCASEFL_RT WHERE BCFID=@BCFID"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If dt3.Rows.Count > 0 Then rst &= "該項目，有檔案資訊(子項)，不可刪除"

        End Select
        Return rst
    End Function

    ''' <summary>最近一次版本送件</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_latestSend1_Click(sender As Object, e As EventArgs) Handles bt_latestSend1.Click
        UTL_LATEST(cst_MTYPE_LATEST_SEND1) '最近一次版本送件
    End Sub

    ''' <summary>最近一次版本-下載</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_latestDown1_Click(sender As Object, e As EventArgs) Handles bt_latestDown1.Click
        UTL_LATEST(cst_MTYPE_LATEST_DOWN1) '最近一次版本-下載
    End Sub

    ''' <summary>最近一次版本-下載2</summary>
    ''' <param name="MTYPE"></param>
    Private Sub UTL_LATEST(MTYPE As String)
        'cst_MTYPE_LATEST_SEND1: 最近一次版本送件 /cst_MTYPE_LATEST_DOWN1: 最近一次版本-下載
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_G01_TTQS評核證書影本, TIMS.cst_W01_TTQS評核證書影本
                Call FILE_COPY1(drOB, drKB, drRR, MTYPE)
            Case TIMS.cst_G02_設立證明文件影本, TIMS.cst_W02_設立登記影本
                Call FILE_COPY1(drOB, drKB, drRR, MTYPE)
            Case TIMS.cst_G03_組織章程影本, TIMS.cst_W03_組織章程影本
                Call FILE_COPY1(drOB, drKB, drRR, MTYPE)
            Case TIMS.cst_G04_法人登記證書, TIMS.cst_G04_法人登記證書
                Call FILE_COPY1(drOB, drKB, drRR, MTYPE)
            Case TIMS.cst_G10_師資助教基本資料表, TIMS.cst_W10_師資助教基本資料表
                Call FILE_COPY1_TT(drOB, drKB, drRR, MTYPE)
            Case TIMS.cst_G11_授課師資學經歷證書影本, TIMS.cst_W11_授課師資學經歷證書影本
                Call FILE_COPY1_TT2(drOB, drKB, drRR, MTYPE)
            Case Else
                Common.MessageBox(Me, String.Concat("(查無資料)!!", vORGKINDGW, vKBID))
                Return
        End Select
    End Sub

    ''' <summary>以目前版本批次送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SENTBATVER_Click(sender As Object, e As EventArgs) Handles BTN_SENTBATVER.Click
        'Dim vUploadPath As String = Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If Hid_BCASENO.Value = "" OrElse Hid_BCID.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(案件號為空)，請重新操作!!")
            Return
        End If
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim vBCID As String = Hid_BCID.Value
        Dim vKBSID As String = Hid_KBSID.Value
        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_W08_訓練班別計畫表, TIMS.cst_G08_訓練班別計畫表
                Dim iBCFID As Integer = TIMS.GET_ORG_BIDCASEFL_iBCFID(sm, objconn, vBCID, vKBSID, cst_08_訓練班別計畫表_WAIVED_PI, drOB)
                If iBCFID <= 0 Then Return

                Dim dtFLPI As DataTable = TIMS.GET_ORG_BIDCASEFL_PI(objconn, vBCID)

                Dim s_TMPMSG1 As String = ""
                Dim CYearsVal As String = (sm.UserInfo.Years - 1911)
                Dim strWinOpenJScriptALL As String = ""
                Dim iRow As Integer = 0
                For Each eItem As DataGridItem In DataGrid08.Items
                    iRow += 1
                    Dim HDG8_PlanID As HtmlInputHidden = eItem.FindControl("HDG8_PlanID")
                    Dim HDG8_ComIDNO As HtmlInputHidden = eItem.FindControl("HDG8_ComIDNO")
                    Dim HDG8_SeqNo As HtmlInputHidden = eItem.FindControl("HDG8_SeqNo")
                    Dim vHDG8_PlanID As String = TIMS.ClearSQM(HDG8_PlanID.Value)
                    Dim vHDG8_ComIDNO As String = TIMS.ClearSQM(HDG8_ComIDNO.Value)
                    Dim vHDG8_SeqNo As String = TIMS.ClearSQM(HDG8_SeqNo.Value)
                    'Dim vPCS As String = String.Concat(vHDG8_PlanID, "x", vHDG8_ComIDNO, "x", vHDG8_SeqNo)

                    Dim iBCPID As Integer = TIMS.GET_ORG_BIDCASEPI_iBCPID(sm, objconn, TIMS.CINT1(vBCID), vHDG8_PlanID, vHDG8_ComIDNO, vHDG8_SeqNo)
                    If iBCPID <= 0 Then Return
                    'Dim iBCFID As Integer = TIMS.GET_ORG_BIDCASEFL_iBCFID(sm, objconn, vBCID, vKBSID, cst_08_訓練班別計畫表_WAIVED_PI, drOB)
                    'If iBCFID <= 0 Then Return
                    Dim rPMS As New Hashtable From {
                        {"ORGKINDGW", drKB("ORGKINDGW")},
                        {"YEARS", drOB("YEARS")},
                        {"APPSTAGE", drOB("APPSTAGE")},
                        {"RID", drOB("RID")},
                        {"BCID", drOB("BCID")},
                        {"KBSID", drKB("KBSID")},
                        {"PlanID", vHDG8_PlanID},
                        {"ComIDNO", vHDG8_ComIDNO},
                        {"SeqNo", vHDG8_SeqNo},
                        {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    Call SAVE_ORG_BIDCASEPI_08(rPMS)

                    Dim tkVal As String = ""
                    TIMS.SetMyValue(tkVal, String.Concat("N1SM", Now.ToString("ss")), Now.ToString("ssmm"))
                    TIMS.SetMyValue(tkVal, "RID", RIDValue.Value)
                    TIMS.SetMyValue(tkVal, "BCID", Hid_BCID.Value)
                    TIMS.SetMyValue(tkVal, "BCASENO", Hid_BCASENO.Value)
                    TIMS.SetMyValue(tkVal, "ORGKINDGW", Hid_ORGKINDGW.Value)
                    TIMS.SetMyValue(tkVal, "KBSID", Hid_KBSID.Value)

                    Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
                    tryFIND = String.Concat("BCPID=", iBCPID, " AND PlanID=", vHDG8_PlanID, " AND ComIDNO='", vHDG8_ComIDNO, "' AND SeqNo=", vHDG8_SeqNo)
                    If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 AndAlso dtFLPI.Select(tryFIND).Length > 0 Then
                        Dim drFLPI As DataRow = dtFLPI.Select(tryFIND)(0)
                        Dim vMODIFY_DAY As String = Convert.ToString(drFLPI("MODIFY_DAY")) 'MODIFY_DAY
                        Dim vMODIFY_MI As String = Convert.ToString(drFLPI("MODIFY_MI")) 'MODIFY_MI
                        fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
                    End If

                    If fg_RUN_REPORT_1 Then
                        Dim sCmdArg As String = ""
                        TIMS.SetMyValue(sCmdArg, "Type", "B") 'Type: A:已轉班查詢 B:未轉班查詢
                        TIMS.SetMyValue(sCmdArg, "PrintOrg", "Y") '顯示訓練單位名稱 Y/M
                        TIMS.SetMyValue(sCmdArg, "Years", CYearsVal)
                        TIMS.SetMyValue(sCmdArg, "PlanID", vHDG8_PlanID)
                        TIMS.SetMyValue(sCmdArg, "ComIDNO", vHDG8_ComIDNO)
                        TIMS.SetMyValue(sCmdArg, "SeqNo", vHDG8_SeqNo)
                        TIMS.SetMyValue(sCmdArg, "FTYPE", "2") '1:細明體/2:標楷體(def)
                        TIMS.SetMyValue(sCmdArg, "PDFOUT", "YB") '以目前版本批次送出
                        TIMS.SetMyValue(sCmdArg, "tk", TIMS.EncryptAes(tkVal))
                        'TIMS.SetMyValue(sCmdArg, "MSD", v_MSD)
                        'Dim ock_Value1 As String = String.Concat("window.open('", sPrintASPX1, sCmdArg, "','','resizable=yes,toolbar=no,scrollbars=yes');")
                        strWinOpenJScriptALL &= ReportQuery.strWOScriptC(String.Concat(sPrintASPX1, sCmdArg))
                    Else
                        s_TMPMSG1 &= String.Concat(If(s_TMPMSG1 <> "", ", ", ""), iRow)
                    End If
                Next
                'ReportQuery.strWOScript(strWinOpen)
                If strWinOpenJScriptALL <> "" Then
                    Dim strScript As String = String.Concat("<script language=""javascript"">", strWinOpenJScriptALL, "</script>")
                    RegisterStartupScript(TIMS.xBlockName(), strScript)
                End If
                If s_TMPMSG1 <> "" Then
                    Common.MessageBox(Me, String.Concat("(部份) 資料表重複處理時間過短(3分鐘1次)，請等待3分鐘後再試!", vbCrLf, s_TMPMSG1))
                    'Return
                End If
            Case Else
                Dim rParms2 As New Hashtable From {{"BCID", TIMS.CINT1(Hid_BCID.Value)}, {"KBSID", TIMS.CINT1(Hid_KBSID.Value)}}
                Dim rSql2 As String = "SELECT 1 FROM ORG_BIDCASEFL WHERE BCID=@BCID AND KBSID=@KBSID"
                Dim drFL2 As DataRow = DbAccess.GetOneRow(rSql2, objconn, rParms2)
                If drFL2 IsNot Nothing Then
                    Common.MessageBox(Me, "已儲存過該文件，不可再次操作!")
                    Return
                End If
        End Select

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)
    End Sub

    Private Sub DataGrid08_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid08.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        Dim vBCFID As String = TIMS.GetMyValue(sCmdArg, "BCFID")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vBCFPID As String = TIMS.GetMyValue(sCmdArg, "BCFPID")
        Select Case e.CommandName
            Case "DOWNLOAD8"
                Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "BCFID", vBCFID)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "BCFPID", vBCFPID)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select
    End Sub

    Private Sub DataGrid08_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid08.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG8_OCID As HtmlInputHidden = e.Item.FindControl("HDG8_OCID")
                Dim HDG8_PlanID As HtmlInputHidden = e.Item.FindControl("HDG8_PlanID")
                Dim HDG8_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG8_ComIDNO")
                Dim HDG8_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG8_SeqNo")
                Dim HDG8_PrintRpt1 As HtmlInputButton = e.Item.FindControl("HDG8_PrintRpt1")
                Dim BTN_DOWNLOAD8 As Button = e.Item.FindControl("BTN_DOWNLOAD8") '下載 

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()

                '0:未轉班,1:已轉班 '未轉班(依計畫查詢) 含列印
                HDG8_PrintRpt1.Visible = True
                HDG8_PlanID.Value = $"{drv("PlanID")}"
                HDG8_ComIDNO.Value = $"{drv("ComIDNO")}"
                HDG8_SeqNo.Value = $"{drv("SeqNo")}"
                HDG8_OCID.Value = $"{drv("OCID")}"
                'Dim v_MSD As String = $"{drv("MSD")}"
                Dim CYearsVal As String = (sm.UserInfo.Years - 1911)
                'Dim sUrl As String = String.Concat(If(iPYNum >= 3, cst_printASPX_R, cst_printASPX_Q), Request("ID"))
                Dim tkVal As String = ""
                TIMS.SetMyValue(tkVal, String.Concat("N1SM", Now.ToString("ss")), Now.ToString("ssmm"))
                TIMS.SetMyValue(tkVal, "RID", RIDValue.Value)
                TIMS.SetMyValue(tkVal, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue(tkVal, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue(tkVal, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue(tkVal, "KBSID", Hid_KBSID.Value)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "Type", "B") 'Type: A:已轉班查詢 B:未轉班查詢
                TIMS.SetMyValue(sCmdArg, "PrintOrg", "Y") '顯示訓練單位名稱 Y/M
                TIMS.SetMyValue(sCmdArg, "Years", CYearsVal)
                TIMS.SetMyValue(sCmdArg, "PlanID", HDG8_PlanID.Value)
                TIMS.SetMyValue(sCmdArg, "ComIDNO", HDG8_ComIDNO.Value)
                TIMS.SetMyValue(sCmdArg, "SeqNo", HDG8_SeqNo.Value)
                TIMS.SetMyValue(sCmdArg, "FTYPE", "2") '1:細明體/2:標楷體(def)
                TIMS.SetMyValue(sCmdArg, "PDFOUT", "Y") '單一輸出
                TIMS.SetMyValue(sCmdArg, "tk", TIMS.EncryptAes(tkVal))
                'TIMS.SetMyValue(sCmdArg, "MSD", v_MSD)
                Dim ock_Value1 As String = String.Concat("window.open('", sPrintASPX1, sCmdArg, "','','resizable=yes,toolbar=no,scrollbars=yes');")
                '20090107 andy edit 報表改為網頁產生的方式
                HDG8_PrintRpt1.Attributes.Add("onclick", ock_Value1)

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = $"{drv("OKFLAG")}"
                    BTN_DOWNLOAD8.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                Else
                    titleMsg = cst_tpmsg_enb9
                    BTN_DOWNLOAD8.Enabled = False
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD8, titleMsg, True)

                Dim sCmdArg8 As String = ""
                TIMS.SetMyValue(sCmdArg8, "BCFID", $"{drv("BCFID")}")
                TIMS.SetMyValue(sCmdArg8, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg8, "KBSID", $"{drv("KBSID")}")
                TIMS.SetMyValue(sCmdArg8, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg8, "BCFPID", $"{drv("BCFPID")}")
                BTN_DOWNLOAD8.CommandArgument = sCmdArg8 '檔案下載
        End Select
    End Sub

    Private Sub DataGrid10_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid10.ItemCommand
        'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vTECHID As String = TIMS.GetMyValue(sCmdArg, "TECHID")
        Dim vBCFTID As String = TIMS.GetMyValue(sCmdArg, "BCFTID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vRID As String = RIDValue.Value
        Dim vBCID As String = Hid_BCID.Value
        Dim vBCASENO As String = Hid_BCASENO.Value
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Select Case e.CommandName
            Case "REPORT10"
                If e.CommandArgument = "" OrElse vTECHID = "" Then Return

                Dim rPMS As New Hashtable
                rPMS.Clear()
                rPMS.Add("TechID", vTECHID)
                rPMS.Add("Years", drRR("YEARS"))
                rPMS.Add("Title", Convert.ToString(drRR("ORGPLANNAME")))
                Call RPT_SD_14_004(rPMS)

            Case "DOWNLOAD10"
                'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                'Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                'Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "BCFTID", vBCFTID)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
            Case "DELFILE10"
                If e.CommandArgument = "" OrElse vBCFTID = "" Then Return

                '刪除檔案 '"ORG_BIDCASEFL"
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
                If drOB Is Nothing Then Return
                Dim oYEARS As String = $"{drOB("YEARS")}"
                Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
                Dim oPLANID As String = $"{drOB("PLANID")}"
                Dim oRID As String = $"{drOB("RID")}"
                Dim oBCASENO As String = $"{drOB("BCASENO")}"
                Dim dtFLTT As DataTable = TIMS.GET_ORG_BIDCASEFL_TT(objconn, vBCID)
                tryFIND = If(vBCFTID <> "" AndAlso vFILENAME1 <> "", String.Concat("BCFTID=", vBCFTID, " AND FILENAME1='", vFILENAME1, "'"), "")
                If dtFLTT IsNot Nothing AndAlso dtFLTT.Rows.Count > 0 AndAlso dtFLTT.Select(tryFIND).Length > 0 Then
                    For Each drFLTT As DataRow In dtFLTT.Select(tryFIND)
                        Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                        Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        Try
                            oFILENAME1 = Convert.ToString(drFLTT("FILENAME1"))
                            oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            Call TIMS.MyFileDelete(s_FilePath1)
                        Catch ex As Exception
                            Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                            strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                            Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                        End Try
                    Next
                End If

                Dim dParms As New Hashtable From {
                    {"BCFTID", vBCFTID}
                }
                Dim rdSql As String = "DELETE ORG_BIDCASEFL_TT WHERE BCFTID=@BCFTID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)
                'DataGrid1.EditItemIndex = -1

                '師資／助教基本資料表
                Call SHOW_DATAGRID_10()
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
                HDG10_TechID.Value = $"{drv("TechID")}"
                HDG10_RID.Value = $"{drv("RID")}"

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = $"{drv("OKFLAG")}"
                    BTN_DOWNLOAD10.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = "Y" Then
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
                TIMS.SetMyValue(sCmdArg, "BCFTID", drv("BCFTID"))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg, "KBSID", $"{drv("KBSID")}")
                BTN_REPORT10.CommandArgument = sCmdArg '報表下載
                BTN_DOWNLOAD10.CommandArgument = sCmdArg '檔案下載
                BTN_DELFILE10.CommandArgument = sCmdArg '刪除檔案
                BTN_DELFILE10.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE10.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
                '(其他原因調整) '送件／退件修正，不提供刪除
                If Hid_BISTATUS.Value = "B" Then
                    BTN_DELFILE10.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE10, cst_tpmsg_enb6, True)
                ElseIf Hid_BISTATUS.Value = "R" AndAlso Hid_RTUREASON.Value <> "" Then
                    BTN_DELFILE10.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE10, cst_tpmsg_enb8, True)
                ElseIf Hid_BISTATUS.Value = "R" AndAlso Hid_RTUREASON.Value = "" Then
                    BTN_DELFILE10.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE10, cst_tpmsg_enb7, True)
                End If
        End Select
    End Sub

    Private Sub DataGrid11_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid11.ItemCommand
        'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vTECHID As String = TIMS.GetMyValue(sCmdArg, "TECHID")
        Dim vBCFT2ID As String = TIMS.GetMyValue(sCmdArg, "BCFT2ID")
        If e.CommandArgument = "" OrElse vBCFT2ID = "" Then Return '(程式有誤中斷執行)
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vRID As String = RIDValue.Value
        Dim vBCID As String = Hid_BCID.Value
        Dim vBCASENO As String = Hid_BCASENO.Value

        Select Case e.CommandName
            Case "DOWNLOAD11"
                If e.CommandArgument = "" OrElse vBCFT2ID = "" Then Return
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "BCFT2ID", vBCFT2ID)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return

            Case "DELFILE11"
                If e.CommandArgument = "" OrElse vBCFT2ID = "" Then Return

                '刪除檔案  
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
                If drOB Is Nothing Then Return
                Dim oYEARS As String = $"{drOB("YEARS")}"
                Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
                Dim oPLANID As String = $"{drOB("PLANID")}"
                Dim oRID As String = $"{drOB("RID")}"
                Dim oBCASENO As String = $"{drOB("BCASENO")}"
                Dim dtFLTT2 As DataTable = TIMS.GET_ORG_BIDCASEFL_TT2(objconn, vBCID)
                tryFIND = If(vBCFT2ID <> "" AndAlso vFILENAME1 <> "", String.Concat("BCFT2ID=", vBCFT2ID, " AND FILENAME1='", vFILENAME1, "'"), "")
                If dtFLTT2 IsNot Nothing AndAlso dtFLTT2.Rows.Count > 0 AndAlso dtFLTT2.Select(tryFIND).Length > 0 Then
                    For Each drFLTT2 As DataRow In dtFLTT2.Select(tryFIND)
                        Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                        Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        Try
                            oFILENAME1 = Convert.ToString(drFLTT2("FILENAME1"))
                            oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            Call TIMS.MyFileDelete(s_FilePath1)
                        Catch ex As Exception
                            Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                            strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                            Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                        End Try
                    Next
                End If

                '"ORG_BIDCASEFL"
                Dim dParms As New Hashtable From {
                    {"BCFT2ID", vBCFT2ID}
                }
                Dim rdSql As String = "DELETE ORG_BIDCASEFL_TT2 WHERE BCFT2ID=@BCFT2ID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)

                'DataGrid1.EditItemIndex = -1
                Call SHOW_DATAGRID_11()
        End Select
    End Sub

    Private Sub DataGrid11_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid11.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim chkItem1 As HtmlInputCheckBox = e.Item.FindControl("chkItem1")
                Dim HDG11_TechID As HtmlInputHidden = e.Item.FindControl("HDG11_TechID")
                Dim HDG11_RID As HtmlInputHidden = e.Item.FindControl("HDG11_RID")

                e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e)

                chkItem1.Attributes("onclick") = String.Concat("selectOnlyThis('", chkItem1.ClientID, "',", iDG11_ROWS, ",'DataGrid11')")
                '0:未轉班,1:已轉班 '未轉班(依計畫查詢) 含列印
                HDG11_TechID.Value = $"{drv("TechID")}"
                HDG11_RID.Value = $"{drv("RID")}"

                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                Dim BTN_DELFILE11 As Button = e.Item.FindControl("BTN_DELFILE11")
                Dim BTN_DOWNLOAD11 As Button = e.Item.FindControl("BTN_DOWNLOAD11")

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    BTN_DOWNLOAD11.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    'LabFileName1.Text = cst_txt_免附文件
                    titleMsg = cst_txt_免附文件
                Else
                    titleMsg = cst_tpmsg_enb9
                    BTN_DOWNLOAD11.Enabled = False
                    BTN_DELFILE11.Enabled = False
                    Call TIMS.Tooltip(BTN_DELFILE11, cst_tpmsg_enb9, True)
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD11, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TECHID", $"{drv("TECHID")}")
                TIMS.SetMyValue(sCmdArg, "BCFT2ID", $"{drv("BCFT2ID")}")
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg, "KBSID", $"{drv("KBSID")}")
                BTN_DELFILE11.CommandArgument = sCmdArg '刪除檔案
                BTN_DOWNLOAD11.CommandArgument = sCmdArg '檔案下載
                BTN_DELFILE11.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE11.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
                '(其他原因調整) '送件／退件修正，不提供刪除
                If Hid_BISTATUS.Value = "B" Then
                    BTN_DELFILE11.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE11, cst_tpmsg_enb6, True)
                ElseIf Hid_BISTATUS.Value = "R" AndAlso Hid_RTUREASON.Value <> "" Then
                    BTN_DELFILE11.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE11, cst_tpmsg_enb8, True)
                ElseIf Hid_BISTATUS.Value = "R" AndAlso Hid_RTUREASON.Value = "" Then
                    BTN_DELFILE11.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE11, cst_tpmsg_enb7, True)
                End If

        End Select
    End Sub

    ''' <summary>以目前版本送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SENDCURRVER_Click(sender As Object, e As EventArgs) Handles BTN_SENDCURRVER.Click
        'Dim vUploadPath As String = Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.ClearSQM(Hid_KBSID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If Hid_BCASENO.Value = "" OrElse Hid_BCID.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(案件號為空)，請重新操作!!")
            Return
        End If
        Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(sm, objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim vKBID As String = $"{drKB("KBID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Select Case String.Concat(vORGKINDGW, vKBID)
            Case TIMS.cst_W13_教學環境資料表, TIMS.cst_G13_教學環境資料表
                'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
                'If drOB Is Nothing Then Return
                'Dim drKB As DataRow = TIMS.GET_KEY_BIDCASE(objconn, Hid_KBSID.Value, Hid_ORGKINDGW.Value)
                'If drKB Is Nothing Then Return
                Dim vBCID As String = $"{drOB("BCID")}"
                Dim vKBSID As String = $"{drKB("KBSID")}"
                Dim vBCASENO As String = $"{drOB("BCASENO")}"
                Dim iBCFID As Integer = TIMS.GET_ORG_BIDCASEFL_iBCFID(sm, objconn, TIMS.CINT1(vBCID), TIMS.CINT1(vKBSID), cst_13_教學環境資料表_WAIVED_PI2, drOB)
                If iBCFID <= 0 Then Return

                Dim rPMS As New Hashtable From {
                    {"ORGKINDGW", drKB("ORGKINDGW")},
                    {"YEARS", drOB("YEARS")},
                    {"APPSTAGE", drOB("APPSTAGE")},
                    {"RID", drOB("RID")},
                    {"BCID", TIMS.CINT1(vBCID)},
                    {"BCASENO", vBCASENO},
                    {"KBSID", TIMS.CINT1(vKBSID)},
                    {"BCFID", iBCFID},
                    {"MODIFYACCT", sm.UserInfo.UserID}
                }
                Call SAVE_ORG_BIDCASE_ALL_13(rPMS)

            Case TIMS.cst_W13_1_混成課程教學環境資料表, TIMS.cst_G13_1_混成課程教學環境資料表
                Dim vBCID As String = $"{drOB("BCID")}"
                Dim vKBSID As String = $"{drKB("KBSID")}"
                Dim vBCASENO As String = $"{drOB("BCASENO")}"
                Dim iBCFID As Integer = TIMS.GET_ORG_BIDCASEFL_iBCFID(sm, objconn, TIMS.CINT1(vBCID), TIMS.CINT1(vKBSID), cst_13_1混成課程教學環境資料表_WAIVED_RT2, drOB)
                If iBCFID <= 0 Then Return

                Dim rPMS As New Hashtable From {
                    {"ORGKINDGW", drKB("ORGKINDGW")},
                    {"YEARS", drOB("YEARS")},
                    {"APPSTAGE", drOB("APPSTAGE")},
                    {"RID", drOB("RID")},
                    {"BCID", TIMS.CINT1(vBCID)},
                    {"BCASENO", vBCASENO},
                    {"KBSID", TIMS.CINT1(vKBSID)},
                    {"BCFID", iBCFID},
                    {"MODIFYACCT", sm.UserInfo.UserID}
                }
                Call SAVE_ORG_BIDCASE_ALL_13B(rPMS)

            Case Else
                Dim rParms2 As New Hashtable From {{"BCID", TIMS.CINT1(Hid_BCID.Value)}, {"KBSID", TIMS.CINT1(Hid_KBSID.Value)}}
                Dim rSql2 As String = "SELECT 1 FROM ORG_BIDCASEFL WHERE BCID=@BCID AND KBSID=@KBSID"
                Dim drFL2 As DataRow = DbAccess.GetOneRow(rSql2, objconn, rParms2)
                If drFL2 IsNot Nothing Then
                    Common.MessageBox(Me, "已儲存過該文件，不可再次操作!!")
                    Return
                End If
        End Select

        Threading.Thread.Sleep(10)

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)
        'Call MOVE_PREV()
    End Sub

    Private Sub DataGrid13_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid13.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
        Dim vSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")
        If e.CommandArgument = "" OrElse vPlanID = "" OrElse vComIDNO = "" OrElse vSeqNo = "" Then Return '(程式有誤中斷執行)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vBCFEID As String = TIMS.GetMyValue(sCmdArg, "BCFEID")
        Select Case e.CommandName
            Case "REPORT13"
                '"13" '教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'PCSVALUE  'Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value) 
                'Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
                Dim rPMS As New Hashtable
                rPMS.Clear()
                rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("selsqlstr", String.Concat(vPlanID, "-", vComIDNO, "-", vSeqNo))
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                'W13_教學環境資料表
                Call RPT_SD_14_014(rPMS)

            Case "DOWNLOAD13"
                'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                'Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                'Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "BCFEID", vBCFEID)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select
    End Sub

    Private Sub DataGrid13_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid13.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG_PlanID As HtmlInputHidden = e.Item.FindControl("HDG_PlanID")
                Dim HDG_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG_ComIDNO")
                Dim HDG_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG_SeqNo")
                Dim BTN_REPORT13 As Button = e.Item.FindControl("BTN_REPORT13") '報表下載
                Dim BTN_DOWNLOAD13 As Button = e.Item.FindControl("BTN_DOWNLOAD13") '檔案下載
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                HDG_PlanID.Value = $"{drv("PlanID")}"
                HDG_ComIDNO.Value = $"{drv("ComIDNO")}"
                HDG_SeqNo.Value = $"{drv("SeqNo")}"

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    BTN_DOWNLOAD13.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    'LabFileName1.Text = cst_txt_免附文件
                    titleMsg = cst_txt_免附文件
                Else
                    titleMsg = cst_tpmsg_enb9
                    BTN_DOWNLOAD13.Enabled = False
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD13, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", $"{drv("PlanID")}")
                TIMS.SetMyValue(sCmdArg, "ComIDNO", $"{drv("ComIDNO")}")
                TIMS.SetMyValue(sCmdArg, "SeqNo", $"{drv("SeqNo")}")
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg, "KBSID", $"{drv("KBSID")}")
                TIMS.SetMyValue(sCmdArg, "BCFEID", $"{drv("BCFEID")}")
                TIMS.SetMyValue(sCmdArg, "CLASSCNAMEX", $"{drv("CLASSCNAMEX")}")
                BTN_REPORT13.CommandArgument = sCmdArg '報表下載
                BTN_DOWNLOAD13.CommandArgument = sCmdArg '檔案下載
        End Select
    End Sub

    ''' <summary>重新查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH2_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH2.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_KBSID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_KBSID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKBSID.Value <> "" Then
            Call SHOW_BIDCASE_KBSID(Hid_FirstKBSID.Value, Hid_ORGKINDGW.Value)
        End If
        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "BCID", Hid_BCID.Value)
        Call SHOW_BIDCASEFL_DG2(rPMS3)
    End Sub

    ''' <summary>帶入機構與業務代碼</summary>
    ''' <param name="drRR"></param>
    Sub SHOW_RIDValue_DATA(ByRef drRR As DataRow)
        If drRR Is Nothing Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Return
        ElseIf RIDValue.Value <> TIMS.ClearSQM(drRR("RID")) Then
            center.Text = TIMS.ClearSQM(drRR("OrgName")) 'OrgName
            RIDValue.Value = TIMS.ClearSQM(drRR("RID")) 'sm.UserInfo.RID
            Return
        End If
    End Sub

    Public Shared Sub DELETE_Detail_BIDCASE(MyPage As Page, oConn As SqlConnection, drRR As DataRow, drOB As DataRow)
        '訓練機構有誤
        If drRR Is Nothing OrElse drOB Is Nothing Then Return
        'Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
        'If drOB Is Nothing Then Return
        Dim vBCID As String = $"{drOB("BCID")}"
        Dim oYEARS As String = $"{drOB("YEARS")}"
        Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
        Dim oPLANID As String = $"{drOB("PLANID")}"
        Dim oRID As String = $"{drOB("RID")}"
        Dim oBCASENO As String = $"{drOB("BCASENO")}"

        '刪除檔案  
        Dim dtFLRT As DataTable = TIMS.GET_ORG_BIDCASEFL_RT(oConn, vBCID)
        If dtFLRT IsNot Nothing AndAlso dtFLRT.Rows.Count > 0 Then
            For Each drFLRT As DataRow In dtFLRT.Rows
                Dim oFILENAME1 As String = Convert.ToString(drFLRT("FILENAME1"))
                If oFILENAME1 = "" Then Continue For
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Dim vKBSID As String = Convert.ToString(drFLRT("KBSID"))
                Try
                    oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        '刪除檔案  
        Dim dtFLEV As DataTable = TIMS.GET_ORG_BIDCASEFL_EV(oConn, vBCID)
        If dtFLEV IsNot Nothing AndAlso dtFLEV.Rows.Count > 0 Then
            For Each drFLEV As DataRow In dtFLEV.Rows
                Dim oFILENAME1 As String = Convert.ToString(drFLEV("FILENAME1"))
                If oFILENAME1 = "" Then Continue For
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Dim vKBSID As String = Convert.ToString(drFLEV("KBSID"))
                Try
                    oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        '刪除檔案  
        Dim dtFLTT2 As DataTable = TIMS.GET_ORG_BIDCASEFL_TT2(oConn, vBCID)
        If dtFLTT2 IsNot Nothing AndAlso dtFLTT2.Rows.Count > 0 Then
            For Each drFLTT2 As DataRow In dtFLTT2.Rows
                Dim oFILENAME1 As String = Convert.ToString(drFLTT2("FILENAME1"))
                If oFILENAME1 = "" Then Continue For
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Dim vKBSID As String = Convert.ToString(drFLTT2("KBSID"))
                Try
                    oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        '刪除檔案  
        Dim dtFLTT As DataTable = TIMS.GET_ORG_BIDCASEFL_TT(oConn, vBCID)
        If dtFLTT IsNot Nothing AndAlso dtFLTT.Rows.Count > 0 Then
            For Each drFLTT As DataRow In dtFLTT.Rows
                Dim oFILENAME1 As String = Convert.ToString(drFLTT("FILENAME1"))
                If oFILENAME1 = "" Then Continue For
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Dim vKBSID As String = Convert.ToString(drFLTT("KBSID"))
                Try
                    oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        '刪除檔案
        Dim dtFLPI As DataTable = TIMS.GET_ORG_BIDCASEFL_PI(oConn, vBCID)
        If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 Then
            For Each drFLPI As DataRow In dtFLPI.Rows
                Dim oFILENAME1 As String = Convert.ToString(drFLPI("FILENAME1"))
                If oFILENAME1 = "" Then Continue For
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Dim vKBSID As String = Convert.ToString(drFLPI("KBSID"))
                Try
                    oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        '刪除檔案
        Dim dtFL As DataTable = TIMS.GET_ORG_BIDCASEFL(oConn, vBCID)
        For Each drFL As DataRow In dtFL.Rows
            Dim oFILENAME1 As String = Convert.ToString(drFL("FILENAME1"))
            If oFILENAME1 = "" Then Continue For
            Dim oUploadPath As String = ""
            Dim s_FilePath1 As String = ""
            Try
                oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Call TIMS.MyFileDelete(s_FilePath1)
            Catch ex As Exception
                Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
            End Try
        Next

        'Dim vRID As String = Convert.ToString(drRR("RID"))
        'Dim vPLANID As String = Convert.ToString(drRR("PLANID"))
        'Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))
        Dim dParms As New Hashtable From {{"BCID", vBCID}}
        Dim dsSql As String = ""
        dsSql = "DELETE ORG_BIDCASE WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEPI WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEFL WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEFL_RT WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEFL_EV WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEFL_PI WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEFL_TT WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE ORG_BIDCASEFL_TT2 WHERE BCID=@BCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
    End Sub

    Private Sub DataGrid14_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid14.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim chkItem1 As HtmlInputCheckBox = e.Item.FindControl("chkItem1")
                Dim HDG_PCS As HtmlInputHidden = e.Item.FindControl("HDG_PCS")
                Dim HDG_PlanID As HtmlInputHidden = e.Item.FindControl("HDG_PlanID")
                Dim HDG_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG_ComIDNO")
                Dim HDG_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG_SeqNo")

                e.Item.Cells(1).Text = TIMS.Get_DGSeqNo(sender, e)
                chkItem1.Attributes("onclick") = String.Concat("selectOnlyThis('", chkItem1.ClientID, "',", iDG14_ROWS, ",'DataGrid14')")

                '0:未轉班,1:已轉班 '未轉班(依計畫查詢) 含列印
                'HDG_PCS.Value = $"{drv("PCS")}"
                HDG_PlanID.Value = $"{drv("PlanID")}"
                HDG_ComIDNO.Value = $"{drv("ComIDNO")}"
                HDG_SeqNo.Value = $"{drv("SeqNo")}"
                HDG_PCS.Value = String.Concat(HDG_PlanID.Value, "x", HDG_ComIDNO.Value, "x", HDG_SeqNo.Value)
                Dim BTN_DELFILE14 As Button = e.Item.FindControl("BTN_DELFILE14")
                Dim BTN_DOWNLOAD14 As Button = e.Item.FindControl("BTN_DOWNLOAD14")
                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    BTN_DOWNLOAD14.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    'LabFileName1.Text = cst_txt_免附文件
                    titleMsg = cst_txt_免附文件
                Else
                    titleMsg = cst_tpmsg_enb9
                    BTN_DELFILE14.Enabled = False
                    BTN_DOWNLOAD14.Enabled = False
                    Call TIMS.Tooltip(BTN_DELFILE14, cst_tpmsg_enb9, True)
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD14, titleMsg, True)

                Dim sCmdArg As String = "" 'BCFP3ID
                TIMS.SetMyValue(sCmdArg, "BCFP3ID", $"{drv("BCFP3ID")}")
                TIMS.SetMyValue(sCmdArg, "PlanID", $"{drv("PlanID")}")
                TIMS.SetMyValue(sCmdArg, "ComIDNO", $"{drv("ComIDNO")}")
                TIMS.SetMyValue(sCmdArg, "SeqNo", $"{drv("SeqNo")}")
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg, "KBSID", $"{drv("KBSID")}")
                BTN_DELFILE14.CommandArgument = sCmdArg '刪除檔案
                BTN_DOWNLOAD14.CommandArgument = sCmdArg '檔案下載

                BTN_DELFILE14.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE14.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
                '(其他原因調整) '送件／退件修正，不提供刪除
                If Hid_BISTATUS.Value = "B" Then
                    BTN_DELFILE14.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE14, cst_tpmsg_enb6, True)
                ElseIf Hid_BISTATUS.Value = "R" AndAlso Hid_RTUREASON.Value <> "" Then
                    BTN_DELFILE14.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE14, cst_tpmsg_enb8, True)
                ElseIf Hid_BISTATUS.Value = "R" AndAlso Hid_RTUREASON.Value = "" Then
                    BTN_DELFILE14.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE14, cst_tpmsg_enb7, True)
                End If
        End Select
    End Sub

    Private Sub DataGrid14_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid14.ItemCommand
        'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vBCFP3ID As String = TIMS.GetMyValue(sCmdArg, "BCFP3ID")
        If e.CommandArgument = "" OrElse vBCFP3ID = "" Then Return '(程式有誤中斷執行)
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vRID As String = RIDValue.Value
        Dim vBCID As String = Hid_BCID.Value
        Dim vBCASENO As String = Hid_BCASENO.Value

        If e.CommandArgument = "" OrElse vBCFP3ID = "" Then Return
        Select Case e.CommandName
            Case "DOWNLOAD14"
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "BCFP3ID", vBCFP3ID)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
            Case "DELFILE14"
                '刪除檔案  
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, vRID, vBCID, vBCASENO)
                If drOB Is Nothing Then Return
                Dim oYEARS As String = $"{drOB("YEARS")}"
                Dim oAPPSTAGE As String = $"{drOB("APPSTAGE")}"
                Dim oPLANID As String = $"{drOB("PLANID")}"
                Dim oRID As String = $"{drOB("RID")}"
                Dim oBCASENO As String = $"{drOB("BCASENO")}"

                Dim dtFLPI3 As DataTable = TIMS.GET_ORG_BIDCASEFL_PI3(objconn, vBCID)
                tryFIND = If(vBCFP3ID <> "" AndAlso vFILENAME1 <> "", String.Concat("BCFP3ID=", vBCFP3ID, " AND FILENAME1='", vFILENAME1, "'"), "")
                If dtFLPI3 IsNot Nothing AndAlso dtFLPI3.Rows.Count > 0 AndAlso dtFLPI3.Select(tryFIND).Length > 0 Then
                    For Each drFLTT2 As DataRow In dtFLPI3.Select(tryFIND)
                        Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                        Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                        Try
                            oFILENAME1 = Convert.ToString(drFLTT2("FILENAME1"))
                            oUploadPath = TIMS.GET_UPLOADPATH1_BI(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                            s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                            Call TIMS.MyFileDelete(s_FilePath1)
                        Catch ex As Exception
                            Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                            strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                            strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                            Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                        End Try
                    Next
                End If

                '"ORG_BIDCASEFL"
                Dim dParms As New Hashtable From {
                    {"BCFP3ID", vBCFP3ID}
                }
                Dim rdSql As String = "DELETE ORG_BIDCASEFL_PI3 WHERE BCFP3ID=@BCFP3ID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)

                'DataGrid1.EditItemIndex = -1
                Call SHOW_DATAGRID_14()
        End Select
    End Sub

    Private Sub DataGrid13B_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid13B.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim HDG_PlanID As HtmlInputHidden = e.Item.FindControl("HDG_PlanID")
                Dim HDG_ComIDNO As HtmlInputHidden = e.Item.FindControl("HDG_ComIDNO")
                Dim HDG_SeqNo As HtmlInputHidden = e.Item.FindControl("HDG_SeqNo")
                Dim BTN_REPORT13B As Button = e.Item.FindControl("BTN_REPORT13B") '報表下載
                Dim BTN_DOWNLOAD13B As Button = e.Item.FindControl("BTN_DOWNLOAD13B") '檔案下載

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                HDG_PlanID.Value = $"{drv("PlanID")}"
                HDG_ComIDNO.Value = $"{drv("ComIDNO")}"
                HDG_SeqNo.Value = $"{drv("SeqNo")}"

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    'HFileName.Value = $"{drv("FILENAME1")}" '.ToString()
                    titleMsg = If($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}", $"{drv("FILENAME1")}", $"{drv("OKFLAG")}")
                    BTN_DOWNLOAD13B.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    'LabFileName1.Text = cst_txt_免附文件
                    titleMsg = cst_txt_免附文件
                Else
                    titleMsg = cst_tpmsg_enb9
                    BTN_DOWNLOAD13B.Enabled = False
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD13B, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PlanID", $"{drv("PlanID")}")
                TIMS.SetMyValue(sCmdArg, "ComIDNO", $"{drv("ComIDNO")}")
                TIMS.SetMyValue(sCmdArg, "SeqNo", $"{drv("SeqNo")}")
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "KBID", $"{drv("KBID")}")
                TIMS.SetMyValue(sCmdArg, "KBSID", $"{drv("KBSID")}")
                TIMS.SetMyValue(sCmdArg, "BCRTID", $"{drv("BCRTID")}")
                TIMS.SetMyValue(sCmdArg, "CLASSCNAMEX", $"{drv("CLASSCNAMEX")}")
                BTN_REPORT13B.CommandArgument = sCmdArg '報表下載
                BTN_DOWNLOAD13B.CommandArgument = sCmdArg '檔案下載
        End Select
    End Sub

    Private Sub DataGrid13B_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid13B.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vPlanID As String = TIMS.GetMyValue(sCmdArg, "PlanID")
        Dim vComIDNO As String = TIMS.GetMyValue(sCmdArg, "ComIDNO")
        Dim vSeqNo As String = TIMS.GetMyValue(sCmdArg, "SeqNo")
        If e.CommandArgument = "" OrElse vPlanID = "" OrElse vComIDNO = "" OrElse vSeqNo = "" Then Return '(程式有誤中斷執行)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vKBID As String = TIMS.GetMyValue(sCmdArg, "KBID")
        Dim vKBSID As String = TIMS.GetMyValue(sCmdArg, "KBSID")
        Dim vBCRTID As String = TIMS.GetMyValue(sCmdArg, "BCRTID")
        Select Case e.CommandName
            Case "REPORT13B"
                'SD_14_014R "13B" '混成課程教學環境資料表'教學環境資料表 'view-source:https://ojtims.wda.gov.tw/SD/14/SD_14_014?ID=309
                'PCSVALUE  'Hid_PCS.Value = TIMS.ClearSQM(Hid_PCS.Value) 
                'Dim selsqlstr As String = Replace(Hid_PCS.Value, "x", "-") 'TIMS.CombiSQLIN(Replace(Hid_PCS.Value, "x", "-"))
                Dim drOB As DataRow = TIMS.GET_ORG_BIDCASE(objconn, RIDValue.Value, Hid_BCID.Value, Hid_BCASENO.Value)
                Dim rPMS As New Hashtable
                rPMS.Clear()
                rPMS.Add("YEARS", drOB("YEARS"))
                rPMS.Add("selsqlstr", String.Concat(vPlanID, "-", vComIDNO, "-", vSeqNo))
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                'W13-1混成課程教學環境資料表
                Call RPT_SD_14_014R(rPMS)

            Case "DOWNLOAD13B"
                'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                'Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
                'Hid_BCASENO.Value = TIMS.ClearSQM(Hid_BCASENO.Value)
                'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "BCID", Hid_BCID.Value)
                TIMS.SetMyValue2(rPMS4, "BCASENO", Hid_BCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "KBID", vKBID)
                TIMS.SetMyValue2(rPMS4, "KBSID", vKBSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "BCRTID", vBCRTID)
                Call TIMS.ResponseZIPFile_BI(sm, objconn, Me, rPMS4)
                Return
        End Select
    End Sub



    Protected Sub BTN_BACK2_Click(sender As Object, e As EventArgs) Handles BTN_BACK2.Click
        '清理隱藏的參數
        Call ClearHidValue()

        Call SHOW_Frame1(0)
    End Sub

    Protected Sub BTN_SAVENEXT2_Click(sender As Object, e As EventArgs) Handles BTN_SAVENEXT2.Click
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Dim sERRMSG As String = ""
        Dim fg_CanADDNEW1 As Boolean = Utl_ADDNEW_DATA2(sERRMSG, drRR)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Return
        End If

        Hid_BCID.Value = TIMS.ClearSQM(Hid_BCID.Value)
        Call SHOW_Detail_BIDCASE(drRR, Hid_BCID.Value, "")
    End Sub
End Class
