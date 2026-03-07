Imports System.IO

Partial Class TC_02_003
    Inherits AuthBasePage

    '基本上委訓單位使用 '申請階段管理-受理期間設定 APPLISTAGE
    'Dim fg_can_applistage As Boolean = False
    'Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試
    'Dim fg_IsSuperUser_S1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
    Dim fg_IsSuperUser_S1 As Boolean = False

    '最近一次版本送件
    Const cst_MTYPE_LATEST_SEND1 As String = "MTYPE_LATEST_SEND1"
    '最近一次版本-下載
    Const cst_MTYPE_LATEST_DOWN1 As String = "MTYPE_LATEST_DOWN1"
    '儲存(暫存
    'Const cst_ACTTYPE_BTN_SAVETMP1 As String = "BTN_SAVETMP1" '儲存(暫存)
    ''' <summary>'儲存後進下一步</summary>
    Const cst_ACTTYPE_BTN_SAVENEXT1 As String = "BTN_SAVENEXT1" '儲存後進下一步

    Dim tryFIND As String = ""
    Dim iDG11_ROWS As Integer = 0
    Dim iDG10_ROWS As Integer = 0
    '以目前版本批次送出
    Const cst_txt_版本批次送出 As String = "(版本批次送出)"
    Const cst_txt_免附文件 As String = "(免附文件)"
    Const cst_REUPLOADED_MSG As String = "(已重新上傳)"

    Const cst_SF_02_申復意見表_WAIVED_PI As String = "PI"

    'Const cst_SF_G01_申復申請書 As String = "G01"
    'Const cst_SF_G02_申復意見表 As String = "G02"
    'Const cst_SF_G03_公文 As String = "G03"
    'Const cst_SF_G04_其他佐證文件 As String = "G04"

    'Const cst_SF_W01_申復申請書 As String = "W01"
    'Const cst_SF_W02_申復意見表 As String = "W02"
    'Const cst_SF_W03_公文 As String = "W03"
    'Const cst_SF_W04_其他佐證文件 As String = "W04"

    Const cst_ss_RqProcessType As String = "RqProcessType" 'Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1
    Const cst_DG1CMDNM_DELETE1 As String = "DELETE1"
    Const cst_DG1CMDNM_VIEW1 As String = "VIEW1"
    Const cst_DG1CMDNM_EDIT1 As String = "EDIT1"
    Const cst_DG1CMDNM_SENDOUT1 As String = "SENDOUT1"
    Const cst_DG1CMDNM_RETURNSEND1 As String = "RETURNSEND1"

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
    Const cst_stopmsg_11 As String = "申請申複階段受理期間未開放，請確認後再操作!"

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

    ''Add new application cases and add reminder messages
    'Const Cst_messages1 As String = "請務必確認此年度/申請階段之所有欲研提班級都已送審，【新增申辦案件】後才送審的班級，將無法納入此次線上申辦案件清單中!"
    'Const cst_ss_messages1 As String = "messages1"

    Dim tmpMSG As String = ""
    Dim ff3 As String = ""
    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'center.Attributes.Add("onfocus", "this.blur();")
        TIMS.INPUT_ReadOnly2(center)

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
        '只有署 或 super 才可使用還原
        fg_IsSuperUser_S1 = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1 '分頁設定
        '(暫存停用／不使用／無用) by 20240105
        'BTN_SAVETMP1.Visible = False
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
        sch_txtSFCASENO.Text = ""
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
        sch_ddlAPPSTAGE = TIMS.Get_APPSTAGE2(sch_ddlAPPSTAGE)
        Dim v_APPSTAGE As String = TIMS.GET_CANUSE_APPSTAGE(objconn, CStr(sm.UserInfo.Years), TIMS.cst_APPSTAGE_PTYPE1_02)
        Dim v_APPSTAGE_SCH_DEF As String = "1"
        Common.SetListItem(sch_ddlAPPSTAGE, If(v_APPSTAGE <> "", v_APPSTAGE, v_APPSTAGE_SCH_DEF))

        '申辦人姓名
        sch_txtSFCNAME.Text = ""
        '申辦日期
        sch_txtSFCDATE1.Text = ""
        sch_txtSFCDATE2.Text = ""

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Get_TitleLab(objconn, MRqID, TitleLab1, TitleLab2)
    End Sub

    ''' <summary>顯示調整</summary>
    ''' <param name="iNum"></param>
    Private Sub SHOW_Frame1(ByVal iNum As Integer)
        FrameTableSch1.Visible = If(iNum = 0, True, False)
        FrameTableEdt1.Visible = If(iNum = 1, True, False)
    End Sub

    ''' <summary>新增申辦案件</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_ADDNEW1_Click(sender As Object, e As EventArgs) Handles BTN_ADDNEW1.Click
        '清理隱藏的參數
        Call ClearHidValue()

        'RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        'Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        'Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE)
        Dim sERRMSG As String = ""
        Dim fg_CanADDNEW1 As Boolean = Utl_ADDNEW_DATA1(sERRMSG, drRR)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Return
        End If

        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Call SHOW_Detail_SFCASE(drRR, Hid_SFCID.Value, "")
    End Sub

    ''' <summary>'檢核新增(查詢)調整</summary>
    ''' <param name="sERRMSG"></param>
    ''' <returns></returns>
    Private Function Utl_ADDNEW_DATA1(ByRef sERRMSG As String, ByRef drRR As DataRow) As Boolean
        If drRR Is Nothing Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構有誤")
            Return False
        End If
        'tr_HISREVIEW.Visible = False '歷程資訊
        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE)
        '申請階段管理-受理期間設定 APPLISTAGE
        Dim aParms As New Hashtable From {{"YEARS", v_sch_ddlYEARS}, {"APPSTAGE", v_sch_ddlAPPSTAGE}}
        '開放受理之申請階段／PLAN_APPSTAGE
        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE02(objconn, aParms)
        '檢核查詢 '開放受理之申請階段／PLAN_APPSTAGE
        If Not fg_can_applistage Then
            sERRMSG = cst_stopmsg_11 '"申請階段受理期間未開放，請確認後再操作!"
            Return False
        End If

        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))
        Dim v_RID As String = Convert.ToString(drRR("RID"))

        Select Case sm.UserInfo.LID
            Case 2
                If v_RID = "" OrElse v_RID <> sm.UserInfo.RID Then
                    sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構業務代碼有誤!!")
                    Return False
                ElseIf v_ORGID = "" OrElse v_ORGID <> sm.UserInfo.OrgID Then
                    sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構有誤!!")
                    Return False
                ElseIf v_PLANID = "" OrElse v_PLANID <> sm.UserInfo.PlanID Then
                    sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構計畫有誤!!")
                    Return False
                ElseIf v_ORGLEVEL = "" OrElse Val(v_ORGLEVEL) < 2 Then
                    sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構層級有誤!!")
                    Return False
                End If
            Case 1
                If v_PLANID = "" OrElse v_PLANID <> sm.UserInfo.PlanID Then
                    sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " (分署)新增訓練機構計畫有誤!")
                    Return False
                End If
            Case 0
                Dim fg_TY1 As Boolean = (v_TPLANID = sm.UserInfo.TPlanID AndAlso v_YEARS = v_sch_ddlYEARS)
                If v_TPLANID = "" OrElse v_TPLANID = "" OrElse Not fg_TY1 Then
                    sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " (署)新增訓練機構計畫有誤")
                    Return False
                End If
        End Select
        'Case Else '0
        If v_RID = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構業務代碼有誤")
            Return False
        ElseIf v_ORGID = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構有誤")
            Return False
        ElseIf v_PLANID = "" OrElse v_TPLANID = "" OrElse v_YEARS = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構計畫有誤")
            Return False
        ElseIf v_ORGLEVEL = "" OrElse Val(v_ORGLEVEL) < 2 Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增訓練機構層級有誤")
            Return False
        ElseIf v_sch_ddlYEARS = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增計畫年度不可為空")
            Return False
        ElseIf v_sch_ddlAPPSTAGE = "" Then
            sERRMSG = String.Concat(TIMS.cst_NODATAMsg2, " 新增申請階段不可為空")
            Return False
        End If

        '同一年度／轄區 (計畫)PLANID／申請階段APPSTAGE，每個(訓練單位)RID只能有一筆申辦案件
        Dim sParms2 As New Hashtable From {{"PLANID", v_PLANID}, {"RID", v_RID}, {"APPSTAGE", v_sch_ddlAPPSTAGE}}
        Dim sSql2 As String = "SELECT 1 FROM ORG_SFCASE WHERE PLANID=@PLANID AND RID=@RID AND APPSTAGE=@APPSTAGE"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, objconn, sParms2)
        If dt2.Rows.Count > 0 Then
            sERRMSG = "同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能1筆申辦案(或使用查詢修改)"
            Return False
        End If

        '於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)
        Dim S1Parms As New Hashtable From {{"TPLANID", v_TPLANID}, {"RID", v_RID}, {"PLANID", v_PLANID}, {"APPSTAGE", v_sch_ddlAPPSTAGE}}
        Dim strCLASSNAME2S As String = TIMS.GET_CLASSNAME2S_SF(objconn, S1Parms, TIMS.cst_outTYPE_CLSNM)
        Dim v_PCS_Value As String = TIMS.GET_CLASSNAME2S_SF(objconn, S1Parms, TIMS.cst_outTYPE_PCSVAL)
        If strCLASSNAME2S = "" OrElse v_PCS_Value = "" Then
            sERRMSG = "查無申復班級清單資料(請先執行-申復申請作業)"
            Return False
        End If

        Dim vSFCASENO_NN As String = ""
        Dim iSFCID As Integer = 0
        Try
            iSFCID = DbAccess.GetNewId(objconn, "ORG_SFCASE_SFCID_SEQ,ORG_SFCASE,SFCID")
            Dim irParms As New Hashtable From {{"DISTID", v_DISTID}, {"RID", v_RID}, {"APPSTAGE", v_sch_ddlAPPSTAGE}}
            vSFCASENO_NN = GET_SFCASENO_NN(objconn, irParms)

            Dim iParms As New Hashtable From {
                {"SFCID", iSFCID},
                {"SFCASENO", vSFCASENO_NN},
                {"YEARS", v_YEARS},
                {"DISTID", v_DISTID},
                {"ORGID", v_ORGID},
                {"PLANID", v_PLANID},
                {"RID", v_RID},
                {"APPSTAGE", v_sch_ddlAPPSTAGE},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_SFCASE(SFCID,SFCASENO,YEARS,DISTID,ORGID" & vbCrLf
            isSql &= " ,PLANID,RID,APPSTAGE,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE )" & vbCrLf
            isSql &= " VALUES (@SFCID,@SFCASENO,@YEARS,@DISTID,@ORGID" & vbCrLf
            isSql &= " ,@PLANID,@RID,@APPSTAGE,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE() )" & vbCrLf
            ',SFCACCT,SFCDATE,SFCSTATUS,@SFCACCT,@SFCDATE,@SFCSTATUS
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Catch ex As Exception
            sERRMSG = "資料庫序號有誤，請重新操作!"
            Return False
        End Try

        '異常狀況有2筆資料產生，刪除後面產生的錯誤資料
        Dim sParms2B As New Hashtable From {{"PLANID", v_PLANID}, {"RID", v_RID}, {"APPSTAGE", v_sch_ddlAPPSTAGE}}
        Dim sSql2B As String = " SELECT PLANID,RID,APPSTAGE,COUNT(1) CNT1,MAX(SFCID) SFCID,MIN(SFCID) MIN_SFCID FROM ORG_SFCASE WITH(NOLOCK) WHERE PLANID=@PLANID AND RID=@RID AND APPSTAGE=@APPSTAGE GROUP BY PLANID ,RID,APPSTAGE HAVING COUNT(1)>1" & vbCrLf
        Dim dt2B As DataTable = DbAccess.GetDataTable(sSql2B, objconn, sParms2B)
        If TIMS.dtHaveDATA(dt2B) Then
            Dim v_MIN_SFCID As String = Convert.ToString(dt2B.Rows(0)("MIN_SFCID"))
            Call DEL_ORG_SFCASE_NG(objconn, v_PLANID, v_RID, v_sch_ddlAPPSTAGE, v_MIN_SFCID)
            sERRMSG = "同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能1筆申辦案(或使用查詢修改)"
            Return False
        End If

        Hid_SFCID.Value = iSFCID
        Hid_SFCASENO.Value = vSFCASENO_NN

        Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))
        labCLASSNAME2S.Text = TIMS.GetResponseWrite(strCLASSNAME2S)
        Hid_PCS.Value = v_PCS_Value

        '審核班級儲存
        Call SAVE_ORG_SFCASEPI(iSFCID)
        Return True
    End Function


    ''' <summary>審核班級儲存</summary>
    ''' <param name="iSFCID"></param>
    Private Sub SAVE_ORG_SFCASEPI(ByVal iSFCID As Integer)
        If iSFCID = 0 OrElse Hid_PCS.Value = "" Then Return
        Dim saPCSALL1 As String() = Hid_PCS.Value.Split(",")

        Dim isSql2 As String = ""
        isSql2 &= " INSERT INTO ORG_SFCASEPI(SFCPID,SFCID,PLANID,COMIDNO,SEQNO,MODIFYACCT,MODIFYDATE)" & vbCrLf
        isSql2 &= " VALUES (@SFCPID,@SFCID,@PLANID,@COMIDNO,@SEQNO,@MODIFYACCT,GETDATE())" & vbCrLf
        ',@SENTBATVER,@SENTBATACCT,@SENTBATDATE

        For Each sPCS1 As String In saPCSALL1
            Dim saPCS1 As String() = sPCS1.Split("x")
            If saPCS1.Length = 3 Then
                Dim pPLANID As String = saPCS1(0)
                Dim pCOMIDNO As String = saPCS1(1)
                Dim pSEQNO As String = saPCS1(2)
                Dim iSFCPID As Integer = 0
                Dim sParms3 As New Hashtable From {{"PLANID", pPLANID}, {"COMIDNO", pCOMIDNO}, {"SEQNO", pSEQNO}}
                Dim sSql3 As String = "SELECT 1 FROM ORG_SFCASEPI WHERE PLANID=@PLANID AND COMIDNO=@COMIDNO AND SEQNO=@SEQNO"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If dt3.Rows.Count = 0 Then
                    iSFCPID = DbAccess.GetNewId(objconn, "ORG_SFCASEPI_SFCPID_SEQ,ORG_SFCASEPI,SFCPID")
                    Dim iParms2 As New Hashtable From {
                        {"SFCPID", iSFCPID},
                        {"SFCID", iSFCID},
                        {"PLANID", Val(pPLANID)},
                        {"COMIDNO", pCOMIDNO},
                        {"SEQNO", Val(pSEQNO)},
                        {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    DbAccess.ExecuteNonQuery(isSql2, objconn, iParms2)
                End If
            End If
        Next
    End Sub

    ''' <summary>清理隱藏的參數</summary>
    Sub ClearHidValue()
        Hid_KSFID.Value = ""
        Hid_SFID.Value = ""
        Hid_LastSFID.Value = ""
        Hid_FirstKSFID.Value = ""

        Hid_SFCID.Value = ""
        Hid_SFCASENO.Value = ""
        Hid_SFCFID.Value = ""
        Hid_ORGKINDGW.Value = ""
        Hid_PCS.Value = ""
    End Sub

    ''' <summary>查詢鈕1</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH1_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH1.Click
        Dim sERRMSG1 As String = ""
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
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
        sch_txtSFCASENO.Text = TIMS.ClearSQM(sch_txtSFCASENO.Text)
        sch_txtSFCNAME.Text = TIMS.ClearSQM(sch_txtSFCNAME.Text)
        sch_txtSFCDATE1.Text = TIMS.Cdate3(sch_txtSFCDATE1.Text)
        sch_txtSFCDATE2.Text = TIMS.Cdate3(sch_txtSFCDATE2.Text)

        'RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        Dim v_sch_ddlAPPSTAG As String = TIMS.GetListValue(sch_ddlAPPSTAGE)

        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        'Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        'Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))

        'NULL 待送審、B 審核中、Y 審核通過、R 退件修正、N 審核不通過。
        Dim pParms As New Hashtable
        If sch_txtSFCASENO.Text <> "" Then pParms.Add("SFCASENO", sch_txtSFCASENO.Text)
        If v_sch_ddlYEARS <> "" Then pParms.Add("YEARS", v_sch_ddlYEARS)
        If v_sch_ddlAPPSTAG <> "" Then pParms.Add("APPSTAGE", v_sch_ddlAPPSTAG)
        If sch_txtSFCNAME.Text <> "" Then pParms.Add("SFCNAME", sch_txtSFCNAME.Text)
        If sch_txtSFCDATE1.Text <> "" Then pParms.Add("SFCDATE1", sch_txtSFCDATE1.Text)
        If sch_txtSFCDATE2.Text <> "" Then pParms.Add("SFCDATE2", sch_txtSFCDATE2.Text)

        Dim sSql As String = ""
        sSql &= " SELECT a.SFCID,a.SFCASENO,a.YEARS,a.DISTID,a.ORGID,a.PLANID,a.RID,a.APPSTAGE" & vbCrLf
        sSql &= " ,dbo.FN_CYEAR2(a.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        sSql &= " ,dbo.FN_GET_DISTNAME(a.DISTID,3) DISTNAME" & vbCrLf
        sSql &= " ,(SELECT oo.ORGNAME FROM dbo.ORG_ORGINFO oo WHERE oo.ORGID=a.ORGID) ORGNAME" & vbCrLf
        sSql &= " ,a.SFCACCT,dbo.FN_GET_USERNAME(a.SFCACCT) SFCNAME" & vbCrLf
        sSql &= " ,(SELECT x.ORGKINDGW FROM dbo.VIEW_RIDNAME x WHERE x.RID=a.RID) ORGKINDGW" & vbCrLf
        sSql &= " ,format(a.SFCDATE,'yyyy/MM/dd') SFCDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.SFCDATE) SFCDATE_ROC" & vbCrLf
        '申辦狀態：暫存 null/ 已送件B /退R
        sSql &= " ,a.SFCSTATUS" & vbCrLf
        sSql &= " ,CASE WHEN a.SFCSTATUS IS NULL THEN '暫存'" & vbCrLf
        sSql &= " WHEN a.SFCSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'" & vbCrLf
        sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'" & vbCrLf
        sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '通過'" & vbCrLf
        sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
        sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END SFCSTATUS_N" & vbCrLf
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL" & vbCrLf
        sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
        sSql &= " FROM ORG_SFCASE a" & vbCrLf
        sSql &= " JOIN VIEW_RIDNAME r on r.RID=a.RID" & vbCrLf
        sSql &= " LEFT JOIN AUTH_ACCOUNT u ON u.ACCOUNT=a.SFCACCT" & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf
        Select Case sm.UserInfo.LID
            Case 0
                If v_ORGLEVEL = 2 Then
                    sSql &= " AND r.RID=@RID" & vbCrLf
                    pParms.Add("RID", RIDValue.Value) 'sm.UserInfo.RID
                Else
                    sSql &= " AND r.DISTID=@DISTID" & vbCrLf
                    sSql &= " AND r.TPLANID=@TPLANID" & vbCrLf
                    pParms.Add("DISTID", v_DISTID)
                    pParms.Add("TPLANID", sm.UserInfo.TPlanID)
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
                sSql &= " AND r.TPLANID=@TPLANID" & vbCrLf
                pParms.Add("RID", RIDValue.Value) 'sm.UserInfo.RID
                pParms.Add("ORGID", sm.UserInfo.OrgID)
                pParms.Add("ORGLEVEL", sm.UserInfo.OrgLevel)
                pParms.Add("DISTID", v_DISTID)
                pParms.Add("TPLANID", sm.UserInfo.TPlanID)
            Case Else
                sSql &= " AND 1<>1" & vbCrLf
        End Select

        If sch_txtSFCASENO.Text <> "" Then sSql &= " AND a.SFCASENO=@SFCASENO" & vbCrLf
        If v_sch_ddlYEARS <> "" Then sSql &= " AND a.YEARS=@YEARS" & vbCrLf
        If v_sch_ddlAPPSTAG <> "" Then sSql &= " AND a.APPSTAGE=@APPSTAGE" & vbCrLf
        If sch_txtSFCNAME.Text <> "" Then sSql &= " AND u.NAME LIKE '%'+@SFCNAME+'%'" & vbCrLf
        If sch_txtSFCDATE1.Text <> "" Then sSql &= " AND a.SFCDATE>=@SFCDATE1" & vbCrLf
        If sch_txtSFCDATE2.Text <> "" Then sSql &= " AND a.SFCDATE<=@SFCDATE2" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            labmsg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        labmsg1.Text = ""
        TableDataGrid1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>檢核查詢 (Search1)／新增 (Search2)</summary>
    ''' <param name="drRR"></param>
    ''' <returns></returns>
    Function CHK_Search1(ByRef sm As SessionModel, ByRef drRR As DataRow, ByRef sERRMSG1 As String) As Boolean
        If drRR Is Nothing Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構有誤!")
            Return False
        End If

        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        Dim v_sch_ddlAPPSTAG As String = TIMS.GetListValue(sch_ddlAPPSTAGE)
        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))

        Select Case sm.UserInfo.LID
            Case 2
                If v_ORGID = "" OrElse v_ORGID <> sm.UserInfo.OrgID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (委訓單位)查詢訓練機構有誤!!", v_ORGID))
                    Return False
                ElseIf v_PLANID = "" OrElse v_PLANID = "0" OrElse v_PLANID <> sm.UserInfo.PlanID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (請選擇對應計畫年度機構)訓練機構計畫有誤!!", v_PLANID))
                    Return False
                ElseIf v_ORGLEVEL = "" OrElse v_ORGLEVEL <> "2" Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" 訓練機構層級有誤!!", v_ORGLEVEL))
                    Return False
                End If
            Case 1
                If v_PLANID = "" OrElse v_PLANID = "0" OrElse v_PLANID <> sm.UserInfo.PlanID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (請選擇對應計畫年度機構)訓練機構計畫有誤!", v_PLANID))
                    Return False
                End If
        End Select
        If v_sch_ddlYEARS = "" AndAlso v_sch_ddlYEARS <> sm.UserInfo.Years Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (請選擇對應計畫年度機構)查詢計畫年度有誤", v_YEARS))
            Return False
        End If
        If v_sch_ddlYEARS = "" Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 計畫年度不可為空!")
            Return False
        ElseIf v_sch_ddlAPPSTAG = "" Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 申請階段不可為空!")
            Return False
        End If
        Return True
    End Function

    ''' <summary>檢核查詢 (Search1)／新增 (Search2) </summary>
    ''' <param name="sm"></param>1
    ''' <param name="drRR"></param>
    ''' <param name="sERRMSG1"></param>
    ''' <returns></returns>
    Function CHK_Search2(ByRef sm As SessionModel, ByRef drRR As DataRow, ByRef sERRMSG1 As String) As Boolean
        If drRR Is Nothing Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, " 訓練機構有誤!")
            Return False
        End If

        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS)
        Dim v_sch_ddlAPPSTAG As String = TIMS.GetListValue(sch_ddlAPPSTAGE)
        'Dim v_ORGID As String = sm.UserInfo.OrgID
        Dim v_ORGID As String = Convert.ToString(drRR("ORGID"))
        Dim v_ORGLEVEL As String = Convert.ToString(drRR("ORGLEVEL"))
        Dim v_PLANID As String = Convert.ToString(drRR("PLANID"))
        Dim v_TPLANID As String = Convert.ToString(drRR("TPLANID"))
        Dim v_YEARS As String = Convert.ToString(drRR("YEARS"))
        Dim v_DISTID As String = Convert.ToString(drRR("DISTID"))

        Select Case sm.UserInfo.LID
            Case 2
                If v_ORGID = "" OrElse v_ORGID <> sm.UserInfo.OrgID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (委訓單位)新增訓練機構有誤!!", v_ORGID))
                    Return False
                ElseIf v_PLANID = "" OrElse v_PLANID = "0" OrElse v_PLANID <> sm.UserInfo.PlanID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (請選擇對應計畫年度機構)新增訓練機構計畫有誤!!", v_PLANID))
                    Return False
                ElseIf v_YEARS = "" OrElse v_YEARS <> sm.UserInfo.Years Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" 新增訓練機構計畫年度有誤!!", v_YEARS))
                    Return False
                ElseIf v_ORGLEVEL = "" OrElse v_ORGLEVEL <> "2" Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" 新增訓練機構層級有誤!!", v_ORGLEVEL))
                    Return False
                End If
            Case 1
                If v_ORGID = "" Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (分署)新增訓練機構有誤!", v_ORGID))
                    Return False
                ElseIf v_PLANID = "" OrElse v_PLANID = "0" OrElse v_PLANID <> sm.UserInfo.PlanID Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" (請選擇對應計畫年度機構)新增訓練機構計畫有誤!", v_PLANID))
                    Return False
                ElseIf v_YEARS = "" OrElse v_YEARS <> sm.UserInfo.Years Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" 新增訓練機構計畫年度有誤!", v_YEARS))
                    Return False
                ElseIf v_ORGLEVEL = "" OrElse v_ORGLEVEL <> "2" Then
                    sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat(" 新增訓練機構層級有誤!", v_ORGLEVEL))
                    Return False
                End If
        End Select
        If v_TPLANID = "" OrElse v_TPLANID <> sm.UserInfo.TPlanID Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat("-(請選擇對應計畫年度機構)新增訓練機構計畫有誤!", v_PLANID))
            Return False
        ElseIf v_YEARS = "" OrElse v_YEARS <> sm.UserInfo.Years Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat("-新增訓練機構計畫年度有誤!", v_YEARS))
            Return False
        ElseIf v_ORGLEVEL = "" OrElse v_ORGLEVEL <> "2" Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, String.Concat("-新增訓練機構層級有誤!", v_ORGLEVEL))
            Return False
        End If
        If v_sch_ddlYEARS = "" Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, "-新增計畫年度不可為空!!")
            Return False
        ElseIf v_sch_ddlAPPSTAG = "" Then
            sERRMSG1 = String.Concat(TIMS.cst_NODATAMsg2, "-新增申請階段不可為空!!")
            Return False
        End If
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
        Hid_KSFID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_FirstKSFID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    ''' <summary>0:暫時儲存／1:正式儲存-UPDATE ORG_SFCASE</summary>
    ''' <param name="iNum"></param>
    Private Sub SAVEDATA1(ByVal iNum As Integer)
        'iNum:0 暫時儲存/1 正式儲存
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim iSFCID As Integer = If(Hid_SFCID.Value <> "", Val(Hid_SFCID.Value), 0)
        If Hid_SFCID.Value = "" OrElse iSFCID <= 0 Then Return

        Dim uParms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"SFCID", iSFCID}, {"RID", RIDValue.Value}, {"SFCASENO", Hid_SFCASENO.Value}}
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_SFCASE" & vbCrLf
        usSql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE SFCID=@SFCID AND RID=@RID AND SFCASENO=@SFCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        '審核班級儲存
    End Sub

    ''' <summary>(申復線上送件-送出)儲存2-UPDATE ORG_SFCASE </summary>
    ''' <param name="rPMS"></param>
    Public Shared Sub SAVEDATA2(oConn As SqlConnection, rPMS As Hashtable)
        Dim vSFCACCT As String = TIMS.GetMyValue2(rPMS, "SFCACCT") ' sm.UserInfo.UserID)
        Dim vSFCSTATUS As String = TIMS.GetMyValue2(rPMS, "SFCSTATUS") 'B/NULL
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT") ' sm.UserInfo.UserID)
        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID") 'SFCID
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID") 'RID
        Dim vSFCASENO As String = TIMS.GetMyValue2(rPMS, "SFCASENO") 'SFCASENO

        If vSFCSTATUS = "" Then
            Dim uParms2 As New Hashtable From {{"MODIFYACCT", vMODIFYACCT}, {"SFCID", vSFCID}, {"RID", vRID}, {"SFCASENO", vSFCASENO}}
            Dim usSql2 As String = ""
            usSql2 &= " UPDATE ORG_SFCASE" & vbCrLf
            usSql2 &= " SET SFCACCT=NULL,SFCDATE=NULL,SFCSTATUS=NULL" & vbCrLf
            usSql2 &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql2 &= " WHERE SFCID=@SFCID AND RID=@RID AND SFCASENO=@SFCASENO" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql2, oConn, uParms2)
            Return '(還原送審資料清空欄位)
        End If

        Dim uParms As New Hashtable From {{"SFCACCT", vSFCACCT}, {"SFCSTATUS", vSFCSTATUS}, {"MODIFYACCT", vMODIFYACCT},
            {"SFCID", vSFCID}, {"RID", vRID}, {"SFCASENO", vSFCASENO}}
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_SFCASE" & vbCrLf
        usSql &= " SET SFCACCT=@SFCACCT,SFCDATE=GETDATE(),SFCSTATUS=@SFCSTATUS" & vbCrLf
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE SFCID=@SFCID AND RID=@RID AND SFCASENO=@SFCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, oConn, uParms)
    End Sub


    ''' <summary>按下 查看／修改／送出 </summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vRID As String = TIMS.GetMyValue(sCmdArg, "RID")
        Dim vSFCID As String = TIMS.GetMyValue(sCmdArg, "SFCID")
        Dim vSFCASENO As String = TIMS.GetMyValue(sCmdArg, "SFCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        If sCmdArg = "" OrElse vSFCID = "" OrElse vRID = "" Then Return

        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, vRID, vSFCID, vSFCASENO)
        Dim drRR As DataRow = TIMS.Get_RID_DR(vRID, objconn) 'If drRR Is Nothing Then Return
        Call SHOW_RIDValue_DATA(drRR)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "申復資訊有誤(查無業務代碼)，請選擇訓練機構!")
            Return
        ElseIf drOB Is Nothing Then
            Common.MessageBox(Me, "申復資訊有誤(查無案件資料)，請選擇訓練機構!")
            Return
        End If
        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))

        '申請階段管理-受理期間設定 APPLISTAGE
        Dim aParms As New Hashtable From {{"YEARS", vYEARS}, {"APPSTAGE", vAPPSTAGE}}
        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE02(objconn, aParms)

        Dim s_RESULTDATE_YMS2 As String = If(Convert.ToString(drOB("RESULTDATE")) <> "", CDate(drOB("RESULTDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
        Dim s_MODIFYDATE_YMS2 As String = If(Convert.ToString(drOB("MODIFYDATE")) <> "", CDate(drOB("MODIFYDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
        Dim fg_RESULTDATE_UPDATE As Boolean = (s_RESULTDATE_YMS2 <> "" AndAlso s_MODIFYDATE_YMS2 <> "" AndAlso DateDiff(DateInterval.Second, CDate(s_RESULTDATE_YMS2), CDate(s_MODIFYDATE_YMS2)) > 0)

        Select Case e.CommandName
            Case cst_DG1CMDNM_DELETE1 'DELETE1 (刪除)
                Call DELETE_Detail_SFCASE(Me, objconn, drRR, drOB)
                Common.MessageBox(Me, TIMS.cst_DELETEOKMsg2)
                Call SSearch1(drRR)

            Case cst_DG1CMDNM_RETURNSEND1 'RETURNSEND1 還原送出 
                'rPMS.Add("SFCACCT", sm.UserInfo.UserID)                'rPMS.Add("SFCSTATUS", "B")
                Dim rPMS As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"SFCID", vSFCID}, {"RID", vRID}, {"SFCASENO", vSFCASENO}}
                Call SAVEDATA2(objconn, rPMS)

                Call SSearch1(drRR)
            Case cst_DG1CMDNM_VIEW1 '"VIEW1 '查看
                Call SHOW_Detail_SFCASE(drRR, vSFCID, cst_DG1CMDNM_VIEW1)

            Case cst_DG1CMDNM_EDIT1 '"EDIT1 '修改
                If Not fg_can_applistage Then
                    Common.MessageBox(Me, cst_stopmsg_11) 'Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                    Return
                End If
                Call SHOW_Detail_SFCASE(drRR, vSFCID, cst_DG1CMDNM_EDIT1)

            Case cst_DG1CMDNM_SENDOUT1 'SENDOUT1 送出 
                If Not fg_can_applistage Then
                    Common.MessageBox(Me, cst_stopmsg_11) 'Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                    Return
                End If

                '線上申辦進度 計算完成度百分比 (0-100)
                Dim vSFCSTATUS As String = Convert.ToString(drOB("SFCSTATUS"))
                Dim iProgress As Integer = TIMS.GET_iPROGRESS_SF(objconn, tmpMSG, vSFCID, vORGKINDGW)
                Dim EMSG As String = ""
                If iProgress < 100 Then
                    EMSG = String.Concat("線上申辦進度 未達100%，不可送出!", vbCrLf, If(tmpMSG <> "", String.Concat("請檢查：(", tmpMSG, ")"), ""))
                    Common.MessageBox(Me, EMSG)
                    Return
                ElseIf vSFCSTATUS = "R" AndAlso Not fg_RESULTDATE_UPDATE Then
                    EMSG = cst_tpmsg_enb5
                    Common.MessageBox(Me, EMSG)
                    Return
                End If

                Dim rPMS As New Hashtable From {{"SFCACCT", sm.UserInfo.UserID}, {"SFCSTATUS", "B"},
                    {"MODIFYACCT", sm.UserInfo.UserID}, {"SFCID", vSFCID}, {"RID", vRID}, {"SFCASENO", vSFCASENO}}
                Call SAVEDATA2(objconn, rPMS)

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
                Dim lBTN_RETURNSEND1 As LinkButton = e.Item.FindControl("lBTN_RETURNSEND1") 'RETURNSEND1 還原送出 
                Call TIMS.Display_None(lBTN_RETURNSEND1)
                lBTN_RETURNSEND1.Visible = (fg_IsSuperUser_S1 OrElse sm.UserInfo.LID > 0) ' False
                lBTN_RETURNSEND1.Enabled = (fg_IsSuperUser_S1 OrElse sm.UserInfo.LID > 0)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SFCID", drv("SFCID"))
                TIMS.SetMyValue(sCmdArg, "RID", drv("RID"))
                TIMS.SetMyValue(sCmdArg, "SFCASENO", drv("SFCASENO"))
                TIMS.SetMyValue(sCmdArg, "SFCSTATUS", drv("SFCSTATUS"))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", drv("ORGKINDGW"))

                'Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                'lBTN_DELETE1.Visible = If(flagS1, True, False)
                'lBTN_DELETE1.Style.Item("display") = "none"
                'lBTN_DELETE1.CommandArgument = sCmdArg
                'lBTN_DELETE1.Attributes("onclick") = "javascript:return confirm('此動作會刪除審核資料，是否確定?');"

                lBTN_DELETE1.CommandArgument = sCmdArg
                lBTN_DELETE1.Attributes("onclick") = "javascript:return confirm('此動作會刪除送件資料，是否確定?');"
                lBTN_VIEW1.CommandArgument = sCmdArg
                lBTN_EDIT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.Attributes("onclick") = "javascript:return confirm('此動作會送出審核資料且不可再次修改，是否確定?');"
                lBTN_RETURNSEND1.CommandArgument = sCmdArg
                lBTN_RETURNSEND1.Attributes("onclick") = "javascript:return confirm('此動作會還原送出審核資料，是否確定?');"

                Dim vSFCSTATUS As String = Convert.ToString(drv("SFCSTATUS"))
                lBTN_DELETE1.Visible = (vSFCSTATUS <> "B")

                If vSFCSTATUS = "R" Then
                    lBTN_EDIT1.Enabled = True
                    TIMS.Tooltip(lBTN_EDIT1, cst_tpmsg_enb5, True)

                    lBTN_SENDOUT1.Enabled = True
                    TIMS.Tooltip(lBTN_SENDOUT1, cst_tpmsg_enb5, True)
                Else
                    lBTN_EDIT1.Enabled = If(vSFCSTATUS <> "", False, True)
                    TIMS.Tooltip(lBTN_EDIT1, If(lBTN_EDIT1.Enabled, "", cst_tpmsg_enb4), True)

                    lBTN_SENDOUT1.Enabled = If(vSFCSTATUS <> "", False, True)
                    TIMS.Tooltip(lBTN_SENDOUT1, If(lBTN_SENDOUT1.Enabled, "", cst_tpmsg_enb4), True)
                End If

        End Select
    End Sub

    ''' <summary>新增使用資料顯示／查詢使用資料顯示 依 ORG_SFCASE - SFCID</summary>
    Private Sub SHOW_Detail_SFCASE(ByRef drRR As DataRow, ByVal vSFCID As String, ByVal vCmdName As String)
        '訓練機構有誤
        If drRR Is Nothing Then
            Common.MessageBox(Me, "訓練機構資料有誤!")
            Return
        End If
        Call SHOW_Frame1(1)

        Dim rLastSFID As String = ""
        Dim rFirstKSFID As String = ""
        Session(cst_ss_RqProcessType) = vCmdName

        Dim vRID As String = Convert.ToString(drRR("RID"))
        Dim vPLANID As String = Convert.ToString(drRR("PLANID"))
        Hid_ORGKINDGW.Value = Convert.ToString(drRR("ORGKINDGW"))

        Call Utl_GET_SWITCHTO_VAL(objconn, Hid_ORGKINDGW.Value, rLastSFID, rFirstKSFID)
        Hid_LastSFID.Value = rLastSFID
        Hid_FirstKSFID.Value = rFirstKSFID

        ddlSwitchTo = TIMS.GET_ddlKEY_SFCASE(objconn, ddlSwitchTo, Hid_ORGKINDGW.Value)
        'Common.SetListItem(ddlSwitchTo, Hid_FirstKSFID.Value)

        Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE)
        Dim vAPPSTAGE As String = ""

        Dim dtB1 As DataTable = Nothing
        If vSFCID <> "" Then
            '查詢資料  'NULL 待送審、B 審核中、Y 審核通過、R 退件修正、N 審核不通過。
            Dim shtParms As New Hashtable From {{"RID", vRID}, {"SFCID", vSFCID}, {"PLANID", vPLANID}}

            Dim sSql As String = ""
            sSql &= " SELECT a.SFCID,a.SFCASENO,a.YEARS,a.PLANID,a.DISTID,a.ORGID,a.RID,a.APPSTAGE" & vbCrLf
            'APPSTAGE_N
            sSql &= " ,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
            'DISTNAME
            sSql &= " ,dbo.FN_GET_DISTNAME(a.DISTID,3) DISTNAME" & vbCrLf
            'ORGNAME
            sSql &= " ,(SELECT ORGNAME FROM ORG_ORGINFO WHERE ORGID=a.ORGID) ORGNAME" & vbCrLf
            'COMIDNO
            sSql &= " ,(SELECT COMIDNO FROM ORG_ORGINFO WHERE ORGID=a.ORGID) COMIDNO" & vbCrLf
            'SFCACCTNAME
            sSql &= " ,a.SFCACCT,dbo.FN_GET_USERNAME(a.SFCACCT) SFCACCTNAME" & vbCrLf
            sSql &= " ,format(a.SFCDATE,'yyyy/MM/dd') SFCDATE" & vbCrLf
            sSql &= " ,dbo.FN_CDATE1B(a.SFCDATE) SFCDATE_ROC" & vbCrLf
            '申辦狀態：暫存/ 已送件
            sSql &= " ,a.SFCSTATUS" & vbCrLf
            sSql &= " ,CASE WHEN a.SFCSTATUS IS NULL THEN '暫存'" & vbCrLf
            sSql &= " WHEN a.SFCSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'" & vbCrLf
            sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'" & vbCrLf
            sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '通過'" & vbCrLf
            sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
            sSql &= " WHEN a.SFCSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END SFCSTATUS_N" & vbCrLf
            '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過 
            sSql &= " ,a.APPLIEDRESULT,a.REASONFORFAIL" & vbCrLf
            sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
            '歷程資訊 'sSql &= " ,a.HISREVIEW" & vbCrLf
            sSql &= " ,(SELECT MAX(KSFID) FROM ORG_SFCASEFL fl WHERE fl.SFCID=a.SFCID) CurrentKSFID" & vbCrLf
            sSql &= " ,(SELECT MIN(KSFID) FROM ORG_SFCASEFL fl WHERE fl.SFCID=a.SFCID AND fl.RTUREASON IS NOT NULL) Curr2KSFID" & vbCrLf
            sSql &= " FROM ORG_SFCASE a" & vbCrLf
            sSql &= " WHERE a.RID=@RID AND SFCID=@SFCID AND a.PLANID=@PLANID" & vbCrLf

            dtB1 = DbAccess.GetDataTable(sSql, objconn, shtParms)
            If dtB1 Is Nothing OrElse dtB1.Rows.Count = 0 Then Return
            Dim drB1 As DataRow = dtB1.Rows(0)

            'Hid_KSFID.Value = ""
            'Hid_SFID.Value = ""
            vAPPSTAGE = Convert.ToString(drB1("APPSTAGE"))
            Hid_SFCID.Value = Convert.ToString(drB1("SFCID"))
            Hid_SFCASENO.Value = Convert.ToString(drB1("SFCASENO"))
            labOrgNAME.Text = Convert.ToString(drB1("ORGNAME"))
            labYEARS.Text = TIMS.GET_YEARS_ROC(drB1("YEARS"))
            labAPPSTAGE.Text = Convert.ToString(drB1("APPSTAGE_N"))

            '退件
            If Convert.ToString(drB1("SFCSTATUS")) = "R" Then
                Hid_KSFID.Value = Convert.ToString(drB1("Curr2KSFID"))
            Else
                Hid_KSFID.Value = Convert.ToString(drB1("CurrentKSFID"))
            End If
            If Hid_KSFID.Value <> "" Then
                Common.SetListItem(ddlSwitchTo, Hid_KSFID.Value)
            ElseIf Hid_FirstKSFID.Value <> "" Then
                Common.SetListItem(ddlSwitchTo, Hid_FirstKSFID.Value)
            End If
        Else
            '(新增資料) '檢核查詢
            Dim sERRMSG1 As String = ""
            Dim flag_CHECKOK As Boolean = CHK_Search2(sm, drRR, sERRMSG1)
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
            labYEARS.Text = TIMS.GET_YEARS_ROC(drRR("YEARS"))
            labAPPSTAGE.Text = TIMS.GET_APPSTAGE2_NM2(v_sch_ddlAPPSTAGE)
            'Call Utl_SwitchTo(Hid_FirstKSFID.Value)
            'If Hid_FirstKSFID.Value <> "" Then
            '    Common.SetListItem(ddlSwitchTo, Hid_FirstKSFID.Value)
            'End If
        End If

        '	於當年度/申請階段之開放班級申請期間，可新增申辦案件(非申辦期間無法新增)。
        '	同一年度(計畫)／轄區／申請階段，每個(訓練單位)只能有一筆申辦案件。
        '	於新增申辦案件功能中，系統會自動帶入當年度/申請階段之所有已送審班級清單(【審核狀態】：班級審核中)。
        '	各項應備文件上傳會依項目依序分頁顯示，訓練單位也可跳頁填選。第一頁介面示意圖如下

        Dim S1Parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"RID", vRID}, {"PLANID", vPLANID}}
        If vSFCID <> "" Then S1Parms.Add("SFCID", vSFCID)
        S1Parms.Add("APPSTAGE", If(vAPPSTAGE <> "", vAPPSTAGE, v_sch_ddlAPPSTAGE))
        Dim strCLASSNAME2S As String = TIMS.GET_CLASSNAME2S_SF(objconn, S1Parms, TIMS.cst_outTYPE_CLSNM)
        Dim v_PCS_Value As String = TIMS.GET_CLASSNAME2S_SF(objconn, S1Parms, TIMS.cst_outTYPE_PCSVAL)

        labCLASSNAME2S.Text = TIMS.GetResponseWrite(strCLASSNAME2S)
        Hid_PCS.Value = v_PCS_Value
        labProgress.Text = "0%"

        If strCLASSNAME2S = "" OrElse v_PCS_Value = "" Then
            Common.MessageBox(Me, "查無申復班級清單資料(請先執行-申復申請作業)!")
            'Return
        End If

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        'Hid_KSFID.Value = TIMS.ClearSQM(Hid_KSFID.Value)
        Hid_FirstKSFID.Value = TIMS.ClearSQM(Hid_FirstKSFID.Value)

        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "SFCID", Hid_SFCID.Value)
        Call SHOW_SFCASEFL_DG2(rPMS3)

        Hid_KSFID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_FirstKSFID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    Private Sub SHOW_SFCASEFL_DG2(rPMS As Hashtable)
        labmsg1.Text = ""
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        'objconn 因為有檔案輸出關閉的問題 所以要檢查
        If Not TIMS.OpenDbConn(objconn) OrElse Not fg_CANSAVE Then Return

        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID")
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If

        Dim dtFL As DataTable = GET_ORG_SFCASEFL_TB(objconn, vSFCID, vORGKINDGW)

        labmsg1.Text = If(dtFL Is Nothing OrElse dtFL.Rows.Count = 0, "(查無文件項目)", "")

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vSFCASENO As String = Convert.ToString(drOB("SFCASENO"))
        Dim download_Path As String = TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, "")
        Call TIMS.Check_dtSFCASEFL(Me, dtFL, download_Path)
        'DataGrid2.Columns(cst_DG2_退件原因_iCOLUMN).Visible = If(Convert.ToString(drOB("APPLIEDRESULT")) = "R", True, False)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

        'Dim iProgress As Integer = If(dtA.Rows.Count > 0, (dt.Rows.Count / dtA.Rows.Count * 100), 0)
        '線上申辦進度 計算完成度百分比 (0-100)
        Dim iProgress As Integer = TIMS.GET_iPROGRESS_SF(objconn, tmpMSG, vSFCID, vORGKINDGW)
        labProgress.Text = String.Concat(iProgress, "%")
        'BTN_SAVETMP1.Visible = (iProgress = 100)
        'BTN_SAVERC2.Visible = (iProgress = 100)
        '儲存(暫存)
        'BTN_SAVETMP1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        'TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, "", cst_tpmsg_enb1), True)
        '進下一步 '儲存後進下一步
        BTN_SAVENEXT1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, "", cst_tpmsg_enb1), True)
    End Sub

    ''' <summary>回上一步</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PREV1_Click(sender As Object, e As EventArgs) Handles BTN_PREV1.Click
        Call MOVE_PREV()
    End Sub

    ''' <summary>回上一步</summary>
    Private Sub MOVE_PREV()
        If (Hid_SFID.Value = "" OrElse Hid_SFID.Value = "01" OrElse ddlSwitchTo.SelectedIndex - 1 = -1) Then
            Common.MessageBox(Me, "(目前沒有上一步)")
            Return
        End If

        'TIMS.AddZero(Val(Hid_KSFID.Value) - 1, 2)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Hid_KSFID.Value = Val(TIMS.GetListValue(ddlSwitchTo)) - 1
        Hid_KSFID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex - 1).Value
        If Hid_KSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_FirstKSFID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    '暫時儲存
    'Protected Sub BTN_SAVETMP1_Click(sender As Object, e As EventArgs) Handles BTN_SAVETMP1.Click
    '    '儲存(暫存)
    '    Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVETMP1)
    'End Sub

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
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        Hid_KSFID.Value = TIMS.ClearSQM(Hid_KSFID.Value)

        Hid_SFID.Value = TIMS.ClearSQM(Hid_SFID.Value)
        Hid_LastSFID.Value = TIMS.ClearSQM(Hid_LastSFID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)

        'SAVE_ORG_SFCASEFL
        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!")
            Return
        End If

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        '申請階段管理-受理期間設定 APPLISTAGE '檢核查詢
        Dim aParms As New Hashtable From {{"YEARS", vYEARS}, {"APPSTAGE", vAPPSTAGE}}
        Dim fg_can_applistage As Boolean = TIMS.CAN_APPLISTAGE_PTYPE02(objconn, aParms)
        If Not fg_can_applistage Then
            Common.MessageBox(Me, cst_stopmsg_11) ' "申請階段受理期間未開放，請確認後再操作!"
            Return
        End If

        '必須上傳檔案
        Dim fg_FILE_MUSTBE_UPLOADED As Boolean = True
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        Dim vKSFID As String = Convert.ToString(drKB("KSFID"))
        Dim vSFID As String = Convert.ToString(drKB("SFID"))
        Dim vORGKINDGW As String = Convert.ToString(drKB("ORGKINDGW"))
        Select Case String.Concat(vORGKINDGW, vSFID)
            Case TIMS.cst_SF_G02_申復意見表, TIMS.cst_SF_W02_申復意見表
                fg_FILE_MUSTBE_UPLOADED = False
                vWAIVED = cst_SF_02_申復意見表_WAIVED_PI
        End Select

        Dim vSFCID As String = TIMS.ClearSQM(Hid_SFCID.Value)
        Dim drFL As DataRow = TIMS.GET_ORG_SFCASEFL(objconn, vSFCID, vKSFID)
        Dim vFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
        'Dim vWAIVED As String = If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
        If fg_FILE_MUSTBE_UPLOADED AndAlso Not CHKB_WAIVED.Checked AndAlso vFILENAME1 = "" Then
            Common.MessageBox(Me, "未上傳檔案也未勾選免附且儲存過該文件，不可再次操作!")
            Return
        End If

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vSFCASENO As String = Convert.ToString(drOB("SFCASENO"))
        'Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, "")
        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'If (drFL IsNot Nothing AndAlso fg_CanSaveAgain_1) Then TIMS.SetMyValue2(rPMS2, "BCFID", drFL("BCFID"))
            TIMS.SetMyValue2(rPMS2, "SFID", vSFID)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "SFCID", vSFCID)
            TIMS.SetMyValue2(rPMS2, "KSFID", vKSFID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED)
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            Select Case vWAIVED
                Case "", "Y"
                    Call SAVE_ORG_SFCASEFL_UPLOAD(rPMS2)
            End Select
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub  'Throw ex
        End Try

        '暫時儲存／正式儲存-UPDATE ORG_SFCASE
        Call SAVEDATA1(0)

        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "SFCID", Hid_SFCID.Value)
        Call SHOW_SFCASEFL_DG2(rPMS3)

        'Case cst_ACTTYPE_BTN_SAVETMP1
        '    '儲存(暫存) '項目(重跑1次)
        '    Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        Select Case s_ACTTYPE
            Case cst_ACTTYPE_BTN_SAVENEXT1
                '進下一步 '儲存後進下一步  '(檢查儲存值)
                Dim rPMS As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"SFCID", Hid_SFCID.Value}, {"KSFID", Hid_KSFID.Value}}
                Dim flag_OK_OBFL As Boolean = CHK_ORG_SFCASEFL(rPMS)
                If Not flag_OK_OBFL Then
                    Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步!")
                    Return
                End If

                '下一步
                Call MOVE_NEXT()
        End Select

    End Sub

    ''' <summary>後進下一步</summary>
    Private Sub MOVE_NEXT()
        If Hid_SFID.Value <> "" AndAlso (Hid_SFID.Value = Hid_LastSFID.Value) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        ElseIf (ddlSwitchTo.SelectedIndex + 1 >= ddlSwitchTo.Items.Count) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        End If

        '下一步
        'Hid_KSFID.Value = If(Hid_SFID.Value = "" OrElse Hid_KSFID.Value = "", 1, Val(Hid_KSFID.Value) + 1)
        Hid_KSFID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex + 1).Value
        Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
    End Sub

    ''' <summary>下載報表</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_DOWNLOADRPT1_Click(sender As Object, e As EventArgs) Handles BTN_DOWNLOADRPT1.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Hid_KSFID.Value = TIMS.ClearSQM(Hid_KSFID.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim vKSFID As String = TIMS.GetListValue(ddlSwitchTo)
        If Hid_SFCASENO.Value = "" OrElse Hid_SFCID.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(案件號為空)，請重新操作!!")
            Return
        ElseIf RIDValue.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(業務代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_KSFID.Value = "" OrElse Hid_ORGKINDGW.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_KSFID.Value <> "" AndAlso Hid_KSFID.Value <> vKSFID Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目序號有誤)，請重新操作!!")
            Return
        End If

        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, Hid_KSFID.Value, Hid_ORGKINDGW.Value)
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

        '取得ID代號／流水號
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vORGID As String = Convert.ToString(drOB("ORGID"))
        Dim vDISTID As String = Convert.ToString(drOB("DISTID"))
        Dim vSFID As String = Convert.ToString(drKB("SFID"))
        Dim vORGKINDGW As String = Convert.ToString(drKB("ORGKINDGW"))
        Dim vSFIDNAME As String = String.Concat(vORGKINDGW, vSFID, drKB("SFNAME"))
        'Dim vRID As String = Convert.ToString(drOB("RID"))
        'Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim rPMS As New Hashtable
        Select Case String.Concat(vORGKINDGW, vSFID)
            Case TIMS.cst_SF_G01_申復申請書, TIMS.cst_SF_W01_申復申請書
                Const cst_printFN1 As String = "TC_02_002B"
                Dim drRR As DataRow = TIMS.Get_RID_DR(Convert.ToString(drOB("RID")), objconn)
                If RIDValue.Value = "" OrElse vORGID = "" OrElse vDISTID = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Return
                End If

                Dim MyValue1 As String = ""
                TIMS.SetMyValue(MyValue1, "TPlanID", sm.UserInfo.TPlanID)
                TIMS.SetMyValue(MyValue1, "DISTID", vDISTID)
                TIMS.SetMyValue(MyValue1, "YEARS", sm.UserInfo.Years)
                TIMS.SetMyValue(MyValue1, "RID", RIDValue.Value)
                TIMS.SetMyValue(MyValue1, "APPSTAGE", vAPPSTAGE)
                TIMS.SetMyValue(MyValue1, "ORGID", vORGID)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)
        End Select

    End Sub

#Region "共用"
    ''' <summary>共用取號</summary>
    ''' <param name="oConn"></param>
    ''' <param name="rParms"></param>
    ''' <returns></returns>
    Public Shared Function GET_SFCASENO_NN(ByRef oConn As SqlConnection, ByRef rParms As Hashtable) As String
        Dim v_YEARS_ROC As String = TIMS.GET_YEARS_ROC(Now.Year)
        Dim v_NMD4 As String = Now.ToString("MMdd") 'TIMS.GET_YEARS_ROC()
        'Dim vDISTID As String = TIMS.GetMyValue2(rParms, "DISTID")
        Dim v_RID As String = TIMS.GetMyValue2(rParms, "RID")
        Dim vRID1 As String = If(v_RID.Length >= 1, Left(v_RID, 1), v_RID)
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rParms, "APPSTAGE")
        Dim v_CASENO_T1 As String = String.Concat(v_YEARS_ROC, vRID1, vAPPSTAGE, v_NMD4)
        'v_YEARS_ROC +RID +APPSTAGE +MMdd
        '3+1+1+4 = 9 前面9位 從第10位取得

        Dim sParms As New Hashtable From {{"SFCASENOT1", v_CASENO_T1}}
        '112B27060001' BCASENO_MAX_SEQ
        Dim sql As String = "SELECT MAX(CAST(SUBSTRING(SFCASENO,10,9) AS INT)) SFCASENO_MAX_SEQ FROM ORG_SFCASE WITH(NOLOCK) WHERE SUBSTRING(SFCASENO,1,9)=SUBSTRING(@SFCASENOT1,1,9)"
        Dim drB1 As DataRow = DbAccess.GetOneRow(sql, oConn, sParms)
        Dim fg_NODATA As Boolean = (drB1 Is Nothing OrElse Convert.ToString(drB1("SFCASENO_MAX_SEQ")) = "")
        Dim iBCASENO_LAST_SEQ As Integer = If(fg_NODATA, 1, Val(drB1("SFCASENO_MAX_SEQ")) + 1)
        Return String.Concat(v_CASENO_T1, TIMS.AddZero(iBCASENO_LAST_SEQ.ToString(), 4))
    End Function
#End Region

#Region "Private1"

    ''' <summary>儲存後-檢核 ORG_SFCASEFL - 正確為true ／異常為false</summary>
    ''' <param name="rPMS"></param>
    ''' <returns></returns>
    Private Function CHK_ORG_SFCASEFL(rPMS As Hashtable) As Boolean
        '(外部參數)
        Dim vKSFID As String = TIMS.GetMyValue2(rPMS, "KSFID")
        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        If vKSFID = "" OrElse vSFCID = "" OrElse Not fg_CANSAVE Then Return False '(異常)查無鍵值1

        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, vKSFID, vORGKINDGW)
        If drKB Is Nothing Then Return False '(異常)查無鍵值2

        Dim rParms As New Hashtable From {
            {"SFCID", Val(vSFCID)},
            {"ORGKINDGW", vORGKINDGW},
            {"KSFID", vKSFID}
        }
        Dim rsSql As String = ""
        rsSql &= " SELECT a.SFCFID,ob.YEARS,ob.APPSTAGE,ob.RID,a.SFCID,a.KSFID" & vbCrLf
        rsSql &= " ,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        rsSql &= " ,kb.SFID,concat(kb.SFID,'.',kb.SFNAME) SFIDNAME" & vbCrLf
        rsSql &= " ,a.WAIVED,a.SRCFILENAME1,a.FILENAME1,a.FILENAME1 OKFLAG,a.FILEPATH1" & vbCrLf
        rsSql &= " FROM ORG_SFCASEFL a" & vbCrLf
        rsSql &= " JOIN KEY_SFCASE kb on kb.KSFID=a.KSFID" & vbCrLf
        rsSql &= " JOIN ORG_SFCASE ob on ob.SFCID=a.SFCID" & vbCrLf
        rsSql &= " JOIN AUTH_RELSHIP rr on rr.RID=ob.RID" & vbCrLf
        rsSql &= " WHERE a.SFCID=@SFCID AND kb.ORGKINDGW=@ORGKINDGW AND a.KSFID=@KSFID" & vbCrLf
        '(單筆資訊)
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, objconn, rParms)
        '(沒有檔案)
        Dim fg_NOFILE1 As Boolean = (dt Is Nothing OrElse dt.Rows.Count = 0)
        '(沒有檔案資料)
        If fg_NOFILE1 Then Return False '(異常)

        Return True
    End Function

    ''' <summary>清除-切換項目(預設)</summary>
    Sub CLEAR_SFCASE_KSFID()
        '取得文字說明
        LiteralSwitchTo.Text = "" 'If(vKSFID <> "", TIMS.HtmlDecode1(vSFDESC1), "(無)")
        tr_LiteralSwitchTo.Visible = False ' If(vNOTKBDESC1 = "Y", False, True) '(Y:不使用KBDESC)
        '檔案格式說明
        tr_FILEDESC1.Visible = False 'If(vNOTFLDESC1 = "Y", False, True) '(Y:不使用FLDESC1)
        '(使用)'(報表名稱)'DOWNLOADRPT '可下載報表
        BTN_DOWNLOADRPT1.Text = "" 'If(vRPTNAME <> "", vRPTNAME, "下載報表")
        tr_DOWNLOADRPT1.Visible = False 'If(vDOWNLOADRPT = "Y", True, False)
        '以目前版本批次送出:SENTBATVER
        tr_SENTBATVER.Visible = False 'If(vSENTBATVER = "Y", True, False)
        '以目前版本送出: SENDCURRVER
        tr_SENDCURRVER.Visible = False 'If(vSENDCURRVER = "Y", True, False)
        '檔案上傳:UPLOADFL1
        tr_UPLOADFL1.Visible = False 'If(vUPLOADFL1 = "Y", True, False)
        ' return checkFile1(sizeLimit);
        'But1.Attributes.Remove("onclick")
        'If vUPLOADFL1 = "Y" AndAlso str_rtn_checkFile1 <> "" Then But1.Attributes.Add("onclick", str_rtn_checkFile1)
        '取得SFID代號／非流水號
        Hid_SFID.Value = "" 'vSFID 'GET_SFID(vKSFID, vORGKINDGW) GET_KBDESC1(vKSFID, vORGKINDGW)
        LabSwitchTo.Text = "" 'If(vKSFID <> "", TIMS.GetListText(ddlSwitchTo), "")
        'USELATESTVER : 以最近一次版本送件
        tr_USELATESTVER.Visible = False 'If(vUSELATESTVER = "Y", True, False)
        'MUSTFILL 必填資訊／WAIVED:免附文件(必填就不顯示)
        tr_WAIVED.Visible = False ' If(vMUSTFILL = "Y", False, True)
        lbWAIVEDESC1.Text = ""
        '備註說明:USEMEMO1
        tr_USEMEMO1.Visible = False 'If(vUSEMEMO1 = "Y", True, False)
        '預設值-免附文件
        CHKB_WAIVED.Checked = False '(預設值不填寫)
        '預設值-(上傳檔案)
        Hid_SFCFID.Value = ""
        '預設值-備註說明
        txtMEMO1.Text = ""
    End Sub

    ''' <summary>'切換項目(預設)</summary>
    ''' <param name="vKSFID"></param>
    ''' <param name="vORGKINDGW"></param>
    Sub SHOW_SFCASE_KSFID(ByVal vKSFID As String, ByVal vORGKINDGW As String)
        'Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Dim v_ddlSwitchTo As String = TIMS.GetListValue(ddlSwitchTo)
        If (v_ddlSwitchTo <> vKSFID) Then
            Hid_KSFID.Value = vKSFID
            Common.SetListItem(ddlSwitchTo, vKSFID)
        End If

        '清除 - 切換項目(預設)
        Call CLEAR_SFCASE_KSFID()

        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, vKSFID, vORGKINDGW)
        If drKB Is Nothing Then Return

        '(取得)SFID代號／非流水號
        Dim vSFID As String = Convert.ToString(drKB("SFID"))
        '取得文字說明
        Dim vSFDESC1 As String = Convert.ToString(drKB("SFDESC1"))
        Dim vNOTKBDESC1 As String = Convert.ToString(drKB("NOTKBDESC1")) '(Y:不顯示文字說明/不使用KBDESC)
        Dim vNOTFLDESC1 As String = Convert.ToString(drKB("NOTFLDESC1")) '(Y:不使用FLDESC1)
        '必填資訊／免附文件(必填就不顯示)
        Dim vMUSTFILL As String = Convert.ToString(drKB("MUSTFILL"))
        '免附文件-其他說明
        Dim vWAIVEDESC1 As String = Convert.ToString(drKB("WAIVEDESC1"))
        'USELATESTVER : 以最近一次版本送件
        Dim vUSELATESTVER As String = Convert.ToString(drKB("USELATESTVER"))
        'DOWNLOADRPT '可下載報表
        Dim vDOWNLOADRPT As String = Convert.ToString(drKB("DOWNLOADRPT"))
        '(報表名稱)
        Dim vRPTNAME As String = Convert.ToString(drKB("RPTNAME"))
        '以目前版本批次送出:SENTBATVER
        Dim vSENTBATVER As String = Convert.ToString(drKB("SENTBATVER"))
        '以目前版本送出: SENDCURRVER
        Dim vSENDCURRVER As String = Convert.ToString(drKB("SENDCURRVER"))
        '檔案上傳:UPLOADFL1
        Dim vUPLOADFL1 As String = Convert.ToString(drKB("UPLOADFL1"))
        '備註說明:USEMEMO1
        Dim vUSEMEMO1 As String = Convert.ToString(drKB("USEMEMO1"))
        '申復意見表/訓練班別計畫表'DataGrid08
        Dim vDataGrid08 As String = Convert.ToString(drKB("DataGrid08"))
        '檔案格式說明
        labFILEDESC1.Text = cst_FileDescMsg_7_10M

        Dim str_rtn_checkFile1 As String = String.Concat("return checkFile1(", cst_PostedFile_MAX_SIZE_10M, ");")
        '訓練班別計畫表'DataGrid08
        tr_DataGrid08.Visible = If(Convert.ToString(drKB("DataGrid08")) = "Y", True, False)
        '申復意見表/訓練班別計畫表 'DataGrid08
        If tr_DataGrid08.Visible Then Call SHOW_DATAGRID_08_SF()

        '取得文字說明
        LiteralSwitchTo.Text = If(vKSFID <> "", TIMS.HtmlDecode1(vSFDESC1), "(無)")
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

        '取得SFID代號／非流水號
        Hid_SFID.Value = vSFID 'GET_SFID(vKSFID, vORGKINDGW) GET_KBDESC1(vKSFID, vORGKINDGW)
        LabSwitchTo.Text = If(vKSFID <> "", TIMS.GetListText(ddlSwitchTo), "")
        'USELATESTVER : 以最近一次版本送件
        tr_USELATESTVER.Visible = If(vUSELATESTVER = "Y", True, False)
        'MUSTFILL 必填資訊／WAIVED:免附文件(必填就不顯示)
        tr_WAIVED.Visible = If(vMUSTFILL = "Y", False, True)
        lbWAIVEDESC1.Text = vWAIVEDESC1
        '備註說明:USEMEMO1
        tr_USEMEMO1.Visible = If(vUSEMEMO1 = "Y", True, False)
        '預設值-免附文件
        CHKB_WAIVED.Checked = False '(預設值不填寫)
        '預設值-(上傳檔案)
        Hid_SFCFID.Value = ""
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

        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        Dim drFL As DataRow = TIMS.GET_ORG_SFCASEFL(objconn, Val(Hid_SFCID.Value), Val(vKSFID))
        If drOB Is Nothing OrElse drFL Is Nothing Then Return

        Hid_SFCFID.Value = Convert.ToString(drFL("SFCFID"))
        '免附文件
        CHKB_WAIVED.Checked = If(Convert.ToString(drFL("WAIVED")) = "Y", True, False)

        txtMEMO1.Text = TIMS.ClearSQM(drFL("MEMO1"))

        '修改狀態，且為退件修正
        Dim vSFCSTATUS As String = Convert.ToString(drOB("SFCSTATUS"))
        Dim fg_UPDATE_SFCSTATUS_R As Boolean = (Session(cst_ss_RqProcessType) = cst_DG1CMDNM_EDIT1 AndAlso Hid_SFCFID.Value <> "" AndAlso vSFCSTATUS = "R")
        If fg_UPDATE_SFCSTATUS_R Then
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
            'BTN_SAVETMP1.Enabled = If(fg_LOCK_INPUT, False, True)
            'TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
            '進下一步 '儲存後進下一步
            BTN_SAVENEXT1.Enabled = If(fg_LOCK_INPUT, False, True)
            TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, cst_tpmsg_enb5, cst_tpmsg_enb7), True)
        End If
    End Sub

    ''' <summary>顯示項目資訊／狀態</summary>
    ''' <param name="oConn"></param>
    ''' <param name="vSFCID"></param>
    ''' <param name="vORGKINDGW"></param>
    ''' <returns></returns>
    Private Function GET_ORG_SFCASEFL_TB(ByRef oConn As SqlConnection, vSFCID As String, vORGKINDGW As String) As DataTable
        Dim rParms As New Hashtable From {
            {"SFCID", Val(vSFCID)},
            {"ORGKINDGW", vORGKINDGW}
        }
        Dim rsSql As String = ""
        rsSql &= " SELECT a.SFCFID,ob.YEARS,ob.APPSTAGE,ob.RID,a.SFCID,a.KSFID" & vbCrLf
        rsSql &= " ,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE" & vbCrLf
        rsSql &= " ,kb.SFID,concat(kb.SFID,'.',kb.SFNAME,CASE WHEN a.MEMO1 IS NOT NULL THEN CONCAT('-',a.MEMO1) END,CASE WHEN a.WAIVED='Y' THEN '(免付文件)' END) KBSFNAME" & vbCrLf
        rsSql &= " ,a.RTUREASON" & vbCrLf
        rsSql &= " ,ob.SFCSTATUS,ob.APPLIEDRESULT" & vbCrLf
        rsSql &= " ,a.WAIVED,a.SRCFILENAME1,a.FILENAME1,a.FILENAME1 OKFLAG,a.FILEPATH1" & vbCrLf
        rsSql &= " FROM ORG_SFCASEFL a" & vbCrLf
        rsSql &= " JOIN KEY_SFCASE kb on kb.KSFID=a.KSFID" & vbCrLf
        rsSql &= " JOIN ORG_SFCASE ob on ob.SFCID=a.SFCID" & vbCrLf
        rsSql &= " JOIN AUTH_RELSHIP rr on rr.RID=ob.RID" & vbCrLf
        rsSql &= " WHERE a.SFCID=@SFCID AND kb.ORGKINDGW=@ORGKINDGW" & vbCrLf
        rsSql &= " ORDER BY kb.KSORT,a.SFCFID" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        Return dt
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

    ''' <summary>檔案上傳</summary>
    ''' <param name="drOB"></param>
    Sub FILE_UPLOAD_SF_1(ByRef drOB As DataRow, ByRef drKB As DataRow)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        'Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件資料)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號資料)，請重新操作!!")
            Return
        End If
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text

        '年度／計畫ID／機構ID／caseno／1
        'Dim vUploadPath As String = String.Concat(G_UPDRV, "/", Hid_SFCASENO.Value, "/")
        Dim vYEARS As String = TIMS.ClearSQM(drOB("YEARS")) 'TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.ClearSQM(drOB("APPSTAGE")) 'TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vPLANID As String = TIMS.ClearSQM(drOB("PLANID"))
        Dim vRID As String = TIMS.ClearSQM(drOB("RID")) ' TIMS.GetMyValue2(rPMS, "RID")
        Dim vSFCID As String = TIMS.ClearSQM(drOB("SFCID"))
        Dim vSFCASENO As String = TIMS.ClearSQM(drOB("SFCASENO"))
        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vSFCASENO <> Hid_SFCASENO.Value Then Return '(此狀況不太可能發生)

        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "")
        If vWAIVED = "Y" Then
            Common.MessageBox(Me, cst_errMsg_21)
            Exit Sub
        End If

        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFilePdf(Me, File1, MyPostedFile) Then Return
        'Dim MyFileColl As HttpFileCollection = HttpContext.Current.Request.Files
        'Dim MyPostedFile As HttpPostedFile = MyFileColl.Item(0)
        'If Not MyPostedFile.ContentType.Equals("application/pdf", StringComparison.OrdinalIgnoreCase) Then
        '    Common.MessageBox(Me, cst_errMsg_11_PDF)
        '    Exit Sub
        'ElseIf File1.Value = "" Then
        '    Common.MessageBox(Me, cst_errMsg_8)
        '    Exit Sub
        'ElseIf File1.PostedFile.ContentLength = 0 Then
        '    Common.MessageBox(Me, cst_errMsg_3)
        '    Exit Sub
        'End If

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

        '取得SFID代號／非流水號
        Dim vSFID As String = Convert.ToString(drKB("SFID"))
        Dim vKSFID As String = Convert.ToString(drKB("KSFID"))
        Dim vORGKINDGW As String = Convert.ToString(drKB("ORGKINDGW"))

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, "")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_SF(vSFCID, vKSFID, "pdf")
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
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "SFCFID", If(vUploadPath <> "", iSFCFID, -1)) '(可再次傳送)
            TIMS.SetMyValue2(rPMS2, "SFID", vSFID)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "SFCID", vSFCID)
            TIMS.SetMyValue2(rPMS2, "KSFID", vKSFID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", "")
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", vMODIFYACCT)
            Call SAVE_ORG_SFCASEFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
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
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Hid_KSFID.Value = TIMS.ClearSQM(Hid_KSFID.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'SAVE_ORG_BIDCASEFL
        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, Hid_KSFID.Value, Hid_ORGKINDGW.Value)

        If Hid_SFCASENO.Value = "" OrElse Hid_SFCID.Value = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(申復案件號為空)，請重新操作!")
            Return
        ElseIf drOB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無申復案件資料)，請重新操作!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無申復項目編號資料)，請重新操作!")
            Return
        End If

        '取得SFID代號／非流水號
        'Dim vKSFID As String = Convert.ToString(drKB("KBSID"))
        Dim vSFID As String = Convert.ToString(drKB("SFID"))
        Dim vORGKINDGW As String = Convert.ToString(drKB("ORGKINDGW"))

        Dim vSFCID As String = TIMS.ClearSQM(Hid_SFCID.Value)
        Dim vKSFID As String = TIMS.ClearSQM(Hid_KSFID.Value)
        Dim drFL As DataRow = TIMS.GET_ORG_SFCASEFL(objconn, vSFCID, vKSFID)
        '(退件修正)有退件原因,可重新上傳
        'Dim flag_NG_UPLOAD_1 As Boolean = (drFL IsNot Nothing) '(有資料 不可再次傳送)
        Dim flag_NG_UPLOAD_2 As Boolean = (drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) = "") '(有資料不可傳送且原因為空 不可再次傳送)
        Select Case String.Concat(vORGKINDGW, vSFID)
            Case TIMS.cst_SF_G04_其他佐證文件, TIMS.cst_SF_W04_其他佐證文件
                txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
                If txtMEMO1.Text = "" Then
                    Common.MessageBox(Me, "備註說明不可為空!")
                    Return
                End If
                Dim dtFL As DataTable = TIMS.GET_ORG_SFCASEFL_tb(objconn, vSFCID, vKSFID)
                tryFIND = String.Concat("MEMO1='", txtMEMO1.Text, "'")
                If dtFL IsNot Nothing AndAlso dtFL.Rows.Count > 0 AndAlso dtFL.Select(tryFIND).Length > 0 Then
                    Common.MessageBox(Me, String.Concat("同備註說明，已有上傳資訊!(若要再次上傳，請先刪除)", txtMEMO1.Text))
                    Return
                End If
            Case Else
                Dim vFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
                'Dim vWAIVED As String = If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
                If vFILENAME1 <> "" AndAlso flag_NG_UPLOAD_2 Then
                    '符合所有 不可再次傳送 'cst_tpmsg_enb8
                    Common.MessageBox(Me, "已上傳儲存過該文件，不可再次操作!")
                    Return
                End If
        End Select

        '有錯誤原因 可再次傳送 並記錄 iBCFID
        'Dim iSFCFID As Integer = If(drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) <> "", Val(drFL("SFCFID")), -1)
        '檔案上傳／確定檔案上傳
        Call FILE_UPLOAD_SF_1(drOB, drKB)

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "SFCID", Hid_SFCID.Value)
        Call SHOW_SFCASEFL_DG2(rPMS3)
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument

        Dim vSFCFID As String = TIMS.GetMyValue(sCmdArg, "SFCFID")
        'Dim vSFCID As String = TIMS.GetMyValue(sCmdArg, "SFCID")
        Dim vKSFID As String = TIMS.GetMyValue(sCmdArg, "KSFID")
        Dim vSFID As String = TIMS.GetMyValue(sCmdArg, "SFID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        Dim vRID As String = RIDValue.Value
        Dim vSFCID As String = Hid_SFCID.Value
        Dim vSFCASENO As String = Hid_SFCASENO.Value
        If e.CommandArgument = "" OrElse vSFCFID = "" Then Return

        Select Case e.CommandName
            Case "DELFILE4"
                Dim sErrMsg1 As String = CHKDEL_ORG_SFCASEFL(vSFCFID)
                If sErrMsg1 <> "" Then
                    Common.MessageBox(Me, sErrMsg1)
                    Return
                End If

                '刪除檔案 '"ORG_BIDCASEFL"
                Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, vRID, vSFCID, vSFCASENO)
                If drOB Is Nothing Then Return
                Dim oYEARS As String = Convert.ToString(drOB("YEARS"))
                Dim oAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
                Dim oPLANID As String = Convert.ToString(drOB("PLANID"))
                Dim oRID As String = Convert.ToString(drOB("RID"))
                Dim oSFCASENO As String = Convert.ToString(drOB("SFCASENO"))
                Dim oFILENAME1 As String = ""
                Dim oUploadPath As String = ""
                Dim s_FilePath1 As String = ""
                Try
                    oFILENAME1 = vFILENAME1
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH1_SF(oYEARS, oAPPSTAGE, oPLANID, oRID, oSFCASENO, ""))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try

                '"ORG_BIDCASEFL"
                Dim dParms As New Hashtable From {{"SFCFID", vSFCFID}}
                Dim rdSql As String = "DELETE ORG_SFCASEFL WHERE SFCFID=@SFCFID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)

            Case "DOWNLOAD4" '下載
                Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
                Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
                RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "SFCID", Hid_SFCID.Value)
                TIMS.SetMyValue2(rPMS4, "SFCASENO", Hid_SFCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "SFCFID", vSFCFID)
                TIMS.SetMyValue2(rPMS4, "SFID", vSFID)
                TIMS.SetMyValue2(rPMS4, "KSFID", vKSFID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)
                Call TIMS.ResponseZIPFile_SF(sm, objconn, Me, rPMS4)
                Return

        End Select

        If Not TIMS.OpenDbConn(objconn) Then Return
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)

        '顯示檔案資料表
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "SFCID", Hid_SFCID.Value)
        Call SHOW_SFCASEFL_DG2(rPMS3)

        Dim drRR As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If RIDValue.Value = "" OrElse drRR Is Nothing Then
            Common.MessageBox(Me, "申復送件資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        Call SHOW_Detail_SFCASE(drRR, Hid_SFCID.Value, Session(cst_ss_RqProcessType))
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
                ElseIf Convert.ToString(drv("WAIVED")) = cst_SF_02_申復意見表_WAIVED_PI Then
                    'LabFileName1.Text = cst_txt_版本批次送出
                    titleMsg = cst_txt_版本批次送出
                    'BTN_DELFILE4.Enabled = False
                    'TIMS.Tooltip(BTN_DELFILE4, "清單式資料(無提供刪除)!", True)
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "SFCFID", Convert.ToString(drv("SFCFID")))
                TIMS.SetMyValue(sCmdArg, "SFCID", Convert.ToString(drv("SFCID")))
                TIMS.SetMyValue(sCmdArg, "KSFID", Convert.ToString(drv("KSFID")))
                TIMS.SetMyValue(sCmdArg, "SFID", Convert.ToString(drv("SFID")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", Convert.ToString(drv("FILEPATH1")))
                BTN_DELFILE4.CommandArgument = sCmdArg '刪除
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                BTN_DELFILE4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE4.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)

                '(其他原因調整) '送件／退件修正，不提供刪除
                Dim vSFCSTATUS As String = Convert.ToString(drv("SFCSTATUS"))
                If vSFCSTATUS = "B" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb6, True)
                ElseIf vSFCSTATUS = "R" AndAlso Convert.ToString(drv("RTUREASON")) <> "" Then
                    BTN_DELFILE4.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb8, True)
                ElseIf vSFCSTATUS = "R" AndAlso Convert.ToString(drv("RTUREASON")) = "" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb7, True)
                End If
        End Select
    End Sub

    Private Sub DataGrid08_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid08.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Return
        Dim vSFCFID As String = TIMS.GetMyValue(sCmdArg, "SFCFID")
        Dim vSFID As String = TIMS.GetMyValue(sCmdArg, "SFID")
        Dim vKSFID As String = TIMS.GetMyValue(sCmdArg, "KSFID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")
        Dim vSFCFPID As String = TIMS.GetMyValue(sCmdArg, "SFCFPID")
        Dim vPSOID As String = TIMS.GetMyValue(sCmdArg, "PSOID")
        Dim vPSNO28 As String = TIMS.GetMyValue(sCmdArg, "PSNO28")
        Dim vDISTID As String = TIMS.GetMyValue(sCmdArg, "DISTID")

        Select Case e.CommandName
            Case "DOWNLOAD8"
                Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
                Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
                RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "SFCID", Hid_SFCID.Value)
                TIMS.SetMyValue2(rPMS4, "SFCASENO", Hid_SFCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "RID", RIDValue.Value)
                TIMS.SetMyValue2(rPMS4, "SFCFID", vSFCFID)
                TIMS.SetMyValue2(rPMS4, "SFID", vSFID)
                TIMS.SetMyValue2(rPMS4, "KSFID", vKSFID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)

                TIMS.SetMyValue2(rPMS4, "SFCFPID", vSFCFPID)
                Call TIMS.ResponseZIPFile_SF(sm, objconn, Me, rPMS4)
                Return

            Case "PRINTDG08" '列印"
                Const cst_printFN1 As String = "TC_02_002A"
                If vPSOID = "" OrElse vPSNO28 = "" OrElse vDISTID = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Return
                End If
                Dim MyValue1 As String = ""
                TIMS.SetMyValue(MyValue1, "TPlanID", sm.UserInfo.TPlanID)
                TIMS.SetMyValue(MyValue1, "YEARS", sm.UserInfo.Years)
                TIMS.SetMyValue(MyValue1, "PSOID", vPSOID)
                TIMS.SetMyValue(MyValue1, "PSNO28", vPSNO28)
                TIMS.SetMyValue(MyValue1, "DISTID", vDISTID)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)
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
                Dim BTN_PRINTDG08 As Button = e.Item.FindControl("BTN_PRINTDG08") '列印
                Dim BTN_DOWNLOAD8 As Button = e.Item.FindControl("BTN_DOWNLOAD8") '下載 

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                'Dim LabFileName1 As Label = e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden = e.Item.FindControl("HFileName")
                'LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                'HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()

                '0:未轉班,1:已轉班 '未轉班(依計畫查詢) 含列印
                'BTN_PRINTDG08.Visible = True
                HDG8_PlanID.Value = Convert.ToString(drv("PlanID"))
                HDG8_ComIDNO.Value = Convert.ToString(drv("ComIDNO"))
                HDG8_SeqNo.Value = Convert.ToString(drv("SeqNo"))
                HDG8_OCID.Value = Convert.ToString(drv("OCID"))
                'Dim v_MSD As String = Convert.ToString(drv("MSD"))
                'Dim CYearsVal As String = (sm.UserInfo.Years - 1911)
                Dim titleMsg As String = If(Not IsDBNull(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")), cst_tpmsg_enb9)
                BTN_DOWNLOAD8.Enabled = If(Not IsDBNull(drv("FILENAME1")), (Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG"))), False)
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD8, titleMsg, True)

                Dim sCmdArg8 As String = ""
                TIMS.SetMyValue(sCmdArg8, "SFCFID", Convert.ToString(drv("SFCFID")))
                TIMS.SetMyValue(sCmdArg8, "SFID", Convert.ToString(drv("SFID")))
                TIMS.SetMyValue(sCmdArg8, "KSFID", Convert.ToString(drv("KSFID")))
                TIMS.SetMyValue(sCmdArg8, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg8, "FILEPATH1", Convert.ToString(drv("FILEPATH1")))
                TIMS.SetMyValue(sCmdArg8, "SFCFPID", Convert.ToString(drv("SFCFPID")))
                TIMS.SetMyValue(sCmdArg8, "PSOID", Convert.ToString(drv("PSOID")))
                TIMS.SetMyValue(sCmdArg8, "PSNO28", Convert.ToString(drv("PSNO28")))
                TIMS.SetMyValue(sCmdArg8, "DISTID", Convert.ToString(drv("DISTID")))
                BTN_PRINTDG08.CommandArgument = sCmdArg8 '列印
                BTN_DOWNLOAD8.CommandArgument = sCmdArg8 '檔案下載
        End Select
    End Sub

    ''' <summary>重新查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH2_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH2.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_KSFID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKSFID.Value <> "" Then
            Call SHOW_SFCASE_KSFID(Hid_FirstKSFID.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    ''' <summary>
    ''' 顯示申請機構
    ''' </summary>
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

    ''' <summary>設定下拉項目(產投G／自主W)，取得 rLastKBID，取得 rFirstKBSID</summary>
    ''' <param name="vORGKINDGW"></param>
    Sub Utl_GET_SWITCHTO_VAL(oConn As SqlConnection, ByVal vORGKINDGW As String, ByRef rLastSFID As String, ByRef rFirstKSFID As String)
        If vORGKINDGW = "" Then Return 'rst

        Dim xsPMS As New Hashtable From {{"ORGKINDGW", vORGKINDGW}}
        Dim xSql As String = "SELECT TOP 1 SFID xSFID FROM KEY_SFCASE WHERE ORGKINDGW=@ORGKINDGW ORDER BY KSORT DESC"
        Dim xSFID As String = DbAccess.ExecuteScalar(xSql, oConn, xsPMS)
        rLastSFID = xSFID

        Dim xsPMS2 As New Hashtable From {{"ORGKINDGW", vORGKINDGW}}
        Dim xSql2 As String = "SELECT TOP 1 KSFID xKSFID FROM KEY_SFCASE WHERE ORGKINDGW=@ORGKINDGW ORDER BY KSORT"
        Dim xKSFID As String = DbAccess.ExecuteScalar(xSql2, oConn, xsPMS2)
        rFirstKSFID = xKSFID
    End Sub

    ''' <summary>
    ''' 儲存資料
    ''' </summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_ORG_SFCASEFL_UPLOAD(rPMS As Hashtable)
        Dim iSFCFID As Integer = -1

        Dim vSFID As String = TIMS.GetMyValue2(rPMS, "SFID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        If vSFID = "" OrElse vORGKINDGW = "" Then Return

        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID")
        Dim vKSFID As String = TIMS.GetMyValue2(rPMS, "KSFID")
        Dim vWAIVED As String = TIMS.GetMyValue2(rPMS, "WAIVED")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS, "FILEPATH1")
        Dim vMEMO1 As String = TIMS.GetMyValue2(rPMS, "MEMO1")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim drFL As DataRow = TIMS.GET_ORG_SFCASEFL(objconn, Val(vSFCID), Val(vKSFID))

        '免附文件或上傳檔案
        Dim fg_NG_SAVE As Boolean = (vWAIVED = "" AndAlso (vFILENAME1 = "" OrElse vSRCFILENAME1 = "" OrElse vFILEPATH1 = ""))
        If fg_NG_SAVE Then Return 'iSFCFID

        '取得SFID代號／非流水號
        'Dim vKSFID As String = Convert.ToString(drKB("KBSID"))
        Select Case String.Concat(vORGKINDGW, vSFID)
            Case TIMS.cst_SF_G04_其他佐證文件, TIMS.cst_SF_W04_其他佐證文件
                drFL = Nothing '(永遠新增)
        End Select

        If drFL Is Nothing Then
            'Dim rsSql As String = "Select 1 FROM ORG_SFCASEFL WHERE SFCID=@SFCID And KSFID=@KSFID" & vbCrLf
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_SFCASEFL(SFCFID, SFCID, KSFID, FILENAME1, SRCFILENAME1, FILEPATH1, MEMO1, MODIFYACCT, MODIFYDATE, WAIVED)" & vbCrLf
            isSql &= " VALUES(@SFCFID,@SFCID,@KSFID,@FILENAME1,@SRCFILENAME1,@FILEPATH1 ,@MEMO1,@MODIFYACCT, GETDATE(),@WAIVED)" & vbCrLf
            iSFCFID = DbAccess.GetNewId(objconn, "ORG_SFCASEFL_SFCFID_SEQ, ORG_SFCASEFL, SFCFID")
            Dim iParms As New Hashtable From {
                {"SFCFID", iSFCFID},
                {"SFCID", vSFCID},
                {"KSFID", vKSFID},
                {"WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull)},
                {"FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull)}, 'vFILENAME1)
                {"SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull)}, 'vSRCFILENAME1)
                {"FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull)}, ' vFILEPATH1)
                {"MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull)},
                {"MODIFYACCT", vMODIFYACCT}
            }
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            iSFCFID = Val(drFL("SFCFID"))
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_SFCASEFL" & vbCrLf
            usSql &= " SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,FILEPATH1=@FILEPATH1" & vbCrLf
            usSql &= " ,WAIVED=@WAIVED,MEMO1=@MEMO1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE SFCFID=@SFCFID And SFCID=@SFCID And KSFID=@KSFID" & vbCrLf
            Dim uParms As New Hashtable From {
                {"SFCFID", iSFCFID},
                {"SFCID", vSFCID},
                {"KSFID", vKSFID},
                {"FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull)}, 'vFILENAME1)
                {"SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull)}, 'vSRCFILENAME1)
                {"FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull)}, ' vFILEPATH1)
                {"WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull)},
                {"MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull)},
                {"MODIFYACCT", vMODIFYACCT}
            }
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        End If

    End Sub

    ''' <summary>
    ''' 顯示資料
    ''' </summary>
    Private Sub SHOW_DATAGRID_08_SF()
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "查無申復案件編號資料，請重新操作!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "查無申復項目編號資料，請重新操作!")
            Return
        End If

        Dim vSFCID As String = Hid_SFCID.Value 'TIMS.ClearSQM(Hid_BCID.Value)
        Dim vAPPSTAGE As String = Convert.ToString(drOB("APPSTAGE"))
        Dim vRID As String = Convert.ToString(drOB("RID"))

        '使用登入者業務權限
        Dim sParms1 As New Hashtable From {{"SFCID", vSFCID}, {"RID", vRID}, {"AppStage", vAPPSTAGE}}

        Dim sSql1 As String = ""
        'WBPI / ORG_SFCASEPI
        sSql1 &= " WITH WBPI As (SELECT a.SFCID,a.SFCPID,a.PLANID,a.COMIDNO,a.SEQNO FROM ORG_SFCASEPI a WHERE a.SFCID=@SFCID)" & vbCrLf
        'WFPI2 / ORG_SFCASEFL_PI
        sSql1 &= " ,WFPI2 As (Select a.SFCID,a.SFCPID,a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql1 &= " ,b.SRCFILENAME1,b.FILENAME1,b.FILEPATH1" & vbCrLf
        sSql1 &= " ,b.SFCFPID,b.SFCFID,kb.KSFID,kb.SFID" & vbCrLf
        sSql1 &= " FROM WBPI a" & vbCrLf
        sSql1 &= " JOIN ORG_SFCASEFL_PI b On b.SFCPID=a.SFCPID" & vbCrLf
        sSql1 &= " JOIN ORG_SFCASEFL f On f.SFCFID=b.SFCFID" & vbCrLf
        sSql1 &= " JOIN ORG_SFCASE ob On ob.SFCID=f.SFCID" & vbCrLf
        sSql1 &= " JOIN KEY_SFCASE kb On kb.KSFID=f.KSFID)" & vbCrLf

        'PLAN_PLANINFO a / PLAN_STAFFOPIN pf
        sSql1 &= " SELECT a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
        sSql1 &= " ,dbo.FN_OCID(a.PLANID,a.COMIDNO,a.SEQNO) OCID" & vbCrLf
        sSql1 &= " ,dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType) CLASSCNAME" & vbCrLf
        sSql1 &= " ,concat(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),'-',dbo.FN_CDATE1B(a.STDate)) CLASSCNAMEX" & vbCrLf
        sSql1 &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sSql1 &= " ,b.OrgName ,a.RID ,ip.DISTID,a.PSNO28" & vbCrLf
        'sSql1 &= " ,(SELECT MAX(x.PSOID) FROM PLAN_STAFFOPIN x WHERE x.PSNO28=a.PSNO28 AND x.SFCONTREASONS IS NOT NULL) PSOID" & vbCrLf
        sSql1 &= " ,pf.PSOID" & vbCrLf
        sSql1 &= " ,FORMAT(a.modifydate,'mmssdd') MSD" & vbCrLf
        sSql1 &= " ,p2.SRCFILENAME1,p2.FILENAME1,p2.FILENAME1 OKFLAG,p2.FILEPATH1" & vbCrLf
        sSql1 &= " ,p2.SFCFPID,p2.SFCFID,p2.SFID,p2.KSFID" & vbCrLf
        sSql1 &= " FROM WBPI p1" & vbCrLf
        sSql1 &= " JOIN dbo.PLAN_PLANINFO a ON a.PLANID=p1.PLANID AND a.COMIDNO=p1.COMIDNO AND a.SEQNO=p1.SEQNO" & vbCrLf
        sSql1 &= " JOIN dbo.PLAN_VERREPORT a3 ON a3.PLANID=a.PLANID AND a3.COMIDNO=a.COMIDNO AND a3.SEQNO=a.SEQNO" & vbCrLf
        sSql1 &= " JOIN dbo.VIEW_RIDNAME b ON a.RID = b.RID" & vbCrLf
        sSql1 &= " JOIN dbo.ID_PLAN ip ON ip.PlanID = a.PlanID" & vbCrLf
        sSql1 &= " JOIN dbo.PLAN_STAFFOPIN pf WITH(NOLOCK) ON pf.PSNO28=a.PSNO28 AND LEN(pf.SFCONTREASONS)>1" & vbCrLf '申復理由及說明
        sSql1 &= " LEFT JOIN WFPI2 p2 ON p2.PLANID=a.PLANID AND p2.COMIDNO=a.COMIDNO AND p2.SEQNO=a.SEQNO" & vbCrLf
        '0:未轉班,1:已轉班
        'sSql1 &= " WHERE a.TransFlag='N' AND a.IsApprPaper='Y' AND a.AppliedResult IS NULL AND a.RESULTBUTTON IS NULL" & vbCrLf
        sSql1 &= " WHERE a.ISAPPRPAPER='Y' AND a3.ISAPPRPAPER='Y' AND a.RESULTBUTTON IS NULL" & vbCrLf
        sSql1 &= " AND a.RID=@RID AND a.AppStage=@AppStage" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sParms1.Add("TPlanID", sm.UserInfo.TPlanID)
            sParms1.Add("Years", sm.UserInfo.Years)
            sSql1 &= " AND ip.TPlanID=@TPlanID" & vbCrLf
            sSql1 &= " AND ip.Years =@Years" & vbCrLf
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

        Dim vYEARS As String = Convert.ToString(drOB("YEARS"))
        Dim vPLANID As String = Convert.ToString(drOB("PLANID"))
        'Dim vRID As String = Convert.ToString(drOB("RID"))
        Dim vSFCASENO As String = Convert.ToString(drOB("SFCASENO"))
        Dim vKSFID As String = Convert.ToString(drKB("KSFID"))
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, vKSFID)
        Call TIMS.Check_dtSFCASEFL(Me, dt2, vUploadPath)

        With DataGrid08
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    Private Function CHKDEL_ORG_SFCASEFL(vSFCFID As String) As String
        'Throw New NotImplementedException()
        Return ""
    End Function

    ''' <summary>
    ''' 刪除
    ''' </summary>
    ''' <param name="MyPage"></param>
    ''' <param name="oConn"></param>
    ''' <param name="drRR"></param>
    ''' <param name="drOB"></param>
    Private Sub DELETE_Detail_SFCASE(MyPage As Page, oConn As SqlConnection, drRR As DataRow, drOB As DataRow)
        If drRR Is Nothing OrElse drOB Is Nothing Then Return
        If Convert.ToString(drRR("RID")) <> Convert.ToString(drOB("RID")) Then Return

        Dim vSFCID As String = Convert.ToString(drOB("SFCID"))
        Call DEL_ALL_ORG_SFCASE(oConn, vSFCID)
    End Sub

    Protected Sub BTN_SENTBATVER_Click(sender As Object, e As EventArgs) Handles BTN_SENTBATVER.Click
        'Dim vUploadPath As String = Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        Hid_KSFID.Value = TIMS.ClearSQM(Hid_KSFID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If Hid_SFCASENO.Value = "" OrElse Hid_KSFID.Value = "" Then
            Common.MessageBox(Me, "查無申復案件號為空，請重新操作!!")
            Return
        End If
        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_SFCASE(objconn, Hid_KSFID.Value, Hid_ORGKINDGW.Value)
        If drOB Is Nothing Then
            Common.MessageBox(Me, "查無申復案件編號資料，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "查無申復項目編號資料，請重新操作!!")
            Return
        End If

        Dim vDISTID As String = Convert.ToString(drOB("DISTID"))
        Dim vSFCID As String = Hid_SFCID.Value
        Dim vKSFID As String = Hid_KSFID.Value
        Dim vSFID As String = Convert.ToString(drKB("SFID"))
        Dim vORGKINDGW As String = Convert.ToString(drKB("ORGKINDGW"))
        Select Case String.Concat(vORGKINDGW, vSFID)
            Case TIMS.cst_SF_G02_申復意見表, TIMS.cst_SF_W02_申復意見表
                Dim iSFCFID As Integer = TIMS.GET_ORG_SFCASEFL_iSFCFID(sm, objconn, vSFCID, vKSFID, cst_SF_02_申復意見表_WAIVED_PI, drOB)
                If iSFCFID <= 0 Then Return

                Dim dtFLPI As DataTable = TIMS.GET_ORG_SFCASEFL_PI(objconn, Val(vSFCID))
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
                    Dim vPCS As String = String.Concat(vHDG8_PlanID, "x", vHDG8_ComIDNO, "x", vHDG8_SeqNo)
                    Dim iSFCPID As Integer = TIMS.GET_ORG_SFCASEPI_iSFCPID(objconn, Val(vSFCID), vHDG8_PlanID, vHDG8_ComIDNO, vHDG8_SeqNo)
                    Dim drPP As DataRow = TIMS.GetPCSDate(vPCS, objconn)
                    If iSFCPID <= 0 OrElse drPP Is Nothing Then Return
                    Dim vPSOID As String = Convert.ToString(drPP("PSOID"))
                    Dim vPSNO28 As String = Convert.ToString(drPP("PSNO28"))
                    Dim rPMS As New Hashtable From {
                        {"SFCFID", iSFCFID},
                        {"SFCPID", iSFCPID},
                        {"SFCID", drOB("SFCID")},
                        {"KSFID", drKB("KSFID")},
                        {"PLANID", vHDG8_PlanID},
                        {"COMIDNO", vHDG8_ComIDNO},
                        {"SEQNO", vHDG8_SeqNo},
                        {"MODIFYACCT", sm.UserInfo.UserID}
                    }
                    Call SAVE_ORG_SFCASEFL_PI_08(rPMS)

                    'Dim tkVal As String = ""
                    'TIMS.SetMyValue(tkVal, String.Concat("N1SM", Now.ToString("ss")), Now.ToString("ssmm"))
                    'TIMS.SetMyValue(tkVal, "RID", RIDValue.Value)
                    'TIMS.SetMyValue(tkVal, "BCID", Hid_BCID.Value)
                    'TIMS.SetMyValue(tkVal, "BCASENO", Hid_BCASENO.Value)
                    'TIMS.SetMyValue(tkVal, "ORGKINDGW", Hid_ORGKINDGW.Value)
                    'TIMS.SetMyValue(tkVal, "KBSID", Hid_KBSID.Value)

                    Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
                    tryFIND = String.Concat("SFCPID=", iSFCPID, " AND PlanID=", vHDG8_PlanID, " AND ComIDNO='", vHDG8_ComIDNO, "' AND SeqNo=", vHDG8_SeqNo)
                    If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 AndAlso dtFLPI.Select(tryFIND).Length > 0 Then
                        Dim drFLPI As DataRow = dtFLPI.Select(tryFIND)(0)
                        Dim vMODIFY_DAY As String = Convert.ToString(drFLPI("MODIFY_DAY")) 'MODIFY_DAY
                        Dim vMODIFY_MI As String = Convert.ToString(drFLPI("MODIFY_MI")) 'MODIFY_MI
                        fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
                    End If

                    If fg_RUN_REPORT_1 Then
                        Dim rPMS4 As New Hashtable From {
                            {"TPlanID", sm.UserInfo.TPlanID},
                            {"YEARS", sm.UserInfo.Years},
                            {"PSOID", vPSOID},
                            {"PSNO28", vPSNO28},
                            {"DISTID", vDISTID}
                        }
                        Dim s_RPTURL As String = GET_RPTURL_TC_02_002A(rPMS4)
                        Dim s_PDF_byte As Byte() = Nothing
                        Try
                            Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
                        Catch ex As Exception
                            Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
                            eErrmsg &= String.Concat(", s_RPTURL: ", s_RPTURL)
                            eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
                            eErrmsg &= String.Concat(", rPMS4: ", TIMS.GetMyValue4(rPMS4))
                            TIMS.LOG.Error(eErrmsg, ex)
                            Common.MessageBox(Me, "資料表下載檔案有誤，請確認檔案是否正確!")
                            Return
                        End Try
                        If s_PDF_byte IsNot Nothing Then
                            Dim xPMS As New Hashtable
                            TIMS.SetMyValue2(xPMS, "PLANID", vHDG8_PlanID)
                            TIMS.SetMyValue2(xPMS, "PCS", vPCS)
                            TIMS.SetMyValue2(xPMS, "SFCFID", iSFCFID)
                            TIMS.SetMyValue2(xPMS, "SFCPID", iSFCPID)
                            TIMS.SetMyValue2(xPMS, "MODIFYACCT", sm.UserInfo.UserID)

                            TIMS.SetMyValue2(xPMS, "YEARS", sm.UserInfo.Years)
                            TIMS.SetMyValue2(xPMS, "APPSTAGE", drOB("APPSTAGE"))
                            TIMS.SetMyValue2(xPMS, "RID", drOB("RID"))
                            TIMS.SetMyValue2(xPMS, "SFCASENO", Hid_SFCASENO.Value)
                            TIMS.SetMyValue2(xPMS, "KSFID", Hid_KSFID.Value)
                            TIMS.SetMyValue2(xPMS, "SFCID", Hid_SFCID.Value)

                            Call SAVE_ORG_SFCASEFL_PI_PDF_FILE(xPMS, s_PDF_byte)
                        End If
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
                    Return
                End If
            Case Else
                Dim rParms2 As New Hashtable From {
                    {"SFCID", Val(Hid_SFCID.Value)},
                    {"KSFID", Val(Hid_KSFID.Value)}
                }
                Dim rSql2 As String = "SELECT 1 FROM ORG_SFCASEFL WHERE SFCID=@SFCID AND KSFID=@KSFID"
                Dim drFL2 As DataRow = DbAccess.GetOneRow(rSql2, objconn, rParms2)
                If drFL2 IsNot Nothing Then
                    Common.MessageBox(Me, "已儲存過該文件，不可再次操作!")
                    Return
                End If
        End Select

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable
        TIMS.SetMyValue2(rPMS3, "ORGKINDGW", Hid_ORGKINDGW.Value)
        TIMS.SetMyValue2(rPMS3, "SFCID", Hid_SFCID.Value)
        Call SHOW_SFCASEFL_DG2(rPMS3)
        Call SHOW_SFCASE_KSFID(Hid_KSFID.Value, Hid_ORGKINDGW.Value)
    End Sub

    ''' <summary>
    ''' 儲存PDF
    ''' </summary>
    ''' <param name="rPMS"></param>
    ''' <param name="s_PDF_byte"></param>
    Private Sub SAVE_ORG_SFCASEFL_PI_PDF_FILE(rPMS As Hashtable, s_PDF_byte() As Byte)

        If rPMS Is Nothing Then Return

        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vPCS As String = TIMS.GetMyValue2(rPMS, "PCS")
        Dim vSFCFID As String = TIMS.GetMyValue2(rPMS, "SFCFID")
        Dim vSFCPID As String = TIMS.GetMyValue2(rPMS, "SFCPID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
        Dim vRID As String = TIMS.GetMyValue2(rPMS, "RID")
        Dim vSFCASENO As String = TIMS.GetMyValue2(rPMS, "SFCASENO")
        Dim vKSFID As String = TIMS.GetMyValue2(rPMS, "KSFID")
        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID")

        Dim vUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = "" 'TIMS.GET_FILENAME1_EV(vBCID, vKBSID, vPCS, "pdf")
        Dim vSRCFILENAME1 As String = "" 'vFILENAME1 'Convert.ToString(oSRCFILENAME1)
        '上傳檔案/存檔：檔名
        Try
            vUploadPath = TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, vKSFID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
            vFILENAME1 = TIMS.GET_FILENAME1_SF_PCS(vSFCID, vKSFID, vPCS, "pdf")
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

        Dim sParms3 As New Hashtable From {{"SFCFID", vSFCFID}, {"SFCID", vSFCID}, {"KSFID", vKSFID}, {"SFCPID", vSFCPID}}
        Dim sSql3 As String = "SELECT SFCFPID FROM ORG_SFCASEFL_PI WHERE SFCFID=@SFCFID AND SFCID=@SFCID AND KSFID=@KSFID AND SFCPID=@SFCPID"
        Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)

        Dim iSFCFPID As Integer = 0
        If dt3.Rows.Count = 0 Then
            iSFCFPID = DbAccess.GetNewId(objconn, "ORG_SFCASEFL_PI_SFCFPID_SEQ,ORG_SFCASEFL_PI,SFCFPID")
            'iParms.Add("WAIVED", cst_SF_02_申復意見表_WAIVED_PI)
            Dim iParms As New Hashtable From {
                {"SFCFPID", iSFCFPID},
                {"SFCPID", Val(vSFCPID)},
                {"SFCFID", Val(vSFCFID)},
                {"SFCID", Val(vSFCID)},
                {"KSFID", Val(vKSFID)},
                {"FILENAME1", vFILENAME1},
                {"SRCFILENAME1", vSRCFILENAME1},
                {"FILEPATH1", vUploadPath},
                {"MODIFYACCT", vMODIFYACCT}
            }
            Dim isSql As String = ""
            isSql &= " INSERT INTO ORG_SFCASEFL_PI(SFCFPID, SFCPID, SFCFID, SFCID, KSFID, FILENAME1, SRCFILENAME1, FILEPATH1, MODIFYACCT, MODIFYDATE)" & vbCrLf
            isSql &= " VALUES(@SFCFPID,@SFCPID,@SFCFID,@SFCID,@KSFID,@FILENAME1,@SRCFILENAME1,@FILEPATH1,@MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            iSFCFPID = dt3.Rows(0)("SFCFPID")
            Dim uParms As New Hashtable From {{"FILENAME1", vFILENAME1}, {"SRCFILENAME1", vSRCFILENAME1}, {"FILEPATH1", vUploadPath},
                {"MODIFYACCT", vMODIFYACCT}, {"SFCFPID", iSFCFPID}}
            Dim usSql As String = ""
            usSql &= " UPDATE ORG_SFCASEFL_PI" & vbCrLf
            usSql &= " SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE SFCFPID=@SFCFPID" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        End If

    End Sub

    ''' <summary>
    ''' 儲存班級
    ''' </summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_ORG_SFCASEFL_PI_08(rPMS As Hashtable)
        Dim iSFCFID As Integer = TIMS.GetMyValue2(rPMS, "SFCFID", TIMS.cst_oType_obj)
        Dim iSFCPID As Integer = TIMS.GetMyValue2(rPMS, "SFCPID", TIMS.cst_oType_obj)
        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID")
        Dim vKSFID As String = TIMS.GetMyValue2(rPMS, "KSFID")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(rPMS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(rPMS, "SEQNO")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        'Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        'Dim drFL As DataRow = TIMS.GET_ORG_SFCASEFL(objconn, Val(Hid_SFCID.Value), Val(vKSFID))
        Return
        Dim iSFCFPID As Integer = DbAccess.GetNewId(objconn, "ORG_SFCASEFL_PI_SFCFPID_SEQ,ORG_SFCASEFL_PI,SFCFPID")
        'iParms.Add("FILENAME1", FILENAME1) 'iParms.Add("SRCFILENAME1", SRCFILENAME1) 'iParms.Add("FILEPATH1", FILEPATH1)
        Dim iParms As New Hashtable From {
            {"SFCFPID", iSFCFPID},
            {"SFCPID", iSFCPID},
            {"SFCFID", iSFCFID},
            {"SFCID", Val(vSFCID)},
            {"KSFID", Val(vKSFID)},
            {"WAIVED", cst_SF_02_申復意見表_WAIVED_PI},
            {"MODIFYACCT", vMODIFYACCT}
        }
        Dim isSql As String = ""
        isSql &= " INSERT INTO ORG_SFCASEFL_PI(SFCFPID, SFCPID, SFCFID, SFCID, KSFID, FILENAME1, SRCFILENAME1, FILEPATH1, PATTERN, MEMO1, WAIVED, MODIFYACCT, MODIFYDATE)" & vbCrLf
        isSql &= " VALUES(@SFCFPID,@SFCPID,@SFCFID,@SFCID,@KSFID,@FILENAME1,@SRCFILENAME1,@FILEPATH1,@PATTERN,@MEMO1,@WAIVED,@MODIFYACCT,@MODIFYDATE)" & vbCrLf
    End Sub

    ''' <summary>
    ''' 資料表下載檔
    ''' </summary>
    ''' <param name="rPMS4"></param>
    ''' <returns></returns>
    Private Function GET_RPTURL_TC_02_002A(rPMS4 As Hashtable) As String

        Const cst_printFN1 As String = "TC_02_002A"
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS4, "TPlanID")
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS4, "YEARS")
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS4, "DISTID")
        Dim vPSOID As String = TIMS.GetMyValue2(rPMS4, "PSOID")
        Dim vPSNO28 As String = TIMS.GetMyValue2(rPMS4, "PSNO28")

        Dim sfilename1 As String = "" 'cst_printFN1
        sfilename1 = cst_printFN1

        Dim sMyValue1 As String = ""
        TIMS.SetMyValue(sMyValue1, "TPlanID", sm.UserInfo.TPlanID)
        TIMS.SetMyValue(sMyValue1, "YEARS", sm.UserInfo.Years)
        TIMS.SetMyValue(sMyValue1, "PSOID", vPSOID)
        TIMS.SetMyValue(sMyValue1, "PSNO28", vPSNO28)
        TIMS.SetMyValue(sMyValue1, "DISTID", vDISTID)
        Return ReportQuery.GetReportUrl2(Me, sfilename1, sMyValue1)
    End Function

    ''' <summary>異常狀況有2筆資料產生，刪除後面產生的錯誤資料</summary>
    ''' <param name="oConn"></param>
    ''' <param name="v_PLANID"></param>
    ''' <param name="v_RID"></param>
    ''' <param name="v_APPSTAGE"></param>
    ''' <param name="v_MIN_SFCID"></param>
    Private Sub DEL_ORG_SFCASE_NG(oConn As SqlConnection, v_PLANID As String, v_RID As String, v_APPSTAGE As String, v_MIN_SFCID As String)
        '排除傳入序號
        Dim sParms2B As New Hashtable From {{"PLANID", v_PLANID}, {"RID", v_RID}, {"APPSTAGE", v_APPSTAGE}, {"SFCID", v_MIN_SFCID}}
        Dim sSql2B As String = " SELECT SFCID FROM ORG_SFCASE WHERE PLANID=@PLANID AND RID=@RID AND APPSTAGE=@APPSTAGE AND APPSTAGE=@APPSTAGE AND SFCID!=@SFCID" & vbCrLf
        Dim dt2B As DataTable = DbAccess.GetDataTable(sSql2B, objconn, sParms2B)
        If TIMS.dtNODATA(dt2B) Then Return
        For Each dr1 As DataRow In dt2B.Rows
            Dim vSFCID As String = Convert.ToString(dr1("SFCID"))
            Call DEL_ALL_ORG_SFCASE(oConn, vSFCID)
        Next
    End Sub

    ''' <summary>
    ''' 完全刪除此筆資料
    ''' </summary>
    ''' <param name="oConn"></param>
    ''' <param name="vSFCID"></param>
    Sub DEL_ALL_ORG_SFCASE(oConn As SqlConnection, vSFCID As String)
        Dim dParms As New Hashtable From {{"SFCID", Val(vSFCID)}}
        Dim dsSql As String = ""
        dsSql = " DELETE ORG_SFCASEFL_PI WHERE SFCID=@SFCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = " DELETE ORG_SFCASEFL WHERE SFCID=@SFCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = " DELETE ORG_SFCASEPI WHERE SFCID=@SFCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = " DELETE ORG_SFCASE WHERE SFCID=@SFCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
    End Sub
End Class
