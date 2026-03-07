Imports System.IO

Partial Class TC_13_001
    Inherits AuthBasePage

    '線上核銷送件--委訓單位使用
    'Dim fg_test As Boolean=TIMS.sUtl_ChkTest() '測試
    'CLASS_VERIFYONLINE / KEY_VERIFY / CLASS_VERIFYONLINE_FL

    Const cst_G01_學員補助申請書 As String = "G01"
    Const cst_G02_支付參訓學員補助經費申請表 As String = "G02"
    Const cst_G03_參訓學員出席紀錄一覽表 As String = "G03"
    Const cst_G04_結訓證書清冊 As String = "G04"
    Const cst_G05_學員繳費收據或發票 As String = "G05"
    Const cst_G06_印花稅加銷花或印花稅大額憑證應繳納稅額繳款書 As String = "G06"
    Const cst_G07_學員簽到退及教學日誌 As String = "G07"
    Const cst_G08_學員線上簽到退明細一覽表 As String = "G08"
    Const cst_G09_其他 As String = "G09"
    Const cst_G10_公文 As String = "G10"

    Const cst_W01_學員補助申請書 As String = "W01"
    Const cst_W02_支付參訓學員補助經費申請表 As String = "W02"
    Const cst_W03_參訓學員出席紀錄一覽表 As String = "W03"
    Const cst_W04_結訓證書清冊 As String = "W04"
    Const cst_W05_學員繳費收據或發票 As String = "W05"
    Const cst_W06_印花稅加銷花或印花稅大額憑證應繳納稅額繳款書 As String = "W06"
    Const cst_W07_學員簽到退及教學日誌 As String = "W07"
    Const cst_W08_學員線上簽到退明細一覽表 As String = "W08"
    Const cst_W09_其他 As String = "W09"
    Const cst_W10_公文 As String = "W10"

    '最近一次版本送件
    Const cst_MTYPE_LATEST_SEND1 As String = "MTYPE_LATEST_SEND1"
    '最近一次版本-下載
    Const cst_MTYPE_LATEST_DOWN1 As String = "MTYPE_LATEST_DOWN1"

    ''' <summary>儲存(暫存)</summary>
    Const cst_ACTTYPE_BTN_SAVETMP1 As String = "BTN_SAVETMP1" '儲存(暫存)
    ''' <summary>'儲存後進下一步</summary>
    Const cst_ACTTYPE_BTN_SAVENEXT1 As String = "BTN_SAVENEXT1" '儲存後進下一步

    Dim tryFIND As String = ""
    '以目前版本批次送出
    Const cst_txt_版本批次送出 As String = "(版本批次送出)"
    Const cst_txt_免附文件 As String = "(免附文件)"
    Const cst_REUPLOADED_MSG As String = "(已重新上傳)"

    'sPrintASPX1=String.Concat(cst_printASPX_R, TIMS.Get_MRqID(Me))'Const cst_printASPX_R As String="../../SD/14/SD_14_002_R.aspx?ID="'Dim sPrintASPX1 As String="" 'SD_14_002_R
    'outTYPE: CLSNM,PCSVAL
    'Const cst_outTYPE_CLSNM As String="CLSNM"
    'Const cst_outTYPE_PCSVAL As String="PCSVAL"

    Const cst_ss_RqProcessType As String = "RqProcessType" 'Session(cst_ss_RqProcessType)=cst_DG1CMDNM_VIEW1
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

    'Dim G_UPDRV As String="~/UPDRV"
    'Dim G_UPDRV_JS As String="../../UPDRV"
    Const cst_errMsg_1 As String = "資料有誤請重新查詢!"
    Const cst_errMsg_2 As String = "上傳檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)" 'Const cst_errMsg_2 As String="上傳檔案壓縮時發生錯誤，請重新確認上傳檔案格式!"
    Const cst_errMsg_3 As String = "檔案位置錯誤!"
    Const cst_errMsg_4 As String = "檔案類型錯誤!"
    Const cst_errMsg_5 As String = "檔案類型錯誤，必須為PDF類型檔案!"
    Const cst_errMsg_5b As String = "檔案類型錯誤，內容必須為PDF檔案!"
    Const cst_errMsg_6 As String = "(檔案上傳失敗／異常，請刪除後重新上傳)"
    Const cst_PostedFile_MAX_SIZE_10M As Integer = 10485760 '10*1024*1024=10,485,760  '2*1024*1024=2,097,152
    Const cst_PostedFile_MAX_SIZE_15M As Integer = 15728640 '1024*1024*15=15728640
    'Const cst_errMsg_7 As String="檔案大小超過2MB!"
    Const cst_errMsg_7_10M As String = "檔案大小超過10MB!"
    Const cst_errMsg_7_15M As String = "檔案大小超過15MB!"
    Const cst_FileDescMsg_7_10M As String = "PDF(掃瞄畫面需清楚，檔案大小限制10MB以下)!"
    Const cst_FileDescMsg_7_15M As String = "PDF(掃瞄畫面需清楚，檔案大小限制15MB以下)!"

    Const cst_errMsg_8 As String = "請選擇上傳檔案(不可為空)!"
    'Const cst_errMsg_9 As String="請選擇場地圖片--隸屬於教室1 或教室2!"
    'Const cst_errMsg_11 As String="無效的檔案格式。"
    Const cst_errMsg_11_PDF As String = "無效的檔案格式(限PDF檔案)。"
    Const cst_errMsg_21 As String = "不可勾選免附文件又按上傳檔案。"

    ''Add new application cases and add reminder messages
    ''Const Cst_messages1 As String="請務必確認此年度/申請階段之所有欲研提班級都已送審，【新增申辦案件】後才送審的班級，將無法納入此次線上申辦案件清單中!"
    ''Const cst_ss_messages1 As String="messages1"

    Dim tmpMSG As String = ""
    Dim objconn As SqlConnection = Nothing

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        '每次執行
        Call CCreate11()

        If Not IsPostBack Then
            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            Call CCreate1(0)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button3.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            'center.Attributes("onclick")="showObj('HistoryList2');ShowFrame();" 'HistoryRID.Attributes("onclick")="ShowFrame();"
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
    End Sub

    ''' <summary>
    ''' 每次執行
    ''' </summary>
    Sub CCreate11()
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        'sPrintASPX1=String.Concat(cst_printASPX_R, TIMS.Get_MRqID(Me))
        '<add key="UPLOAD_OJT_Path" value="~/UPDRV" />
        '<add key="DOWNLOAD_OJT_Path" value="../../UPDRV" />
        'Dim vUPLOAD_OJT_Path As String=TIMS.Utl_GetConfigSet("UPLOAD_OJT_Path")
        'Dim vDOWNLOAD_OJT_Path As String=TIMS.Utl_GetConfigSet("DOWNLOAD_OJT_Path")
        'Dim G_UPDRV As String="~/UPDRV"  'Dim G_UPDRV_JS As String="../../UPDRV"
        'If (vUPLOAD_OJT_Path <> "") Then G_UPDRV=vUPLOAD_OJT_Path
        'If (vDOWNLOAD_OJT_Path <> "") Then G_UPDRV_JS=vDOWNLOAD_OJT_Path
    End Sub

    ''' <summary>
    ''' 呼叫頁面時執行第1次'設定 資料與顯示 狀況！
    ''' </summary>
    ''' <param name="iNum"></param>
    Private Sub CCreate1(ByVal iNum As Integer)

        labmsg1.Text = ""
        TableDataGrid1.Visible = False

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        Call SHOW_Frame1(0)

        'Dim MRqID As String=TIMS.Get_MRqID(Me)
        'TIMS.Get_TitleLab(objconn, MRqID, TitleLab1, TitleLab2)

        Call UseKeepSearch_SD13003()
    End Sub

    ''' <summary>
    ''' '帶入查詢參數
    ''' </summary>
    Sub UseKeepSearch_SD13003()
        '帶入查詢參數
        Dim MyVale As String = ""
        If Session("_Search") IsNot Nothing AndAlso Convert.ToString(Session("_Search")) <> "" Then
            MyVale = TIMS.GetMyValue(Session("_Search"), "prg")
            If MyVale = "SD13003" Then
                center.Text = TIMS.GetMyValue(Session("_Search"), "center")
                RIDValue.Value = TIMS.GetMyValue(Session("_Search"), "RIDValue")
                TMID1.Text = TIMS.GetMyValue(Session("_Search"), "TMID1")
                OCID1.Text = TIMS.GetMyValue(Session("_Search"), "OCID1")
                TMIDValue1.Value = TIMS.GetMyValue(Session("_Search"), "TMIDValue1")
                OCIDValue1.Value = TIMS.GetMyValue(Session("_Search"), "OCIDValue1")
                If TIMS.GetMyValue(Session("_Search"), "Button1") = "TRUE" Then
                    Session("_Search") = Nothing
                    'Button1_Click(sender, e)
                    Call SSearch1()
                End If
            End If
            Session("_Search") = Nothing
        End If
    End Sub

    ''' <summary>顯示調整 0 查詢 1 顯示</summary>
    ''' <param name="iNum"></param>
    Private Sub SHOW_Frame1(ByVal iNum As Integer)
        FrameTableSch1.Visible = If(iNum = 0, True, False)
        FrameTableEdt1.Visible = If(iNum = 1, True, False)
    End Sub

    ''' <summary>
    ''' 清理隱藏的參數
    ''' </summary>
    Sub ClearHidValue()
        Hid_ORGKINDGW.Value = ""

        Hid_KVSID.Value = ""
        Hid_KVID.Value = ""
        Hid_LastKVID.Value = ""
        Hid_FirstKVSID.Value = ""

        Hid_CVOCFID.Value = ""
        Hid_CVOCID.Value = ""
        Hid_OCIDVal.Value = ""
        Hid_SEQ_ID.Value = ""
    End Sub

    ''' <summary>查詢鈕1</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH1_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH1.Click

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call SSearch1()

    End Sub
    '查詢原因-INQUIRY
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        If center.Text <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Text)
        If OCID1.Text <> "" Then RstMemo &= String.Concat("&職類/班別=", OCID1.Text)
        If sch_txtSENDDATE1.Text <> "" Then RstMemo &= String.Concat("&申辦日期1=", sch_txtSENDDATE1.Text)
        If sch_txtSENDDATE2.Text <> "" Then RstMemo &= String.Concat("&申辦日期2=", sch_txtSENDDATE2.Text)
        Return RstMemo
    End Function
    ''' <summary>
    ''' 查詢1
    ''' </summary>
    Private Sub SSearch1()
        '清理隱藏的參數
        Call ClearHidValue()

        labmsg1.Text = TIMS.cst_NODATAMsg1
        TableDataGrid1.Visible = False

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'sch_txtSENDACCTNAME.Text=TIMS.ClearSQM(sch_txtSENDACCTNAME.Text)
        sch_txtSENDDATE1.Text = TIMS.Cdate3(sch_txtSENDDATE1.Text)
        sch_txtSENDDATE2.Text = TIMS.Cdate3(sch_txtSENDDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(sch_txtSENDDATE1.Text, sch_txtSENDDATE2.Text) Then
            Dim T_DATE1 As String = sch_txtSENDDATE1.Text
            sch_txtSENDDATE1.Text = sch_txtSENDDATE2.Text
            sch_txtSENDDATE2.Text = T_DATE1
        End If

        Dim pms_s1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"RIDValue", RIDValue.Value}}
        Dim sSql As String = ""
        sSql &= " SELECT a.CVOCID, a.OCID,a.SEQ_ID,a.APPLIEDRESULT" & vbCrLf
        sSql &= " ,IP.YEARS,dbo.FN_CYEAR2(ip.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,pp.APPSTAGE,CASE pp.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N" & vbCrLf
        sSql &= " ,oo.ORGNAME" & vbCrLf
        sSql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME2" & vbCrLf
        'sSql &= " /* '申辦狀態：暫存/ 已送件 */" & vbCrLf
        sSql &= " ,a.SENDSTATUS ,CASE WHEN a.SENDSTATUS IS NULL THEN '暫存'" & vbCrLf
        sSql &= "  WHEN a.SENDSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'" & vbCrLf
        sSql &= "  WHEN a.SENDSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'" & vbCrLf
        sSql &= "  WHEN a.SENDSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '通過'" & vbCrLf
        sSql &= "  WHEN a.SENDSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'" & vbCrLf
        sSql &= "  WHEN a.SENDSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END SENDSTATUS_N" & vbCrLf
        sSql &= " ,a.SENDDATE,dbo.FN_CDATE1B(a.SENDDATE) SENDDATE_ROC,a.SENDACCT,aa.NAME SENDACCTNAME" & vbCrLf
        sSql &= " ,a.RESULTACCT,a.RESULTDATE,a.REASONFORFAIL"
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        sSql &= " ,a.APPLIEDRESULT"
        sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf

        sSql &= " FROM CLASS_VERIFYONLINE a" & vbCrLf
        sSql &= " JOIN CLASS_CLASSINFO cc ON cc.OCID=a.OCID" & vbCrLf
        sSql &= " JOIN PLAN_PLANINFO pp ON cc.PLANID=pp.PLANID AND cc.COMIDNO=pp.COMIDNO AND cc.SEQNO=pp.SEQNO" & vbCrLf
        sSql &= " JOIN ID_PLAN ip on ip.PLANID=cc.PLANID" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.COMIDNO=cc.COMIDNO" & vbCrLf
        sSql &= " LEFT JOIN AUTH_ACCOUNT aa on aa.ACCOUNT=a.SENDACCT" & vbCrLf
        sSql &= " WHERE pp.IsApprPaper='Y' AND cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
        sSql &= " AND ip.TPLANID=@TPLANID and ip.YEARS=@YEARS" & vbCrLf
        sSql &= " AND cc.RID=@RIDValue" & vbCrLf
        If OCIDValue1.Value <> "" Then
            pms_s1.Add("OCID", OCIDValue1.Value)
            sSql &= " and cc.OCID=@OCID" & vbCrLf
        End If
        'If sch_txtSENDACCTNAME.Text <> "" Then
        '    pms_s1.Add("SENDACCTNAME", sch_txtSENDACCTNAME.Text)
        '    sSql &= " and aa.NAME like '%'+@SENDACCTNAME+'%'" & vbCrLf
        'End If
        If sch_txtSENDDATE1.Text <> "" Then
            pms_s1.Add("SENDDATE1", sch_txtSENDDATE1.Text)
            sSql &= " and a.SENDDATE>=@SENDDATE1" & vbCrLf
        End If
        If sch_txtSENDDATE2.Text <> "" Then
            pms_s1.Add("SENDDATE2", sch_txtSENDDATE2.Text)
            sSql &= " and a.SENDDATE<=@SENDDATE2" & vbCrLf
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, pms_s1)

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        Dim sMemo As String = GET_SEARCH_MEMO()
        '查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "YEARS_ROC,APPSTAGE_N,ORGNAME,CLASSCNAME2,SENDDATE_ROC,SENDSTATUS_N,APPLIEDRESULT_N")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        If TIMS.dtNODATA(dt) Then
            labmsg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        labmsg1.Text = ""
        TableDataGrid1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
        'DataGrid1.DataSource=dt
        'DataGrid1.DataBind()
    End Sub

    ''' <summary>切換至(文件項目)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub ddlSwitchTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSwitchTo.SelectedIndexChanged
        Hid_KVSID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_KVSID.Value)
        ElseIf Hid_FirstKVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_FirstKVSID.Value)
        End If
    End Sub

    ''' <summary>顯示-文件項目-資訊</summary>
    ''' <param name="vKVSID"></param>
    Private Sub SHOW_KEY_VERIFY_KVSID(vKVSID As String)
        vKVSID = TIMS.ClearSQM(vKVSID)
        Dim v_ddlSwitchTo As String = TIMS.GetListValue(ddlSwitchTo)
        If (v_ddlSwitchTo <> vKVSID) Then
            Hid_KVSID.Value = vKVSID
            Common.SetListItem(ddlSwitchTo, vKVSID)
        End If

        'Dim pms_sw As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"ORGKINDGW", Hid_ORGKINDGW.Value}}
        Dim vTPLANID As String = sm.UserInfo.TPlanID
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim drKV As DataRow = TIMS.GET_KEY_VERIFY(objconn, vKVSID, vTPLANID, vORGKINDGW)
        If drKV Is Nothing Then Return

        Dim vKVID As String = Convert.ToString(drKV("KVID"))
        '取得文字說明
        Dim vKBDESC1 As String = Convert.ToString(drKV("KBDESC1"))
        '(Y:不使用KBDESC)
        Dim vNOTKBDESC1 As String = Convert.ToString(drKV("NOTKBDESC1"))
        '(Y:不使用FLDESC1)
        Dim vNOTFLDESC1 As String = Convert.ToString(drKV("NOTFLDESC1"))
        '必填資訊／免附文件(必填就不顯示)
        Dim vMUSTFILL As String = Convert.ToString(drKV("MUSTFILL"))
        'USELATESTVER : 以最近一次版本送件
        Dim vUSELATESTVER As String = Convert.ToString(drKV("USELATESTVER"))
        'DOWNLOADRPT '可下載報表
        Dim vDOWNLOADRPT As String = Convert.ToString(drKV("DOWNLOADRPT"))
        '(報表名稱)
        Dim vRPTNAME As String = Convert.ToString(drKV("RPTNAME"))
        '以目前版本批次送出:SENTBATVER
        Dim vSENTBATVER As String = Convert.ToString(drKV("SENTBATVER"))
        '以目前版本送出: SENDCURRVER
        Dim vSENDCURRVER As String = Convert.ToString(drKV("SENDCURRVER"))
        '檔案上傳:UPLOADFL1
        Dim vUPLOADFL1 As String = Convert.ToString(drKV("UPLOADFL1"))
        '備註說明:USEMEMO1
        Dim vUSEMEMO1 As String = Convert.ToString(drKV("USEMEMO1"))
        '檔案格式說明
        labFILEDESC1.Text = cst_FileDescMsg_7_10M

        Dim str_rtn_checkFile1 As String = String.Concat("return checkFile1(", cst_PostedFile_MAX_SIZE_10M, ");")

        '取得文字說明
        LiteralSwitchTo.Text = If(vKBDESC1 <> "", TIMS.HtmlDecode1(vKBDESC1), "(無)")
        '(Y:不使用KBDESC)
        tr_LiteralSwitchTo.Visible = If(vNOTKBDESC1 = "Y", False, True)
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

        '取得 KVID 代號／非流水號
        Hid_KVID.Value = vKVID 'GET_KBID(vKBSID, vORGKINDGW) GET_KBDESC1(vKBSID, vORGKINDGW)
        LabSwitchTo.Text = If(vKVSID <> "", TIMS.GetListText(ddlSwitchTo), "")
        'USELATESTVER : 以最近一次版本送件
        tr_USELATESTVER.Visible = If(vUSELATESTVER = "Y", True, False)
        'MUSTFILL 必填資訊／WAIVED:免附文件(必填就不顯示)
        tr_WAIVED.Visible = If(vMUSTFILL = "Y", False, True)
        '備註說明:USEMEMO1
        tr_USEMEMO1.Visible = If(vUSEMEMO1 = "Y", True, False)
        '預設值-免附文件
        CHKB_WAIVED.Checked = False '(預設值不填寫)
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

        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        Dim vCVOCID As String = Hid_CVOCID.Value
        Dim vOCIDVal As String = Hid_OCIDVal.Value
        Dim vSEQ_ID As String = Hid_SEQ_ID.Value
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCIDVal, vSEQ_ID)
        If drCV Is Nothing Then Return
        If Hid_CVOCID.Value = "" Then Hid_CVOCID.Value = drCV("CVOCID")
        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, Hid_CVOCID.Value, Hid_KVSID.Value)
        If drFL Is Nothing Then Return

        Hid_CVOCFID.Value = Convert.ToString(drFL("CVOCFID"))
        Hid_CVOCID.Value = Convert.ToString(drCV("CVOCID"))
        Hid_OCIDVal.Value = Convert.ToString(drCV("OCID"))
        Hid_SEQ_ID.Value = Convert.ToString(drCV("SEQ_ID"))
        '免附文件
        CHKB_WAIVED.Checked = If(Convert.ToString(drFL("WAIVED")) = "Y", True, False)

        txtMEMO1.Text = TIMS.ClearSQM(drFL("MEMO1"))

        '修改狀態，且為退件修正
        Dim fg_UPDATE_SENDSTATUS_R As Boolean = (Session(cst_ss_RqProcessType) = cst_DG1CMDNM_EDIT1 AndAlso Hid_CVOCFID.Value <> "" AndAlso Convert.ToString(drCV("SENDSTATUS")) = "R")
        If fg_UPDATE_SENDSTATUS_R Then
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

    ''' <summary>本班尚未進行補助申請！</summary>
    ''' <param name="v_OCIDVal"></param>
    ''' <returns></returns>
    Function CHK_SUBSIDYCOST_NODATA(v_OCIDVal As String) As Boolean
        Dim pms_1 As New Hashtable From {{"OCID", v_OCIDVal}}
        Dim sSql As String = ""
        sSql &= " SELECT a.SOCID" & vbCrLf
        sSql &= " FROM dbo.V_STUDENTINFO a" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBSIDYCOST sc ON sc.SOCID=a.SOCID" & vbCrLf
        sSql &= " WHERE a.STUDSTATUS NOT IN (2,3)" & vbCrLf
        sSql &= " AND a.OCID=@OCID" & vbCrLf
        Dim dtSS As DataTable = DbAccess.GetDataTable(sSql, objconn, pms_1)
        Return TIMS.dtNODATA(dtSS)
    End Function

    ''' <summary>新增</summary>
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
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
            Return
        End If

        Dim fg_SUB1 As Boolean = CHK_SUBSIDYCOST_NODATA(OCIDValue1.Value)
        If fg_SUB1 Then
            Common.MessageBox(Me, "本班尚未進行補助申請！")
            Return
        End If

        'OCIDValue1.Value=TIMS.ClearSQM(OCIDValue1.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        Hid_OCIDVal.Value = Convert.ToString(drCC("OCID")) 'Value1.Value
        If Hid_SEQ_ID.Value = "" Then Hid_SEQ_ID.Value = "1"

        Hid_CVOCFID.Value = GET_VERIFYONLINE_CVOCFID(objconn, Hid_OCIDVal.Value, Hid_SEQ_ID.Value)

        If Hid_CVOCFID.Value = "0" Then Hid_CVOCFID.Value = ""

        If Hid_CVOCFID.Value = "" OrElse Hid_OCIDVal.Value = "" OrElse Hid_SEQ_ID.Value = "" Then
            String.Concat("資訊有誤(查無職類/班別代碼)，請選擇職類/班別,,{0},{1},{2}!!", Hid_CVOCFID.Value, Hid_OCIDVal.Value, Hid_SEQ_ID.Value)
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!!")
            Return
        End If

        '依目前核銷代號查詢詳細資料
        Call SHOW_Detail_VERIFYONLINE(drCC, Hid_OCIDVal.Value, Hid_SEQ_ID.Value, cst_DG1CMDNM_EDIT1)
    End Sub

    ''' <summary>新增／取得-線上核銷送件-案件號</summary>
    ''' <param name="oConn"></param>
    ''' <param name="OCIDVal"></param>
    ''' <param name="SEQ_ID"></param>
    ''' <returns></returns>
    Private Function GET_VERIFYONLINE_CVOCFID(oConn As SqlConnection, OCIDVal As String, SEQ_ID As String) As Integer
        'SAVE_CLASS_VERIFYONLINE
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(oConn, OCIDVal, SEQ_ID)
        Dim iCVOCID As Integer = 0
        If drCV Is Nothing Then
            iCVOCID = DbAccess.GetNewId(oConn, "CLASS_VERIFYONLINE_CVOCID_SEQ,CLASS_VERIFYONLINE,CVOCID")
            'iParms.Add("CREATEDATE", CREATEDATE) 'iParms.Add("MODIFYDATE", MODIFYDATE)
            Dim iParms As New Hashtable From {
                {"CVOCID", iCVOCID},
                {"OCID", TIMS.CINT1(OCIDVal)},
                {"SEQ_ID", TIMS.CINT1(SEQ_ID)},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim sSql_i As String = ""
            sSql_i &= " INSERT INTO CLASS_VERIFYONLINE(CVOCID,OCID,SEQ_ID,CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE )" & vbCrLf
            sSql_i &= " VALUES (@CVOCID,@OCID,@SEQ_ID, @CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(sSql_i, oConn, iParms)
            Return iCVOCID
        End If
        iCVOCID = drCV("CVOCID")
        Return iCVOCID
    End Function

    ''' <summary>下拉選項</summary>
    ''' <param name="oConn"></param>
    ''' <param name="ddlobj"></param>
    ''' <param name="htPMS_r"></param>
    ''' <returns></returns>
    Function GET_ddlVERIFY(oConn As SqlConnection, ddlobj As DropDownList, htPMS_r As Hashtable) As DropDownList
        ddlobj.Items.Clear()
        Dim vTPLANID As String = TIMS.GetMyValue2(htPMS_r, "TPLANID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(htPMS_r, "ORGKINDGW")
        If vTPLANID = "" OrElse vORGKINDGW = "" Then Return ddlobj

        Dim sPMS As New Hashtable From {{"TPLANID", vTPLANID}, {"ORGKINDGW", vORGKINDGW}}
        Dim sSql As String = ""
        sSql &= " SELECT KVSID,concat(KVID,'.',KBNAME) KBNAME,KSORT,ORGKINDGW,TPLANID" & vbCrLf
        sSql &= " FROM KEY_VERIFY" & vbCrLf
        sSql &= " WHERE TPLANID=@TPLANID AND ORGKINDGW=@ORGKINDGW" & vbCrLf
        sSql &= " ORDER BY KSORT" & vbCrLf
        DbAccess.MakeListItem(ddlobj, sSql, oConn, sPMS)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose2, ""))
        Return ddlobj
    End Function

    ''' <summary>依目前核銷代號查詢詳細資料-查詢並顯示</summary>
    ''' <param name="drCC"></param>
    ''' <param name="vOCID1"></param>
    ''' <param name="vSEQ_ID"></param>
    Private Sub SHOW_Detail_VERIFYONLINE(drCC As DataRow, vOCID1 As String, vSEQ_ID As String, vCmdName As String)
        If drCC Is Nothing Then Return
        If vOCID1 = "" AndAlso vSEQ_ID = "" Then Return
        Session(cst_ss_RqProcessType) = vCmdName

        SHOW_Frame1(1)

        labOrgNAME.Text = Convert.ToString(drCC("ORGNAME"))
        'labBIYEARS.Text=Convert.ToString(drCC("YEARS")) 'labBIYEARS.Text=Convert.ToString(drCC("PLANYEARS"))
        labSEND_YEARS_ROC.Text = Convert.ToString(drCC("YEARS_ROC"))
        Dim vAPPSTAGE2 As String = Convert.ToString(drCC("APPSTAGE"))
        labAPPSTAGE.Text = TIMS.GET_APPSTAGE2_NM2(vAPPSTAGE2)
        labCLASSNAME2S.Text = Convert.ToString(drCC("CLASSCNAME2"))
        Hid_ORGKINDGW.Value = Convert.ToString(drCC("ORGKINDGW"))

        Dim rLastKVID As String = ""
        Dim rFirstKVSID As String = ""
        Call Utl_GET_ddlVERIFY_VALS(sm, objconn, Hid_ORGKINDGW.Value, rLastKVID, rFirstKVSID)
        Hid_LastKVID.Value = rLastKVID
        Hid_FirstKVSID.Value = rFirstKVSID

        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCID1, vSEQ_ID)
        Hid_CVOCID.Value = TIMS.ClearSQM(drCV("CVOCID"))
        Hid_OCIDVal.Value = TIMS.ClearSQM(drCV("OCID"))
        Hid_SEQ_ID.Value = TIMS.ClearSQM(drCV("SEQ_ID"))
        Dim vCVOCID As String = TIMS.ClearSQM(drCV("CVOCID"))
        '檢視目前上傳檔案 '顯示上傳檔案／細項 -'線上申辦進度
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)

        Dim pms_sw As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"ORGKINDGW", Hid_ORGKINDGW.Value}}
        ddlSwitchTo = GET_ddlVERIFY(objconn, ddlSwitchTo, pms_sw)
        If Hid_KVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_KVSID.Value)
        ElseIf Hid_FirstKVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_FirstKVSID.Value)
        End If
    End Sub

    ''' <summary>取得項目list 最大最小值</summary>
    ''' <param name="sm"></param>
    ''' <param name="oConn"></param>
    ''' <param name="vORGKINDGW"></param>
    ''' <param name="rLastKVID"></param>
    ''' <param name="rFirstKVSID"></param>
    Private Sub Utl_GET_ddlVERIFY_VALS(sm As SessionModel, oConn As SqlConnection, vORGKINDGW As String, ByRef rLastKVID As String, ByRef rFirstKVSID As String)
        If vORGKINDGW = "" Then Return 'rst
        'DECLARE @ORGKINDGW VARCHAR(10)='G'; 'DECLARE @TPLANID VARCHAR(10)='28';
        Dim xsPMS As New Hashtable From {{"ORGKINDGW", vORGKINDGW}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim xSql As String = "SELECT TOP 1 KVID xKVID FROM KEY_VERIFY WHERE ORGKINDGW=@ORGKINDGW AND TPLANID=@TPLANID ORDER BY KSORT DESC"
        Dim xKVID As String = DbAccess.ExecuteScalar(xSql, oConn, xsPMS)
        rLastKVID = xKVID

        Dim xsPMS2 As New Hashtable From {{"ORGKINDGW", vORGKINDGW}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim xSql2 As String = "SELECT TOP 1 KVSID FROM KEY_VERIFY WHERE ORGKINDGW=@ORGKINDGW AND TPLANID=@TPLANID ORDER BY KSORT"
        Dim xKVSID As String = DbAccess.ExecuteScalar(xSql2, oConn, xsPMS2)
        rFirstKVSID = xKVSID
    End Sub

    ''' <summary>重新查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SEARCH2_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH2.Click
        Hid_KVSID.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_KVSID.Value)
        ElseIf Hid_FirstKVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_FirstKVSID.Value)
        End If
    End Sub

    ''' <summary>
    ''' 回上一步
    ''' </summary>
    Private Sub MOVE_PREV()
        If (Hid_KVID.Value = "" OrElse Hid_KVID.Value = "01" OrElse ddlSwitchTo.SelectedIndex - 1 = -1) Then
            Common.MessageBox(Me, "(目前沒有上一步)")
            Return
        End If

        Hid_KVSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex - 1).Value
        If Hid_KVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_KVSID.Value)
        ElseIf Hid_FirstKVSID.Value <> "" Then
            Call SHOW_KEY_VERIFY_KVSID(Hid_FirstKVSID.Value)
        End If
    End Sub

    ''' <summary>儲存</summary>
    ''' <param name="s_ACTTYPE"></param>
    Private Sub SAVEDATA2_BTN_ACTION1(s_ACTTYPE As String)
        's_ACTTYPE : "BTN_SAVETMP1" '儲存(暫存)／ 儲存後進下一步
        Const cst_ACTTYPE_BTN_SAVENEXT1 As String = "BTN_SAVENEXT1" '儲存後進下一步

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)

        Hid_KVSID.Value = TIMS.ClearSQM(Hid_KVSID.Value)
        Hid_KVID.Value = TIMS.ClearSQM(Hid_KVID.Value)
        Hid_LastKVID.Value = TIMS.ClearSQM(Hid_LastKVID.Value)
        Hid_FirstKVSID.Value = TIMS.ClearSQM(Hid_FirstKVSID.Value)

        Hid_CVOCFID.Value = TIMS.ClearSQM(Hid_CVOCFID.Value)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)

        'SAVE,,CLASS_VERIFYONLINE
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, Hid_OCIDVal.Value, Hid_SEQ_ID.Value, Hid_CVOCID.Value)
        If drCV Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無課程編號)，請重新操作!")
            Return
        End If
        Dim vCVOCID As String = Convert.ToString(drCV("CVOCID"))

        'Dim vKVSID As String=Hid_KVSID.Value
        'Dim pms_sw As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"ORGKINDGW", Hid_ORGKINDGW.Value}}
        'Dim drKV As DataRow=GET_KEY_VERIFY(objconn, Hid_KVSID.Value, pms_sw)
        Dim vKVSID As String = TIMS.ClearSQM(Hid_KVSID.Value)
        Dim vTPLANID As String = sm.UserInfo.TPlanID
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim drKV As DataRow = TIMS.GET_KEY_VERIFY(objconn, vKVSID, vTPLANID, vORGKINDGW)
        If drKV Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!")
            Return
        End If

        'Dim fg_FILE_MUSTBE_UPLOADED As Boolean=True '必須上傳檔案
        Dim fg_FILE_MUSTBE_UPLOADED As Boolean = True
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "") '免附文件
        'Dim vKVSID As String=Convert.ToString(drKV("KVSID"))
        Dim vKVID As String = Convert.ToString(drKV("KVID"))

        'Dim vBCID As String=TIMS.ClearSQM(Hid_BCID.Value)
        ''Dim vKBSID As String=TIMS.ClearSQM(Hid_KBSID.Value)
        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID)
        Dim vFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
        If vFILENAME1 = "" AndAlso Not CHKB_WAIVED.Checked Then
            Common.MessageBox(Me, "未上傳檔案也未勾選免附，不可再次操作!")
            Return
        End If

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vYEARS As String = Convert.ToString(drCV("YEARS"))
        Dim vPLANID As String = Convert.ToString(drCV("PLANID"))
        Dim vCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
        Dim vSEQNO As String = Convert.ToString(drCV("SEQNO"))
        'Dim vBCASENO As String=Convert.ToString(drOB("BCASENO"))
        'Dim vUploadPath As String=TIMS.GET_UPLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")
        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "CVOCID", vCVOCID)
            'TIMS.SetMyValue2(rPMS2, "KVSID", vKVSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED) ' If(CHKB_WAIVED.Checked, "Y", ""))
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            Select Case vWAIVED
                Case "Y", ""
                    '儲存／上傳-CLASS_VERIFYONLINE_FL
                    Call SAVE_CLASS_VERIFYONLINE_FL_UPLOAD(vCVOCID, vKVSID, rPMS2)
            End Select
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub  'Throw ex
        End Try

        '暫時儲存／正式儲存-UPDATE CLASS_VERIFYONLINE 
        'Call SAVEDATE1(0)

        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)

        Select Case s_ACTTYPE
            Case cst_ACTTYPE_BTN_SAVETMP1
                '儲存(暫存) 

                '項目(重跑1次)
                Call SHOW_KEY_VERIFY_KVSID(Hid_KVSID.Value)

            Case cst_ACTTYPE_BTN_SAVENEXT1
                '儲存後進下一步 

                '(檢查儲存值)
                Dim rPMS As New Hashtable From {{"CVOCID", Hid_CVOCID.Value}}
                Dim flag_OK_OBFL As Boolean = CHK_CLASS_VERIFYONLINE_FL(rPMS, Hid_KVSID.Value)
                If Not flag_OK_OBFL Then
                    Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步")
                    Return
                End If

                '下一步
                Call MOVE_NEXT()
        End Select
    End Sub

    ''' <summary>檢核</summary>
    ''' <param name="rPMS"></param>
    ''' <param name="vKVSID"></param>
    ''' <returns></returns>
    Private Function CHK_CLASS_VERIFYONLINE_FL(rPMS As Hashtable, vKVSID As String) As Boolean
        Dim vCVOCFID As String = TIMS.GetMyValue2(rPMS, "CVOCFID") '可為空
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID, vCVOCFID)
        Return (drFL IsNot Nothing)
    End Function

    ''' <summary>直接顯示項目</summary>
    ''' <param name="rPMS"></param>
    Private Sub SHOW_VERIFYONLINE_FL_DG2(rPMS As Hashtable)
        labmsg1.Text = ""
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim fg_CANSAVE As Boolean = (vCVOCID <> "" AndAlso TIMS.CINT1(vCVOCID) >= 0)
        'objconn 因為有檔案輸出關閉的問題 所以要檢查
        If Not TIMS.OpenDbConn(objconn) OrElse Not fg_CANSAVE Then Return

        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        'SAVE,,CLASS_VERIFYONLINE
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, Hid_OCIDVal.Value, Hid_SEQ_ID.Value, Hid_CVOCID.Value)
        If drCV Is Nothing OrElse vCVOCID <> Hid_CVOCID.Value Then
            Common.MessageBox(Me, "查詢資訊有誤(查無課程編號)，請重新操作!!")
            Return
        End If

        '查詢目前全部文件項目
        Dim dtFL As DataTable = TIMS.GET_CLASS_VERIFYONLINE_FL_TB(objconn, vCVOCID, Hid_OCIDVal.Value, Hid_SEQ_ID.Value)

        labmsg1.Text = If(TIMS.dtNODATA(dtFL), "(查無文件項目)", "")

        Dim vYEARS As String = Convert.ToString(drCV("YEARS"))
        Dim vPLANID As String = Convert.ToString(drCV("PLANID"))
        Dim vCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
        Dim vSEQNO As String = Convert.ToString(drCV("SEQNO"))
        Dim vORGKINDGW As String = Convert.ToString(drCV("ORGKINDGW"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")
        Call Check_dtVERIFYONLINE_FL(Me, dtFL, download_Path)
        DataGrid2.Columns(cst_DG2_退件原因_iCOLUMN).Visible = If(Convert.ToString(drCV("APPLIEDRESULT")) = "R", True, False)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

        'Dim iProgress As Integer=If(dtA.Rows.Count > 0, (dt.Rows.Count / dtA.Rows.Count * 100), 0)
        '線上申辦進度 計算完成度百分比 (0-100)
        Dim iProgress As Integer = GET_iPROGRESS_CVO(sm, objconn, tmpMSG, vCVOCID, vORGKINDGW)
        labProgress.Text = String.Concat(iProgress, "%")
        'BTN_SAVETMP1.Visible=(iProgress=100) 'BTN_SAVERC2.Visible=(iProgress=100)
        '儲存(暫存)
        BTN_SAVETMP1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, "", cst_tpmsg_enb1), True)
        '儲存後進下一步
        BTN_SAVENEXT1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, "", cst_tpmsg_enb1), True)
    End Sub

    ''' <summary>'線上申辦進度 計算完成度百分比 (0-100)</summary>
    ''' <param name="sm"></param>
    ''' <param name="oConn"></param>
    ''' <param name="showMsg"></param>
    ''' <param name="vCVOCID"></param>
    ''' <param name="vORGKINDGW"></param>
    ''' <returns></returns>
    Public Shared Function GET_iPROGRESS_CVO(sm As SessionModel, oConn As SqlConnection, ByRef showMsg As String, vCVOCID As String, vORGKINDGW As String) As Integer
        Dim iProgress As Integer = 0 'Return 0
        showMsg = ""

        Dim rParmsA As New Hashtable From {{"ORGKINDGW", vORGKINDGW}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim sSqlA As String = "SELECT 1 FROM KEY_VERIFY WHERE ORGKINDGW=@ORGKINDGW AND TPLANID=@TPLANID"
        Dim dtA As DataTable = DbAccess.GetDataTable(sSqlA, oConn, rParmsA)
        If dtA Is Nothing Then Return iProgress

        Dim rParms As New Hashtable From {{"CVOCID", TIMS.CINT1(vCVOCID)}, {"ORGKINDGW", vORGKINDGW}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim rsSql As String = ""
        rsSql &= " SELECT 1 FROM CLASS_VERIFYONLINE_FL a" & vbCrLf
        rsSql &= " JOIN KEY_VERIFY kv ON kv.KVSID=a.KVSID" & vbCrLf
        rsSql &= " JOIN CLASS_VERIFYONLINE cv on cv.CVOCID=a.CVOCID" & vbCrLf
        rsSql &= " WHERE a.CVOCID=@CVOCID AND kv.TPLANID=@TPLANID AND kv.ORGKINDGW=@ORGKINDGW" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        If dt Is Nothing Then Return iProgress

        iProgress = If(dtA.Rows.Count > 0, (dt.Rows.Count / dtA.Rows.Count * 100), 0)

        If iProgress > 30 AndAlso iProgress < 100 Then
            Dim rParms2 As New Hashtable From {{"CVOCID", TIMS.CINT1(vCVOCID)}, {"ORGKINDGW", vORGKINDGW}, {"TPLANID", sm.UserInfo.TPlanID}}
            Dim sSql2 As String = ""
            sSql2 &= " SELECT a.CVOCID,kv.KVID,kv.KBNAME,kv.ORGKINDGW" & vbCrLf
            sSql2 &= " ,concat(kv.KVID,'.',kv.KBNAME) KBNAME2" & vbCrLf
            sSql2 &= " ,CASE WHEN a.CVOCFID IS NULL THEN 'Y' END CVOCFIDNG" & vbCrLf
            sSql2 &= " FROM CLASS_VERIFYONLINE_FL a" & vbCrLf
            sSql2 &= " JOIN KEY_VERIFY kv ON kv.KVSID=a.KVSID" & vbCrLf
            sSql2 &= " JOIN CLASS_VERIFYONLINE cv on cv.CVOCID=a.CVOCID" & vbCrLf
            rsSql &= " WHERE a.CVOCID=@CVOCID AND kv.TPLANID=@TPLANID AND kv.ORGKINDGW=@ORGKINDGW" & vbCrLf
            sSql2 &= " ORDER BY kv.KSORT" & vbCrLf
            Dim dt2 As DataTable = DbAccess.GetDataTable(sSql2, oConn, rParms2)
            If TIMS.dtNODATA(dt2) Then Return iProgress
            For Each dr2 As DataRow In dt2.Rows
                If (Convert.ToString(dr2("CVOCFIDNG")) = "Y") Then
                    showMsg &= String.Concat(If(showMsg <> "", "、", ""), dr2("KBNAME2"))
                End If
            Next
        End If
        Return iProgress
    End Function

    ''' <summary>檢核實際檔案-"(檔案上傳失敗／異常，請刪除後重新上傳)"</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="dt"></param>
    ''' <param name="download_Path"></param>
    Private Sub Check_dtVERIFYONLINE_FL(MyPage As Page, dt As DataTable, download_Path As String)
        'Dim filename As String=""
        If TIMS.dtNODATA(dt) Then Return
        Const cst_errMsg_6 As String = "(檔案上傳失敗／異常，請刪除後重新上傳)"
        For i As Int16 = 0 To dt.Rows.Count - 1
            Dim filename As String = Convert.ToString(dt.Rows(i).Item("FILENAME1"))
            'Dim s_waived As String=Convert.ToString(dt.Rows(i).Item("WAIVED"))
            'Dim download_Path As String=String.Concat(G_UPDRV_JS, "/", dt.Rows(i).Item("PATH1"))
            If filename <> "" Then
                Dim flag_FL_EXISTS As Boolean = TIMS.CHK_PIC_EXISTS(MyPage.Server, download_Path, filename)
                dt.Rows(i)("OKFLAG") = If(Not flag_FL_EXISTS, cst_errMsg_6, filename)
                'Dim urlA1 As String=String.Concat("<a class='l' target='_blank' href=""", download_Path, filename, """>", filename, "</a>")
                'If Not flag_FL_EXISTS Then urlA1=String.Concat("<font color='red'>", cst_errMsg_6, "</font>") '表示 檔案不存在
                'dt.Rows(i)("OKFLAG")=urlA1
            End If
        Next
    End Sub

    ''' <summary>儲存／上傳-CLASS_VERIFYONLINE_FL</summary>
    ''' <param name="vCVOCID"></param>
    ''' <param name="vKVSID"></param>
    ''' <param name="rPMS2"></param>
    ''' <returns></returns>
    Private Function SAVE_CLASS_VERIFYONLINE_FL_UPLOAD(vCVOCID As String, vKVSID As String, rPMS2 As Hashtable) As Integer
        Dim iCVOCFID As Integer = 0
        'Dim vCVOCID As String=TIMS.GetMyValue2(rPMS2, "CVOCID")
        'Dim vKVSID As String=TIMS.GetMyValue2(rPMS2, "KVSID")
        vCVOCID = TIMS.ClearSQM(vCVOCID)
        vKVSID = TIMS.ClearSQM(vKVSID)
        If vCVOCID = "" OrElse vKVSID = "" Then Return iCVOCFID

        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS2, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS2, "SRCFILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS2, "UploadPath")
        Dim vMEMO1 As String = TIMS.GetMyValue2(rPMS2, "MEMO1")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS2, "MODIFYACCT")
        Dim vWAIVED As String = TIMS.GetMyValue2(rPMS2, "WAIVED")

        '免附文件或上傳檔案 無資料不儲存
        Dim fg_NG_SAVE As Boolean = (vWAIVED = "" AndAlso (vFILENAME1 = "" AndAlso vSRCFILENAME1 = ""))
        If fg_NG_SAVE Then Return iCVOCFID

        'Dim sSql_1 As String=" SELECT * FROM CLASS_VERIFYONLINE_FL WHERE CVOCID=@CVOCID AND KVSID=@KVSID" & vbCrLf
        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID)
        If drFL Is Nothing Then
            iCVOCFID = DbAccess.GetNewId(objconn, "CLASS_VERIFYONLINE_FL_CVOCFID_SEQ,CLASS_VERIFYONLINE_FL,CVOCFID")
            Dim isSql As String = ""
            isSql &= " INSERT INTO CLASS_VERIFYONLINE_FL(CVOCFID,CVOCID,KVSID,FILENAME1,SRCFILENAME1,FILEPATH1 ,MEMO1,MODIFYACCT,MODIFYDATE,WAIVED)" & vbCrLf
            isSql &= " VALUES (@CVOCFID,@CVOCID,@KVSID,@FILENAME1,@SRCFILENAME1,@FILEPATH1 ,@MEMO1,@MODIFYACCT,GETDATE(),@WAIVED)" & vbCrLf
            Dim iParms As New Hashtable
            iParms.Add("CVOCFID", iCVOCFID)
            iParms.Add("CVOCID", TIMS.CINT1(vCVOCID))
            iParms.Add("KVSID", TIMS.CINT1(vKVSID))
            iParms.Add("FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull))
            iParms.Add("SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull))
            iParms.Add("FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull)) 'iParms.Add("PATTERN", PATTERN)
            iParms.Add("MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull))
            iParms.Add("MODIFYACCT", If(vMODIFYACCT <> "", vMODIFYACCT, sm.UserInfo.UserID)) 'iParms.Add("MODIFYDATE", MODIFYDATE)
            iParms.Add("WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull))
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            iCVOCFID = drFL("CVOCFID")
            If iCVOCFID > 0 Then
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_VERIFYONLINE_FL" & vbCrLf
                usSql &= " SET FILENAME1=@FILENAME1,SRCFILENAME1=@SRCFILENAME1,FILEPATH1=@FILEPATH1" & vbCrLf ',PATTERN=@PATTERN
                usSql &= " ,MEMO1=@MEMO1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE(),WAIVED=@WAIVED" & vbCrLf ',RTUREASON=@RTUREASON,RTURESACCT=@RTURESACCT,RTURESDATE=@RTURESDATE
                usSql &= " WHERE CVOCFID=@CVOCFID AND CVOCID=@CVOCID AND KVSID=@KVSID" & vbCrLf
                Dim uParms As New Hashtable
                uParms.Add("FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull))
                uParms.Add("SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull))
                uParms.Add("FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull))
                uParms.Add("MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull))
                uParms.Add("MODIFYACCT", If(vMODIFYACCT <> "", vMODIFYACCT, sm.UserInfo.UserID))
                uParms.Add("WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull))
                uParms.Add("CVOCFID", iCVOCFID)
                uParms.Add("CVOCID", TIMS.CINT1(vCVOCID))
                uParms.Add("KVSID", TIMS.CINT1(vKVSID))
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
            End If
        End If

        '暫時儲存／正式儲存-UPDATE CLASS_VERIFYONLINE 
        Call SAVEDATE1(0)

        Return iCVOCFID
    End Function

    ''' <summary>再次儲存-CLASS_VERIFYONLINE</summary>
    ''' <param name="iNum"></param>
    Private Sub SAVEDATE1(ByVal iNum As Integer)
        'iNum:0 暫時儲存/1 正式儲存
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        Dim iCVOCID As Integer = If(Hid_CVOCID.Value <> "", TIMS.CINT1(Hid_CVOCID.Value), 0)
        If Hid_CVOCID.Value = "" OrElse iCVOCID <= 0 Then Return
        Dim iOCID As Integer = If(Hid_OCIDVal.Value <> "", TIMS.CINT1(Hid_OCIDVal.Value), 0)
        If Hid_OCIDVal.Value = "" OrElse iOCID <= 0 Then Return
        Dim iSEQ_ID As Integer = If(Hid_SEQ_ID.Value <> "", TIMS.CINT1(Hid_SEQ_ID.Value), 0)
        If Hid_SEQ_ID.Value = "" OrElse iSEQ_ID <= 0 Then Return

        Dim uParms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"CVOCID", iCVOCID},
            {"OCID", iOCID},
            {"SEQ_ID", iSEQ_ID}
        }
        Dim usSql As String = ""
        usSql &= " UPDATE CLASS_VERIFYONLINE" & vbCrLf
        usSql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE CVOCID=@CVOCID AND OCID=@OCID AND SEQ_ID=@SEQ_ID" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
    End Sub

    ''' <summary>
    ''' 後進下一步
    ''' </summary>
    Private Sub MOVE_NEXT()
        Hid_KVID.Value = TIMS.ClearSQM(Hid_KVID.Value)
        Hid_LastKVID.Value = TIMS.ClearSQM(Hid_LastKVID.Value)
        If Hid_KVID.Value <> "" AndAlso (Hid_KVID.Value = Hid_LastKVID.Value) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        ElseIf (ddlSwitchTo.SelectedIndex + 1 >= ddlSwitchTo.Items.Count) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        End If

        '下一步
        Hid_KVSID.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex + 1).Value
        Call SHOW_KEY_VERIFY_KVSID(Hid_KVSID.Value)
    End Sub

    ''' <summary>回上一步</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PREV1_Click(sender As Object, e As EventArgs) Handles BTN_PREV1.Click
        Call MOVE_PREV()
    End Sub

    ''' <summary>儲存(暫存)</summary>
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

    ''' <summary>不儲存返回查詢</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_BACK1_Click(sender As Object, e As EventArgs) Handles BTN_BACK1.Click
        '清理隱藏的參數
        Call ClearHidValue()

        Call SHOW_Frame1(0)
    End Sub

    ''' <summary>確定檔案上傳</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But1_Click(sender As Object, e As EventArgs) Handles But1.Click
        'Dim vUploadPath As String=Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        Hid_KVSID.Value = TIMS.ClearSQM(Hid_KVSID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If Hid_OCIDVal.Value = "" OrElse Hid_SEQ_ID.Value = "" OrElse Hid_CVOCID.Value = "" OrElse Hid_KVSID.Value = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(案件號為空)，請重新操作!")
            Return
        End If
        Dim vCVOCID As String = Hid_CVOCID.Value
        Dim vOCIDVal As String = Hid_OCIDVal.Value
        Dim vSEQ_ID As String = Hid_SEQ_ID.Value
        Dim vKVSID As String = Hid_KVSID.Value
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCIDVal, vSEQ_ID, vCVOCID)

        'Dim vKVSID As String=TIMS.ClearSQM(Hid_KVSID.Value)
        'Dim pms_sw As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"ORGKINDGW", Hid_ORGKINDGW.Value}}
        'Dim drKV As DataRow=GET_KEY_VERIFY(objconn, vKVSID, pms_sw)
        Dim vTPLANID As String = sm.UserInfo.TPlanID
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim drKV As DataRow = TIMS.GET_KEY_VERIFY(objconn, vKVSID, vTPLANID, vORGKINDGW)
        If drCV Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKV Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, Hid_CVOCID.Value, Hid_KVSID.Value)
        '(退件修正)有退件原因,可重新上傳
        'Dim flag_NG_UPLOAD_1 As Boolean=(drFL IsNot Nothing) '(有資料 不可再次傳送)
        Dim flag_NG_UPLOAD_2 As Boolean = (drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) = "") '(有資料不可傳送且原因為空 不可再次傳送)

        Dim vFILENAME1 As String = If(drFL IsNot Nothing, Convert.ToString(drFL("FILENAME1")), "")
        'Dim vWAIVED As String=If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
        If vFILENAME1 <> "" AndAlso flag_NG_UPLOAD_2 Then
            '符合所有 不可再次傳送 'cst_tpmsg_enb8
            Common.MessageBox(Me, "已上傳儲存過該文件，不可再次操作!")
            Return
        End If

        '有錯誤原因 可再次傳送 並記錄 iCVOCFID
        Dim iCVOCFID As Integer = If(drFL IsNot Nothing AndAlso Convert.ToString(drFL("RTUREASON")) <> "", TIMS.CINT1(drFL("CVOCFID")), -1)
        '檔案上傳／確定檔案上傳
        Call FILE_UPLOAD_1(drCV, drKV, iCVOCFID)

        '檢視目前上傳檔案 '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)
    End Sub

    ''' <summary>檔案上傳-資訊</summary>
    ''' <param name="drCV"></param>
    ''' <param name="drKV"></param>
    ''' <param name="iCVOCFID"></param>
    Private Sub FILE_UPLOAD_1(drCV As DataRow, drKV As DataRow, iCVOCFID As Integer)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        If drCV Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKV Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")

        Dim vCVOCID As String = TIMS.ClearSQM(drCV("CVOCID"))
        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vCVOCID <> Hid_CVOCID.Value Then Return '(此狀況不太可能發生)

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
        If Not TIMS.IsFileTypeValid(MyPostedFile, "pdf") Then
            Common.MessageBox(Me, cst_errMsg_5b)
            Exit Sub
        End If

        '取得KVID代號／非流水號
        Dim vKVSID As String = Convert.ToString(drKV("KVSID"))
        Dim vKVID As String = Convert.ToString(drKV("KVID"))

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vYEARS As String = Convert.ToString(drCV("YEARS"))
        Dim vPLANID As String = Convert.ToString(drCV("PLANID"))
        Dim vCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
        Dim vSEQNO As String = Convert.ToString(drCV("SEQNO"))
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_CVO(vCVOCID, vKVSID, "pdf")
        Dim vSRCFILENAME1 As String = MyFileName
        '上傳檔案/存檔：檔名
        Try
            '上傳檔案
            TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
            'File1.PostedFile.SaveAs(Server.MapPath(Cst_Upload_Path & MyFileName))
            'GUIDfilename=GetThumbNail(MyPostedFile.FileName, cst_pic_iWidth, cst_pic_iHeight, MyPostedFile.ContentType.ToString(), False, MyPostedFile.InputStream, Upload_Path)
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
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "CVOCFID", If(vUploadPath <> "", iCVOCFID, -1)) '(可再次傳送)

            'TIMS.SetMyValue2(rPMS2, "CVOCID", vCVOCID)
            'TIMS.SetMyValue2(rPMS2, "KVSID", vKVSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED) ' If(CHKB_WAIVED.Checked, "Y", ""))
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", vMODIFYACCT)

            '儲存／上傳-CLASS_VERIFYONLINE_FL
            Call SAVE_CLASS_VERIFYONLINE_FL_UPLOAD(vCVOCID, vKVSID, rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try

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
                TIMS.SetMyValue(sCmdArg, "CVOCID", drv("CVOCID"))
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                TIMS.SetMyValue(sCmdArg, "SEQ_ID", drv("SEQ_ID"))
                TIMS.SetMyValue(sCmdArg, "SENDSTATUS", drv("SENDSTATUS"))
                'TIMS.SetMyValue(sCmdArg, "SENDDATE", TIMS.cdate3(drv("SENDDATE")))

                Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                lBTN_DELETE1.Visible = If(flagS1, True, False)
                lBTN_DELETE1.Style.Item("display") = "none"
                lBTN_DELETE1.CommandArgument = sCmdArg
                lBTN_DELETE1.Attributes("onclick") = "javascript:return confirm('此動作會刪除審核資料，是否確定?');"

                lBTN_VIEW1.CommandArgument = sCmdArg
                lBTN_EDIT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.Attributes("onclick") = "javascript:return confirm('此動作會送出審核資料且不可再次修改，是否確定?');"

                'SENDSTATUS
                'lBTN_VIEW1.Enabled=If(Convert.ToString(drv("SENDSTATUS")) <> "", False, True)
                If Convert.ToString(drv("SENDSTATUS")) = "R" Then
                    lBTN_EDIT1.Enabled = True
                    TIMS.Tooltip(lBTN_EDIT1, cst_tpmsg_enb5, True)
                    lBTN_SENDOUT1.Enabled = True
                    TIMS.Tooltip(lBTN_SENDOUT1, cst_tpmsg_enb5, True)
                Else
                    lBTN_EDIT1.Enabled = If(Convert.ToString(drv("SENDSTATUS")) <> "", False, True)
                    TIMS.Tooltip(lBTN_EDIT1, If(lBTN_EDIT1.Enabled, "", cst_tpmsg_enb4), True)
                    lBTN_SENDOUT1.Enabled = If(Convert.ToString(drv("SENDSTATUS")) <> "", False, True)
                    TIMS.Tooltip(lBTN_SENDOUT1, If(lBTN_SENDOUT1.Enabled, "", cst_tpmsg_enb4), True)
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vCVOCID As String = TIMS.GetMyValue(sCmdArg, "CVOCID")
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim vSEQ_ID As String = TIMS.GetMyValue(sCmdArg, "SEQ_ID")
        Dim vSENDSTATUS As String = TIMS.GetMyValue(sCmdArg, "SENDSTATUS")
        If sCmdArg = "" OrElse vCVOCID = "" OrElse vOCID = "" OrElse vSEQ_ID = "" Then Return

        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCID, vSEQ_ID, vCVOCID) 'If drRR Is Nothing Then Return
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If vOCID = "" OrElse drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無課程代碼)，請選擇課程資料!!")
            Return
        ElseIf vCVOCID = "" OrElse drCV Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無核銷代碼)，請選擇核銷資料!!")
            Return
        End If

        Dim vORGKINDGW As String = Convert.ToString(drCV("ORGKINDGW"))
        Dim s_RESULTDATE_YMS2 As String = If(Convert.ToString(drCV("RESULTDATE")) <> "", CDate(drCV("RESULTDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
        Dim s_MODIFYDATE_YMS2 As String = If(Convert.ToString(drCV("MODIFYDATE")) <> "", CDate(drCV("MODIFYDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
        Dim fg_RESULTDATE_UPDATE As Boolean = (s_RESULTDATE_YMS2 <> "" AndAlso s_MODIFYDATE_YMS2 <> "" AndAlso DateDiff(DateInterval.Second, CDate(s_RESULTDATE_YMS2), CDate(s_MODIFYDATE_YMS2)) > 0)

        Select Case e.CommandName
            Case cst_DG1CMDNM_DELETE1 'DELETE1 (刪除)
                Call DELETE_Detail_CLASS_VERIFYONLINE(Me, objconn, drCV, vCVOCID)
                Common.MessageBox(Me, TIMS.cst_DELETEOKMsg2)
                Call SSearch1()

            Case cst_DG1CMDNM_VIEW1 '"VIEW1 '查看
                '依目前核銷代號查詢詳細資料
                Call SHOW_Detail_VERIFYONLINE(drCC, vOCID, vSEQ_ID, cst_DG1CMDNM_VIEW1)

            Case cst_DG1CMDNM_EDIT1 '"EDIT1 '修改
                '依目前核銷代號查詢詳細資料
                Call SHOW_Detail_VERIFYONLINE(drCC, vOCID, vSEQ_ID, cst_DG1CMDNM_EDIT1)

            Case cst_DG1CMDNM_SENDOUT1 'SENDOUT1 送出 
                '線上申辦進度 計算完成度百分比 (0-100)
                Dim iProgress As Integer = GET_iPROGRESS_CVO(sm, objconn, tmpMSG, vCVOCID, vORGKINDGW)
                Dim EMSG As String = ""
                If iProgress < 100 Then
                    EMSG = String.Concat("線上核銷送件進度 未達100%，不可送出! (", iProgress, "%)", vbCrLf, If(tmpMSG <> "", String.Concat("請檢查：(", tmpMSG, ")"), ""))
                    Common.MessageBox(Me, EMSG)
                    Return
                ElseIf Convert.ToString(drCV("SENDSTATUS")) = "R" AndAlso Not fg_RESULTDATE_UPDATE Then
                    EMSG = cst_tpmsg_enb5
                    Common.MessageBox(Me, EMSG)
                    Return
                End If

                ',CASE a.SENDSTATUS WHEN 'B' THEN '已送件' WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過'" & vbCrLf
                'uParms.Add("SENDDATE", SENDDATE)
                Dim uParms As New Hashtable From {
                    {"SENDACCT", sm.UserInfo.UserID},
                    {"SENDSTATUS", "B"},
                    {"MODIFYACCT", sm.UserInfo.UserID},
                    {"CVOCID", vCVOCID},
                    {"OCID", vOCID},
                    {"SEQ_ID", vSEQ_ID}
                }
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_VERIFYONLINE" & vbCrLf
                usSql &= " SET SENDACCT=@SENDACCT,SENDDATE=GETDATE(),SENDSTATUS=@SENDSTATUS" & vbCrLf
                usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE CVOCID=@CVOCID AND OCID=@OCID AND SEQ_ID=@SEQ_ID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

                Call SSearch1()
        End Select
    End Sub

    ''' <summary>刪除 核銷資料 CLASS_VERIFYONLINE／CLASS_VERIFYONLINE_FL</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="oConn"></param>
    ''' <param name="drCV"></param>
    ''' <param name="vCVOCID"></param>
    Public Shared Sub DELETE_Detail_CLASS_VERIFYONLINE(MyPage As Page, oConn As SqlConnection, drCV As DataRow, vCVOCID As String)
        vCVOCID = TIMS.ClearSQM(vCVOCID)
        If vCVOCID = "" Then Return

        Dim vYEARS As String = Convert.ToString(drCV("YEARS"))
        Dim vPLANID As String = Convert.ToString(drCV("PLANID"))
        Dim vCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
        Dim vSEQNO As String = Convert.ToString(drCV("SEQNO"))
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")

        '核銷資料 上傳檔案 刪除檔案
        Dim dtFL As DataTable = TIMS.GET_CLASS_VERIFYONLINE_FL(oConn, vCVOCID)
        For Each drFL As DataRow In dtFL.Rows
            Dim oFILENAME1 As String = Convert.ToString(drFL("FILENAME1"))
            If oFILENAME1 = "" Then Continue For
            Dim oUploadPath As String = ""
            Dim s_FilePath1 As String = ""
            Try
                oUploadPath = vUploadPath
                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Call TIMS.MyFileDelete(s_FilePath1)
            Catch ex As Exception
                Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
            End Try
        Next

        Dim dParms As New Hashtable From {{"CVOCID", vCVOCID}}
        Dim dsSql As String = ""
        dsSql = "DELETE CLASS_VERIFYONLINE WHERE CVOCID=@CVOCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)

        Dim dParms2 As New Hashtable From {{"CVOCID", vCVOCID}}
        Dim dsSql2 As String = ""
        dsSql2 = "DELETE CLASS_VERIFYONLINE_FL WHERE CVOCID=@CVOCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql2, oConn, dParms2)
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim LabdepID As Label=e.Item.FindControl("LabdepID")
                'Dim LabFileName1 As Label=e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden=e.Item.FindControl("HFileName")
                Dim BTN_DELFILE4 As Button = e.Item.FindControl("BTN_DELFILE4") '刪除
                Dim BTN_DOWNLOAD4 As Button = e.Item.FindControl("BTN_DOWNLOAD4") '下載 
                Dim labRTUREASON As Label = e.Item.FindControl("labRTUREASON") '退件原因
                labRTUREASON.Text = Convert.ToString(drv("RTUREASON"))

                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    'LabFileName1.Text=If(Convert.ToString(drv("FILENAME1"))=Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                    'HFileName.Value=Convert.ToString(drv("FILENAME1")) '.ToString()
                    titleMsg = Convert.ToString(drv("OKFLAG"))
                    BTN_DOWNLOAD4.Enabled = (Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")))
                ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "CVOCFID", Convert.ToString(drv("CVOCFID")))
                TIMS.SetMyValue(sCmdArg, "KVID", Convert.ToString(drv("KVID")))
                TIMS.SetMyValue(sCmdArg, "KVSID", Convert.ToString(drv("KVSID")))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", Convert.ToString(drv("FILENAME1")))
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", Convert.ToString(drv("FILEPATH1")))

                BTN_DELFILE4.CommandArgument = sCmdArg '刪除
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                BTN_DELFILE4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE4.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)

                '(其他原因調整) '送件／退件修正，不提供刪除
                If Convert.ToString(drv("SENDSTATUS")) = "B" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb6, True)

                ElseIf Convert.ToString(drv("SENDSTATUS")) = "R" AndAlso Convert.ToString(drv("RTUREASON")) <> "" Then
                    BTN_DELFILE4.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb8, True)

                ElseIf Convert.ToString(drv("SENDSTATUS")) = "R" AndAlso Convert.ToString(drv("RTUREASON")) = "" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb7, True)

                End If
        End Select
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim HFileName As HtmlInputHidden=e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vCVOCFID As String = TIMS.ClearSQM(TIMS.GetMyValue(sCmdArg, "CVOCFID"))
        Dim vKVID As String = TIMS.GetMyValue(sCmdArg, "KVID")
        Dim vKVSID As String = TIMS.GetMyValue(sCmdArg, "KVSID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Hid_CVOCFID.Value = TIMS.ClearSQM(vCVOCFID)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        If e.CommandArgument = "" OrElse vCVOCFID = "" Then Return

        'Dim vCVOCFID As String=TIMS.ClearSQM(Hid_CVOCFID.Value)
        Dim vCVOCID As String = TIMS.ClearSQM(Hid_CVOCID.Value)
        Dim vOCIDVal As String = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Dim vSEQ_ID As String = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        Select Case e.CommandName
            Case "DELFILE4"
                Dim sErrMsg1 As String = CHKDEL_CLASS_VERIFYONLINE_FL(vCVOCFID)
                If sErrMsg1 <> "" Then
                    Common.MessageBox(Me, sErrMsg1)
                    Return
                End If

                Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCIDVal, vSEQ_ID, vCVOCID)
                If drCV Is Nothing Then Return
                If Hid_CVOCID.Value = "" Then Hid_CVOCID.Value = drCV("CVOCID")
                Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID, vCVOCFID)
                If drFL Is Nothing Then Return

                '刪除檔案 '"_CLASS_VERIFYONLINE_FL"
                Dim oYEARS As String = Convert.ToString(drCV("YEARS"))
                Dim oPLANID As String = Convert.ToString(drCV("PLANID"))
                Dim oCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
                Dim oSEQNO As String = Convert.ToString(drCV("SEQNO"))
                Dim oCVOCID As String = Convert.ToString(drCV("CVOCID"))
                'Dim vUploadPath As String=TIMS.GET_UPLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")
                Dim oFILENAME1 As String = ""
                Dim oUploadPath As String = ""
                Dim s_FilePath1 As String = ""
                Try
                    oFILENAME1 = vFILENAME1
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH1_CVO(oYEARS, oPLANID, oCOMIDNO, oSEQNO, oCVOCID, ""))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try

                '"CLASS_VERIFYONLINE_FL"
                Dim dParms As New Hashtable From {{"CVOCFID", vCVOCFID}}
                Dim rdSql As String = "DELETE CLASS_VERIFYONLINE_FL WHERE CVOCFID=@CVOCFID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)
                'DataGrid1.EditItemIndex=-1

            Case "DOWNLOAD4" '下載
                'Hid_ORGKINDGW.Value=TIMS.ClearSQM(Hid_ORGKINDGW.Value)
                'Hid_BCID.Value=TIMS.ClearSQM(Hid_BCID.Value)
                'Hid_BCASENO.Value=TIMS.ClearSQM(Hid_BCASENO.Value)
                'RIDValue.Value=TIMS.ClearSQM(RIDValue.Value)
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "CVOCFID", vCVOCFID)
                TIMS.SetMyValue2(rPMS4, "CVOCID", vCVOCID)
                TIMS.SetMyValue2(rPMS4, "OCID", vOCIDVal)
                TIMS.SetMyValue2(rPMS4, "SEQ_ID", vSEQ_ID)
                TIMS.SetMyValue2(rPMS4, "KVSID", vKVSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)

                Call TIMS.ResponseZIPFile_CVO(sm, objconn, Me, rPMS4)
                Return
        End Select

        If Not TIMS.OpenDbConn(objconn) Then Return

        '檢視目前上傳檔案 '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)

        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCIDVal, objconn)
        If vOCIDVal = "" OrElse drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無課程代碼)，請選擇課程資料!!")
            Return
        End If

        '依目前核銷代號查詢詳細資料
        Call SHOW_Detail_VERIFYONLINE(drCC, Hid_OCIDVal.Value, Hid_SEQ_ID.Value, "")
    End Sub

    ''' <summary>查核銷代碼存在與否</summary>
    ''' <param name="vCVOCFID"></param>
    ''' <returns></returns>
    Private Function CHKDEL_CLASS_VERIFYONLINE_FL(vCVOCFID As String) As String
        Dim rst As String = ""
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        Hid_KVSID.Value = TIMS.ClearSQM(Hid_KVSID.Value)
        Hid_KVID.Value = TIMS.ClearSQM(Hid_KVID.Value)

        Dim vCVOCID As String = Hid_CVOCID.Value
        Dim vOCIDVal As String = Hid_OCIDVal.Value
        Dim vSEQ_ID As String = Hid_SEQ_ID.Value
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCIDVal, vSEQ_ID, vCVOCID)
        If drCV Is Nothing Then
            rst &= "資訊有誤(查無核銷代碼)，請重新操作!!"
            Return rst
        End If

        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, Hid_CVOCID.Value, Hid_KVSID.Value, vCVOCFID)

        Dim sParms1 As New Hashtable From {{"CVOCFID", vCVOCFID}}
        Dim sSql As String = ""
        sSql &= " SELECT a.CVOCFID,a.CVOCID" & vbCrLf
        sSql &= " ,a.KVSID,a.RTUREASON,a.RTURESACCT,a.RTURESDATE" & vbCrLf
        sSql &= " ,kv.ORGKINDGW,kv.KVID,kv.KBNAME" & vbCrLf
        sSql &= " FROM CLASS_VERIFYONLINE_FL a" & vbCrLf
        sSql &= " JOIN KEY_VERIFY kv ON kv.KVSID=a.KVSID" & vbCrLf
        sSql &= " WHERE a.CVOCFID=@CVOCFID" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sSql, objconn, sParms1)
        If dr1 Is Nothing Then Return rst

        Return rst
    End Function

    ''' <summary>下載報表</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_DOWNLOADRPT1_Click(sender As Object, e As EventArgs) Handles BTN_DOWNLOADRPT1.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_KVSID.Value = TIMS.ClearSQM(Hid_KVSID.Value)
        Hid_KVID.Value = TIMS.ClearSQM(Hid_KVID.Value)

        Hid_CVOCFID.Value = TIMS.ClearSQM(Hid_CVOCFID.Value)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        If Hid_CVOCID.Value = "" OrElse Hid_CVOCID.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無課程編號)，請重新操作!")
            Return
        ElseIf Hid_KVSID.Value = "" OrElse Hid_KVID.Value = "" OrElse Hid_ORGKINDGW.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目代碼為空)，請重新操作!!")
            Return
        End If
        Dim vKVSID As String = TIMS.ClearSQM(Hid_KVSID.Value)
        Dim vTPLANID As String = sm.UserInfo.TPlanID
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'SAVE,,CLASS_VERIFYONLINE
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, Hid_OCIDVal.Value, Hid_SEQ_ID.Value, Hid_CVOCID.Value)
        If drCV Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無課程核銷編號)，請重新操作!")
            Return
        End If
        Dim drKV As DataRow = TIMS.GET_KEY_VERIFY(objconn, vKVSID, vTPLANID, vORGKINDGW)
        If drKV Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!")
            Return
        End If

        '首頁>>訓練機構管理>>表單列印>>訓練單位基本資料表 '訓練單位基本資料
        'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_001_18G&Years=112&RSID=47877&planid=5093&rid=%27E6762%27&AppStage=1&UserID=snoopy
        '列印-REPORT-PDF
        Call UTL_PRINT1GW(drCV, drKV)
    End Sub

    ''' <summary>'列印-REPORT-PDF</summary>
    ''' <param name="drCV"></param>
    ''' <param name="drKV"></param>
    Sub UTL_PRINT1GW(ByRef drCV As DataRow, ByRef drKV As DataRow)
        If drCV Is Nothing OrElse drKV Is Nothing Then Return

        Dim vOCIDVal As String = Convert.ToString(drCV("OCID"))
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCIDVal, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
            Return
        End If

        '取得KVID代號／非流水號
        Dim vKVID As String = Convert.ToString(drKV("KVID"))
        Dim vORGKINDGW As String = Convert.ToString(drKV("ORGKINDGW"))
        Dim vKBNAME2 As String = String.Concat(vORGKINDGW, vKVID, drKV("KBNAME"))
        Dim rPMS As New Hashtable
        Select Case String.Concat(vORGKINDGW, vKVID)
            Case cst_G01_學員補助申請書, cst_W01_學員補助申請書
                'Dim vOCIDVal As String=Convert.ToString(drCV("OCID"))
                'Dim drCC As DataRow=TIMS.GetOCIDDate(vOCIDVal, objconn)
                'If OCIDValue1.Value="" OrElse drCC Is Nothing Then
                '    Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
                '    Return
                'End If

                '學員資料審核鈕
                If Convert.ToString(drCC("AppliedResultR")) <> "Y" Then
                    Common.MessageBox(Me, String.Concat("學員資料尚未審核!! (", vOCIDVal, ")"))
                    Exit Sub
                End If

                TIMS.SetMyValue2(rPMS, "TPLANID", drCC("TPLANID"))
                TIMS.SetMyValue2(rPMS, "OCID", drCC("OCID"))
                TIMS.SetMyValue2(rPMS, "ORGKINDGW", drCC("ORGKINDGW"))
                TIMS.SetMyValue2(rPMS, "YEARS_ROC", drCC("YEARS_ROC"))
                TIMS.SetMyValue2(rPMS, "RID", drCC("RID"))
                '補助經費申請書
                'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_013_2018G&RID=D4871&OCID=156430&SOCID=2897866,2897867&Years=113&UserID=snoopy
                Call RPT_SD_14_013(rPMS)

            Case cst_G02_支付參訓學員補助經費申請表, cst_W02_支付參訓學員補助經費申請表
                'Dim vOCIDVal As String=Convert.ToString(drCV("OCID"))
                'Dim drCC As DataRow=TIMS.GetOCIDDate(vOCIDVal, objconn)
                'If OCIDValue1.Value="" OrElse drCC Is Nothing Then
                '    Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
                '    Return
                'End If

                TIMS.SetMyValue2(rPMS, "OCID", drCC("OCID"))
                TIMS.SetMyValue2(rPMS, "YEARS_ROC", drCC("YEARS_ROC"))
                TIMS.SetMyValue2(rPMS, "RID", drCC("RID"))
                '支付參訓學員補助經費申請表
                'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_012_1_2021&Years=113&OCID=157055&RID=B7860&Printtype=2&Printtype2=Y&UserID=snoopy
                Call RPT_SD_14_012(rPMS)

            Case cst_G03_參訓學員出席紀錄一覽表, cst_W03_參訓學員出席紀錄一覽表
                'Dim vOCIDVal As String=Convert.ToString(drCV("OCID"))
                'Dim drCC As DataRow=TIMS.GetOCIDDate(vOCIDVal, objconn)
                'If OCIDValue1.Value="" OrElse drCC Is Nothing Then
                '    Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
                '    Return
                'End If

                TIMS.SetMyValue2(rPMS, "OCID", drCC("OCID"))
                TIMS.SetMyValue2(rPMS, "YEARS_ROC", drCC("YEARS_ROC"))
                '參訓學員出席紀錄一覽表
                'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_009_16&Years=113&OCID=157060&UserID=snoopy
                Call RPT_SD_14_009(rPMS)

            Case cst_G04_結訓證書清冊, cst_W04_結訓證書清冊
                'Dim vOCIDVal As String=Convert.ToString(drCV("OCID"))
                'Dim drCC As DataRow=TIMS.GetOCIDDate(vOCIDVal, objconn)
                'If OCIDValue1.Value="" OrElse drCC Is Nothing Then
                '    Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
                '    Return
                'End If

                TIMS.SetMyValue2(rPMS, "TPlanID", drCC("TPlanID"))
                TIMS.SetMyValue2(rPMS, "OCID", drCC("OCID"))
                TIMS.SetMyValue2(rPMS, "MSD", drCC("MSD"))
                '結訓證書清冊
                'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_030&TPlanID=28&MSD=232003&OCID=157060&UserID=snoopy
                Call RPT_SD_14_030(rPMS)

            Case cst_G08_學員線上簽到退明細一覽表, cst_W08_學員線上簽到退明細一覽表
                Dim rPMShp As New Hashtable From {{"HP", 2}, {"OCID", vOCIDVal}}

                Call TIMS.UPDATE_ADP_STUDATTEND(sm, objconn, vOCIDVal, rPMShp)

                Dim sErrMsg3 As String = ""
                Dim iStd As Integer = SD_14_029.CHK_STUD_ADP_STUDATTEND(objconn, vOCIDVal)
                If iStd = 0 Then sErrMsg3 &= "查無 學員線上簽到(退)明細 資料!" & vbCrLf
                If sErrMsg3 <> "" Then
                    Common.MessageBox(Me, sErrMsg3) 'labmsg.Text=sErrMsg3
                    Exit Sub
                End If

                TIMS.SetMyValue2(rPMS, "TPlanID", drCC("TPlanID"))
                TIMS.SetMyValue2(rPMS, "OCID", drCC("OCID"))
                TIMS.SetMyValue2(rPMS, "MSD", drCC("MSD"))
                '學員線上簽到退明細一覽表
                Call RPT_SD_14_029(rPMS)

        End Select

    End Sub

    Private Sub RPT_SD_14_029(rPMS As Hashtable)
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim vOCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vMSD As String = TIMS.GetMyValue2(rPMS, "MSD")

        Const cst_printFN1 As String = "SD_14_029" '2011
        'Dim s_MSD As String=Convert.ToString(drCC("MSD"))
        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "TPlanID", vTPlanID)
        TIMS.SetMyValue(MyValue, "OCID", vOCIDVal)
        TIMS.SetMyValue(MyValue, "MSD", vMSD)

        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    ''' <summary>結訓證書清冊</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_030(rPMS As Hashtable)
        'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_030&TPlanID=28&MSD=232003&OCID=157060&UserID=snoopy
        Dim v_TPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim v_OCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim s_MSD As String = TIMS.GetMyValue2(rPMS, "MSD")

        Const cst_printFN1 As String = "SD_14_030" '2011

        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "TPlanID", v_TPlanID)
        TIMS.SetMyValue(MyValue, "MSD", s_MSD)
        TIMS.SetMyValue(MyValue, "OCID", v_OCIDVal)
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    ''' <summary>參訓學員出席紀錄一覽表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_009(rPMS As Hashtable)
        'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_009_16&Years=113&OCID=157060&UserID=snoopy
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS, "YEARS_ROC")
        Dim v_OCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Const cst_printFN3 As String = "SD_14_009_16" '2016
        Dim MyValue As String = ""
        MyValue = "Years=" & vYEARS_ROC
        MyValue &= "&OCID=" & v_OCIDVal
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN3, MyValue)
    End Sub

    ''' <summary>支付參訓學員補助經費申請表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_012(rPMS As Hashtable)
        'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_012_1_2021&Years=113&OCID=157055&RID=B7860&Printtype=2&Printtype2=Y&UserID=snoopy
        'Dim vTPLANID As String=TIMS.GetMyValue2(rPMS, "TPLANID")
        Dim v_OCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        'Dim vORGKINDGW As String=TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS, "YEARS_ROC")
        Dim v_RIDValue As String = TIMS.GetMyValue2(rPMS, "RID")

        Const cst_printFN1 As String = "SD_14_012_1_2021"
        Dim MyValue1 As String = ""
        TIMS.SetMyValue(MyValue1, "Years", vYEARS_ROC)
        TIMS.SetMyValue(MyValue1, "OCID", v_OCIDVal)
        TIMS.SetMyValue(MyValue1, "RID", v_RIDValue)
        TIMS.SetMyValue(MyValue1, "Printtype", "2")
        TIMS.SetMyValue(MyValue1, "Printtype2", "Y")
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)
    End Sub

    ''' <summary>補助經費申請書-SD_14_013_2018</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_013(rPMS As Hashtable)
        '補助經費申請書
        'http://192.168.0.76:8080/ReportServer3/report?RptID=SD_14_013_2018G&RID=D4871&OCID=156430&SOCID=2897866,2897867&Years=113&UserID=snoopy
        Dim vTPLANID As String = TIMS.GetMyValue2(rPMS, "TPLANID")
        Dim v_OCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS, "YEARS_ROC")
        Dim v_RIDValue As String = TIMS.GetMyValue2(rPMS, "RID")

        Dim pms_1 As New Hashtable From {{"TPLANID", vTPLANID}, {"OCID", v_OCIDVal}}
        Dim sSql As String = ""
        sSql &= " SELECT a.SOCID" & vbCrLf
        sSql &= " FROM dbo.V_STUDENTINFO a" & vbCrLf
        sSql &= " JOIN dbo.STUD_SUBSIDYCOST sc ON sc.SOCID=a.SOCID" & vbCrLf
        sSql &= " WHERE a.STUDSTATUS NOT IN (2,3)" & vbCrLf
        sSql &= " AND a.TPLANID=@TPLANID AND a.OCID=@OCID" & vbCrLf
        Dim dtSS As DataTable = DbAccess.GetDataTable(sSql, objconn, pms_1)
        If TIMS.dtNODATA(dtSS) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return
        End If

        Dim s_SOCIDVals As String = ""
        For Each drS1 As DataRow In dtSS.Rows
            Dim SOCIDVal As String = Convert.ToString(drS1("SOCID"))
            s_SOCIDVals &= String.Concat(If(s_SOCIDVals <> "", ",", ""), SOCIDVal)
        Next

        Const cst_printFNoth2 As String = "SD_14_013_2019" '非產投計畫使用
        Const cst_printFN18G As String = "SD_14_013_2018G"
        Const cst_printFN18W As String = "SD_14_013_2018W"

        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "RID", v_RIDValue)
        TIMS.SetMyValue(myValue, "OCID", v_OCIDVal)
        TIMS.SetMyValue(myValue, "SOCID", s_SOCIDVals)
        TIMS.SetMyValue(myValue, "Years", vYEARS_ROC)
        Dim filename As String = cst_printFNoth2 '非產投計畫使用
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case vORGKINDGW
                Case "G"
                    filename = cst_printFN18G
                Case "W"
                    filename = cst_printFN18W
            End Select
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, filename, myValue)
    End Sub

    ''' <summary>以目前版本送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_SENDCURRVER_Click(sender As Object, e As EventArgs) Handles BTN_SENDCURRVER.Click
        'Dim vUploadPath As String=Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)

        Hid_KVSID.Value = TIMS.ClearSQM(Hid_KVSID.Value)
        Hid_KVID.Value = TIMS.ClearSQM(Hid_KVID.Value)
        Hid_LastKVID.Value = TIMS.ClearSQM(Hid_LastKVID.Value)
        Hid_FirstKVSID.Value = TIMS.ClearSQM(Hid_FirstKVSID.Value)

        Hid_CVOCFID.Value = TIMS.ClearSQM(Hid_CVOCFID.Value)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        If Hid_CVOCID.Value = "" OrElse Hid_OCIDVal.Value = "" OrElse Hid_SEQ_ID.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(核銷案件號為空)，請重新操作!!")
            Return
        End If

        Dim vCVOCID As String = TIMS.ClearSQM(Hid_CVOCID.Value)
        Dim vOCIDVal As String = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Dim vSEQ_ID As String = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        Dim vKVSID As String = TIMS.ClearSQM(Hid_KVSID.Value)
        Dim vTPLANID As String = sm.UserInfo.TPlanID
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)

        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCIDVal, vSEQ_ID, vCVOCID)
        'Dim drFL As DataRow=TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, Hid_CVOCID.Value, Hid_KVSID.Value)
        Dim drKV As DataRow = TIMS.GET_KEY_VERIFY(objconn, vKVSID, vTPLANID, vORGKINDGW)
        If drCV Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無核銷案件編號)，請重新操作!!")
            Return
        ElseIf drKV Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無核銷項目編號)，請重新操作!!")
            Return
        End If
        'Dim vOCIDVal As String=Convert.ToString(drCV("OCID"))
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCIDVal, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇職類/班別!")
            Return
        End If

        Dim vKVID As String = Convert.ToString(drKV("KVID"))
        'Dim vORGKINDGW As String=Convert.ToString(drKV("ORGKINDGW"))
        Select Case String.Concat(vORGKINDGW, vKVID)
            Case cst_G03_參訓學員出席紀錄一覽表, cst_W03_參訓學員出席紀錄一覽表
                Dim rPMS As New Hashtable
                rPMS.Add("C_KBNAME", drKV("C_KBNAME"))
                rPMS.Add("C_KBNAME2", drKV("C_KBNAME2"))
                rPMS.Add("ORGKINDGW", drKV("ORGKINDGW"))
                rPMS.Add("CVOCID", TIMS.CINT1(vCVOCID))
                rPMS.Add("OCID", vOCIDVal)
                rPMS.Add("SEQ_ID", vSEQ_ID)
                rPMS.Add("KVSID", TIMS.CINT1(vKVSID))
                'rPMS.Add("CVOCFID", iCVOCFID)
                rPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
                rPMS.Add("YEARS", drCV("YEARS"))
                rPMS.Add("PLANID", drCC("PLANID"))
                rPMS.Add("COMIDNO", drCC("COMIDNO"))
                rPMS.Add("SEQNO", drCC("SEQNO"))
                rPMS.Add("MSD", drCC("MSD"))
                '參訓學員出席紀錄一覽表
                Call SAVE_CLASS_VERIFYONLINE_ALL_GW03(rPMS)

            Case cst_G04_結訓證書清冊, cst_W04_結訓證書清冊
                Dim rPMS As New Hashtable
                rPMS.Add("C_KBNAME", drKV("C_KBNAME"))
                rPMS.Add("C_KBNAME2", drKV("C_KBNAME2"))
                rPMS.Add("ORGKINDGW", drKV("ORGKINDGW"))
                rPMS.Add("CVOCID", TIMS.CINT1(vCVOCID))
                rPMS.Add("OCID", vOCIDVal)
                rPMS.Add("SEQ_ID", vSEQ_ID)
                rPMS.Add("KVSID", TIMS.CINT1(vKVSID))
                'rPMS.Add("CVOCFID", iCVOCFID)
                rPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
                rPMS.Add("YEARS", drCV("YEARS"))
                rPMS.Add("PLANID", drCC("PLANID"))
                rPMS.Add("COMIDNO", drCC("COMIDNO"))
                rPMS.Add("SEQNO", drCC("SEQNO"))
                rPMS.Add("MSD", drCC("MSD"))
                '結訓證書清冊
                Call SAVE_CLASS_VERIFYONLINE_ALL_GW04(rPMS)

            Case cst_G08_學員線上簽到退明細一覽表, cst_W08_學員線上簽到退明細一覽表
                '--view-source:https://localhost:44383/SD/14/SD_14_029?ID=945
                'Const cst_printFN1 As String="SD_14_029" '2011
                'OCIDValue1.Value=TIMS.ClearSQM(OCIDValue1.Value)
                Dim rPMS As New Hashtable
                rPMS.Add("C_KBNAME", drKV("C_KBNAME"))
                rPMS.Add("C_KBNAME2", drKV("C_KBNAME2"))
                rPMS.Add("ORGKINDGW", drKV("ORGKINDGW"))
                rPMS.Add("CVOCID", TIMS.CINT1(vCVOCID))
                rPMS.Add("OCID", vOCIDVal)
                rPMS.Add("SEQ_ID", vSEQ_ID)
                rPMS.Add("KVSID", TIMS.CINT1(vKVSID))
                'rPMS.Add("CVOCFID", iCVOCFID)
                rPMS.Add("MODIFYACCT", sm.UserInfo.UserID)
                rPMS.Add("YEARS", drCV("YEARS"))
                rPMS.Add("PLANID", drCC("PLANID"))
                rPMS.Add("COMIDNO", drCC("COMIDNO"))
                rPMS.Add("SEQNO", drCC("SEQNO"))
                rPMS.Add("MSD", drCC("MSD"))
                '學員線上簽到退明細一覽表
                Call SAVE_CLASS_VERIFYONLINE_ALL_GW08(rPMS)

            Case Else
                Dim rParms2 As New Hashtable From {{"CVOCID", TIMS.CINT1(Hid_CVOCID.Value)}, {"KVSID", TIMS.CINT1(Hid_KVSID.Value)}}
                Dim rSql2 As String = "SELECT 1 FROM CLASS_VERIFYONLINE_FL WHERE CVOCID=@CVOCID AND KVSID=@KVSID"
                Dim drFL2 As DataRow = DbAccess.GetOneRow(rSql2, objconn, rParms2)
                If drFL2 IsNot Nothing Then
                    Common.MessageBox(Me, "已儲存過該文件，不可再次操作!!")
                    Return
                End If
                Common.MessageBox(Me, "此按鈕無功能(請連絡系統管理者)，不可再次操作!")
                Return
        End Select

        Threading.Thread.Sleep(10)
        '檢視目前上傳檔案 '顯示上傳檔案／細項 -'線上申辦進度
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)

    End Sub

    ''' <summary>結訓證書清冊</summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_CLASS_VERIFYONLINE_ALL_GW04(rPMS As Hashtable)
        Dim vC_KBNAME As String = TIMS.GetMyValue2(rPMS, "C_KBNAME")
        Dim vC_KBNAME2 As String = TIMS.GetMyValue2(rPMS, "C_KBNAME2")
        Dim vOCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS, "YEARS_ROC")
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim vKVSID As String = TIMS.GetMyValue2(rPMS, "KVSID")

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(rPMS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(rPMS, "SEQNO")
        Dim vMSD As String = TIMS.GetMyValue2(rPMS, "MSD")

        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID)
        Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
        If drFL IsNot Nothing Then
            Dim vMODIFY_DAY As String = Convert.ToString(drFL("MODIFY_DAY")) 'MODIFY_DAY
            Dim vMODIFY_MI As String = Convert.ToString(drFL("MODIFY_MI")) 'MODIFY_MI
            fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
        End If

        If Not fg_RUN_REPORT_1 Then
            's_TMPMSG1 &= String.Concat(If(s_TMPMSG1 <> "", ", ", ""), iRow)
            Common.MessageBox(Me, String.Concat("項目：", vC_KBNAME, ",重複處理時間過短(3分鐘1次)，請等待3分鐘後再試!", vbCrLf))
            Return
        End If

        'rPMS4.Clear()
        Dim rPMS4 As New Hashtable From {{"TPlanID", sm.UserInfo.TPlanID}, {"OCID", vOCIDVal}, {"MSD", vMSD}}
        Dim s_RPTURL As String = GET_RPTURL_SD_14_030(rPMS4)
        Dim s_PDF_byte As Byte() = Nothing
        Try
            Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
        Catch ex As Exception
            Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
            eErrmsg &= String.Concat(", s_RPTURL: ", s_RPTURL)
            eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
            eErrmsg &= String.Concat(", rPMS4: ", TIMS.GetMyValue4(rPMS4))
            TIMS.LOG.Error(eErrmsg, ex)
            Common.MessageBox(Me, "下載檔案有誤，請確認檔案是否正確!")
            Return
        End Try
        If s_PDF_byte IsNot Nothing Then
            Dim xPMS As New Hashtable
            TIMS.SetMyValue2(xPMS, "C_KBNAME2", vC_KBNAME2)
            TIMS.SetMyValue2(xPMS, "YEARS", vYEARS)
            TIMS.SetMyValue2(xPMS, "PLANID", vPLANID)
            TIMS.SetMyValue2(xPMS, "COMIDNO", vCOMIDNO)
            TIMS.SetMyValue2(xPMS, "SEQNO", vSEQNO)
            TIMS.SetMyValue2(xPMS, "CVOCID", vCVOCID)
            TIMS.SetMyValue2(xPMS, "KVSID", vKVSID)
            TIMS.SetMyValue2(xPMS, "MODIFYACCT", sm.UserInfo.UserID)
            '儲存／上傳-CLASS_VERIFYONLINE_FL
            Call SAVE_CLASS_VERIFYONLINE_FL_PDF_FILE(xPMS, s_PDF_byte)
        End If
    End Sub

    ''' <summary>學員線上簽到退明細一覽表</summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_CLASS_VERIFYONLINE_ALL_GW08(rPMS As Hashtable)
        Dim vC_KBNAME As String = TIMS.GetMyValue2(rPMS, "C_KBNAME")
        Dim vC_KBNAME2 As String = TIMS.GetMyValue2(rPMS, "C_KBNAME2")
        Dim vOCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS, "YEARS_ROC")
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim vKVSID As String = TIMS.GetMyValue2(rPMS, "KVSID")

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(rPMS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(rPMS, "SEQNO")
        Dim vMSD As String = TIMS.GetMyValue2(rPMS, "MSD")

        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID)
        Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
        If drFL IsNot Nothing Then
            Dim vMODIFY_DAY As String = Convert.ToString(drFL("MODIFY_DAY")) 'MODIFY_DAY
            Dim vMODIFY_MI As String = Convert.ToString(drFL("MODIFY_MI")) 'MODIFY_MI
            fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
        End If

        If Not fg_RUN_REPORT_1 Then
            's_TMPMSG1 &= String.Concat(If(s_TMPMSG1 <> "", ", ", ""), iRow)
            Common.MessageBox(Me, String.Concat("項目：", vC_KBNAME, ",重複處理時間過短(3分鐘1次)，請等待3分鐘後再試!", vbCrLf))
            Return
        End If

        'rPMS4.Clear()
        Dim rPMS4 As New Hashtable From {{"TPlanID", sm.UserInfo.TPlanID}, {"OCID", vOCIDVal}, {"MSD", vMSD}}
        Dim s_RPTURL As String = GET_RPTURL_SD_14_029(rPMS4)
        Dim s_PDF_byte As Byte() = Nothing
        Try
            Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
        Catch ex As Exception
            Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
            eErrmsg &= String.Concat(", s_RPTURL: ", s_RPTURL)
            eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
            eErrmsg &= String.Concat(", rPMS4: ", TIMS.GetMyValue4(rPMS4))
            TIMS.LOG.Error(eErrmsg, ex)
            Common.MessageBox(Me, "下載檔案有誤，請確認檔案是否正確!")
            Return
        End Try
        If s_PDF_byte IsNot Nothing Then
            Dim xPMS As New Hashtable
            TIMS.SetMyValue2(xPMS, "C_KBNAME2", vC_KBNAME2)
            TIMS.SetMyValue2(xPMS, "YEARS", vYEARS)
            TIMS.SetMyValue2(xPMS, "PLANID", vPLANID)
            TIMS.SetMyValue2(xPMS, "COMIDNO", vCOMIDNO)
            TIMS.SetMyValue2(xPMS, "SEQNO", vSEQNO)
            TIMS.SetMyValue2(xPMS, "CVOCID", vCVOCID)
            TIMS.SetMyValue2(xPMS, "KVSID", vKVSID)
            TIMS.SetMyValue2(xPMS, "MODIFYACCT", sm.UserInfo.UserID)
            '儲存／上傳-CLASS_VERIFYONLINE_FL
            Call SAVE_CLASS_VERIFYONLINE_FL_PDF_FILE(xPMS, s_PDF_byte)
        End If

    End Sub

    ''' <summary>'參訓學員出席紀錄一覽表</summary>
    ''' <param name="rPMS"></param>
    Private Sub SAVE_CLASS_VERIFYONLINE_ALL_GW03(rPMS As Hashtable)
        Dim vC_KBNAME As String = TIMS.GetMyValue2(rPMS, "C_KBNAME")
        Dim vC_KBNAME2 As String = TIMS.GetMyValue2(rPMS, "C_KBNAME2")
        Dim vOCIDVal As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS, "YEARS_ROC")
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim vKVSID As String = TIMS.GetMyValue2(rPMS, "KVSID")

        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(rPMS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(rPMS, "SEQNO")

        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID)
        Dim fg_RUN_REPORT_1 As Boolean = True '(執行報表)(試著搜尋看看有無資料)
        If drFL IsNot Nothing Then
            Dim vMODIFY_DAY As String = Convert.ToString(drFL("MODIFY_DAY")) 'MODIFY_DAY
            Dim vMODIFY_MI As String = Convert.ToString(drFL("MODIFY_MI")) 'MODIFY_MI
            fg_RUN_REPORT_1 = (vMODIFY_DAY <> "0" OrElse vMODIFY_MI <> "0") '(有資料 且異動時間不為0)
        End If

        If Not fg_RUN_REPORT_1 Then
            's_TMPMSG1 &= String.Concat(If(s_TMPMSG1 <> "", ", ", ""), iRow)
            Common.MessageBox(Me, String.Concat("項目：", vC_KBNAME, ",重複處理時間過短(3分鐘1次)，請等待3分鐘後再試!", vbCrLf))
            Return
        End If

        'rPMS4.Clear()
        Dim rPMS4 As New Hashtable From {{"OCID", vOCIDVal}, {"YEARS_ROC", vYEARS_ROC}}
        Dim s_RPTURL As String = GET_RPTURL_SD_14_009(rPMS4)
        Dim s_PDF_byte As Byte() = Nothing
        Try
            Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
        Catch ex As Exception
            Dim eErrmsg As String = String.Concat("##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte), ex.Message: ", ex.Message)
            eErrmsg &= String.Concat(", s_RPTURL: ", s_RPTURL)
            eErrmsg &= String.Concat(", s_PDF_byte: ", If(s_PDF_byte Is Nothing, "Is Nothing!", Convert.ToString(s_PDF_byte.Length)))
            eErrmsg &= String.Concat(", rPMS4: ", TIMS.GetMyValue4(rPMS4))
            TIMS.LOG.Error(eErrmsg, ex)
            Common.MessageBox(Me, "下載檔案有誤，請確認檔案是否正確!")
            Return
        End Try
        If s_PDF_byte IsNot Nothing Then
            Dim xPMS As New Hashtable
            TIMS.SetMyValue2(xPMS, "C_KBNAME2", vC_KBNAME2)
            TIMS.SetMyValue2(xPMS, "YEARS", vYEARS)
            TIMS.SetMyValue2(xPMS, "PLANID", vPLANID)
            TIMS.SetMyValue2(xPMS, "COMIDNO", vCOMIDNO)
            TIMS.SetMyValue2(xPMS, "SEQNO", vSEQNO)
            TIMS.SetMyValue2(xPMS, "CVOCID", vCVOCID)
            TIMS.SetMyValue2(xPMS, "KVSID", vKVSID)
            TIMS.SetMyValue2(xPMS, "MODIFYACCT", sm.UserInfo.UserID)
            '儲存／上傳-CLASS_VERIFYONLINE_FL
            Call SAVE_CLASS_VERIFYONLINE_FL_PDF_FILE(xPMS, s_PDF_byte)
        End If

    End Sub

    ''' <summary>儲存／上傳-CLASS_VERIFYONLINE_FL</summary>
    ''' <param name="rPMS"></param>
    ''' <param name="s_PDF_byte"></param>
    Private Sub SAVE_CLASS_VERIFYONLINE_FL_PDF_FILE(rPMS As Hashtable, s_PDF_byte() As Byte)

        If rPMS Is Nothing Then Return

        Dim vC_KBNAME2 As String = TIMS.GetMyValue2(rPMS, "C_KBNAME2")
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
        Dim vPLANID As String = TIMS.GetMyValue2(rPMS, "PLANID")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(rPMS, "COMIDNO")
        Dim vSEQNO As String = TIMS.GetMyValue2(rPMS, "SEQNO")
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim vKVSID As String = TIMS.GetMyValue2(rPMS, "KVSID")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim vPATTERN As String = ""
        Dim vUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(vYEARS, vAPPSTAGE, vPLANID, vRID, vBCASENO, vKBSID) 'String.Concat(G_UPDRV, "/", vYEARS, "/", vPLANID, "/", vRID, "/", vBCASENO, "/", vKBSID, "/")
        Dim vFILENAME1 As String = "" 'TIMS.GET_FILENAME1_EV(vBCID, vKBSID, vPCS, "pdf")
        Dim vSRCFILENAME1 As String = "" 'vFILENAME1 'Convert.ToString(oSRCFILENAME1)
        '上傳檔案/存檔：檔名
        Try
            vUploadPath = TIMS.GET_UPLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")
            vFILENAME1 = TIMS.GET_FILENAME1_CVU(vCVOCID, vKVSID, "pdf")
            vSRCFILENAME1 = TIMS.GET_SRCFILENAME1(vC_KBNAME2, "pdf") 'Convert.ToString(oSRCFILENAME1)
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
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim vWAIVED As String = ""
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "CVOCFID", If(vUploadPath <> "", iCVOCFID, -1)) '(可再次傳送)

            'TIMS.SetMyValue2(rPMS2, "CVOCID", vCVOCID)
            'TIMS.SetMyValue2(rPMS2, "KVSID", vKVSID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED) ' If(CHKB_WAIVED.Checked, "Y", ""))
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "PATTERN", vPATTERN)
            'TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", vMODIFYACCT)
            '儲存／上傳-CLASS_VERIFYONLINE_FL
            Call SAVE_CLASS_VERIFYONLINE_FL_UPLOAD(vCVOCID, vKVSID, rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try

    End Sub

    ''' <summary>參訓學員出席紀錄一覽表</summary>
    ''' <param name="rPMS4"></param>
    ''' <returns></returns>
    Private Function GET_RPTURL_SD_14_009(rPMS4 As Hashtable) As String
        Dim vYEARS_ROC As String = TIMS.GetMyValue2(rPMS4, "YEARS_ROC")
        Dim v_OCIDVal As String = TIMS.GetMyValue2(rPMS4, "OCID")
        Const cst_printFN3 As String = "SD_14_009_16" '2016
        Dim sfilename1 As String = cst_printFN3
        Dim MyValue As String = ""
        MyValue = "Years=" & vYEARS_ROC
        MyValue &= "&OCID=" & v_OCIDVal
        Return ReportQuery.GetReportUrl2(Me, sfilename1, MyValue)
    End Function

    ''' <summary>學員線上簽到退明細一覽表</summary>
    ''' <param name="rPMS4"></param>
    ''' <returns></returns>
    Private Function GET_RPTURL_SD_14_029(rPMS4 As Hashtable) As String
        Const cst_printFN1 As String = "SD_14_029" '2011
        Dim sfilename1 As String = cst_printFN1 'cst_printFN1

        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS4, "TPlanID")
        Dim v_OCID As String = TIMS.GetMyValue2(rPMS4, "OCID")
        Dim v_MSD As String = TIMS.GetMyValue2(rPMS4, "MSD")

        Dim sMyValue1 As String = ""
        TIMS.SetMyValue(sMyValue1, "TPlanID", vTPlanID)
        TIMS.SetMyValue(sMyValue1, "OCID", v_OCID)
        TIMS.SetMyValue(sMyValue1, "MSD", v_MSD)
        Return ReportQuery.GetReportUrl2(Me, cst_printFN1, sMyValue1)
    End Function

    ''' <summary>結訓證書清冊</summary>
    ''' <param name="rPMS4"></param>
    ''' <returns></returns>
    Private Function GET_RPTURL_SD_14_030(rPMS4 As Hashtable) As String
        Const cst_printFN1 As String = "SD_14_030" '2011
        Dim sfilename1 As String = cst_printFN1 'cst_printFN1

        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS4, "TPlanID")
        Dim v_OCID As String = TIMS.GetMyValue2(rPMS4, "OCID")
        Dim v_MSD As String = TIMS.GetMyValue2(rPMS4, "MSD")

        Dim sMyValue1 As String = ""
        TIMS.SetMyValue(sMyValue1, "TPlanID", vTPlanID)
        TIMS.SetMyValue(sMyValue1, "OCID", v_OCID)
        TIMS.SetMyValue(sMyValue1, "MSD", v_MSD)
        Return ReportQuery.GetReportUrl2(Me, cst_printFN1, sMyValue1)
    End Function

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
