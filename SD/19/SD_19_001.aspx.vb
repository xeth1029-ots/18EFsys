Imports System.IO
Imports System.Linq

Partial Class SD_19_001
    Inherits AuthBasePage

    'SELECT CONCAT('const ',ORGKINDGW,KTID ,'_',KTNAME,' as string="', ORGKINDGW,KTID ,'"') KTNAME2 FROM KEY_STD14TH ORDER BY ORGKINDGW,KTID
    'KEY_STD14TH,CLASS_STD14OA,CLASS_STD14OAFL,
    Const G01_參訓學員簽訂之契約書 As String = "G01"
    Const G02_參訓學員身分證影本與存摺影本黏貼表 As String = "G02"
    Const G03_訓練課程開班學員名冊 As String = "G03"
    Const G04_預估參訓學員補助經費清冊 As String = "G04"
    Const G05_參訓學員投保狀況檢核表 As String = "G05"
    Const G06_其他補充資料 As String = "G06"
    Const G07_公文 As String = "G07"
    Const W01_參訓學員簽訂之契約書 As String = "W01"
    Const W02_參訓學員身分證影本與存摺影本黏貼表 As String = "W02"
    Const W03_訓練課程開班學員名冊 As String = "W03"
    Const W04_預估參訓學員補助經費清冊 As String = "W04"
    Const W05_參訓學員投保狀況檢核表 As String = "W05"
    Const W06_其他補充資料 As String = "W06"
    Const W07_公文 As String = "W07"

    Dim iDG06_ROWS As Integer = 0
    Const cst_txt_其他補充資料 As String = "(其他補充資料)"
    Const cst_06_其他補充資料_WAIVED_OTH1 As String = "OTH1"

    Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
    '最近一次版本送件
    Const cst_MTYPE_LATEST_SEND1 As String = "MTYPE_LATEST_SEND1"
    '最近一次版本-下載
    Const cst_MTYPE_LATEST_DOWN1 As String = "MTYPE_LATEST_DOWN1"
    ''' <summary>儲存(暫存)</summary>
    Const cst_ACTTYPE_BTN_SAVETMP1 As String = "BTN_SAVETMP1" '儲存(暫存)
    ''' <summary>'儲存後進下一步</summary>
    Const cst_ACTTYPE_BTN_SAVENEXT1 As String = "BTN_SAVENEXT1" '儲存後進下一步
    '以目前版本批次送出
    Const cst_txt_版本批次送出 As String = "(版本批次送出)"
    Const cst_txt_免附文件 As String = "(免附文件)"
    Const cst_REUPLOADED_MSG As String = "(已重新上傳)"
    'sPrintASPX1=String.Concat(cst_printASPX_R, TIMS.Get_MRqID(Me))
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
    'Const cst_PostedFile_MAX_SIZE_15M As Integer = 15728640 '1024*1024*15=15728640
    Const cst_PostedFile_MAX_SIZE_20M As Integer = 20971520 '1024*1024*20=20971520
    'Const cst_errMsg_7 As String="檔案大小超過2MB!"
    Const cst_errMsg_7_10M As String = "檔案大小超過10MB!"
    'Const cst_errMsg_7_15M As String = "檔案大小超過15MB!"
    Const cst_errMsg_7_20M As String = "檔案大小超過20MB!"
    Const cst_FileDescMsg_7_10M As String = "PDF(掃瞄畫面需清楚，檔案大小限制10MB以下)!"
    'Const cst_FileDescMsg_7_15M As String = "PDF(掃瞄畫面需清楚，檔案大小限制15MB以下)!"
    Const cst_FileDescMsg_7_20M As String = "PDF(掃瞄畫面需清楚，檔案大小限制20MB以下)!"

    Const cst_errMsg_8 As String = "請選擇上傳檔案(不可為空)!"
    'Const cst_errMsg_9 As String="請選擇場地圖片--隸屬於教室1 或教室2!"
    'Const cst_errMsg_11 As String="無效的檔案格式。"
    Const cst_errMsg_11_PDF As String = "無效的檔案格式(限PDF檔案)。"
    Const cst_errMsg_21 As String = "不可勾選免附文件又按上傳檔案。"

    'Add new application cases and add reminder messages
    Const Cst_messages1 As String = "請務必確認此年度/申請階段之所有欲研提班級都已送審，【新增申辦案件】後才送審的班級，將無法納入此次線上申辦案件清單中!"
    Const cst_ss_messages1 As String = "messages1"
    Const Cst_messages2 As String = "單位按下「確認更新」後，為確保資料正確性，原已上傳之班級相關文件從第7項開始均會自動清除，單位須重新上傳。"

    Dim tmpMSG As String = ""
    Dim ff3 As String = ""
    Dim objconn As SqlConnection = Nothing

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        If Not IsPostBack Then
            Call CCreate1(0)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'Dim jsScript1 As String = Common.GetJsString($"{Cst_messages1}{vbCrLf}確定是否要新增?")
        Dim jsScript1 As String = Common.GetJsString($"確定是否要新增?")
        BTN_ADDNEW1.Attributes("onclick") = $"return confirm('{jsScript1}');"

        '(顯示1次訊息)(有顯示過就存session)
        'If Session(cst_ss_messages1) Is Nothing OrElse $"{Session(cst_ss_messages1)}" <> Cst_messages1 Then
        '    Session(cst_ss_messages1) = Cst_messages1
        '    Common.MessageBox(Me, Cst_messages1)
        'End If

    End Sub

    ''' <summary>
    ''' '設定 資料與顯示 狀況！iNum:0:search,1:edit
    ''' </summary>
    ''' <param name="iNum"></param>
    Private Sub CCreate1(ByVal iNum As Integer)
        labmsg1.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        If Session(cst_ss_messages1) IsNot Nothing Then Session(cst_ss_messages1) = Nothing
        TB_DataGrid1.Visible = False

        Call SHOW_Frame1(0)
        '案件編號
        SCH_TBCASENO.Text = ""

        '申請階段
        Dim v_APPSTAGE As String = TIMS.GET_CANUSE_APPSTAGE(objconn, CStr(sm.UserInfo.Years), TIMS.cst_APPSTAGE_PTYPE1_01)
        SCH_DDLAPPSTAGE = TIMS.Get_APPSTAGE2(SCH_DDLAPPSTAGE)
        Common.SetListItem(SCH_DDLAPPSTAGE, v_APPSTAGE)

        '申辦日期
        SCH_BIDATE1.Text = ""
        SCH_BIDATE2.Text = ""

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Get_TitleLab(objconn, MRqID, TitleLab1, TitleLab2)
    End Sub

    ''' <summary>顯示調整 iNum:0:search,1:edit</summary>
    ''' <param name="iNum"></param>
    Private Sub SHOW_Frame1(ByVal iNum As Integer)
        FrameTableSch1.Visible = If(iNum = 0, True, False)
        FrameTableEdt1.Visible = If(iNum = 1, True, False)
    End Sub

    Sub UTL_FMTINPUTVAL()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim v_SCH_DDLAPPSTAGE As String = TIMS.GetListValue(SCH_DDLAPPSTAGE)
        SCH_TBCASENO.Text = TIMS.ClearSQM(SCH_TBCASENO.Text)
        SCH_STDATE1.Text = TIMS.Cdate3(SCH_STDATE1.Text)
        SCH_STDATE2.Text = TIMS.Cdate3(SCH_STDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(SCH_STDATE1.Text, SCH_STDATE2.Text) Then
            Dim T_DATE1 As String = SCH_STDATE1.Text
            SCH_STDATE1.Text = SCH_STDATE2.Text
            SCH_STDATE2.Text = T_DATE1
        End If
        SCH_FTDATE1.Text = TIMS.Cdate3(SCH_FTDATE1.Text)
        SCH_FTDATE2.Text = TIMS.Cdate3(SCH_FTDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(SCH_FTDATE1.Text, SCH_FTDATE2.Text) Then
            Dim T_DATE1 As String = SCH_FTDATE1.Text
            SCH_FTDATE1.Text = SCH_FTDATE2.Text
            SCH_FTDATE2.Text = T_DATE1
        End If
        SCH_BIDATE1.Text = TIMS.Cdate3(SCH_BIDATE1.Text)
        SCH_BIDATE2.Text = TIMS.Cdate3(SCH_BIDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(SCH_BIDATE1.Text, SCH_BIDATE2.Text) Then
            Dim T_DATE1 As String = SCH_BIDATE1.Text
            SCH_BIDATE1.Text = SCH_BIDATE2.Text
            SCH_BIDATE2.Text = T_DATE1
        End If
    End Sub

    ''' <summary>案件資料查詢</summary>
    ''' <param name="rPMS"></param>
    ''' <returns></returns>
    Function SEARCH_DATA1_ROW(rPMS As Hashtable) As DataRow
        Dim V_TBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim V_OCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        If V_TBCID = "" AndAlso V_OCID = "" Then Return Nothing

        'labTCASENO.Text = $"{drOA("TBCASENO")}"
        'LabCREATEDATE.Text = $"{drOA("TBCDATE_ROC")}"
        Dim PMS1 As New Hashtable From {{"YEARS", $"{sm.UserInfo.Years}"}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim SQL1 As String = "SELECT a.TBCID,a.TBCASENO,a.OCID,a.APPSTAGE
,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N
,a.TBCACCT,dbo.FN_GET_USERNAME(a.TBCACCT) TBCNAME,a.TBCDATE,dbo.FN_CDATE1B(a.TBCDATE) TBCDATE_ROC
,a.TBCSTATUS,CASE WHEN a.TBCSTATUS IS NULL THEN '暫存'
WHEN a.TBCSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '分署已收件'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END TBCSTATUS_N
,a.APPLIEDRESULT,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N
,a.REASONFORFAIL,a.RESULTACCT,a.RESULTDATE,a.HISREVIEW
,cc.DISTNAME,cc.ORGNAME,cc.CLASSCNAME2,cc.ORGKINDGW
FROM CLASS_STD14OA a
JOIN VIEW2 cc on cc.OCID=a.OCID 
WHERE cc.YEARS=@YEARS AND cc.TPLANID=@TPLANID"
        If V_TBCID <> "" Then
            PMS1.Add("TBCID", TIMS.CINT1(V_TBCID))
            SQL1 &= " AND a.TBCID=@TBCID"
        End If
        If V_OCID <> "" Then
            PMS1.Add("OCID", TIMS.CINT1(V_OCID))
            SQL1 &= " AND a.OCID=@OCID"
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(SQL1, objconn, PMS1)
        If TIMS.dtNODATA(dt) Then Return Nothing
        Return dt.Rows(0)
    End Function

    Sub SEARCH_1()
        labmsg1.Text = TIMS.cst_NODATAMsg1
        TB_DataGrid1.Visible = False

        Call UTL_FMTINPUTVAL()
        Dim v_SCH_DDLAPPSTAGE As String = TIMS.GetListValue(SCH_DDLAPPSTAGE)

        Dim PMS1 As New Hashtable From {{"YEARS", $"{sm.UserInfo.Years}"}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim SQL1 As String = "
SELECT a.TBCID,a.TBCASENO,a.OCID,a.APPSTAGE
,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N
,a.TBCACCT,dbo.FN_GET_USERNAME(a.TBCACCT) TBCNAME,a.TBCDATE,dbo.FN_CDATE1B(a.TBCDATE) TBCDATE_ROC
,a.TBCSTATUS, CASE WHEN a.TBCSTATUS IS NULL THEN '暫存'
WHEN a.TBCSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '分署已收件'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END TBCSTATUS_N
,a.APPLIEDRESULT,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N
,cc.DISTNAME,cc.ORGNAME,cc.CLASSCNAME2,cc.ORGKINDGW
FROM CLASS_STD14OA a
JOIN VIEW2 cc on cc.OCID=a.OCID
WHERE cc.YEARS=@YEARS AND cc.TPLANID=@TPLANID"
        Select Case sm.UserInfo.LID
            Case 0
                If RIDValue.Value.Length > 1 Then
                    PMS1.Add("RID", RIDValue.Value)
                    SQL1 &= " AND cc.RID=@RID"
                ElseIf RIDValue.Value <> "A" AndAlso RIDValue.Value.Length = 1 Then
                    Dim V_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    PMS1.Add("DISTID", V_DistID)
                    SQL1 &= " AND cc.DISTID=@DISTID"
                End If
            Case 1
                PMS1.Add("DISTID", sm.UserInfo.DistID)
                SQL1 &= " AND cc.DISTID=@DISTID"
                If RIDValue.Value.Length > 1 Then
                    PMS1.Add("RID", RIDValue.Value)
                    SQL1 &= " AND cc.RID=@RID"
                End If
            Case 2
                PMS1.Add("RID", RIDValue.Value)
                SQL1 &= " AND cc.RID=@RID"
        End Select
        If OCIDValue1.Value <> "" Then
            PMS1.Add("OCID", TIMS.CINT1(OCIDValue1.Value))
            SQL1 &= " AND a.OCID=@OCID"
        End If
        If v_SCH_DDLAPPSTAGE <> "" Then
            PMS1.Add("APPSTAGE", v_SCH_DDLAPPSTAGE)
            SQL1 &= " AND a.APPSTAGE=@APPSTAGE"
        End If
        If SCH_TBCASENO.Text <> "" Then
            PMS1.Add("TBCASENO", SCH_TBCASENO.Text)
            SQL1 &= " AND a.TBCASENO like '%'+@TBCASENO+'%'"
        End If
        If SCH_STDATE1.Text <> "" Then
            PMS1.Add("STDATE1", SCH_STDATE1.Text)
            SQL1 &= " AND a.STDATE>=@STDATE1"
        End If
        If SCH_STDATE2.Text <> "" Then
            PMS1.Add("STDATE2", SCH_STDATE2.Text)
            SQL1 &= " AND a.STDATE<=@STDATE2"
        End If
        If SCH_FTDATE1.Text <> "" Then
            PMS1.Add("FTDATE1", SCH_FTDATE1.Text)
            SQL1 &= " AND a.FTDATE>=@FTDATE1"
        End If
        If SCH_FTDATE2.Text <> "" Then
            PMS1.Add("FTDATE2", SCH_FTDATE2.Text)
            SQL1 &= " AND a.FTDATE<=@FTDATE2"
        End If
        If SCH_BIDATE1.Text <> "" Then
            PMS1.Add("TBCDATE1", SCH_BIDATE1.Text)
            SQL1 &= " AND a.TBCDATE>=@TBCDATE1"
        End If
        If SCH_BIDATE2.Text <> "" Then
            PMS1.Add("TBCDATE2", SCH_BIDATE2.Text)
            SQL1 &= " AND a.TBCDATE<=@TBCDATE2"
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(SQL1, objconn, PMS1)

        If TIMS.dtNODATA(dt) Then
            labmsg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        labmsg1.Text = ""
        TB_DataGrid1.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Function GET_TBCASENO_NN(oConn As SqlConnection, rParms As Hashtable) As String
        Dim v_YEARS_ROC As String = TIMS.GET_YEARS_ROC(Now.Year)
        Dim v_NMD4 As String = Now.ToString("MMdd") 'TIMS.GET_YEARS_ROC()
        Dim v_RID As String = TIMS.GetMyValue2(rParms, "RID")
        Dim vRID1 As String = If(v_RID.Length >= 1, Left(v_RID, 1), v_RID)
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rParms, "APPSTAGE")
        Dim v_CASENO_T1 As String = String.Concat(v_YEARS_ROC, vRID1, vAPPSTAGE, v_NMD4)
        'ROC_YEAR+RID1+APPSTAGE+MMdd
        '3+1+1+4=9 
        Dim sParms As New Hashtable From {{"TBCASENOT1", v_CASENO_T1}}
        '112B27060001' BCASENO_MAX_SEQ
        Dim sql As String = "SELECT MAX(CAST(SUBSTRING(TBCASENO,10,9) AS INT)) BCASENO_MAX_SEQ FROM CLASS_STD14OA WITH(NOLOCK) WHERE SUBSTRING(TBCASENO,1,9)=SUBSTRING(@TBCASENOT1,1,9)"
        Dim drB1 As DataRow = DbAccess.GetOneRow(sql, oConn, sParms)
        Dim fg_NODATA As Boolean = (drB1 Is Nothing OrElse $"{drB1("BCASENO_MAX_SEQ")}" = "")
        Dim iBCASENO_LAST_SEQ As Integer = If(fg_NODATA, 1, TIMS.CINT1(drB1("BCASENO_MAX_SEQ")) + 1)
        Return String.Concat(v_CASENO_T1, TIMS.AddZero(iBCASENO_LAST_SEQ.ToString(), 4))
    End Function

    Function Utl_ADDNEW_DATA1(ByRef sERRMSG As String, ByRef drCC As DataRow) As Boolean
        '清理隱藏的參數
        Call ClearHidValue()
        sERRMSG = ""
        If drCC Is Nothing Then
            sERRMSG = "資訊有誤(查無職類/班別代碼)"
            Return False
        End If

        Dim PMS_S As New Hashtable From {{"OCID", TIMS.CINT1(drCC("OCID"))}}
        Dim SQL_S As String = "SELECT 1 FROM CLASS_STD14OA WHERE OCID=@OCID"
        Dim dtS1 As DataTable = DbAccess.GetDataTable(SQL_S, objconn, PMS_S)
        If TIMS.dtHaveDATA(dtS1) Then
            sERRMSG = "班級已有線上申辦案件(一個線上申辦案件只能包含一個班級)!"
            Return False
        End If

        Dim vTBCASENO_NN As String = ""
        Dim iTBCID As Integer = 0
        Try
            iTBCID = DbAccess.GetNewId(objconn, "CLASS_STD14OA_TBCID_SEQ,CLASS_STD14OA,TBCID")
            Dim TbcPMS1 As New Hashtable From {{"RID", $"{drCC("RID")}"}, {"APPSTAGE", $"{drCC("APPSTAGE")}"}}
            vTBCASENO_NN = GET_TBCASENO_NN(objconn, TbcPMS1)

            Dim PMS_A As New Hashtable From {
                {"TBCID", iTBCID},
                {"TBCASENO", vTBCASENO_NN},
                {"OCID", TIMS.CINT1(drCC("OCID"))},
                {"APPSTAGE", $"{drCC("APPSTAGE")}"},
                {"CREATEACCT", sm.UserInfo.UserID},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            Dim SQL_A As String = "INSERT INTO CLASS_STD14OA(TBCID,TBCASENO,OCID,APPSTAGE, CREATEACCT,CREATEDATE,MODIFYACCT,MODIFYDATE)
VALUES (@TBCID,@TBCASENO,@OCID,@APPSTAGE ,@CREATEACCT,GETDATE(),@MODIFYACCT,GETDATE())"
            DbAccess.ExecuteNonQuery(SQL_A, objconn, PMS_A)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            sERRMSG = "資料庫序號有誤，請重新操作!"
            Return False
        End Try

        Hid_TBCID.Value = $"{iTBCID}"
        Hid_TBCASENO.Value = $"{vTBCASENO_NN}"
        Hid_ORGKINDGW.Value = $"{drCC("ORGKINDGW")}"
        Return True
    End Function

    Sub UTL_ADDNEW1()
        Call UTL_FMTINPUTVAL()
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "資訊有誤，請選擇 職類/班別!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇 職類/班別!")
            Return
        End If

        Dim sERRMSG As String = ""
        Dim fg_CanADDNEW1 As Boolean = Utl_ADDNEW_DATA1(sERRMSG, drCC)
        If sERRMSG <> "" Then
            Common.MessageBox(Me, sERRMSG)
            Return
        End If

        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Call SHOW_Detail_STD14OA(drCC, Hid_TBCID.Value, "")
    End Sub

    Private Sub SHOW_Detail_STD14OA(drCC As DataRow, V_TBCID As String, vCmdName As String)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)，請選擇 職類/班別!")
            Return
        End If
        If V_TBCID = "" Then
            Common.MessageBox(Me, "線上申辦案件序號有誤!")
            Return
        End If
        Dim V_PMS1 As New Hashtable From {{"TBCID", TIMS.CINT1(V_TBCID)}, {"OCID", TIMS.CINT1(drCC("OCID"))}}
        Dim drOA As DataRow = SEARCH_DATA1_ROW(V_PMS1)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無線上申辦案件)!!")
            Return
        End If

        Session(cst_ss_RqProcessType) = $"{vCmdName}"
        Call SHOW_Frame1(1)

        labOrgNAME.Text = $"{drCC("ORGNAME")}"
        labTBCASENO.Text = $"{drOA("TBCASENO")}"
        LabCREATEDATE.Text = $"{drOA("TBCDATE_ROC")}"
        labBIYEARS.Text = $"{drCC("PLANYEARS")}"
        labAPPSTAGE.Text = $"{drOA("APPSTAGE_N")}"
        labCLASSNAME2S.Text = $"{drCC("CLASSCNAME2")}"
        tr_HISREVIEW.Visible = False
        'labHISREVIEW.Text = "" ' $"{drCC("ORGNAME")}"

        ClearHidValue()
        Hid_TBCID.Value = V_TBCID
        Hid_TBCASENO.Value = $"{drOA("TBCASENO")}"
        Hid_ORGKINDGW.Value = $"{drCC("ORGKINDGW")}"
        ddlSwitchTo = TIMS.GET_DDL_STD14TH(sm, objconn, ddlSwitchTo, Hid_ORGKINDGW.Value)
        Dim rLastKTID As String = ""
        Dim rFirstKTSEQ As String = ""
        Call TIMS.GET_SWITCHTO_VAL_OA(sm, objconn, Hid_ORGKINDGW.Value, rLastKTID, rFirstKTSEQ)
        Hid_LastKTID.Value = rLastKTID
        Hid_FirstKTSEQ.Value = rFirstKTSEQ

        labProgress.Text = "0%"

        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)

        'Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_KTSEQ.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_FirstKTSEQ.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    Protected Sub BTN_SEARCH1_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH1.Click
        Call SEARCH_1()
    End Sub

    Protected Sub BTN_ADDNEW1_Click(sender As Object, e As EventArgs) Handles BTN_ADDNEW1.Click
        Call UTL_ADDNEW1()
    End Sub

    Protected Sub BTN_SEARCH2_Click(sender As Object, e As EventArgs) Handles BTN_SEARCH2.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_KTSEQ.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_FirstKTSEQ.Value, Hid_ORGKINDGW.Value)
        End If
        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)
    End Sub

    Private Sub SHOW_STD14OAFL_DG2(rPMS As Hashtable)
        labmsg1.Text = ""
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        'objconn 因為有檔案輸出關閉的問題 所以要檢查
        If Not TIMS.OpenDbConn(objconn) OrElse Not fg_CANSAVE Then Return

        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)!")
            Return
        End If

        Dim dtFL As DataTable = GET_CLASS_STD14OAFL_TB(objconn, vTBCID, vORGKINDGW)
        labmsg1.Text = If(dtFL Is Nothing OrElse dtFL.Rows.Count = 0, "(查無文件項目)", "")

        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vPLANID As String = $"{drCC("PLANID")}"
        Dim vTBCASENO As String = $"{drOA("TBCASENO")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, "")
        Call TIMS.Check_dtSTD14OAFL(Me, dtFL, download_Path)
        DataGrid2.Columns(cst_DG2_退件原因_iCOLUMN).Visible = If($"{drOA("APPLIEDRESULT")}" = "R", True, False)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

        'Dim iProgress As Integer=If(dtA.Rows.Count > 0, (dt.Rows.Count / dtA.Rows.Count * 100), 0)
        '線上申辦進度 計算完成度百分比 (0-100)
        Dim iProgress As Integer = TIMS.GET_iPROGRESS_OA(sm, objconn, tmpMSG, vTBCID, vORGKINDGW)
        labProgress.Text = $"{iProgress}%"
        'BTN_SAVETMP1.Visible=(iProgress=100) 'BTN_SAVERC2.Visible=(iProgress=100)
        '儲存(暫存)
        BTN_SAVETMP1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, "", cst_tpmsg_enb1), True)
        '儲存後進下一步
        BTN_SAVENEXT1.Enabled = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
        TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, "", cst_tpmsg_enb1), True)
    End Sub

    Private Function GET_CLASS_STD14OAFL_TB(oConn As SqlConnection, vTBCID As String, vORGKINDGW As String) As DataTable
        Dim rParms As New Hashtable From {{"TBCID", TIMS.CINT1(vTBCID)}, {"ORGKINDGW", vORGKINDGW}}
        Dim rsSql As String = "
SELECT A.TBCFID,A.TBCID,A.KTSEQ,A.PATTERN,A.MEMO1,A.MODIFYACCT,A.MODIFYDATE
,KB.KTID,CONCAT(KB.KTID,'.',KB.KTNAME) KTNAME2
,OA.TBCSTATUS,OA.APPLIEDRESULT
,A.RTUREASON,A.WAIVED,A.SRCFILENAME1,A.FILEPATH1,A.FILENAME1,A.FILENAME1 OKFLAG
FROM CLASS_STD14OAFL A
JOIN KEY_STD14TH KB ON KB.KTSEQ=A.KTSEQ
JOIN CLASS_STD14OA OA ON OA.TBCID=A.TBCID
JOIN CLASS_CLASSINFO CC ON CC.OCID=OA.OCID
WHERE A.TBCID=@TBCID AND KB.ORGKINDGW=@ORGKINDGW
ORDER BY KB.KSORT,A.TBCFID"
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        Return dt
    End Function
    ''' <summary>清理隱藏的參數</summary>
    Sub ClearHidValue()
        Hid_TBCID.Value = ""
        Hid_TBCASENO.Value = ""
        Hid_ORGKINDGW.Value = ""
        Hid_FirstKTSEQ.Value = ""
        Hid_LastKTID.Value = ""

        Hid_KTSEQ.Value = ""
        Hid_KTID.Value = ""
        Hid_TBCFID.Value = ""
    End Sub
    Sub SAVEDATA2_BTN_ACTION1(ByVal s_ACTTYPE As String)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_KTSEQ.Value = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Hid_KTID.Value = TIMS.ClearSQM(Hid_KTID.Value)
        Hid_TBCFID.Value = TIMS.ClearSQM(Hid_TBCFID.Value)
        Hid_FirstKTSEQ.Value = TIMS.ClearSQM(Hid_FirstKTSEQ.Value)
        Hid_LastKTID.Value = TIMS.ClearSQM(Hid_LastKTID.Value)
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)

        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無案件編號)，請重新操作!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無項目編號)，請重新操作!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無職類/班別代碼)，請重新操作!")
            Return
        End If

        'Dim fg_FILE_MUSTBE_UPLOADED As Boolean=True '必須上傳檔案
        Dim fg_FILE_MUSTBE_UPLOADED As Boolean = True
        Dim vWAIVED As String = If(CHKB_WAIVED.Checked, "Y", "") '免附文件
        Dim vNOPROCESD As String = $"{drKB("NOPROCESD")}"
        Dim vKTSEQ As String = $"{drKB("KTSEQ")}"
        Dim vKTID As String = $"{drKB("KTID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Dim vTBCID As String = TIMS.ClearSQM(Hid_TBCID.Value)
        'Dim vKBSID As String=TIMS.ClearSQM(Hid_KBSID.Value)
        Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, vTBCID, vKTSEQ)
        '(退件修正)有退件原因,可重新儲存
        Dim fg_CanSaveAgain_1 As Boolean = (drFL IsNot Nothing AndAlso $"{drFL("RTUREASON")}" <> "") '(有資料 且原因不為空 可再次傳送)
        Dim vFILENAME1 As String = If(drFL IsNot Nothing, $"{drFL("FILENAME1")}", "")
        'Dim vWAIVED As String=If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
        If vNOPROCESD = "" AndAlso fg_FILE_MUSTBE_UPLOADED AndAlso Not fg_CanSaveAgain_1 AndAlso drFL IsNot Nothing AndAlso Not CHKB_WAIVED.Checked AndAlso vFILENAME1 = "" Then
            Common.MessageBox(Me, "未上傳檔案也未勾選免附且儲存過該文件，不可再次操作!")
            Return
        End If

        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vPLANID As String = $"{drCC("PLANID")}"
        Dim vTBCASENO As String = $"{drOA("TBCASENO")}"
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, "")
        Try
            'SAVE CLASS_STD14OAFL
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            If (drFL IsNot Nothing AndAlso fg_CanSaveAgain_1) Then TIMS.SetMyValue2(rPMS2, "TBCFID", drFL("TBCFID"))

            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "TBCID", vTBCID)
            TIMS.SetMyValue2(rPMS2, "KTSEQ", vKTSEQ)
            TIMS.SetMyValue2(rPMS2, "MEMO1", txtMEMO1.Text)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            TIMS.SetMyValue2(rPMS2, "WAIVED", vWAIVED)
            Select Case vWAIVED
                Case "Y", ""
                    Call SAVE_CLASS_STD14OAFL_UPLOAD(rPMS2)
            End Select
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.Message)

            Dim strErrmsg As String = $"ex.ToString:{ex.ToString}{vbCrLf}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub  'Throw ex
        End Try

        '儲存/UPDATE CLASS_STD14OA'暫時儲存／正式儲存
        Call SAVEDATE1()

        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)

        Select Case s_ACTTYPE
            Case cst_ACTTYPE_BTN_SAVETMP1
                '儲存(暫存) 

                '項目(重跑1次)
                Call SHOW_STD14TH_KTSEQ(Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)

            Case cst_ACTTYPE_BTN_SAVENEXT1
                '儲存後進下一步 

                '(檢查儲存值)
                Dim rPMSCHK As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}, {"KTSEQ", Hid_KTSEQ.Value}}
                Dim flag_OK_OBFL As Boolean = CHK_CLASS_STD14OAFL(rPMSCHK)
                If Not flag_OK_OBFL Then
                    Common.MessageBox(Me, "請確認 上傳資料或勾選內容 再進行下一步")
                    Return
                End If

                '下一步
                Call MOVE_NEXT()
        End Select
    End Sub

    Private Function CHK_CLASS_STD14OAFL(rPMS As Hashtable) As Boolean
        '(外部參數)
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim vKTSEQ As String = TIMS.GetMyValue2(rPMS, "KTSEQ")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        If vTBCID = "" OrElse vKTSEQ = "" OrElse Not fg_CANSAVE Then Return False '(異常)

        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, vKTSEQ, vORGKINDGW)
        If drKB Is Nothing Then Return False '(異常)

        Dim rParms As New Hashtable From {
            {"TBCID", TIMS.CINT1(vTBCID)},
            {"ORGKINDGW", vORGKINDGW},
            {"KTSEQ", vKTSEQ}
        }
        Dim rsSql As String = "
SELECT a.TBCFID,a.TBCID,a.KTSEQ,a.PATTERN,a.MEMO1,a.MODIFYACCT,a.MODIFYDATE,kb.KTID,concat(kb.KTID,'.',kb.KTNAME) KTNAME2
,a.WAIVED,a.SRCFILENAME1,A.FILEPATH1,a.FILENAME1,a.FILENAME1 OKFLAG
FROM CLASS_STD14OAFL a
JOIN KEY_STD14TH kb on kb.KTSEQ=a.KTSEQ
JOIN CLASS_STD14OA OA on OA.TBCID=a.TBCID
JOIN CLASS_CLASSINFO CC on CC.OCID=OA.OCID
WHERE a.TBCID=@TBCID AND kb.ORGKINDGW=@ORGKINDGW AND a.KTSEQ=@KTSEQ"
        '(若有多筆排序)rsSql &= " ORDER BY kb.KBID,a.KBSID,a.BCFID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, objconn, rParms)
        If TIMS.dtNODATA(dt) Then Return False '(異常)
        Return True '(異常)
    End Function

    Private Sub MOVE_NEXT()
        If Hid_KTID.Value <> "" AndAlso (Hid_KTID.Value = Hid_LastKTID.Value) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        ElseIf (ddlSwitchTo.SelectedIndex + 1 >= ddlSwitchTo.Items.Count) Then
            Common.MessageBox(Me, "(目前沒有下一步)")
            Return
        End If
        '下一步 
        Hid_KTSEQ.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex + 1).Value
        Call SHOW_STD14TH_KTSEQ(Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
    End Sub

    ''' <summary>儲存/UPDATE CLASS_STD14OA,MODIFYACCT,MODIFYDATE</summary>
    Private Sub SAVEDATE1()
        'iNum As Integer'iNum:0 暫時儲存/1 正式儲存
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Dim iTBCID As Integer = If(Hid_TBCID.Value <> "", TIMS.CINT1(Hid_TBCID.Value), 0)
        If Hid_TBCID.Value = "" OrElse iTBCID <= 0 Then Return

        Dim uParms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"TBCID", iTBCID}, {"TBCASENO", Hid_TBCASENO.Value}}
        Dim usSql As String = ""
        usSql &= " UPDATE CLASS_STD14OA" & vbCrLf
        usSql &= " SET MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE TBCID=@TBCID AND TBCASENO=@TBCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
    End Sub

    Function SAVE_CLASS_STD14OAFL_UPLOAD(rPMS As Hashtable) As Integer
        Dim iTBCFID As Integer = -1

        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        If Not fg_CANSAVE Then Return iTBCFID

        '重審 vUploadPath / vBCFID 
        Dim vUploadPath As String = TIMS.GetMyValue2(rPMS, "UploadPath")
        Dim vTBCFID As String = TIMS.GetMyValue2(rPMS, "TBCFID")
        iTBCFID = If(vTBCFID <> "" AndAlso TIMS.CINT1(vTBCFID) > 0, TIMS.CINT1(vTBCFID), -1)
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS, "FILEPATH1")
        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim vKTSEQ As String = TIMS.GetMyValue2(rPMS, "KTSEQ")
        Dim vWAIVED As String = TIMS.GetMyValue2(rPMS, "WAIVED")
        Dim vMEMO1 As String = TIMS.GetMyValue2(rPMS, "MEMO1")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        '(若檔名或路徑是空白則全部清空)
        If vFILENAME1 = "" OrElse vFILEPATH1 = "" Then
            vFILENAME1 = "" : vFILEPATH1 = "" : vSRCFILENAME1 = ""
        End If
        '免附文件或上傳檔案
        Dim fg_NG_SAVE As Boolean = (vWAIVED = "" AndAlso (vFILENAME1 = "" AndAlso vSRCFILENAME1 = ""))
        If fg_NG_SAVE Then Return iTBCFID
        'WAIVED: 只能是Y/ ""
        'Dim fg_WAIVED_CAN_SAVE As Boolean=(vWAIVED="" OrElse vWAIVED="Y")
        'If Not fg_WAIVED_CAN_SAVE Then Return iBCFID

        Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, vTBCID, vKTSEQ)
        If drFL IsNot Nothing Then
            Dim OldvFILENAME1 As String = If(drFL IsNot Nothing, $"{drFL("FILENAME1")}", "")
            Dim OldvSRCFILENAME1 As String = If(drFL IsNot Nothing, $"{drFL("SRCFILENAME1")}", "")
            Dim vRTUREASON As String = If(drFL IsNot Nothing, $"{drFL("RTUREASON")}", "")

            '(重新儲存) (加入訊息) 'iBCFID > 0 
            Dim vREUPLOADED_MSG As String = ""
            If iTBCFID > 0 Then
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
                {"WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull)},
                {"MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull)},
                {"MODIFYACCT", vMODIFYACCT},
                {"TBCFID", TIMS.CINT1(drFL("TBCFID"))}
            }
            If vWAIVED = "Y" OrElse vFILENAME1 <> "" Then uParms.Add("FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull)) 'vFILENAME1)
            If vWAIVED = "Y" OrElse vSRCFILENAME1 <> "" Then uParms.Add("SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull)) 'vSRCFILENAME1)
            If vWAIVED = "Y" OrElse vFILEPATH1 <> "" Then uParms.Add("FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull))
            If iTBCFID > 0 Then uParms.Add("RTUREASON", If(vREUPLOADED_MSG <> "", vREUPLOADED_MSG, If(vRTUREASON <> "", vRTUREASON, Convert.DBNull)))

            Dim usSql As String = ""
            usSql &= " UPDATE CLASS_STD14OAFL" & vbCrLf
            usSql &= " SET WAIVED=@WAIVED,MEMO1=@MEMO1,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            If vWAIVED = "Y" OrElse vFILENAME1 <> "" Then usSql &= " ,FILENAME1=@FILENAME1" & vbCrLf
            If vWAIVED = "Y" OrElse vSRCFILENAME1 <> "" Then usSql &= " ,SRCFILENAME1=@SRCFILENAME1" & vbCrLf
            If vWAIVED = "Y" OrElse vFILEPATH1 <> "" Then usSql &= " ,FILEPATH1=@FILEPATH1" & vbCrLf
            If iTBCFID > 0 Then usSql &= " ,RTUREASON=@RTUREASON" & vbCrLf
            usSql &= " WHERE TBCFID=@TBCFID" & vbCrLf
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
                    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                    TIMS.WriteTraceLog(Me, ex, strErrmsg)
                End Try
            End If

            iTBCFID = TIMS.CINT1(drFL("TBCFID"))
        Else
            iTBCFID = DbAccess.GetNewId(objconn, "CLASS_STD14OAFL_TBCFID_SEQ,CLASS_STD14OAFL,TBCFID")
            'k_KBGWID="KBGID" 'k_BCFGWID="BCFGID"
            Dim iParms As New Hashtable From {
                {"TBCFID", iTBCFID},
                {"TBCID", TIMS.CINT1(vTBCID)},
                {"KTSEQ", TIMS.CINT1(vKTSEQ)},
                {"WAIVED", If(vWAIVED <> "", vWAIVED, Convert.DBNull)},
                {"FILENAME1", If(vFILENAME1 <> "", vFILENAME1, Convert.DBNull)}, 'vFILENAME1)
                {"SRCFILENAME1", If(vSRCFILENAME1 <> "", vSRCFILENAME1, Convert.DBNull)}, 'vSRCFILENAME1)
                {"FILEPATH1", If(vFILEPATH1 <> "", vFILEPATH1, Convert.DBNull)},
                {"MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull)},
                {"MODIFYACCT", vMODIFYACCT}
            }
                'iParms.Add("MODIFYDATE", MODIFYDATE)
                Dim isSql As String = "
INSERT INTO CLASS_STD14OAFL(TBCFID,TBCID,KTSEQ,FILENAME1,SRCFILENAME1,FILEPATH1,MEMO1,MODIFYACCT,MODIFYDATE,WAIVED)
VALUES (@TBCFID,@TBCID,@KTSEQ,@FILENAME1,@SRCFILENAME1,@FILEPATH1,@MEMO1,@MODIFYACCT,GETDATE(),@WAIVED)"
                DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
            End If

            '儲存/UPDATE CLASS_STD14OA'暫時儲存／正式儲存
            Call SAVEDATE1()

        Return iTBCFID
    End Function
    ''' <summary>'切換項目(預設)KEY_STD14TH</summary>
    ''' <param name="vKTSEQ"></param>
    ''' <param name="vORGKINDGW"></param>
    Sub SHOW_STD14TH_KTSEQ(vKTSEQ As String, vORGKINDGW As String)
        'Dim vORGKINDGW As String=Hid_ORGKINDGW.Value
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Dim v_ddlSwitchTo As String = TIMS.GetListValue(ddlSwitchTo)
        If (v_ddlSwitchTo <> vKTSEQ) Then
            Hid_KTSEQ.Value = vKTSEQ
            Common.SetListItem(ddlSwitchTo, vKTSEQ)
        End If

        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, vKTSEQ, vORGKINDGW)
        If drKB Is Nothing Then Return
        ',MUSTFILL,TPLANID,ORGKINDGW,KSORT,USELATESTVER,DOWNLOADRPT,RPTNAME,UPLOADFL1,SENTBATVER,USEMEMO1,SENDCURRVER,NOTKBDESC1,NOTFLDESC1,NOPROCESD FROM KEY_STD14TH
        '代號／非流水號
        Dim vKTID As String = $"{drKB("KTID")}"
        '取得文字說明
        Dim vKTDESC1 As String = $"{drKB("KTDESC1")}"
        Dim vNOTKBDESC1 As String = $"{drKB("NOTKBDESC1")}" '(Y:不使用KBDESC)
        Dim vNOTFLDESC1 As String = $"{drKB("NOTFLDESC1")}" '(Y:不使用FLDESC1)
        '必填資訊／免附文件(必填就不顯示)
        Dim vMUSTFILL As String = $"{drKB("MUSTFILL")}"
        'USELATESTVER : 以最近一次版本送件
        Dim vUSELATESTVER As String = $"{drKB("USELATESTVER")}"
        '檔案上傳:UPLOADFL1
        Dim vUPLOADFL1 As String = $"{drKB("UPLOADFL1")}"
        'DOWNLOADRPT '可下載報表
        Dim vDOWNLOADRPT As String = $"{drKB("DOWNLOADRPT")}"
        Dim vDOWNLOADRPT2 As String = $"{drKB("DOWNLOADRPT2")}"
        '(報表名稱)
        Dim vRPTNAME As String = $"{drKB("RPTNAME")}"
        Dim vRPTNAME2 As String = $"{drKB("RPTNAME2")}"
        '以目前版本批次送出:SENTBATVER
        Dim vSENTBATVER As String = $"{drKB("SENTBATVER")}"
        '以目前版本送出: SENDCURRVER
        Dim vSENDCURRVER As String = $"{drKB("SENDCURRVER")}"
        '備註說明:USEMEMO1
        Dim vUSEMEMO1 As String = $"{drKB("USEMEMO1")}"
        Dim vUPLOADSIZE As String = $"{drKB("UPLOADSIZE")}"

        '取得文字說明
        LiteralSwitchTo.Text = If(vKTSEQ <> "", TIMS.HtmlDecode1(vKTDESC1), "(無)")
        tr_LiteralSwitchTo.Visible = If(vNOTKBDESC1 = "Y", False, True) '(Y:不使用KBDESC)
        '檔案格式說明
        tr_FILEDESC1.Visible = If(vNOTFLDESC1 = "Y", False, True) '(Y:不使用FLDESC1)
        '(使用)'(報表名稱)'DOWNLOADRPT '可下載報表
        BTN_DOWNLOADRPT1.Text = If(vRPTNAME <> "", vRPTNAME, "下載報表")
        tr_DOWNLOADRPT1.Visible = If(vDOWNLOADRPT = "Y", True, False)
        BTN_DOWNLOADRPT2.Text = If(vRPTNAME2 <> "", vRPTNAME2, "下載報表2")
        BTN_DOWNLOADRPT2.Visible = If(vDOWNLOADRPT2 = "Y", True, False)
        '以目前版本批次送出:SENTBATVER
        tr_SENTBATVER.Visible = If(vSENTBATVER = "Y", True, False)
        '以目前版本送出: SENDCURRVER
        tr_SENDCURRVER.Visible = If(vSENDCURRVER = "Y", True, False)

        '檔案格式說明-預設10MB
        labFILEDESC1.Text = cst_FileDescMsg_7_10M
        Dim str_rtn_checkFile1 As String = $"return checkFile1({cst_PostedFile_MAX_SIZE_10M});"
        If vUPLOADSIZE = "20" Then
            labFILEDESC1.Text = cst_FileDescMsg_7_20M '特20MB
            str_rtn_checkFile1 = $"return checkFile1({cst_PostedFile_MAX_SIZE_20M});"
        End If
        '檔案上傳:UPLOADFL1
        tr_UPLOADFL1.Visible = If(vUPLOADFL1 = "Y", True, False)
        ' return checkFile1(sizeLimit);
        But1.Attributes.Remove("onclick")
        If vUPLOADFL1 = "Y" AndAlso str_rtn_checkFile1 <> "" Then But1.Attributes.Add("onclick", str_rtn_checkFile1)

        '方式 1：使用 Attributes 集合（最通用於 HtmlControl）
        td_USEMEMO1.Attributes("class") = "bluecol"
        tr_DataGrid06.Visible = If($"{drKB("DataGrid06")}" = "Y", True, False)
        If tr_DataGrid06.Visible Then SHOW_DATAGRID_06()

        '代號／非流水號
        Hid_KTID.Value = vKTID
        LabSwitchTo.Text = If(vKTSEQ <> "", TIMS.GetListText(ddlSwitchTo), "")
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

        '預設值-(上傳檔案) Hid_TBCFID.Value = ""
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

        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then Return
        Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, Hid_TBCID.Value, vKTSEQ)
        If drFL Is Nothing Then Return
        Hid_TBCFID.Value = $"{drFL("TBCFID")}"
        '免附文件
        CHKB_WAIVED.Checked = If($"{drFL("WAIVED")}" = "Y", True, False)
        txtMEMO1.Text = TIMS.ClearSQM(drFL("MEMO1"))

        '修改狀態，且為退件修正
        Dim fg_UPDATE_TBCSTATUS_R As Boolean = (Session(cst_ss_RqProcessType) = cst_DG1CMDNM_EDIT1 AndAlso Hid_TBCFID.Value <> "" AndAlso $"{drOA("TBCSTATUS")}" = "R")
        If fg_UPDATE_TBCSTATUS_R Then
            '沒有退件原因就不能改
            Dim fg_LOCK_INPUT As Boolean = ($"{drFL("RTUREASON")}" = "")
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

    ''' <summary>回上一步</summary>
    Private Sub MOVE_PREV()
        If (Hid_KTID.Value = "" OrElse Hid_KTID.Value = "01" OrElse ddlSwitchTo.SelectedIndex - 1 = -1) Then
            Common.MessageBox(Me, "(目前沒有上一步)")
            Return
        End If

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_KTSEQ.Value = ddlSwitchTo.Items(ddlSwitchTo.SelectedIndex - 1).Value
        If Hid_KTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_FirstKTSEQ.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    Protected Sub BTN_PREV1_Click(sender As Object, e As EventArgs) Handles BTN_PREV1.Click
        Call MOVE_PREV()
    End Sub
    Protected Sub BTN_SAVETMP1_Click(sender As Object, e As EventArgs) Handles BTN_SAVETMP1.Click
        '儲存(暫存)
        Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVETMP1)
    End Sub
    Protected Sub BTN_SAVENEXT1_Click(sender As Object, e As EventArgs) Handles BTN_SAVENEXT1.Click
        '儲存(暫存)
        Call SAVEDATA2_BTN_ACTION1(cst_ACTTYPE_BTN_SAVENEXT1)
    End Sub
    Protected Sub BTN_BACK1_Click(sender As Object, e As EventArgs) Handles BTN_BACK1.Click
        '清理隱藏的參數
        Call ClearHidValue()
        Call SHOW_Frame1(0)
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                '序號>,TBCASENO,案件編號>,APPSTAGE_N,申請階段>,ORGNAME,訓練機構>,CLASSCNAME2,班級名稱>,TBCNAME,申辦人姓名>
                ',TBCDATE_ROC,申辦日期>,TBCSTATUS_N,申辦狀態>,APPLIEDRESULT_N,審查狀態>,功能>
                ',lBTN_DELETE1,刪除,DELETE1,lBTN_VIEW1,查看,VIEW1,lBTN_EDIT1,修改,EDIT1,lBTN_SENDOUT1,送出,SENDOUT1,
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim lBTN_DELETE1 As LinkButton = e.Item.FindControl("lBTN_DELETE1") '(刪除)
                Dim lBTN_VIEW1 As LinkButton = e.Item.FindControl("lBTN_VIEW1") '查看
                Dim lBTN_EDIT1 As LinkButton = e.Item.FindControl("lBTN_EDIT1") '修改 
                Dim lBTN_SENDOUT1 As LinkButton = e.Item.FindControl("lBTN_SENDOUT1") '送出 

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TBCID", drv("TBCID"))
                TIMS.SetMyValue(sCmdArg, "TBCASENO", drv("TBCASENO"))
                TIMS.SetMyValue(sCmdArg, "TBCSTATUS", drv("TBCSTATUS"))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", drv("ORGKINDGW"))

                Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                lBTN_DELETE1.Visible = If(flagS1, True, False)
                lBTN_DELETE1.Style.Item("display") = If(flagS1, "", "none")
                lBTN_DELETE1.CommandArgument = sCmdArg
                Dim vMsgB As String = "請注意：此筆線上申辦案件原已上傳之相關文件均會一併刪除，確定要刪除此筆資料嗎?"
                lBTN_DELETE1.Attributes("onclick") = $"javascript:return confirm('{vMsgB}');"

                lBTN_VIEW1.CommandArgument = sCmdArg
                lBTN_EDIT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.CommandArgument = sCmdArg
                lBTN_SENDOUT1.Attributes("onclick") = "javascript:return confirm('此動作會送出資料審核不可再次修改，是否確定?');"

                If $"{drv("TBCSTATUS")}" = "R" Then
                    lBTN_EDIT1.Enabled = True
                    TIMS.Tooltip(lBTN_EDIT1, cst_tpmsg_enb5, True)
                    lBTN_SENDOUT1.Enabled = True
                    TIMS.Tooltip(lBTN_SENDOUT1, cst_tpmsg_enb5, True)
                Else
                    lBTN_EDIT1.Enabled = If($"{drv("TBCSTATUS")}" <> "", False, True)
                    TIMS.Tooltip(lBTN_EDIT1, If(lBTN_EDIT1.Enabled, "", cst_tpmsg_enb4), True)
                    lBTN_SENDOUT1.Enabled = If($"{drv("TBCSTATUS")}" <> "", False, True)
                    TIMS.Tooltip(lBTN_SENDOUT1, If(lBTN_SENDOUT1.Enabled, "", cst_tpmsg_enb4), True)
                End If

        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vTBCID As String = TIMS.GetMyValue(sCmdArg, "TBCID")
        Dim vTBCASENO As String = TIMS.GetMyValue(sCmdArg, "TBCASENO")
        Dim vTBCSTATUS As String = TIMS.GetMyValue(sCmdArg, "TBCSTATUS")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        If sCmdArg = "" OrElse vTBCID = "" OrElse vTBCASENO = "" Then Return

        Hid_TBCID.Value = vTBCID
        Hid_TBCASENO.Value = vTBCASENO
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無線上申辦案件)!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)!")
            Return
        End If
        'Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, Hid_TBCID.Value, vKTSEQ)
        'If drFL Is Nothing Then Return
        Select Case e.CommandName
            Case cst_DG1CMDNM_DELETE1 'DELETE1 (刪除)
                Call DELETE_Detail_STD14OA(Me, objconn, drOA)
                Common.MessageBox(Me, TIMS.cst_DELETEOKMsg2)
                Call SEARCH_1()

            Case cst_DG1CMDNM_VIEW1 '"VIEW1 '查看
                Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
                Call SHOW_Detail_STD14OA(drCC, Hid_TBCID.Value, cst_DG1CMDNM_VIEW1)

            Case cst_DG1CMDNM_EDIT1 '"EDIT1 '修改
                Call SHOW_Detail_STD14OA(drCC, Hid_TBCID.Value, cst_DG1CMDNM_EDIT1)

            Case cst_DG1CMDNM_SENDOUT1 'SENDOUT1 送出 
                '線上申辦進度 計算完成度百分比 (0-100) 
                Dim iProgress As Integer = TIMS.GET_iPROGRESS_OA(sm, objconn, tmpMSG, vTBCID, vORGKINDGW)
                '(比對異動日期)
                Dim s_RESULTDATE_YMS2 As String = If($"{drOA("RESULTDATE")}" <> "", CDate(drOA("RESULTDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
                Dim s_MODIFYDATE_YMS2 As String = If($"{drOA("MODIFYDATE")}" <> "", CDate(drOA("MODIFYDATE")).ToString("yyyy/MM/dd HH:mm:ss"), "")
                Dim fg_RESULTDATE_UPDATE As Boolean = (s_RESULTDATE_YMS2 <> "" AndAlso s_MODIFYDATE_YMS2 <> "" AndAlso DateDiff(DateInterval.Second, CDate(s_RESULTDATE_YMS2), CDate(s_MODIFYDATE_YMS2)) > 0)
                Dim EMSG As String = ""
                If iProgress < 100 Then
                    EMSG = $"線上申辦進度 未達100%，不可送出!{If(tmpMSG <> "", $"{vbCrLf}請檢查：({tmpMSG})", "")}.{iProgress}."
                    Common.MessageBox(Me, EMSG)
                    Return
                ElseIf $"{drOA("TBCSTATUS")}" = "R" AndAlso Not fg_RESULTDATE_UPDATE Then
                    EMSG = cst_tpmsg_enb5
                    Common.MessageBox(Me, EMSG)
                    Return
                End If

                Dim uParms As New Hashtable From {
                    {"TBCACCT", sm.UserInfo.UserID},
                    {"TBCSTATUS", "B"},
                    {"MODIFYACCT", sm.UserInfo.UserID},
                    {"TBCID", vTBCID},
                    {"TBCASENO", vTBCASENO}
                }
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_STD14OA" & vbCrLf
                usSql &= " SET TBCACCT=@TBCACCT,TBCDATE=GETDATE(),TBCSTATUS=@TBCSTATUS" & vbCrLf
                usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE TBCID=@TBCID and TBCASENO=@TBCASENO" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
#Region "SendMail_STD14OA"
                Dim drOA_U As DataRow = TIMS.GET_CLASS_STD14OA(objconn, vTBCID, vTBCASENO)
                If drOA_U Is Nothing Then
                    Common.MessageBox(Me, "送出資訊有誤(查無案件編號)，請重新操作!!")
                    Return
                End If
                Dim QPMS1 As New Hashtable From {
                    {"TBCID", vTBCID},
                    {"TBCASENO", vTBCASENO},
                    {"TBCDATE", TIMS.Cdate3t(drOA_U("TBCDATE"))},
                    {"OCID", $"{drOA_U("OCID")}"},
                    {"DISTID", $"{drCC("DISTID")}"}, 'sm.UserInfo.DistID)
                    {"DateLongStr", Now.ToString()}
                }
                Call TIMS.SendMail_STD14OA(objconn, QPMS1)
#End Region
                Call SEARCH_1()

        End Select
    End Sub

    Private Sub DELETE_Detail_STD14OA(MyPage As Page, oConn As SqlConnection, drOA As DataRow)
        If drOA Is Nothing Then Return
        Dim vTBCID As String = $"{drOA("TBCID")}"
        Dim vTBCASENO As String = $"{drOA("TBCASENO")}"
        Dim vAPPSTAGE As String = $"{drOA("APPSTAGE")}"
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then Return
        Dim vORGKINDGW As String = $"{drCC("ORGKINDGW")}"
        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vPLANID As String = $"{drCC("PLANID")}"

        Dim dtFL As DataTable = GET_CLASS_STD14OAFL_TB(objconn, vTBCID, vORGKINDGW)
        If TIMS.dtHaveDATA(dtFL) Then
            For Each drFL As DataRow In dtFL.Rows
                Dim iTBCFID As Integer = TIMS.CINT1(drFL("TBCFID"))
                Dim vFILENAME1 As String = $"{drFL("FILENAME1")}"
                Dim vFILEPATH1 As String = $"{drFL("FILEPATH1")}"
                If vFILENAME1 = "" Then Continue For
                '刪除檔案
                Dim oFILENAME1 As String = ""
                Dim oUploadPath As String = ""
                Dim s_FilePath1 As String = ""
                Try
                    oFILENAME1 = vFILENAME1
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, ""))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try

                'DELETE CLASS_STD14OAFL
                Dim dPMS As New Hashtable From {{"TBCID", TIMS.CINT1(vTBCID)}, {"TBCFID", iTBCFID}}
                Dim rdSql As String = "DELETE CLASS_STD14OAFL WHERE TBCID=@TBCID AND TBCFID=@TBCFID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dPMS)
            Next
        End If

        Dim dsSql As String = ""
        Dim dParms As New Hashtable From {{"TBCID", TIMS.CINT1(vTBCID)}}
        dsSql = "DELETE CLASS_STD14OAFL WHERE TBCID=@TBCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
        dsSql = "DELETE CLASS_STD14OA WHERE TBCID=@TBCID" & vbCrLf
        DbAccess.ExecuteNonQuery(dsSql, oConn, dParms)
    End Sub

    Protected Sub ddlSwitchTo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlSwitchTo.SelectedIndexChanged
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_KTSEQ.Value = TIMS.GetListValue(ddlSwitchTo)
        If Hid_KTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        ElseIf Hid_FirstKTSEQ.Value <> "" Then
            Call SHOW_STD14TH_KTSEQ(Hid_FirstKTSEQ.Value, Hid_ORGKINDGW.Value)
        End If
    End Sub

    ''' <summary>
    ''' 確定檔案上傳
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But1_Click(sender As Object, e As EventArgs) Handles But1.Click
        'Dim vUploadPath As String=Now.ToString("yyyyMMddHHmmss")
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_KTSEQ.Value = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        If Hid_TBCASENO.Value = "" OrElse Hid_TBCID.Value = "" Then
            Common.MessageBox(Me, "上傳資訊有誤(案件號為空)，請重新操作!!")
            Return
        End If
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Dim vKTID As String = $"{drKB("KTID")}"
        Dim vMULTIUPLOAD As String = $"{drKB("MULTIUPLOAD")}"
        Dim vTBCID As String = TIMS.ClearSQM(Hid_TBCID.Value)
        Dim vKTSEQ As String = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, vTBCID, vKTSEQ)
        '(退件修正)有退件原因,可重新上傳
        'Dim flag_NG_UPLOAD_1 As Boolean=(drFL IsNot Nothing) '(有資料 不可再次傳送)
        Dim flag_NG_UPLOAD_2 As Boolean = (drFL IsNot Nothing AndAlso $"{drFL("RTUREASON")}" = "") '(有資料不可傳送且原因為空 不可再次傳送)
        Dim vFILENAME1 As String = If(drFL IsNot Nothing, $"{drFL("FILENAME1")}", "")
        'Dim vWAIVED As String=If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
        If vMULTIUPLOAD = "" AndAlso vFILENAME1 <> "" AndAlso flag_NG_UPLOAD_2 Then
            '符合所有 不可再次傳送 'cst_tpmsg_enb8
            Common.MessageBox(Me, "已上傳儲存過該文件，不可再次操作!")
            Return
        End If

        Select Case $"{vORGKINDGW}{vKTID}"
            Case G06_其他補充資料, W06_其他補充資料
                txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
                If txtMEMO1.Text = "" Then
                    Common.MessageBox(Me, "備註說明不可為空，請輸入備註說明!")
                    Return
                End If

                Call FILE_UPLOAD_06(drOA, drKB)
                Call SHOW_DATAGRID_06()
            Case Else
                '有錯誤原因 可再次傳送 並記錄 iBCFID
                Dim iTBCFID As Integer = If(drFL IsNot Nothing AndAlso $"{drFL("RTUREASON")}" <> "", TIMS.CINT1(drFL("TBCFID")), -1)
                '檔案上傳／確定檔案上傳
                Call FILE_UPLOAD_1(drOA, iTBCFID)
        End Select

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)
    End Sub

    Private Sub FILE_UPLOAD_06(drOA As DataRow, drKB As DataRow)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If
        If drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無職類/班別代碼)，請重新操作!")
            Return
        End If
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
        If LCase(MyFileType) <> "pdf" Then
            Common.MessageBox(Me, cst_errMsg_5)
            Exit Sub
        End If
        'Case "pdf"
        If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
            Common.MessageBox(Me, cst_errMsg_7_10M)
            Exit Sub
        End If
        'Case "pdf"
        If Not TIMS.IsFileTypeValid(MyPostedFile, "pdf") Then
            Common.MessageBox(Me, cst_errMsg_5b)
            Exit Sub
        End If

        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text
        Dim vTBCID As String = TIMS.ClearSQM(Hid_TBCID.Value)
        Dim vKTSEQ As String = TIMS.ClearSQM(Hid_KTSEQ.Value)

        Dim vKTID As String = $"{drKB("KTID")}"
        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vPLANID As String = $"{drCC("PLANID")}"
        Dim vTBCASENO As String = $"{drOA("TBCASENO")}"
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, vKTSEQ)
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_S1_OTH(vTBCID, vKTSEQ, "", "pdf")
        Dim vSRCFILENAME1 As String = MyFileName

        Dim vTBCFID As Integer = -1
        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "TBCID", vTBCID)
            TIMS.SetMyValue2(rPMS2, "KTSEQ", vKTSEQ)
            TIMS.SetMyValue2(rPMS2, "WAIVED", cst_06_其他補充資料_WAIVED_OTH1)
            'TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            vTBCFID = SAVE_CLASS_STD14OAFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = $"ex.Message:{ex.Message}{vbCrLf},ex.ToString:{ex.ToString}"
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try

        '上傳檔案/存檔：檔名
        Try
            '上傳檔案
            TIMS.MyFileSaveAs(Me, File1, vUploadPath, vFILENAME1)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, cst_errMsg_2)

            Dim strErrmsg As String = $"ex.Message:{ex.Message}{vbCrLf}"
            strErrmsg &= String.Concat("vUploadPath: ", vUploadPath, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.FileName: ", MyPostedFile.FileName, vbCrLf)
            strErrmsg &= String.Concat("vFILENAME1: ", vFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("vSRCFILENAME1(MyFileName): ", vSRCFILENAME1, vbCrLf)
            strErrmsg &= String.Concat("MyPostedFile.ContentType: ", MyPostedFile.ContentType, vbCrLf)
            strErrmsg &= String.Concat("Server.MapPath(vUploadPath, vFILENAME1): ", Server.MapPath($"{vUploadPath}{vFILENAME1}"), vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Dim rPMS As New Hashtable From {
            {"TBCFID", vTBCFID},
            {"TBCID", vTBCID},
            {"KTSEQ", vKTSEQ},
            {"FILENAME1", vFILENAME1},
            {"SRCFILENAME1", vSRCFILENAME1},
            {"FILEPATH1", vUploadPath},
            {"MEMO1", vMEMO1},
            {"MODIFYACCT", sm.UserInfo.UserID}
        }
        Call SAVE_CLASS_STD14OAFL_OTH(rPMS)
    End Sub

    Private Sub SAVE_CLASS_STD14OAFL_OTH(rPMS As Hashtable)
        'drOA As DataRow, drKB As DataRow, drCC As DataRow, 
        Dim vTBCFID As String = TIMS.GetMyValue2(rPMS, "TBCFID")
        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim vKTSEQ As String = TIMS.GetMyValue2(rPMS, "KTSEQ")
        Dim vFILENAME1 As String = TIMS.GetMyValue2(rPMS, "FILENAME1")
        Dim vSRCFILENAME1 As String = TIMS.GetMyValue2(rPMS, "SRCFILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue2(rPMS, "FILEPATH1")
        Dim vMEMO1 As String = TIMS.GetMyValue2(rPMS, "MEMO1")
        'Dim vWAIVED As String = TIMS.GetMyValue2(rPMS, "WAIVED")
        Dim vMODIFYACCT As String = TIMS.GetMyValue2(rPMS, "MODIFYACCT")

        Dim iCS14OFID As Integer = DbAccess.GetNewId(objconn, "CLASS_STD14OAFL_OTH_CS14OFID_SEQ,CLASS_STD14OAFL_OTH,CS14OFID")
        Dim iParmsT As New Hashtable From {
            {"CS14OFID", iCS14OFID},
            {"TBCFID", TIMS.CINT1(vTBCFID)},
            {"TBCID", TIMS.CINT1(vTBCID)},
            {"KTSEQ", TIMS.CINT1(vKTSEQ)},
            {"FILENAME1", vFILENAME1},
            {"SRCFILENAME1", vSRCFILENAME1},
            {"FILEPATH1", vFILEPATH1},
            {"MEMO1", If(vMEMO1 <> "", vMEMO1, Convert.DBNull)},
            {"MODIFYACCT", vMODIFYACCT}
        }
        Dim isSqlT As String = "
INSERT INTO CLASS_STD14OAFL_OTH(CS14OFID,TBCFID,TBCID,KTSEQ,FILENAME1,SRCFILENAME1,FILEPATH1,MEMO1,MODIFYACCT,MODIFYDATE)
VALUES (@CS14OFID,@TBCFID,@TBCID,@KTSEQ,@FILENAME1,@SRCFILENAME1,@FILEPATH1,@MEMO1,@MODIFYACCT,GETDATE())
"
        DbAccess.ExecuteNonQuery(isSqlT, objconn, iParmsT)
    End Sub

    Private Sub SHOW_DATAGRID_06()
        '方式 1：使用 Attributes 集合（最通用於 HtmlControl）
        td_USEMEMO1.Attributes("class") = "bluecol_need"
        iDG06_ROWS = 0
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_KTSEQ.Value = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Dim vTBCID As String = Hid_TBCID.Value
        Dim vKTSEQ As String = Hid_KTSEQ.Value 'TIMS.GetListValue(ddlSwitchTo)
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        If drOA Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            'Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            'Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)!")
            Return
        End If

        'GET_CLASS_STD14OAFL_OTH'TBCFID,TBCID,KTSEQ
        Dim sParms1 As New Hashtable From {{"TBCID", TIMS.CINT1(vTBCID)}, {"KTSEQ", TIMS.CINT1(vKTSEQ)}}
        Dim SQL_ORDER1 As String = "ORDER BY a.MODIFYDATE"
        Dim sSql1 As String = $"
SELECT ROW_NUMBER() OVER ({SQL_ORDER1}) ROWNUM1
,a.CS14OFID,a.TBCFID,a.TBCID,a.KTSEQ,a.PATTERN,a.MEMO1,a.WAIVED,a.MODIFYACCT,a.MODIFYDATE
,b.KTID,c.OCID,c.TBCSTATUS,f.RTUREASON,a.FILEPATH1,a.SRCFILENAME1,a.FILENAME1,a.FILENAME1 OKFLAG
FROM CLASS_STD14OAFL_OTH a
JOIN KEY_STD14TH b on b.KTSEQ=a.KTSEQ
JOIN CLASS_STD14OA c on c.TBCID=a.TBCID
JOIN CLASS_STD14OAFL f on f.TBCFID=a.TBCFID
WHERE a.TBCID=@TBCID AND a.KTSEQ=@KTSEQ
{SQL_ORDER1}
"
        Dim dt2 As DataTable = DbAccess.GetDataTable(sSql1, objconn, sParms1)

        labmsg2.Text = ""
        If TIMS.dtNODATA(dt2) Then
            DataGrid06.DataSource = Nothing
            DataGrid06.DataBind()
            labmsg2.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        iDG06_ROWS = dt2.Rows.Count

        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vPLANID As String = $"{drCC("PLANID")}"
        Dim vTBCASENO As String = $"{drOA("TBCASENO")}"
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, vKTSEQ)
        Call TIMS.Check_dtBIDCASEFL(Me, dt2, download_Path)
        With DataGrid06
            .DataSource = dt2
            .DataBind()
        End With
    End Sub

    Private Sub FILE_UPLOAD_1(drOA As Object, iTBCFID As Object)
        '(上傳路徑) 'If drOB Is Nothing Then Return
        If drOA Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無案件編號)，請重新操作!!")
            Return
        End If
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, Hid_KTSEQ.Value, Hid_ORGKINDGW.Value)
        If drKB Is Nothing Then
            Common.MessageBox(Me, "上傳資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無職類/班別代碼)，請重新操作!!")
            Return
        End If

        '取得KBID代號／非流水號
        Dim vKTSEQ As String = $"{drKB("KTSEQ")}"
        Dim vKTID As String = $"{drKB("KTID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Dim vUPLOADSIZE As String = $"{drKB("UPLOADSIZE")}"
        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vPLANID As String = $"{drCC("PLANID")}"

        'Hid_TECHID.Value=TIMS.ClearSQM(Hid_TECHID) 'Dim vTECHID As String=Hid_TECHID.Value
        txtMEMO1.Text = TIMS.ClearSQM(txtMEMO1.Text)
        Dim vMEMO1 As String = txtMEMO1.Text  'TIMS.GetMyValue2(rPMS, "MEMO1")

        Dim vTBCID As String = TIMS.ClearSQM(drOA("TBCID")) 'TIMS.GetMyValue2(rPMS, "BCID")
        Dim vTBCASENO As String = TIMS.ClearSQM(drOA("TBCASENO"))

        Dim vMODIFYACCT As String = sm.UserInfo.UserID 'TIMS.GetMyValue2(rPMS, "MODIFYACCT")
        If vTBCASENO <> Hid_TBCASENO.Value Then Return '(此狀況不太可能發生)
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
        If LCase(MyFileType) <> "pdf" Then
            Common.MessageBox(Me, cst_errMsg_5)
            Exit Sub
        End If
        If vUPLOADSIZE = "20" Then
            If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_20M Then
                Common.MessageBox(Me, cst_errMsg_7_20M)
                Exit Sub
            End If
        Else
            If File1.PostedFile.ContentLength > cst_PostedFile_MAX_SIZE_10M Then
                Common.MessageBox(Me, cst_errMsg_7_10M)
                Exit Sub
            End If
        End If
        If Not TIMS.IsFileTypeValid(MyPostedFile, "pdf") Then
            Common.MessageBox(Me, cst_errMsg_5b)
            Exit Sub
        End If

        '上傳檔案 '年度／計畫ID／機構ID／caseno／1
        Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, "")
        Dim vFILENAME1 As String = TIMS.GET_FILENAME1_S1(vTBCID, vKTSEQ, "", "pdf") 'String.Concat("B", TIMS.GetDateNo2(4), "x", vBCID, "x", vKBSID, vPATTERN, ".pdf")
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
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim rPMS2 As New Hashtable
            TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "TBCFID", If(vUploadPath <> "", iTBCFID, -1)) '(可再次傳送)

            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "TBCID", vTBCID)
            TIMS.SetMyValue2(rPMS2, "KTSEQ", vKTSEQ)
            'TIMS.SetMyValue2(rPMS2, "FILENAME1", FILENAME1)
            'TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", SRCFILENAME1)
            'TIMS.SetMyValue2(rPMS2, "PATTERN", PATTERN)
            TIMS.SetMyValue2(rPMS2, "WAIVED", "")
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", vMODIFYACCT)
            Call SAVE_CLASS_STD14OAFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim HFileName As HtmlInputHidden=e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vTBCFID As String = TIMS.GetMyValue(sCmdArg, "TBCFID")
        Dim vKTID As String = TIMS.GetMyValue(sCmdArg, "KTID")
        Dim vKTSEQ As String = TIMS.GetMyValue(sCmdArg, "KTSEQ")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Dim vORGKINDGW As String = Hid_ORGKINDGW.Value
        Dim vTBCID As String = Hid_TBCID.Value
        Dim vTBCASENO As String = Hid_TBCASENO.Value
        If e.CommandArgument = "" OrElse vTBCFID = "" Then Return

        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then Return

        Select Case e.CommandName
            Case "DELFILE4"
                'Common.MessageBox(Me, $"{vORGKINDGW},{vKTID},{vTBCFID}")'Return
                Dim sErrMsg1 As String = CHKDEL_CLASS_STD14OAFL(vORGKINDGW, vKTID, vTBCFID)
                If sErrMsg1 <> "" Then
                    Common.MessageBox(Me, sErrMsg1)
                    Return
                End If

                '刪除檔案
                Dim oTBCASENO As String = $"{drOA("TBCASENO")}"
                Dim oFILENAME1 As String = ""
                Dim oUploadPath As String = ""
                Dim s_FilePath1 As String = ""
                Try
                    Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
                    If drCC Is Nothing Then Return
                    Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
                    Dim vDISTID As String = $"{drCC("DISTID")}"
                    Dim vPLANID As String = $"{drCC("PLANID")}"
                    oFILENAME1 = vFILENAME1
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, ""))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try

                'DELETE CLASS_STD14OAFL
                Dim dParms As New Hashtable From {{"TBCFID", vTBCFID}}
                Dim rdSql As String = "DELETE CLASS_STD14OAFL WHERE TBCFID=@TBCFID"
                DbAccess.ExecuteNonQuery(rdSql, objconn, dParms)
                'DataGrid1.EditItemIndex=-1
                Select Case $"{vORGKINDGW}{vKTID}"
                    Case G06_其他補充資料, W06_其他補充資料
                        'DELETE CLASS_STD14OAFL_OTH
                        Dim dParms2 As New Hashtable From {{"TBCFID", vTBCFID}, {"TBCID", vTBCID}, {"KTSEQ", vKTSEQ}}
                        Dim rdSql2 As String = "DELETE CLASS_STD14OAFL_OTH WHERE TBCFID=@TBCFID AND TBCID=@TBCID AND KTSEQ=@KTSEQ"
                        DbAccess.ExecuteNonQuery(rdSql2, objconn, dParms2)
                End Select

            Case "DOWNLOAD4" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "TBCID", Hid_TBCID.Value)
                TIMS.SetMyValue2(rPMS4, "TBCASENO", Hid_TBCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "TBCFID", vTBCFID)
                TIMS.SetMyValue2(rPMS4, "KTID", vKTID)
                TIMS.SetMyValue2(rPMS4, "KTSEQ", vKTSEQ)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)
                Call TIMS.ResponseZIPFile_OA(sm, objconn, Me, rPMS4)
                Return
        End Select

        If Not TIMS.OpenDbConn(objconn) Then Return

        '顯示檔案資料表
        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)

        Dim drCC2 As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        Call SHOW_Detail_STD14OA(drCC2, Hid_TBCID.Value, Session(cst_ss_RqProcessType))
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
                labRTUREASON.Text = $"{drv("RTUREASON")}"
                Dim titleMsg As String = ""
                If Not IsDBNull(drv("FILENAME1")) Then
                    titleMsg = $"{drv("OKFLAG")}"
                    BTN_DOWNLOAD4.Enabled = ($"{drv("FILENAME1")}" = $"{drv("OKFLAG")}")
                ElseIf $"{drv("WAIVED")}" = cst_06_其他補充資料_WAIVED_OTH1 Then
                    titleMsg = cst_txt_其他補充資料
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TBCFID", $"{drv("TBCFID")}")
                TIMS.SetMyValue(sCmdArg, "KTID", $"{drv("KTID")}")
                TIMS.SetMyValue(sCmdArg, "KTSEQ", $"{drv("KTSEQ")}")
                TIMS.SetMyValue(sCmdArg, "FILENAME1", $"{drv("FILENAME1")}")
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", $"{drv("FILEPATH1")}")
                BTN_DELFILE4.CommandArgument = sCmdArg '刪除
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                BTN_DELFILE4.Attributes("onclick") = TIMS.cst_confirm_delmsg1
                '檢視不能修改
                BTN_DELFILE4.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)

                '(其他原因調整) '送件／退件修正，不提供刪除
                If $"{drv("TBCSTATUS")}" = "B" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb6, True)

                ElseIf $"{drv("TBCSTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" <> "" Then
                    BTN_DELFILE4.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb8, True)

                ElseIf $"{drv("TBCSTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" = "" Then
                    BTN_DELFILE4.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE4, cst_tpmsg_enb7, True)

                End If
        End Select
    End Sub

    Protected Sub BTN_DOWNLOADRPT1_Click(sender As Object, e As EventArgs) Handles BTN_DOWNLOADRPT1.Click, BTN_DOWNLOADRPT2.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_KTSEQ.Value = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Dim vKTSEQ As String = TIMS.GetListValue(ddlSwitchTo)
        If Hid_TBCASENO.Value = "" OrElse Hid_TBCID.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(案件號為空)，請重新操作!!")
            Return
        ElseIf Hid_KTSEQ.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_ORGKINDGW.Value = "" Then
            Common.MessageBox(Me, "下載報表資訊有誤(計畫代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_KTSEQ.Value <> "" AndAlso Hid_KTSEQ.Value <> vKTSEQ Then
            Common.MessageBox(Me, "下載報表資訊有誤(項目代碼／序號有誤)，請重新操作!!")
            Return
        End If

        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, vKTSEQ, Hid_ORGKINDGW.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        '首頁>>訓練機構管理>>表單列印>>訓練單位基本資料表 '訓練單位基本資料
        'https://ojrept.wda.gov.tw/ReportServer3/report.do?RptID=SD_14_001_18G&Years=112&RSID=47877&planid=5093&rid=%27E6762%27&AppStage=1&UserID=snoopy
        '列印
        Select Case TIMS.GET_ButtonID1(sender)
            Case "BTN_DOWNLOADRPT1"
                Call UTL_PRINT1GW(drOA, drKB, 1, False)
            Case "BTN_DOWNLOADRPT2"
                Call UTL_PRINT1GW(drOA, drKB, 2, False)
            Case Else
                Common.MessageBox(Me, "下載報表資訊有誤(ButtonID異常)，請重新操作!!")
                Return
        End Select
    End Sub

    Private Sub UTL_PRINT1GW(drOA As DataRow, drKB As DataRow, iNum As Integer, isSENDCURRVER As Boolean)
        If drOA Is Nothing OrElse drKB Is Nothing Then Return
        Dim NUMS As Integer() = {1, 2}
        If Not NUMS.Contains(iNum) Then Return
        'SELECT CONCAT('const ',ORGKINDGW,KTID ,'_',KTNAME,' as string="', ORGKINDGW,KTID ,'"') KTNAME2 FROM KEY_STD14TH ORDER BY ORGKINDGW,KTID

        '取得KBID代號／非流水號
        Dim vKTID As String = $"{drKB("KTID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Dim vKBNAME2 As String = $"{vORGKINDGW}{vKTID}{drKB("KTNAME")}"
        Dim rPMS As New Hashtable
        Select Case $"{vORGKINDGW}{vKTID}"
            Case G01_參訓學員簽訂之契約書, W01_參訓學員簽訂之契約書
                '辦理本計畫訓練課程之專職人員名冊 'rPMS.Clear()
                rPMS.Add("INUM", iNum)
                rPMS.Add("ORGKINDGW", vORGKINDGW)
                rPMS.Add("OCID", $"{drOA("OCID")}")
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                Call RPT_SD_14_021(rPMS)

            Case G02_參訓學員身分證影本與存摺影本黏貼表, W02_參訓學員身分證影本與存摺影本黏貼表
                rPMS.Add("INUM", iNum)
                rPMS.Add("ORGKINDGW", vORGKINDGW)
                rPMS.Add("OCID", $"{drOA("OCID")}")
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                Call RPT_SD_14_033(rPMS)

            Case G03_訓練課程開班學員名冊, W03_訓練課程開班學員名冊
                rPMS.Add("INUM", iNum)
                rPMS.Add("ORGKINDGW", vORGKINDGW)
                rPMS.Add("OCID", $"{drOA("OCID")}")
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                Call RPT_SD_14_008(rPMS)

            Case G04_預估參訓學員補助經費清冊, W04_預估參訓學員補助經費清冊
                rPMS.Add("INUM", iNum)
                rPMS.Add("ORGKINDGW", vORGKINDGW)
                rPMS.Add("OCID", $"{drOA("OCID")}")
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                Call RPT_SD_14_011(rPMS)

            Case G05_參訓學員投保狀況檢核表, W05_參訓學員投保狀況檢核表
                rPMS.Add("INUM", iNum)
                rPMS.Add("ORGKINDGW", vORGKINDGW)
                rPMS.Add("OCID", $"{drOA("OCID")}")
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                Call RPT_SD_01_007_R(rPMS)

        End Select
    End Sub

    ''' <summary>參訓學員投保狀況檢核表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_01_007_R(rPMS As Hashtable)
        '3.產業人才投資方案 [28] '4:充電起飛計畫（非在職／是產投）[54]--參訓學員投保狀況檢核表 'SD_01_007_R2*_b.jrxml
        Const cst_printFN2b As String = "SD_01_007_R2_b" 'STUD_BLIGATEDATA28[署(局)、分署(中心)]
        Const cst_printFN2bb As String = "SD_01_007_R2b_b" 'STUD_BLIGATEDATA28[委訓單位]
        '06:在職進修訓練 66:主題產業職業訓練(在職) --參訓學員投保狀況檢核表 SD_01_007_R3_b*.jrxml
        'Const cst_printFN3b As String = "SD_01_007_R3_b" '開訓日 STUD_SELRESULTBLI
        'Const cst_printFN3bb As String = "SD_01_007_R3_bb" '開訓日 STUD_SELRESULTBLI
        '06:在職進修訓練 66:主題產業職業訓練(在職) --甄試學員投保狀況檢核表 --LpfBatchGet28es.vbproj SD_01_007_R4_b*.jrxml
        'Const cst_printFN4b As String = "SD_01_007_R4_b" '甄試日 STUD_BLIGATEDATA28E '署(局)、分署(中心)
        'Const cst_printFN4bb As String = "SD_01_007_R4_bb" '甄試日 STUD_BLIGATEDATA28E '委訓單位
        '70:區域產業據點職業訓練計畫(在職)
        'Const cst_printFN5b As String = "SD_01_007_R5_b" '甄試日 STUD_BLIGATEDATA28E '署(局)、分署(中心) SD_01_007_R5_b*.jrxml
        'Const cst_printFN5bb As String = "SD_01_007_R5_bb" '甄試日 STUD_BLIGATEDATA28E '委訓單位

        Const cst_Printtype_B1 As String = "B1" '"b.IDNO" 'B1
        Const cst_Printtype_S1 As String = "S1" '"cs.StudentID" 'S1
        'Const cst_Printtype3 As String = "c.IDNO" '"b.IDNO" '06 66
        'Const cst_Printtype4B As String = "a.IDNO" '06 66 /70

        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "請選擇有效班級資料!!")
            Exit Sub
        End If

        Dim MyValue As String = ""
        Dim V_PrintTypeValue As String = ""
        Dim prtFilename As String = ""
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then 'End If
        Select Case vINUM 'print_type.SelectedValue
            Case "1" '"&Printtype=b.IDNO" '依身分證號(排序)
                V_PrintTypeValue = cst_Printtype_B1 '"&Printtype=b.IDNO" '依身分證號(排序)
            Case "2" '"&Printtype=cs.StudentID" '依學號(排序)
                V_PrintTypeValue = cst_Printtype_S1 '"&Printtype=cs.StudentID" '依學號(排序)
        End Select
        'HidPrinttype.Value=cst_Printtype3 '"&Printtype=b.IDNO" '依身分證號(排序)
        '產投報表。'3.產業人才投資方案 [28] '4:充電起飛計畫（在職）[54] 
        Select Case $"{sm.UserInfo.LID}"
            Case "0", "1" '署(局)、分署(中心)
                prtFilename = cst_printFN2b '"SD_01_007_R2_b"
            Case Else '委訓單位
                prtFilename = cst_printFN2bb '"SD_01_007_R2b_b"
        End Select

        MyValue = ""
        TIMS.SetMyValue(MyValue, "OCID", $"{drCC("OCID")}")
        TIMS.SetMyValue(MyValue, "MSD", $"{drCC("MSD")}")
        TIMS.SetMyValue(MyValue, "Printtype", V_PrintTypeValue)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, MyValue)
    End Sub

    ''' <summary>參訓學員簽訂之契約書</summary>
    ''' <param name="rPMS"></param>
    Sub RPT_SD_14_021(rPMS As Hashtable)
        Const cst_printFN_G1 As String = "SD_14_021" '1:2017前
        Const cst_printFN_G1_2 As String = "SD_14_021_2page" '1:2017前 / SD_14_021
        'NEW 補助學員參訓契約書 參訓學員簽訂之契約書
        Const cst_printFN_G3 As String = "SD_14_021G3" '3:2018之後
        Const cst_printFN_W3 As String = "SD_14_021W3" '3:2018之後
        Const cst_printFN_G3_2 As String = "SD_14_021G3_2page" '3:2018之後 / SD_14_021G3
        Const cst_printFN_W3_2 As String = "SD_14_021W3_2page" '3:2018之後 / SD_14_021W3

        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")

        Dim prtFilename As String = ""
        If TIMS.Cst_TPlanID54AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If vINUM = "1" Then
                prtFilename = cst_printFN_G1
            ElseIf vINUM = "2" Then
                '20180723 加入 列印方式(可分兩頁印)
                prtFilename = cst_printFN_G1_2
            End If
        Else
            If vINUM = "1" Then
                prtFilename = If(vORGKINDGW = "G", cst_printFN_G3, If(vORGKINDGW = "W", cst_printFN_W3, "")) 'cst_printFN_G3 '3:2018
            ElseIf vINUM = "2" Then
                '20180723 加入 列印方式(可分兩頁印)
                prtFilename = If(vORGKINDGW = "G", cst_printFN_G3_2, If(vORGKINDGW = "W", cst_printFN_W3_2, ""))  'cst_printFN_W3_2 '3:2018
            End If
        End If

        If prtFilename = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim prtstr As String = $"&TPlanID={vTPlanID}&OCID={vOCID}"
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, prtFilename, prtstr)
    End Sub

    ''' <summary>參訓學員身分證影本與存摺影本黏貼表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_033(rPMS As Hashtable)
        Const cst_printFN_1 As String = "SD_14_033"
        'Const cst_printFN_2 As String = "SD_14_033_S"
        Const cst_printFN_3 As String = "SD_14_033_SS"
        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        'WHERE cs.TPlanID=#{TPlanID} AND cs.RID=#{RID} AND cs.OCID=#{OCID}
        Dim vMSD As String = $"{drCC("MSD")}"
        Dim vRID As String = $"{drCC("RID")}"
        Select Case vINUM
            Case "1"
                Dim vSOCID_S As String = GET_SD_14_033_STDVAL(vOCID, "SOCID") '取得有資料的學員
                Dim vSETID_S As String = GET_SD_14_033_STDVAL(vOCID, "SETID") '取得有資料的學員
                If vSOCID_S = "" OrElse vSETID_S = "" Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
                Dim vTSTPRINT As String = If($"{TIMS.Utl_GetConfigSet("printtest")}" = "Y", "2", "1") '正式區1／'測試區2
                Dim myValue1 As String = ""
                TIMS.SetMyValue(myValue1, "TPlanID", vTPlanID)
                TIMS.SetMyValue(myValue1, "RID", vRID)
                TIMS.SetMyValue(myValue1, "OCID", vOCID)
                TIMS.SetMyValue(myValue1, "SOCID", vSOCID_S)
                TIMS.SetMyValue(myValue1, "SETID", vSETID_S)
                TIMS.SetMyValue(myValue1, "TSTPRINT", vTSTPRINT)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN_3, myValue1)
            Case "2"
                Dim myValue1 As String = ""
                TIMS.SetMyValue(myValue1, "TPlanID", vTPlanID)
                TIMS.SetMyValue(myValue1, "RID", vRID)
                TIMS.SetMyValue(myValue1, "OCID", vOCID)
                TIMS.SetMyValue(myValue1, "MSD", vMSD)
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN_1, myValue1)
        End Select
    End Sub
    ''' <summary>'取得有資料的內容 </summary>
    ''' <param name="vOCID"></param>
    ''' <param name="RTNCOLUMN"></param>
    ''' <returns></returns>
    Private Function GET_SD_14_033_STDVAL(vOCID As String, RTNCOLUMN As String) As String
        Dim PMS1 As New Hashtable From {{"OCID", TIMS.CINT1(vOCID)}}
        Dim SQL1 As String = "
SELECT a.OCID,a.SOCID,a.SETID
FROM V_STUDENTINFO a
JOIN V_EIMG12 b on b.IDNO=a.IDNO
WHERE b.ISUSE='Y' AND b.ISDEL IS NULL AND a.OCID=@OCID"
        Dim dt As DataTable = DbAccess.GetDataTable(SQL1, objconn, PMS1)
        If TIMS.dtNODATA(dt) Then Return ""
        Return TIMS.ECHOVAL(dt, RTNCOLUMN)
    End Function

    ''' <summary>訓練課程開班學員名冊</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_008(rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_008_2009_b"
        Const cst_printFN2 As String = "SD_14_008_2025_c"

        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "未選擇職類/班別，請選擇班級!!")
            Exit Sub
        End If

        '(「職場續航」之課程勾稽投保年資)
        Dim gfg_WYROLE As String = TIMS.CHECK_WYROLE(objconn, vOCID) 'Dim gfg_WYROLE As Boolean = CHECK_WYROLE()
        Dim rptFILENAME As String = If(gfg_WYROLE, cst_printFN2, cst_printFN1)
        Dim V_MSD As String = $"{drCC("MSD")}"
        Dim V_PRINT_ORDERYBY As String = If(vINUM = "2", "a.StudentID", "c.IDNO")
        Dim vYEARS_ROC As String = $"{drCC("YEARS_ROC")}"  ' TIMS.GET_YEARS_ROC(Now.Year)

        Dim myvalue As String = ""
        TIMS.SetMyValue(myvalue, "MSD", V_MSD)
        TIMS.SetMyValue(myvalue, "Years", vYEARS_ROC) 'sm.UserInfo.Years - 1911
        TIMS.SetMyValue(myvalue, "OCID", drCC("OCID")) ' OCIDValue1.Value)
        TIMS.SetMyValue(myvalue, "Printtype", V_PRINT_ORDERYBY)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, rptFILENAME, myvalue)
    End Sub

    ''' <summary>訓練課程開班學員名冊</summary>
    ''' <param name="rPMS"></param>
    ''' <returns></returns>
    Function GET_RPTURL_SD_14_008(rPMS As Hashtable) As String
        Const cst_printFN1 As String = "SD_14_008_2009_b"
        Const cst_printFN2 As String = "SD_14_008_2025_c"

        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If drCC Is Nothing Then Return ""
        'If drCC Is Nothing Then,Common.MessageBox(Me, "未選擇職類/班別，請選擇班級!!"),Exit Sub,End If,
        '(「職場續航」之課程勾稽投保年資)
        Dim gfg_WYROLE As String = TIMS.CHECK_WYROLE(objconn, vOCID) 'Dim gfg_WYROLE As Boolean = CHECK_WYROLE()
        Dim rptFILENAME As String = If(gfg_WYROLE, cst_printFN2, cst_printFN1)
        Dim V_MSD As String = $"{drCC("MSD")}"
        Dim V_PRINT_ORDERYBY As String = If(vINUM = "2", "a.StudentID", "c.IDNO")
        Dim vYEARS_ROC As String = $"{drCC("YEARS_ROC")}"  ' TIMS.GET_YEARS_ROC(Now.Year)

        Dim myvalue As String = ""
        TIMS.SetMyValue(myvalue, "MSD", V_MSD)
        TIMS.SetMyValue(myvalue, "Years", vYEARS_ROC) 'sm.UserInfo.Years - 1911
        TIMS.SetMyValue(myvalue, "OCID", $"{drCC("OCID")}")
        TIMS.SetMyValue(myvalue, "Printtype", V_PRINT_ORDERYBY)
        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, rptFILENAME, myvalue)
        Return ReportQuery.GetReportUrl2(Me, rptFILENAME, myvalue)
    End Function

    Sub SAVE_SD_14_008(rPMS As Hashtable)
        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")

        Dim rPMS4 As New Hashtable From {{"INUM", vINUM}, {"ORGKINDGW", vORGKINDGW}, {"OCID", vOCID}, {"TPlanID", vTPlanID}} 'rPMS4.Clear()
        Dim s_RPTURL As String = GET_RPTURL_SD_14_008(rPMS4)
        Dim s_PDF_byte As Byte() = Nothing

        'Dim infomsg As String = $"##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte),{s_RPTURL}"
        'infomsg &= $",訓練課程開班學員名冊 s_RPTURL: {s_RPTURL}"
        'infomsg &= $", s_PDF_byte: {If(s_PDF_byte Is Nothing, "Is Nothing!", $"{s_PDF_byte.Length}")}"
        'infomsg &= $", rPMS4: {TIMS.GetMyValue4(rPMS4)}"
        'TIMS.LOG.Info(infomsg)

        Try
            Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
        Catch ex As Exception
            Dim eErrmsg As String = $"##TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte),{s_RPTURL}, ex.Message: {ex.Message}"
            eErrmsg &= $",訓練課程開班學員名冊 s_RPTURL: {s_RPTURL}"
            eErrmsg &= $", s_PDF_byte: {If(s_PDF_byte Is Nothing, "Is Nothing!", $"{s_PDF_byte.Length}")}"
            eErrmsg &= $", rPMS4: {TIMS.GetMyValue4(rPMS4)}"
            TIMS.LOG.Error(eErrmsg, ex)
            Common.MessageBox(Me, "訓練課程開班學員名冊 報表下載檔案有誤，請確認檔案是否正確!")
            Return
        End Try
        If s_PDF_byte Is Nothing Then Return

        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim vKTSEQ As String = TIMS.GetMyValue2(rPMS, "KTSEQ")
        Dim vFILENAME1 As String = TIMS.GetValidFileName(TIMS.GetMyValue2(rPMS, "FILENAME1"))
        Dim vSRCFILENAME1 As String = TIMS.GetValidFileName(TIMS.GetMyValue2(rPMS, "SRCFILENAME1"))
        Dim vUploadPath As String = TIMS.GetMyValue2(rPMS, "UploadPath")
        '上傳檔案/存檔：檔名
        Try
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
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
        End Try

        Try
            Dim rPMS2 As New Hashtable
            'TIMS.SetMyValue2(rPMS2, "UploadPath", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "TBCFID", If(vUploadPath <> "", iTBCFID, -1)) '(可再次傳送)
            TIMS.SetMyValue2(rPMS2, "ORGKINDGW", vORGKINDGW)
            TIMS.SetMyValue2(rPMS2, "TBCID", vTBCID)
            TIMS.SetMyValue2(rPMS2, "KTSEQ", vKTSEQ)
            TIMS.SetMyValue2(rPMS2, "WAIVED", "")
            TIMS.SetMyValue2(rPMS2, "FILENAME1", vFILENAME1)
            TIMS.SetMyValue2(rPMS2, "SRCFILENAME1", vSRCFILENAME1)
            TIMS.SetMyValue2(rPMS2, "FILEPATH1", vUploadPath)
            'TIMS.SetMyValue2(rPMS2, "MEMO1", vMEMO1)
            TIMS.SetMyValue2(rPMS2, "MODIFYACCT", sm.UserInfo.UserID)
            Call SAVE_CLASS_STD14OAFL_UPLOAD(rPMS2)
        Catch ex As Exception
            TIMS.LOG.Warn(ex.Message, ex)
            Common.MessageBox(Me, ex.ToString)

            Dim strErrmsg As String = String.Concat("ex.ToString:", ex.ToString, vbCrLf)
            TIMS.WriteTraceLog(Me, ex, strErrmsg)
            Exit Sub
            'Throw ex
        End Try

    End Sub

    ''' <summary>預估參訓學員補助經費清冊</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_011(rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_011_1_2020_b"
        Const cst_errMsg1 As String = "請完成[學員資料維護]的[學員資料確認]與[學員資料審核]，才可使用此列印功能!!!"

        Dim vINUM As String = TIMS.GetMyValue2(rPMS, "INUM")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vTPlanID As String = TIMS.GetMyValue2(rPMS, "TPlanID")
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "未選擇職類/班別，請選擇班級!!")
            Exit Sub
        End If

        Dim v_PRINT_ORDERYBY As String = If(vINUM = "2", "2", "1")  '2:by StudentID,'1:by IDNO
        Dim vYEARS_ROC As String = $"{drCC("YEARS_ROC")}"  ' TIMS.GET_YEARS_ROC(Now.Year)

        Dim fg_CANPRINT As Boolean = True '已完成學員資料審核 與 學員資料確認
        Dim R_PMS2 As New Hashtable From {{"RID", drCC("RID")}}
        Dim R_SQL2 As String = " SELECT TOP 11 RID,PLANID FROM VIEW2 WHERE RID=@RID AND ISNULL(AppliedResultR,'N')<>'Y' "
        Dim drR2 As DataRow = DbAccess.GetOneRow(R_SQL2, objconn, R_PMS2)

        If drR2 IsNot Nothing Then fg_CANPRINT = False '未完成學員資料審核 與 學員資料確認
        If fg_test Then fg_CANPRINT = True '(測試機跳過驗證)
        If Not fg_CANPRINT Then
            'Common.MessageBox(Me, $"{cst_errMsg1}.{drCC("RID")}")
            Common.MessageBox(Me, $"{cst_errMsg1}")
            Exit Sub
        End If

        '已完成學員資料審核 與 學員資料確認 'prtFilename="SD_14_011_1_2013_b"'vOCID
        Dim MyValue1 As String = $"RID={drCC("RID")}&PlanID={drCC("PLANID")}&OCID={drCC("OCID")}&Years={vYEARS_ROC}&Printtype={v_PRINT_ORDERYBY}"
        Select Case v_PRINT_ORDERYBY
            Case "1"
                MyValue1 &= "&Printtype1=Y"
            Case Else
                MyValue1 &= "&Printtype2=Y"
        End Select
        '已完成學員資料審核 與 學員資料確認 'prtFilename="SD_14_011_1_2013_b"
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)
    End Sub

    Sub UPDATE_HISREVIEW(drOA As DataRow, ddlAPPLIEDRESULT As DropDownList, RPMS1 As Hashtable)
        'Dim s_HIS As String = String.Concat(TIMS.cdate3t(Now), "-", TIMS.GetListText(ddlAPPLIEDRESULT))
        Dim v_Reasonforfail As String = TIMS.GetMyValue2(RPMS1, "Reasonforfail")
        Dim v_TBCID As String = TIMS.GetMyValue2(RPMS1, "TBCID")
        Dim v_TBCASENO As String = TIMS.GetMyValue2(RPMS1, "TBCASENO")
        If v_TBCID = "" OrElse v_TBCASENO = "" Then Return
        If drOA Is Nothing Then Return
        Dim v_ddlAPPLIEDRESULT As String = TIMS.GetListValue(ddlAPPLIEDRESULT)
        If v_ddlAPPLIEDRESULT = "" Then Return
        Dim s_HISREVIEW As String = $"{drOA("HISREVIEW")}"
        s_HISREVIEW &= $"{If(s_HISREVIEW <> "", "，", "")}{TIMS.Cdate3t(Now)}-{TIMS.GetListText(ddlAPPLIEDRESULT)}"

        Dim uParms As New Hashtable
        uParms.Add("HISREVIEW", s_HISREVIEW)
        uParms.Add("APPLIEDRESULT", v_ddlAPPLIEDRESULT)
        uParms.Add("REASONFORFAIL", If(v_Reasonforfail <> "", v_Reasonforfail, Convert.DBNull))
        uParms.Add("RESULTACCT", sm.UserInfo.UserID)
        'uParms.Add("RESULTDATE", RESULTDATE)
        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
        uParms.Add("TBCID", v_TBCID)
        uParms.Add("TBCASENO", v_TBCASENO)
        Dim usSql As String = ""
        usSql &= " UPDATE CLASS_STD14OA" & vbCrLf
        usSql &= " SET HISREVIEW=@HISREVIEW,APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        usSql &= " ,REASONFORFAIL=@REASONFORFAIL,RESULTACCT=@RESULTACCT,RESULTDATE=GETDATE()" & vbCrLf
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE TBCID=@TBCID AND TBCASENO=@TBCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
    End Sub

    Protected Sub BTN_SENDCURRVER_Click(sender As Object, e As EventArgs) Handles BTN_SENDCURRVER.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_KTSEQ.Value = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Dim vTBCID As String = $"{Hid_TBCID.Value}"
        Dim vKTSEQ As String = TIMS.GetListValue(ddlSwitchTo)
        Dim vFUNCNM As String = "以目前版本送出"
        If Hid_TBCASENO.Value = "" OrElse Hid_TBCID.Value = "" Then
            Common.MessageBox(Me, $"{vFUNCNM}資訊有誤(案件號為空)，請重新操作!!")
            Return
        ElseIf Hid_KTSEQ.Value = "" Then
            Common.MessageBox(Me, $"{vFUNCNM}資訊有誤(項目代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_ORGKINDGW.Value = "" Then
            Common.MessageBox(Me, $"{vFUNCNM}資訊有誤(計畫代碼為空)，請重新操作!!")
            Return
        ElseIf Hid_KTSEQ.Value <> "" AndAlso Hid_KTSEQ.Value <> vKTSEQ Then
            Common.MessageBox(Me, $"{vFUNCNM}資訊有誤(項目代碼／序號有誤)，請重新操作!!")
            Return
        End If
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, vKTSEQ, Hid_ORGKINDGW.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無案件編號)，請重新操作!!")
            Return
        ElseIf drKB Is Nothing Then
            Common.MessageBox(Me, "下載報表資訊有誤(查無項目編號)，請重新操作!!")
            Return
        End If

        Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, vTBCID, vKTSEQ)
        '(退件修正)有退件原因,可重新上傳
        'Dim flag_NG_UPLOAD_1 As Boolean=(drFL IsNot Nothing) '(有資料 不可再次傳送)
        Dim flag_NG_UPLOAD_2 As Boolean = (drFL IsNot Nothing AndAlso $"{drFL("RTUREASON")}" = "") '(有資料不可傳送且原因為空 不可再次傳送)
        Dim vFILENAME1 As String = If(drFL IsNot Nothing, $"{drFL("FILENAME1")}", "")
        'Dim vWAIVED As String=If(drFL IsNot Nothing, Convert.ToString(drFL("WAIVED")), "")
        If vFILENAME1 <> "" AndAlso flag_NG_UPLOAD_2 Then
            '符合所有 不可再次傳送 'cst_tpmsg_enb8
            Common.MessageBox(Me, "已上傳儲存過該文件，不可再次操作!")
            Return
        End If

        Dim rPMS As New Hashtable
        '取得KBID代號／非流水號
        Dim vKTID As String = $"{drKB("KTID")}"
        Dim vORGKINDGW As String = $"{drKB("ORGKINDGW")}"
        Dim vKBNAME2 As String = $"{vORGKINDGW}{vKTID}{drKB("KTNAME")}"
        Select Case $"{vORGKINDGW}{vKTID}"
            Case G03_訓練課程開班學員名冊, W03_訓練課程開班學員名冊
                Dim iNum As Integer = 2 '排序方式:If(vINUM = "2", "a.StudentID", "c.IDNO")
                Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
                If drCC Is Nothing Then Return

                Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
                Dim vDISTID As String = $"{drCC("DISTID")}"
                Dim vPLANID As String = $"{drCC("PLANID")}"
                Dim vTBCASENO As String = $"{drOA("TBCASENO")}"
                Dim vUploadPath As String = TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, "")

                'Dim vTBCID As String = $"{Hid_TBCID.Value}"
                'Dim vFILENAME1 As String = TIMS.GET_FILENAME1_S1(vTBCID, vKTSEQ, "", "pdf")
                vFILENAME1 = TIMS.GET_FILENAME1_S1(vTBCID, vKTSEQ, "", "pdf")
                Dim vSRCFILENAME1 As String = $"{vKBNAME2}x{Now.ToString("fffss")}.pdf" 'MyFileName
                'Dim vFILEPATH1 As String = vUploadPath ' TIMS.GetMyValue2(rPMS, "FILEPATH1")

                rPMS.Add("TBCID", TIMS.CINT1(vTBCID))
                rPMS.Add("KTSEQ", vKTSEQ)
                rPMS.Add("FILENAME1", vFILENAME1)
                rPMS.Add("SRCFILENAME1", vSRCFILENAME1)
                rPMS.Add("UploadPath", vUploadPath)

                rPMS.Add("INUM", iNum)
                rPMS.Add("ORGKINDGW", vORGKINDGW)
                rPMS.Add("OCID", $"{drOA("OCID")}")
                rPMS.Add("TPlanID", sm.UserInfo.TPlanID)
                Call SAVE_SD_14_008(rPMS)
        End Select

        Threading.Thread.Sleep(10)

        '檢視目前上傳檔案
        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)
    End Sub

    Private Sub DataGrid06_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid06.ItemDataBound
        'ROWNUM1,序號>,SRCFILENAME1,上傳檔案>,MEMO1,備註說明>,Hid_CS14OFID,BTN_DOWNLOAD06,檔案下載,DOWNLOAD06,BTN_DELFILE06,刪除檔案,DELFILE06,
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                'e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                Dim BTN_DELFILE06 As Button = e.Item.FindControl("BTN_DELFILE06") '(刪除)
                Dim BTN_DOWNLOAD06 As Button = e.Item.FindControl("BTN_DOWNLOAD06") '查看
                ',a.CS14OFID,a.TBCFID,a.TBCID,a.KTSEQ
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "CS14OFID", drv("CS14OFID"))
                TIMS.SetMyValue(sCmdArg, "TBCFID", drv("TBCFID"))
                TIMS.SetMyValue(sCmdArg, "TBCID", drv("TBCID"))
                TIMS.SetMyValue(sCmdArg, "KTSEQ", drv("KTSEQ"))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", drv("FILENAME1"))
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", drv("FILEPATH1"))
                TIMS.SetMyValue(sCmdArg, "KTID", drv("KTID"))
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                'Dim flagS1 As Boolean = TIMS.IsSuperUser(sm, 1) '是否為(後台)系統管理者 
                'BTN_DELFILE06.Visible = If(flagS1, True, False)
                'BTN_DELFILE06.Style.Item("display") = If(flagS1, "", "none")
                BTN_DELFILE06.CommandArgument = sCmdArg
                Dim vMsgB As String = "請注意：此筆線上申辦案件原已上傳之相關文件均會一併刪除，確定要刪除此筆資料嗎?"
                BTN_DELFILE06.Attributes("onclick") = $"javascript:return confirm('{vMsgB}');"
                '檢視不能修改
                BTN_DELFILE06.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
                '(其他原因調整) '送件／退件修正，不提供刪除
                If $"{drv("TBCSTATUS")}" = "B" Then
                    BTN_DELFILE06.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE06, cst_tpmsg_enb6, True)
                ElseIf $"{drv("TBCSTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" <> "" Then
                    BTN_DELFILE06.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                    TIMS.Tooltip(BTN_DELFILE06, cst_tpmsg_enb8, True)
                ElseIf $"{drv("TBCSTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" = "" Then
                    BTN_DELFILE06.Enabled = False
                    TIMS.Tooltip(BTN_DELFILE06, cst_tpmsg_enb7, True)
                End If
                BTN_DOWNLOAD06.CommandArgument = sCmdArg
        End Select
    End Sub

    Private Sub DataGrid06_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid06.ItemCommand
        'ROWNUM1,序號>,SRCFILENAME1,上傳檔案>,MEMO1,備註說明>,Hid_CS14OFID,BTN_DOWNLOAD06,檔案下載,DOWNLOAD06,BTN_DELFILE06,刪除檔案,DELFILE06,
        'Dim HFileName As HtmlInputHidden=e.Item.FindControl("HFileName")
        Dim sCmdArg As String = e.CommandArgument
        Dim vCS14OFID As String = TIMS.GetMyValue(sCmdArg, "CS14OFID")
        Dim vTBCFID As String = TIMS.GetMyValue(sCmdArg, "TBCFID")
        Dim vKTID As String = TIMS.GetMyValue(sCmdArg, "KTID")
        Dim vKTSEQ As String = TIMS.GetMyValue(sCmdArg, "KTSEQ")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        If e.CommandArgument = "" OrElse vCS14OFID = "" OrElse vOCID = "" Then Return

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Dim vTBCASENO As String = Hid_TBCASENO.Value
        Select Case e.CommandName
            Case "DELFILE06"
                'Dim sErrMsg1 As String = CHKDEL_CLASS_STD14OAFL_OTH(vCS14OFID)
                'If sErrMsg1 <> "" Then
                '    Common.MessageBox(Me, sErrMsg1)
                '    Return
                'End If

                Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
                If drCC Is Nothing Then
                    Common.MessageBox(Me, "上傳資訊有誤(查無職類/班別代碼)，請重新操作!")
                    Return
                End If
                Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
                Dim vDISTID As String = $"{drCC("DISTID")}"
                Dim vPLANID As String = $"{drCC("PLANID")}"
                Dim oFILENAME1 As String = ""
                Dim oUploadPath As String = ""
                Dim s_FilePath1 As String = ""
                Try
                    oFILENAME1 = vFILENAME1
                    oUploadPath = If(vFILEPATH1 <> "", vFILEPATH1, TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, vKTSEQ))
                    s_FilePath1 = Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    Call TIMS.MyFileDelete(s_FilePath1)
                Catch ex As Exception
                    Dim strErrmsg As String = String.Concat(New Diagnostics.StackFrame(True).GetMethod().Name, vbCrLf)
                    strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf, "s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try

                'DELETE CLASS_STD14OAFL_OTH
                Dim dParms2 As New Hashtable From {{"CS14OFID", vCS14OFID}}
                Dim rdSql2 As String = "DELETE CLASS_STD14OAFL_OTH WHERE CS14OFID=@CS14OFID"
                DbAccess.ExecuteNonQuery(rdSql2, objconn, dParms2)

            Case "DOWNLOAD06" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "TBCID", Hid_TBCID.Value)
                TIMS.SetMyValue2(rPMS4, "TBCASENO", Hid_TBCASENO.Value)
                TIMS.SetMyValue2(rPMS4, "CS14OFID", vCS14OFID)
                TIMS.SetMyValue2(rPMS4, "TBCFID", vTBCFID)
                TIMS.SetMyValue2(rPMS4, "KTID", vKTID)
                TIMS.SetMyValue2(rPMS4, "KTSEQ", vKTSEQ)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)
                Call TIMS.ResponseZIPFile_OA(sm, objconn, Me, rPMS4)
                Return
        End Select

        If Not TIMS.OpenDbConn(objconn) Then Return
        Call SHOW_DATAGRID_06()

        '顯示上傳檔案／細項
        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)
    End Sub

    Private Function CHKDEL_CLASS_STD14OAFL(vORGKINDGW As String, vKTID As String, vTBCFID As String) As String
        Dim rst As String = ""
        Select Case $"{vORGKINDGW}{vKTID}"
            Case G06_其他補充資料, W06_其他補充資料
                'DELETE CLASS_STD14OAFL_OTH
                Dim sParms3 As New Hashtable From {{"TBCFID ", vTBCFID}}
                Dim sSql3 As String = "SELECT 1 FROM CLASS_STD14OAFL_OTH WHERE TBCFID=@TBCFID"
                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql3, objconn, sParms3)
                If TIMS.dtHaveDATA(dt3) Then rst &= "該項目，有檔案資訊(子項)，不可刪除"
        End Select
        Return rst
    End Function

    Protected Sub DataGrid06_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid06.SelectedIndexChanged

    End Sub

    Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged

    End Sub
End Class


