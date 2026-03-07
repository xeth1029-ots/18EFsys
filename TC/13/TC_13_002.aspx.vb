Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Partial Class TC_13_002
    Inherits AuthBasePage

    '線上核銷送件--委訓單位使用
    'Dim fg_test As Boolean=TIMS.sUtl_ChkTest() '測試

    '基本上分署單位使用
    '申請階段管理-受理期間設定 APPLISTAGE
    'Dim fg_can_applistage As Boolean=False
    Dim tryFIND As String = ""

    Const cst_MaxLen500_i As Integer = 500
    Const cst_MaxLen500_TPMSG1 As String = "限定500字元"

    '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
    'Dim vAPPLIEDRESULT As String=""

    Dim iDG11_ROWS As Integer = 0
    Dim iDG10_ROWS As Integer = 0
    '以目前版本批次送出
    Const cst_txt_版本批次送出 As String = "(版本批次送出)"
    Const cst_txt_免附文件 As String = "(免附文件)"
    'Const cst_printASPX_R As String="../../SD/14/SD_14_002_R.aspx?ID="
    'Dim sPrintASPX1 As String=""

    'outTYPE: CLSNM,PCSVAL
    'Const cst_outTYPE_CLSNM As String="CLSNM"
    'Const cst_outTYPE_PCSVAL As String="PCSVAL"

    Const cst_ss_RqProcessType As String = "RqProcessType"
    'Const cst_DG1CMDNM_VIEW1 As String="VIEW1"
    Const cst_DG1CMDNM_EDIT1 As String = "EDIT1" '審核/審查/確認
    Const cst_DG1CMDNM_REVERT2 As String = "REVERT2" '還原"-確認

    Const cst_DG2CMDNM_RtuBACK1 As String = "RtuBACK1" '退回開放修改
    Const cst_DG2CMDNM_REVERT1 As String = "REVERT1" '還原"
    Const cst_DG2CMDNM_VIEWFILE4 As String = "VIEWFILE4" '查詢
    Const cst_DG2CMDNM_DOWNLOAD4 As String = "DOWNLOAD4" '下載

    'Dim G_UPDRV As String="~/UPDRV"
    'Dim G_UPDRV_JS As String="../../UPDRV"

    'Const cst_errMsg_1 As String="資料有誤請重新查詢!"
    'Const cst_errMsg_2 As String="上傳檔案時發生錯誤，請重新操作!(若持續發生請連絡系統管理者)" 'Const cst_errMsg_2 As String="上傳檔案壓縮時發生錯誤，請重新確認上傳檔案格式!"
    'Const cst_errMsg_3 As String="檔案位置錯誤!"
    'Const cst_errMsg_4 As String="檔案類型錯誤!"
    'Const cst_errMsg_5 As String="檔案類型錯誤，必須為PDF類型檔案!"
    Const cst_errMsg_6 As String = "(檔案上傳失敗／異常，請刪除後重新上傳)"
    'Const cst_PostedFile_MAX_SIZE As Integer=2097152 '10*1024*1024 '2*1024*1024
    'Const cst_errMsg_7 As String="檔案大小超過2MB!"
    'Const cst_errMsg_8 As String="請選擇上傳檔案(不可為空)!"
    ''Const cst_errMsg_9 As String="請選擇場地圖片--隸屬於教室1 或教室2!"
    'Const cst_errMsg_11 As String="無效的檔案格式。"
    'Const cst_errMsg_21 As String="不可勾選免附文件又按上傳檔案。"


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

        Call CCreate11() '每次執行

        If Not IsPostBack Then
            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            Call cCreate1(0)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button3.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            'HistoryRID.Attributes("onclick")="ShowFrame();"
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If
    End Sub

    ''' <summary>每次執行</summary>
    Sub CCreate11()
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1 '分頁設定
        'PageControler1.PageDataTable=dt
        'PageControler1.ControlerLoad()
        'sPrintASPX1=String.Concat(cst_printASPX_R, TIMS.Get_MRqID(Me))
        '<add key="UPLOAD_OJT_Path" value="~/UPDRV" />
        '<add key="DOWNLOAD_OJT_Path" value="../../UPDRV" />
        'Dim vUPLOAD_OJT_Path As String=TIMS.Utl_GetConfigSet("UPLOAD_OJT_Path")
        'Dim vDOWNLOAD_OJT_Path As String=TIMS.Utl_GetConfigSet("DOWNLOAD_OJT_Path")
        'Dim G_UPDRV As String="~/UPDRV"
        'Dim G_UPDRV_JS As String="../../UPDRV"
        'If (vUPLOAD_OJT_Path <> "") Then G_UPDRV=vUPLOAD_OJT_Path
        'If (vDOWNLOAD_OJT_Path <> "") Then G_UPDRV_JS=vDOWNLOAD_OJT_Path
    End Sub

    '設定 資料與顯示 狀況！
    Private Sub cCreate1(ByVal iNum As Integer)
        TableDataGrid1.Visible = False
        labmsg1.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        Call SHOW_Frame1(0)

        '申辦人姓名
        sch_txtSENDACCTNAME.Text = ""
        '申辦日期
        sch_txtSENDDATE1.Text = ""
        sch_txtSENDDATE2.Text = ""

        Dim MRqID As String = TIMS.Get_MRqID(Me)
        TIMS.Get_TitleLab(objconn, MRqID, TitleLab1, TitleLab2)
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

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        '清理隱藏的參數
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
    ''' <summary>查詢1</summary>
    Private Sub SSearch1()
        '清理隱藏的參數
        Call ClearHidValue()

        Call SHOW_Frame1(0)
        labmsg1.Text = TIMS.cst_NODATAMsg1
        TableDataGrid1.Visible = False

        'RIDValue.Value=If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If

        sch_txtSENDACCTNAME.Text = TIMS.ClearSQM(sch_txtSENDACCTNAME.Text)
        sch_txtSENDDATE1.Text = TIMS.Cdate3(sch_txtSENDDATE1.Text)
        sch_txtSENDDATE2.Text = TIMS.Cdate3(sch_txtSENDDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(sch_txtSENDDATE1.Text, sch_txtSENDDATE2.Text) Then
            Dim T_DATE1 As String = sch_txtSENDDATE1.Text
            sch_txtSENDDATE1.Text = sch_txtSENDDATE2.Text
            sch_txtSENDDATE2.Text = T_DATE1
        End If

        Dim vDISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        Dim pms_s1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"DISTID", vDISTID}}
        Dim sSql As String = ""
        sSql &= " SELECT a.CVOCID, a.OCID,a.SEQ_ID,a.APPLIEDRESULT" & vbCrLf
        sSql &= " ,ip.YEARS,dbo.FN_CYEAR2(ip.YEARS) YEARS_ROC" & vbCrLf
        sSql &= " ,ip.DISTID,ip.DISTNAME" & vbCrLf
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
        sSql &= " ,format(a.SENDDATE,'yyyy/MM/dd') SENDDATE" & vbCrLf
        sSql &= " ,dbo.FN_CDATE1B(a.SENDDATE) SENDDATE_ROC" & vbCrLf
        sSql &= " ,a.SENDACCT,aa.NAME SENDACCTNAME" & vbCrLf

        sSql &= " ,a.RESULTACCT,a.RESULTDATE,a.REASONFORFAIL"
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        sSql &= " ,a.APPLIEDRESULT"
        sSql &= " ,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N" & vbCrLf
        sSql &= " ,CASE oo.ORGKIND WHEN '10' THEN 'W' ELSE 'G' END COLLATE Chinese_Taiwan_Stroke_CS_AS ORGKINDGW"
        sSql &= " FROM CLASS_VERIFYONLINE a" & vbCrLf
        sSql &= " JOIN CLASS_CLASSINFO cc ON cc.OCID=a.OCID" & vbCrLf
        sSql &= " JOIN PLAN_PLANINFO pp ON cc.PLANID=pp.PLANID AND cc.COMIDNO=pp.COMIDNO AND cc.SEQNO=pp.SEQNO" & vbCrLf
        sSql &= " JOIN VIEW_PLAN ip on ip.PLANID=cc.PLANID" & vbCrLf
        sSql &= " JOIN ORG_ORGINFO oo on oo.COMIDNO=cc.COMIDNO" & vbCrLf
        sSql &= " LEFT JOIN AUTH_ACCOUNT aa on aa.ACCOUNT=a.SENDACCT" & vbCrLf
        sSql &= " WHERE pp.IsApprPaper='Y' AND cc.IsSuccess='Y' AND cc.NotOpen='N'" & vbCrLf
        sSql &= " AND ip.TPLANID=@TPLANID and ip.YEARS=@YEARS" & vbCrLf
        sSql &= " AND ip.DISTID=@DISTID" & vbCrLf
        sSql &= " AND a.SENDSTATUS IS NOT NULL" & vbCrLf
        If OCIDValue1.Value <> "" Then
            pms_s1.Add("OCID", OCIDValue1.Value)
            sSql &= " and cc.OCID=@OCID" & vbCrLf
        ElseIf RIDValue.Value.Length <> 1 Then
            pms_s1.Add("RIDValue", RIDValue.Value)
            sSql &= " AND cc.RID=@RIDValue" & vbCrLf
        End If
        If sch_txtSENDACCTNAME.Text <> "" Then
            pms_s1.Add("SENDACCTNAME", sch_txtSENDACCTNAME.Text)
            sSql &= " and aa.NAME like '%'+@SENDACCTNAME+'%'" & vbCrLf
        End If
        If sch_txtSENDDATE1.Text <> "" Then
            pms_s1.Add("SENDDATE1", sch_txtSENDDATE1.Text)
            sSql &= " and a.SENDDATE>=@SENDDATE1" & vbCrLf
        End If
        If sch_txtSENDDATE2.Text <> "" Then
            pms_s1.Add("SENDDATE2", sch_txtSENDDATE2.Text)
            sSql &= " and a.SENDDATE<=@SENDDATE2" & vbCrLf
        End If

        Dim v_rbAPPLIEDRESULT As String = TIMS.GetListValue(rbAPPLIEDRESULT)
        If v_rbAPPLIEDRESULT = "A" Then
            sSql &= " AND a.SENDSTATUS IN ('R','B')" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "B" Then
            sSql &= " AND a.APPLIEDRESULT IS NULL" & vbCrLf
            sSql &= " AND a.SENDSTATUS='B'" & vbCrLf
            'sSql &= " AND (a.APPLIEDRESULT='R' OR a.APPLIEDRESULT IS NULL)" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "R" Then
            sSql &= " AND a.APPLIEDRESULT='R'" & vbCrLf
            sSql &= " AND a.SENDSTATUS IN ('R','B')" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "Y" OrElse v_rbAPPLIEDRESULT = "N" Then 'Y:申辦確認 / N:申辦不通過
            pms_s1.Add("APPLIEDRESULT", v_rbAPPLIEDRESULT)
            sSql &= " AND a.APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        End If
        sSql &= " ORDER BY a.SENDDATE DESC,a.CVOCID DESC" & vbCrLf

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

    ''' <summary>清理隱藏的參數</summary>
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

    ''' <summary>按下 審查</summary>
    ''' <param name="source"></param>
    ''' <param name="e"></param>
    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vCVOCID As String = TIMS.GetMyValue(sCmdArg, "CVOCID")
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim vSEQ_ID As String = TIMS.GetMyValue(sCmdArg, "SEQ_ID")
        Dim vSENDSTATUS As String = TIMS.GetMyValue(sCmdArg, "SENDSTATUS")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        If sCmdArg = "" OrElse vCVOCID = "" OrElse vOCID = "" OrElse vSEQ_ID = "" Then Return

        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCID, vSEQ_ID, vCVOCID) 'If drRR Is Nothing Then Return
        If drCV Is Nothing Then Return
        Dim drCC As DataRow = TIMS.GetOCIDDate(vOCID, objconn)
        If drCC Is Nothing Then Return

        Select Case e.CommandName
            Case cst_DG1CMDNM_EDIT1 '"EDIT1"審核
                'If Not fg_can_applistage Then
                '    Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                '    Return
                'End If
                If Convert.ToString(drCV("SENDSTATUS")) = "R" Then
                    Common.MessageBox(Me, "申辦狀態 退件待修正，待修正後送審，再行審核!")
                    Return 'If Not fg_test Then Return
                ElseIf Convert.ToString(drCV("SENDSTATUS")) = "" Then
                    Common.MessageBox(Me, "申辦狀態 未填寫，待送審後，再行審核!")
                    Return 'If Not fg_test Then Return
                End If
                '查詢使用資料顯示 依 CLASS_VERIFYONLINE-CVOCID '"EDIT1"審核
                Call SHOW_Detail_VERIFYONLINE(drCV, drCC, vOCID, vSEQ_ID, vCVOCID, cst_DG1CMDNM_EDIT1)

            Case cst_DG1CMDNM_REVERT2 '"REVERT2" '還原"-確認
                '當【審核狀態】：已申辦確認、已申辦不通過，按下【還原】鈕即清空【審核狀態】且【申辦狀態】：已送件， 如圖二狀態
                '當【審核狀態】：退件修正，【還原】鈕反灰不可按
                If Convert.ToString(drCV("APPLIEDRESULT")) = "R" Then
                    Common.MessageBox(Me, "審核狀態 退件待修正，不可還原!")
                    Return
                ElseIf Convert.ToString(drCV("APPLIEDRESULT")) = "" Then
                    Common.MessageBox(Me, "審核狀態 未填寫，不可還原!")
                    Return
                End If

                Const cst_REVERT2_N As String = "還原"
                Dim s_HISREVIEW As String = Convert.ToString(drCV("HISREVIEW"))
                s_HISREVIEW &= String.Concat(If(s_HISREVIEW <> "", "，", ""), String.Concat(TIMS.Cdate3t(Now), "-", cst_REVERT2_N))
                '線上送件審核，新增還原按鈕 UPDATE ORG_BIDCASE / BISTATUS='B',APPLIEDRESULT=NULL
                Call UPDATE_REVERT2(s_HISREVIEW, vOCID, vSEQ_ID, vCVOCID)
                '查詢1
                Call SSearch1()

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                Dim drv As DataRowView = e.Item.DataItem
                Dim lBTN_EDIT1 As LinkButton = e.Item.FindControl("lBTN_EDIT1") '審核/審查/確認
                Dim lBTN_REVERT2 As LinkButton = e.Item.FindControl("lBTN_REVERT2") '還原
                'OJT-20231128:線上送件審核-還原按鈕 NULL/Y:可使用
                'lBTN_REVERT2.Visible=(Hid_USE_ORG_BIDCASE_REVERT2.Value="Y")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "CVOCID", drv("CVOCID"))
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                TIMS.SetMyValue(sCmdArg, "SEQ_ID", drv("SEQ_ID"))
                TIMS.SetMyValue(sCmdArg, "SENDSTATUS", drv("SENDSTATUS"))
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

    ''' <summary>新增使用資料顯示／查詢使用資料顯示 依 CLASS_VERIFYONLINE-CVOCID</summary>
    Private Sub SHOW_Detail_VERIFYONLINE(drCV As DataRow, drCC As DataRow, vOCID As String, vSEQ_ID As String, vCVOCID As String, vCmdName As String)
        '資料有誤
        If drCV Is Nothing Then Return
        '資料有誤
        If drCC Is Nothing Then Return
        Call SHOW_Frame1(1)
        If vCVOCID = "" Then Return

        Session(cst_ss_RqProcessType) = vCmdName

        Hid_ORGKINDGW.Value = Convert.ToString(drCC("ORGKINDGW"))

        tr_HISREVIEW.Visible = False '歷程資訊

        '審查狀態
        Common.SetListItem(ddlAPPLIEDRESULT, Convert.ToString(drCV("APPLIEDRESULT")))
        '不通過原因
        Reasonforfail.Text = Convert.ToString(drCV("REASONFORFAIL"))
        TIMS.Tooltip(Reasonforfail, cst_MaxLen500_TPMSG1, True)

        Hid_ORGKINDGW.Value = Convert.ToString(drCV("ORGKINDGW"))

        'Hid_KVSID.Value=""
        'Hid_KVID.Value=""
        'Hid_LastKVID.Value=""
        'Hid_FirstKVSID.Value=""

        'Hid_CVOCFID.Value=""
        Hid_CVOCID.Value = Convert.ToString(drCV("CVOCID"))
        Hid_OCIDVal.Value = Convert.ToString(drCV("OCID"))
        Hid_SEQ_ID.Value = Convert.ToString(drCV("SEQ_ID"))

        labOrgNAME.Text = Convert.ToString(drCV("ORGNAME"))
        labSEND_YEARS_ROC.Text = TIMS.GET_YEARS_ROC(drCV("YEARS"))
        labAPPSTAGE.Text = Convert.ToString(drCV("APPSTAGE_N"))
        labCLASSNAME2.Text = Convert.ToString(drCV("CLASSCNAME2"))  'TIMS.GetResponseWrite(strCLASSNAME2S)

        Dim vCurrentKVSID As String = GET_CLASS_VERIFYONLINE_FL_CurrentKVSID(objconn, Hid_CVOCID.Value)
        Hid_KVSID.Value = vCurrentKVSID 'Convert.ToString(drCV("CurrentKVSID"))

        'Dim t_ddlAPPLIEDRESULT As String=TIMS.GetListText(ddlAPPLIEDRESULT)
        Hid_APPLIEDRESULT.Value = Convert.ToString(drCV("APPLIEDRESULT"))
        Dim fg_have_APPLIEDRESULT As Boolean = (Hid_APPLIEDRESULT.Value = "N" OrElse Hid_APPLIEDRESULT.Value = "Y")
        Dim t_ddlAPPLIEDRESULT As String = If(fg_have_APPLIEDRESULT, TIMS.GetListText(ddlAPPLIEDRESULT), "")
        Reasonforfail.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        ddlAPPLIEDRESULT.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        But_Sub.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        TIMS.Tooltip(Reasonforfail, t_ddlAPPLIEDRESULT, True)
        TIMS.Tooltip(ddlAPPLIEDRESULT, t_ddlAPPLIEDRESULT, True)
        TIMS.Tooltip(But_Sub, t_ddlAPPLIEDRESULT, True)

        '歷程資訊
        If Convert.ToString(drCV("HISREVIEW")) <> "" AndAlso Convert.ToString(drCV("HISREVIEW")).Length > 1 Then
            tr_HISREVIEW.Visible = True '歷程資訊
            labHISREVIEW.Text = Convert.ToString(drCV("HISREVIEW"))
        End If

        Hid_CVOCID.Value = TIMS.ClearSQM(drCV("CVOCID"))
        Hid_OCIDVal.Value = TIMS.ClearSQM(drCV("OCID"))
        Hid_SEQ_ID.Value = TIMS.ClearSQM(drCV("SEQ_ID"))
        '檢視目前上傳檔案 '顯示上傳檔案／細項 -'線上申辦進度
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)

    End Sub

    Private Function GET_CLASS_VERIFYONLINE_FL_CurrentKVSID(oConn As SqlConnection, vCVOCID As String) As String
        Dim rst As String = ""
        If vCVOCID = "" Then Return rst
        Dim pms_s1 As New Hashtable From {{"CVOCID", Val(vCVOCID)}}
        '返回第1項
        Dim sSql As String = " SELECT MIN(KVSID) CurrentKBSID FROM CLASS_VERIFYONLINE_FL WHERE CVOCID=@CVOCID" & vbCrLf
        Dim drFL As DataRow = DbAccess.GetOneRow(sSql, oConn, pms_s1)
        If drFL Is Nothing Then Return rst
        rst = Convert.ToString(drFL("CurrentKBSID"))
        Return rst
    End Function

    ''' <summary>直接顯示項目-檢視目前上傳檔案</summary>
    ''' <param name="rPMS"></param>
    Private Sub SHOW_VERIFYONLINE_FL_DG2(rPMS As Hashtable)
        labmsg1.Text = ""
        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim fg_CANSAVE As Boolean = (vCVOCID <> "" AndAlso Val(vCVOCID) >= 0)
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

        labmsg1.Text = If(dtFL Is Nothing OrElse dtFL.Rows.Count = 0, "(查無文件項目)", "")

        Dim vYEARS As String = Convert.ToString(drCV("YEARS"))
        Dim vPLANID As String = Convert.ToString(drCV("PLANID"))
        Dim vCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
        Dim vSEQNO As String = Convert.ToString(drCV("SEQNO"))
        Dim vORGKINDGW As String = Convert.ToString(drCV("ORGKINDGW"))
        Dim download_Path As String = TIMS.GET_DOWNLOADPATH1_CVO(vYEARS, vPLANID, vCOMIDNO, vSEQNO, vCVOCID, "")
        Call Check_dtVERIFYONLINE_FL(Me, dtFL, download_Path)
        'DataGrid2.Columns(cst_DG2_退件原因_iCOLUMN).Visible=If(Convert.ToString(drCV("APPLIEDRESULT"))="R", True, False)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

        'Dim iProgress As Integer=If(dtA.Rows.Count > 0, (dt.Rows.Count / dtA.Rows.Count * 100), 0)
        ''線上申辦進度 計算完成度百分比 (0-100)
        'Dim iProgress As Integer=GET_iPROGRESS_CVO(sm, objconn, tmpMSG, vCVOCID, vORGKINDGW)
        'labProgress.Text=String.Concat(iProgress, "%")
        ''BTN_SAVETMP1.Visible=(iProgress=100)
        ''BTN_SAVERC2.Visible=(iProgress=100)
        ''儲存(暫存)
        'BTN_SAVETMP1.Enabled=If(Session(cst_ss_RqProcessType)=cst_DG1CMDNM_VIEW1, False, True)
        'TIMS.Tooltip(BTN_SAVETMP1, If(BTN_SAVETMP1.Enabled, "", cst_tpmsg_enb1), True)
        ''儲存後進下一步
        'BTN_SAVENEXT1.Enabled=If(Session(cst_ss_RqProcessType)=cst_DG1CMDNM_VIEW1, False, True)
        'TIMS.Tooltip(BTN_SAVENEXT1, If(BTN_SAVENEXT1.Enabled, "", cst_tpmsg_enb1), True)
    End Sub


    ''' <summary>檢核實際檔案-"(檔案上傳失敗／異常，請刪除後重新上傳)"</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="dt"></param>
    ''' <param name="download_Path"></param>
    Private Sub Check_dtVERIFYONLINE_FL(MyPage As Page, dt As DataTable, download_Path As String)
        'Dim filename As String=""
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return
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


#Region "PRINT_1"
    ''' <summary>cst_W13_教學環境資料表 '"13" '教學環境資料表</summary>
    ''' <param name="rPMS"></param>
    Private Sub RPT_SD_14_014(ByRef rPMS As Hashtable)
        Const cst_printFN1 As String = "SD_14_014" '0:未轉班' 1:已轉班
        Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
        Dim TSTPRINT As String = If(sPrint_Test = "Y", "2", "1") '測試區2／'正式區1 

        'Const cst_printFN2 As String="SD_14_014_1" '2:變更待審
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
#End Region


#Region "NO USE"
    '教學環境資料表
    'Private Shared Function GET_ORG_BIDCASEFL_EV(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String=""
    '    sSql &= " SELECT b.BCFEID,a.BCID,a.BCPID, a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,pp.PSNO28" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_EV b on b.BCPID=a.BCPID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
    '    sSql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID and pp.COMIDNO=a.COMIDNO and pp.SEQNO=a.SEQNO" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable=DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function

    'Private Shared Function GET_ORG_BIDCASEFL_TT(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String=""
    '    sSql &= " SELECT a.BCID,f.BCFID" & vbCrLf
    '    sSql &= " ,b.BCFTID,b.TECHID" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,tt.TEACHCNAME,tt.TEACHERID" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASE a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCID=a.BCID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_TT b on b.BCFID=f.BCFID" & vbCrLf
    '    sSql &= " JOIN TEACH_TEACHERINFO tt on tt.TECHID=b.TECHID" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable=DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function

    'Private Shared Function GET_ORG_BIDCASEFL_TT2(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String=""
    '    sSql="" & vbCrLf
    '    sSql &= " SELECT a.BCID,f.BCFID" & vbCrLf
    '    sSql &= " ,b.BCFT2ID,b.TECHID" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,tt.TEACHCNAME,tt.TEACHERID" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASE a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCID=a.BCID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_TT2 b on b.BCFID=f.BCFID" & vbCrLf
    '    sSql &= " JOIN TEACH_TEACHERINFO tt on tt.TECHID=b.TECHID" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable=DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function

    '取得班級申請資料
    'Private Shared Function GET_ORG_BIDCASEFL_PI(vBCID As String) As DataTable
    '    Dim sPMS As New Hashtable
    '    sPMS.Add("BCID", vBCID)
    '    Dim sSql As String=""
    '    sSql &= " SELECT a.BCID,a.BCPID, a.PLANID,a.COMIDNO,a.SEQNO" & vbCrLf
    '    sSql &= " ,b.SRCFILENAME1,b.FILENAME1" & vbCrLf
    '    sSql &= " ,pp.PSNO28" & vbCrLf
    '    sSql &= " FROM ORG_BIDCASEPI a" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL_PI b on b.BCPID=a.BCPID" & vbCrLf
    '    sSql &= " JOIN ORG_BIDCASEFL f on f.BCFID=b.BCFID" & vbCrLf
    '    sSql &= " JOIN PLAN_PLANINFO pp on pp.PLANID=a.PLANID and pp.COMIDNO=a.COMIDNO and pp.SEQNO=a.SEQNO" & vbCrLf
    '    sSql &= " WHERE a.BCID=@BCID" & vbCrLf
    '    Dim dt As DataTable=DbAccess.GetDataTable(sSql, objconn, sPMS)
    '    Return dt
    'End Function
#End Region

    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        'Dim LabFileName1 As Label=e.Item.FindControl("LabFileName1")
        'Dim HFileName As HtmlInputHidden=e.Item.FindControl("HFileName")
        '退回原因說明／ '退回開放修改／ '還原
        Dim txtRtuReason As TextBox = e.Item.FindControl("txtRtuReason") '退回原因說明
        Dim Btn_RtuBACK1 As Button = e.Item.FindControl("Btn_RtuBACK1") '退回開放修改
        Dim Btn_REVERT1 As Button = e.Item.FindControl("Btn_REVERT1") '還原
        Dim sCmdArg As String = e.CommandArgument

        Dim vCVOCID As String = TIMS.GetMyValue(sCmdArg, "CVOCID")
        'Dim vOCID As String=TIMS.GetMyValue(sCmdArg, "OCID")
        'Dim vSEQ_ID As String=TIMS.GetMyValue(sCmdArg, "SEQ_ID")
        Dim vCVOCFID As String = TIMS.GetMyValue(sCmdArg, "CVOCFID")
        Dim vKVID As String = TIMS.GetMyValue(sCmdArg, "KVID")
        Dim vKVSID As String = TIMS.GetMyValue(sCmdArg, "KVSID")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")

        Hid_CVOCID.Value = vCVOCID 'TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        If e.CommandArgument = "" OrElse vCVOCID = "" OrElse vCVOCFID = "" OrElse vKVID = "" OrElse vKVSID = "" Then Return

        If txtRtuReason IsNot Nothing Then txtRtuReason.Text = TIMS.ClearSQM(txtRtuReason.Text)

        Dim drFL As DataRow = TIMS.GET_CLASS_VERIFYONLINE_FL(objconn, vCVOCID, vKVSID, vCVOCFID)
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
                Dim uParms As New Hashtable From {
                    {"CVOCID", Val(vCVOCID)},
                    {"CVOCFID", Val(vCVOCFID)},
                    {"RTUREASON", If(v_txtRtuReason <> "", v_txtRtuReason, Convert.DBNull)},
                    {"RTURESACCT", sm.UserInfo.UserID}
                }
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_VERIFYONLINE_FL" & vbCrLf
                usSql &= " SET RTUREASON=@RTUREASON,RTURESACCT=@RTURESACCT,RTURESDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE CVOCID=@CVOCID AND CVOCFID=@CVOCFID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Case cst_DG2CMDNM_REVERT1 '還原
                'txtRtuReason.Text=TIMS.ClearSQM(txtRtuReason.Text)
                Dim uParms As New Hashtable From {{"CVOCID", Val(vCVOCID)}, {"CVOCFID", Val(vCVOCFID)}}
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_VERIFYONLINE_FL" & vbCrLf
                usSql &= " SET RTUREASON=NULL,RTURESACCT=NULL,RTURESDATE=NULL" & vbCrLf
                usSql &= " WHERE CVOCID=@CVOCID AND CVOCFID=@CVOCFID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Case cst_DG2CMDNM_DOWNLOAD4
                Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, Hid_OCIDVal.Value, Hid_SEQ_ID.Value, vCVOCID) 'If drRR Is Nothing Then Return
                If drCV Is Nothing Then Return
                'Hid_ORGKINDGW.Value=Convert.ToString(drCV("ORGKINDGW"))
                'Hid_CVOCID.Value=Convert.ToString(drCV("CVOCID"))
                'Hid_OCIDVal.Value=Convert.ToString(drCV("OCID"))
                'Hid_SEQ_ID.Value=Convert.ToString(drCV("SEQ_ID"))

                ' "DOWNLOAD4" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "CVOCID", Hid_CVOCID.Value)
                TIMS.SetMyValue2(rPMS4, "OCID", Hid_OCIDVal.Value)
                TIMS.SetMyValue2(rPMS4, "SEQ_ID", Hid_SEQ_ID.Value)
                TIMS.SetMyValue2(rPMS4, "CVOCFID", vCVOCFID)
                TIMS.SetMyValue2(rPMS4, "KVID", vKVID)
                TIMS.SetMyValue2(rPMS4, "KVSID", vKVSID)
                TIMS.SetMyValue2(rPMS4, "FILENAME1", vFILENAME1)
                TIMS.SetMyValue2(rPMS4, "FILEPATH1", vFILEPATH1)
                Call TIMS.ResponseZIPFile_CVO(sm, objconn, Me, rPMS4)
                Return

        End Select
        If Not TIMS.OpenDbConn(objconn) Then Return
        '顯示檔案資料表
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        'Hid_ORGKINDGW.Value=Convert.ToString(drCV("ORGKINDGW"))
        'Hid_CVOCID.Value=Convert.ToString(drCV("CVOCID"))
        'Hid_OCIDVal.Value=Convert.ToString(drCV("OCID"))
        'Hid_SEQ_ID.Value=Convert.ToString(drCV("SEQ_ID"))
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        '檢視目前上傳檔案 '顯示上傳檔案／細項 -'線上申辦進度
        Dim rPMS3 As New Hashtable From {{"CVOCID", vCVOCID}}
        Call SHOW_VERIFYONLINE_FL_DG2(rPMS3)

    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'Dim LabdepID As Label=e.Item.FindControl("LabdepID")
                'Dim LabFileName1 As Label=e.Item.FindControl("LabFileName1")
                'Dim HFileName As HtmlInputHidden=e.Item.FindControl("HFileName")
                Dim txtRtuReason As TextBox = e.Item.FindControl("txtRtuReason") '退回原因說明
                Dim Btn_RtuBACK1 As Button = e.Item.FindControl("Btn_RtuBACK1") '退回開放修改
                Dim Btn_REVERT1 As Button = e.Item.FindControl("Btn_REVERT1") '還原
                Dim BTN_VIEWFILE4 As Button = e.Item.FindControl("BTN_VIEWFILE4")
                Dim BTN_DOWNLOAD4 As Button = e.Item.FindControl("BTN_DOWNLOAD4")
                BTN_VIEWFILE4.Visible = False
                'BTN_DOWNLOAD4.Visible=False

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
                    'LabFileName1.Text=If(Convert.ToString(drv("FILENAME1"))=Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                    'HFileName.Value=Convert.ToString(drv("FILENAME1")) '.ToString()
                ElseIf Convert.ToString(drv("WAIVED")) = "Y" Then
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                    'LabFileName1.Text=cst_txt_免附文件

                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "CVOCID", drv("CVOCID"))
                TIMS.SetMyValue(sCmdArg, "OCID", drv("OCID"))
                TIMS.SetMyValue(sCmdArg, "SEQ_ID", drv("SEQ_ID"))
                TIMS.SetMyValue(sCmdArg, "CVOCFID", drv("CVOCFID"))
                TIMS.SetMyValue(sCmdArg, "KVID", drv("KVID"))
                TIMS.SetMyValue(sCmdArg, "KVSID", drv("KVSID"))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", drv("FILENAME1"))
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", drv("FILEPATH1"))

                BTN_VIEWFILE4.CommandArgument = sCmdArg '查看(查詢細項)
                BTN_DOWNLOAD4.CommandArgument = sCmdArg '下載 
                Btn_RtuBACK1.CommandArgument = sCmdArg '退回開放修改
                Btn_REVERT1.CommandArgument = sCmdArg '還原

                Btn_RtuBACK1.Attributes("onclick") = "return confirm('您確定要「退回開放修改」這一筆資料?');"
                Btn_REVERT1.Attributes("onclick") = "return confirm('您確定要「還原」這一筆資料?');"
                '檢視不能修改
                'BTN_DELFILE4.Visible=If(Session(cst_ss_RqProcessType)=cst_DG1CMDNM_VIEW1, False, True)
        End Select
    End Sub

    ''' <summary>回上一頁</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub But_BACK1_Click(sender As Object, e As EventArgs) Handles But_BACK1.Click
        '清理隱藏的參數
        Call SSearch1()
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
        Dim flag_NG1 As Boolean = CHK_CLASS_VERIFYONLINE_FL_RTUREASON()

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

    Private Function CHK_CLASS_VERIFYONLINE_FL_RTUREASON() As Boolean
        'Dim rst As Boolean=False
        Dim vCVOCID As String = TIMS.ClearSQM(Hid_CVOCID.Value)
        Dim vOCIDVal As String = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Dim vSEQ_ID As String = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        Dim dtFL As DataTable = TIMS.GET_CLASS_VERIFYONLINE_FL_TB(objconn, vCVOCID, vOCIDVal, vSEQ_ID)
        For Each drFL As DataRow In dtFL.Rows
            If Convert.ToString(drFL("RTUREASON")) <> "" Then Return True
        Next
        Return False
    End Function



#Region "Private1"
    ''' <summary>檢核 ORG_BIDCASEFL RTUREASON 有值:true 無值:false </summary>
    ''' <returns></returns>
    Private Function CHK_CLASS_VERIFYONLINE_FL__RTUREASON() As Boolean
        'Dim rst As Boolean=False
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        Dim vCVOCID As String = TIMS.ClearSQM(Hid_CVOCID.Value)
        Dim vOCIDVal As String = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Dim vSEQ_ID As String = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim dtFL As DataTable = TIMS.GET_CLASS_VERIFYONLINE_FL_TB(objconn, vCVOCID, vOCIDVal, vSEQ_ID)
        For Each drFL As DataRow In dtFL.Rows
            If Convert.ToString(drFL("RTUREASON")) <> "" Then Return True
        Next
        Return False
    End Function

    ''' <summary> '儲存 </summary>
    Sub SAVEDATA1()
        '申辦狀態：暫存/ 已送件
        '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)

        Dim vCVOCID As String = Hid_CVOCID.Value
        Dim vOCID As String = Hid_OCIDVal.Value
        Dim vSEQ_ID As String = Hid_SEQ_ID.Value

        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCID, vSEQ_ID, vCVOCID) 'If drRR Is Nothing Then Return
        If drCV Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無核銷編號)，請重新操作!!!")
            Return
        End If

        Dim v_ddlAPPLIEDRESULT As String = TIMS.GetListValue(ddlAPPLIEDRESULT)
        Reasonforfail.Text = TIMS.ClearSQM(Reasonforfail.Text)
        Dim v_Reasonforfail As String = TIMS.Get_Substr1(Reasonforfail.Text, cst_MaxLen500_i)

        Dim s_HISREVIEW As String = Convert.ToString(drCV("HISREVIEW"))
        s_HISREVIEW &= String.Concat(If(s_HISREVIEW <> "", "，", ""), String.Concat(TIMS.Cdate3t(Now), "-", TIMS.GetListText(ddlAPPLIEDRESULT)))

        Dim uParms As New Hashtable
        uParms.Add("HISREVIEW", s_HISREVIEW)
        uParms.Add("APPLIEDRESULT", v_ddlAPPLIEDRESULT)
        uParms.Add("REASONFORFAIL", If(v_Reasonforfail <> "", v_Reasonforfail, Convert.DBNull))
        uParms.Add("RESULTACCT", sm.UserInfo.UserID)
        'uParms.Add("RESULTDATE", RESULTDATE)
        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
        'uParms.Add("MODIFYDATE", MODIFYDATE)
        uParms.Add("OCID", vOCID)
        uParms.Add("SEQ_ID", vSEQ_ID)
        uParms.Add("CVOCID", vCVOCID)
        Dim usSql As String = ""
        usSql &= " UPDATE CLASS_VERIFYONLINE" & vbCrLf
        usSql &= " SET HISREVIEW=@HISREVIEW,APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        usSql &= " ,REASONFORFAIL=@REASONFORFAIL,RESULTACCT=@RESULTACCT,RESULTDATE=GETDATE()" & vbCrLf
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE OCID=@OCID AND SEQ_ID=@SEQ_ID AND CVOCID=@CVOCID" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        If v_ddlAPPLIEDRESULT = "R" Then
            '退件修正 BISTATUS='R'
            Dim uParms2 As New Hashtable
            uParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
            uParms2.Add("OCID", vOCID)
            uParms2.Add("SEQ_ID", vSEQ_ID)
            uParms2.Add("CVOCID", vCVOCID)
            Dim usSql2 As String = ""
            usSql2 &= " UPDATE CLASS_VERIFYONLINE" & vbCrLf
            usSql2 &= " SET SENDSTATUS='R',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql2 &= " WHERE OCID=@OCID AND SEQ_ID=@SEQ_ID AND CVOCID=@CVOCID" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        End If

        Call SSearch1()
    End Sub

    ''' <summary>線上送件審核，還原按鈕 UPDATE CLASS_VERIFYONLINE / SENDSTATUS='B',APPLIEDRESULT=NULL </summary>
    Sub UPDATE_REVERT2(s_HISREVIEW As String, vOCID As String, vSEQ_ID As String, vCVOCID As String)

        vOCID = TIMS.ClearSQM(vOCID)
        vSEQ_ID = TIMS.ClearSQM(vSEQ_ID)
        vCVOCID = TIMS.ClearSQM(vCVOCID)
        If vOCID = "" OrElse vSEQ_ID = "" OrElse vCVOCID = "" Then Return
        '退件修正 BISTATUS
        Dim uParms2 As New Hashtable
        If s_HISREVIEW <> "" Then uParms2.Add("HISREVIEW", s_HISREVIEW)
        uParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
        uParms2.Add("OCID", vOCID)
        uParms2.Add("SEQ_ID", vSEQ_ID)
        uParms2.Add("CVOCID", vCVOCID)
        Dim usSql2 As String = ""
        usSql2 &= " UPDATE CLASS_VERIFYONLINE" & vbCrLf
        usSql2 &= " SET SENDSTATUS='B',APPLIEDRESULT=NULL" & vbCrLf
        If s_HISREVIEW <> "" Then usSql2 &= " ,HISREVIEW=@HISREVIEW" & vbCrLf
        usSql2 &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql2 &= " WHERE OCID=@OCID AND SEQ_ID=@SEQ_ID AND CVOCID=@CVOCID" & vbCrLf
        Dim iRst As Integer = DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        If iRst > 0 Then Common.MessageBox(Me, "審核狀態 已還原，尚未確認狀態!")
    End Sub
#End Region

    ''' <summary>檔案打包下載(機構班級／機構案件)</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_PACKAGE_DOWNLOAD1.Click
        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_CVOCID.Value = TIMS.ClearSQM(Hid_CVOCID.Value)
        Hid_OCIDVal.Value = TIMS.ClearSQM(Hid_OCIDVal.Value)
        Hid_SEQ_ID.Value = TIMS.ClearSQM(Hid_SEQ_ID.Value)
        Dim vCVOCID As String = Hid_CVOCID.Value
        Dim vOCID As String = Hid_OCIDVal.Value
        Dim vSEQ_ID As String = Hid_SEQ_ID.Value
        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCID, vSEQ_ID, vCVOCID) 'If drRR Is Nothing Then Return
        If drCV Is Nothing Then Return

        ' "DOWNLOAD4" '下載
        Dim rPMS As New Hashtable From {
            {"ORGKINDGW", Hid_ORGKINDGW.Value},
            {"CVOCID", Hid_CVOCID.Value},
            {"OCID", Hid_OCIDVal.Value},
            {"SEQ_ID", Hid_SEQ_ID.Value}
        }
        Call ResponseZIPFileALL_CVO(Me, objconn, rPMS)
    End Sub

    Private Sub ResponseZIPFileALL_CVO(MyPage As Page, oConn As SqlConnection, rPMS As Hashtable)
        Const cst_UtlSubName As String = "/*ResponseZIPFileALL_CVO(MyPage As Page, oConn As SqlConnection, rPMS As Hashtable)*/"

        Dim vCVOCID As String = TIMS.GetMyValue2(rPMS, "CVOCID")
        Dim vOCID As String = TIMS.GetMyValue2(rPMS, "OCID")
        Dim vSEQ_ID As String = TIMS.GetMyValue2(rPMS, "SEQ_ID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")

        Dim drCV As DataRow = TIMS.GET_CLASS_VERIFYONLINE(objconn, vOCID, vSEQ_ID, vCVOCID) 'If drRR Is Nothing Then Return
        If drCV Is Nothing Then Return

        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(drCV("YEARS"))
        Dim vDISTID As String = Convert.ToString(drCV("DISTID"))
        Dim vDISTNAME3 As String = TIMS.GET_DISTNAME3(oConn, vDISTID)
        Dim vORGNAME As String = Convert.ToString(drCV("ORGNAME")) 'TIMS.GET_ORGNAME(Convert.ToString(drCV("ORGID")), oConn)
        Dim vCLASSCNAME2 As String = Convert.ToString(drCV("CLASSCNAME2"))

        Dim Template_ZipPath1 As String = TIMS.GET_Template_ZipPath1(vCVOCID)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath1)) Then
            Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath1))
        End If

        '查詢目前全部文件項目
        Dim dtFL As DataTable = TIMS.GET_CLASS_VERIFYONLINE_FL_TB(objconn, vCVOCID, vOCID, vSEQ_ID)

        Dim oYEARS As String = Convert.ToString(drCV("YEARS"))
        Dim oPLANID As String = Convert.ToString(drCV("PLANID"))
        Dim oCOMIDNO As String = Convert.ToString(drCV("COMIDNO"))
        Dim oSEQNO As String = Convert.ToString(drCV("SEQNO"))

        For Each drFL As DataRow In dtFL.Rows
            Dim vKVID As String = Convert.ToString(drFL("KVID"))
            Dim vKVSID As String = Convert.ToString(drFL("KVSID"))
            Dim vKBNAME2 As String = Convert.ToString(drFL("KBNAME2"))

            Dim oFILENAME1 As String = "" 'Convert.ToString(drFL("FILENAME1"))
            Dim oFILEPATH1 As String = ""
            Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
            Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
            '年度申請階段_單位名稱_項目編號+項目名稱
            Dim t_FILENAME As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf")
            Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))
            'Dim t_FilePath1 As String=Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
            Try
                oFILENAME1 = Convert.ToString(drFL("FILENAME1"))
                oFILEPATH1 = Convert.ToString(drFL("FILEPATH1"))
                oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH1_CVO(oYEARS, oPLANID, oCOMIDNO, oSEQNO, vCVOCID, ""))
                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                '年度申請階段_單位名稱_項目編號+項目名稱
                t_FILENAME = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, "_", vORGNAME, "_", vCLASSCNAME2, "_", vKBNAME2, ".pdf"))
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
        Next

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim filenames As String() = Directory.GetFiles(MyPage.Server.MapPath(String.Concat(Template_ZipPath1, "/")))
        Dim zipFileName As String = TIMS.GetValidFileName(String.Concat("CV", vYEARS_ROC, vDISTNAME3, "_", vORGNAME, "_", vOCID, "_", vCVOCID, "_", strNOW, ".zip"))
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
                    Dim strErrmsg As String = "/*ResponseZIPFileALL_CVO*/" & vbCrLf
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
            TIMS.SAVE_ADP_ZIPFILE(oConn, "-tc13002cvo", File)
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
            TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
        End With

    End Sub

End Class
