Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Partial Class SD_19_002
    Inherits AuthBasePage

    Dim tryFIND As String = ""

    Const cst_MaxLen500_i As Integer = 500
    Const cst_MaxLen500_TPMSG1 As String = "限定500字元"

    '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
    'Dim vAPPLIEDRESULT As String = ""

    'Dim iDG11_ROWS As Integer = 0
    'Dim iDG10_ROWS As Integer = 0

    Const G06_其他補充資料 As String = "G06"
    Const W06_其他補充資料 As String = "W06"
    Dim iDG06_ROWS As Integer = 0
    Const cst_txt_其他補充資料 As String = "(其他補充資料)"
    Const cst_06_其他補充資料_WAIVED_OTH1 As String = "OTH1"

    '以目前版本批次送出
    Const cst_txt_版本批次送出 As String = "(版本批次送出)"
    Const cst_txt_免附文件 As String = "(免附文件)"

    Const cst_ss_RqProcessType As String = "RqProcessType"
    'Const cst_DG1CMDNM_VIEW1 As String = "VIEW1"
    Const cst_DG1CMDNM_EDIT1 As String = "EDIT1" '審核/審查/確認
    Const cst_DG1CMDNM_REVERT2 As String = "REVERT2" '還原"-確認

    Const cst_DG2CMDNM_RtuBACK1 As String = "RtuBACK1" '退回開放修改
    Const cst_DG2CMDNM_REVERT1 As String = "REVERT1" '還原"
    Const cst_DG2CMDNM_VIEWFILE4 As String = "VIEWFILE4" '查詢
    Const cst_DG2CMDNM_DOWNLOAD4 As String = "DOWNLOAD4" '下載

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
    Dim objconn As SqlConnection = Nothing

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        '線上送件審核-還原按鈕 NULL/Y:可使用
        USE_CLASS_STD14OA_REVERT2.Value = TIMS.Utl_GetConfigVAL(objconn, "USE_CLASS_STD14OA_REVERT2")

        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        If Not IsPostBack Then
            Call CCreate1(0)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button3.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True)
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If
    End Sub

    '設定 資料與顯示 狀況！
    Private Sub CCreate1(ByVal iNum As Integer)
        TB_DataGrid1.Visible = False
        labmsg1.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        Call SHOW_Frame1(0)
        ''案件編號
        'sch_txtTBCASENO.Text = ""
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
        'sch_txtBINAME.Text = ""
        ''申辦日期
        'sch_txtBIDATE1.Text = ""
        'sch_txtBIDATE2.Text = ""

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

        '清理隱藏的參數
        Call SEARCH_1()
    End Sub

    Sub SEARCH_1()
        '清理隱藏的參數
        Call ClearHidValue()

        Call SHOW_Frame1(0)
        labmsg1.Text = TIMS.cst_NODATAMsg1
        TB_DataGrid1.Visible = False

        'RIDValue.Value = If(RIDValue.Value <> "", RIDValue.Value, sm.UserInfo.RID)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS) '計畫年度
        Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE) '申請階段
        sch_txtTBCNAME.Text = TIMS.ClearSQM(sch_txtTBCNAME.Text)
        sch_txtTBCDATE1.Text = TIMS.Cdate3(sch_txtTBCDATE1.Text)
        sch_txtTBCDATE2.Text = TIMS.Cdate3(sch_txtTBCDATE2.Text)
        '檢核日期順序 異常:TRUE 執行對調
        If TIMS.ChkDateErr3(sch_txtTBCDATE1.Text, sch_txtTBCDATE2.Text) Then
            Dim T_DATE1 As String = sch_txtTBCDATE1.Text
            sch_txtTBCDATE1.Text = sch_txtTBCDATE2.Text
            sch_txtTBCDATE2.Text = T_DATE1
        End If
        '檢核查詢
        If RIDValue.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(查無業務代碼)，請選擇訓練機構!!")
            Return
        End If
        Dim vDISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        Dim PMS1 As New Hashtable 'From {{"YEARS", $"{sm.UserInfo.Years}"}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim SQL1 As String = "
SELECT a.TBCID,a.TBCASENO,a.OCID,a.APPSTAGE
,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N
,a.TBCACCT,u.NAME TBCNAME,a.TBCDATE,dbo.FN_CDATE1B(a.TBCDATE) TBCDATE_ROC
,a.TBCSTATUS, CASE WHEN a.TBCSTATUS IS NULL THEN '暫存'
WHEN a.TBCSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '分署已收件'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END TBCSTATUS_N
,a.APPLIEDRESULT,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N
,cc.DISTNAME,cc.ORGNAME,cc.CLASSCNAME2,cc.ORGKINDGW
,cc.YEARS,cc.TPLANID,cc.DISTID,dbo.FN_CYEAR2(cc.YEARS) YEARS_ROC
FROM CLASS_STD14OA a
JOIN AUTH_ACCOUNT u ON u.ACCOUNT=a.TBCACCT
JOIN VIEW2 cc on cc.OCID=a.OCID
WHERE cc.YEARS=@YEARS AND cc.TPLANID=@TPLANID AND a.TBCSTATUS IS NOT NULL 
"
        'AND a.APPSTAGE=@APPSTAGE AND cc.DISTID=@DISTID
        PMS1.Add("YEARS", v_sch_ddlYEARS)
        PMS1.Add("TPLANID", sm.UserInfo.TPlanID)
        'PMS1.Add("APPSTAGE", v_sch_ddlAPPSTAGE)
        'PMS1.Add("DISTID", vDISTID)
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
        If v_sch_ddlAPPSTAGE <> "" Then
            PMS1.Add("APPSTAGE", v_sch_ddlAPPSTAGE)
            SQL1 &= " AND a.APPSTAGE=@APPSTAGE"
        End If
        If sch_txtTBCNAME.Text <> "" Then
            PMS1.Add("TBCNAME", sch_txtTBCNAME.Text)
            SQL1 &= " AND u.NAME like '%'+@TBCNAME+'%'" & vbCrLf
        End If
        If sch_txtTBCDATE1.Text <> "" Then
            PMS1.Add("TBCDATE1", TIMS.Cdate2(sch_txtTBCDATE1.Text))
            SQL1 &= " AND A.TBCDATE>=@TBCDATE1" & vbCrLf
        End If
        If sch_txtTBCDATE2.Text <> "" Then
            PMS1.Add("TBCDATE2", TIMS.Cdate2(sch_txtTBCDATE2.Text))
            SQL1 &= " AND A.TBCDATE<=@TBCDATE2" & vbCrLf
        End If
        ' a.TBCSTATUS WHEN 'B' THEN '已送件' WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過
        'TBCSTATUS: 申辦狀態：NULL:暫存/B:已送件/R:退件修正
        'APPLIEDRESULT: 審查狀態：Y:申辦確認/R:申辦退件修正/N:申辦不通過
        Dim v_rbAPPLIEDRESULT As String = TIMS.GetListValue(rbAPPLIEDRESULT)
        If v_rbAPPLIEDRESULT = "A" Then
            SQL1 &= " AND a.TBCSTATUS IN ('R','B')" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "B" Then
            SQL1 &= " AND a.APPLIEDRESULT IS NULL" & vbCrLf
            SQL1 &= " AND a.TBCSTATUS='B'" & vbCrLf
            'sSql &= " AND (a.APPLIEDRESULT='R' OR a.APPLIEDRESULT IS NULL)" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "R" Then
            SQL1 &= " AND a.APPLIEDRESULT='R'" & vbCrLf
            SQL1 &= " AND (a.TBCSTATUS='R' OR a.TBCSTATUS='B')" & vbCrLf
        ElseIf v_rbAPPLIEDRESULT = "Y" OrElse v_rbAPPLIEDRESULT = "N" Then 'Y:申辦確認 / N:申辦不通過
            PMS1.Add("APPLIEDRESULT", v_rbAPPLIEDRESULT)
            SQL1 &= " AND a.APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        End If
        SQL1 &= " ORDER BY A.TBCID DESC" & vbCrLf

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

    ''' <summary>清理隱藏的參數</summary>
    Sub ClearHidValue()
        Hid_TBCID.Value = ""
        Hid_TBCASENO.Value = ""
        Hid_ORGKINDGW.Value = ""
        'Hid_FirstKTSEQ.Value = ""
        'Hid_LastKTID.Value = ""
        Hid_RID.Value = ""
        Hid_APPLIEDRESULT.Value = ""
        'Hid_KTSEQ.Value = ""
        'Hid_KTID.Value = ""
        'Hid_TBCFID.Value = ""
    End Sub

    '分署打包下載
    'Protected Sub BTN_TRAIN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_TRAIN_PACKAGE_DOWNLOAD1.Click
    '    '清理隱藏的參數
    '    Call ClearHidValue()

    '    If (sm.UserInfo.LID > 1) Then
    '        Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
    '        Return
    '    End If

    '    Dim v_sch_ddlYEARS As String = TIMS.GetListValue(sch_ddlYEARS) '計畫年度
    '    Dim v_sch_ddlAPPSTAGE As String = TIMS.GetListValue(sch_ddlAPPSTAGE) '申請階段
    '    Hid_RID.Value = ""
    '    If Hid_RID.Value = "" AndAlso RIDValue.Value.Length > 0 Then Hid_RID.Value = RIDValue.Value.Substring(0, 1)
    '    If Hid_RID.Value = "" AndAlso sm.UserInfo.LID = 1 Then Hid_RID.Value = sm.UserInfo.RID
    '    If Hid_RID.Value.Length > 1 Then Hid_RID.Value = Hid_RID.Value.Substring(0, 1)
    '    Dim vDISTID As String = TIMS.Get_DistID_RID(Hid_RID.Value, objconn)

    '    If v_sch_ddlYEARS = "" OrElse v_sch_ddlAPPSTAGE = "" Then
    '        Common.MessageBox(Me, "計畫年度與申請階段為必選!")
    '        Return
    '    ElseIf Hid_RID.Value = "" OrElse vDISTID = "" Then
    '        Common.MessageBox(Me, "訓練機構為必選!")
    '        Return
    '    ElseIf vDISTID = "" Then
    '        Common.MessageBox(Me, "轄區分署有誤，請重新選擇!")
    '        Return
    '    End If

    '    'DOWNLOAD '下載 分署打包下載 (單位有上傳就算)
    '    Dim rPMS As New Hashtable
    '    rPMS.Add("DISTID", vDISTID)
    '    rPMS.Add("YEARS", v_sch_ddlYEARS)
    '    rPMS.Add("APPSTAGE", v_sch_ddlAPPSTAGE)
    '    '打包 - 分署打包下載 (單位有上傳就算)
    '    Dim dtOAW As DataTable = GET_CLASS_STD14OA_WAIVED(objconn, rPMS)
    '    If TIMS.dtNODATA(dtOAW) Then
    '        Common.MessageBox(Me, String.Concat("【分署打包下載】", TIMS.cst_NODATAMsg1))
    '        Return
    '    End If
    '    '分署打包下載
    '    Call ResponseZIPFileOA(Me, objconn, rPMS, dtOAW)
    'End Sub

    '分署打包下載
    'Private Sub ResponseZIPFileOA(MyPage As Page, oConn As SqlConnection, rPMS As Hashtable, dtOAW As DataTable)
    '    Const cst_UtlSubName As String = "/*ResponseZIPFileOA*/"
    '    Dim vDISTID As String = TIMS.GetMyValue2(rPMS, "DISTID")
    '    Dim vYEARS As String = TIMS.GetMyValue2(rPMS, "YEARS")
    '    Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS, "APPSTAGE")
    '    Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(vYEARS)
    '    Dim vAPPSTAGE_S As String = TIMS.Get_APPSTAGE_S(vAPPSTAGE)
    '    Dim vDISTNAME As String = TIMS.GET_DISTNAME(oConn, vDISTID)

    '    '打包 - 分署打包下載 (單位有上傳就算)
    '    'Dim dtOAW As DataTable = GET_CLASS_STD14OA_WAIVED(objconn, rPMS)
    '    If TIMS.dtNODATA(dtOAW) Then Return

    '    Dim Template_ZipPath1 As String = TIMS.GET_Template_ZipPath1(vDISTID)
    '    '判斷是否有資料夾
    '    If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath1)) Then
    '        Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath1))
    '    End If

    '    For Each drOP As DataRow In dtOAW.Rows
    '        Dim vTBCID As String = $"{drOP("TBCID")}"
    '        Dim vTBCFID As String = $"{drOP("TBCFID")}"
    '        Dim vORGNAME As String = $"{drOP("ORGNAME")}"
    '        Dim oYEARS As String = $"{drOP("YEARS")}"
    '        Dim oAPPSTAGE As String = $"{drOP("APPSTAGE")}"
    '        Dim oTBCASENO As String = $"{drOP("TBCASENO")}"
    '        Dim vKTSEQ As String = $"{drOP("KTSEQ")}"
    '        Dim vORGKINDGW As String = $"{drOP("ORGKINDGW")}"
    '        Dim vKTID As String = $"{drOP("KTID")}"
    '        Dim vKTNAME2x As String = String.Concat(vORGKINDGW, vKTID, drOP("KTNAME"))
    '        'Dim vPSNO28 As String = $"{drOP("PSNO28")}"
    '        Dim oFILENAME1 As String = $"{drOP("FILENAME1")}"
    '        Dim oFILEPATH1 As String = $"{drOP("FILEPATH1")}"
    '        Dim vPLANID As String = $"{drOP("PLANID")}"
    '        Dim vTBCASENO As String = $"{drOP("TBCASENO")}"

    '        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
    '        Dim s_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
    '        Dim t_FILENAME_OA As String = ""
    '        Dim t_FilePath1 As String = "" 'MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_PI))
    '        Try
    '            'vPSNO28 = Convert.ToString(drFLPI("PSNO28"))
    '            'oFILENAME1 = Convert.ToString(drFLPI("FILENAME1"))
    '            oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH1_OA(oYEARS, vDISTID, vPLANID, vTBCASENO, ""))
    '            s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
    '            t_FILENAME_OA = TIMS.GetValidFileName($"{vYEARS_ROC}{vAPPSTAGE_S}_{vORGNAME}_{vTBCASENO}_{vKTNAME2x}.pdf")
    '            t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME_OA))
    '            Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
    '            If IO.File.Exists(s_FilePath1) Then
    '                Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
    '                File.WriteAllBytes(t_FilePath1, dbyte)
    '            End If
    '        Catch ex As Exception
    '            Dim strErrmsg As String = String.Concat(cst_UtlSubName, vbCrLf)
    '            strErrmsg &= String.Concat("t_FILENAME_OA: ", t_FILENAME_OA, vbCrLf)
    '            strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
    '            strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
    '            strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
    '            Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
    '        End Try
    '    Next

    '    Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
    '    Dim zipFileName As String = String.Concat("p", vYEARS_ROC, vAPPSTAGE_S, "_", vDISTNAME, "_", strNOW, ".zip")
    '    Dim filenames As String() = Directory.GetFiles(MyPage.Server.MapPath(String.Concat(Template_ZipPath1, "/")))
    '    Dim full_zipFileName As String = String.Concat(Template_ZipPath1, "/", zipFileName)
    '    Using zp As New ZipOutputStream(System.IO.File.Create(MyPage.Server.MapPath(full_zipFileName)))
    '        zp.SetLevel(6) ' 設定壓縮比
    '        ' 逐一將資料夾內的檔案抓出來壓縮，並寫入至目的檔(.ZIP)
    '        For Each filename As String In filenames
    '            Dim entry As New ZipEntry(Path.GetFileName(filename)) With {.IsUnicodeText = True}
    '            zp.PutNextEntry(entry) '建立下一個壓縮檔案或資料夾條目
    '            Try
    '                Using fs As New FileStream(filename, FileMode.Open)
    '                    Dim buffer As Byte() = New Byte(fs.Length - 1) {}
    '                    Dim i_readLength As Integer
    '                    Do
    '                        i_readLength = fs.Read(buffer, 0, buffer.Length)
    '                        If i_readLength > 0 Then zp.Write(buffer, 0, i_readLength)
    '                    Loop While i_readLength > 0
    '                End Using
    '            Catch ex As Exception
    '                Dim strErrmsg As String = cst_UtlSubName & vbCrLf
    '                strErrmsg &= String.Concat("full_zipFileName: ", full_zipFileName, vbCrLf)
    '                strErrmsg &= String.Concat("filename: ", filename, vbCrLf)
    '                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
    '                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
    '                Common.MessageBox(MyPage, String.Concat("檔案下載有誤，請重新操作!", ex.Message))
    '                Return
    '            End Try
    '            '假設處理某段程序需花費1毫秒 (避免機器不同步)
    '            Threading.Thread.Sleep(1)
    '            '刪除檔案
    '            Call TIMS.MyFileDelete(filename)
    '        Next
    '    End Using

    '    With MyPage
    '        Dim File As New FileInfo(.Server.MapPath(full_zipFileName))
    '        TIMS.SAVE_ADP_ZIPFILE(oConn, "-sd19002oa", File)
    '        ' Clear the content of the response
    '        .Response.ClearContent()
    '        ' LINE1 Add the file name And attachment, which will force the open/cance/save dialog To show, to the header
    '        .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
    '        'Response.Headers["Content-Disposition"] = "attachment; filename=" + zipFileName;
    '        ' Add the file size into the response header
    '        .Response.AddHeader("Content-Length", File.Length.ToString())
    '        ' Set the ContentType
    '        .Response.ContentType = "application/zip"
    '        .Response.TransmitFile(File.FullName)
    '        ' End the response
    '        TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
    '    End With

    'End Sub

    ''' <summary>dtOAW-分署打包下載</summary>
    ''' <param name="oConn"></param>
    ''' <param name="rPMS2"></param>
    ''' <returns></returns>
    Private Function GET_CLASS_STD14OA_WAIVED(oConn As SqlConnection, rPMS2 As Hashtable) As DataTable
        Dim vDISTID As String = TIMS.GetMyValue2(rPMS2, "DISTID")
        Dim vYEARS As String = TIMS.GetMyValue2(rPMS2, "YEARS")
        Dim vAPPSTAGE As String = TIMS.GetMyValue2(rPMS2, "APPSTAGE")
        Dim sParms As New Hashtable
        sParms.Add("DISTID", vDISTID)
        sParms.Add("YEARS", vYEARS)
        sParms.Add("APPSTAGE", vAPPSTAGE)
        Dim sSql As String = "
SELECT a.TBCID,a.TBCASENO,cc.YEARS,cc.DISTID,cc.ORGID,cc.RID,a.APPSTAGE
,cc.ORGNAME,cc.PLANID
,a.TBCACCT,a.TBCDATE,a.TBCSTATUS
,a.APPLIEDRESULT,a.REASONFORFAIL
,a.RESULTACCT,a.RESULTDATE,a.HISREVIEW
,b.TBCFID,b.KTSEQ,c.KTID,c.KTNAME,c.ORGKINDGW
,b.FILENAME1,b.FILEPATH1,b.SRCFILENAME1
,b.WAIVED,b.RTUREASON,b.RTURESACCT,b.RTURESDATE
FROM CLASS_STD14OA a
JOIN VIEW2 cc on cc.OCID=a.OCID
JOIN CLASS_STD14OAFL b on b.TBCID=a.TBCID
JOIN KEY_STD14TH c on c.KTSEQ =b.KTSEQ
WHERE a.TBCSTATUS='B'
AND cc.YEARS=@YEARS AND cc.DISTID=@DISTID AND a.APPSTAGE=@APPSTAGE
"
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, oConn, sParms)
        Return dt
    End Function

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        '清理隱藏的參數
        Call ClearHidValue()

        Dim sCmdArg As String = e.CommandArgument
        Dim vOCID As String = TIMS.GetMyValue(sCmdArg, "OCID")
        Dim vTBCID As String = TIMS.GetMyValue(sCmdArg, "TBCID")
        Dim vTBCASENO As String = TIMS.GetMyValue(sCmdArg, "TBCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue(sCmdArg, "ORGKINDGW")
        If sCmdArg = "" OrElse vTBCID = "" OrElse vTBCASENO = "" OrElse vOCID = "" OrElse vORGKINDGW = "" Then Return

        Hid_TBCID.Value = vTBCID
        Hid_TBCASENO.Value = vTBCASENO
        Hid_ORGKINDGW.Value = vORGKINDGW
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無線上申辦案件)!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{vOCID}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)!")
            Return
        End If

        Select Case e.CommandName
            Case cst_DG1CMDNM_EDIT1 '"EDIT1"審核
                'If Not fg_can_applistage Then
                '    Common.MessageBox(Me, "申請階段受理期間未開放，請確認後再操作!")
                '    Return
                'End If
                If $"{drOA("TBCSTATUS")}" = "R" Then
                    Common.MessageBox(Me, "申辦狀態 退件待修正，待修正後送審，再行審核!")
                    Return
                ElseIf $"{drOA("TBCSTATUS")}" = "" Then
                    Common.MessageBox(Me, "申辦狀態 未填寫，待送審後，再行審核!")
                    Return
                End If
                '查詢使用資料顯示 依 ORG_BIDCASE-BCID
                Call SHOW_Detail_STD14OA(drCC, vTBCID, cst_DG1CMDNM_EDIT1) '"EDIT1"審核

            Case cst_DG1CMDNM_REVERT2 '"REVERT2" '還原"-確認
                '當【審核狀態】：已申辦確認、已申辦不通過，按下【還原】鈕即清空【審核狀態】且【申辦狀態】：已送件， 如圖二狀態
                '當【審核狀態】：退件修正，【還原】鈕反灰不可按
                If $"{drOA("APPLIEDRESULT")}" = "R" Then
                    Common.MessageBox(Me, "審核狀態 退件待修正，不可還原!")
                    Return
                ElseIf Convert.ToString(drOA("APPLIEDRESULT")) = "" Then
                    Common.MessageBox(Me, "審核狀態 未填寫，不可還原!")
                    Return
                End If

                Const cst_REVERT2_N As String = "還原"
                Dim s_HISREVIEW As String = $"{drOA("HISREVIEW")}"
                s_HISREVIEW &= $"{If(s_HISREVIEW <> "", "，", "")}{TIMS.Cdate3t(Now)}-{cst_REVERT2_N}"
                '線上送件審核，新增還原按鈕 UPDATE ORG_BIDCASE / BISTATUS='B',APPLIEDRESULT=NULL
                Call UPDATE_REVERT2(s_HISREVIEW, vTBCID, vTBCASENO)
                '查詢1
                Call SEARCH_1()

        End Select
    End Sub

    ''' <summary>新增使用資料顯示／查詢使用資料顯示 依 ORG_BIDCASE-BCID</summary>
    Private Sub SHOW_Detail_STD14OA(drCC As DataRow, V_TBCID As String, vCmdName As String)
        '訓練機構有誤
        If drCC Is Nothing Then Return
        Call SHOW_Frame1(1)
        If V_TBCID = "" Then
            Common.MessageBox(Me, "傳入參數為空，異常!")
            Return
        End If

        Session(cst_ss_RqProcessType) = vCmdName

        tr_HISREVIEW.Visible = False '歷程資訊

        '查詢資料 'NULL 待送審、B 審核中、Y 審核通過、R 退件修正、N 審核不通過。
        Dim pParms As New Hashtable From {{"TBCID", V_TBCID}, {"OCID", $"{drCC("OCID")}"}}

        Dim sSql As String = "
SELECT a.TBCID,a.TBCASENO,a.OCID,a.APPSTAGE,a.TBCACCT,a.TBCDATE,a.TBCSTATUS
,CASE a.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N
,cc.YEARS,cc.DISTNAME,cc.ORGNAME,cc.COMIDNO,cc.ORGKINDGW,cc.CLASSCNAME2
,a.TBCACCT,u.NAME TBCNAME
,a.TBCDATE,dbo.FN_CDATE1B(a.TBCDATE) TBCDATE_ROC
/* '申辦狀態：暫存/ 已送件 */
,a.TBCSTATUS, CASE WHEN a.TBCSTATUS IS NULL THEN '暫存'
WHEN a.TBCSTATUS='R' AND a.APPLIEDRESULT='R' THEN '退件待修正'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='R' THEN '修正再送審'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='Y' THEN '通過'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT='N' THEN '不通過'
WHEN a.TBCSTATUS='B' AND a.APPLIEDRESULT IS NULL THEN '已送件' END TBCSTATUS_N
,a.APPLIEDRESULT,a.REASONFORFAIL
,CASE a.APPLIEDRESULT WHEN 'Y' THEN '申辦確認' WHEN 'R' THEN '申辦退件修正' WHEN 'N' THEN '申辦不通過' END APPLIEDRESULT_N
,a.HISREVIEW
,(SELECT MIN(KTSEQ) FROM CLASS_STD14OAFL fl WHERE fl.TBCID=a.TBCID) CurrentKTSEQ
/* ,a.CREATEACCT,a.CREATEDATE,a.MODIFYACCT,a.MODIFYDATE */
,a.APPLIEDRESULT,a.REASONFORFAIL
FROM CLASS_STD14OA a
JOIN AUTH_ACCOUNT u ON u.ACCOUNT=a.TBCACCT
JOIN VIEW2 cc on cc.OCID=a.OCID
WHERE TBCID=@TBCID AND a.OCID=@OCID 
"
        Dim dtB1 As DataTable = DbAccess.GetDataTable(sSql, objconn, pParms)

        If TIMS.dtNODATA(dtB1) Then
            Common.MessageBox(Me, "查無有效資料，異常!")
            Return
        End If

        Dim drB1 As DataRow = dtB1.Rows(0)

        Dim vAPPSTAGE As String = $"{drB1("APPSTAGE")}"

        Common.SetListItem(ddlAPPLIEDRESULT, $"{drB1("APPLIEDRESULT")}") ' Convert.ToString(drB1("APPLIEDRESULT")))
        Reasonforfail.Text = Convert.ToString(drB1("REASONFORFAIL"))
        TIMS.Tooltip(Reasonforfail, cst_MaxLen500_TPMSG1, True)

        Hid_TBCID.Value = $"{drB1("TBCID")}"
        Hid_TBCASENO.Value = $"{drB1("TBCASENO")}"
        Hid_ORGKINDGW.Value = $"{drB1("ORGKINDGW")}"

        'labAPPSTAGE,labCLASSNAME2S,BTN_PACKAGE_DOWNLOAD1,檔案打包下載,PACKAGE_DOWNLOAD1,tr_HISREVIEW,labHISREVIEW,></asp:Label>,
        labOrgNAME.Text = $"{drB1("ORGNAME")}"
        labTBCASENO.Text = $"{drB1("TBCASENO")}"
        LabCREATEDATE.Text = TIMS.Cdate17t1($"{drB1("TBCDATE")}")
        labBIYEARS.Text = TIMS.GET_YEARS_ROC(drB1("YEARS"))
        labAPPSTAGE.Text = $"{drB1("APPSTAGE_N")}"
        labCLASSNAME2S.Text = $"{drB1("CLASSCNAME2")}" 'CLASSCNAME2
        'Hid_KTSEQ.Value = $"{drB1("CurrentKTSEQ")}"
        Hid_APPLIEDRESULT.Value = $"{drB1("APPLIEDRESULT")}"
        Dim fg_have_APPLIEDRESULT As Boolean = (Hid_APPLIEDRESULT.Value = "N" OrElse Hid_APPLIEDRESULT.Value = "Y")
        Dim t_ddlAPPLIEDRESULT As String = If(fg_have_APPLIEDRESULT, TIMS.GetListText(ddlAPPLIEDRESULT), "")
        Reasonforfail.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        ddlAPPLIEDRESULT.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        But_Sub.Enabled = If(Not fg_have_APPLIEDRESULT, True, False)
        TIMS.Tooltip(Reasonforfail, t_ddlAPPLIEDRESULT, True)
        TIMS.Tooltip(ddlAPPLIEDRESULT, t_ddlAPPLIEDRESULT, True)
        TIMS.Tooltip(But_Sub, t_ddlAPPLIEDRESULT, True)

        '歷程資訊
        If $"{drB1("HISREVIEW")}" <> "" AndAlso $"{drB1("HISREVIEW")}".Length > 1 Then
            tr_HISREVIEW.Visible = True '歷程資訊
            labHISREVIEW.Text = Convert.ToString(drB1("HISREVIEW"))
        End If

        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)
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

    Private Sub SHOW_STD14OAFL_DG2(rPMS As Hashtable)
        labmsg1.Text = ""
        td_title06.Visible = False
        tr_DataGrid06.Visible = False
        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")
        Dim fg_CANSAVE As Boolean = (vORGKINDGW = "G" OrElse vORGKINDGW = "W")
        'objconn 因為有檔案輸出關閉的問題 所以要檢查
        If Not TIMS.OpenDbConn(objconn) OrElse Not fg_CANSAVE Then Return

        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        If vTBCID = "" OrElse Hid_TBCID.Value = "" OrElse Hid_TBCASENO.Value = "" Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!(1)")
            Return
        End If
        If vTBCID <> Hid_TBCID.Value Then
            Common.MessageBox(Me, "資訊有誤(案件編號有誤)，請重新操作!(2)")
            Return
        End If
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!(3)")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無職類/班別代碼)!(4)")
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
        'DataGrid2.Columns(cst_DG2_退件原因_iCOLUMN).Visible = If($"{drOA("APPLIEDRESULT")}" = "R", True, False)
        DataGrid2.DataSource = dtFL
        DataGrid2.DataBind()

    End Sub

    ''' <summary>線上送件審核，還原按鈕 UPDATE CLASS_STD14OA, TBCSTATUS='B',APPLIEDRESULT=NULL </summary>
    ''' <param name="s_HISREVIEW"></param>
    ''' <param name="vTBCID"></param>
    ''' <param name="vTBCASENO"></param>
    Sub UPDATE_REVERT2(s_HISREVIEW As String, vTBCID As String, vTBCASENO As String)
        vTBCID = TIMS.ClearSQM(vTBCID)
        vTBCASENO = TIMS.ClearSQM(vTBCASENO)
        If vTBCID = "" OrElse vTBCASENO = "" Then Return

        '退件修正 
        Dim USSQL2_W1 As String = ""
        Dim uParms2 As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"TBCID", vTBCID}, {"TBCASENO", vTBCASENO}}
        If s_HISREVIEW <> "" Then
            uParms2.Add("HISREVIEW", s_HISREVIEW)
            USSQL2_W1 = " ,HISREVIEW=@HISREVIEW"
        End If
        Dim usSql2 As String = ""
        usSql2 &= $" UPDATE CLASS_STD14OA SET TBCSTATUS='B',APPLIEDRESULT=NULL{USSQL2_W1},MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql2 &= " WHERE TBCID=@TBCID AND TBCASENO=@TBCASENO" & vbCrLf
        Dim iRst As Integer = DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        If iRst > 0 Then Common.MessageBox(Me, "審核狀態 已還原，尚未確認狀態!")
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem,
                Dim drv As DataRowView = e.Item.DataItem
                Dim lBTN_EDIT1 As LinkButton = e.Item.FindControl("lBTN_EDIT1") '審核/審查/確認
                Dim lBTN_REVERT2 As LinkButton = e.Item.FindControl("lBTN_REVERT2") '還原
                '線上送件審核-還原按鈕 NULL/Y:可使用
                lBTN_REVERT2.Visible = (USE_CLASS_STD14OA_REVERT2.Value = "Y")

                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OCID", $"{drv("OCID")}")
                TIMS.SetMyValue(sCmdArg, "TBCID", $"{drv("TBCID")}") ' drv("TBCID"))
                TIMS.SetMyValue(sCmdArg, "TBCASENO", $"{drv("TBCASENO")}") ' drv("TBCASENO"))
                TIMS.SetMyValue(sCmdArg, "ORGKINDGW", $"{drv("ORGKINDGW")}") 'drv("ORGKINDGW"))

                Dim s_APPLIEDRESULT As String = $"{drv("APPLIEDRESULT")}"
                Dim tipMsg1 As String = $"{drv("APPLIEDRESULT_N")}"
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
    Private Sub DataGrid2_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Dim txtRtuReason As TextBox = e.Item.FindControl("txtRtuReason") '退回原因說明
        Dim Btn_RtuBACK1 As Button = e.Item.FindControl("Btn_RtuBACK1") '退回開放修改
        Dim Btn_REVERT1 As Button = e.Item.FindControl("Btn_REVERT1") '還原
        Dim sCmdArg As String = e.CommandArgument
        Dim vTBCID As String = TIMS.GetMyValue(sCmdArg, "TBCID")
        'Dim vTBCASENO As String = TIMS.GetMyValue(sCmdArg, "TBCASENO")
        Dim vTBCFID As String = TIMS.GetMyValue(sCmdArg, "TBCFID")
        Dim vKTID As String = TIMS.GetMyValue(sCmdArg, "KTID")
        Dim vKTSEQ As String = TIMS.GetMyValue(sCmdArg, "KTSEQ")
        Dim vFILENAME1 As String = TIMS.GetMyValue(sCmdArg, "FILENAME1")
        Dim vFILEPATH1 As String = TIMS.GetMyValue(sCmdArg, "FILEPATH1")
        If e.CommandArgument = "" OrElse vTBCID = "" OrElse vTBCFID = "" OrElse vKTID = "" OrElse vKTSEQ = "" Then Return
        '顯示檔案資料表
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)

        If txtRtuReason IsNot Nothing Then txtRtuReason.Text = TIMS.ClearSQM(txtRtuReason.Text)

        Dim drFL As DataRow = TIMS.GET_CLASS_STD14OAFL(objconn, vTBCID, vKTSEQ)
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
                uParms.Add("TBCFID", TIMS.CINT1(vTBCFID))
                uParms.Add("TBCID", TIMS.CINT1(vTBCID))
                uParms.Add("RTUREASON", If(v_txtRtuReason <> "", v_txtRtuReason, Convert.DBNull))
                uParms.Add("RTURESACCT", sm.UserInfo.UserID)
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_STD14OAFL" & vbCrLf
                usSql &= " SET RTUREASON=@RTUREASON,RTURESACCT=@RTURESACCT,RTURESDATE=GETDATE()" & vbCrLf
                usSql &= " WHERE TBCFID=@TBCFID AND TBCID=@TBCID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Case cst_DG2CMDNM_REVERT1 '還原
                'txtRtuReason.Text = TIMS.ClearSQM(txtRtuReason.Text)
                Dim uParms As New Hashtable
                uParms.Add("TBCFID", TIMS.CINT1(vTBCFID))
                uParms.Add("TBCID", TIMS.CINT1(vTBCID))
                Dim usSql As String = ""
                usSql &= " UPDATE CLASS_STD14OAFL" & vbCrLf
                usSql &= " SET RTUREASON=NULL,RTURESACCT=NULL,RTURESDATE=NULL" & vbCrLf
                usSql &= " WHERE TBCFID=@TBCFID AND TBCID=@TBCID" & vbCrLf
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

            Case cst_DG2CMDNM_VIEWFILE4
                '"VIEWFILE4" '查詢

            Case cst_DG2CMDNM_DOWNLOAD4
                Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, vTBCID, Hid_TBCASENO.Value)
                If drOA Is Nothing Then Return
                ' "DOWNLOAD4" '下載
                Dim rPMS4 As New Hashtable
                TIMS.SetMyValue2(rPMS4, "ORGKINDGW", Hid_ORGKINDGW.Value)
                TIMS.SetMyValue2(rPMS4, "TBCID", vTBCID)
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

        Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        Call SHOW_STD14OAFL_DG2(rPMS3)
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
                    titleMsg = $"{drv("FILENAME1")}"
                    'LabFileName1.Text = If(Convert.ToString(drv("FILENAME1")) = Convert.ToString(drv("OKFLAG")), Convert.ToString(drv("FILENAME1")), Convert.ToString(drv("OKFLAG")))
                    'HFileName.Value = Convert.ToString(drv("FILENAME1")) '.ToString()
                ElseIf $"{drv("WAIVED")}" = "Y" Then
                    titleMsg = cst_txt_免附文件
                    BTN_DOWNLOAD4.Enabled = False
                    'LabFileName1.Text = cst_txt_免附文件
                ElseIf $"{drv("WAIVED")}" = cst_06_其他補充資料_WAIVED_OTH1 Then
                    'Dim vKTSEQ As String = $"{drv("KTSEQ")}"
                    SHOW_DATAGRID_06($"{drv("KTSEQ")}")
                End If
                If titleMsg <> "" Then TIMS.Tooltip(BTN_DOWNLOAD4, titleMsg, True)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "TBCID", drv("TBCID"))
                'TIMS.SetMyValue(sCmdArg, "TBCASENO", drv("TBCASENO"))
                TIMS.SetMyValue(sCmdArg, "TBCFID", drv("TBCFID"))
                TIMS.SetMyValue(sCmdArg, "KTID", drv("KTID"))
                TIMS.SetMyValue(sCmdArg, "KTSEQ", drv("KTSEQ"))
                TIMS.SetMyValue(sCmdArg, "FILENAME1", drv("FILENAME1"))
                TIMS.SetMyValue(sCmdArg, "FILEPATH1", drv("FILEPATH1"))

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

    Protected Sub But_BACK1_Click(sender As Object, e As EventArgs) Handles But_BACK1.Click
        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        '清理隱藏的參數
        Call SEARCH_1()
    End Sub

    Protected Sub But_Sub_Click(sender As Object, e As EventArgs) Handles But_Sub.Click
        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        Dim sErrMsg1 As String = ""
        'ddlAPPLIEDRESULT:／審核通過:Y／審核不通過:N／退件修正:R
        Dim v_ddlAPPLIEDRESULT As String = TIMS.GetListValue(ddlAPPLIEDRESULT)
        Reasonforfail.Text = TIMS.ClearSQM(Reasonforfail.Text)

        If v_ddlAPPLIEDRESULT = "" Then sErrMsg1 &= "請選擇，審查狀態結果!" & vbCrLf
        '(6) 下方審核結果系統檢核， 當上面文件項目有任一個是有點選退回開放修改，【審核結果】就必須要整份退件修正。 有值:true 無值:false
        Dim flag_NG1 As Boolean = CHK_CLASS_STD14OAFL_RTUREASON()

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

    Private Function CHK_CLASS_STD14OAFL_RTUREASON() As Boolean
        'Dim rst As Boolean = False
        Dim vTBCID As String = TIMS.ClearSQM(Hid_TBCID.Value)
        Dim vORGKINDGW As String = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Dim dtFL As DataTable = GET_CLASS_STD14OAFL(objconn, vTBCID, vORGKINDGW)
        For Each drFL As DataRow In dtFL.Rows
            If $"{drFL("RTUREASON")}" <> "" Then Return True
        Next
        Return False
    End Function

    Private Function GET_CLASS_STD14OAFL(oConn As SqlConnection, vTBCID As String, vORGKINDGW As String) As DataTable
        Dim rParms As New Hashtable
        rParms.Add("TBCID", TIMS.CINT1(vTBCID))
        rParms.Add("ORGKINDGW", vORGKINDGW)
        Dim rsSql As String = "
SELECT a.TBCFID,cc.YEARS,ob.APPSTAGE,a.TBCID,a.KTSEQ,a.PATTERN,a.MEMO1
,dbo.FN_CYEAR2(cc.YEARS) YEARS_ROC
,CASE ob.APPSTAGE WHEN 1 THEN '上半年' WHEN 2 THEN '下半年' WHEN 3 THEN '政策性產業' WHEN 4 THEN '進階政策性產業' END APPSTAGE_N
,CASE ob.APPSTAGE WHEN 1 THEN '上' WHEN 2 THEN '下' WHEN 3 THEN '政' WHEN 4 THEN '進' END APPSTAGE_S
/* a.MODIFYACCT,a.MODIFYDATE */
,kb.KTID,kb.KTNAME,concat(kb.KTID,'.',kb.KTNAME) KTNAME1,concat(kb.ORGKINDGW,kb.KTID,kb.KTNAME) KTNAME2
,cc.ORGNAME
,a.WAIVED,a.SRCFILENAME1,a.FILENAME1,a.FILEPATH1,a.FILENAME1 OKFLAG
,a.RTUREASON,a.RTURESACCT,a.RTURESDATE
FROM CLASS_STD14OAFL a
JOIN KEY_STD14TH kb on kb.KTSEQ=a.KTSEQ
JOIN CLASS_STD14OA ob on ob.TBCID =a.TBCID 
JOIN VIEW2 CC ON CC.OCID=ob.OCID
WHERE a.TBCID =@TBCID AND kb.ORGKINDGW=@ORGKINDGW
ORDER BY kb.KSORT,a.TBCFID
"
        Dim dt As DataTable = DbAccess.GetDataTable(rsSql, oConn, rParms)
        Return dt
    End Function

    Private Sub SAVEDATA1()
        '申辦狀態：暫存/ 已送件 '審查狀態：申辦確認/ 申辦退件修正 / 申辦不通過
        'Hid_RID.Value = TIMS.ClearSQM(Hid_RID.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!!!")
            Return
        End If

        Dim vTBCID As String = Hid_TBCID.Value
        Dim vTBCASENO As String = Hid_TBCASENO.Value

        Dim v_ddlAPPLIEDRESULT As String = TIMS.GetListValue(ddlAPPLIEDRESULT)
        Reasonforfail.Text = TIMS.ClearSQM(Reasonforfail.Text)
        Dim v_Reasonforfail As String = TIMS.Get_Substr1(Reasonforfail.Text, cst_MaxLen500_i)
        'Reasonforfail.Text = TIMS.Get_Substr1(Reasonforfail.Text, cst_MaxLen500_i)

        Dim s_HISREVIEW As String = $"{drOA("HISREVIEW")}"
        s_HISREVIEW &= $"{If(s_HISREVIEW <> "", "，", "")}{TIMS.Cdate3t(Now)}-{TIMS.GetListText(ddlAPPLIEDRESULT)}"

        Dim uParms As New Hashtable
        uParms.Add("HISREVIEW", s_HISREVIEW)
        uParms.Add("APPLIEDRESULT", v_ddlAPPLIEDRESULT)
        uParms.Add("REASONFORFAIL", If(v_Reasonforfail <> "", v_Reasonforfail, Convert.DBNull))
        uParms.Add("RESULTACCT", sm.UserInfo.UserID)
        'uParms.Add("RESULTDATE", RESULTDATE)
        uParms.Add("MODIFYACCT", sm.UserInfo.UserID)
        'uParms.Add("MODIFYDATE", MODIFYDATE)
        uParms.Add("TBCID", vTBCID)
        uParms.Add("TBCASENO", vTBCASENO)
        Dim usSql As String = ""
        usSql &= " UPDATE CLASS_STD14OA" & vbCrLf
        usSql &= " SET HISREVIEW=@HISREVIEW,APPLIEDRESULT=@APPLIEDRESULT" & vbCrLf
        usSql &= " ,REASONFORFAIL=@REASONFORFAIL,RESULTACCT=@RESULTACCT,RESULTDATE=GETDATE()" & vbCrLf
        usSql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        usSql &= " WHERE TBCID=@TBCID AND TBCASENO=@TBCASENO" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

        If v_ddlAPPLIEDRESULT = "R" Then
            '退件修正 BISTATUS='R'
            Dim uParms2 As New Hashtable
            uParms2.Add("MODIFYACCT", sm.UserInfo.UserID)
            uParms2.Add("TBCID", vTBCID)
            uParms2.Add("TBCASENO", vTBCASENO)
            Dim usSql2 As String = ""
            usSql2 &= " UPDATE CLASS_STD14OA" & vbCrLf
            usSql2 &= " SET TBCSTATUS='R',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql2 &= " WHERE TBCID=@TBCID AND TBCASENO=@TBCASENO" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql2, objconn, uParms2)
        End If

        Call SEARCH_1()

    End Sub

    Protected Sub BTN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_PACKAGE_DOWNLOAD1.Click
        If (sm.UserInfo.LID > 1) Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg16)
            Return
        End If

        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)

        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無線上申辦案件)!")
            Return
        End If

        Dim rPMS As New Hashtable From {{"TBCID", Hid_TBCID.Value}, {"TBCASENO", Hid_TBCASENO.Value}, {"ORGKINDGW", Hid_ORGKINDGW.Value}}
        ' "DOWNLOAD4" '下載
        Call ResponseZIPFileALL_OA(Me, objconn, rPMS)
    End Sub

    Private Sub ResponseZIPFileALL_OA(MyPage As Page, oConn As SqlConnection, rPMS As Hashtable)
        Const cst_UtlSubName As String = "/*ResponseZIPFileALL_OA*/"
        Dim vTBCID As String = TIMS.GetMyValue2(rPMS, "TBCID")
        Dim vTBCASENO As String = TIMS.GetMyValue2(rPMS, "TBCASENO")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")

        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, vTBCID, vTBCASENO)
        If drOA Is Nothing Then
            Common.MessageBox(Me, "資訊有誤(查無案件編號)，請重新操作!")
            Return
        End If
        Dim drCC As DataRow = TIMS.GetOCIDDate($"{drOA("OCID")}", objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "儲存資訊有誤(查無職類/班別代碼)，請重新操作!")
            Return
        End If

        Dim Template_ZipPath1 As String = TIMS.GET_Template_ZipPath1(vTBCID)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath1)) Then
            Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath1))
        End If

        Dim vPLANYEARS As String = $"{drCC("PLANYEARS")}"
        Dim vDISTID As String = $"{drCC("DISTID")}"
        Dim vDISTNAME3 As String = TIMS.GET_DISTNAME3(oConn, vDISTID)
        Dim vPLANID As String = $"{drCC("PLANID")}"
        Dim vORGNAME As String = $"{drCC("ORGNAME")}"
        Dim vAPPSTAGE_S As String = TIMS.Get_APPSTAGE_S($"{drOA("APPSTAGE")}")
        Dim vYEARS_ROC As String = $"{drCC("YEARS_ROC")}"

        Dim dtFL As DataTable = GET_CLASS_STD14OAFL_TB(objconn, vTBCID, vORGKINDGW)

        Dim oFILENAME1 As String = ""
        Dim oFILEPATH1 As String = ""
        Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
        Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
        '年度申請階段_單位名稱_項目編號+項目名稱
        Dim t_FILENAME As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf")
        Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))

        For Each drFL As DataRow In dtFL.Rows
            Dim vKTID As String = $"{drFL("KTID")}"
            Dim vKTSEQ As String = $"{drFL("KTSEQ")}"
            Dim vKTNAME2X As String = $"{drFL("KTNAME2")}".Replace(".", "x")
            Dim vTBCFID As String = $"{drFL("TBCFID")}"

            'Dim t_FilePath1 As String = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
            Try
                oFILENAME1 = $"{drFL("FILENAME1")}"
                oFILEPATH1 = $"{drFL("FILEPATH1")}"
                oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH1_OA(vPLANYEARS, vDISTID, vPLANID, vTBCASENO, ""))
                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                '年度申請階段_單位名稱_項目編號+項目名稱
                t_FILENAME = TIMS.GetValidFileName($"{vYEARS_ROC}{vAPPSTAGE_S}_{vORGNAME}_{vKTNAME2X}.pdf")
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

            Select Case $"{vORGKINDGW}{vKTID}"
                Case G06_其他補充資料, W06_其他補充資料

                    Dim dtFLOTH As DataTable = TIMS.GET_CLASS_STD14OAFL_OTH(oConn, vTBCFID)
                    If TIMS.dtHaveDATA(dtFLOTH) Then
                        For Each drFLOTH As DataRow In dtFLOTH.Rows
                            Try
                                oFILENAME1 = $"{drFLOTH("FILENAME1")}"
                                oUploadPath = $"{drFLOTH("FILEPATH1")}"
                                s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                                t_FILENAME = TIMS.GetValidFileName($"{vKTNAME2X}x{drFLOTH("TBCID")}x{drFLOTH("ROWNUM1")}.pdf")
                                t_FilePath1 = MyPage.Server.MapPath(Path.Combine($"{Template_ZipPath1}/", t_FILENAME))
                                If IO.File.Exists(s_FilePath1) Then
                                    Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                                    File.WriteAllBytes(t_FilePath1, dbyte)
                                End If
                            Catch ex As Exception
                                Dim strErrmsg As String = "/*Public Shared Sub ResponseZIPFile*/" & vbCrLf
                                strErrmsg &= String.Concat("oFILENAME1: ", oFILENAME1, vbCrLf, "oUploadPath: ", oUploadPath, vbCrLf)
                                strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("t_FILENAME: ", t_FILENAME, vbCrLf)
                                strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                                strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                                Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                            End Try
                        Next
                    End If

            End Select

        Next

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim filenames As String() = Directory.GetFiles(MyPage.Server.MapPath(String.Concat(Template_ZipPath1, "/")))
        Dim zipFileName As String = TIMS.GetValidFileName(String.Concat("p", vYEARS_ROC, vAPPSTAGE_S, "_", vDISTNAME3, "_", vORGNAME, "_", vTBCID, "_", strNOW, ".zip"))
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
                    Dim strErrmsg As String = cst_UtlSubName & vbCrLf
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
            TIMS.SAVE_ADP_ZIPFILE(oConn, "-sd19002all", File)
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

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Protected Sub DataGrid06_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid06.SelectedIndexChanged

    End Sub

    Private Sub DataGrid06_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid06.ItemDataBound
        'ROWNUM1,序號>,SRCFILENAME1,上傳檔案>,MEMO1,備註說明>,Hid_CS14OFID,BTN_DOWNLOAD06,檔案下載,DOWNLOAD06,BTN_DELFILE06,刪除檔案,DELFILE06,
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item 'ListItemType.EditItem, 
                'e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                Dim drv As DataRowView = e.Item.DataItem
                'Dim BTN_DELFILE06 As Button = e.Item.FindControl("BTN_DELFILE06") '(刪除)
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
                'BTN_DELFILE06.CommandArgument = sCmdArg
                'Dim vMsgB As String = "請注意：此筆線上申辦案件原已上傳之相關文件均會一併刪除，確定要刪除此筆資料嗎?"
                'BTN_DELFILE06.Attributes("onclick") = $"javascript:return confirm('{vMsgB}');"
                ''檢視不能修改
                'BTN_DELFILE06.Visible = If(Session(cst_ss_RqProcessType) = cst_DG1CMDNM_VIEW1, False, True)
                ''(其他原因調整) '送件／退件修正，不提供刪除
                'If $"{drv("TBCSTATUS")}" = "B" Then
                '    BTN_DELFILE06.Enabled = False
                '    TIMS.Tooltip(BTN_DELFILE06, cst_tpmsg_enb6, True)
                'ElseIf $"{drv("TBCSTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" <> "" Then
                '    BTN_DELFILE06.Enabled = False '"(退件修正)有退件原因,可重新上傳"
                '    TIMS.Tooltip(BTN_DELFILE06, cst_tpmsg_enb8, True)
                'ElseIf $"{drv("TBCSTATUS")}" = "R" AndAlso $"{drv("RTUREASON")}" = "" Then
                '    BTN_DELFILE06.Enabled = False
                '    TIMS.Tooltip(BTN_DELFILE06, cst_tpmsg_enb7, True)
                'End If
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
        Call SHOW_DATAGRID_06(vKTSEQ)

        '顯示上傳檔案/ 細項
        'Dim rPMS3 As New Hashtable From {{"ORGKINDGW", Hid_ORGKINDGW.Value}, {"TBCID", Hid_TBCID.Value}}
        'Call SHOW_STD14OAFL_DG2(rPMS3)
    End Sub
    Private Sub SHOW_DATAGRID_06(vKTSEQ As String)
        td_title06.Visible = True
        tr_DataGrid06.Visible = True 'If($"{drKB("DataGrid06")}" = "Y", True, False)
        '方式 1：使用 Attributes 集合（最通用於 HtmlControl）
        'td_USEMEMO1.Attributes("class") = "bluecol_need"
        iDG06_ROWS = 0
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_TBCASENO.Value = TIMS.ClearSQM(Hid_TBCASENO.Value)
        Hid_TBCID.Value = TIMS.ClearSQM(Hid_TBCID.Value)
        'Hid_KTSEQ.Value = TIMS.ClearSQM(Hid_KTSEQ.Value)
        Dim vTBCID As String = Hid_TBCID.Value
        'Dim vKTSEQ As String = Hid_KTSEQ.Value 'TIMS.GetListValue(ddlSwitchTo)
        Dim drOA As DataRow = TIMS.GET_CLASS_STD14OA(objconn, Hid_TBCID.Value, Hid_TBCASENO.Value)
        Dim drKB As DataRow = TIMS.GET_KEY_STD14TH(sm, objconn, vKTSEQ, Hid_ORGKINDGW.Value)
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

End Class
