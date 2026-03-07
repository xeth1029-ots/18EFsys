Public Class CR_01_005
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'PLAN_STAFFOPIN / PSOID / PSNO28
    Dim s_COL_PSNO28 As String = "" '課程申請流水號
    Dim g_IMP_ERR1 As Boolean = False

    Const cst_col_PSNO28 As Integer = 0 '課程申請流水號
    Const cst_col_CURESULT As Integer = 1 '核班結果,核班結果
    Const cst_col_NGREASON As Integer = 2 '核班結果,未核班原因 
    Const cst_col_iMaxLen As Integer = 3

    '分署確認課程分類 / 職類課程/ 訓練業別
    Dim dtGCODE3 As DataTable = Nothing

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1
        Call TIMS.OpenDbConn(objconn)
        '審查職類代碼
        dtGCODE3 = TIMS.Get_GOVCODE3dt(objconn)

        If Not IsPostBack Then
            cCreate1()
        End If

        '委訓
        Select Case sm.UserInfo.LID
            Case 2
                Button2.Visible = False
            Case Else
                'Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
                'If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
                '    '署(局) 或 分署(中心)
                '    TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                '    If HistoryRID.Rows.Count <> 0 Then
                '        center.Attributes("onclick") = "showObj('HistoryList2');"
                '        center.Style("CURSOR") = "hand"
                '    End If
                '    Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
                'End If
                TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
                If HistoryRID.Rows.Count <> 0 Then
                    center.Attributes("onclick") = "showObj('HistoryList2');"
                    center.Style("CURSOR") = "hand"
                End If
                Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
                Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End Select
    End Sub

    Sub cCreate1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False

        Const cst_Title_msg1 As String = "當有勾選，於資料匯入時，系統不檢核允許直接覆蓋匯入資料。"
        TIMS.Tooltip(ChkBxCover1, cst_Title_msg1, True)

        msg1.Text = ""
        tbDataGrid1.Visible = False

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim v_APPSTAGE_SCH_DEF As String = "1"
        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, v_APPSTAGE_SCH_DEF)

        '訓練機構
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ 

        '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位 '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位

        '初審建議結論 --Y 通過、N 不通過、P 調整後通過
        ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        ddlGCODE = TIMS.Get_GOVCODE3(dtGCODE3, ddlGCODE, False)
        '一階審查結果 RESULT 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
        ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)
        '核班結果,核班結果'Y 通過、N 不通過
        ddlCURESULT = TIMS.Get_CURESULT(ddlCURESULT)
    End Sub

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Call SSearch1()
    End Sub

    Function GET_ORG_SQL1() As String
        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @APPSTAGE NUMERIC(10,0)=2;
        Dim sql As String = ""
        sql &= " SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
        sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
        sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
        Return sql
    End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt() As DataTable
        Dim dt As DataTable = Nothing
        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過
        Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        '審查結果'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        Dim v_RBL_RESULT_SCH As String = TIMS.GetListValue(RBL_RESULT_SCH)
        'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim v_RBL_CURESULT_SCH As String = TIMS.GetListValue(RBL_CURESULT_SCH)

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        If s_DISTID = "" AndAlso v_RBL_CrossDist_SCH <> "C" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If
        If v_RBL_CrossDist_SCH = "C" Then s_DISTID = ""
        If v_RBL_CrossDist_SCH = "C" Then RIDValue.Value = ""

        '篩選範圍 1:不區分 2:轄區單位 3:19大類主責課程 SYS_GCODEREVIE
        Dim v_RBL_RANGE1_SCH As String = TIMS.GetListValue(RBL_RANGE1_SCH)

        '計畫'TRPlanPoint28
        'Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        '開訓日期
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        '課程申請流水號
        schPSNO28.Text = TIMS.ClearSQM(schPSNO28.Text)

        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

        Dim sql_WORG1 As String = String.Format("WITH WORG1 AS ({0})", GET_ORG_SQL1())

        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @TPLANID VARCHAR(3)='28';DECLARE @APPSTAGE NUMERIC(10,0)=2;
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Dim parms As New Hashtable From {{"YEARS", sm.UserInfo.Years}, {"TPLANID", sm.UserInfo.TPlanID}, {"APPSTAGE", v_APPSTAGE_SCH}}
        Dim sql As String = ""
        sql &= sql_WORG1
        sql &= " SELECT pp.YEARS" & vbCrLf
        sql &= " ,dbo.FN_CYEAR2(pp.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,pp.APPSTAGE" & vbCrLf
        sql &= " ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" & vbCrLf
        sql &= " ,pp.PLANNAME" & vbCrLf
        sql &= " ,pp.PSNO28 ,pp.RID" & vbCrLf
        sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        sql &= " ,pp.OCID" & vbCrLf
        sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf
        sql &= " ,pp.CLASSCNAME" & vbCrLf
        sql &= " ,FORMAT(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,FORMAT(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        sql &= " ,pp.GCID3" & vbCrLf
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '訓練業別編碼" & vbCrLf
        sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sql &= " ,ig3.CNAME GCNAME" & vbCrLf '/訓練業別名稱" & vbCrLf

        sql &= " ,pf.PSOID" & vbCrLf '審查幕僚意見 --SEQNO
        sql &= " ,pf.ST1SUGGEST" & vbCrLf '初審幕僚建議/分署幕僚建議
        sql &= " ,pf.OTHFIXCONT" & vbCrLf '其他應修正內容
        sql &= " ,pf.CONFIRMCONT" & vbCrLf '送請委員確認內容
        '初審建議結論	(轄區分署)顯示通過、不通過、調整後通過
        '(19大類主責分署)下拉選單，選項包括：==請選擇==、通過、不通過、調整後通過
        '初審建議結論 Y 通過、N 不通過、P 調整後通過 '1:通過/2:調整後通過/3:不通過
        sql &= " ,pf.ST1RESULT" & vbCrLf '初審建議結論
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END ST1RESULT_N" & vbCrLf
        ''審查結果 1:通過/2:調整後通過/3:不通過
        sql &= " ,pf.RESULT" & vbCrLf '審查結果
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END RESULT_C" & vbCrLf
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END RESULT_N" & vbCrLf
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf
        '分署確認課程分類
        sql &= " ,gc.PFCNAME" & vbCrLf
        '19大類主責課程 SYS_GCODEREVIE
        sql &= " ,gr1.DISTID GRDISTID " & vbCrLf

        sql &= " ,pf.CURESULT" & vbCrLf ' 核班結果,核班結果'Y 通過、N 不通過
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3'  END CURESULT_C" & vbCrLf
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過'  END CURESULT_N" & vbCrLf
        sql &= " ,pf.NGREASON" & vbCrLf '核班結果,未核班原因

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=pp.TMID" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        sql &= " JOIN WORG1 wo on wo.ORGID=pp.ORGID" & vbCrLf
        '19大類主責課程 SYS_GCODEREVIE
        'sql &= " JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        '申請階段2 (1:上半年/2:下半年/3:政策性產業/4:進階政策性產業) (請選擇)
        Select Case v_APPSTAGE_SCH
            Case "1", "2"
                sql &= " JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
            Case Else
                sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        End Select
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf

        sql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sql &= " AND pp.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND pp.YEARS=@YEARS" & vbCrLf
        sql &= " AND pp.APPSTAGE=@APPSTAGE" & vbCrLf

        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
            sql &= " AND pp.RID =@RID"
            parms.Add("RID", RIDValue.Value)
        End If

        '計畫'TRPlanPoint28
        If TRPlanPoint28.Visible Then
            Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
            Select Case v_rblOrgKind2'rblOrgKind2.SelectedValue
                Case "G", "W"
                    sql &= " AND pp.ORGKIND2=@ORGKIND2"
                    parms.Add("ORGKIND2", v_rblOrgKind2)
            End Select
        End If

        'STDate1.Text = TIMS.cdate3(STDate1.Text)
        'STDate2.Text = TIMS.cdate3(STDate2.Text)
        If STDate1.Text <> "" Then
            sql &= " AND pp.STDATE >=@STDATE1"
            parms.Add("STDATE1", TIMS.Cdate2(STDate1.Text))
        End If
        If STDate2.Text <> "" Then
            sql &= " AND pp.STDATE <=@STDATE2"
            parms.Add("STDATE2", TIMS.Cdate2(STDate2.Text))
        End If
        '課程申請流水號
        If schPSNO28.Text <> "" Then
            sql &= " AND pp.PSNO28=@PSNO28"
            parms.Add("PSNO28", schPSNO28.Text)
        End If

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 / ST1RESULT
        'Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        Select Case v_RBL_ST1RESULT_SCH
            Case "2"
                sql &= " AND pf.ST1RESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.ST1RESULT IS NULL" & vbCrLf
            Case "Y", "N", "P"
                sql &= " AND pf.ST1RESULT=@ST1RESULT" & vbCrLf
                parms.Add("ST1RESULT", v_RBL_ST1RESULT_SCH)
        End Select
        '審查結果'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        Select Case v_RBL_RESULT_SCH
            Case "2"
                sql &= " AND pf.RESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.RESULT IS NULL" & vbCrLf
            Case "Y", "N", "P"
                sql &= " AND pf.RESULT=@RESULT" & vbCrLf
                parms.Add("RESULT", v_RBL_RESULT_SCH)
        End Select
        'CURESULT 核班結果,核班結果 不區分 'Y 通過、N 不通過
        Select Case v_RBL_CURESULT_SCH
            Case "2"
                sql &= " AND pf.CURESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.CURESULT IS NULL" & vbCrLf
            Case "Y", "N"
                sql &= " AND pf.CURESULT=@CURESULT" & vbCrLf
                parms.Add("CURESULT", v_RBL_CURESULT_SCH)
        End Select

        '篩選範圍 1:不區分 2:轄區單位 3:19大類主責課程 SYS_GCODEREVIE
        If s_DISTID <> "" Then
            Select Case Val(v_RBL_RANGE1_SCH)
                Case 1
                    sql &= " AND (pp.DISTID=@DISTID OR gr1.DISTID=@DISTID)"
                    parms.Add("DISTID", s_DISTID)
                Case 2
                    sql &= " AND pp.DISTID=@DISTID"
                    parms.Add("DISTID", s_DISTID)
                Case 3
                    sql &= " AND gr1.DISTID=@DISTID"
                    parms.Add("DISTID", s_DISTID)
            End Select
        End If

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                sql &= " AND wo.I_CROSSDIST !=-1" & vbCrLf
            Case "J" 'J:轄區提案單位
                sql &= " AND wo.I_CROSSDIST =-1 " & vbCrLf
        End Select

        'ROW_NUMBER() OVER(ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE) SEQNUM
        sql &= " ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE"

        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, $"--{vbCrLf}{TIMS.GetMyValue5(parms)}{vbCrLf}--##CR_01_005:{vbCrLf}{sql}")

        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    Sub SSearch1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        msg1.Text = TIMS.cst_NODATAMsg1
        tbDataGrid1.Visible = False

        Dim dt As DataTable = SEARCH_DATA1_dt()

        If TIMS.dtNODATA(dt) Then
            msg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If

        msg1.Text = ""
        tbDataGrid1.Visible = True
        'DataGrid1.DataSource = dt 'DataGrid1.DataBind()
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing Then Return
        Dim sCmdArg As String = e.CommandArgument
        Dim sCMDNM As String = e.CommandName
        If sCmdArg = "" OrElse sCMDNM = "" Then Return

        Call CLEAR_DATA1()

        Hid_PSOID.Value = TIMS.GetMyValue(sCmdArg, "PSOID")
        Hid_PSNO28.Value = TIMS.GetMyValue(sCmdArg, "PSNO28")
        Hid_GCODE.Value = TIMS.GetMyValue(sCmdArg, "GCODE")
        Hid_PFGCODE.Value = TIMS.GetMyValue(sCmdArg, "PFGCODE")
        Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))
        If Hid_PSNO28.Value = "" Then Return

        Dim dr1 As DataRow = GET_DATA1()
        Select Case sCMDNM'e.CommandName
            Case "EDT1"
                btnSAVE1.Visible = True
                Call SHOW_DATA1(dr1)
                Call DISABLE_SHOW1(dr1) 'EDT1
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg9)
        End Select
    End Sub

    Private Sub DISABLE_SHOW1(ByRef dr1 As DataRow)
        If dr1 Is Nothing Then Return

        Dim flag_OTHFIXCONT_OK As Boolean = If(sm.UserInfo.DistID = Convert.ToString(dr1("DISTID")), True, False)
        'Dim flag_CONFIRMCONT_OK As Boolean = If(sm.UserInfo.DistID = Convert.ToString(dr1("GRDISTID")), True, False)

        ddlCURESULT.Enabled = flag_OTHFIXCONT_OK  'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        NGREASON.Enabled = flag_OTHFIXCONT_OK '核班結果,未核班原因

        If (Not ddlCURESULT.Enabled) Then TIMS.Tooltip(ddlCURESULT, "所屬轄區之訓練單位，可填寫「核班結果」", True)
        If (Not NGREASON.Enabled) Then TIMS.Tooltip(NGREASON, "所屬轄區之訓練單位，可填寫「未核班原因」", True)
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)
                '初審建議結論 Y 通過、N 不通過、P 調整後通過
                Dim labST1RESULT_N As Label = e.Item.FindControl("labST1RESULT_N")
                labST1RESULT_N.Text = Convert.ToString(drv("ST1RESULT_N")) 'TIMS.Get_ST1RESULT_N(Convert.ToString(drv("ST1RESULT_N")))

                'dg_RESULT 審查結果 RESULT 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
                'Dim dg_RESULT As DropDownList = e.Item.FindControl("dg_RESULT")
                'dg_RESULT = TIMS.Get_ST1RESULT(dg_RESULT)
                'Common.SetListItem(dg_RESULT, Convert.ToString(drv("RESULT")))
                '審查結果 RESULT 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
                Dim labRESULT_N As Label = e.Item.FindControl("labRESULT_N")
                labRESULT_N.Text = Convert.ToString(drv("RESULT_N")) 'TIMS.Get_ST1RESULT_N(Convert.ToString(drv("ST1RESULT_N")))
                'CURESULT 核班結果,核班結果'Y 通過、N 不通過
                Dim labCURESULT_N As Label = e.Item.FindControl("labCURESULT_N")
                labCURESULT_N.Text = Convert.ToString(drv("CURESULT_N"))

                Dim BtnEDT1 As Button = e.Item.FindControl("BtnEDT1")    '編輯
                BtnEDT1.Visible = If(Convert.ToString(drv("PSOID")) <> "", True, False)

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "PSOID", Convert.ToString(drv("PSOID")))
                TIMS.SetMyValue(sCmdArg, "PSNO28", Convert.ToString(drv("PSNO28")))
                TIMS.SetMyValue(sCmdArg, "GCODE", Convert.ToString(drv("GCODE")))
                TIMS.SetMyValue(sCmdArg, "PFGCODE", Convert.ToString(drv("PFGCODE")))
                BtnEDT1.CommandArgument = sCmdArg

        End Select
    End Sub

    Sub CLEAR_DATA1()
        lbYEARS_ROC.Text = ""
        lbDistName.Text = ""
        lbOrgName.Text = ""
        lbPSNO28.Text = ""
        lbClassName.Text = ""
        lbSFTDate.Text = ""
        lbGCNAME.Text = ""
        lbCCNAME.Text = ""
        lbTNum.Text = ""
        lbTHours.Text = ""
        lbACTHUMCOST.Text = ""
        lbMETSUMCOST.Text = ""
        lbIsCROSSDIST.Text = ""
        lbiCAPNUM.Text = ""
        lbD20KNAME.Text = "" '$"{dr("D20KNAME")}" '政府政策性產業
        lbD25KNAME.Text = "" '$"{dr("D25KNAME")}" '政府政策性產業
        lbD26KNAME.Text = "" '$"{dr("D26KNAME")}" '政府政策性產業

        ST1SUGGEST.Text = ""
        OTHFIXCONT.Text = ""
        ddlST1RESULT.SelectedIndex = -1
        Common.SetListItem(ddlST1RESULT, "")
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        ddlGCODE.SelectedIndex = -1
        Common.SetListItem(ddlGCODE, "")

        CONFIRMCONT.Text = ""
        COMMENTS.Text = "" 'Convert.ToString(dr("COMMENTS")) '委員審查意見與建議
        ddlRESULT.SelectedIndex = -1
        Common.SetListItem(ddlRESULT, "")

        'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        ddlCURESULT.SelectedIndex = -1
        '核班結果,未核班原因
        Common.SetListItem(ddlCURESULT, "")
        NGREASON.Text = ""
    End Sub

    Function GET_DATA1() As DataRow
        Dim dr1 As DataRow = Nothing
        If Hid_PSNO28.Value = "" Then Return dr1
        Dim parms As New Hashtable
        parms.Add("PSNO28", Hid_PSNO28.Value)

        Dim sql As String = ""
        sql &= " SELECT rr.ORGPLANNAME " & vbCrLf '-- 計畫別、" & vbCrLf
        sql &= " ,pp.PSNO28" & vbCrLf
        sql &= " ,pp.YEARS" & vbCrLf
        sql &= " ,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        'sql &= " ,pp.DISTID,pp.DISTNAME" & vbCrLf '分署別" & vbCrLf
        'sql &= " ,pp.ORGNAME" & vbCrLf '訓練單位名稱" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf 'FIRSTSORT
        sql &= " ,pp.PSNO28" & vbCrLf '課程申請流水號" & vbCrLf
        sql &= " ,pp.CLASSCNAME" & vbCrLf '班級名稱" & vbCrLf
        sql &= " ,format(pp.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期" & vbCrLf
        sql &= " ,format(pp.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期" & vbCrLf
        sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '/: 訓練業別編碼" & vbCrLf
        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,ig3.GCODE31 GCODE" & vbCrLf
        sql &= " ,ig3.PNAME GCODEPNAME" & vbCrLf
        sql &= " ,ig3.CNAME GCNAME" & vbCrLf '/訓練業別名稱" & vbCrLf
        'sql &= " ,ig3.GCODE2 GCODENAME" & vbCrLf '訓練業別編碼" & vbCrLf
        'sql &= " ,ig3.CNAME GCNAME" & vbCrLf '訓練業別名稱" & vbCrLf
        sql &= " ,kc.CCNAME " & vbCrLf '訓練職能" & vbCrLf

        sql &= " ,pp.TNUM " & vbCrLf '訓練人次" & vbCrLf
        sql &= " ,pp.THOURS " & vbCrLf '訓練時數" & vbCrLf
        sql &= " ,pp.ACTHUMCOST" & vbCrLf '實際人時成本" & vbCrLf
        sql &= " ,pp.METSUMCOST" & vbCrLf '實際材料費" & vbCrLf
        sql &= " ,dd.D20KNAME,dd.D25KNAME,dd.D26KNAME" & vbCrLf '政府政策性產業
        '5+2產業,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策
        sql &= " ,dd.D20KNAME1,dd.D20KNAME2,dd.D20KNAME3,dd.D20KNAME4,dd.D20KNAME5,dd.D20KNAME6" & vbCrLf '"5+2產業創新計畫"" & vbCrLf
        ' 5+2產業--新南向政策--台灣AI行動計畫--數位國家創新經濟發展方案--國家資通安全發展方案--前瞻基礎建設計畫" & vbCrLf
        sql &= " ,dbo.FN_GET_CROSSDIST(pp.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST " & vbCrLf '是否跨區提案" & vbCrLf
        sql &= " ,pp.ICAPNUM  " & vbCrLf 'iCAP標章證號" & vbCrLf
        sql &= " ,pf.PSOID " & vbCrLf
        sql &= " ,pf.ST1SUGGEST " & vbCrLf '初審幕僚建議/分署幕僚意見
        sql &= " ,pf.OTHFIXCONT " & vbCrLf '其他應修正內容" & vbCrLf
        sql &= " ,pf.CONFIRMCONT" & vbCrLf '送請委員確認內容" & vbCrLf
        sql &= " ,pf.ST1RESULT" & vbCrLf '初審建議結論" & vbCrLf'初審建議結論 Y 通過、N 不通過、P 調整後通過
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END ST1RESULT_C" & vbCrLf
        '初審建議結論 Y 通過、N 不通過、P 調整後通過
        sql &= " ,CASE pf.ST1RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END ST1RESULT_N" & vbCrLf
        sql &= " ,pf.RESULT" & vbCrLf '審查結果
        '1:通過/2:調整後通過/3:不通過
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3' WHEN 'P' THEN '2' END RESULT_C" & vbCrLf
        sql &= " ,CASE pf.RESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' WHEN 'P' THEN '調整後通過' END RESULT_N" & vbCrLf
        sql &= " ,pf.COMMENTS" & vbCrLf '委員審查意見與建議

        '分署確認課程分類 / 職類課程 / 訓練業別
        sql &= " ,pf.GCODE PFGCODE" & vbCrLf
        '分署確認課程分類
        sql &= " ,gc.PFCNAME" & vbCrLf
        '19大類主責課程 SYS_GCODEREVIE
        sql &= " ,gr1.DISTID GRDISTID " & vbCrLf

        sql &= " ,pf.CURESULT" & vbCrLf ' 核班結果,核班結果'Y 通過、N 不通過
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3'  END CURESULT_C" & vbCrLf
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過'  END CURESULT_N" & vbCrLf
        sql &= " ,pf.NGREASON" & vbCrLf '核班結果,未核班原因

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " LEFT JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        '審查計分等級'19大類主責課程 SYS_GCODEREVIE
        'sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf
        If Hid_PSOID.Value <> "" Then
            sql &= " AND pf.PSOID=@PSOID" & vbCrLf
            parms.Add("PSOID", Hid_PSOID.Value)
        End If
        'sql &= " AND CC.YEARS='2022'" & vbCrLf
        sql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)

        sql &= " AND pp.TPLANID='28'" & vbCrLf
        sql &= " AND pp.PSNO28=@PSNO28" & vbCrLf
        dr1 = DbAccess.GetOneRow(sql, objconn, parms)
        Return dr1
    End Function

    Sub SHOW_DATA1(ByRef dr As DataRow)
        PanelSch1.Visible = False
        PanelEdit1.Visible = True
        '初審幕僚建議
        ST1SUGGEST.ReadOnly = True '分署幕僚意見
        OTHFIXCONT.ReadOnly = True  '其他應修正內容
        ST1SUGGEST.ApplyStyle(TIMS.GET_RO_STYLE())
        OTHFIXCONT.ApplyStyle(TIMS.GET_RO_STYLE())
        TIMS.Tooltip(ST1SUGGEST, "僅提供顯示", True)
        TIMS.Tooltip(OTHFIXCONT, "僅提供顯示", True)
        ddlST1RESULT.Enabled = False '初審建議結論 通過、不通過、調整後通過
        TIMS.Tooltip(ddlST1RESULT, "僅提供顯示", True)
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        ddlGCODE.Enabled = False
        TIMS.Tooltip(ddlGCODE, "僅提供顯示", True)

        CONFIRMCONT.ReadOnly = True  '送請委員確認內容
        CONFIRMCONT.ApplyStyle(TIMS.GET_RO_STYLE())
        TIMS.Tooltip(CONFIRMCONT, "僅提供顯示", True)

        '審查結果
        COMMENTS.ReadOnly = True '委員審查意見與建議
        COMMENTS.ApplyStyle(TIMS.GET_RO_STYLE())
        TIMS.Tooltip(COMMENTS, "僅提供顯示", True)
        ddlRESULT.Enabled = False '審查結果/一階審查結果
        TIMS.Tooltip(ddlRESULT, "僅提供顯示", True)

        If dr Is Nothing Then Return
        Hid_PSOID.Value = Convert.ToString(dr("PSOID"))
        Hid_PSNO28.Value = Convert.ToString(dr("PSNO28"))

        lbYEARS_ROC.Text = TIMS.GET_YEARS_ROC(dr("YEARS"))
        lbDistName.Text = Convert.ToString(dr("DISTNAME"))
        lbOrgName.Text = Convert.ToString(dr("ORGNAME"))
        lbPSNO28.Text = Convert.ToString(dr("PSNO28"))
        lbClassName.Text = Convert.ToString(dr("CLASSCNAME"))
        lbSFTDate.Text = String.Format("{0}~{1}", dr("STDATE"), dr("STDATE"))

        lbGCNAME.Text = Convert.ToString(dr("GCNAME")) '訓練業別
        lbCCNAME.Text = Convert.ToString(dr("CCNAME")) '訓練職能
        lbTNum.Text = Convert.ToString(dr("TNUM"))
        lbTHours.Text = Convert.ToString(dr("THOURS"))
        lbACTHUMCOST.Text = Convert.ToString(dr("ACTHUMCOST"))
        lbMETSUMCOST.Text = Convert.ToString(dr("METSUMCOST"))
        '是否跨區提案
        Dim s_CROSSDIST As String = If(Convert.ToString(dr("CROSSDIST")) <> "", If(Val(dr("CROSSDIST")) > -1, "是", "否"), "")
        lbIsCROSSDIST.Text = s_CROSSDIST 'Convert.ToString(dr("CROSSDIST")) '是否跨區提案
        lbiCAPNUM.Text = Convert.ToString(dr("ICAPNUM"))
        'lbD20KNAME.Text = TIMS.NullToStr(dr("D20KNAME"), "無") '政府政策性產業
        lbD20KNAME.Text = $"{dr("D20KNAME")}" '政府政策性產業
        lbD25KNAME.Text = $"{dr("D25KNAME")}" '政府政策性產業
        lbD26KNAME.Text = $"{dr("D26KNAME")}" '政府政策性產業
        If $"{dr("D20KNAME")}{dr("D25KNAME")}{dr("D26KNAME")}" = "" Then lbD20KNAME.Text = "無"

        '初審幕僚建議
        ST1SUGGEST.Text = Convert.ToString(dr("ST1SUGGEST")) '分署幕僚意見
        OTHFIXCONT.Text = Convert.ToString(dr("OTHFIXCONT")) '其他應修正內容
        Common.SetListItem(ddlST1RESULT, dr("ST1RESULT")) '初審建議結論 --通過、不通過、調整後通過
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        Hid_GCODE.Value = Convert.ToString(dr("GCODE"))
        Hid_PFGCODE.Value = Convert.ToString(dr("PFGCODE"))
        Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))

        CONFIRMCONT.Text = Convert.ToString(dr("CONFIRMCONT")) '送請委員確認內容
        '審查結果
        COMMENTS.Text = Convert.ToString(dr("COMMENTS")) '委員審查意見與建議
        Common.SetListItem(ddlRESULT, dr("RESULT")) '審查結果

        '核班結果
        Common.SetListItem(ddlCURESULT, dr("CURESULT")) 'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        NGREASON.Text = Convert.ToString(dr("NGREASON")) '核班結果,未核班原因
    End Sub

    Protected Sub btnBACK1_Click(sender As Object, e As EventArgs) Handles btnBACK1.Click
        Call CLEAR_DATA1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False
    End Sub

    Protected Sub btnSAVE1_Click(sender As Object, e As EventArgs) Handles btnSAVE1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Return 'Exit Sub
        End If
        Call SAVE_DATA1()
    End Sub

    Private Sub CheckData1(ByRef s_ERRMSG As String)
        '核班結果,核班結果
        Dim v_ddlCURESULT As String = TIMS.GetListValue(ddlCURESULT)
        If NGREASON.Text <> "" Then NGREASON.Text = Trim(NGREASON.Text)
        If v_ddlCURESULT = "" Then s_ERRMSG &= " 請選擇 核班結果/核班結果" & vbCrLf
        If v_ddlCURESULT = "N" AndAlso NGREASON.Text = "" Then s_ERRMSG += "核班結果 為不通過/未核班原因 必填" & vbCrLf
    End Sub

    Sub SAVE_DATA1()
        Hid_PSOID.Value = TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSOID.Value = "" Then Return
        If Hid_PSNO28.Value = "" Then Return

        '核班結果,核班結果
        Dim v_ddlCURESULT As String = TIMS.GetListValue(ddlCURESULT)

        Dim iRst As Integer = 0
        Dim iPSOID As Integer = Val(Hid_PSOID.Value)
        Dim parms As New Hashtable From {{"PSOID", Val(Hid_PSOID.Value)}, {"PSNO28", Hid_PSNO28.Value}}
        Dim s_sql As String = "SELECT PSOID FROM PLAN_STAFFOPIN WHERE PSOID=@PSOID AND PSNO28=@PSNO28"
        Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then Return '檢核正確性

        Dim u_parms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"CURESULT", v_ddlCURESULT},
            {"NGREASON", If(NGREASON.Text <> "", NGREASON.Text, Convert.DBNull)},
            {"CURESULTACCT", sm.UserInfo.UserID},
            {"PSOID", iPSOID},
            {"PSNO28", Hid_PSNO28.Value}
        }

        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        u_sql &= " SET MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,CURESULT=@CURESULT" & vbCrLf
        u_sql &= " ,NGREASON=@NGREASON" & vbCrLf

        u_sql &= " ,CURESULTACCT=@CURESULTACCT" & vbCrLf
        u_sql &= " ,CURESULTDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Call SSearch1()
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

    Protected Sub BtnIMPORT1_Click(sender As Object, e As EventArgs) Handles BtnIMPORT1.Click
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        ElseIf Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Call ImportXLSX_1()
        Call cCreate1()
    End Sub

    Function CheckImp1(ByRef ErrMsg1 As String) As Boolean
        Dim rst As Boolean = False '正常:true '異常:false 
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        If v_ddlAPPSTAGE_SCH = "" Then
            ErrMsg1 &= "申請階段，未選擇無法匯入，請先選擇申請階段!" & vbCrLf
            Return rst
        End If
        'Common.MessageBox(Me, "分署未選擇，無法匯入，請先選擇分署!")
        'Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        'If v_ddlSCORING = "" Then
        '    ErrMsg1 &= "審查計分區間未選擇，無法匯入，請先選擇審查計分區間!" & vbCrLf
        '    Return rst
        'End If
        rst = True '正常:true '異常:false 
        Return rst
    End Function

    ''' <summary>'匯入等級/分數</summary>
    Sub ImportXLSX_1()
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        ElseIf Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Const cst_Upload_Path As String = "~/CR/01/Temp/" '暫存路徑
        Call TIMS.MyCreateDir(Me, cst_Upload_Path)
        Const Cst_Filetype As String = "xlsx" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, Cst_Filetype) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        '檢查檔案格式與大小 Start
        If File1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        ElseIf File1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱
        MyFileName = Split(File1.PostedFile.FileName, "\")((Split(File1.PostedFile.FileName, "\")).Length - 1)
        '取出檔案類型
        If MyFileName.IndexOf(".") = -1 Then
            Common.MessageBox(Me, "檔案類型錯誤!")
            Exit Sub
        End If
        MyFileType = Split(MyFileName, ".")((Split(MyFileName, ".")).Length - 1)
        If LCase(MyFileType) <> LCase(Cst_Filetype) Then
            Common.MessageBox(Me, "檔案類型錯誤，必須為" & UCase(Cst_Filetype) & "檔!")
            Exit Sub
        End If
        '檢查檔案格式與大小 End

        Dim Errmag As String = ""
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim filePath1 As String = Server.MapPath($"{cst_Upload_Path}{MyFileName}")
        File1.PostedFile.SaveAs(filePath1) '上傳檔案
        '(讀取XLSX檔案轉為dt_xls)
        Dim dt_xls As DataTable = TIMS.ReadXLSX(New IO.FileInfo(filePath1), Errmag)
        '刪除檔案 'If IO.File.Exists(FullFileName1) Then IO.File.Delete(FullFileName1)
        Call TIMS.MyFileDelete(filePath1)

        If Errmag <> "" Then
            Errmag &= "資料有誤，故無法匯入，請修正Excel檔案!"
            Common.MessageBox(Me, Errmag)
            Exit Sub
        End If

        If TIMS.dtNODATA(dt_xls) Then
            If dt_xls Is Nothing Then '有資料
                Common.MessageBox(Me, "資料為空，故無法匯入，請修正Excel檔案!")
                Exit Sub
            End If
            Common.MessageBox(Me, "查無資料，故無法匯入，請修正Excel檔案!")
            Exit Sub
        End If

        '建立錯誤資料格式Table Start
        'Dim Reason As String  '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("PSNO28"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        Dim sHtb As New Hashtable
        Dim iRowIndex As Integer = 0 '讀取行累計數
        Dim Reason As String = "" '做一次驗証的即可
        If Reason = "" Then
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                Reason = ""
                Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)

                Reason = SAVE_PLAN_STAFFOPIN(colArray, sHtb)  '驗証(單筆) 並 儲存

                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = String.Concat("第", CStr(iRowIndex + 2), "列")
                    drWrong("PSNO28") = s_COL_PSNO28 '統一編號
                    drWrong("Reason") = If(Reason <> "", Reason, "(錯誤)") 'Reason
                End If

                iRowIndex += 1 '讀取行累計數
                If g_IMP_ERR1 Then Exit For
            Next
            'Loop
        End If

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""
        '開始判別欄位存入 End
        If TIMS.dtNODATA(dtWrong) Then
            explain = ""
            explain = String.Concat(explain, "匯入資料共", iRowIndex, "筆", vbCrLf)
            explain = String.Concat(explain, "成功：", (iRowIndex - dtWrong.Rows.Count), "筆", vbCrLf)
            explain = String.Concat(explain, "失敗：", dtWrong.Rows.Count, "筆", vbCrLf)
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Reason = String.Concat("錯誤訊息如下:", vbCrLf, Reason)
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            explain2 = String.Concat(explain2, "匯入資料共", iRowIndex, "筆\n")
            explain2 = String.Concat(explain2, "成功：", (iRowIndex - dtWrong.Rows.Count), "筆\n")
            explain2 = String.Concat(explain2, "失敗：", dtWrong.Rows.Count, "筆\n")
            Session("MyWrongTable") = dtWrong
            Const CST_WRONG_ASPX_1 As String = "CR_01_003_Wrong.aspx"
            Dim s_FMT1 As String = String.Format("window.open('{0}','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');", CST_WRONG_ASPX_1)
            Dim s_DYW2CRES As String = String.Concat(explain2, "是否要檢視原因?")
            Dim s_JS1 As String = String.Concat("<script>if(confirm('", s_DYW2CRES, "')){", s_FMT1, "}</script>")
            Page.RegisterStartupScript("", s_JS1)
        End If
    End Sub

    ''' <summary>匯入驗証</summary>
    ''' <param name="colArray">比對資料</param>
    ''' <param name="Htb">輸入查詢</param>
    ''' <param name="o_parms">取得有效值</param>
    ''' <returns></returns>
    Function CheckImportData(ByRef colArray As Array, ByRef Htb As Hashtable, ByRef o_parms As Hashtable) As String
        Dim Reason As String = ""
        s_COL_PSNO28 = ""

        If colArray.Length < cst_col_iMaxLen Then
            g_IMP_ERR1 = True
            Reason += "欄位對應有誤<BR>,請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        'Dim s_col_PSNO28 As String = "" '課程申請流水號
        If colArray.Length > cst_col_PSNO28 Then s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col_PSNO28)) '課程申請流水號

        Dim s_COL_CURESULT As String = "" 'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim s_COL_CURESULT_YN As String = "" '核班結果(YN)
        Dim s_COL_NGREASON As String = "" '核班結果,未核班原因

        s_COL_PSNO28 = TIMS.ClearSQM(colArray(cst_col_PSNO28)) '課程申請流水號

        s_COL_CURESULT = TIMS.NullToStr(colArray(cst_col_CURESULT)) 'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        s_COL_NGREASON = TIMS.ClearSQM(colArray(cst_col_NGREASON)) '核班結果,未核班原因
        'Dim flag_TXT_NG As Boolean = (s_COL_NGREASON = "")

        '先確認資料不為空
        If s_COL_PSNO28 = "" Then Reason += "課程申請流水號 不可為空<br>"

        If s_COL_CURESULT = "" AndAlso s_COL_NGREASON = "" Then Reason += "核班結果/未核班原因 不可皆為空<br>"
        If Reason <> "" Then Return Reason

        Dim s_COL_PCS As String = "" '課程流水號
        s_COL_PCS = TIMS.Get_PCSforPSNO28(sm, s_COL_PSNO28, objconn)
        If s_COL_PCS = "" Then Reason += String.Format("課程申請流水號 有誤，查無班級資料({0})<br>", s_COL_PSNO28)
        If Reason <> "" Then Return Reason

        'Y/N/P'用文字方式輸入
        s_COL_CURESULT_YN = If(s_COL_CURESULT = "通過", "Y", If(s_COL_CURESULT = "不通過", "N", ""))
        If s_COL_CURESULT_YN = "" AndAlso s_COL_CURESULT <> "" Then
            Select Case s_COL_CURESULT '用代碼方式輸入
                Case "Y", "N"
                    s_COL_CURESULT_YN = s_COL_CURESULT
            End Select
        End If
        If s_COL_CURESULT_YN = "" AndAlso s_COL_NGREASON = "" Then Reason += "核班結果/未核班原因 資料皆為空<br>"
        If s_COL_CURESULT_YN = "N" AndAlso s_COL_NGREASON = "" Then Reason += "核班結果 為不通過/未核班原因 必填<br>"

        'Dim flag_EXISTS_1 As Boolean = TIMS.CHK_STAFFOPIN_RESULT_EXISTS(objconn, (s_COL_PSNO28)
        'If flag_EXISTS_1 Then Reason += String.Format("課程申請流水號 已有審核結果，不再匯入({0})<br>", s_COL_PSNO28)
        'If Reason <> "" Then Return Reason

        If Not ChkBxCover1.Checked Then
            Dim flag_EXISTS_2 As Boolean = CHK_STAFFOPIN_CURESULT_EXISTS(s_COL_PSNO28)
            If flag_EXISTS_2 Then Reason += String.Format("課程申請流水號 已有核班結果，不再匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If

        If sm.UserInfo.LID <> 0 Then
            '轄區分署有誤
            Dim flag_DISTID_NG As Boolean = CHK_STAFFOPIN_DISTID_NG(s_COL_PSNO28, sm.UserInfo.DistID)
            If flag_DISTID_NG Then Reason += String.Format("課程申請流水號 轄區分署有誤，不可匯入({0})<br>", s_COL_PSNO28)
            If Reason <> "" Then Return Reason
        End If

        If o_parms Is Nothing Then o_parms = New Hashtable
        o_parms.Add("PSNO28", s_COL_PSNO28)
        o_parms.Add("CURESULT", s_COL_CURESULT_YN)
        o_parms.Add("NGREASON", s_COL_NGREASON)
        Return Reason
    End Function

    ''' <summary>檢核資料是否 已有核班結果 CURESULT IS NOT NULL</summary>
    ''' <param name="s_PSNO28"></param>
    ''' <returns></returns>
    Function CHK_STAFFOPIN_CURESULT_EXISTS(ByVal s_PSNO28 As String) As Boolean
        Dim rst As Boolean = False
        If s_PSNO28 = "" Then Return rst
        Dim dt1 As New DataTable
        Dim s_sql As String = " SELECT PSOID FROM PLAN_STAFFOPIN WITH(NOLOCK) WHERE CURESULT IS NOT NULL AND PSNO28=@PSNO28" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            dt1.Load(.ExecuteReader)
        End With
        rst = (dt1.Rows.Count > 0)
        Return rst
    End Function

    ''' <summary>課程申請流水號 轄區分署有誤，不可匯入</summary>
    ''' <param name="s_PSNO28"></param>
    ''' <param name="s_DISTID"></param>
    ''' <returns></returns>
    Function CHK_STAFFOPIN_DISTID_NG(ByVal s_PSNO28 As String, ByVal s_DISTID As String) As Boolean
        Dim rst As Boolean = False
        If s_PSNO28 = "" Then Return rst
        Dim dt1 As New DataTable
        Dim s_sql As String = ""
        s_sql &= " SELECT pf.PSOID,pp.DISTID FROM PLAN_STAFFOPIN pf WITH(NOLOCK)" & vbCrLf
        s_sql &= " JOIN dbo.VIEW2B pp on pp.PSNO28=pf.PSNO28" & vbCrLf
        s_sql &= " WHERE pf.PSNO28=@PSNO28 AND pp.DISTID!=@DISTID" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            .Parameters.Add("DISTID", SqlDbType.VarChar).Value = s_DISTID ' pp.DISTID!=@DISTID
            dt1.Load(.ExecuteReader)
        End With
        rst = (dt1.Rows.Count > 0)
        Return rst
    End Function

    ''' <summary>取得流水號</summary>
    ''' <param name="s_PSNO28"></param>
    ''' <returns></returns>
    Function GET_PLAN_STAFFOPIN_PSOID(ByVal s_PSNO28 As String) As Integer
        Dim rst As Integer = -1
        If s_PSNO28 = "" Then Return rst
        Dim dt1 As New DataTable
        Dim s_sql As String = " SELECT PSOID FROM PLAN_STAFFOPIN WITH(NOLOCK) WHERE PSNO28=@PSNO28" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("PSNO28", SqlDbType.VarChar).Value = s_PSNO28
            dt1.Load(.ExecuteReader)
        End With
        If (dt1.Rows.Count = 1) Then rst = dt1.Rows(0)("PSOID")
        Return rst
    End Function

    ''' <summary>匯入檔[儲存]</summary>
    ''' <param name="colArray"></param>
    ''' <param name="Htb"></param>
    ''' <returns></returns>
    Function SAVE_PLAN_STAFFOPIN(ByRef colArray As Array, ByRef Htb As Hashtable) As String
        Dim o_parms As New Hashtable
        Dim rst As String = CheckImportData(colArray, Htb, o_parms)
        If rst <> "" Then Return rst

        Dim s_PSNO28 As String = TIMS.GetMyValue2(o_parms, "PSNO28")
        'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim vCURESULT As String = TIMS.GetMyValue2(o_parms, "CURESULT")
        Dim vNGREASON As String = TIMS.GetMyValue2(o_parms, "NGREASON")

        Dim iPSOID As Integer = GET_PLAN_STAFFOPIN_PSOID(s_PSNO28)
        If iPSOID = -1 Then Return rst

        'u_parms.Add("CONFIRMCONT", If(vCONFIRMCONT <> "", vCONFIRMCONT, Convert.DBNull))
        Dim u_parms As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"CURESULT", If(vCURESULT <> "", vCURESULT, Convert.DBNull)},
            {"NGREASON", If(vNGREASON <> "", vNGREASON, Convert.DBNull)},
            {"CURESULTACCT", sm.UserInfo.UserID},
            {"PSOID", iPSOID},
            {"PSNO28", s_PSNO28}
        }
        'i_parms.Add("ST1ACCT", sm.UserInfo.UserID)
        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        u_sql &= " SET MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " ,CURESULT=@CURESULT,NGREASON=@NGREASON " & vbCrLf
        u_sql &= " ,CURESULTACCT=@CURESULTACCT ,CURESULTDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        Return rst
    End Function

End Class
