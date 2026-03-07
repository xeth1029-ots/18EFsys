Imports System.IO
Imports ICSharpCode.SharpZipLib.Zip

Public Class CR_01_008
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

        '2:委訓/1:分署
        Select Case sm.UserInfo.LID
            Case 2
                Button2.Visible = False
                BtnSearch.Visible = False
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

        msg1.Text = ""
        tbDataGrid1.Visible = False

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

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
        'ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        'ddlGCODE = TIMS.Get_GOVCODE3(dtGCODE3, ddlGCODE, False)
        '一階審查結果 Result 初審建議結論/審查結果 --Y 通過、N 不通過、P 調整後通過
        'ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)

        '核班結果,核班結果'Y 通過、N 不通過
        ddlCURESULT = TIMS.Get_CURESULT(ddlCURESULT)

        '申復理由及說明
        'SFCONTREASONS.Text = ""
        '申復類別
        ddlSFCATELOG = TIMS.GET_SFCATELOG(objconn, ddlSFCATELOG)
        ddlSFCATELOG.SelectedIndex = -1
        Common.SetListItem(ddlSFCATELOG, "")

        '申復類別-其它-說明
        Call TIMS.Display_None(tr_SFCATELOG_OTH)
        '申復類別-其它-說明
        'txtSFCATELOG_OTH.Text = ""

        '申復核班結果
        ddlSFRESULT = TIMS.GET_SFRESULT(ddlSFRESULT)
        '申復核班結果
        ddlSFRESULT.SelectedIndex = -1
        Common.SetListItem(ddlSFRESULT, "")
        '申復未核班原因
        'SFRULTREASON.Text = ""
    End Sub

    Protected Sub BtnSearch_Click(sender As Object, e As EventArgs) Handles BtnSearch.Click
        Call sSearch1()
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
        '初審建議結論 '1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過
        'Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        '審查結果 '1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        'Dim v_RBL_RESULT_SCH As String = TIMS.GetListValue(RBL_RESULT_SCH)

        'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim v_RBL_CURESULT_SCH As String = TIMS.GetListValue(RBL_CURESULT_SCH)
        'SFRESULT:【申復核班結果】：1:不區分、2:有值、3:無值、Y:通過、N:不通過
        Dim v_RBL_SFRESULT_SCH As String = TIMS.GetListValue(RBL_SFRESULT_SCH)

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
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

        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

        Dim sql_WORG1 As String = String.Format("WITH WORG1 AS ({0}) {1}", GET_ORG_SQL1(), vbCrLf)

        'DECLARE @YEARS VARCHAR(4)='2021';DECLARE @TPLANID VARCHAR(3)='28';DECLARE @APPSTAGE NUMERIC(10,0)=2;
        Dim parms As New Hashtable From {
            {"YEARS", sm.UserInfo.Years},
            {"TPLANID", sm.UserInfo.TPlanID},
            {"APPSTAGE", v_APPSTAGE_SCH}
        }

        Dim sql As String = ""
        sql &= sql_WORG1
        sql &= " SELECT pp.YEARS ,dbo.FN_CYEAR2(pp.YEARS) YEARS_ROC" & vbCrLf
        sql &= " ,pp.APPSTAGE ,dbo.FN_GET_APPSTAGE(pp.APPSTAGE) APPSTAGE_N" & vbCrLf
        sql &= " ,pp.PLANNAME ,pp.PSNO28 ,pp.RID" & vbCrLf
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
        sql &= " ,gr1.DISTID GRDISTID" & vbCrLf

        sql &= " ,pf.CURESULT" & vbCrLf ' 核班結果 'Y 通過、N 不通過
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3'  END CURESULT_C" & vbCrLf
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' END CURESULT_N" & vbCrLf
        sql &= " ,pf.NGREASON" & vbCrLf '核班結果,未核班原因

        sql &= " ,pf.SFCONTREASONS" & vbCrLf '申復理由及說明
        sql &= " ,pf.SFLID" & vbCrLf '申復類別
        sql &= " ,pf.SFCATELOG_OTH" & vbCrLf '申復類別-其它 - 說明
        sql &= " ,pf.SFRESULT" & vbCrLf ' 申複核班結果 'Y 通過、N 不通過
        sql &= " ,CASE pf.SFRESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' END SFRESULT_N" & vbCrLf
        sql &= " ,pf.SFRULTREASON" & vbCrLf  '申復未核班原因

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        sql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=pp.TMID" & vbCrLf
        sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID AND dd.COMIDNO=pp.COMIDNO AND dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 AND LEN(pf.SFCONTREASONS)>1" & vbCrLf '申復理由及說明
        sql &= " JOIN dbo.ORG_SFCASEPI sf on sf.PLANID=pp.PLANID AND sf.COMIDNO=pp.COMIDNO AND sf.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " JOIN dbo.ORG_SFCASE ob on ob.SFCID=sf.SFCID AND ob.SFCSTATUS IS NOT NULL" & vbCrLf
        sql &= " JOIN WORG1 wo on wo.ORGID=pp.ORGID" & vbCrLf
        '19大類主責課程 SYS_GCODEREVIE
        'sql &= " JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf

        sql &= " WHERE pp.ISAPPRPAPER='Y' AND pp.PVR_ISAPPRPAPER='Y' AND pp.RESULTBUTTON IS NULL" & vbCrLf '審核送出(已送審) '正式
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sql &= " AND pp.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND pp.YEARS=@YEARS" & vbCrLf
        sql &= " AND pp.APPSTAGE=@APPSTAGE" & vbCrLf

        '初審建議結論'1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 / ST1RESULT
        'Dim v_RBL_ST1RESULT_SCH As String = TIMS.GetListValue(RBL_ST1RESULT_SCH)
        'Select Case v_RBL_ST1RESULT_SCH
        '    Case "2"
        '        sql &= " AND pf.ST1RESULT IS NOT NULL" & vbCrLf
        '    Case "3"
        '        sql &= " AND pf.ST1RESULT IS NULL" & vbCrLf
        '    Case "Y", "N", "P"
        '        sql &= " AND pf.ST1RESULT=@ST1RESULT" & vbCrLf
        '        parms.Add("ST1RESULT", v_RBL_ST1RESULT_SCH)
        'End Select
        '審查結果 '1:不區分2:有值3:無值Y:通過N:不通過P:調整後通過 /RESULT
        'Select Case v_RBL_RESULT_SCH
        '    Case "2"
        '        sql &= " AND pf.RESULT IS NOT NULL" & vbCrLf
        '    Case "3"
        '        sql &= " AND pf.RESULT IS NULL" & vbCrLf
        '    Case "Y", "N", "P"
        '        sql &= " AND pf.RESULT=@RESULT" & vbCrLf
        '        parms.Add("RESULT", v_RBL_RESULT_SCH)
        'End Select
        'CURESULT 核班結果, 核班結果'Y 通過、N 不通過
        Select Case v_RBL_CURESULT_SCH
            Case "2"
                sql &= " AND pf.CURESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.CURESULT IS NULL" & vbCrLf
            Case "Y", "N"
                sql &= " AND pf.CURESULT=@CURESULT" & vbCrLf
                parms.Add("CURESULT", v_RBL_CURESULT_SCH)
        End Select
        'SFRESULT:【申復核班結果】：1:不區分、2:有值、3:無值、Y:通過、N:不通過
        Select Case v_RBL_SFRESULT_SCH
            Case "2"
                sql &= " AND pf.SFRESULT IS NOT NULL" & vbCrLf
            Case "3"
                sql &= " AND pf.SFRESULT IS NULL" & vbCrLf
            Case "Y", "N"
                sql &= " AND pf.SFRESULT=@SFRESULT" & vbCrLf
                parms.Add("SFRESULT", v_RBL_SFRESULT_SCH)
        End Select

        Select Case v_APPSTAGE_SCH
            Case "3"
                If s_DISTID <> "" Then
                    sql &= " AND pp.DISTID=@DISTID" & vbCrLf
                    parms.Add("DISTID", s_DISTID)
                End If
            Case Else
                '篩選範圍 1:不區分 2:轄區單位 3:19大類主責課程 SYS_GCODEREVIE
                If s_DISTID <> "" AndAlso v_RBL_RANGE1_SCH <> "" Then
                    Select Case Val(v_RBL_RANGE1_SCH)
                        Case 1
                            sql &= " AND (pp.DISTID=@DISTID OR gr1.DISTID=@DISTID)" & vbCrLf
                            parms.Add("DISTID", s_DISTID)
                        Case 2
                            sql &= " AND pp.DISTID=@DISTID" & vbCrLf
                            parms.Add("DISTID", s_DISTID)
                        Case 3
                            sql &= " AND gr1.DISTID=@DISTID" & vbCrLf
                            parms.Add("DISTID", s_DISTID)
                    End Select
                End If
        End Select

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                sql &= " and wo.I_CROSSDIST !=-1" & vbCrLf
            Case "J" 'J:轄區提案單位
                sql &= " and wo.I_CROSSDIST =-1" & vbCrLf
        End Select

        '計畫'TRPlanPoint28
        If TRPlanPoint28.Visible Then
            Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
            Select Case v_rblOrgKind2'rblOrgKind2.SelectedValue
                Case "G", "W"
                    sql &= " and pp.ORGKIND2=@ORGKIND2" & vbCrLf
                    parms.Add("ORGKIND2", v_rblOrgKind2)
            End Select
        End If

        'STDate1.Text = TIMS.cdate3(STDate1.Text)
        'STDate2.Text = TIMS.cdate3(STDate2.Text)
        If STDate1.Text <> "" Then
            sql &= " and pp.STDATE >=@STDATE1" & vbCrLf
            parms.Add("STDATE1", TIMS.Cdate2(STDate1.Text))
        End If
        If STDate2.Text <> "" Then
            sql &= " and pp.STDATE <=@STDATE2" & vbCrLf
            parms.Add("STDATE2", TIMS.Cdate2(STDate2.Text))
        End If

        If RIDValue.Value <> "" AndAlso RIDValue.Value.Length > 1 Then
            sql &= " AND pp.RID=@RID" & vbCrLf
            parms.Add("RID", RIDValue.Value)
        End If

        TXT_PSNO28_SCH.Text = TIMS.Get_Substr1(TIMS.ClearSQM(TXT_PSNO28_SCH.Text), 11)
        If TXT_PSNO28_SCH.Text <> "" Then
            sql &= " AND pp.PSNO28=@PSNO28" & vbCrLf
            parms.Add("PSNO28", TXT_PSNO28_SCH.Text)
        End If
        'ROW_NUMBER() OVER(ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE) SEQNUM
        sql &= " ORDER BY pp.ORGNAME,pp.FIRSTSORT,pp.STDATE" & vbCrLf

        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    '尚未完成，設定19大類主責分署!
    Private Function CHECK_SYS_GCODEREVIE() As String
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then Return TIMS.cst_NODATAMsg2

        If v_APPSTAGE_SCH = "1" OrElse v_APPSTAGE_SCH = "2" Then
            Dim parms As New Hashtable
            parms.Add("YEARS", sm.UserInfo.Years)
            parms.Add("APPSTAGE", v_APPSTAGE_SCH)
            Dim sql As String = ""
            sql = "SELECT 1 FROM SYS_GCODEREVIE WHERE YEARS=@YEARS AND APPSTAGE=@APPSTAGE"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
            If dt Is Nothing OrElse dt.Rows.Count < 19 Then Return "尚未完成，設定19大類主責分署!"
        End If

        Return ""
    End Function

    Sub sSearch1()
        PanelSch1.Visible = True
        PanelEdit1.Visible = False

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        msg1.Text = TIMS.cst_NODATAMsg1
        tbDataGrid1.Visible = False

        Dim sERRMSG1 As String = CHECK_SYS_GCODEREVIE()
        If sERRMSG1 <> "" Then
            Common.MessageBox(Me, sERRMSG1)
            Return
        End If

        Dim dt As DataTable = SEARCH_DATA1_dt()
        If dt Is Nothing Then
            msg1.Text = TIMS.cst_NODATAMsg1
            Return
        End If
        If dt.Rows.Count = 0 Then
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
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Return
        Dim sCMDNM As String = e.CommandName
        If sCMDNM = "" Then Return

        Call CLEAR_DATA1()

        Hid_PSOID.Value = TIMS.GetMyValue(sCmdArg, "PSOID")
        Hid_PSNO28.Value = TIMS.GetMyValue(sCmdArg, "PSNO28")
        Hid_GCODE.Value = TIMS.GetMyValue(sCmdArg, "GCODE")
        Hid_PFGCODE.Value = TIMS.GetMyValue(sCmdArg, "PFGCODE")
        'Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))
        If Hid_PSNO28.Value = "" Then Return

        Select Case sCMDNM'e.CommandName
            Case "EDT1"
                btnSAVE1.Visible = True
                Dim drPP As DataRow = GET_DATA1()
                Call SHOW_DATA1(drPP)
                Call DISABLE_SHOW1(drPP) 'EDT1
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg9)
        End Select
    End Sub

    Private Sub DISABLE_SHOW1(ByRef drPP As DataRow)
        If drPP Is Nothing Then Return

        Dim flag_OTHFIXCONT_OK As Boolean = If(sm.UserInfo.DistID = Convert.ToString(drPP("DISTID")), True, False)
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
                'CURESULT 核班結果,核班結果 'Y 通過、N 不通過
                Dim labCURESULT_N As Label = e.Item.FindControl("labCURESULT_N")
                labCURESULT_N.Text = Convert.ToString(drv("CURESULT_N"))
                '申復核班結果'labSFRESULT_N 'Y 通過、N 不通過
                Dim labSFRESULT_N As Label = e.Item.FindControl("labSFRESULT_N")
                labSFRESULT_N.Text = Convert.ToString(drv("SFRESULT_N"))
                '編輯
                Dim BtnEDT1 As Button = e.Item.FindControl("BtnEDT1")
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
        lbD20KNAME.Text = ""

        'ST1SUGGEST.Text = ""
        'OTHFIXCONT.Text = ""
        'ddlST1RESULT.SelectedIndex = -1
        'Common.SetListItem(ddlST1RESULT, "")
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        'ddlGCODE.SelectedIndex = -1
        'Common.SetListItem(ddlGCODE, "")

        'CONFIRMCONT.Text = ""
        'COMMENTS.Text = "" 'Convert.ToString(dr("COMMENTS")) '委員審查意見與建議
        'ddlRESULT.SelectedIndex = -1
        'Common.SetListItem(ddlRESULT, "")

        'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        ddlCURESULT.SelectedIndex = -1
        '核班結果,未核班原因
        Common.SetListItem(ddlCURESULT, "")
        NGREASON.Text = ""
        '申復理由及說明
        SFCONTREASONS.Text = ""

        '申復類別
        ddlSFCATELOG.SelectedIndex = -1
        Common.SetListItem(ddlSFCATELOG, "")
        '申復類別-其它-說明
        Call TIMS.Display_None(tr_SFCATELOG_OTH)
        '申復類別-其它-說明
        SFCATELOG_OTH.Text = ""
        '申復核班結果
        ddlSFRESULT.SelectedIndex = -1
        Common.SetListItem(ddlSFRESULT, "")
        '申復未核班原因
        SFRULTREASON.Text = ""
    End Sub

    ''' <summary>單筆資料查詢(依 Hid_PSNO28.Value )</summary>
    ''' <returns></returns>
    Function GET_DATA1() As DataRow
        Dim dr1 As DataRow = Nothing
        If Hid_PSNO28.Value = "" Then Return dr1

        Dim parms As New Hashtable
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("PSNO28", Hid_PSNO28.Value)

        Dim sql As String = ""
        sql &= " SELECT rr.ORGPLANNAME" & vbCrLf '計畫別、" & vbCrLf
        '課程申請流水號" & vbCrLf
        sql &= " ,pp.ORGKINDGW,pp.PSNO28,pp.YEARS" & vbCrLf
        sql &= " ,pp.RID,pp.ORGNAME,pp.DISTID,pp.DISTNAME" & vbCrLf
        sql &= " ,(SELECT MAX(a.SFCPID) FROM ORG_SFCASEPI a WHERE a.PLANID=pp.PLANID AND a.COMIDNO=pp.COMIDNO AND a.SEQNO=pp.SEQNO) SFCPID" & vbCrLf
        'sql &= " ,pp.DISTID,pp.DISTNAME" & vbCrLf '分署別" & vbCrLf
        'sql &= " ,pp.ORGNAME" & vbCrLf '訓練單位名稱" & vbCrLf
        sql &= " ,pp.FIRSTSORT" & vbCrLf 'FIRSTSORT
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
        sql &= " ,(SELECT kc.CCNAME FROM dbo.KEY_CLASSCATELOG kc WHERE kc.CCID=pp.CLASSCATE) CCNAME" & vbCrLf '訓練職能" & vbCrLf
        sql &= " ,pp.PLANID,pp.COMIDNO,pp.SEQNO" & vbCrLf
        '訓練人次'訓練時數" & vbCrLf
        sql &= " ,pp.TNUM ,pp.THOURS" & vbCrLf
        sql &= " ,pp.ACTHUMCOST" & vbCrLf '實際人時成本" & vbCrLf
        sql &= " ,pp.METSUMCOST" & vbCrLf '實際材料費" & vbCrLf
        sql &= " ,dbo.FN_GET_KID20NAME(pp.PLANID,pp.COMIDNO,pp.SEQNO) D20KNAME" & vbCrLf '政府政策性產業
        sql &= " ,dd.D20KNAME1" & vbCrLf '5+2產業創新計畫"" & vbCrLf
        sql &= " ,dd.D20KNAME2" & vbCrLf '台灣AI行動計畫"" & vbCrLf'
        sql &= " ,dd.D20KNAME3" & vbCrLf '數位國家創新經濟發展方案"" & vbCrLf
        sql &= " ,dd.D20KNAME4" & vbCrLf '國家資通安全發展方案"" & vbCrLf
        sql &= " ,dd.D20KNAME5" & vbCrLf '前瞻基礎建設計畫"" & vbCrLf
        sql &= " ,dd.D20KNAME6" & vbCrLf '新南向政策"" & vbCrLf
        ' 5+2產業--新南向政策--台灣AI行動計畫--數位國家創新經濟發展方案--國家資通安全發展方案--前瞻基礎建設計畫" & vbCrLf
        sql &= " ,dbo.FN_GET_CROSSDIST(pp.YEARS,pp.COMIDNO,pp.APPSTAGE) CROSSDIST" & vbCrLf '是否跨區提案" & vbCrLf
        sql &= " ,pp.ICAPNUM" & vbCrLf 'iCAP標章證號" & vbCrLf

        sql &= " ,pf.PSOID" & vbCrLf
        sql &= " ,pf.ST1SUGGEST" & vbCrLf '初審幕僚建議/分署幕僚意見
        sql &= " ,pf.OTHFIXCONT" & vbCrLf '其他應修正內容" & vbCrLf
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
        sql &= " ,gr1.DISTID GRDISTID" & vbCrLf

        sql &= " ,pf.CURESULT" & vbCrLf ' 核班結果,核班結果'Y 通過、N 不通過
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '1' WHEN 'N' THEN '3'  END CURESULT_C" & vbCrLf
        sql &= " ,CASE pf.CURESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過'  END CURESULT_N" & vbCrLf
        sql &= " ,pf.NGREASON" & vbCrLf '核班結果,未核班原因
        '申復理由及說明
        sql &= " ,pf.SFCONTREASONS" & vbCrLf
        '申復類別/'申復類別-其它-說明
        sql &= " ,pf.SFLID ,pf.SFCATELOG_OTH" & vbCrLf
        '申複核班結果 'Y 通過、N 不通過
        sql &= " ,pf.SFRESULT" & vbCrLf
        sql &= " ,CASE pf.SFRESULT WHEN 'Y' THEN '通過' WHEN 'N' THEN '不通過' END SFRESULT_N" & vbCrLf
        '申復未核班原因
        sql &= " ,pf.SFRULTREASON" & vbCrLf

        sql &= " FROM dbo.VIEW2B pp" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME rr on rr.RID=pp.RID" & vbCrLf
        'sql &= " JOIN dbo.KEY_CLASSCATELOG kc on kc.CCID=pp.CLASSCATE" & vbCrLf
        sql &= " JOIN dbo.V_GOVCLASSCAST3 ig3 on ig3.GCID3=pp.GCID3" & vbCrLf
        sql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=pp.PLANID and dd.COMIDNO=pp.COMIDNO and dd.SEQNO=pp.SEQNO" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        '審查計分等級'19大類主責課程 SYS_GCODEREVIE
        'sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=ig3.GCODE31" & vbCrLf
        sql &= " LEFT JOIN dbo.SYS_GCODEREVIE gr1 on gr1.YEARS=pp.YEARS AND gr1.APPSTAGE=pp.APPSTAGE AND gr1.GCODE=pf.GCODE" & vbCrLf
        sql &= " LEFT JOIN dbo.V_GOVCLASS gc on gc.GCODE=pf.GCODE" & vbCrLf
        If Hid_PSOID.Value <> "" Then
            sql &= " AND pf.PSOID=@PSOID" & vbCrLf
            parms.Add("PSOID", Hid_PSOID.Value)
        End If
        '正式 'sql &= " AND CC.YEARS='2022'" & vbCrLf
        sql &= " WHERE pp.ISAPPRPAPER='Y' AND pp.PVR_ISAPPRPAPER='Y' AND (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND pp.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sql &= " AND pp.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND pp.PSNO28=@PSNO28" & vbCrLf

        dr1 = DbAccess.GetOneRow(sql, objconn, parms)

        Return dr1
    End Function

    Sub SHOW_DATA1(ByRef drPP As DataRow)
        PanelSch1.Visible = False
        PanelEdit1.Visible = True

        '核班結果
        ddlCURESULT.Enabled = False
        TIMS.Tooltip(ddlCURESULT, "僅提供顯示", True)
        '核班結果,未核班原因
        NGREASON.ReadOnly = True  '.Text = Convert.ToString(dr("NGREASON")) '核班結果,未核班原因
        NGREASON.ApplyStyle(TIMS.GET_RO_STYLE())
        '申復理由及說明
        SFCONTREASONS.ReadOnly = True
        SFCONTREASONS.ApplyStyle(TIMS.GET_RO_STYLE())

        If drPP Is Nothing Then Return

        center.Text = Convert.ToString(drPP("ORGNAME"))
        RIDValue.Value = Convert.ToString(drPP("RID"))
        Hid_ORGKINDGW.Value = Convert.ToString(drPP("ORGKINDGW"))
        Hid_SFCPID.Value = Convert.ToString(drPP("SFCPID"))
        Dim drPI As DataRow = Nothing
        If Hid_SFCPID.Value <> "" Then drPI = TIMS.GET_ORG_SFCASEPI_row(objconn, Val(Hid_SFCPID.Value))
        If Hid_SFCPID.Value <> "" AndAlso drPI IsNot Nothing Then
            Hid_SFCASENO.Value = Convert.ToString(drPI("SFCASENO"))
            Hid_SFCID.Value = Convert.ToString(drPI("SFCID"))
        End If
        Hid_PSOID.Value = Convert.ToString(drPP("PSOID"))
        Hid_PSNO28.Value = Convert.ToString(drPP("PSNO28"))

        lbYEARS_ROC.Text = TIMS.GET_YEARS_ROC(drPP("YEARS"))
        lbDistName.Text = Convert.ToString(drPP("DISTNAME"))
        lbOrgName.Text = Convert.ToString(drPP("ORGNAME"))
        lbPSNO28.Text = Convert.ToString(drPP("PSNO28"))
        lbClassName.Text = Convert.ToString(drPP("CLASSCNAME"))
        lbSFTDate.Text = String.Format("{0}~{1}", drPP("STDATE"), drPP("STDATE"))

        lbGCNAME.Text = Convert.ToString(drPP("GCNAME")) '訓練業別
        lbCCNAME.Text = Convert.ToString(drPP("CCNAME")) '訓練職能
        lbTNum.Text = Convert.ToString(drPP("TNUM"))
        lbTHours.Text = Convert.ToString(drPP("THOURS"))
        lbACTHUMCOST.Text = Convert.ToString(drPP("ACTHUMCOST"))
        lbMETSUMCOST.Text = Convert.ToString(drPP("METSUMCOST"))
        '是否跨區提案
        Dim s_CROSSDIST As String = If(Convert.ToString(drPP("CROSSDIST")) <> "", If(Val(drPP("CROSSDIST")) > -1, "是", "否"), "")
        lbIsCROSSDIST.Text = s_CROSSDIST 'Convert.ToString(dr("CROSSDIST")) '是否跨區提案
        lbiCAPNUM.Text = Convert.ToString(drPP("ICAPNUM"))
        lbD20KNAME.Text = TIMS.NullToStr(drPP("D20KNAME"), "無") '政府政策性產業

        '初審幕僚建議
        'ST1SUGGEST.Text = Convert.ToString(dr("ST1SUGGEST")) '分署幕僚意見
        'OTHFIXCONT.Text = Convert.ToString(dr("OTHFIXCONT")) '其他應修正內容
        'Common.SetListItem(ddlST1RESULT, dr("ST1RESULT")) '初審建議結論 --通過、不通過、調整後通過
        '審查課程職類  '分署確認課程分類 / 職類課程 / 訓練業別
        'Hid_GCODE.Value = Convert.ToString(dr("GCODE"))
        'Hid_PFGCODE.Value = Convert.ToString(dr("PFGCODE"))
        'Common.SetListItem(ddlGCODE, If(Hid_PFGCODE.Value <> "", Hid_PFGCODE.Value, Hid_GCODE.Value))
        'CONFIRMCONT.Text = Convert.ToString(dr("CONFIRMCONT")) '送請委員確認內容
        '審查結果
        'COMMENTS.Text = Convert.ToString(dr("COMMENTS")) '委員審查意見與建議
        'Common.SetListItem(ddlRESULT, dr("RESULT")) '審查結果

        '核班結果
        Common.SetListItem(ddlCURESULT, drPP("CURESULT")) 'CURESULT 核班結果,核班結果'Y 通過、N 不通過
        NGREASON.Text = Convert.ToString(drPP("NGREASON")) '核班結果,未核班原因
        '申復理由及說明
        SFCONTREASONS.Text = Convert.ToString(drPP("SFCONTREASONS")) '核班結果,未核班原因
        '申復類別
        Common.SetListItem(ddlSFCATELOG, Convert.ToString(drPP("SFLID")))
        '申復類別-其它 - 說明
        SFCATELOG_OTH.Text = Convert.ToString(drPP("SFCATELOG_OTH"))
        Call TIMS.Display_None(tr_SFCATELOG_OTH)
        If SFCATELOG_OTH.Text <> "" Then Call TIMS.Display_Inline(tr_SFCATELOG_OTH)
        '申復核班結果
        Common.SetListItem(ddlSFRESULT, Convert.ToString(drPP("SFRESULT")))
        '申復未核班原因
        SFRULTREASON.Text = Convert.ToString(drPP("SFRULTREASON"))
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

    ''' <summary>儲存前檢核</summary>
    ''' <param name="s_ERRMSG"></param>
    Private Sub CheckData1(ByRef s_ERRMSG As String)
        '核班結果,核班結果
        'Dim v_ddlCURESULT As String = TIMS.GetListValue(ddlCURESULT)
        'If NGREASON.Text <> "" Then NGREASON.Text = Trim(NGREASON.Text)
        'If v_ddlCURESULT = "" Then s_ERRMSG &= " 請選擇 核班結果/核班結果" & vbCrLf
        'If v_ddlCURESULT = "N" AndAlso NGREASON.Text = "" Then s_ERRMSG += "核班結果 為不通過/未核班原因 必填" & vbCrLf

        '申復類別-11(其它)
        Const cst_ddlSFCATELOG_ID_11 As String = "11"
        '申復類別
        Dim v_ddlSFCATELOG As String = TIMS.GetListValue(ddlSFCATELOG)
        Dim t_ddlSFCATELOG As String = TIMS.GetListText(ddlSFCATELOG)
        'tr_txtSFCATELOG_OTH '申復類別-其它-說明
        SFCATELOG_OTH.Text = TIMS.ClearSQM(SFCATELOG_OTH.Text)
        '申復核班結果
        Dim v_ddlSFRESULT As String = TIMS.GetListValue(ddlSFRESULT)
        '申復未核班原因
        SFRULTREASON.Text = TIMS.ClearSQM(SFRULTREASON.Text)

        If v_ddlSFCATELOG = "" Then s_ERRMSG &= " 請選擇 申復結果/申復類別" & vbCrLf
        If v_ddlSFCATELOG = cst_ddlSFCATELOG_ID_11 AndAlso SFCATELOG_OTH.Text = "" Then s_ERRMSG += String.Concat("申復類別 為", t_ddlSFCATELOG, "/申復類別-其它說明 必填", vbCrLf)
        If v_ddlSFRESULT = "" Then s_ERRMSG &= " 請選擇 申復結果/申復核班結果" & vbCrLf
        If v_ddlSFRESULT = "N" AndAlso SFRULTREASON.Text = "" Then s_ERRMSG += "申復核班結果 為不通過/申復未核班原因 必填" & vbCrLf
    End Sub

    Sub SAVE_DATA1()
        Hid_PSOID.Value = TIMS.ClearSQM(Hid_PSOID.Value)
        Hid_PSNO28.Value = TIMS.ClearSQM(Hid_PSNO28.Value)
        If Hid_PSOID.Value = "" OrElse Hid_PSNO28.Value = "" Then Return

        '核班結果,核班結果
        'Dim v_ddlCURESULT As String = TIMS.GetListValue(ddlCURESULT)
        '申復類別
        Dim v_ddlSFCATELOG As String = TIMS.GetListValue(ddlSFCATELOG)
        'Dim t_ddlSFCATELOG As String = TIMS.GetListText(ddlSFCATELOG)
        'tr_txtSFCATELOG_OTH '申復類別-其它-說明
        SFCATELOG_OTH.Text = TIMS.ClearSQM(SFCATELOG_OTH.Text)
        '申復核班結果
        Dim v_ddlSFRESULT As String = TIMS.GetListValue(ddlSFRESULT)
        '申復未核班原因
        SFRULTREASON.Text = TIMS.ClearSQM(SFRULTREASON.Text)

        Dim iRst As Integer = 0
        Dim iPSOID As Integer = Val(Hid_PSOID.Value)
        Dim parms As New Hashtable
        parms.Add("PSOID", Val(Hid_PSOID.Value))
        parms.Add("PSNO28", Hid_PSNO28.Value)

        Dim s_sql As String = ""
        s_sql &= " SELECT PSOID FROM PLAN_STAFFOPIN" & vbCrLf
        s_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(s_sql, objconn, parms)
        If dt.Rows.Count = 0 Then Return '檢核正確性

        Dim u_parms As New Hashtable
        u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)

        u_parms.Add("SFLID", v_ddlSFCATELOG)
        u_parms.Add("SFCATELOG_OTH", If(SFCATELOG_OTH.Text <> "", SFCATELOG_OTH.Text, Convert.DBNull))
        u_parms.Add("SFRESULT", v_ddlSFRESULT)
        u_parms.Add("SFRULTREASON", If(SFRULTREASON.Text <> "", SFRULTREASON.Text, Convert.DBNull))
        u_parms.Add("SFRULTACCT", sm.UserInfo.UserID)
        u_parms.Add("PSOID", iPSOID)
        u_parms.Add("PSNO28", Hid_PSNO28.Value)

        Dim u_sql As String = ""
        u_sql &= " UPDATE PLAN_STAFFOPIN" & vbCrLf
        u_sql &= " SET MODIFYDATE=GETDATE(),MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,SFLID=@SFLID ,SFCATELOG_OTH=@SFCATELOG_OTH" & vbCrLf
        u_sql &= " ,SFRESULT=@SFRESULT ,SFRULTREASON=@SFRULTREASON" & vbCrLf
        u_sql &= " ,SFRULTACCT=@SFRULTACCT ,SFRULTDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE PSOID=@PSOID AND PSNO28=@PSNO28" & vbCrLf
        iRst += DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        'Dim iRst As Integer = 0
        If iRst = 0 Then
            Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3b)
            Return
        End If
        Call sSearch1()
        Common.MessageBox(Me, TIMS.cst_SAVEOKMsg3)
    End Sub

    ''' <summary>檔案下載/檔案打包下載</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BTN_PACKAGE_DOWNLOAD1_Click(sender As Object, e As EventArgs) Handles BTN_PACKAGE_DOWNLOAD1.Click
        Hid_ORGKINDGW.Value = TIMS.ClearSQM(Hid_ORGKINDGW.Value)
        Hid_SFCASENO.Value = TIMS.ClearSQM(Hid_SFCASENO.Value)
        Hid_SFCID.Value = TIMS.ClearSQM(Hid_SFCID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim drOB As DataRow = TIMS.GET_ORG_SFCASE(objconn, RIDValue.Value, Hid_SFCID.Value, Hid_SFCASENO.Value)
        If Hid_SFCASENO.Value = "" OrElse Hid_SFCID.Value = "" Then
            Common.MessageBox(Me, "申復結果資訊有誤(案件號為空)，請重新操作!")
            Return
        ElseIf drOB Is Nothing Then
            Common.MessageBox(Me, "申復結果資訊有誤(查無案件編號)，請重新操作!")
            Return
        End If

        Dim rPMS As New Hashtable
        rPMS.Add("ORGKINDGW", Hid_ORGKINDGW.Value)
        rPMS.Add("RID", RIDValue.Value)
        rPMS.Add("SFCID", Hid_SFCID.Value)
        rPMS.Add("SFCASENO", Hid_SFCASENO.Value)
        rPMS.Add("ORGNAME", lbOrgName.Text)
        Call ResponseZIPFileALL_SF(Me, objconn, rPMS)
    End Sub

    ''' <summary>檔案下載/檔案打包下載</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="objconn"></param>
    ''' <param name="rPMS"></param>
    Public Shared Sub ResponseZIPFileALL_SF(MyPage As Page, objconn As SqlConnection, rPMS As Hashtable)
        Dim vSFCID As String = TIMS.GetMyValue2(rPMS, "SFCID")
        Dim vORGNAME As String = TIMS.GetMyValue2(rPMS, "ORGNAME")
        Dim vORGKINDGW As String = TIMS.GetMyValue2(rPMS, "ORGKINDGW")

        Dim dtFL As DataTable = TIMS.GET_ORG_SFCASEFL_dt(objconn, Val(vSFCID))
        Dim dtFLPI As DataTable = TIMS.GET_ORG_SFCASEFL_PI(objconn, Val(vSFCID))
        '申復線上送件
        If dtFL Is Nothing OrElse dtFL.Rows.Count = 0 Then
            Common.MessageBox(MyPage, "申復線上送件查無資料!")
            Return
        ElseIf dtFLPI Is Nothing OrElse dtFLPI.Rows.Count = 0 Then
            Common.MessageBox(MyPage, "計畫資訊查無資料!")
            Return
        End If

        Dim drFL1 As DataRow = dtFL.Rows(0)
        Dim vYEARS As String = Convert.ToString(drFL1("YEARS"))
        Dim vAPPSTAGE As String = Convert.ToString(drFL1("APPSTAGE"))
        Dim vPLANID As String = Convert.ToString(drFL1("PLANID"))
        Dim vRID As String = Convert.ToString(drFL1("RID"))
        Dim vSFCASENO As String = Convert.ToString(drFL1("SFCASENO"))
        Dim vYEARS_ROC As String = TIMS.GET_YEARS_ROC(drFL1("YEARS"))
        Dim vAPPSTAGE_S As String = TIMS.Get_APPSTAGE_S(drFL1("APPSTAGE"))
        Dim Template_ZipPath2 As String = TIMS.GET_Template_ZipPath2(vSFCID)
        '判斷是否有資料夾
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath2)) Then Directory.CreateDirectory(MyPage.Server.MapPath(Template_ZipPath2))

        'ORG_SFCASEFL
        If dtFL IsNot Nothing AndAlso dtFL.Rows.Count > 0 Then
            For Each drFL As DataRow In dtFL.Rows
                Dim vGWSFID As String = ""
                Dim vMEMO1 As String = ""
                Dim vSFIDNAME3 As String = ""
                Dim oFILENAME1 As String = "" 'Convert.ToString(drFL("FILENAME1"))
                Dim oFILEPATH1 As String = ""
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, "")
                Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                '年度申請階段_單位名稱_項目編號+項目名稱
                Dim t_FILENAME As String = "" 'String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vKBNAME2, ".pdf")
                Dim t_FilePath1 As String = "" 'Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), t_FILENAME))
                'Dim t_FilePath1 As String = Server.MapPath(Path.Combine(String.Concat(Template_ZipPath1, "/"), oFILENAME1))
                Try
                    vGWSFID = String.Concat(vORGKINDGW, drFL("SFID"))
                    vMEMO1 = Convert.ToString(drFL("MEMO1"))
                    Const cst_vMEMO1_maxlength As Integer = 9
                    If (vMEMO1 <> "" AndAlso vMEMO1.Length > cst_vMEMO1_maxlength) Then vMEMO1 = TIMS.Get_Substr1(vMEMO1, cst_vMEMO1_maxlength)
                    vSFIDNAME3 = Convert.ToString(drFL("SFIDNAME3"))
                    oFILENAME1 = Convert.ToString(drFL("FILENAME1"))
                    oFILEPATH1 = Convert.ToString(drFL("FILEPATH1"))
                    oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, ""))
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度_單位名稱_項目編號名稱
                    Select Case vGWSFID
                        Case TIMS.cst_SF_G04_其他佐證文件, TIMS.cst_SF_W04_其他佐證文件
                            t_FILENAME = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vSFIDNAME3, "_", vMEMO1, ".pdf"))
                        Case Else
                            t_FILENAME = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vSFIDNAME3, ".pdf"))
                    End Select
                    t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME))
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

        'ORG_SFCASEFL_PI
        If dtFLPI IsNot Nothing AndAlso dtFLPI.Rows.Count > 0 Then
            For Each drFLPI As DataRow In dtFLPI.Rows
                Dim oKSFID As String = ""
                Dim vPSNO28 As String = ""
                Dim vSFIDNAME As String = ""
                Dim oFILENAME1 As String = "" 'Convert.ToString(drFLTT("FILENAME1"))
                Dim oFILEPATH1 As String = ""
                Dim oUploadPath As String = "" 'TIMS.GET_UPLOADPATH1(oYEARS, oAPPSTAGE, oPLANID, oRID, oBCASENO, vKBSID)
                Dim s_FilePath1 As String = "" 'Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                Dim t_FILENAME_PI As String = ""
                Dim t_FilePath1 As String = ""
                Try
                    oKSFID = Convert.ToString(drFLPI("KSFID"))
                    vPSNO28 = Convert.ToString(drFLPI("PSNO28"))
                    vSFIDNAME = Convert.ToString(drFLPI("SFIDNAME"))
                    oFILENAME1 = Convert.ToString(drFLPI("FILENAME1"))
                    oFILEPATH1 = Convert.ToString(drFLPI("FILEPATH1"))
                    oUploadPath = If(oFILEPATH1 <> "", oFILEPATH1, TIMS.GET_UPLOADPATH1_SF(vYEARS, vAPPSTAGE, vPLANID, vRID, vSFCASENO, oKSFID))
                    s_FilePath1 = MyPage.Server.MapPath(Path.Combine(oUploadPath, oFILENAME1))
                    '年度申請階段_班級課程流水號_項目編號+項目名稱
                    t_FILENAME_PI = TIMS.GetValidFileName(String.Concat(vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vSFIDNAME, "_", vPSNO28, ".pdf"))
                    t_FilePath1 = MyPage.Server.MapPath(Path.Combine(String.Concat(Template_ZipPath2, "/"), t_FILENAME_PI))
                    If IO.File.Exists(s_FilePath1) Then
                        Dim dbyte As Byte() = File.ReadAllBytes(s_FilePath1)
                        File.WriteAllBytes(t_FilePath1, dbyte)
                    End If
                Catch ex As Exception
                    Dim strErrmsg As String = "/*Sub ResponseZIPFileALL(ByRef MyPage As Page)*/" & vbCrLf
                    strErrmsg &= String.Concat("vPSNO28: ", vPSNO28, vbCrLf, "vSFIDNAME: ", vSFIDNAME, vbCrLf, "oFILENAME1: ", oFILENAME1, vbCrLf)
                    strErrmsg &= String.Concat("s_FilePath1: ", s_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("t_FILENAME_PI: ", t_FILENAME_PI, vbCrLf)
                    strErrmsg &= String.Concat("t_FilePath1: ", t_FilePath1, vbCrLf)
                    strErrmsg &= String.Concat("ex.Message: ", ex.Message, vbCrLf)
                    Call TIMS.WriteTraceLog(strErrmsg, ex) 'Throw ex
                End Try
            Next
        End If

        Dim strNOW As String = DateTime.Now.ToString("yyyyMMddHHmmss")
        Dim zipFileName As String = TIMS.GetValidFileName(String.Concat("SF", vYEARS_ROC, vAPPSTAGE_S, "_", vORGNAME, "_", vSFCID, "_", strNOW, ".zip"))
        'Dim zipFileName As String = String.Concat("SF", vYEARS_ROC, "_", vDISTNAME3, "_", vORGNAME, "_", rSCDateNT, "_", strNOW, ".zip")
        If Not Directory.Exists(MyPage.Server.MapPath(Template_ZipPath2)) Then
            Common.MessageBox(MyPage, String.Concat(Template_ZipPath2, "下載檔案資料夾有誤!"))
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
                    Dim strErrmsg As String = "/*ResponseZIPFileALL_SF*/" & vbCrLf
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
            TIMS.SAVE_ADP_ZIPFILE(objconn, "-cr01008sf", File)
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

End Class
