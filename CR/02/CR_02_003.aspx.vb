Public Class CR_02_003
    Inherits AuthBasePage 'System.Web.UI.Page

    'OJT-22063001
    'iType 1:查詢用 11:匯出(1) 12:匯出(2)
    'Const cst_iType_查詢1 As Integer = 1
    'Const cst_iType_匯出1 As Integer = 11
    'Const cst_iType_匯出2 As Integer = 12
    'Const cst_SCORELEVEL_A As String = "A"
    'Const cst_SCORELEVEL_B As String = "B"
    'Const cst_SCORELEVEL_C As String = "C"
    'Const cst_SCORELEVEL_D As String = "D"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            cCreate1()
        End If

        '委訓
        'Select Case sm.UserInfo.LID
        '    Case 2
        '        Button2.Visible = False
        '    Case Else
        '        'Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        '        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
        '            '署(局) 或 分署(中心)
        '            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        '            If HistoryRID.Rows.Count <> 0 Then
        '                center.Attributes("onclick") = "showObj('HistoryList2');"
        '                center.Style("CURSOR") = "hand"
        '            End If
        '            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        '        End If
        'End Select
    End Sub

    Sub cCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, TIMS.Get_DISTIDT2(objconn))
        Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID)

        ddlAPPSTAGE_SCH = TIMS.Get_APPSTAGE2(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '訓練機構
        'center.Text = sm.UserInfo.OrgName
        'RIDValue.Value = sm.UserInfo.RID

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        'Common.SetListItem(rblOrgKind2, "G")
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        '開訓日期～ 

        '跨區/ 轄區提案 不區分跨區提案單位轄區提案單位 '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位

        '初審建議結論 --Y 通過、N 不通過、P 調整後通過
        'ddlST1RESULT = TIMS.Get_ST1RESULT(ddlST1RESULT)
        'Result 初審建議結論 / 審查結果 - -Y 通過、N 不通過、P 調整後通過
        'ddlRESULT = TIMS.Get_ST1RESULT(ddlRESULT)
    End Sub

    'Function GET_ORG_SQL1() As String
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
    '    sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
    '    sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
    '    Return sql
    'End Function

    Function GET_CLASS_SQL1(ByRef parms As Hashtable) As String
        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        '訓練機構
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        'Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH)

        '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        If v_APPSTAGE_SCH = "3" Then v_rblOrgKind2 = ""
        '開訓日期
        parms.Add("YEARS", sm.UserInfo.Years)
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("DISTID", v_ddlDISTID_SCH)
        parms.Add("APPSTAGE", v_APPSTAGE_SCH)
        '計畫'TRPlanPoint28
        If v_rblOrgKind2 <> "" Then parms.Add("ORGKIND2", v_rblOrgKind2)

        Dim sql As String = ""
        sql &= " SELECT cc.RID,cc.PSNO28,cc.DISTID,cc.DISTNAME" & vbCrLf
        sql &= " ,cc.APPSTAGE,cc.STDATE,cc.FTDATE" & vbCrLf
        sql &= " ,cc.TOTALCOST,cc.DEFGOVCOST,cc.DEFSTDCOST" & vbCrLf
        sql &= " ,cc.TPLANID,cc.YEARS " & vbCrLf
        sql &= " ,oo.RSID, OO.ORGLEVEL,oo.PLANID,cc.ORGKIND2,OO.ORGID" & vbCrLf
        sql &= " ,OO.COMIDNO" & vbCrLf
        sql &= " ,OO.ORGTYPENAME" & vbCrLf
        sql &= " ,OO.ORGNAME" & vbCrLf
        sql &= " ,OO.MASTERNAME" & vbCrLf
        'sql &= " ,dbo.FN_GET_CSTUDCNT14(cc.OCID) CSTUDCNT14" & vbCrLf
        sql &= " ,cc.TNUM CSTUDTNUM" & vbCrLf '訓練人次
        '建議結論 Y 通過、N 不通過、P 調整後通過
        'sql &= " AND cp.RESULT IN ('Y','P')" & vbCrLf
        sql &= " ,cp.RESULT" & vbCrLf
        sql &= " ,cp.CURESULT" & vbCrLf '改 核班結果：通過/不通過 判斷
        sql &= " FROM dbo.VIEW2B cc" & vbCrLf
        sql &= " JOIN dbo.VIEW_ORGPLANINFO oo on oo.RID=cc.RID" & vbCrLf
        sql &= " JOIN dbo.PLAN_STAFFOPIN cp on cp.PSNO28=cc.PSNO28" & vbCrLf
        sql &= " WHERE (cc.RESULTBUTTON IS NULL OR cc.APPLIEDRESULT='Y')" & vbCrLf '審核送出(已送審)
        sql &= " AND cc.PVR_ISAPPRPAPER='Y'" & vbCrLf '正式
        sql &= " AND cc.DATANOTSENT IS NULL" & vbCrLf '未檢送資料註記(排除有勾選)
        sql &= " AND OO.ORGLEVEL=2" & vbCrLf
        'sql &= " AND cc.TPLANID='28' AND cc.YEARS='2022' AND cc.DISTID='001' AND cc.APPSTAGE=1 AND cc.ORGKIND2='G'" & vbCrLf
        '建議結論 Y 通過、N 不通過、P 調整後通過
        'sql &= " AND cp.RESULT IN ('Y','P')" & vbCrLf
        sql &= " AND cc.YEARS=@YEARS" & vbCrLf
        sql &= " AND cc.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND cc.DISTID=@DISTID" & vbCrLf
        sql &= " AND cc.APPSTAGE=@APPSTAGE" & vbCrLf
        '計畫'TRPlanPoint28
        If v_rblOrgKind2 <> "" Then sql &= " AND cc.ORGKIND2=@ORGKIND2" & vbCrLf

        Return sql
    End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Function SEARCH_DATA1_dt(ByVal iType2 As Integer) As DataTable
        Dim dt As DataTable = Nothing

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = TIMS.cst_NODATAMsg2
            Return dt
        End If

        Dim parms As New Hashtable

        Dim sql As String = ""

        'iType2'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        Select Case iType2
            Case 1
                Dim sql_WC1 As String = String.Format("WITH WC1 AS ({0})", GET_CLASS_SQL1(parms))
                sql = ""
                sql &= sql_WC1 'WITH WC1 
                sql &= " SELECT cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf
                sql &= " ,COUNT(1) CLASSCNT1" & vbCrLf
                sql &= " ,ISNULL(SUM(cc.CSTUDTNUM),0) CSTUDTNUM" & vbCrLf
                sql &= " ,ISNULL(SUM(cc.DEFGOVCOST),0) DEFGOVCOST" & vbCrLf
                sql &= " FROM WC1 cc" & vbCrLf
                '/*建議結論 Y 通過、N 不通過、P 調整後通過*/
                'sql &= " WHERE cc.RESULT IN ('Y','P')" & vbCrLf
                '改 核班結果：通過/不通過 判斷
                sql &= " WHERE cc.CURESULT='Y'" & vbCrLf
                sql &= " GROUP BY cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf

            Case 3
                Dim sql_WC1 As String = String.Format("WITH WC1 AS ({0})", GET_CLASS_SQL1(parms))
                sql = ""
                sql &= sql_WC1 'WITH WC1 
                sql &= " SELECT cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf
                sql &= " ,COUNT(1) CLASSCNT1" & vbCrLf
                sql &= " ,ISNULL(SUM(cc.CSTUDTNUM),0) CSTUDTNUM" & vbCrLf
                sql &= " ,ISNULL(SUM(cc.DEFGOVCOST),0) DEFGOVCOST" & vbCrLf
                sql &= " FROM WC1 cc" & vbCrLf
                '/*建議結論 Y 通過、N 不通過、P 調整後通過*/
                'sql &= " WHERE cc.RESULT IN ('N')" & vbCrLf
                '改 核班結果：通過/不通過 判斷
                sql &= " WHERE cc.CURESULT='N'" & vbCrLf
                sql &= " GROUP BY cc.DISTID,cc.DISTNAME,cc.COMIDNO,cc.ORGNAME" & vbCrLf

        End Select

        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function


    '匯出  (請參考附件：表單05_通過課程審查結果彙整總表+明細表.xls、表單05_未通過課程審查結果彙整總表+明細表.xls)
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        Dim v_RBLExpType2 As String = TIMS.GetListValue(RBLExpType2)
        If v_RBLExpType2 = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return
        End If

        'iType2 'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        Dim iType2 As Integer = Val(v_RBLExpType2) '1/3
        Call EXPORT_5(iType2)
    End Sub

    '(請參考附件：表單05_通過課程審查結果彙整總表+明細表.xls、表單05_未通過課程審查結果彙整總表+明細表.xls)
    Sub EXPORT_5(ByVal iType2 As Integer)
        'Dim dtXls As DataTable = Nothing
        Dim dtXls As DataTable = SEARCH_DATA1_dt(iType2)
        If dtXls Is Nothing Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        If dtXls.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        '年度 + 申請階段 + 計畫 + 通過彙整總表 / 未通過彙整總表 / 通過明細表 / 未通過明細表
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_APPSTAGE_SCH) '申請階段
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        '111年度上半年提升勞工自主學習計畫通過課程彙整表
        If v_APPSTAGE_SCH = "3" Then v_rblOrgKind2 = ""

        'iType2 'RBLExpType2:1:通過彙整總表/2:通過明細表/3:未通過彙整總表/4:未通過明細表
        Dim sTYPE2_NM As String = If(iType2 = 1, "通過彙整總表", If(iType2 = 2, "通過明細表", If(iType2 = 3, "未通過彙整總表", If(iType2 = 4, "未通過明細表", "未通過明細表-"))))
        Dim sTitle1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, sTYPE2_NM)

        Dim s_DISTID_NM As String = TIMS.GetListText(ddlDISTID_SCH)

        '匯出excel /ods
        Dim s_FILENAME1 As String = String.Concat(sTYPE2_NM, "-", s_DISTID_NM, "_", TIMS.GetDateNo2(3))

        Dim v_ExpType As String = TIMS.GetListValue(RBListExpType) 'EXCEL/PDF/ODS
        Dim in_parms As New Hashtable
        in_parms.Clear()
        in_parms.Add("EXP", "Y") '匯出查詢條件
        in_parms.Add("ExpType", v_ExpType) '匯出查詢條件

        Dim parms As New Hashtable
        'parms.Clear()
        parms.Add("ExpType", v_ExpType) 'EXCEL/PDF/ODS
        parms.Add("FileName", s_FILENAME1)
        parms.Add("TitleName", TIMS.ClearSQM(sTitle1))
        Call EXPORT_5_13(dtXls, parms)

    End Sub

    Sub EXPORT_5_13(ByRef dtXls As DataTable, ByRef parms As Hashtable)
        Dim iCLASSCNT1 As Integer = 0 '訓練班次'CLASSCNT1
        Dim iCSTUDTNUM As Integer = 0  '訓練人次'CSTUDTNUM
        Dim iDEFGOVCOST As Integer = 0  '訓練補助費'DEFGOVCOST

        For Each dr1 As DataRow In dtXls.Rows
            iCLASSCNT1 += If(Val(dr1("CLASSCNT1")) > 0, Val(dr1("CLASSCNT1")), 0)
            iCSTUDTNUM += If(Val(dr1("CSTUDTNUM")) > 0, Val(dr1("CSTUDTNUM")), 0)
            iDEFGOVCOST += If(Val(dr1("DEFGOVCOST")) > 0, Val(dr1("DEFGOVCOST")), 0)
        Next

        Dim sPattern As String = "分署,訓練單位,班次,總訓練人次,總補助費(元)〈以訓練費用之80%估算〉"
        Dim sColumn As String = "DISTNAME,ORGNAME,CLASSCNT1,CSTUDTNUM,DEFGOVCOST"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        'Dim iColSpanCount As Integer = 5
        Dim iColSpanCount As Integer = sColumnA.Length

        Dim s_FootHtml2 As String = ""
        s_FootHtml2 &= "<tr>"
        s_FootHtml2 &= String.Format("<td colspan=2>{0}</td>", "合計") '合計
        s_FootHtml2 &= String.Format("<td>{0}</td>", iCLASSCNT1) '訓練班次
        s_FootHtml2 &= String.Format("<td>{0}</td>", iCSTUDTNUM) '訓練人次
        s_FootHtml2 &= String.Format("<td>{0}</td>", iDEFGOVCOST) '總補助費(元)〈以訓練費用之80%估算〉
        s_FootHtml2 &= "</tr>"

        'parms.Add("TitleHtml2", s_TitleHtml2)
        parms.Add("FootHtml2", s_FootHtml2)
        parms.Add("TitleColSpanCnt", iColSpanCount)
        parms.Add("sPatternA", sPatternA)
        parms.Add("sColumnA", sColumnA)
        TIMS.Utl_Export(Me, dtXls, parms)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

End Class
