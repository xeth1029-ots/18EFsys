Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CR_03_001
    Inherits AuthBasePage 'System.Web.UI.Page

    '114年確定性需求5：<系統> 產投兩計畫_報表1：訓練課程申請/核定差異統計表
    ' 共用設定
    Dim fontName As String = "標楷體"
    Dim fontSize12s As Single = 12.0F
    Dim print_lock As New Object '(); //lock

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            CCreate1()
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

    Sub CCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH)
        'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        ddlAPPSTAGE_SCH = TIMS.GET_APPSTAGE2_N34(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '訓練機構
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

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

        If TIMS.sUtl_ChkTest() Then
            Common.SetListItem(ddlAPPSTAGE_SCH, "1")
            Common.SetListItem(rblOrgKind2, "G")
            '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
            Common.SetListItem(RBL_CrossDist_SCH, "C")
        End If

    End Sub

    'Function GET_ORG_SQL1() As String
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql &= " SELECT dbo.FN_GET_CROSSDIST(@YEARS,oo.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
    '    sql &= " ,oo.COMIDNO,oo.ORGID" & vbCrLf
    '    sql &= " FROM ORG_ORGINFO oo WITH(NOLOCK)" & vbCrLf
    '    Return sql
    'End Function

    ''' <summary>'C:跨區提案單位</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Public Function GET_WX_SQL1(ByRef parms As Hashtable) As String
        Dim vSUM As String = TIMS.GetMyValue2(parms, "SUM")
        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "申請階段不可為空!")
            Common.MessageBox(Me, msg1.Text)
            Return "" 'dt
        End If
        Dim vSCORINGID As String = TIMS.GET_SCORINGID_VAL(sm.UserInfo.Years, v_APPSTAGE_SCH, objconn) '審查計分區間
        If vSCORINGID = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "審查計分區間不可為空!")
            Common.MessageBox(Me, msg1.Text)
            Return "" 'dt
        End If
        Dim drSCOR As DataRow = TIMS.GET_SCORINGID_DR(vSCORINGID, objconn) '審查計分區間
        If drSCOR Is Nothing Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "審查計分區間查無資料!-", vSCORINGID)
            Common.MessageBox(Me, msg1.Text)
            Return "" 'dt
        End If

        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)

        'parms = New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years.ToString()}, {"ORGKIND2", v_rblOrgKind2}, {"APPSTAGE", v_APPSTAGE_SCH}}
        parms = New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years.ToString()}, {"ORGKIND2", v_rblOrgKind2}, {"APPSTAGE", v_APPSTAGE_SCH}, {"SCORINGID", vSCORINGID}}

        Dim sSql As String = ""
        'VIEW_RIDNAME
        sSql &= " WITH WOD1 AS ( SELECT vr.RID,vr.ORGID,vr.COMIDNO,vr.ORGNAME,vr.ORGKIND2" & vbCrLf
        sSql &= " ,ip.TPLANID,ip.YEARS,ip.PLANID,ip.DISTID,ip.DISTNAME,IP.DISTNAME2" & vbCrLf
        'sSql &= " ,CASE WHEN dbo.FN_GET_CROSSDIST4(@YEARS,oo.COMIDNO,@APPSTAGE)>1 THEN 'Y' END CROSSDIST41" & vbCrLf
        sSql &= " ,dbo.FN_GET_CROSSDIST(@YEARS,vr.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
        sSql &= " ,sr.RLEVEL2" & vbCrLf
        sSql &= " ,dbo.FN_CLASSQUOTA(vr.ORGKIND2,ip.YEARS,@APPSTAGE,ISNULL(sr.RLEVEL2,'C')) CLASSQUOTA" & vbCrLf
        sSql &= " ,ROW_NUMBER() OVER (ORDER BY vr.ORGNAME ASC,vr.DISTID ASC) ROWSEQNO" & vbCrLf
        sSql &= " FROM VIEW_RIDNAME vr" & vbCrLf
        sSql &= " JOIN VIEW_PLAN ip on ip.PLANID=vr.PLANID" & vbCrLf
        sSql &= " LEFT JOIN VIEW_SCORING2 sr on sr.TPLANID=ip.TPLANID AND sr.DISTID=ip.DISTID AND sr.OrgID=vr.ORGID" & vbCrLf
        sSql &= " AND sr.SCORINGID=@SCORINGID" & vbCrLf
        sSql &= " WHERE ip.TPLANID=@TPLANID AND IP.YEARS=@YEARS AND vr.ORGKIND2=@ORGKIND2" & vbCrLf
        ' ORGKIND1/ORGTYPEID1 IN (7,8,9),7:全國性總工會,8:全國性各業聯合會,9:省級各業聯合會
        If v_rblOrgKind2 = "W" Then sSql &= " AND vr.ORGKIND1 IN (7,8,9)" & vbCrLf
        If s_DISTID <> "" AndAlso RIDValue.Value.Length = 1 AndAlso RIDValue.Value <> "A" Then
            parms.Add("DISTID", s_DISTID)
            sSql &= " AND IP.DISTID=@DISTID" & vbCrLf
        ElseIf s_DISTID <> "" AndAlso RIDValue.Value.Length > 1 Then
            parms.Add("DISTID", s_DISTID)
            parms.Add("RID", RIDValue.Value)
            sSql &= " AND IP.DISTID=@DISTID AND vr.RID=@RID" & vbCrLf
        End If
        sSql &= " )" & vbCrLf

        'VIEW2B
        sSql &= " ,WC1 AS ( SELECT pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.TPLANID,pp.DISTID" & vbCrLf
        sSql &= " ,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS,pp.TNUM" & vbCrLf
        sSql &= " ,PP.TOTALCOST,PP.DEFGOVCOST" & vbCrLf
        sSql &= " ,pf.CURESULT" & vbCrLf ' 核班結果,核班結果'Y 通過、N 不通過
        sSql &= " FROM VIEW2B pp" & vbCrLf
        sSql &= " JOIN WOD1 oo on oo.RID=pp.RID" & vbCrLf
        sSql &= " LEFT JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        sSql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')"   '審核送出(已送審)
        sSql &= " AND pp.PVR_ISAPPRPAPER='Y'"  '正式
        sSql &= " AND pp.DATANOTSENT IS NULL"  '未檢送資料註記(排除有勾選)
        'sSql &= " AND pp.TPLANID=@TPLANID" & vbCrLf sSql &= " AND pp.YEARS=@YEARS" & vbCrLf
        sSql &= " AND pp.APPSTAGE=@APPSTAGE )" & vbCrLf
        sSql &= " ,WC2 AS ( SELECT DISTINCT COMIDNO,PLANID FROM WC1)" & vbCrLf

        'WA1-第1階計算        
        sSql &= " ,WA1 AS ( SELECT oo.ROWSEQNO,oo.ORGID,oo.COMIDNO,oo.ORGNAME,oo.DISTID,oo.I_CROSSDIST
 ,ISNULL(oo.RLEVEL2,'C') RLEVEL2 ,oo.CLASSQUOTA
 ,ROW_NUMBER() OVER (PARTITION BY oo.COMIDNO ORDER BY ISNULL(oo.RLEVEL2,'C') ASC) PARROWID
 ,(SELECT COUNT(1) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID) CLSCNT
 ,(SELECT COUNT(1) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID AND x.CURESULT='Y') CLSCNTY
 ,(SELECT SUM(x.TNUM) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID) STDCNT
 ,(SELECT SUM(x.TNUM) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID AND x.CURESULT='Y') STDCNTY
 ,(SELECT SUM(x.DEFGOVCOST) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID) DEFGOVCOST
 ,(SELECT SUM(x.DEFGOVCOST) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID AND x.CURESULT='Y') DEFGOVCOSTY
 FROM WOD1 oo" & vbCrLf
        'sSql &= " JOIN WC2 ON WC2.COMIDNO=oo.COMIDNO AND WC2.PLANID=oo.PLANID" & vbCrLf
        If v_rblOrgKind2 = "G" Then
            sSql &= " WHERE oo.I_CROSSDIST!=-1" & vbCrLf 'I_CROSSDIST !=-1:跨區提案單位 'I_CROSSDIST =-1:轄區提案單位
        Else
            sSql &= " WHERE 1=1" & vbCrLf 'I_CROSSDIST !=-1:跨區提案單位 'I_CROSSDIST =-1:轄區提案單位
        End If
        sSql &= " AND EXISTS(SELECT 1 FROM WC2 WHERE WC2.COMIDNO=oo.COMIDNO AND WC2.PLANID=oo.PLANID)"
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        sSql &= " )" & vbCrLf

        If vSUM = "Y" Then
            ' 總合計算調整 {{"SUM", "Y"}}
            sSql &= " SELECT COUNT(DISTINCT COMIDNO) ORGCNT" & vbCrLf
            sSql &= " ,(SELECT COUNT(1) CNT FROM (SELECT COMIDNO,SUM(CLSCNT) SUMCLSCNT,SUM(CLSCNTY) SUMCLSCNTY FROM WA1 GROUP BY COMIDNO HAVING SUM(CLSCNTY)>0) G) ORGCNTY" & vbCrLf
            sSql &= " ,SUM(CLSCNT) CLSCNT ,SUM(CLSCNTY) CLSCNTY" & vbCrLf
            sSql &= " ,CASE WHEN SUM(CLSCNT)>0 THEN concat(FORMAT(100.0*SUM(CLSCNTY)/SUM(CLSCNT),'0.00'),'%') END CLSRATE2" & vbCrLf
            sSql &= " ,SUM(STDCNT) STDCNT ,SUM(STDCNTY) STDCNTY" & vbCrLf
            sSql &= " ,SUM(DEFGOVCOST) DEFGOVCOST ,SUM(DEFGOVCOSTY) DEFGOVCOSTY" & vbCrLf
            sSql &= " FROM WA1" & vbCrLf
            Return sSql
        End If

        'MAIN 'SELECT  ROW_NUMBER() OVER(ORDER BY a3.ORGNAME,a3.ORGID,a3.TBN,a3.DISTID) AS ROWSEQNO
        sSql &= " SELECT a3.ROWSEQNO,a3.ORGID,a3.COMIDNO,a3.ORGNAME,a3.DISTID,a3.PARROWID
,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=a3.DISTID) DISTNAME3
,(SELECT MASTERNAME FROM V_ORGINFO x WHERE x.COMIDNO=a3.COMIDNO) MASTERNAME
,a3.RLEVEL2,a3.CLASSQUOTA,a3.CLSCNT,a3.CLSCNTY
,a3.STDCNT,ISNULL(a3.STDCNTY,0) STDCNTY
,a3.DEFGOVCOST,ISNULL(a3.DEFGOVCOSTY,0) DEFGOVCOSTY 
,'0.00%' CLSRATE1
FROM WA1 a3
ORDER BY a3.ROWSEQNO" & vbCrLf

        Return sSql
    End Function

    ''' <summary>'J:轄區提案單位</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Public Function GET_WX_SQL2(ByRef parms As Hashtable) As String
        Dim vSUM As String = TIMS.GetMyValue2(parms, "SUM")
        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "申請階段不可為空!!")
            Common.MessageBox(Me, msg1.Text)
            Return "" 'dt
        End If
        Dim vSCORINGID As String = TIMS.GET_SCORINGID_VAL(sm.UserInfo.Years, v_APPSTAGE_SCH, objconn) '審查計分區間
        If vSCORINGID = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "審查計分區間不可為空!!")
            Common.MessageBox(Me, msg1.Text)
            Return "" 'dt
        End If
        Dim drSCOR As DataRow = TIMS.GET_SCORINGID_DR(vSCORINGID, objconn) '審查計分區間
        If drSCOR Is Nothing Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "審查計分區間查無資料!!-", vSCORINGID)
            Common.MessageBox(Me, msg1.Text)
            Return "" 'dt
        End If

        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)

        'parms = New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years.ToString()}, {"ORGKIND2", v_rblOrgKind2}, {"APPSTAGE", v_APPSTAGE_SCH}}
        parms = New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years.ToString()}, {"ORGKIND2", v_rblOrgKind2}, {"APPSTAGE", v_APPSTAGE_SCH}, {"SCORINGID", vSCORINGID}}

        Dim sSql As String = ""
        'VIEW_RIDNAME
        sSql &= " WITH WOD1 AS ( SELECT vr.RID,vr.ORGID,vr.COMIDNO,vr.ORGNAME,vr.ORGKIND2" & vbCrLf
        sSql &= " ,ip.TPLANID,ip.YEARS,ip.PLANID,ip.DISTID,ip.DISTNAME,IP.DISTNAME2" & vbCrLf
        'sSql &= " ,CASE WHEN dbo.FN_GET_CROSSDIST4(@YEARS,oo.COMIDNO,@APPSTAGE)>1 THEN 'Y' END CROSSDIST41" & vbCrLf
        sSql &= " ,dbo.FN_GET_CROSSDIST(@YEARS,vr.COMIDNO,@APPSTAGE) I_CROSSDIST" & vbCrLf
        sSql &= " ,sr.RLEVEL2" & vbCrLf
        sSql &= " ,dbo.FN_CLASSQUOTA(vr.ORGKIND2,ip.YEARS,@APPSTAGE,ISNULL(sr.RLEVEL2,'C')) CLASSQUOTA" & vbCrLf
        sSql &= " ,ROW_NUMBER() OVER (ORDER BY vr.ORGNAME ASC,vr.DISTID ASC) ROWSEQNO" & vbCrLf
        sSql &= " FROM VIEW_RIDNAME vr" & vbCrLf
        sSql &= " JOIN VIEW_PLAN ip on ip.PLANID=vr.PLANID" & vbCrLf
        sSql &= " LEFT JOIN VIEW_SCORING2 sr on sr.TPLANID=ip.TPLANID AND sr.DISTID=ip.DISTID AND sr.OrgID=vr.ORGID" & vbCrLf
        sSql &= " AND sr.SCORINGID=@SCORINGID" & vbCrLf
        sSql &= " WHERE ip.TPLANID=@TPLANID AND IP.YEARS=@YEARS AND vr.ORGKIND2=@ORGKIND2" & vbCrLf
        ' ORGKIND1/ORGTYPEID1 IN (10,11),10:各縣市總工會,11:各縣市職業工會
        ' ORGKIND1/ORGTYPEID1 IN (7,8,9),7:全國性總工會,8:全國性各業聯合會,9:省級各業聯合會
        If v_rblOrgKind2 = "W" Then sSql &= " AND vr.ORGKIND1 NOT IN (7,8,9)" & vbCrLf
        If s_DISTID <> "" AndAlso RIDValue.Value.Length = 1 AndAlso RIDValue.Value <> "A" Then
            parms.Add("DISTID", s_DISTID)
            sSql &= " AND IP.DISTID=@DISTID" & vbCrLf
        ElseIf s_DISTID <> "" AndAlso RIDValue.Value.Length > 1 Then
            parms.Add("DISTID", s_DISTID)
            parms.Add("RID", RIDValue.Value)
            sSql &= " AND IP.DISTID=@DISTID AND vr.RID=@RID" & vbCrLf
        End If
        sSql &= " )" & vbCrLf

        'VIEW2B
        sSql &= " ,WC1 AS ( SELECT pp.PCS,pp.PLANID,pp.COMIDNO,pp.SEQNO,pp.TPLANID,pp.DISTID" & vbCrLf
        sSql &= " ,PP.ISAPPRPAPER,PP.TRANSFLAG,PP.OCID,pp.ISSUCCESS,pp.TNUM" & vbCrLf
        sSql &= " ,PP.TOTALCOST,PP.DEFGOVCOST" & vbCrLf
        sSql &= " ,pf.CURESULT" & vbCrLf ' 核班結果,核班結果'Y 通過、N 不通過
        sSql &= " FROM VIEW2B pp" & vbCrLf
        sSql &= " JOIN WOD1 oo on oo.RID=pp.RID" & vbCrLf
        sSql &= " LEFT JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28" & vbCrLf
        sSql &= " WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y')"  '審核送出(已送審)
        sSql &= " AND pp.PVR_ISAPPRPAPER='Y'" '正式
        sSql &= " AND pp.DATANOTSENT IS NULL" '未檢送資料註記(排除有勾選)
        'sSql &= " AND pp.TPLANID=@TPLANID" & vbCrLf sSql &= " AND pp.YEARS=@YEARS" & vbCrLf
        sSql &= " AND pp.APPSTAGE=@APPSTAGE )" & vbCrLf
        sSql &= " ,WC2 AS ( SELECT DISTINCT COMIDNO,PLANID FROM WC1)" & vbCrLf

        'WA1-第1階計算        
        sSql &= " ,WA1 AS ( SELECT oo.ROWSEQNO,oo.ORGID,oo.COMIDNO,oo.ORGNAME,oo.DISTID,oo.I_CROSSDIST
 ,ISNULL(oo.RLEVEL2,'C') RLEVEL2 ,oo.CLASSQUOTA
 ,(SELECT COUNT(1) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID) CLSCNT
 ,(SELECT COUNT(1) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID AND x.CURESULT='Y') CLSCNTY
 ,(SELECT SUM(x.TNUM) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID) STDCNT
 ,(SELECT SUM(x.TNUM) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID AND x.CURESULT='Y') STDCNTY
 ,(SELECT SUM(x.DEFGOVCOST) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID) DEFGOVCOST
 ,(SELECT SUM(x.DEFGOVCOST) FROM WC1 x WHERE x.COMIDNO=oo.COMIDNO AND x.PLANID=oo.PLANID AND x.CURESULT='Y') DEFGOVCOSTY
 FROM WOD1 oo" & vbCrLf
        'sSql &= " JOIN WC2 ON WC2.COMIDNO=oo.COMIDNO AND WC2.PLANID=oo.PLANID" & vbCrLf
        If v_rblOrgKind2 = "G" Then
            sSql &= " WHERE oo.I_CROSSDIST=-1" & vbCrLf 'I_CROSSDIST !=-1:跨區提案單位 'I_CROSSDIST =-1:轄區提案單位
        Else
            sSql &= " WHERE 1=1" & vbCrLf 'I_CROSSDIST !=-1:跨區提案單位 'I_CROSSDIST =-1:轄區提案單位
        End If
        sSql &= " AND EXISTS(SELECT 1 FROM WC2 WHERE WC2.COMIDNO=oo.COMIDNO AND WC2.PLANID=oo.PLANID)"
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        sSql &= " )" & vbCrLf

        If vSUM = "Y" Then
            ' 總合計算調整 {{"SUM", "Y"}}
            sSql &= " SELECT COUNT(DISTINCT COMIDNO) ORGCNT" & vbCrLf
            sSql &= " ,(SELECT COUNT(1) CNT FROM (SELECT COMIDNO,SUM(CLSCNT) SUMCLSCNT,SUM(CLSCNTY) SUMCLSCNTY FROM WA1 GROUP BY COMIDNO HAVING SUM(CLSCNTY)>0) G) ORGCNTY" & vbCrLf
            sSql &= " ,SUM(CLSCNT) CLSCNT ,SUM(CLSCNTY) CLSCNTY" & vbCrLf
            sSql &= " ,CASE WHEN SUM(CLSCNT)>0 THEN concat(FORMAT(100.0*SUM(CLSCNTY)/SUM(CLSCNT),'0.00'),'%') END CLSRATE2" & vbCrLf
            sSql &= " ,SUM(STDCNT) STDCNT ,SUM(STDCNTY) STDCNTY" & vbCrLf
            sSql &= " ,SUM(DEFGOVCOST) DEFGOVCOST ,SUM(DEFGOVCOSTY) DEFGOVCOSTY" & vbCrLf
            sSql &= " FROM WA1" & vbCrLf
            Return sSql
        End If

        'MAIN 'SELECT  ROW_NUMBER() OVER(ORDER BY a3.ORGNAME,a3.ORGID,a3.TBN,a3.DISTID) AS ROWSEQNO
        sSql &= " SELECT a3.ROWSEQNO
,a3.ORGID,a3.COMIDNO,a3.ORGNAME,a3.DISTID
,(SELECT x.DISTNAME3 FROM V_DISTRICT x WHERE x.DISTID=a3.DISTID) DISTNAME3
,(SELECT MASTERNAME FROM V_ORGINFO x WHERE x.COMIDNO=a3.COMIDNO) MASTERNAME
,a3.RLEVEL2,a3.CLASSQUOTA,a3.CLSCNT,a3.CLSCNTY
,a3.STDCNT,ISNULL(a3.STDCNTY,0) STDCNTY
,a3.DEFGOVCOST,ISNULL(a3.DEFGOVCOSTY,0) DEFGOVCOSTY 
,'0.00%' CLSRATE1
FROM WA1 a3
ORDER BY a3.ROWSEQNO" & vbCrLf

        Return sSql
    End Function

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Public Function SEARCH_DATA1_dt() As DataTable
        'sm As SessionModel, hPMS As Hashtable
        Dim dt As DataTable = Nothing

        'Dim v_YEARS_SCH As String = TIMS.GetListValue(ddlYEARS_SCH) '年度
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "申請階段不可為空")
            Common.MessageBox(Me, msg1.Text)
            Return dt
        End If
        'Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        Dim vSCORINGID As String = TIMS.GET_SCORINGID_VAL(sm.UserInfo.Years, v_APPSTAGE_SCH, objconn) '審查計分區間
        'Dim drSCOR As DataRow = TIMS.GET_SCORINGID_DR(vSCORINGID, objconn) '審查計分區間
        If vSCORINGID = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "-", "審查計分區間不可為空")
            Common.MessageBox(Me, msg1.Text)
            Return dt
        End If

        '訓練機構
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
        '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)

        'sSql &= " DECLARE @TPLANID VARCHAR(4)='28';" & vbCrLf'sSql &= " DECLARE @V_YEARS VARCHAR(4)='2024';" & vbCrLf
        'sSql &= " DECLARE @V_APPSTAGE NUMERIC(4,0)=2;" & vbCrLf'sSql &= " DECLARE @V_ORGKIND2 VARCHAR(4)='G';" & vbCrLf

        Dim parms As New Hashtable
        Dim sSql As String = ""
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                sSql &= GET_WX_SQL1(parms)
            Case "J" 'J:轄區提案單位
                sSql &= GET_WX_SQL2(parms)
        End Select
        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms), vbCrLf, "--##CR_03_001 , RBL_CrossDist: ", v_RBL_CrossDist_SCH, ", sSql:", vbCrLf, sSql))
        End If
        'Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        If TIMS.dtHaveDATA(dt) Then

            For Each dr As DataRow In dt.Rows
                Dim iCLSCNT As Double = TIMS.GetValue2(dr("CLSCNT"))
                Dim iCLSCNTY As Double = TIMS.GetValue2(dr("CLSCNTY"))
                If iCLSCNT > 0 Then dr("CLSRATE1") = String.Concat(TIMS.ROUND(iCLSCNTY / iCLSCNT * 100, 2), "%")
            Next

        End If

        Return dt
    End Function

    Sub EXPXLSX_3(ws As ExcelWorksheet, dr2a As DataRow, idxStr As Integer, vOrgKind2 As String)
        'Dim idxStr As Integer = 5
        Dim a3 As String = String.Concat("申請訓練單位數(含跨區及非跨區)", vbCrLf, "(不重複)")
        Dim a6 As String = String.Concat("跨區申請訓練單位數", vbCrLf, "(不重複)")
        Dim aVAL As String = If(idxStr = 3, a3, a6)

        Select Case vOrgKind2
            Case "G"
                Using exlRow1 As ExcelRange = ws.Cells(String.Format("A{0}:C{0}", idxStr))
                    With exlRow1
                        .Merge = True
                        .Style.Font.Name = fontName
                        .Style.Font.Size = 12
                        .Style.Font.Bold = True
                        .Value = aVAL 'String.Concat("跨區申請訓練單位數", vbLf, "(不重複)")
                        .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                        .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                        .AutoFitColumns(25.0, 250.0)
                    End With
                End Using
                'SetCellValue(ws, "D" & idxStr, String.Concat("審查計分表", vbLf, "等級")) '15
                'SetCellValue(ws, "E" & idxStr, "可核配上限") '20
                SetCellValue(ws, "D" & idxStr, String.Concat("申請", vbLf, "班次")) '10
                SetCellValue(ws, "E" & idxStr, String.Concat("核定", vbLf, "班次")) '10
                SetCellValue(ws, "F" & idxStr, String.Concat("核班", vbLf, "比率")) '20
                SetCellValue(ws, "G" & idxStr, String.Concat("申請", vbLf, "人次")) '10
                SetCellValue(ws, "H" & idxStr, String.Concat("核定", vbLf, "人次")) '20
                SetCellValue(ws, "I" & idxStr, "申請補助費") '20
                'SetCellValue(ws, "J" & idxStr, String.Concat("核定補助費")) '20
                Using exlRow1 As ExcelRange = ws.Cells(String.Format("J{0}:L{0}", idxStr))
                    With exlRow1
                        .Merge = True
                        .Style.Font.Name = fontName
                        .Style.Font.Size = 12
                        .Value = "核定補助費"
                        .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                        .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                        .AutoFitColumns(25.0, 250.0)
                    End With
                End Using

                Dim cellsCOLSPNumF2 As String = String.Concat("A{0}:L{0}") '(畫格子使用)
                SetCellBorder(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)))
                ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).Style.Font.Bold = True
                ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).AutoFitColumns(25.0, 250.0)
                '增加列的高度
                ws.Row(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).Start.Row).Height = 43

                'SetCellValue(ws, "D" & idxStr, dr1("RLEVEL2")) '"審查計分表等級")
                'SetCellValue(ws, "E" & idxStr, dr1("CLASSQUOTA"), ExcelHorizontalAlignment.Right) '"可核配上限")
                idxStr += 1
                Dim sV2A As String = String.Format("申請{0}/核定{1}", dr2a("ORGCNT"), dr2a("ORGCNTY"))
                SetCellValue(ws, String.Format("A{0}:C{0}", idxStr), sV2A) '"申請{0}/核定")
                ws.Cells(String.Format("A{0}:C{0}", idxStr)).Merge = True
                SetCellValue(ws, "D" & idxStr, dr2a("CLSCNT")) '"申請班次")
                SetCellValue(ws, "E" & idxStr, dr2a("CLSCNTY")) '"核定班次")
                SetCellValue(ws, "F" & idxStr, dr2a("CLSRATE2")) '"核班率
                SetCellValue(ws, "G" & idxStr, dr2a("STDCNT")) '"申請人次
                SetCellValue(ws, "H" & idxStr, dr2a("STDCNTY")) '"核定人次
                SetCellValue(ws, "I" & idxStr, dr2a("DEFGOVCOST")) '"申請補助費")
                SetCellValue(ws, String.Format("J{0}:L{0}", idxStr), dr2a("DEFGOVCOSTY")) '"核定補助費")
                ws.Cells(String.Format("J{0}:L{0}", idxStr)).Merge = True
                SetCellBorder(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)))
                ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).AutoFitColumns(25.0, 250.0)
                '增加列的高度
                ws.Row(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).Start.Row).Height = 33

                ws.Cells("I" & idxStr).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
                ws.Cells(String.Format("J{0}:L{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改

            Case "W"
                Using exlRow1 As ExcelRange = ws.Cells(String.Format("A{0}:D{0}", idxStr))
                    With exlRow1
                        .Merge = True
                        .Style.Font.Name = fontName
                        .Style.Font.Size = 12
                        .Style.Font.Bold = True
                        .Value = aVAL 'String.Concat("跨區申請訓練單位數", vbLf, "(不重複)")
                        .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                        .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                        .AutoFitColumns(25.0, 250.0)
                    End With
                End Using
                SetCellValue(ws, "E" & idxStr, String.Concat("申請", vbLf, "班次")) '10
                SetCellValue(ws, "F" & idxStr, String.Concat("核定", vbLf, "班次")) '10
                SetCellValue(ws, "G" & idxStr, String.Concat("核班", vbLf, "比率")) '20
                SetCellValue(ws, "H" & idxStr, String.Concat("申請", vbLf, "人次")) '10
                SetCellValue(ws, "I" & idxStr, String.Concat("核定", vbLf, "人次")) '20
                SetCellValue(ws, "J" & idxStr, "申請補助費") '20
                'SetCellValue(ws, "J" & idxStr, String.Concat("核定補助費")) '20
                Using exlRow1 As ExcelRange = ws.Cells(String.Format("K{0}:M{0}", idxStr))
                    With exlRow1
                        .Merge = True
                        .Style.Font.Name = fontName
                        .Style.Font.Size = 12
                        .Value = "核定補助費"
                        .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                        .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                        .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                        .AutoFitColumns(25.0, 250.0)
                    End With
                End Using
                Dim cellsCOLSPNumF2 As String = String.Concat("A{0}:M{0}") '(畫格子使用)
                SetCellBorder(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)))
                ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).Style.Font.Bold = True
                ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).AutoFitColumns(25.0, 250.0)
                '增加列的高度
                ws.Row(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).Start.Row).Height = 43

                'SetCellValue(ws, "D" & idxStr, dr1("RLEVEL2")) '"審查計分表等級")
                'SetCellValue(ws, "E" & idxStr, dr1("CLASSQUOTA"), ExcelHorizontalAlignment.Right) '"可核配上限")
                idxStr += 1
                Dim sV2A As String = String.Format("申請{0}/核定{1}", dr2a("ORGCNT"), dr2a("ORGCNTY"))
                SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), sV2A) '"申請{0}/核定")
                ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
                SetCellValue(ws, "E" & idxStr, dr2a("CLSCNT")) '"申請班次")
                SetCellValue(ws, "F" & idxStr, dr2a("CLSCNTY")) '"核定班次")
                SetCellValue(ws, "G" & idxStr, dr2a("CLSRATE2")) '"核班率
                SetCellValue(ws, "H" & idxStr, dr2a("STDCNT")) '"申請人次
                SetCellValue(ws, "I" & idxStr, dr2a("STDCNTY")) '"核定人次
                SetCellValue(ws, "J" & idxStr, dr2a("DEFGOVCOST")) '"申請補助費")
                SetCellValue(ws, String.Format("K{0}:M{0}", idxStr), dr2a("DEFGOVCOSTY")) '"核定補助費")
                ws.Cells(String.Format("K{0}:M{0}", idxStr)).Merge = True
                SetCellBorder(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)))
                ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).AutoFitColumns(25.0, 250.0)
                '增加列的高度
                ws.Row(ws.Cells(String.Format(cellsCOLSPNumF2, idxStr)).Start.Row).Height = 33

                ws.Cells("J" & idxStr).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
                ws.Cells(String.Format("K{0}:M{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改

        End Select

    End Sub

    ''' <summary>'C:跨區提案單位 </summary>
    Sub EXPXLSX_1CG()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!")
            Exit Sub
        End If

        Dim dtXls2a As DataTable = Nothing 'C:跨區提案單位
        Dim dtXls2b As DataTable = Nothing 'J:轄區提案單位
        dtXls2a = ALL_DATA_TOTAL_dt2(dtXls2b)
        If TIMS.dtNODATA(dtXls2a) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If
        If TIMS.dtNODATA(dtXls2b) Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        Dim dr2a As DataRow = dtXls2a.Rows(0) 'C:跨區提案單位
        Dim dr2b As DataRow = dtXls2b.Rows(0) 'J:轄區提案單位'申請訓練單位數(含跨區及非跨區)

        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        Dim V_ORGKIND2_NM1 As String = If(v_rblOrgKind2 = "G", "產投", If(v_rblOrgKind2 = "W", "自主", "")) '113下自主-北分署
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫

        Dim END_COL_NM As String = "L"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim strErrmsg As String = ""

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Dim V_CROSSDIST_NM1 As String = If(v_RBL_CrossDist_SCH = "C", "跨區", If(v_RBL_CrossDist_SCH = "J", "轄區", "不區分"))
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim V_APPSTAGE_NM1 As String = If(v_ddlAPPSTAGE_SCH = "1", "上", If(v_ddlAPPSTAGE_SCH = "2", "下", "X"))
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度

        '"產投差異表(依明細)-113下(跨區)"
        Dim V_SHEETNM1 As String = String.Concat(V_ORGKIND2_NM1, "差異表(依明細)-", s_ROCYEAR1, V_APPSTAGE_NM1, "(", V_CROSSDIST_NM1, ")")
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "訓練課程申請/核定差異統計表")
        Dim s_TITLENAME2 As String = "跨區單位明細"
        '114年度上半年產投計畫申請核定差異統計表_XXXXXXXXX
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "申請核定差異統計表_", TIMS.GetDateNo())

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            'Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage()

            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定 'Dim fontName As String = "標楷體" 'Dim fontSize12s As Single = 12.0F '報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "2"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 10
                    .Value = String.Concat(Now.Year - 1911, Now.ToString(".MM.dd"))
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    '.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Call EXPXLSX_3(ws, dr2b, 3, v_rblOrgKind2)
            Call EXPXLSX_3(ws, dr2a, 6, v_rblOrgKind2)

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "8"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME2
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
            Dim idxStr As Integer = 9
            SetCellValue(ws, "A" & idxStr, "序號") '10
            SetCellValue(ws, "B" & idxStr, "訓練單位") '33
            SetCellValue(ws, "C" & idxStr, "分署") '10
            SetCellValue(ws, "D" & idxStr, String.Concat("審查計分表", vbLf, "等級")) '15
            SetCellValue(ws, "E" & idxStr, "可核配上限") '20
            SetCellValue(ws, "F" & idxStr, String.Concat("申請", vbLf, "班次")) '10
            SetCellValue(ws, "G" & idxStr, String.Concat("核定", vbLf, "班次")) '10
            SetCellValue(ws, "H" & idxStr, String.Concat("申請", vbLf, "人次")) '10
            SetCellValue(ws, "I" & idxStr, "核定人次") '20
            SetCellValue(ws, "J" & idxStr, String.Concat("申請", vbLf, "補助費")) '20
            SetCellValue(ws, "K" & idxStr, String.Concat("核定", vbLf, "補助費")) '20
            SetCellValue(ws, "L" & idxStr, "核班率") '20
            ws.Cells(String.Format("A{0}:L{0}", 9)).Style.Font.Bold = True

            idxStr = 10
            Dim idx1 As Integer = idxStr
            Dim idx2b As Integer = 0
            Dim V_COMIDNO As String = dtXls.Rows(0)("COMIDNO")
            Dim iCNT_G As Integer = 0 '(-群組筆數跳號)
            Dim iCNTX As Integer = 0 '(-筆數跳號)
            Dim iCNT_O As Integer = 1 '(-機構數跳號)
            'Dim iTHOURS As Integer = 0
            'Dim iTNUM As Integer = 0
            'Dim iTOTAL As Integer = 0
            'Dim iTOTALCOST As Integer = 0
            'Dim iDEFGOVCOST As Integer = 0
            Dim i_CLASSQUOTA As Integer = 0 '"可核配上限")
            Dim i_CLSCNT As Integer = 0 '"申請班次")
            Dim i_CLSCNTY As Integer = 0 '"核定班次")
            Dim i_STDCNT As Integer = 0 '"申請人次
            Dim i_STDCNTY As Integer = 0 '"核定人次
            Dim i_DEFGOVCOST As Integer = 0 '"申請補助費")
            Dim i_DEFGOVCOSTY As Integer = 0 '"核定補助費")
            '(總計)
            Dim i_CLASSQUOTA_A As Integer = 0 '"可核配上限")
            Dim i_CLSCNT_A As Integer = 0 '"申請班次")
            Dim i_CLSCNTY_A As Integer = 0 '"核定班次")
            Dim i_STDCNT_A As Integer = 0 '"申請人次
            Dim i_STDCNTY_A As Integer = 0 '"核定人次
            Dim i_DEFGOVCOST_A As Integer = 0 '"申請補助費")
            Dim i_DEFGOVCOSTY_A As Integer = 0 '"核定補助費")

            Dim v_CLSRATE1 As String = "" '"核班率")

            For Each dr1 As DataRow In dtXls.Rows
                iCNTX += 1
                'iTHOURS += TIMS.VAL1(dr1("THOURS")) 'iTNUM += TIMS.VAL1(dr1("TNUM"))
                'iTOTAL += TIMS.VAL1(dr1("TOTAL"))'iTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))'iDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
                If V_COMIDNO = dr1("COMIDNO") Then
                    iCNT_G += 1
                    If TIMS.GetValue2(dr1("PARROWID")) <= 3 Then i_CLASSQUOTA += TIMS.GetValue2(dr1("CLASSQUOTA"))
                    'i_CLASSQUOTA += TIMS.GetValue2(dr1("CLASSQUOTA"))
                    i_CLSCNT += TIMS.GetValue2(dr1("CLSCNT"))
                    i_CLSCNTY += TIMS.GetValue2(dr1("CLSCNTY"))
                    i_STDCNT += TIMS.GetValue2(dr1("STDCNT"))
                    i_STDCNTY += TIMS.GetValue2(dr1("STDCNTY"))
                    i_DEFGOVCOST += TIMS.GetValue2(dr1("DEFGOVCOST"))
                    i_DEFGOVCOSTY += TIMS.GetValue2(dr1("DEFGOVCOSTY"))
                Else
                    '(合併儲存格)
                    idx2b = idx1 + iCNT_G - 1
                    ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
                    ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True

                    '(小計) '
                    ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
                    SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), "小計")
                    SetCellValue(ws, "E" & idxStr, i_CLASSQUOTA, ExcelHorizontalAlignment.Right) '"可核配上限")
                    SetCellValue(ws, "F" & idxStr, i_CLSCNT, ExcelHorizontalAlignment.Right) '"申請班次")
                    SetCellValue(ws, "G" & idxStr, i_CLSCNTY, ExcelHorizontalAlignment.Right) '"核定班次")
                    SetCellValue(ws, "H" & idxStr, i_STDCNT, ExcelHorizontalAlignment.Right) '"申請人次
                    SetCellValue(ws, "I" & idxStr, i_STDCNTY, ExcelHorizontalAlignment.Right) '"核定人次
                    SetCellValue(ws, "J" & idxStr, i_DEFGOVCOST, ExcelHorizontalAlignment.Right) '"申請補助費")
                    SetCellValue(ws, "K" & idxStr, i_DEFGOVCOSTY, ExcelHorizontalAlignment.Right) '"核定補助費")
                    v_CLSRATE1 = String.Concat(If(i_CLSCNT > 0, TIMS.ROUND(100.0 * (i_CLSCNTY / i_CLSCNT), 2), "0.00"), "%") '核班率
                    SetCellValue(ws, "L" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
                    ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

                    idxStr += 1 '(小計)'(總計)
                    i_CLASSQUOTA_A += i_CLASSQUOTA
                    i_CLSCNT_A += i_CLSCNT
                    i_CLSCNTY_A += i_CLSCNTY
                    i_STDCNT_A += i_STDCNT
                    i_STDCNTY_A += i_STDCNTY
                    i_DEFGOVCOST_A += i_DEFGOVCOST
                    i_DEFGOVCOSTY_A += i_DEFGOVCOSTY

                    idx1 = idxStr
                    V_COMIDNO = dr1("COMIDNO")
                    iCNT_G = 1
                    iCNT_O += 1
                    i_CLASSQUOTA = If(TIMS.GetValue2(dr1("PARROWID")) <= 3, TIMS.GetValue2(dr1("CLASSQUOTA")), 0)
                    'i_CLASSQUOTA = TIMS.GetValue2(dr1("CLASSQUOTA"))
                    i_CLSCNT = TIMS.GetValue2(dr1("CLSCNT"))
                    i_CLSCNTY = TIMS.GetValue2(dr1("CLSCNTY"))
                    i_STDCNT = TIMS.GetValue2(dr1("STDCNT"))
                    i_STDCNTY = TIMS.GetValue2(dr1("STDCNTY"))
                    i_DEFGOVCOST = TIMS.GetValue2(dr1("DEFGOVCOST"))
                    i_DEFGOVCOSTY = TIMS.GetValue2(dr1("DEFGOVCOSTY"))
                End If

                '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
                ' ROWSEQNO,ORGNAME,DISTNAME3,RLEVEL2,CLASSQUOTA,CLSCNT,CLSCNTY,STDCNT,STDCNTY,DEFGOVCOST,DEFGOVCOSTY,CLSRATE1
                SetCellValue(ws, "A" & idxStr, iCNT_O) '"訓練單位名稱")
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '"訓練單位")
                SetCellValue(ws, "C" & idxStr, dr1("DISTNAME3")) '"分署")
                SetCellValue(ws, "D" & idxStr, dr1("RLEVEL2")) '"審查計分表等級")
                SetCellValue(ws, "E" & idxStr, dr1("CLASSQUOTA"), ExcelHorizontalAlignment.Right) '"可核配上限")
                SetCellValue(ws, "F" & idxStr, dr1("CLSCNT"), ExcelHorizontalAlignment.Right) '"申請班次")
                SetCellValue(ws, "G" & idxStr, dr1("CLSCNTY"), ExcelHorizontalAlignment.Right) '"核定班次")
                SetCellValue(ws, "H" & idxStr, dr1("STDCNT"), ExcelHorizontalAlignment.Right) '"申請人次
                SetCellValue(ws, "I" & idxStr, dr1("STDCNTY"), ExcelHorizontalAlignment.Right) '"核定人次
                SetCellValue(ws, "J" & idxStr, dr1("DEFGOVCOST"), ExcelHorizontalAlignment.Right) '"申請補助費")
                SetCellValue(ws, "K" & idxStr, dr1("DEFGOVCOSTY"), ExcelHorizontalAlignment.Right) '"核定補助費")
                SetCellValue(ws, "L" & idxStr, dr1("CLSRATE1"), ExcelHorizontalAlignment.Right) '"核班率")

                idxStr += 1
            Next
            '(合併儲存格) idxStr += 1
            idx2b = idx1 + iCNT_G - 1
            ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
            ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True

            '(小計) '
            ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
            SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), "小計")
            SetCellValue(ws, "E" & idxStr, i_CLASSQUOTA, ExcelHorizontalAlignment.Right) '"可核配上限")
            SetCellValue(ws, "F" & idxStr, i_CLSCNT, ExcelHorizontalAlignment.Right) '"申請班次")
            SetCellValue(ws, "G" & idxStr, i_CLSCNTY, ExcelHorizontalAlignment.Right) '"核定班次")
            SetCellValue(ws, "H" & idxStr, i_STDCNT, ExcelHorizontalAlignment.Right) '"申請人次
            SetCellValue(ws, "I" & idxStr, i_STDCNTY, ExcelHorizontalAlignment.Right) '"核定人次
            SetCellValue(ws, "J" & idxStr, i_DEFGOVCOST, ExcelHorizontalAlignment.Right) '"申請補助費")
            SetCellValue(ws, "K" & idxStr, i_DEFGOVCOSTY, ExcelHorizontalAlignment.Right) '"核定補助費")
            v_CLSRATE1 = String.Concat(If(i_CLSCNT > 0, TIMS.ROUND(100.0 * (i_CLSCNTY / i_CLSCNT), 2), "0.00"), "%") '核班率
            SetCellValue(ws, "L" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
            ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

            '(總計)
            i_CLASSQUOTA_A += i_CLASSQUOTA
            i_CLSCNT_A += i_CLSCNT
            i_CLSCNTY_A += i_CLSCNTY
            i_STDCNT_A += i_STDCNT
            i_STDCNTY_A += i_STDCNTY
            i_DEFGOVCOST_A += i_DEFGOVCOST
            i_DEFGOVCOSTY_A += i_DEFGOVCOSTY
            '(總計)
            idxStr += 1
            ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
            SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), "總計")
            SetCellValue(ws, "E" & idxStr, i_CLASSQUOTA_A, ExcelHorizontalAlignment.Right) '"可核配上限")
            SetCellValue(ws, "F" & idxStr, i_CLSCNT_A, ExcelHorizontalAlignment.Right) '"申請班次")
            SetCellValue(ws, "G" & idxStr, i_CLSCNTY_A, ExcelHorizontalAlignment.Right) '"核定班次")
            SetCellValue(ws, "H" & idxStr, i_STDCNT_A, ExcelHorizontalAlignment.Right) '"申請人次
            SetCellValue(ws, "I" & idxStr, i_STDCNTY_A, ExcelHorizontalAlignment.Right) '"核定人次
            SetCellValue(ws, "J" & idxStr, i_DEFGOVCOST_A, ExcelHorizontalAlignment.Right) '"申請補助費")
            SetCellValue(ws, "K" & idxStr, i_DEFGOVCOSTY_A, ExcelHorizontalAlignment.Right) '"核定補助費")
            v_CLSRATE1 = String.Concat(If(i_CLSCNT_A > 0, TIMS.ROUND(100.0 * (i_CLSCNTY_A / i_CLSCNT_A), 2), "0.00"), "%") '核班率
            SetCellValue(ws, "L" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
            ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

            'idxStr -= 1 '(畫線)
            Dim cellsCOLSPNumF2 As String = String.Concat("A9:", END_COL_NM, "{0}") '(畫格子使用)
            Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
                With exlRow3X
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s 'FontSize
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    .AutoFitColumns(25.0, 250.0)
                End With
                SetCellBorder(exlRow3X)
            End Using

            ' 設定貨幣格式，小數位數為 0
            ws.Cells(String.Format("E9:E{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("J9:J{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("K9:K{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("A9:A{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("B9:B{0}", idxStr)).Start.Column).Width = 33
            ws.Column(ws.Cells(String.Format("C9:C{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("D9:D{0}", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("E9:E{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("F9:F{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("G9:G{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("H9:H{0}", idxStr)).Start.Column).Width = 10

            ws.Column(ws.Cells(String.Format("I9:I{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("J9:J{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("K9:K{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("L9:L{0}", idxStr)).Start.Column).Width = 15

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90

            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub

    Sub EXPXLSX_1JG()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!")
            Exit Sub
        End If

        Dim dtXls2a As DataTable = Nothing 'C:跨區提案單位
        Dim dtXls2b As DataTable = Nothing 'J:轄區提案單位
        dtXls2a = ALL_DATA_TOTAL_dt2(dtXls2b)
        If TIMS.dtNODATA(dtXls2a) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If
        If TIMS.dtNODATA(dtXls2b) Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        Dim dr2a As DataRow = dtXls2a.Rows(0) 'C:跨區提案單位
        Dim dr2b As DataRow = dtXls2b.Rows(0) 'J:轄區提案單位'申請訓練單位數(含跨區及非跨區)

        Dim END_COL_NM As String = "L"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim strErrmsg As String = ""

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Dim V_CROSSDIST_NM1 As String = If(v_RBL_CrossDist_SCH = "C", "跨區", If(v_RBL_CrossDist_SCH = "J", "轄區", "不區分"))
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim V_APPSTAGE_NM1 As String = If(v_ddlAPPSTAGE_SCH = "1", "上", If(v_ddlAPPSTAGE_SCH = "2", "下", "X"))
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim V_ORGKIND2_NM1 As String = If(v_rblOrgKind2 = "G", "產投", If(v_rblOrgKind2 = "W", "自主", "")) '113下自主-北分署

        '"產投差異表(依明細)-113下(跨區)"
        Dim V_SHEETNM1 As String = String.Concat(V_ORGKIND2_NM1, "差異表(依明細)-", s_ROCYEAR1, V_APPSTAGE_NM1, "(", V_CROSSDIST_NM1, ")")
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "訓練課程申請/核定差異統計表")
        Dim s_TITLENAME2 As String = "轄區單位明細"
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "申請核定差異統計表_", TIMS.GetDateNo())

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            'Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage()

            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定 'Dim fontName As String = "標楷體" 'Dim fontSize12s As Single = 12.0F '報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "2"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 10
                    .Value = String.Concat(Now.Year - 1911, Now.ToString(".MM.dd"))
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    '.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Call EXPXLSX_3(ws, dr2b, 3, v_rblOrgKind2)
            Call EXPXLSX_3(ws, dr2a, 6, v_rblOrgKind2)

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "8"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME2
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
            Dim idxStr As Integer = 9
            SetCellValue(ws, "A" & idxStr, "序號") '10
            SetCellValue(ws, "B" & idxStr, "訓練單位") '33
            SetCellValue(ws, "C" & idxStr, "分署") '10
            SetCellValue(ws, "D" & idxStr, String.Concat("審查計分表", vbLf, "等級")) '15
            SetCellValue(ws, "E" & idxStr, "可核配上限") '20
            SetCellValue(ws, "F" & idxStr, String.Concat("申請", vbLf, "班次")) '10
            SetCellValue(ws, "G" & idxStr, String.Concat("核定", vbLf, "班次")) '10
            SetCellValue(ws, "H" & idxStr, String.Concat("申請", vbLf, "人次")) '10
            SetCellValue(ws, "I" & idxStr, "核定人次") '20
            SetCellValue(ws, "J" & idxStr, String.Concat("申請", vbLf, "補助費")) '20
            SetCellValue(ws, "K" & idxStr, String.Concat("核定", vbLf, "補助費")) '20
            SetCellValue(ws, "L" & idxStr, "核班率") '20
            ws.Cells(String.Format("A{0}:L{0}", 9)).Style.Font.Bold = True

            idxStr = 10
            Dim idx1 As Integer = idxStr
            Dim idx2b As Integer = 0
            Dim V_COMIDNO As String = dtXls.Rows(0)("COMIDNO")
            Dim iCNT As Integer = 0 '(-筆數跳號)
            '(總計)
            Dim i_CLASSQUOTA_A As Integer = 0 '"可核配上限")
            Dim i_CLSCNT_A As Integer = 0 '"申請班次")
            Dim i_CLSCNTY_A As Integer = 0 '"核定班次")
            Dim i_STDCNT_A As Integer = 0 '"申請人次
            Dim i_STDCNTY_A As Integer = 0 '"核定人次
            Dim i_DEFGOVCOST_A As Integer = 0 '"申請補助費")
            Dim i_DEFGOVCOSTY_A As Integer = 0 '"核定補助費")

            Dim v_CLSRATE1 As String = "" '"核班率")

            For Each dr1 As DataRow In dtXls.Rows
                iCNT += 1
                'iTHOURS += TIMS.VAL1(dr1("THOURS")) 'iTNUM += TIMS.VAL1(dr1("TNUM"))
                'iTOTAL += TIMS.VAL1(dr1("TOTAL"))'iTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))'iDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
                'If V_COMIDNO = dr1("COMIDNO") Then
                '    iCNT_G += 1
                i_CLASSQUOTA_A += TIMS.GetValue2(dr1("CLASSQUOTA"))
                i_CLSCNT_A += TIMS.GetValue2(dr1("CLSCNT"))
                i_CLSCNTY_A += TIMS.GetValue2(dr1("CLSCNTY"))
                i_STDCNT_A += TIMS.GetValue2(dr1("STDCNT"))
                i_STDCNTY_A += TIMS.GetValue2(dr1("STDCNTY"))
                i_DEFGOVCOST_A += TIMS.GetValue2(dr1("DEFGOVCOST"))
                i_DEFGOVCOSTY_A += TIMS.GetValue2(dr1("DEFGOVCOSTY"))

                '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
                'ROWSEQNO,ORGNAME,DISTNAME3,RLEVEL2,CLASSQUOTA,CLSCNT,CLSCNTY,STDCNT,STDCNTY,DEFGOVCOST,DEFGOVCOSTY,CLSRATE1
                SetCellValue(ws, "A" & idxStr, iCNT) '"訓練單位名稱")
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '"訓練單位")
                SetCellValue(ws, "C" & idxStr, dr1("DISTNAME3")) '"分署")
                SetCellValue(ws, "D" & idxStr, dr1("RLEVEL2")) '"審查計分表等級")
                SetCellValue(ws, "E" & idxStr, dr1("CLASSQUOTA"), ExcelHorizontalAlignment.Right) '"可核配上限")
                SetCellValue(ws, "F" & idxStr, dr1("CLSCNT"), ExcelHorizontalAlignment.Right) '"申請班次")
                SetCellValue(ws, "G" & idxStr, dr1("CLSCNTY"), ExcelHorizontalAlignment.Right) '"核定班次")
                SetCellValue(ws, "H" & idxStr, dr1("STDCNT"), ExcelHorizontalAlignment.Right) '"申請人次
                SetCellValue(ws, "I" & idxStr, dr1("STDCNTY"), ExcelHorizontalAlignment.Right) '"核定人次
                SetCellValue(ws, "J" & idxStr, dr1("DEFGOVCOST"), ExcelHorizontalAlignment.Right) '"申請補助費")
                SetCellValue(ws, "K" & idxStr, dr1("DEFGOVCOSTY"), ExcelHorizontalAlignment.Right) '"核定補助費")
                SetCellValue(ws, "L" & idxStr, dr1("CLSRATE1"), ExcelHorizontalAlignment.Right) '"核班率")

                idxStr += 1
            Next
            '(合併儲存格) idxStr += 1
            'idx2b = idx1 + iCNT_G - 1
            'ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
            'ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True

            ''(小計) '
            'ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
            'SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), "小計")
            'SetCellValue(ws, "E" & idxStr, i_CLASSQUOTA, ExcelHorizontalAlignment.Right) '"可核配上限")
            'SetCellValue(ws, "F" & idxStr, i_CLSCNT, ExcelHorizontalAlignment.Right) '"申請班次")
            'SetCellValue(ws, "G" & idxStr, i_CLSCNTY, ExcelHorizontalAlignment.Right) '"核定班次")
            'SetCellValue(ws, "H" & idxStr, i_STDCNT, ExcelHorizontalAlignment.Right) '"申請人次
            'SetCellValue(ws, "I" & idxStr, i_STDCNTY, ExcelHorizontalAlignment.Right) '"核定人次
            'SetCellValue(ws, "J" & idxStr, i_DEFGOVCOST, ExcelHorizontalAlignment.Right) '"申請補助費")
            'SetCellValue(ws, "K" & idxStr, i_DEFGOVCOSTY, ExcelHorizontalAlignment.Right) '"核定補助費")
            'v_CLSRATE1 = String.Concat(If(i_CLSCNT > 0, TIMS.ROUND(100.0 * (i_CLSCNTY / i_CLSCNT), 2), "0.00"), "%") '核班率
            'SetCellValue(ws, "L" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")

            '(總計) 'idxStr += 1
            ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
            SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), "總計")
            SetCellValue(ws, "E" & idxStr, i_CLASSQUOTA_A, ExcelHorizontalAlignment.Right) '"可核配上限")
            SetCellValue(ws, "F" & idxStr, i_CLSCNT_A, ExcelHorizontalAlignment.Right) '"申請班次")
            SetCellValue(ws, "G" & idxStr, i_CLSCNTY_A, ExcelHorizontalAlignment.Right) '"核定班次")
            SetCellValue(ws, "H" & idxStr, i_STDCNT_A, ExcelHorizontalAlignment.Right) '"申請人次
            SetCellValue(ws, "I" & idxStr, i_STDCNTY_A, ExcelHorizontalAlignment.Right) '"核定人次
            SetCellValue(ws, "J" & idxStr, i_DEFGOVCOST_A, ExcelHorizontalAlignment.Right) '"申請補助費")
            SetCellValue(ws, "K" & idxStr, i_DEFGOVCOSTY_A, ExcelHorizontalAlignment.Right) '"核定補助費")
            v_CLSRATE1 = String.Concat(If(i_CLSCNT_A > 0, TIMS.ROUND(100.0 * (i_CLSCNTY_A / i_CLSCNT_A), 2), "0.00"), "%") '核班率
            SetCellValue(ws, "L" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
            ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

            'idxStr -= 1 '(畫線)
            Dim cellsCOLSPNumF2 As String = String.Concat("A9:", END_COL_NM, "{0}") '(畫格子使用)
            Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
                With exlRow3X
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s 'FontSize
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    .AutoFitColumns(25.0, 250.0)
                End With
                SetCellBorder(exlRow3X)
            End Using

            ' 設定貨幣格式，小數位數為 0
            ws.Cells(String.Format("E9:E{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("J9:J{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("K9:K{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("A9:A{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("B9:B{0}", idxStr)).Start.Column).Width = 33
            ws.Column(ws.Cells(String.Format("C9:C{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("D9:D{0}", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("E9:E{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("F9:F{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("G9:G{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("H9:H{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("I9:I{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("J9:J{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("K9:K{0}", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("L9:L{0}", idxStr)).Start.Column).Width = 15

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90

            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
    End Sub

    ''' <summary> 'C:跨區提案單位 </summary>
    Sub EXPXLSX_1CW()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!")
            Exit Sub
        End If

        Dim dtXls2a As DataTable = Nothing 'C:跨區提案單位
        Dim dtXls2b As DataTable = Nothing 'J:轄區提案單位
        dtXls2a = ALL_DATA_TOTAL_dt2(dtXls2b)
        If TIMS.dtNODATA(dtXls2a) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If
        If TIMS.dtNODATA(dtXls2b) Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        Dim dr2a As DataRow = dtXls2a.Rows(0) 'C:跨區提案單位
        Dim dr2b As DataRow = dtXls2b.Rows(0) 'J:轄區提案單位'申請訓練單位數(含跨區及非跨區)

        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        Dim V_ORGKIND2_NM1 As String = If(v_rblOrgKind2 = "G", "產投", If(v_rblOrgKind2 = "W", "自主", "")) '113下自主-北分署
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫

        Dim END_COL_NM As String = "M"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim strErrmsg As String = ""

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Dim V_CROSSDIST_NM1 As String = If(v_RBL_CrossDist_SCH = "C", "跨區", If(v_RBL_CrossDist_SCH = "J", "轄區", "不區分"))
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim V_APPSTAGE_NM1 As String = If(v_ddlAPPSTAGE_SCH = "1", "上", If(v_ddlAPPSTAGE_SCH = "2", "下", "X"))
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度

        '"產投差異表(依明細)-113下(跨區)"
        Dim V_SHEETNM1 As String = String.Concat(V_ORGKIND2_NM1, "差異表(依明細)-", s_ROCYEAR1, V_APPSTAGE_NM1, "(", V_CROSSDIST_NM1, ")")
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "訓練課程申請/核定差異統計表")
        Dim s_TITLENAME2 As String = "跨區單位明細"
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "申請核定差異統計表_", TIMS.GetDateNo())

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            'Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage()

            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定 'Dim fontName As String = "標楷體" 'Dim fontSize12s As Single = 12.0F '報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "2"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 10
                    .Value = String.Concat(Now.Year - 1911, Now.ToString(".MM.dd"))
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    '.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Call EXPXLSX_3(ws, dr2b, 3, v_rblOrgKind2)
            Call EXPXLSX_3(ws, dr2a, 6, v_rblOrgKind2)

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "8"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME2
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
            Dim idxStr As Integer = 9
            SetCellValue(ws, "A" & idxStr, "序號") '10
            SetCellValue(ws, "B" & idxStr, "訓練單位") '33
            SetCellValue(ws, "C" & idxStr, "理事長") '10
            SetCellValue(ws, "D" & idxStr, "分署") '10
            SetCellValue(ws, "E" & idxStr, String.Concat("審查計分表", vbLf, "等級")) '15
            SetCellValue(ws, "F" & idxStr, "可核配上限") '20
            SetCellValue(ws, "G" & idxStr, String.Concat("申請", vbLf, "班次")) '10
            SetCellValue(ws, "H" & idxStr, String.Concat("核定", vbLf, "班次")) '10
            SetCellValue(ws, "I" & idxStr, String.Concat("申請", vbLf, "人次")) '10
            SetCellValue(ws, "J" & idxStr, "核定人次") '20
            SetCellValue(ws, "K" & idxStr, String.Concat("申請", vbLf, "補助費")) '20
            SetCellValue(ws, "L" & idxStr, String.Concat("核定", vbLf, "補助費")) '20
            SetCellValue(ws, "M" & idxStr, "核班率") '20
            ws.Cells(String.Format(cellsCOLSPNumF, 9)).Style.Font.Bold = True

            idxStr = 10
            Dim idx1 As Integer = idxStr
            Dim idx2b As Integer = 0
            Dim V_COMIDNO As String = dtXls.Rows(0)("COMIDNO")
            Dim iCNT_G As Integer = 0 '(-群組筆數跳號)
            Dim iCNTX As Integer = 0 '(-筆數跳號)
            Dim iCNT_O As Integer = 1 '(-機構數跳號)
            'Dim iTHOURS As Integer = 0
            'Dim iTNUM As Integer = 0
            'Dim iTOTAL As Integer = 0
            'Dim iTOTALCOST As Integer = 0
            'Dim iDEFGOVCOST As Integer = 0
            Dim i_CLASSQUOTA As Integer = 0 '"可核配上限")
            Dim i_CLSCNT As Integer = 0 '"申請班次")
            Dim i_CLSCNTY As Integer = 0 '"核定班次")
            Dim i_STDCNT As Integer = 0 '"申請人次
            Dim i_STDCNTY As Integer = 0 '"核定人次
            Dim i_DEFGOVCOST As Integer = 0 '"申請補助費")
            Dim i_DEFGOVCOSTY As Integer = 0 '"核定補助費")
            '(總計)
            Dim i_CLASSQUOTA_A As Integer = 0 '"可核配上限")
            Dim i_CLSCNT_A As Integer = 0 '"申請班次")
            Dim i_CLSCNTY_A As Integer = 0 '"核定班次")
            Dim i_STDCNT_A As Integer = 0 '"申請人次
            Dim i_STDCNTY_A As Integer = 0 '"核定人次
            Dim i_DEFGOVCOST_A As Integer = 0 '"申請補助費")
            Dim i_DEFGOVCOSTY_A As Integer = 0 '"核定補助費")

            Dim v_CLSRATE1 As String = "" '"核班率")

            For Each dr1 As DataRow In dtXls.Rows
                iCNTX += 1
                'iTHOURS += TIMS.VAL1(dr1("THOURS")) 'iTNUM += TIMS.VAL1(dr1("TNUM"))
                'iTOTAL += TIMS.VAL1(dr1("TOTAL"))'iTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))'iDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
                If V_COMIDNO = dr1("COMIDNO") Then
                    iCNT_G += 1
                    If TIMS.GetValue2(dr1("PARROWID")) <= 3 Then i_CLASSQUOTA += TIMS.GetValue2(dr1("CLASSQUOTA"))
                    i_CLSCNT += TIMS.GetValue2(dr1("CLSCNT"))
                    i_CLSCNTY += TIMS.GetValue2(dr1("CLSCNTY"))
                    i_STDCNT += TIMS.GetValue2(dr1("STDCNT"))
                    i_STDCNTY += TIMS.GetValue2(dr1("STDCNTY"))
                    i_DEFGOVCOST += TIMS.GetValue2(dr1("DEFGOVCOST"))
                    i_DEFGOVCOSTY += TIMS.GetValue2(dr1("DEFGOVCOSTY"))
                Else
                    '(合併儲存格)
                    idx2b = idx1 + iCNT_G - 1
                    ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
                    ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True
                    ws.Cells(String.Format("C{0}:C{1}", idx1, idx2b)).Merge = True
                    '(小計) '
                    v_CLSRATE1 = String.Concat(If(i_CLSCNT > 0, TIMS.ROUND(100.0 * (i_CLSCNTY / i_CLSCNT), 2), "0.00"), "%") '核班率
                    ws.Cells(String.Format("A{0}:E{0}", idxStr)).Merge = True
                    SetCellValue(ws, String.Format("A{0}:E{0}", idxStr), "小計")
                    SetCellValue(ws, "F" & idxStr, i_CLASSQUOTA, ExcelHorizontalAlignment.Right) '"可核配上限")
                    SetCellValue(ws, "G" & idxStr, i_CLSCNT, ExcelHorizontalAlignment.Right) '"申請班次")
                    SetCellValue(ws, "H" & idxStr, i_CLSCNTY, ExcelHorizontalAlignment.Right) '"核定班次")
                    SetCellValue(ws, "I" & idxStr, i_STDCNT, ExcelHorizontalAlignment.Right) '"申請人次
                    SetCellValue(ws, "J" & idxStr, i_STDCNTY, ExcelHorizontalAlignment.Right) '"核定人次
                    SetCellValue(ws, "K" & idxStr, i_DEFGOVCOST, ExcelHorizontalAlignment.Right) '"申請補助費")
                    SetCellValue(ws, "L" & idxStr, i_DEFGOVCOSTY, ExcelHorizontalAlignment.Right) '"核定補助費")
                    SetCellValue(ws, "M" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
                    ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

                    idxStr += 1 '(小計)'(總計)

                    i_CLASSQUOTA_A += i_CLASSQUOTA
                    i_CLSCNT_A += i_CLSCNT
                    i_CLSCNTY_A += i_CLSCNTY
                    i_STDCNT_A += i_STDCNT
                    i_STDCNTY_A += i_STDCNTY
                    i_DEFGOVCOST_A += i_DEFGOVCOST
                    i_DEFGOVCOSTY_A += i_DEFGOVCOSTY

                    idx1 = idxStr
                    V_COMIDNO = dr1("COMIDNO")
                    iCNT_G = 1
                    iCNT_O += 1
                    i_CLASSQUOTA = If(TIMS.GetValue2(dr1("PARROWID")) <= 3, TIMS.GetValue2(dr1("CLASSQUOTA")), 0)
                    'i_CLASSQUOTA = TIMS.GetValue2(dr1("CLASSQUOTA"))
                    i_CLSCNT = TIMS.GetValue2(dr1("CLSCNT"))
                    i_CLSCNTY = TIMS.GetValue2(dr1("CLSCNTY"))
                    i_STDCNT = TIMS.GetValue2(dr1("STDCNT"))
                    i_STDCNTY = TIMS.GetValue2(dr1("STDCNTY"))
                    i_DEFGOVCOST = TIMS.GetValue2(dr1("DEFGOVCOST"))
                    i_DEFGOVCOSTY = TIMS.GetValue2(dr1("DEFGOVCOSTY"))
                End If

                '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
                'ROWSEQNO,ORGNAME,DISTNAME3,RLEVEL2,CLASSQUOTA,CLSCNT,CLSCNTY,STDCNT,STDCNTY,DEFGOVCOST,DEFGOVCOSTY,CLSRATE1
                SetCellValue(ws, "A" & idxStr, iCNT_O) '"訓練單位名稱")
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '"訓練單位 
                SetCellValue(ws, "C" & idxStr, dr1("MASTERNAME")) '理事長
                SetCellValue(ws, "D" & idxStr, dr1("DISTNAME3")) '分署
                SetCellValue(ws, "E" & idxStr, dr1("RLEVEL2")) '"審查計分表等級
                SetCellValue(ws, "F" & idxStr, dr1("CLASSQUOTA"), ExcelHorizontalAlignment.Right) '"可核配上限
                SetCellValue(ws, "G" & idxStr, dr1("CLSCNT"), ExcelHorizontalAlignment.Right) '"申請班次
                SetCellValue(ws, "H" & idxStr, dr1("CLSCNTY"), ExcelHorizontalAlignment.Right) '"核定班次
                SetCellValue(ws, "I" & idxStr, dr1("STDCNT"), ExcelHorizontalAlignment.Right) '"申請人次
                SetCellValue(ws, "J" & idxStr, dr1("STDCNTY"), ExcelHorizontalAlignment.Right) '"核定人次
                SetCellValue(ws, "K" & idxStr, dr1("DEFGOVCOST"), ExcelHorizontalAlignment.Right) '"申請補助費
                SetCellValue(ws, "L" & idxStr, dr1("DEFGOVCOSTY"), ExcelHorizontalAlignment.Right) '"核定補助費
                SetCellValue(ws, "M" & idxStr, dr1("CLSRATE1"), ExcelHorizontalAlignment.Right) '"核班率

                idxStr += 1
            Next
            '(合併儲存格) idxStr += 1
            idx2b = idx1 + iCNT_G - 1
            ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
            ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True
            ws.Cells(String.Format("C{0}:C{1}", idx1, idx2b)).Merge = True
            '(小計) '
            v_CLSRATE1 = String.Concat(If(i_CLSCNT > 0, TIMS.ROUND(100.0 * (i_CLSCNTY / i_CLSCNT), 2), "0.00"), "%") '核班率
            ws.Cells(String.Format("A{0}:E{0}", idxStr)).Merge = True
            SetCellValue(ws, String.Format("A{0}:E{0}", idxStr), "小計")
            SetCellValue(ws, "F" & idxStr, i_CLASSQUOTA, ExcelHorizontalAlignment.Right) '"可核配上限")
            SetCellValue(ws, "G" & idxStr, i_CLSCNT, ExcelHorizontalAlignment.Right) '"申請班次")
            SetCellValue(ws, "H" & idxStr, i_CLSCNTY, ExcelHorizontalAlignment.Right) '"核定班次")
            SetCellValue(ws, "I" & idxStr, i_STDCNT, ExcelHorizontalAlignment.Right) '"申請人次
            SetCellValue(ws, "J" & idxStr, i_STDCNTY, ExcelHorizontalAlignment.Right) '"核定人次
            SetCellValue(ws, "K" & idxStr, i_DEFGOVCOST, ExcelHorizontalAlignment.Right) '"申請補助費")
            SetCellValue(ws, "L" & idxStr, i_DEFGOVCOSTY, ExcelHorizontalAlignment.Right) '"核定補助費")
            SetCellValue(ws, "M" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
            ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True
            '(總計)
            i_CLASSQUOTA_A += i_CLASSQUOTA
            i_CLSCNT_A += i_CLSCNT
            i_CLSCNTY_A += i_CLSCNTY
            i_STDCNT_A += i_STDCNT
            i_STDCNTY_A += i_STDCNTY
            i_DEFGOVCOST_A += i_DEFGOVCOST
            i_DEFGOVCOSTY_A += i_DEFGOVCOSTY
            '(總計)
            idxStr += 1
            v_CLSRATE1 = String.Concat(If(i_CLSCNT_A > 0, TIMS.ROUND(100.0 * (i_CLSCNTY_A / i_CLSCNT_A), 2), "0.00"), "%") '核班率
            ws.Cells(String.Format("A{0}:E{0}", idxStr)).Merge = True
            SetCellValue(ws, String.Format("A{0}:E{0}", idxStr), "總計")
            SetCellValue(ws, "F" & idxStr, i_CLASSQUOTA_A, ExcelHorizontalAlignment.Right) '"可核配上限")
            SetCellValue(ws, "G" & idxStr, i_CLSCNT_A, ExcelHorizontalAlignment.Right) '"申請班次")
            SetCellValue(ws, "H" & idxStr, i_CLSCNTY_A, ExcelHorizontalAlignment.Right) '"核定班次")
            SetCellValue(ws, "I" & idxStr, i_STDCNT_A, ExcelHorizontalAlignment.Right) '"申請人次
            SetCellValue(ws, "J" & idxStr, i_STDCNTY_A, ExcelHorizontalAlignment.Right) '"核定人次
            SetCellValue(ws, "K" & idxStr, i_DEFGOVCOST_A, ExcelHorizontalAlignment.Right) '"申請補助費")
            SetCellValue(ws, "L" & idxStr, i_DEFGOVCOSTY_A, ExcelHorizontalAlignment.Right) '"核定補助費")
            SetCellValue(ws, "M" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
            ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

            'idxStr -= 1 '(畫線)
            Dim cellsCOLSPNumF2 As String = String.Concat("A9:", END_COL_NM, "{0}") '(畫格子使用)
            Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
                With exlRow3X
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s 'FontSize
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    .AutoFitColumns(25.0, 250.0)
                End With
                SetCellBorder(exlRow3X)
            End Using

            ' 設定貨幣格式，小數位數為 0
            ws.Cells(String.Format("{0}9:{0}{1}", "F", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("{0}9:{0}{1}", "K", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("{0}9:{0}{1}", "L", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "A", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "B", idxStr)).Start.Column).Width = 44
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "C", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "D", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "E", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "F", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "G", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "H", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "I", idxStr)).Start.Column).Width = 10

            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "J", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "K", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "L", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "M", idxStr)).Start.Column).Width = 15

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90

            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub

    Sub EXPXLSX_1JW()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!")
            Exit Sub
        End If

        Dim dtXls2a As DataTable = Nothing 'C:跨區提案單位
        Dim dtXls2b As DataTable = Nothing 'J:轄區提案單位
        dtXls2a = ALL_DATA_TOTAL_dt2(dtXls2b)
        If TIMS.dtNODATA(dtXls2a) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If
        If TIMS.dtNODATA(dtXls2b) Then
            Common.MessageBox(Me, "查無匯出資料!!!")
            Exit Sub
        End If
        Dim dr2a As DataRow = dtXls2a.Rows(0) 'C:跨區提案單位
        Dim dr2b As DataRow = dtXls2b.Rows(0) 'J:轄區提案單位'申請訓練單位數(含跨區及非跨區)

        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim V_ORGKIND2_NM1 As String = If(v_rblOrgKind2 = "G", "產投", If(v_rblOrgKind2 = "W", "自主", "")) '113下自主-北分署

        Dim END_COL_NM As String = "M"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim strErrmsg As String = ""

        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Dim V_CROSSDIST_NM1 As String = If(v_RBL_CrossDist_SCH = "C", "跨區", If(v_RBL_CrossDist_SCH = "J", "轄區", "不區分"))
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim V_APPSTAGE_NM1 As String = If(v_ddlAPPSTAGE_SCH = "1", "上", If(v_ddlAPPSTAGE_SCH = "2", "下", "X"))
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度

        '"產投差異表(依明細)-113下(跨區)"
        Dim V_SHEETNM1 As String = String.Concat(V_ORGKIND2_NM1, "差異表(依明細)-", s_ROCYEAR1, V_APPSTAGE_NM1, "(", V_CROSSDIST_NM1, ")")
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "訓練課程申請/核定差異統計表")
        Dim s_TITLENAME2 As String = "轄區單位明細"
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "申請核定差異統計表_", TIMS.GetDateNo())

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            'Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage()

            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定 'Dim fontName As String = "標楷體" 'Dim fontSize12s As Single = 12.0F '報表標題
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "2"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 10
                    .Value = String.Concat(Now.Year - 1911, Now.ToString(".MM.dd"))
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Right
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    '.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            Call EXPXLSX_3(ws, dr2b, 3, v_rblOrgKind2) 'J:轄區提案單位'申請訓練單位數(含跨區及非跨區)
            Call EXPXLSX_3(ws, dr2a, 6, v_rblOrgKind2) 'C:跨區提案單位

            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "8"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Bold = True
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME2
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
            Dim idxStr As Integer = 9
            SetCellValue(ws, "A" & idxStr, "序號") '10
            SetCellValue(ws, "B" & idxStr, "訓練單位") '33
            SetCellValue(ws, "C" & idxStr, "理事長") '10'MASTERNAME
            SetCellValue(ws, "D" & idxStr, "分署") '10
            SetCellValue(ws, "E" & idxStr, String.Concat("審查計分表", vbLf, "等級")) '15
            SetCellValue(ws, "F" & idxStr, "可核配上限") '20
            SetCellValue(ws, "G" & idxStr, String.Concat("申請", vbLf, "班次")) '10
            SetCellValue(ws, "H" & idxStr, String.Concat("核定", vbLf, "班次")) '10
            SetCellValue(ws, "I" & idxStr, String.Concat("申請", vbLf, "人次")) '10
            SetCellValue(ws, "J" & idxStr, "核定人次") '20
            SetCellValue(ws, "K" & idxStr, String.Concat("申請", vbLf, "補助費")) '20
            SetCellValue(ws, "L" & idxStr, String.Concat("核定", vbLf, "補助費")) '20
            SetCellValue(ws, "M" & idxStr, "核班率") '20
            ws.Cells(String.Format(cellsCOLSPNumF, 9)).Style.Font.Bold = True

            idxStr = 10
            Dim idx1 As Integer = idxStr
            Dim idx2b As Integer = 0
            Dim V_COMIDNO As String = dtXls.Rows(0)("COMIDNO")
            Dim iCNT As Integer = 0 '(-筆數跳號)
            '(總計)
            Dim i_CLASSQUOTA_A As Integer = 0 '"可核配上限")
            Dim i_CLSCNT_A As Integer = 0 '"申請班次")
            Dim i_CLSCNTY_A As Integer = 0 '"核定班次")
            Dim i_STDCNT_A As Integer = 0 '"申請人次
            Dim i_STDCNTY_A As Integer = 0 '"核定人次
            Dim i_DEFGOVCOST_A As Integer = 0 '"申請補助費")
            Dim i_DEFGOVCOSTY_A As Integer = 0 '"核定補助費")

            Dim v_CLSRATE1 As String = "" '"核班率")

            For Each dr1 As DataRow In dtXls.Rows
                iCNT += 1
                'iTHOURS += TIMS.VAL1(dr1("THOURS")) 'iTNUM += TIMS.VAL1(dr1("TNUM"))
                'iTOTAL += TIMS.VAL1(dr1("TOTAL"))'iTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))'iDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
                'If V_COMIDNO = dr1("COMIDNO") Then
                '    iCNT_G += 1
                i_CLASSQUOTA_A += TIMS.GetValue2(dr1("CLASSQUOTA"))
                i_CLSCNT_A += TIMS.GetValue2(dr1("CLSCNT"))
                i_CLSCNTY_A += TIMS.GetValue2(dr1("CLSCNTY"))
                i_STDCNT_A += TIMS.GetValue2(dr1("STDCNT"))
                i_STDCNTY_A += TIMS.GetValue2(dr1("STDCNTY"))
                i_DEFGOVCOST_A += TIMS.GetValue2(dr1("DEFGOVCOST"))
                i_DEFGOVCOSTY_A += TIMS.GetValue2(dr1("DEFGOVCOSTY"))

                '序號,訓練單位,分署,審查計分表等級,可核配上限,申請班次,核定班次,申請人次,核定人次,申請補助費,核定補助費,核班率
                'ROWSEQNO,ORGNAME,DISTNAME3,RLEVEL2,CLASSQUOTA,CLSCNT,CLSCNTY,STDCNT,STDCNTY,DEFGOVCOST,DEFGOVCOSTY,CLSRATE1
                SetCellValue(ws, "A" & idxStr, iCNT) '"訓練單位名稱")
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '"訓練單位
                SetCellValue(ws, "C" & idxStr, dr1("MASTERNAME")) '理事長
                SetCellValue(ws, "D" & idxStr, dr1("DISTNAME3")) '分署
                SetCellValue(ws, "E" & idxStr, dr1("RLEVEL2")) '"審查計分表等級
                SetCellValue(ws, "F" & idxStr, dr1("CLASSQUOTA"), ExcelHorizontalAlignment.Right) '"可核配上限
                SetCellValue(ws, "G" & idxStr, dr1("CLSCNT"), ExcelHorizontalAlignment.Right) '"申請班次
                SetCellValue(ws, "H" & idxStr, dr1("CLSCNTY"), ExcelHorizontalAlignment.Right) '"核定班次
                SetCellValue(ws, "I" & idxStr, dr1("STDCNT"), ExcelHorizontalAlignment.Right) '"申請人次
                SetCellValue(ws, "J" & idxStr, dr1("STDCNTY"), ExcelHorizontalAlignment.Right) '"核定人次
                SetCellValue(ws, "K" & idxStr, dr1("DEFGOVCOST"), ExcelHorizontalAlignment.Right) '"申請補助費
                SetCellValue(ws, "L" & idxStr, dr1("DEFGOVCOSTY"), ExcelHorizontalAlignment.Right) '"核定補助費
                SetCellValue(ws, "M" & idxStr, dr1("CLSRATE1"), ExcelHorizontalAlignment.Right) '"核班率")

                idxStr += 1
            Next
            '(合併儲存格) idxStr += 1
            'idx2b = idx1 + iCNT_G - 1
            'ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
            'ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True

            ''(小計) '
            'ws.Cells(String.Format("A{0}:D{0}", idxStr)).Merge = True
            'SetCellValue(ws, String.Format("A{0}:D{0}", idxStr), "小計")
            'SetCellValue(ws, "E" & idxStr, i_CLASSQUOTA, ExcelHorizontalAlignment.Right) '"可核配上限")
            'SetCellValue(ws, "F" & idxStr, i_CLSCNT, ExcelHorizontalAlignment.Right) '"申請班次")
            'SetCellValue(ws, "G" & idxStr, i_CLSCNTY, ExcelHorizontalAlignment.Right) '"核定班次")
            'SetCellValue(ws, "H" & idxStr, i_STDCNT, ExcelHorizontalAlignment.Right) '"申請人次
            'SetCellValue(ws, "I" & idxStr, i_STDCNTY, ExcelHorizontalAlignment.Right) '"核定人次
            'SetCellValue(ws, "J" & idxStr, i_DEFGOVCOST, ExcelHorizontalAlignment.Right) '"申請補助費")
            'SetCellValue(ws, "K" & idxStr, i_DEFGOVCOSTY, ExcelHorizontalAlignment.Right) '"核定補助費")
            'v_CLSRATE1 = String.Concat(If(i_CLSCNT > 0, TIMS.ROUND(100.0 * (i_CLSCNTY / i_CLSCNT), 2), "0.00"), "%") '核班率
            'SetCellValue(ws, "L" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")

            '(總計) 'idxStr += 1
            v_CLSRATE1 = String.Concat(If(i_CLSCNT_A > 0, TIMS.ROUND(100.0 * (i_CLSCNTY_A / i_CLSCNT_A), 2), "0.00"), "%") '核班率
            ws.Cells(String.Format("A{0}:E{0}", idxStr)).Merge = True
            SetCellValue(ws, String.Format("A{0}:E{0}", idxStr), "總計")
            SetCellValue(ws, "F" & idxStr, i_CLASSQUOTA_A, ExcelHorizontalAlignment.Right) '"可核配上限")
            SetCellValue(ws, "G" & idxStr, i_CLSCNT_A, ExcelHorizontalAlignment.Right) '"申請班次")
            SetCellValue(ws, "H" & idxStr, i_CLSCNTY_A, ExcelHorizontalAlignment.Right) '"核定班次")
            SetCellValue(ws, "I" & idxStr, i_STDCNT_A, ExcelHorizontalAlignment.Right) '"申請人次
            SetCellValue(ws, "J" & idxStr, i_STDCNTY_A, ExcelHorizontalAlignment.Right) '"核定人次
            SetCellValue(ws, "K" & idxStr, i_DEFGOVCOST_A, ExcelHorizontalAlignment.Right) '"申請補助費")
            SetCellValue(ws, "L" & idxStr, i_DEFGOVCOSTY_A, ExcelHorizontalAlignment.Right) '"核定補助費")
            SetCellValue(ws, "M" & idxStr, v_CLSRATE1, ExcelHorizontalAlignment.Right) '"核班率")
            ws.Cells(String.Format(cellsCOLSPNumF, idxStr)).Style.Font.Bold = True

            'idxStr -= 1 '(畫線)
            Dim cellsCOLSPNumF2 As String = String.Concat("A9:", END_COL_NM, "{0}") '(畫格子使用)
            Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
                With exlRow3X
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s 'FontSize
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    .AutoFitColumns(25.0, 250.0)
                End With
                SetCellBorder(exlRow3X)
            End Using

            ' 設定貨幣格式，小數位數為 0
            ws.Cells(String.Format("{0}9:{0}{1}", "F", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("{0}9:{0}{1}", "K", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("{0}9:{0}{1}", "L", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "A", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "B", idxStr)).Start.Column).Width = 44
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "C", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "D", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "E", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "F", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "G", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "H", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "I", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "J", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "K", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "L", idxStr)).Start.Column).Width = 20
            ws.Column(ws.Cells(String.Format("{0}9:{0}{1}", "M", idxStr)).Start.Column).Width = 15

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90

            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case Else
                    Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                    Common.MessageBox(Me, s_log1)
                    Return ' Exit Sub
            End Select
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
    End Sub

    ''' <summary>總合計算調整 {{"SUM", "Y"}}</summary>
    ''' <param name="dt2"></param>
    ''' <returns></returns>
    Private Function ALL_DATA_TOTAL_dt2(ByRef dt2 As DataTable) As DataTable
        Dim parms1 As New Hashtable From {{"SUM", "Y"}} ' 總合計算調整 {{"SUM", "Y"}}
        Dim sSql1 As String = GET_WX_SQL1(parms1) 'C:跨區提案單位
        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms1), vbCrLf, "--##CR_03_001:", vbCrLf, sSql1))
        Dim dt As DataTable = DbAccess.GetDataTable(sSql1, objconn, parms1)

        Dim parms2 As New Hashtable From {{"SUM", "Y"}} ' 總合計算調整 {{"SUM", "Y"}}
        Dim sSql2 As String = GET_WX_SQL2(parms2) 'J:轄區提案單位
        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms2), vbCrLf, "--##CR_03_001:", vbCrLf, sSql2))
        dt2 = DbAccess.GetDataTable(sSql2, objconn, parms2)

        If TIMS.dtNODATA(dt) OrElse TIMS.dtNODATA(dt2) Then Return dt
        Dim dr1 As DataRow = dt.Rows(0)
        Dim dr2 As DataRow = dt2.Rows(0)

        Dim sCODL1 As String = "ORGCNT,ORGCNTY,CLSCNT,CLSCNTY,STDCNT,STDCNTY,DEFGOVCOST,DEFGOVCOSTY"
        Dim sCODL1S As String() = sCODL1.Split(",")
        For Each COL1 As String In sCODL1S
            dr2(COL1) = TIMS.VAL1(dr1(COL1)) + TIMS.VAL1(dr2(COL1))
        Next
        If Not IsDBNull(dr2("CLSCNT")) AndAlso Val(dr2("CLSCNT")) > 0 Then
            ',concat(FORMAT(100.0*SUM(CLSCNTY)/SUM(CLSCNT),'0.00'),'%') CLSRATE
            dr2("CLSRATE2") = String.Concat(TIMS.ROUND(100.0 * Val(dr2("CLSCNTY")) / Val(dr2("CLSCNT")), 2), "%")
        Else
            dr2("CLSRATE2") = "-"
        End If
        Return dt
    End Function

    '匯出  '表單04_申請核定差異統計表.xls
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        'Call EXPORT_4()
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                If v_rblOrgKind2 = "G" Then
                    Call EXPXLSX_1CG()
                Else
                    Call EXPXLSX_1CW()
                End If
            Case "J" 'J:轄區提案單位
                If v_rblOrgKind2 = "G" Then
                    Call EXPXLSX_1JG()
                Else
                    Call EXPXLSX_1JW()
                End If
        End Select
    End Sub

    ''' <summary>設定 Cell 儲存格值</summary>
    ''' <param name="sheet">Excel 工作表</param>
    ''' <param name="cellAddress">Cell 儲存格位址 (例如 A4、A1:L5)</param>
    ''' <param name="V_OBJ">Cell 儲存格值</param>
    ''' <param name="alignH">水平對齊方式</param>
    ''' <param name="alignV">垂直對齊方式</param>
    Private Sub SetCellValue(ByVal sheet As ExcelWorksheet, ByVal cellAddress As String, ByVal V_OBJ As Object, Optional ByVal alignH As ExcelHorizontalAlignment = ExcelHorizontalAlignment.Center, Optional ByVal alignV As ExcelVerticalAlignment = ExcelVerticalAlignment.Center)
        If sheet Is Nothing OrElse V_OBJ Is Nothing OrElse IsDBNull(V_OBJ) Then Return
        Dim nCells As ExcelRange = sheet.Cells(cellAddress)
        If nCells.Merge AndAlso cellAddress.IndexOf(":") > -1 Then
            sheet.Cells(cellAddress.Split(":")(0)).Value = V_OBJ
        Else
            nCells.Value = V_OBJ
        End If
        nCells.Style.HorizontalAlignment = alignH
        nCells.Style.VerticalAlignment = alignV
        nCells.Style.Font.Name = fontName
        nCells.Style.Font.Size = fontSize12s
        ' 設定自動換行
        nCells.Style.WrapText = True
        ' 設定欄寬為 40 (單位是字元寬度)
        'sheet.Column(nCells.Start.Column).Width = 40
        ' 自動調整列高以適應內容 (在設定值和自動換行後執行)
        nCells.AutoFitColumns(30, 60) ' 注意這裡用的是 AutoFitColumns，它會根據內容調整欄寬，但我們已經設定了固定的欄寬
        'nCells.AutoFitRows()    ' 這個方法會根據儲存格內容和自動換行調整列高

        ' 取得目前儲存格的欄索引 Dim columnIndex As Integer = nCells.Start.Column
        ' 自動調整該欄的寬度以適應內容 sheet.Column(columnIndex).AutoFitColumns()
        'nCells.AutoFitColumns(10, 1000)

        ' 設定框線樣式
        ' With nCells.Style.Border' .Left.Style = ExcelBorderStyle.Thin ' = BorderStyle
        '.Right.Style = ExcelBorderStyle.Thin 'BorderStyle' .Top.Style = ExcelBorderStyle.Thin 'BorderStyle
        '.Bottom.Style = ExcelBorderStyle.Thin ' BorderStyle' ' 設定框線顏色 (只有在指定顏色時才設定)
        ''If borderColor <> Color.Empty AndAlso borderColor IsNot Nothing Then'    .Left.Color.SetColor(borderColor)
        '.Right.Color.SetColor(borderColor)'    .Top.Color.SetColor(borderColor)'    .Bottom.Color.SetColor(borderColor)'End If
        'End With
    End Sub

    Private Sub SetCellBorder(ByVal exlRow As ExcelRange, Optional ByVal borderStyle As ExcelBorderStyle = ExcelBorderStyle.Thin)
        If exlRow Is Nothing Then Return
        For Each nERB As ExcelRangeBase In exlRow
            nERB.Style.Border.BorderAround(borderStyle)
        Next
    End Sub

End Class

