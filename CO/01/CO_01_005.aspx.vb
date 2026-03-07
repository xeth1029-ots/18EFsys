Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CO_01_005
    Inherits AuthBasePage 'System.Web.UI.Page

    '排程[Co_OrgScoring] Co_OrgScoring.exe.config  '排程[-xx-CO_ORGSCORING] /'ORG_SCORING2 'Lab_SUSPENDED_msg1
    Const cst_SUSPENDED_msgFM1 As String = "此單位因有{0}班停班經認列屬「不可抗力因素」，將不列入核定總班數計算。"
    ' 共用設定 
    Dim fontName As String = "標楷體"
    Dim fontSize12s As Single = 12.0F
    Dim fontSize14s As Single = 14.0F
    Dim fontSize16s As Single = 16.0F
    Dim print_lock As New Object '(); //lock

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1

        '匯出等級比率統計表(署用)
        tr_rblSCORESTAGE.Visible = (sm.UserInfo.LID = 0)
        BtnExp2.Visible = (sm.UserInfo.LID = 0)
        BtnExp3.Visible = (sm.UserInfo.LID = 0)
        BtnExp4.Visible = (sm.UserInfo.LID = 0)

        'autorecsubtotal

        '產投/非產投判斷
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    '產投
        '    Me.LabTMID.Text="訓練業別"
        'End If

        If Not IsPostBack Then
            CCreate1()
        End If

        If sm.UserInfo.DistID <> "000" Then
            Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
            ddlDISTID.Enabled = False
        End If

        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        'If HistoryRID.Rows.Count <> 0 Then
        '    'center.Attributes("onclick")="showObj('HistoryList2');ShowFrame();"
        '    'HistoryRID.Attributes("onclick")="ShowFrame();"
        '    center.Attributes("onclick")="showObj('HistoryList2');"
        '    center.Style("CURSOR")="hand"
        'End If
    End Sub

    Sub KeepSearch1()
        '(加強操作便利性)
        Dim s_kpSearch1 As String = ""
        TIMS.SetMyValue(s_kpSearch1, "ddlDISTID", TIMS.GetListValue(ddlDISTID))
        TIMS.SetMyValue(s_kpSearch1, "ddlSCORING", TIMS.GetListValue(ddlSCORING))
        TIMS.SetMyValue(s_kpSearch1, "OrgName", OrgName.Text)
        TIMS.SetMyValue(s_kpSearch1, "COMIDNO", COMIDNO.Text)
        TIMS.SetMyValue(s_kpSearch1, "OrgPlanKind", TIMS.GetListValue(OrgPlanKind))
        TIMS.SetMyValue(s_kpSearch1, "OrgKindList", TIMS.GetListValue(OrgKindList))
        Session("CO_01_004_Search1") = s_kpSearch1
    End Sub

    Sub UseKeepSearch1()
        '(加強操作便利性)
        If Session("CO_01_004_Search1") Is Nothing Then Return
        Dim s_kpSearch1 As String = Session("CO_01_004_Search1")
        If s_kpSearch1 = "" Then Return
        Session("CO_01_004_Search1") = Nothing
        Common.SetListItem(ddlDISTID, TIMS.GetMyValue(s_kpSearch1, "ddlDISTID"))
        Common.SetListItem(ddlSCORING, TIMS.GetMyValue(s_kpSearch1, "ddlSCORING"))
        OrgName.Text = TIMS.GetMyValue(s_kpSearch1, "OrgName")
        COMIDNO.Text = TIMS.GetMyValue(s_kpSearch1, "COMIDNO")
        Common.SetListItem(OrgPlanKind, TIMS.GetMyValue(s_kpSearch1, "OrgPlanKind"))
        Common.SetListItem(OrgKindList, TIMS.GetMyValue(s_kpSearch1, "OrgKindList"))
        Call SSearch1()
    End Sub

    Sub CCreate1()
        '(加強操作便利性)
        BtnSaveData1.Visible = False
        'changeMINISTERADD
        Dim js42 As String = "javascript:changeMINISUB(2);"
        MINISTERADD.Attributes.Add("onchange", js42)
        MINISTERADD.Attributes.Add("onblur", js42)
        MINISTERADD.Attributes.Add("onclick", js42)
        Dim js43 As String = "javascript:changeMINISUB(3);"
        DEPTADD.Attributes.Add("onchange", js43)
        DEPTADD.Attributes.Add("onblur", js43)
        DEPTADD.Attributes.Add("onclick", js43)

        'autorecsubtotal
        'Dim js_auto1 As String="autorecsubtotal();"
        'SCORE4_1.Attributes("onclick")=js_auto1 '"javascript:autorecsubtotal();"
        'SCORE4_1.Attributes("onblur")=js_auto1 '"javascript:autorecsubtotal();"
        'SCORE4_1.Attributes("onchange")=js_auto1 '"javascript:autorecsubtotal();"

        'SUBTOTAL.Attributes("onclick")=js_auto1 '"javascript:autorecsubtotal();"
        'SUBTOTAL.Attributes("onblur")=js_auto1 '"javascript:autorecsubtotal();"
        'SUBTOTAL.Attributes("onchange")=js_auto1 '"javascript:autorecsubtotal();"

        'SCORE4_2_RATE.Attributes("onclick")=js_auto1 '"javascript:autorecsubtotal();"
        'SCORE4_2_RATE.Attributes("onblur")=js_auto1 '"javascript:autorecsubtotal();"
        'SCORE4_2_RATE.Attributes("onchange")=js_auto1 '"javascript:autorecsubtotal();"

        divSch1.Visible = True
        divEdt1.Visible = False
        msg1.Text = ""
        PageControler1.Visible = False
        '評核版本 'ddlSENDVER=Get_SENDVER_TS(ddlSENDVER)
        '評核結果 'ddlRESULT=Get_RESULT_TS(ddlRESULT)
        '初擬等級-審核等級
        ddlIMPLEVEL_1 = TIMS.Get_SCORELEVEL(ddlIMPLEVEL_1)
        '部加分等級
        ddl_MINISTERLEVEL = TIMS.Get_SCORELEVEL(ddl_MINISTERLEVEL)
        '複審等級
        ddlRLEVEL_2R = TIMS.Get_SCORELEVEL(ddlRLEVEL_2R)

        ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        If (ddlDISTID.Items.FindByValue("000") IsNot Nothing) Then ddlDISTID.Items.Remove(ddlDISTID.Items.FindByValue("000"))
        Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)

        ddlSCORING = TIMS.Get_ddlSCORING(ddlSCORING, objconn)
        'SYEARlist=TIMS.GetSyear(SYEARlist)
        'Common.SetListItem(SYEARlist, sm.UserInfo.Years)

        'SearchPlan=TIMS.Get_RblOrgKind2(Me, SearchPlan) ', objconn
        OrgPlanKind = TIMS.Get_RblOrgPlanKind(OrgPlanKind, objconn)
        Common.SetListItem(OrgPlanKind, "G")

        OrgKindList = TIMS.Get_OrgType(OrgKindList, objconn)

        '依登入者機構判斷計畫種類
        '依登入者 LID 判斷是否可自由輸入
        If sm.UserInfo.LID = 2 Then '委訓單位動作
            Dim droo As DataRow = TIMS.Get_ORGINFOdr(sm.UserInfo.OrgID, objconn)
            If droo Is Nothing Then Throw New Exception("機構資料取得異常失敗")
            OrgName.Text = droo("OrgName")
            COMIDNO.Text = droo("ComIDNO")
            Select Case Convert.ToString(droo("OrgKind2"))
                Case "G", "W"
                    Common.SetListItem(OrgPlanKind, Convert.ToString(droo("OrgKind2")))
            End Select
            OrgName.Enabled = False
            COMIDNO.Enabled = False
            OrgPlanKind.Enabled = False
        End If

        '登入年度轉民國年份
        'Years.Value=sm.UserInfo.Years - 1911
        If TIMS.sUtl_ChkTest() Then
            Common.SetListItem(ddlDISTID, "001")
            'Common.SetListItem(ddlSCORING, "2025-01-2024-1-2024-2")
            '114年07月(113年下半年~114年上半年)	2025-07-2024-2-2025-1	114	下半年	115年度上半年	115上	2025	07
            Common.SetListItem(ddlSCORING, "2025-07-2024-2-2025-1")
        End If

        '選擇清除工作
        'SelectValue.Value=""
        DataGridTable.Visible = False
        Call UseKeepSearch1()
        'RIDValue.Value=sm.UserInfo.RID 'center.Text=sm.UserInfo.OrgName
    End Sub

    Protected Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = ""
        BtnSaveData1.Visible = False

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_IMP As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        Dim fg_SECONDCHK As Boolean = TIMS.CHK_ORG_TTQSLOCK_SECONDCHK(objconn, vSCORINGID) '有初審資料1年內
        If fg_SECONDCHK Then
            Call UPDATE_ORG_SCORING2_FIRSTCHK_Y()
            'ElseIf Not fg_SECONDCHK Then 'Common.MessageBox(Me, $"審查計分區間-初審或複審階段!(查無資料)請再確認查詢參數!") Return
        ElseIf fg_CAN_IMP Then
            Dim Msg_CAN_IMP As String = "請注意：審查計分表目前尚在開放分署確認之「初審」階段 !"
            Common.MessageBox(Me, Msg_CAN_IMP) 'Return
        End If

        'Call sClearlist1()
        Call SSearch1()
    End Sub

    ''' <summary>查詢多筆</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function GET_SQL_DT1(ByRef parms As Hashtable) As String
        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")
        Dim vORGNAME As String = TIMS.GetMyValue2(parms, "ORGNAME")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(parms, "COMIDNO")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")
        Dim vORGKIND As String = TIMS.GetMyValue2(parms, "ORGKIND")

        Dim sql As String = ""
        sql &= " SELECT a.OSID2" & vbCrLf '/*PK*/
        sql &= " ,a.OrgID,a.TPLANID,a.DISTID" & vbCrLf
        'SCORING_N 審查計分區間
        sql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS) ,'年',a.MONTHS,'月'" & vbCrLf
        sql &= "   ,'(',dbo.FN_CYEAR2(a.YEARS1) ,'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end ,'~'" & vbCrLf
        sql &= "   ,dbo.FN_CYEAR2(a.YEARS2) ,'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end ,')') SCORING_N" & vbCrLf
        sql &= " ,CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,a.MONTHS" & vbCrLf
        '【實際開班數】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
        '【政策型課程班】：顯示政策性課程全部的班數 (政策性(含停辦))
        sql &= " ,a.CLSACTCNT" & vbCrLf '實際開班數 (核定-停辦-政策性)(核定課程數-停辦課程數)
        sql &= " ,a.CLSACTCNT2" & vbCrLf '政策型課程班  (政策性(含停辦))

        sql &= " ,a.CLSAPPCNT" & vbCrLf
        sql &= " ,a.SCORE1_1A" & vbCrLf
        sql &= " ,a.SCORE1_1" & vbCrLf
        '【實際開訓人次】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
        '【政策性課程核定人次】：顯示政策性課程全部班級的核定人次 (政策性)
        sql &= " ,a.STDACTCNT" & vbCrLf '實際開訓人次
        sql &= " ,a.STDACTCNT2" & vbCrLf '政策性課程核定人次

        sql &= " ,a.STDAPPCNT" & vbCrLf
        sql &= " ,a.SCORE1_2A" & vbCrLf
        sql &= " ,a.SCORE1_2" & vbCrLf
        'sql &= " ,a.SCORE2_1_1A" & vbCrLf
        'sql &= " ,a.SCORE2_1_1B" & vbCrLf
        'sql &= " ,a.SCORE2_1_1C" & vbCrLf
        'sql &= " ,a.SCORE2_1_1D" & vbCrLf
        sql &= " ,a.SCORE2_1_1_ALL" & vbCrLf

        sql &= " ,a.SCORE2_1_2A_DIS" & vbCrLf
        sql &= " ,a.SCORE2_1_2B_DIS" & vbCrLf
        sql &= " ,a.SCORE2_1_2C_DIS" & vbCrLf
        sql &= " ,a.SCORE2_1_2D_DIS" & vbCrLf
        sql &= " ,a.SCORE2_1_2_SUM_ALL" & vbCrLf

        sql &= " ,a.SCORE2_1_3A" & vbCrLf 'SCORE2_1_3A-核定總班數
        sql &= " ,a.SCORE2_1_3" & vbCrLf
        sql &= " ,a.SCORE2_2_1" & vbCrLf
        sql &= " ,a.SCORE2_2_2_DIS" & vbCrLf
        sql &= " ,a.SCORE2_2_2" & vbCrLf
        sql &= " ,a.SCORE2_3_1" & vbCrLf

        sql &= " ,a.SCORE3_1" & vbCrLf
        sql &= " ,a.SCORE3_2" & vbCrLf

        sql &= " ,a.SCORE4_1" & vbCrLf
        sql &= " ,a.DEPTPNT" & vbCrLf
        sql &= " ,a.UNITPNT" & vbCrLf

        sql &= " ,a.SCORE4_2A" & vbCrLf
        sql &= " ,a.SCORE4_2_CNT" & vbCrLf
        sql &= " ,a.SCORE4_2_RATE" & vbCrLf
        sql &= " ,a.SCORE4_2" & vbCrLf
        sql &= " ,a.SCORE4_1_2" & vbCrLf
        '署／部加分項目
        sql &= " ,ISNULL(a.BRANCHPNT,0) BRANCHPNT" & vbCrLf
        'SUBTOTAL 分署小計/初審分數
        sql &= " ,ISNULL(a.SUBTOTAL,0) SUBTOTAL" & vbCrLf
        'LEVEL_1 初擬等級/初審等級
        'sql &= " ,convert(varchar(3),NULL) LEVEL_1" & vbCrLf
        sql &= " ,a.IMPSCORE_1" & vbCrLf 'IMPSCORE_1 初擬分數／匯入成績
        sql &= " ,a.IMPLEVEL_1" & vbCrLf 'IMPLEVEL_1 初擬等級／匯入等級
        sql &= " ,a.IMODIFYDATE,a.IMODIFYACCT" & vbCrLf '匯入時間/匯入帳號 

        'RLEVEL_2 複審等級  '有複審等級，使用複審等級，複審等級為空，使用匯入等級 'sql &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RLEVEL_2" & vbCrLf '複審等級
        sql &= " ,a.RLEVEL_2" & vbCrLf '複審等級
        '部長加分,部長加分小計,部長加分等級,署加分,
        sql &= ",a.MINISTERADD,a.MINISTERSUB,a.MINISTERLEVEL,a.DEPTADD" & vbCrLf '

        'TOTALSCORE  總分
        '計算優先順序：'【分署小計】+署/部加分項目/【匯入成績】'【匯入成績】優先 【分署小計】
        '(1)	當【匯入成績】有資料時 '總分 =【匯入成績】+【署/部加分項目】
        '(2)	當【匯入成績】為空、【分署小計】有資料時 '總分 =【分署小計】+【署/部加分項目】
        '(3)	當【匯入成績】、【分署小計】皆為空時     '總分=顯示 "-"
        sql &= " ,Case When a.IMPSCORE_1>0 Then convert(varchar,a.IMPSCORE_1+isnull(a.BRANCHPNT,0))" & vbCrLf
        sql &= "  When a.SUBTOTAL>0 Then convert(varchar,a.SUBTOTAL+isnull(a.BRANCHPNT,0)) Else '-' end TOTALSCORE" & vbCrLf
        sql &= " ,ISNULL(a.TOTAL,0) TOTAL" & vbCrLf
        'sql &= ",a.TOTAL" & vbCrLf
        '審核狀況(初審)/(複審)
        sql &= " ,a.FIRSTCHK ,a.SECONDCHK" & vbCrLf

        sql &= " ,a.CREATEDATE ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.FIRSTDATE  ,a.FIRSTACCT" & vbCrLf
        sql &= " ,a.SECONDATE  ,a.SECONACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE ,a.MODIFYACCT" & vbCrLf

        sql &= " ,oo.ORGNAME ,oo.COMIDNO" & vbCrLf
        sql &= " ,kd.NAME DISTNAME" & vbCrLf
        sql &= " ,k1.NAME ORGKIND_N" & vbCrLf
        '評核版本 'ddlSENDVER=Get_SENDVER_TS(ddlSENDVER) '評核結果 'ddlRESULT=Get_RESULT_TS(ddlRESULT)
        sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
        sql &= " ,v2.VNAME RESULT_N" & vbCrLf '獎牌
        sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
        sql &= " JOIN dbo.ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
        sql &= " LEFT JOIN dbo.ORG_TTQS2 b ON concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
        sql &= " LEFT JOIN dbo.V_SENDVER v1 ON v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.V_RESULT v2 ON v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
        sql &= " WHERE a.FIRSTCHK='Y'" & vbCrLf '(初審通過)
        '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
        sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        If vDISTID <> "" Then sql &= " AND a.DISTID=@DISTID" & vbCrLf
        sql &= " AND CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        'sql &= " AND a.YEARS=@YEARS" & vbCrLf
        'If vHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf '1:上年度 /2:下年度
        If vORGNAME <> "" Then sql &= " AND oo.ORGNAME Like '%" & vORGNAME & "%'" & vbCrLf
        If vCOMIDNO <> "" Then sql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
        'If vORGKIND2 <> "" Then sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf 'G/W
        Select Case vORGKIND2
            Case "G", "W"
                sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf 'G/W
        End Select

        If vORGKIND <> "" Then sql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
        'sql &= " ORDER BY a.OSID2" & vbCrLf
        sql &= " ORDER BY a.DISTID,a.SCORE4_1_2 DESC,oo.ORGNAME,a.OSID2" & vbCrLf

        Return sql
    End Function

    ''' <summary>Y:匯出 SQL</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function GET_SQL_EXP1(ByRef parms As Hashtable) As String
        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")
        Dim vORGNAME As String = TIMS.GetMyValue2(parms, "ORGNAME")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(parms, "COMIDNO")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")
        Dim vORGKIND As String = TIMS.GetMyValue2(parms, "ORGKIND")

        Dim sql As String = ""
        sql &= " SELECT oo.COMIDNO" & vbCrLf
        sql &= " ,ROW_NUMBER() OVER(ORDER BY a.DISTID,a.SCORE4_1_2 DESC,oo.ORGNAME,a.OSID2) AS ROWID" & vbCrLf
        sql &= " ,oo.ORGNAME,k1.NAME ORGKIND_N" & vbCrLf
        sql &= " ,oo.ORGKIND1,(SELECT x.ORGTYPE FROM dbo.VIEW_ORGTYPE1 x WHERE x.ORGTYPEID1=oo.ORGKIND1) ORGKIND1_N" & vbCrLf
        sql &= " ,kd.NAME DISTNAME" & vbCrLf
        'SCORING_N 審查計分區間
        sql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS) ,'年',a.MONTHS,'月'" & vbCrLf
        sql &= "    ,'(',dbo.FN_CYEAR2(a.YEARS1) ,'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end ,'~'" & vbCrLf
        sql &= "    ,dbo.FN_CYEAR2(a.YEARS2) ,'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end ,')') SCORING_N" & vbCrLf
        sql &= " ,CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
        sql &= " ,a.SCORE1_1" & vbCrLf '"1-1開班率(5%)(開班率=實際開班數/核定總班數)"
        sql &= " ,a.SCORE1_2" & vbCrLf '"1-2訓練人次達成率(8%)(訓練人次達成率=實際開訓人次/核定總人次)"
        'sql &= " ,a.SCORE2_1_1a + a.SCORE2_1_1b + a.SCORE2_1_1c + a.SCORE2_1_1d SCORE2_1_1" & vbCrLf
        'sql &= " ,a.SCORE2_1_2a + a.SCORE2_1_2b + a.SCORE2_1_2c + a.SCORE2_1_2d SCORE2_1_2" & vbCrLf
        'sql &= " ,CASE WHEN (a.SCORE2_1_2A_DIS+a.SCORE2_1_2B_DIS+a.SCORE2_1_2C_DIS+a.SCORE2_1_2D_DIS)<10.0" & vbCrLf
        'sql &= " THEN 10.0-(a.SCORE2_1_2A_DIS+a.SCORE2_1_2B_DIS+a.SCORE2_1_2C_DIS+a.SCORE2_1_2D_DIS) ELSE 0 END SCORE2_1_2" & vbCrLf
        sql &= " ,a.SCORE2_1_1_ALL" & vbCrLf 'SCORE2_1_1 '"2-1-1各項函送資料及資訊登錄作業時效(11%)(各班分數加總/核定總班數)"
        sql &= " ,a.SCORE2_1_2_SUM_ALL" & vbCrLf 'SCORE2_1_2 '"2-1-2函送資料內容及資訊登錄正確性(10%)(各班分數加總/核定總班數)"
        sql &= " ,a.SCORE2_1_3" & vbCrLf '"2-1-3訓練計畫變更項次數(7%)(各班分數加總/核定總班數)"
        sql &= " ,a.SCORE2_2_1" & vbCrLf '"2-2-1學員管理(4%)(各班分數加總/核定總班數)"
        sql &= " ,a.SCORE2_2_2_DIS" & vbCrLf
        sql &= " ,a.SCORE2_2_2" & vbCrLf '"2-2-2課程辦理情形(30%)
        sql &= " ,a.SCORE2_3_1" & vbCrLf '"2-3-1計畫說明、訓練活動及相關會議之出席率(5%)(計畫參與度=實際出席總場次/應出席總場次)"
        sql &= " ,a.SCORE3_1" & vbCrLf '"3-1最近一次TTQS評核結果等級(10%)
        sql &= " ,a.SCORE3_2" & vbCrLf '3-2學員滿意程度(5%)
        sql &= " ,a.SCORE4_1" & vbCrLf 'SCORE4_1 配合分署辦理相關活動或政策宣導(3%)
        sql &= " ,a.SCORE4_2A" & vbCrLf 'SCORE4_2A 參訓學員訓後動態調查表填寫人次
        sql &= " ,a.SCORE4_2_CNT" & vbCrLf 'SCORE4_2_CNT 結訓學員總人次
        sql &= " ,a.SCORE4_2_RATE" & vbCrLf 'SCORE4_2_RATE 參訓學員平均填答率
        sql &= " ,a.SCORE4_2" & vbCrLf ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
        sql &= " ,a.SCORE4_1_2" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
        sql &= " ,ISNULL(a.BRANCHPNT,0) BRANCHPNT" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
        sql &= " ,ISNULL(a.SUBTOTAL,0) SUBTOTAL" & vbCrLf '分署小計/初審分數
        sql &= " ,ISNULL(a.TOTAL,0) TOTAL" & vbCrLf
        '評核版本 'ddlSENDVER=Get_SENDVER_TS(ddlSENDVER)
        '評核結果 'ddlRESULT=Get_RESULT_TS(ddlRESULT)
        sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
        sql &= " ,v2.VNAME RESULT_N" & vbCrLf '獎牌
        sql &= " ,a.IMPSCORE_1" & vbCrLf 'IMPSCORE_1 初擬分數／匯入成績
        sql &= " ,a.IMPLEVEL_1" & vbCrLf 'IMPLEVEL_1 初擬等級／匯入等級/初審等級
        '有複審等級，使用複審等級，複審等級為空，使用匯入等級 'sql &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RLEVEL_2" & vbCrLf '複審等級
        sql &= " ,a.SECONDCHK" & vbCrLf '複審通過
        '總分-複審等級-通過才顯示
        '計算優先順序：'【分署小計】+署/部加分項目/【匯入成績】'【匯入成績】優先 【分署小計】
        '(1)	當【匯入成績】有資料時 '總分 =【匯入成績】+【署/部加分項目】
        '(2)	當【匯入成績】為空、【分署小計】有資料時 '總分 =【分署小計】+【署/部加分項目】
        '(3)	當【匯入成績】、【分署小計】皆為空時     '總分=顯示 "-"
        sql &= " ,case when a.IMPSCORE_1>0 then convert(varchar,a.IMPSCORE_1+isnull(a.BRANCHPNT,0))" & vbCrLf
        sql &= "  when a.SUBTOTAL>0 then convert(varchar,a.SUBTOTAL+isnull(a.BRANCHPNT,0)) else '-' end TOTALSCORE" & vbCrLf
        '複審等級-通過才顯示 'sql &= " ,case when a.SECONDCHK='Y' THEN a.RLEVEL_2 END RLEVEL_2" & vbCrLf
        sql &= " ,a.RLEVEL_2,a.MINISTERADD,a.MINISTERSUB,a.MINISTERLEVEL,a.DEPTADD" & vbCrLf
        'sql &= " ,a.IMODIFYDATE" & vbCrLf'sql &= " ,a.IMODIFYACCT" & vbCrLf
        sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
        sql &= " JOIN ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
        sql &= " LEFT JOIN dbo.ORG_TTQS2 b ON concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
        sql &= " LEFT JOIN dbo.V_SENDVER v1 ON v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.V_RESULT v2 ON v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
        sql &= " WHERE a.FIRSTCHK='Y'" & vbCrLf '(初審通過)
        '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
        sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        sql &= " AND a.TPLANID=@TPLANID" & vbCrLf
        If vDISTID <> "" Then sql &= " AND a.DISTID=@DISTID" & vbCrLf
        sql &= " AND CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        'sql &= " AND a.YEARS=@YEARS" & vbCrLf
        'If vHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf '1:上年度 /2:下年度
        If vORGNAME <> "" Then sql &= " AND oo.ORGNAME LIKE '%" & vORGNAME & "%'" & vbCrLf
        If vCOMIDNO <> "" Then sql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
        'If vORGKIND2 <> "" Then sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf 'G/W
        Select Case vORGKIND2
            Case "G", "W"
                sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
                'parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then sql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
        sql &= " ORDER BY a.DISTID,a.SCORE4_1_2 DESC,oo.ORGNAME,a.OSID2" & vbCrLf
        Return sql
    End Function

    ''' <summary>'1:查詢一筆資料[ORG_SCORING2]</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function GET_SQL_ONEDATA1(ByRef parms As Hashtable) As String
        Dim vOSID2 As String = TIMS.GetMyValue2(parms, "OSID2")
        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")

        Dim sql As String = ""
        sql &= " SELECT a.*" & vbCrLf 'dbo.ORG_SCORING2 a WITH(NOLOCK)
        sql &= " ,oo.ORGNAME" & vbCrLf 'sql &= " ,oo.COMIDNO" & vbCrLf
        sql &= " ,kd.NAME DISTNAME" & vbCrLf
        sql &= " ,k1.NAME ORGKIND_N" & vbCrLf
        'SCORING_N 審查計分區間
        sql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS) ,'年',a.MONTHS,'月'" & vbCrLf
        sql &= "  ,'(',dbo.FN_CYEAR2(a.YEARS1) ,'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end ,'~'" & vbCrLf
        sql &= "  ,dbo.FN_CYEAR2(a.YEARS2) ,'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end ,')') SCORING_N" & vbCrLf
        sql &= " ,CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
        'TOTALSCORE  總分
        '計算優先順序：'【分署小計】+署/部加分項目/【匯入成績】'【匯入成績】優先 【分署小計】
        '(1)	當【匯入成績】有資料時 '總分 =【匯入成績】+【署/部加分項目】
        '(2)	當【匯入成績】為空、【分署小計】有資料時 '總分 =【分署小計】+【署/部加分項目】
        '(3)	當【匯入成績】、【分署小計】皆為空時     '總分=顯示 "-"
        sql &= " ,case when a.IMPSCORE_1>0 then convert(varchar,a.IMPSCORE_1+isnull(a.BRANCHPNT,0))" & vbCrLf
        sql &= "  when a.SUBTOTAL>0 then convert(varchar,a.SUBTOTAL+isnull(a.BRANCHPNT,0)) else '-' end TOTALSCORE" & vbCrLf
        '評核版本 'ddlSENDVER=Get_SENDVER_TS(ddlSENDVER)
        '評核結果 'ddlRESULT=Get_RESULT_TS(ddlRESULT)
        sql &= " ,v1.VNAME SENDVER_N ,v2.VNAME RESULT_N" & vbCrLf
        sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
        sql &= " JOIN ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
        sql &= " LEFT JOIN dbo.ORG_TTQS2 b ON concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
        sql &= " LEFT JOIN dbo.V_SENDVER v1 ON v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
        sql &= " LEFT JOIN dbo.V_RESULT v2 ON v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
        sql &= " WHERE a.FIRSTCHK='Y'" & vbCrLf '(初審通過)
        '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
        sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        sql &= " AND a.OSID2=@OSID2" & vbCrLf
        Return sql
    End Function

    ''' <summary> OUTPUT ORG_SCORING2 ALL SQL get TABLE </summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Public Function Get_dtORGSCORING2(ByVal parms As Hashtable) As DataTable
        Dim vEXP As String = TIMS.GetMyValue2(parms, "EXP") 'Y:匯出 /1:一筆資料[ORG_SCORING2]
        Dim vOSID2 As String = TIMS.GetMyValue2(parms, "OSID2")
        Dim vTPLANID As String = TIMS.GetMyValue2(parms, "TPLANID")
        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")
        'Dim vYEARS As String=TIMS.GetMyValue2(parms, "YEARS")
        'Dim vHALFYEAR As String=TIMS.GetMyValue2(parms, "HALFYEAR") '1:上年度 /2:下年度
        Dim vORGNAME As String = TIMS.GetMyValue2(parms, "ORGNAME")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(parms, "COMIDNO")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")
        Dim vORGKIND As String = TIMS.GetMyValue2(parms, "ORGKIND")

        Dim dt As DataTable = Nothing
        Dim sql As String = ""
        '(Y:匯出 / 1:查詢一筆資料[ORG_SCORING2] /S:查詢多筆)
        Select Case vEXP
            Case "Y" 'Y:匯出 
                sql = GET_SQL_EXP1(parms)
            Case "1" '1:查詢一筆資料[ORG_SCORING2]  
                sql = GET_SQL_ONEDATA1(parms)
            Case "S" 'S:查詢多筆 'SEARCH1 DATAGRID
                sql = GET_SQL_DT1(parms)
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                Return dt
        End Select

        dt = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    ''' <summary>DG儲存-多筆勾選儲存</summary>
    Sub SSaveData1()
        Dim iChkCnt As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim drv As DataRowView = eItem.DataItem
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            Dim HidSUBTOTAL As HiddenField = eItem.FindControl("HidSUBTOTAL")
            '部長加分,部長加分小計,部長加分等級,署加分,
            Dim tMINISTERADD As TextBox = eItem.FindControl("tMINISTERADD") '部長加分
            Dim LabMINISTERSUB As Label = eItem.FindControl("LabMINISTERSUB") '部長加分小計
            Dim ddlMINISTERLEVEL As DropDownList = eItem.FindControl("ddlMINISTERLEVEL") '部長加分等級
            Dim tDEPTADD As TextBox = eItem.FindControl("tDEPTADD") '署加分
            Dim LabTOTALSCORE As Label = eItem.FindControl("LabTOTALSCORE") '總分
            'Dim Hid_BRANCHPNTorg As HiddenField = eItem.FindControl("Hid_BRANCHPNTorg") '署／部加分項目
            'Dim tBRANCHPNT As TextBox = eItem.FindControl("tBRANCHPNT") '分署<br>加分項目
            Dim Hid_RLEVEL_2 As HiddenField = eItem.FindControl("Hid_RLEVEL_2") '複審等級
            Dim ddlRLEVEL_2 As DropDownList = eItem.FindControl("ddlRLEVEL_2") '複審等級
            Dim vRLEVEL_2 As String = TIMS.GetListValue(ddlRLEVEL_2)
            Dim HidOSID2 As HiddenField = eItem.FindControl("HidOSID2")
            'Hid_SECONDCHKorg
            Dim Hid_SECONDCHKorg As HiddenField = eItem.FindControl("Hid_SECONDCHKorg")
            Dim ddlSECONDCHK As DropDownList = eItem.FindControl("ddlSECONDCHK")
            Dim vSECONDCHKorg As String = TIMS.ClearSQM(Hid_SECONDCHKorg.Value)
            Dim vddlSECONDCHK As String = TIMS.GetListValue(ddlSECONDCHK)

            Dim vOSID2 As String = TIMS.ClearSQM(HidOSID2.Value)
            Dim fgCanSave1 As Boolean = (checkbox1.Checked AndAlso vOSID2 <> "") '有勾選且有序號可以儲存
            If checkbox1.Checked Then iChkCnt += 1
            If fgCanSave1 Then
                Select Case vddlSECONDCHK'vFIRSTCHK
                    Case "Y", "N"
                    Case Else
                        Common.MessageBox(Me, String.Concat("(", iChkCnt, ")請選擇審核狀態!"))
                        Exit Sub
                End Select

                Dim errmsg1 As String = ""
                HidSUBTOTAL.Value = TIMS.ClearSQM(HidSUBTOTAL.Value)
                Dim vSUBTOTAL As Double = TIMS.VAL1(HidSUBTOTAL.Value)
                'BRANCHPNT.Text = TIMS.ClearSQM(BRANCHPNT.Text)
                'Dim vBRANCHPNT As Double = TIMS.VAL1(BRANCHPNT.Text) '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                tMINISTERADD.Text = TIMS.ClearSQM(tMINISTERADD.Text) '部加分
                If tMINISTERADD.Text = "" Then tMINISTERADD.Text = "0"
                Dim vMINISTERADD As Double = TIMS.VAL1(tMINISTERADD.Text)
                Dim FG_CHK_MINISTERADD_ISNUM As Boolean = TIMS.IsNumeric1(tMINISTERADD.Text) 'TRUE:數字 /FALSE:(非數字)
                If tMINISTERADD.Text = "" OrElse Not FG_CHK_MINISTERADD_ISNUM Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")部加分，請填寫0-7的數字!") & vbCrLf
                ElseIf vMINISTERADD > 7 OrElse vMINISTERADD < 0 Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")部加分，請填寫0-7的數字!!") & vbCrLf
                End If
                tDEPTADD.Text = TIMS.ClearSQM(tDEPTADD.Text) '署加分
                If tDEPTADD.Text = "" Then tDEPTADD.Text = "0"
                Dim vDEPTADD As Double = TIMS.VAL1(tDEPTADD.Text)
                Dim FG_CHK_DEPTADD_ISNUM As Boolean = TIMS.IsNumeric1(tDEPTADD.Text) 'TRUE:數字 /FALSE:(非數字)
                If tDEPTADD.Text = "" OrElse Not FG_CHK_DEPTADD_ISNUM Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")署加分，請填寫0-7的數字!") & vbCrLf
                ElseIf vDEPTADD > 7 OrElse vDEPTADD < 0 Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")署加分，請填寫0-7的數字!!") & vbCrLf
                End If
                Dim vBRANCHPNT As Double = vMINISTERADD + vDEPTADD '部長加分+署加分
                If vBRANCHPNT > 7 OrElse vBRANCHPNT < 0 Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")署部加分合計，請確認為0-7的數字!!") & vbCrLf
                End If
                Dim vMINISTERSUB As Double = vSUBTOTAL + vMINISTERADD
                If vMINISTERSUB > 100 OrElse vMINISTERSUB < 0 Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")部加分小計，請確認為0-100的數字!!") & vbCrLf
                End If
                Dim vSCORE4_1_2 As Double = vSUBTOTAL + vMINISTERADD + vDEPTADD
                If vSCORE4_1_2 > 100 OrElse vSCORE4_1_2 < 0 Then
                    errmsg1 &= String.Concat("(", iChkCnt, ")總分，請確認為0-100的數字!!") & vbCrLf
                End If

                'MINISTERSUB.Text = TIMS.ClearSQM(MINISTERSUB.Text) '部加分小計
                'Dim vMINISTERSUB As Double = TIMS.VAL1(MINISTERSUB.Text) '部加分小計
                'Dim FG_CHK_MINISTERSUB_ISNUM As Boolean = TIMS.IsNumeric1(MINISTERSUB.Text) 'TRUE:數字 /FALSE:(非數字)
                'If MINISTERSUB.Text = "" OrElse Not FG_CHK_MINISTERSUB_ISNUM Then
                '    errmsg1 &= ("部加分小計，請填寫0-100的數字!") & vbCrLf
                'ElseIf vMINISTERSUB > 100 OrElse vMINISTERSUB < 0 Then
                '    errmsg1 &= ("部加分小計，請填寫0-100的數字!!") & vbCrLf
                'End If
                'SCORE4_1_2.Text = TIMS.ClearSQM(SCORE4_1_2.Text) '總分
                'Dim vSCORE4_1_2 As Double = TIMS.VAL1(SCORE4_1_2.Text)
                'Dim FG_CHK_SCORE4_1_2_ISNUM As Boolean = TIMS.IsNumeric1(SCORE4_1_2.Text) 'TRUE:數字 /FALSE:(非數字)
                'If SCORE4_1_2.Text = "" OrElse Not FG_CHK_SCORE4_1_2_ISNUM Then
                '    errmsg1 &= ("總分，請填寫0-100的數字!") & vbCrLf
                'ElseIf vSCORE4_1_2 > 100 OrElse vSCORE4_1_2 < 0 Then
                '    errmsg1 &= ("總分，請填寫0-100的數字!!") & vbCrLf
                'End If
                vSUBTOTAL = TIMS.ROUND(vSUBTOTAL, 1)
                vMINISTERADD = TIMS.ROUND(vMINISTERADD, 1)
                vDEPTADD = TIMS.ROUND(vDEPTADD, 1)
                vBRANCHPNT = TIMS.ROUND(vBRANCHPNT, 1)
                Select Case vBRANCHPNT'SCORE4_1.Text
                    Case 0
                    Case Else
                        'If Not TIMS.IsNumeric1(BRANCHPNT.Text) Then 'TRUE:數字 /FALSE:(非數字)
                        '    errmsg1 &= String.Concat("加分項目由本部、本署自填，請填寫數字，至多加 7 分!", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vbCrLf)
                        'ElseIf (vBRANCHPNT > 7) OrElse (vBRANCHPNT < 0) Then
                        '    errmsg1 &= String.Concat("加分項目由本部、本署自填，請填寫數字，至多加 7 分!!", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vbCrLf)
                        'End If
                        If (vBRANCHPNT > 7) OrElse (vBRANCHPNT < 0) Then
                            errmsg1 &= String.Concat("加分項目由本部、本署自填，請填寫數字，至多加 7 分!!", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vbCrLf)
                        End If
                End Select
                If errmsg1 <> "" Then
                    Common.MessageBox(Me, errmsg1)
                    Exit Sub
                End If

                'vMINISTERSUB = TIMS.ROUND(vMINISTERSUB, 1)
                'vSCORE4_1_2 = TIMS.ROUND(vSCORE4_1_2, 1)

                'If Not TIMS.VAL1_Equal(vBRANCHPNT, vMINISTERADD + vDEPTADD) Then
                '    errmsg1 &= String.Concat("加分項目由本部、本署自填，部署合計有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vBRANCHPNT, ",", (vMINISTERADD + vDEPTADD), vbCrLf)
                'ElseIf Not TIMS.VAL1_Equal(vMINISTERSUB, vSUBTOTAL + vMINISTERADD) Then
                '    errmsg1 &= String.Concat("加分項目由本部、本署自填，部加分小計有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vMINISTERSUB, ",", (vSUBTOTAL + vMINISTERADD), vbCrLf)
                'ElseIf Not TIMS.VAL1_Equal(vSCORE4_1_2, vSUBTOTAL + vMINISTERADD + vDEPTADD) Then
                '    errmsg1 &= String.Concat("加分項目由本部、本署自填，總分合計有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vSCORE4_1_2, ",", (vSUBTOTAL + vMINISTERADD + vDEPTADD), vbCrLf)
                'ElseIf Not TIMS.VAL1_Equal(vSCORE4_1_2, vSUBTOTAL + vBRANCHPNT) Then
                '    errmsg1 &= String.Concat("加分項目由本部、本署自填，部署加分有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vSCORE4_1_2, ",", (vSUBTOTAL + vBRANCHPNT), vbCrLf)
                'End If
                'If errmsg1 <> "" Then Return False 'Exit Sub 

                'Dim flag_chk1_NG As Boolean = False '非數字
                'Dim flag_chk2_NG As Boolean = False '數字超過範圍
                'If vBRANCHPNT <> "" Then
                '    '有變動過且不為空白
                '    If (Not TIMS.IsNumeric1(vBRANCHPNT)) Then flag_chk1_NG = True
                '    If (Not flag_chk1_NG) AndAlso ((Val(vBRANCHPNT) > 7) OrElse (Val(vBRANCHPNT) < 0)) Then flag_chk2_NG = True
                '    If flag_chk1_NG Then
                '        Common.MessageBox(Me, "加分項目配合本部、本署辦理相關活動 或政策宣導(7%)，應為數字格式，至多加 7 分!")
                '        Exit Sub
                '    ElseIf flag_chk2_NG Then
                '        Common.MessageBox(Me, "加分項目配合本部、本署辦理相關活動 或政策宣導(7%)，至多加 7 分!")
                '        Exit Sub
                '    End If
                'End If

            Else
                If vddlSECONDCHK <> vSECONDCHKorg Then
                    Common.MessageBox(Me, String.Concat("(", iChkCnt, ")請勾選要儲存的項目!(有變更過審核)"))
                    Exit Sub
                End If
                'If vBRANCHPNT <> vBRANCHPNTorg Then
                '    Common.MessageBox(Me, "請勾選要儲存的項目!(有變更過加分項目)")
                '    Exit Sub
                'End If
            End If
        Next
        If iChkCnt = 0 Then
            Common.MessageBox(Me, "請勾選要儲存的項目!(未有勾選)")
            Exit Sub
        End If

        For Each eItem As DataGridItem In DataGrid1.Items
            Dim drv As DataRowView = eItem.DataItem
            'Dim CheckBox1 As CheckBox=eItem.FindControl("CheckBox1")
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            Dim HidSUBTOTAL As HiddenField = eItem.FindControl("HidSUBTOTAL")
            '部長加分,部長加分小計,部長加分等級,署加分,
            Dim tMINISTERADD As TextBox = eItem.FindControl("tMINISTERADD") '部長加分
            Dim LabMINISTERSUB As Label = eItem.FindControl("LabMINISTERSUB") '部長加分小計
            Dim ddlMINISTERLEVEL As DropDownList = eItem.FindControl("ddlMINISTERLEVEL") '部長加分等級
            Dim tDEPTADD As TextBox = eItem.FindControl("tDEPTADD") '署加分
            Dim LabTOTALSCORE As Label = eItem.FindControl("LabTOTALSCORE") '總分
            'Dim Hid_BRANCHPNTorg As HiddenField = eItem.FindControl("Hid_BRANCHPNTorg") '署／部加分項目
            'Dim tBRANCHPNT As TextBox = eItem.FindControl("tBRANCHPNT") '分署<br>加分項目
            Dim Hid_RLEVEL_2 As HiddenField = eItem.FindControl("Hid_RLEVEL_2") '複審等級
            Dim ddlRLEVEL_2 As DropDownList = eItem.FindControl("ddlRLEVEL_2") '複審等級
            Dim vRLEVEL_2 As String = TIMS.GetListValue(ddlRLEVEL_2)
            Dim HidOSID2 As HiddenField = eItem.FindControl("HidOSID2")
            'Hid_SECONDCHKorg
            Dim Hid_SECONDCHKorg As HiddenField = eItem.FindControl("Hid_SECONDCHKorg")
            Dim ddlSECONDCHK As DropDownList = eItem.FindControl("ddlSECONDCHK")
            Dim vSECONDCHKorg As String = TIMS.ClearSQM(Hid_SECONDCHKorg.Value)
            Dim vddlSECONDCHK As String = TIMS.GetListValue(ddlSECONDCHK)

            tMINISTERADD.Text = TIMS.ClearSQM(tMINISTERADD.Text) '部長加分
            If tMINISTERADD.Text = "" Then tMINISTERADD.Text = "0"
            tDEPTADD.Text = TIMS.ClearSQM(tDEPTADD.Text) '署加分
            If tDEPTADD.Text = "" Then tDEPTADD.Text = "0"

            HidSUBTOTAL.Value = TIMS.ClearSQM(HidSUBTOTAL.Value)
            Dim vSUBTOTAL As String = TIMS.VAL1(HidSUBTOTAL.Value)
            Dim vMINISTERADD As String = TIMS.VAL1(tMINISTERADD.Text)
            Dim vMINISTERSUB As String = TIMS.VAL1(HidSUBTOTAL.Value) + TIMS.VAL1(tMINISTERADD.Text)
            Dim vMINISTERLEVEL As String = TIMS.GetListValue(ddlMINISTERLEVEL)
            Dim vDEPTADD As String = TIMS.VAL1(tDEPTADD.Text)
            Dim vBRANCHPNT As String = TIMS.VAL1(vMINISTERADD) + TIMS.VAL1(vDEPTADD) '部長加分+署加分
            'Dim vSECONDCHK As String=TIMS.ClearSQM(ddlSECONDCHK.SelectedValue)
            Dim vSECONDCHK As String = TIMS.GetListValue(ddlSECONDCHK)
            Dim vOSID2 As String = TIMS.ClearSQM(HidOSID2.Value)
            Dim flagCanSave1 As Boolean = If(checkbox1.Checked AndAlso vOSID2 <> "", True, False)
            If flagCanSave1 Then
                'parms.Add("SUBTOTAL", vSUBTOTAL)
                Dim pms_u1 As New Hashtable From {
                    {"SECONDCHK", If(vSECONDCHK <> "", vSECONDCHK, Convert.DBNull)},
                    {"SECONACCT", sm.UserInfo.UserID},
                    {"BRANCHPNT", If(vBRANCHPNT <> "", Val(vBRANCHPNT), 0)}, '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                    {"MINISTERADD", If(vMINISTERADD <> "", Val(vMINISTERADD), 0)},
                    {"MINISTERSUB", If(vMINISTERSUB <> "", Val(vMINISTERSUB), Val(vSUBTOTAL))},
                    {"MINISTERLEVEL", If(vMINISTERLEVEL <> "", vMINISTERLEVEL, Convert.DBNull)}, '部加分等級
                    {"DEPTADD", If(vDEPTADD <> "", Val(vDEPTADD), 0)},
                    {"RLEVEL_2", If(vRLEVEL_2 <> "", vRLEVEL_2, Convert.DBNull)},
                    {"MODIFYACCT", sm.UserInfo.UserID},
                    {"OSID2", Val(vOSID2)}
                }

                Dim u_sql As String = ""
                u_sql &= " UPDATE ORG_SCORING2" & vbCrLf
                u_sql &= " SET SECONDCHK=@SECONDCHK,SECONACCT=@SECONACCT,SECONDATE=GETDATE()" & vbCrLf
                u_sql &= " ,BRANCHPNT=@BRANCHPNT" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                '部長加分,部長加分小計,部長加分等級,署加分,
                u_sql &= " ,MINISTERADD=@MINISTERADD,MINISTERSUB=@MINISTERSUB,MINISTERLEVEL=@MINISTERLEVEL,DEPTADD=@DEPTADD" & vbCrLf
                u_sql &= " ,RLEVEL_2=@RLEVEL_2" & vbCrLf 'sql &= " ,SUBTOTAL=@SUBTOTAL" & vbCrLf '複審等級
                u_sql &= " ,SCORE4_1_2=ISNULL(@BRANCHPNT,0.0)+ISNULL(SUBTOTAL,0.0)" & vbCrLf '總分
                u_sql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
                u_sql &= " WHERE OSID2=@OSID2" & vbCrLf
                DbAccess.ExecuteNonQuery(u_sql, objconn, pms_u1)
            End If
        Next

        'divSch1.Visible=True 'divEdt1.Visible=False
        sm.LastResultMessage = "儲存完畢"
        Call SSearch1()
    End Sub

    ''' <summary> 單一資料儲存-檢核 </summary>
    ''' <param name="errmsg1"></param>
    ''' <param name="htSS"></param>
    ''' <returns></returns>
    Function CheckData2(ByRef errmsg1 As String, ByRef htSS As Hashtable) As Boolean
        'Dim vBRANCHPNT As String=TIMS.ClearSQM(vBRANCHPNT)
        'SUBTOTAL.Text=TIMS.ClearSQM(SUBTOTAL.Text)
        'Dim vSUBTOTAL As String=SUBTOTAL.Text
        'If vSUBTOTAL="" Then vSUBTOTAL="0"
        Dim vOSID2 As String = TIMS.ClearSQM(Hid_OSID2.Value)
        If vOSID2 = "" Then
            errmsg1 &= "儲存資料有誤!" & vbCrLf
            Return False
        End If

        '部長加分,部長加分小計,部長加分等級,署加分,
        'Dim vMINISTERADD As String = TIMS.ClearSQM(MINISTERADD.Text)
        'Dim vDEPTADD As String = TIMS.ClearSQM(DEPTADD.Text)
        'Dim vMINISTERSUB As String = TIMS.ClearSQM(MINISTERSUB.Text) '部加分小計
        'Dim vMINISTERLEVEL As String = TIMS.GetListValue(ddl_MINISTERLEVEL)
        '配合分署辦理相關活動或政策宣導 '0-4
        'SCORE4_1.Text=TIMS.ClearSQM(SCORE4_1.Text) '分署 加分項目 加分項目 配合分署辦理相關活動或政策宣導(3%)
        'BRANCHPNT.Text = TIMS.ClearSQM(BRANCHPNT.Text) '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
        '部長加分,部長加分小計,部長加分等級,署加分,
        'MINISTERADD.Text = TIMS.ClearSQM(MINISTERADD.Text) '部加分
        'DEPTADD.Text = TIMS.ClearSQM(DEPTADD.Text) '署加分
        'MINISTERSUB.Text = TIMS.ClearSQM(MINISTERSUB.Text)
        'Hid_MINISTERLEVEL.Value = TIMS.ClearSQM(Hid_MINISTERLEVEL.Value)
        'Dim vddl_MINISTERLEVEL As String = TIMS.GetListValue(ddl_MINISTERLEVEL)
        ''Dim vBRANCHPNT As String=SCORE4_1.Text '分署 加分項目
        'Dim vBRANCHPNT As String = TIMS.VAL1(BRANCHPNT.Text) '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
        'Dim vMINISTERADD As String = TIMS.VAL1(MINISTERADD.Text) '部加分
        'Dim vDEPTADD As String = TIMS.VAL1(DEPTADD.Text)  '署加分
        'If vOSID2="" Then Exit Sub
        'Dim vFIRSTCHK As String=TIMS.ClearSQM(ddlFIRSTCHK_1.SelectedValue)
        'Select Case vFIRSTCHK
        '    Case "Y", "N"
        '    Case Else
        '        errmsg1 &= "請選擇審核狀態!" & vbCrLf
        '        Return False
        '        'Common.MessageBox(Me, "請選擇審核狀態!") Exit Sub
        'End Select

        SUBTOTAL.Text = TIMS.ClearSQM(SUBTOTAL.Text)
        Dim vSUBTOTAL As Double = TIMS.VAL1(SUBTOTAL.Text)
        BRANCHPNT.Text = TIMS.ClearSQM(BRANCHPNT.Text)
        Dim vBRANCHPNT As Double = TIMS.VAL1(BRANCHPNT.Text) '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)

        MINISTERADD.Text = TIMS.ClearSQM(MINISTERADD.Text) '部加分
        Dim vMINISTERADD As Double = TIMS.VAL1(MINISTERADD.Text)
        Dim FG_CHK_MINISTERADD_ISNUM As Boolean = TIMS.IsNumeric1(MINISTERADD.Text) 'TRUE:數字 /FALSE:(非數字)
        If MINISTERADD.Text = "" OrElse Not FG_CHK_MINISTERADD_ISNUM Then
            errmsg1 &= ("部加分，請填寫0-7的數字!") & vbCrLf
        ElseIf vMINISTERADD > 7 OrElse vMINISTERADD < 0 Then
            errmsg1 &= ("部加分，請填寫0-7的數字!!") & vbCrLf
        End If
        DEPTADD.Text = TIMS.ClearSQM(DEPTADD.Text) '署加分
        Dim vDEPTADD As Double = TIMS.VAL1(DEPTADD.Text)
        Dim FG_CHK_DEPTADD_ISNUM As Boolean = TIMS.IsNumeric1(DEPTADD.Text) 'TRUE:數字 /FALSE:(非數字)
        If DEPTADD.Text = "" OrElse Not FG_CHK_DEPTADD_ISNUM Then
            errmsg1 &= ("署加分，請填寫0-7的數字!") & vbCrLf
        ElseIf vDEPTADD > 7 OrElse vDEPTADD < 0 Then
            errmsg1 &= ("署加分，請填寫0-7的數字!!") & vbCrLf
        End If
        MINISTERSUB.Text = TIMS.ClearSQM(MINISTERSUB.Text) '部加分小計
        Dim vMINISTERSUB As Double = TIMS.VAL1(MINISTERSUB.Text) '部加分小計
        Dim FG_CHK_MINISTERSUB_ISNUM As Boolean = TIMS.IsNumeric1(MINISTERSUB.Text) 'TRUE:數字 /FALSE:(非數字)
        If MINISTERSUB.Text = "" OrElse Not FG_CHK_MINISTERSUB_ISNUM Then
            errmsg1 &= ("部加分小計，請填寫0-100的數字!") & vbCrLf
        ElseIf vMINISTERSUB > 100 OrElse vMINISTERSUB < 0 Then
            errmsg1 &= ("部加分小計，請填寫0-100的數字!!") & vbCrLf
        End If
        SCORE4_1_2.Text = TIMS.ClearSQM(SCORE4_1_2.Text) '總分
        Dim vSCORE4_1_2 As Double = TIMS.VAL1(SCORE4_1_2.Text)
        Dim FG_CHK_SCORE4_1_2_ISNUM As Boolean = TIMS.IsNumeric1(SCORE4_1_2.Text) 'TRUE:數字 /FALSE:(非數字)
        If SCORE4_1_2.Text = "" OrElse Not FG_CHK_SCORE4_1_2_ISNUM Then
            errmsg1 &= ("總分，請填寫0-100的數字!") & vbCrLf
        ElseIf vSCORE4_1_2 > 100 OrElse vSCORE4_1_2 < 0 Then
            errmsg1 &= ("總分，請填寫0-100的數字!!") & vbCrLf
        End If
        Select Case vBRANCHPNT'SCORE4_1.Text
            Case 0
            Case Else
                If Not TIMS.IsNumeric1(BRANCHPNT.Text) Then 'TRUE:數字 /FALSE:(非數字)
                    errmsg1 &= String.Concat("加分項目由本部、本署自填，請填寫數字，至多加 7 分!", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vbCrLf)
                ElseIf (vBRANCHPNT > 7) OrElse (vBRANCHPNT < 0) Then
                    errmsg1 &= String.Concat("加分項目由本部、本署自填，請填寫數字，至多加 7 分!!", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vbCrLf)
                End If
        End Select
        If errmsg1 <> "" Then Return False 'Exit Sub 

        vSUBTOTAL = TIMS.ROUND(vSUBTOTAL, 1)
        vMINISTERADD = TIMS.ROUND(vMINISTERADD, 1)
        vDEPTADD = TIMS.ROUND(vDEPTADD, 1)
        vBRANCHPNT = TIMS.ROUND(vBRANCHPNT, 1)
        vMINISTERSUB = TIMS.ROUND(vMINISTERSUB, 1)
        vSCORE4_1_2 = TIMS.ROUND(vSCORE4_1_2, 1)

        If Not TIMS.VAL1_Equal(vBRANCHPNT, vMINISTERADD + vDEPTADD) Then
            errmsg1 &= String.Concat("加分項目由本部、本署自填，部署合計有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vBRANCHPNT, ",", (vMINISTERADD + vDEPTADD), vbCrLf)
        ElseIf Not TIMS.VAL1_Equal(vMINISTERSUB, vSUBTOTAL + vMINISTERADD) Then
            errmsg1 &= String.Concat("加分項目由本部、本署自填，部加分小計有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vMINISTERSUB, ",", (vSUBTOTAL + vMINISTERADD), vbCrLf)
        ElseIf Not TIMS.VAL1_Equal(vSCORE4_1_2, vSUBTOTAL + vMINISTERADD + vDEPTADD) Then
            errmsg1 &= String.Concat("加分項目由本部、本署自填，總分合計有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vSCORE4_1_2, ",", (vSUBTOTAL + vMINISTERADD + vDEPTADD), vbCrLf)
        ElseIf Not TIMS.VAL1_Equal(vSCORE4_1_2, vSUBTOTAL + vBRANCHPNT) Then
            errmsg1 &= String.Concat("加分項目由本部、本署自填，部署加分有誤！", vbCrLf, "配合本部、本署辦理相關活動 或政策宣導 (7%)", vSCORE4_1_2, ",", (vSUBTOTAL + vBRANCHPNT), vbCrLf)
        End If
        If errmsg1 <> "" Then Return False 'Exit Sub 

        Dim vMINISTERLEVEL As String = TIMS.GetListValue(ddl_MINISTERLEVEL) '.ClearSQM(ddlRLEVEL_2R.SelectedValue)
        If vMINISTERLEVEL = "" Then
            errmsg1 &= "請選擇 部加分等級!" & vbCrLf 'Return False
        End If
        '複審等級 ddlRLEVEL_2R 
        Dim vddlRLEVEL_2R As String = TIMS.GetListValue(ddlRLEVEL_2R) '.ClearSQM(ddlRLEVEL_2R.SelectedValue)
        If vddlRLEVEL_2R = "" Then
            errmsg1 &= "請選擇 複審等級!" & vbCrLf 'Return False
        End If
        'ddlSECONDCHK_1
        Dim vddlSECONDCHK_1 As String = TIMS.GetListValue(ddlSECONDCHK_1)
        If vddlSECONDCHK_1 = "" Then
            errmsg1 &= "請選擇審核狀態!" & vbCrLf 'Return False
        End If
        If errmsg1 <> "" Then Return False 'Exit Sub 

        'htSS.Add("vSUBTOTAL", vSUBTOTAL) '小計 '部長加分,部長加分小計,部長加分等級,署加分,
        htSS = New Hashtable From {
            {"vSECONDCHK", vddlSECONDCHK_1}, 'vFIRSTCHK)
            {"vBRANCHPNT", vBRANCHPNT}, 'vFIRSTCHK)
            {"vSCORE4_1_2", vSCORE4_1_2}, '總分
            {"vRLEVEL_2", vddlRLEVEL_2R}, '複審等級
            {"vSUBTOTAL", vSUBTOTAL},
            {"vMINISTERADD", vMINISTERADD},
            {"vMINISTERSUB", vMINISTERSUB},
            {"vMINISTERLEVEL", vMINISTERLEVEL},
            {"vDEPTADD", vDEPTADD},
            {"vOSID2", vOSID2}
        }
        Return True
    End Function

    ''' <summary>單一資料儲存</summary>
    ''' <param name="htSS"></param>
    Sub SSaveData2(ByRef htSS As Hashtable)
        'Dim vFIRSTCHK As String=TIMS.GetMyValue2(htSS, "vFIRSTCHK")
        Dim vSECONDCHK As String = TIMS.GetMyValue2(htSS, "vSECONDCHK") 'SECONDCHK'審核結果
        Dim vBRANCHPNT As String = TIMS.GetMyValue2(htSS, "vBRANCHPNT") '配合本部、本署辦理相關活動 或政策宣導 (7%)
        'Dim vSUBTOTAL As String=TIMS.GetMyValue2(htSS, "vSUBTOTAL") '小計
        Dim vSCORE4_1_2 As String = TIMS.GetMyValue2(htSS, "vSCORE4_1_2") '總分
        Dim vRLEVEL_2 As String = TIMS.GetMyValue2(htSS, "vRLEVEL_2") '複審等級

        Dim vSUBTOTAL As String = TIMS.GetMyValue2(htSS, "vSUBTOTAL")
        Dim vMINISTERADD As String = TIMS.GetMyValue2(htSS, "vMINISTERADD")
        Dim vMINISTERSUB As String = TIMS.GetMyValue2(htSS, "vMINISTERSUB")
        Dim vMINISTERLEVEL As String = TIMS.GetMyValue2(htSS, "vMINISTERLEVEL") '部加分等級
        Dim vDEPTADD As String = TIMS.GetMyValue2(htSS, "vDEPTADD")
        Dim vOSID2 As String = TIMS.GetMyValue2(htSS, "vOSID2")
        If vOSID2 = "" Then Return

        'parms.Add("BRANCHPNT", vBRANCHPNT)
        'parms.Add("SCORE4_1", If(SCORE4_1.Text <> "", Val(SCORE4_1.Text), Convert.DBNull)) '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
        'parms.Add("SUBTOTAL", If(vSUBTOTAL <> "", Val(vSUBTOTAL), Convert.DBNull)) '小計
        Dim parms As New Hashtable From {
            {"SECONDCHK", If(vSECONDCHK <> "", vSECONDCHK, Convert.DBNull)},
            {"SECONACCT", sm.UserInfo.UserID},
            {"BRANCHPNT", If(vBRANCHPNT <> "", Val(vBRANCHPNT), Convert.DBNull)}, '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
            {"MINISTERADD", If(vMINISTERADD <> "", Val(vMINISTERADD), 0)},
            {"MINISTERSUB", If(vMINISTERSUB <> "", Val(vMINISTERSUB), Val(vSUBTOTAL))},
            {"MINISTERLEVEL", If(vMINISTERLEVEL <> "", vMINISTERLEVEL, Convert.DBNull)}, '部加分等級
            {"DEPTADD", If(vDEPTADD <> "", Val(vDEPTADD), 0)},
            {"SCORE4_1_2", If(vSCORE4_1_2 <> "", Val(vSCORE4_1_2), Convert.DBNull)}, '總分
            {"RLEVEL_2", If(vRLEVEL_2 <> "", vRLEVEL_2, Convert.DBNull)},
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"OSID2", Val(vOSID2)}
        }

        'sql &= " SET FIRSTCHK=@FIRSTCHK ,FIRSTACCT=@FIRSTACCT ,FIRSTDATE=GETDATE()" & vbCrLf
        Dim u_sql As String = "" 'updata
        u_sql &= " UPDATE ORG_SCORING2" & vbCrLf
        'SECONDCHK'審核結果
        u_sql &= " SET SECONDCHK=@SECONDCHK,SECONACCT=@SECONACCT,SECONDATE=GETDATE()" & vbCrLf
        'u_sql &= " ,SCORE4_1=@SCORE4_1" & vbCrLf '分署 加分項目 加分項目 配合分署辦理相關活動或政策宣導(3%)
        u_sql &= " ,BRANCHPNT=@BRANCHPNT" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
        '部長加分,部長加分小計,部長加分等級,署加分,
        u_sql &= " ,MINISTERADD=@MINISTERADD,MINISTERSUB=@MINISTERSUB,MINISTERLEVEL=@MINISTERLEVEL,DEPTADD=@DEPTADD" & vbCrLf
        'u_sql &= " ,SUBTOTAL=@SUBTOTAL" & vbCrLf '小計
        u_sql &= " ,SCORE4_1_2=@SCORE4_1_2" & vbCrLf '總分
        u_sql &= " ,RLEVEL_2=@RLEVEL_2" & vbCrLf '複審等級
        u_sql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE OSID2=@OSID2" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, parms)

        divEdt1.Visible = False
        divSch1.Visible = True
        sm.LastResultMessage = "儲存完畢"
    End Sub

    Sub SSearch1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        PageControler1.Visible = False
        DataGridTable.Visible = False
        msg1.Text = "查無資料"
        BtnSaveData1.Visible = False

        Dim vDISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        'Dim vYEARS As String=TIMS.ClearSQM(SYEARlist.SelectedValue)
        'Dim vHALFYEAR As String=TIMS.ClearSQM(halfYear.SelectedValue) '1:上年度 /2:下年度

        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        Dim vORGKIND2 As String = TIMS.ClearSQM(OrgPlanKind.SelectedValue) '計畫
        Dim vORGKIND As String = TIMS.ClearSQM(OrgKindList.SelectedValue) '機構別

        'Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        Dim fg_CanNoDistSch As Boolean = TIMS.ChkUserLID(sm, 0) '(署)不卡控分署欄位為必選
        Dim eErrMsg1 As String = ""
        If Not fg_CanNoDistSch AndAlso vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf

        'If vYEARS="" Then eErrMsg1 &= "請選擇年度" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        Call KeepSearch1()

        'parms.Add("YEARS", sm.UserInfo.Years)
        Dim parms As New Hashtable From {
            {"EXP", "S"}, '查詢多筆資料 DT
            {"TPLANID", sm.UserInfo.TPlanID},
            {"SCORINGID", vSCORINGID}
        }
        If vDISTID <> "" Then parms.Add("DISTID", vDISTID) 'sql &= " AND t.DISTID=@DISTID" & vbCrLf
        'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf

        Dim dt As DataTable = Get_dtORGSCORING2(parms)
        'PageControler1.Visible=False 'DataGridTable.Visible=False 'msg1.Text="查無資料"
        If TIMS.dtNODATA(dt) Then Return

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""
        BtnSaveData1.Visible = True

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Function GET_CAL_1(ByVal oSUM1 As Object, ByVal oCLSAPPCNT As Object) As String
        Dim Rst1 As String = ""
        Dim flagCanCal As Boolean = False 'flagCanCal=False
        If Convert.ToString(oSUM1) <> "" AndAlso Convert.ToString(oCLSAPPCNT) <> "" Then
            If Val(oCLSAPPCNT) > 0 Then flagCanCal = True
        End If
        If flagCanCal Then
            Dim vEQU As Double = TIMS.ROUND(CDbl(oSUM1) / CDbl(oCLSAPPCNT) * 100, 1)
            Rst1 = CStr(vEQU)
        End If
        Return Rst1
    End Function

    Sub SClearlist1()
        Hid_OSID2.Value = ""
        LabOrgName.Text = "" 'Convert.ToString(dr("OrgName"))
        labSCORING_N.Text = "" '審查計分區間
        LabDISTNAME.Text = "" 'Convert.ToString(dr("DISTNAME"))
        'ddlFIRSTCHK_1.Text=Convert.ToString(dr("OrgName"))
        'Common.SetListItem(ddlFIRSTCHK_1, Convert.ToString(dr("FIRSTCHK")))

        'ddlFIRSTCHK_1.SelectedIndex=-1
        'Common.SetListItem(ddlFIRSTCHK_1, "")
        'ddlSECONDCHK_1.SelectedIndex=-1
        'Common.SetListItem(ddlSECONDCHK_1, "")

        CLSACTCNT.Text = "" ' Convert.ToString(dr("CLSACTCNT"))
        CLSACTCNT2.Text = ""
        CLSAPPCNT.Text = "" ' Convert.ToString(dr("CLSAPPCNT")) '核定班數
        CLSAPPCNT_t2.Text = "" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t3.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t4.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t5.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t6.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t7.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t8.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t9.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t10.Text="" ' Convert.ToString(dr("CLSAPPCNT"))
        CLSAPPCNT_t11.Text = "" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t12.Text="" 'Convert.ToString(dr("CLSAPPCNT"))

        SCORE1_1A.Text = "" 'Convert.ToString(dr("SCORE1_1A"))
        SCORE1_1.Text = "" 'Convert.ToString(dr("SCORE1_1"))
        STDACTCNT.Text = "" 'Convert.ToString(dr("STDACTCNT"))
        STDACTCNT2.Text = ""
        STDAPPCNT.Text = "" 'Convert.ToString(dr("STDAPPCNT"))
        SCORE1_2A.Text = "" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE1_2.Text = "" 'Convert.ToString(dr("SCORE1_2A"))

        SCORE2_1_1_SUM_A.Text = "" 'Convert.ToString(dr("SCORE2_1_1_SUM_A"))
        'SCORE2_1_1A.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_SUM_B.Text = "" 'Convert.ToString(dr("SCORE2_1_1_SUM_B"))
        'SCORE2_1_1B.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_SUM_C.Text = "" 'Convert.ToString(dr("SCORE2_1_1_SUM_C"))
        'SCORE2_1_1C.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_SUM_D.Text = "" 'Convert.ToString(dr("SCORE2_1_1_SUM_D"))
        'SCORE2_1_1D.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_ALL.Text = ""

        SCORE2_1_2A_DIS.Text = ""
        SCORE2_1_2B_DIS.Text = ""
        SCORE2_1_2C_DIS.Text = ""
        SCORE2_1_2D_DIS.Text = ""
        SCORE2_1_2_SUM_ALL.Text = ""

        SCORE2_1_3_SUM.Text = "" 'Convert.ToString(dr("SCORE2_1_3_SUM"))
        'SCORE2_1_3_EQU.Text="" 'Convert.ToString(dr("SCORE2_1_3_EQU"))
        SCORE2_2_1_SUM.Text = "" 'Convert.ToString(dr("SCORE2_2_1_SUM"))
        'SCORE2_2_1_EQU.Text="" 'Convert.ToString(dr("SCORE2_2_1_EQU"))
        SCORE2_2_1.Text = "" 'Convert.ToString(dr("SCORE2_2_1"))

        'SCORE2_2_2_SUM.Text="" 'Convert.ToString(dr("SCORE2_2_2_SUM"))
        'SCORE2_2_2_EQU.Text="" 'Convert.ToString(dr("SCORE2_2_2_EQU"))
        SCORE2_2_2_DIS.Text = ""
        SCORE2_2_2.Text = "" 'Convert.ToString(dr("SCORE2_2_2"))

        SCORE2_3_1_SUM.Text = "" 'Convert.ToString(dr("SCORE2_3_1_SUM"))
        SCORE2_3_1_CNT.Text = "" 'Convert.ToString(dr("SCORE2_3_1_CNT"))
        SCORE2_3_1_EQU.Text = "" 'Convert.ToString(dr("SCORE2_3_1_EQU"))
        SCORE2_3_1.Text = "" 'Convert.ToString(dr("SCORE2_3_1"))

        SCORE3_1_N.Text = "" 'Convert.ToString(dr("SCORE3_1_N"))
        SCORE3_1.Text = "" 'Convert.ToString(dr("SCORE3_1"))
        SCORE3_2_SUM.Text = "" ' Convert.ToString(dr("SCORE3_2_SUM"))
        SCORE3_2_CNT.Text = "" 'Convert.ToString(dr("SCORE3_2_CNT"))
        SCORE3_2_EQU.Text = "" 'Convert.ToString(dr("SCORE3_2_EQU"))
        SCORE3_2.Text = "" 'Convert.ToString(dr("SCORE3_2"))

        'SCORE4_1_A.Text="" 'Convert.ToString(dr("SCORE4_1_A")) '配合活動得分
        SCORE4_1.Text = "" 'Convert.ToString(dr("SCORE4_1"))'分署 加分項目
        'Sql &= " ,a.SCORE4_2A" & vbCrLf
        'Sql &= " ,a.SCORE4_2_CNT" & vbCrLf
        'Sql &= " ,a.SCORE4_2_RATE" & vbCrLf
        'Sql &= " ,a.SCORE4_2" & vbCrLf
        SCORE4_2A.Text = ""
        SCORE4_2_CNT.Text = ""
        SCORE4_2_RATE.Text = ""
        SCORE4_2.Text = "" 'Convert.ToString(dr("SCORE4_2_RATE")) '參訓學員平均填答率

        SUBTOTAL.Text = "" 'Convert.ToString(dr("SUBTOTAL"))
        BRANCHPNT.Text = "" 'Convert.ToString(dr("BRANCHPNT"))
        '部長加分,部長加分小計,部長加分等級,署加分,
        MINISTERADD.Text = "" ' TIMS.ClearSQM(MINISTERADD.Text)
        DEPTADD.Text = "" 'TIMS.ClearSQM(DEPTADD.Text)
        MINISTERSUB.Text = "" 'TIMS.ClearSQM(MINISTERSUB.Text)
        Hid_MINISTERLEVEL.Value = "" 'TIMS.ClearSQM(Hid_MINISTERLEVEL.Value)
        ddl_MINISTERLEVEL.SelectedIndex = -1 'Dim vddl_MINISTERLEVEL As String = TIMS.GetListValue(ddl_MINISTERLEVEL)
        Common.SetListItem(ddl_MINISTERLEVEL, "")

        SCORE4_1_2.Text = "" 'Convert.ToString(dr("SCORE4_1_2"))
        ddlSECONDCHK_1.SelectedIndex = -1
        Common.SetListItem(ddlSECONDCHK_1, "")
    End Sub

    ''' <summary>
    ''' 鎖定輸入格-true:Lock
    ''' </summary>
    ''' <param name="bLock"></param>
    Sub Utl_LockData1(ByRef bLock As Boolean)
        Dim flag_Enabled_1 As Boolean = If(bLock, False, True)

        CLSACTCNT.Enabled = flag_Enabled_1  '.Text="" ' Convert.ToString(dr("CLSACTCNT"))
        CLSACTCNT2.Enabled = flag_Enabled_1
        CLSAPPCNT.Enabled = flag_Enabled_1  '.Text="" ' Convert.ToString(dr("CLSAPPCNT")) '核定班數
        CLSAPPCNT_t2.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t3.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t4.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t5.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t6.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t7.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t8.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t9.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        CLSAPPCNT_t10.Enabled = flag_Enabled_1
        'CLSAPPCNT_t10.Enabled=False'.Text="" ' Convert.ToString(dr("CLSAPPCNT"))
        CLSAPPCNT_t11.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t12.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))

        SCORE1_1A.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_1A"))
        SCORE1_1.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_1"))
        STDACTCNT.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("STDACTCNT"))
        STDACTCNT2.Enabled = flag_Enabled_1
        STDAPPCNT.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("STDAPPCNT"))
        SCORE1_2A.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE1_2.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_2A"))

        SCORE2_1_1_SUM_A.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_1_1_SUM_A"))
        'SCORE2_1_1A.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_SUM_B.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_1_1_SUM_B"))
        'SCORE2_1_1B.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_SUM_C.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_1_1_SUM_C"))
        'SCORE2_1_1C.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_SUM_D.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_1_1_SUM_D"))
        'SCORE2_1_1D.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_2A"))
        SCORE2_1_1_ALL.Enabled = flag_Enabled_1

        SCORE2_1_2A_DIS.Enabled = flag_Enabled_1
        SCORE2_1_2B_DIS.Enabled = flag_Enabled_1
        SCORE2_1_2C_DIS.Enabled = flag_Enabled_1
        SCORE2_1_2D_DIS.Enabled = flag_Enabled_1
        SCORE2_1_2_SUM_ALL.Enabled = flag_Enabled_1

        SCORE2_1_3_SUM.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_1_3_SUM"))
        'SCORE2_1_3A.Enabled=flag_Enabled_1 'SCORE2_1_3A-核定總班數
        'SCORE2_1_3_EQU.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_1_3_EQU"))
        SCORE2_1_3.Enabled = flag_Enabled_1 '得分
        SCORE2_2_1_SUM.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_2_1_SUM"))
        'SCORE2_2_1_EQU.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_2_1_EQU"))
        SCORE2_2_1.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_2_1"))

        'SCORE2_2_2_SUM.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_2_2_SUM"))
        'SCORE2_2_2_EQU.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_2_2_EQU"))
        SCORE2_2_2_DIS.Enabled = flag_Enabled_1
        SCORE2_2_2.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_2_2"))

        SCORE2_3_1_SUM.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_3_1_SUM"))
        SCORE2_3_1_CNT.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_3_1_CNT"))
        SCORE2_3_1_EQU.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_3_1_EQU"))
        SCORE2_3_1.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE2_3_1"))

        SCORE3_1_N.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE3_1_N"))
        SCORE3_1.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE3_1"))
        SCORE3_2_SUM.Enabled = flag_Enabled_1  '.Text="" ' Convert.ToString(dr("SCORE3_2_SUM"))
        SCORE3_2_CNT.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE3_2_CNT"))
        SCORE3_2_EQU.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE3_2_EQU"))
        SCORE3_2.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE3_2"))

        SCORE4_2A.Enabled = flag_Enabled_1
        SCORE4_2_CNT.Enabled = flag_Enabled_1
        SCORE4_2_RATE.Enabled = flag_Enabled_1 '參訓學員平均填答率
        SCORE4_2.Enabled = flag_Enabled_1

        SCORE4_1.Enabled = flag_Enabled_1

        SUBTOTAL.Enabled = flag_Enabled_1 '分署小計
        '部長加分,部長加分小計,部長加分等級,署加分,
        'BRANCHPNT.Enabled = flag_Enabled_1 '配合本部、本署辦理相關活動 或政策宣導 (7%)
        'MINISTERSUB.Enabled = flag_Enabled_1 '部加分小計
        TIMS.INPUT_ReadOnly(BRANCHPNT, "自動計算")
        TIMS.INPUT_ReadOnly(MINISTERSUB, "自動計算")
        'MINISTERADD.Enabled = flag_Enabled_1 '部加分
        'DEPTADD.Enabled = flag_Enabled_1 '署加分
        'ddl_MINISTERLEVEL.Enabled = flag_Enabled_1 '部加分等級

        'Hid_MINISTERLEVEL.Value = TIMS.ClearSQM(Hid_MINISTERLEVEL.Value)
        'Dim vMINISTERLEVEL As String = TIMS.GetListValue(ddl_MINISTERLEVEL)
        'SCORE4_1_2.Enabled = flag_Enabled_1 '總分
        'ddlRLEVEL_2R.Enabled = flag_Enabled_1 '複審等級

        'SCORE4_1_A.Text="" 'Convert.ToString(dr("SCORE4_1_A")) '配合活動得分
        'SCORE4_1.Text="" 'Convert.ToString(dr("SCORE4_1"))
    End Sub

    Sub SShowData1(ByRef dr As DataRow)
        If dr Is Nothing Then Exit Sub
        'ORG_SCORING
        divSch1.Visible = False
        divEdt1.Visible = True
        Hid_OSID2.Value = Convert.ToString(dr("OSID2"))
        If Hid_OSID2.Value = "" Then Exit Sub

        Dim iCLSBEDCNT As Integer = If(Convert.ToString(dr("CLSBEDCNT")) <> "", Val(dr("CLSBEDCNT")), 0)
        tr_Lab_SUSPENDED_msg1.Visible = (iCLSBEDCNT > 0)
        Dim str_SUSPENDED_msg1 As String = String.Format(cst_SUSPENDED_msgFM1, iCLSBEDCNT)
        Lab_SUSPENDED_msg1.Text = If(iCLSBEDCNT > 0, str_SUSPENDED_msg1, "")
        '【初審審核】隱藏 (因為也只有通過的才會進到複審)'審查計分表(初審)
        'Common.SetListItem(ddlFIRSTCHK_1, Convert.ToString(dr("FIRSTCHK")))
        'tIMPSCORE_1.Text=Convert.ToString(dr("IMPSCORE_1"))
        Hid_IMPLEVEL_1.Value = Convert.ToString(dr("IMPLEVEL_1"))
        Common.SetListItem(ddlIMPLEVEL_1, Hid_IMPLEVEL_1.Value) '初審等級
        '【初審審核】隱藏 (因為也只有通過的才會進到複審)
        'ddlFIRSTCHK_1.Enabled=False
        'tIMPSCORE_1.Enabled=False
        'Hid_IMPLEVEL_1.Enabled=False
        ddlIMPLEVEL_1.Enabled = False
        '【初審審核】隱藏 (因為也只有通過的才會進到複審)
        'TIMS.Tooltip(ddlFIRSTCHK_1, "僅供顯示", True)
        'TIMS.Tooltip(tIMPSCORE_1, "僅供顯示", True)
        TIMS.Tooltip(ddlIMPLEVEL_1, "僅供顯示", True) '初審等級

        LabOrgName.Text = Convert.ToString(dr("OrgName"))
        labSCORING_N.Text = Convert.ToString(dr("SCORING_N")) '審查計分區間
        LabDISTNAME.Text = Convert.ToString(dr("DISTNAME"))
        'ddlFIRSTCHK_1.Text=Convert.ToString(dr("OrgName"))
        'Common.SetListItem(ddlFIRSTCHK_1, Convert.ToString(dr("FIRSTCHK")))
        'Common.SetListItem(ddlSECONDCHK_1, Convert.ToString(dr("SECONDCHK")))

        CLSACTCNT.Text = Convert.ToString(dr("CLSACTCNT")) '實際開班數 (核定-停辦-政策性)(核定課程數-停辦課程數)
        CLSACTCNT2.Text = Convert.ToString(dr("CLSACTCNT2")) '政策型課程班  (政策性(含停辦))

        CLSAPPCNT.Text = Convert.ToString(dr("CLSAPPCNT")) '核定班數
        CLSAPPCNT_t2.Text = Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t3.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t4.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t5.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t6.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t7.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t8.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t9.Text=Convert.ToString(dr("CLSAPPCNT"))
        CLSAPPCNT_t10.Text = Convert.ToString(dr("CLSAPPCNT")) 'SCORE2_1_3A
        CLSAPPCNT_t11.Text = Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t12.Text=Convert.ToString(dr("CLSAPPCNT"))

        SCORE1_1A.Text = Convert.ToString(dr("SCORE1_1A"))
        SCORE1_1.Text = Convert.ToString(dr("SCORE1_1"))
        STDACTCNT.Text = Convert.ToString(dr("STDACTCNT")) '實際開訓人次(核定-停辦-政策性)
        STDACTCNT2.Text = Convert.ToString(dr("STDACTCNT2")) '政策性課程核定人次(政策性)
        STDAPPCNT.Text = Convert.ToString(dr("STDAPPCNT")) '核定總人次
        SCORE1_2A.Text = Convert.ToString(dr("SCORE1_2A"))
        SCORE1_2.Text = Convert.ToString(dr("SCORE1_2"))

        SCORE2_1_1_SUM_A.Text = Convert.ToString(dr("SCORE2_1_1_SUM_A"))
        'SCORE2_1_1A.Text=Convert.ToString(dr("SCORE2_1_1A"))
        SCORE2_1_1_SUM_B.Text = Convert.ToString(dr("SCORE2_1_1_SUM_B"))
        'SCORE2_1_1B.Text=Convert.ToString(dr("SCORE2_1_1B"))
        SCORE2_1_1_SUM_C.Text = Convert.ToString(dr("SCORE2_1_1_SUM_C"))
        'SCORE2_1_1C.Text=Convert.ToString(dr("SCORE2_1_1C"))
        SCORE2_1_1_SUM_D.Text = Convert.ToString(dr("SCORE2_1_1_SUM_D"))
        'SCORE2_1_1D.Text=Convert.ToString(dr("SCORE2_1_1D"))
        SCORE2_1_1_ALL.Text = Convert.ToString(If(Convert.ToString(dr("SCORE2_1_1_ALL")) <> "", dr("SCORE2_1_1_ALL"), 0))

        SCORE2_1_2A_DIS.Text = Convert.ToString(dr("SCORE2_1_2A_DIS"))
        SCORE2_1_2B_DIS.Text = Convert.ToString(dr("SCORE2_1_2B_DIS"))
        SCORE2_1_2C_DIS.Text = Convert.ToString(dr("SCORE2_1_2C_DIS"))
        SCORE2_1_2D_DIS.Text = Convert.ToString(dr("SCORE2_1_2D_DIS"))
        SCORE2_1_2_SUM_ALL.Text = Convert.ToString(dr("SCORE2_1_2_SUM_ALL"))

        SCORE2_1_3_SUM.Text = Convert.ToString(dr("SCORE2_1_3_SUM"))
        'SCORE2_1_3A.Text=Convert.ToString(dr("SCORE2_1_3A")) 'SCORE2_1_3A-核定總班數
        'SCORE2_1_3_EQU.Text=GET_CAL_1(dr("SCORE2_1_3_SUM"), dr("SCORE2_1_3A"))
        SCORE2_1_3.Text = Convert.ToString(dr("SCORE2_1_3"))

        SCORE2_2_1_SUM.Text = Convert.ToString(dr("SCORE2_2_1_SUM"))
        'SCORE2_2_1_EQU.Text=GET_CAL_1(dr("SCORE2_2_1_SUM"), dr("CLSAPPCNT")) 'Convert.ToString(dr("SCORE2_2_1_EQU"))
        SCORE2_2_1.Text = Convert.ToString(dr("SCORE2_2_1"))

        'SCORE2_2_2_SUM.Text=Convert.ToString(dr("SCORE2_2_2_SUM"))
        'SCORE2_2_2_EQU.Text=GET_CAL_1(dr("SCORE2_2_2_SUM"), dr("CLSAPPCNT")) 'Convert.ToString(dr("SCORE2_2_2_EQU"))
        SCORE2_2_2_DIS.Text = Convert.ToString(dr("SCORE2_2_2_DIS"))
        SCORE2_2_2.Text = Convert.ToString(dr("SCORE2_2_2"))

        SCORE2_3_1_SUM.Text = Convert.ToString(dr("SCORE2_3_1_SUM"))
        SCORE2_3_1_CNT.Text = Convert.ToString(dr("SCORE2_3_1_CNT"))
        SCORE2_3_1_EQU.Text = GET_CAL_1(dr("SCORE2_3_1_SUM"), dr("SCORE2_3_1_CNT")) 'Convert.ToString(dr("SCORE2_3_1_EQU"))
        SCORE2_3_1.Text = Convert.ToString(dr("SCORE2_3_1"))

        'SELECT concat('''',vid,'.',vname) FROM V_RESULT
        '1.金牌'2.銀牌'3.銅牌'4.通過'5.未通過'6.合格'7.不合格'8.通過門檻'9.未通過門檻
        'Dim vSCORE3_1_N As String=""
        'Select Case Convert.ToString(dr("RESULT_N"))
        '    Case "1"
        '        vSCORE3_1_N="金"
        '    Case "2"
        '        vSCORE3_1_N="銀"
        '    Case "3"
        '        vSCORE3_1_N="銅"
        'End Select
        SCORE3_1_N.Text = Convert.ToString(dr("RESULT_N")) 'vSCORE3_1_N
        'SCORE3_1_N.Text=Convert.ToString(dr("SCORE3_1_N"))
        SCORE3_1.Text = Convert.ToString(dr("SCORE3_1"))

        SCORE3_2_SUM.Text = Convert.ToString(dr("SCORE3_2_SUM"))
        SCORE3_2_CNT.Text = Convert.ToString(dr("SCORE3_2_CNT"))

        Dim vSCORE3_2_CNT As Integer = If(Convert.ToString(dr("SCORE3_2_CNT")) <> "", Val(dr("SCORE3_2_CNT")), 0)
        SCORE3_2_EQU.Text = GET_CAL_1(dr("SCORE3_2_SUM"), (vSCORE3_2_CNT * 2)) 'Convert.ToString(dr("SCORE3_2_EQU"))
        SCORE3_2.Text = Convert.ToString(dr("SCORE3_2"))

        'SCORE4_1_A.Text=Convert.ToString(dr("SCORE4_1_A")) '配合活動得分
        SCORE4_1.Text = Convert.ToString(dr("SCORE4_1")) '分署 加分項目'配合活動得分
        'Sql &= " ,a.SCORE4_2A" & vbCrLf
        'Sql &= " ,a.SCORE4_2_CNT" & vbCrLf
        'Sql &= " ,a.SCORE4_2_RATE" & vbCrLf
        'Sql &= " ,a.SCORE4_2" & vbCrLf
        ' (1)加分項目分為分署、署
        '    分署項目標題調整為：4 加分項目(分署)(如圖一)
        '配合分署辦理相關活動或政策宣導(3%)
        '預設空白， 由分署自填， 至多加 3 分
        '參訓學員訓後動態調查表單位平均填答率達80%(2%) ： 目前都是空的， 應由系統計算
        '    計分公式： 參訓學員訓後動態調查表填寫人次 / 結訓學員總人次
        ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
        '    參訓學員訓後動態調查表填寫人次=該訓練單位之所有開訓課程有填寫參訓學員訓後動態調查表之人次總計
        '結訓學員總人次=該訓練單位之所有開訓課程的結訓人次總計
        SCORE4_2A.Text = Convert.ToString(dr("SCORE4_2A")) ' 參訓學員訓後動態調查表填寫人次
        SCORE4_2_CNT.Text = Convert.ToString(dr("SCORE4_2_CNT")) '結訓學員總人次
        SCORE4_2_RATE.Text = Convert.ToString(dr("SCORE4_2_RATE")) '參訓學員平均填答率
        SCORE4_2.Text = Convert.ToString(dr("SCORE4_2")) ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
        '分署小計/'小計
        SUBTOTAL.Text = Convert.ToString(dr("SUBTOTAL"))
        If (SUBTOTAL.Text = "") Then SUBTOTAL.Text = "0"
        Dim v_TOTALSCORE As String = If(Convert.ToString(dr("TOTALSCORE")) = "-", "", Convert.ToString(dr("TOTALSCORE")))
        Hid_TOTALSCORE.Value = v_TOTALSCORE
        '5 加分項目(署) >配合本部、本署辦理相關活動 或政策宣導 (7%)
        BRANCHPNT.Text = Convert.ToString(dr("BRANCHPNT"))
        If (BRANCHPNT.Text = "") Then BRANCHPNT.Text = "0"
        '部長加分,部長加分小計,部長加分等級,署加分,
        MINISTERADD.Text = Convert.ToString(dr("MINISTERADD")) ' TIMS.ClearSQM(MINISTERADD.Text)
        If (MINISTERADD.Text = "") Then MINISTERADD.Text = "0"
        DEPTADD.Text = Convert.ToString(dr("DEPTADD")) 'TIMS.ClearSQM(DEPTADD.Text)
        If (DEPTADD.Text = "") Then DEPTADD.Text = "0"

        MINISTERSUB.Text = Convert.ToString(dr("MINISTERSUB")) 'TIMS.ClearSQM(MINISTERSUB.Text)
        If (MINISTERSUB.Text = "") Then MINISTERSUB.Text = TIMS.VAL1(SUBTOTAL.Text) + TIMS.VAL1(MINISTERADD.Text)

        Hid_MINISTERLEVEL.Value = Convert.ToString(dr("MINISTERLEVEL")) 'TIMS.ClearSQM(Hid_MINISTERLEVEL.Value)
        Common.SetListItem(ddl_MINISTERLEVEL, Hid_MINISTERLEVEL.Value)
        '5 加分項目(署) >總分
        SCORE4_1_2.Text = Convert.ToString(dr("SCORE4_1_2"))
        If SCORE4_1_2.Text = "" Then SCORE4_1_2.Text = v_TOTALSCORE
        '(至多100)
        If TIMS.VAL1(SCORE4_1_2.Text) > 100 Then SCORE4_1_2.Text = "100"

        '5 加分項目(署) >複審等級
        Common.SetListItem(ddlRLEVEL_2R, Convert.ToString(dr("RLEVEL_2")))
        '5 加分項目(署) >審核結果
        Common.SetListItem(ddlSECONDCHK_1, Convert.ToString(dr("SECONDCHK")))
    End Sub

    ''' <summary>'載入資料動作1</summary>
    ''' <param name="sCmdArg"></param>
    Sub SLoadData1(ByRef sCmdArg As String)
        If sCmdArg = "" Then Exit Sub
        Dim OSID2 As String = TIMS.GetMyValue(sCmdArg, "OSID2")
        If OSID2 = "" Then Exit Sub

        '一筆資料
        Dim parms As New Hashtable From {{"EXP", "1"}, {"OSID2", OSID2}}
        Dim dt As DataTable = Get_dtORGSCORING2(parms)

        divSch1.Visible = True
        divEdt1.Visible = False
        If TIMS.dtNODATA(dt) Then
            sm.LastErrorMessage = "查無資料"
            Exit Sub
        End If

        divSch1.Visible = False
        divEdt1.Visible = True
        Dim dr1 As DataRow = dt.Rows(0)
        Call SClearlist1()

        Call SShowData1(dr1)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        'Dim RTSID As String=TIMS.GetMyValue(sCmdArg, "RTSID")
        'Dim ORGID As String=TIMS.GetMyValue(sCmdArg, "ORGID")
        Select Case e.CommandName
            Case "btnView"
                Dim OSID2 As String = TIMS.GetMyValue(sCmdArg, "OSID2")
                If OSID2 = "" Then Exit Sub
                SLoadData1(sCmdArg)
                Utl_LockData1(True)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim lbtView As LinkButton = e.Item.FindControl("lbtView")
                Dim HidSUBTOTAL As HiddenField = e.Item.FindControl("HidSUBTOTAL")
                '部長加分,部長加分小計,部長加分等級,署加分,
                Dim tMINISTERADD As TextBox = e.Item.FindControl("tMINISTERADD") '部長加分
                Dim LabMINISTERSUB As Label = e.Item.FindControl("LabMINISTERSUB") '部長加分小計
                Dim ddlMINISTERLEVEL As DropDownList = e.Item.FindControl("ddlMINISTERLEVEL") '部長加分等級
                Dim tDEPTADD As TextBox = e.Item.FindControl("tDEPTADD") '署加分
                'Dim Hid_BRANCHPNTorg As HiddenField = e.Item.FindControl("Hid_BRANCHPNTorg")
                'Dim tBRANCHPNT As TextBox = e.Item.FindControl("tBRANCHPNT") '分署<br>加分項目
                'Dim LSUBTOTAL As Label=e.Item.FindControl("LSUBTOTAL") '小計
                'Dim lRlevel_1 As Label=e.Item.FindControl("lRlevel_1") '初審<br>等級"
                'Dim lRlevel_2 As Label=e.Item.FindControl("lRlevel_2") '複審等級/複審<br>等級"
                'lRlevel_2.Text="-"  '複審等級/複審<br>等級" 'ddlRLEVEL_2 /複審等級
                Dim Hid_RLEVEL_2 As HiddenField = e.Item.FindControl("Hid_RLEVEL_2")
                Dim LabTOTALSCORE As Label = e.Item.FindControl("LabTOTALSCORE")
                Dim ddlRLEVEL_2 As DropDownList = e.Item.FindControl("ddlRLEVEL_2")

                HidSUBTOTAL.Value = TIMS.ClearSQM(drv("SUBTOTAL"))
                ddlMINISTERLEVEL = TIMS.Get_SCORELEVEL(ddlMINISTERLEVEL)
                '部長加分,部長加分小計,部長加分等級,署加分,
                tMINISTERADD.Text = Convert.ToString(drv("MINISTERADD")) '部長加分
                If tMINISTERADD.Text = "" Then tMINISTERADD.Text = "0"
                tDEPTADD.Text = Convert.ToString(drv("DEPTADD")) '署加分
                If tDEPTADD.Text = "" Then tDEPTADD.Text = "0"
                Dim iSUBTOTAL As Double = TIMS.VAL1(drv("SUBTOTAL"))

                '部長加分
                Dim jjadd1 As String = String.Concat("changeSUB1(", iSUBTOTAL, ",$(this).val(),'", LabMINISTERSUB.ClientID, "','", tMINISTERADD.ClientID, "','", tDEPTADD.ClientID, "','", LabTOTALSCORE.ClientID, "',", TIMS.VAL1(tMINISTERADD.Text), ");")
                tMINISTERADD.Attributes.Add("onchange", jjadd1)
                tMINISTERADD.Attributes.Add("onblur", jjadd1)
                tMINISTERADD.Attributes.Add("onclick", jjadd1)

                '署加分
                Dim jjadd2 As String = String.Concat("changeSUB2(", iSUBTOTAL, ",$(this).val(),'", LabTOTALSCORE.ClientID, "','", tDEPTADD.ClientID, "','", tMINISTERADD.ClientID, "',", TIMS.VAL1(tDEPTADD.Text), ");")
                tDEPTADD.Attributes.Add("onchange", jjadd2)
                tDEPTADD.Attributes.Add("onblur", jjadd2)
                tDEPTADD.Attributes.Add("onclick", jjadd2)

                If Convert.ToString(drv("MINISTERSUB")) <> "" Then
                    LabMINISTERSUB.Text = Convert.ToString(drv("MINISTERSUB"))
                ElseIf Convert.ToString(drv("MINISTERADD")) <> "" Then
                    LabMINISTERSUB.Text = (iSUBTOTAL + TIMS.VAL1(drv("MINISTERADD")))
                Else
                    LabMINISTERSUB.Text = iSUBTOTAL
                End If
                'If Convert.ToString(drv("MINISTERLEVEL")) <> "" Then
                '    Common.SetListItem(ddlMINISTERLEVEL, Convert.ToString(drv("MINISTERLEVEL")))
                'ElseIf Convert.ToString(drv("IMPLEVEL_1")) <> "" Then
                '    Common.SetListItem(ddlMINISTERLEVEL, Convert.ToString(drv("IMPLEVEL_1")))
                'ElseIf Convert.ToString(drv("RLEVEL_2")) <> "" Then
                '    Common.SetListItem(ddlMINISTERLEVEL, Convert.ToString(drv("RLEVEL_2")))
                'End If
                Common.SetListItem(ddlMINISTERLEVEL, Convert.ToString(drv("MINISTERLEVEL")))

                LabTOTALSCORE.Text = Convert.ToString(drv("TOTALSCORE"))
                ddlRLEVEL_2 = TIMS.Get_SCORELEVEL(ddlRLEVEL_2)
                Hid_RLEVEL_2.Value = Convert.ToString(drv("RLEVEL_2"))
                'If Hid_RLEVEL_2.Value="" Then Hid_RLEVEL_2.Value=Convert.ToString(drv("IMPLEVEL_1"))
                Common.SetListItem(ddlRLEVEL_2, Convert.ToString(drv("RLEVEL_2")))

                Dim HidOSID2 As HiddenField = e.Item.FindControl("HidOSID2")
                Dim Hid_SECONDCHKorg As HiddenField = e.Item.FindControl("Hid_SECONDCHKorg")
                Dim ddlSECONDCHK As DropDownList = e.Item.FindControl("ddlSECONDCHK")
                'h_SUBTOTAL.Value=Convert.ToString(drv("SUBTOTAL")) 'h_TOTALSCORE.Value=Convert.ToString(drv("TOTALSCORE"))
                'tBRANCHPNT.Text = Convert.ToString(drv("BRANCHPNT")) '分署<br>加分項目
                'Hid_BRANCHPNTorg.Value = Convert.ToString(drv("BRANCHPNT")) 'tBRANCHPNT.Text
                'LSUBTOTAL.Text=Convert.ToString(drv("SUBTOTAL")) '小計
                If Convert.ToString(drv("SECONDCHK")) <> "" Then
                    Hid_SECONDCHKorg.Value = Convert.ToString(drv("SECONDCHK"))
                    Common.SetListItem(ddlSECONDCHK, Convert.ToString(drv("SECONDCHK")))
                End If

                HidOSID2.Value = Convert.ToString(drv("OSID2"))
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OSID2", Convert.ToString(drv("OSID2")))
                lbtView.CommandArgument = sCmdArg

                'Dim Hid_RTSID As HiddenField=e.Item.FindControl("Hid_RTSID")
                'Dim Hid_ORGID As HiddenField=e.Item.FindControl("Hid_ORGID")
                'Hid_RTSID.Value=Convert.ToString(drv("OTSID"))
                'Hid_ORGID.Value=Convert.ToString(drv("ORGID"))
                'TIMS.SetMyValue(sCmdArg, "RTSID", Convert.ToString(drv("RTSID")))
                'TIMS.SetMyValue(sCmdArg, "ORGID", Convert.ToString(drv("ORGID")))
        End Select
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Call SSaveData1()
    End Sub

#Region "SAMPLE NO USE"
    ''' <summary> 本功能為範例-暫無使用 </summary>
    ''' <param name="dt"></param>
    Sub ExpSampleXLS(ByRef dt As DataTable)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim strErrmsg As String = ""
        Dim sFileName As String = ""
        sFileName = ""
        sFileName &= "~\CO\01\Temp\"
        sFileName &= TIMS.GetDateNo()
        sFileName &= ".xls"
        Dim sMyFile1 As String = Server.MapPath(sFileName)

        Const cst_SampleXLS As String = "~\CO\01\SampleC.xls"
        'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), sMyFile1, True)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Exit Sub
        End Try
        If strErrmsg <> "" Then
            Try
                strErrmsg += "Path/File: " & sMyFile1 & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Exit Sub
        End If
        '除去sample檔的唯讀屬性
        'MyFile.SetAttributes(Server.MapPath("~\SD\03\Temp\" & Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "") & ".xls"), IO.FileAttributes.Normal)
        IO.File.SetAttributes(sMyFile1, IO.FileAttributes.Normal)
        'copy一份sample資料---End

        '根據路徑建立資料庫連線，並取出學員資料填入---------------   Start
        Using MyConn As New OleDb.OleDbConnection
            MyConn.ConnectionString = TIMS.Get_OleDbStr(sMyFile1)
            Try
                MyConn.Open()
            Catch ex As Exception
                Const cst_err_msg_1 As String = "Excel資料無法開啟連線!"
                strErrmsg = ""
                strErrmsg += cst_err_msg_1 & vbCrLf
                strErrmsg += "ex.ToString:" & ex.ToString & vbCrLf '取得錯誤資訊寫入
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
                Common.MessageBox(Me, cst_err_msg_1)
                Exit Sub
            End Try

            For Each drV As DataRow In dt.Rows
                Dim rCOMIDNO As String = Convert.ToString(drV("COMIDNO"))
                Dim rROWID As String = Convert.ToString(drV("ROWID"))
                Dim rOrgName As String = Convert.ToString(drV("OrgName"))
                Dim rORGKIND_N As String = Convert.ToString(drV("ORGKIND_N"))
                Dim rDISTNAME As String = Convert.ToString(drV("DISTNAME"))
                Dim rSCORE1_1 As String = Convert.ToString(drV("SCORE1_1"))
                Dim rSCORE1_2 As String = Convert.ToString(drV("SCORE1_2"))

                Dim rSCO2_1_1 As String = Convert.ToString(drV("SCORE2_1_1"))
                Dim rSCO2_1_2 As String = Convert.ToString(drV("SCORE2_1_2"))
                Dim rSCO2_1_3 As String = Convert.ToString(drV("SCORE2_1_3"))

                Dim rSCO2_2_1 As String = Convert.ToString(drV("SCORE2_2_1"))
                Dim rSCO2_2_2 As String = Convert.ToString(drV("SCORE2_2_2"))
                Dim rSCO2_3_1 As String = Convert.ToString(drV("SCORE2_3_1"))

                Dim rSCORE3_1 As String = Convert.ToString(drV("RESULT_N")) ' Convert.ToString(drV("SCORE3_1"))
                Dim rSCORE3_2 As String = Convert.ToString(drV("SCORE3_2"))

                Dim rSCORE4_1 As String = Convert.ToString(drV("SCORE4_1")) '分署 加分項目
                Dim rSCORE4_2 As String = Convert.ToString(drV("SCORE4_2"))

                Dim sql As String = ""
                sql &= " INSERT INTO [Sheet1$]([統一編號], [序號], [訓練單位名稱], [屬性], [分署]" & vbCrLf
                sql &= " ,[1-1],[1-2]" & vbCrLf
                sql &= " ,[2-1-1],[2-1-2],[2-1-3]" & vbCrLf
                sql &= " ,[2-2-1],[2-2-2],[2-3-1]" & vbCrLf
                sql &= " ,[3-1], [獎牌],[3-2]" & vbCrLf
                sql &= " ,[4-1],[4-2]" & vbCrLf
                sql &= " , [小計], [初擬等級],[4-1-2]" & vbCrLf
                sql &= " , [合計], [等級], [備註] )" & vbCrLf
                sql &= " VALUES ('" & rCOMIDNO & "','" & rROWID & "','" & rOrgName & "','" & rORGKIND_N & "','" & rDISTNAME & "'" & vbCrLf
                sql &= " ,'" & rSCORE1_1 & "','" & rSCORE1_2 & "'" & vbCrLf
                sql &= " ,'" & rSCO2_1_1 & "','" & rSCO2_1_2 & "','" & rSCO2_1_3 & "'" & vbCrLf
                sql &= " ,'" & rSCO2_2_1 & "','" & rSCO2_2_2 & "','" & rSCO2_3_1 & "'" & vbCrLf
                sql &= " ,'" & rSCORE3_1 & "','獎牌','" & rSCORE3_2 & "'" & vbCrLf
                sql &= " ,'" & rSCORE4_1 & "','" & rSCORE4_2 & "'" & vbCrLf
                sql &= " ,'[小計]','[初擬等級]','[4-1-2]'" & vbCrLf
                sql &= " ,'[合計]','[等級]','[備註]' )" & vbCrLf

                Using OleCmd1 As New OleDb.OleDbCommand(sql, MyConn)
                    Try
                        If MyConn.State = ConnectionState.Closed Then MyConn.Open()
                        OleCmd1.ExecuteNonQuery()  'edit，by:20181011
                        'If conn.State=ConnectionState.Open Then conn.Close()
                        'DbAccess.ExecuteNonQuery(sql)  'edit，by:20181011
                    Catch ex As Exception
                        If MyConn.State = ConnectionState.Open Then MyConn.Close()
                        strErrmsg = ""
                        strErrmsg += "程式錯誤!!!" & vbCrLf
                        strErrmsg += "sql:" & sql & vbCrLf '取得錯誤資訊寫入
                        strErrmsg += "ex.ToString:" & ex.ToString & vbCrLf '取得錯誤資訊寫入
                        strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                        'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                        Call TIMS.WriteTraceLog(strErrmsg)
                        Throw ex
                    End Try
                End Using

            Next
            If MyConn.State = ConnectionState.Open Then MyConn.Close()
            '根據路徑建立資料庫連線，並取出學員資料填入---------------   End
        End Using

        '將新建立的excel存入記憶體下載-----   Start
        Dim myFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng
        'Dim MyFileName As String=TIMS.r.ChangeIDNO(Replace(Replace(Replace(OCID1.Text, ")", ""), "(", ""), "/", "")) & ".xls"
        myFileName1 = TIMS.ClearSQM(myFileName1) & ".xls"
        strErrmsg = ""

        Try
            Dim fr As New System.IO.FileStream(sMyFile1, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte
            fr.Read(buf, 0, fr.Length)
            fr.Close()
            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(myFileName1, System.Text.Encoding.UTF8))
            Response.ContentType = "Application/vnd.ms-Excel"
            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
            Response.BinaryWrite(buf)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "無法存取該檔案!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
        Finally
            '刪除Temp中的資料
            Call TIMS.MyFileDelete(sMyFile1)
            If strErrmsg = "" Then
                Call TIMS.CloseDbConn(objconn)
                Response.End()
            End If
        End Try
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
        End If
        '將新建立的excel存入記憶體下載-----   End
    End Sub

#End Region

    ''' <summary>
    ''' 匯出審查計分表
    ''' </summary>
    Sub SExprot21()
        'Dim dtXls As DataTable=Nothing
        Dim vDISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        'Dim vYEARS As String=TIMS.ClearSQM(SYEARlist.SelectedValue)
        'Dim vHALFYEAR As String=TIMS.ClearSQM(halfYear.SelectedValue) '1:上年度 /2:下年度

        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        Dim vORGKIND2 As String = TIMS.ClearSQM(OrgPlanKind.SelectedValue) '計畫
        Dim vORGKIND As String = TIMS.ClearSQM(OrgKindList.SelectedValue) '機構別

        Dim eErrMsg1 As String = ""
        '署可不選分署
        If vDISTID = "" AndAlso sm.UserInfo.LID <> 0 Then
            eErrMsg1 &= "請選擇分署" & vbCrLf
        End If
        'If vYEARS="" Then eErrMsg1 &= "請選擇年度" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        '匯出查詢條件'parms.Clear() 'parms.Add("YEARS", sm.UserInfo.Years)
        Dim parms As New Hashtable From {{"EXP", "Y"}, {"TPLANID", sm.UserInfo.TPlanID}, {"SCORINGID", vSCORINGID}}
        'parms.Add("YEARS", vYEARS) 'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        '{"DISTID", vDISTID}, 'sql &= " AND t.DISTID=@DISTID" & vbCrLf
        If vDISTID <> "" Then parms.Add("DISTID", vDISTID)
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf

        'Dim dtXls As DataTable=Nothing
        Dim dtXls As DataTable = Get_dtORGSCORING2(parms)
        'If dtXls Is Nothing Then
        '    Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
        '    Return
        'End If
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If
        'Call ExpSampleXLS(dtXls)
        Call ExpXMLXLS(dtXls)
    End Sub

    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles BtnExp1.Click
        Call SExprot21()
    End Sub

    '回上頁
    Protected Sub BtnBack2_Click(sender As Object, e As EventArgs) Handles BtnBack2.Click
        Call SClearlist1()
        divEdt1.Visible = False
        divSch1.Visible = True
        'Call sShowData1(dr1) 'Call sClearlist1()
        Call SSearch1()
    End Sub

    '儲存
    Protected Sub BtnSaveData2_Click(sender As Object, e As EventArgs) Handles BtnSaveData2.Click
        Dim errmsg1 As String = ""
        Dim h_parms As New Hashtable
        CheckData2(errmsg1, h_parms)
        If errmsg1 <> "" Then
            Common.MessageBox(Me, errmsg1)
            Exit Sub
        End If

        'select SCORE4_1--配合分署辦理相關活動或政策宣導(3%)
        ',SCORE4_2A--參訓學員訓後動態調查表填寫人次
        ',SCORE4_2_CNT--結訓學員總人次
        ',SCORE4_2_RATE--參訓學員平均填答率
        ',SCORE4_2--參訓學員訓後動態調查表單位平均填答率達80% (2%)
        ',SUBTOTAL--小計
        ',SCORE4_1_2--配合本部、本署辦理相關活動或政策宣導 (7%)
        ',[TOTAL]--"合計 I+II+III+IV"
        '--等級	'--備註
        'FROM ORG_SCORING2 WHERE 0=0

        Call SSaveData2(h_parms)
        Call CCreate1()
    End Sub

    ''' <summary> 匯出xls審查計分表 </summary>
    ''' <param name="dt"></param>
    Sub ExpXMLXLS(ByRef dt As DataTable)
        If TIMS.dtNODATA(dt) Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        'Const cst_files_ext As String=".xlsx" ' ".xls" 
        Dim strErrmsg As String = ""
        Dim sFileName As String = String.Concat("~\CO\01\Temp\", TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim sMyFile1 As String = Server.MapPath(sFileName) '複製一份(Sample)

        Const cst_SampleXLS As String = "~\CO\01\SampleD.xlsx" '& cst_files_ext
        'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), sMyFile1, True)
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "目錄名稱或磁碟區標籤語法錯誤!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            Common.MessageBox(Me, strErrmsg)
            'Exit Sub
        End Try
        If strErrmsg <> "" Then
            Try
                strErrmsg += "Path/File: " & sMyFile1 & vbCrLf
                strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.WriteTraceLog(strErrmsg)
            Catch ex As Exception
            End Try
            Exit Sub
        End If

        '開檔
        Dim fs As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        'Dim fs As FileInfo=New FileInfo(sMyFile1)
        Dim ep As ExcelPackage = New ExcelPackage(fs)
        Dim sheet As ExcelWorksheet = ep.Workbook.Worksheets(0) '取得Sheet1 'ep.Workbook.Worksheets(1) '取得Sheet1
        'Dim i_startRowNumber As Integer=sheet.Dimension.Start.Row '起始列編號， 從1算起
        'Dim i_endRowNumber As Integer=sheet.Dimension.End.Row '結束列編號， 從1算起
        'Dim i_startColumn As Integer=sheet.Dimension.Start.Column '開始欄編號， 從1算起
        'Dim i_endColumn As Integer=sheet.Dimension.End.Column '結束欄編號， 從1算起
        'i_startRowNumber=6 ' 有包含標題 Then
        'sheet.Cells(i_startRowNumber, 1).Value="第1欄"
        'sheet.Cells(i_startRowNumber, 2).Value="第2欄"
        'For i_currentRow As Integer=i_startRowNumber To dt.Rows.Count + i_startRowNumber
        '    'Dim Range As ExcelRange=sheet.Cells(i_currentRow, i_startColumn, i_currentRow, i_endColumn) '抓出目前的Excel列
        '    '寫值
        '    sheet.Cells(i_currentRow, 1).Value=cellValue + "test";
        'Next

        Dim i_currentRow As Integer = 5
        For Each drV As DataRow In dt.Rows
            i_currentRow += 1
            Dim rCOMIDNO As String = Convert.ToString(drV("COMIDNO"))
            Dim rROWID As String = Convert.ToString(drV("ROWID"))
            Dim rOrgName As String = Convert.ToString(drV("OrgName"))
            'Dim rORGKIND_N As String=Convert.ToString(drV("ORGKIND_N"))
            Dim rORGKIND1_N As String = Convert.ToString(drV("ORGKIND1_N"))
            Dim rDISTNAME As String = Convert.ToString(drV("DISTNAME"))
            Dim rSCORE1_1 As String = Convert.ToString(drV("SCORE1_1"))
            Dim rSCORE1_2 As String = Convert.ToString(drV("SCORE1_2"))

            Dim rSCORE2_1_1_ALL As String = Convert.ToString(drV("SCORE2_1_1_ALL"))    '"2-1-1各項函送資料及資訊登錄作業時效(11%)(各班分數加總/核定總班數)"        '
            Dim rSCORE2_1_2_SUM_ALL As String = Convert.ToString(drV("SCORE2_1_2_SUM_ALL")) '"2-1-2函送資料內容及資訊登錄正確性(10%)(各班分數加總/核定總班數)"
            Dim rSCO2_1_3 As String = Convert.ToString(drV("SCORE2_1_3")) '"2-1-3訓練計畫變更項次數(7%)(各班分數加總/核定總班數)"
            'Dim rSCO2_1_1 As String=Convert.ToString(drV("SCORE2_1_1"))
            'Dim rSCO2_1_2 As String=Convert.ToString(drV("SCORE2_1_2"))

            Dim rSCO2_2_1 As String = Convert.ToString(drV("SCORE2_2_1")) '"2-2-1學員管理(4%)(各班分數加總/核定總班數)"
            Dim rSCO2_2_2 As String = Convert.ToString(drV("SCORE2_2_2")) '"2-2-2課程辦理情形(30%)
            Dim rSCO2_3_1 As String = Convert.ToString(drV("SCORE2_3_1")) '"2-3-1計畫說明、訓練活動及相關會議之出席率(5%)(計畫參與度=實際出席總場次/應出席總場次)"

            Dim rSCORE3_1 As String = Convert.ToString(drV("SCORE3_1")) '"3-1最近一次TTQS評核結果等級(10%)
            Dim rSCORE3_1b As String = Convert.ToString(drV("RESULT_N")) '獎牌
            Dim rSCORE3_2 As String = Convert.ToString(drV("SCORE3_2")) '3-2學員滿意程度(5%)

            Dim rSCORE4_1 As String = Convert.ToString(drV("SCORE4_1")) 'SCORE4_1 配合分署辦理相關活動或政策宣導(3%)

            Dim rSCORE4_2A As String = Convert.ToString(drV("SCORE4_2A")) 'SCORE4_2A 參訓學員訓後動態調查表填寫人次
            Dim rSCORE4_2_CNT As String = Convert.ToString(drV("SCORE4_2_CNT")) 'SCORE4_2_CNT 結訓學員總人次
            Dim rSCORE4_2_RATE As String = Convert.ToString(drV("SCORE4_2_RATE")) 'SCORE4_2_RATE 參訓學員平均填答率
            Dim rSCORE4_2 As String = Convert.ToString(drV("SCORE4_2")) ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
            Dim rSUBTOTAL As String = Convert.ToString(drV("SUBTOTAL")) 'SUBTOTAL 分署小計/初審分數
            Dim rSCORE4_1_2 As String = Convert.ToString(drV("SCORE4_1_2"))  'rSCORE4_1_2-總分
            Dim rBRANCHPNT As String = Convert.ToString(drV("BRANCHPNT"))  'BRANCHPNT 署／部加分項目 7%'【IV加分項目(本部)】= 圖一的【署/部加分項目】
            Dim rIMPSCORE1 As String = Convert.ToString(drV("IMPSCORE_1")) 'IMPSCORE_1 初擬分數／匯入成績
            Dim rIMPLEVEL1 As String = Convert.ToString(drV("IMPLEVEL_1")) 'IMPLEVEL_1 初擬等級／匯入等級/初審等級
            Dim rRLEVEL_2 As String = Convert.ToString(drV("RLEVEL_2")) 'RLEVEL_2 複審等級

            '【小計】邏輯
            '圖二的【小計】邏輯：(【匯入分數】優先於【分署小計】)
            '(1)當圖一的【匯入分數】有資料時 【小計】=【匯入分數】
            '(2)當圖一的【匯入分數】為空、【分署小計】有資料時 【小計】=【分署小計】
            '(3)當【匯入分數】、【分署小計】皆為空時 【小計】=顯示 0
            Dim s_SUBTOTAL As String = If(rIMPSCORE1 <> "" AndAlso Val(rIMPSCORE1) > 0, rIMPSCORE1, If(rSUBTOTAL <> "" AndAlso Val(rSUBTOTAL) > 0, rSUBTOTAL, "0"))

            '寫值
            sheet.Cells(i_currentRow, 1).Value = rCOMIDNO
            sheet.Cells(i_currentRow, 2).Value = rROWID
            sheet.Cells(i_currentRow, 3).Value = rOrgName
            sheet.Cells(i_currentRow, 4).Value = rORGKIND1_N
            sheet.Cells(i_currentRow, 5).Value = rDISTNAME
            sheet.Cells(i_currentRow, 6).Value = rSCORE1_1
            sheet.Cells(i_currentRow, 7).Value = rSCORE1_2

            sheet.Cells(i_currentRow, 8).Value = rSCORE2_1_1_ALL 'rSCO2_1_1
            sheet.Cells(i_currentRow, 9).Value = rSCORE2_1_2_SUM_ALL 'rSCO2_1_2
            sheet.Cells(i_currentRow, 10).Value = rSCO2_1_3

            sheet.Cells(i_currentRow, 11).Value = rSCO2_2_1
            sheet.Cells(i_currentRow, 12).Value = rSCO2_2_2
            sheet.Cells(i_currentRow, 13).Value = rSCO2_3_1

            sheet.Cells(i_currentRow, 14).Value = rSCORE3_1 '3_1a
            sheet.Cells(i_currentRow, 15).Value = rSCORE3_1b '3_1b
            sheet.Cells(i_currentRow, 16).Value = rSCORE3_2 '3_1b

            sheet.Cells(i_currentRow, 17).Value = rSCORE4_1
            sheet.Cells(i_currentRow, 18).Value = rSCORE4_2

            sheet.Cells(i_currentRow, 19).Value = s_SUBTOTAL 'rSUBTOTAL '"0" '小計
            sheet.Cells(i_currentRow, 20).Value = rIMPLEVEL1 '"0" '初擬等級  '【初擬等級】= 【匯入等級】
            sheet.Cells(i_currentRow, 21).Value = rBRANCHPNT  'rSCORE4_1_2 【IV加分項目(本部)】= 【署/部加分項目】

            sheet.Cells(i_currentRow, 22).Value = rSCORE4_1_2 'rSCORE4_1_2-總分【合計】= 【總分】
            sheet.Cells(i_currentRow, 23).Value = rRLEVEL_2 '【等級】(W)= 【複審等級】

            '下列 本程式功能沒有介面
            sheet.Cells(i_currentRow, 24).Value = "" '備註

            sheet.Cells(i_currentRow, 1, i_currentRow, 24).Style.Font.Size = 9
        Next
        'ep.Save()

        Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case V_ExpType
            Case "EXCEL"
                ExpExccl_1(strErrmsg, ep)
                'Response.Flush()
                'Response.Close()
                'TIMS.Utl_RespWriteEnd(Me, objconn, "")
                'Return

            Case "ODS"
                Dim myFileName1 As String = TIMS.ClearSQM("ExpFile" & TIMS.GetRnd6Eng) & ".xlsx" '檔名
                Dim myFileName2 As String = "~\CO\01\Temp\" & myFileName1 '複製
                Dim sMyFile2 As String = Server.MapPath(myFileName2)
                Dim createStream As FileStream = New FileStream(sMyFile2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
                ep.SaveAs(createStream) '存檔
                createStream.Close()
                createStream = Nothing

                Dim fr As New System.IO.FileStream(sMyFile2, IO.FileMode.Open)
                Dim br As New System.IO.BinaryReader(fr)
                Dim buf(fr.Length) As Byte
                fr.Read(buf, 0, fr.Length)
                fr.Close()

                Dim sFileName1 As String = String.Concat("審查計分表-", TIMS.GetRnd6Eng())
                'parmsExp.Add("strHTML", strHTML)
                Dim parmsExp As New Hashtable From {
                    {"ExpType", V_ExpType}, 'EXCEL/PDF/ODS
                    {"FileName", sFileName1},
                    {"xlsx_buf", buf},
                    {"ResponseNoEnd", "Y"}
                }
                TIMS.Utl_ExportRp1(Me, parmsExp)

            Case Else
                Dim s_log1 As String = $"ExpType(參數有誤)!!{V_ExpType}"
                Common.MessageBox(Me, s_log1)
                Exit Sub

        End Select

        '刪除Temp中的資料
        'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg = "" Then
            TIMS.Utl_RespWriteEnd(Me, objconn, "")
            'Call TIMS.CloseDbConn(objconn)
            'Response.End()
        End If
        If strErrmsg <> "" Then Common.MessageBox(Me, strErrmsg)
    End Sub
    Sub ExpExccl_1(ByRef strErrmsg As String, ByRef ep As ExcelPackage)
        '將新建立的excel存入記憶體下載-----   Start
        'Dim myFileName1 As String=TIMS.ClearSQM("ExpFile" & TIMS.GetRnd6Eng) & cst_files_ext '檔名
        'Dim myFileName2 As String="~\CO\01\Temp\" & myFileName1 '複製
        'Dim sMyFile2 As String=Server.MapPath(myFileName2)
        'Dim createStream As FileStream=New FileStream(sMyFile2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
        'ep.SaveAs(createStream) '存檔
        Dim myFileName1 As String = TIMS.ClearSQM(String.Concat("審查計分表-", TIMS.GetRnd6Eng(), ".xlsx")) '檔名
        Dim myFileName2 As String = String.Concat("~\CO\01\Temp\", myFileName1) '複製
        Dim sMyFile2 As String = Server.MapPath(myFileName2)

        Dim createStream As FileStream = New FileStream(sMyFile2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
        ep.SaveAs(createStream) '存檔
        createStream.Close()
        createStream = Nothing

        'https://dotblogs.com.tw/malonestudyrecord/2018/03/21/103124

        '建立檔案
        Try
            Dim fr As New System.IO.FileStream(sMyFile2, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte
            fr.Read(buf, 0, fr.Length)
            fr.Close()

            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
            'Response.ContentEncoding=System.Text.Encoding.ASCII
            Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(myFileName1, System.Text.Encoding.UTF8))
            Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            'Response.ContentType="Application/vnd.ms-Excel"
            'Common.RespWrite(Me, br.ReadBytes(fr.Length))
            Response.BinaryWrite(buf)
            'Response.End()
        Catch ex As Exception
            strErrmsg = ""
            strErrmsg += "無法存取該檔案!!!" & vbCrLf
            strErrmsg += " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            'Finally
        End Try
    End Sub
    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub

    Function GET_SCORING2_R(oVAL As String) As DataRow
        Dim PMS1 As New Hashtable From {{"VALUEFD", oVAL}}
        Dim SSQL As String = "SELECT a.TEXTFD,a.VALUEFD,a.ROC_YEARS,a.MONTHS_N,a.NEXT_YMN FROM V_SCORING2 a WHERE a.VALUEFD=@VALUEFD" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(SSQL, objconn, PMS1)
        Return dr
    End Function

    Function SEARCH_DATA1_dt1() As DataTable
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim vSCORESTAGE As String = TIMS.GetListValue(rblSCORESTAGE)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】

        Dim PMS1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}

        Dim SSQL3 As String = " AND a.IMPLEVEL_1 IS NOT NULL" '(單1條件)
        Dim SSQL As String = ""
        'SSQL &= " declare @SCORINGID varchar(99)='2025-01-2024-1-2024-2';" & vbCrLf
        'SSQL &= " declare @TPLANID varchar(99)='28';" & vbCrLf
        'SSQL &= " --declare @DISTID varchar(99)='001';" & vbCrLf
        SSQL &= " WITH WC1 AS ( SELECT oo.COMIDNO,oo.ORGNAME,oo.ORGKIND2,oo.ORGKIND1,ko.ORGTYPE ORGKIND1_N,a.DISTID" & vbCrLf
        'SUBTOTAL 初審分數 'IMPLEVEL_1/IMP1 初審等級  'SSQL &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1 IMP1" & vbCrLf
        Select Case vSCORESTAGE
            Case "01"
                SSQL &= " ,a.IMPLEVEL_1 IMP1" & vbCrLf
                SSQL3 = " AND a.IMPLEVEL_1 IS NOT NULL"
            Case "02"
                SSQL &= " ,a.MINISTERLEVEL IMP1" & vbCrLf
                SSQL3 = " AND a.MINISTERLEVEL IS NOT NULL"
            Case "03"
                SSQL &= " ,a.RLEVEL_2 IMP1" & vbCrLf
                SSQL3 = " AND a.RLEVEL_2 IS NOT NULL"
            Case Else
                SSQL &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1 IMP1" & vbCrLf
        End Select

        SSQL &= " FROM dbo.ORG_SCORING2 a" & vbCrLf
        SSQL &= " JOIN dbo.ORG_ORGINFO oo ON oo.OrgID=a.OrgID" & vbCrLf
        SSQL &= " LEFT JOIN dbo.VIEW_ORGTYPE1 ko ON ko.ORGTYPEID1=oo.ORGKIND1" & vbCrLf
        '2024-01-2023-1-2023-2'--@SCORINGID
        SSQL &= " WHERE CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        SSQL &= " AND a.TPLANID=@TPLANID" & vbCrLf
        '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
        SSQL &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        SSQL &= String.Concat(SSQL3, ")", vbCrLf)

        SSQL &= " ,WC2 AS ( SELECT a.ORGKIND2,a.DISTID,COUNT(1) CNT2 FROM WC1 a GROUP BY a.ORGKIND2,a.DISTID )" & vbCrLf
        SSQL &= " ,WC3 AS ( SELECT a.ORGKIND2,a.DISTID,a.IMP1,COUNT(1) CNT3 FROM WC1 a GROUP BY a.ORGKIND2,a.DISTID,a.IMP1  )" & vbCrLf
        SSQL &= " ,WC4G AS ( SELECT c.ORGKIND2,c.DISTID,c.IMP1,c.CNT3,b.CNT2" & vbCrLf
        SSQL &= " ,CONVERT(FLOAT,C.CNT3)/CONVERT(FLOAT,B.CNT2)*100/100 CNT4" & vbCrLf
        SSQL &= " FROM WC3 c JOIN WC2 b on b.ORGKIND2=c.ORGKIND2 AND b.DISTID=c.DISTID AND c.ORGKIND2='G')" & vbCrLf
        SSQL &= " ,WC4W AS ( SELECT c.ORGKIND2,c.DISTID,c.IMP1,c.CNT3,b.CNT2" & vbCrLf
        SSQL &= " ,CONVERT(FLOAT,C.CNT3)/CONVERT(FLOAT,B.CNT2)*100/100 CNT4" & vbCrLf
        SSQL &= " FROM WC3 c JOIN WC2 b on b.ORGKIND2=c.ORGKIND2 AND b.DISTID=c.DISTID AND c.ORGKIND2='W' )" & vbCrLf
        SSQL &= " ,WT1 AS ( SELECT kd.DISTID,t.IMP1,kd.DISTNAME,kd.DISTNAME3 FROM dbo.V_DISTRICT kd CROSS JOIN (VALUES ('A'),('B'),('C'),('D'),('NG'),('ZZ')) AS t(IMP1) WHERE kd.DISTID!='000' )" & vbCrLf
        SSQL &= " SELECT kd.DISTID,kd.IMP1,kd.DISTNAME,kd.DISTNAME3,'-' CNVL" & vbCrLf
        SSQL &= " ,c.CNT3 CNT3G,c.CNT2 CNT2G,c.CNT4 CNT4G" & vbCrLf
        SSQL &= " ,w.CNT3 CNT3W,w.CNT2 CNT2W,w.CNT4 CNT4W" & vbCrLf
        SSQL &= " FROM WT1 kd" & vbCrLf
        SSQL &= " LEFT JOIN WC4G c ON c.DISTID=kd.DISTID AND c.IMP1=kd.IMP1" & vbCrLf
        SSQL &= " LEFT JOIN WC4W w ON w.DISTID=kd.DISTID AND w.IMP1=kd.IMP1" & vbCrLf
        SSQL &= " ORDER BY kd.DISTID,kd.IMP1" & vbCrLf
        'If TIMS.sUtl_ChkTest() Then
        '    TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(PMS1), vbCrLf, "--##CO_01_005, SSQL:", vbCrLf, SSQL))
        'End If
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt
    End Function
    ''' <summary>
    ''' 匯出等級比率統計表(署用)
    ''' </summary>
    Sub ExportXlsStd28_2()
        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const cst_SampleXLS As String = "~\CO\01\sampleC01005a.xlsx" '& cst_files_ext 'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If

        Dim strErrmsg As String = ""
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim sMyFile1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), sMyFile1, True)
        Catch ex As Exception
            strErrmsg = String.Concat("目錄名稱或磁碟區標籤語法錯誤!!!", vbCrLf, " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)", vbCrLf, ex.Message, vbCrLf)
            Common.MessageBox(Me, strErrmsg)
            TIMS.LOG.Error(ex.Message, ex)
            Return 'Exit Sub
        End Try

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim drSC2 As DataRow = GET_SCORING2_R(vSCORINGID)
        If drSC2 Is Nothing Then
            Common.MessageBox(Me, "查無 匯出資料!")
            Exit Sub
        End If
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt1()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim v_SCORESTAGE As String = TIMS.GetListText(rblSCORESTAGE)
        'Dim vROC_YEARS As String = Convert.ToString(drSC2("ROC_YEARS")) 'Dim vMONTHS_N As String = Convert.ToString(drSC2("MONTHS_N"))
        '"OOO年度<申請階段>產業人才投資方案審查計分結果統計表 (<分數階段>)"
        'Dim SF_TITLE1 As String = String.Concat(vROC_YEARS, "年度", vMONTHS_N, "產業人才投資方案審查計分結果統計表 (", v_SCORESTAGE, ")")
        Dim vNEXT_YMN As String = Convert.ToString(drSC2("NEXT_YMN"))
        Dim SF_TITLE1 As String = String.Concat(vNEXT_YMN, "產業人才投資方案審查計分結果統計表", " (", v_SCORESTAGE, ")")
        Dim s_FILENAME1 As String = String.Concat(vNEXT_YMN, "產業人才投資方案審查計分結果統計表", "x", v_SCORESTAGE, "x", TIMS.GetDateNo2(3))
        Dim fg_RespWriteEnd As Boolean = False
        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial
            'ExcelPackage.LicenseContext=LicenseContext.NonCommercial

            'Dim file1 As New FileInfo(filePath1)
            'Dim ndt As DateTime = Now

            '開檔
            Using fs1 As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim ep As New ExcelPackage(fs1)
                Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)
                'Dim ep As New ExcelPackage()
                'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)

                ws.Cells("A1:I1").Value = SF_TITLE1
                ws.Cells("A1:I1").Style.Font.Bold = True

                Dim idxStr1 As Integer = 4
                'DISTID	IMP1	DISTNAME	DISTNAME3	CNVL	CNT3G	CNT2G	CNT4G	CNT3W	CNT2W	CNT4W
                '001 A	勞動力發展署北基宜花金馬分署	北分署	-	50	199	25.13%	41	165	24.85%
                Dim AG_CT As Integer = 0
                Dim BG_CT As Integer = 0
                Dim CG_CT As Integer = 0
                Dim DG_CT As Integer = 0

                Dim AW_CT As Integer = 0
                Dim BW_CT As Integer = 0
                Dim CW_CT As Integer = 0
                Dim DW_CT As Integer = 0
                Dim G_ALCT As Double = 0
                Dim W_ALCT As Double = 0
                Dim ivCNT2G As Integer = 0 'TIMS.VAL1(dr("CNT2G"))
                Dim ivCNT2W As Integer = 0  'TIMS.VAL1(dr("CNT2W"))
                For Each dr As DataRow In dtXls1.Rows
                    Select Case Convert.ToString(dr("IMP1"))
                        Case "ZZ"
                            ws.Cells("C" & idxStr1).Value = ivCNT2G 'TIMS.VAL1(dr("CNT2G"))
                            ws.Cells("D" & idxStr1).Value = If(ivCNT2G > 0, TIMS.VAL1("1"), 0)
                            ws.Cells("H" & idxStr1).Value = ivCNT2W 'TIMS.VAL1(dr("CNT2W"))
                            ws.Cells("I" & idxStr1).Value = If(ivCNT2W > 0, TIMS.VAL1("1"), 0)
                            ivCNT2G = 0 'TIMS.VAL1(dr("CNT2G"))
                            ivCNT2W = 0 'TIMS.VAL1(dr("CNT2W"))
                        Case "NG"
                            ws.Cells("C" & idxStr1).Value = 0
                            ws.Cells("D" & idxStr1).Value = TIMS.VAL1("0")
                            ws.Cells("H" & idxStr1).Value = 0
                            ws.Cells("I" & idxStr1).Value = TIMS.VAL1("0")
                        Case Else
                            If (ivCNT2G = 0) Then ivCNT2G = TIMS.VAL1(dr("CNT2G"))
                            If (ivCNT2W = 0) Then ivCNT2W = TIMS.VAL1(dr("CNT2W"))
                            ws.Cells("C" & idxStr1).Value = TIMS.VAL1(dr("CNT3G"))
                            ws.Cells("D" & idxStr1).Value = TIMS.VAL1(dr("CNT4G"))
                            ws.Cells("H" & idxStr1).Value = TIMS.VAL1(dr("CNT3W"))
                            ws.Cells("I" & idxStr1).Value = TIMS.VAL1(dr("CNT4W"))

                            Select Case Convert.ToString(dr("IMP1"))
                                Case "A"
                                    AG_CT += TIMS.VAL1(dr("CNT3G"))
                                    AW_CT += TIMS.VAL1(dr("CNT3W"))
                                Case "B"
                                    BG_CT += TIMS.VAL1(dr("CNT3G"))
                                    BW_CT += TIMS.VAL1(dr("CNT3W"))
                                Case "C"
                                    CG_CT += TIMS.VAL1(dr("CNT3G"))
                                    CW_CT += TIMS.VAL1(dr("CNT3W"))
                                Case "D"
                                    DG_CT += TIMS.VAL1(dr("CNT3G"))
                                    DW_CT += TIMS.VAL1(dr("CNT3W"))
                            End Select
                            G_ALCT += TIMS.VAL1(dr("CNT3G"))
                            W_ALCT += TIMS.VAL1(dr("CNT3W"))
                    End Select
                    ws.Cells("D" & idxStr1).Style.Numberformat.Format = "0.0%"
                    ws.Cells("I" & idxStr1).Style.Numberformat.Format = "0.0%"
                    idxStr1 += 1
                Next

                For idx1 As Integer = 34 To 39
                    Dim iG_CT As Double = 0
                    Dim iW_CT As Double = 0
                    Dim iG_RTE As Double = 0
                    Dim iW_RTE As Double = 0
                    Select Case idx1
                        Case 34
                            iG_CT = AG_CT
                            iG_RTE = If(G_ALCT > 0, AG_CT / G_ALCT * 100 / 100, 0)
                            iW_CT = AW_CT
                            iW_RTE = If(W_ALCT > 0, AW_CT / W_ALCT * 100 / 100, 0)
                        Case 35
                            iG_CT = BG_CT
                            iG_RTE = If(G_ALCT > 0, BG_CT / G_ALCT * 100 / 100, 0)
                            iW_CT = BW_CT
                            iW_RTE = If(W_ALCT > 0, BW_CT / W_ALCT * 100 / 100, 0)
                        Case 36
                            iG_CT = CG_CT
                            iG_RTE = If(G_ALCT > 0, CG_CT / G_ALCT * 100 / 100, 0)
                            iW_CT = CW_CT
                            iW_RTE = If(W_ALCT > 0, CW_CT / W_ALCT * 100 / 100, 0)
                        Case 37
                            iG_CT = DG_CT
                            iG_RTE = If(G_ALCT > 0, DG_CT / G_ALCT * 100 / 100, 0)
                            iW_CT = DW_CT
                            iW_RTE = If(W_ALCT > 0, DW_CT / W_ALCT * 100 / 100, 0)
                        Case 38
                            iG_CT = 0
                            iG_RTE = 0
                            iW_CT = 0
                            iW_RTE = 0
                        Case 39
                            iG_CT = G_ALCT
                            iG_RTE = If(G_ALCT > 0, 1, 0)
                            iW_CT = W_ALCT
                            iW_RTE = If(W_ALCT > 0, 1, 0)
                    End Select
                    ws.Cells("C" & idx1).Value = iG_CT
                    ws.Cells("D" & idx1).Value = iG_RTE
                    ws.Cells("H" & idx1).Value = iW_CT
                    ws.Cells("I" & idx1).Value = iW_RTE
                    ws.Cells("D" & idx1).Style.Numberformat.Format = "0.0%"
                    ws.Cells("I" & idx1).Style.Numberformat.Format = "0.0%"
                Next

                'Dim iFivePoint As Double = 0
                'Dim iFiveCnt As Double = 0
                'Dim iFivePoint2 As Double = 0
                'Dim iFiveCnt2 As Double = 0
                'Dim i5 As Double = 5
                'For idx As Integer = 4 To 8
                '    iFivePoint += TIMS.VAL1(ws.Cells("I" & idx).Value) * i5
                '    iFiveCnt += TIMS.VAL1(ws.Cells("I" & idx).Value)
                '    iFivePoint2 += TIMS.VAL1(ws.Cells("I" & (idx + 5)).Value) * i5
                '    iFiveCnt2 += TIMS.VAL1(ws.Cells("I" & (idx + 5)).Value)
                '    i5 -= 1
                'Next
                'ws.Cells("L4").Value = iFivePoint / iFiveCnt
                'ws.Cells("L9").Value = iFivePoint2 / iFiveCnt2
                'ws.Cells("M4").Value = iFivePoint / iFiveCnt * 20
                'ws.Cells("M9").Value = iFivePoint2 / iFiveCnt2 * 20
                'ws.Cells("L4").Style.Numberformat.Format = "0.00"
                'ws.Cells("L9").Style.Numberformat.Format = "0.00"
                'ws.Cells("M4").Style.Numberformat.Format = "0.00"
                'ws.Cells("M9").Style.Numberformat.Format = "0.00"

                ' 設定貨幣格式，小數位數為 0
                'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format="$#,##0" ' 美元符號，您可以根據需要更改
                'ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width=33

                ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom=70 無法運行 修正為 ws.View.ZoomScale=70 才可運行
                'ws.View.ZoomScale = 90

                Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
                Select Case V_ExpType
                    Case "EXCEL"
                        TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case "ODS"
                        TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case Else
                        Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                        Common.MessageBox(Me, s_log1)
                        Return ' Exit Sub
                End Select
            End Using
            Call TIMS.MyFileDelete(sMyFile1)
            If fg_RespWriteEnd Then TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub

    Protected Sub BtnExp2_Click(sender As Object, e As EventArgs) Handles BtnExp2.Click
        ExportXlsStd28_2()
    End Sub

    '匯出統計表(署用)
    Function SEARCH_DATA1_dt3(vORGKIND2 As String) As DataTable
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim vSCORESTAGE As String = TIMS.GetListValue(rblSCORESTAGE)
        Dim PMS1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}, {"ORGKIND2", vORGKIND2}}
        'Declare @SCORINGID NVarChar(21)='2025-01-2024-1-2024-2';/*1*/'Declare @TPLANID NVarChar(2)='28';/*2*/ 'Declare @ORGKIND2 NVarChar(1)='G';/*3*/

        Dim SSQL3 As String = " AND a.IMPLEVEL_1 IS NOT NULL" '(單1條件) 條件修改
        Dim SSQL_IMP1 As String = "" '欄位選擇
        Dim SSQL_ORDER1 As String = "" '排序調整1
        Dim SSQL_ORDER2 As String = "" '排序調整2
        Select Case vSCORESTAGE
            Case "01"
                SSQL_ORDER1 = "SELECT ROW_NUMBER() OVER(PARTITION BY a.DISTID ORDER BY a.DISTID ASC,a.SUBTOTAL DESC,oo.ORGNAME ASC) ROWID"
                SSQL_IMP1 = " ,A.SUBTOTAL SCORE412,a.IMPLEVEL_1 IMP1"
                SSQL3 = " AND a.IMPLEVEL_1 IS NOT NULL"
                SSQL_ORDER2 = " ORDER BY a.DISTID ASC,a.SUBTOTAL DESC,oo.ORGNAME ASC"
            Case "02"
                SSQL_ORDER1 = "SELECT ROW_NUMBER() OVER(PARTITION BY a.DISTID ORDER BY a.DISTID ASC,a.MINISTERSUB DESC,oo.ORGNAME ASC) ROWID"
                SSQL_IMP1 = " ,A.MINISTERSUB SCORE412,a.MINISTERLEVEL IMP1"
                SSQL3 = " AND a.MINISTERLEVEL IS NOT NULL"
                SSQL_ORDER2 = " ORDER BY a.DISTID ASC,a.MINISTERSUB DESC,oo.ORGNAME ASC"
            Case "03"
                SSQL_ORDER1 = "SELECT ROW_NUMBER() OVER(PARTITION BY a.DISTID ORDER BY a.DISTID ASC,a.SCORE4_1_2 DESC,oo.ORGNAME ASC) ROWID"
                SSQL_IMP1 = " ,A.SCORE4_1_2 SCORE412,a.RLEVEL_2 IMP1"
                SSQL3 = " AND a.RLEVEL_2 IS NOT NULL"
                SSQL_ORDER2 = " ORDER BY a.DISTID ASC,a.SCORE4_1_2 DESC,oo.ORGNAME ASC"
            Case Else
                SSQL_ORDER1 = "SELECT ROW_NUMBER() OVER(PARTITION BY a.DISTID ORDER BY a.DISTID ASC,a.SUBTOTAL DESC,oo.ORGNAME ASC) ROWID"
                SSQL_IMP1 = " ,ISNULL(a.SCORE4_1_2,a.SUBTOTAL) SCORE412,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1 IMP1" & vbCrLf
                SSQL_ORDER2 = " ORDER BY a.DISTID ASC,a.SUBTOTAL DESC,oo.ORGNAME ASC"
        End Select

        Dim SSQL As String = ""
        SSQL &= SSQL_ORDER1
        SSQL &= " ,oo.COMIDNO,oo.ORGNAME ,oo.ORGKIND ,oo.ORGKIND1 ,kd.DISTNAME3" & vbCrLf
        'SSQL &= " ,(SELECT x.MASTERNAME FROM V_ORGINFO x WHERE x.COMIDNO=oo.COMIDNO) MASTERNAME" & vbCrLf
        SSQL &= " ,(SELECT op.MASTERNAME FROM dbo.VIEW_ORGPLANINFO op WHERE op.ORGID=a.ORGID AND op.TPLANID=a.TPLANID AND op.DISTID=a.DISTID AND op.YEARS=a.YEARS) MASTERNAME" & vbCrLf
        SSQL &= " ,(SELECT x.MASTERNAME FROM V_ORGINFO x WHERE x.COMIDNO=oo.COMIDNO) MASTERNAME1" & vbCrLf
        SSQL &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS) ,'年',a.MONTHS,'月'" & vbCrLf
        SSQL &= " ,'(',dbo.FN_CYEAR2(a.YEARS1) ,'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end ,'~'" & vbCrLf
        SSQL &= " ,dbo.FN_CYEAR2(a.YEARS2) ,'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end ,')') SCORING_N" & vbCrLf
        SSQL &= " ,CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
        SSQL &= " ,a.MINISTERADD,a.SCORE4_1,a.SUBTOTAL" & vbCrLf
        'SSQL &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1" & vbCrLf
        SSQL &= SSQL_IMP1
        SSQL &= " ,a.FIRSTCHK,a.SECONDCHK,b.APPLIEDRESULT,'' bMEMO" & vbCrLf
        SSQL &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
        SSQL &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
        SSQL &= " JOIN dbo.V_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID" & vbCrLf
        SSQL &= " LEFT JOIN dbo.ORG_TTQS2 b WITH(NOLOCK) ON concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
        SSQL &= " WHERE CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        SSQL &= " AND a.TPLANID=@TPLANID AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
        SSQL &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        'SSQL &= " AND a.IMPLEVEL_1 IS NOT NULL" & vbCrLf
        SSQL &= $"{SSQL3}{SSQL_ORDER2}"

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(PMS1), vbCrLf, "--##CO_01_005:", vbCrLf, SSQL))
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt
    End Function

    '匯出統計表(署用)
    Sub ExportXlsStd28_3()
        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim drSC2 As DataRow = GET_SCORING2_R(vSCORINGID)
        If drSC2 Is Nothing Then
            Common.MessageBox(Me, "查無 匯出資料!")
            Exit Sub
        End If
        Dim dtXls1G As DataTable = SEARCH_DATA1_dt3("G")
        Dim dtXls1W As DataTable = SEARCH_DATA1_dt3("W")
        If TIMS.dtNODATA(dtXls1G) AndAlso TIMS.dtNODATA(dtXls1W) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        'Dim drX1 As DataRow = dtXls1.Rows(0)
        Dim END_COL_NM As String = "G"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim cellsCOLSPNumF2 As String = String.Concat("A2:", END_COL_NM, "{0}") '(畫格子使用)
        Dim strErrmsg As String = ""

        '114年度下半年產業人才投資計畫審查計分統計表(五分署)/114年度下半年提升勞工自主學習計畫審查計分統計表(五分署)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim v_SCORESTAGE As String = TIMS.GetListText(rblSCORESTAGE)
        'Dim vROC_YEARS As String = Convert.ToString(drSC2("ROC_YEARS"))'Dim vMONTHS_N As String = Convert.ToString(drSC2("MONTHS_N"))
        Dim vNEXT_YMN As String = Convert.ToString(drSC2("NEXT_YMN"))
        Dim SF_TITLE1 As String = String.Concat(vNEXT_YMN, "產業人才投資計畫審查計分統計表(五分署) (", v_SCORESTAGE, ")")
        Dim SF_TITLE2 As String = String.Concat(vNEXT_YMN, "提升勞工自主學習計畫審查計分統計表(五分署) (", v_SCORESTAGE, ")")
        Dim s_FILENAME1 As String = String.Concat(vNEXT_YMN, "計畫審查計分統計表(五分署)", v_SCORESTAGE, "x", TIMS.GetDateNo2(3))

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial 'ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            'Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage()
#Region "ORGKIND2-G"
            Dim V_SHEETNM1 As String = "產投-統計"
            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 16
                    .Value = SF_TITLE1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    '.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using
            '編號	分署	訓練單位名稱	負責人	等級	總分	備註
            'ROWID	DISTNAME3	ORGNAME	MASTERNAME	IMP1	SCORE4_1_2	bMEMO
            Dim idxStr As Integer = 2
            TIMS.SetCellValue(ws, "A" & idxStr, "編號") 'ROWID
            TIMS.SetCellValue(ws, "B" & idxStr, "分署") 'DISTNAME3
            TIMS.SetCellValue(ws, "C" & idxStr, "訓練單位名稱") 'ORGNAME
            TIMS.SetCellValue(ws, "D" & idxStr, "負責人") 'MASTERNAME
            TIMS.SetCellValue(ws, "E" & idxStr, "等級") 'RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1
            TIMS.SetCellValue(ws, "F" & idxStr, "總分") 'SCORE4_1_2
            TIMS.SetCellValue(ws, "G" & idxStr, "備註") 'bMEMO
            ws.Cells(String.Concat("A2:", END_COL_NM, "2")).Style.Font.Bold = True
            ws.Cells(String.Concat("A2:", END_COL_NM, "2")).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)

            idxStr = 3
            For Each dr1 As DataRow In dtXls1G.Rows
                '編號	分署	訓練單位名稱	負責人	等級	總分	備註
                'ROWID	DISTNAME3	ORGNAME	MASTERNAME	IMPLEVEL_1	SCORE4_1_2	bMEMO
                TIMS.SetCellValue(ws, "A" & idxStr, dr1("ROWID")) '序號
                TIMS.SetCellValue(ws, "B" & idxStr, dr1("DISTNAME3"))
                TIMS.SetCellValue(ws, "C" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left)
                Dim vMASTERNAME As String = $"{dr1("MASTERNAME")}"
                If vMASTERNAME = "" Then vMASTERNAME = $"{dr1("MASTERNAME1")}"
                TIMS.SetCellValue(ws, "D" & idxStr, vMASTERNAME)
                TIMS.SetCellValue(ws, "E" & idxStr, dr1("IMP1"))
                TIMS.SetCellValue(ws, "F" & idxStr, dr1("SCORE412"))
                TIMS.SetCellValue(ws, "G" & idxStr, dr1("bMEMO"))
                idxStr += 1
            Next
            idxStr -= 1 '(畫線)
            Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
                With exlRow3X
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s 'FontSize
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    .AutoFitColumns(25.0, 250.0)
                End With
                TIMS.SetCellBorder(exlRow3X)
            End Using
            ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format = "00.0"
            ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width = 8
            ws.Column(ws.Cells(String.Format("B3:B{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("C3:C{0}", idxStr)).Start.Column).Width = 88
            ws.Column(ws.Cells(String.Format("D3:D{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("E3:E{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("F3:F{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("G3:G{0}", idxStr)).Start.Column).Width = 12

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90
#End Region
#Region "ORGKIND2-W"
            Dim V_SHEETNM2 As String = "自主-統計"
            ws = ep.Workbook.Worksheets.Add(V_SHEETNM2)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)
            Using exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
                With exlRow1
                    .Merge = True
                    .Style.Font.Name = fontName
                    .Style.Font.Size = 16
                    .Value = SF_TITLE2
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    '.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using
            idxStr = 2 'Dim idxStr As Integer = 2
            TIMS.SetCellValue(ws, "A" & idxStr, "編號") 'ROWID
            TIMS.SetCellValue(ws, "B" & idxStr, "分署") 'DISTNAME3
            TIMS.SetCellValue(ws, "C" & idxStr, "訓練單位名稱") 'ORGNAME
            TIMS.SetCellValue(ws, "D" & idxStr, "理事長") '理事長/負責人'MASTERNAME
            TIMS.SetCellValue(ws, "E" & idxStr, "等級") 'RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1
            TIMS.SetCellValue(ws, "F" & idxStr, "總分") 'SCORE4_1_2
            TIMS.SetCellValue(ws, "G" & idxStr, "備註") 'bMEMO
            ws.Cells(String.Concat("A2:", END_COL_NM, "2")).Style.Font.Bold = True
            ws.Cells(String.Concat("A2:", END_COL_NM, "2")).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
            idxStr = 3
            For Each dr1 As DataRow In dtXls1W.Rows
                TIMS.SetCellValue(ws, "A" & idxStr, dr1("ROWID")) '序號
                TIMS.SetCellValue(ws, "B" & idxStr, dr1("DISTNAME3"))
                TIMS.SetCellValue(ws, "C" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left)
                Dim vMASTERNAME As String = $"{dr1("MASTERNAME")}"
                If vMASTERNAME = "" Then vMASTERNAME = $"{dr1("MASTERNAME1")}"
                TIMS.SetCellValue(ws, "D" & idxStr, vMASTERNAME)
                TIMS.SetCellValue(ws, "E" & idxStr, dr1("IMP1"))
                TIMS.SetCellValue(ws, "F" & idxStr, dr1("SCORE412"))
                TIMS.SetCellValue(ws, "G" & idxStr, dr1("bMEMO"))
                idxStr += 1
            Next
            idxStr -= 1 '(畫線)
            Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
                With exlRow3X
                    .Style.Font.Name = fontName
                    .Style.Font.Size = fontSize12s 'FontSize
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    .AutoFitColumns(25.0, 250.0)
                End With
                TIMS.SetCellBorder(exlRow3X)
            End Using

            '設定格式，小數位數為 1
            ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format = "00.0"
            ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width = 8
            ws.Column(ws.Cells(String.Format("B3:B{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("C3:C{0}", idxStr)).Start.Column).Width = 88
            ws.Column(ws.Cells(String.Format("D3:D{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("E3:E{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("F3:F{0}", idxStr)).Start.Column).Width = 12
            ws.Column(ws.Cells(String.Format("G3:G{0}", idxStr)).Start.Column).Width = 12

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90
#End Region

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

    '匯出統計表(署用)
    Protected Sub BtnExp3_Click(sender As Object, e As EventArgs) Handles BtnExp3.Click
        ExportXlsStd28_3()
    End Sub

    Function SEARCH_DATA1_dt4() As DataTable
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim vSCORESTAGE As String = TIMS.GetListValue(rblSCORESTAGE)
        Dim PMS1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}

        Dim SSQL3 As String = " AND a.IMPLEVEL_1 IS NOT NULL" '(單1條件)
        Dim SSQL As String = ""
        SSQL &= " WITH WC1 AS ( SELECT oo.COMIDNO,oo.ORGNAME,oo.ORGKIND2,oo.ORGKIND1,ko.ORGTYPE ORGKIND1_N,a.DISTID" & vbCrLf
        Select Case vSCORESTAGE
            Case "01"
                SSQL &= " ,A.SUBTOTAL SCORE412,a.IMPLEVEL_1 IMP1" & vbCrLf
                SSQL3 = " AND a.IMPLEVEL_1 IS NOT NULL"
            Case "02"
                SSQL &= " ,A.MINISTERSUB SCORE412,a.MINISTERLEVEL IMP1" & vbCrLf
                SSQL3 = " AND a.MINISTERLEVEL IS NOT NULL"
            Case "03"
                SSQL &= " ,A.SCORE4_1_2 SCORE412,a.RLEVEL_2 IMP1" & vbCrLf
                SSQL3 = " AND a.RLEVEL_2 IS NOT NULL"
            Case Else
                SSQL &= " ,A.SCORE4_1_2 SCORE412,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) IMP1" & vbCrLf
                SSQL3 = " AND a.IMPLEVEL_1 IS NOT NULL"
        End Select
        SSQL &= " FROM dbo.ORG_SCORING2 a" & vbCrLf
        SSQL &= " JOIN dbo.ORG_ORGINFO oo ON oo.OrgID=a.OrgID" & vbCrLf
        SSQL &= " LEFT JOIN dbo.VIEW_ORGTYPE1 ko ON ko.ORGTYPEID1=oo.ORGKIND1" & vbCrLf
        SSQL &= " WHERE CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        SSQL &= " AND a.TPLANID=@TPLANID" & vbCrLf
        SSQL &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
        SSQL &= String.Concat(SSQL3, ")", vbCrLf)
        SSQL &= " ,WC2 AS ( SELECT a.ORGKIND2,a.DISTID,COUNT(1) CNT2 FROM WC1 a GROUP BY a.ORGKIND2,a.DISTID )" & vbCrLf
        SSQL &= " ,WC3 AS ( SELECT a.ORGKIND2,a.DISTID,a.IMP1,COUNT(1) CNT3,MIN(a.SCORE412) SCORE412 FROM WC1 a GROUP BY a.ORGKIND2,a.DISTID,a.IMP1 )" & vbCrLf
        SSQL &= " ,WC4W AS ( SELECT c.ORGKIND2,c.DISTID,c.IMP1,c.CNT3,b.CNT2,c.SCORE412" & vbCrLf
        SSQL &= " ,CONVERT(FLOAT,C.CNT3)/CONVERT(FLOAT,B.CNT2)*100/100 CNT4" & vbCrLf
        SSQL &= " FROM WC3 c JOIN WC2 b ON b.ORGKIND2=c.ORGKIND2 AND b.DISTID=c.DISTID AND c.ORGKIND2='W' )" & vbCrLf
        SSQL &= " SELECT DISTID" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'A' THEN CNT3 END) CNT3A" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'B' THEN CNT3 END) CNT3B" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'C' THEN CNT3 END) CNT3C" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'D' THEN CNT3 END) CNT3D" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'A' THEN CNT4 END) CNT4A" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'B' THEN CNT4 END) CNT4B" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'C' THEN CNT4 END) CNT4C" & vbCrLf
        SSQL &= " ,MAX(CASE IMP1 WHEN 'D' THEN CNT4 END) CNT4D" & vbCrLf
        SSQL &= " ,MAX(CNT2) CNT2" & vbCrLf
        SSQL &= " ,MIN(CASE IMP1 WHEN 'A' THEN SCORE412 END) SCORE412A" & vbCrLf
        SSQL &= " ,MIN(CASE IMP1 WHEN 'B' THEN SCORE412 END) SCORE412B" & vbCrLf
        SSQL &= " ,MIN(CASE IMP1 WHEN 'C' THEN SCORE412 END) SCORE412C" & vbCrLf
        SSQL &= " ,MIN(CASE IMP1 WHEN 'D' THEN SCORE412 END) SCORE412D" & vbCrLf
        SSQL &= " FROM WC4W" & vbCrLf
        SSQL &= " GROUP BY DISTID" & vbCrLf
        SSQL &= " ORDER BY DISTID" & vbCrLf

        'If TIMS.sUtl_ChkTest() Then
        '    TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(PMS1), vbCrLf, "--##CO_01_005, SSQL:", vbCrLf, SSQL))
        'End If
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt
    End Function
    ''' <summary>
    ''' (自主)各等級分配比率
    ''' </summary>
    Sub ExportXlsStd28_4()
        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const cst_SampleXLS As String = "~\CO\01\sampleC01005c.xlsx" '& cst_files_ext 'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If

        Dim strErrmsg As String = ""
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xlsx") '複製一份(Sample)
        Dim sMyFile1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), sMyFile1, True)
        Catch ex As Exception
            strErrmsg = String.Concat("目錄名稱或磁碟區標籤語法錯誤!!!", vbCrLf, " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉)", vbCrLf, ex.Message, vbCrLf)
            Common.MessageBox(Me, strErrmsg)
            TIMS.LOG.Error(ex.Message, ex)
            Return 'Exit Sub
        End Try

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim drSC2 As DataRow = GET_SCORING2_R(vSCORINGID)
        If drSC2 Is Nothing Then
            Common.MessageBox(Me, "查無 匯出資料!")
            Exit Sub
        End If
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt4()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim vROC_YMD_NOW As String = TIMS.GetROCTWDate(Now)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim v_SCORESTAGE As String = TIMS.GetListText(rblSCORESTAGE)
        Dim vNEXT_YMN As String = Convert.ToString(drSC2("NEXT_YMN"))
        '114年度下半年提升勞工自主學習計畫各分署各等級分配比率
        Dim SF_TITX As String = String.Concat(vNEXT_YMN, "提升勞工自主學習計畫各分署各等級分配比率")
        Dim SF_TITLE1 As String = String.Concat(SF_TITX, " (", v_SCORESTAGE, ")")
        Dim SF_TITX2 As String = String.Concat(vNEXT_YMN, "提升勞工自主學習計畫各分署各等級最低分數")
        Dim SF_TITLE2 As String = String.Concat(SF_TITX2, " (", v_SCORESTAGE, ")")
        Dim s_FILENAME1 As String = String.Concat(SF_TITX, "x", v_SCORESTAGE, "x", TIMS.GetDateNo2(3))
        Dim fg_RespWriteEnd As Boolean = False
        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial 'ExcelPackage.LicenseContext=LicenseContext.NonCommercial
            'Dim file1 As New FileInfo(filePath1) 'Dim ndt As DateTime = Now

            '開檔
            Using fs1 As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim ep As New ExcelPackage(fs1)
                Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)
                'Dim ep As New ExcelPackage() 'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)

                Dim S_COLNM1 As String = "A1:F1"
                ws.Cells(S_COLNM1).Value = SF_TITLE1
                ws.Cells(S_COLNM1).Style.Font.Bold = True
                ws.Cells(S_COLNM1).Style.Font.Size = fontSize12s
                ws.Cells("F2").Value = vROC_YMD_NOW
                Dim idxStr1 As Integer = 5
                For Each dr As DataRow In dtXls1.Rows
                    ws.Cells("B" & idxStr1).Value = String.Concat(dr("CNT3A"), "家")
                    ws.Cells("C" & idxStr1).Value = String.Concat(dr("CNT3B"), "家")
                    ws.Cells("D" & idxStr1).Value = String.Concat(dr("CNT3C"), "家")
                    ws.Cells("E" & idxStr1).Value = String.Concat(dr("CNT3D"), "家")
                    ws.Cells("F" & idxStr1).Value = dr("CNT2")
                    idxStr1 += 1
                    ws.Cells("B" & idxStr1).Value = TIMS.VAL1(dr("CNT4A"))
                    ws.Cells("C" & idxStr1).Value = TIMS.VAL1(dr("CNT4B"))
                    ws.Cells("D" & idxStr1).Value = TIMS.VAL1(dr("CNT4C"))
                    ws.Cells("E" & idxStr1).Value = TIMS.VAL1(dr("CNT4D"))
                    ws.Cells("B" & idxStr1).Style.Numberformat.Format = "0.0%"
                    ws.Cells("C" & idxStr1).Style.Numberformat.Format = "0.0%"
                    ws.Cells("D" & idxStr1).Style.Numberformat.Format = "0.0%"
                    ws.Cells("E" & idxStr1).Style.Numberformat.Format = "0.0%"
                    idxStr1 += 1
                Next

                S_COLNM1 = "A18:F18"
                ws.Cells(S_COLNM1).Value = SF_TITLE2
                ws.Cells(S_COLNM1).Style.Font.Bold = True
                ws.Cells(S_COLNM1).Style.Font.Size = fontSize12s
                ws.Cells("F19").Value = vROC_YMD_NOW
                idxStr1 = 21
                For Each dr As DataRow In dtXls1.Rows
                    ws.Cells("B" & idxStr1).Value = TIMS.VAL1(dr("SCORE412A"))
                    ws.Cells("C" & idxStr1).Value = TIMS.VAL1(dr("SCORE412B"))
                    ws.Cells("D" & idxStr1).Value = TIMS.VAL1(dr("SCORE412C"))
                    ws.Cells("E" & idxStr1).Value = TIMS.VAL1(dr("SCORE412D"))
                    ws.Cells("B" & idxStr1).Style.Numberformat.Format = "00.0"
                    ws.Cells("C" & idxStr1).Style.Numberformat.Format = "00.0"
                    ws.Cells("D" & idxStr1).Style.Numberformat.Format = "00.0"
                    ws.Cells("E" & idxStr1).Style.Numberformat.Format = "00.0"
                    idxStr1 += 1
                Next

                ' 設定貨幣格式，小數位數為 0
                'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format="$#,##0" ' 美元符號，您可以根據需要更改
                'ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width=33

                ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom=70 無法運行 修正為 ws.View.ZoomScale=70 才可運行
                'ws.View.ZoomScale = 90

                Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
                Select Case V_ExpType
                    Case "EXCEL"
                        TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case "ODS"
                        TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case Else
                        Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                        Common.MessageBox(Me, s_log1)
                        Return ' Exit Sub
                End Select
            End Using
            Call TIMS.MyFileDelete(sMyFile1)
            If fg_RespWriteEnd Then TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub
    Protected Sub BtnExp4_Click(sender As Object, e As EventArgs) Handles BtnExp4.Click
        ExportXlsStd28_4()
    End Sub
    Function GET_BLACKLIST_dt1() As DataTable
        Dim SSQL As String = ""
        SSQL &= " SELECT b.OBSN,b.COMIDNO,b.DISTID,b.OBSDATE,DATEADD(YEAR,b.OBYEARS,b.OBSDATE) OBSDATE_F,b.OBYEARS,b.TPLANID,b.AVAIL" & vbCrLf
        SSQL &= " FROM dbo.ORG_BLACKLIST b WHERE b.OBSDATE<=GETDATE() AND DATEADD(YEAR,b.OBYEARS,b.OBSDATE)>=GETDATE() AND b.AVAIL='Y'" & vbCrLf
        Return DbAccess.GetDataTable(SSQL, objconn)
    End Function
    Function GET_SCORING2_dt5_6(oVAL As String) As DataTable
        Dim PMS1 As New Hashtable From {{"VALUEFD", oVAL}}
        Dim SSQL As String = ""
        SSQL &= " SELECT TOP 6 TEXTFD,VALUEFD,ROC_YEARS,MONTHS_N,NEXT_YMN,NEXT_YMN2,YEARS,MONTHS" & vbCrLf
        SSQL &= " FROM V_SCORING2 WHERE VALUEFD<=@VALUEFD" & vbCrLf
        SSQL &= " ORDER BY YEARS DESC,MONTHS DESC" & vbCrLf
        Dim dt6 As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt6
    End Function
    Function SEARCH_DATA1_dt5(iTYPEK1 As Integer, dt5_6 As DataTable) As DataTable
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim vSCORESTAGE As String = TIMS.GetListValue(rblSCORESTAGE)
        Dim PMS1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}, {"ORGKIND2", "W"}}

        'Dim dt5_6 As DataTable = GET_SCORING2_dt5_6(vSCORINGID)
        Dim SC_WJ As String = If(dt5_6.Rows.Count > 1, Convert.ToString(dt5_6.Rows(1)("VALUEFD")), "")
        Dim SC_WI As String = If(dt5_6.Rows.Count > 2, Convert.ToString(dt5_6.Rows(2)("VALUEFD")), "")
        Dim SC_WH As String = If(dt5_6.Rows.Count > 3, Convert.ToString(dt5_6.Rows(3)("VALUEFD")), "")
        Dim SC_WG As String = If(dt5_6.Rows.Count > 4, Convert.ToString(dt5_6.Rows(4)("VALUEFD")), "")
        Dim SC_WF As String = If(dt5_6.Rows.Count > 5, Convert.ToString(dt5_6.Rows(5)("VALUEFD")), "")
        'sSql &= " DECLARE @SCORINGID NVarChar(21)='2025-01-2024-1-2024-2';/*1*/" & vbCrLf
        'sSql &= " DECLARE @TPLANID NVarChar(2)='28';/*2*/" & vbCrLf
        'sSql &= " DECLARE @ORGKIND2 NVarChar(1)='W';/*3*/" & vbCrLf

        Dim SSQL As String = ""
        Select Case iTYPEK1
            Case 1
                SSQL &= " WITH WO1 AS (SELECT ORGID,COMIDNO,ORGKIND1,ORGKIND2,ORGNAME FROM ORG_ORGINFO oo WITH(NOLOCK) WHERE oo.ORGKIND2=@ORGKIND2 AND oo.ORGKIND1 IN (7,8,9))" & vbCrLf '全國性工會
            Case 2
                SSQL &= "  WITH WO1 AS (SELECT ORGID,COMIDNO,ORGKIND1,ORGKIND2,ORGNAME FROM ORG_ORGINFO oo WITH(NOLOCK) WHERE oo.ORGKIND2=@ORGKIND2 AND oo.ORGKIND1 IN (10,11))" & vbCrLf '縣市級工會
        End Select
        SSQL &= $" ,WJ AS (SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID FROM dbo.ORG_SCORING2 a where CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)='{SC_WJ}')"
        SSQL &= $" ,WI AS (SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID FROM dbo.ORG_SCORING2 a where CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)='{SC_WI}')"
        SSQL &= $" ,WH AS (SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID FROM dbo.ORG_SCORING2 a where CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)='{SC_WH}')"
        SSQL &= $" ,WG AS (SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID FROM dbo.ORG_SCORING2 a where CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)='{SC_WG}')"
        SSQL &= $" ,WF AS (SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID FROM dbo.ORG_SCORING2 a where CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)='{SC_WF}')"
        SSQL &= " ,WC1 AS (SELECT a.YEARS,a.MONTHS,a.YEARS1,a.HALFYEAR1,a.YEARS2,a.HALFYEAR2,a.COMIDNO,a.ORGID,a.DISTID,a.TPLANID" & vbCrLf
        SSQL &= " ,a.SUBTOTAL,a.IMPLEVEL_1,a.MINISTERSUB,a.MINISTERLEVEL,A.SCORE4_1_2,a.SCORE4_1 ,a.RLEVEL_2,a.BRANCHPNT,a.MINISTERADD,a.DEPTADD" & vbCrLf
        SSQL &= " ,dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2) CLASSCNT2B" & vbCrLf
        SSQL &= " FROM dbo.ORG_SCORING2 a WHERE a.TPLANID=@TPLANID AND CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID)" & vbCrLf
        SSQL &= " ,WC2 AS (SELECT a.YEARS,a.MONTHS,a.YEARS1,a.HALFYEAR1,a.YEARS2,a.HALFYEAR2,a.COMIDNO,a.ORGID,a.DISTID,a.TPLANID" & vbCrLf
        SSQL &= " ,op.MASTERNAME FROM WC1 a JOIN dbo.VIEW_ORGPLANINFO op ON op.ORGID=a.ORGID AND op.TPLANID=a.TPLANID AND op.DISTID=a.DISTID AND op.YEARS=a.YEARS)" & vbCrLf
        SSQL &= " ,WC3 AS (SELECT a.YEARS,a.MONTHS,a.YEARS1,a.HALFYEAR1,a.YEARS2,a.HALFYEAR2,a.COMIDNO,a.ORGID,a.DISTID,a.TPLANID" & vbCrLf
        SSQL &= " ,vo.MASTERNAME FROM WC1 a JOIN dbo.V_ORGINFO vo ON vo.COMIDNO=a.COMIDNO)" & vbCrLf
        'SSQL &= " ,WC4 AS (SELECT a.YEARS,a.MONTHS,a.YEARS1,a.HALFYEAR1,a.YEARS2,a.HALFYEAR2,a.COMIDNO,a.ORGID,a.DISTID,a.TPLANID" & vbCrLf
        'SSQL &= " ,vc.SUBTOTALA,vc.SUBTOTALB,vc.SUBTOTALC FROM WC1 a JOIN dbo.V_SCORING2_MIN vc ON A.YEARS=vc.YEARS AND A.MONTHS=vc.MONTHS AND A.TPLANID=vc.TPLANID AND A.DISTID=vc.DISTID)" & vbCrLf

        Dim SSQL_ORDERBY As String = ""
        Select Case iTYPEK1
            Case 1
                SSQL &= " SELECT ROW_NUMBER() OVER (PARTITION BY oo.ORGNAME ORDER BY oo.ORGNAME DESC) RN,DENSE_RANK() OVER (ORDER BY oo.ORGNAME) DRK" & vbCrLf
                SSQL_ORDERBY = " ORDER BY oo.ORGNAME,a.DISTID" & vbCrLf',oo.ORGKIND1
            Case 2
                SSQL &= " SELECT 1 RN,DENSE_RANK() OVER (ORDER BY a.DISTID,a.SUBTOTAL DESC,oo.ORGNAME) DRK" & vbCrLf
                SSQL_ORDERBY = " ORDER BY a.DISTID,a.SUBTOTAL DESC,oo.ORGNAME" & vbCrLf ',oo.ORGNAME,oo.ORGKIND1
        End Select

        SSQL &= " ,oo.COMIDNO,oo.ORGNAME,oo.ORGKIND2,oo.ORGKIND1,a.DISTID, dd.DISTNAME3" & vbCrLf
        'SSQL &= " ,(SELECT ko.ORGTYPE FROM dbo.VIEW_ORGTYPE1 ko WHERE ko.ORGTYPEID1=oo.ORGKIND1) ORGKIND1_N" & vbCrLf
        SSQL &= " ,ko.ORGTYPE ORGKIND1_N" & vbCrLf
        ',OP.MASTERNAME
        SSQL &= " ,(SELECT op.MASTERNAME FROM WC2 op WHERE op.ORGID=a.ORGID AND op.TPLANID=a.TPLANID AND op.DISTID=a.DISTID AND op.YEARS=a.YEARS AND op.MONTHS=a.MONTHS) MASTERNAME" & vbCrLf
        SSQL &= " ,(SELECT x.MASTERNAME FROM WC3 x WHERE x.ORGID=a.ORGID AND x.TPLANID=a.TPLANID AND x.DISTID=a.DISTID AND x.YEARS=a.YEARS AND x.MONTHS=a.MONTHS) MASTERNAME1" & vbCrLf
        SSQL &= " ,a.BRANCHPNT,a.SUBTOTAL,a.IMPLEVEL_1,a.MINISTERADD,a.MINISTERSUB,a.MINISTERLEVEL,a.DEPTADD,A.SCORE4_1_2,a.RLEVEL_2,a.SCORE4_1" & vbCrLf
        'SSQL &= " ,format(CASE a.IMPLEVEL_1 WHEN 'B' THEN dbo.FN_GET_LEVELUPADD('A',A.SUBTOTAL,A.TPLANID,A.DISTID,A.YEARS,A.MONTHS)" & vbCrLf
        'SSQL &= " WHEN 'C' THEN dbo.FN_GET_LEVELUPADD('B',A.SUBTOTAL,A.TPLANID,A.DISTID,A.YEARS,A.MONTHS)" & vbCrLf
        'SSQL &= " WHEN 'D' THEN dbo.FN_GET_LEVELUPADD('C',A.SUBTOTAL,A.TPLANID,A.DISTID,A.YEARS,A.MONTHS) END,'N1') LEVELUPADD" & vbCrLf
        'SSQL &= " ,(SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD) FROM WJ WHERE ORGID=a.ORGID AND DISTID=a.DISTID AND TPLANID=a.TPLANID) RLEVEL_2J" & vbCrLf
        'SSQL &= " ,(SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD) FROM WI WHERE ORGID=a.ORGID AND DISTID=a.DISTID AND TPLANID=a.TPLANID) RLEVEL_2I" & vbCrLf
        'SSQL &= " ,(SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD) FROM WH WHERE ORGID=a.ORGID AND DISTID=a.DISTID AND TPLANID=a.TPLANID) RLEVEL_2H" & vbCrLf
        'SSQL &= " ,(SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD) FROM WG WHERE ORGID=a.ORGID AND DISTID=a.DISTID AND TPLANID=a.TPLANID) RLEVEL_2G" & vbCrLf
        'SSQL &= " ,(SELECT RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD) FROM WF WHERE ORGID=a.ORGID AND DISTID=a.DISTID AND TPLANID=a.TPLANID) RLEVEL_2F" & vbCrLf
        SSQL &= " ,format(CASE a.IMPLEVEL_1 WHEN 'B' THEN dbo.FN_GET_LEVELUP('A',A.SUBTOTAL,co.SUBTOTALA,co.SUBTOTALB,co.SUBTOTALC)" & vbCrLf
        SSQL &= " WHEN 'C' THEN dbo.FN_GET_LEVELUP('B',A.SUBTOTAL,co.SUBTOTALA,co.SUBTOTALB,co.SUBTOTALC)" & vbCrLf
        SSQL &= " WHEN 'D' THEN dbo.FN_GET_LEVELUP('C',A.SUBTOTAL,co.SUBTOTALA,co.SUBTOTALB,co.SUBTOTALC) END,'N1') LEVELUPADD" & vbCrLf
        SSQL &= " ,dbo.FN_GET_MEMOV4(1,A.IMPLEVEL_1,A.SUBTOTAL,A.SCORE4_1,co.SUBTOTALA,co.SUBTOTALB,co.SUBTOTALC) MEMOV41" & vbCrLf
        SSQL &= " ,dbo.FN_GET_MEMOV4(2,A.IMPLEVEL_1,A.SUBTOTAL,A.SCORE4_1,co.SUBTOTALA,co.SUBTOTALB,co.SUBTOTALC) MEMOV42" & vbCrLf
        SSQL &= " ,dbo.FN_GET_MEMOV4(3,A.IMPLEVEL_1,A.SUBTOTAL,A.SCORE4_1,co.SUBTOTALA,co.SUBTOTALB,co.SUBTOTALC) MEMOV43" & vbCrLf
        SSQL &= " ,WJ.RLEVEL_2+dbo.FN_RTN_RLEVEL2ADD(WJ.MINISTERADD,WJ.DEPTADD) RLEVEL_2J" & vbCrLf
        SSQL &= " ,WI.RLEVEL_2+dbo.FN_RTN_RLEVEL2ADD(WI.MINISTERADD,WI.DEPTADD) RLEVEL_2I" & vbCrLf
        SSQL &= " ,WH.RLEVEL_2+dbo.FN_RTN_RLEVEL2ADD(WH.MINISTERADD,WH.DEPTADD) RLEVEL_2H" & vbCrLf
        SSQL &= " ,WG.RLEVEL_2+dbo.FN_RTN_RLEVEL2ADD(WG.MINISTERADD,WG.DEPTADD) RLEVEL_2G" & vbCrLf
        SSQL &= " ,WF.RLEVEL_2+dbo.FN_RTN_RLEVEL2ADD(WF.MINISTERADD,WF.DEPTADD) RLEVEL_2F" & vbCrLf
        'SSQL &= " ,WJ.RLEVEL_2+CASE WHEN WJ.RLEVEL_2 IS NOT NULL AND (ISNULL(dbo.FN_CAST2FLOAT(WJ.BRANCHPNT),0)+ISNULL(dbo.FN_CAST2FLOAT(WJ.MINISTERADD),0)+ISNULL(dbo.FN_CAST2FLOAT(WJ.DEPTADD),0))>0 THEN '*' else '' END RLEVEL_2J" & vbCrLf
        'SSQL &= " ,WI.RLEVEL_2+CASE WHEN WI.RLEVEL_2 IS NOT NULL AND (ISNULL(dbo.FN_CAST2FLOAT(WI.BRANCHPNT),0)+ISNULL(dbo.FN_CAST2FLOAT(WI.MINISTERADD),0)+ISNULL(dbo.FN_CAST2FLOAT(WI.DEPTADD),0))>0 THEN '*' else '' END RLEVEL_2I" & vbCrLf
        'SSQL &= " ,WH.RLEVEL_2+CASE WHEN WH.RLEVEL_2 IS NOT NULL AND (ISNULL(dbo.FN_CAST2FLOAT(WH.BRANCHPNT),0)+ISNULL(dbo.FN_CAST2FLOAT(WH.MINISTERADD),0)+ISNULL(dbo.FN_CAST2FLOAT(WH.DEPTADD),0))>0 THEN '*' else '' END RLEVEL_2H" & vbCrLf
        'SSQL &= " ,WG.RLEVEL_2+CASE WHEN WG.RLEVEL_2 IS NOT NULL AND (ISNULL(dbo.FN_CAST2FLOAT(WG.BRANCHPNT),0)+ISNULL(dbo.FN_CAST2FLOAT(WG.MINISTERADD),0)+ISNULL(dbo.FN_CAST2FLOAT(WG.DEPTADD),0))>0 THEN '*' else '' END RLEVEL_2G" & vbCrLf
        'SSQL &= " ,WF.RLEVEL_2+CASE WHEN WF.RLEVEL_2 IS NOT NULL AND (ISNULL(dbo.FN_CAST2FLOAT(WF.BRANCHPNT),0)+ISNULL(dbo.FN_CAST2FLOAT(WF.MINISTERADD),0)+ISNULL(dbo.FN_CAST2FLOAT(WF.DEPTADD),0))>0 THEN '*' else '' END RLEVEL_2F" & vbCrLf
        SSQL &= " FROM WC1 a" & vbCrLf
        SSQL &= " JOIN WO1 oo ON oo.OrgID=a.OrgID" & vbCrLf
        SSQL &= " JOIN dbo.V_DISTRICT dd on dd.DISTID=a.DISTID" & vbCrLf
        SSQL &= " JOIN dbo.V_SCORING2_MIN co ON co.YEARS=a.YEARS AND co.MONTHS=a.MONTHS AND co.TPLANID=a.TPLANID AND co.DISTID=a.DISTID" & vbCrLf
        'SSQL &= " JOIN dbo.VIEW_ORGPLANINFO op on op.ORGID=a.ORGID AND op.TPLANID=a.TPLANID AND op.DISTID=a.DISTID AND op.YEARS=a.YEARS" & vbCrLf
        SSQL &= " LEFT JOIN dbo.VIEW_ORGTYPE1 ko ON ko.ORGTYPEID1=oo.ORGKIND1" & vbCrLf
        SSQL &= " LEFT JOIN WJ ON WJ.ORGID=a.ORGID AND WJ.DISTID=a.DISTID AND WJ.TPLANID=a.TPLANID" & vbCrLf
        SSQL &= " LEFT JOIN WI ON WI.ORGID=a.ORGID AND WI.DISTID=a.DISTID AND WI.TPLANID=a.TPLANID" & vbCrLf
        SSQL &= " LEFT JOIN WH ON WH.ORGID=a.ORGID AND WH.DISTID=a.DISTID AND WH.TPLANID=a.TPLANID" & vbCrLf
        SSQL &= " LEFT JOIN WG ON WG.ORGID=a.ORGID AND WG.DISTID=a.DISTID AND WG.TPLANID=a.TPLANID" & vbCrLf
        SSQL &= " LEFT JOIN WF ON WF.ORGID=a.ORGID AND WF.DISTID=a.DISTID AND WF.TPLANID=a.TPLANID" & vbCrLf
        'SSQL &= " WHERE CONCAT(a.YEARS ,'-',a.MONTHS,'-',a.YEARS1 ,'-',a.HALFYEAR1,'-',a.YEARS2 ,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
        'SSQL &= " LEFT JOIN WC4 co ON co.ORGID=a.ORGID AND co.DISTID=a.DISTID AND co.TPLANID=a.TPLANID" & vbCrLf
        SSQL &= $" WHERE a.CLASSCNT2B>0 {SSQL_ORDERBY}" & vbCrLf
        'SELECT * FROM VIEW_ORGTYPE1 WHERE ORGTYPEID1  IN (7,8,9)'全國性工會,,'SELECT * FROM VIEW_ORGTYPE1 WHERE ORGTYPEID1  IN (10,11)'縣市級工會
        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, String.Concat("--iTYPEK1: ", iTYPEK1, vbCrLf, TIMS.GetMyValue5(PMS1), vbCrLf, "--##CO_01_005:", vbCrLf, SSQL))
        Dim dt As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt
    End Function
    Public Sub XLS_SHEET1_5(ep As ExcelPackage, ws As ExcelWorksheet, dtXls2 As DataTable, dt5_6 As DataTable, dtBlack As DataTable, hPP As Hashtable)
        Dim SF_TITLE2 As String = TIMS.GetMyValue2(hPP, "SF_TITLE2")
        Dim S_COLNM1 As String = TIMS.GetMyValue2(hPP, "S_COLNM1")
        Dim str_wsname2 As String = TIMS.GetMyValue2(hPP, "str_wsname2") '"五分署名單(縣市級工會)"
        Dim vSCORESTAGE As String = TIMS.GetMyValue2(hPP, "vSCORESTAGE")
        Dim cellsCOLSPNumF2 As String = TIMS.GetMyValue2(hPP, "cellsCOLSPNumF2")

        'Dim str_wsname2 As String = "五分署名單(縣市級工會)"
        'ws = ep.Workbook.Worksheets(1)
        'Dim ep As New ExcelPackage() 'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
        ws.Name = str_wsname2
        'Dim END_COL_NM As String = "V"
        'Dim cellsCOLSPNumF2 As String = String.Concat("A5:", END_COL_NM, "{0}") '(畫格子使用)

        'Dim S_COLNM1 As String = "A1:V1"
        ws.Cells(S_COLNM1).Value = SF_TITLE2
        ws.Cells(S_COLNM1).Style.Font.Bold = True
        ws.Cells(S_COLNM1).Style.Font.Size = fontSize16s

        Dim idxStr1 As String = 3
        ws.Cells("K" & idxStr1).Value = String.Concat(dt5_6.Rows(0)("NEXT_YMN2"), "等級")
        ws.Cells("J" & idxStr1).Value = If(dt5_6.Rows.Count > 1, String.Concat(dt5_6.Rows(1)("NEXT_YMN2"), "等級"), "-")
        ws.Cells("I" & idxStr1).Value = If(dt5_6.Rows.Count > 2, String.Concat(dt5_6.Rows(2)("NEXT_YMN2"), "等級"), "-")
        ws.Cells("H" & idxStr1).Value = If(dt5_6.Rows.Count > 3, String.Concat(dt5_6.Rows(3)("NEXT_YMN2"), "等級"), "-")
        ws.Cells("G" & idxStr1).Value = If(dt5_6.Rows.Count > 4, String.Concat(dt5_6.Rows(4)("NEXT_YMN2"), "等級"), "-")
        ws.Cells("F" & idxStr1).Value = If(dt5_6.Rows.Count > 5, String.Concat(dt5_6.Rows(5)("NEXT_YMN2"), "等級"), "-")

        'ws.Cells("F2").Value = vROC_YMD_NOW
        Dim vRN As Integer = 1
        idxStr1 = 5 'Dim idxStr1 As Integer = 5
        Dim iROWNUM As Integer = 0
        For Each dr As DataRow In dtXls2.Rows
            iROWNUM += 1
            ws.Cells("A" & idxStr1).Value = Convert.ToString(dr("DRK")) 'iROWNUM
            ws.Cells("B" & idxStr1).Value = Convert.ToString(dr("ORGNAME"))
            ws.Cells("C" & idxStr1).Value = Convert.ToString(dr("ORGKIND1_N"))
            ws.Cells("D" & idxStr1).Value = Convert.ToString(dr("DISTNAME3"))
            Dim vMASTERNAME As String = $"{dr("MASTERNAME")}"
            If vMASTERNAME = "" Then vMASTERNAME = $"{dr("MASTERNAME1")}"
            ws.Cells("E" & idxStr1).Value = vMASTERNAME

            ws.Cells("K" & idxStr1).Value = TIMS.GetValue3(dr("IMPLEVEL_1"), "-") '初擬等級
            ws.Cells("L" & idxStr1).Value = TIMS.VAL1(dr("SCORE4_1"), Nothing) '初審＞分署加分
            ws.Cells("M" & idxStr1).Value = TIMS.VAL1(dr("SUBTOTAL"), Nothing) '初審＞分署小計
            Dim oLEVELUPADD As Object = TIMS.VAL1(dr("LEVELUPADD"), Nothing) '升1級須加分數
            If oLEVELUPADD IsNot Nothing AndAlso oLEVELUPADD < 0 Then oLEVELUPADD = Nothing '升1級須加分數(負數有誤)
            ws.Cells("N" & idxStr1).Value = oLEVELUPADD '升1級須加分數

            Select Case vSCORESTAGE
                Case "02", "03"
                    Dim oMINISTERADD As Object = TIMS.VAL1(dr("MINISTERADD"), Nothing)
                    If oMINISTERADD IsNot Nothing Then
                        If Not TIMS.VAL1_Equal(oMINISTERADD, 0) Then ws.Cells("O" & idxStr1).Value = "Ⅴ" '本部加分
                        ws.Cells("P" & idxStr1).Value = TIMS.VAL1(dr("MINISTERADD"), Nothing) '本部加分
                    End If
                    ws.Cells("Q" & idxStr1).Value = TIMS.VAL1(dr("MINISTERSUB"), Nothing) '本部加分後總分
                    ws.Cells("R" & idxStr1).Value = TIMS.GetValue3(dr("MINISTERLEVEL"), Nothing) '本部加分後等級
            End Select
            If vSCORESTAGE = "03" Then
                ws.Cells("S" & idxStr1).Value = TIMS.VAL1(dr("DEPTADD"), Nothing) '本署加分
                ws.Cells("T" & idxStr1).Value = TIMS.VAL1(dr("SCORE4_1_2"), Nothing) '本署加分後總分
                ws.Cells("U" & idxStr1).Value = TIMS.GetValue3(dr("RLEVEL_2"), Nothing) '本署加分後等級
            End If

            Dim vMEMOV4 As String = ""
            Dim vMEMOV41 As String = $"{dr("MEMOV41")}"
            Dim vMEMOV42 As String = $"{dr("MEMOV42")}"
            Dim vMEMOV43 As String = $"{dr("MEMOV43")}"
            If vMEMOV41 <> "" Then vMEMOV4 &= $"{If(vMEMOV4 <> "", ";", "")}{vMEMOV41}"
            If vMEMOV42 <> "" Then vMEMOV4 &= $"{If(vMEMOV4 <> "", ";", "")}{vMEMOV42}"
            If vMEMOV43 <> "" Then vMEMOV4 &= $"{If(vMEMOV4 <> "", ";", "")}{vMEMOV43}"
            If TIMS.dtHaveDATA(dtBlack) Then
                Dim ff3 As String = $"COMIDNO='{dr("COMIDNO")}' AND DISTID='{dr("DISTID")}'"
                If dtBlack.Select(ff3).Length > 0 Then vMEMOV4 &= $"{If(vMEMOV4 <> "", ";", "")}違反計畫，停權處分中"
            End If
            ws.Cells("V" & idxStr1).Value = $"{vMEMOV4}" '備註
            ws.Cells("V" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left

            ws.Cells("A" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("D" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("E" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("K" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("L" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("M" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("N" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("O" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("P" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("Q" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("R" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("S" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("T" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("U" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            ws.Cells("J" & idxStr1).Value = TIMS.GetValue3(dr("RLEVEL_2J"), "-")
            ws.Cells("I" & idxStr1).Value = TIMS.GetValue3(dr("RLEVEL_2I"), "-")
            ws.Cells("H" & idxStr1).Value = TIMS.GetValue3(dr("RLEVEL_2H"), "-")
            ws.Cells("G" & idxStr1).Value = TIMS.GetValue3(dr("RLEVEL_2G"), "-")
            ws.Cells("F" & idxStr1).Value = TIMS.GetValue3(dr("RLEVEL_2F"), "-")
            ws.Cells("J" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("I" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("H" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("G" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            ws.Cells("F" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

            ws.Cells("L" & idxStr1).Style.Numberformat.Format = "#0.0"
            ws.Cells("M" & idxStr1).Style.Numberformat.Format = "#0.0"
            ws.Cells("N" & idxStr1).Style.Numberformat.Format = "#0.0"
            ws.Cells("P" & idxStr1).Style.Numberformat.Format = "#0.0"
            ws.Cells("Q" & idxStr1).Style.Numberformat.Format = "#0.0"
            ws.Cells("S" & idxStr1).Style.Numberformat.Format = "#0.0"
            ws.Cells("T" & idxStr1).Style.Numberformat.Format = "#0.0"

            If dr("RN") > 1 OrElse dr("RN") > vRN Then
                vRN = dr("RN")
            ElseIf dr("RN") = 1 AndAlso vRN > 1 Then
                'Dim idx1 As Integer = idxStr1 - TIMS.VAL1(dr("RN")) + 1
                Dim idx2b As Integer = idxStr1 - 1
                Dim idx1b As Integer = idx2b - vRN + 1
                ws.Cells(String.Format("A{0}:A{1}", idx1b, idx2b)).Merge = True
                ws.Cells(String.Format("B{0}:B{1}", idx1b, idx2b)).Merge = True
                ws.Cells(String.Format("C{0}:C{1}", idx1b, idx2b)).Merge = True
                vRN = 1
            End If
            idxStr1 += 1
        Next
        If vRN > 1 Then
            'Dim idx1 As Integer = idxStr1 - TIMS.VAL1(dr("RN")) + 1
            Dim idx2b As Integer = idxStr1 - 1
            Dim idx1b As Integer = idx2b - vRN + 1
            ws.Cells(String.Format("A{0}:A{1}", idx1b, idx2b)).Merge = True
            ws.Cells(String.Format("B{0}:B{1}", idx1b, idx2b)).Merge = True
            ws.Cells(String.Format("C{0}:C{1}", idx1b, idx2b)).Merge = True
            vRN = 1
        End If

        idxStr1 -= 1 '(畫線)
        Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr1))
            With exlRow3X
                .Style.Font.Name = fontName
                .Style.Font.Size = fontSize14s 'FontSize
                .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                .AutoFitColumns(25.0, 250.0)
            End With
            TIMS.SetCellBorder(exlRow3X)
        End Using

        'ws.Column("Q:V").Width = Convert.ToDouble(10)/ws.Column("S").Width = Convert.ToDouble(6) 24.35 : 10 
        ws.Column(ws.Cells(String.Format("Q2:Q{0}", idxStr1)).Start.Column).Width = Convert.ToDouble(11)
        ws.Column(ws.Cells(String.Format("R2:R{0}", idxStr1)).Start.Column).Width = Convert.ToDouble(11)
        ws.Column(ws.Cells(String.Format("S2:S{0}", idxStr1)).Start.Column).Width = Convert.ToDouble(7)
        ws.Column(ws.Cells(String.Format("T2:T{0}", idxStr1)).Start.Column).Width = Convert.ToDouble(11)
        ws.Column(ws.Cells(String.Format("U2:U{0}", idxStr1)).Start.Column).Width = Convert.ToDouble(11)
        ws.Column(ws.Cells(String.Format("V2:V{0}", idxStr1)).Start.Column).Width = Convert.ToDouble(66)

        idxStr1 += 2
        Using exlRow3X As ExcelRange = ws.Cells(String.Format("A{0}:L{0}", idxStr1))
            With exlRow3X
                .Style.Font.Name = fontName
                .Style.Font.Size = fontSize12s 'FontSize
                .Style.Font.Bold = True
                .Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                .Merge = True
                .Value = "※註：等級符號說明 *部有加分,#署有加分"
            End With
            ws.Row(idxStr1).Height = Convert.ToDouble(22)
        End Using

        '設定工作表的顯示比例為 70%  worksheet.View.Zoom=70 無法運行 修正為 ws.View.ZoomScale=70 才可運行
        ws.View.ZoomScale = 88

        '設定貨幣格式，小數位數為 0
        'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format="$#,##0" ' 美元符號，您可以根據需要更改
        'ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width=33
    End Sub
    ''' <summary>
    ''' (自主)各等級分配比率
    ''' </summary>
    Sub ExportXlsStd28_5()
        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const cst_SampleXLS As String = "~\CO\01\sampleC01005d2.xlsx" '& cst_files_ext 'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If

        Dim strErrmsg As String = ""
        Dim sFileName As String = $"{Cst_FileSavePath}{TIMS.GetDateNo()}.xlsx" '複製一份(Sample)
        Dim sMyFile1 As String = Server.MapPath(sFileName) '複製一份(Sample)
        Try
            IO.File.Copy(Server.MapPath(cst_SampleXLS), sMyFile1, True)
        Catch ex As Exception
            strErrmsg = $"目錄名稱或磁碟區標籤語法錯誤!!!{vbCrLf} (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉){vbCrLf}{ex.Message}{vbCrLf}"
            Common.MessageBox(Me, strErrmsg)
            TIMS.LOG.Error(ex.Message, ex)
            Return 'Exit Sub
        End Try

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim drSC2 As DataRow = GET_SCORING2_R(vSCORINGID)
        If drSC2 Is Nothing Then
            Common.MessageBox(Me, "查無 匯出資料!")
            Exit Sub
        End If

        Dim dt5_6 As DataTable = GET_SCORING2_dt5_6(vSCORINGID)
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt5(1, dt5_6)
        Dim dtXls2 As DataTable = SEARCH_DATA1_dt5(2, dt5_6)
        Dim dtBlack As DataTable = GET_BLACKLIST_dt1()
        If TIMS.dtNODATA(dtXls1) AndAlso TIMS.dtNODATA(dtXls2) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim vROC_YMD_NOW As String = TIMS.GetROCTWDate(Now)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim vSCORESTAGE As String = TIMS.GetListValue(rblSCORESTAGE)
        'rblSCORESTAGE {"SCORESTAGE", vSCORESTAGE}01：部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        Dim txt_SCORESTAGE As String = TIMS.GetListText(rblSCORESTAGE)
        Dim vNEXT_YMN As String = Convert.ToString(drSC2("NEXT_YMN"))
        'OOO年度<申請階段>提升勞工自主學習計畫訓練單位審查計分表初擬等級及分數(全國性工會)  
        Dim SF_TITX As String = String.Concat(vNEXT_YMN, "提升勞工自主學習計畫訓練單位審查計分表初擬等級及分數")
        Dim SF_TITLE1 As String = String.Concat(SF_TITX, "(全國性工會) (", txt_SCORESTAGE, ")")
        '114年度下半年提升勞工自主學習計畫訓練單位審查計分表初擬等級及分數(縣市級工會)
        Dim SF_TITLE2 As String = String.Concat(SF_TITX, "(縣市級工會) (", txt_SCORESTAGE, ")")
        Dim s_FILENAME1 As String = String.Concat(SF_TITX, "x", txt_SCORESTAGE, "x", TIMS.GetDateNo2(3))
        Dim fg_RespWriteEnd As Boolean = False
        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial, ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            'Dim file1 As New FileInfo(filePath1) 'Dim ndt As DateTime = Now
            '開檔
            Using fs1 As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim ep As New ExcelPackage(fs1)

                Dim ws0 As ExcelWorksheet = ep.Workbook.Worksheets(0)
                Dim ws1 As ExcelWorksheet = ep.Workbook.Worksheets(1)
                'Dim ep As New ExcelPackage() 'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)

                Dim S_COLNM1 As String = "A1:V1"
                Dim END_COL_NM As String = "V"
                Dim cellsCOLSPNumF2 As String = String.Concat("A5:", END_COL_NM, "{0}") '(畫格子使用)

                'Dim str_wsname1 As String = "五分署名單(全國性工會)"
                Dim myPMS1 As New Hashtable From {
                    {"SF_TITLE2", SF_TITLE1},
                    {"S_COLNM1", S_COLNM1},
                    {"str_wsname2", "五分署名單(全國性工會)"},
                    {"vSCORESTAGE", vSCORESTAGE},
                    {"cellsCOLSPNumF2", cellsCOLSPNumF2}
                }
                Call XLS_SHEET1_5(ep, ws0, dtXls1, dt5_6, dtBlack, myPMS1)

                'Dim str_wsname2 As String = "五分署名單(縣市級工會)"
                Dim myPMS2 As New Hashtable From {
                    {"SF_TITLE2", SF_TITLE2},
                    {"S_COLNM1", S_COLNM1},
                    {"str_wsname2", "五分署名單(縣市級工會)"},
                    {"vSCORESTAGE", vSCORESTAGE},
                    {"cellsCOLSPNumF2", cellsCOLSPNumF2}
                }
                Call XLS_SHEET1_5(ep, ws1, dtXls2, dt5_6, dtBlack, myPMS2)

                Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
                Select Case V_ExpType
                    Case "EXCEL"
                        TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case "ODS"
                        TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case Else
                        Dim s_log1 As String = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                        Common.MessageBox(Me, s_log1)
                        Return ' Exit Sub
                End Select
            End Using
            Call TIMS.MyFileDelete(sMyFile1)
            If fg_RespWriteEnd Then TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End SyncLock

        '刪除Temp中的資料 'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If

    End Sub
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        ExportXlsStd28_5()
    End Sub

    Sub UPDATE_ORG_SCORING2_FIRSTCHK_Y()
        Dim V_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim V_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        Dim V_OrgPlanKind As String = TIMS.GetListValue(OrgPlanKind)
        'DECLARE @TPLANID VARCHAR(4)='28';DECLARE @DISTID  VARCHAR(4)='001';DECLARE @ORGKIND2  VARCHAR(3)='G';DECLARE @SCORINGID  VARCHAR(22)='2026-01-2025-1-2025-2';
        Dim PMS_S1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"DISTID", V_ddlDISTID}, {"ORGKIND2", V_OrgPlanKind}, {"SCORINGID", V_ddlSCORING}}
        Dim SQL_S1 As String = "
SELECT A.OSID2
FROM dbo.ORG_SCORING2 a WITH(NOLOCK)
JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID
WHERE a.FIRSTCHK IS NULL AND A.IMPLEVEL_1 IS NOT NULL AND A.SUBTOTAL IS NOT NULL AND A.RLEVEL_2 IS NOT NULL AND A.MINISTERLEVEL IS NOT NULL AND A.IMODIFYACCT IS NOT NULL
AND a.TPLANID=@TPLANID AND a.DISTID=@DISTID AND a.DISTID=@DISTID AND oo.ORGKIND2=@ORGKIND2
AND CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID
"
        Dim dt1 As DataTable = DbAccess.GetDataTable(SQL_S1, objconn, PMS_S1)
        If TIMS.dtNODATA(dt1) Then Return

        Dim PMS_U1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"DISTID", V_ddlDISTID}, {"ORGKIND2", V_OrgPlanKind}, {"SCORINGID", V_ddlSCORING}}
        Dim SQL_U1 As String = "
UPDATE ORG_SCORING2 SET FIRSTCHK='Y'
FROM ORG_SCORING2 a
JOIN ORG_ORGINFO oo ON oo.OrgID=a.OrgID
WHERE a.FIRSTCHK IS NULL AND A.IMPLEVEL_1 IS NOT NULL AND A.SUBTOTAL IS NOT NULL AND A.RLEVEL_2 IS NOT NULL AND A.MINISTERLEVEL IS NOT NULL AND A.IMODIFYACCT IS NOT NULL
AND a.TPLANID=@TPLANID AND a.DISTID=@DISTID AND a.DISTID=@DISTID AND oo.ORGKIND2=@ORGKIND2
AND CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID
"
        TIMS.ExecuteNonQuery(SQL_U1, objconn, PMS_U1)
    End Sub

End Class
