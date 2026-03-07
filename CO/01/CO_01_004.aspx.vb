Imports System.IO
Imports System.Text.RegularExpressions
Imports OfficeOpenXml

Public Class CO_01_004
    Inherits AuthBasePage 'System.Web.UI.Page

    '啟動審查計分排程 ORG_SCORING2 '排程[Co_OrgScoring] Co_OrgScoring.exe.config '\SVN\WDAIIP\SRC\Batch\Co_OrgScoring '\SVN\WDAIIP\SRC\Batch\Co_OrgScoring\bin\Debug

    ''' <summary>匯入等級/分數(停用)</summary>
    Dim fg_trImport1_NoUse As Boolean = True
    'Lab_SUSPENDED_msg1
    Const cst_SUSPENDED_msgFM1 As String = "此單位因有{0}班停班經認列屬「不可抗力因素」，將不列入核定總班數計算。"
    'Const cst_s_LEVEL1s As String="A,B,C,D"
    Const cst_col1_COMIDNO As Integer = 0 '統一編號
    Const cst_col1_SCORE1 As Integer = 1 '匯入分數
    Const cst_col1_LEVEL1 As Integer = 2 '匯入等級

    Const cst_col2_COMIDNO As Integer = 0 '統一編號
    Const cst_col2_SCORE4_1 As Integer = 1 '分署加分

    Const cst_col3_COMIDNO As Integer = 0 '統一編號
    Const cst_col3_LEVEL1 As Integer = 1 '匯入等級

    'ORG_SCORING2'統一編號    匯入分數	匯入等級
    Dim s_col_COMIDNO As String = "" '統一編號
    Dim s_col_SCORE1 As String = "" '匯入分數
    Dim s_col_LEVEL1 As String = "" '匯入等級
    'Dim s_col_COMIDNO As String="" '統一編號
    Dim s_col_SCORE4_1 As String = "" '分署加分

    'Const CST_CAPIDX_UP As String = "可提升等級!"
    Const CST_CAPIDX_DOWN As String = "已超出等級核配比率!"
    'Const CST_CAPIDX_WARN As String = "注意!可能超出比率!"
    Const CST_CAPIDX_WARN2 As String = "此為該等級最後序位!"

    Const CST_NON_REVIEWSCORE As String = "（非）審查計分表(初審)時間" '(Non-)Review Score Sheet (Preliminary Review) Time
    Dim GB_FG_CAN_SAVE_1 As Boolean = False '(判斷)審查計分表(初審)時間 

    Const cst_tit1 As String = "複審選通過，初審資料鎖定!"
    Const cst_tit2 As String = "請填寫數字，至多加3分，輸入後按Tab鍵 或失去焦點會自動計算"
    Const cst_tit3 As String = "（非）審查計分表(初審)時間,資料鎖定!"

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1
        '產投/非產投判斷 'autorecsubtotal 
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then  Me.LabTMID.Text="訓練業別"
        '測試環境啟用
        'Dim flag_chktest As Boolean=TIMS.sUtl_ChkTest()
        'btnExp3.Visible=flag_chktest '測試環境啟用 'False
        'trImport1.Visible=False
        'Dim FG_CAN_SAVE_1 As Boolean = CHK_TTQSQUERY_2()
        GB_FG_CAN_SAVE_1 = CHK_TTQSQUERY_2()
        trImport1.Visible = GB_FG_CAN_SAVE_1 '匯入等級/分數
        trImport2.Visible = GB_FG_CAN_SAVE_1 '匯入分署加分
        trImport3.Visible = GB_FG_CAN_SAVE_1 '匯入初擬等級

        If (fg_trImport1_NoUse) Then TIMS.Display_None(trImport1) '匯入等級/分數

        If Not IsPostBack Then
            Call CCreate1() '重新載入資訊 
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
        'TIMS.SetMyValue(s_kpSearch1, "rblFIRSTCHK_SCH", TIMS.GetListValue(rblFIRSTCHK_SCH))
        Session("CO_01_004_Search1") = s_kpSearch1
    End Sub

    Sub UseKeepSearch1()
        '(加強操作便利性)
        If Session("CO_01_004_Search1") Is Nothing Then Return
        Dim s_kpSearch1 As String = Session("CO_01_004_Search1")
        Session("CO_01_004_Search1") = Nothing
        If s_kpSearch1 = "" Then Return

        Common.SetListItem(ddlDISTID, TIMS.GetMyValue(s_kpSearch1, "ddlDISTID"))
        Common.SetListItem(ddlSCORING, TIMS.GetMyValue(s_kpSearch1, "ddlSCORING"))
        OrgName.Text = TIMS.GetMyValue(s_kpSearch1, "OrgName")
        COMIDNO.Text = TIMS.GetMyValue(s_kpSearch1, "COMIDNO")
        Common.SetListItem(OrgPlanKind, TIMS.GetMyValue(s_kpSearch1, "OrgPlanKind"))
        Common.SetListItem(OrgKindList, TIMS.GetMyValue(s_kpSearch1, "OrgKindList"))
        'Common.SetListItem(rblFIRSTCHK_SCH, TIMS.GetMyValue(s_kpSearch1, "rblFIRSTCHK_SCH"))
        Call SSearch1()
    End Sub

    ''' <summary>重新載入資訊</summary>
    Sub CCreate1()
        BtnSaveData1.Visible = False
        Labmsg2.Text = "" ' "白天每2小時會計算結果：9:00、11:00、13:00、15:00、17:00"
        '審核等級
        ddlIMPLEVEL_1 = TIMS.Get_SCORELEVEL(ddlIMPLEVEL_1)
        'autorecsubtotal
        Dim js_auto1 As String = "autorecsubtotal();"
        '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
        SCORE4_1.Attributes("onclick") = js_auto1 '"javascript:autorecsubtotal();"
        SCORE4_1.Attributes("onblur") = js_auto1 '"javascript:autorecsubtotal();"
        SCORE4_1.Attributes("onchange") = js_auto1 '"javascript:autorecsubtotal();"

        SUBTOTAL.Attributes("onclick") = js_auto1 '"javascript:autorecsubtotal();"
        SUBTOTAL.Attributes("onblur") = js_auto1 '"javascript:autorecsubtotal();"
        SUBTOTAL.Attributes("onchange") = js_auto1 '"javascript:autorecsubtotal();"
        SUBTOTAL.ReadOnly = True
        SUBTOTAL.ApplyStyle(TIMS.GET_RO_STYLE())

        btnResetSUBTOTAL.Attributes("onclick") = "click_btnResetSUBTOTAL();" '"javascript:autorecsubtotal();"

        'SCORE4_2_RATE.Attributes("onclick")=js_auto1 '"javascript:autorecsubtotal();"
        'SCORE4_2_RATE.Attributes("onblur")=js_auto1 '"javascript:autorecsubtotal();"
        'SCORE4_2_RATE.Attributes("onchange")=js_auto1 '"javascript:autorecsubtotal();"

        divSch1.Visible = True
        divEdt1.Visible = False
        msg1.Text = ""
        PageControler1.Visible = False
        '評核版本
        'ddlSENDVER=Get_SENDVER_TS(ddlSENDVER)
        '評核結果
        'ddlRESULT=Get_RESULT_TS(ddlRESULT)

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

        '依登入者機構判斷計畫種類 '依登入者 LID 判斷是否可自由輸入
        If sm.UserInfo.LID = 2 Then '委訓單位動作
            Dim droo As DataRow = TIMS.Get_ORGINFOdr(sm.UserInfo.OrgID, objconn)
            If droo Is Nothing Then Throw New ArgumentException("登入資訊有誤,請洽系統管理者!")
            OrgName.Text = Convert.ToString(droo("OrgName"))
            COMIDNO.Text = Convert.ToString(droo("ComIDNO"))
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
            Common.SetListItem(ddlSCORING, "2025-01-2024-1-2024-2")
        End If

        '選擇清除工作 'SelectValue.Value=""
        DataGridTable.Visible = False
        Call UseKeepSearch1()
    End Sub

    Protected Sub BtnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        'Call sClearlist1()
        Call UPDATE_ORG_SCORING2_SUBTOTAL()

        Call SSearch1()
    End Sub

    ''' <summary>(儲存DB)直接計算小計,更新1年內的資料</summary>
    Sub UPDATE_ORG_SCORING2_SUBTOTAL()
        Dim vDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim vORGKIND2 As String = TIMS.GetListValue(OrgPlanKind) '計畫
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        If vDISTID = "" OrElse vORGKIND2 = "" OrElse vSCORINGID = "" Then Return

        Dim PMS_U As New Hashtable From {{"DISTID", vDISTID}, {"ORGKIND2", vORGKIND2}, {"SCORINGID", vSCORINGID}}
        Dim SQL_U As String = "
UPDATE ORG_SCORING2 
SET SUBTOTAL=(SCORE1_1+SCORE1_2)+(SCORE2_1_1_ALL+SCORE2_1_2_SUM_ALL+SCORE2_1_3)+(SCORE2_2_1+SCORE2_2_2+SCORE2_3_1)+(SCORE3_1+SCORE3_2)+isnull(SCORE4_1,0.0)+isnull(SCORE4_2,0.0)
FROM ORG_SCORING2 A 
JOIN ORG_ORGINFO oo ON oo.OrgID=a.OrgID
WHERE a.DISTID=@DISTID AND oo.ORGKIND2=@ORGKIND2 and CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID
AND (A.SUBTOTAL IS NULL OR A.SUBTOTAL!=(A.SCORE1_1+A.SCORE1_2)+(A.SCORE2_1_1_ALL+A.SCORE2_1_2_SUM_ALL+A.SCORE2_1_3)+(A.SCORE2_2_1+A.SCORE2_2_2+A.SCORE2_3_1)+(A.SCORE3_1+A.SCORE3_2)+isnull(A.SCORE4_1,0.0)+isnull(A.SCORE4_2,0.0))
AND A.MODIFYDATE>=DATEADD(YY,-1,GETDATE())
"
        TIMS.ExecuteNonQuery(SQL_U, objconn, PMS_U)
    End Sub

    ''' <summary>調整-儲存-系統排序等級</summary>
    ''' <param name="dt2"></param>
    Private Sub UPDATE_CAL_SORT_RL(ByRef dt2 As DataTable)
        Dim SQL_1 As String = "SELECT SRTL,SRTR FROM V_SORT1 WITH(NOLOCK) ORDER BY SRTL"
        Dim dt1 As DataTable = DbAccess.GetDataTable(SQL_1, objconn)
        '沒有4筆資料異常
        If TIMS.dtNODATA(dt1) OrElse dt1.Rows.Count <> 4 Then Return
        Dim FSS As String = ""
        FSS = "SRTL='A'" : Dim v_M1_A As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        FSS = "SRTL='B'" : Dim v_M1_B As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        FSS = "SRTL='C'" : Dim v_M1_C As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        FSS = "SRTL='D'" : Dim v_M1_D As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        If v_M1_A = 0 OrElse v_M1_B = 0 OrElse v_M1_C = 0 OrElse v_M1_D = 0 Then Return
        '沒有資料無須計算
        If TIMS.dtNODATA(dt2) Then Return
        Dim iCNT_all As Double = dt2.Rows.Count

        Dim vDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim vORGKIND2 As String = TIMS.GetListValue(OrgPlanKind) '計畫
        Dim R_PMS As New Hashtable From {{"DISTID", vDISTID}, {"SCORINGID", vSCORINGID}, {"ORGKIND2", vORGKIND2}}
        Dim dtRL As DataTable = GET_SORT_RL_DT(R_PMS)
        '(沒有資料為異常)
        If TIMS.dtNODATA(dtRL) Then Return

        Dim i_ROW As Double = 0
        For Each dr2 As DataRow In dt2.Rows
            Dim FSS2 As String = $"OSID2={dr2("OSID2")}"
            If dtRL.Select(FSS2).Length > 0 Then
                Dim drL As DataRow = dtRL.Select(FSS2)(0)
                dr2("SORTRATIO1_N") = drL("SORTRATIO1_N")
                dr2("SORTRATIO1") = drL("SORTRATIO1")
                dr2("SORTLEVEL1") = drL("SORTLEVEL1")
                dr2("CAPIDX1") = drL("CAPIDX1") '說明
            End If
        Next
        dt2.AcceptChanges()
    End Sub

    ''' <summary>取得-系統排序等級-調整顯示內容</summary>
    ''' <param name="R_PMS"></param>
    ''' <returns></returns>
    Function GET_SORT_RL_DT(R_PMS As Hashtable) As DataTable
        ' Dim R_PMS As New Hashtable From {{"DISTID", vDISTID}, {"SCORINGID", vSCORINGID}, {"ORGKIND2", vORGKIND2}}
        Dim vDISTID As String = TIMS.GetMyValue2(R_PMS, "DISTID")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(R_PMS, "ORGKIND2")
        Dim vSCORINGID As String = TIMS.GetMyValue2(R_PMS, "SCORINGID")

        Dim SQL_1 As String = "SELECT SRTL,SRTR FROM V_SORT1 WITH(NOLOCK) ORDER BY SRTL"
        Dim dt1 As DataTable = DbAccess.GetDataTable(SQL_1, objconn)
        '沒有4筆資料異常
        If TIMS.dtNODATA(dt1) OrElse dt1.Rows.Count <> 4 Then Return TIMS.dtNew()
        Dim FSS As String = ""
        FSS = "SRTL='A'" : Dim v_M1_A As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        FSS = "SRTL='B'" : Dim v_M1_B As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        FSS = "SRTL='C'" : Dim v_M1_C As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        FSS = "SRTL='D'" : Dim v_M1_D As Double = If(dt1.Select(FSS).Length > 0, dt1.Select(FSS)(0)("SRTR"), 0)
        If v_M1_A = 0 OrElse v_M1_B = 0 OrElse v_M1_C = 0 OrElse v_M1_D = 0 Then Return TIMS.dtNew()

        Dim PMS_2 As New Hashtable From {{"DISTID", vDISTID}, {"ORGKIND2", vORGKIND2}, {"SCORINGID", vSCORINGID}}

        'DECLARE @DISTID  VARCHAR(4)='001';DECLARE @ORGKIND2  VARCHAR(3)='G';DECLARE @SCORINGID  VARCHAR(22)='2025-01-2024-1-2024-2';
        Dim SQL_2 As String = "
SELECT a.OSID2,a.SUBTOTAL,oo.ORGKIND2,CONCAT(dbo.FN_CYEAR2(a.YEARS),'年',a.MONTHS,'月'
,'(',dbo.FN_CYEAR2(a.YEARS1),'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end,'~'
,dbo.FN_CYEAR2(a.YEARS2),'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end,')') SCORING_N
,CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2) SCORINGID
,a.IMPLEVEL_1,a.DISTID,a.SORTRATIO1,a.SORTLEVEL1,'' SORTRATIO1_N,'' CAPIDX1
from ORG_SCORING2 a WITH(NOLOCK) 
JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID
where a.DISTID=@DISTID and oo.ORGKIND2=@ORGKIND2 and CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID
ORDER BY a.SUBTOTAL DESC,a.IMPLEVEL_1
"
        Dim dt2 As DataTable = DbAccess.GetDataTable(SQL_2, objconn, PMS_2)
        '沒有資料無須計算
        If TIMS.dtNODATA(dt2) Then Return dt2
        Dim iCNT_all As Double = dt2.Rows.Count
        'Dim SQL_U1 As String = "update ORG_SCORING2 set SORTRATIO1=@SORTRATIO1, SORTLEVEL1=@SORTLEVEL1 where OSID2=@OSID2"

        Dim last_SORTRATIO1 As Decimal = 0
        Dim last_LEVEL As String = ""
        Dim last_SUBTOTAL As Double = 0
        'Dim last_OSID2 As Integer = 0
        Dim i_ROW As Double = 0
        For Each dr2 As DataRow In dt2.Rows
            i_ROW += 1
            Dim v_M1 As Double = i_ROW / iCNT_all * 100
            Dim v_M2 As Decimal = TIMS.TruncateDecimal(i_ROW / iCNT_all * 100, 2)
            Dim V_SORTLEVEL1 As String = ""
            Dim V_CAPIDX1 As String = "" '說明
            Select Case v_M1
                Case <= v_M1_A
                    V_SORTLEVEL1 = "A"
                    'If $"{dr2("IMPLEVEL_1")}" <> "" AndAlso V_SORTLEVEL1 <> $"{dr2("IMPLEVEL_1")}" Then V_CAPIDX1 = CST_CAPIDX_UP 'CST_CAPIDX_DOWN
                    last_LEVEL = V_SORTLEVEL1
                Case <= v_M1_B
                    V_SORTLEVEL1 = "B"
                    If $"{dr2("IMPLEVEL_1")}" <> "" AndAlso V_SORTLEVEL1 <> $"{dr2("IMPLEVEL_1")}" Then
                        Select Case $"{dr2("IMPLEVEL_1")}"
                            Case "A"
                                V_CAPIDX1 = CST_CAPIDX_DOWN
                                'Case Else V_CAPIDX1 = CST_CAPIDX_UP
                        End Select
                    End If
                    If last_LEVEL = "A" AndAlso TIMS.VAL1_Equal(last_SUBTOTAL, TIMS.VAL1(dr2("SUBTOTAL"))) Then
                        UPDATA_CAPIDX1_SUBTOTAL(dt2, last_SUBTOTAL, CST_CAPIDX_WARN2)
                        'V_CAPIDX1 = CST_CAPIDX_WARN
                    Else
                        last_LEVEL = V_SORTLEVEL1
                    End If
                Case <= v_M1_C
                    V_SORTLEVEL1 = "C"
                    If $"{dr2("IMPLEVEL_1")}" <> "" AndAlso V_SORTLEVEL1 <> $"{dr2("IMPLEVEL_1")}" Then
                        Select Case $"{dr2("IMPLEVEL_1")}"
                            Case "A", "B"
                                V_CAPIDX1 = CST_CAPIDX_DOWN
                                'Case Else V_CAPIDX1 = CST_CAPIDX_UP
                        End Select
                    End If
                    If last_LEVEL = "B" AndAlso TIMS.VAL1_Equal(last_SUBTOTAL, TIMS.VAL1(dr2("SUBTOTAL"))) Then
                        UPDATA_CAPIDX1_SUBTOTAL(dt2, last_SUBTOTAL, CST_CAPIDX_WARN2)
                        'V_CAPIDX1 = CST_CAPIDX_WARN
                    Else
                        last_LEVEL = V_SORTLEVEL1
                    End If
                Case Else
                    V_SORTLEVEL1 = "D"
                    If $"{dr2("IMPLEVEL_1")}" <> "" AndAlso V_SORTLEVEL1 <> $"{dr2("IMPLEVEL_1")}" Then V_CAPIDX1 = CST_CAPIDX_DOWN 'CST_CAPIDX_UP
                    If last_LEVEL = "C" AndAlso TIMS.VAL1_Equal(last_SUBTOTAL, TIMS.VAL1(dr2("SUBTOTAL"))) Then
                        UPDATA_CAPIDX1_SUBTOTAL(dt2, last_SUBTOTAL, CST_CAPIDX_WARN2)
                        'V_CAPIDX1 = CST_CAPIDX_WARN
                    Else
                        last_LEVEL = V_SORTLEVEL1
                    End If
            End Select
            If TIMS.VAL1_Equal(last_SUBTOTAL, TIMS.VAL1(dr2("SUBTOTAL"))) Then
                '(同分)等級暫不變動
                dr2("SORTRATIO1_N") = $"{last_SORTRATIO1}%"
                dr2("SORTRATIO1") = last_SORTRATIO1
            Else
                dr2("SORTRATIO1_N") = $"{v_M2}%"
                dr2("SORTRATIO1") = v_M2
                last_SORTRATIO1 = v_M2
            End If
            dr2("SORTLEVEL1") = V_SORTLEVEL1
            If V_CAPIDX1 <> "" Then dr2("CAPIDX1") = V_CAPIDX1 'TIMS.Str1V($"{dr2("CAPIDX1")", $"{V_CAPIDX1}")
            last_SUBTOTAL = TIMS.VAL1(dr2("SUBTOTAL"))
            'last_OSID2 = TIMS.CINT1(dr2("OSID2"))
        Next
        Return dt2
    End Function

    Private Sub UPDATA_CAPIDX1_SUBTOTAL(dt3 As DataTable, last_SUBTOTAL As Double, str_CAPIDX1 As String)
        Dim ff3 As String = $"SUBTOTAL={last_SUBTOTAL}"
        If dt3.Select(ff3).Length = 0 Then Return
        For Each dr3 As DataRow In dt3.Select(ff3)
            dr3("CAPIDX1") = str_CAPIDX1
        Next
    End Sub


#Region "NO USE"
    'Function GET_Y3_dtENG(ByRef parms As Hashtable, ByRef sParms1 As Hashtable) As DataTable
    '    'EXP: 1:一筆資料[ORG_SCORING2] ／"":(list)一般查詢)
    '    'EXP: Y3:匯出班級明細計分 :Y:匯出-審查計分表／Y2:匯出單位計分／Y3:匯出班級明細計分
    '    Dim vEXP As String=TIMS.GetMyValue2(parms, "EXP")
    '    Dim vOSID2 As String=TIMS.GetMyValue2(parms, "OSID2")
    '    Dim vTPLANID As String=TIMS.GetMyValue2(parms, "TPLANID")
    '    Dim vDISTID As String=TIMS.GetMyValue2(parms, "DISTID")
    '    'Dim vYEARS As String=TIMS.GetMyValue2(parms, "YEARS")
    '    'Dim vHALFYEAR As String=TIMS.GetMyValue2(parms, "HALFYEAR") '1:上年度 /2:下年度
    '    Dim vORGNAME As String=TIMS.GetMyValue2(parms, "ORGNAME")
    '    Dim vCOMIDNO As String=TIMS.GetMyValue2(parms, "COMIDNO")
    '    Dim vORGKIND2 As String=TIMS.GetMyValue2(parms, "ORGKIND2")
    '    Dim vORGKIND As String=TIMS.GetMyValue2(parms, "ORGKIND")
    '    Dim vSCORINGID As String=TIMS.GetMyValue2(parms, "SCORINGID")
    '    Dim vFIRSTCHK_SCH As String=TIMS.GetMyValue2(parms, "FIRSTCHK_SCH")

    '    sParms1.Clear()
    '    sParms1.Add("SCORINGID", vSCORINGID)
    '    sParms1.Add("TPLANID", vTPLANID)
    '    sParms1.Add("DISTID", vDISTID)
    '    Dim sSql As String=""
    '    'dbo.VIEW_SCORING2
    '    sSql &= " WITH WO1 AS ( SELECT a.OSID2,a.SCORINGID,a.SCORING_N" & vbCrLf
    '    sSql &= " ,a.COMIDNO,a.TPLANID,a.YEARS,a.MONTHS" & vbCrLf
    '    sSql &= " ,a.STDATE1,a.STDATE2,a.ORGNAME,a.ORGKIND2" & vbCrLf
    '    sSql &= " FROM dbo.VIEW_SCORING2 a" & vbCrLf
    '    sSql &= " WHERE a.SCORINGID=@SCORINGID" & vbCrLf
    '    sSql &= " AND a.TPLANID=@TPLANID AND a.DISTID=@DISTID" & vbCrLf
    '    If vCOMIDNO <> "" Then
    '        sParms1.Add("COMIDNO", vCOMIDNO)
    '        sSql &= " AND a.COMIDNO=@COMIDNO" & vbCrLf
    '    End If
    '    Select Case vORGKIND2
    '        Case "G", "W"
    '            sParms1.Add("ORGKIND2", vORGKIND2)
    '            sSql &= " AND a.ORGKIND2=@ORGKIND2" & vbCrLf
    '    End Select
    '    If vORGKIND <> "" Then
    '        sParms1.Add("ORGKIND", vORGKIND)
    '        sSql &= " AND a.ORGKIND=@ORGKIND" & vbCrLf
    '    End If
    '    Select Case vFIRSTCHK_SCH
    '        Case "Y", "N"
    '            sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
    '            sSql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
    '    End Select
    '    sSql &= " )" & vbCrLf
    '    'dbo.VIEW2B 
    '    sSql &= " ,WC1 AS (" & vbCrLf
    '    sSql &= " SELECT cc.YEARS,cc.DISTID,cc.TPLANID" & vbCrLf
    '    sSql &= " ,cc.DISTNAME,cc.ORGNAME" & vbCrLf
    '    sSql &= " ,cc.PLANID,cc.COMIDNO,cc.SEQNO" & vbCrLf
    '    sSql &= " ,cc.ORGTYPENAME" & vbCrLf ',單位屬性" & vbCrLf
    '    sSql &= " ,cc.OSID2,cc.TMID,cc.PSNO28" & vbCrLf ' --課程申請流水號" & vbCrLf
    '    sSql &= " ,cc.CLASSCNAME" & vbCrLf '班別名稱" & vbCrLf
    '    sSql &= " ,cc.CYCLTYPE" & vbCrLf '期別" & vbCrLf
    '    sSql &= " ,cc.OCID" & vbCrLf '課程代碼" & vbCrLf
    '    sSql &= " ,cc.APPSTAGE" & vbCrLf '申請階段" & vbCrLf
    '    sSql &= " ,cc.THOURS" & vbCrLf '訓練時數" & vbCrLf
    '    sSql &= " ,cc.TNUM" & vbCrLf
    '    sSql &= " ,cc.STDATE" & vbCrLf '開訓日期" & vbCrLf
    '    sSql &= " ,cc.FTDATE" & vbCrLf '結訓日期" & vbCrLf
    '    sSql &= " ,cc.NOTOPEN" & vbCrLf
    '    sSql &= " ,cc.SUSPENDED" & vbCrLf
    '    sSql &= " FROM WO1 o2" & vbCrLf
    '    sSql &= " JOIN dbo.VIEW2B cc on cc.OSID2=o2.OSID2)" & vbCrLf
    '    'dbo.V_STUDENTINFO 
    '    sSql &= " ,WS1 AS ( SELECT cc.OCID" & vbCrLf
    '    sSql &= " ,COUNT(CASE WHEN cc.NOTOPEN='N' AND dd.KID20 IS NULL AND cc.SUSPENDED IS NULL" & vbCrLf
    '    sSql &= " 	AND dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cs.STDATE)=1 THEN 1 END) STDACTCNT--實際開訓人次" & vbCrLf
    '    sSql &= " ,COUNT(CASE WHEN cc.NOTOPEN='N' AND dd.KID20 IS NULL AND cc.SUSPENDED IS NULL" & vbCrLf
    '    sSql &= " 	AND cs.STUDSTATUS=5 THEN 1 END) STDCLOSECNT--結訓人次" & vbCrLf
    '    sSql &= " FROM WC1 cc" & vbCrLf
    '    sSql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=cc.PLANID and dd.COMIDNO=cc.COMIDNO and dd.SEQNO=cc.SEQNO" & vbCrLf
    '    sSql &= " JOIN dbo.V_STUDENTINFO cs on cs.OCID=cc.OCID" & vbCrLf
    '    sSql &= " GROUP BY cc.OCID )" & vbCrLf
    '    'ALL
    '    sSql &= " SELECT cc.YEARS" & vbCrLf
    '    sSql &= " ,cc.DISTNAME,cc.ORGNAME,cc.COMIDNO" & vbCrLf
    '    sSql &= " ,cc.ORGTYPENAME" & vbCrLf '單位屬性" & vbCrLf
    '    sSql &= " ,cc.PSNO28" & vbCrLf '課程申請流水號" & vbCrLf
    '    sSql &= " ,cc.CLASSCNAME" & vbCrLf '班別名稱" & vbCrLf
    '    sSql &= " ,cc.CYCLTYPE" & vbCrLf '期別" & vbCrLf
    '    sSql &= " ,cc.OCID" & vbCrLf '課程代碼" & vbCrLf
    '    sSql &= " ,cc.APPSTAGE" & vbCrLf '申請階段" & vbCrLf
    '    sSql &= " ,dbo.FN_GET_APPSTAGE(cc.APPSTAGE) APPSTAGE_N" & vbCrLf                '
    '    sSql &= " ,cc.THOURS" & vbCrLf '訓練時數" & vbCrLf
    '    sSql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期" & vbCrLf
    '    sSql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期" & vbCrLf
    '    sSql &= " ,tt.JOBNAME" & vbCrLf '課程分類" & vbCrLf
    '    sSql &= " ,dd.D20KNAME1" & vbCrLf '5+2產業創新計畫" & vbCrLf
    '    sSql &= " ,dd.D20KNAME2" & vbCrLf '台灣AI行動計畫" & vbCrLf
    '    sSql &= " ,dd.D20KNAME3" & vbCrLf '數位國家創新經濟發展方案" & vbCrLf
    '    sSql &= " ,dd.D20KNAME4" & vbCrLf '國家資通安全發展方案" & vbCrLf
    '    sSql &= " ,dd.D20KNAME5" & vbCrLf '前瞻基礎建設計畫" & vbCrLf
    '    sSql &= " ,dd.D20KNAME6" & vbCrLf '新南向政策" & vbCrLf
    '    sSql &= " ,cc.TNUM" & vbCrLf '核定人次" & vbCrLf
    '    sSql &= " ,cs.STDACTCNT" & vbCrLf '實際開訓人次" & vbCrLf
    '    sSql &= " ,cs.STDCLOSECNT" & vbCrLf '結訓人次" & vbCrLf
    '    sSql &= " ,cc.NOTOPEN,CASE cc.NOTOPEN WHEN 'N' THEN '否' WHEN 'Y' THEN '是' END NOTOPEN_N" & vbCrLf '是否停辦" & vbCrLf
    '    sSql &= " ,c2.SENDDATE1" & vbCrLf '招訓資料函送日期" & vbCrLf
    '    sSql &= " ,c2.STATUS1" & vbCrLf '招訓資料函送狀態" & vbCrLf
    '    sSql &= " ,c2.OVERWEEK1" & vbCrLf '招訓資料逾期週數" & vbCrLf
    '    sSql &= " ,c2.SENDDATE2" & vbCrLf '開訓資料函送日期" & vbCrLf
    '    sSql &= " ,c2.STATUS2" & vbCrLf '開訓資料函送狀態" & vbCrLf
    '    sSql &= " ,c2.OVERWEEK2" & vbCrLf '開訓資料逾期週數" & vbCrLf
    '    sSql &= " ,c2.SENDDATE3" & vbCrLf '結訓資料函送日期" & vbCrLf
    '    sSql &= " ,c2.STATUS3" & vbCrLf '結訓資料函送狀態" & vbCrLf
    '    sSql &= " ,c2.OVERWEEK3" & vbCrLf '結訓資料逾期週數" & vbCrLf
    '    sSql &= " ,(SELECT STUFF((SELECT CONCAT(',',x.ALTDATAID) FROM PLAN_REVISE x" & vbCrLf
    '    sSql &= " WHERE x.PLANID=cc.PLANID AND x.COMIDNO=cc.COMIDNO AND x.SEQNO=cc.SEQNO FOR XML PATH('')),1,1,'')) ALTDATAID" & vbCrLf '變更項目" & vbCrLf
    '    sSql &= " ,(SELECT STUFF((SELECT ','+FORMAT(x.SENDDATE4,'yyyy/MM/dd') FROM PLAN_REVISE x" & vbCrLf
    '    sSql &= " WHERE x.PLANID=cc.PLANID AND x.COMIDNO=cc.COMIDNO AND x.SEQNO=cc.SEQNO FOR XML PATH('')),1,1,'')) SENDDATE4" & vbCrLf '申請變更函送日期" & vbCrLf
    '    sSql &= " ,(SELECT STUFF((SELECT CONCAT(',',x.STATUS4) FROM PLAN_REVISE x" & vbCrLf
    '    sSql &= " WHERE x.PLANID=cc.PLANID AND x.COMIDNO=cc.COMIDNO AND x.SEQNO=cc.SEQNO FOR XML PATH('')),1,1,'')) STATUS4" & vbCrLf '申請變更函送狀態" & vbCrLf
    '    sSql &= " ,(SELECT STUFF((SELECT CONCAT(',',x.OVERWEEK4) FROM PLAN_REVISE x" & vbCrLf
    '    sSql &= " WHERE x.PLANID=cc.PLANID AND x.COMIDNO=cc.COMIDNO AND x.SEQNO=cc.SEQNO FOR XML PATH('')),1,1,'')) OVERWEEK4" & vbCrLf '逾期週數" & vbCrLf
    '    sSql &= " ,(SELECT STUFF((SELECT CONCAT(',',x.NOINC4) FROM PLAN_REVISE x" & vbCrLf
    '    sSql &= " WHERE x.PLANID=cc.PLANID AND x.COMIDNO=cc.COMIDNO AND x.SEQNO=cc.SEQNO FOR XML PATH('')),1,1,'')) NOINC4" & vbCrLf '不納入審查計分變更次數
    '    sSql &= " ,(SELECT STUFF((SELECT CONCAT(',',x.NODEDUC4) FROM PLAN_REVISE x" & vbCrLf
    '    sSql &= " WHERE x.PLANID=cc.PLANID AND x.COMIDNO=cc.COMIDNO AND x.SEQNO=cc.SEQNO FOR XML PATH('')),1,1,'')) NODEDUC4" & vbCrLf '政策性課程不扣分
    '    sSql &= " FROM WC1 cc" & vbCrLf
    '    sSql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
    '    sSql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=cc.PLANID and dd.COMIDNO=cc.COMIDNO and dd.SEQNO=cc.SEQNO" & vbCrLf
    '    sSql &= " LEFT JOIN WS1 cs on cs.OCID=cc.OCID" & vbCrLf
    '    sSql &= " LEFT JOIN dbo.CLASS_SCORE c2 WITH(NOLOCK) on c2.OCID=cc.OCID" & vbCrLf
    '    Dim dt As DataTable=DbAccess.GetDataTable(sSql, objconn, sParms1)
    '    Return dt
    'End Function
#End Region

    ''' <summary>匯出班級明細計分-1/2/3</summary>
    ''' <param name="parms"></param>
    ''' <param name="iTYPE2"></param>
    ''' <returns></returns>
    Function GET_Y3_dtCHN(ByRef parms As Hashtable, ByVal iTYPE2 As Integer) As DataTable
        'iTYPE2: 1.班級資訊, 2.班級變更資訊, 3.班級不預告實地抽訪紀錄表
        'EXP: 1:一筆資料[ORG_SCORING2] ／"":(list)一般查詢)
        'EXP: Y3:匯出班級明細計分 :Y:匯出-審查計分表／Y2:匯出單位計分／Y3:匯出班級明細計分
        Dim vEXP As String = TIMS.GetMyValue2(parms, "EXP")
        'Dim vOSID2 As String=TIMS.GetMyValue2(parms, "OSID2")
        Dim vSCORINGID As String = TIMS.GetMyValue2(parms, "SCORINGID")
        Dim vTPLANID As String = TIMS.GetMyValue2(parms, "TPLANID")
        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")
        'Dim vYEARS As String=TIMS.GetMyValue2(parms, "YEARS")
        'Dim vHALFYEAR As String=TIMS.GetMyValue2(parms, "HALFYEAR") '1:上年度 /2:下年度
        Dim vORGNAME As String = TIMS.GetMyValue2(parms, "ORGNAME")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(parms, "COMIDNO")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")
        Dim vORGKIND As String = TIMS.GetMyValue2(parms, "ORGKIND")
        'Dim vFIRSTCHK_SCH As String = TIMS.GetMyValue2(parms, "FIRSTCHK_SCH")

        'sParms1.Clear()
        Dim sParms1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", vTPLANID}, {"DISTID", vDISTID}}
        Dim sSql As String = ""

        'dbo.VIEW_SCORING2
        sSql &= " WITH WO1 AS (" & " SELECT a.OSID2,a.SCORINGID,a.SCORING_N" & vbCrLf
        sSql &= " ,a.COMIDNO,a.TPLANID,a.YEARS,a.MONTHS" & vbCrLf
        sSql &= " ,a.STDATE1,a.STDATE2,a.ORGNAME,a.ORGKIND2" & vbCrLf
        sSql &= " ,a.SP_STDATE1,a.SP_STDATE2,a.SP32_FTDATE2,a.SP42_FTDATE2,a.SP21X_FTDATE" & vbCrLf
        sSql &= " FROM dbo.VIEW_SCORING2 a" & vbCrLf
        sSql &= " WHERE a.SCORINGID=@SCORINGID AND a.TPLANID=@TPLANID AND a.DISTID=@DISTID" & vbCrLf
        If vORGNAME <> "" Then
            sParms1.Add("ORGNAME", vORGNAME)
            sSql &= " AND a.ORGNAME like '%'+@ORGNAME+'%'" & vbCrLf
        End If
        If vCOMIDNO <> "" Then
            sParms1.Add("COMIDNO", vCOMIDNO)
            sSql &= " AND a.COMIDNO=@COMIDNO" & vbCrLf
        End If
        Select Case vORGKIND2
            Case "G", "W"
                sParms1.Add("ORGKIND2", vORGKIND2)
                sSql &= " AND a.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then
            sParms1.Add("ORGKIND", vORGKIND)
            sSql &= " AND a.ORGKIND=@ORGKIND" & vbCrLf
        End If
        'Select Case vFIRSTCHK_SCH
        '    Case "Y", "N"
        '        sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
        '        sSql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
        'End Select
        sSql &= " )" & vbCrLf

        'dbo.VIEW2B 
        'SUSPENDED 此單位因有{0}班停班經認列屬「不可抗力因素」，將不列入核定總班數計算。
        sSql &= " ,WC1 AS (" & " SELECT cc.YEARS,cc.DISTID,cc.TPLANID" & vbCrLf
        sSql &= " ,cc.DISTNAME,cc.ORGNAME" & vbCrLf
        sSql &= " ,cc.PLANID,cc.COMIDNO,cc.SEQNO" & vbCrLf
        sSql &= " ,cc.ORGTYPENAME" & vbCrLf ',單位屬性" & vbCrLf
        sSql &= " ,cc.OSID2,cc.TMID,cc.PSNO28" & vbCrLf ' --課程申請流水號" & vbCrLf
        sSql &= " ,cc.CLASSCNAME" & vbCrLf '班別名稱" & vbCrLf
        sSql &= " ,cc.CYCLTYPE" & vbCrLf '期別" & vbCrLf
        sSql &= " ,cc.OCID" & vbCrLf '課程代碼" & vbCrLf
        sSql &= " ,cc.APPSTAGE" & vbCrLf '申請階段" & vbCrLf
        sSql &= " ,cc.THOURS" & vbCrLf '訓練時數" & vbCrLf
        sSql &= " ,cc.TNUM" & vbCrLf
        sSql &= " ,cc.STDATE" & vbCrLf '開訓日期" & vbCrLf
        sSql &= " ,cc.FTDATE" & vbCrLf '結訓日期" & vbCrLf
        sSql &= " ,cc.NOTOPEN" & vbCrLf
        sSql &= " ,cc.SUSPENDED" & vbCrLf '不可抗力因素
        'sSql &= " ,dd.KID20" & vbCrLf '政府政策性產業 政策型課程班 KID20
        'sSql &= " ,dd.D20KNAME1" & vbCrLf '5+2產業創新計畫" & vbCrLf
        'sSql &= " ,dd.D20KNAME2" & vbCrLf '台灣AI行動計畫" & vbCrLf
        'sSql &= " ,dd.D20KNAME3" & vbCrLf '數位國家創新經濟發展方案" & vbCrLf
        'sSql &= " ,dd.D20KNAME4" & vbCrLf '國家資通安全發展方案" & vbCrLf
        'sSql &= " ,dd.D20KNAME5" & vbCrLf '前瞻基礎建設計畫" & vbCrLf
        'sSql &= " ,dd.D20KNAME6" & vbCrLf '新南向政策" & vbCrLf
        sSql &= " ,CASE WHEN dd.KID20 IS NOT NULL THEN 1 WHEN dd.KID25 IS NOT NULL THEN 1 ELSE 0 END USE_KID2025" & vbCrLf '是否為政策性產業
        sSql &= " ,o2.SP_STDATE1,o2.SP_STDATE2,o2.SP32_FTDATE2,o2.SP42_FTDATE2,o2.SP21X_FTDATE" & vbCrLf
        sSql &= " FROM WO1 o2" & vbCrLf
        sSql &= " JOIN dbo.VIEW2B cc on cc.OSID2=o2.OSID2 AND cc.OCID IS NOT NULL" & vbCrLf
        sSql &= " JOIN dbo.V_PLAN_DEPOT dd on dd.PLANID=cc.PLANID and dd.COMIDNO=cc.COMIDNO and dd.SEQNO=cc.SEQNO and dd.OCID=cc.OCID" & vbCrLf
        sSql &= " WHERE cc.SUSPENDED IS NULL" & " )" & vbCrLf
        'iTYPE2: 1.班級資訊, 2.班級變更資訊, 3.班級不預告實地抽訪紀錄表
        If iTYPE2 = 1 Then 'WCSC2
            sSql &= " ,WCSC2 AS (" & " SELECT a.OCID,b.OVERWEEK1,b.ISPASS1,b.SENDDATE1,b.STATUS1" & vbCrLf
            sSql &= " ,b.OVERWEEK2,b.ISPASS2,b.SENDDATE2,b.STATUS2" & vbCrLf
            sSql &= " ,b.OVERWEEK3,b.ISPASS3,b.SENDDATE3,b.STATUS3" & vbCrLf
            sSql &= " ,ISNULL(dbo.FN_CHK_SPDATE(b.SENDDATE1,a.SP_STDATE1,a.SP_STDATE2),1) SPDATE_SENDDATE1" & vbCrLf
            sSql &= " ,ISNULL(dbo.FN_CHK_SPDATE(b.SENDDATE2,a.SP_STDATE1,a.SP_STDATE2),1) SPDATE_SENDDATE2" & vbCrLf
            sSql &= " ,ISNULL(dbo.FN_CHK_SPDATE(b.SENDDATE3,a.SP_STDATE1,a.SP_STDATE2),1) SPDATE_SENDDATE3" & vbCrLf
            sSql &= " FROM WC1 a" & vbCrLf
            sSql &= " JOIN dbo.CLASS_SCORE b WITH(NOLOCK) ON b.OCID =a.OCID" & " )" & vbCrLf
        End If
        'dbo.V_STUDENTINFO 
        sSql &= " ,WS1 AS ( SELECT cc.OCID" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN cc.NOTOPEN='N' AND dbo.FN_GET_STUDCNT14B(cs.STUDSTATUS,cs.REJECTTDATE1,cs.REJECTTDATE2,cs.STDATE)=1 THEN 1 END) STDACTCNT" & vbCrLf '實際開訓人次" & vbCrLf
        sSql &= " ,COUNT(CASE WHEN cc.NOTOPEN='N' AND cs.STUDSTATUS=5 THEN 1 END) STDCLOSECNT" & vbCrLf '結訓人次" & vbCrLf
        sSql &= " FROM WC1 cc" & vbCrLf
        sSql &= " JOIN dbo.V_STUDENTINFO cs on cs.OCID=cc.OCID" & vbCrLf
        sSql &= " GROUP BY cc.OCID )" & vbCrLf

        'iTYPE2: 1.班級資訊, 2.班級變更資訊, 3.班級不預告實地抽訪紀錄表
        If iTYPE2 = 2 Then 'WRP1
            sSql &= " ,WRP1 AS ( SELECT pr.PLANID,pr.COMIDNO,pr.SEQNO,pr.ALTDATAID,pr.SENDDATE4,pr.STATUS4,pr.OVERWEEK4,pr.ISPASS4,pr.NOINC4,pr.NODEDUC4" & vbCrLf
            sSql &= " ,dbo.FN_CHK_SPDATE(pr.SENDDATE4,cc.SP_STDATE1,cc.SP_STDATE2) SPDATE_SENDDATE4" & vbCrLf
            sSql &= " FROM WC1 cc JOIN dbo.PLAN_REVISE pr WITH(NOLOCK) on pr.PLANID=cc.PLANID AND pr.COMIDNO=cc.COMIDNO AND pr.SEQNO=cc.SEQNO )" & vbCrLf
        ElseIf iTYPE2 = 3 Then 'WUN1
            sSql &= " ,WUN1 AS ( SELECT dbo.FN_CHK_SPDATE(cu.APPLYDATE,cc.SP_STDATE1,cc.SP_STDATE2) SPDATE_APPLYDATE" & vbCrLf
            sSql &= " ,cc.OCID,cu.APPLYDATE,cu.AtteRate,cu.DATA81,cu.SITEM51,cu.SITEM61,cu.SITEM71,cu.SITEM81" & vbCrLf
            sSql &= " FROM WC1 cc JOIN dbo.CLASS_UNEXPECTVISITOR cu WITH(NOLOCK) ON cu.OCID=cc.OCID  )" & vbCrLf
        End If
        'dbo.FN_CHK_KID20(cc.OCID)>0 
        'iTYPE2: 1.班級資訊, 2.班級變更資訊, 3.班級不預告實地抽訪紀錄表
        Select Case iTYPE2
            Case 1
                'ALL
                sSql &= " SELECT cc.YEARS 計畫年度" & vbCrLf
                sSql &= " ,cc.DISTNAME 分署,cc.ORGNAME 訓練單位,cc.COMIDNO 統一編號" & vbCrLf
                sSql &= " ,cc.ORGTYPENAME 單位屬性" & vbCrLf '單位屬性" & vbCrLf
                sSql &= " ,cc.PSNO28 課程申請流水號" & vbCrLf '課程申請流水號" & vbCrLf
                sSql &= " ,cc.CLASSCNAME 班別名稱" & vbCrLf '班別名稱" & vbCrLf
                sSql &= " ,cc.CYCLTYPE 期別" & vbCrLf '期別" & vbCrLf
                sSql &= " ,cc.OCID 課程代碼" & vbCrLf '課程代碼" & vbCrLf
                'sSql &= " ,cc.APPSTAGE" & vbCrLf '申請階段" & vbCrLf
                sSql &= " ,dbo.FN_GET_APPSTAGE(cc.APPSTAGE) 申請階段" & vbCrLf                '
                sSql &= " ,cc.THOURS 訓練時數" & vbCrLf '訓練時數" & vbCrLf
                sSql &= " ,format(cc.STDATE,'yyyy/MM/dd') 開訓日期" & vbCrLf '開訓日期" & vbCrLf
                sSql &= " ,format(cc.FTDATE,'yyyy/MM/dd') 結訓日期" & vbCrLf '結訓日期" & vbCrLf
                sSql &= " ,tt.JOBNAME 課程分類" & vbCrLf '課程分類" & vbCrLf
                'sSql &= " ,cc.D20KNAME1 ""5+2產業創新計畫""" & vbCrLf '5+2產業創新計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME2 台灣AI行動計畫" & vbCrLf '台灣AI行動計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME3 數位國家創新經濟發展方案" & vbCrLf '數位國家創新經濟發展方案" & vbCrLf
                'sSql &= " ,cc.D20KNAME4 國家資通安全發展方案" & vbCrLf '國家資通安全發展方案" & vbCrLf
                'sSql &= " ,cc.D20KNAME5 前瞻基礎建設計畫" & vbCrLf '前瞻基礎建設計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME6 新南向政策" & vbCrLf '新南向政策" & vbCrLf
                sSql &= " ,CASE cc.USE_KID2025 WHEN 1 THEN '是' ELSE '否' END 是否為政策性產業" & vbCrLf '是否為政策性產業
                sSql &= " ,cc.TNUM 核定人次" & vbCrLf '核定人次" & vbCrLf
                sSql &= " ,cs.STDACTCNT 實際開訓人次" & vbCrLf '實際開訓人次" & vbCrLf
                sSql &= " ,cs.STDCLOSECNT 結訓人次" & vbCrLf '結訓人次" & vbCrLf
                'sSql &= " ,cc.NOTOPEN,CASE cc.NOTOPEN WHEN 'N' THEN '否' WHEN 'Y' THEN '是' END 是否停辦" & vbCrLf '是否停辦" & vbCrLf
                sSql &= " ,CASE cc.NOTOPEN WHEN 'N' THEN '否' WHEN 'Y' THEN '是' END 是否停辦" & vbCrLf '是否停辦" & vbCrLf

                'sSql &= " ,ba.SENDDATE1 招訓資料函送日期" & vbCrLf '招訓資料函送日期" & vbCrLf
                'sSql &= " ,dbo.FN_STATUS1_N(ba.STATUS1) 招訓資料函送狀態" & vbCrLf
                'sSql &= " ,dbo.FN_OVERWEEK1_N(ba.OVERWEEK1) 招訓資料逾期週數" & vbCrLf
                'sSql &= " ,dbo.FN_ISPASSCNT_N(ba.ISPASS1) 招訓資資料內容符合規定" & vbCrLf
                'sSql &= " ,bb.SENDDATE2 開訓資料函送日期" & vbCrLf
                'sSql &= " ,dbo.FN_STATUS1_N(bb.STATUS2) 開訓資料函送狀態" & vbCrLf
                'sSql &= " ,dbo.FN_OVERWEEK1_N(bb.OVERWEEK2) 開訓資料逾期週數" & vbCrLf
                'sSql &= " ,dbo.FN_ISPASSCNT_N(bb.ISPASS2) 開訓資料內容符合規定" & vbCrLf
                'sSql &= " ,bc.SENDDATE3 結訓資料函送日期" & vbCrLf
                'sSql &= " ,dbo.FN_STATUS1_N(bc.STATUS3) 結訓資料函送狀態" & vbCrLf
                'sSql &= " ,dbo.FN_OVERWEEK1_N(bc.OVERWEEK3) 結訓資料逾期週數" & vbCrLf
                'sSql &= " ,dbo.FN_ISPASSCNT_N(bc.ISPASS3) 結訓資料內容符合規定" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE1 WHEN 1 THEN bb.SENDDATE1 END 招訓資料函送日期" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE1 WHEN 1 THEN dbo.FN_STATUS1_N(bb.STATUS1) END 招訓資料函送狀態" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE1 WHEN 1 THEN dbo.FN_OVERWEEK1_N(bb.OVERWEEK1) END 招訓資料逾期週數" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE1 WHEN 1 THEN dbo.FN_ISPASSCNT_N(bb.ISPASS1) END 招訓資資料內容符合規定" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE2 WHEN 1 THEN bb.SENDDATE2 END 開訓資料函送日期" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE2 WHEN 1 THEN dbo.FN_STATUS1_N(bb.STATUS2) END 開訓資料函送狀態" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE2 WHEN 1 THEN dbo.FN_OVERWEEK1_N(bb.OVERWEEK2) END 開訓資料逾期週數" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE2 WHEN 1 THEN dbo.FN_ISPASSCNT_N(bb.ISPASS2) END 開訓資料內容符合規定" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE3 WHEN 1 THEN bb.SENDDATE3 END 結訓資料函送日期" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE3 WHEN 1 THEN dbo.FN_STATUS1_N(bb.STATUS3) END 結訓資料函送狀態" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE3 WHEN 1 THEN dbo.FN_OVERWEEK1_N(bb.OVERWEEK3) END 結訓資料逾期週數" & vbCrLf
                sSql &= " ,CASE bb.SPDATE_SENDDATE3 WHEN 1 THEN dbo.FN_ISPASSCNT_N(bb.ISPASS3) END 結訓資料內容符合規定" & vbCrLf

                sSql &= " FROM WC1 cc" & vbCrLf
                sSql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sSql &= " LEFT JOIN WS1 cs on cs.OCID=cc.OCID" & vbCrLf
                sSql &= " LEFT JOIN WCSC2 bb WITH(NOLOCK) on bb.OCID=cc.OCID" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 ba WITH(NOLOCK) on ba.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(ba.SENDDATE1,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 bb WITH(NOLOCK) on bb.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(bb.SENDDATE2,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 bc WITH(NOLOCK) on bc.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(bc.SENDDATE3,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf

            Case 2
                'ALL PLAN_REVISE
                sSql &= " SELECT cc.YEARS 計畫年度" & vbCrLf
                sSql &= " ,cc.DISTNAME 分署,cc.ORGNAME 訓練單位,cc.COMIDNO 統一編號" & vbCrLf
                sSql &= " ,cc.ORGTYPENAME 單位屬性" & vbCrLf
                sSql &= " ,cc.PSNO28 課程申請流水號" & vbCrLf
                sSql &= " ,cc.CLASSCNAME 班別名稱" & vbCrLf
                sSql &= " ,cc.CYCLTYPE 期別" & vbCrLf
                sSql &= " ,cc.OCID 課程代碼" & vbCrLf
                sSql &= " ,dbo.FN_GET_APPSTAGE(cc.APPSTAGE) 申請階段" & vbCrLf
                sSql &= " ,cc.THOURS 訓練時數" & vbCrLf
                sSql &= " ,format(cc.STDATE,'yyyy/MM/dd') 開訓日期" & vbCrLf
                sSql &= " ,format(cc.FTDATE,'yyyy/MM/dd') 結訓日期" & vbCrLf
                sSql &= " ,tt.JOBNAME 課程分類" & vbCrLf
                'sSql &= " ,cc.D20KNAME1 ""5+2產業創新計畫""" & vbCrLf
                'sSql &= " ,cc.D20KNAME2 台灣AI行動計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME3 數位國家創新經濟發展方案" & vbCrLf
                'sSql &= " ,cc.D20KNAME4 國家資通安全發展方案" & vbCrLf
                'sSql &= " ,cc.D20KNAME5 前瞻基礎建設計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME6 新南向政策" & vbCrLf
                'sSql &= " ,CASE cc.USE_KID2025 WHEN 1 THEN '是' ELSE '否' END 是否為政策性產業" & vbCrLf '是否為政策性產業
                sSql &= " ,cc.TNUM 核定人次" & vbCrLf
                sSql &= " ,cs.STDACTCNT 實際開訓人次" & vbCrLf
                sSql &= " ,cs.STDCLOSECNT 結訓人次" & vbCrLf
                sSql &= " ,CASE cc.NOTOPEN WHEN 'N' THEN '否' WHEN 'Y' THEN '是' END 是否停辦" & vbCrLf

                sSql &= " ,dbo.FN_GET_ALTDATAID_N(pr.ALTDATAID) 變更項目" & vbCrLf
                sSql &= " ,FORMAT(pr.SENDDATE4,'yyyy/MM/dd') 申請變更函送日期" & vbCrLf
                sSql &= " ,dbo.FN_STATUS4_N(pr.STATUS4) 申請變更函送狀態" & vbCrLf
                sSql &= " ,dbo.FN_OVERWEEK4_N(pr.OVERWEEK4) 申請變更逾期週數" & vbCrLf
                sSql &= " ,dbo.FN_ISPASSCNT_N(pr.ISPASS4) 申請變更資料內容符合規定" & vbCrLf
                sSql &= " ,CASE pr.NOINC4 WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END 不納入審查計分變更次數" & vbCrLf
                sSql &= " ,CASE pr.NODEDUC4 WHEN 'Y' THEN '是' WHEN 'N' THEN '否' END 政策性課程不扣分" & vbCrLf

                sSql &= " FROM WC1 cc" & vbCrLf
                sSql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sSql &= " JOIN WRP1 pr WITH(NOLOCK) on pr.PLANID=cc.PLANID AND pr.COMIDNO=cc.COMIDNO AND pr.SEQNO=cc.SEQNO AND SPDATE_SENDDATE4=1" & vbCrLf
                sSql &= " LEFT JOIN WS1 cs on cs.OCID=cc.OCID" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 ba WITH(NOLOCK) on ba.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(ba.SENDDATE1,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 bb WITH(NOLOCK) on bb.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(bb.SENDDATE2,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 bc WITH(NOLOCK) on bc.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(bc.SENDDATE3,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf

            Case 3
                Dim sql51 As String = "SELECT CODE_ID,CODE_CNAME,CODE_KIND,SORT_ID FROM SYS_SHAREDCODE WHERE CODE_KIND='CP_01_006_SITEM51'"
                Dim sql61 As String = "SELECT CODE_ID,CODE_CNAME,CODE_KIND,SORT_ID FROM SYS_SHAREDCODE WHERE CODE_KIND='CP_01_006_SITEM61'"
                Dim sql71 As String = "SELECT CODE_ID,CODE_CNAME,CODE_KIND,SORT_ID FROM SYS_SHAREDCODE WHERE CODE_KIND='CP_01_006_SITEM71'"
                Dim sql81 As String = "SELECT CODE_ID,CODE_CNAME,CODE_KIND,SORT_ID FROM SYS_SHAREDCODE WHERE CODE_KIND='CP_01_006_SITEM81'"
                Dim dtSITEM51 As DataTable = DbAccess.GetDataTable(sql51, objconn)
                Dim dtSITEM61 As DataTable = DbAccess.GetDataTable(sql61, objconn)
                Dim dtSITEM71 As DataTable = DbAccess.GetDataTable(sql71, objconn)
                Dim dtSITEM81 As DataTable = DbAccess.GetDataTable(sql81, objconn)
                'ALL CLASS_UNEXPECTVISITOR
                sSql &= " SELECT cc.YEARS 計畫年度" & vbCrLf
                sSql &= " ,cc.DISTNAME 分署,cc.ORGNAME 訓練單位,cc.COMIDNO 統一編號" & vbCrLf
                sSql &= " ,cc.ORGTYPENAME 單位屬性" & vbCrLf
                sSql &= " ,cc.PSNO28 課程申請流水號" & vbCrLf
                sSql &= " ,cc.CLASSCNAME 班別名稱" & vbCrLf
                sSql &= " ,cc.CYCLTYPE 期別" & vbCrLf
                sSql &= " ,cc.OCID 課程代碼" & vbCrLf
                sSql &= " ,dbo.FN_GET_APPSTAGE(cc.APPSTAGE) 申請階段" & vbCrLf
                sSql &= " ,cc.THOURS 訓練時數" & vbCrLf
                sSql &= " ,format(cc.STDATE,'yyyy/MM/dd') 開訓日期" & vbCrLf
                sSql &= " ,format(cc.FTDATE,'yyyy/MM/dd') 結訓日期" & vbCrLf
                sSql &= " ,tt.JOBNAME 課程分類" & vbCrLf
                'sSql &= " ,cc.D20KNAME1 ""5+2產業創新計畫""" & vbCrLf
                'sSql &= " ,cc.D20KNAME2 台灣AI行動計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME3 數位國家創新經濟發展方案" & vbCrLf
                'sSql &= " ,cc.D20KNAME4 國家資通安全發展方案" & vbCrLf
                'sSql &= " ,cc.D20KNAME5 前瞻基礎建設計畫" & vbCrLf
                'sSql &= " ,cc.D20KNAME6 新南向政策" & vbCrLf
                sSql &= " ,CASE cc.USE_KID2025 WHEN 1 THEN '是' ELSE '否' END 是否為政策性產業" & vbCrLf '是否為政策性產業
                sSql &= " ,cc.TNUM 核定人次" & vbCrLf
                sSql &= " ,cs.STDACTCNT 實際開訓人次" & vbCrLf
                sSql &= " ,cs.STDCLOSECNT 結訓人次" & vbCrLf

                sSql &= " ,CASE cc.NOTOPEN WHEN 'N' THEN '否' WHEN 'Y' THEN '是' END 是否停辦" & vbCrLf
                sSql &= " ,format(cu.APPLYDATE,'yyyy/MM/dd') 實地訪視日期" & vbCrLf
                sSql &= " ,CASE WHEN cu.AtteRate>=0 THEN CONCAT(ROUND(CAST(cu.AtteRate AS FLOAT)*CAST(100.0 AS float),2),'%') END 出席率" & vbCrLf
                sSql &= " ,CASE cu.DATA81 WHEN '1' THEN '齊' WHEN '2' THEN '缺' WHEN '5' THEN '其他' END ""學員簽到(退)及教學日誌齊全""" & vbCrLf
                sSql &= " ,cu.SITEM51 重要工作事項未依核定課程施訓" & vbCrLf
                sSql &= " ,cu.SITEM61 課程異常狀況" & vbCrLf
                sSql &= " ,cu.SITEM71 其他未依核定課程施訓" & vbCrLf
                sSql &= " ,cu.SITEM81 其他重大異常狀況" & vbCrLf

                sSql &= " FROM WC1 cc" & vbCrLf
                sSql &= " JOIN dbo.VIEW_TRAINTYPE tt on tt.TMID=cc.TMID" & vbCrLf
                sSql &= " JOIN WUN1 cu ON cu.OCID=cc.OCID and cu.SPDATE_APPLYDATE=1" & vbCrLf
                sSql &= " LEFT JOIN WS1 cs on cs.OCID=cc.OCID" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 ba WITH(NOLOCK) on ba.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(ba.SENDDATE1,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 bb WITH(NOLOCK) on bb.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(bb.SENDDATE2,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                'sSql &= " LEFT JOIN WCSC2 bc WITH(NOLOCK) on bc.OCID=cc.OCID AND ISNULL(dbo.FN_CHK_SPDATE(bc.SENDDATE3,cc.SP_STDATE1,cc.SP_STDATE2),1)=1" & vbCrLf
                If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, String.Concat("--iTYPE2: ", iTYPE2, vbCrLf, TIMS.GetMyValue5(sParms1), vbCrLf, "--##CO_01_004:", vbCrLf, sSql))

                Dim dt3 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)
                For Each drV As DataRow In dt3.Rows
                    If Not IsDBNull(drV("重要工作事項未依核定課程施訓")) Then drV("重要工作事項未依核定課程施訓") = GET_REPLACE_SITEM567(dtSITEM51, Convert.ToString(drV("重要工作事項未依核定課程施訓")))
                    If Not IsDBNull(drV("課程異常狀況")) Then drV("課程異常狀況") = GET_REPLACE_SITEM567(dtSITEM61, Convert.ToString(drV("課程異常狀況")))
                    If Not IsDBNull(drV("其他未依核定課程施訓")) Then drV("其他未依核定課程施訓") = GET_REPLACE_SITEM567(dtSITEM71, Convert.ToString(drV("其他未依核定課程施訓")))
                    If Not IsDBNull(drV("其他重大異常狀況")) Then drV("其他重大異常狀況") = GET_REPLACE_SITEM567(dtSITEM81, Convert.ToString(drV("其他重大異常狀況")))
                Next
                Return dt3
            Case Else
                Return Nothing
        End Select

        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, $"--iTYPE2: {iTYPE2}{vbCrLf}{TIMS.GetMyValue5(sParms1)}{vbCrLf}--##CO_01_004:{vbCrLf}{sSql}")

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)
        Return dt
    End Function

    Private Function GET_REPLACE_SITEM567(dtS1 As DataTable, STR1 As String) As String
        If dtS1 Is Nothing OrElse STR1 Is Nothing OrElse String.IsNullOrEmpty(STR1) OrElse Len(STR1) > 1 Then Return STR1
        Dim rst As String = STR1
        For Each drS1 As DataRow In dtS1.Rows
            Dim sCID As String = Convert.ToString(drS1("CODE_ID"))
            If rst.IndexOf(sCID) > -1 Then
                Dim sCNAME As String = Convert.ToString(drS1("CODE_CNAME"))
                rst = Replace(rst, sCID, sCNAME)
            End If
        Next
        Return rst
    End Function

    ''' <summary>OUTPUT ORG_SCORING2 ALL SQL get TABLE</summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Public Function Get_dtORGSCORING2(ByVal parms As Hashtable) As DataTable
        'EXP: 1:一筆資料[ORG_SCORING2] ／"":(list)一般查詢)
        'EXP: Y3:匯出班級明細計分 :Y:匯出-審查計分表／Y2:匯出單位計分／Y3:匯出班級明細計分
        Dim vEXP As String = TIMS.GetMyValue2(parms, "EXP")
        Dim vOSID2 As String = TIMS.GetMyValue2(parms, "OSID2")
        Dim vTPLANID As String = TIMS.GetMyValue2(parms, "TPLANID")
        Dim vDISTID As String = TIMS.GetMyValue2(parms, "DISTID")
        'Dim vYEARS As String=TIMS.GetMyValue2(parms, "YEARS")
        'Dim vHALFYEAR As String=TIMS.GetMyValue2(parms, "HALFYEAR") '1:上年度 /2:下年度
        Dim vORGNAME As String = TIMS.GetMyValue2(parms, "ORGNAME")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(parms, "COMIDNO")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(parms, "ORGKIND2")
        Dim vORGKIND As String = TIMS.GetMyValue2(parms, "ORGKIND")
        Dim vSCORINGID As String = TIMS.GetMyValue2(parms, "SCORINGID")
        'Dim vFIRSTCHK_SCH As String = TIMS.GetMyValue2(parms, "FIRSTCHK_SCH")

        Dim sParms1 As New Hashtable
        Dim sql As String = ""
        'EXP: 1:一筆資料[ORG_SCORING2] ／"":(list)一般查詢)
        'EXP: Y3:匯出班級明細計分 :Y:匯出／Y2:匯出單位計分／Y3:匯出班級明細計分
        Select Case vEXP
            Case "Y3" 'Y3:匯出班級明細計分 'Return GET_Y3_dtCHN(parms, sParms1) 'Return GET_Y3_dtENG(parms, sParms1)
            Case "Y" 'Y:匯出
                sParms1 = New Hashtable From {{"TPLANID", vTPLANID}, {"DISTID", vDISTID}, {"SCORINGID", vSCORINGID}}
                sql = ""
                sql &= " SELECT oo.COMIDNO" & vbCrLf
                sql &= " ,ROW_NUMBER() OVER(ORDER BY a.SUBTOTAL DESC,a.IMPLEVEL_1,a.OSID2) AS ROWID" & vbCrLf
                'sql &= " ,ROW_NUMBER() OVER(ORDER BY a.SUBTOTAL DESC,oo.ORGNAME ASC) AS ROWID" & vbCrLf
                sql &= " ,oo.ORGNAME" & vbCrLf
                sql &= " ,oo.ORGKIND,k1.NAME ORGKIND_N" & vbCrLf
                sql &= " ,oo.ORGKIND1,(SELECT x.ORGTYPE FROM dbo.VIEW_ORGTYPE1 x WHERE x.ORGTYPEID1=oo.ORGKIND1) ORGKIND1_N" & vbCrLf
                sql &= " ,kd.NAME DISTNAME" & vbCrLf
                'SCORING_N 審查計分區間
                sql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS),'年',a.MONTHS,'月'" & vbCrLf
                sql &= "    ,'(',dbo.FN_CYEAR2(a.YEARS1),'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end,'~'" & vbCrLf
                sql &= "    ,dbo.FN_CYEAR2(a.YEARS2),'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end,')') SCORING_N" & vbCrLf
                sql &= " ,CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
                sql &= " ,a.SCORE1_1" & vbCrLf
                sql &= " ,a.SCORE1_2" & vbCrLf
                'sql &= " ,a.SCORE2_1_1a + a.SCORE2_1_1b + a.SCORE2_1_1c + a.SCORE2_1_1d SCORE2_1_1" & vbCrLf
                sql &= " ,a.SCORE2_1_1_ALL" & vbCrLf
                'sql &= " ,a.SCORE2_1_2a + a.SCORE2_1_2b + a.SCORE2_1_2c + a.SCORE2_1_2d SCORE2_1_2" & vbCrLf
                'sql &= " ,CASE WHEN (a.SCORE2_1_2A_DIS+a.SCORE2_1_2B_DIS+a.SCORE2_1_2C_DIS+a.SCORE2_1_2D_DIS)<10.0" & vbCrLf
                'sql &= "  THEN 10.0-(a.SCORE2_1_2A_DIS+a.SCORE2_1_2B_DIS+a.SCORE2_1_2C_DIS+a.SCORE2_1_2D_DIS) ELSE 0 END SCORE2_1_2" & vbCrLf
                sql &= " ,a.SCORE2_1_2_SUM_ALL,a.SCORE2_1_3,a.SCORE2_2_1,a.SCORE2_2_2_DIS,a.SCORE2_2_2,a.SCORE2_3_1,a.SCORE3_1,a.SCORE3_2" & vbCrLf
                '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                sql &= " ,a.SCORE4_1,a.SCORE4_2A,a.SCORE4_2_CNT,a.SCORE4_2_RATE,a.SCORE4_2,a.SCORE4_1_2" & vbCrLf
                '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                'sql &= " ,ISNULL(a.BRANCHPNT,0) BRANCHPNT" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                'sql &= " ,ISNULL(a.SUBTOTAL,0) SUBTOTAL" & vbCrLf
                sql &= " ,a.SUBTOTAL" & vbCrLf
                sql &= " ,ISNULL(a.TOTAL,0) TOTAL" & vbCrLf
                sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
                sql &= " ,v2.VNAME RESULT_N" & vbCrLf
                sql &= " ,a.IMPSCORE_1,a.IMPLEVEL_1" & vbCrLf
                'RLEVEL_2 複審等級  '有複審等級，使用複審等級，複審等級為空，使用匯入等級
                sql &= " ,a.RLEVEL_2" & vbCrLf
                '部長加分,部長加分小計,部長加分等級,署加分,
                sql &= " ,a.MINISTERADD,a.MINISTERSUB,a.MINISTERLEVEL,a.DEPTADD" & vbCrLf
                sql &= " ,a.FIRSTCHK,a.SECONDCHK" & vbCrLf
                'sql &= " ,b.APPLIEDRESULT,a.SORTRATIO1,a.SORTLEVEL1,ROUND(a.SORTRATIO1,2,1) SORTRATIO1_N" & vbCrLf
                sql &= " ,b.APPLIEDRESULT,a.SORTRATIO1,a.SORTLEVEL1,'' SORTRATIO1_N,'' CAPIDX1" & vbCrLf
                'sql &= " ,a.IMODIFYDATE ,a.IMODIFYACCT" & vbCrLf
                sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
                sql &= " JOIN dbo.ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
                sql &= " LEFT JOIN dbo.ORG_TTQS2 b On concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
                sql &= " LEFT JOIN dbo.V_SENDVER v1 On v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.V_RESULT v2 On v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
                sql &= " WHERE a.TPLANID=@TPLANID" & vbCrLf
                '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
                sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
                sql &= " AND a.DISTID=@DISTID" & vbCrLf
                sql &= " AND CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
                'sql &= " AND a.YEARS=@YEARS" & vbCrLf 'If vHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf '1:上年度 /2:下年度
                If vORGNAME <> "" Then sql &= " AND oo.ORGNAME LIKE '%" & vORGNAME & "%'" & vbCrLf
                If vCOMIDNO <> "" Then
                    sParms1.Add("COMIDNO", vCOMIDNO)
                    sql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
                End If
                Select Case vORGKIND2
                    Case "G", "W"
                        sParms1.Add("ORGKIND2", vORGKIND2)
                        sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
                End Select
                If vORGKIND <> "" Then
                    sParms1.Add("ORGKIND", vORGKIND)
                    sql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
                End If
                'Select Case vFIRSTCHK_SCH
                '    Case "Y", "N"
                '        sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
                '        sql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
                'End Select
                sql &= " ORDER BY a.SUBTOTAL DESC,a.IMPLEVEL_1,a.OSID2" & vbCrLf

            Case "Y2" 'Y2:匯出單位計分
                sParms1 = New Hashtable From {{"TPLANID", vTPLANID}, {"DISTID", vDISTID}, {"SCORINGID", vSCORINGID}}
                If vCOMIDNO <> "" Then sParms1.Add("COMIDNO", vCOMIDNO)
                Select Case vORGKIND2
                    Case "G", "W"
                        sParms1.Add("ORGKIND2", vORGKIND2)
                End Select
                If vORGKIND <> "" Then sParms1.Add("ORGKIND", vORGKIND)
                'Select Case vFIRSTCHK_SCH
                '    Case "Y", "N"
                '        sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
                'End Select

                sql = ""
                sql &= " SELECT kd.NAME DISTNAME,oo.ORGNAME,a.CLSACTCNT,a.CLSACTCNT2,a.CLSAPPCNT,a.SCORE1_1A,a.STDACTCNT,a.STDACTCNT2" & vbCrLf
                'sql &= " ,a.STDAPPCNT,a.SCORE1_2A,a.SCORE2_1_1A,a.SCORE2_1_1B,a.SCORE2_1_1C,a.SCORE2_1_1D" & vbCrLf
                'sql &= " ,(a.SCORE2_1_1A+a.SCORE2_1_1B+a.SCORE2_1_1C+a.SCORE2_1_1D) SCORE2_1_1_ITEM" & vbCrLf
                sql &= " ,a.STDAPPCNT,a.SCORE1_2A,a.SCORE2_1_1_SUM_A,a.SCORE2_1_1_SUM_B,a.SCORE2_1_1_SUM_C,a.SCORE2_1_1_SUM_D ,a.SCORE2_1_1_ALL" & vbCrLf
                sql &= " ,a.SCORE2_1_2A_DIS,a.SCORE2_1_2B_DIS,a.SCORE2_1_2C_DIS,a.SCORE2_1_2D_DIS" & vbCrLf
                sql &= " ,a.SCORE2_1_2_SUM_ALL,a.SCORE2_1_3" & vbCrLf
                sql &= " ,case when a.CLSAPPCNT>0 then ROUND(a.SCORE2_2_1_SUM/cast(a.CLSAPPCNT as float)*100,1) end SCORE2_2_1_EQU,a.SCORE2_2_1" & vbCrLf
                'sql &= " ,case when a.CLSAPPCNT>0 then ROUND(a.SCORE2_2_2_SUM/cast(a.CLSAPPCNT as float)*100,1) end SCORE2_2_2_EQU" & vbCrLf
                sql &= " ,a.SCORE2_2_2_DIS,a.SCORE2_2_2" & vbCrLf
                sql &= " ,case when a.SCORE2_3_1_CNT>0 then ROUND(a.SCORE2_3_1_SUM/cast(a.SCORE2_3_1_CNT as float)*100,1) end SCORE2_3_1_EQU,a.SCORE2_3_1" & vbCrLf
                sql &= " ,v2.VNAME SCORE3_1_N,a.SCORE3_1" & vbCrLf
                sql &= " ,case when a.SCORE3_2_CNT>0 then ROUND(a.SCORE3_2_SUM/cast(a.SCORE3_2_CNT as float)/2.0*100,1) end SCORE3_2_EQU,a.SCORE3_2" & vbCrLf
                sql &= " ,a.SCORE4_2_RATE,a.SCORE4_2" & vbCrLf
                'sql &= " ,b.APPLIEDRESULT,a.SORTRATIO1,a.SORTLEVEL1,ROUND(a.SORTRATIO1,2,1) SORTRATIO1_N" & vbCrLf
                sql &= " ,b.APPLIEDRESULT,a.SORTRATIO1,a.SORTLEVEL1,'' SORTRATIO1_N,'' CAPIDX1" & vbCrLf
                sql &= " ,a.IMPSCORE_1,a.IMPLEVEL_1,a.FIRSTCHK,a.SECONDCHK" & vbCrLf
                'sql &= " ,v1.VNAME SENDVER_N ,v2.VNAME RESULT_N" & vbCrLf
                sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
                sql &= " JOIN dbo.ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID" & vbCrLf
                sql &= " LEFT JOIN dbo.KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
                sql &= " LEFT JOIN dbo.ORG_TTQS2 b WITH(NOLOCK) ON concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
                sql &= " LEFT JOIN dbo.V_SENDVER v1 ON v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.V_RESULT v2 ON v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
                sql &= " WHERE a.TPLANID=@TPLANID" & vbCrLf
                '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
                sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
                sql &= " AND a.DISTID=@DISTID" & vbCrLf
                sql &= " AND CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
                If vORGNAME <> "" Then sql &= " AND oo.ORGNAME LIKE '%" & vORGNAME & "%'" & vbCrLf
                If vCOMIDNO <> "" Then sql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
                Select Case vORGKIND2
                    Case "G", "W"
                        sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
                End Select
                If vORGKIND <> "" Then sql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
                'Select Case vFIRSTCHK_SCH
                '    Case "Y", "N"
                '        sql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
                'End Select
                sql &= " ORDER BY a.SUBTOTAL DESC,a.IMPLEVEL_1,a.OSID2" & vbCrLf

            Case "1" '1:一筆資料[ORG_SCORING2] --'(Y:匯出 / 1:一筆資料[ORG_SCORING2])
                sParms1 = New Hashtable From {{"OSID2", vOSID2}}
                sql = "
SELECT a.OSID2,a.OrgID,a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS,a.YEARS1,a.HALFYEAR1,a.YEARS2,a.HALFYEAR2,a.STDATE1,a.STDATE2,a.FTDATE1,a.FTDATE2
,a.CLSACTCNT,a.CLSAPPCNT,a.SCORE1_1A,a.SCORE1_1,a.STDACTCNT,a.STDAPPCNT,a.SCORE1_2A,a.SCORE1_2,a.SCORE2_1_1A,a.SCORE2_1_1B,a.SCORE2_1_1C,a.SCORE2_1_1D
,a.SCORE2_1_2A,a.SCORE2_1_2B,a.SCORE2_1_2C,a.SCORE2_1_2D,a.SCORE2_1_3A,a.SCORE2_1_3,a.SCORE2_2_1,a.SCORE2_2_2,a.SCORE2_3_1,a.SCORE3_1,a.SCORE3_2
,a.SCORE4_1,a.BRANCHPNT,a.DEPTPNT,a.UNITPNT,a.SCORE4_2A,a.SCORE4_2,a.SUBTOTAL,a.TOTAL,a.FIRSTCHK,a.SECONDCHK
,a.CREATEDATE,a.CREATEACCT,a.FIRSTDATE,a.FIRSTACCT,a.SECONDATE,a.SECONACCT,a.MODIFYDATE,a.MODIFYACCT
,a.SCORE2_1_1_SUM_A,a.SCORE2_1_1_SUM_B,a.SCORE2_1_1_SUM_C,a.SCORE2_1_1_SUM_D,a.SCORE2_1_2_SUM_A,a.SCORE2_1_2_SUM_B,a.SCORE2_1_2_SUM_C,a.SCORE2_1_2_SUM_D
,a.SCORE2_1_3_SUM,a.SCORE2_2_1_SUM,a.SCORE2_2_2_SUM,a.SCORE2_3_1_SUM,a.SCORE2_3_1_CNT,a.SCORE3_2_SUM,a.SCORE3_2_CNT,a.SCORE4_1_A
,a.SCORE4_2_RATE,a.SCORE4_2_CNT,a.SCORE4_1_2,a.CLSACTCNT2,a.STDACTCNT2,a.IMPSCORE_1,a.IMPLEVEL_1,a.IMODIFYDATE,a.IMODIFYACCT,a.RLEVEL_2
,a.SCORE2_1_2A_DIS,a.SCORE2_1_2B_DIS,a.SCORE2_1_2C_DIS,a.SCORE2_1_2D_DIS,a.SCORE2_1_2_SUM_ALL,a.SCORE2_2_2_DIS,a.SCORE2_1_1_ALL
,a.SP_STDATE1,a.SP_STDATE2,a.SP32_FTDATE2,a.SP42_FTDATE2,a.SP21X_FTDATE,a.CLSBEDCNT,a.STDBEDCNT,a.STDBEDCNT2
,a.MINISTERADD,a.MINISTERSUB,a.MINISTERLEVEL,a.DEPTADD,a.SORTRATIO1,a.SORTLEVEL1,'' SORTRATIO1_N,'' CAPIDX1
"
                ',ROUND(a.SORTRATIO1,2,1) SORTRATIO1_N
                'FROM dbo.ORG_SCORING2 a WITH(NOLOCK) - datarow 'sql &= " SELECT a.*" & vbCrLf
                sql &= " ,oo.ORGNAME,oo.COMIDNO COMIDNO_N2" & vbCrLf
                sql &= " ,kd.NAME DISTNAME" & vbCrLf
                sql &= " ,k1.NAME ORGKIND_N" & vbCrLf
                'SCORING_N 審查計分區間
                sql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS),'年',a.MONTHS,'月'" & vbCrLf
                sql &= "    ,'(',dbo.FN_CYEAR2(a.YEARS1),'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end,'~'" & vbCrLf
                sql &= "    ,dbo.FN_CYEAR2(a.YEARS2),'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end,')') SCORING_N" & vbCrLf
                sql &= " ,CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
                sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
                sql &= " ,v2.VNAME RESULT_N" & vbCrLf
                sql &= " ,b.APPLIEDRESULT" & vbCrLf
                sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
                sql &= " JOIN dbo.ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
                sql &= " LEFT JOIN dbo.ORG_TTQS2 b On concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
                sql &= " LEFT JOIN dbo.V_SENDVER v1 On v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.V_RESULT v2 On v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
                sql &= " WHERE a.OSID2=@OSID2" & vbCrLf
                '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
                sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf

            Case Else '"":(list)一般查詢)
                sParms1 = New Hashtable From {{"TPLANID", vTPLANID}, {"SCORINGID", vSCORINGID}}
                If vDISTID <> "" Then sParms1.Add("DISTID", vDISTID)
                If vCOMIDNO <> "" Then sParms1.Add("COMIDNO", vCOMIDNO)
                Select Case vORGKIND2
                    Case "G", "W"
                        sParms1.Add("ORGKIND2", vORGKIND2)
                End Select
                If vORGKIND <> "" Then sParms1.Add("ORGKIND", vORGKIND)
                'Select Case vFIRSTCHK_SCH
                '    Case "Y", "N"
                '        sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
                'End Select

                '"":(list)一般查詢)
                sql = ""
                sql &= " SELECT a.OSID2,a.OrgID,a.TPLANID,a.DISTID" & vbCrLf
                'SCORING_N 審查計分區間
                sql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS),'年',a.MONTHS,'月'" & vbCrLf
                sql &= "   ,'(',dbo.FN_CYEAR2(a.YEARS1),'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end,'~'" & vbCrLf
                sql &= "   ,dbo.FN_CYEAR2(a.YEARS2),'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end,')') SCORING_N" & vbCrLf
                sql &= " ,CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
                sql &= " ,a.YEARS ,a.MONTHS" & vbCrLf
                '【實際開班數】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
                '【政策型課程班】：顯示政策性課程全部的班數 (政策性(含停辦))
                sql &= " ,a.CLSACTCNT" & vbCrLf '實際開班數 (核定-停辦-政策性)(核定課程數-停辦課程數)
                sql &= " ,a.CLSACTCNT2" & vbCrLf '政策型課程班  (政策性(含停辦))
                sql &= " ,a.CLSAPPCNT" & vbCrLf '核定總班數
                sql &= " ,a.SCORE1_1A" & vbCrLf '實際開班數/核定總班數
                sql &= " ,a.SCORE1_1" & vbCrLf
                '【實際開訓人次】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
                '【政策性課程核定人次】：顯示政策性課程全部班級的核定人次 (政策性)
                sql &= " ,a.STDACTCNT" & vbCrLf '實際開訓人次
                sql &= " ,a.STDACTCNT2" & vbCrLf '政策性課程核定人次
                sql &= " ,a.STDAPPCNT" & vbCrLf
                sql &= " ,a.SCORE1_2A" & vbCrLf
                sql &= " ,a.SCORE1_2" & vbCrLf
                'sql &= " ,a.SCORE2_1_1A" & vbCrLf'sql &= " ,a.SCORE2_1_1B" & vbCrLf'sql &= " ,a.SCORE2_1_1C" & vbCrLf'sql &= " ,a.SCORE2_1_1D" & vbCrLf
                sql &= " ,a.SCORE2_1_1_ALL,a.SCORE2_1_2A_DIS,a.SCORE2_1_2B_DIS,a.SCORE2_1_2C_DIS,a.SCORE2_1_2D_DIS,a.SCORE2_1_2_SUM_ALL" & vbCrLf
                'SCORE2_1_3A-核定總班數
                sql &= " ,a.SCORE2_1_3A,a.SCORE2_1_3,a.SCORE2_2_1,a.SCORE2_2_2_DIS,a.SCORE2_2_2,a.SCORE2_3_1" & vbCrLf
                sql &= " ,a.SCORE3_1 ,a.SCORE3_2" & vbCrLf
                '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                sql &= " ,a.SCORE4_1" & vbCrLf
                'sql &= ",a.BRANCHPNT" & vbCrLf
                sql &= " ,a.DEPTPNT" & vbCrLf
                sql &= " ,a.UNITPNT" & vbCrLf
                sql &= " ,a.SCORE4_2A" & vbCrLf
                sql &= " ,a.SCORE4_2_CNT" & vbCrLf
                sql &= " ,a.SCORE4_2_RATE" & vbCrLf
                sql &= " ,a.SCORE4_2" & vbCrLf
                sql &= " ,a.SCORE4_1_2" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                'sql &= " ,ISNULL(a.BRANCHPNT,0) BRANCHPNT" & vbCrLf '署 加分項目 配合本部、本署辦理相關活動 或政策宣導 (7%)
                'sql &= " ,ISNULL(a.SUBTOTAL,0) SUBTOTAL" & vbCrLf
                sql &= " ,a.SUBTOTAL" & vbCrLf
                sql &= " ,ISNULL(a.TOTAL,0) TOTAL" & vbCrLf
                'sql &= ",a.TOTAL" & vbCrLf
                sql &= " ,a.FIRSTCHK ,a.SECONDCHK" & vbCrLf
                sql &= " ,a.CREATEDATE ,a.CREATEACCT" & vbCrLf
                sql &= " ,a.FIRSTDATE ,a.FIRSTACCT" & vbCrLf
                sql &= " ,a.SECONDATE ,a.SECONACCT" & vbCrLf
                sql &= " ,a.MODIFYDATE ,a.MODIFYACCT" & vbCrLf
                sql &= " ,a.IMPSCORE_1 ,a.IMPLEVEL_1" & vbCrLf
                sql &= " ,a.IMODIFYDATE ,a.IMODIFYACCT" & vbCrLf
                'RLEVEL_2 複審等級 '有複審等級，使用複審等級，複審等級為空，使用匯入等級
                sql &= " ,a.RLEVEL_2" & vbCrLf
                '部長加分,部長加分小計,部長加分等級,署加分,
                sql &= " ,a.MINISTERADD,a.MINISTERSUB,a.MINISTERLEVEL,a.DEPTADD" & vbCrLf
                sql &= " ,oo.ORGNAME ,oo.COMIDNO" & vbCrLf
                sql &= " ,kd.NAME DISTNAME" & vbCrLf
                sql &= " ,k1.NAME ORGKIND_N" & vbCrLf
                sql &= " ,v1.VNAME SENDVER_N" & vbCrLf
                sql &= " ,v2.VNAME RESULT_N" & vbCrLf
                'sql &= " ,b.APPLIEDRESULT,a.SORTRATIO1,a.SORTLEVEL1,ROUND(a.SORTRATIO1,2,1) SORTRATIO1_N" & vbCrLf
                sql &= " ,b.APPLIEDRESULT,a.SORTRATIO1,a.SORTLEVEL1,'' SORTRATIO1_N,'' CAPIDX1" & vbCrLf
                sql &= " FROM dbo.ORG_SCORING2 a WITH(NOLOCK)" & vbCrLf
                sql &= " JOIN dbo.ORG_ORGINFO oo WITH(NOLOCK) ON oo.OrgID=a.OrgID" & vbCrLf
                sql &= " JOIN dbo.ID_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.KEY_ORGTYPE k1 WITH(NOLOCK) ON k1.ORGTYPEID=oo.ORGKIND" & vbCrLf
                sql &= " LEFT JOIN dbo.ORG_TTQS2 b On concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
                sql &= " LEFT JOIN dbo.V_SENDVER v1 On v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                sql &= " LEFT JOIN dbo.V_RESULT v2 On v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf

                sql &= " WHERE a.TPLANID=@TPLANID" & vbCrLf
                '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
                sql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
                If vDISTID <> "" Then sql &= " AND a.DISTID=@DISTID" & vbCrLf
                sql &= " AND CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
                'sql &= " AND a.YEARS=@YEARS" & vbCrLf 'If vHALFYEAR <> "" Then sql &= " AND a.HALFYEAR=@HALFYEAR" & vbCrLf '1:上年度 /2:下年度
                If vORGNAME <> "" Then sql &= " AND oo.ORGNAME Like '%" & vORGNAME & "%'" & vbCrLf
                If vCOMIDNO <> "" Then sql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
                Select Case vORGKIND2
                    Case "G", "W"
                        sql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
                End Select
                If vORGKIND <> "" Then sql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
                'Select Case vFIRSTCHK_SCH
                '    Case "Y", "N"
                '        sql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
                'End Select
                sql &= " ORDER BY a.SUBTOTAL DESC,a.IMPLEVEL_1,a.OSID2" & vbCrLf

        End Select

        If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, $"--##CO_01_004:{vbCrLf}{TIMS.GetMyValue5(sParms1)}{vbCrLf},--##CO_01_004:{vbCrLf}{sql}")

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, sParms1)
        Return dt
    End Function

    ''' <summary> DG儲存-多筆勾選儲存 </summary>
    Sub SSaveData1()
        Dim iChkCnt As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim drv As DataRowView = eItem.DataItem
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            Dim Hid_SCORE4_1org As HiddenField = eItem.FindControl("Hid_SCORE4_1org")
            Dim tSCORE4_1 As TextBox = eItem.FindControl("tSCORE4_1") 'SCORE4_1'分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
            'Dim LSUBTOTAL As Label=eItem.FindControl("LSUBTOTAL") '小計 'Dim tSUBTOTAL As TextBox=eItem.FindControl("tSUBTOTAL") '小計
            'Dim lRlevel_1 As Label=eItem.FindControl("lRlevel_1") '初審等級/初審<br>等級"
            '初審審核,ddlFIRSTCHK_ALL,HidOSID2,Hid_FIRSTCHKorg,ddlFIRSTCHK,
            Dim HidOSID2 As HiddenField = eItem.FindControl("HidOSID2")
            'Dim Hid_FIRSTCHKorg As HiddenField = eItem.FindControl("Hid_FIRSTCHKorg")
            'Dim vFIRSTCHKorg As String = TIMS.ClearSQM(Hid_FIRSTCHKorg.Value)
            'Dim ddlFIRSTCHK As DropDownList = eItem.FindControl("ddlFIRSTCHK")
            'Dim vddlFIRSTCHK As String = TIMS.GetListValue(ddlFIRSTCHK)

            Dim vSCORE4_1org As String = TIMS.ClearSQM(Hid_SCORE4_1org.Value)
            If vSCORE4_1org = "" Then vSCORE4_1org = "0"
            Dim vSCORE4_1 As String = TIMS.ClearSQM(tSCORE4_1.Text) '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
            If vSCORE4_1 = "" Then vSCORE4_1 = "0"

            Dim vOSID2 As String = TIMS.ClearSQM(HidOSID2.Value)
            Dim flagCanSave1 As Boolean = If(checkbox1.Checked AndAlso vOSID2 <> "", True, False) 'false
            If checkbox1.Checked Then iChkCnt += 1
            If flagCanSave1 Then
                If vSCORE4_1 <> "" Then
                    '有變動過且,不為空白
                    Dim flag_chk1_NG As Boolean = (Not TIMS.IsNumeric1(vSCORE4_1))  '非數字
                    Dim flag_chk2_NG As Boolean = (Not flag_chk1_NG) AndAlso ((Val(vSCORE4_1) > 3) OrElse (Val(vSCORE4_1) < 0)) '數字超過範圍
                    If flag_chk1_NG Then
                        Common.MessageBox(Me, "加分項目由分署自填，請填寫數字，至多加 3 分!")
                        Exit Sub
                    End If
                    If flag_chk2_NG Then
                        Common.MessageBox(Me, "加分項目由分署自填，至多加 3 分!")
                        Exit Sub
                    End If
                End If
                'Select Case vddlFIRSTCHK'vFIRSTCHK,Case "Y", "N",
                'Common.MessageBox(Me, "請選擇審核狀態!"),Exit Sub,
            Else
                'If vddlFIRSTCHK <> vFIRSTCHKorg Then
                '    Common.MessageBox(Me, "請勾選要儲存的項目!(有變更過審核)")
                '    Exit Sub
                'End If
                If vSCORE4_1 <> vSCORE4_1org Then
                    Common.MessageBox(Me, "請勾選要儲存的項目!(有變更過加分項目)")
                    Exit Sub
                End If
            End If
        Next
        If iChkCnt = 0 Then
            Common.MessageBox(Me, "請勾選要儲存的項目!(未有勾選)")
            Exit Sub
        End If

        'updata /* FIRSTCHK=@FIRSTCHK,,SUBTOTAL=@SUBTOTAL */'SCORE4_1：分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
        Dim u_sql As String = "
UPDATE ORG_SCORING2
SET FIRSTACCT=@FIRSTACCT,FIRSTDATE=GETDATE(),SCORE4_1=@SCORE4_1
,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()
WHERE OSID2=@OSID2
"
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim drv As DataRowView = eItem.DataItem
            'Dim CheckBox1 As CheckBox=eItem.FindControl("CheckBox1")
            Dim checkbox1 As HtmlInputCheckBox = eItem.FindControl("checkbox1")
            'Dim Hid_BRANCHPNTorg As HiddenField=eItem.FindControl("Hid_BRANCHPNTorg")
            'Dim tBRANCHPNT As TextBox=eItem.FindControl("tBRANCHPNT") '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
            Dim Hid_SCORE4_1org As HiddenField = eItem.FindControl("Hid_SCORE4_1org")
            Dim tSCORE4_1 As TextBox = eItem.FindControl("tSCORE4_1") 'SCORE4_1'分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
            'Dim LSUBTOTAL As Label=eItem.FindControl("LSUBTOTAL") '小計 'Dim tSUBTOTAL As TextBox=eItem.FindControl("tSUBTOTAL") '小計
            'Dim lRlevel_1 As Label=eItem.FindControl("lRlevel_1") '初審等級/初審<br>等級"
            Dim HidOSID2 As HiddenField = eItem.FindControl("HidOSID2")
            'Dim Hid_FIRSTCHKorg As HiddenField = eItem.FindControl("Hid_FIRSTCHKorg")
            'Dim vFIRSTCHKorg As String = TIMS.ClearSQM(Hid_FIRSTCHKorg.Value)
            'Dim ddlFIRSTCHK As DropDownList = eItem.FindControl("ddlFIRSTCHK")
            'Dim vddlFIRSTCHK As String = TIMS.GetListValue(ddlFIRSTCHK)

            'tSUBTOTAL.Text=TIMS.ClearSQM(tSUBTOTAL.Text)
            'Dim vSUBTOTAL As String=TIMS.VAL1(tSUBTOTAL.Text)
            'If vSUBTOTAL="" Then vSUBTOTAL="0"
            'If Convert.ToString(drv("FIRSTCHK")) <> "" Then Common.SetListItem(ddlFIRSTCHK, Convert.ToString(drv("FIRSTCHK")))
            '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
            Dim vSCORE4_1 As String = TIMS.ClearSQM(tSCORE4_1.Text)
            If vSCORE4_1 = "" Then vSCORE4_1 = "0"

            Dim vOSID2 As String = TIMS.ClearSQM(HidOSID2.Value)
            Dim flagCanSave1 As Boolean = False
            If checkbox1.Checked AndAlso vOSID2 <> "" Then flagCanSave1 = True
            If flagCanSave1 Then
                'u_parms.Add("SUBTOTAL", vSUBTOTAL)
                '{"FIRSTCHK", If(vddlFIRSTCHK <> "", vddlFIRSTCHK, Convert.DBNull)},
                Dim u_parms As New Hashtable From {
                    {"FIRSTACCT", sm.UserInfo.UserID},
                    {"SCORE4_1", vSCORE4_1}, 'SCORE4_1 分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                    {"MODIFYACCT", sm.UserInfo.UserID},
                    {"OSID2", TIMS.CINT1(vOSID2)}
                }
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

                Dim uParms As New Hashtable From {{"OSID2", TIMS.CINT1(vOSID2)}}
                Dim usSql As String = "
UPDATE ORG_SCORING2
SET SUBTOTAL=(SCORE1_1+SCORE1_2)+(SCORE2_1_1_ALL+SCORE2_1_2_SUM_ALL+SCORE2_1_3)+(SCORE2_2_1+SCORE2_2_2+SCORE2_3_1)+(SCORE3_1+SCORE3_2)+isnull(SCORE4_1,0.0)+isnull(SCORE4_2,0.0)
FROM ORG_SCORING2 WHERE OSID2=@OSID2
"
                DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
            End If
        Next

        'divSch1.Visible=True 'divEdt1.Visible=False
        sm.LastResultMessage = "儲存完畢"
    End Sub

    ''' <summary> 單一資料儲存-檢核 </summary>
    ''' <param name="errmsg1"></param>
    ''' <param name="htSS"></param>
    ''' <returns></returns>
    Function CheckData2(ByRef errmsg1 As String, ByRef htSS As Hashtable) As Boolean
        'Dim vBRANCHPNT As String=TIMS.ClearSQM(vBRANCHPNT)
        SUBTOTAL.Text = TIMS.ClearSQM(SUBTOTAL.Text)
        Dim vSUBTOTAL As String = SUBTOTAL.Text 'TIMS.ClearSQM(SUBTOTAL.Text)

        Dim vOSID2 As String = TIMS.ClearSQM(Hid_OSID2.Value)
        If vOSID2 = "" Then
            errmsg1 &= "儲存資料有誤!" & vbCrLf
            Return False
        End If
        'If vOSID2="" Then Exit Sub
        'Dim vddlFIRSTCHK_1 As String = TIMS.GetListValue(ddlFIRSTCHK_1)
        ''Dim vFIRSTCHK As String=TIMS.ClearSQM(ddlFIRSTCHK_1.SelectedValue)
        'Select Case vddlFIRSTCHK_1
        '    Case "Y", "N"
        '    Case Else
        '        errmsg1 &= "請選擇審核狀態!" & vbCrLf
        '        Return False
        '        'Common.MessageBox(Me, "請選擇審核狀態!") Exit Sub
        'End Select

        'tIMPSCORE_1.Text=TIMS.ClearSQM(tIMPSCORE_1.Text)
        'If tIMPSCORE_1.Text <> "" Then
        '    If (Not TIMS.IsNumeric1(tIMPSCORE_1.Text)) Then
        '        errmsg1 &= ("匯入分數 格式有誤，請填寫數字，0-100 分!")
        '        Return False 'Exit Sub
        '    Else
        '        If ((Val(tIMPSCORE_1.Text) > 100) OrElse (Val(tIMPSCORE_1.Text) < 0)) Then
        '            errmsg1 &= ("匯入分數 範圍有誤，請填寫數字，0-100 分!")
        '            Return False 'Exit Sub
        '        End If
        '    End If
        'End If
        Dim v_ddlIMPLEVEL_1 As String = TIMS.GetListValue(ddlIMPLEVEL_1)
        If Hid_IMPLEVEL_1.Value <> "" AndAlso v_ddlIMPLEVEL_1 = "" Then
            errmsg1 &= ("匯入等級 不可為空，請重新選擇!")
            Return False 'Exit Sub
        End If

        '配合分署辦理相關活動或政策宣導 '0-4
        SCORE4_1.Text = TIMS.ClearSQM(SCORE4_1.Text) '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
        Dim vSCORE4_1 As String = SCORE4_1.Text '分署 加分項目
        Select Case vSCORE4_1'SCORE4_1.Text
            Case "0"
            Case Else
                Dim flag_chk1_NG As Boolean = False '非數字
                Dim flag_chk2_NG As Boolean = False '數字超過範圍
                If vSCORE4_1 <> "" Then
                    '有變動過且不為空白
                    If (Not TIMS.IsNumeric1(vSCORE4_1)) Then flag_chk1_NG = True
                    If (Not flag_chk1_NG) AndAlso ((Val(vSCORE4_1) > 3) OrElse (Val(vSCORE4_1) < 0)) Then flag_chk2_NG = True

                    If flag_chk1_NG Then
                        errmsg1 &= ("加分項目由分署自填，請填寫數字，至多加 3 分!")
                        errmsg1 &= ("配合分署辦理相關活動或政策宣導(3%)")
                        Return False 'Exit Sub
                    End If
                    If flag_chk2_NG Then
                        errmsg1 &= ("加分項目由分署自填，至多加 3 分!")
                        errmsg1 &= ("配合分署辦理相關活動或政策宣導(3%)")
                        Return False 'Exit Sub
                    End If
                End If

        End Select

        'htSS.Add("vIMPSCORE1", tIMPSCORE_1.Text) '匯入分數
        'RLEVEL_2 複審等級 '有複審等級，使用複審等級，複審等級為空，使用匯入等級
        '{"vFIRSTCHK", vddlFIRSTCHK_1},
        htSS = New Hashtable From {
            {"vOSID2", vOSID2},
            {"vSCORE4_1", vSCORE4_1},
            {"vSUBTOTAL", vSUBTOTAL},
            {"vIMPLEVEL1", v_ddlIMPLEVEL_1}, '匯入初擬等級
            {"vRLEVEL2", Hid_RLEVEL_2.Value}, 'RLEVEL_2 複審等級
            {"vMINISTERLEVEL", Hid_MINISTERLEVEL.Value} '部加分等級
            }
        Return True
    End Function

    ''' <summary> 單一資料儲存 </summary>
    ''' <param name="htSS"></param>
    Sub SSaveData2(ByRef htSS As Hashtable)
        Dim vFIRSTCHK As String = TIMS.GetMyValue2(htSS, "vFIRSTCHK")
        Dim vOSID2 As String = TIMS.GetMyValue2(htSS, "vOSID2")
        Dim vSCORE4_1 As String = TIMS.GetMyValue2(htSS, "vSCORE4_1")
        Dim vSUBTOTAL As String = TIMS.GetMyValue2(htSS, "vSUBTOTAL")
        Dim vIMPSCORE1 As String = TIMS.GetMyValue2(htSS, "vIMPSCORE1")
        Dim vIMPLEVEL1 As String = TIMS.GetMyValue2(htSS, "vIMPLEVEL1")
        '複審等級'有複審等級，使用複審等級，複審等級為空，使用匯入等級
        Dim vRLEVEL2 As String = TIMS.GetMyValue2(htSS, "vRLEVEL2")
        If vRLEVEL2 = "" Then vRLEVEL2 = vIMPLEVEL1 '複審等級
        Dim vMINISTERLEVEL As String = TIMS.GetMyValue2(htSS, "vMINISTERLEVEL")
        If vMINISTERLEVEL = "" Then vMINISTERLEVEL = vIMPLEVEL1 '部加分等級

        'parms.Add("SECONDCHK", If(vSECONDCHK <> "", vSECONDCHK, Convert.DBNull)) 'parms.Add("SECONACCT", sm.UserInfo.UserID)
        Dim u_parms As New Hashtable From {
            {"FIRSTCHK", If(vFIRSTCHK <> "", vFIRSTCHK, Convert.DBNull)},
            {"FIRSTACCT", sm.UserInfo.UserID},
            {"SCORE4_1", If(vSCORE4_1 <> "", Val(vSCORE4_1), Convert.DBNull)}, '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
            {"SUBTOTAL", If(vSUBTOTAL <> "", Val(vSUBTOTAL), Convert.DBNull)},
            {"IMPSCORE_1", If(vIMPSCORE1 <> "", Val(vIMPSCORE1), Convert.DBNull)}, '匯入分數
            {"IMPLEVEL_1", If(vIMPLEVEL1 <> "", vIMPLEVEL1, Convert.DBNull)}, '匯入等級
            {"RLEVEL_2", If(vRLEVEL2 <> "", vRLEVEL2, Convert.DBNull)}, '複審等級
            {"MINISTERLEVEL", If(vMINISTERLEVEL <> "", vMINISTERLEVEL, Convert.DBNull)}, '部加分等級
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"OSID2", TIMS.CINT1(vOSID2)}
        }
        '-updata
        Dim u_sql As String = ""
        u_sql &= " UPDATE ORG_SCORING2" & vbCrLf
        u_sql &= " SET FIRSTCHK=@FIRSTCHK,FIRSTACCT=@FIRSTACCT ,FIRSTDATE=GETDATE()" & vbCrLf
        u_sql &= " ,SCORE4_1=@SCORE4_1" & vbCrLf '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
        u_sql &= " ,SUBTOTAL=@SUBTOTAL,IMPSCORE_1=@IMPSCORE_1,IMPLEVEL_1=@IMPLEVEL_1" & vbCrLf '匯入分數/'匯入等級
        u_sql &= " ,RLEVEL_2=@RLEVEL_2" & vbCrLf '複審等級
        u_sql &= " ,MINISTERLEVEL=@MINISTERLEVEL" & vbCrLf '部加分等級
        u_sql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " WHERE OSID2=@OSID2" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        Dim uParms As New Hashtable From {{"OSID2", TIMS.CINT1(vOSID2)}}
        Dim usSql As String = ""
        usSql &= " UPDATE ORG_SCORING2" & vbCrLf
        usSql &= " SET SUBTOTAL=(SCORE1_1+SCORE1_2)+(SCORE2_1_1_ALL+SCORE2_1_2_SUM_ALL+SCORE2_1_3)" & vbCrLf
        usSql &= " +(SCORE2_2_1+SCORE2_2_2+SCORE2_3_1)+(SCORE3_1+SCORE3_2)+isnull(SCORE4_1,0.0)+isnull(SCORE4_2,0.0)" & vbCrLf '小計,SCORE4_1
        usSql &= " FROM ORG_SCORING2 WHERE OSID2=@OSID2" & vbCrLf
        DbAccess.ExecuteNonQuery(usSql, objconn, uParms)

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
        'Dim vFIRSTCHK_SCH As String = TIMS.GetListValue(rblFIRSTCHK_SCH) 'A不區分/Y通過/N不通過

        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        Dim vORGKIND2 As String = TIMS.ClearSQM(OrgPlanKind.SelectedValue) '計畫
        Dim vORGKIND As String = TIMS.ClearSQM(OrgKindList.SelectedValue) '機構別

        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        'Dim flag_CanNoDistIDValue As Boolean
        Dim eErrMsg1 As String = ""
        If Not flagS1 AndAlso vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf
        'If vYEARS="" Then eErrMsg1 &= "請選擇年度" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        Call KeepSearch1()

        'EXP:""
        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"SCORINGID", vSCORINGID}}
        'parms.Add("YEARS", sm.UserInfo.Years) 'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        If vDISTID <> "" Then parms.Add("DISTID", vDISTID) 'sql &= " AND t.DISTID=@DISTID" & vbCrLf
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf
        'If vFIRSTCHK_SCH <> "" Then parms.Add("FIRSTCHK_SCH", vFIRSTCHK_SCH) 'A不區分/Y通過/N不通過

        Dim dt As DataTable = Get_dtORGSCORING2(parms)
        'PageControler1.Visible=False 'DataGridTable.Visible=False 'msg1.Text="查無資料"
        If TIMS.dtNODATA(dt) Then Return
        'If dt.Rows.Count = 0 Then Exit Sub
        Call UPDATE_CAL_SORT_RL(dt)

        PageControler1.Visible = True
        DataGridTable.Visible = True
        msg1.Text = ""
        Labmsg2.Text = ""

        Dim fg_CAN_SAVE_1 As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        GB_FG_CAN_SAVE_1 = fg_CAN_SAVE_1
        '(符合-審查計分表(初審))
        BtnSaveData1.Visible = GB_FG_CAN_SAVE_1
        'NOT (符合-審查計分表(初審))
        If Not GB_FG_CAN_SAVE_1 Then Labmsg2.Text = CST_NON_REVIEWSCORE '（非）審查計分表(初審)時間

        PageControler1.PageDataTable = dt
        PageControler1.Sort = "SUBTOTAL DESC,IMPLEVEL_1,SORTLEVEL1,SORTRATIO1"
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>求百分比</summary>
    ''' <param name="oSUM1"></param>
    ''' <param name="oCLSAPPCNT"></param>
    ''' <returns></returns>
    Public Shared Function GET_CAL_1(ByVal oSUM1 As Object, ByVal oCLSAPPCNT As Object) As String
        Dim Rst1 As String = ""
        Dim flagCanCal As Boolean = False 'flagCanCal=False
        If Convert.ToString(oSUM1) <> "" AndAlso Convert.ToString(oCLSAPPCNT) <> "" Then
            If TIMS.CINT1(oCLSAPPCNT) > 0 Then flagCanCal = True
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
        'ddlFIRSTCHK_1.SelectedIndex = -1
        'Common.SetListItem(ddlFIRSTCHK_1, "")

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
        CLSAPPCNT_t10.Text = "" ' Convert.ToString(dr("CLSAPPCNT"))
        CLSAPPCNT_t11.Text = "" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t12.Text="" 'Convert.ToString(dr("CLSAPPCNT"))

        SCORE1_1A.Text = "" 'Convert.ToString(dr("SCORE1_1A"))
        SCORE1_1.Text = "" 'Convert.ToString(dr("SCORE1_1"))

        '【實際開訓人次】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
        '【政策性課程核定人次】：顯示政策性課程全部班級的核定人次 (政策性)
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
        'SCORE2_1_2A_DIS.Text=""
        'SCORE2_1_2B_DIS.Text=""
        'SCORE2_1_2C_DIS.Text=""
        'SCORE2_1_2D_DIS.Text=""
        'SCORE2_1_2_SUM_ALL.Text=""

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
        labSCORE3_1_N.Text = ""
        SCORE3_1.Text = "" 'Convert.ToString(dr("SCORE3_1"))
        SCORE3_2_SUM.Text = "" ' Convert.ToString(dr("SCORE3_2_SUM"))
        SCORE3_2_CNT.Text = "" 'Convert.ToString(dr("SCORE3_2_CNT"))
        SCORE3_2_EQU.Text = "" 'Convert.ToString(dr("SCORE3_2_EQU"))
        SCORE3_2.Text = "" 'Convert.ToString(dr("SCORE3_2"))

        'SCORE4_1_A.Text="" 'Convert.ToString(dr("SCORE4_1_A")) '配合活動得分
        SCORE4_1.Text = "" 'Convert.ToString(dr("SCORE4_1"))'分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
        'tIMPSCORE_1.Text=""
        Hid_IMPLEVEL_1.Value = ""
        Common.SetListItem(ddlIMPLEVEL_1, "")
        'Sql &= " ,a.SCORE4_2A" & vbCrLf
        'Sql &= " ,a.SCORE4_2_CNT" & vbCrLf
        'Sql &= " ,a.SCORE4_2_RATE" & vbCrLf
        'Sql &= " ,a.SCORE4_2" & vbCrLf
        SCORE4_2A.Text = ""
        SCORE4_2_CNT.Text = ""
        SCORE4_2_RATE.Text = ""
        SCORE4_2.Text = "" 'Convert.ToString(dr("SCORE4_2_RATE")) '參訓學員平均填答率
    End Sub

    ''' <summary> 鎖定輸入格-true:Lock </summary>
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
        CLSAPPCNT_t11.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t12.Enabled=flag_Enabled_1  '.Text="" 'Convert.ToString(dr("CLSAPPCNT"))

        SCORE1_1A.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_1A"))
        SCORE1_1.Enabled = flag_Enabled_1  '.Text="" 'Convert.ToString(dr("SCORE1_1"))
        '【實際開訓人次】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
        '【政策性課程核定人次】：顯示政策性課程全部班級的核定人次 (政策性)
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
        'SCORE4_1_A.Text="" 'Convert.ToString(dr("SCORE4_1_A")) '配合活動得分
        'SCORE4_1.Text="" 'Convert.ToString(dr("SCORE4_1"))
    End Sub

    ''' <summary>顯示資料 </summary>
    ''' <param name="dr"></param>
    Sub SShowData1(ByRef dr As DataRow)
        If dr Is Nothing Then Exit Sub

        'ORG_SCORING
        divSch1.Visible = False
        divEdt1.Visible = True

        Hid_OSID2.Value = Convert.ToString(dr("OSID2"))
        If Hid_OSID2.Value = "" Then Exit Sub

        Dim iCLSBEDCNT As Integer = If($"{dr("CLSBEDCNT")}" <> "", TIMS.CINT1(dr("CLSBEDCNT")), 0)
        tr_Lab_SUSPENDED_msg1.Visible = (iCLSBEDCNT > 0)
        Dim str_SUSPENDED_msg1 As String = String.Format(cst_SUSPENDED_msgFM1, iCLSBEDCNT)
        Lab_SUSPENDED_msg1.Text = If(iCLSBEDCNT > 0, str_SUSPENDED_msg1, "")

        Hid_SECONDCHK.Value = Convert.ToString(dr("SECONDCHK"))
        Hid_RLEVEL_2.Value = Convert.ToString(dr("RLEVEL_2")) '複審等級
        Hid_MINISTERLEVEL.Value = Convert.ToString(dr("MINISTERLEVEL")) '部加分等級
        LabOrgName.Text = Convert.ToString(dr("OrgName"))
        labSCORING_N.Text = Convert.ToString(dr("SCORING_N")) '審查計分區間
        LabDISTNAME.Text = Convert.ToString(dr("DISTNAME"))
        'ddlFIRSTCHK_1.Text=Convert.ToString(dr("OrgName"))
        'Common.SetListItem(ddlFIRSTCHK_1, Convert.ToString(dr("FIRSTCHK")))

        CLSACTCNT.Text = Convert.ToString(dr("CLSACTCNT"))
        CLSACTCNT2.Text = Convert.ToString(dr("CLSACTCNT2"))
        CLSAPPCNT.Text = Convert.ToString(dr("CLSAPPCNT")) '核定班數
        CLSAPPCNT_t2.Text = Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t3.Text=Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t9.Text=Convert.ToString(dr("CLSAPPCNT"))
        CLSAPPCNT_t10.Text = Convert.ToString(dr("CLSAPPCNT")) 'SCORE2_1_3A
        CLSAPPCNT_t11.Text = Convert.ToString(dr("CLSAPPCNT"))
        'CLSAPPCNT_t12.Text=Convert.ToString(dr("CLSAPPCNT"))

        SCORE1_1A.Text = Convert.ToString(dr("SCORE1_1A"))
        SCORE1_1.Text = Convert.ToString(dr("SCORE1_1"))
        '【實際開訓人次】顯示一般課程(非政策性課程)有開班的班數 (核定-停辦-政策性)
        '【政策性課程核定人次】：顯示政策性課程全部班級的核定人次 (政策性)
        STDACTCNT.Text = Convert.ToString(dr("STDACTCNT")) '實際開訓人次
        STDACTCNT2.Text = Convert.ToString(dr("STDACTCNT2")) '政策性課程核定人次
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

        'Dim iSCORE2_1_1_ALL As Double=0
        'iSCORE2_1_1_ALL=If(Val(dr("CLSAPPCNT")) > 0, (Val(dr("SCORE2_1_1_SUM_A")) + Val(dr("SCORE2_1_1_SUM_B")) + Val(dr("SCORE2_1_1_SUM_C")) + Val(dr("SCORE2_1_1_SUM_D"))) / Val(dr("CLSAPPCNT")), 0)
        'iSCORE2_1_1_ALL=Math.Round(iSCORE2_1_1_ALL, 1)
        'SCORE2_1_1_ALL.Text=iSCORE2_1_1_ALL
        'SCORE2_1_1_ALL.Text=Convert.ToString(dr("SCORE2_1_1_ALL"))
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
        'APPLIEDRESULT Y: 分署確認
        labSCORE3_1_N.Text = If(Convert.ToString(dr("APPLIEDRESULT")) = "Y", "", "(分署未確認)")
        'SCORE3_1_N.Text=Convert.ToString(dr("SCORE3_1_N"))
        SCORE3_1.Text = Convert.ToString(dr("SCORE3_1"))

        SCORE3_2_SUM.Text = Convert.ToString(dr("SCORE3_2_SUM"))
        SCORE3_2_CNT.Text = Convert.ToString(dr("SCORE3_2_CNT"))

        Dim vSCORE3_2_CNT As Integer = If(Convert.ToString(dr("SCORE3_2_CNT")) <> "", Val(dr("SCORE3_2_CNT")), 0)
        SCORE3_2_EQU.Text = GET_CAL_1(dr("SCORE3_2_SUM"), (vSCORE3_2_CNT * 2)) 'Convert.ToString(dr("SCORE3_2_EQU"))
        SCORE3_2.Text = Convert.ToString(dr("SCORE3_2"))

        'SCORE4_1_A.Text=Convert.ToString(dr("SCORE4_1_A")) '配合活動得分
        SCORE4_1.Text = Convert.ToString(dr("SCORE4_1")) '分署 加分項目'配合活動得分 配合分署辦理相關活動或政策宣導(3%)

        'tIMPSCORE_1.Text=Convert.ToString(dr("IMPSCORE_1"))
        Hid_IMPLEVEL_1.Value = Convert.ToString(dr("IMPLEVEL_1"))
        Common.SetListItem(ddlIMPLEVEL_1, Hid_IMPLEVEL_1.Value)
        'Sql &= " ,a.SCORE4_2A  ,a.SCORE4_2_CNT  ,a.SCORE4_2_RATE ,a.SCORE4_2" & vbCrLf

        ' (1)加分項目分為分署、署,'分署項目標題調整為：4 加分項目(分署)(如圖一),'配合分署辦理相關活動或政策宣導(3%),'預設空白， 由分署自填，至多加 3 分
        '參訓學員訓後動態調查表單位平均填答率達80%(2%) ： 目前都是空的， 應由系統計算
        '    計分公式： 參訓學員訓後動態調查表填寫人次 / 結訓學員總人次 ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
        '    參訓學員訓後動態調查表填寫人次=該訓練單位之所有開訓課程有填寫參訓學員訓後動態調查表之人次總計
        '結訓學員總人次=該訓練單位之所有開訓課程的結訓人次總計
        SCORE4_2A.Text = Convert.ToString(dr("SCORE4_2A")) ' 參訓學員訓後動態調查表填寫人次
        SCORE4_2_CNT.Text = Convert.ToString(dr("SCORE4_2_CNT")) '結訓學員總人次
        SCORE4_2_RATE.Text = Convert.ToString(dr("SCORE4_2_RATE")) '參訓學員平均填答率
        SCORE4_2.Text = $"{dr("SCORE4_2")}" ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
        SUBTOTAL.Text = $"{dr("SUBTOTAL")}"
    End Sub

    Function GET_AUTOCAL_SUBTOTAL(rPMS As Hashtable) As Double
        Dim iSubTotal As Double = 0

        Dim rSCORE1_1 As String = TIMS.GetMyValue2(rPMS, "SCORE1_1")
        Dim rSCORE1_2 As String = TIMS.GetMyValue2(rPMS, "SCORE1_2")
        Dim rSCORE2_1_1_ALL As String = TIMS.GetMyValue2(rPMS, "SCORE2_1_1_ALL")
        Dim rSCORE2_1_2_SUM_ALL As String = TIMS.GetMyValue2(rPMS, "SCORE2_1_2_SUM_ALL")
        Dim rSCORE2_1_3 As String = TIMS.GetMyValue2(rPMS, "SCORE2_1_3")

        Dim rSCORE2_2_1 As String = TIMS.GetMyValue2(rPMS, "SCORE2_2_1")
        Dim rSCORE2_2_2 As String = TIMS.GetMyValue2(rPMS, "SCORE2_2_2")
        Dim rSCORE2_3_1 As String = TIMS.GetMyValue2(rPMS, "SCORE2_3_1")

        Dim rSCORE3_1 As String = TIMS.GetMyValue2(rPMS, "SCORE3_1")
        Dim rSCORE3_2 As String = TIMS.GetMyValue2(rPMS, "SCORE3_2")
        Dim rSCORE4_1 As String = TIMS.GetMyValue2(rPMS, "SCORE4_1")
        Dim rSCORE4_2 As String = TIMS.GetMyValue2(rPMS, "SCORE4_2")

        Dim vSCORE1_1 As Double = Val(rSCORE1_1)
        Dim vSCORE1_2 As Double = Val(rSCORE1_2)
        Dim vSCORE2_1_1_ALL As Double = Val(rSCORE2_1_1_ALL)
        Dim vSCORE2_1_2_SUM_ALL As Double = Val(rSCORE2_1_2_SUM_ALL)
        Dim vSCORE2_1_3 As Double = Val(rSCORE2_1_3)

        Dim vSCORE2_2_1 As Double = Val(rSCORE2_2_1)
        Dim vSCORE2_2_2 As Double = Val(rSCORE2_2_2)
        Dim vSCORE2_3_1 As Double = Val(rSCORE2_3_1)

        Dim vSCORE3_1 As Double = Val(rSCORE3_1)
        Dim vSCORE3_2 As Double = Val(rSCORE3_2)
        Dim vSCORE4_1 As Double = Val(rSCORE4_1)
        Dim vSCORE4_2 As Double = Val(rSCORE4_2)

        iSubTotal += vSCORE1_1
        iSubTotal += vSCORE1_2
        iSubTotal += vSCORE2_1_1_ALL
        iSubTotal += vSCORE2_1_2_SUM_ALL
        iSubTotal += vSCORE2_1_3

        iSubTotal += vSCORE2_2_1
        iSubTotal += vSCORE2_2_2
        iSubTotal += vSCORE2_3_1

        iSubTotal += vSCORE3_1
        iSubTotal += vSCORE3_2

        If (vSCORE4_1 > 0) Then iSubTotal += vSCORE4_1
        If (vSCORE4_2 > 0) Then iSubTotal += vSCORE4_2
        'iSubTotal=Math.Round(iSubTotal, 2)
        'SUBTOTAL.Text=iSubTotal
        Return iSubTotal
    End Function

    ''' <summary> 依參數 OSID2 載入資料顯示 </summary>
    ''' <param name="sCmdArg"></param>
    Sub SLoadData1(ByRef sCmdArg As String)
        If sCmdArg = "" Then Exit Sub
        Dim OSID2 As String = TIMS.GetMyValue(sCmdArg, "OSID2")
        If OSID2 = "" Then Exit Sub

        Dim parms As New Hashtable From {{"EXP", "1"}, {"OSID2", OSID2}}
        Dim dt As DataTable = Get_dtORGSCORING2(parms)
        If dt.Rows.Count = 0 Then Exit Sub
        Dim dr As DataRow = dt.Rows(0)

        divSch1.Visible = True
        divEdt1.Visible = False
        If dt.Rows.Count = 0 Then
            sm.LastErrorMessage = "查無資料"
            Exit Sub
        End If

        divSch1.Visible = False
        divEdt1.Visible = True
        Dim dr1 As DataRow = dt.Rows(0)
        Call SClearlist1()
        Call SShowData1(dr1)
        'debugger;
        Dim aPMS As New Hashtable From {
            {"SCORE1_1", TIMS.ClearSQM(SCORE1_1.Text)},
            {"SCORE1_2", TIMS.ClearSQM(SCORE1_2.Text)},
            {"SCORE2_1_1_ALL", TIMS.ClearSQM(SCORE2_1_1_ALL.Text)},
            {"SCORE2_1_2_SUM_ALL", TIMS.ClearSQM(SCORE2_1_2_SUM_ALL.Text)},
            {"SCORE2_1_3", TIMS.ClearSQM(SCORE2_1_3.Text)},
            {"SCORE2_2_1", TIMS.ClearSQM(SCORE2_2_1.Text)},
            {"SCORE2_2_2", TIMS.ClearSQM(SCORE2_2_2.Text)},
            {"SCORE2_3_1", TIMS.ClearSQM(SCORE2_3_1.Text)},
            {"SCORE3_1", TIMS.ClearSQM(SCORE3_1.Text)},
            {"SCORE3_2", TIMS.ClearSQM(SCORE3_2.Text)},
            {"SCORE4_1", TIMS.ClearSQM(SCORE4_1.Text)},
            {"SCORE4_2", TIMS.ClearSQM(SCORE4_2.Text)}
        }
        Dim iSubTotal As Double = GET_AUTOCAL_SUBTOTAL(aPMS)
        iSubTotal = Math.Round(iSubTotal, 2)
        SUBTOTAL.Text = iSubTotal

        '複審通過，初審鎖定不可修改
        Hid_SECONDCHK.Value = TIMS.ClearSQM(Hid_SECONDCHK.Value)
        Dim fg_SECONDCHK_Y_lock As Boolean = (Hid_SECONDCHK.Value = "Y")

        SCORE4_1.Enabled = If(fg_SECONDCHK_Y_lock, False, True)
        btnResetSUBTOTAL.Disabled = fg_SECONDCHK_Y_lock
        btnResetSUBTOTAL.Visible = If(fg_SECONDCHK_Y_lock, False, True)
        'ddlFIRSTCHK_1.Enabled = If(fg_SECONDCHK_Y_lock, False, True)
        ddlIMPLEVEL_1.Enabled = If(fg_SECONDCHK_Y_lock, False, True)
        SUBTOTAL.Enabled = If(fg_SECONDCHK_Y_lock, False, True)
        BtnSaveData2.Enabled = If(fg_SECONDCHK_Y_lock, False, True)
        BtnSaveData2.Visible = If(fg_SECONDCHK_Y_lock, False, True)

        Dim s_tit1 As String = If(fg_SECONDCHK_Y_lock, cst_tit1, "")
        TIMS.Tooltip(SCORE4_1, s_tit1, True)
        TIMS.Tooltip(btnResetSUBTOTAL, s_tit1, True)
        'TIMS.Tooltip(ddlFIRSTCHK_1, s_tit1, True)
        TIMS.Tooltip(ddlIMPLEVEL_1, s_tit1, True)
        TIMS.Tooltip(SUBTOTAL, s_tit1, True)

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_SAVE_1 As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        GB_FG_CAN_SAVE_1 = fg_CAN_SAVE_1
        '(符合-審查計分表(初審))
        BtnSaveData2.Visible = GB_FG_CAN_SAVE_1
        'NOT (符合-審查計分表(初審))
        If Not GB_FG_CAN_SAVE_1 Then labmsg3.Text = CST_NON_REVIEWSCORE '（非）審查計分表(初審)時間
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

                '依參數 OSID2 載入資料顯示
                Call SLoadData1(sCmdArg)
                '鎖定輸入格-true:Lock 
                Utl_LockData1(True)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dg1 As DataGrid = DataGrid1
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                Dim checkbox1 As HtmlInputCheckBox = e.Item.FindControl("checkbox1")
                Dim lbtView As LinkButton = e.Item.FindControl("lbtView")
                'Dim Hid_BRANCHPNTorg As HiddenField=e.Item.FindControl("Hid_BRANCHPNTorg")
                Dim tBRANCHPNT As TextBox = e.Item.FindControl("tBRANCHPNT") '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                Dim Hid_SCORE4_1org As HiddenField = e.Item.FindControl("Hid_SCORE4_1org")
                Dim tSCORE4_1 As TextBox = e.Item.FindControl("tSCORE4_1") 'SCORE4_1'分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                'Dim LSUBTOTAL As Label=e.Item.FindControl("LSUBTOTAL") '小計
                Dim tSUBTOTAL As TextBox = e.Item.FindControl("tSUBTOTAL") '小計
                Dim Hid_SUBTOTALorg As HiddenField = e.Item.FindControl("Hid_SUBTOTALorg") '小計
                'Dim lRlevel_1 As Label=eItem.FindControl("lRlevel_1") '初審等級/初審<br>等級"
                Dim HidOSID2 As HiddenField = e.Item.FindControl("HidOSID2")
                'Dim Hid_FIRSTCHKorg As HiddenField = e.Item.FindControl("Hid_FIRSTCHKorg")
                'Dim ddlFIRSTCHK As DropDownList = e.Item.FindControl("ddlFIRSTCHK")
                Dim labCAPIDX1 As Label = e.Item.FindControl("labCAPIDX1") '說明
                'tBRANCHPNT.Text=Convert.ToString(drv("BRANCHPNT")) '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                'Hid_BRANCHPNTorg.Value=tBRANCHPNT.Text
                labCAPIDX1.Text = $"{drv("CAPIDX1")}" '說明
                Select Case $"{drv("CAPIDX1")}"
                    Case CST_CAPIDX_DOWN 'CST_CAPIDX_UP
                        e.Item.Cells(9).CssClass = "TC_TD3"
                        e.Item.Cells(10).CssClass = "TC_TD3"
                        'Case CST_CAPIDX_DOWN,e.Item.Cells(9).CssClass = "TC_TD4",e.Item.Cells(10).CssClass = "TC_TD4"
                    Case CST_CAPIDX_WARN2 'CST_CAPIDX_WARN, 
                        e.Item.Cells(9).CssClass = "TC_TD5"
                        e.Item.Cells(10).CssClass = "TC_TD5"
                    Case Else
                        e.Item.Cells(9).CssClass = "whitecol"
                        e.Item.Cells(10).CssClass = "whitecol"
                End Select

                tSCORE4_1.Text = $"{drv("SCORE4_1")}" '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                Hid_SCORE4_1org.Value = $"{drv("SCORE4_1")}" 'tSCORE4_1.Text
                'LSUBTOTAL.Text=Convert.ToString(drv("SUBTOTAL")) '分署小計

                tSUBTOTAL.ReadOnly = True
                tSUBTOTAL.ApplyStyle(TIMS.GET_RO_STYLE())
                If $"{drv("SUBTOTAL")}" <> "" AndAlso TIMS.IsNumeric1(drv("SUBTOTAL")) Then
                    'tSUBTOTAL.Text=Val(drv("SUBTOTAL"))
                    tSUBTOTAL.Text = $"{drv("SUBTOTAL")}" '分署小計
                Else
                    Dim aPMS As New Hashtable From {
                        {"SCORE1_1", drv("SCORE1_1")},
                        {"SCORE1_2", drv("SCORE1_2")},
                        {"SCORE2_1_1_ALL", drv("SCORE2_1_1_ALL")},
                        {"SCORE2_1_2_SUM_ALL", drv("SCORE2_1_2_SUM_ALL")},
                        {"SCORE2_1_3", drv("SCORE2_1_3")},
                        {"SCORE2_2_1", drv("SCORE2_2_1")},
                        {"SCORE2_2_2", drv("SCORE2_2_2")},
                        {"SCORE2_3_1", drv("SCORE2_3_1")},
                        {"SCORE3_1", drv("SCORE3_1")},
                        {"SCORE3_2", drv("SCORE3_2")},
                        {"SCORE4_1", drv("SCORE4_1")},
                        {"SCORE4_2", drv("SCORE4_2")}
                    }
                    Dim iSubTotal As Double = GET_AUTOCAL_SUBTOTAL(aPMS)
                    iSubTotal = Math.Round(iSubTotal, 2)
                    tSUBTOTAL.Text = iSubTotal
                End If
                Hid_SUBTOTALorg.Value = tSUBTOTAL.Text

                Dim js_auto2 As String = String.Format("AUTO_CAL_2('{0}','{1}','{2}','{3}');", Hid_SCORE4_1org.ClientID, tSCORE4_1.ClientID, Hid_SUBTOTALorg.ClientID, tSUBTOTAL.ClientID)
                tSCORE4_1.Attributes("onblur") = js_auto2 '"javascript:autorecsubtotal();"
                tSCORE4_1.Attributes("onchange") = js_auto2 '"javascript:autorecsubtotal();"

                'lRlevel_1.Text="-" '初審等級/初審<br>等級"
                'Hid_FIRSTCHKorg.Value = $"{drv("FIRSTCHK")}"
                'If Convert.ToString(drv("FIRSTCHK")) <> "" Then
                '    Hid_FIRSTCHKorg.Value = Convert.ToString(drv("FIRSTCHK"))
                '    Common.SetListItem(ddlFIRSTCHK, Convert.ToString(drv("FIRSTCHK")))
                'End If

                HidOSID2.Value = $"{drv("OSID2")}"
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OSID2", $"{drv("OSID2")}")
                lbtView.CommandArgument = sCmdArg

                'Dim Hid_RTSID As HiddenField=e.Item.FindControl("Hid_RTSID")
                'Dim Hid_ORGID As HiddenField=e.Item.FindControl("Hid_ORGID")
                'Hid_RTSID.Value=Convert.ToString(drv("OTSID"))
                'Hid_ORGID.Value=Convert.ToString(drv("ORGID"))
                'TIMS.SetMyValue(sCmdArg, "RTSID", Convert.ToString(drv("RTSID")))
                'TIMS.SetMyValue(sCmdArg, "ORGID", Convert.ToString(drv("ORGID")))

                '複審通過，初審鎖定不可修改
                Dim fg_SECONDCHK_Y_lock As Boolean = ($"{drv("SECONDCHK")}" = "Y")
                checkbox1.Disabled = False 'fg_SECONDCHK_Y_lock
                tSCORE4_1.Enabled = True 'If(fg_SECONDCHK_Y_lock, False, True)
                'tSUBTOTAL.Enabled = If(fg_SECONDCHK_Y_lock, False, True) 'ddlFIRSTCHK.Enabled = If(fg_SECONDCHK_Y_lock, False, True)

                Dim s_tit1 As String = If(fg_SECONDCHK_Y_lock, cst_tit1, "")
                Dim s_tit3 As String = If(Not GB_FG_CAN_SAVE_1, cst_tit3, "")
                If s_tit1 <> "" Then
                    checkbox1.Disabled = True
                    tSCORE4_1.Enabled = False
                    TIMS.Tooltip(checkbox1, s_tit1, True)
                    TIMS.Tooltip(tSCORE4_1, s_tit1, True)
                    'TIMS.Tooltip(tSUBTOTAL, s_tit1, True)
                    'TIMS.Tooltip(ddlFIRSTCHK, s_tit1, True)
                ElseIf s_tit3 <> "" Then
                    checkbox1.Disabled = True
                    tSCORE4_1.Enabled = False
                    TIMS.Tooltip(checkbox1, s_tit3, True)
                    TIMS.Tooltip(tSCORE4_1, s_tit3, True)
                    'TIMS.Tooltip(tSUBTOTAL, s_tit2, True)
                ElseIf (tSCORE4_1.Enabled) Then
                    TIMS.Tooltip(tSCORE4_1, cst_tit2, True)
                End If
        End Select
    End Sub

    'DG儲存-多筆勾選儲存 
    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_IMP As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        If Not fg_CAN_IMP Then
            Common.MessageBox(Me, $"{CST_NON_REVIEWSCORE}不可儲存!請再確認審查計分區間!")
            Return
        End If

        'DG儲存-多筆勾選儲存 
        Call SSaveData1()
    End Sub

#Region "Sample NO USE"
    ''' <summary> 本功能為範例-暫無使用 </summary>
    ''' <param name="dt"></param>
    Sub ExpSampleXLS(ByRef dt As DataTable)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Const cst_SampleXLS As String = "~\CO\01\SampleC.xls"
        'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If

        Const Cst_FileSavePath As String = "~/CO/01/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim strErrmsg As String = ""
        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), ".xls")
        Dim sMyFile1 As String = Server.MapPath(sFileName)

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

                Dim rSCORE4_1 As String = Convert.ToString(drV("SCORE4_1")) '分署 加分項目 配合分署辦理相關活動或政策宣導(3%)
                Dim rSCORE4_2 As String = Convert.ToString(drV("SCORE4_2"))

                Dim sql As String = ""
                sql &= " INSERT INTO [Sheet1$]( [統一編號], [序號], [訓練單位名稱], [屬性], [分署]" & vbCrLf
                sql &= " ,[1-1],[1-2]" & vbCrLf
                sql &= " ,[2-1-1],[2-1-2],[2-1-3]" & vbCrLf
                sql &= " ,[2-2-1],[2-2-2],[2-3-1]" & vbCrLf
                sql &= " ,[3-1], [獎牌],[3-2]" & vbCrLf
                sql &= " ,[4-1],[4-2]" & vbCrLf
                sql &= " , [小計], [初擬等級],[4-1-2]" & vbCrLf
                sql &= " , [合計], [等級], [備註] )" & vbCrLf
                sql &= "  VALUES( '" & rCOMIDNO & "','" & rROWID & "','" & rOrgName & "','" & rORGKIND_N & "','" & rDISTNAME & "'" & vbCrLf
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

    ''' <summary>匯出審查計分表</summary>
    Sub SExprot2_Y1()
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
        If vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf
        'If vYEARS="" Then eErrMsg1 &= "請選擇年度" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        'parms.Add("YEARS", sm.UserInfo.Years)
        Dim parms As New Hashtable From {
            {"EXP", "Y"}, '匯出查詢條件
            {"TPLANID", sm.UserInfo.TPlanID},
            {"DISTID", vDISTID}, 'sql &= " AND t.DISTID=@DISTID" & vbCrLf
            {"SCORINGID", vSCORINGID}
        }
        'parms.Add("YEARS", vYEARS)
        'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf

        Dim dtXls As DataTable = Get_dtORGSCORING2(parms)
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        '調整
        For Each drV As DataRow In dtXls.Rows
            Dim rSUBTOTAL As String = $"{drV("SUBTOTAL")}" '小計
            'Dim rIMPLEVEL_1 As String=Convert.ToString(drV("IMPLEVEL_1")) '初擬等級
            If rSUBTOTAL = "" OrElse rSUBTOTAL = "0" Then
                Dim aPMS As New Hashtable From {
                    {"SCORE1_1", drV("SCORE1_1")},
                    {"SCORE1_2", drV("SCORE1_2")},
                    {"SCORE2_1_1_ALL", drV("SCORE2_1_1_ALL")},
                    {"SCORE2_1_2_SUM_ALL", drV("SCORE2_1_2_SUM_ALL")},
                    {"SCORE2_1_3", drV("SCORE2_1_3")},
                    {"SCORE2_2_1", drV("SCORE2_2_1")},
                    {"SCORE2_2_2", drV("SCORE2_2_2")},
                    {"SCORE2_3_1", drV("SCORE2_3_1")},
                    {"SCORE3_1", drV("SCORE3_1")},
                    {"SCORE3_2", drV("SCORE3_2")},
                    {"SCORE4_1", drV("SCORE4_1")},
                    {"SCORE4_2", drV("SCORE4_2")}
                }
                Dim iSubTotal As Double = GET_AUTOCAL_SUBTOTAL(aPMS)
                iSubTotal = Math.Round(iSubTotal, 2)
                rSUBTOTAL = iSubTotal
                drV("SUBTOTAL") = iSubTotal
            End If
        Next

        dtXls.DefaultView.Sort = "SUBTOTAL DESC,IMPLEVEL_1,ORGNAME ASC"
        dtXls = dtXls.DefaultView.ToTable()
        'Call ExpSampleXLS(dtXls)
        Call ExpXMLXLS_1(dtXls)
    End Sub

    ''' <summary>匯出審查計分表</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles btnExp1.Click
        Call SExprot2_Y1()
    End Sub

    Protected Sub BtnBack2_Click(sender As Object, e As EventArgs) Handles BtnBack2.Click
        Call SClearlist1()
        divEdt1.Visible = False
        divSch1.Visible = True
        'Call sShowData1(dr1)
        Call SSearch1()
    End Sub

    ''' <summary>'儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSaveData2_Click(sender As Object, e As EventArgs) Handles BtnSaveData2.Click
        Dim errmsg1 As String = ""
        Dim h_parms As New Hashtable
        CheckData2(errmsg1, h_parms)
        If errmsg1 <> "" Then
            Common.MessageBox(Me, errmsg1)
            Exit Sub
        End If
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_IMP As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        If Not fg_CAN_IMP Then
            Common.MessageBox(Me, $"{CST_NON_REVIEWSCORE}不可儲存!請再確認審查計分區間!")
            Return
        End If
        'select SCORE4_1--配合分署辦理相關活動或政策宣導(3%)
        ',SCORE4_2A--參訓學員訓後動態調查表填寫人次
        ',SCORE4_2_CNT--結訓學員總人次
        ',SCORE4_2_RATE--參訓學員平均填答率
        ',SCORE4_2--參訓學員訓後動態調查表單位平均填答率達80% (2%)
        ',SUBTOTAL--小計
        ',SCORE4_1_2--配合本部、本署辦理相關活動或政策宣導 (7%)
        ',[TOTAL]--"合計 I+II+III+IV"
        '等級 '備註
        'FROM ORG_SCORING2 ' WHERE 0=0
        Call SSaveData2(h_parms)
        '重新載入資訊
        Call CCreate1()
    End Sub

    ''' <summary> 匯出xls-審查計分表</summary>
    ''' <param name="dtXLS"></param>
    Sub ExpXMLXLS_1(ByRef dtXLS As DataTable)
        Dim strErrmsg As String = ""
        If dtXLS.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        'Const cst_files_ext As String=".xlsx" ' ".xls" 
        '~\CO\01\SampleD.xlsx
        Const cst_SampleXLS As String = "~\CO\01\SampleD.xlsx" '& cst_files_ext
        'copy一份sample資料---Start
        If Not IO.File.Exists(Server.MapPath(cst_SampleXLS)) Then
            Common.MessageBox(Me, "Sample檔案不存在")
            Exit Sub
        End If

        Const Cst_FileSavePath As String = "~/CO/01/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim sFileName As String = String.Concat(Cst_FileSavePath, TIMS.GetDateNo(), "xlsx") '複製一份(Sample)
        Dim sMyFile1 As String = Server.MapPath(sFileName) '複製一份(Sample)

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
        Dim fs1 As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
        'Dim fs As FileInfo=New FileInfo(sMyFile1)
        Dim ep As ExcelPackage = New ExcelPackage(fs1)
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


        Dim i_ROWID As Integer = 0
        Dim i_currentRow As Integer = 5
        For Each drV As DataRow In dtXLS.Rows
            i_ROWID += 1
            i_currentRow += 1
            Dim rCOMIDNO As String = Convert.ToString(drV("COMIDNO"))
            Dim rROWID As String = i_ROWID 'Convert.ToString(drV("ROWID"))
            Dim rOrgName As String = Convert.ToString(drV("OrgName"))
            'Dim rORGKIND_N As String=Convert.ToString(drV("ORGKIND_N"))
            Dim rORGKIND1_N As String = Convert.ToString(drV("ORGKIND1_N"))
            Dim rDISTNAME As String = Convert.ToString(drV("DISTNAME"))
            Dim rSCORE1_1 As String = Convert.ToString(drV("SCORE1_1"))
            Dim rSCORE1_2 As String = Convert.ToString(drV("SCORE1_2"))

            Dim rSCORE2_1_1_ALL As String = Convert.ToString(drV("SCORE2_1_1_ALL"))            '
            Dim rSCORE2_1_2_SUM_ALL As String = Convert.ToString(drV("SCORE2_1_2_SUM_ALL"))
            Dim rSCO2_1_3 As String = Convert.ToString(drV("SCORE2_1_3"))

            Dim rSCO2_2_1 As String = Convert.ToString(drV("SCORE2_2_1"))
            Dim rSCO2_2_2 As String = Convert.ToString(drV("SCORE2_2_2"))
            Dim rSCO2_3_1 As String = Convert.ToString(drV("SCORE2_3_1"))

            Dim rSCORE3_1 As String = Convert.ToString(drV("SCORE3_1"))
            Dim rSCORE3_1b As String = Convert.ToString(drV("RESULT_N"))
            Dim rSCORE3_2 As String = Convert.ToString(drV("SCORE3_2"))

            Dim rSCORE4_1 As String = Convert.ToString(drV("SCORE4_1")) '配合分署辦理相關活動或政策宣導(3%)
            Dim rSCORE4_2A As String = Convert.ToString(drV("SCORE4_2A")) ' 參訓學員訓後動態調查表填寫人次
            Dim rSCORE4_2_CNT As String = Convert.ToString(drV("SCORE4_2_CNT")) '結訓學員總人次
            Dim rSCORE4_2_RATE As String = Convert.ToString(drV("SCORE4_2_RATE")) '參訓學員平均填答率
            Dim rSCORE4_2 As String = Convert.ToString(drV("SCORE4_2")) ' >= 80% --> 得 2 分  // < 80% --> 得 0 分
            'rSUBTOTAL
            Dim rSUBTOTAL As String = Convert.ToString(drV("SUBTOTAL")) '小計
            Dim rIMPLEVEL_1 As String = Convert.ToString(drV("IMPLEVEL_1")) '初擬等級

            'rSCORE4_1_2-總分
            Dim rSCORE4_1_2 As String = Convert.ToString(drV("SCORE4_1_2"))

            '寫值
            sheet.Cells(i_currentRow, 1).Value = rCOMIDNO
            sheet.Cells(i_currentRow, 2).Value = rROWID
            sheet.Cells(i_currentRow, 3).Value = rOrgName
            sheet.Cells(i_currentRow, 4).Value = rORGKIND1_N 'rORGKIND_N
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

            sheet.Cells(i_currentRow, 17).Value = rSCORE4_1 '配合分署辦理相關活動或政策宣導(3%)
            sheet.Cells(i_currentRow, 18).Value = rSCORE4_2

            sheet.Cells(i_currentRow, 19).Value = rSUBTOTAL '"0" '小計
            sheet.Cells(i_currentRow, 20).Value = rIMPLEVEL_1 '"0" '初擬等級
            sheet.Cells(i_currentRow, 21).Value = "" 'rSCORE4_1_2

            '下列 本程式功能沒有介面
            sheet.Cells(i_currentRow, 22).Value = "" '合計		
            sheet.Cells(i_currentRow, 23).Value = "" '等級
            sheet.Cells(i_currentRow, 24).Value = "" '備註

            sheet.Cells(i_currentRow, 1, i_currentRow, 24).Style.Font.Size = 9
        Next
        'ep.Save()
        'Dim parmsExp As New Hashtable
        'parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        'parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        'parmsExp.Add("strHTML", strHTML)
        'parmsExp.Add("ResponseNoEnd", "Y")
        'TIMS.Utl_ExportRp1(Me, parmsExp)

        Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
        Select Case V_ExpType
            Case "EXCEL"
                ExpExccl_1(strErrmsg, ep)
            Case "ODS"
                'Const Cst_FileSavePath As String="~/CO/01/Temp/"
                'Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
                Dim myFileName1 As String = TIMS.ClearSQM(String.Concat("審查計分表-", TIMS.GetRnd6Eng(), ".xlsx")) '檔名
                Dim myFileName2 As String = Cst_FileSavePath & myFileName1 '複製

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

                Dim sFileName1 As String = "ExpFile" & TIMS.GetRnd6Eng()

                'parmsExp.Add("strHTML", strHTML)
                Dim parmsExp As New Hashtable From {
                    {"ExpType", TIMS.GetListValue(RBListExpType)}, 'EXCEL/PDF/ODS
                    {"FileName", sFileName1},
                    {"xlsx_buf", buf},
                    {"ResponseNoEnd", "Y"}
                }
                TIMS.Utl_ExportRp1(Me, parmsExp)
            Case Else
                Dim s_log1 As String = ""
                s_log1 = String.Format("ExpType(參數有誤)!!{0}", V_ExpType)
                Common.MessageBox(Me, s_log1)
                Exit Sub
        End Select

        '刪除Temp中的資料
        'Call TIMS.MyFileDelete(myFileName1)
        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Return
        End If
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    ''' <summary>匯出使用(xlsx) ExcelPackage</summary>
    ''' <param name="strErrmsg"></param>
    ''' <param name="ep"></param>
    Sub ExpExccl_1(ByRef strErrmsg As String, ByRef ep As ExcelPackage)
        '將新建立的excel存入記憶體下載-----   Start
        'Dim myFileName1 As String=TIMS.ClearSQM("ExpFile" & TIMS.GetRnd6Eng) & cst_files_ext '檔名
        'Dim createStream As FileStream=New FileStream(sMyFile2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
        'ep.SaveAs(createStream) '存檔
        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim myFileName1 As String = TIMS.ClearSQM(String.Concat("審查計分表-", TIMS.GetRnd6Eng(), ".xlsx")) '檔名
        Dim myFileName2 As String = Cst_FileSavePath & myFileName1 '複製
        Dim sMyFile2 As String = Server.MapPath(myFileName2)
        Dim createStream As FileStream = New FileStream(sMyFile2, FileMode.Create, FileAccess.Write, FileShare.ReadWrite)
        ep.SaveAs(createStream) '存檔
        createStream.Close()
        createStream = Nothing
        'https://dotblogs.com.tw/malonestudyrecord/2018/03/21/103124

        '建立檔案
        Try
            'Dim fr As New System.IO.FileStream(sMyFile2, IO.FileMode.Open)
            Dim fr As New System.IO.FileStream(sMyFile2, IO.FileMode.Open)
            Dim br As New System.IO.BinaryReader(fr)
            Dim buf(fr.Length) As Byte
            fr.Read(buf, 0, fr.Length)
            fr.Close()

            Response.Clear()
            Response.ClearHeaders()
            Response.Buffer = True
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
    Protected Sub BtnImport1_Click(sender As Object, e As EventArgs) Handles btnImport1.Click
        If fg_trImport1_NoUse Then Return

        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_IMP As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        If Not fg_CAN_IMP Then
            Common.MessageBox(Me, $"{CST_NON_REVIEWSCORE}不可匯入!請再確認匯入參數!")
            Return
        End If
        Call ImportXLS_1(File1)
        '重新載入資訊
        Call CCreate1()
    End Sub

    ''' <summary>匯入檢核1 </summary>
    ''' <param name="ErrMsg1"></param>
    ''' <returns></returns>
    Function CheckImp1(ByRef ErrMsg1 As String) As Boolean
        Dim rst As Boolean = False
        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        If v_ddlDISTID = "" Then
            ErrMsg1 &= "分署未選擇，無法匯入，請先選擇分署!" & vbCrLf
            'Common.MessageBox(Me, "分署未選擇，無法匯入，請先選擇分署!")
            Return rst
        End If
        If v_ddlSCORING = "" Then
            ErrMsg1 &= "審查計分區間未選擇，無法匯入，請先選擇審查計分區間!" & vbCrLf
            Return rst
        End If
        rst = True
        Return rst
    End Function

    ''' <summary>'匯入等級/分數</summary>
    Sub ImportXLS_1(ByRef oFile1 As HtmlInputFile)
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp1(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        Dim htB1 As New Hashtable From {{"IMP", "1"}, {"DISTID", v_ddlDISTID}, {"SCORING", v_ddlSCORING}}
        Const cst_Upload_Path As String = "~/CO/01/Temp/" '暫存路徑
        Call TIMS.MyCreateDir(Me, cst_Upload_Path)

        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, oFile1, MyPostedFile, Cst_Filetype) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""
        '檢查檔案格式與大小 Start
        If oFile1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        End If
        If oFile1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If

        '取出檔案名稱
        MyFileName = Split(oFile1.PostedFile.FileName, "\")((Split(oFile1.PostedFile.FileName, "\")).Length - 1)
        'FileOCIDValue=Split(Split(MyFileName, "-")(1), ".")(0)

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

        'Dim dt_xls As DataTable
        Dim Errmag As String = ""

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        oFile1.PostedFile.SaveAs(Server.MapPath(cst_Upload_Path & MyFileName)) '上傳檔案
        'dt_xls = TIMS.GetDataTable_XlsFile(Server.MapPath(cst_Upload_Path & MyFileName).ToString, "", Errmag, "統一編號") '取得內容
        Dim dt_xls As DataTable = TIMS.ReadExceldtT1(Server.MapPath(cst_Upload_Path & MyFileName), Errmag) '取得內容
        IO.File.Delete(Server.MapPath(cst_Upload_Path & MyFileName)) '刪除檔案

        If Errmag <> "" Then
            Errmag &= "資料有誤，故無法匯入，請修正Excel檔案!"
            Common.MessageBox(Me, Errmag)
            Exit Sub
        End If
        If dt_xls Is Nothing Then '有資料
            Common.MessageBox(Me, "資料為空，故無法匯入，請修正Excel檔案!")
            Exit Sub
        ElseIf dt_xls.Rows.Count = 0 Then '有資料
            Common.MessageBox(Me, "查無資料，故無法匯入，請修正Excel檔案!!")
            Exit Sub
        End If

        '建立錯誤資料格式Table Start 'Dim Reason As String '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("COMIDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        Dim iRowIndex As Integer = 0 '讀取行累計數
        Dim Reason As String = "" '做一次驗証的即可
        If Reason = "" Then
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                Reason = ""
                Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                Reason = SAVE_ORG_SCORING2(colArray, htB1)  '驗証 並 儲存
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = $"第{iRowIndex + 2}列"
                    drWrong("COMIDNO") = s_col_COMIDNO '統一編號
                    drWrong("Reason") = If(Reason <> "", Reason, "(錯誤)") 'Reason
                End If
                iRowIndex += 1 '讀取行累計數
            Next 'Loop
        End If

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""

        '開始判別欄位存入   End
        If dtWrong.Rows.Count = 0 Then
            explain = ""
            explain &= "匯入資料共" & iRowIndex & "筆" & vbCrLf
            explain &= "成功：" & (iRowIndex - dtWrong.Rows.Count) & "筆" & vbCrLf
            explain &= "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Reason = "錯誤訊息如下:" & vbCrLf & Reason
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            explain2 = ""
            explain2 &= "匯入資料共" & iRowIndex & "筆\n"
            explain2 &= "成功：" & (iRowIndex - dtWrong.Rows.Count) & "筆\n"
            explain2 &= "失敗：" & dtWrong.Rows.Count & "筆\n"
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('CO_01_004_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If

        'sr.Close()
        'srr.Close()
        'MyFile.Delete(Server.MapPath(Upload_Path & MyFileName)) '刪除暫存檔案
    End Sub


    '只檢核統編，是否有資料 ORG_SCORING2, 有資料時回傳 o_parms
    Function CHK_SCORING2(ByRef i_parms As Hashtable, ByRef Htb As Hashtable, ByRef s_OSID2 As String) As Boolean
        Dim rst As Boolean = False
        If Htb Is Nothing Then Return rst
        Dim vDISTID As String = TIMS.GetMyValue2(Htb, "DISTID")
        Dim vSCORING As String = TIMS.GetMyValue2(Htb, "SCORING")
        If i_parms Is Nothing Then Return rst
        Dim vCOMIDNO As String = TIMS.GetMyValue2(i_parms, "COMIDNO")

        Dim parms As New Hashtable From {{"COMIDNO", vCOMIDNO}, {"DISTID", vDISTID}, {"SCORING", vSCORING}}
        Dim sql As String = ""
        sql &= " SELECT OSID2,ORGID FROM ORG_SCORING2" & vbCrLf
        sql &= " WHERE COMIDNO=@COMIDNO AND DISTID =@DISTID" & vbCrLf
        sql &= " AND CONCAT(YEARS,'-',MONTHS,'-',YEARS1,'-',HALFYEAR1,'-',YEARS2,'-',HALFYEAR2)=@SCORING" & vbCrLf
        'sql &= " AND COMIDNO='82198634'" & vbCrLf
        'sql &= " AND CONCAT(YEARS ,'-',MONTHS,'-',YEARS1 ,'-',HALFYEAR1,'-',YEARS2 ,'-',HALFYEAR2)='2021-01-2020-1-2020-2'" & vbCrLf
        'sql &= " AND DISTID ='001'" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If TIMS.dtNODATA(dt) Then Return rst

        Dim dr As DataRow = dt.Rows(0)
        s_OSID2 = $"{dr("OSID2")}"
        'If o_parms Is Nothing Then o_parms=New Hashtable 'o_parms.Add("OSID2", dr("OSID2")) 'o_parms.Add("ORGID", dr("ORGID"))
        rst = True
        Return rst
    End Function

    ''' <summary>(匯入) IMPORT '驗証 並 儲存</summary>
    ''' <param name="colArray"></param>
    ''' <param name="Htb"></param>
    ''' <returns></returns>
    Function SAVE_ORG_SCORING2(ByRef colArray As Array, ByRef Htb As Hashtable) As String
        Dim o_parms As New Hashtable
        Dim rst As String = CheckImportData(colArray, Htb, o_parms)
        If rst <> "" Then Return rst
        If Htb Is Nothing Then
            rst = TIMS.cst_NODATAMsg3
            Return rst
        End If
        'IMP:(匯入種類)1:匯入等級/分數 (暫不使用)／2:分署加分匯入
        Dim vIMP As String = TIMS.GetMyValue2(Htb, "IMP")
        Dim vDISTID As String = TIMS.GetMyValue2(Htb, "DISTID")
        Dim vSCORING As String = TIMS.GetMyValue2(Htb, "SCORING")
        If o_parms Is Nothing Then
            rst = TIMS.cst_NODATAMsg3
            Return rst
        End If
        Dim s_OSID2 As String = TIMS.GetMyValue2(o_parms, "OSID2")
        Dim vCOMIDNO As String = TIMS.GetMyValue2(o_parms, "COMIDNO")
        'Dim vSCORE1 As String=TIMS.GetMyValue2(o_parms, "SCORE1")
        Dim vLEVEL1 As String = TIMS.GetMyValue2(o_parms, "LEVEL1") '匯入等級
        'RLEVEL_2 複審等級'有複審等級，使用複審等級，複審等級為空，使用匯入等級
        Dim vRLEVEL2 As String = TIMS.GetMyValue2(o_parms, "RLEVEL2")
        Dim vSCORE4_1 As String = TIMS.GetMyValue2(o_parms, "SCORE4_1") '分署加分
        If vRLEVEL2 = "" Then vRLEVEL2 = vLEVEL1 '複審等級
        Dim vMINISTERLEVEL As String = TIMS.GetMyValue2(o_parms, "MINISTERLEVEL")
        If vMINISTERLEVEL = "" Then vMINISTERLEVEL = vLEVEL1 '部加分等級

        Select Case vIMP
            Case "1"
                'u_parms.Add("IMPSCORE_1", vSCORE1)
                Dim u_parms As New Hashtable From {
                    {"IMPLEVEL_1", vLEVEL1},
                    {"RLEVEL_2", vRLEVEL2},  'RLEVEL_2 複審等級
                    {"MINISTERLEVEL", vMINISTERLEVEL},  '部加分等級
                    {"IMODIFYACCT", sm.UserInfo.UserID},
                    {"OSID2", TIMS.CINT1(s_OSID2)},
                    {"COMIDNO", vCOMIDNO},
                    {"DISTID", vDISTID}
                }
                Dim u_sql As String = ""
                u_sql &= " UPDATE ORG_SCORING2" & vbCrLf
                u_sql &= " SET IMPSCORE_1=@IMPSCORE_1 ,IMPLEVEL_1=@IMPLEVEL_1 ,RLEVEL_2=@RLEVEL_2" & vbCrLf  'RLEVEL_2 複審等級
                u_sql &= " ,MINISTERLEVEL=@MINISTERLEVEL" & vbCrLf  '部加分等級
                u_sql &= " ,IMODIFYACCT=@IMODIFYACCT,IMODIFYDATE=GETDATE()" & vbCrLf
                u_sql &= " WHERE OSID2=@OSID2 AND COMIDNO=@COMIDNO AND DISTID=@DISTID" & vbCrLf
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

            Case "2"
                Dim u_parms As New Hashtable From {
                    {"SCORE4_1", vSCORE4_1}, '分署加分
                    {"IMODIFYACCT", sm.UserInfo.UserID},
                    {"OSID2", TIMS.CINT1(s_OSID2)},
                    {"COMIDNO", vCOMIDNO},
                    {"DISTID", vDISTID}
                }
                Dim u_sql As String = ""
                u_sql &= " UPDATE ORG_SCORING2" & vbCrLf
                u_sql &= " SET SCORE4_1=@SCORE4_1" & vbCrLf
                u_sql &= " ,IMODIFYACCT=@IMODIFYACCT ,IMODIFYDATE=GETDATE()" & vbCrLf
                u_sql &= " WHERE OSID2=@OSID2 AND COMIDNO=@COMIDNO AND DISTID=@DISTID" & vbCrLf
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

            Case "3"
                Dim u_parms As New Hashtable From {
                    {"IMPLEVEL_1", vLEVEL1}, '初擬等級
                    {"RLEVEL_2", vRLEVEL2},  'RLEVEL_2 複審等級
                    {"MINISTERLEVEL", vMINISTERLEVEL},  '部加分等級
                    {"IMODIFYACCT", sm.UserInfo.UserID},
                    {"OSID2", TIMS.CINT1(s_OSID2)},
                    {"COMIDNO", vCOMIDNO},
                    {"DISTID", vDISTID}
                }
                Dim u_sql As String = ""
                u_sql &= " UPDATE ORG_SCORING2" & vbCrLf
                u_sql &= " SET IMPLEVEL_1=@IMPLEVEL_1 ,RLEVEL_2=@RLEVEL_2" & vbCrLf   '初擬等級／'RLEVEL_2 複審等級
                u_sql &= " ,MINISTERLEVEL=@MINISTERLEVEL" & vbCrLf  '部加分等級
                u_sql &= " ,IMODIFYACCT=@IMODIFYACCT ,IMODIFYDATE=GETDATE()" & vbCrLf
                u_sql &= " WHERE OSID2=@OSID2 AND COMIDNO=@COMIDNO AND DISTID=@DISTID" & vbCrLf
                DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        End Select
        Return rst
    End Function

    ''' <summary>匯入驗証</summary>
    ''' <param name="colArray">比對資料</param>
    ''' <param name="Htb">輸入查詢</param>
    ''' <param name="o_parms">取得有效值</param>
    ''' <returns></returns>
    Function CheckImportData(ByRef colArray As Array, ByRef Htb As Hashtable, ByRef o_parms As Hashtable) As String
        Dim Reason As String = ""
        'clear
        s_col_COMIDNO = ""
        s_col_SCORE1 = ""
        s_col_LEVEL1 = ""
        s_col_SCORE4_1 = ""

        'IMP:(匯入種類)1:匯入等級/分數 (暫不使用)／2:分署加分匯入
        Dim vIMP As String = TIMS.GetMyValue2(Htb, "IMP")
        Dim vDISTID As String = TIMS.GetMyValue2(Htb, "DISTID")
        Dim vSCORING As String = TIMS.GetMyValue2(Htb, "SCORING")

        '欄位最大數 1:3／2:2／3:2
        Dim i_MaxColumnCount1 As Integer = If(vIMP = "1", 3, If(vIMP = "2", 2, If(vIMP = "3", 2, 0)))
        If colArray.Length < i_MaxColumnCount1 Then
            'Reason += "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            Return Reason
        End If

        'IMP:(匯入種類)1:匯入等級/分數 (暫不使用)／2:分署加分匯入
        Select Case vIMP
            Case "1"
                s_col_COMIDNO = TIMS.ClearSQM(colArray(cst_col1_COMIDNO)) '統一編號
                s_col_SCORE1 = TIMS.ClearSQM(colArray(cst_col1_SCORE1)) '匯入分數
                s_col_LEVEL1 = TIMS.ClearSQM(colArray(cst_col1_LEVEL1)) '匯入等級

                '先確認資料不為空
                If s_col_COMIDNO = "" Then Reason += "統一編號不可為空<br>"
                If s_col_SCORE1 = "" Then Reason += "匯入分數不可為空<br>"
                If s_col_LEVEL1 = "" Then Reason += "匯入等級不可為空<br>"
                If Reason <> "" Then Return Reason

                If Not TIMS.IsNumeric1(s_col_SCORE1) Then
                    Reason += String.Format("匯入分數有誤(數字格式): {0}<br>", s_col_SCORE1)
                    Return Reason
                ElseIf Val(s_col_SCORE1) > 100 OrElse Val(s_col_SCORE1) < 0 Then
                    Reason += String.Format("匯入分數有誤(數字0-100): {0}<br>", s_col_SCORE1)
                    Return Reason
                End If
                'TIMS.cst_SCORELEVEL_all 'As String="A,B,C,D" '審核等級
                If Not TIMS.cst_SCORELEVEL_all.Contains(s_col_LEVEL1) Then
                    Reason += String.Format("匯入等級有誤: {0}<br>", s_col_LEVEL1)
                    Return Reason
                End If

            Case "2"
                s_col_COMIDNO = TIMS.ClearSQM(colArray(cst_col2_COMIDNO)) '統一編號
                s_col_SCORE4_1 = TIMS.ClearSQM(colArray(cst_col2_SCORE4_1)) '分署加分

                '先確認資料不為空
                If s_col_COMIDNO = "" Then Reason += "統一編號不可為空<br>"
                If s_col_SCORE4_1 = "" Then Reason += "分署加分不可為空<br>"
                If Reason <> "" Then Return Reason

                Dim fg_match As Match = Regex.Match(s_col_SCORE4_1, "^[0-3]+(.[0-9]{1})?$")
                If Not TIMS.IsNumeric1(s_col_SCORE4_1) Then
                    Reason += String.Format("分署加分-匯入數字有誤(數字格式): {0}<br>", s_col_SCORE4_1)
                    Return Reason
                ElseIf Val(s_col_SCORE4_1) > Val(3.0) OrElse Val(s_col_SCORE4_1) < Val(0.0) Then
                    Reason += String.Format("分署加分-匯入數字有誤!(數字0.0-3.0): {0}<br>", s_col_SCORE4_1)
                    Return Reason
                ElseIf Not fg_match.Success Then
                    Reason += String.Format("分署加分-匯入數字有誤!!(數字0.0-3.0): {0}<br>", s_col_SCORE4_1)
                    Return Reason
                End If

            Case "3"
                s_col_COMIDNO = TIMS.ClearSQM(colArray(cst_col3_COMIDNO)) '統一編號
                s_col_LEVEL1 = TIMS.ClearSQM(colArray(cst_col3_LEVEL1)) '匯入等級
                '先確認資料不為空
                If s_col_COMIDNO = "" Then Reason += "統一編號不可為空<br>"
                If s_col_LEVEL1 = "" Then Reason += "匯入初擬等級不可為空<br>"
                If Reason <> "" Then Return Reason
                'TIMS.cst_SCORELEVEL_all 'As String="A,B,C,D" '審核等級
                If Not TIMS.cst_SCORELEVEL_all.Contains(s_col_LEVEL1) Then
                    Reason += String.Format("匯入初擬等級有誤: {0}<br>", s_col_LEVEL1)
                    Return Reason
                End If

            Case Else
                Reason = "[匯入檔]選擇有誤，請重新選擇!(程式目前有n種匯入格式)<br>"
                Return Reason

        End Select

        Dim i_parms As New Hashtable From {{"COMIDNO", s_col_COMIDNO}}
        Dim s_OSID2 As String = ""
        Dim flag_chkOK As Boolean = CHK_SCORING2(i_parms, Htb, s_OSID2)
        If Not flag_chkOK OrElse s_OSID2 = "" Then
            Reason += "查無匯入資料，依「統編、分署、審查計分區間」<br>"
            Return Reason
        End If

        If o_parms Is Nothing Then o_parms = New Hashtable '(if Nothing -> New Hashtable )
        o_parms.Add("OSID2", TIMS.CINT1(s_OSID2))
        o_parms.Add("COMIDNO", s_col_COMIDNO)
        o_parms.Add("SCORE1", s_col_SCORE1)
        o_parms.Add("LEVEL1", s_col_LEVEL1)
        'RLEVEL_2 複審等級  '有複審等級，使用複審等級，複審等級為空，使用匯入等級
        o_parms.Add("RLEVEL2", s_col_LEVEL1) '複審等級
        o_parms.Add("MINISTERLEVEL", s_col_LEVEL1) '部加分等級
        o_parms.Add("SCORE4_1", s_col_SCORE4_1) '分署加分
        Return Reason
    End Function

    Protected Sub BtnImport2_Click(sender As Object, e As EventArgs) Handles btnImport2.Click
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp2(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_IMP As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        If Not fg_CAN_IMP Then
            Common.MessageBox(Me, $"{CST_NON_REVIEWSCORE}不可匯入!請再確認匯入參數!")
            Return
        End If
        Call ImportXLS_2(File2)
        '重新載入資訊
        Call CCreate1()
    End Sub

    ''' <summary> 匯入動作2</summary>
    Private Sub ImportXLS_2(ByRef oFile1 As HtmlInputFile)
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp2(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        Dim htB2 As New Hashtable
        htB2.Add("IMP", "2")
        htB2.Add("DISTID", v_ddlDISTID)
        htB2.Add("SCORING", v_ddlSCORING)

        Const cst_Upload_Path As String = "~/CO/01/Temp/" '暫存路徑
        Call TIMS.MyCreateDir(Me, cst_Upload_Path)

        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, oFile1, MyPostedFile, Cst_Filetype, 1) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        '檢查檔案格式與大小 Start
        If oFile1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        End If
        If oFile1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If

        '取出檔案名稱
        MyFileName = Split(oFile1.PostedFile.FileName, "\")((Split(oFile1.PostedFile.FileName, "\")).Length - 1)
        'FileOCIDValue=Split(Split(MyFileName, "-")(1), ".")(0)

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
        '檢查檔案格式與大小----------   End

        Dim Errmag As String = ""
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        oFile1.PostedFile.SaveAs(Server.MapPath(cst_Upload_Path & MyFileName)) '上傳檔案
        'dt_xls = TIMS.GetDataTable_XlsFile(Server.MapPath(cst_Upload_Path & MyFileName).ToString, "", Errmag, "統一編號") '取得內容
        Dim dt_xls As DataTable = TIMS.ReadExceldtT1(Server.MapPath(cst_Upload_Path & MyFileName), Errmag) '取得內容
        IO.File.Delete(Server.MapPath(cst_Upload_Path & MyFileName)) '刪除檔案
        If Errmag <> "" Then
            Errmag &= "資料有誤，故無法匯入，請修正Excel檔案!"
            Common.MessageBox(Me, Errmag)
            Exit Sub
        End If
        If dt_xls Is Nothing Then '有資料
            Common.MessageBox(Me, "資料為空，故無法匯入，請修正Excel檔案!")
            Exit Sub
        ElseIf dt_xls.Rows.Count = 0 Then '有資料
            Common.MessageBox(Me, "查無資料，故無法匯入，請修正Excel檔案!!")
            Exit Sub
        End If

        '取出資料庫的所有欄位    Start
        'Dim sql As String=""
        'Dim da As SqlDataAdapter=Nothing

        '建立錯誤資料格式Table Start 'Dim Reason As String '儲存錯誤的原因
        Dim dtWrong As New DataTable            '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("COMIDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        Dim iRowIndex As Integer = 0 '讀取行累計數
        Dim Reason As String = "" '做一次驗証的即可
        If Reason = "" Then
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                Reason = ""
                Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                Reason = SAVE_ORG_SCORING2(colArray, htB2)  '驗証 並 儲存
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    drWrong = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = String.Concat("第", CStr(iRowIndex + 2), "列")
                    drWrong("COMIDNO") = s_col_COMIDNO '統一編號
                    drWrong("Reason") = If(Reason <> "", Reason, "(錯誤)") 'Reason
                End If
                iRowIndex += 1 '讀取行累計數
            Next 'Loop
        End If

        '判斷匯出資料是否有誤
        Dim explain As String = ""
        Dim explain2 As String = ""

        '開始判別欄位存入   End
        If dtWrong.Rows.Count = 0 Then
            explain = ""
            explain &= "匯入資料共" & iRowIndex & "筆" & vbCrLf
            explain &= "成功：" & (iRowIndex - dtWrong.Rows.Count) & "筆" & vbCrLf
            explain &= "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Reason = "錯誤訊息如下:" & vbCrLf & Reason
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            explain2 = ""
            explain2 &= "匯入資料共" & iRowIndex & "筆\n"
            explain2 &= "成功：" & (iRowIndex - dtWrong.Rows.Count) & "筆\n"
            explain2 &= "失敗：" & dtWrong.Rows.Count & "筆\n"
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('CO_01_004_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
    End Sub

    ''' <summary> 匯入檢核 </summary>
    ''' <param name="errMsg1"></param>
    ''' <returns></returns>
    Private Function CheckImp2(ByRef errMsg1 As String) As Boolean
        Dim rst As Boolean = False
        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        If v_ddlDISTID = "" Then
            errMsg1 &= "分署未選擇，無法匯入，請先選擇分署!" & vbCrLf
            'Common.MessageBox(Me, "分署未選擇，無法匯入，請先選擇分署!")
            Return rst
        End If
        If v_ddlSCORING = "" Then
            errMsg1 &= "審查計分區間未選擇，無法匯入，請先選擇審查計分區間!" & vbCrLf
            Return rst
        End If
        rst = True
        Return rst
    End Function

    ''' <summary>匯出單位計分</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp2_Click(sender As Object, e As EventArgs) Handles btnExp2.Click
        Call SExprot2_Y2()
    End Sub

    ''' <summary>匯出單位計分</summary>
    Private Sub SExprot2_Y2()
        Dim vDISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        Dim vORGKIND2 As String = TIMS.ClearSQM(OrgPlanKind.SelectedValue) '計畫
        Dim vORGKIND As String = TIMS.ClearSQM(OrgKindList.SelectedValue) '機構別

        Dim eErrMsg1 As String = ""
        If vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        'parms.Add("YEARS", sm.UserInfo.Years)
        Dim parms As New Hashtable From {
            {"EXP", "Y2"}, '匯出查詢條件
            {"TPLANID", sm.UserInfo.TPlanID},
            {"DISTID", vDISTID}, 'sql &= " AND t.DISTID=@DISTID" & vbCrLf
            {"SCORINGID", vSCORINGID}
        }
        'parms.Add("YEARS", vYEARS)
        'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf

        Dim dtXls As DataTable = Get_dtORGSCORING2(parms)
        If TIMS.dtNODATA(dtXls) Then
            Common.MessageBox(Me, "查無匯出資料!!")
            Exit Sub
        End If

        '匯出單位計分
        Call ExpXMLXLS2_2(dtXls)
    End Sub

    ''' <summary>匯出單位計分</summary>
    ''' <param name="dtXls"></param>
    Sub ExpXMLXLS2_2(ByRef dtXls As DataTable)
        Dim sPattern As String = ""
        sPattern &= "序號,分署,訓練單位,實際開班數,政策性課程班數,核定總班數,開班率(1-1),實際開訓人數,政策性課程核定人次"
        sPattern &= ",核定總人次,訓練人次達成率(1-2),招訓資料時效,開訓資料時效,結訓資料時效,變更申請時效,各項資料時效分數(2-1-1)"
        sPattern &= ",招訓資料內容(總扣分),開訓資料內容(總扣分),結訓資料內容(總扣分),變更申請內容(總扣分),資料內容正確性(2-1-2),訓練計畫變更次數分數(2-1-3),學員管理(2-2-1),學員管理得分"
        sPattern &= ",課程辦理情形(2-2-2)(總扣分),課程辦理情形得分,計畫參與度(2-3),計畫參與度得分,TTQS評核結果(3-1),TTQS得分,學員滿意度(3-2),學員滿意度得分"
        sPattern &= ",訓後動態調查表填答率,訓後調查得分"

        'SCORE2_1_1_ITEM, SCORE2_2_1_EQU, SCORE2_2_2_EQU, SCORE2_3_1_EQU
        Dim sColumn As String = ""
        sColumn &= "SEQNO,DISTNAME,ORGNAME,CLSACTCNT,CLSACTCNT2,CLSAPPCNT,SCORE1_1A,STDACTCNT,STDACTCNT2"
        sColumn &= ",STDAPPCNT,SCORE1_2A,SCORE2_1_1_SUM_A,SCORE2_1_1_SUM_B,SCORE2_1_1_SUM_C,SCORE2_1_1_SUM_D,SCORE2_1_1_ALL"
        sColumn &= ",SCORE2_1_2A_DIS,SCORE2_1_2B_DIS,SCORE2_1_2C_DIS,SCORE2_1_2D_DIS,SCORE2_1_2_SUM_ALL,SCORE2_1_3,SCORE2_2_1_EQU,SCORE2_2_1"
        sColumn &= ",SCORE2_2_2_DIS,SCORE2_2_2,SCORE2_3_1_EQU,SCORE2_3_1,SCORE3_1_N,SCORE3_1,SCORE3_2_EQU,SCORE3_2"
        sColumn &= ",SCORE4_2_RATE,SCORE4_2"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")

        Dim sFileName1 As String = String.Concat("匯出單位計分", TIMS.GetDateNo())

        '套CSS值
        Dim strSTYLE As String = String.Concat("<style>", "td{mso-number-format:""\@"";}", ".noDecFormat{mso-number-format:""0"";}", "</style>")

        Dim sbHTML As New StringBuilder
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '建立輸出文字
        Dim ExportStr As String = ""
        '標題抬頭1
        ExportStr = String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFileName1) '& vbTab
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '標題抬頭2
        ExportStr = ""
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
        Next
        sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

        '建立資料面
        Dim iRows As Integer = 0
        For Each dr As DataRow In dtXls.Rows
            iRows += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                Dim sCOLTXT As String = ""
                Select Case sColumnA(i)
                    Case "SEQNO"
                        sCOLTXT = iRows.ToString()
                    Case Else
                        sCOLTXT = Convert.ToString(dr(sColumnA(i)))
                End Select
                ExportStr &= String.Format("<td>{0}</td>", sCOLTXT)
            Next
            ExportStr &= "</tr>" & vbCrLf
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
        'parmsExp.Add("ExpType", "EXCEL") 'EXCEL/PDF/ODS
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.CloseDbConn(objconn)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    ''' <summary>匯入初擬等級</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnImport3_Click(sender As Object, e As EventArgs) Handles btnImport3.Click
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp3(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim fg_CAN_IMP As Boolean = TIMS.CHK_ORG_TTQSLOCK_SCORINGID(objconn, vSCORINGID)
        If Not fg_CAN_IMP Then
            Common.MessageBox(Me, $"{CST_NON_REVIEWSCORE}不可匯入!請再確認匯入參數!")
            Return
        End If
        Call ImportXLS_3(File3)
        '重新載入資訊
        Call CCreate1()
    End Sub

    ''' <summary>檢核-匯入初擬等級</summary>
    ''' <param name="errMsg1"></param>
    ''' <returns></returns>
    Private Function CheckImp3(ByRef errMsg1 As String) As Boolean
        Dim rst As Boolean = False
        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        If v_ddlDISTID = "" Then
            errMsg1 &= "分署未選擇，無法匯入，請先選擇分署!" & vbCrLf
            Return rst
        End If
        If v_ddlSCORING = "" Then
            errMsg1 &= "審查計分區間未選擇，無法匯入，請先選擇審查計分區間!" & vbCrLf
            Return rst
        End If
        rst = True
        Return rst
    End Function

    ''' <summary>匯入初擬等級</summary>
    ''' <param name="oFile1"></param>
    Private Sub ImportXLS_3(ByRef oFile1 As HtmlInputFile)
        Dim ErrMsg1 As String = ""
        Dim flag_OK As Boolean = CheckImp3(ErrMsg1)
        If ErrMsg1 <> "" Then
            Common.MessageBox(Me, ErrMsg1)
            Return
        End If
        If Not flag_OK Then
            Common.MessageBox(Me, "匯入檢核有誤!請再確認匯入參數!")
            Return
        End If

        Dim v_ddlDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim v_ddlSCORING As String = TIMS.GetListValue(ddlSCORING)
        Dim htB3 As New Hashtable From {{"IMP", "3"}, {"DISTID", v_ddlDISTID}, {"SCORING", v_ddlSCORING}}
        Const cst_Upload_Path As String = "~/CO/01/Temp/" '暫存路徑
        Call TIMS.MyCreateDir(Me, cst_Upload_Path)
        Const Cst_Filetype As String = "xls" '匯入檔案類型
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If Not TIMS.HttpCHKFile(Me, oFile1, MyPostedFile, Cst_Filetype, 2) Then Return

        Dim MyFileName As String = ""
        Dim MyFileType As String = ""

        '檢查檔案格式與大小 Start
        If oFile1.Value = "" Then
            Common.MessageBox(Me, "未輸入匯入檔案位置!!")
            Exit Sub
        ElseIf oFile1.PostedFile.ContentLength = 0 Then
            Common.MessageBox(Me, "檔案位置錯誤!")
            Exit Sub
        End If
        '取出檔案名稱 / '取出檔案類型 
        MyFileName = Split(oFile1.PostedFile.FileName, "\")((Split(oFile1.PostedFile.FileName, "\")).Length - 1)
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
        'Dim dt_xls As DataTable
        Dim Errmag As String = ""
        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(oFile1.PostedFile.FileName).ToLower()
        MyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        oFile1.PostedFile.SaveAs(Server.MapPath(cst_Upload_Path & MyFileName)) '上傳檔案
        'dt_xls = TIMS.GetDataTable_XlsFile(Server.MapPath(cst_Upload_Path & MyFileName).ToString, "", Errmag, "統一編號") '取得內容
        Dim dt_xls As DataTable = TIMS.ReadExceldtT1(Server.MapPath(cst_Upload_Path & MyFileName), Errmag) '取得內容
        IO.File.Delete(Server.MapPath(cst_Upload_Path & MyFileName)) '刪除檔案
        If Errmag <> "" Then
            Errmag &= "資料有誤，故無法匯入，請修正Excel檔案!"
            Common.MessageBox(Me, Errmag)
            Exit Sub
        End If
        If dt_xls Is Nothing Then '有資料
            Common.MessageBox(Me, "資料為空，故無法匯入，請修正Excel檔案!")
            Exit Sub
        ElseIf dt_xls.Rows.Count = 0 Then '有資料
            Common.MessageBox(Me, "查無資料，故無法匯入，請修正Excel檔案!!")
            Exit Sub
        End If
        '建立錯誤資料格式Table Start '儲存錯誤的原因
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        'Dim drWrong As DataRow
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("COMIDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table End

        Dim iRowIndex As Integer = 0 '讀取行累計數
        Dim Reason As String = "" '做一次驗証的即可
        If Reason = "" Then
            For i As Integer = 0 To dt_xls.Rows.Count - 1
                Reason = ""
                Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
                Reason = SAVE_ORG_SCORING2(colArray, htB3)  '驗証 並 儲存
                If Reason <> "" Then
                    '錯誤資料，填入錯誤資料表
                    Dim drWrong As DataRow = dtWrong.NewRow
                    dtWrong.Rows.Add(drWrong)
                    drWrong("Index") = String.Concat("第", CStr(iRowIndex + 2), "列")
                    drWrong("COMIDNO") = s_col_COMIDNO '統一編號
                    drWrong("Reason") = If(Reason <> "", Reason, "(錯誤)") 'Reason
                End If
                iRowIndex += 1 '讀取行累計數
            Next 'Loop
        End If

        '判斷匯出資料是否有誤 'Dim explain As String="" 'Dim explain2 As String="" '開始判別欄位存入   End
        If dtWrong.Rows.Count = 0 Then
            Dim explain As String = ""
            explain &= "匯入資料共" & iRowIndex & "筆" & vbCrLf
            explain &= "成功：" & (iRowIndex - dtWrong.Rows.Count) & "筆" & vbCrLf
            explain &= "失敗：" & dtWrong.Rows.Count & "筆" & vbCrLf
            If Reason = "" Then
                Common.MessageBox(Me, explain)
            Else
                Reason = "錯誤訊息如下:" & vbCrLf & Reason
                Common.MessageBox(Me, explain & Reason)
            End If
        Else
            Dim explain2 As String = ""
            explain2 &= "匯入資料共" & iRowIndex & "筆\n"
            explain2 &= "成功：" & (iRowIndex - dtWrong.Rows.Count) & "筆\n"
            explain2 &= "失敗：" & dtWrong.Rows.Count & "筆\n"
            Session("MyWrongTable") = dtWrong
            Page.RegisterStartupScript("", "<script>if(confirm('" & explain2 & "是否要檢視原因?')){window.open('CO_01_004_Wrong.aspx','','width=500,height=500,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
        End If
    End Sub

    ''' <summary>匯出班級明細計分</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp3_Click(sender As Object, e As EventArgs) Handles btnExp3.Click
        Call SExprot3_Y3()
    End Sub

    ''' <summary>匯出班級明細計分</summary>
    Private Sub SExprot3_Y3()
        Dim vDISTID As String = TIMS.ClearSQM(ddlDISTID.SelectedValue)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING)
        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        Dim vORGKIND2 As String = TIMS.ClearSQM(OrgPlanKind.SelectedValue) '計畫
        Dim vORGKIND As String = TIMS.ClearSQM(OrgKindList.SelectedValue) '機構別
        'Dim vFIRSTCHK_SCH As String = TIMS.GetListValue(rblFIRSTCHK_SCH)

        '匯出班級明細計分
        Dim eErrMsg1 As String = ""
        If vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        '匯出查詢條件-匯出班級明細計分 'parms.Add("YEARS", sm.UserInfo.Years) 'parms.Add("YEARS", vYEARS)
        'sql &= " AND t.DISTID=@DISTID" & vbCrLf
        Dim parms As New Hashtable From {{"EXP", "Y3"}, {"TPLANID", sm.UserInfo.TPlanID}, {"DISTID", vDISTID}, {"SCORINGID", vSCORINGID}}
        'If vHALFYEAR <> "" Then parms.Add("HALFYEAR", vHALFYEAR) '1:上年度 /2:下年度
        If vORGNAME <> "" Then parms.Add("ORGNAME", vORGNAME) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        If vCOMIDNO <> "" Then parms.Add("COMIDNO", vCOMIDNO) 'sql &= " AND tr.COMIDNO=@COMIDNO" & vbCrLf
        Select Case vORGKIND2
            Case "G", "W"
                parms.Add("ORGKIND2", vORGKIND2) 'sql &= " AND o.ORGKIND2=@ORGKIND2" & vbCrLf
        End Select
        If vORGKIND <> "" Then parms.Add("ORGKIND", vORGKIND) 'sql &= " AND o.ORGKIND=@ORGKIND" & vbCrLf
        'Select Case vFIRSTCHK_SCH
        '    Case "Y", "N"
        '        parms.Add("FIRSTCHK", vFIRSTCHK_SCH)
        'End Select

        'iTYPE2: 1.班級資訊, 2.班級變更資訊, 3.班級不預告實地抽訪紀錄表
        Using dsXlsALL As New DataSet
            Using dtXls1 As DataTable = GET_Y3_dtCHN(parms, 1)
                If dtXls1 Is Nothing OrElse dtXls1.Rows.Count = 0 Then
                    Common.MessageBox(Me, "查無匯出資料!!")
                    Return ' Exit Sub
                End If
                dtXls1.TableName = "班級資訊"
                dsXlsALL.Tables.Add(dtXls1)

                Using dtXls2 As DataTable = GET_Y3_dtCHN(parms, 2)
                    dtXls2.TableName = "班級變更資訊"
                    dsXlsALL.Tables.Add(dtXls2)

                    Using dtXls3 As DataTable = GET_Y3_dtCHN(parms, 3)
                        dtXls3.TableName = "班級不預告實地抽訪紀錄表"
                        dsXlsALL.Tables.Add(dtXls3)

                        Dim sFileName1 As String = String.Concat("匯出班級明細計分", TIMS.GetDateNo())
                        'Dim s_titleRange As String="A1:AG1,A1:AG1,A1:AG1"
                        ExpClass1.Utl_Export1_XLSX(Me, dsXlsALL, sFileName1)
                        TIMS.Utl_RespWriteEnd(Me, objconn, "")
                        '匯出班級明細計分 'Call ExpXMLXLS3_1(dtXls)
                    End Using
                End Using
            End Using
        End Using
    End Sub

    ''' <summary>
    ''' '匯出統計表(Y4)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp4_Click(sender As Object, e As EventArgs) Handles btnExp4.Click
        Call SExprot4_Y4()
    End Sub

    ''' <summary>
    ''' '匯出等級比率統計表(Y5)
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnExp5_Click(sender As Object, e As EventArgs) Handles btnExp5.Click
        Call SExprot4_Y5()
    End Sub

    ''' <summary>
    ''' 匯出統計表
    ''' </summary>
    Private Sub SExprot4_Y4()
        Dim vDISTID As String = TIMS.GetListValue(ddlDISTID) '分署
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING) '審查計分區間
        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text) '訓練機構
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text) '機構統一編號
        Dim vORGKIND2 As String = TIMS.GetListValue(OrgPlanKind) '計畫別(產投／提升勞工自主)
        Dim vORGKIND As String = TIMS.GetListValue(OrgKindList) '機構別
        'Dim vFIRSTCHK_SCH As String = TIMS.GetListValue(rblFIRSTCHK_SCH) '初審審核狀態

        '匯出統計表
        Dim eErrMsg1 As String = ""
        If vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        If vORGKIND2 = "" Then eErrMsg1 &= "請選擇 計畫別(產投／提升勞工自主)" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        Dim sParms1 As New Hashtable From {{"EXP", "Y4"}, {"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}
        sParms1.Add("DISTID", vDISTID)
        sParms1.Add("ORGNAME", vORGNAME)
        sParms1.Add("COMIDNO", vCOMIDNO)
        sParms1.Add("ORGKIND2", vORGKIND2)
        sParms1.Add("ORGKIND", vORGKIND)
        'sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)

        Using dtXls1 As DataTable = Get_dtORGSCORING2_B(sParms1) ' DbAccess.GetDataTable(sSql, objconn, sParms1)
            If dtXls1 Is Nothing OrElse dtXls1.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無匯出資料!!")
                Return ' Exit Sub
            End If
            '○年度○半年○計畫審查計分統計表(○分署)
            Dim t_SCORING_N2 As String = GET_SCORING_N2_TT(sParms1)

            Dim strFileNM1 As String = String.Concat("匯出統計表x", TIMS.GetDateNo())
            Dim sPattern As String = "編號,分署,訓練單位名稱,負責人,等級,總分,備註"
            'Dim sColumn As String="ROWID,DISTNAME3,ORGNAME,MASTERNAME,RL2_IMP1,SCORE4_1_2,bMEMO"
            'IMPLEVEL_1 初審等級 'SUBTOTAL 初審分數
            Dim sColumn As String = "ROWID,DISTNAME3,ORGNAME,MASTERNAME,IMPLEVEL_1,SUBTOTAL,bMEMO"

            Dim sPatternA() As String = Split(sPattern, ",")
            Dim sColumnA() As String = Split(sColumn, ",")
            Dim iColSpanCount As Integer = sColumnA.Length

            Dim parms As New Hashtable
            parms.Add("ExpType", TIMS.GetListValue(RBListExpType))
            parms.Add("FileName", strFileNM1)
            parms.Add("TitleName", TIMS.ClearSQM(t_SCORING_N2))
            parms.Add("TitleColSpanCnt", iColSpanCount)
            parms.Add("sPatternA", sPatternA)
            parms.Add("sColumnA", sColumnA)
            TIMS.Utl_Export(Me, dtXls1, parms)
            TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End Using

    End Sub

    ''' <summary>
    ''' 匯出等級比率統計表
    ''' </summary>
    Private Sub SExprot4_Y5()
        Dim vDISTID As String = TIMS.GetListValue(ddlDISTID)
        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING) '審查計分區間
        Dim vORGNAME As String = TIMS.ClearSQM(OrgName.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO.Text)
        'Dim vORGKIND2 As String=TIMS.GetListValue(OrgPlanKind) '計畫別(產投／提升勞工自主)
        Dim vORGKIND As String = TIMS.GetListValue(OrgKindList) '機構別
        'Dim vFIRSTCHK_SCH As String = TIMS.GetListValue(rblFIRSTCHK_SCH) '初審審核狀態

        '匯出等級比率統計表
        Dim eErrMsg1 As String = ""
        If sm.UserInfo.LID <> 0 AndAlso vDISTID = "" Then eErrMsg1 &= "請選擇分署" & vbCrLf
        If vSCORINGID = "" Then eErrMsg1 &= "請選擇 審查計分區間" & vbCrLf
        'If vORGKIND2="" Then eErrMsg1 &= "請選擇 計畫別(產投／提升勞工自主)" & vbCrLf
        If eErrMsg1 <> "" Then
            Common.MessageBox(Me, eErrMsg1)
            Exit Sub
        End If

        Dim sParms1 As New Hashtable From {{"EXP", "Y5"}, {"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}
        sParms1.Add("DISTID", vDISTID)
        sParms1.Add("ORGNAME", vORGNAME)
        sParms1.Add("COMIDNO", vCOMIDNO)
        'sParms1.Add("ORGKIND2", vORGKIND2)
        sParms1.Add("ORGKIND", vORGKIND)
        'sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)

        Using dtXls1 As DataTable = Get_dtORGSCORING2_B(sParms1) ' DbAccess.GetDataTable(sSql, objconn, sParms1)
            If dtXls1 Is Nothing OrElse dtXls1.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無匯出資料!!")
                Return ' Exit Sub
            End If

            'Dim s_DISTNAME As String=String.Concat(TIMS.GET_DISTNAME(objconn, vDISTID), " 合計 ")
            'Dim s_FootHtml2 As String=""
            's_FootHtml2 &= "<tr>"
            's_FootHtml2 &= String.Format("<td colspan=3>{0}</td>", s_DISTNAME)
            's_FootHtml2 &= String.Format("<td colspan=2>{0}</td>", s_CNT2)
            's_FootHtml2 &= "</tr>"

            '○年度○半年○計畫審查計分統計表(○分署)
            Dim t_SCORING_N2 As String = GET_SCORING_N2_TT(sParms1)

            Dim s_ORGKIND2NM_G As String = GET_ORGKIND2NM("G")
            Dim s_ORGKIND2NM_W As String = GET_ORGKIND2NM("W")

            Dim strFileNM1 As String = String.Concat("匯出等級比率統計表x", TIMS.GetDateNo())
            Dim sPattern As String = "分署,等級,單位數,比率,,分署,等級,單位數,比率"
            'IMPLEVEL_1/IMP1 初審等級 'SUBTOTAL 初審分數
            'IMPLEVEL_1 初審等級 'SUBTOTAL 初審分數 'RL2_IMP1 複審/初審等級
            Const cst_colNM_RL2_IMP1 As String = "IMP1"
            Dim sColumn As String = "DISTNAME,IMP1,CNT3G,CNT4G,CNVL,DISTNAME,IMP1,CNT3W,CNT4W"

            Dim sPatternA() As String = Split(sPattern, ",")
            Dim sColumnA() As String = Split(sColumn, ",")
            Dim iColSpanCount As Integer = 9 ' sColumnA.Length
            'Dim i_ENDROWS As Integer=CInt(dtXls1.Rows.Count / 2)

            Dim strSTYLE As String = "<style> .text { mso-number-format:\@; text-align:center;} td { mso-number-format:\@;} </style>" & vbCrLf

            Dim sbHTML As New StringBuilder
            sbHTML.Append("<div>")
            sbHTML.Append("<table border='1' cellspacing='0' style='border-collapse:collapse;border:solid thin #000000;'>")

            '表頭及查詢條件列
            sbHTML.Append("<tr><td align='center' style='font-weight:bold' colspan='" & iColSpanCount & "'>" & t_SCORING_N2 & "</td></tr>")

            '建立資料面
            Dim iRows As Integer = 0
            For Each dr As DataRow In dtXls1.DefaultView.Table.Rows
                iRows += 1
                If Convert.ToString(dr(cst_colNM_RL2_IMP1)) = "A" Then
                    '建立表頭-空白1行
                    sbHTML.Append("<tr><td align='center' style='font-weight:bold' colspan='" & iColSpanCount & "'></td></tr>")
                    '表頭及查詢條件列--產業人才投資計畫	／ 提升勞工自主學習計畫			
                    sbHTML.Append("<tr><td align='center' style='font-weight:bold' colspan=4>" & s_ORGKIND2NM_G & "</td>")
                    sbHTML.Append("<td></td>")
                    sbHTML.Append("<td align='center' style='font-weight:bold' colspan=4>" & s_ORGKIND2NM_W & "</td></tr>")
                    '建立表頭1
                    sbHTML.Append("<tr>")
                    For i As Integer = 0 To sPatternA.Length - 1
                        sbHTML.Append(" <td align='center' style='font-weight:bold'>" & sPatternA(i) & "</td>") '& vbTab
                    Next
                    sbHTML.Append("</tr>")

                    '建立資料1
                    sbHTML.Append("<tr>")
                    For i As Integer = 0 To sColumnA.Length - 1
                        Select Case Convert.ToString(sColumnA(i))
                            Case "DISTNAME"
                                sbHTML.Append("<td align='center' rowspan=4>" & Convert.ToString(dr(sColumnA(i))) & "</td>") '& vbTab
                            Case Else
                                sbHTML.Append("<td align='center'>" & Convert.ToString(dr(sColumnA(i))) & "</td>") '& vbTab
                        End Select
                    Next
                    sbHTML.Append("</tr>")

                ElseIf Convert.ToString(dr(cst_colNM_RL2_IMP1)) = "D" Then
                    '建立資料1-尾1
                    sbHTML.Append("<tr>")
                    For i As Integer = 1 To sColumnA.Length - 1
                        Select Case Convert.ToString(sColumnA(i))
                            Case "DISTNAME"
                            Case Else
                                sbHTML.Append("<td align='center'>" & Convert.ToString(dr(sColumnA(i))) & "</td>") '& vbTab
                        End Select
                    Next
                    sbHTML.Append("</tr>")

                    '建立資料1-尾2
                    Dim s_CNT2G As String = Convert.ToString(dr("CNT2G"))
                    Dim s_CNT2W As String = Convert.ToString(dr("CNT2W"))
                    Dim s_DNMTOT As String = String.Concat(Convert.ToString(dr("DISTNAME")), " 合計 ")
                    Dim s_FootHtml2 As String = ""
                    s_FootHtml2 &= "<tr>"
                    s_FootHtml2 &= String.Format("<td align='center' colspan=2>{0}</td>", s_DNMTOT)
                    s_FootHtml2 &= String.Format("<td align='center'>{0}</td>", s_CNT2G)
                    s_FootHtml2 &= "<td></td>"
                    s_FootHtml2 &= "<td></td>"
                    s_FootHtml2 &= String.Format("<td align='center' colspan=2>{0}</td>", s_DNMTOT)
                    s_FootHtml2 &= String.Format("<td align='center'>{0}</td>", s_CNT2W)
                    s_FootHtml2 &= "<td></td>"
                    s_FootHtml2 &= "</tr>"
                    sbHTML.Append(s_FootHtml2)

                ElseIf Convert.ToString(dr(cst_colNM_RL2_IMP1)) <> "A" Then
                    '建立資料1
                    sbHTML.Append("<tr>")
                    For i As Integer = 1 To sColumnA.Length - 1
                        Select Case Convert.ToString(sColumnA(i))
                            Case "DISTNAME"
                            Case Else
                                sbHTML.Append("<td align='center'>" & Convert.ToString(dr(sColumnA(i))) & "</td>") '& vbTab
                        End Select
                    Next
                    sbHTML.Append("</tr>")
                End If
                'If (i_ENDROWS=iRows) Then Exit For
            Next

            sbHTML.Append("</table>")
            sbHTML.Append("</div>")

            'Dim strHTML As String=sbHTML.ToString()
            Dim parmsExp As New Hashtable From {
                {"ExpType", TIMS.GetListValue(RBListExpType)},
                {"FileName", strFileNM1},
                {"strSTYLE", strSTYLE},
                {"strHTML", sbHTML.ToString()},
                {"ResponseNoEnd", "Y"}
            }
            TIMS.Utl_ExportRp1(Me, parmsExp)
            TIMS.Utl_RespWriteEnd(Me, objconn, "")
        End Using
    End Sub

    Function Get_dtORGSCORING2_B(ByRef pms_B1 As Hashtable) As DataTable
        Dim vEXP As String = TIMS.GetMyValue2(pms_B1, "EXP")
        Dim vDISTID As String = TIMS.GetMyValue2(pms_B1, "DISTID") ' TIMS.GetListValue(ddlDISTID) '分署
        Dim vSCORINGID As String = TIMS.GetMyValue2(pms_B1, "SCORINGID") ' TIMS.GetListValue(ddlSCORING) '審查計分區間
        Dim vORGNAME As String = TIMS.GetMyValue2(pms_B1, "ORGNAME") 'TIMS.ClearSQM(OrgName.Text) '訓練機構
        Dim vCOMIDNO As String = TIMS.GetMyValue2(pms_B1, "COMIDNO") 'TIMS.ClearSQM(COMIDNO.Text) '機構別
        Dim vORGKIND2 As String = TIMS.GetMyValue2(pms_B1, "ORGKIND2") 'TIMS.GetListValue(OrgPlanKind) '計畫別(產投／提升勞工自主)
        Dim vORGKIND As String = TIMS.GetMyValue2(pms_B1, "ORGKIND") 'TIMS.GetListValue(OrgKindList) '機構別
        'Dim vFIRSTCHK_SCH As String = TIMS.GetMyValue2(pms_B1, "FIRSTCHK_SCH") 'TIMS.GetListValue(rblFIRSTCHK_SCH) '初審審核狀態

        If sm.UserInfo.LID <> 0 AndAlso vDISTID = "" Then Return Nothing
        If vSCORINGID = "" Then Return Nothing

        Dim sParms1 As New Hashtable
        Dim sSql As String = ""
        Select Case vEXP
            Case "Y4" '匯出統計表
#Region "EXP Y4"
                sParms1 = New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}
                sSql = ""
                'sSql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.DISTID ASC,ISNULL(a.SCORE4_1_2,a.SUBTOTAL) DESC,oo.ORGNAME ASC) ROWID" & vbCrLf
                sSql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.DISTID ASC,a.SUBTOTAL DESC,oo.ORGNAME ASC,a.IMPLEVEL_1) ROWID" & vbCrLf
                sSql &= " ,oo.COMIDNO,oo.ORGNAME" & vbCrLf
                sSql &= " ,oo.ORGKIND" & vbCrLf
                'sSql &= " ,(SELECT x.NAME FROM dbo.KEY_ORGTYPE x WHERE x.ORGTYPEID=oo.ORGKIND) ORGKIND_N" & vbCrLf
                sSql &= " ,oo.ORGKIND1" & vbCrLf
                'sSql &= " ,(SELECT x.ORGTYPE FROM dbo.VIEW_ORGTYPE1 x WHERE x.ORGTYPEID1=oo.ORGKIND1) ORGKIND1_N" & vbCrLf
                sSql &= " ,kd.DISTNAME3" & vbCrLf
                sSql &= " ,(SELECT x.MASTERNAME FROM V_ORGINFO x WHERE x.COMIDNO=oo.COMIDNO) MASTERNAME" & vbCrLf
                sSql &= " ,CONCAT(dbo.FN_CYEAR2(a.YEARS),'年',a.MONTHS,'月'" & vbCrLf
                sSql &= "   ,'(',dbo.FN_CYEAR2(a.YEARS1),'年',case when a.HALFYEAR1=1 then '上半年' else '下半年' end,'~'" & vbCrLf
                sSql &= "   ,dbo.FN_CYEAR2(a.YEARS2),'年',case when a.HALFYEAR2=1 then '上半年' else '下半年' end,')') SCORING_N" & vbCrLf
                sSql &= " ,CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2) SCORINGID" & vbCrLf
                '分署 加分項目 配合分署辦理相關活動或政策宣導(3%) 'SUBTOTAL 初審分數 'IMPLEVEL_1 初審等級
                sSql &= " ,a.SCORE4_1 ,a.SUBTOTAL ,ISNULL(a.SCORE4_1_2,a.SUBTOTAL) SCORE4_1_2" & vbCrLf
                'sSql &= " ,(a.SCORE1_1+a.SCORE1_2)+(a.SCORE2_1_1_ALL+a.SCORE2_1_2_SUM_ALL+a.SCORE2_1_3)+(a.SCORE2_2_1+a.SCORE2_2_2+a.SCORE2_3_1)" & vbCrLf
                'sSql &= " +(a.SCORE3_1+a.SCORE3_2)+isnull(a.SCORE4_1,0.0)+isnull(a.SCORE4_2,0.0) TOTALX" & vbCrLf
                'sSql &= " ,v1.VNAME SENDVER_N ,v2.VNAME RESULT_N" & vbCrLf
                sSql &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1" & vbCrLf
                sSql &= " ,a.FIRSTCHK,a.SECONDCHK,b.APPLIEDRESULT,'' bMEMO" & vbCrLf
                sSql &= " FROM dbo.ORG_SCORING2 a" & vbCrLf
                sSql &= " JOIN dbo.ORG_ORGINFO oo ON oo.OrgID=a.OrgID" & vbCrLf
                sSql &= " JOIN dbo.V_DISTRICT kd WITH(NOLOCK) ON kd.DISTID=a.DISTID" & vbCrLf
                sSql &= " LEFT JOIN dbo.ORG_TTQS2 b ON concat(b.ORGID,'x',b.COMIDNO,b.TPLANID,b.DISTID,b.YEARS,b.MONTHS)=concat(a.ORGID,'x',a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS,a.MONTHS)" & vbCrLf
                'sSql &= " LEFT JOIN dbo.V_SENDVER v1 ON v1.VID=b.SENDVER COLLATE Chinese_Taiwan_Stroke_CS_AS" & vbCrLf
                'sSql &= " LEFT JOIN dbo.V_RESULT v2 ON v2.VID=b.RESULT COLLATE Chinese_Taiwan_Stroke_CS_AS AND v2.VID<='4'" & vbCrLf
                '2024-01-2023-1-2023-2'--@SCORINGID
                sSql &= " WHERE CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf '審查計分區間
                sSql &= " AND a.TPLANID=@TPLANID" & vbCrLf
                '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
                sSql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
                sSql &= " AND a.IMPLEVEL_1 IS NOT NULL" & vbCrLf 'sSql &= " AND ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) IS NOT NULL" & vbCrLf
                If vDISTID <> "" Then '分署
                    sParms1.Add("DISTID", vDISTID)
                    sSql &= " AND a.DISTID=@DISTID" & vbCrLf
                End If
                If vORGNAME <> "" Then '訓練機構
                    sParms1.Add("ORGNAME", vORGNAME)
                    sSql &= " AND oo.ORGNAME like '%'+@ORGNAME+'%'" & vbCrLf
                End If
                If vCOMIDNO <> "" Then '統一編號
                    sParms1.Add("COMIDNO", vCOMIDNO)
                    sSql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
                End If
                Select Case vORGKIND2'計畫別
                    Case "G", "W"
                        sParms1.Add("ORGKIND2", vORGKIND2)
                        sSql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
                End Select
                If vORGKIND <> "" Then '機構別
                    sParms1.Add("ORGKIND", vORGKIND)
                    sSql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
                End If
                'Select Case vFIRSTCHK_SCH'初審審核狀態
                '    Case "Y", "N"
                '        sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
                '        sSql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
                'End Select
                'sSql &= " ORDER BY a.DISTID ASC,ISNULL(a.SCORE4_1_2,a.SUBTOTAL) DESC,oo.ORGNAME ASC" & vbCrLf
                sSql &= " ORDER BY a.DISTID ASC,a.SUBTOTAL DESC,oo.ORGNAME ASC,a.IMPLEVEL_1" & vbCrLf
#End Region

            Case "Y5" '匯出等級比率統計表
#Region "EXP Y5"
                sParms1 = New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}
                sSql = ""
                sSql &= " WITH WC1 AS ( SELECT oo.COMIDNO,oo.ORGNAME,oo.ORGKIND2" & vbCrLf
                sSql &= " ,oo.ORGKIND1,(SELECT x.ORGTYPE FROM dbo.VIEW_ORGTYPE1 x WHERE x.ORGTYPEID1=oo.ORGKIND1) ORGKIND1_N" & vbCrLf
                sSql &= " ,a.DISTID" & vbCrLf
                'SUBTOTAL 初審分數 'IMPLEVEL_1/IMP1 初審等級
                sSql &= " ,ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) RL2_IMP1,a.RLEVEL_2,a.IMPLEVEL_1 IMP1" & vbCrLf
                sSql &= " FROM dbo.ORG_SCORING2 a" & vbCrLf
                sSql &= " JOIN dbo.ORG_ORGINFO oo ON oo.OrgID=a.OrgID" & vbCrLf
                '2024-01-2023-1-2023-2'--@SCORINGID
                sSql &= " WHERE CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SCORINGID" & vbCrLf
                sSql &= " AND a.TPLANID=@TPLANID" & vbCrLf
                '審查計分表-調修1：審查計分表(初審)、(複審)查詢清單顯示邏輯調整
                sSql &= " AND dbo.FN_GET_CLASSCNT2B(a.COMIDNO,a.TPLANID,a.DISTID,a.YEARS1,a.YEARS2)>0" & vbCrLf
                sSql &= " AND a.IMPLEVEL_1 IS NOT NULL" & vbCrLf 'sSql &= " AND ISNULL(a.RLEVEL_2,a.IMPLEVEL_1) IS NOT NULL" & vbCrLf
                'sSql &= " --AND a.DISTID='001'" & vbCrLf

                If vDISTID <> "" Then '分署
                    sParms1.Add("DISTID", vDISTID)
                    sSql &= " AND a.DISTID=@DISTID" & vbCrLf
                End If
                If vORGNAME <> "" Then '訓練機構
                    sParms1.Add("ORGNAME", vORGNAME)
                    sSql &= " AND oo.ORGNAME like '%'+@ORGNAME+'%'" & vbCrLf
                End If
                If vCOMIDNO <> "" Then '統一編號
                    sParms1.Add("COMIDNO", vCOMIDNO)
                    sSql &= " AND oo.COMIDNO=@COMIDNO" & vbCrLf
                End If
                Select Case vORGKIND2'計畫別
                    Case "G", "W"
                        sParms1.Add("ORGKIND2", vORGKIND2)
                        sSql &= " AND oo.ORGKIND2=@ORGKIND2" & vbCrLf
                End Select
                If vORGKIND <> "" Then '機構別
                    sParms1.Add("ORGKIND", vORGKIND)
                    sSql &= " AND oo.ORGKIND=@ORGKIND" & vbCrLf
                End If
                'Select Case vFIRSTCHK_SCH'初審審核狀態
                '    Case "Y", "N"
                '        sParms1.Add("FIRSTCHK", vFIRSTCHK_SCH)
                '        sSql &= " AND a.FIRSTCHK=@FIRSTCHK" & vbCrLf
                'End Select
                sSql &= " )" & vbCrLf

                sSql &= " ,WC2 AS ( SELECT a.ORGKIND2,a.DISTID,COUNT(1) CNT2" & vbCrLf
                sSql &= " FROM WC1 a" & vbCrLf
                sSql &= " GROUP BY a.ORGKIND2,a.DISTID )" & vbCrLf

                sSql &= " ,WC3 AS ( SELECT a.ORGKIND2,a.DISTID,a.IMP1,COUNT(1) CNT3" & vbCrLf
                sSql &= " FROM WC1 a" & vbCrLf
                sSql &= " GROUP BY a.ORGKIND2,a.DISTID,a.IMP1 )" & vbCrLf

                sSql &= " ,WC4G AS ( SELECT c.ORGKIND2,c.DISTID,c.IMP1,c.CNT3,b.CNT2" & vbCrLf
                sSql &= " ,CONCAT(ROUND(convert(float,c.CNT3)/convert(float,b.CNT2)*100,2),'%') CNT4" & vbCrLf
                sSql &= " FROM WC3 c" & vbCrLf
                sSql &= " JOIN WC2 b on b.ORGKIND2=c.ORGKIND2 AND b.DISTID=c.DISTID AND c.ORGKIND2='G')" & vbCrLf

                sSql &= " ,WC4W AS ( SELECT c.ORGKIND2,c.DISTID,c.IMP1,c.CNT3,b.CNT2" & vbCrLf
                sSql &= " ,CONCAT(ROUND(convert(float,c.CNT3)/convert(float,b.CNT2)*100,2),'%') CNT4" & vbCrLf
                sSql &= " FROM WC3 c" & vbCrLf
                sSql &= " JOIN WC2 b on b.ORGKIND2=c.ORGKIND2 AND b.DISTID=c.DISTID AND c.ORGKIND2='W' )" & vbCrLf

                sSql &= " SELECT kd.DISTID,t.IMP1,kd.DISTNAME,kd.DISTNAME3,'-' CNVL" & vbCrLf
                sSql &= " ,c.CNT3 CNT3G,c.CNT2 CNT2G,c.CNT4 CNT4G" & vbCrLf
                sSql &= " ,w.CNT3 CNT3W,w.CNT2 CNT2W,w.CNT4 CNT4W" & vbCrLf
                sSql &= " FROM dbo.V_DISTRICT kd" & vbCrLf
                sSql &= " CROSS JOIN (VALUES ('A'),('B'),('C'),('D')) AS t(IMP1)" & vbCrLf
                sSql &= " LEFT JOIN WC4G c ON c.DISTID=kd.DISTID AND c.IMP1=t.IMP1" & vbCrLf
                sSql &= " LEFT JOIN WC4W w ON w.DISTID=kd.DISTID AND w.IMP1=t.IMP1" & vbCrLf
                sSql &= " WHERE kd.DISTID!='000'" & vbCrLf
                If vDISTID <> "" Then '分署
                    sParms1.Add("kDISTID", vDISTID)
                    sSql &= " AND kd.DISTID=@kDISTID" & vbCrLf
                End If
                sSql &= " ORDER BY kd.DISTID,t.IMP1" & vbCrLf
#End Region
        End Select

        'If TIMS.sUtl_ChkTest() Then TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(sParms1), vbCrLf, "--##CO_01_004, SQL: ", vbCrLf, sSql)) 

        Using dtXls1 As DataTable = DbAccess.GetDataTable(sSql, objconn, sParms1)
            Return dtXls1
        End Using
    End Function

    Function GET_ORGKIND2NM(vORGKIND2 As String) As String
        Dim rst As String = ""
        Dim pms_3 As New Hashtable From {{"ORGKIND2", vORGKIND2}}
        Dim sql_3 As String = " SELECT NAME ORGKIND2NM FROM dbo.V_ORGKIND1 WHERE [VALUE]=@ORGKIND2" & vbCrLf
        Dim dr3 As DataRow = DbAccess.GetOneRow(sql_3, objconn, pms_3)
        If dr3 Is Nothing Then Return rst
        rst = $"{dr3("ORGKIND2NM")}"
        Return rst
    End Function

    Function GET_SCORING_N2NM(vSCORINGID As String) As String
        Dim rst As String = ""
        Dim pms_1 As New Hashtable From {{"SCORINGID", vSCORINGID}}
        Dim sql_1 As String = " SELECT MAX(SCORING_N2) SCORING_N2 FROM dbo.VIEW_SCORING2 WHERE SCORINGID=@SCORINGID" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql_1, objconn, pms_1)
        If dr1 Is Nothing Then Return rst
        rst = $"{dr1("SCORING_N2")}"
        Return rst
    End Function

    Function GET_DISTNAME3(vDISTID As String) As String
        Dim rst As String = ""
        Dim pms_2 As New Hashtable From {{"DISTID", vDISTID}}
        Dim sql_2 As String = " SELECT DISTNAME3 FROM dbo.V_DISTRICT WHERE DISTID=@DISTID" & vbCrLf
        Dim dr2 As DataRow = DbAccess.GetOneRow(sql_2, objconn, pms_2)
        If dr2 Is Nothing Then Return rst
        rst = $"{dr2("DISTNAME3")}" 'If(dr2 Is Nothing, sm.UserInfo.DistID, )
        Return rst
    End Function

    Function GET_TPLAN_NAME(vTPLANID As String) As String
        Dim rst As String = ""
        Dim pms_3 As New Hashtable From {{"TPLANID", vTPLANID}}
        Dim sql_3 As String = " SELECT PLANNAME FROM KEY_PLAN WHERE TPLANID=@TPLANID" & vbCrLf
        Dim dr3 As DataRow = DbAccess.GetOneRow(sql_3, objconn, pms_3)
        If dr3 Is Nothing Then Return rst
        rst = $"{dr3("PLANNAME")}"
        Return rst
    End Function

    Private Function GET_SCORING_N2_TT(pms_R1 As Hashtable) As String
        Dim rst As String = ""
        'DIM vEXP As String=TIMS.GetMyValue2("",), vSCORINGID As String, vDISTID As String, vORGKIND2 As String
        Dim vEXP As String = TIMS.GetMyValue2(pms_R1, "EXP")
        Dim vSCORINGID As String = TIMS.GetMyValue2(pms_R1, "SCORINGID")
        Dim vDISTID As String = TIMS.GetMyValue2(pms_R1, "DISTID")
        Dim vORGKIND2 As String = TIMS.GetMyValue2(pms_R1, "ORGKIND2")

        If vEXP = "Y4" Then
            Dim s_SCORING_N2 As String = GET_SCORING_N2NM(vSCORINGID)
            Dim s_ORGKIND2NM As String = GET_ORGKIND2NM(vORGKIND2)
            Dim s_DISTNAME3 As String = GET_DISTNAME3(vDISTID) 'If(dr2 Is Nothing, sm.UserInfo.DistID, )
            Dim EXPORT_PRT_NM As String = "審查計分統計表"
            rst = String.Concat(s_SCORING_N2, s_ORGKIND2NM, EXPORT_PRT_NM, "(", s_DISTNAME3, ")")
            Return rst

        ElseIf vEXP = "Y5" Then
            Dim s_SCORING_N2 As String = GET_SCORING_N2NM(vSCORINGID)
            Dim s_TPLANNM As String = GET_TPLAN_NAME(sm.UserInfo.TPlanID)
            Dim s_DISTNAME3 As String = GET_DISTNAME3(vDISTID)
            Dim s_DISTNAME3_2 As String = If(s_DISTNAME3 = "", "", String.Concat("(", s_DISTNAME3, ")"))
            Dim EXPORT_PRT_NM As String = "等級單位數分配說明"
            rst = String.Concat(s_SCORING_N2, s_TPLANNM, EXPORT_PRT_NM, s_DISTNAME3_2)
            Return rst
        End If
        Return rst
    End Function

    ''' <summary>作業提醒,'查詢審查計分等級開關機制,審查計分表開關機制,Hid_REMIND1.Value</summary>
    ''' <returns></returns>
    Function CHK_TTQSQUERY_2() As Boolean
        Dim dtTQSQUERY_2 As DataTable = TIMS.Get_TTQSQUERY_TB(objconn, 2)
        If TIMS.dtHaveDATA(dtTQSQUERY_2) Then
            For Each drTQSQUERY_2 As DataRow In dtTQSQUERY_2.Rows
                Hid_REMIND1.Value = $"{drTQSQUERY_2("REMIND1")}"
            Next
        End If
        Return (Hid_REMIND1.Value <> "")
    End Function

#Region "NO USE"
    '匯出班級明細計分
    'Private Sub ExpXMLXLS3_1(ByRef dtXls As DataTable)
    '    'Dim sPattern As String=""
    '    'sPattern &= "計畫年度,分署,訓練單位,統一編號,單位屬性,課程申請流水號,班別名稱,期別,課程代碼,申請階段,訓練時數,開訓日期,結訓日期,課程分類"
    '    'sPattern &= ",5+2產業創新計畫,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策,核定人次,實際開訓人次,結訓人次,是否停辦"
    '    'sPattern &= ",招訓資料函送日期,招訓資料函送狀態,招訓資料逾期週數,招訓資料時效分數(2-1-1),招訓資料內容(2-1-2)"
    '    'sPattern &= ",開訓資料函送日期,開訓資料函送狀態,開訓資料逾期週數,開訓資料時效分數(2-1-1),開訓資料內容(2-1-2)"
    '    'sPattern &= ",結訓資料函送日期,結訓資料函送狀態,結訓資料逾期週數,結訓資料時效分數(2-1-1),結訓資料內容(2-1-2)"
    '    'sPattern &= ",變更項目,申請變更函送日期,申請變更函送狀態,逾期週數,變更資料時效分數(2-1-1),變更資料內容(2-1-2),不納入審查計分變更次數,政策性課程不扣分"
    '    'sPattern &= ",實地訪視日期,出席率,學員簽到(退)及教學日誌齊全,學員管理扣分(2-2-1),重要工作事項未依核定課程施訓,課程異常狀況,其他未依核定課程施訓,其他重大異常狀況,課程辦理情形扣分(2-2-2)"
    '    'sPattern &= ",會議應出席總場次,會議實際出席總場次,計畫參與度(%),計畫參與度得分(2-3)"
    '    'sPattern &= ",TTQS評核結果等級,TTQS評核結果得分(3-1)"
    '    'sPattern &= ",滿意學員人次,結訓學員總人次,學員滿意度(%),學員滿意程度得分(3-2)"
    '    'sPattern &= ",參訓學員訓後動態調查表填寫人次,訓後動態調查表填答率(%),填答率達80%得分"

    '    Dim sPattern As String=""
    '    sPattern &= "計畫年度,分署,訓練單位,統一編號,單位屬性,課程申請流水號,班別名稱,期別,課程代碼,申請階段,訓練時數,開訓日期,結訓日期,課程分類"
    '    sPattern &= ",5+2產業創新計畫,台灣AI行動計畫,數位國家創新經濟發展方案,國家資通安全發展方案,前瞻基礎建設計畫,新南向政策,核定人次,實際開訓人次,結訓人次,是否停辦"
    '    sPattern &= ",招訓資料函送日期,招訓資料函送狀態,招訓資料逾期週數"
    '    sPattern &= ",開訓資料函送日期,開訓資料函送狀態,開訓資料逾期週數"
    '    sPattern &= ",結訓資料函送日期,結訓資料函送狀態,結訓資料逾期週數"
    '    sPattern &= ",變更項目,申請變更函送日期,申請變更函送狀態,逾期週數,不納入審查計分變更次數,政策性課程不扣分"
    '    'sPattern &= ",實地訪視日期,出席率,學員簽到(退)及教學日誌齊全,學員管理扣分(2-2-1),重要工作事項未依核定課程施訓,課程異常狀況,其他未依核定課程施訓,其他重大異常狀況,課程辦理情形扣分(2-2-2)"
    '    'sPattern &= ",會議應出席總場次,會議實際出席總場次,計畫參與度(%),計畫參與度得分(2-3)"
    '    'sPattern &= ",TTQS評核結果等級,TTQS評核結果得分(3-1)"
    '    'sPattern &= ",滿意學員人次,結訓學員總人次,學員滿意度(%),學員滿意程度得分(3-2)"
    '    'sPattern &= ",參訓學員訓後動態調查表填寫人次,訓後動態調查表填答率(%),填答率達80%得分"
    '    Dim sColumn As String=""
    '    sColumn &= "YEARS,DISTNAME,ORGNAME,COMIDNO,ORGTYPENAME,PSNO28,CLASSCNAME,CYCLTYPE,OCID,APPSTAGE_N,THOURS,STDATE,FTDATE,JOBNAME"
    '    sColumn &= ",D20KNAME1,D20KNAME2,D20KNAME3,D20KNAME4,D20KNAME5,D20KNAME6,TNUM,STDACTCNT,STDCLOSECNT,NOTOPEN_N"
    '    sColumn &= ",SENDDATE1,STATUS1,OVERWEEK1"
    '    sColumn &= ",SENDDATE2,STATUS2,OVERWEEK2"
    '    sColumn &= ",SENDDATE3,STATUS3,OVERWEEK3"
    '    sColumn &= ",ALTDATAID,SENDDATE4,STATUS4,OVERWEEK4,NOINC4,NODEDUC4"
    '    'sColumn &= ",實地訪視日期,出席率,學員簽到(退)及教學日誌齊全,學員管理扣分(2-2-1),重要工作事項未依核定課程施訓,課程異常狀況,其他未依核定課程施訓,其他重大異常狀況,課程辦理情形扣分(2-2-2)"
    '    'sColumn &= ",會議應出席總場次,會議實際出席總場次,計畫參與度(%),計畫參與度得分(2-3)"
    '    'sColumn &= ",TTQS評核結果等級,TTQS評核結果得分(3-1)"
    '    'sColumn &= ",滿意學員人次,結訓學員總人次,學員滿意度(%),學員滿意程度得分(3-2)"
    '    'sColumn &= ",參訓學員訓後動態調查表填寫人次,訓後動態調查表填答率(%),填答率達80%得分"

    '    Dim sPatternA() As String=Split(sPattern, ",")
    '    Dim sColumnA() As String=Split(sColumn, ",")

    '    Dim sFileName1 As String=String.Concat("匯出班級明細計分", TIMS.GetDateNo())

    '    '套CSS值
    '    Dim strSTYLE As String=String.Concat("<style>", "td{mso-number-format:""\@"";}", ".noDecFormat{mso-number-format:""0"";}", "</style>")

    '    Dim sbHTML As New StringBuilder
    '    sbHTML.Append("<div>")
    '    sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

    '    '建立輸出文字
    '    Dim ExportStr As String=""
    '    '標題抬頭1
    '    ExportStr=String.Format("<td colspan={0}>{1}</td>", sPatternA.Length, sFileName1) '& vbTab
    '    sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

    '    '標題抬頭2
    '    ExportStr=""
    '    For i As Integer=0 To sPatternA.Length - 1
    '        ExportStr &= String.Format("<td>{0}</td>", sPatternA(i)) '& vbTab
    '    Next
    '    sbHTML.Append(String.Concat("<tr>", ExportStr, "</tr>"))

    '    '建立資料面
    '    Dim iRows As Integer=0
    '    For Each dr As DataRow In dtXls.Rows
    '        iRows += 1
    '        ExportStr="<tr>"
    '        For i As Integer=0 To sColumnA.Length - 1
    '            Dim sCOLTXT As String=""
    '            Select Case sColumnA(i)
    '                Case "SEQNO"
    '                    sCOLTXT=iRows.ToString()
    '                Case Else
    '                    sCOLTXT=Convert.ToString(dr(sColumnA(i)))
    '            End Select
    '            ExportStr &= String.Format("<td>{0}</td>", sCOLTXT)
    '        Next
    '        ExportStr &= "</tr>" & vbCrLf
    '        sbHTML.Append(ExportStr)
    '    Next
    '    sbHTML.Append("</table>")
    '    sbHTML.Append("</div>")

    '    'Call ExpClass1.Utl_Export1_XLSX(dtG, s_FNAMEIN, cst_sheetN1, cst_titleRange)

    '    Dim parmsExp As New Hashtable
    '    parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType)) 'EXCEL/PDF/ODS
    '    'parmsExp.Add("ExpType", "EXCEL") 'EXCEL/PDF/ODS
    '    parmsExp.Add("FileName", sFileName1)
    '    parmsExp.Add("strSTYLE", strSTYLE)
    '    parmsExp.Add("strHTML", sbHTML.ToString())
    '    parmsExp.Add("ResponseNoEnd", "Y")
    '    TIMS.Utl_ExportRp1(Me, parmsExp)

    '    TIMS.CloseDbConn(objconn)
    '    TIMS.Utl_RespWriteEnd(Me, objconn, "")
    'End Sub
#End Region

End Class

