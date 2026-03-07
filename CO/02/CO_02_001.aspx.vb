Imports System.IO
Imports System.Linq
Imports System.Text.RegularExpressions
Imports NPOI.SS.UserModel
Imports NPOI.XSSF.UserModel
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CO_02_001
    Inherits AuthBasePage 'System.Web.UI.Page

    'Dim fontName As String = "標楷體"
    'Dim fontSize12s As Single = 12.0F
    'Dim fontSize14s As Single = 14.0F
    Const cst_fontSize16s As Single = 16.0F
    Dim print_lock As New Object '(); //lock

    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If Not IsPostBack Then
            Call CCREATE1() '重新載入資訊
        End If

    End Sub

    Sub CCREATE1()
        ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, Nothing, objconn)
        If (ddlDISTID_SCH.Items.FindByValue("000") IsNot Nothing) Then ddlDISTID_SCH.Items.Remove(ddlDISTID_SCH.Items.FindByValue("000"))
        Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID)
        ddlSCORING_SCH = TIMS.Get_ddlSCORING(ddlSCORING_SCH, objconn)
        Dim V_RBL_ORGPLANKIND_SCH As String = TIMS.GetListValue(RBL_ORGPLANKIND_SCH)
        TIMS.GET_DDL_TYPEID12(DDL_TYPEID2_SCH, objconn, V_RBL_ORGPLANKIND_SCH)
        Call GET_EXITCELL(CBEXIT_SCH, CBEXIT2_SCH) '匯出欄位 
        Common.SetListItem(CBEXIT_SCH, "0")
        If TIMS.sUtl_ChkTest() Then
            Common.SetListItem(ddlDISTID_SCH, "004")
        End If
    End Sub

    Sub GET_EXITCELL(ByRef obj As ListControl, ByRef obj2 As ListControl)
        obj.Items.Clear()
        Dim listStr1 As String = "全部,計畫,分署,訓練單位名稱,理事長(負責人),單位屬性"
        Dim aStr1 As String() = Split(listStr1, ",")
        With obj.Items
            For i As Integer = 0 To aStr1.Length - 1
                .Insert(i, New ListItem(aStr1(i), i))
            Next
        End With

        obj2.Items.Clear()
        Dim listStr2 As String = ""
        listStr2 &= "前2個半年(等級/分數/申請班數/申請補助費/核定班數/核定補助費)"
        listStr2 &= ",當年度階段(分署加分前分數/分署加分前等級/分署加分/加分後分數/加分後等級)"
        Dim aStr2 As String() = Split(listStr2, ",")
        With obj2.Items
            For i As Integer = 0 To aStr2.Length - 1
                .Insert(i, New ListItem(aStr2(i), i))
            Next
        End With
    End Sub

    Protected Sub RBL_ORGPLANKIND_SCH_SelectedIndexChanged(sender As Object, e As EventArgs) Handles RBL_ORGPLANKIND_SCH.SelectedIndexChanged
        Dim v_RBL_ORGPLANKIND_SCH As String = TIMS.GetListValue(RBL_ORGPLANKIND_SCH)
        TIMS.GET_DDL_TYPEID12(DDL_TYPEID2_SCH, objconn, v_RBL_ORGPLANKIND_SCH)
    End Sub

    Protected Sub BTN_EXP1_Click(sender As Object, e As EventArgs) Handles BTN_EXP1.Click
        ExportXlsStd28_5()
    End Sub

    Function GET_SCORING2_R(oVAL As String) As DataRow
        Dim PMS1 As New Hashtable From {{"VALUEFD", oVAL}}
        Dim SSQL As String = "SELECT a.TEXTFD,a.VALUEFD,a.ROC_YEARS,a.MONTHS_N,a.NEXT_YMN,a.YEARS,a.MONTHS FROM V_SCORING2 a WHERE a.VALUEFD=@VALUEFD" & vbCrLf
        Return DbAccess.GetOneRow(SSQL, objconn, PMS1)
    End Function

    Function GET_SCORING2_dt2_3(oVAL As String) As DataTable
        Dim PMS1 As New Hashtable From {{"VALUEFD", oVAL}}
        Dim SSQL As String = "
SELECT TOP 3 TEXTFD,VALUEFD,ROC_YEARS,MONTHS_N,NEXT_YMN,NEXT_YMN2,YEARS,MONTHS
FROM V_SCORING2 WHERE VALUEFD<=@VALUEFD
ORDER BY YEARS DESC,MONTHS DESC
"
        Dim dt6 As DataTable = DbAccess.GetDataTable(SSQL, objconn, PMS1)
        Return dt6
    End Function

    Function GET_ORGTYPEID1(VTypeID1 As String, VTypeID2 As String) As String
        Dim CUTypeID1 As String() = {"1", "2", "G", "W"}
        If CUTypeID1.Contains(VTypeID1) Then
            Dim VVTypeID1 As String = If(VTypeID1 = "G", "1", If(VTypeID1 = "W", "2", VTypeID1))
            Dim SQL1 As String = "SELECT ORGTYPEID1 FROM VIEW_ORGTYPE1 WHERE TypeID1=@TypeID1 and TypeID2=@TypeID2"
            Dim SCMD1 As New SqlCommand(SQL1, objconn)
            With SCMD1
                .Parameters.Add("TypeID1", VVTypeID1)
                .Parameters.Add("TypeID2", VTypeID2)
            End With
            Dim DR As DataRow = TIMS.GetOneRow(SCMD1, objconn)
            If DR Is Nothing Then Return ""
            Return $"{DR("ORGTYPEID1")}"
        End If
        Return ""
    End Function
    Function SEARCH_DATA1_dt3(dt2_3 As DataTable, vSCORINGID As String) As DataTable
        'ddlDISTID_SCH,ddlSCORING_SCH,ORGNAME_SCH,COMIDNO_SCH,TRPlanPoint28,RBL_ORGPLANKIND_SCH,DDL_TYPEID2_SCH,RBL_CrossDist_SCH,MASTERNAME_SCH,
        Dim vDISTID As String = TIMS.GetListValue(ddlDISTID_SCH)
        'Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING_SCH)
        Dim vORGNAME As String = TIMS.ClearSQM(ORGNAME_SCH.Text)
        Dim vCOMIDNO As String = TIMS.ClearSQM(COMIDNO_SCH.Text)
        Dim vORGKIND2 As String = TIMS.GetListValue(RBL_ORGPLANKIND_SCH)
        Dim vTYPEID2 As String = TIMS.GetListValue(DDL_TYPEID2_SCH)
        Dim V_ORGTYPEID1 As String = If(vTYPEID2 <> "", GET_ORGTYPEID1(vORGKIND2, vTYPEID2), "")
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Dim v_RBL_CrossDist_SCH As String = TIMS.GetListValue(RBL_CrossDist_SCH)
        Dim vMASTERNAME As String = TIMS.ClearSQM(MASTERNAME_SCH.Text)
        'Dim vSCORINGID As String = TIMS.GetMyValue2(PMSR, "SCORINGID") 'TIMS.GetListValue(ddlSCORING)
        ', {"SCORESTAGE", vSCORESTAGE}01： 部-加分前：等級顯示【初擬等級】,02:部-加分後：等級顯示【部加分等級】,03:署-加分後：等級顯示【複審等級】
        'Dim vSCORESTAGE As String = 2 'TIMS.GetListValue(rblSCORESTAGE)
        'Dim PMS1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}, {"ORGKIND2", "W"}}
        Dim PMS1 As New Hashtable From {{"SCORINGID", vSCORINGID}, {"TPLANID", sm.UserInfo.TPlanID}}
        Dim SC_WJ As String = If(dt2_3.Rows.Count > 1, Convert.ToString(dt2_3.Rows(1)("VALUEFD")), "")
        Dim SC_WI As String = If(dt2_3.Rows.Count > 2, Convert.ToString(dt2_3.Rows(2)("VALUEFD")), "")
        PMS1.Add("SC_WJ", SC_WJ)
        PMS1.Add("SC_WI", SC_WI)
        PMS1.Add("ORGKIND2", vORGKIND2)

        'DECLARE @SCORINGID VARCHAR(111)='2025-07-2024-2-2025-1';'DECLARE @SC_WJ VARCHAR(111)='2025-01-2024-1-2024-2';'DECLARE @SC_WI VARCHAR(111)='2025-01-2024-1-2024-2';
        Dim SQL_VNG As String = "WITH WVN1 AS (SELECT TPLANID,DISTID,YEARS,MONTHS,SUBTOTALA,SUBTOTALB,SUBTOTALC,SUBTOTALD FROM V_SCORING2_MIN_G)"
        Dim SQL_VNW As String = "WITH WVN1 AS (SELECT TPLANID,DISTID,YEARS,MONTHS,SUBTOTALA,SUBTOTALB,SUBTOTALC,SUBTOTALD FROM V_SCORING2_MIN)"
        Dim SQL_VN As String = If(vORGKIND2 = "G", SQL_VNG, SQL_VNW)

        Dim SQL_1 As String = $"{SQL_VN}
,WJ AS (SELECT OSID2,CLSAPPCNT,CLSACTCNT,SCORE4_1_2,RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID
,dbo.FN_GET_TOTALCOST(COMIDNO,OSID2,1) TCOST1,dbo.FN_GET_TOTALCOST(COMIDNO,OSID2,2) TCOST2 FROM dbo.ORG_SCORING2 a 
	WHERE CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SC_WJ)
,WI AS (SELECT OSID2,CLSAPPCNT,CLSACTCNT,SCORE4_1_2,RLEVEL_2,BRANCHPNT,MINISTERADD,DEPTADD,ORGID,DISTID,TPLANID
,dbo.FN_GET_TOTALCOST(COMIDNO,OSID2,1) TCOST1,dbo.FN_GET_TOTALCOST(COMIDNO,OSID2,2) TCOST2 FROM dbo.ORG_SCORING2 a 
	WHERE CONCAT(a.YEARS,'-',a.MONTHS,'-',a.YEARS1,'-',a.HALFYEAR1,'-',a.YEARS2,'-',a.HALFYEAR2)=@SC_WI)
,WM1 AS (
SELECT A.OSID2,A.YEARS,A.COMIDNO,A.ORGID,A.DISTID,A.TPLANID
,OO.ORGKIND2,OO.ORGNAME,DD.DISTNAME3
,OO.ORGKIND1,KO.ORGTYPE ORGKIND1_N
,(SELECT OP.MASTERNAME FROM VIEW_ORGPLANINFO OP WHERE OP.ORGID=A.ORGID AND OP.TPLANID=A.TPLANID AND OP.DISTID=A.DISTID AND OP.YEARS=A.YEARS) MASTERNAME
,(SELECT X.MASTERNAME FROM V_ORGINFO X WHERE X.ORGID=A.ORGID AND X.DISTID=A.DISTID) MASTERNAME1
,A.SUBTOTAL, A.IMPLEVEL_1, A.SCORE4_1_2, A.RLEVEL_2
,A.SCORE4_1,A.MINISTERADD,A.DEPTADD,A.BRANCHPNT 
,WJ.SCORE4_1_2 WJSCORE4_1_2, WJ.RLEVEL_2 WJRLEVEL_2,WJ.CLSAPPCNT WJCLSAPPCNT,WJ.CLSACTCNT WJCLSACTCNT
,WI.SCORE4_1_2 WISCORE4_1_2, WI.RLEVEL_2 WIRLEVEL_2,WI.CLSAPPCNT WICLSAPPCNT,WI.CLSACTCNT WICLSACTCNT
,WJ.TCOST1 WJTCOST1,WJ.TCOST2 WJTCOST2
,WI.TCOST1 WITCOST1,WI.TCOST2 WITCOST2
,CONCAT(A.YEARS,'-',A.MONTHS,'-',A.YEARS1,'-',A.HALFYEAR1,'-',A.YEARS2,'-',A.HALFYEAR2) SCORINGID,A.HALFYEAR1,A.HALFYEAR2
,dbo.FN_GET_CROSSDIST(A.YEARS,A.COMIDNO,A.HALFYEAR1) I_CROSSDIST
,VN.SUBTOTALA,VN.SUBTOTALB,VN.SUBTOTALC,VN.SUBTOTALD
FROM ORG_SCORING2 A
JOIN ORG_ORGINFO OO ON OO.COMIDNO=A.COMIDNO
JOIN V_DISTRICT DD ON DD.DISTID=A.DISTID
LEFT JOIN VIEW_ORGTYPE1 KO ON KO.ORGTYPEID1=OO.ORGKIND1
LEFT JOIN WJ ON WJ.ORGID=a.ORGID AND WJ.DISTID=a.DISTID AND WJ.TPLANID=a.TPLANID
LEFT JOIN WI ON WI.ORGID=a.ORGID AND WI.DISTID=a.DISTID AND WI.TPLANID=a.TPLANID
LEFT JOIN WVN1 VN ON VN.TPLANID=A.TPLANID AND VN.DISTID=A.DISTID AND VN.YEARS=A.YEARS AND VN.MONTHS=A.MONTHS
"
        SQL_1 &= " WHERE A.TPLANID=@TPLANID AND CONCAT(A.YEARS,'-',A.MONTHS,'-',A.YEARS1,'-',A.HALFYEAR1,'-',A.YEARS2,'-',A.HALFYEAR2)=@SCORINGID" & vbCrLf
        SQL_1 &= " AND OO.ORGKIND2=@ORGKIND2" & vbCrLf
        If (vDISTID <> "") Then
            PMS1.Add("DISTID", vDISTID)
            SQL_1 &= " AND A.DISTID=@DISTID" & vbCrLf
        End If
        If (vORGNAME <> "") Then
            PMS1.Add("ORGNAME", vORGNAME)
            SQL_1 &= " AND OO.ORGNAME LIKE '%'+@ORGNAME+'%'" & vbCrLf
        End If
        If (vCOMIDNO <> "") Then
            PMS1.Add("COMIDNO", vCOMIDNO)
            SQL_1 &= " AND A.COMIDNO=@COMIDNO" & vbCrLf
        End If
        'If (vTYPEID2 <> "") Then PMS1.Add("TYPEID2", vTYPEID2)
        If (V_ORGTYPEID1 <> "") Then
            PMS1.Add("ORGTYPEID1", V_ORGTYPEID1)
            SQL_1 &= " AND OO.ORGKIND1=@ORGTYPEID1" & vbCrLf
        End If
        If (vMASTERNAME <> "") Then
            PMS1.Add("MASTERNAME", vMASTERNAME)
            SQL_1 &= " AND EXISTS (SELECT 1 FROM VIEW_ORGPLANINFO X WHERE X.ORGID=A.ORGID AND X.TPLANID=A.TPLANID AND X.DISTID=A.DISTID AND X.YEARS=A.YEARS AND X.MASTERNAME LIKE '%'+@MASTERNAME+'%')" & vbCrLf
        End If
        ',WM1 AS (
        SQL_1 &= " )
SELECT * FROM WM1 M"
        '跨區/轄區提案 'D>不區分 C>跨區提案單位 J>轄區提案單位
        Select Case v_RBL_CrossDist_SCH
            Case "C" 'C:跨區提案單位
                SQL_1 &= " WHERE M.I_CROSSDIST!=-1" & vbCrLf
            Case "J" 'J:轄區提案單位
                SQL_1 &= " WHERE M.I_CROSSDIST=-1" & vbCrLf
        End Select
        SQL_1 &= " ORDER BY M.DISTID,M.SUBTOTAL DESC,M.ORGNAME"
        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, $"--PMS1:{vbCrLf}{TIMS.GetMyValue5(PMS1)}{vbCrLf}--CO_02_001:{vbCrLf}{SQL_1}{vbCrLf}")
        End If
        Dim dt As DataTable = DbAccess.GetDataTable(SQL_1, objconn, PMS1)
        Return dt
    End Function

    Function GET_ORGKIND2_N(STR1 As String) As String
        If (STR1 = "") Then Return STR1
        Return If(STR1 = "G", "產投", If(STR1 = "W", "自主", STR1))
    End Function

    Sub ExportXlsStd28_5()
        Const Cst_FileSavePath As String = "~/CO/01/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)
        Const cst_SampleXLS As String = "~\CO\02\sp2_CO02001.xlsx" '& cst_files_ext 'copy一份sample資料---Start
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

        Dim vSCORINGID As String = TIMS.GetListValue(ddlSCORING_SCH)
        Dim drSC2 As DataRow = GET_SCORING2_R(vSCORINGID)
        If drSC2 Is Nothing Then
            Common.MessageBox(Me, "查無 匯出資料!")
            Exit Sub
        End If
        Dim dt2_3 As DataTable = GET_SCORING2_dt2_3(vSCORINGID)
        If TIMS.dtNODATA(dt2_3) Then
            Common.MessageBox(Me, "查無 匯出資料!!")
            Exit Sub
        End If
        Dim drSC2_WJ As DataRow = If(dt2_3.Rows.Count > 1, dt2_3.Rows(1), Nothing)
        Dim drSC2_WI As DataRow = If(dt2_3.Rows.Count > 2, dt2_3.Rows(2), Nothing)
        Dim dtXls1 As DataTable = SEARCH_DATA1_dt3(dt2_3, vSCORINGID)
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim ROC_YEARSJ As String = $"{drSC2_WJ("ROC_YEARS")}"
        Dim MONTHS_NJ As String = $"{drSC2_WJ("MONTHS_N")}"
        Dim ROC_YEARSI As String = $"{drSC2_WI("ROC_YEARS")}"
        Dim MONTHS_NI As String = $"{drSC2_WI("MONTHS_N")}"

        Dim ROC_YEARS As String = $"{drSC2("ROC_YEARS")}"
        Dim MONTHS_N As String = $"{drSC2("MONTHS_N")}"
        Dim TITLE_NM1 As String = $"{ROC_YEARS}年度{MONTHS_N}產業人才投資方案-審查計分綜合動態報表"
        Dim SheetNM As String = "審查計分綜合動態報表"
        Dim vROC_YMD_NOW As String = TIMS.GetROCTWDate(Now)
        Dim s_FILENAME1 As String = $"{TITLE_NM1}x{TIMS.GetDateNo2(3)}"

        Dim CG23 As String = $"{ROC_YEARSI}{MONTHS_NI(0)}"
        Dim CM23 As String = $"{ROC_YEARSJ}{MONTHS_NJ(0)}"

        Dim CS23 As String = $"{ROC_YEARS}{MONTHS_N(0)}分署加分前分數"
        Dim CT23 As String = $"{ROC_YEARS}{MONTHS_N(0)}分署加分前等級"
        Dim CU23 As String = $"{ROC_YEARS}{MONTHS_N(0)}分署加分"
        Dim CV23 As String = $"{ROC_YEARS}{MONTHS_N(0)}分署加分後分數"
        Dim CW23 As String = $"{ROC_YEARS}{MONTHS_N(0)}分署加分後等級"

        Dim fg_RespWriteEnd As Boolean = False
        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial, ExcelPackage.LicenseContext = LicenseContext.NonCommercial
            'Dim file1 As New FileInfo(filePath1) 'Dim ndt As DateTime = Now
            '開檔
            Using fs1 As FileStream = New FileStream(sMyFile1, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                Dim ep As New ExcelPackage(fs1)

                Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)
                ws.Name = SheetNM
                'Dim ws1 As ExcelWorksheet = ep.Workbook.Worksheets(1)
                'Dim ep As New ExcelPackage() 'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
                Dim END_COL_NM As String = "W"
                Dim S_COLNM1 As String = $"A1:{END_COL_NM}1"
                Dim cellsCOLSPNumF2 As String = String.Concat($"A5:{END_COL_NM}", "{0}") '(畫格子使用)

                ws.Cells(S_COLNM1).Value = TITLE_NM1
                ws.Cells(S_COLNM1).Style.Font.Bold = True
                ws.Cells(S_COLNM1).Style.Font.Size = cst_fontSize16s

                TIMS.SetCellValue2(ws, "G2", CG23)
                TIMS.SetCellValue2(ws, "M2", CM23)

                TIMS.SetCellValue2(ws, "S2", CS23)
                TIMS.SetCellValue2(ws, "T2", CT23)
                TIMS.SetCellValue2(ws, "U2", CU23)
                TIMS.SetCellValue2(ws, "V2", CV23)
                TIMS.SetCellValue2(ws, "W2", CW23)

                Dim idxStr1 As Integer = 5
                Dim iROWNUM As Integer = 0
                For Each dr As DataRow In dtXls1.Rows
                    iROWNUM += 1
                    TIMS.SetCellValue3(ws, $"A{idxStr1}", iROWNUM)
                    TIMS.SetCellValue3(ws, $"B{idxStr1}", GET_ORGKIND2_N($"{dr("ORGKIND2")}"))
                    TIMS.SetCellValue3(ws, $"C{idxStr1}", $"{dr("DISTNAME3")}")
                    TIMS.SetCellValue3(ws, $"D{idxStr1}", $"{dr("ORGNAME")}")
                    Dim V_MASTERNAME As String = If($"{dr("MASTERNAME")}" <> "", $"{dr("MASTERNAME")}", $"{dr("MASTERNAME1")}")
                    TIMS.SetCellValue3(ws, $"E{idxStr1}", $"{dr("MASTERNAME")}")
                    TIMS.SetCellValue3(ws, $"F{idxStr1}", $"{dr("ORGKIND1_N")}")

                    TIMS.SetCellValue3(ws, $"G{idxStr1}", $"{dr("WIRLEVEL_2")}")
                    TIMS.SetCellValue3(ws, $"H{idxStr1}", TIMS.VAL1($"{dr("WISCORE4_1_2")}"))
                    TIMS.SetCellValue3(ws, $"I{idxStr1}", TIMS.VAL1($"{dr("WICLSAPPCNT")}"))
                    TIMS.SetCellValue3(ws, $"J{idxStr1}", TIMS.VAL1($"{dr("WITCOST1")}"))
                    TIMS.SetCellValue3(ws, $"K{idxStr1}", TIMS.VAL1($"{dr("WICLSACTCNT")}"))
                    TIMS.SetCellValue3(ws, $"L{idxStr1}", TIMS.VAL1($"{dr("WITCOST2")}"))

                    TIMS.SetCellValue3(ws, $"M{idxStr1}", $"{dr("WJRLEVEL_2")}")
                    TIMS.SetCellValue3(ws, $"N{idxStr1}", TIMS.VAL1($"{dr("WJSCORE4_1_2")}"))
                    TIMS.SetCellValue3(ws, $"O{idxStr1}", TIMS.VAL1($"{dr("WJCLSAPPCNT")}"))
                    TIMS.SetCellValue3(ws, $"P{idxStr1}", TIMS.VAL1($"{dr("WJTCOST1")}"))
                    TIMS.SetCellValue3(ws, $"Q{idxStr1}", TIMS.VAL1($"{dr("WJCLSACTCNT")}"))
                    TIMS.SetCellValue3(ws, $"R{idxStr1}", TIMS.VAL1($"{dr("WJTCOST2")}"))

                    Dim V_SUBTOTAL As Double = TIMS.VAL1($"{dr("SUBTOTAL")}")
                    Dim V_SCORE4_1 As Double = TIMS.VAL1($"{dr("SCORE4_1")}")
                    Dim V_SUBTOTAL_BEF As Double = V_SUBTOTAL - V_SCORE4_1
                    Dim V_IMPLEVEL_1 As String = $"{dr("IMPLEVEL_1")}"
                    Dim V_SUBTOTALA As Double = TIMS.VAL1($"{dr("SUBTOTALA")}")
                    Dim V_SUBTOTALB As Double = TIMS.VAL1($"{dr("SUBTOTALB")}")
                    Dim V_SUBTOTALC As Double = TIMS.VAL1($"{dr("SUBTOTALC")}")
                    Dim V_SUBTOTALD As Double = TIMS.VAL1($"{dr("SUBTOTALD")}")
                    Dim V_LEVEL_N As String = "D"
                    If (V_SUBTOTAL_BEF = V_SUBTOTAL) Then
                        V_LEVEL_N = V_IMPLEVEL_1
                    ElseIf (V_SUBTOTAL_BEF >= V_SUBTOTALA) Then
                        V_LEVEL_N = "A"
                    ElseIf (V_SUBTOTAL_BEF >= V_SUBTOTALB) Then
                        V_LEVEL_N = "B"
                    ElseIf (V_SUBTOTAL_BEF >= V_SUBTOTALC) Then
                        V_LEVEL_N = "C"
                    ElseIf (V_SUBTOTAL_BEF >= V_SUBTOTALD) Then
                        V_LEVEL_N = "D"
                    End If
                    TIMS.SetCellValue3(ws, $"S{idxStr1}", V_SUBTOTAL_BEF)
                    TIMS.SetCellValue3(ws, $"T{idxStr1}", V_LEVEL_N)
                    TIMS.SetCellValue3(ws, $"U{idxStr1}", V_SCORE4_1)
                    TIMS.SetCellValue3(ws, $"V{idxStr1}", V_SUBTOTAL)
                    TIMS.SetCellValue3(ws, $"W{idxStr1}", V_IMPLEVEL_1)

                    'TIMS.SetCellValue3(ws, $"S{idxStr1}", TIMS.VAL1($"{dr("SUBTOTAL")}"))
                    'TIMS.SetCellValue3(ws, $"T{idxStr1}", $"{dr("IMPLEVEL_1")}")
                    'TIMS.SetCellValue3(ws, $"U{idxStr1}", TIMS.VAL1($"{dr("BRANCHPNT")}"))
                    'TIMS.SetCellValue3(ws, $"V{idxStr1}", TIMS.VAL1($"{dr("SCORE4_1_2")}"))
                    'TIMS.SetCellValue3(ws, $"W{idxStr1}", $"{dr("RLEVEL_2")}")

                    ws.Cells("G" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells("M" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells("T" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    ws.Cells("W" & idxStr1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center

                    ws.Cells("J" & idxStr1).Style.Numberformat.Format = "#,##0"
                    ws.Cells("L" & idxStr1).Style.Numberformat.Format = "#,##0"
                    ws.Cells("P" & idxStr1).Style.Numberformat.Format = "#,##0"
                    ws.Cells("R" & idxStr1).Style.Numberformat.Format = "#,##0"

                    ws.Cells("H" & idxStr1).Style.Numberformat.Format = "#0.0"
                    ws.Cells("N" & idxStr1).Style.Numberformat.Format = "#0.0"
                    ws.Cells("S" & idxStr1).Style.Numberformat.Format = "#0.0"
                    ws.Cells("U" & idxStr1).Style.Numberformat.Format = "#0.0"
                    ws.Cells("V" & idxStr1).Style.Numberformat.Format = "#0.0"

                    idxStr1 += 1
                Next

                ws.Column(ws.Cells($"D{idxStr1}").Start.Column).Width = Convert.ToDouble(55)
                ws.Column(ws.Cells($"J{idxStr1}").Start.Column).Width = Convert.ToDouble(11)
                ws.Column(ws.Cells($"L{idxStr1}").Start.Column).Width = Convert.ToDouble(11)
                ws.Column(ws.Cells($"P{idxStr1}").Start.Column).Width = Convert.ToDouble(11)
                ws.Column(ws.Cells($"R{idxStr1}").Start.Column).Width = Convert.ToDouble(11)

                idxStr1 -= 1 '(畫線)
                Using exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr1))
                    exlRow3X.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                    TIMS.SetCellBorder(exlRow3X)
                End Using

                ws.View.ZoomScale = 90

                Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
                Select Case V_ExpType
                    Case "EXCEL"
                        TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case "ODS"
                        TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath, s_FILENAME1)
                        fg_RespWriteEnd = True
                    Case Else
                        Dim s_log1 As String = $"ExpType(參數有誤)!!{V_ExpType}"
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
End Class
