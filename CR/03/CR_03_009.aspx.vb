Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CR_03_009
    Inherits AuthBasePage 'System.Web.UI.Page

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
    End Sub

    Sub CCreate1()
        PanelSch1.Visible = True
        msg1.Text = ""

        'ddlDISTID_SCH = TIMS.Get_DistID(ddlDISTID_SCH, TIMS.Get_DISTIDT2(objconn))
        'Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID)

        ddlAPPSTAGE_SCH = TIMS.GET_APPSTAGE2_N34(ddlAPPSTAGE_SCH)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '計畫  產業人才投資計畫 / 提升勞工自主學習計畫
        Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        Common.SetListItem(rblOrgKind2, vsOrgKind2)

        If TIMS.sUtl_ChkTest() Then
            'Common.SetListItem(ddlDISTID_SCH, "001")
            Common.SetListItem(ddlAPPSTAGE_SCH, "2")
            Common.SetListItem(rblOrgKind2, "G")
        End If
    End Sub

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Public Function SEARCH_DATA1_dt() As DataTable
        Dim dt As DataTable = Nothing

        'Dim v_ddlDISTID_SCH As String = TIMS.GetListValue(ddlDISTID_SCH) '轄區分署
        'If v_ddlDISTID_SCH = "" Then
        '    msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇轄區分署")
        '    Return dt
        'End If

        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
            Return dt
        End If
        '訓練機構 '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        If v_rblOrgKind2 = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇計畫")
            Return dt
        End If

        'DECLARE @TPLANID NVARCHAR(2)='28';DECLARE @YEARS SMALLINT=CONVERT(SMALLINT,'2024');DECLARE @APPSTAGE NVARCHAR(1)='2';DECLARE @ORGKIND2 NVARCHAR(1)='G';DECLARE @DISTID NVARCHAR(4)='001';
        ' CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}, {"ORGKIND2", v_rblOrgKind2}}

        Dim sSql As String = ""
        sSql = "
WITH WP1 AS ( SELECT PP.RID,PP.GCID3,PP.TPLANID,PP.YEARS,PP.APPSTAGE,PP.ORGKIND2,PP.DISTID,PP.ORGID,PP.COMIDNO
,PP.ORGNAME,dbo.FN_GET_DISTNAME(PP.DISTID,3) DISTNAME3 
,PP.PSNO28,PP.CLASSCNAME2,PF.CURESULT,PF.NGREASON
,dbo.FN_GET_CROSSDIST_YN(PP.YEARS,PP.COMIDNO,PP.APPSTAGE) CROSSDIST_YN
FROM dbo.VIEW2B PP
JOIN dbo.PLAN_STAFFOPIN PF ON PF.PSNO28=PP.PSNO28 
WHERE (PP.RESULTBUTTON IS NULL OR PP.APPLIEDRESULT='Y') 
AND PP.PVR_ISAPPRPAPER='Y' AND PP.DATANOTSENT IS NULL 
AND PF.CURESULT='N'
AND PP.TPLANID=@TPLANID AND PP.YEARS=@YEARS AND PP.APPSTAGE=@APPSTAGE AND PP.ORGKIND2=@ORGKIND2 )
SELECT PP.TPLANID,PP.YEARS,PP.APPSTAGE,PP.ORGKIND2,PP.DISTID,PP.ORGID,PP.COMIDNO
,PP.ORGNAME,dbo.FN_GET_DISTNAME(PP.DISTID,3) DISTNAME3,OO.MASTERNAME
,PP.PSNO28,PP.CLASSCNAME2
,PP.CURESULT,PP.NGREASON
,dbo.FN_GET_CROSSDIST_YN(PP.YEARS,PP.COMIDNO,PP.APPSTAGE) CROSSDIST_YN
FROM WP1 PP
JOIN dbo.VIEW_ORGPLANINFO OO ON OO.RID=PP.RID
JOIN dbo.V_GOVCLASSCAST3 IG3 ON IG3.GCID3=PP.GCID3 
ORDER BY PP.CROSSDIST_YN DESC,PP.ORGNAME,PP.DISTID,PP.PSNO28"

        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms), vbCrLf, "--##CR_03_009:", vbCrLf, sSql))
        End If

        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        Return dt
    End Function

    '匯出 
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
            Return 'dt
        End If
        '訓練機構 '計畫'TRPlanPoint28
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2) 'G/W
        If v_rblOrgKind2 = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇計畫")
            Return 'dt
        End If
        Select Case v_rblOrgKind2
            Case "G"
                Call ExportXlsG1()
            Case "W"
                Call ExportXlsW1()
        End Select

    End Sub

    Sub ExportXlsG1()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls1 As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        'Dim drX1 As DataRow = dtXls1.Rows(0)
        Dim END_COL_NM As String = "F"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim cellsCOLSPNumF2 As String = String.Concat("A2:", END_COL_NM, "{0}") '(畫格子使用)
        Dim strErrmsg As String = ""

        '114年度上半年自主計畫未核班課程明細彙整表							
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        'Dim v_DISTID As String = TIMS.GetListValue(ddlDISTID_SCH)
        'Dim txt_DISTNAME As String = TIMS.GetListText(ddlDISTID_SCH)
        'Dim txt_DISTNAME2 As String = Replace(txt_DISTNAME, "勞動力發展署", "")
        'Dim dtDIST3 As DataTable = TIMS.Get_DISTNAME3dt(objconn)
        'Dim FFF As String = String.Concat("DISTID='", v_DISTID, "'")
        'Dim V_DISTNAME3 As String = If(dtDIST3.Select(FFF).Length > 0, dtDIST3.Select(FFF)(0)("DISTNAME3"), "")
        'Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim V_SHTNM1 As String = If(v_rblOrgKind2 = "G", "產投", If(v_rblOrgKind2 = "W", "自主", "")) '113下自主-北分署
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, V_SHTNM1, "計畫未核班課程明細彙整表")
        '未核班單位明細-自主-上半年
        Dim V_SHEETNM1 As String = String.Concat("未核班單位明細-", V_SHTNM1, "-", s_APPSTAGE_NM2)
        '114上產投計畫未核班課程明細彙整表-1131218(ok)
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, V_SHTNM1, "計畫未核班課程明細彙整表_", TIMS.GetDateNo())

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
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            '差異表序號"	訓練單位名稱	理事長	分署別	課程名稱(含期別)	未核班原因	是否為跨區單位課程(Y/N)
            '序號"	訓練單位名稱	理事長	分署別	課程名稱(含期別)	未核班原因	是否為跨區單位課程(Y/N)
            Dim idxStr As Integer = 2
            SetCellValue(ws, "A" & idxStr, "序號")
            SetCellValue(ws, "B" & idxStr, "訓練單位名稱") 'SetCellValue(ws, "C" & idxStr, "理事長")
            SetCellValue(ws, "C" & idxStr, "分署別")
            SetCellValue(ws, "D" & idxStr, "課程名稱(含期別)")
            SetCellValue(ws, "E" & idxStr, "未核班原因")
            SetCellValue(ws, "F" & idxStr, "是否為跨區單位課程(Y/N)")
            ws.Cells(String.Concat("A1:", END_COL_NM, "2")).Style.Font.Bold = True
            ws.Cells(String.Concat("A1:", END_COL_NM, "2")).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)

            idxStr = 3
            Dim iCNTX As Integer = 0
            For Each dr1 As DataRow In dtXls1.Rows
                iCNTX += 1
                '序號"	訓練單位名稱	理事長	分署別	課程名稱(含期別)	未核班原因	是否為跨區單位課程(Y/N)
                SetCellValue(ws, "A" & idxStr, iCNTX) '序號
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '訓練單位名稱 'SetCellValue(ws, "C" & idxStr, dr1("MASTERNAME")) '理事長
                SetCellValue(ws, "C" & idxStr, dr1("DISTNAME3")) '分署別
                SetCellValue(ws, "D" & idxStr, dr1("CLASSCNAME2"), ExcelHorizontalAlignment.Left) '"課程名稱(含期別)")
                SetCellValue(ws, "E" & idxStr, dr1("NGREASON"), ExcelHorizontalAlignment.Left) '"未核班原因")
                SetCellValue(ws, "F" & idxStr, dr1("CROSSDIST_YN")) '"是否為跨區單位課程(Y/N)")
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
                SetCellBorder(exlRow3X)
            End Using

            '設定貨幣格式，小數位數為 0
            'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            'ws.Cells(String.Format("G3:G{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            'ws.Cells(String.Format("H3:H{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("{0}3:{0}{1}", "A", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}3:{0}{1}", "B", idxStr)).Start.Column).Width = 22 'ws.Column(ws.Cells(String.Format("C3:C{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}3:{0}{1}", "C", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("{0}3:{0}{1}", "D", idxStr)).Start.Column).Width = 22
            ws.Column(ws.Cells(String.Format("{0}3:{0}{1}", "E", idxStr)).Start.Column).Width = 133
            ws.Column(ws.Cells(String.Format("{0}3:{0}{1}", "F", idxStr)).Start.Column).Width = 16

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

    Sub ExportXlsW1()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls1 As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        'Dim drX1 As DataRow = dtXls1.Rows(0)
        Dim END_COL_NM As String = "G"
        Dim cellsCOLSPNumF As String = String.Concat("A{0}:", END_COL_NM, "{0}")
        Dim cellsCOLSPNumF2 As String = String.Concat("A2:", END_COL_NM, "{0}") '資料(畫格子使用)
        Dim strErrmsg As String = ""

        '114年度上半年自主計畫未核班課程明細彙整表							
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        'Dim v_DISTID As String = TIMS.GetListValue(ddlDISTID_SCH)
        'Dim txt_DISTNAME As String = TIMS.GetListText(ddlDISTID_SCH)
        'Dim txt_DISTNAME2 As String = Replace(txt_DISTNAME, "勞動力發展署", "")
        'Dim dtDIST3 As DataTable = TIMS.Get_DISTNAME3dt(objconn)
        'Dim FFF As String = String.Concat("DISTID='", v_DISTID, "'")
        'Dim V_DISTNAME3 As String = If(dtDIST3.Select(FFF).Length > 0, dtDIST3.Select(FFF)(0)("DISTNAME3"), "")
        'Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        Dim V_SHTNM1 As String = If(v_rblOrgKind2 = "G", "產投", If(v_rblOrgKind2 = "W", "自主", "")) '113下自主-北分署
        Dim s_TITLENAME1 As String = String.Concat(s_ROCYEAR1, "年度", s_APPSTAGE_NM2, V_SHTNM1, "計畫未核班課程明細彙整表")
        '未核班單位明細-自主-上半年
        Dim V_SHEETNM1 As String = String.Concat("未核班單位明細-", V_SHTNM1, "-", s_APPSTAGE_NM2)
        '114上產投計畫未核班課程明細彙整表-1131218(ok)
        Dim s_FILENAME1 As String = String.Concat(s_ROCYEAR1, s_APPSTAGE_NM2, V_SHTNM1, "計畫未核班課程明細彙整表_", TIMS.GetDateNo())

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
                    .Style.Font.Size = 16
                    .Value = s_TITLENAME1
                    .Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
                    .Style.VerticalAlignment = ExcelVerticalAlignment.Center
                    .Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
                End With
            End Using

            '差異表序號"	訓練單位名稱	理事長	分署別	課程名稱(含期別)	未核班原因	是否為跨區單位課程(Y/N)
            '序號"	訓練單位名稱	理事長	分署別	課程名稱(含期別)	未核班原因	是否為跨區單位課程(Y/N)
            Dim idxStr As Integer = 2
            SetCellValue(ws, "A" & idxStr, "序號")
            SetCellValue(ws, "B" & idxStr, "訓練單位名稱")
            SetCellValue(ws, "C" & idxStr, "理事長")
            SetCellValue(ws, "D" & idxStr, "分署別")
            SetCellValue(ws, "E" & idxStr, "課程名稱(含期別)")
            SetCellValue(ws, "F" & idxStr, "未核班原因")
            SetCellValue(ws, "G" & idxStr, "是否為跨區單位課程(Y/N)")
            ws.Cells(String.Concat("A1:", END_COL_NM, "2")).Style.Font.Bold = True
            ws.Cells(String.Concat("A1:", END_COL_NM, "2")).Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)

            idxStr = 3
            Dim iCNTX As Integer = 0
            For Each dr1 As DataRow In dtXls1.Rows
                iCNTX += 1
                '序號"	訓練單位名稱	理事長	分署別	課程名稱(含期別)	未核班原因	是否為跨區單位課程(Y/N)
                SetCellValue(ws, "A" & idxStr, iCNTX) '序號
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '訓練單位名稱
                SetCellValue(ws, "C" & idxStr, dr1("MASTERNAME")) '理事長
                SetCellValue(ws, "D" & idxStr, dr1("DISTNAME3")) '分署別
                SetCellValue(ws, "E" & idxStr, dr1("CLASSCNAME2"), ExcelHorizontalAlignment.Left) '"課程名稱(含期別)")
                SetCellValue(ws, "F" & idxStr, dr1("NGREASON"), ExcelHorizontalAlignment.Left) '"未核班原因")
                SetCellValue(ws, "G" & idxStr, dr1("CROSSDIST_YN")) '"是否為跨區單位課程(Y/N)")
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
                SetCellBorder(exlRow3X)
            End Using

            '設定貨幣格式，小數位數為 0
            'ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            'ws.Cells(String.Format("G3:G{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            'ws.Cells(String.Format("H3:H{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("B3:B{0}", idxStr)).Start.Column).Width = 22
            ws.Column(ws.Cells(String.Format("C3:C{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("D3:D{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("E3:E{0}", idxStr)).Start.Column).Width = 22
            ws.Column(ws.Cells(String.Format("F3:F{0}", idxStr)).Start.Column).Width = 133
            ws.Column(ws.Cells(String.Format("G3:G{0}", idxStr)).Start.Column).Width = 16

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


