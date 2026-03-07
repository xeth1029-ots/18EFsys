Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CR_03_006
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
            cCreate1()
        End If
    End Sub

    Sub cCreate1()
        PanelSch1.Visible = True

        msg1.Text = ""

        'ddlYEARS_SCH = TIMS.GetSyear(ddlYEARS_SCH) 'Common.SetListItem(ddlYEARS_SCH, sm.UserInfo.Years)

        'cblDistid = TIMS.Get_DistID(cblDistid)
        'Dim objid000 As Web.UI.WebControls.ListItem = cblDistid.Items.FindByValue("000")
        'If objid000 IsNot Nothing Then cblDistid.Items.Remove(objid000)
        'cblDistid.Items.Insert(0, New ListItem("全部", 0))
        'cblDistid.Enabled = True
        'If sm.UserInfo.DistID <> "000" Then
        '    'Common.SetListItem(cblDistid, sm.UserInfo.DistID)
        '    TIMS.SetCblValue(cblDistid, sm.UserInfo.DistID)
        '    cblDistid.Enabled = False
        'End If
        'cblDistid.Attributes("onclick") = "SelectAll('cblDistid','DistHidden');"

        ddlAPPSTAGE_SCH = TIMS.GET_APPSTAGE2_N34(ddlAPPSTAGE_SCH)
        'Dim objItem3 As Web.UI.WebControls.ListItem = ddlAPPSTAGE_SCH.Items.FindByValue("3")
        'If objItem3 IsNot Nothing Then ddlAPPSTAGE_SCH.Items.Remove(objItem3)
        'Dim objItem4 As Web.UI.WebControls.ListItem = ddlAPPSTAGE_SCH.Items.FindByValue("4")
        'If objItem4 IsNot Nothing Then ddlAPPSTAGE_SCH.Items.Remove(objItem4)
        Common.SetListItem(ddlAPPSTAGE_SCH, "1")

        '計畫  產業人才投資計畫/提升勞工自主學習計畫
        'Dim vsOrgKind2 As String = TIMS.Get_OrgKind2(sm.UserInfo.OrgID, TIMS.c_ORGID, objconn)
        'If (vsOrgKind2 = "") Then vsOrgKind2 = "G"
        'rblOrgKind2 = TIMS.Get_RblSearchPlan(rblOrgKind2, objconn, False)
        'Common.SetListItem(rblOrgKind2, vsOrgKind2)

        If TIMS.sUtl_ChkTest() Then
            Common.SetListItem(ddlAPPSTAGE_SCH, "2")
        End If
    End Sub

    ''' <summary> 查詢SQL DataTable </summary>
    ''' <returns></returns>
    Public Function SEARCH_DATA1_dt() As DataTable
        Dim dt As DataTable = Nothing

        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
            Return dt
        End If
        'sSql &= " DECLARE @TPLANID VARCHAR(4) ='28' DECLARE @YEARS VARCHAR(4) ='2024' DECLARE @APPSTAGE INT =1; DECLARE @ORGKIND2 VARCHAR(4) ='W';" & vbCrLf
        'DECLARE @TPLANID NVARCHAR(2)='28';/*2*/DECLARE @YEARS SMALLINT=CONVERT(SMALLINT,'2024');/*3*/DECLARE @APPSTAGE NVARCHAR(1)='2';/*4*/DECLARE @ORGKIND2 NVARCHAR(1)='G';/*1*/
        ' CURESULT 核班結果,核班結果'Y 通過、N 不通過
        Dim parms As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}, {"YEARS", sm.UserInfo.Years}, {"APPSTAGE", v_APPSTAGE_SCH}}
        Dim sSql As String = ""
        sSql = "
SELECT pp.DISTID,pp.COMIDNO,pp.TPLANID
,DENSE_RANK() OVER (ORDER BY pp.ORGNAME ) AS ROWID
,pp.ORGNAME,pp.CLASSCNAME2 ,pp.THOURS,pp.TNUM,pp.TOTAL,pp.TOTALCOST,pp.DEFGOVCOST
,format(pp.STDATE,'yyyy/MM/dd') STDATE
,format(pp.FTDATE,'yyyy/MM/dd') FTDATE
,pp.CTNAME,pp.CONTACTNAME,pp.CONTACTPHONE,pp.CONTACTMOBILE 
FROM dbo.VIEW2B PP
JOIN dbo.V_GOVCLASSCAST3 IG3 ON IG3.GCID3=PP.GCID3
JOIN dbo.PLAN_STAFFOPIN pf on pf.PSNO28=pp.PSNO28 
WHERE (pp.RESULTBUTTON IS NULL OR pp.APPLIEDRESULT='Y') 
AND pp.PVR_ISAPPRPAPER='Y' AND pp.DATANOTSENT IS NULL /*'審核送出(已送審)/正式/未檢送資料註記(排除有勾選)*/
AND PF.CURESULT='Y' AND pp.POINTYN='Y'
AND pp.TPLANID=@TPLANID AND pp.YEARS=@YEARS AND PP.APPSTAGE=@APPSTAGE
ORDER BY pp.ORGNAME
"
        If TIMS.sUtl_ChkTest() Then
            TIMS.WriteLog(Me, String.Concat("--", vbCrLf, TIMS.GetMyValue5(parms)))
            TIMS.WriteLog(Me, String.Concat("--##CR_03_006.aspx , sSql:", vbCrLf, sSql))
        End If
        dt = DbAccess.GetDataTable(sSql, objconn, parms)

        Return dt
    End Function

    '匯出 
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        'Dim V_cblDistid As String = TIMS.GetCblValue(cblDistid)
        'If V_cblDistid = "" Then
        '    msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇轄區分署")
        '    Return 'dt
        'End If
        ''訓練機構 '計畫'TRPlanPoint28
        'Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        'If v_rblOrgKind2 = "" Then
        '    msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇計畫")
        '    Return 'dt
        'End If
        Dim v_APPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH) '申請階段
        If v_APPSTAGE_SCH = "" Then
            msg1.Text = String.Concat(TIMS.cst_NODATAMsg2, "請選擇申請階段")
            Return 'dt
        End If

        Call ExportXls1()
    End Sub

    Sub ExportXls1()
        Const Cst_FileSavePath As String = "~/CR/03/Temp/" '"~\CO\01\Temp\"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        Dim dtXls1 As DataTable = SEARCH_DATA1_dt()
        If TIMS.dtNODATA(dtXls1) Then
            Common.MessageBox(Me, "查無 匯出資料。")
            Exit Sub
        End If

        Dim cellsCOLSPNumF As String = "A{0}:M{0}"
        Dim cellsCOLSPNumF2 As String = "A2:M{0}"
        Dim strErrmsg As String = ""
        '勞動部勞動力發展署各分署113年度下半年產業人才投資計畫學分班訓練課程統計表
        'Dim vDISTNAME As String = Convert.ToString(dtXls1.Rows(0)("DISTNAME"))
        'Dim v_cblDistid As String = TIMS.GetCblValue(cblDistid)
        'Dim sp_cblDistid As String() = v_cblDistid.Split(",")
        Dim V_SHEETNM1 As String = "學分班彙整(陳核版)"
        Dim s_TITLE1 As String = "勞動部勞動力發展署各分署"
        Dim s_ROCYEAR1 As String = CStr(CInt(sm.UserInfo.Years) - 1911) '年度
        Dim v_ddlAPPSTAGE_SCH As String = TIMS.GetListValue(ddlAPPSTAGE_SCH)
        Dim s_APPSTAGE_NM2 As String = TIMS.GET_APPSTAGE2_NM2(v_ddlAPPSTAGE_SCH) '申請階段
        'Dim v_rblOrgKind2 As String = TIMS.GetListValue(rblOrgKind2)
        'Dim s_PLANNAME As String = TIMS.GetListText(rblOrgKind2) '計畫
        Dim s_PLANNAME As String = String.Concat(Replace(TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn), "方案", ""), "計畫")
        Dim s_TITLENAME1 As String = String.Concat(s_TITLE1, s_ROCYEAR1, "年度", s_APPSTAGE_NM2, s_PLANNAME, "學分班訓練課程統計表")

        SyncLock print_lock
            'ExcelPackage.LicenseContext = LicenseContext.Commercial
            'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

            'Dim file1 As New FileInfo(filePath1)
            Dim ndt As DateTime = Now
            Dim ep As New ExcelPackage()

            Dim ws As ExcelWorksheet = ep.Workbook.Worksheets.Add(V_SHEETNM1)
            'Dim ws As ExcelWorksheet = ep.Workbook.Worksheets(0)

            ' 共用設定 'Dim fontName As String = "標楷體" 'Dim fontSize12s As Single = 12.0F '報表標題
            Dim exlRow1 As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF, "1"))
            exlRow1.Merge = True
            exlRow1.Style.Font.Name = fontName
            exlRow1.Style.Font.Size = 16
            exlRow1.Value = s_TITLENAME1
            exlRow1.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center
            exlRow1.Style.VerticalAlignment = ExcelVerticalAlignment.Center
            exlRow1.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)

            SetCellValue(ws, "A2", "序號")
            SetCellValue(ws, "B2", "訓練單位名稱")
            SetCellValue(ws, "C2", "課程名稱(含期別)")
            SetCellValue(ws, "D2", "訓練時數")
            SetCellValue(ws, "E2", "訓練人次")
            SetCellValue(ws, "F2", "每人訓練費用(元)")
            SetCellValue(ws, "G2", "訓練單位可向學員收取之訓練費用(元)")
            SetCellValue(ws, "H2", "總補助費(元)(以訓練費用之80%估算)")
            SetCellValue(ws, "I2", "開訓日期")
            SetCellValue(ws, "J2", "結訓日期")
            SetCellValue(ws, "K2", "辦訓縣市別")
            SetCellValue(ws, "L2", "聯絡人")
            SetCellValue(ws, "M2", "聯絡電話")
            ws.Cells("A1:M2").Style.Font.Bold = True

            Dim idxStr As Integer = 3
            Dim idx1 As Integer = idxStr
            Dim V_COMIDNO As String = dtXls1.Rows(0)("COMIDNO")
            Dim iCNT As Integer = 0
            Dim iCNTX As Integer = 0
            Dim iTHOURS As Integer = 0
            Dim iTNUM As Integer = 0
            Dim iTOTAL As Integer = 0
            Dim iTOTALCOST As Integer = 0
            Dim iDEFGOVCOST As Integer = 0

            For Each dr1 As DataRow In dtXls1.Rows
                iCNTX += 1
                iTHOURS += TIMS.VAL1(dr1("THOURS"))
                iTNUM += TIMS.VAL1(dr1("TNUM"))
                iTOTAL += TIMS.VAL1(dr1("TOTAL"))
                iTOTALCOST += TIMS.VAL1(dr1("TOTALCOST"))
                iDEFGOVCOST += TIMS.VAL1(dr1("DEFGOVCOST"))
                If V_COMIDNO = dr1("COMIDNO") Then
                    iCNT += 1
                Else
                    '(合計) 'idxStr += 1
                    Dim idx2b As Integer = idx1 + iCNT - 1
                    ws.Cells(String.Format("A{0}:A{1}", idx1, idx2b)).Merge = True
                    ws.Cells(String.Format("B{0}:B{1}", idx1, idx2b)).Merge = True

                    idx1 = idxStr
                    V_COMIDNO = dr1("COMIDNO")
                    iCNT = 1
                End If

                Dim V_CONTACTPHONE As String = ""
                If Convert.ToString(dr1("CONTACTPHONE")) <> "" AndAlso Convert.ToString(dr1("CONTACTMOBILE")) <> "" Then
                    V_CONTACTPHONE = String.Concat(dr1("CONTACTPHONE"), "、", dr1("CONTACTMOBILE"))
                ElseIf Convert.ToString(dr1("CONTACTPHONE")) <> "" Then
                    V_CONTACTPHONE = Convert.ToString(dr1("CONTACTPHONE"))
                ElseIf Convert.ToString(dr1("CONTACTMOBILE")) <> "" Then
                    V_CONTACTPHONE = Convert.ToString(dr1("CONTACTMOBILE"))
                End If

                SetCellValue(ws, "A" & idxStr, dr1("ROWID")) '"序號")
                SetCellValue(ws, "B" & idxStr, dr1("ORGNAME"), ExcelHorizontalAlignment.Left) '"訓練單位名稱")
                SetCellValue(ws, "C" & idxStr, dr1("CLASSCNAME2"), ExcelHorizontalAlignment.Left) '"課程名稱(含期別)")
                SetCellValue(ws, "D" & idxStr, dr1("THOURS")) '"訓練時數")
                SetCellValue(ws, "E" & idxStr, dr1("TNUM")) '"訓練人次")
                SetCellValue(ws, "F" & idxStr, dr1("TOTAL")) '"每人訓練費用(元)")
                SetCellValue(ws, "G" & idxStr, dr1("TOTALCOST")) '"訓練單位可向學員收取之訓練費用(元)")
                SetCellValue(ws, "H" & idxStr, dr1("DEFGOVCOST")) '"總補助費(元)(以訓練費用之")
                SetCellValue(ws, "I" & idxStr, dr1("STDATE")) '"開訓日期")
                SetCellValue(ws, "J" & idxStr, dr1("FTDATE")) '"開訓日期")
                SetCellValue(ws, "K" & idxStr, dr1("CTNAME")) '"辦訓縣市別")
                SetCellValue(ws, "L" & idxStr, dr1("CONTACTNAME")) '"聯絡人")
                SetCellValue(ws, "M" & idxStr, V_CONTACTPHONE, ExcelHorizontalAlignment.Left) '"聯絡電話")

                idxStr += 1
            Next
            '(合計) idxStr += 1
            Dim idx2 As Integer = idx1 + iCNT - 1
            ws.Cells(String.Format("A{0}:A{1}", idx1, idx2)).Merge = True
            ws.Cells(String.Format("B{0}:B{1}", idx1, idx2)).Merge = True

            SetCellValue(ws, String.Format("A{0}:B{0}", idxStr), "合計")
            SetCellValue(ws, "C" & idxStr, iCNTX)
            SetCellValue(ws, "D" & idxStr, iTHOURS) '"訓練人次")
            SetCellValue(ws, "E" & idxStr, iTNUM) '"訓練人次")
            SetCellValue(ws, "F" & idxStr, iTOTAL) '"每人訓練費用(元)")
            SetCellValue(ws, "G" & idxStr, iTOTALCOST) '"訓練單位可向學員收取之訓練費用(元)")
            SetCellValue(ws, "H" & idxStr, iDEFGOVCOST) '"總補助費(元)(以訓練費用之80%估算)")
            SetCellValue(ws, "I" & idxStr, "") '"結訓日期")
            SetCellValue(ws, "J" & idxStr, "") '"單位屬性")
            SetCellValue(ws, "K" & idxStr, "") '"辦訓縣市別")
            SetCellValue(ws, "L" & idxStr, "") '"課程流水號")
            SetCellValue(ws, "M" & idxStr, "") '"聯絡人")

            ws.Cells(String.Format("A{0}:B{0}", idxStr)).Merge = True
            ws.Cells(String.Format("A{0}:M{0}", idxStr)).Style.Font.Bold = True

            'var rangeTxt = String.Format(cellsCOLSPNumF2, idxStr) // "A5:K" + rowIdx.ToString();
            Dim exlRow3X As ExcelRange = ws.Cells(String.Format(cellsCOLSPNumF2, idxStr))
            exlRow3X.Style.Font.Name = fontName
            exlRow3X.Style.Font.Size = fontSize12s 'FontSize
            exlRow3X.Style.Border.BorderAround(ExcelBorderStyle.Thin, Color.Black)
            exlRow3X.AutoFitColumns(25.0, 250.0)
            SetCellBorder(exlRow3X)

            ' 設定貨幣格式，小數位數為 0
            ws.Cells(String.Format("F3:F{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("G3:G{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Cells(String.Format("H3:H{0}", idxStr)).Style.Numberformat.Format = "$#,##0" ' 美元符號，您可以根據需要更改
            ws.Column(ws.Cells(String.Format("A3:A{0}", idxStr)).Start.Column).Width = 10
            ws.Column(ws.Cells(String.Format("D3:D{0}", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("E3:E{0}", idxStr)).Start.Column).Width = 15

            ws.Column(ws.Cells(String.Format("I3:I{0}", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("J3:J{0}", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("K3:K{0}", idxStr)).Start.Column).Width = 15
            ws.Column(ws.Cells(String.Format("L3:L{0}", idxStr)).Start.Column).Width = 15

            ' 設定工作表的顯示比例為 70%  worksheet.View.Zoom = 70 無法運行 修正為 ws.View.ZoomScale = 70 才可運行
            ws.View.ZoomScale = 90

            Dim V_ExpType As String = TIMS.GetListValue(RBListExpType)
            Select Case V_ExpType
                Case "EXCEL"
                    TIMS.ExpExcel_1(Me, strErrmsg, ep, Cst_FileSavePath)
                    TIMS.Utl_RespWriteEnd(Me, objconn, "")
                Case "ODS"
                    TIMS.ExpODSl_1(Me, strErrmsg, ep, Cst_FileSavePath)
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

