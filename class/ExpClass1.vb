Imports System.ComponentModel
Imports System.IO
Imports System.Drawing
Imports System.Net
Imports System.Net.Mail
Imports System.Text
Imports NPOI.HSSF.UserModel
Imports NPOI.XSSF.UserModel
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports Spire.Xls

Public Class ExpClass1

    Private Shared ReadOnly print_lock As New Object
    'Dim print_lock As New Object '(); //lock

    'EXCEL 輸出
    'Public Shared Sub Utl_Export1_XLSX(ByRef MyPage As Page, ByRef dtG As DataTable, ByRef s_fileN1 As String, ByRef s_sheetN1 As String, ByRef s_titleRange As String)
    '    Call Utl_Export1_XLSX(MyPage, dtG, s_fileN1, s_sheetN1, s_titleRange, "")
    'End Sub

    ''' <summary> EPPLUS EXCEL 輸出 </summary>
    ''' <param name="dtG"></param>
    ''' <param name="s_fileN1"></param>
    ''' <param name="s_sheetN1"></param>
    ''' <param name="s_titleRange"></param>
    ''' <param name="s_SQL1"></param>
    Public Shared Sub Utl_Export1_XLSX(ByRef MyPage As Page, ByRef dtG As DataTable, ByRef s_fileN1 As String, ByRef s_sheetN1 As String, ByRef s_titleRange As String, ByRef s_SQL1 As String)
        'Dim s_fileN1 As String = String.Format("SQLTMP1Lt-{0}.xlsx", Now.ToString("yyyy-MM-dd-HH-mm-ss"))
        'Dim s_path1 As String = "\SVN\WDAIIP\SRC\Batch\dbt_20210217\dbt_20210217\XLSX\"
        If Not s_fileN1.EndsWith(".xlsx") Then s_fileN1 &= ".xlsx" '補字
        s_fileN1 = TIMS.GetValidFileName(s_fileN1)

        If dtG Is Nothing Then Return

        'If dtG Is Nothing Then
        '    sLogMsg = String.Concat("fileN1:", s_fileN1, vbCrLf, "=> dtG Is Nothing!!!")
        '    Call writeMailBody(sb_gMailBody3, sLogMsg)
        '    writeLog(gLogFiles1, sLogMsg)
        '    Return
        'End If
        'sLogMsg = String.Concat("fileN1:", s_fileN1, vbCrLf, "=> 資料筆數: ", dtG.Rows.Count)
        'Call writeMailBody(sb_gMailBody3, sLogMsg)
        'writeLog(gLogFiles1, sLogMsg)

        'Dim s_path1_XLSX As String = DbAccess.Utl_GetConfigSet("FilePathXLSX") 'Log路徑
        'If s_path1_XLSX = "" Then s_path1_XLSX = ".\"

        '上傳資料夾尾部沒有斜線
        Dim s_G_UPDRV As String = TIMS.GET_G_UPDRV()
        Dim s_DateNo25 As String = TIMS.GetDateNo2(5)
        Dim s_DateNo26 As String = TIMS.GetDateNo2(6)
        Dim s_DateNo3 As String = TIMS.GetDateNo3()
        Dim s_path1_XLSX As String = String.Concat(s_G_UPDRV, "/ExpXlsx", "/", s_DateNo26, "/", s_DateNo25, "/")

        Dim s_ser_MapPath2 As String = MyPage.Server.MapPath(s_path1_XLSX)
        If Not Directory.Exists(s_ser_MapPath2) Then Directory.CreateDirectory(s_ser_MapPath2)
        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        Dim s_ful_xlsFileName As String = String.Format("{0}{1}", s_path1_XLSX, s_fileN1)
        Dim s_filePath1 As String = String.Format("{0}{1}", s_ser_MapPath2, s_fileN1)
        Dim o_filePath1 As System.IO.FileInfo = New System.IO.FileInfo(s_filePath1)

        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial 'ExcelPackage.LicenseContext=LicenseContext.NonCommercial
            'ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
            'save file packages\EPPlus.5.8.14\lib\net45\EPPlus.dll
            Using Epk As New ExcelPackage()

                '新增worksheet
                Dim ws As ExcelWorksheet = Epk.Workbook.Worksheets.Add(s_sheetN1)
                For Each dataCol As DataColumn In dtG.Columns
                    If dataCol.DataType = GetType(Date) Then
                        Dim i_colNumber As Integer = dataCol.Ordinal + 1
                        ws.Column(i_colNumber).Style.Numberformat.Format = "yyyy/MM/dd"
                    End If
                Next

                '將DataTable資料塞到sheet中
                ws.Cells("A1").LoadFromDataTable(dtG, True)
                '自適應寬度設定
                ws.Cells.AutoFitColumns(10, 1000)
                '自適應寬度設定
                ws.Row(1).CustomHeight = True

                Dim totalCols As Integer = If(ws.Dimension Is Nothing, -1, ws.Dimension.End.Column)
                'sLogMsg = String.Format("totalCols: {0} ..", totalCols)
                'writeLog(gLogFiles1, sLogMsg)

                ' "A1:E1"
                Dim rng As ExcelRange = ws.Cells(s_titleRange)
                rng.Style.Font.Bold = True
                rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                rng.Style.Font.Color.SetColor(Color.White)


                If s_SQL1 <> "" Then
                    'https://dotblogs.com.tw/shadow/2012/07/13/73385
                    '寫入標題文字  sheet.Cells[1, 1].Value = "第1欄";
                    'epplus套件-產製、讀、寫excel/
                    'https://tynadesigner.wordpress.com/2019/11/13/epplus%E5%A5%97%E4%BB%B6-%E7%94%A2%E8%A3%BD%E3%80%81%E8%AE%80%E3%80%81%E5%AF%ABexcel/
                    '新增worksheet
                    Using ws2 As ExcelWorksheet = Epk.Workbook.Worksheets.Add("SQL1")
                        '遇\n或(char)10自動斷行
                        ws2.Cells.Style.WrapText = True
                        '自適應寬度設定
                        ws2.Cells.AutoFitColumns(10, 1000)
                        '自適應寬度設定
                        ws2.Row(1).CustomHeight = True
                        'vbCrLf->vbLf
                        If (s_SQL1.IndexOf(vbCrLf) > -1) Then s_SQL1 = Replace(s_SQL1, vbCrLf, vbLf)
                        '將DataTable資料塞到sheet中
                        ws2.Cells(1, 1).Value = s_SQL1
                    End Using
                End If

                Epk.SaveAs(o_filePath1)
            End Using

            'Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)

            Using oConn As SqlConnection = DbAccess.GetConnection()
                'Dim File As New FileInfo(MyPage.Server.MapPath(full_xlsFileName))
                With MyPage
                    Dim File As New FileInfo(.Server.MapPath(s_ful_xlsFileName))
                    ' Clear the content of the response
                    .Response.ClearContent()
                    ' LINE1 Add the file name And attachment, which will force the open/cance/save dialog To show, to the header
                    .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
                    'Response.Headers["Content-Disposition"] = "attachment; filename=" + zipFileName;
                    ' Add the file size into the response header
                    .Response.AddHeader("Content-Length", File.Length.ToString())
                    ' Set the ContentType
                    '.Response.ContentType = "application/zip"
                    '.xlsx	Microsoft Excel (OpenXML)	application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                    .Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                    .Response.TransmitFile(File.FullName)
                    ' End the response
                    TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
                End With
            End Using
            'sLogMsg = String.Format("s_filePath1: {0}匯出檔案完成..", (s_filePath1 & vbCrLf)) 'Call writeMailBody(sb_gMailBody3, sLogMsg) 'writeLog(gLogFiles1, sLogMsg)

        End SyncLock

    End Sub

    Friend Shared Sub Utl_Export1_XLSX(MyPage As Page, dsXlsALL As DataSet, s_fileN1 As String)
        'Dim s_fileN1 As String = String.Format("SQLTMP1Lt-{0}.xlsx", Now.ToString("yyyy-MM-dd-HH-mm-ss"))
        'Dim s_path1 As String = "\SVN\WDAIIP\SRC\Batch\dbt_20210217\dbt_20210217\XLSX\"
        If Not s_fileN1.EndsWith(".xlsx") Then s_fileN1 &= ".xlsx" '補字
        s_fileN1 = TIMS.GetValidFileName(s_fileN1)

        If dsXlsALL Is Nothing OrElse dsXlsALL.Tables.Count = 0 Then Return

        '上傳資料夾尾部沒有斜線
        Dim s_G_UPDRV As String = TIMS.GET_G_UPDRV()
        Dim s_DateNo25 As String = TIMS.GetDateNo2(5)
        Dim s_DateNo26 As String = TIMS.GetDateNo2(6)
        Dim s_DateNo3 As String = TIMS.GetDateNo3()
        Dim s_path1_XLSX As String = String.Concat(s_G_UPDRV, "/ExpXlsx", "/", s_DateNo26, "/", s_DateNo25, "/")

        Dim s_ser_MapPath2 As String = MyPage.Server.MapPath(s_path1_XLSX)
        If Not Directory.Exists(s_ser_MapPath2) Then Directory.CreateDirectory(s_ser_MapPath2)
        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)

        Dim s_ful_xlsFileName As String = String.Format("{0}{1}", s_path1_XLSX, s_fileN1)
        Dim s_filePath1 As String = String.Format("{0}{1}", s_ser_MapPath2, s_fileN1)
        Dim o_filePath1 As System.IO.FileInfo = New System.IO.FileInfo(s_filePath1)

        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial
            'ExcelPackage.LicenseContext=LicenseContext.NonCommercial
            'ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
            'save file
            Using Epk As New ExcelPackage()

                For Each dtG As DataTable In dsXlsALL.Tables
                    Dim s_titleRange As String = GET_XLStitleRange(dtG.Columns.Count)
                    Dim s_sheetN1 As String = dtG.TableName

                    '新增worksheet
                    Dim ws As ExcelWorksheet = Epk.Workbook.Worksheets.Add(s_sheetN1)

                    For Each dataCol As DataColumn In dtG.Columns
                        If dataCol.DataType = GetType(Date) Then
                            Dim i_colNumber As Integer = dataCol.Ordinal + 1
                            ws.Column(i_colNumber).Style.Numberformat.Format = "yyyy/MM/dd"
                        End If
                    Next

                    '將DataTable資料塞到sheet中
                    ws.Cells("A1").LoadFromDataTable(dtG, True)
                    '自適應寬度設定
                    ws.Cells.AutoFitColumns(10, 1000)
                    '自適應寬度設定
                    ws.Row(1).CustomHeight = True

                    Dim totalCols As Integer = If(ws.Dimension Is Nothing, -1, ws.Dimension.End.Column)
                    'sLogMsg = String.Format("totalCols: {0} ..", totalCols)
                    'writeLog(gLogFiles1, sLogMsg)

                    ' "A1:E1" Dim rng As ExcelRange = Nothing
                    Dim rng As ExcelRange = ws.Cells(s_titleRange)
                    rng.Style.Font.Bold = True
                    rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                    rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    rng.Style.Font.Color.SetColor(Color.White)

                Next

                Epk.SaveAs(o_filePath1)
            End Using

            'Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)

            Using oConn As SqlConnection = DbAccess.GetConnection()
                'https://dotblogs.com.tw/malonestudyrecord/2018/03/21/103124
                'Dim sMyFile2 As String = MyPage.Server.MapPath(String.Concat(tmpFileSavePath, myFileName1))
                Using fr As New System.IO.FileStream(s_filePath1, IO.FileMode.Open)
                    'Dim br As New System.IO.BinaryReader(fr)
                    Dim buf(fr.Length - 1) As Byte
                    fr.Read(buf, 0, fr.Length) 'fr.Close()
                    With MyPage
                        .Response.Clear()
                        .Response.ClearHeaders()
                        .Response.Buffer = True
                        .Response.AddHeader("Content-Length", buf.Length.ToString())
                        .Response.AddHeader("content-disposition", "attachment;filename=" & HttpUtility.UrlEncode(s_fileN1, System.Text.Encoding.UTF8))
                        .Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        'Response.ContentType="Application/vnd.ms-Excel" 'Common.RespWrite(Me, br.ReadBytes(fr.Length))
                        .Response.BinaryWrite(buf)
                        .Response.Flush()
                        '.Response.End()
                    End With
                    TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
                End Using
            End Using

            'Using oConn As SqlConnection = DbAccess.GetConnection()
            '    'Dim File As New FileInfo(MyPage.Server.MapPath(full_xlsFileName))
            '    With MyPage
            '        Dim File As New FileInfo(.Server.MapPath(s_ful_xlsFileName))
            '        .Response.ClearContent()
            '        .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
            '        .Response.AddHeader("Content-Length", File.Length.ToString())
            '        ' Set the ContentType '.Response.ContentType = "application/zip" '.xlsx	Microsoft Excel (OpenXML)	application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
            '        .Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            '        .Response.TransmitFile(File.FullName)
            '        TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
            '    End With
            'End Using
            'sLogMsg = String.Format("s_filePath1: {0}匯出檔案完成..", (s_filePath1 & vbCrLf)) 'Call writeMailBody(sb_gMailBody3, sLogMsg) 'writeLog(gLogFiles1, sLogMsg)
        End SyncLock
    End Sub

    Friend Shared Sub Utl_Export2_XLSX(MyPage As Page, dsXlsALL As DataSet, s_fileN1 As String)
        'Dim s_fileN1 As String = String.Format("SQLTMP1Lt-{0}.xlsx", Now.ToString("yyyy-MM-dd-HH-mm-ss"))
        'Dim s_path1 As String = "SVN\WDAIIP\SRC\Batch\dbt_20210217\dbt_20210217\XLSX\"
        If Not s_fileN1.EndsWith(".xlsx") Then s_fileN1 &= ".xlsx" '補字
        s_fileN1 = TIMS.GetValidFileName(s_fileN1)

        If dsXlsALL Is Nothing OrElse dsXlsALL.Tables.Count = 0 Then Return

        '上傳資料夾尾部沒有斜線
        Dim s_G_UPDRV As String = TIMS.GET_G_UPDRV()
        Dim s_DateNo25 As String = TIMS.GetDateNo2(5)
        Dim s_DateNo26 As String = TIMS.GetDateNo2(6)
        Dim s_DateNo3 As String = TIMS.GetDateNo3()
        Dim s_path1_XLSX As String = String.Concat(s_G_UPDRV, "/ExpXlsx", "/", s_DateNo26, "/", s_DateNo25, "/")

        Dim s_ser_MapPath2 As String = MyPage.Server.MapPath(s_path1_XLSX)
        If Not Directory.Exists(s_ser_MapPath2) Then Directory.CreateDirectory(s_ser_MapPath2)
        'Threading.Thread.Sleep(1) '假設處理某段程序需花費1毫秒 (避免機器不同步)
        Dim s_ful_xlsFileName As String = String.Format("{0}{1}", s_path1_XLSX, s_fileN1)
        Dim s_filePath1 As String = String.Format("{0}{1}", s_ser_MapPath2, s_fileN1)
        Dim o_filePath1 As System.IO.FileInfo = New System.IO.FileInfo(s_filePath1)

        SyncLock print_lock
            'ExcelPackage.LicenseContext=LicenseContext.Commercial
            'ExcelPackage.LicenseContext=LicenseContext.NonCommercial
            'ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial
            'save file
            Using Epk As New ExcelPackage()

                For Each dtG As DataTable In dsXlsALL.Tables
                    Dim s_titleRange As String = GET_XLStitleRange(dtG.Columns.Count)
                    Dim s_sheetN1 As String = dtG.TableName

                    '新增worksheet
                    Dim ws As ExcelWorksheet = Epk.Workbook.Worksheets.Add(s_sheetN1)

                    For Each dataCol As DataColumn In dtG.Columns
                        If dataCol.DataType = GetType(Date) Then
                            Dim i_colNumber As Integer = dataCol.Ordinal + 1
                            ws.Column(i_colNumber).Style.Numberformat.Format = "yyyy/MM/dd"
                        End If
                    Next

                    '將DataTable資料塞到sheet中
                    ws.Cells("A1").LoadFromDataTable(dtG, True)
                    '自適應寬度設定
                    ws.Cells.AutoFitColumns(10, 1000)
                    '自適應寬度設定
                    ws.Row(1).CustomHeight = True

                    Dim totalCols As Integer = If(ws.Dimension Is Nothing, -1, ws.Dimension.End.Column)
                    'sLogMsg = String.Format("totalCols: {0} ..", totalCols)
                    'writeLog(gLogFiles1, sLogMsg)

                    ' "A1:E1"
                    'Dim rng As ExcelRange = Nothing
                    'rng = ws.Cells(s_titleRange)
                    'rng.Style.Font.Bold = True
                    'rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                    'rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'rng.Style.Font.Color.SetColor(Color.White)

                Next

                Epk.SaveAs(o_filePath1)
            End Using

            'Call TIMS.WebClientDownloadData(s_RPTURL, s_PDF_byte)
            Using oConn As SqlConnection = DbAccess.GetConnection()
                'Dim File As New FileInfo(MyPage.Server.MapPath(full_xlsFileName))
                With MyPage
                    Dim File As New FileInfo(.Server.MapPath(s_ful_xlsFileName))
                    ' Clear the content of the response
                    .Response.ClearContent()
                    ' LINE1 Add the file name And attachment, which will force the open/cance/save dialog To show, to the header
                    .Response.AddHeader("Content-Disposition", String.Concat("attachment; filename=", File.Name))
                    'Response.Headers["Content-Disposition"] = "attachment; filename=" + zipFileName;
                    ' Add the file size into the response header
                    .Response.AddHeader("Content-Length", File.Length.ToString())
                    ' Set the ContentType
                    '.Response.ContentType = "application/zip"
                    '.xlsx	Microsoft Excel (OpenXML)	application/vnd.openxmlformats-officedocument.spreadsheetml.sheet
                    .Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

                    .Response.TransmitFile(File.FullName)
                    ' End the response
                End With
                TIMS.Utl_RespWriteEnd(MyPage, oConn, "") '.Response.End()
            End Using
            'sLogMsg = String.Format("s_filePath1: {0}匯出檔案完成..", (s_filePath1 & vbCrLf)) 'Call writeMailBody(sb_gMailBody3, sLogMsg) 'writeLog(gLogFiles1, sLogMsg)

        End SyncLock

    End Sub

    ''' <summary>
    ''' 將 DataSet 資料直接匯出為 XLSX 格式並輸出至瀏覽器，不產生伺服器臨時檔案。
    ''' </summary>
    ''' <param name="MyPage">當前 Page 物件。</param>
    ''' <param name="dsXlsALL">包含要匯出資料的 DataSet。</param>
    ''' <param name="s_fileN1">要匯出的檔案名稱，應包含 .xlsx 副檔名。</param>
    Friend Shared Sub Utl_Export2_XLSX_Direct(MyPage As Page, dsXlsALL As DataSet, s_fileN1 As String)
        ' 確保檔案名稱以 .xlsx 結尾
        If Not s_fileN1.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) Then s_fileN1 &= ".xlsx"

        ' 確保 DataSet 和 DataTable 存在
        If dsXlsALL Is Nothing OrElse dsXlsALL.Tables.Count = 0 Then Return

        ' 設定 EPPlus 的授權模式
        'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Using Epk As New ExcelPackage()
            For Each dtG As DataTable In dsXlsALL.Tables
                Dim s_sheetN1 As String = dtG.TableName

                ' 新增一個 Worksheet
                Dim ws As ExcelWorksheet = Epk.Workbook.Worksheets.Add(s_sheetN1)

                ' 設定日期欄位格式
                For Each dataCol As DataColumn In dtG.Columns
                    If dataCol.DataType = GetType(Date) Then
                        Dim i_colNumber As Integer = dataCol.Ordinal + 1
                        ws.Column(i_colNumber).Style.Numberformat.Format = "yyyy/MM/dd"
                    End If
                Next

                ' 將 DataTable 資料載入到 Worksheet
                ws.Cells("A1").LoadFromDataTable(dtG, True)
                ' 自適應寬度設定
                ws.Cells.AutoFitColumns(10, 1000)
                ws.Row(1).CustomHeight = True

                ' 設定標題欄位樣式（如果需要）
                ' Dim s_titleRange As String = GET_XLStitleRange(dtG.Columns.Count)
                ' If Not String.IsNullOrEmpty(s_titleRange) Then
                '     Dim rng As ExcelRange = ws.Cells(s_titleRange)
                '     rng.Style.Font.Bold = True
                '     rng.Style.Fill.PatternType = ExcelFillStyle.Solid
                '     rng.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                '     rng.Style.Font.Color.SetColor(Color.White)
                ' End If
            Next
            ' 將 ExcelPackage 內容寫入一個記憶體流
            Using ms As New MemoryStream()
                Epk.SaveAs(ms)
                ms.Position = 0

                ' 清空當前網頁 Response
                With MyPage.Response
                    .Clear()
                    .ClearContent()
                    .ClearHeaders()
                    ' 設定 Response Headers 以便下載檔案
                    .ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    .AddHeader("Content-Disposition", $"attachment; filename={HttpUtility.UrlEncode(s_fileN1, System.Text.Encoding.UTF8)}")
                    .AddHeader("Content-Length", ms.Length.ToString())
                    ' 將記憶體流的內容寫入 Response
                    ms.WriteTo(.OutputStream)
                    ' **[修改]** Response.Flush() 保持不變，它用於立即傳送緩衝區內容
                    .Flush()
                End With
                '替代 .End()，避免 System.Threading.ThreadAbortException 
                TIMS.Utl_RespWriteEnd(MyPage, "")
            End Using

        End Using
    End Sub

    ''' <summary>
    ''' 將 DataSet 資料直接匯出為 ODS 格式並輸出至瀏覽器，不產生伺服器臨時檔案。
    ''' </summary>
    ''' <param name="MyPage">當前 Page 物件。</param>
    ''' <param name="dsOdsALL">包含要匯出資料的 DataSet。</param>
    ''' <param name="s_fileN1">要匯出的檔案名稱，應包含 .ods 副檔名。</param>
    Friend Shared Sub Utl_Export2_ODS_Direct(MyPage As Page, dsOdsALL As DataSet, s_fileN1 As String)
        ' 確保檔案名稱以 .ods 結尾
        If Not s_fileN1.EndsWith(".ods", StringComparison.OrdinalIgnoreCase) Then s_fileN1 &= ".ods"

        ' 確保 DataSet 和 DataTable 存在
        If dsOdsALL Is Nothing OrElse dsOdsALL.Tables.Count = 0 Then Return

        ' 如果您有授權檔案
        Spire.License.LicenseProvider.SetLicenseFileName("license.elic.xml")

        ' 或者，如果官方提供授權碼 'License.SetLicenseKey("您的授權碼")

        ' 創建一個新的 Workbook 物件
        Using workbook As New Workbook()
            For Each dtG As DataTable In dsOdsALL.Tables
                Dim s_sheetN1 As String = dtG.TableName

                Dim ws As Worksheet = Nothing
                If workbook.Worksheets.Count > 0 Then
                    If workbook.Worksheets.Count > 4 Then workbook.Worksheets(4).Remove()
                    If workbook.Worksheets.Count > 3 Then workbook.Worksheets(3).Remove()
                    If workbook.Worksheets.Count > 2 Then workbook.Worksheets(2).Remove()
                    If workbook.Worksheets.Count > 1 Then workbook.Worksheets(1).Remove()
                    '使用第1個 Worksheet 並改名
                    ws = workbook.Worksheets(0)
                    ws.Name = s_sheetN1
                Else
                    '新增一個 Worksheet
                    ws = workbook.Worksheets.Add(s_sheetN1)
                End If

                ' 將 DataTable 資料載入到 Worksheet
                ws.InsertDataTable(dtG, True, 1, 1)

                ' 設定日期欄位格式
                Dim i_Col As Integer = 1
                For Each dataCol As DataColumn In dtG.Columns
                    If dataCol.DataType = GetType(Date) Then
                        Dim i_colNumber As Integer = dataCol.Ordinal + 1
                        ' Spire.XLS 的日期格式設定
                        ws.Columns(i_colNumber - 1).Style.NumberFormat = "yyyy/MM/dd"
                    End If
                    ' 自動調整欄寬
                    ws.AutoFitColumn(i_Col)
                    i_Col += 1
                Next

                ' 可選：設定標題樣式
                ' Dim titleRange As CellRange = ws.Range(1, 1, 1, dtG.Columns.Count)
                ' titleRange.Style.Font.IsBold = True
                ' titleRange.Style.FillPattern = ExcelPatternType.Solid
                ' titleRange.Style.KnownColor = ExcelColors.CornflowerBlue
                ' titleRange.Style.Font.Color = Color.White
            Next

            ' 將 Workbook 內容以 ODS 格式寫入記憶體流
            Using ms As New MemoryStream()
                workbook.SaveToStream(ms, FileFormat.ODS)
                ms.Position = 0

                ' 清空當前網頁 Response
                With MyPage.Response
                    .Clear()
                    .ClearContent()
                    .ClearHeaders()

                    ' 設定 Response Headers 以便下載檔案
                    .ContentType = "application/vnd.oasis.opendocument.spreadsheet"
                    .AddHeader("Content-Disposition", String.Concat("attachment; filename=", HttpUtility.UrlEncode(s_fileN1, System.Text.Encoding.UTF8)))
                    .AddHeader("Content-Length", ms.Length.ToString())

                    ' 將記憶體流的內容寫入 Response
                    ms.WriteTo(.OutputStream)

                    ' 結束 Response 流程
                    .Flush()
                End With
                '替代 .End()，避免 System.Threading.ThreadAbortException 
                TIMS.Utl_RespWriteEnd(MyPage, "")
            End Using
        End Using
    End Sub

    ''' <summary>計算EXCEL RANGE 返回文字 A1:AG1</summary>
    ''' <param name="xls_cnt"></param>
    ''' <returns></returns>
    Public Shared Function GET_XLStitleRange(ByVal xls_cnt As Integer) As String
        Dim rst As String = "A1:AG1"
        Dim CharEng As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        Dim tmp1 As String = ""

        Dim i As Integer = 1
        For Each s2 As String In CharEng
            If i = xls_cnt Then
                tmp1 = String.Concat("A1:", s2, 1)
                Return tmp1
                Exit For
            End If
            i += 1
        Next
        For Each s1 As String In CharEng
            For Each s2 As String In CharEng
                If i = xls_cnt Then
                    tmp1 = String.Concat("A1:", s1, s2, 1)
                    Return tmp1
                    Exit For
                End If
                i += 1
            Next
        Next
        Return rst
    End Function

    ''' <summary>
    ''' 匯出資料到 XLSX 檔案
    ''' </summary>
    ''' <param name="filePath">要儲存的完整檔案路徑 (例如: "C:\Temp\ExportData.xlsx")</param>
    Public Shared Sub ExportDataToXlsx(ByVal filePath As String)

        ' 1. 建立一個新的工作簿 (Workbook) 實例 (適用於 XLSX 格式)
        Dim workbook As IWorkbook = New XSSFWorkbook()

        ' 2. 建立一個名為 "Sheet1" 的工作表 (Sheet)
        Dim sheet As ISheet = workbook.CreateSheet("Sheet1")

        ' 3. 設定標題列 (第一列: 索引 0) 
        Dim headerRow As IRow = sheet.CreateRow(0)

        ' 建立標題單元格 (Cell)
        headerRow.CreateCell(0).SetCellValue("欄位 A")
        headerRow.CreateCell(1).SetCellValue("欄位 B")
        headerRow.CreateCell(2).SetCellValue("欄位 C")

        ' --- 4. 寫入資料列 (從第二列開始: 索引 1) ---

        ' 第 2 列 (索引 1)
        Dim dataRow1 As IRow = sheet.CreateRow(1)
        dataRow1.CreateCell(0).SetCellValue("資料-1A")
        dataRow1.CreateCell(1).SetCellValue(123) ' 數值型資料
        dataRow1.CreateCell(2).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"))

        ' 第 3 列 (索引 2)
        Dim dataRow2 As IRow = sheet.CreateRow(2)
        dataRow2.CreateCell(0).SetCellValue("資料-2A")
        dataRow2.CreateCell(1).SetCellValue(456.78) ' 小數型資料
        dataRow2.CreateCell(2).SetCellValue(True) ' 布林型資料

        ' 為了美觀，可以讓欄位寬度自動調整
        sheet.AutoSizeColumn(0)
        sheet.AutoSizeColumn(1)
        sheet.AutoSizeColumn(2)

        ' --- 5. 將工作簿寫入檔案 ---
        Try
            Using fs As New FileStream(filePath, FileMode.Create, FileAccess.Write)
                workbook.Write(fs)
            End Using

            ' 選用：顯示成功訊息或日誌記錄
            Console.WriteLine($"Excel 檔案已成功匯出到: {filePath}")

        Catch ex As Exception
            ' 處理寫入檔案時發生的錯誤
            Console.WriteLine($"匯出 Excel 檔案時發生錯誤: {ex.Message}")
        Finally
            ' 釋放資源 (儘管 Using 塊已經處理了 FileStream，但對於 Workbook 而言，這樣做是良好的習慣)
            If Not workbook Is Nothing Then
                ' 在較新版本的 NPOI 中，IWorkbook 實作了 IDisposable，但在舊版本可能沒有，
                ' 但為了檔案輸出完整，將其設為 Nothing 也是一種釋放物件參考的方式。
                workbook = Nothing
            End If
        End Try

    End Sub
    ''' <summary>
    ''' 匯出資料並觸發瀏覽器下載 XLSX 檔案
    ''' </summary>
    Public Shared Sub ExportDataForWebDownload()
        'Imports System.IO
        'Imports System.Web ' 引入 System.Web 命名空間以使用 Response 物件
        'Imports NPOI.XSSF.UserModel
        'Imports NPOI.SS.UserModel
        Dim fileName As String = $"ExportData_{DateTime.Now.ToString("yyyyMMddHHmmss")}.xlsx"

        ' 1. 建立工作簿並寫入資料 (與之前範例相同)
        Dim workbook As IWorkbook = New XSSFWorkbook()
        Dim sheet As ISheet = workbook.CreateSheet("Sheet1")

        ' --- 建立標題列 ---
        Dim headerRow As IRow = sheet.CreateRow(0)
        headerRow.CreateCell(0).SetCellValue("產品名稱")
        headerRow.CreateCell(1).SetCellValue("庫存量")
        headerRow.CreateCell(2).SetCellValue("更新時間")

        ' --- 寫入一些範例資料 ---
        Dim dataRow1 As IRow = sheet.CreateRow(1)
        dataRow1.CreateCell(0).SetCellValue("筆記型電腦")
        dataRow1.CreateCell(1).SetCellValue(50)
        dataRow1.CreateCell(2).SetCellValue(DateTime.Now.ToString("yyyy/MM/dd HH:mm"))

        sheet.AutoSizeColumn(0)
        sheet.AutoSizeColumn(1)
        sheet.AutoSizeColumn(2)

        ' 2. 將 Workbook 寫入 MemoryStream
        Using ms As New MemoryStream()
            workbook.Write(ms)
            ms.Flush()

            ' 3. 設定 HTTP 響應頭，觸發瀏覽器下載
            ' 清空 Response 緩衝區
            HttpContext.Current.Response.Clear()

            ' 設定內容類型為 XLSX 格式
            HttpContext.Current.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

            ' 設定檔案下載的名稱 (Content-Disposition)
            ' 注意: 如果檔名包含中文，可能需要進行 URL 編碼或使用特定的編碼設定
            HttpContext.Current.Response.AddHeader("Content-Disposition", $"attachment; filename={HttpUtility.UrlEncode(fileName, System.Text.Encoding.UTF8)}")

            ' 設定內容長度
            HttpContext.Current.Response.AddHeader("Content-Length", ms.Length.ToString())

            ' 4. 將 MemoryStream 內容寫入 Response 輸出流
            HttpContext.Current.Response.BinaryWrite(ms.ToArray())

            ' 5. 停止處理頁面，確保檔案下載成功
            HttpContext.Current.Response.End()

        End Using ' Using 塊會自動釋放 MemoryStream 資源

    End Sub

End Class
