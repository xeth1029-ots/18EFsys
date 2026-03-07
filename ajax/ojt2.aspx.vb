Imports System.Net
Imports System.IO
Imports System.Text
Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Net.Security

Public Class ojt2
    Inherits System.Web.UI.Page


    'Partial Public Class _Default
    '    Inherits System.Web.UI.Page
    '    ' 待測試的檔案清單
    '    Private ReadOnly fileUrls As String() = {
    '        "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-73502108-13/20250808/1/A.pdf",
    '        "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-73502108-13/20250808/1/B.pdf",
    '        "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-78965330-8/20250807/1/C.pdf"
    '    }
    '    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    '    End Sub
    'End Class

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            CCREATE1()
        End If
    End Sub
    Sub CCREATE1()
        ' 待測試的檔案清單
        Dim fileUrls As String() = {
            "https://ojtims.wda.gov.tw/upojt/2024/2/5113/G5664/113G204290002/B2404290609x1445x32.pdf",
            "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-72455161-1/20250807/1/PR2508081935xF5543x41x14.pdf",
            "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-72455161-1/20250807/1/PR2508085112xF5543x42x14.pdf",
            "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-78965330-4/20250807/1/PR2508085533xF5468x101x1.pdf",
            "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-73502108-13/20250808/1/PR2508081232xF5547x53x15.pdf",
            "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-73502108-13/20250808/1/PR2508080052xF5547x51x15.pdf",
            "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-78965330-8/20250807/1/PR2508081507xF5468x101x1.pdf"
        }

        'Dim fileUrls As String() = {
        '    "https://ojtims.wda.gov.tw/upojt/2024/2/5113/G5664/113G204290002/B2404290609x1445x32.pdf",
        '    "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-72455161-1/20250807/1/PR2508081935xF5543x41x14.pdf",
        '    "https://ojtims.wda.gov.tw/upojt/2025/REVISE/5162-72455161-1/20250807/1/PR2508085112xF5543x42x14.pdf"
        '}

        For Each tmp1 As String In fileUrls
            TextBox1.Text += tmp1 & vbCrLf
        Next
    End Sub

    Protected Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim txtBox1 As String = TextBox1.Text
        Dim uurl1 As String = ""
        For Each sTMP1 As String In txtBox1.Split(vbCrLf)
            If sTMP1 <> "" Then
                uurl1 &= String.Concat(If(uurl1 <> "", "^", ""), sTMP1)
            End If
        Next

        Dim fileUrls As String() = uurl1.Split("^")
        'StartTest(fileUrls)
        ' 顯示訊息給使用者，告知測試正在背景執行
        'labResult.Text = "檔案下載測試正在進行中，請稍候..."

        'WebRequest物件如何忽略憑證問題
        ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 'HttpClient
        System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 ' 3072 'System.Net.SecurityProtocolType.Tls12

        'Dim report As String = TestFileDownloads(fileUrls)
        'labResult.Text = report

        Dim report As String = TestFileDownloads2(fileUrls)
        labResult.Text = report

        ' 將下載邏輯移至背景 Task
        'Dim task As Task = task.Run(
        '    Sub()
        '        Try
        '            Dim report As String = TestFileDownloads(fileUrls)

        '            ' 使用 RegisterAsyncTask 安全地更新 UI
        '            Me.RegisterAsyncTask(New PageAsyncTask(
        '                Sub()
        '                    txtResult.Text = report
        '                End Sub
        '            ))
        '        Catch ex As Exception
        '            ' 處理下載過程中可能發生的錯誤
        '            Me.RegisterAsyncTask(New PageAsyncTask(
        '                Sub()
        '                    txtResult.Text = $"發生錯誤：{ex.Message}"
        '                End Sub
        '            ))
        '        End Try
        '    End Sub
        ')

    End Sub

    ' 測試多個檔案的下載耗時並生成報告。包含下載測試結果的文字報告
    'Private Function TestFileDownloads(fileUrls As String()) As String
    '    Dim reportBuilder As New StringBuilder()
    '    reportBuilder.AppendLine("檔案下載耗時測試報告")
    '    reportBuilder.AppendLine($"測試開始時間: {DateTime.Now.ToString()}")
    '    reportBuilder.AppendLine("--------------------------------------------------")

    '    'WebRequest物件如何忽略憑證問題
    '    'ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
    '    'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 'HttpClient
    '    'System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 ' 3072 'System.Net.SecurityProtocolType.Tls12

    '    Using webClient As New WebClient()
    '        For Each url As String In fileUrls
    '            url = Replace(Replace(Replace(url, vbCrLf, ""), vbLf, ""), vbCr, "")
    '            If url = "" Then Continue For

    '            Dim exMessage As String = ""
    '            Dim fgOK As Boolean = False
    '            reportBuilder.AppendLine($"開始下載: {url}")
    '            Dim stopwatch As New Stopwatch()
    '            stopwatch.Start()
    '            Dim data1 As Byte() = Nothing
    '            Try
    '                data1 = webClient.DownloadData(url)
    '                fgOK = True
    '            Catch ex As Exception
    '                exMessage = ex.Message
    '            End Try
    '            stopwatch.Stop()

    '            Dim elapsedMs As Long = stopwatch.ElapsedMilliseconds
    '            Dim fileSize As String = If(data1 IsNot Nothing, FormatFileSize(data1.Length), "0")
    '            If fgOK Then
    '                reportBuilder.AppendLine($"完成下載: {url}")
    '                reportBuilder.AppendLine($"檔案大小: {fileSize}")
    '                reportBuilder.AppendLine($"下載耗時: {elapsedMs} 毫秒")
    '                reportBuilder.AppendLine("--------------------------------------------------")
    '                TIMS.LOG.Debug($"完成下載: {url}{vbCrLf}檔案大小: {fileSize}{vbCrLf}下載耗時: {elapsedMs} 毫秒")
    '            Else
    '                reportBuilder.AppendLine($"下載失敗: {url}")
    '                reportBuilder.AppendLine($"錯誤訊息: {exMessage}")
    '                reportBuilder.AppendLine($"下載耗時: {elapsedMs} 毫秒")
    '                reportBuilder.AppendLine("--------------------------------------------------")
    '                TIMS.LOG.Error($"下載失敗: {url}{vbCrLf}錯誤訊息: {exMessage}{vbCrLf}下載耗時: {elapsedMs} 毫秒")
    '            End If
    '        Next
    '    End Using

    '    Return reportBuilder.ToString().Replace(vbCrLf, "<br />")
    'End Function

    ''' <summary>
    ''' 測試多個檔案的下載耗時並生成報告。
    ''' </summary>
    ''' <returns>包含下載測試結果的文字報告。</returns>
    Private Function TestFileDownloads2(fileUrls As String()) As String
        Dim reportBuilder As New StringBuilder()
        reportBuilder.AppendLine("檔案下載耗時測試報告")
        reportBuilder.AppendLine($"測試開始時間: {DateTime.Now.ToString()}")
        reportBuilder.AppendLine("--------------------------------------------------")

        'WebRequest物件如何忽略憑證問題
        'ServicePointManager.ServerCertificateValidationCallback = New RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 'HttpClient
        'System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12 ' 3072 'System.Net.SecurityProtocolType.Tls12

        Const SAVE_PATH As String = "~/SAVE_PATH/"
        Dim fullSavePath As String = Server.MapPath(SAVE_PATH)
        If Not Directory.Exists(fullSavePath) Then Directory.CreateDirectory(fullSavePath)
        Dim saveDirectory As String = fullSavePath

        Using webClient As New CustomWebClient()
            webClient.Timeout = 600000 ' 設定為 10 分鐘

            For Each url As String In fileUrls
                url = Replace(Replace(Replace(url, vbCrLf, ""), vbLf, ""), vbCr, "")
                If url = "" Then Continue For

                Dim exMessage As String = ""
                Dim fgOK As Boolean = False
                reportBuilder.AppendLine($"開始下載: {url}")
                Dim stopwatch As New Stopwatch()
                stopwatch.Start()
                Dim data1 As Byte() = Nothing
                Try
                    data1 = webClient.DownloadData(url)

                    fgOK = True
                Catch ex As Exception
                    exMessage = ex.Message
                End Try
                stopwatch.Stop()

                Dim elapsedMs As Long = stopwatch.ElapsedMilliseconds
                Dim fileSize As String = If(data1 IsNot Nothing, FormatFileSize(data1.Length), "0")

                Dim saveMessage As String = ""
                Dim filePath As String = ""
                Dim fileName As String = ""
                Try
                    If data1 IsNot Nothing AndAlso data1.Length > 0 Then
                        ' 取得檔案名稱
                        fileName = Path.GetFileName(url)
                        fileName = TIMS.GetValidFileName(HttpUtility.UrlDecode(fileName))
                        ' 組合完整的儲存檔案路徑
                        filePath = Path.Combine(saveDirectory, fileName)
                        ' 將下載的資料寫入檔案
                        File.WriteAllBytes(filePath, data1)
                    End If
                    saveMessage &= $",,WriteAllBytes! ,,{fileSize},,{saveDirectory},,{fileName}"
                Catch ex As Exception
                    saveMessage &= $",,WriteAllBytes: {ex.Message} ,,{fileSize},,{saveDirectory},,{fileName}"
                End Try

                If fgOK Then
                    reportBuilder.AppendLine($"完成下載: {url}")
                    reportBuilder.AppendLine($"檔案大小: {fileSize}")
                    reportBuilder.AppendLine($"下載耗時: {elapsedMs} 毫秒.{saveMessage}")
                    reportBuilder.AppendLine("---")
                    TIMS.LOG.Debug($"完成下載: {url}{vbCrLf}檔案大小: {fileSize}{vbCrLf}下載耗時: {elapsedMs} 毫秒.{saveMessage}")
                Else
                    reportBuilder.AppendLine($"下載失敗: {url}")
                    reportBuilder.AppendLine($"錯誤訊息: {exMessage}")
                    reportBuilder.AppendLine($"下載耗時: {elapsedMs} 毫秒.{saveMessage}")
                    reportBuilder.AppendLine("---")
                    TIMS.LOG.Error($"下載失敗: {url}{vbCrLf}錯誤訊息: {exMessage}{vbCrLf}下載耗時: {elapsedMs} 毫秒.{saveMessage}")
                End If
            Next
        End Using

        Return reportBuilder.ToString().Replace(vbCrLf, "<br />")
    End Function


    ''' <summary>
    ''' 格式化檔案大小，使其更易讀。
    ''' </summary>
    Private Function FormatFileSize(ByVal byteCount As Long) As String
        Dim suffixes As String() = {"B", "KB", "MB", "GB", "TB"}
        Dim i As Integer = 0
        Dim dblSByte As Double = byteCount

        If byteCount = 0 Then Return "0 B"

        While dblSByte >= 1024 AndAlso i < suffixes.Length - 1
            dblSByte /= 1024
            i += 1
        End While

        Return String.Format("{0:0.##} {1}", dblSByte, suffixes(i))
    End Function

End Class


