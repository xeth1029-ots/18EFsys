Imports System.Diagnostics
Imports System.Threading.Tasks
Imports System.Web
Imports System.Web.Services
Imports System.Net.Http
Imports System.Threading

Public Class SolrAPI
    Implements System.Web.IHttpHandler

    Public Shared LOG As ILog = LogManager.GetLogger("SolrAPI") 'log4net

    Public Shared Function Get_SolrApi(s_q1 As String, keyword As String) As String
        Const def_WebApi_URL1 As String = "https://job.taiwanjobs.gov.tw/SolrApi.aspx"
        Dim s_fntn_n As String = "*Get_SolrApi"
        If (s_q1 <> "") Then LOG.Info(String.Format(s_fntn_n + " ,q1={0}", s_q1))
        If (keyword <> "") Then LOG.Info(String.Format(s_fntn_n + " ,keyword={0}", keyword))

        'HttpClient
        'WebRequest物件如何忽略憑證問題
        System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf TIMS.ValidateServerCertificate)
        'TLS 1.2-基礎連接已關閉: 傳送時發生未預期的錯誤 
        System.Net.ServicePointManager.SecurityProtocol = Net.SecurityProtocolType.Tls12 '3072

        Using client As New HttpClient()
            'Dim sw As New Stopwatch() 'sw.Start()

            ' 從 App.config 或其他配置檔案讀取 Web API URL
            Dim WebApi_SolrApi_URL1 As String = ConfigurationManager.AppSettings("SolrApi_URL1")
            If (String.IsNullOrEmpty(WebApi_SolrApi_URL1)) Then WebApi_SolrApi_URL1 = def_WebApi_URL1

            'Dim s_q1 As String=""
            If s_q1 = "" Then
                Dim tfS As String = "2"
                Dim tfE As String = "999999"
                Dim s_cat As String = "APP_TEXT APP_TEXT_OTHER APP_TEXT_OTHER_SEEK"
                s_q1 = String.Concat("key:(", keyword, "* *", keyword, ") AND cat:(", s_cat, ") NOT artificial:1")
                s_q1 &= String.Concat(" AND tf:[", tfS, " TO ", tfE, "]")
            End If

            Dim urlWithParams As String = String.Format("{0}?q={1}", WebApi_SolrApi_URL1, s_q1)
            LOG.Info(String.Concat(s_fntn_n, " ,urlWithParams: ", WebApi_SolrApi_URL1, " , ", urlWithParams))

            '使用 GET 請求
            Dim response As HttpResponseMessage = client.GetAsync(urlWithParams).GetAwaiter().GetResult()

            Dim strResponse As String = ""
            If response.IsSuccessStatusCode Then
                strResponse = response.Content.ReadAsStringAsync().GetAwaiter().GetResult()
                Return strResponse 'JsonConvert.DeserializeObject(Of SearchResponse)(strResponse)
            End If

            ' 處理非成功狀態碼
            LOG.Error($"HTTP request failed with status code: {response.StatusCode}")
            Return Nothing
            'sw.Stop() 'LOG.Debug(String.Format(s_fntn_n + " sw= {0} tms", sw.Elapsed.TotalMilliseconds.ToString()))
        End Using
    End Function

    Sub ProcessRequest(ByVal context As HttpContext) Implements IHttpHandler.ProcessRequest

        'context.Response.ContentType = "text/plain"
        'context.Response.Write("Hello World!")
        'var keyword = $('#' + schInput1).val();
        '    If (keyword.length < 2) Then { Return; }
        '    //var urlxStrSegment = 'https://job.taiwanjobs.gov.tw/StrSegment.ashx?str=' + keyword;
        '    //var urlxStrSegment = 'http://192.168.0.56:8531/StrSegment.ashx?str=' + keyword;

        '    var cat = "APP_TEXTAPP_TEXT_OTHERAPP_TEXT_OTHER_SEEK";
        '    var q = "key:(" + keyword + "* *" + keyword + ") AND cat:" + cat;
        '    // 20200316, eric, 去除已標記刪除的項目
        '    q += " NOT artificial:1";
        '    var urlxSolrApi = 'http://192.168.0.56:8531/SolrApi.aspx?q=' + q;
        '    $.ajax({
        '        url: urlxSolrApi,
        '        Type 'GET',
        '        success: Function (data) {
        '            $('#' + result2).html(data);
        '        },
        '        Error: Function (error) {
        '            Console.Error('發生錯誤：', error);
        '            $('#' + result2).html('搜尋失敗，請稍後再試。');
        '        }
        '    });
        '}

        context.Response.ContentType = "application/json"

        Dim UTF8bytes As Byte()

        Dim result As New AjaxResultStruct

        Dim keyword As String = TIMS.ClearSQM(context.Request("keyword"))
        Dim s_q1 As String = TIMS.ClearSQM(context.Request("q"))
        Dim check_parms_ok As Boolean = ((keyword <> "") OrElse (s_q1 <> ""))
        If Not check_parms_ok Then
            TIMS.LOG.Warn(String.Format("#SolrAPI check parms is Empty ,{0},{1} ", "keyword", keyword))
            TIMS.LOG.Warn(String.Format("#SolrAPI check parms is Empty ,{0},{1} ", "q", s_q1))
            result.status = False
            result.message = "fail"
            result.data = "check parms is Empty"
            UTF8bytes = Encoding.UTF8.GetBytes(result.Serialize())
            context.Response.AddHeader("Content-Length", UTF8bytes.Length.ToString())
            context.Response.BinaryWrite(UTF8bytes)
            context.Response.Flush()
            'context.Response.StatusCode = 401
            context.Response.End()
            Return
        End If

        'Dim s_data As String = ""
        Dim objLock_SolrAPI As New Object
        SyncLock objLock_SolrAPI
            Dim fg_error As Boolean = False
            Try
                If ((s_q1 <> "")) Then
                    UTF8bytes = Encoding.UTF8.GetBytes(Get_SolrApi(s_q1, ""))
                ElseIf ((keyword <> "")) Then
                    UTF8bytes = Encoding.UTF8.GetBytes(Get_SolrApi("", keyword))
                Else
                    context.Response.StatusCode = 401
                    context.Response.End()
                    Return
                End If
                'UTF8bytes = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(JsonConvert.DeserializeObject(Of SearchResponse)(s_data)))
                'UTF8bytes = Encoding.UTF8.GetBytes(s_data)
                context.Response.AddHeader("Content-Length", UTF8bytes.Length.ToString())
                context.Response.BinaryWrite(UTF8bytes)
                context.Response.Flush()
                'context.Response.StatusCode = 401
                context.Response.End()
                Return
            Catch threx As ThreadAbortException
                '執行緒已經中止
                'TIMS.LOG.Error(String.Concat("#ThreadAbortException! ", threx.Message), threx)
                fg_error = True
            Catch httpex As HttpException
                '遠端主機已關閉連接。錯誤碼為 0x800704CD  要求已經逾時。
                TIMS.LOG.Error(String.Concat("#HttpException! ", httpex.Message), httpex)
                fg_error = True
            Catch ex As Exception
                '執行緒已經中止 / 遠端主機已關閉連接
                TIMS.LOG.Error(String.Concat("#Exception! ", ex.Message), ex)
                'fg_error = True
            End Try
            '(無法調適的錯誤直接離開)
            If fg_error Then Return

            context.Response.StatusCode = 401
            context.Response.End()
            Return
        End SyncLock

        'result.status = True
        'result.message = "ok"
        'result.data = s_data
        'UTF8bytes = Encoding.UTF8.GetBytes(result.Serialize())
        'context.Response.AddHeader("Content-Length", UTF8bytes.Length.ToString())
        'context.Response.BinaryWrite(UTF8bytes)
        'context.Response.Flush()
        'context.Response.End()
    End Sub

    ReadOnly Property IsReusable() As Boolean Implements IHttpHandler.IsReusable
        Get
            Return False
        End Get
    End Property

End Class