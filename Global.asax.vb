Imports System.IO

Public Class [Global]
    Inherits System.Web.HttpApplication

    ReadOnly LOG As ILog = log4net.LogManager.GetLogger("WDAIIP.SYS.Net")

    Sub Application_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 啟動應用程式時引發
        ' init log4net configuration
        log4net.Config.XmlConfigurator.Configure(
            New System.IO.FileInfo(Path.Combine(Server.MapPath("~/"), "log4net.config")))

        'Spire.License.LicenseProvider.SetLicenseFileName("license.elic.DOC.xml")
        'Spire.License.LicenseProvider.SetLicenseFileName("license.elic.XLS.xml")

        LOG.Info("##Application_Start")
        ' Routing Config
        RouteConfig.RegisterRoutes(Routing.RouteTable.Routes)

        Try
            Using oConn As SqlConnection = DbAccess.GetConnection()
                Call TIMS.Utl_SetConfigVAL(oConn)
            End Using
        Catch ex As Exception
            TIMS.WriteTraceLog($"#Application_Start:{ex.Message}", ex)
            Return
        End Try

        ' init REST Service Handler
        'PoC.Web.Services.ReSTServiceHandler.setIgnoreAppCode(True)
        'PoC.Web.Services.ReSTServiceHandler.setPackageNamePreix("TIMS.CORE")
        'PoC.Web.Services.ReSTServiceHandler.Initialize()
    End Sub

    Sub Session_Start(ByVal sender As Object, ByVal e As EventArgs)
        ' 啟動工作階段時引發
    End Sub

    Sub Utl_Response404()
        'Dim oRan As New Random
        'System.Threading.Thread.Sleep(oRan.Next(oRan.Next(1, 2) * 1000))
        System.Threading.Thread.Sleep(TIMS.GetNextRND(TIMS.GetNextRND(1, 2) * 1000))
        If (Response.StatusCode <> 404) Then Response.StatusCode = 404
        Response.StatusDescription = "Not Found"
        Response.End()
    End Sub

    Sub Utl_Response404(i_seconds As Integer)
        'Dim oRan As New Random
        'System.Threading.Thread.Sleep((i_seconds * 1000) + oRan.Next(oRan.Next(1, 2) * 1000))
        System.Threading.Thread.Sleep((i_seconds * 1000) + TIMS.GetNextRND(TIMS.GetNextRND(1, 2) * 1000))
        If (Response.StatusCode <> 404) Then Response.StatusCode = 404
        Response.StatusDescription = "Not Found"
        Response.End()
    End Sub

    Sub Get_HeaderValue_info()
        Dim s_loginfo As New StringBuilder
        s_loginfo.Append(String.Format("##Application_BeginRequest-Request.Headers.AllKeys:{0}", vbCrLf))
        Dim i As Integer = 1
        For Each s_ky As String In Request.Headers.AllKeys
            'If s_ky.Equals("X-Scan-Memo") Then 'End If
            s_loginfo.Append(String.Format("[{0}]={1}", s_ky, vbCrLf))
            Dim a_ky() As String = Request.Headers.GetValues(s_ky)
            For Each s_ky_val As String In a_ky
                'If s_ky.Equals("Engine") Then  End If
                s_loginfo.Append(String.Format("{0}> {1};{2}", i, s_ky_val, vbCrLf))
                i += 1
            Next
        Next
        LOG.Info(s_loginfo.ToString())
    End Sub

    Function Get_HeaderValue(ByVal s_hd_value As String) As String
        Dim rst As String = ""
        For Each s_ky As String In Request.Headers.AllKeys
            'If s_ky.Equals("X-Scan-Memo") Then 'End If
            If s_ky.Equals(s_hd_value, StringComparison.OrdinalIgnoreCase) Then
                Dim a_ky() As String = Request.Headers.GetValues(s_ky)
                For Each s_ky_val As String In a_ky
                    rst &= String.Concat(If(rst <> "", ";", ""), s_ky_val)
                Next
                Exit For
            End If
        Next
        Return rst
    End Function

    Function CHK_HK1() As Boolean
        Dim rst As Boolean = False
        Dim s_v1 As String = Get_HeaderValue("X-Scan-Memo")
        If s_v1.Length > 0 Then
            If s_v1.IndexOf("Engine=""Http+Request+Smuggling""", StringComparison.OrdinalIgnoreCase) > -1 Then
                LOG.Error("##Application_BeginRequest [CHK_HK1]:X-Scan-Memo:" & s_v1)
                rst = True
            Else
                LOG.Info("##Application_BeginRequest [CHK_HK1]:X-Scan-Memo:" & s_v1)
            End If
        End If
        Return rst
    End Function

    Function CHK_HK2() As Boolean
        Dim rst As Boolean = False
        Dim vWebFmHost1 As String = TIMS.Utl_GetConfigSet("WebFmHost1")
        If (vWebFmHost1 IsNot Nothing AndAlso vWebFmHost1.Length > 0) Then
            Dim s_v1 As String = Get_HeaderValue("Host")
            If s_v1.Length > 0 Then
                If vWebFmHost1.IndexOf(s_v1, StringComparison.OrdinalIgnoreCase) = -1 Then
                    LOG.Error("##Application_BeginRequest [CHK_HK2]:Host:" & s_v1)
                    rst = True
                End If
            End If
        End If
        Return rst
    End Function

    Sub Application_BeginRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 於每一個要求開始時引發
        'Dim url As String = Request.Url.AbsolutePath
        'LOG.Info("##Application_BeginRequest")
        'LOG.Info(String.Format("##Application_BeginRequest ; url={0}", url))
        'If (url.Contains(".aspx")) Then
        '    'add CSP header
        '    Dim s_csp As String = "default-src 'self'; frame-ancestors 'self'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; font-src 'self' fonts.gstatic.com; frame-src 'self'; img-src 'self';object-src 'self';"
        '    Response.AppendHeader("Content-Security-Policy", s_csp)
        'End If
        'Get_HeaderValue_info()
        'Dim s_v1 As String = Get_HeaderValue("X-Scan-Memo")
        'LOG.Error("##Application_BeginRequest-X-Scan-Memo:" & s_v1)
        'Dim s_v1 As String = Get_HeaderValue("COOKIE")
        'LOG.Error("##Application_BeginRequest-COOKIE:" & s_v1)

        '(取得使用者正確ip)
        Dim v_IpAddress As String = Common.GetIpAddress() 'MyPage.Request.UserHostAddress

        Dim flag_hk1 As Boolean = CHK_HK1()
        If flag_hk1 Then
            If Not String.IsNullOrEmpty(v_IpAddress) Then
                '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
                Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(v_IpAddress)
                If flag_ChkHISTORY1 Then
                    Dim i_ChkHISTORY1_cnt As Integer = TIMS.SUtl_ChkHISTORY1_CNT(v_IpAddress)
                    If i_ChkHISTORY1_cnt > 2 Then
                        Utl_Response404(i_ChkHISTORY1_cnt)
                        Return
                    Else
                        Utl_Response404()
                        Return
                    End If
                End If
            Else
                Utl_Response404()
                Return
            End If
            Return
        End If

        Dim flag_hk2 As Boolean = CHK_HK2()
        If flag_hk2 Then
            If Not String.IsNullOrEmpty(v_IpAddress) Then
                '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
                Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(v_IpAddress)
                If flag_ChkHISTORY1 Then
                    Dim i_ChkHISTORY1_cnt As Integer = TIMS.SUtl_ChkHISTORY1_CNT(v_IpAddress)
                    If i_ChkHISTORY1_cnt > 2 Then
                        Utl_Response404(i_ChkHISTORY1_cnt)
                        Return
                    Else
                        Utl_Response404()
                        Return
                    End If
                End If
            Else
                Utl_Response404()
                Return
            End If
            Return
        End If

        'Dim iMaxCanMailCount As Integer = TIMS.cst_iMaxCanMailCount '(寄狀況信最大容忍數)
        'Dim iGlobalMailCount As Integer = TIMS.GlobalMailCount '目前寄信總數量
        If TIMS.GlobalMailCount >= TIMS.cst_iMaxCanMailCount Then
            If Not String.IsNullOrEmpty(v_IpAddress) Then
                '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
                Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(v_IpAddress)
                If flag_ChkHISTORY1 Then
                    Dim i_ChkHISTORY1_cnt As Integer = TIMS.SUtl_ChkHISTORY1_CNT(v_IpAddress)
                    If i_ChkHISTORY1_cnt > 2 Then
                        Utl_Response404(i_ChkHISTORY1_cnt)
                        Return
                    Else
                        Utl_Response404()
                        Return
                    End If
                    Return
                End If
            End If
        End If

        'If (Request.HttpMethod <> "GET" OrElse Request.ContentLength > 0) Then
        '    Dim s_loginfo As New StringBuilder
        '    s_loginfo.Append("##Application_BeginRequest-Request.HttpMethod != GET")
        '    s_loginfo.Append(String.Format(" ,Request.Url.AbsolutePath {0}", url))
        '    s_loginfo.Append(String.Format(" ,Request.HttpMethod {0}", Request.HttpMethod))
        '    s_loginfo.Append(String.Format(" ,Request.ContentLength {0}", Request.ContentLength))
        '    LOG.Debug(s_loginfo.ToString())
        '    Dim flag_NG As Boolean = True
        '    If (flag_NG) Then Utl_Response404()
        '    Return
        'End If

    End Sub

    Sub Application_AuthenticateRequest(ByVal sender As Object, ByVal e As EventArgs)
        ' 嘗試驗證使用時引發
    End Sub

    '錯誤信件內容
    Function GetErrorMsg(sm As SessionModel) As String
        '取得錯誤訊息
        'Dim sm As SessionModel = SessionModel.Instance()
        'sm查無有效資訊1/2/3
        Dim flag_sm_is_nothing As Boolean = (sm Is Nothing OrElse sm.UserInfo Is Nothing OrElse Not sm.IsLogin)
        Dim LastError As Exception = Server.GetLastError
        'lerr查無資訊1
        Dim flag_LastError_is_nothing As Boolean = (LastError Is Nothing)

        Dim sMailBodyB As New StringBuilder
        If flag_LastError_is_nothing AndAlso flag_sm_is_nothing Then Return "" '查無有效資訊-不須回傳
        Try
            '寫入新的錯誤訊息
            sMailBodyB.Append(String.Concat("時間：", Now, vbCrLf)) '存入時間
            sMailBodyB.Append(String.Concat("Request.UserAgent：", Me.Request.UserAgent, vbCrLf))
            sMailBodyB.Append(String.Concat("IpAddress：", Common.GetIpAddress(), vbCrLf))
            sMailBodyB.Append(String.Concat("UserHostAddress：", Me.Request.UserHostAddress, vbCrLf))
            sMailBodyB.Append(String.Concat("UserHostName：", Me.Request.UserHostName, vbCrLf))
            sMailBodyB.Append(String.Concat("MachineName：", Server.MachineName, vbCrLf))
            sMailBodyB.Append(String.Concat("Url：", Context.Request.Url, vbCrLf))
            sMailBodyB.Append(String.Concat("RawUrl：", Context.Request.RawUrl, vbCrLf))
            If LastError IsNot Nothing Then
                sMailBodyB.AppendLine(String.Concat("GetType： ", LastError.GetType())) '存入錯誤訊息1
                sMailBodyB.AppendLine(String.Concat("Messager： ", LastError.Message)) '存入錯誤訊息2
                sMailBodyB.AppendLine(String.Concat("ToStr： ", LastError.ToString())) '存入錯誤訊息3
                sMailBodyB.AppendLine(String.Concat("StackTrace： ", LastError.StackTrace)) '存入錯誤訊息4
            End If
            If Not flag_sm_is_nothing Then sMailBodyB.AppendLine(GetErrorMsgSys(sm))
        Catch ex As Exception
            LOG.Warn(ex.Message, ex)
        End Try
        'sMailBody 為空離開
        If sMailBodyB.ToString() = "" Then Return ""

        '無需通知的關鍵字 使用逗號分隔
        Dim s_NOMAIL As String = TIMS.Utl_GetConfigSet("NOMAILStr")
        If $"{s_NOMAIL}" <> "" Then
            Dim NOMAIL_ARRAY As String() = s_NOMAIL.Split(",")
            If Not flag_LastError_is_nothing Then
                '有錯誤資訊
                For Each s_VAL As String In NOMAIL_ARRAY
                    If s_VAL <> "" AndAlso $"{sMailBodyB}".Contains(s_VAL) Then Return "" '該錯誤為網路常態性錯誤狀況-不須回傳 
                Next
            End If
        End If
        Return $"{sMailBodyB}"
    End Function

    '取得錯誤訊息 系統資訊
    Function GetErrorMsgSys(ByRef sm As SessionModel) As String
        'Dim sm As SessionModel = SessionModel.Instance()
        Dim pText1 As String = ""
        If (sm Is Nothing OrElse sm.UserInfo Is Nothing) Then Return pText1
        Try
            pText1 &= $"UserID={sm.UserInfo.UserID}{vbCrLf}"
            pText1 &= $"RoleID={sm.UserInfo.RoleID}{vbCrLf}" '  Session("RoleID").ToString}{vbCrLf}"'角色代碼
            pText1 &= $"LID={sm.UserInfo.LID}{vbCrLf}" 'Session("LID").ToString}{vbCrLf}"      '階層代碼
            pText1 &= $"OrgID={sm.UserInfo.OrgID}{vbCrLf}" 'Session("OrgID").ToString & vbCrLf
            pText1 &= $"OrgName={sm.UserInfo.OrgName}{vbCrLf}" 'Session("OrgName").ToString & vbCrLf
            pText1 &= $"RID={sm.UserInfo.RID}{vbCrLf}" 'Session("RID").ToString & vbCrLf
            pText1 &= $"OrgLevel={sm.UserInfo.OrgLevel}{vbCrLf}" 'Session("OrgLevel").ToString & vbCrLf
            pText1 &= $"DistID={sm.UserInfo.DistID}{vbCrLf}" 'Session("DistID").ToString & vbCrLf
            pText1 &= $"PlanID={sm.UserInfo.PlanID}{vbCrLf}" 'Session("PlanID").ToString & vbCrLf
            pText1 &= $"TPlanID={sm.UserInfo.TPlanID}{vbCrLf}" 'Session("TPlanID").ToString & vbCrLf
            pText1 &= $"Years={sm.UserInfo.Years}{vbCrLf}" 'Session("Years").ToString & vbCrLf
            pText1 &= vbCrLf
            pText1 = TIMS.EncryptAes(pText1)
        Catch ex As Exception
            LOG.Warn(ex.Message, ex)
        End Try
        Return pText1
    End Function

    '發生錯誤時引發
    Sub Application_Error(ByVal sender As Object, ByVal e As EventArgs)
        '例外狀況儲存變數
        Dim lastError As Exception = Server.GetLastError()
        If lastError Is Nothing Then Return '(異常離開)
        Dim s_Uri As String = TryCast(Request.Url.AbsoluteUri, String)
        LOG.Error($"##Application_Error: {s_Uri}: {lastError.Message}", lastError)

        Dim oWDAP As New WDAP
        '寄送錯誤信件
        Dim sGlobalErrorSendEmail As String = TIMS.Utl_GetConfigSet("GlobalErrorSendEmail") '錯誤信寄送 N:停止
        Dim sMailBody As String = ""
        Try
            '檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Nothing, Nothing)
            If sGlobalErrorSendEmail <> "N" Then
                Dim sm As SessionModel = SessionModel.Instance()
                sMailBody = GetErrorMsg(sm)
                If sMailBody <> "" Then Call TIMS.SendMailTest(sMailBody)
            End If
        Catch ex As Exception
            LOG.Error(ex.Message, ex)
        End Try
        sMailBody = String.Empty

        '(清除先前的例外狀況)
        Server.ClearError()

        'If Not (HttpContext.Current Is Nothing) Then Session("LastException") = lastError
        Try
            If (HttpContext.Current.Session IsNot Nothing) Then
                HttpContext.Current.Session.RemoveAll() '(清除所有SESSION)
                Session.RemoveAll() '(清除所有SESSION)
                Session("LastException") = lastError '(記載錯誤訊息)
            End If
        Catch ex As Exception
            LOG.Error(ex.Message, ex)
        End Try
        'If Not (HttpContext.Current Is Nothing) Then
        '    Dim context As HttpContext = HttpContext.Current
        '    context.Session("LastException") = lastError
        'End If
        '直接跳錯誤頁面
        If Not Request.IsLocal Then Response.Redirect("~/AppError")

        '顯示錯誤狀況協助開發
        Call ShowErrorMsg1(lastError)
    End Sub

    ''' <summary>'顯示錯誤狀況 (測試環境或偵錯狀況)</summary>
    ''' <param name="lastError"></param>
    Sub ShowErrorMsg1(ByRef lastError As Exception)
        Dim fg_test As Boolean = TIMS.sUtl_ChkTest() '測試
        Dim fg_InfoShow As Boolean = (fg_test OrElse TIMS.Utl_GetConfigSet("DEBUGINFO").Equals("Y"))

        Dim strStackTrace As String = ""
        strStackTrace &= "GetLastError()：<br/>"
        strStackTrace &= lastError.ToString().Replace(vbLf, "<br/>")
        'strStackTrace &= "<br/>StackTrace()：<br/>"
        'strStackTrace &= lastError.StackTrace.Replace(vbLf, "<br/>")
        While lastError.InnerException IsNot Nothing
            lastError = lastError.InnerException
            If lastError.StackTrace Is Nothing Then Exit While
            strStackTrace &= "<br/>InnerException()：<br/>"
            strStackTrace &= lastError.StackTrace.Replace(vbLf, "<br/>")
        End While

        If fg_InfoShow Then
            'DEBUGINFO 'Response.Flush()
            Dim js_str As String = "<script type=""text/javascript"" src=""/Scripts/jquery-3.7.1.min.js""></script>"
            js_str &= String.Concat("<script>", "if (document.body) { window.scroll(0, document.body.scrollHeight); }")
            js_str &= String.Concat("if (parent && parent.setMainFrameHeight() != undefined) { parent.setMainFrameHeight(); }", "</script>")
            Response.Write(String.Concat(js_str & strStackTrace))
            'Response.End()
        End If
    End Sub

    Sub Session_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 於工作階段結束時引發
    End Sub

    Sub Application_End(ByVal sender As Object, ByVal e As EventArgs)
        ' 於應用程式結束時引發
    End Sub

End Class
