Imports System.Net
'Imports System,Imports System.Runtime.InteropServices ' 需要匯入這個命名空間,
Public Class Login
    Inherits System.Web.UI.Page

    Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Dim sTxtIdUnchange As String = "" 'TIMS.ClearSQM(txtUserId.Text)
    Dim sTxtPass As String = "" 'TIMS.ClearSQM(txtUserPass.Text)
    Dim sVCode As String = "" 'TIMS.ClearSQM(txtVCode.Text)

    Const cst_alert_msg_1 As String = "請輸入帳號/密碼"
    Const cst_alert_msg_1b As String = "帳號或密碼錯誤!"
    Const cst_alert_msg_1bx As String = "系統一律以自然人憑證登入，請按「自然人憑證登入」鈕進行系統登入" '"帳號或密碼錯誤!!!"
    Const cst_alert_msg_1c As String = "該帳號尚未啟用!!!"
    Const cst_alert_msg_1d As String = "帳號已被停用,請洽詢系統管理者!!!"
    Const cst_alert_msg_1f As String = "帳號沒有系統使用權限,請洽詢系統管理者!""帳號已被停用,請洽詢系統管理者!!"
    Const cst_alert_msg_1g As String = "查無此帳號!"
    Const cst_alert_msg_2 As String = "請輸入驗證碼"
    Const cst_alert_msg_2b As String = "驗證碼不正確!!!"
    Const cst_alert_msg_2c As String = "請輸入帳號與驗證碼"
    Const cst_alert_msg_3 As String = "此帳號未設定EMAIL!"
    '帳號三個月未使用，系統將自動停用。 [判斷是否超過三個月未登入]
    '帳號一年未使用，系統將自動清除自然人憑證序號。

    Dim g_parms As Hashtable
    'Public Const cst_ErrorMsg2 As String = "請勿嘗試在頁面輸入具有危險性的字元!"

    Public BaseUrl As String
    Public sm As SessionModel '= SessionModel.Instance()

    Dim oWDAP As New WDAP
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'If Session(TIMS.cst_MOICA_Login) Is Nothing Then Return
        'If Not Session(TIMS.cst_MOICA_Login).Equals("xxx") Then Return
        If Session(TIMS.cst_MOICA_Login) Is Nothing Then Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        'atest
        bt_atest.Visible = False
        'atest button 是否顯示
        Call At_test() 'atest
        '資安檢核
        Call Critical_Issues_1()
        '網站停止功能2(只能使用在後台停止-login)
        Call TIMS.Stop2(Me, objconn)

        If Not TIMS.OpenDbConn(Me, objconn) Then Return ' Exit Sub

        sm = SessionModel.Instance()

        BaseUrl = ResolveUrl("~/")

        'If Not String.IsNullOrEmpty(BaseUrl) AndAlso Not BaseUrl.EndsWith("/") Then BaseUrl = String.Concat(BaseUrl, "/")

        '測試程式碼
        Call Test1_prg()

        Hidversion1.Value = TIMS.ClearSQM(oWDAP.FileLastModified(Me, 1))
        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)
        If Not IsPostBack Then
            ' Get 連結到 Login 視為登出/重登
            AuthUtil.LogoutLog()
            ' 清除登入狀態
            sm.ClearSession()
        End If

    End Sub

    ''' <summary> 資安檢核 </summary>
    Sub Critical_Issues_1()
        'openhttps.Value = TIMS.ClearSQM(openhttps.Value)  'g_parms.Clear()
        If (g_parms IsNot Nothing) AndAlso g_parms.Count > 0 Then g_parms.Clear()
        g_parms = New Hashtable From {
            {"1.txtUserId.Text", txtUserId.Text},
            {"2.txtUserPass.Text", txtUserPass.Text},
            {"3.txtVCode.Text", txtVCode.Text}
        }

        sTxtIdUnchange = TIMS.ClearSQM(txtUserId.Text)
        sTxtPass = TIMS.ClearSQM(txtUserPass.Text)
        sVCode = TIMS.ClearSQM(txtVCode.Text)
        txtUserId.Text = sTxtIdUnchange
        txtUserPass.Text = sTxtPass
        txtVCode.Text = sVCode

        g_parms.Add("1.sTxtIdUnchange", sTxtIdUnchange)
        g_parms.Add("2.sTxtPass", sTxtPass)
        g_parms.Add("3.sVCode", sVCode)
    End Sub

    ''' <summary> 測試程式碼 </summary>
    Sub Test1_prg()
        Const cst_cc_sRtnUrl As String = "/" 'TIMS.Utl_GetConfigSet("RtnUrl")'不可接受/
        Dim rqReturnUrl As String = Request("ReturnUrl") ': rqMsgid = TIMS.ClearSQM(rqMsgid)
        If rqReturnUrl <> "" AndAlso cst_cc_sRtnUrl.IndexOf(rqReturnUrl) = -1 Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 'Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            TIMS.sUtl_404NOTFOUND(Me, objconn)
            Return
        End If

        '測試程式碼
        Dim test1 As String = TIMS.ClearSQM(Request("test1"))
        Select Case UCase(test1)
            Case UCase("test1")
                'ERROR 測試
                Dim i As Integer = 0
                i = ""
            Case UCase("EMAIL1")
                '(本機)Subject-測試信-EMAIL測試
                Dim path3 As String = TIMS.Utl_GetConfigSet("from_emailaddress")
                Dim vpath3 As String = If(String.IsNullOrEmpty(path3), TIMS.Cst_SendMail3_from_emailaddress, path3)
                Dim htSS As New Hashtable From {
                    {"MachineName", HttpContext.Current.Server.MachineName},
                    {"TPlanID", If(sm.UserInfo IsNot Nothing, Convert.ToString(sm.UserInfo.TPlanID), TIMS.Cst_TPlanID28)},
                    {"Stud_Name", "STUD_NAME"},
                    {"Subject", "Subject-測試信"},
                    {"ExamNo", "ExamNo"},
                    {"RelEnterDate", "RelEnterDate"},
                    {"ExamDate", "ExamDate"},
                    {"CheckInDate", "CheckInDate"},
                    {"EComment", "EComment"},
                    {"Email", TIMS.Cst_EmailtoMe},
                    {"from_emailaddress", vpath3},
                    {"signUpMemo", ""},
                    {"sRIDOrgName", ""},
                    {"sType", TIMS.Cst_SendMail3_CheckedOK}
                } 'htSS Hashtable() 
                Dim mail_msg As String = TIMS.SendMail3(htSS)
                'If mail_msg <> "" Then Common.RespWrite(Me, "<script>alert('" & mail_msg & "');</script>")
                Dim str_MSG1 As String = "Email OK!"
                If mail_msg <> "" Then
                    str_MSG1 = $"Email ERROR!-{mail_msg}"
                    sm.LastErrorMessage = Common.GetJsString(str_MSG1)
                    Return
                End If
                sm.LastResultMessage = Common.GetJsString(str_MSG1)
                Return
            Case UCase("EMAIL2")
                'EMAIL測試 [WDAIIP錯誤狀況提醒]'寄送錯誤信件
                Dim sGlobalErrorSendEmail As String = TIMS.Utl_GetConfigSet("GlobalErrorSendEmail") '錯誤信寄送 N:停止
                Dim sMailBody As String = ""
                Try
                    '檢核'防止駭客攻擊(紀錄) Call oWDAP.sUtl_SaveLoginData1(Nothing, Nothing)
                    If sGlobalErrorSendEmail <> "N" Then
                        'Dim sm As SessionModel = SessionModel.Instance()
                        sMailBody = TIMS.GetErrorMsg()
                        Call TIMS.SendMailTest(sMailBody)
                        sm.LastResultMessage = Common.GetJsString($"EMAIL:2: Email OK! length: {sMailBody.Length}")
                    Else
                        sm.LastErrorMessage = Common.GetJsString("EMAIL:2: sGlobalErrorSendEmail:`N`")
                    End If
                Catch ex As Exception
                    TIMS.LOG.Error(ex.Message, ex)
                    sm.LastErrorMessage = Common.GetJsString($"EMAIL:2: Exception: {ex.Message}")
                End Try
                'sMailBody = String.Empty
            Case UCase("test2")
                'GetHTTP_HOST
                Dim strErrmsg1 As String = $"GetHTTP_HOST: {vbCrLf}{TIMS.GetHTTP_HOST(Me)}{vbCrLf}"
                TIMS.WriteTraceLog(strErrmsg1)
                sm.LastResultMessage = Common.GetJsString($"GetHTTP_HOST! length: {strErrmsg1.Length}")
            Case UCase("test3")
                '防止駭客攻擊(紀錄) --錯誤才記
                Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
                sm.LastResultMessage = Common.GetJsString("防止駭客攻擊(紀錄)+1")
            Case UCase("test33")
                '防止駭客攻擊(清理紀錄)
                Call oWDAP.SUTL_SYS_HISTORY1_RST(Me, objconn)
                sm.LastResultMessage = Common.GetJsString("防止駭客攻擊(清理紀錄)+1")
            Case UCase("test4")
                '查詢檔案-最後異動日期
                Dim str_MSG1 As String = oWDAP.FileLastModified(Me, 2)
                sm.LastResultMessage = Common.GetJsString(str_MSG1)
            Case UCase("test5")
                '(勞保勾稽)試著隨意勾稽一筆資料
                Dim s_ERRMSG As String = ""
                Dim strCNAME As String = ""
                Dim strBIRTH As String = ""
                Dim strIDNO As String = GET_NEW_IDNO(strCNAME, strBIRTH)
                Dim str_MSG1 As String = TIMS.GET_BLI_WSV4(Me, objconn, strIDNO, strBIRTH, strCNAME, s_ERRMSG)
                If s_ERRMSG <> "" Then
                    sm.LastErrorMessage = Common.GetJsString(s_ERRMSG)
                    Return
                End If
                Dim sMemo As String = $"(任意勾稽1筆)&ACT=勞保明細查詢,IDNO={strIDNO},BIRTH={strBIRTH},CNAME={strCNAME}{vbCrLf}{str_MSG1}"
                sm.LastResultMessage = Common.GetJsString(sMemo)
                Return
            Case "TEST6"
                '查詢目前主機的作業系統 (OS) 版本
                Dim str_MSG1 As String = oWDAP.GetEnvOsVersion()
                sm.LastResultMessage = Common.GetJsString(str_MSG1)
            Case UCase("srvip")
                '取得伺服器端ip
                Dim s_MachineName As String = Environment.MachineName
                Dim str_IPXS As String = ""
                Dim iRow As Integer = 0
                For Each xIP As IPAddress In System.Net.Dns.GetHostAddresses(s_MachineName)
                    iRow += 1
                    str_IPXS &= $"{If(str_IPXS <> "", ",", "")}IP {iRow}: {xIP}{vbTab}"
                Next
                Dim str_MSG1 As String = $"MachineName: {s_MachineName}{vbCrLf}IPADDR: {str_IPXS}{vbCrLf}"
                sm.LastResultMessage = Common.GetJsString(str_MSG1)
            Case UCase("myip")
                '取得客戶端ip
                Dim str_MSG1 As String = $"IpAddress: {Common.GetIpAddress()}{vbCrLf}"  'MyPage.Request.UserHostAddress v_IpAddress '
                sm.LastResultMessage = Common.GetJsString(str_MSG1)
            Case UCase("DELADP1")
                '刪除無用暫存檔
                Dim i_DEL As Integer = TIMS.DEL_ADP_ZIPFILE(objconn)
                Dim str_MSG1 As String = $",, -1:有誤, 0:查無資料. {vbCrLf}"
                str_MSG1 &= $"(刪除無用暫存檔)DEL_ADP_ZIPFILE: {i_DEL}.{vbCrLf}"
                str_MSG1 &= $"(目前寄信日期)GlobalMailDate: {TIMS.GET_GlobalMailDate_T()}.{vbCrLf}"
                str_MSG1 &= $"(目前寄信總數量)GlobalMailCount: { TIMS.GlobalMailCount}.{vbCrLf}"
                sm.LastResultMessage = Common.GetJsString(str_MSG1)
            Case UCase("RPT1")
                '查詢指定目錄，產生文字報告
                Try
                    Dim str_MSG1 As String = TIMS.GenerateFileReportByDate()
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                    'txtReport.Text = report
                Catch ex As Exception
                    Dim str_MSG1 As String = $"發生錯誤：{ex.Message}"
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                End Try
            Case UCase("RPT2"), UCase("RPT22")
                Dim iRPT As Integer = 2
                If UCase(test1) = UCase("RPT22") Then iRPT = 22
                '查詢指定目錄，產生文字報告
                Try
                    Dim str_MSG1 As String = TIMS.GenerateFileReportX(iRPT)
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                    'txtReport.Text = report
                Catch ex As Exception
                    Dim str_MSG1 As String = $"發生錯誤：{ex.Message}"
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                End Try
            Case UCase("RPT3")
                '查詢指定目錄，產生文字報告
                Try
                    Dim str_MSG1 As String = TIMS.GenerateFileReportX(3)
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                    'txtReport.Text = report
                Catch ex As Exception
                    Dim str_MSG1 As String = $"發生錯誤：{ex.Message}"
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                End Try
            Case UCase("RPT4"), UCase("RPT44")
                Dim iRPT As Integer = 4
                If UCase(test1) = UCase("RPT44") Then iRPT = 44
                '查詢指定目錄，產生文字報告
                Try
                    Dim str_MSG1 As String = TIMS.GenerateFileReportX(iRPT)
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                    'txtReport.Text = report
                Catch ex As Exception
                    Dim str_MSG1 As String = $"發生錯誤：{ex.Message}"
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                End Try
            Case "MEM1"
                '查詢本機系統的記憶體狀態（包括總量、已使用與空閒空間）
                Try
                    Dim str_MSG1 As String = oWDAP.GetMemoryStatus()
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                    'txtReport.Text = report
                Catch ex As Exception
                    Dim str_MSG1 As String = $"發生錯誤：{ex.Message}"
                    sm.LastResultMessage = Common.GetJsString(str_MSG1)
                End Try
            Case "SBT1" 'UCase("SBT1")
                '系統開機時間
                Dim str_MSG1 As String = TIMS.SystemBootTime()
                sm.LastResultMessage = Common.GetJsString(str_MSG1)

            Case "UPDATE_MAX_ESERNUM_SEQ2" '將序號與資料表最大值同步
                DbAccess.Open(objconn)
                Dim SQLYP As String = "SELECT ISNULL(MAX(ESERNUM), 0) FROM dbo.STUD_ENTERTYPE2 WITH(NOLOCK) WHERE LEN(ESERNUM)<8"
                Dim SQLQ2 As String = "SELECT ISNULL(MAX(CURVAL), 0) FROM dbo.SYS_AUTONUM WITH(NOLOCK) WHERE TABLENAME='STUD_ENTERTYPE2_ESERNUM_SEQ2'"

                Dim vYP As Object = (New SqlCommand(SQLYP, objconn).ExecuteScalar())
                Dim vSEQ2 As Object = (New SqlCommand(SQLQ2, objconn).ExecuteScalar())
                Dim s_MSG1 As String = $"將序號與資料表最大值同步{vbCrLf}"
                s_MSG1 &= $"修正前：MAX,YP: {vYP} , SEQ2: {vSEQ2}{vbCrLf}"

                Dim fg_OK As Boolean = TIMS.UPDATE_MAX_ESERNUM_SEQ2(objconn)
                vYP = (New SqlCommand(SQLYP, objconn).ExecuteScalar())
                vSEQ2 = (New SqlCommand(SQLQ2, objconn).ExecuteScalar())
                s_MSG1 &= $"修正後：MAX,YP: {vYP} , SEQ2: {vSEQ2}{vbCrLf} , {fg_OK}"
                sm.LastResultMessage = Common.GetJsString(s_MSG1)

            Case "QPH" '查詢路徑
                Dim s_MSG1 As String = ""
                Dim PH As String = $"{Request("PH")}"
                If PH = "" Then
                    s_MSG1 = "PH,不可為空"
                    sm.LastResultMessage = Common.GetJsString(s_MSG1)
                    Return
                End If
                Dim V_PH As String = TIMS.DecryptAes(PH)
                If V_PH = "" Then
                    s_MSG1 = "V_PH,不可為空"
                    sm.LastResultMessage = Common.GetJsString(s_MSG1)
                    Return
                End If
                Dim V_CNT As Integer = TIMS.GetFileCount(Me, V_PH)
                If V_CNT >= 0 Then
                    Dim V_FNSTR1 As String = TIMS.GetFolderSum(Me, V_PH, V_CNT)
                    s_MSG1 = $"目錄: {V_PH},檔案總數: {V_CNT},檔名: {V_FNSTR1}"
                Else
                    s_MSG1 = $"查詢失敗，請檢查權限或路徑設定。目錄: {V_PH},檔案總數: {V_CNT}"
                End If
                sm.LastResultMessage = Common.GetJsString(s_MSG1)
                Return

        End Select
    End Sub

    ''' <summary>取得100天內任意學員身分證號</summary>
    ''' <returns></returns>
    Private Function GET_NEW_IDNO(ByRef strNAME As String, ByRef strBIRTH As String) As String
        If Not TIMS.OpenDbConn(Me, objconn) Then Return "" ' Exit Sub
        Dim SQL As String = "SELECT TOP 1 IDNO,NAME,CONVERT(VARCHAR(8),BIRTHDAY, 112) BIRTHDAY FROM V_STUDENTINFO WHERE TPLANID='28' AND MODIFYDATE>=GETDATE()-100 ORDER BY NEWID()"
        Dim dt As New DataTable
        Dim sCmd As New SqlCommand(SQL, objconn)
        With sCmd
            dt.Load(.ExecuteReader())
        End With
        If TIMS.dtNODATA(dt) Then Return ""
        strNAME = Convert.ToString(dt.Rows(0)("NAME"))
        strBIRTH = Convert.ToString(dt.Rows(0)("BIRTHDAY"))
        Return Convert.ToString(dt.Rows(0)("IDNO"))
    End Function

    ''' <summary>
    ''' 分析 危險性字元優化程式效能
    ''' </summary>
    ''' <param name="STYPE"></param>
    ''' <param name="f_parms"></param>
    Sub MAIL_TYPE1(ByVal STYPE As String, ByRef f_parms As Hashtable)
        Dim ary_parms As New ArrayList(f_parms.Keys)
        ary_parms.Sort() '排序
        'ary_parms.Reverse() '//反向排序
        Dim flag_use_httpcontext As Boolean = If(IsNothing(Me), True, False)
        Dim s_USERAGENT_INFO As String = ""
        Try
            s_USERAGENT_INFO = TIMS.GetUserAgent(Me, flag_use_httpcontext)
        Catch ex As Exception
        End Try
        Dim sMailBody As String = $"{s_USERAGENT_INFO}{vbCrLf}Login: {TIMS.cst_ErrorMsg2}{vbCrLf}STYPE: {STYPE}{vbCrLf}"
        'For Each oItem As DictionaryEntry In f_parms
        '    If oItem.Key IsNot Nothing AndAlso oItem.Value IsNot Nothing Then
        '        sMailBody &= Convert.ToString(oItem.Key) & " : " & Convert.ToString(oItem.Value) & vbCrLf
        '    End If
        'Next
        For Each strItem As String In ary_parms
            sMailBody &= $"{strItem} : {f_parms(strItem)}{vbCrLf}"
        Next
        Try
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
        Catch ex As Exception
        End Try
        TIMS.SendMailTest(sMailBody)
    End Sub

    ''' <summary>'輸入帳號 登入送出</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_submit_Click(sender As Object, e As System.EventArgs) Handles bt_submit.Click
        'Dim sm As SessionModel = SessionModel.Instance()
        'sm = SessionModel.Instance()

        '所有超過3個月未登入的帳號，設為不啟用
        Call TIMS.UPDATE_STOP_ACCOUNT(objconn)

        Dim flag_can_go_auto_login As Boolean = False 'false:不可使用自動登入功能
        Dim flag_BX As Boolean = False 'false:不可使用測試登入功能
        If txtUserPass.Text = "" Then flag_can_go_auto_login = True '密碼為空可測試自動登入功能
        If flag_can_go_auto_login Then
            '測試/'測試3
            Call TIMS.Chk_TEST_Login1(Me, txtVCode, txtUserId, txtUserPass)
            Call TIMS.Chk_TEST_Login3(Me, txtVCode, txtUserId, txtUserPass)
            flag_BX = TIMS.Chk_TEST_Login3Aes(Me, txtVCode, txtUserId, txtUserPass)
        End If
        If Not flag_BX Then flag_BX = TIMS.sUtl_ChkTest() '測試環境啟用

        sTxtIdUnchange = TIMS.ClearSQM(txtUserId.Text)
        sTxtPass = TIMS.ClearSQM(txtUserPass.Text)
        sVCode = TIMS.ClearSQM(txtVCode.Text)

        '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
        Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(objconn)
        If flag_ChkHISTORY1 Then
            Dim strErrmsg1 As String = ""
            strErrmsg1 = "防止駭客攻擊(紀錄) 啟動-login.bt_submit_Click" & vbCrLf
            strErrmsg1 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg1 &= "GetHTTP_HOST:" & vbCrLf & TIMS.GetHTTP_HOST(Me) & vbCrLf
            'strErrmsg1 = Replace(strErrmsg1, vbCrLf, "<br>" & vbCrLf)
            logger.Warn(strErrmsg1)

            sm.LastErrorMessage = TIMS.cst_ErrorMsg1
            'AuthUtil.LoginLog(txtIDNO.Text, False)

            Const cst_iMaxCanMailCount As Integer = 30 '(寄狀況信數量)
            Dim iGlobalMailCount As Integer = TIMS.GlobalMailCount '目前寄信總數量
            '(狀況信)超過n:30
            If iGlobalMailCount >= cst_iMaxCanMailCount Then
                '(取得使用者正確ip)
                Dim v_IpAddress As String = Common.GetIpAddress() 'MyPage.Request.UserHostAddress
                Dim i_ChkHISTORY1_cnt As Integer = TIMS.SUtl_ChkHISTORY1_CNT(v_IpAddress)
                If i_ChkHISTORY1_cnt > 2 Then
                    TIMS.sUtl_404NOTFOUND(Me, objconn, i_ChkHISTORY1_cnt)
                Else
                    TIMS.sUtl_404NOTFOUND(Me, objconn)
                End If
            End If
            '登出/ 重登
            AuthUtil.LogoutLog()
            '清除登入狀態
            sm.ClearSession()
            Return 'redirectUrl
        ElseIf Not flag_BX AndAlso TIMS.Utl_ChkHISTORY1(Me, objconn, sTxtIdUnchange) Then
            '未使用測試-正式者-檢核 true:攻擊異常達標
            Return
        End If

        'cst_ErrorMsg2
        If sTxtIdUnchange <> "" AndAlso TIMS.CheckInput(sTxtIdUnchange) Then
            Dim sCheckInput As String = TIMS.CheckInputRtn(sTxtIdUnchange)
            If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
            MAIL_TYPE1("sTxtIdUnchange", g_parms)
            '檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = TIMS.cst_ErrorMsg2
            Return
        End If
        'If sTxtPass <> "" AndAlso TIMS.CheckInput(sTxtPass) Then
        '    Dim sCheckInput As String = TIMS.CheckInputRtn(sTxtPass)
        '    If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
        '    MAIL_TYPE1("sTxtPass", g_parms)
        '    sm.LastErrorMessage = TIMS.cst_ErrorMsg2
        '    Return
        'End If
        If sVCode <> "" AndAlso TIMS.CheckInput(sVCode) Then
            Dim sCheckInput As String = TIMS.CheckInputRtn(sVCode)
            If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
            MAIL_TYPE1("sVCode", g_parms)
            '檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = TIMS.cst_ErrorMsg2
            Return
        End If

        Dim s_exMessage As String = ""
        Dim sTxtId As String = ""
        '有值整理為大寫
        If Not String.IsNullOrWhiteSpace(sTxtIdUnchange) Then sTxtId = UCase(sTxtIdUnchange)
        Dim s_exTxtId As String = $",{sTxtId}"
        If sTxtIdUnchange = "" Then
            s_exMessage = $"#GetAccount 發生錯誤: {cst_alert_msg_1}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_1
            Return
        End If
        If sTxtPass = "" Then
            s_exMessage = $"#GetAccount 發生錯誤: {cst_alert_msg_1}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_1
            Return
        End If
        If sVCode = "" Then
            s_exMessage = $"#GetAccount 發生錯誤: {cst_alert_msg_2}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_2
            Return
        End If
        If sVCode <> sm.LoginValidateCode Then
            s_exMessage = $"#GetAccount 發生錯誤: {cst_alert_msg_2b}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_2b
            Return
        End If

        Dim dr As DataRow = Nothing
        Try
            '依帳號取得
            '測試時，不限定(false)/正式，有限定(true)
            Dim flag_Only_BX As Boolean = If(flag_BX, False, True)
            dr = TIMS.sUtl_GetAccount(sTxtId, objconn, flag_Only_BX)
        Catch ex As Exception
            s_exMessage = $"GetAccount 發生錯誤: {ex.Message}{s_exTxtId}"
            TIMS.LOG.Error(s_exMessage, ex)

            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            Throw New Exception(s_exMessage, ex)
        End Try

        If dr Is Nothing Then
            s_exMessage = $"#GetAccount 發生錯誤( dr Is Nothing ): {cst_alert_msg_1bx}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)

            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)

            sm.LastErrorMessage = cst_alert_msg_1bx
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Return
        End If

        'pXssXXrd
        If Convert.ToString(dr.Item("HASHPWD1")) <> TIMS.CreateHash(sTxtPass) Then
            s_exMessage = $"#GetAccount 發生錯誤( HASHPWD1 ):  {cst_alert_msg_1b}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            'Dim vMsg1 As String = "##pXssXXrd"
            'vMsg1 &= String.Format("HASHPWD1:{0}", dr.Item("HASHPWD1")) & vbCrLf
            'vMsg1 &= String.Format("CreateHash:{0}", TIMS.CreateHash(sTxtPass)) & vbCrLf
            'vMsg1 &= String.Format("sTxtPass:{0}", sTxtPass) & vbCrLf
            'TIMS.LOG.Debug(vMsg1)

            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_1b
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        If Convert.ToString(dr.Item("IsUsed")) <> "Y" Then
            s_exMessage = $"GetAccount 發生錯誤: {cst_alert_msg_1c}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            sm.LastErrorMessage = cst_alert_msg_1c
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        If Convert.ToString(dr.Item("Stopmsg")) = "Y" Then '過了帳號停用日
            s_exMessage = $"GetAccount 發生錯誤: {cst_alert_msg_1d}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            sm.LastErrorMessage = cst_alert_msg_1d
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        ' 判斷是否超過三個月未登入
        Dim ReturnMsg As String = Nothing
        Dim flag3 As Boolean = TIMS.Check_AccoutLoginDate(sm, dr("Account"), ReturnMsg, objconn)
        If Not flag3 Then
            s_exMessage = $"GetAccount 發生錯誤: {ReturnMsg}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)
            sm.LastErrorMessage = ReturnMsg
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        If dr.Item("RoleID") <> -1 Then
            ' 帳密登入驗證成功
            Dim userInfo As LoginUserInfo = New LoginUserInfo()
            TIMS.SET_SESSIONMODEL1(sm, userInfo, dr)
            logger.Info("User Logined:" & vbCrLf & userInfo.ToString())
            AuthUtil.LoginLog(sTxtIdUnchange, True)

            ' 取得使用者預設的年度及計畫(由前次年度計畫選擇後儲存)  
            ' 若找不到預設的年度計畫, 則導向年度計畫選擇頁面
            If userInfo.DefaultYear = "" Or userInfo.DefaultPlanID = "" Then
                DbAccess.CloseDbConn(objconn)
                ' 導向計畫選擇頁面
                Response.Redirect(ResolveUrl("~/SelectPlan"))
            Else
                If SelectPlan.SetPlan(objconn, userInfo.DefaultYear, userInfo.DefaultPlanID, False) Then
                    DbAccess.CloseDbConn(objconn)
                    '設定 年度/計畫 成功, 導向登入後首頁
                    Response.Redirect(ResolveUrl("~/Index"))
                Else
                    DbAccess.CloseDbConn(objconn)
                    '預設的 年度/計劃 失敗, 導向計畫選擇頁面
                    sm.LastErrorMessage = Nothing   ' 忽略預設 年度/計劃 的錯誤訊息
                    Response.Redirect(ResolveUrl("~/SelectPlan"))
                End If
            End If
        Else
            s_exMessage = $"GetAccount 發生錯誤: {cst_alert_msg_1f}{s_exTxtId}"
            TIMS.LOG.Warn(s_exMessage)

            'RoleID:-1  使用者 無此功能
            sm.LastErrorMessage = cst_alert_msg_1f
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If
    End Sub

    'Protected Sub bt_MOICA_Click(sender As Object, e As EventArgs) Handles bt_MOICA.Click
    '    'Response.Redirect(ResolveUrl("~/MOICA_Login"))
    '    Response.Redirect(ResolveUrl("~/MOICA_Login.aspx"))
    'End Sub

    ''' <summary>清除登入狀態</summary>
    Sub Utl_reset_1()
        Const cst_test_usr As String = "test_usr"
        Const cst_test_pwd As String = "test_pwd"
        'Dim flag_clearSess As Boolean = False
        'If Not flag_clearSess AndAlso Session(cst_test_usr) Is Nothing Then flag_clearSess = True
        'If Not flag_clearSess AndAlso Convert.ToString(Session(cst_test_usr)) = "" Then flag_clearSess = True
        Session(cst_test_usr) = Nothing
        Session(cst_test_pwd) = Nothing
        sm.ClearSession() '清除登入狀態
        'If flag_clearSess Then sm.ClearSession() '清除登入狀態
    End Sub

    ''' <summary> 重設 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_reset_Click(sender As Object, e As EventArgs) Handles bt_reset.Click
        Call Utl_reset_1()
    End Sub

    ''' <summary> 忘記-FRGTPXSWXD </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_FRGTPXSWXD_Click(sender As Object, e As EventArgs) Handles bt_FRGTPXSWXD.Click
        '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
        Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(objconn)
        If flag_ChkHISTORY1 Then
            Dim strErrmsg1 As String = ""
            strErrmsg1 = "防止駭客攻擊(紀錄) 啟動-bt_submit_Click" & vbCrLf
            strErrmsg1 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg1 &= "GetHTTP_HOST:" & vbCrLf & TIMS.GetHTTP_HOST(Me) & vbCrLf
            'strErrmsg1 = Replace(strErrmsg1, vbCrLf, "<br>" & vbCrLf)
            logger.Warn(strErrmsg1)

            sm.LastErrorMessage = TIMS.cst_ErrorMsg1
            'AuthUtil.LoginLog(txtIDNO.Text, False)

            Const cst_iMaxCanMailCount As Integer = 30 '(寄狀況信數量)
            Dim iGlobalMailCount As Integer = TIMS.GlobalMailCount '目前寄信總數量
            '(狀況信)超過n:30
            If iGlobalMailCount >= cst_iMaxCanMailCount Then
                '(取得使用者正確ip)
                Dim v_IpAddress As String = Common.GetIpAddress() 'MyPage.Request.UserHostAddress
                Dim i_ChkHISTORY1_cnt As Integer = TIMS.SUtl_ChkHISTORY1_CNT(v_IpAddress)
                If i_ChkHISTORY1_cnt > 2 Then
                    TIMS.sUtl_404NOTFOUND(Me, objconn, i_ChkHISTORY1_cnt)
                Else
                    TIMS.sUtl_404NOTFOUND(Me, objconn)
                End If
            End If
            '登出/ 重登
            AuthUtil.LogoutLog()
            '清除登入狀態
            sm.ClearSession()
            Return 'redirectUrl
        ElseIf TIMS.Utl_ChkHISTORY1(Me, objconn, sTxtIdUnchange) Then
            '正式者-檢核 true:攻擊異常達標
            Return 'redirectUrl
        End If

        'sm.LastResultMessage = String.Format("系統已寄發密碼重設通知函至您的E-Mail帳號! [{0}]", vEMAIL)

        Dim redirectUrl As String = ResolveUrl("~/eforgetPwd") 'String.Empty
        'redirectUrl = ResolveUrl("~/login")
        If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
            '檢核成功, 導向首頁
            Response.Redirect(redirectUrl)
        End If

    End Sub

    ''' <summary>atest button 是否顯示</summary>
    Sub At_test()
        Const cst_s_test_ur As String = "test_usr"
        Const cst_s_test_pd As String = "test_pwd"
        bt_atest.Visible = False
        Dim flag_test As Boolean = TIMS.sUtl_ChkTest() '測試環境啟用
        If Not flag_test Then Return '環境錯誤離開此功能

        Dim flag_s_test_ur As Boolean = (Session(cst_s_test_ur) IsNot Nothing AndAlso Convert.ToString(Session(cst_s_test_ur)).Length > 1)
        Dim flag_s_test_pd As Boolean = (Session(cst_s_test_pd) IsNot Nothing AndAlso Convert.ToString(Session(cst_s_test_pd)).Length > 1)
        '(已登入不再顯示)
        bt_atest.Visible = Not (flag_s_test_ur AndAlso flag_s_test_pd)

        'Dim LOGIN_PAGE As String = ""
        'If flag_test Then LOGIN_PAGE = ResolveUrl("~/atestWebForm1.aspx")
    End Sub

    ''' <summary>atest 導向首頁</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub Bt_atest_Click(sender As Object, e As EventArgs) Handles bt_atest.Click
        'Response.Redirect("~/atestWebForm1.aspx")
        Dim redirectUrl As String = ResolveUrl("~/atestWebForm1.aspx") 'String.Empty
        'redirectUrl = ResolveUrl("~/login")
        If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
            '檢核成功, 導向首頁
            Response.Redirect(redirectUrl)
        End If
    End Sub

End Class