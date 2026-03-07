Public Class eforgetPwd
    Inherits System.Web.UI.Page

    'Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Dim sTxtIdUnchange As String = "" '使用帳號
    Dim sTxtIDNO As String = ""
    Dim sTxtEMAIL As String = ""
    Dim sVCode As String = ""

    'Dim s_GID As String = ""
    'Dim s_OTK As String = ""
    'Dim vSENDMAIL As String = ""
    'Dim vHASH1 As String = ""
    Const cst_alert_msg_1 As String = "請輸入帳號/密碼"
    Const cst_alert_msg_1b As String = "帳號或密碼錯誤!"
    Const cst_alert_msg_1c As String = "該帳號尚未啟用!!!"
    Const cst_alert_msg_1d As String = "帳號已被停用,請洽詢系統管理者 !!"
    'Const cst_alert_msg_1f As String = "帳號沒有系統使用權限,請洽詢系統管理者!""帳號已被停用,請洽詢系統管理者!"
    Const cst_alert_msg_1g As String = "查無此帳號!"
    Const cst_alert_msg_2 As String = "請輸入驗證碼"
    Const cst_alert_msg_2b As String = "驗證碼不正確!!!"
    Const cst_alert_msg_2c As String = "請輸入帳號、身分證號碼、E-Mail與驗證碼"
    Const cst_alert_msg_3 As String = "此帳號未設定EMAIL!"

    Const cst_alert_msg_98 As String = "系統已寄發密碼重設通知函至您的E-Mail帳號! "
    Const cst_alert_msg_99 As String = "資料填寫有誤，請重新輸入，或洽系統管理員，謝謝!"

    Dim g_parms As Hashtable
    'Dim g_parms2 As Hashtable
    'Public Const cst_ErrorMsg2 As String = "請勿嘗試在頁面輸入具有危險性的字元!"

    Public BaseUrl As String
    Public sm As SessionModel '= SessionModel.Instance()

    Dim oWDAP As New WDAP
    Dim objconn As SqlConnection

    ''' <summary>
    ''' PageUnload
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    ''' <summary>
    ''' 資安檢核-1
    ''' </summary>
    Sub Critical_Issues_1()
        'openhttps.Value = TIMS.ClearSQM(openhttps.Value)
        g_parms = New Hashtable
        g_parms.Clear()
        g_parms.Add("1.txtUserId.Text", txtUserId.Text)
        g_parms.Add("2.txtUserIdno.Text", txtUserIdno.Text)
        g_parms.Add("3.txtUserEMAIL.Text", txtUserEMAIL.Text)
        g_parms.Add("4.txtVCode.Text", txtVCode.Text)

        sTxtIdUnchange = TIMS.ClearSQM(txtUserId.Text)
        sTxtIDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(txtUserIdno.Text))
        sTxtEMAIL = TIMS.ChangeEmail(TIMS.ClearSQM(txtUserEMAIL.Text))
        sVCode = TIMS.ClearSQM(txtVCode.Text)

        txtUserId.Text = sTxtIdUnchange
        txtUserIdno.Text = sTxtIDNO
        txtUserEMAIL.Text = sTxtEMAIL
        txtVCode.Text = sVCode

        g_parms.Add("sTxtIdUnchg", sTxtIdUnchange)
        g_parms.Add("sTxtIDNO", sTxtIDNO)
        g_parms.Add("sTxtEMAIL", sTxtEMAIL)
        g_parms.Add("sVCode", sVCode)
        '"https://ojtims.wda.gov.tw/emailChgPwd?GID={0}&OTK={1}&OTK2={2}"
        'TIMS.sUtl_404NOTFOUND(Page, objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call Critical_Issues_1()
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub

        sm = SessionModel.Instance()

        BaseUrl = ResolveUrl("~/")

        'If Not String.IsNullOrEmpty(BaseUrl) AndAlso Not BaseUrl.EndsWith("/") Then BaseUrl = String.Concat(BaseUrl, "/")

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)

        If Not IsPostBack Then
            '登出/ 重登
            AuthUtil.LogoutLog()
            '清除登入狀態
            sm.ClearSession()

        End If

    End Sub

    ''' <summary>寄送密碼函</summary>
    Sub SendPWDletter()
        Call Critical_Issues_1()

        '未使用測試-正式者-檢核 true:攻擊異常達標
        If TIMS.Utl_ChkHISTORY1(Me, objconn, sTxtIdUnchange) Then Exit Sub

        Dim flag_test_ENVC As Boolean = TIMS.CHK_IS_TEST_ENVC() '檢測為測試環境:true 正式環境為:false

        'Const cst_alert_msg_1b As String = "查無此帳號!"
        'Const cst_alert_msg_3 As String = "此帳號未設定EMAIL!"
        sVCode = TIMS.ClearSQM(txtVCode.Text)
        txtUserId.Text = TIMS.ClearSQM(txtUserId.Text)
        If sVCode = "" OrElse txtUserId.Text = "" OrElse txtUserIdno.Text = "" OrElse txtUserEMAIL.Text = "" Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_2c
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Return
        End If
        'If sVCode = "" Then
        '    '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
        '    Call TIMS.sUtl_SaveLoginData1(Me, objconn)
        '    sm.LastErrorMessage = cst_alert_msg_2
        '    Return
        'End If
        If sVCode <> sm.LoginValidateCode Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_2b
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Return
        End If
        'If txtUserId.Text = "" Then
        '    sm.LastErrorMessage = cst_alert_msg_1
        '    AuthUtil.LoginLog(sTxtIdUnchange, False)
        '    Exit Sub
        'End If
        'sm.LastErrorMessage = "使用忘記密碼功能，自動登出"
        'AuthUtil.LoginLog(sTxtIdUnchange, False)

        Dim drAA As DataRow = TIMS.sUtl_GetAccount(txtUserId.Text, objconn, False)
        If drAA Is Nothing Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = cst_alert_msg_1g
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If
        Dim vEMAIL As String = TIMS.ChangeEmail(TIMS.ClearSQM(drAA("EMAIL")))
        If vEMAIL = "" Then
            sm.LastErrorMessage = cst_alert_msg_3
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If
        If Not LCase(vEMAIL).Equals(LCase(sTxtEMAIL)) Then
            sm.LastErrorMessage = cst_alert_msg_99 '"資料填寫有誤，請重新輸入，或洽系統管理員，謝謝!"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        Dim vIDNO As String = TIMS.ChangeIDNO(TIMS.ClearSQM(drAA("IDNO")))
        If Not vIDNO.Equals(sTxtIDNO) Then
            sm.LastErrorMessage = cst_alert_msg_99 '"資料填寫有誤，請重新輸入，或洽系統管理員，謝謝!"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        Dim vMODIFYACCT As String = ""
        vMODIFYACCT = If(sm.UserInfo IsNot Nothing, sm.UserInfo.UserID, Convert.ToString(drAA("ACCOUNT")))
        Dim s_parma As New Hashtable
        s_parma.Clear()
        s_parma.Add("ACCOUNT", Convert.ToString(drAA("ACCOUNT")))
        s_parma.Add("SENDMAIL", vEMAIL)
        '/*密碼寄信種類-1:新設密碼-2:忘記密碼(修改密碼)-3:修改密碼 */ SENDMAILTYPE /SDMATYPE
        s_parma.Add("SDMATYPE", "2")
        s_parma.Add("MODIFYACCT", vMODIFYACCT)
        Dim htSS As Hashtable = TIMS.INSERT_PXSSWARDHIS(objconn, s_parma)

        'sm.LastResultMessage = String.Format("系統已寄發密碼重設通知函至您的E-Mail帳號! [{0}]", vEMAIL)
        sm.LastResultMessage = String.Format("{0} {1}", cst_alert_msg_98, TIMS.strMask(vEMAIL, 4)) '"系統已寄發密碼重設通知函至您的E-Mail帳號! "

        If flag_test_ENVC Then
            Dim xMybody As String = TIMS.GetMyValue2(htSS, "xMybody")
            If xMybody <> "" Then Response.Write(xMybody)
        End If
        'eforgetPwd
        'Dim redirectUrl As String = ResolveUrl("~/login") 'String.Empty
        Dim redirectUrl As String = ResolveUrl("~/eforgetPwd") 'String.Empty
        'redirectUrl = ResolveUrl("~/login")
        If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
            '檢核成功, 導向首頁
            Response.Redirect(redirectUrl)
            If Not flag_test_ENVC Then Response.Redirect(redirectUrl)
        End If

    End Sub

    ''' <summary>寄送密碼函</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_submit_Click(sender As Object, e As System.EventArgs) Handles bt_submit.Click
        '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
        Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(objconn)
        If flag_ChkHISTORY1 Then
            Dim strErrmsg1 As String = ""
            strErrmsg1 = "防止駭客攻擊(紀錄) 啟動-eforgetPwd.bt_submit_Click" & vbCrLf
            strErrmsg1 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg1 &= "GetHTTP_HOST:" & vbCrLf & TIMS.GetHTTP_HOST(Me) & vbCrLf
            'strErrmsg1 = Replace(strErrmsg1, vbCrLf, "<br>" & vbCrLf)
            TIMS.LOG.Warn(strErrmsg1)

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
            Return
        End If

        Call SendPWDletter()
    End Sub

    ''' <summary>
    ''' 重設
    ''' </summary>
    Sub Utl_reset_1()
        txtUserId.Text = ""
        txtUserIdno.Text = ""
        txtUserEMAIL.Text = ""
        txtVCode.Text = ""
    End Sub

    ''' <summary>
    ''' 重設
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_reset_Click(sender As Object, e As EventArgs) Handles bt_reset.Click
        Call Utl_reset_1()
    End Sub

    ''' <summary>
    ''' 關閉
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_close1_Click(sender As Object, e As EventArgs) Handles bt_close1.Click

    End Sub

    Protected Sub bt_back1_Click(sender As Object, e As EventArgs) Handles bt_back1.Click
        Dim redirectUrl As String = ResolveUrl("~/login")
        Response.Redirect(redirectUrl)
    End Sub
End Class