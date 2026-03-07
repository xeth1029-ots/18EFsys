Partial Class MOICA_Login
    Inherits System.Web.UI.Page

    Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Private sm As SessionModel
    Dim vc As Turbo.Commons.ValidateCode = New Turbo.Commons.ValidateCode()

    Dim g_parms As Hashtable
    'Public Const cst_ErrorMsg2 As String = "請勿嘗試在頁面輸入具有危險性的字元!"

    Const cst_IndexPage As String = "index"

    'Const cst_MOICA_Login1 As String = "top.location.href='MOICA_Login';"
    'Const cst_LoginPage1 As String = "login"
    Const cst_errMsg1 As String = "*****自然人憑證連線有誤(請重試)!!*****" & vbCrLf
    Const cst_errMsg2 As String = " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf

    Dim sExToString As String = "" '(本機)錯誤訊息儲存

    Dim oWDAP As New WDAP
    Dim objconn As SqlConnection

    Private Sub SUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    ''' <summary> 資安檢核 </summary>
    Sub Critical_Issues_1()
        g_parms = New Hashtable
        g_parms.Clear()
        g_parms.Add("Hide_PinVerify.Value", Hide_PinVerify.Value)
        g_parms.Add("Hide_cadata.Value", Hide_cadata.Value)
        g_parms.Add("Hide_enccert.Value", Hide_enccert.Value)
        g_parms.Add("txtIDNO.Text", txtIDNO.Text)
        g_parms.Add("txtPin.Text", txtPin.Text)
        g_parms.Add("Hide_sign.Value", Hide_sign.Value)

        'openhttps.Value = TIMS.ClearSQM(openhttps.Value)
        Hide_PinVerify.Value = TIMS.ClearSQM(Hide_PinVerify.Value)
        Hide_cadata.Value = TIMS.ClearSQM(Hide_cadata.Value)
        Hide_enccert.Value = TIMS.ClearSQM(Hide_enccert.Value)
        txtIDNO.Text = TIMS.ClearSQM(txtIDNO.Text)
        Hide_sign.Value = TIMS.ClearSQM(Hide_sign.Value)

        g_parms.Add("cHide_PinVerify.Value", Hide_PinVerify.Value)
        g_parms.Add("cHide_cadata.Value", Hide_cadata.Value)
        g_parms.Add("cHide_enccert.Value", Hide_enccert.Value)
        g_parms.Add("ctxtIDNO.Text", txtIDNO.Text)
        g_parms.Add("cHide_sign.Value", Hide_sign.Value)

        'ReturnUrl.Value = TIMS.ClearSQM(ReturnUrl.Value)
        'txtVCode.Text = TIMS.ClearSQM(txtVCode.Text)
        'txtUserId.Text = TIMS.ClearSQM(txtUserId.Text)
    End Sub

    Sub MAIL_TYPE1(ByVal STYPE As String, ByRef f_parms As Hashtable)
        Dim flag_use_httpcontext As Boolean = If(IsNothing(Me), True, False)
        Dim s_USERAGENT_INFO As String = ""
        Try
            s_USERAGENT_INFO = TIMS.GetUserAgent(Me, flag_use_httpcontext)
        Catch ex As Exception
        End Try
        Dim sMailBody As String = s_USERAGENT_INFO
        sMailBody &= String.Concat("MOICA_Login:", TIMS.cst_ErrorMsg2, vbCrLf)
        sMailBody &= String.Concat("STYPE:", STYPE, vbCrLf)
        For Each oItem As DictionaryEntry In f_parms
            If oItem.Key IsNot Nothing AndAlso oItem.Value IsNot Nothing Then
                sMailBody &= String.Concat(oItem.Key, " : ", oItem.Value, vbCrLf)
            End If
        Next
        Try
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
        Catch ex As Exception
        End Try
        'sMailBody &= "Hide_PinVerify.Value:" & Hide_PinVerify.Value & vbCrLf
        'sMailBody &= "Hide_cadata.Value:" & Hide_cadata.Value & vbCrLf
        'sMailBody &= "Hide_enccert.Value:" & Hide_enccert.Value & vbCrLf
        'sMailBody &= "txtIDNO.Text :" & txtIDNO.Text & vbCrLf
        'sMailBody &= "Hide_sign.Value :" & Hide_sign.Value & vbCrLf
        TIMS.SendMailTest(sMailBody)
    End Sub

    '第1次載入
    Sub SCreatex1()
        'nonce.Value = Session.SessionID
        Const cst_cc_sRtnUrl As String = "/" 'TIMS.Utl_GetConfigSet("RtnUrl")'不可接受/
        Dim rqReturnUrl As String = TIMS.ClearSQM(Request("ReturnUrl")) ': rqMsgid = TIMS.ClearSQM(rqMsgid)
        If rqReturnUrl <> "" AndAlso cst_cc_sRtnUrl.IndexOf(rqReturnUrl) = -1 Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 'Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            TIMS.sUtl_404NOTFOUND(Me, objconn)
            Return
        End If

        '所有超過3個月未登入的帳號，設為不啟用
        Call TIMS.UPDATE_STOP_ACCOUNT(objconn)

        '隱藏【帳號密碼登入】按鈕'將「帳號密碼登入」鈕隱藏。(正式/測試) u SYS_VAR set ITEMVALUE='Y' FROM SYS_VAR WHERE SVID=85 AND TPLANID='00' AND SPAGE='MOICA_LOGIN' AND ITEMVALUE='N'
        'S * FROM SYS_VAR WHERE SVID=85 AND TPLANID='00' AND SPAGE='MOICA_LOGIN' 
        'cst_SPAGE1, "PWDLOGIN", "00", objconn / U SYS_VAR set ITEMVALUE='N' FROM SYS_VAR WHERE SVID=85 AND TPLANID='00' AND SPAGE='MOICA_LOGIN' AND ITEMVALUE='Y'
        Dim s_PWDLOGIN As String = TIMS.GetSystemValue("MOICA_LOGIN", "00", "PWDLOGIN", objconn)
        btnPWDLOGIN.Visible = (s_PWDLOGIN <> "" AndAlso s_PWDLOGIN.Equals("Y"))

        Dim sAltMsg As String = "" '訊息
        Dim AltMsgSDate As String = "" '訊息公佈日
        Dim AltMsgEDate As String = "" '訊息結束日
        sAltMsg = TIMS.Get_SHAREDCODE_MSG("ALTMSG", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsgSDate = TIMS.Get_SHAREDCODE_MSG("SDATE", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsgEDate = TIMS.Get_SHAREDCODE_MSG("EDATE", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)

        Dim strShowMsg As String = TIMS.Get_AltMsg_System_Msg(sAltMsg, AltMsgSDate, AltMsgEDate)
        If strShowMsg <> "" Then sm.LastResultMessage = strShowMsg

        ' 產生驗證碼, 用來防止不正常操作及攻擊
        sm.LoginValidateCode = vc.CreateValidateCode(6)

    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        sm = SessionModel.Instance()
        objconn = DbAccess.GetConnection
        AddHandler MyBase.Unload, AddressOf SUtl_PageUnload

        Call Critical_Issues_1()

        ' 檢測資料庫連線, 無法連線時導向停機頁
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub
        ' 判斷是否為計劃停機期間
        If TIMS.Stop1(Me, objconn) Then Exit Sub

        Call TIMS.Stop4(Me, objconn)

        If Hide_PinVerify.Value <> "" AndAlso TIMS.CheckInput(Hide_PinVerify.Value) Then
            Dim sCheckInput As String = TIMS.CheckInputRtn(Hide_PinVerify.Value)
            If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
            MAIL_TYPE1("Hide_PinVerify", g_parms)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            Me.Lit_LastErrorMessage.Text = TIMS.cst_ErrorMsg2
            Return
        End If
        If Hide_cadata.Value <> "" AndAlso TIMS.CheckInput(Hide_cadata.Value) Then
            Dim sCheckInput As String = TIMS.CheckInputRtn(Hide_cadata.Value)
            If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
            MAIL_TYPE1("Hide_cadata", g_parms)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            Me.Lit_LastErrorMessage.Text = TIMS.cst_ErrorMsg2
            Return
        End If
        'If Hide_enccert.Value <> "" AndAlso TIMS.CheckInput(Hide_enccert.Value) Then
        '    Dim sCheckInput As String = TIMS.CheckInputRtn(Hide_enccert.Value)
        '    If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
        '    MAIL_TYPE1("Hide_enccert", g_parms)
        '    Me.Lit_LastErrorMessage.Text = TIMS.cst_ErrorMsg2
        '    Return
        'End If
        If txtIDNO.Text <> "" AndAlso TIMS.CheckInput(txtIDNO.Text) Then
            Dim sCheckInput As String = TIMS.CheckInputRtn(txtIDNO.Text)
            If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
            MAIL_TYPE1("txtIDNO", g_parms)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            Me.Lit_LastErrorMessage.Text = TIMS.cst_ErrorMsg2
            Return
        End If
        If Hide_sign.Value <> "" AndAlso TIMS.CheckInput(Hide_sign.Value) Then
            Dim sCheckInput As String = TIMS.CheckInputRtn(Hide_sign.Value)
            If sCheckInput <> "" Then g_parms.Add("sCheckInput", sCheckInput)
            MAIL_TYPE1("Hide_sign", g_parms)
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            Me.Lit_LastErrorMessage.Text = TIMS.cst_ErrorMsg2
            Return
        End If

        If Hide_PinVerify.Value = "Ok" And Hide_cadata.Value <> "" And Hide_enccert.Value <> "" Then
            ' 登入作業: 前端 PIN 檢核通過, 檢核憑證效期及帳號綁定

            ' 驗證碼符合 If Hide_vcode.Value = sm.LoginValidateCode Then End If
            Dim redirectUrl As String = String.Empty
            redirectUrl = SUtl_CheckCAData(Hide_cadata.Value, Hide_enccert.Value)

            If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
                '檢核成功, 導向首頁
                Response.Redirect(redirectUrl)
            End If

        End If

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)
        If Not Page.IsPostBack Then
            ' Get 連結到 Login 視為登出/重登 ' 清除登入狀態
            AuthUtil.LogoutLog()
            ' 清除登入狀態
            sm.ClearSession()
            ' 清除所有Session
            Session.RemoveAll()

            Call SCreatex1()
        End If

        'Hide_vcode.Value = sm.LoginValidateCode
        Me.Lit_LastResultMessage.Text = sm.LastResultMessage
        Me.Lit_LastErrorMessage.Text = sm.LastErrorMessage
        Me.Lit_RedirectUrlAfterBlock.Text = sm.RedirectUrlAfterBlock
    End Sub

    ''' <summary>自然人憑證登入檢核
    ''' <para>檢核失敗時會原因記錄在 SessionModel.LastErrorMessage, 並回傳 Nothing</para>
    ''' <para>檢核成功, 則回傳待重導的頁面 url</para>
    ''' </summary>
    ''' <param name="sCAData"></param>
    ''' <param name="sEncCert"></param>
    ''' <returns></returns>
    Function SUtl_CheckCAData(ByVal sCAData As String, ByVal sEncCert As String) As String

        Dim sArrData() As String = Split(sCAData, "~~")
        Dim sCardId As String
        Dim sName As String
        Dim sLastIDNO As String

        Dim redirectUrl As String = String.Empty
        '防止駭客攻擊(紀錄) 啟動-Login / true:攻擊異常達標 false:未達標
        Dim flag_ChkHISTORY1 As Boolean = TIMS.sUtl_ChkHISTORY1(objconn)
        If flag_ChkHISTORY1 Then
            '取得錯誤資訊寫入
            Dim strErrmsg1 As String = String.Concat("防止駭客攻擊(紀錄) 啟動-MOICA_Login.sUtl_CheckCAData", vbCrLf, TIMS.GetErrorMsg(Me), "GetHTTP_HOST:", vbCrLf, TIMS.GetHTTP_HOST(Me), vbCrLf)
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
            Return redirectUrl
        End If

        If sArrData.Length <> 3 Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = "憑證資料有誤!!"
            AuthUtil.LoginLog(txtIDNO.Text, False)
            Return redirectUrl
        End If

        sName = sArrData(0)
        sCardId = sArrData(1)
        sLastIDNO = sArrData(2)
        logger.Info(String.Format("Name={0}, CardId={1}, LatIDNO4={2}", sName, sCardId, sLastIDNO))

        Dim x509Helper As X509CertsHelper
        Try
            x509Helper = New X509CertsHelper(sEncCert)
            logger.Info(x509Helper.ToString())

        Catch ex As Exception
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            logger.Error("憑證資料有誤: " & ex.Message, ex)
            sm.LastErrorMessage = "憑證資料有誤!!"
            AuthUtil.LoginLog(txtIDNO.Text, False)
            Return redirectUrl
        End Try

        If x509Helper.IsExpired Then
            sm.LastErrorMessage = "憑證已過期"
        ElseIf x509Helper.IsCRL Then
            sm.LastErrorMessage = "憑證已廢止(無效)"
        ElseIf x509Helper.CardID <> sCardId Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = "憑證序號不一致"
        ElseIf Not txtIDNO.Text.EndsWith(sLastIDNO) Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            sm.LastErrorMessage = "憑證簽發對象身分證號不符合"
        Else
            ' 檢核使用者綁定
            Dim rtn As Integer = SUtl_CheckUser(txtIDNO.Text, x509Helper.SerialNumber)

            If rtn = 0 Then
                ' 使用者帳號驗證成功

                Dim userInfo As LoginUserInfo = sm.UserInfo
                AuthUtil.LoginLog(sm.UserInfo.UserID, True)

                ' 取得使用者預設的年度及計畫(由前次年度計畫選擇後儲存)  
                ' 若找不到預設的年度計畫, 則導向年度計畫選擇頁面
                If userInfo.DefaultYear = "" Or userInfo.DefaultPlanID = "" Then
                    ' 導向計畫選擇頁面
                    redirectUrl = ResolveUrl("~/SelectPlan")
                Else
                    If SelectPlan.SetPlan(objconn, userInfo.DefaultYear, userInfo.DefaultPlanID, False) Then
                        ' 設定 年度/計畫 成功, 導向登入後首頁
                        redirectUrl = ResolveUrl("~/Index")
                    Else
                        ' 預設的 年度/計劃 失敗, 導向計畫選擇頁面
                        sm.LastErrorMessage = Nothing   ' 忽略預設 年度/計劃 的錯誤訊息
                        redirectUrl = ResolveUrl("~/SelectPlan")
                    End If
                End If

            ElseIf rtn = 1 Then
                '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
                Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
                ' 1.帳號存在但檢核失敗(訊息記錄到 SessionModel.LastErrorMessage)

                Dim message As String = sm.LastErrorMessage
                If (message = "") Then message = "帳號存在但憑證檢核失敗"
                sm.LastErrorMessage = message
                AuthUtil.LoginLog(sm.UserInfo.UserID, False)

            ElseIf rtn = 2 Then
                ' 2.指定的憑證內部序號找不到綁定使用者(或第一次登入)

                ' 導向自然人憑證上傳綁定頁面
                Session("MOICA_IDNO") = txtIDNO.Text
                Session("MOICA_SerialNumber") = x509Helper.SerialNumber

                sm.LastResultMessage = "您的憑證找不到綁定使用者！<br/>如果這是您第一次登入，須進行系統帳號密碼驗證以完成綁定動作。"
                AuthUtil.LoginLog(txtIDNO.Text, False)

                redirectUrl = ResolveUrl("~/MOICA_upload")

            Else
                ' 未知的回傳值(不應該跑到這裡, 單純防呆)
                Dim msg As String = "系統異常: sUtl_CheckUser() 未知的回傳值, rtn=" & rtn
                logger.Error(msg)
                sm.LastErrorMessage = msg
                AuthUtil.LoginLog(sm.UserInfo.UserID, False)
            End If

        End If

        Return redirectUrl
    End Function

    ''' <summary>自然人憑證檢核通過後, 檢核使用者帳號及綁定, 回傳值:
    ''' <para>0.檢核通過</para>
    ''' <para>1.帳號存在但檢核失敗(訊息記錄到 SessionModel.LastErrorMessage)</para>
    ''' <para>2.指定的憑證內部序號找不到綁定使用者</para>
    ''' </summary>
    ''' <param name="sIDNO">身分證字號</param>
    ''' <param name="sSerialNumber">憑證內部序號</param>
    ''' <returns></returns>
    Function SUtl_CheckUser(ByVal sIDNO As String, ByVal sSerialNumber As String) As Integer
        Dim dr As DataRow = Nothing
        Try
            dr = TIMS.sUtl_GetAccount1(sSerialNumber, objconn)
        Catch ex As Exception
            Throw New Exception("GetAccount 發生錯誤:" & ex.Message, ex)
        End Try

        ' 指定的憑證內部序號找不到綁定使用者
        If dr Is Nothing Then Return 2

        ' 帳號存成文字檔
        Dim sUserNo As String = Convert.ToString(dr.Item("ACCOUNT"))

        Dim userInfo As LoginUserInfo = New LoginUserInfo()
        '使用者資訊保存到 userInfo
        TIMS.SET_SESSIONMODEL1(userInfo, dr)
        ' 使用者資訊保存到 SessionModel
        sm.UserInfo = userInfo

        If Convert.ToString(dr.Item("IDNO")) <> sIDNO Then
            sm.LastErrorMessage = "憑證綁定的身分證號不符"
            'AuthUtil.LoginLog(sIDNO, False)
            Return 1
        End If

        If Convert.ToString(dr.Item("IsUsed")) <> "Y" Then
            sm.LastErrorMessage = "該帳號尚未啟用!!!"
            'AuthUtil.LoginLog(sUserNo, False)
            Return 1
        End If

        If Convert.ToString(dr.Item("Stopmsg")) = "Y" Then '過了帳號停用日
            sm.LastErrorMessage = "帳號已被停用,請洽詢系統管理者!"
            'AuthUtil.LoginLog(sUserNo, False)
            Return 1
        End If

        ' 判斷是否超過三個月未登入
        Dim ReturnMsg As String = Nothing
        Dim flag3 As Boolean = TIMS.Check_AccoutLoginDate(sm, sUserNo, ReturnMsg, objconn)
        If Not flag3 Then
            sm.LastErrorMessage = ReturnMsg
            'AuthUtil.LoginLog(sUserNo, False)
            Return 1
        End If

        If dr.Item("RoleID") <> -1 Then
            ' 帳密登入驗證成功
            TIMS.SET_SESSIONMODEL1(sm, userInfo, dr)

            logger.Info("User Logined:" & vbCrLf & userInfo.ToString())
            'AuthUtil.LoginLog(sUserNo, True)

            Return 0
        Else
            'RoleID:-1  使用者 無此功能
            sm.LastErrorMessage = "帳號沒有系統使用權限,請洽詢系統管理者!!!"
            'AuthUtil.LoginLog(sUserNo, False)
            Return 1
        End If

    End Function

End Class
