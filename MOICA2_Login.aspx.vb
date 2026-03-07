Imports System.Web.Security
Imports System.Threading

Partial Class MOICA2_Login
    Inherits System.Web.UI.Page
    '檢測網頁：
    'http://163.29.199.217/HiPKIDemoRoot2/util/checkHiPKI.asp

    Const cst_MOICA_Login1 As String = "top.location.href='MOICA2_Login.aspx';"
    Const cst_LoginPage1 As String = "login.aspx"
    Const cst_errMsg1 As String = "*****自然人憑證連線有誤(請重試)!!*****" & vbCrLf
    Const cst_errMsg2 As String = " (若連續出現多次(3次以上)請連絡系統管理者協助，謝謝)(造成您的不便深感抱歉) " & vbCrLf
    Dim sExToString As String = "" '(本機)錯誤訊息儲存

#Region "SUB S1"
    '顯示超連結
    Sub Show_Hyperlink()
        '測試用
        'Me.Btn_5.Visible = False
        'Me.Btn_6.Visible = False
        'If TIMS.sUtl_ChkTest() Then
        '    '線上甄試
        '    Me.Btn_5.Visible = True
        '    '學員問卷測試
        '    Me.Btn_6.Visible = True
        'End If

        '測試用
        'If TIMS.sUtl_ChkTest() Then
        '    Dim SecretLogin As String = TIMS.Utl_GetConfigSet("Secret_Login") 'System.Configuration.ConfigurationSettings.AppSettings("Secret_Login")
        '    Btn_X.Visible = True
        '    Btn_X.Attributes.Add("OnClick", "window.open('" & SecretLogin & "','_self');return false;")
        '    Btn_X.Style.Add("cursor", "hand")
        'End If
        'Dim TestStr As String = ConfigurationSettings.AppSettings("AmuTestString") '測試用
        'If TestStr = "AmuTest" Then '測試用
        'End If '測試用

        Btn_3.Style.Add("cursor", "hand")
        Btn_7.Style.Add("cursor", "hand")
        '學員填寫期末滿意度
        AddHandler Btn_3.Click, AddressOf sUtl_BtnGroupClick
        '學員受訓期間意見調查表填寫
        AddHandler Btn_7.Click, AddressOf sUtl_BtnGroupClick
        Btn_3.CommandName = TIMS.cst_Btn3_CmdName
        Btn_7.CommandName = TIMS.cst_Btn7_CmdName

        'Btn_1.Style.Add("cursor", "hand")
        'Btn_2.Style.Add("cursor", "hand")
        'Btn_4.Style.Add("cursor", "hand")
        'Btn_5.Style.Add("cursor", "hand")
        'Btn_6.Style.Add("cursor", "hand")
        'Btn_8.Style.Add("cursor", "pointer")

        ''申請帳號
        'AddHandler Btn_1.Click, AddressOf sUtl_BtnGroupClick
        ''學員報名狀況查詢
        'AddHandler Btn_2.Click, AddressOf sUtl_BtnGroupClick
        ''線上甄試
        'AddHandler Btn_5.Click, AddressOf sUtl_BtnGroupClick
        ''學員問卷測試
        'AddHandler Btn_6.Click, AddressOf sUtl_BtnGroupClick
        ''測試區
        'AddHandler Btn_4.Click, AddressOf sUtl_BtnGroupClick
        ''學員線上請假
        'AddHandler Btn_8.Click, AddressOf sUtl_BtnGroupClick

        'Btn_1.CommandName = TIMS.cst_Btn1_CmdName
        'Btn_2.CommandName = TIMS.cst_Btn2_CmdName
        'Btn_4.CommandName = TIMS.cst_Btn4_CmdName
        'Btn_5.CommandName = TIMS.cst_Btn5_CmdName
        'Btn_6.CommandName = TIMS.cst_Btn6_CmdName
        'Btn_8.CommandName = TIMS.cst_Btn8_CmdName

    End Sub

    '網頁切換
    Sub showList(ByVal iType As Integer)
        div12.Visible = False
        tbMoica1.Visible = False
        tbBtnGroup1.Visible = False
        tbBtnGroup2.Visible = False
        Select Case iType
            Case 0 '載入
                tbMoica1.Visible = True
                tbBtnGroup1.Visible = True
                div12.Visible = True
            Case 1 '回上頁
                tbMoica1.Visible = True
                tbBtnGroup1.Visible = True
                div12.Visible = True
            Case 2 '切換網頁 等待使用者輸入 驗證碼
                tbBtnGroup2.Visible = True
            Case Else
                '(不知道發生什麼事)
                '防止駭客攻擊(紀錄)
                Call TIMS.sUtl_SaveLoginData1(Me, objconn)
        End Select

    End Sub

    '按鈕應該做什麼  顯示  驗證碼 網頁
    Sub sUtl_BtnGroupClick(sender As Object, e As EventArgs)
        Hid_BtnV1.Value = ""
        Try
            Select Case DirectCast(sender, System.Web.UI.WebControls.Button).CommandName
                Case TIMS.cst_Btn1_CmdName, TIMS.cst_Btn2_CmdName, TIMS.cst_Btn3_CmdName, TIMS.cst_Btn4_CmdName, TIMS.cst_Btn5_CmdName, TIMS.cst_Btn6_CmdName, TIMS.cst_Btn7_CmdName, TIMS.cst_Btn8_CmdName
                    '合理輸入資訊
                Case Else
                    '防止駭客攻擊(紀錄)
                    Call TIMS.sUtl_SaveLoginData1(Me, objconn)
                    '(異常) 不合理
                    TIMS.sUtl_404NOTFOUND(Me, objconn)
            End Select
        Catch ex As Exception
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            '(異常) 
            TIMS.sUtl_404NOTFOUND(Me, objconn)
        End Try

        '(無異常) 顯示 驗證碼 button 並 存取 Hid_BtnV1 CommandName
        'cst_Btn1_CmdName~cst_Btn8_CmdName
        Hid_BtnV1.Value = DirectCast(sender, System.Web.UI.WebControls.Button).CommandName
        'Session(TIMS.cst_MOICA_Login) = Hid_BtnV1.Value

        showList(2)
        tbMoica1.Visible = False
        tbBtnGroup1.Visible = False
        tbBtnGroup2.Visible = True
    End Sub

    '驗證碼 - 送出鈕
    Sub sUtl_BtnSubmit2()
        Dim vsStrScript As String = ""
        If Session("vnum") Is Nothing Then
            '防止駭客攻擊(紀錄)
            'Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            vsStrScript = "<script>alert('驗證碼尚未產生，請重新整理!!!');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If
        If Convert.ToString(Session("vnum")) = "" Then
            '防止駭客攻擊(紀錄)
            'Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            vsStrScript = "<script>alert('驗證碼尚未產生，請重新整理!!!');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If
        txtvnum.Text = TIMS.ClearSQM(txtvnum.Text)
        If txtvnum.Text = "" Then
            '防止駭客攻擊(紀錄)
            'Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            vsStrScript = "<script>alert('驗證碼不正確!!!');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If

        Dim SSvnum1 As String = TIMS.ClearSQM(txtvnum.Text)
        Dim SSvnum2 As String = TIMS.ClearSQM(Session("vnum"))
        If LCase(SSvnum1) <> LCase(SSvnum2) Then
            txtvnum.Text = ""
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)

            vsStrScript = "<script>alert('驗證碼不正確!!!');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If

        If TIMS.sUtl_ChkHISTORY1(Me, objconn) Then
            Dim strErrmsg1 As String = ""
            strErrmsg1 = ""
            strErrmsg1 = "防止駭客攻擊(紀錄) 啟動-MOICA2_Login.sUtl_BtnSubmit2" & vbCrLf
            strErrmsg1 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg1 &= "GetHTTP_HOST:" & vbCrLf & TIMS.GetHTTP_HOST(Me) & vbCrLf
            strErrmsg1 = Replace(strErrmsg1, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.SendMailTest(strErrmsg1)

            Common.MessageBox(Me, TIMS.cst_ErrorMsg1)
            Exit Sub
        End If

        '(無異常)Hid_BtnV1 CommandName 置換session
        'cst_Btn1_CmdName~cst_Btn8_CmdName
        'Hid_BtnV1.Value = DirectCast(sender, System.Web.UI.WebControls.Button).CommandName
        If Hid_BtnV1.Value = "" Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            '(異常) 
            TIMS.sUtl_404NOTFOUND(Me, objconn)
        End If
        Session(TIMS.cst_MOICA_Login) = Hid_BtnV1.Value

        '驗證碼重整
        txtvnum.Text = ""
        Session("vnum") = ValidateCode.rndnum(4)

        '驗證碼通過
        Dim url1 As String = ""
        Select Case Hid_BtnV1.Value
            Case TIMS.cst_Btn1_CmdName '申請帳號
                url1 = "AC/01/AC_01_001.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case TIMS.cst_Btn2_CmdName '學員報名狀況查詢
                url1 = "AC/02/AC_02_001.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case TIMS.cst_Btn3_CmdName '學員填寫期末滿意度
                url1 = "AC/03/AC_03_001.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case TIMS.cst_Btn5_CmdName '線上甄試
                url1 = "AC/05/AC_05_001.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case TIMS.cst_Btn6_CmdName '學員問卷測試
                url1 = "AC/06/AC_06_001.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case TIMS.cst_Btn7_CmdName '學員受訓期間意見調查表填寫
                url1 = "AC/07/AC_07_001.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case TIMS.cst_Btn4_CmdName '測試區
                '測試區
                'http://163.29.199.236/DEMOTIME/USERID_Login.aspx
                Dim testUrl1 As String = TIMS.Utl_GetConfigSet("DemoHttpUrl")
                If testUrl1 <> "" Then
                    TIMS.Utl_Redirect(Me, objconn, testUrl1)
                    Exit Sub
                End If
                Common.MessageBox(Me, "測試區暫不開放!!!")
                Exit Sub

            Case TIMS.cst_Btn8_CmdName '學員線上請假
                url1 = "AC/03/AC_03_002.aspx"
                TIMS.Utl_Redirect(Me, objconn, url1)

            Case Else
                '防止駭客攻擊(紀錄)
                Call TIMS.sUtl_SaveLoginData1(Me, objconn)

        End Select

    End Sub

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    '第1次載入
    Sub sCreatex1()
        Dim rqReturnUrl As String = Request("ReturnUrl") ': rqMsgid = TIMS.ClearSQM(rqMsgid)
        Dim sRtnUrl As String = TIMS.Utl_GetConfigSet("RtnUrl")
        If rqReturnUrl <> "" Then
            '若不為空可接受字元檢核
            If sRtnUrl.IndexOf(rqReturnUrl) = -1 Then
                '防止駭客攻擊(紀錄)
                Call TIMS.sUtl_SaveLoginData1(Me, objconn)
                TIMS.sUtl_404NOTFOUND(Me, objconn)
                Exit Sub
            End If
        End If

        Dim rqMsgid As String = Request("msgid") ': rqMsgid = TIMS.ClearSQM(rqMsgid)
        Call showList(0)
        bt_submit.Attributes.Add("OnClick", "return CheckCard();")
        txtvnum.Attributes.Add("autocomplete", "off")
        txtname.Attributes.Add("autocomplete", "off")
        txtpass.Attributes.Add("autocomplete", "off")
        AltMsg.Attributes.Add("autocomplete", "off")

        Dim sAltMsg As String = "" '訊息
        Dim AltMsgSDate As String = "" '訊息公佈日
        Dim AltMsgEDate As String = "" '訊息結束日
        sAltMsg = TIMS.Get_SHAREDCODE_MSG("ALTMSG", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsgSDate = TIMS.Get_SHAREDCODE_MSG("SDATE", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsgEDate = TIMS.Get_SHAREDCODE_MSG("EDATE", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsg.Value = TIMS.Get_AltMsg_System_Msg(sAltMsg, AltMsgSDate, AltMsgEDate)

        'Session("serialno") = Nothing
        nonce.Value = Session.SessionID
        'Me.bt_submit.Attributes.Add("OnClick", "return CheckCard();")
        'Me.signeddata1.Value = Now().ToString & TIMS.GetGUID().ToString
        'Session("signeddata1") = Me.signeddata1.Value
        If rqMsgid = "RELOGIN" Then
            'Common.RespWrite(Me, "<script>alert('由於你太久沒有操作系統，系統遺失你的登入資訊，請重新登入!!');</script>")
            'Common.RespWrite(Me, "<script>top.location.href='MOICA_Login.aspx';</script>")
            'Response.End()
            'Exit Sub
            Dim Script1 As String = "<script>alert('" & TIMS.cst_NODATAMsg5 & "');" & cst_MOICA_Login1 & "</script>"
            Call TIMS.Utl_RespWriteEnd(Me, objconn, Script1)
        End If

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)
    End Sub

    'MyBase.Load, Me.Load
    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        'Response.Cache.SetExpires(DateTime.Now)'設定 Expires HTTP 標頭為絕對日期和時間。
        'Response.Cache.SetExpires(DateTime.Now.AddSeconds(360))
        '在這裡放置使用者程式碼以初始化網頁
        Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        If Not TIMS.Get_httpsProtocol Then Me.openhttps.Value = "0" Else Me.openhttps.Value = ""

        Btn_X.Visible = False
        Dim rqMsgid As String = Request("msgid") '.net 預設資訊': rqMsgid = TIMS.ClearSQM(rqMsgid)
        'Dim rqSig As String = Request("credential") ': rqSig = TIMS.ClearSQM(rqSig)
        'Dim rqTxtname As String = Request("txtname") ': rqTxtname = TIMS.ClearSQM(rqTxtname)
        Dim moica_divC2 As String = TIMS.Utl_GetConfigSet("moica_divC2") ' System.Configuration.ConfigurationSettings.AppSettings("moica_divC2")
        If moica_divC2 = "Y" Then divC2.Visible = False

        '顯示超連結
        Call Show_Hyperlink()

        objconn = DbAccess.GetConnection
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub

        '顯示orgID資訊
        labBottomContent.Text = TIMS.ShowBottomContent("162", objconn)

        If Not Page.IsPostBack Then
            Call sCreatex1()
        End If

    End Sub

    '錯誤寄信提醒
    Public Shared Sub sUtl_MailError1(ByRef MyPage As Page, ByVal oConn As SqlConnection, _
                                      ByVal htSS As Hashtable)

        Dim rqSig As String = TIMS.GetMyValue2(htSS, "rqSig")
        Dim Errmsg As String = TIMS.GetMyValue2(htSS, "Errmsg")
        Dim strErrmsg As String = TIMS.GetMyValue2(htSS, "strErrmsg")
        Dim sExToString As String = TIMS.GetMyValue2(htSS, "sExToString")

        Dim flagErrMail As Boolean = False '(是否已寄信 true:寄了(不用再寄) false:還沒)

        'Errmsg：Nonce比對錯誤(s):0x85001003
        If Errmsg.IndexOf("Nonce比對錯誤(s)") > -1 Then
            flagErrMail = True 'true:寄了(不用再寄)
        End If

        '沒有寄信，且有錯誤，再寄一次信
        If Not flagErrMail AndAlso sExToString <> "" Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(MyPage, oConn)

            'Common.MessageBox(Me, strErrmsg)
            '置換換行符號
            Try
                Dim sMailMsg As String = ""
                Dim iSendMailCount As Integer = TIMS.SendMailCount() '目前寄信總數量
                sMailMsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
                sMailMsg &= "rqSig：" & vbCrLf & rqSig & vbCrLf
                If Not MyPage.Session Is Nothing Then sMailMsg &= "Session.SessionID：" & CStr(MyPage.Session.SessionID) & vbCrLf 'End If
                sMailMsg &= "Errmsg：" & Errmsg & vbCrLf '伺服器有回應的錯誤
                sMailMsg &= "strErrmsg：" & strErrmsg & vbCrLf '本機錯誤訊息(提供給使用者)
                sMailMsg &= "sExToString：" & sExToString & vbCrLf '(本機)錯誤訊息儲存
                sMailMsg &= "寄件日期：" & cls_test.GlobalMailDate & vbCrLf
                sMailMsg &= "寄件數量：" & iSendMailCount & vbCrLf
                '置換換行符號 'sMailMsg = Replace(sMailMsg, vbCrLf, "<br>" & vbCrLf)
                flagErrMail = True
                Call TIMS.SendMailTest(sMailMsg, "Y", cls_test.gCst_MaxCanMailCount * 3)
            Catch ex As Exception
            End Try
            'Exit Sub
        End If

        '沒有寄信，且有錯誤，再寄一次信
        If Not flagErrMail AndAlso strErrmsg <> "" Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(MyPage, oConn)

            'Common.MessageBox(Me, strErrmsg)
            '置換換行符號
            Try
                Dim sMailMsg As String = ""
                Dim iSendMailCount As Integer = TIMS.SendMailCount() '目前寄信總數量
                sMailMsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
                sMailMsg &= "rqSig：" & vbCrLf & rqSig & vbCrLf
                If Not MyPage.Session Is Nothing Then sMailMsg &= "Session.SessionID：" & CStr(MyPage.Session.SessionID) & vbCrLf 'End If
                sMailMsg &= "Errmsg：" & Errmsg & vbCrLf
                sMailMsg &= "strErrmsg：" & strErrmsg & vbCrLf
                'sMailMsg &= "sExToString：" & sExToString & vbCrLf
                sMailMsg &= "寄件日期：" & cls_test.GlobalMailDate & vbCrLf
                sMailMsg &= "寄件數量：" & iSendMailCount & vbCrLf
                '置換換行符號 'sMailMsg = Replace(sMailMsg, vbCrLf, "<br>" & vbCrLf)
                flagErrMail = True
                Call TIMS.SendMailTest(sMailMsg, "Y", cls_test.gCst_MaxCanMailCount * 3)
            Catch ex As Exception
            End Try
            'Exit Sub
        End If

        If Not flagErrMail AndAlso Errmsg <> "" Then
            Try
                Dim sMailMsg As String = ""
                Dim iSendMailCount As Integer = TIMS.SendMailCount() '目前寄信總數量
                sMailMsg &= TIMS.GetErrorMsg(MyPage) '取得錯誤資訊寫入
                sMailMsg &= "rqSig：" & vbCrLf & rqSig & vbCrLf
                If Not MyPage.Session Is Nothing Then sMailMsg &= "Session.SessionID：" & CStr(MyPage.Session.SessionID) & vbCrLf 'End If
                sMailMsg &= "Errmsg：" & Errmsg & vbCrLf
                'sMailMsg &= "strErrmsg：" & strErrmsg & vbCrLf
                'sMailMsg &= "sExToString：" & sExToString & vbCrLf
                sMailMsg &= "寄件日期：" & cls_test.GlobalMailDate & vbCrLf
                sMailMsg &= "寄件數量：" & iSendMailCount & vbCrLf
                '置換換行符號 'sMailMsg = Replace(sMailMsg, vbCrLf, "<br>" & vbCrLf)
                flagErrMail = True
                Call TIMS.SendMailTest(sMailMsg, "Y", cls_test.gCst_MaxCanMailCount * 3)
            Catch ex As Exception
            End Try
        End If
    End Sub

    '若有登入資訊，且完整 將可進入TIMS系統
    Sub sUtl_btnSubmit()
        Call TIMS.OpenDbConn(objconn)
        Dim rtnPath As String = Request.FilePath
        Dim rqSig As String = Request("credential") ': rqSig = TIMS.ClearSQM(rqSig)
        Dim rqTxtname As String = Request("txtname") ': rqTxtname = TIMS.ClearSQM(rqTxtname)
        Dim Request_txtname As String = "" 'Me.ViewState("Request_txtname") = ""
        '自然人憑證登入開始1
        'If rqSig = "" Then rqSig = credential.Value
        'If rqSig = "" Then
        '    Common.MessageBox(Me, "憑證資訊不可為空!!")
        '    Exit Sub
        'End If
        If rqSig = "" Then Exit Sub

        'Me.ViewState("Request_txtname") = TIMS.ChangeSQM(UCase(rqTxtname).Trim)
        Request_txtname = TIMS.ChangeSQM(UCase(rqTxtname))
        'Me.ViewState("Request_sig") = rqSig
        If TIMS.CheckInput(Request_txtname) Then
            Common.MessageBox(Me, "請勿嘗試在頁面輸入具有危險性的字元!", rtnPath)
            Exit Sub
        End If

        If TIMS.ChangeSQM(UCase(Me.txtname.Text).Trim) <> TIMS.ChangeSQM(UCase(rqTxtname).Trim) Then
            Common.MessageBox(Me, "請勿嘗試在頁面輸入具有危險性的字元!", rtnPath)
            Exit Sub
        End If

        If TIMS.sUtl_ChkHISTORY1(Me, objconn) Then
            Dim strErrmsg1 As String = ""
            strErrmsg1 = ""
            strErrmsg1 = "防止駭客攻擊(紀錄) 啟動-MOICA2_Login.sUtl_btnSubmit" & vbCrLf
            strErrmsg1 &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg1 &= "GetHTTP_HOST:" & vbCrLf & TIMS.GetHTTP_HOST(Me) & vbCrLf
            strErrmsg1 = Replace(strErrmsg1, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.SendMailTest(strErrmsg1)

            Common.MessageBox(Me, TIMS.cst_ErrorMsg1)
            Exit Sub
        End If

        Dim strErrmsg As String = "" '本機錯誤訊息(提供給使用者)
        Dim Errmsg As String = "" '伺服器有回應的錯誤
        Dim HolderID As String = ""
        'http://163.29.199.217/moica/Service1.asmx
        'http://moicahost/moica/Service1.asmx
        'Dim mo As New moica.Service1
        'http://localhost/HiPKIDemoRoot2/Service2.asmx
        'http://localhost/moica2/Service2.asmx
        'http://moicahost/moica2/Service2.asmx
        Dim mo As New moica2.Service2
        'Session(TIMS.cst_MOICA_Login) = "xxx"
        'Common.MessageBox(Me, "讀取到SessionID:" & CStr(Session.SessionID))
        'Common.MessageBox(Me, "讀取到sig:" & CStr(Me.Request("sig")))
        'Exit Sub
        Dim blnB64 As Boolean = TIMS.Chk_Base64Decode1(rqSig)
        If Not blnB64 Then
            Common.MessageBox(Me, "網路連線有誤!!!請關閉瀏覽器，重新操作登入!!!謝謝")
            Exit Sub
        End If
        Dim blnCA1 As Boolean = TIMS.CheckABC123(CStr(Session.SessionID))
        If Not blnCA1 Then
            Common.MessageBox(Me, "網路連線有誤!!!請關閉瀏覽器，重新操作登入!!!謝謝")
            Exit Sub
        End If
        If blnB64 AndAlso blnCA1 AndAlso rqSig <> "" Then
            '憑證資料解析
            Dim ds As DataSet
            Try
                'http://163.29.199.236/moica2/service2.asmx
                'Dim sMessage As String = ""
                'sMessage = ""
                'sMessage &= "." & vbCrLf
                'sMessage &= "http://163.29.199.217/moica2/service2.asmx/GetMoicaData2"
                'sMessage &= "?credential=" & rqSig
                'sMessage &= "&nonce=" & CStr(Session.SessionID)
                'writeLog(sMessage)
                ds = mo.GetMoicaData(rqSig, CStr(Session.SessionID))
                'writeLog("", Me.ViewState("Request_txtname"))
                If ds.Tables.Count > 0 Then
                    Dim dt1 As DataTable = ds.Tables(0)
                    If dt1.Rows.Count = 0 Then Exit Sub
                    Errmsg = ds.Tables(0).Rows(0)("Errmsg")
                    HolderID = ds.Tables(0).Rows(0)("HolderID")
                    Session("serialno") = ds.Tables(0).Rows(0)("serialno")
                    'If Errmsg <> "" Then sExToString &= "Errmsg:" & Errmsg & vbCrLf
                End If

            Catch ex As Exception
                Common.MessageBox(Me, ex.Message.ToString)
                'Common.MessageBox(Me, ex.ToString)

                strErrmsg = "MOICA2_Login!!" & vbCrLf
                strErrmsg &= cst_errMsg1 & vbCrLf
                strErrmsg &= "ex.Message:" & ex.Message & vbCrLf
                strErrmsg &= "ex.ToString:" & ex.ToString & vbCrLf
                strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
                strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
                Call TIMS.SendMailTest(strErrmsg)

                Exit Sub
                '(本機)錯誤訊息儲存
                sExToString = ex.ToString & vbCrLf
                If Errmsg <> "" Then sExToString &= "Errmsg:" & Errmsg & vbCrLf

                strErrmsg = ""
                strErrmsg += cst_errMsg1
                strErrmsg += cst_errMsg2
                If Errmsg <> "" Then strErrmsg += "Errmsg:" & Errmsg & vbCrLf
                'strErrmsg += ex.ToString & vbCrLf
                'Common.MessageBox(Me, strErrmsg)
            End Try
        Else
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)

            strErrmsg = ""
            strErrmsg += cst_errMsg1
            strErrmsg += cst_errMsg2
            strErrmsg += "計算 sig 有誤!!" & vbCrLf
            'Common.MessageBox(Me, strErrmsg)
        End If

        mo = Nothing

        Call TIMS.OpenDbConn(objconn) '開啟連線。
        Dim htSS As New Hashtable
        htSS.Add("rqSig", rqSig)
        htSS.Add("Errmsg", Errmsg)
        htSS.Add("strErrmsg", strErrmsg)
        htSS.Add("sExToString", sExToString)
        '錯誤寄信提醒
        Call sUtl_MailError1(Me, objconn, htSS)

        If Errmsg <> "" Then
            '伺服器有回應的錯誤(提供給使用者)
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If
        If strErrmsg <> "" Then
            '本機錯誤訊息(提供給使用者)
            Common.MessageBox(Me, strErrmsg)
            Exit Sub
        End If

        Dim oRan As New Random
        Thread.Sleep(oRan.Next(oRan.Next(1, 2) * 1000))

        Dim odt As New DataTable
        '測試資料庫連線狀況1(並取得該身份證號):
        Try
            Dim objstr As String = "SELECT * FROM AUTH_ACCOUNT WHERE IDNO=@IDNO AND ISUSED='Y'"
            Dim sCmd As New SqlCommand(objstr, objconn)
            Call TIMS.OpenDbConn(objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("IDNO", SqlDbType.VarChar).Value = TIMS.ChangeSQM(UCase(Me.txtname.Text).Trim)
                odt.Load(.ExecuteReader())
            End With

        Catch ex As Exception
            odt = Nothing
            Dim strScript As String = ""
            strScript = "<script>alert('資料庫或網路暫時性異常，無法連線到資料庫\n\n稍待一下，請再重試\n\n若持續出現此問題，請連絡系統管理人員!!!謝謝');</script>"
            'Page.RegisterStartupScript("wrong", Me.ViewState("strScript").ToString)
            Page.ClientScript.RegisterStartupScript(Page.GetType, "wrong", strScript)
            Exit Sub

        End Try
        If odt Is Nothing Then
            Common.MessageBox(Me, "網路連線有誤!!!請關閉瀏覽器，重新操作登入!!!謝謝", rtnPath)
            Exit Sub
        End If

        If odt.Rows.Count <> 1 Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            'mo = Nothing
            'Common.AddClientScript(Page, "alert('系統無此人員身分證號碼，請洽系統管理員!!!');")
            'Common.AddClientScript(Page, "form1.txtname.value = '';")
            'Exit Sub
            Common.MessageBox(Me, "系統無此人員身分證號碼，請洽系統管理員!!!")
            Exit Sub
        End If

        '判斷IDNO 取得使用者資訊
        'If odt.Rows.Count = 1 Then
        Dim objrow As DataRow = odt.Rows(0)
        ' 20080805 andy 判斷是否超過三個月未登入
        '-----------------------------------------------
        Dim ReturnMsg As String
        ReturnMsg = "登入帳號超過三個月未登入已暫時停權，請洽系統管理員!!!"
        Try
            '檢查是否帳號已超過三個月未登入
            If TIMS.Check_AccoutLoginDate(Me, objrow("Account").ToString, ReturnMsg, objconn) = False Then
                'mo = Nothing
                Exit Sub
            End If
        Catch ex As Exception
            Dim strScript As String = ""
            strScript = "<script>alert('資料庫或網路暫時性異常，無法連線到資料庫\n\n稍待一下，請再重試\n\n若持續出現此問題，請連絡系統管理人員!!!謝謝');</script>"
            'Page.RegisterStartupScript("wrong", Me.ViewState("strScript").ToString)
            Page.ClientScript.RegisterStartupScript(Page.GetType, "wrong", strScript)
            Exit Sub
        End Try
        '------------------------------------------------
        If Convert.ToString(objrow("IsUsed")) <> "Y" Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)
            'writeLog("失敗 ", Me.ViewState("Request_txtname").Trim)
            'Page.RegisterStartupScript("Error", "<script>alert('此帳號不啟用,請洽系統管理者!');</script>")
            'mo = Nothing
            Page.ClientScript.RegisterStartupScript(Page.GetType, "Error", "<script>alert('此帳號不啟用,請洽系統管理者!');</script>")
            Exit Sub
        End If

        Session("IDNO") = UCase(Me.txtname.Text.Trim) '身份證字號(重要記載)
        If Convert.ToString(Session("IDNO")).Length <= 6 Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)

            Common.MessageBox(Me, "輸入身分證號有誤!!")
            Exit Sub
        End If
        If Convert.ToString(Session("IDNO")).Substring(6) <> HolderID Then
            '防止駭客攻擊(紀錄)
            Call TIMS.sUtl_SaveLoginData1(Me, objconn)

            Common.MessageBox(Me, "輸入身分證號與自然人憑證資料不相符!!")
            Exit Sub
        End If
        Session("AuthCookie") = "Linus"
        FormsAuthentication.SetAuthCookie(Session("AuthCookie"), False)
        Call TIMS.Utl_Redirect(Me, objconn, cst_LoginPage1) '轉移login.aspx
    End Sub

    '(隱藏按鈕btnSubmit2)
    Protected Sub btnSubmit2_Click(sender As Object, e As EventArgs) Handles btnSubmit2.Click
        Call sUtl_btnSubmit()
    End Sub
End Class
