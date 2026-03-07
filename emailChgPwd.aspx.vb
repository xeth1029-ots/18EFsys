Public Class emailChgPwd
    Inherits System.Web.UI.Page

    'Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    Dim sTxtIdUnchange As String = "" 'TIMS.ClearSQM(txtUserId.Text)
    Dim sTxtPass As String = "" 'TIMS.ClearSQM(txtUserPass.Text)
    Dim sTxtPass2 As String = "" 'TIMS.ClearSQM(txtUserPass2.Text)
    Dim sVCode As String = "" 'TIMS.ClearSQM(txtVCode.Text)
    'Dim s_GID As String = ""
    'Dim s_OTK As String = ""
    'Dim vSENDMAIL As String = ""
    'Dim vHASH1 As String = ""

    '密碼設定原則：12~16碼，須包含英文大、小寫、數字及符號的組合
    Const cst_errorMsg4e As String = "輸入新密碼，不得使用身分證字號!"
    Const cst_errorMsg7b As String = "密碼 具有危險性的字元，請重新設定!"
    Const cst_errorMsg18b As String = "密碼 請輸入12~16碼(限定數字.英文.符號)！"
    Const cst_errorMsg19 As String = "密碼 應包含大寫英文字母！"
    Const cst_errorMsg20 As String = "密碼 應包含小寫英文字母！"
    Const cst_errorMsg21 As String = "密碼 應包含數字字元！"
    Const cst_errorMsg21s As String = "密碼 應包含符號字元！"
    'Const cst_errorMsg21ns As String = "密碼 應包含數字或符號字元！"

    Const cst_errorMsg22 As String = "請輸入確認密碼"
    Const cst_errorMsg22a As String = "請輸入新密碼"
    Const cst_errorMsg22b As String = "請輸入原密碼"

    Const cst_errorMsg23 As String = "輸入確認密碼 與密碼 不同"
    Const cst_errorMsg23b As String = "3代密碼不得相同！"
    Const cst_errorMsg23c As String = "輸入原密碼 有誤!!"

    'Dim g_parms As Hashtable
    Dim g_parms2 As Hashtable
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

    '<summary>
    '資安檢核
    ' </summary>
    'Sub Critical_Issues_1()
    '    'openhttps.Value = TIMS.ClearSQM(openhttps.Value)
    '    g_parms = New Hashtable
    '    g_parms.Clear()
    '    g_parms.Add("1.txtUserId.Text", txtUserId.Text)
    '    g_parms.Add("2.txtUserPass.Text", txtUserPass.Text)
    '    g_parms.Add("3.txtUserPass2.Text", txtUserPass2.Text)
    '    g_parms.Add("4.txtVCode.Text", txtVCode.Text)

    '    sTxtIdUnchange = TIMS.ClearSQM(txtUserId.Text)
    '    sTxtPass = TIMS.ClearSQM(txtUserPass.Text)
    '    sTxtPass2 = TIMS.ClearSQM(txtUserPass2.Text)
    '    sVCode = TIMS.ClearSQM(txtVCode.Text)

    '    txtUserId.Text = sTxtIdUnchange
    '    txtUserPass.Text = sTxtPass
    '    txtUserPass2.Text = sTxtPass2
    '    txtVCode.Text = sVCode

    '    g_parms.Add("sTxtIdUnchange", sTxtIdUnchange)
    '    g_parms.Add("sTxtPass", sTxtPass)
    '    g_parms.Add("sTxtPass2", sTxtPass2)
    '    g_parms.Add("sVCode", sVCode)
    '    '"https://ojtims.wda.gov.tw/emailChgPwd?GID={0}&OTK={1}&OTK2={2}"
    '    'TIMS.sUtl_404NOTFOUND(Page, objconn)
    'End Sub

    ''' <summary>
    ''' 資安檢核-2
    ''' </summary>
    ''' <returns></returns>
    Function Critical_Issues_2() As Boolean
        Dim rst As Boolean = False '取得資料失敗為false
        g_parms2 = New Hashtable
        g_parms2.Clear()

        Dim dr As DataRow = Nothing

        divOrgUserPass.Visible = False
        '有登入資訊可顯示-輸入原密碼
        If sm.IsLogin Then divOrgUserPass.Visible = True

        '登入狀況
        Dim s_IsLogin As String = "Y"
        If Not sm.IsLogin Then
            '未登入--要有request 值
            s_IsLogin = "N"
            Dim s_GID As String = TIMS.ClearSQM(Request("GID"))
            Dim s_OTK As String = TIMS.ClearSQM(Request("OTK"))
            If s_GID = "" OrElse s_OTK = "" Then Return rst
            'HSEQ,HASH1
            Dim parms As New Hashtable
            parms.Clear()
            parms.Add("HSEQ", s_GID)
            parms.Add("HASH1", s_OTK)
            '取得外部 序號 傳入-檢核1
            Dim flag_chk_ok As Boolean = TIMS.CHK_HASHLOG(objconn, parms)
            If Not flag_chk_ok Then Return rst
            dr = TIMS.GET_drPXSSWARDHIS(objconn, parms) '取得內部資訊

        Else
            '已登入
            Dim drAA As DataRow = TIMS.sUtl_GetAccount(sm.UserInfo.UserID, objconn, False)
            If drAA Is Nothing Then Return rst '取得資料失敗為false

            Hid_HASHPXWXD1.Value = TIMS.ClearSQM(drAA("HASHPWD1"))
            'Dim vSENDMAIL As String = TIMS.ClearSQM(drAA("EMAIL"))
            Dim s_parma As New Hashtable
            s_parma.Clear()
            s_parma.Add("ACCOUNT", sm.UserInfo.UserID)
            s_parma.Add("SENDMAIL", "") '不提供EMAIL 不寄信
            '/*密碼寄信種類-1:新設密碼-2:忘記密碼(修改密碼)-3:修改密碼 */ SENDMAILTYPE /SDMATYPE
            s_parma.Add("SDMATYPE", "2")
            s_parma.Add("MODIFYACCT", sm.UserInfo.UserID)
            Dim htSS As Hashtable = TIMS.INSERT_PXSSWARDHIS(objconn, s_parma)
            Dim vHSEQ As String = TIMS.GetMyValue2(htSS, "HSEQ")
            Dim vHASH1 As String = TIMS.GetMyValue2(htSS, "HASH1")

            Dim parms As New Hashtable
            parms.Clear()
            parms.Add("HSEQ", vHSEQ)
            parms.Add("HASH1", vHASH1)
            '取得外部/內部 序號 傳入-檢核1
            Dim flag_chk_ok As Boolean = TIMS.CHK_HASHLOG(objconn, parms)
            If Not flag_chk_ok Then Return rst
            dr = TIMS.GET_drPXSSWARDHIS(objconn, parms) '取得內部資訊

        End If
        If dr IsNot Nothing Then
            g_parms2.Add("ISLOGIN", s_IsLogin)
            g_parms2.Add("PXSEQ", Convert.ToString(dr("PXSEQ")))
            g_parms2.Add("ACCOUNT", Convert.ToString(dr("ACCOUNT")))
            g_parms2.Add("HSEQ", Convert.ToString(dr("HSEQ")))
            g_parms2.Add("HASH1", Convert.ToString(dr("HASH1")))
            g_parms2.Add("SENDMAIL", Convert.ToString(dr("SENDMAIL")))
            g_parms2.Add("SENDMAILDATE", Convert.ToString(dr("SENDMAILDATE")))
            rst = True 'true 取得資料 成功
        End If

        Return rst '取得資料失敗為false
    End Function

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub

        sm = SessionModel.Instance()

        BaseUrl = ResolveUrl("~/")

        'If Not String.IsNullOrEmpty(BaseUrl) AndAlso Not BaseUrl.EndsWith("/") Then BaseUrl = String.Concat(BaseUrl, "/")

        'txtUserId.Enabled = False
        txtUserId.Enabled = False

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)
        If Not IsPostBack Then

            If Not sm.IsLogin Then
                ' 登出/重登
                AuthUtil.LogoutLog()
                ' 清除登入狀態
                sm.ClearSession()
            End If

            If Not Critical_Issues_2() Then TIMS.sUtl_404NOTFOUND(Me, objconn)

            txtUserId.Text = TIMS.GetMyValue2(g_parms2, "ACCOUNT")
        End If

    End Sub

    ''' <summary>
    ''' 檢核3代密碼是否相同 :true(有相同) /false:沒有相同
    ''' </summary>
    ''' <returns></returns>
    Function CHK_3TPXWA(ByRef dt1 As DataTable, ByVal vPWD1 As String) As Boolean
        Dim rst As Boolean = False
        If dt1 Is Nothing Then Return rst
        If dt1.Rows.Count = 0 Then Return rst
        For Each dr1 As DataRow In dt1.Rows
            If Convert.ToString(dr1("HASHPWD1")).Equals(TIMS.CreateHash(vPWD1)) Then
                'vSENDMAIL = Convert.ToString(dr1("SENDMAIL"))
                rst = True
                Return rst
            End If
        Next
        Return rst
    End Function

    ''' <summary>檢核1-密碼</summary>
    ''' <param name="sErrMsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef sErrMsg As String) As Boolean
        Dim rst As Boolean = False '異常為false
        '請輸入帳號
        '帳號請輸入數字或英文字
        'userpass
        '請輸入密碼
        '密碼請輸入12~14碼(限定數字.英文)
        'https://jira.turbotech.com.tw/browse/TIMSC-237
        Dim flagOK1 As Boolean = False 'false:證號檢核失敗
        flagOK1 = False 'Dim flagOK1 As Boolean = False

        Dim oOrgUserPassTxt As String = txtOrgUserPass.Text
        Dim oUserpassTxt As String = txtUserPass.Text
        Dim oUserpass2Txt As String = txtUserPass2.Text

        'Dim str_userpass As String = oUserpassTxt
        Dim ok_userpass As String = TIMS.ClearSQM(oUserpassTxt)
        Dim ok_OrgUserPass As String = TIMS.ClearSQM(oOrgUserPassTxt)

        If divOrgUserPass.Visible Then
            If oOrgUserPassTxt = "" Then
                sErrMsg &= cst_errorMsg22b & vbCrLf
                Return rst
            End If
            If Hid_HASHPXWXD1.Value <> TIMS.CreateHash(oOrgUserPassTxt) Then
                sErrMsg &= cst_errorMsg23c & vbCrLf
                Return rst
            End If
        End If
        If oUserpassTxt = "" Then
            sErrMsg &= cst_errorMsg22a & vbCrLf
            Return rst
        End If
        If oUserpass2Txt = "" Then
            sErrMsg &= cst_errorMsg22 & vbCrLf
            Return rst
        End If

        If oUserpassTxt <> ok_userpass Then
            oUserpassTxt = TIMS.ClearSQM(oUserpassTxt)
            sErrMsg &= "密碼" & TIMS.cst_ErrorMsg10 & vbCrLf
            Return rst
        End If

        If oUserpassTxt <> "" Then
            If oUserpassTxt.Length >= 12 AndAlso oUserpassTxt.Length <= 16 Then flagOK1 = True
            If Not flagOK1 Then sErrMsg &= cst_errorMsg18b & vbCrLf

            flagOK1 = False
            If TIMS.ChkUpper(oUserpassTxt) Then flagOK1 = True
            If Not flagOK1 Then sErrMsg &= cst_errorMsg19 & vbCrLf

            flagOK1 = False
            If TIMS.ChkLower(oUserpassTxt) Then flagOK1 = True
            If Not flagOK1 Then sErrMsg &= cst_errorMsg20 & vbCrLf

            flagOK1 = False
            If TIMS.ChkNumber(oUserpassTxt) Then flagOK1 = True
            If Not flagOK1 Then sErrMsg &= cst_errorMsg21 & vbCrLf

            flagOK1 = False
            If TIMS.ChkSymbol(oUserpassTxt) Then flagOK1 = True
            If Not flagOK1 Then sErrMsg &= cst_errorMsg21s & vbCrLf

            'flagOK1 = False
            'If TIMS.ChkNumber(oUserpassTxt) OrElse TIMS.ChkSymbol(oUserpassTxt) Then flagOK1 = True
            'If Not flagOK1 Then sErrMsg &= cst_errorMsg21ns & vbCrLf

            If sErrMsg <> "" Then Return rst

            'userpass2
            '請輸入確認密碼
            'If oUserpass2Txt = "" Then
            '    sErrMsg &= cst_errorMsg22 & vbCrLf
            '    Return rst
            'End If

            '確認密碼輸入錯誤
            If oUserpassTxt <> oUserpass2Txt Then
                sErrMsg &= cst_errorMsg23 & vbCrLf
                'Return rst
            End If
            '確認密碼請輸入12~14碼(限定數字.英文)

            Dim V_userpass As String = oUserpassTxt
            oUserpassTxt = TIMS.ClearSQM(oUserpassTxt) '密碼
            If oUserpassTxt <> V_userpass Then
                sErrMsg &= cst_errorMsg7b & vbCrLf
                Return rst
            End If

            If TIMS.CheckIDNO(oUserpassTxt) Then
                sErrMsg &= cst_errorMsg4e & vbCrLf
                Return rst
            End If
        End If
        If sErrMsg <> "" Then Return rst

        Dim dt1 As DataTable 'AUTH_PXSSWARD_HIS
        dt1 = TIMS.GET_dtPXSSWARDHIS(objconn, g_parms2)
        If dt1 Is Nothing Then
            sErrMsg &= TIMS.cst_NODATAMsg2
            Return rst
        End If

        '檢核3代密碼是否相同 :true(有相同) /false:沒有相同 AUTH_PXSSWARD_HIS
        If CHK_3TPXWA(dt1, oUserpassTxt) Then
            sErrMsg &= cst_errorMsg23b & vbCrLf
            Return rst
        End If

        If sErrMsg <> "" Then Return rst
        Return True
    End Function

    '修改密碼-儲存 '輸入帳號 登入送出
    Protected Sub bt_submit_Click(sender As Object, e As System.EventArgs) Handles bt_submit.Click
        'Dim sm As SessionModel = SessionModel.Instance()
        'sm = SessionModel.Instance()

        If Not Critical_Issues_2() Then TIMS.sUtl_404NOTFOUND(Me, objconn)

        Const cst_alert_msg_2b As String = "驗證碼不正確!!!"

        Dim sVCode As String = ""
        sVCode = TIMS.ClearSQM(txtVCode.Text)

        If sVCode <> sm.LoginValidateCode Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            'Common.MessageBox(Me, cst_alert_msg_2b) 'Exit Sub
            sm.LastErrorMessage = cst_alert_msg_2b
            Return
        End If

        Dim sErrMsg As String = ""
        Call CheckData1(sErrMsg)
        If sErrMsg <> "" Then
            'Common.MessageBox(Me, sErrMsg) 'Exit Sub
            Dim jsErrMsg As String = Common.GetJsString(sErrMsg)
            sm.LastErrorMessage = jsErrMsg
            Return
        End If

        Dim vISLOGIN As String = TIMS.GetMyValue2(g_parms2, "ISLOGIN")
        Dim vSENDMAIL As String = TIMS.GetMyValue2(g_parms2, "SENDMAIL")
        Dim vHSEQ As String = TIMS.GetMyValue2(g_parms2, "HSEQ")
        Dim s_parma As New Hashtable
        s_parma.Clear()
        s_parma.Add("ISLOGIN", vISLOGIN)
        s_parma.Add("ACCOUNT", txtUserId.Text)
        s_parma.Add("SENDMAIL", vSENDMAIL)
        s_parma.Add("HSEQ", vHSEQ)
        s_parma.Add("HASHPWD1", TIMS.CreateHash(txtUserPass.Text))
        s_parma.Add("PXSSENC1", TIMS.EncryptAes(txtUserPass.Text))
        '/*SDMATYPE 密碼寄信種類-1:新設密碼-2:忘記密碼(修改密碼)-3:修改密碼 */ SENDMAILTYPE /SDMATYPE
        s_parma.Add("SDMATYPE", "3")
        s_parma.Add("MODIFYACCT", txtUserId.Text)
        '修改密碼-儲存 
        Dim htSS As Hashtable = TIMS.INSERT_PXSSWARDHIS(objconn, s_parma)

        sm.LastResultMessage = "密碼修改完成"

        Dim redirectUrl As String = ResolveUrl("~/login") 'String.Empty
        'redirectUrl = ResolveUrl("~/login")
        If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
            '檢核成功, 導向首頁
            Response.Redirect(redirectUrl)
        End If

        'Call Utl_LoginAgain()
        'Dim flagCanLogin As Boolean = False '確認是否可再登入
        'If flagCanLogin Then
        '    Call Utl_LoginAgain()
        'End If
    End Sub

    ''' <summary>
    ''' 重設
    ''' </summary>
    Sub Utl_reset_1()
        txtOrgUserPass.Text = ""
        txtUserPass.Text = ""
        txtUserPass2.Text = ""
        txtVCode.Text = ""
    End Sub

    ''' <summary>重設</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_reset_Click(sender As Object, e As EventArgs) Handles bt_reset.Click
        Call Utl_reset_1()
    End Sub

    Protected Sub bt_back1_Click(sender As Object, e As EventArgs) Handles bt_back1.Click
        Dim redirectUrl As String = ResolveUrl("~/login")
        Response.Redirect(redirectUrl)
    End Sub
End Class