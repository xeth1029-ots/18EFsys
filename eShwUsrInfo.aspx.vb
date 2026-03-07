Public Class eShwUsrInfo
    Inherits System.Web.UI.Page

    'Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

    'Dim sTxtIdUnchange As String = ""
    'Dim sTxtIDNO As String = ""
    'Dim sTxtEMAIL As String = ""
    'Dim sVCode As String = ""

    'Dim g_parms As Hashtable
    'Dim g_parms2 As Hashtable
    'Public Const cst_ErrorMsg2 As String = "請勿嘗試在頁面輸入具有危險性的字元!"

    Public BaseUrl As String
    Public sm As SessionModel '= SessionModel.Instance()

    Dim oWDAP As New WDAP
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call Critical_Issues_1()
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub

        sm = SessionModel.Instance()

        If sm.UserInfo Is Nothing Then
            '沒有登入, 導向登入頁面
            'logger.Debug(Request.Path & " Login Required, Redirect to LOGIN_PAGE")
            sm.LastErrorMessage = "您尚未登入或登入資訊已經遺失，請重新登入!"
            Dim redirectUrl As String = ResolveUrl("~/MOICA_login")
            Response.Redirect(redirectUrl)
            'TIMS.sUtl_404NOTFOUND(Me, objconn)
            Return
        End If

        '檢查使用者登入狀態資訊
        If Not sm.IsLogin Then
            '沒有登入, 導向登入頁面
            'logger.Debug(Request.Path & " Login Required, Redirect to LOGIN_PAGE")
            sm.LastErrorMessage = "您尚未登入或登入資訊已經遺失，請重新登入!"
            Dim redirectUrl As String = ResolveUrl("~/MOICA_login")
            Response.Redirect(redirectUrl)
            'Response.End()
            Return
        End If

        BaseUrl = ResolveUrl("~/")

        'If Not String.IsNullOrEmpty(BaseUrl) AndAlso Not BaseUrl.EndsWith("/") Then BaseUrl = String.Concat(BaseUrl, "/")

        Labmsg1.Text = TIMS.Get_LOCALADDR(Me, 2)

        If Not IsPostBack Then
            '登出/ 重登
            'AuthUtil.LogoutLog()
            '清除登入狀態
            'sm.ClearSession()
            cCreate1()
        End If

    End Sub

    Sub cCreate1()
        txtOrgName.Enabled = False
        txtUserId.Enabled = False
        txtUserName.Enabled = False
        TIMS.Tooltip(txtOrgName, "不提供修改", True)
        TIMS.Tooltip(txtUserId, "不提供修改", True)
        TIMS.Tooltip(txtUserName, "不提供修改", True)

        If Not sm.IsLogin Then Exit Sub
        Dim drAA As DataRow = TIMS.sUtl_GetAccount(sm.UserInfo.UserID, objconn, False)
        If drAA Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        txtOrgName.Text = TIMS.ClearSQM(drAA("ORGNAME"))
        txtUserId.Text = TIMS.ClearSQM(drAA("ACCOUNT"))
        txtUserName.Text = TIMS.ClearSQM(drAA("NAME"))
        txtUserEMAIL.Text = TIMS.ClearSQM(drAA("EMAIL"))
        txtUserPhone.Text = TIMS.ClearSQM(drAA("PHONE"))
        Hid_acount_1.Value = sm.UserInfo.UserID
    End Sub

    '修改密碼
    Protected Sub bt_submit_Click(sender As Object, e As System.EventArgs) Handles bt_submit.Click
        Dim redirectUrl As String = ResolveUrl("~/emailChgPwd") 'String.Empty

        'redirectUrl = ResolveUrl("~/login")
        If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
            '檢核成功, 導向首頁
            Response.Redirect(redirectUrl)
        End If

    End Sub

    ''' <summary>儲存</summary>
    ''' <returns></returns>
    Function cSaveData1() As Integer
        Dim iRst As Integer = 0

        Dim s_prams As New Hashtable
        s_prams.Clear()
        s_prams.Add("ACCOUNT", UCase(sm.UserInfo.UserID))
        Dim sqlstr As String = ""
        sqlstr = " SELECT 'X' FROM dbo.AUTH_ACCOUNT" & vbCrLf
        sqlstr &= " WHERE UPPER(ACCOUNT)=@ACCOUNT" & vbCrLf
        Dim dtA As DataTable = DbAccess.GetDataTable(sqlstr, objconn, s_prams)
        If dtA.Rows.Count <> 1 Then Return iRst '只能是1才儲存

        '修改動作
        '帳號基本檔
        'Dim sqlstr As String = ""
        Dim u_prams As New Hashtable
        u_prams.Clear()
        'u_prams.Add("NAME", txtUserName.Text)
        u_prams.Add("EMAIL", txtUserEMAIL.Text)
        u_prams.Add("PHONE", If(txtUserPhone.Text <> "", txtUserPhone.Text, Convert.DBNull))
        u_prams.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_prams.Add("ACCOUNT", UCase(sm.UserInfo.UserID))

        sqlstr = " UPDATE AUTH_ACCOUNT" & vbCrLf
        'sqlstr &= " SET NAME=@NAME" & vbCrLf
        sqlstr &= " SET EMAIL=@EMAIL" & vbCrLf
        sqlstr &= " ,PHONE=@PHONE" & vbCrLf
        sqlstr &= " ,LAST_LOGINDATE=GETDATE()" & vbCrLf
        sqlstr &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sqlstr &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        sqlstr &= " WHERE 1=1" & vbCrLf
        sqlstr &= " AND UPPER(ACCOUNT)=@ACCOUNT" & vbCrLf
        iRst = DbAccess.ExecuteNonQuery(sqlstr, objconn, u_prams)
        Return iRst
    End Function

    ''' <summary>儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub bt_save1_Click(sender As Object, e As EventArgs) Handles bt_save1.Click

        'Const cst_errorMsg1 As String = "姓名 必填欄位不可為空!"
        'Const cst_errorMsg2 As String = "E_Mail 必填欄位不可為空!"
        Const cst_errorMsg31 As String = "E_Mail欄位改為必填不可為空白，請輸入!"
        Const cst_errorMsg32 As String = "E_Mail欄位格式有誤，請重新輸入!"
        Const cst_errorMsg33 As String = "電話欄位格式有誤，至少要有數字8碼以上!"
        Const cst_SaveErrMsg13 As String = "帳號資訊 儲存有誤!!!"
        Const cst_alert_msg_2b As String = "驗證碼不正確!!!"
        Const cst_alert_msg_2d As String = "請輸入姓名、E-Mail與驗證碼"

        Dim sVCode As String = ""
        sVCode = TIMS.ClearSQM(txtVCode.Text)
        txtUserName.Text = TIMS.ClearSQM(txtUserName.Text)
        txtUserEMAIL.Text = TIMS.ChangeEmail(TIMS.ClearSQM(txtUserEMAIL.Text)) '轉EMAIL
        If sVCode = "" OrElse txtUserName.Text = "" OrElse txtUserEMAIL.Text = "" Then
            'sm.LastErrorMessage = cst_alert_msg_2c Return
            Common.MessageBox(Me, cst_alert_msg_2d)
            Exit Sub
        End If

        If sVCode <> sm.LoginValidateCode Then
            '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
            Call oWDAP.SUtl_SaveLoginData1(Me, objconn)
            'sm.LastErrorMessage = cst_alert_msg_2b 'Return
            Common.MessageBox(Me, cst_alert_msg_2b)
            Exit Sub
        End If

        'txtUserName.Text = TIMS.ClearSQM(txtUserName.Text)
        'If txtUserName.Text = "" Then
        '    Common.MessageBox(Me, cst_errorMsg1)
        '    Exit Sub
        'End If

        'txtUserEMAIL.Text = TIMS.ChangeEmail(TIMS.ClearSQM(txtUserEMAIL.Text)) '轉EMAIL
        If txtUserEMAIL.Text = "" Then
            Common.MessageBox(Me, cst_errorMsg31)
            Exit Sub
        End If

        If Not TIMS.CheckEmail(txtUserEMAIL.Text) Then
            Common.MessageBox(Me, cst_errorMsg32)
            Exit Sub
        End If

        txtUserPhone.Text = TIMS.ChangeEmail(TIMS.ClearSQM(txtUserPhone.Text))
        If txtUserPhone.Text <> "" AndAlso Not TIMS.ChkPhone123(txtUserPhone.Text) Then
            Common.MessageBox(Me, cst_errorMsg33)
            Exit Sub
        End If

        If Hid_acount_1.Value <> sm.UserInfo.UserID Then
            sm.LastResultMessage = cst_SaveErrMsg13 '"帳號資訊 儲存有誤!!!"
            Common.MessageBox(Me, cst_SaveErrMsg13)
            Exit Sub
        End If
        Dim iRst As Integer = 0
        iRst = cSaveData1()
        If iRst <> 1 Then
            sm.LastResultMessage = cst_SaveErrMsg13 '"帳號資訊 儲存有誤!!!"
            Common.MessageBox(Me, cst_SaveErrMsg13)
            Exit Sub
        End If

        'sm.LastResultMessage = "帳號資訊 儲存完成!"
        '導向計畫選擇頁面
        'sm.RedirectUrlAfterBlock = ResolveUrl("~/SelectPlan")
        'Response.Redirect(ResolveUrl("~/SelectPlan"))
        'Dim redirectUrl As String = ResolveUrl("~/SelectPlan")
        'Dim redirectUrl As String = ResolveUrl("~/index")
        'Response.Redirect(redirectUrl)
        'sm = SessionModel.Instance()
        'sm.LastResultMessage = "帳號資訊 儲存完成!"
        Common.MessageBox(Me, "帳號資訊 儲存完成!")
        Dim redirectUrl As String = ResolveUrl("~/index")
        Response.Redirect(redirectUrl)
        'Common.MessageBox(Me, "帳號資訊 儲存完成!", "~/index")
    End Sub

    Protected Sub bt_back1_Click(sender As Object, e As EventArgs) Handles bt_back1.Click
        'Dim redirectUrl As String = ResolveUrl("~/login")
        Dim redirectUrl As String = ResolveUrl("~/index")
        Response.Redirect(redirectUrl)
    End Sub
End Class
