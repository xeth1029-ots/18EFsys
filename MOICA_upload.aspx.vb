Public Class upload
    Inherits System.Web.UI.Page

    Private logger As ILog = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)
    Public sm As SessionModel

    Dim sIDNO As String
    Dim sSerialNumber As String
    Dim clickExec As Boolean = False

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)

        'If clickExec Then
        '    Dim resultMsg As String=sm.LastResultMessage
        '    Dim errorMsg As String=sm.LastErrorMessage
        '    Dim redirUrl As String=sm.RedirectUrlAfterBlock

        '    Me.Lit_LastResultMessage.Text=resultMsg
        '    Me.Lit_LastErrorMessage.Text=errorMsg
        '    Me.Lit_RedirectUrlAfterBlock.Text=redirUrl
        'End If

    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        sm = SessionModel.Instance()
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Return ' Exit Sub

        sIDNO = If(Session("MOICA_IDNO") IsNot Nothing, Convert.ToString(Session("MOICA_IDNO")), String.Empty)
        sSerialNumber = If(Session("MOICA_SerialNumber") IsNot Nothing, Convert.ToString(Session("MOICA_SerialNumber")), String.Empty)

        Dim flag_NG As Boolean = (String.IsNullOrEmpty(sIDNO) OrElse String.IsNullOrEmpty(sSerialNumber))
        If (TIMS.sUtl_ChkTest()) Then flag_NG = False '測試
        If flag_NG Then
            ' 自然人憑證綁定 Session 值不存在, 不允許操作, 導向登入頁
            sm.LastErrorMessage = "請依正常動線操作!"
            Response.Redirect(ResolveUrl("~/MOICA_login"))
            Return 'Exit Sub
        End If

        'If Not IsPostBack Then

        'End If

        'Me.Lit_LastResultMessage.Text=sm.LastResultMessage
        'Me.Lit_LastErrorMessage.Text=sm.LastErrorMessage
        'Me.Lit_RedirectUrlAfterBlock.Text=sm.RedirectUrlAfterBlock
    End Sub

    '輸入帳號 確認送出
    Protected Sub bt_submit_Click(sender As Object, e As System.EventArgs) Handles bt_submit.Click

        clickExec = True

        Dim sm As SessionModel = SessionModel.Instance()

        Dim sTxtIdUnchange As String = TIMS.ClearSQM(txtUserId.Text)
        Dim sTxtId As String = ""
        Dim sTxtPass As String = ""
        Dim sVCode As String = ""

        sIDNO = If(Session("MOICA_IDNO") IsNot Nothing, Convert.ToString(Session("MOICA_IDNO")), String.Empty)
        sSerialNumber = If(Session("MOICA_SerialNumber") IsNot Nothing, Convert.ToString(Session("MOICA_SerialNumber")), String.Empty)

        If String.IsNullOrEmpty(sIDNO) OrElse String.IsNullOrEmpty(sSerialNumber) Then
            ' 自然人憑證綁定 Session 值不存在, 不允許操作, 導向登入頁
            sm.LastErrorMessage = "請依正常動線操作!"
            Response.Redirect(ResolveUrl("~/MOICA_login"))
            Exit Sub
        End If

        If Not String.IsNullOrWhiteSpace(sTxtIdUnchange) Then
            sTxtId = UCase(sTxtIdUnchange)
        Else
            sm.LastErrorMessage = "請輸入帳號/密碼"
            Return
        End If

        If Me.txtUserPass.Text <> "" Then
            sTxtPass = Me.txtUserPass.Text
        Else
            sm.LastErrorMessage = "請輸入帳號/密碼"
            Return
        End If

        If txtVCode.Text <> "" Then txtVCode.Text = txtVCode.Text.Trim
        If Me.txtVCode.Text <> "" Then
            sVCode = Me.txtVCode.Text
        Else
            sm.LastErrorMessage = "請輸入驗證碼"
            Return
        End If

        If sVCode <> sm.LoginValidateCode Then
            sm.LastErrorMessage = "驗證碼不正確！"
            Return
        End If

        Dim dr As DataRow = Nothing
        Try
            dr = sUtl_GetAccount(sTxtId)
        Catch ex As Exception
            Throw New Exception("GetAccount 發生錯誤:" & ex.Message, ex)
        End Try

        If dr Is Nothing Then
            sm.LastErrorMessage = "帳號或密碼錯誤！"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Return
        End If

        'pXssXXrd
        If Convert.ToString(dr.Item("HASHPWD1")) <> TIMS.CreateHash(sTxtPass) Then
            sm.LastErrorMessage = "帳號或密碼錯誤！"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        If Convert.ToString(dr.Item("IsUsed")) <> "Y" Then
            sm.LastErrorMessage = "該帳號尚未啟用！"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        If Convert.ToString(dr.Item("Stopmsg")) = "Y" Then '過了帳號停用日
            sm.LastErrorMessage = "帳號已被停用，請洽詢系統管理者！"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        Dim flag_check_idno_ng_1 As Boolean = False '身分證 檢查有誤true 無誤false (default)
        If Not flag_check_idno_ng_1 AndAlso Convert.IsDBNull(dr.Item("IDNO")) Then flag_check_idno_ng_1 = True

        If Not flag_check_idno_ng_1 AndAlso sIDNO <> Convert.ToString(dr.Item("IDNO")) Then flag_check_idno_ng_1 = True

        If flag_check_idno_ng_1 Then
            sm.LastErrorMessage = "此帳號登記的身分證號跟您的不一樣，無法綁定，請洽詢系統管理者！"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If

        'If Convert.ToString(dr.Item("SERIALNO")) <> "" AndAlso Convert.ToString(dr.Item("SERIALNO")) <> sSerialNumber Then
        '    sm.LastErrorMessage="此帳號已有綁定其他自然人憑證，請洽詢系統管理者！"
        '    AuthUtil.LoginLog(sTxtIdUnchange, False)
        '    Exit Sub
        'End If

        ' 判斷是否超過三個月未登入
        '-----------------------------------------------
        Dim ReturnMsg As String = Nothing
        Dim flag3 As Boolean = TIMS.Check_AccoutLoginDate(sm, dr("Account"), ReturnMsg, objconn)
        If Not flag3 Then
            sm.LastErrorMessage = ReturnMsg
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If
        '------------------------------------------------
        Dim iRoleID As Integer = TIMS.GetValue2(dr.Item("RoleID"), -1)
        If iRoleID <> -1 Then
            ' 帳密驗證且有系統權限，綁定憑證序號並完成登入
            ' 這二個 Session 變數, 已無作用, 清除
            Session("MOICA_IDNO") = ""
            Session("MOICA_SerialNumber") = ""

            If Convert.ToString(dr.Item("SERIALNO")) <> "" Then
                If Convert.ToString(dr.Item("SERIALNO")) <> sSerialNumber Then
                    '更新綁定(序號不相同) 相同不更新
                    sUtl_UpdateAccountSerialNum(Convert.ToString(dr("Account")), sIDNO, sSerialNumber)
                End If
            Else
                '更新綁定(無序號)
                sUtl_UpdateAccountSerialNum(Convert.ToString(dr("Account")), sIDNO, sSerialNumber)
            End If

            ' 設定登入資訊
            Dim userInfo As LoginUserInfo = New LoginUserInfo()
            userInfo.UserID = Convert.ToString(dr("Account"))
            userInfo.UserName = Convert.ToString(dr("Name"))
            userInfo.RoleID = Convert.ToString(dr("RoleID"))
            userInfo.OrgID = Convert.ToString(dr("OrgID"))
            userInfo.OrgName = Convert.ToString(dr("OrgName"))
            userInfo.OrgLevel = Convert.ToInt16(dr("OrgLevel"))
            userInfo.DistID = Convert.ToString(dr("DistID"))
            userInfo.LID = Convert.ToInt16(dr("LID"))
            userInfo.RID = Convert.ToString(dr("RID"))
            userInfo.RelShip = Convert.ToString(dr("Relship"))
            'userInfo.Years=Convert.ToString(dr("DEFAULT_YEAR"))
            userInfo.DefaultPlanID = Convert.ToString(dr("DEFAULT_PLANID"))
            userInfo.DefaultYear = Convert.ToString(dr("DEFAULT_YEAR"))

            ' 登入成功 Flag (重要, 有設定才是真的登入成功)
            userInfo.LoginSuccess = True

            ' 登入使用者資訊保存到 SessionModel
            sm.UserInfo = userInfo

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
                    ' 設定 年度/計畫 成功, 導向登入後首頁
                    Response.Redirect(ResolveUrl("~/Index"))
                Else
                    DbAccess.CloseDbConn(objconn)
                    ' 預設的 年度/計劃 失敗, 導向計畫選擇頁面
                    sm.LastErrorMessage = Nothing   ' 忽略預設 年度/計劃 的錯誤訊息
                    Response.Redirect(ResolveUrl("~/SelectPlan"))
                End If
            End If
        Else
            'RoleID:-1  使用者 無此功能
            sm.LastErrorMessage = "帳號沒有系統使用權限，無法綁定，請洽詢系統管理者!!!"
            AuthUtil.LoginLog(sTxtIdUnchange, False)
            Exit Sub
        End If
    End Sub

    ''' <summary>
    ''' 更新綁定憑證序號
    ''' </summary>
    ''' <param name="sUserID"></param>
    ''' <param name="sIDNO"></param>
    ''' <param name="sSerialNum"></param>
    Sub sUtl_UpdateAccountSerialNum(ByVal sUserID As String, ByVal sIDNO As String, ByVal sSerialNum As String)
        Dim old_sSerialNum As String = ""
        Dim dr1 As DataRow = Nothing

        TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql &= " SELECT SERIALNO FROM AUTH_ACCOUNT" & vbCrLf
        sql &= " WHERE UPPER(ACCOUNT)=@account And UPPER(IDNO)=@idno And LEN(SERIALNO)>1" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("account", SqlDbType.VarChar).Value = sUserID
            .Parameters.Add("idno", SqlDbType.VarChar).Value = sIDNO
            dt.Load(.ExecuteReader())
            If dt.Rows.Count > 0 Then dr1 = dt.Rows(0)
        End With
        If dr1 IsNot Nothing Then old_sSerialNum = TIMS.ClearSQM(dr1("SERIALNO"))

        If sSerialNum <> "" AndAlso old_sSerialNum <> "" AndAlso sSerialNum <> old_sSerialNum Then
            '有舊序號，且不等同新序號
            Dim u_sql As String = ""
            u_sql &= " UPDATE AUTH_ACCOUNT" & vbCrLf
            u_sql &= " SET STOP_SERIALNO=@STOP_SERIALNO" & vbCrLf
            u_sql &= " ,MODIFYACCT=@MODIFYACCT ,MODIFYDATE=GETDATE()" & vbCrLf
            u_sql &= " WHERE UPPER(ACCOUNT)=@account And UPPER(IDNO)=@idno" & vbCrLf
            Dim u_parms As Hashtable = New Hashtable()
            u_parms.Clear()
            u_parms.Add("STOP_SERIALNO", old_sSerialNum)
            u_parms.Add("MODIFYACCT", sUserID)
            u_parms.Add("account", UCase(sUserID))
            u_parms.Add("idno", UCase(sIDNO))
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        End If

        If sSerialNum <> "" AndAlso sSerialNum <> old_sSerialNum Then
            '不管有無舊序號，只要不相同就更新
            Dim u_sql As String = ""
            u_sql &= " Update AUTH_ACCOUNT" & vbCrLf
            u_sql &= " SET SERIALNO=@serialno" & vbCrLf
            u_sql &= " WHERE UPPER(ACCOUNT)=@account And UPPER(IDNO)=@idno" & vbCrLf
            Dim u_parms As Hashtable = New Hashtable()
            u_parms.Add("serialno", sSerialNum)
            u_parms.Add("account", UCase(sUserID))
            u_parms.Add("idno", UCase(sIDNO))
            DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
        End If
    End Sub

    ''' <summary>
    ''' 取得帳號資料, 若找不到回傳 nothing
    ''' </summary>
    ''' <param name="sTxtUserID">帳號</param>
    ''' <returns></returns>
    Function sUtl_GetAccount(ByVal sTxtUserID As String) As DataRow
        sTxtUserID = UCase(TIMS.ClearSQM(sTxtUserID))

        Dim parms As New Hashtable From {{"account", sTxtUserID}}
        Dim sql As String = ""
        sql &= " SELECT a.ACCOUNT, a.IDNO, a.SERIALNO, a.ROLEID, a.LID, a.HASHPWD1" & vbCrLf
        sql &= " ,a.NAME, a.ISUSED, a.DEFAULT_PLANID, a.DEFAULT_YEAR" & vbCrLf
        sql &= " ,c.RID,c.RELSHIP, c.ORGLEVEL, c.DISTID, d.ORGID, d.ORGNAME" & vbCrLf
        sql &= " ,CASE WHEN GETDATE()>=STOPDATE Then 'Y' ELSE 'N' END STOPMSG" & vbCrLf
        sql &= " FROM AUTH_ACCOUNT a" & vbCrLf
        sql &= " JOIN AUTH_RELSHIP c ON a.ORGID=c.ORGID" & vbCrLf
        sql &= " JOIN ORG_ORGINFO d ON c.ORGID=d.ORGID" & vbCrLf
        sql &= " WHERE UPPER(a.ACCOUNT)=@account" & vbCrLf

        Dim odt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        '取得1筆資料。
        If TIMS.dtNODATA(odt) Then Return Nothing
        Dim rst As DataRow = odt.Rows(0)
        Return rst
    End Function

End Class