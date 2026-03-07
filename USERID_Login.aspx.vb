Public Class USERID_Login
    Inherits System.Web.UI.Page

    'Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
    '    'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
    '    '請勿使用程式碼編輯器進行修改。
    '    'InitializeComponent()
    '    Response.Cache.SetCacheability(HttpCacheability.NoCache)
    '    Response.Cache.SetExpires(DateTime.MinValue)
    '    Response.Cache.SetNoStore()
    'End Sub

    Const cst_alertMsg1 As String = " 帳號或密碼錯誤 "
    Const cst_UserID_Login As String = "UserID_Login"
    Const cst_Secret_Login As String = "Secret_Login"
    Const cst_Secret_Login_URL As String = "Secret_Login_URL"
    Const cst_USERID_LOGINaspx As String = "USERID_LOGIN.aspx"

    '測試機用的網頁
    Const cst_login1 As String = "login.aspx"
    'Dim TestStr As String
    Dim SecretLogin As String = ""
    Dim vsStrScript As String = ""
    Dim aNow As DateTime

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Session(TIMS.cst_MOICA_Login) = "xxx" '防止session id跳動
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(Me, objconn) Then Exit Sub

        'bt_Plansubmit.Click
        'AddHandler bt_Plansubmit.Click, AddressOf sUtl_ImageButtonLink1

        aNow = TIMS.GetSysDateNow(objconn)
        '日曆顯示功能
        'Call ShowCalendar()

        '顯示超連結
        'Call Show_Hyperlink()

        '顯示orgID資訊
        'labBottomContent.Text = TIMS.ShowBottomContent("162", objconn)

        If Not Page.IsPostBack Then
            Call sCreatex1()
        End If

        'div1.Visible = AccountTable.Visible
        'divtable1.Visible = AccountTable.Visible

    End Sub

    '第1次載入
    Sub sCreatex1()
        txtname.Attributes.Add("autocomplete", "off")
        txtpass.Attributes.Add("autocomplete", "off")
        txtvnum.Attributes.Add("autocomplete", "off")
        'AltMsg.Attributes.Add("autocomplete", "off")

        Dim sAltMsg As String = "" '訊息
        Dim AltMsgSDate As String = "" '訊息公佈日
        Dim AltMsgEDate As String = "" '訊息結束日
        sAltMsg = TIMS.Get_SHAREDCODE_MSG("ALTMSG", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsgSDate = TIMS.Get_SHAREDCODE_MSG("SDATE", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        AltMsgEDate = TIMS.Get_SHAREDCODE_MSG("EDATE", TIMS.cst_gWEB_MSG, TIMS.cst_gANM_NUM, objconn)
        'AltMsg.Value = TIMS.Get_AltMsg_System_Msg(sAltMsg, AltMsgSDate, AltMsgEDate)

        Me.ViewState("txtpass") = Nothing
        If Request("msgid") = "RELOGIN" Then
            Dim Script1 As String = "<script>alert('" & TIMS.cst_NODATAMsg5 & "');top.location.href='" & cst_login1 & "';</script>"
            Call TIMS.Utl_RespWriteEnd(Me, objconn, Script1)
        End If
        'Me.txtPlan.Items.Insert(0, New ListItem("===請選擇===", ""))

        '登入帳號與輸入帳號不同，清除 sm.UserInfo.UserID = Nothing
        If Not sm.UserInfo.UserID Is Nothing Then
            '自動登入配合TIMS
            If Request("User") <> "" And Request("Pwd") <> "" Then
                Me.txtname.Text = Convert.ToString(Request("User"))
            End If
            If Me.txtname.Text <> "" Then
                If sm.UserInfo.UserID <> Me.txtname.Text Then
                    sm.UserInfo.UserID = Nothing
                    Session.Abandon()
                End If
            End If
        End If

        If sm.UserInfo.UserID Is Nothing Then
            'AccountTable.Visible = True
            'FunctionTable.Visible = False
        Else
            'AccountTable.Visible = False
            'FunctionTable.Visible = True
            'Call ShowFunction()
            'Call Show_ListCOrgName()
            'Call show_labWelcome(Convert.ToString(sm.UserInfo.UserID))
            'Call Show_HomeNewsS3(Convert.ToString(sm.UserInfo.UserID))
        End If

        '自動登入配合TIMS
        If Request("User") <> "" And Request("Pwd") <> "" Then
            Me.txtname.Text = Convert.ToString(Request("User"))
            Me.ViewState("txtpass") = Convert.ToString(Request("Pwd"))

            Dim XxScript As String = ""
            'XxScript = "<script>if (document.getElementById('txtpass')) { if (document.getElementById('txtpass').value=='') {document.getElementById('txtpass').value = '" & Me.ViewState("txtpass") & "';btsubmitclick();}}</script>"
            XxScript = "<script>if (document.getElementById('txtpass')) { if (document.getElementById('txtpass').value=='') {document.getElementById('txtpass').value = '" & Me.ViewState("txtpass") & "';}}</script>"
            Page.RegisterStartupScript("TestStr", XxScript)
        End If
    End Sub


    '取得帳號1資料。
    Function sUtl_GetAccount(ByVal sTxtname As String) As DataRow
        Dim rst As DataRow = Nothing
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        Dim cmd As SqlCommand
        Dim odt As New DataTable
        'Dim dr As DataRow
        sql = "" & vbCrLf
        sql += " select Account,RoleID,Password,IsUsed" & vbCrLf
        sql += " ,case when getdate()>=Stopdate then 'Y' else 'N' end Stopmsg" & vbCrLf
        sql += " from Auth_Account where upper(account) = @account " & vbCrLf '& Replace(Me.txtname.Text, "'", "''") & "'" & vbCrLf
        cmd = New SqlCommand(sql, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("account", SqlDbType.VarChar).Value = sTxtname
            odt.Load(.ExecuteReader())
        End With
        '取得1筆資料。
        If odt.Rows.Count <> 1 Then
            'vsStrScript = "<script>alert('查無此帳號!!!');"
            vsStrScript = "<script>alert('" & cst_alertMsg1 & "!!!');"
            vsStrScript += "location.href='" & SecretLogin & "'</script>"
            Page.RegisterStartupScript("", vsStrScript)
            Return rst
            'Exit Function
        End If

        rst = odt.Rows(0)
        Return rst
    End Function

    '登入最新時間更新
    Sub sUtl_UpdAccount(ByVal sTxtname As String)
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        Dim cmd As SqlCommand
        'Dim odt As New DataTable
        'Dim dr As DataRow
        sql = "" & vbCrLf
        sql = "  UPDATE Auth_Account  "
        sql += " SET last_loginDate=getdate(), IsUsed ='Y'"
        sql += " WHERE upper(account) = @account "
        cmd = New SqlCommand(sql, objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("account", SqlDbType.VarChar).Value = sTxtname
            .ExecuteNonQuery()
        End With
    End Sub

    '輸入帳號 登入送出
    Protected Sub bt_submit_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        'Dim strScript As String
        If Me.txtname.Text <> "" Then Me.txtname.Text = Me.txtname.Text.Trim
        Dim sTxtname As String = UCase(txtname.Text)

        If Not Me.ViewState("txtpass") Is Nothing Then txtvnum.Text = Session("vnum")
        If Session("vnum") <> "" Then
            If txtvnum.Text.Trim.ToLower <> Session("vnum") Then
                vsStrScript = "<script>alert('驗證碼不正確!!!');</script>"
                Page.RegisterStartupScript("wrong", vsStrScript)
                Exit Sub
            End If
        Else
            vsStrScript = "<script>alert('驗證碼尚未產生，請重新整理!!!');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If

        Dim dr As DataRow = Nothing
        Try
            vsStrScript = ""
            dr = sUtl_GetAccount(sTxtname)
        Catch ex As Exception
            vsStrScript = "<script>alert('" & Common.GetJsString(ex.ToString) & "');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End Try

        vsStrScript = ""
        If dr Is Nothing Then
            'vsStrScript = "<script>alert('查無此帳號!!!');"
            vsStrScript = "<script>alert('" & cst_alertMsg1 & "!!!!!" & "');"
            vsStrScript += "location.href='" & SecretLogin & "'</script>"
            Page.RegisterStartupScript("", vsStrScript)
            Exit Sub
        End If
        If Convert.ToString(dr.Item("password")) <> Me.txtpass.Text Then
            'vsStrScript = "<script>alert('密碼不正確!!!');</script>"
            vsStrScript = "<script>alert('" & cst_alertMsg1 & "!" & "');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If
        If Convert.ToString(dr.Item("IsUsed")) <> "Y" Then
            vsStrScript = "<script>alert('該帳號尚未啟用!!!');</script>"
            'vsStrScript = "<script>alert('" & cst_alertMsg1 & "!!" & "');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If
        If Convert.ToString(dr.Item("Stopmsg")) = "Y" Then '過了帳號停用日
            vsStrScript = "<script>alert('帳號已被停用,請洽詢系統管理者!!!');</script>"
            'vsStrScript = "<script>alert('" & cst_alertMsg1 & "!!!" & "');</script>"
            Page.RegisterStartupScript("wrong", vsStrScript)
            Exit Sub
        End If
        ' 20080805 andy 判斷是否超過三個月未登入
        '-----------------------------------------------
        Dim ReturnMsg As String
        ReturnMsg = "登入帳號超過三個月未登入已暫時停權，請洽系統管理員!!!"
        'ReturnMsg = cst_alertMsg1 & "!!!!"
        Dim flag3 As Boolean = TIMS.Check_AccoutLoginDate(Me, dr("Account"), ReturnMsg, objconn)
        If Not flag3 Then Exit Sub


        If dr.Item("RoleID") <> -1 Then
            sm.UserInfo.UserID = dr("Account")
            sm.UserInfo.RoleID = dr("RoleID")
            'AccountTable.Visible = False
            'FunctionTable.Visible = True

            '有基礎 sm.UserInfo.UserID
            'div1.Visible = AccountTable.Visible
            'divtable1.Visible = AccountTable.Visible


            Dim dtNoMsg As DataTable = Nothing
            If dtNoMsg Is Nothing Then
                dtNoMsg = New DataTable
                dtNoMsg.Columns.Add("Subject")
                Dim drNoMsg As DataRow
                drNoMsg = dtNoMsg.NewRow
                drNoMsg("Subject") = "本日無系統作業提醒。"
                dtNoMsg.Rows.Add(drNoMsg)
                'GridView1.DataSource = dtNoMsg
                'GridView1.DataBind()
            End If

            'Call ShowFunction()
        Else
            'RoleID:-1  使用者 無此功能
            Dim sqlstr As String = ""
            sqlstr = "" & vbCrLf
            sqlstr += " select a.Account,a.RoleID,a.LID,a.Password,a.Name,a.IsUsed" & vbCrLf
            sqlstr += " ,c.RID,c.Relship,c.OrgLevel,c.DistID,d.OrgID,d.OrgName " & vbCrLf
            sqlstr += " from Auth_Account a " & vbCrLf
            sqlstr += " JOIN Auth_Relship c on a.OrgID=c.OrgID" & vbCrLf
            sqlstr += " JOIN Org_OrgInfo d on c.OrgID=d.OrgID" & vbCrLf
            sqlstr += " where a.IsUsed='Y' " & vbCrLf
            sqlstr += " and a.Account = '" & Me.txtname.Text & "'" & vbCrLf
            dr = DbAccess.GetOneRow(sqlstr, objconn)
            Dim dt As DataTable = Nothing
            Call set_session(dr("account"), dr("password"), dr("roleid"), dr("lid"), dr("orgid"), dr("orgname"), dr("rid"), dr("relship"), dr("orglevel"), dr("DistID"), 0, dt, 0, 0, "")
        End If

        ''登入顯示 歡迎 訊息
        'Call show_labWelcome(dr("account"))
        ''顯示 我的行事曆
        'Call Show_HomeNewsS3(dr("account"))
        'Call Show_ListCOrgName()

    End Sub

    '登入顯示 歡迎 訊息
    'Sub show_labWelcome(ByVal sAccount As String)
    '    Dim sql As String = ""
    '    sql = "" & vbCrLf
    '    sql += " select Account,RoleID,Password,IsUsed ,Name" & vbCrLf
    '    sql += " ,case when getdate()>=Stopdate then 'Y' else 'N' end Stopmsg" & vbCrLf
    '    sql += " from Auth_Account " & vbCrLf
    '    sql += " where UPPER(Account) = UPPER('" & sAccount & "')" & vbCrLf
    '    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
    '    Me.labWelcome.Text = "您好  先生/小姐 歡迎登入TIMS系統。"
    '    If Not dr Is Nothing Then
    '        Me.labWelcome.Text = "您好，" & Convert.ToString(dr("Name")) & " 先生/小姐 歡迎登入TIMS系統。"
    '    End If
    'End Sub



    '存取使用者 Session，並跳離 (登入)
    Private Sub set_session(ByVal name As Object, ByVal pwd As Object, ByVal role As Object, ByVal lid As Object, ByVal orgid As Object, ByVal orgname As Object, ByVal rid As Object, ByVal relship As Object, ByVal orglevel As Object, ByVal distid As Object, ByVal planid As Object, _
                        ByVal dt As Object, ByVal tplanid As Object, ByVal Years As Object, _
                        ByVal FunctionPage As String)
        sm.UserInfo.UserID = name        '使用者帳號
        Session("UserPwd") = pwd        '使用者密碼
        sm.UserInfo.RoleID = role        '角色代碼 (id_Role) 系統管理者、承辦人.. select * from auth_account select * from id_role
        sm.UserInfo.LID = lid            '階層代碼 (0:局 1:中心 2:委訓) NUMBER
        sm.UserInfo.OrgID = orgid        '機構代碼
        sm.UserInfo.OrgName = orgname    '機構名稱
        sm.UserInfo.RID = rid            '機構業務 ID
        sm.UserInfo.Relship = relship    '機構業務 連結
        sm.UserInfo.OrgLevel = orglevel  '機構階層 (0:職訓局 1:中心 2:委訓(補助單位) 3:(委訓))
        sm.UserInfo.DistID = distid      '轄區
        sm.UserInfo.PlanID = planid      '小計畫
        sm.UserInfo.TPlanID = tplanid    '大計畫
        sm.UserInfo.FunDt = dt           '功能權限的資料表，指定DataTable=sm.UserInfo.FunDt就可以抓到
        sm.UserInfo.Years = Years        '計畫年度
        Session("FunctionPage") = FunctionPage

        System.Web.Security.FormsAuthentication.SetAuthCookie(name, False)
        Session("Secret_Login") = True

        Dim url1 As String = ""
        If Not Session(TIMS.Cst_UploadFile) Is Nothing _
            AndAlso Not Session(TIMS.Cst_UploadFile & "rtnPath") Is Nothing Then

            'Response.Redirect(Session(TIMS.Cst_UploadFile & "rtnPath"))
            url1 = Session(TIMS.Cst_UploadFile & "rtnPath")
            Call TIMS.Utl_Redirect(Me, objconn, url1)
        End If

        '不啟用https Not TIMS.Get_httpsProtocol 
        'Dim url1 As String = ""
        If Not TIMS.Get_httpsProtocol Then
            '不啟用https
            Session(cst_UserID_Login) = True '使用測試環境登出
            url1 = TIMS.cst_indexB1
        Else
            '啟用https
            url1 = TIMS.cst_indexA1
        End If
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    '帳號清除 重填
    Protected Sub bt_reset_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) 'Handles bt_reset.Click
        Me.txtname.Text = ""
        Me.txtpass.Text = ""
    End Sub

    '不選擇計畫 重填
    Protected Sub bt_reset2_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) 'Handles bt_reset2.Click
        sm.UserInfo.UserID = Nothing
        'Response.Redirect("logout.aspx")
        Session.Abandon()
        'Session.RemoveAll()
        'Response.Redirect("Login.aspx")
        'Response.Redirect("MOICA_Login.aspx")
        'Response.Redirect("Secret_Login.aspx")
        'Response.Redirect(SecretLogin)
        Session(cst_UserID_Login) = True '使用測試環境登出
        Call TIMS.Utl_Redirect(Me, objconn, SecretLogin) '轉移login.aspx
    End Sub

    'bt_Plansubmit
End Class