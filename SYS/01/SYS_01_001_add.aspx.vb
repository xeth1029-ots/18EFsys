Partial Class SYS_01_001_add
    Inherits AuthBasePage

    Dim g_parms As Hashtable
    Dim sTxtIdUnchange As String = ""
    Dim sTxtIDNO As String = ""
    Dim sTxtEMAIL As String = ""
    'Dim sVCode As String=""

    Const cst_alert_msg_1g As String = "查無此帳號!"
    Const cst_alert_msg_2c As String = "請輸入帳號、身分證號碼、E-Mail與驗證碼"
    Const cst_alert_msg_3 As String = "此帳號未設定EMAIL!"
    Const cst_alert_msg_98ot As String = "系統已寄發密碼重設通知函-該帳號-(E-Mail)! "
    Const cst_alert_msg_99ot As String = "資料填寫有誤，請重新輸入!!"

    Dim u_nameid As String = "" '大寫的帳號測試 'Dim u_nameid As String=UCase(nameid.Text)

    Const cst_errorMsg1 As String = "未輸入身分證或居留證號!"
    Const cst_errorMsg2 As String = "身分證或居留證號有誤!!!"
    Const cst_errorMsg3 As String = "輸入帳號有誤!!!請重新輸入!!!(系統規則無提供帳號修改)"
    Const cst_errorMsg4a As String = "輸入新增帳號字串太短，請超過5字以上(含)!"
    Const cst_errorMsg4b As String = "輸入新增帳號字串太長，請勿超過15字!"
    Const cst_errorMsg4c As String = "輸入新增帳號，不得與身分證字號相同!"
    Const cst_errorMsg4d As String = "輸入新增帳號，不得使用身分證字號!"
    Const cst_errorMsg4e As String = "輸入密碼，不得使用身分證字號!"

    Const cst_errorMsg5 As String = "輸入帳號(動作)有誤!!!請重新操作!!!"
    Const cst_errorMsg6 As String = "未輸入帳號!!!請重新輸入!!!"
    Const cst_errorMsg7 As String = "帳號 具有危險性的字元，請重新設定!"
    Const cst_errorMsg7b As String = "密碼 具有危險性的字元，請重新設定!"

    Const cst_errorMsg8 As String = "請先「檢查身分證」，身分證或居留證 無誤後才可進行「儲存」!"
    Const cst_errorMsg9 As String = "角色選擇資料有誤，請先確認「角色」資料!"
    Const cst_errorMsg9b As String = "拒絕修改，該「角色」只同意本人自行修改資料 或有權限者修改!"
    Const cst_errorMsg11 As String = "帳號重覆!!!請重新輸入!!!"
    Const cst_errorMsg12 As String = "修改帳號重覆!!!請洽系統管理者!!!"
    Const cst_errorMsg13 As String = "儲存時產生錯誤，請重新檢查儲存資料!!"
    Const cst_errorMsg14 As String = "身分證或居留證號重覆!!!"
    Const cst_errorMsg15 As String = "請選擇計畫階層!"

    Const cst_errorMsg16 As String = "請輸入帳號英數字,長度限定(5~15)!"
    Const cst_errorMsg16b As String = "帳號檢核不符合規則,請重新輸入!"
    Const cst_errorMsg17 As String = "帳號開頭字，須為英文字母!"

    Const cst_errorMsg18 As String = "密碼請輸入12~14碼(限定數字.英文)！"
    Const cst_errorMsg19 As String = "密碼 應包含大寫英文字母！"
    Const cst_errorMsg20 As String = "密碼 應包含小寫英文字母！"
    Const cst_errorMsg21 As String = "密碼 應包含數字字元！"

    Const cst_errorMsg22 As String = "請輸入確認密碼"
    Const cst_errorMsg23 As String = "輸入確認密碼 與密碼 不同"
    Const cst_errorMsg24 As String = "機構選擇有誤，請重新選擇!"
    Const cst_errorMsg31 As String = "E_Mail欄位改為必填不可為空白，請輸入!"
    Const cst_errorMsg32 As String = "E_Mail欄位格式有誤，請重新輸入!"
    Const cst_errorMsg33 As String = "E_Mail寄發「密碼通知函」有誤!"
    Const cst_errorMsg34 As String = "姓名欄位改為必填不可為空白，請輸入!"

    '[1]有效期限：3 hr-[2]資訊加密-[3]直接連結至【修改密碼】頁-[4]使用者連續點選2次寄送密碼函，僅保留最後一次的密碼連結為有效之連結，其餘失效。

    Const cst_OK_msg1 As String = "您可以使用該帳號!!!"

    Const cst_NG_msg1 As String = "該帳號已有人使用!!!"
    Const cst_NG_msg2 As String = "請輸入帳號!!!"
    Const cst_NG_msg3 As String = "身分證號碼已經存在!!!"
    Const cst_NG_msg4 As String = "身分證號碼不可為空!!!"

    'Dim rqAct As String=TIMS.ClearSQM(Request("act"))
    Const cst_rq_act_edit As String = "edit"
    Const cst_rq_act_add As String = "add"

    Const cst_PXSSWADX As String = "PASSWORD"
    Const cst_sess_check_idno As String = "check_idno"

    Dim sErrmsg As String = ""
    Dim oReIDNOmsg As String = ""

    Dim chk_UserIsSupper As Boolean = False '是否要檢核-本人 登入者UserID 是否為-SNOOPY-/有權者本人 (管理者) 。
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim flgLIDx0xROLEIDx1 As Boolean = False '署-系統管理者 

    Dim objconn As SqlConnection

    Sub sUtl_PageInit1()
        Dim dt As DataTable = TIMS.Get_USERTABCOLUMNS("AUTH_ACCOUNT", objconn)
        Call TIMS.sUtl_SetMaxLen(dt, "ACCOUNT", nameid)
        'Call TIMS.sUtl_SetMaxLen(30, userpass)
        'Call TIMS.sUtl_SetMaxLen(30, userpass2)
        Call TIMS.sUtl_SetMaxLen(dt, "NAME", txtname)
        Call TIMS.sUtl_SetMaxLen(dt, "PHONE", telphone)
        Call TIMS.sUtl_SetMaxLen(dt, "IDNO", IDNO)
        Call TIMS.sUtl_SetMaxLen(dt, "EMAIL", email)
    End Sub

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        Call sUtl_PageInit1()

        '檢查Session是否存在 End
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        ''是否要檢核 登入者UserID 是否為SNOOPY (管理者)。
        'If UCase(sm.UserInfo.UserID)=UCase("snoopy") Then
        '    'chk_UserIDsnoopy=True '不用檢核角色。
        '    If nameid.Text <> "" Then If UCase(nameid.Text)=UCase("snoopy") Then chk_UserIDsnoopy=True '不用檢核角色。
        'End If
        '超級使用者--判斷登入者的權限。
        flgROLEIDx0xLIDx0 = TIMS.IsSuperUser(Me, 1) 'False
        'Dim flag_test As Boolean=TIMS.sUtl_ChkTest() '測試
        BtnrResetPXD.Visible = False ' If((flag_test AndAlso flgROLEIDx0xLIDx0), True, False)

        chkLoginWay2.Enabled = False '不可使用
        chkLoginWay2.Visible = False '不可檢視
        If sm.UserInfo.LID = 0 Then chkLoginWay2.Visible = True '可檢視
        '如果是系統管理者開啟功能。
        If flgROLEIDx0xLIDx0 Then
            'ROLEID=0 LID=0
            chkLoginWay2.Visible = True '可檢視
            chkLoginWay2.Enabled = True '可使用
            If nameid.Text <> "" Then If UCase(nameid.Text) = UCase(sm.UserInfo.UserID) Then chk_UserIsSupper = True '登入角色等同，且為管理者 
        End If

        '署-系統管理者--'判斷登入者的權限
        flgLIDx0xROLEIDx1 = TIMS.ChkUserLIDRole(sm, 0, 1)
        If flgLIDx0xROLEIDx1 Then
            '署-系統管理者
            chkLoginWay2.Visible = True '可檢視
            chkLoginWay2.Enabled = True '可使用
        End If

        If Not Page.IsPostBack Then
            Call CreateDistID()
            Call LoadData1()
        Else
            '異常狀況下啟動
            Call LoadData2()
        End If
        Call SUB_CHECK_IDNO(Me, IDNO.Text, nameid.Text, oReIDNOmsg)

        '若機構是分署(中心) 【開啟跨區支援】
        cblDistID.Enabled = False
        Select Case Convert.ToString(sm.UserInfo.LID)
            Case "0", "1" '署(局)、分署(中心) '登入者 
                '判斷選擇的機構是否為分署(中心)
                If Len(RIDValue.Value) = 1 Then cblDistID.Enabled = True   '"A", "B", "C", "D", "E", "F", "G"  若機構是分署(中心) 【開啟跨區支援】
                If Not cblDistID.Enabled Then TIMS.Tooltip(cblDistID, "此帳號不提供跨區支援!", True)
            Case Else
                TIMS.Tooltip(cblDistID, "登入者，無法使用跨區支援功能!", True)
        End Select

        Dim sAccount As String = TIMS.ClearSQM(Request("account"))
        '每次檢查動作
        If sAccount <> "" Then
            nameid.Enabled = False
            TBplan.Enabled = False
            but_LevPlan.Disabled = True '失效為真
            TIMS.Tooltip(but_LevPlan, "不可更改!", True)
            update11.Enabled = False
            If CheckSerialNo(sAccount) Then update11.Enabled = True
            Button2.Enabled = False
            If CheckAccountIDNO(sAccount) Then Button2.Enabled = True
        End If

        If nameid.Enabled = False Then Button3.Visible = False
        'back.Attributes("onclick")="history.go(" & ViewState("PageReload") & ");"
        Button3.Attributes("onclick") = "GetID();"

#Region "[用來處理JS呼叫Server端的Function，by:20180926]"
        If Page.IsPostBack Then
            Dim eventTarget As String = If(Request("__EVENTTARGET") Is Nothing, String.Empty, Request("__EVENTTARGET"))
            Dim eventArgument As String = If(Request("__EVENTARGUMENT") Is Nothing, String.Empty, Request("__EVENTARGUMENT"))
            If eventTarget = "CustomPostBack" Then CheckAccountRepeat() 'myCheckID
        End If
#End Region
    End Sub

    '建立cblDistID checkbox資料
    Sub CreateDistID()
        lab_LastDATE.Text = ""
        Dim Sql As String = " SELECT * FROM ID_DISTRICT WHERE DISTID!='000' ORDER BY DISTID" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(Sql, objconn)
        cblDistID = TIMS.Get_DistID(cblDistID, dt)
    End Sub

    '顯示設定cblDistID  checkbox資料
    Sub Sub_SetCblDistID(ByVal sAccount As String)
        If sAccount = "" Then Return
        Dim sAccountUP As String = sAccount.ToUpper()

        Dim dt As DataTable = Nothing
        Dim myParam As New Hashtable From {{"AccountUP", sAccountUP}}
        Dim sql As String = " SELECT ACCOUNT ,DISTID FROM AUTH_ACCRWDIST WHERE UPPER(Account)=UPPER(@AccountUP)"
        dt = DbAccess.GetDataTable(sql, objconn, myParam)
        If TIMS.dtNODATA(dt) Then Return
        Call TIMS.SetCblValue(cblDistID, "DISTID", dt)
    End Sub

    ''' <summary>
    ''' PXD是否有異常
    ''' </summary>
    ''' <returns></returns>
    Function CHK_PWD_ERR(ByVal sACCOUNT As String) As Boolean
        Dim rst As Boolean = True 'true:異常 false:正常 
        Const cst_err1 As String = "err-1"
        Dim pxxxwd1 As String = ""
        Dim pxxxwd2 As String = ""
        Dim parms As New Hashtable From {{"ACCOUNT", sACCOUNT}}
        Dim sql As String = "SELECT PASSWORD FROM AUTH_ACCOUNT WHERE ACCOUNT=@ACCOUNT "
        Try
            pxxxwd1 = DbAccess.ExecuteScalar(sql, objconn, parms)
        Catch ex As Exception
        End Try

        Try
            If pxxxwd1.Length >= 32 Then
                Dim Aes As New AesTk 'Turbo.Crypto4
                pxxxwd2 = Aes.Decrypt(pxxxwd1)
            End If
        Catch ex As Exception
            '產生錯誤，無法反解-異常
            pxxxwd2 = cst_err1
        End Try
        '無法反解-異常
        If pxxxwd2.Equals(cst_err1) Then Return rst
        '(沒有加密)-異常
        If pxxxwd2 = "" Then Return rst

        '(有加密) 比對加密資訊
        Dim HASHPWD1 As String = ""
        Dim parms_2 As New Hashtable From {{"ACCOUNT", sACCOUNT}}
        Dim sql_2 As String = "SELECT HASHPWD1 FROM AUTH_ACCOUNT WHERE ACCOUNT=@ACCOUNT "
        Try
            HASHPWD1 = DbAccess.ExecuteScalar(sql_2, objconn, parms_2)
        Catch ex As Exception
        End Try
        '比對有誤1-異常
        If Not TIMS.CreateHash(pxxxwd2).Equals(HASHPWD1) Then Return rst
        '比對有誤2-異常
        If Not TIMS.EncryptAes(pxxxwd2).Equals(pxxxwd1) Then Return rst
        rst = False
        Return rst
    End Function

    Function LOAD_ACCRWPLAN_drA1(ByRef sAccount As String, ByRef sCreateByAcc As String) As DataRow
        Dim drA1 As DataRow = Nothing
        If sAccount = "" Then Return drA1

        sAccount = TIMS.ClearSQM(sAccount)
        sCreateByAcc = TIMS.ClearSQM(sCreateByAcc)
        Dim sql As String = ""
        sql &= " SELECT a.account ,d.Years ,f.Name DistName ,e.PlanName ,d.Seq" & vbCrLf
        sql &= " ,c.OrgName ,b.PlanID ,b.RID ,a.RoleID ,a.Name ,a.IDNO ,a.Phone ,a.EMail" & vbCrLf
        sql &= " ,a.IsUsed ,a.BACKUSE" & vbCrLf
        sql &= " ,format(a.StopDate,'yyyy/MM/dd') StopDate" & vbCrLf
        sql &= " ,format(a.Last_LoginDate,'yyyy-MM-dd HH:mm:ss') LastDate" & vbCrLf
        sql &= " FROM AUTH_ACCOUNT A" & vbCrLf
        sql &= " JOIN AUTH_ACCRWPLAN B ON A.ACCOUNT=B.ACCOUNT" & vbCrLf
        sql &= " AND EXISTS (SELECT 'X' FROM AUTH_RELSHIP X WHERE x.RID=b.RID)" & vbCrLf
        sql &= " JOIN ORG_ORGINFO c ON a.OrgID=c.OrgID" & vbCrLf
        sql &= " LEFT JOIN ID_Plan d ON b.PlanID=d.PlanID" & vbCrLf
        sql &= " LEFT JOIN Key_Plan e ON e.TPlanID=d.TPlanID" & vbCrLf
        sql &= " LEFT JOIN ID_District f ON f.DistID=d.DistID" & vbCrLf
        sql &= $" WHERE UPPER(a.account)=UPPER('{sAccount}') "
        If sCreateByAcc <> "" Then sql &= $" AND b.CreateByAcc='{sCreateByAcc}'"

        drA1 = DbAccess.GetOneRow(sql, objconn)
        Return drA1
    End Function

    ''' <summary>取得資料 (第1次)</summary>
    Sub LoadData1()
        Dim iRoleID As Integer = 0
        BtnReUsed1.Visible = False '不可視
        BtnSendPXDEMAIL.Visible = False '不可視
        BtnReEmail1.Visible = False '不可視

        '<input id="but_LevPlan" onclick="javascript: wopen('../../Common/LevPlan.aspx?winreload=1&amp;OrgField=orgname&amp;fisBlack=isBlack&amp;SAH=Y', '計畫階段', 850, 400, 1); document.form1.winreload.value=1;" type="button" value="選擇" runat="server" Class="asp_button_M" />
        Dim flag_SAH3 As Boolean = TIMS.IsSuperUser(sm, 3)
        '帳號設定使用-跨轄區與計畫
        Dim s_SHAR3_YN As String = If(flag_SAH3, "&SAH=Y", "")

        'winreload : 呼叫頁取得值SUBMIT
        Dim s_LevPlan_OpenWin As String = "../../Common/LevPlan.aspx?winreload=1&OrgField=orgname&fisBlack=isBlack"
        Dim s_Wopen_FM As String = "wopen('{0}','計畫階段',850,620,1);document.form1.winreload.value=1;"
        Dim s_but_LevPlan_onclick_JS As String = String.Format(s_Wopen_FM, String.Concat(s_LevPlan_OpenWin, s_SHAR3_YN))
        but_LevPlan.Attributes("onclick") = s_but_LevPlan_onclick_JS

        Dim sAccount As String = TIMS.ClearSQM(Request("account"))
        'If Convert.ToString(Request("account")) <> "" Then sAccount=Convert.ToString(Request("account")).Replace("'", "''")
        'nowdate.Value=Now.Year.ToString() & "/" & "/" & Now.Month.ToString() & "/" & Now.Day.ToString()
        'ViewState("PageReload")=-1 '計算Reload次數
        nameid.Text = TIMS.ClearSQM(Request("myid"))

        'userpass.TextMode=TextBoxMode.Password
        'userpass2.TextMode=TextBoxMode.Password
        'If sAccount <> "" Then
        '    userpass.TextMode=TextBoxMode.SingleLine
        '    userpass2.TextMode=TextBoxMode.SingleLine
        'End If

        If sAccount <> "" Then
            Call Sub_SetCblDistID(sAccount)
            'Dim selreader As SqlDataReader

            Dim drA1 As DataRow = LOAD_ACCRWPLAN_drA1(sAccount, "Y")
            If drA1 Is Nothing Then
                '自動修正 帳號預設計畫 1次-1
                drA1 = LOAD_ACCRWPLAN_drA1(sAccount, "")
                '自動修正 帳號預設計畫 1次-2
                Call UPDATE_ACCRWPLAN(drA1)
                '自動修正 帳號預設計畫 1次-3
                drA1 = LOAD_ACCRWPLAN_drA1(sAccount, "Y")
            End If
            If drA1 Is Nothing Then
                '帳號設為停用，並註記停用日期
                UPDATE_STOP_ACCOUNT(sAccount)
                Common.MessageBox(Me, "查無該帳號預設計畫!!請執行計畫賦予修復此問題!!(帳號設為停用)")
                'Exit Sub
            End If

            If drA1 IsNot Nothing Then
                nameid.Text = drA1("account")
                nameid.Enabled = False

                If flgROLEIDx0xLIDx0 Then
                    Dim flag_pxd_err As Boolean = CHK_PWD_ERR(nameid.Text)
                    BtnReUsed1.Visible = True '可視
                    BtnSendPXDEMAIL.Visible = True '可視
                    BtnReEmail1.Visible = True '可視
                    If flag_pxd_err Then BtnrResetPXD.Visible = True
                End If
                'userpass.Text="" 'drA1("Password")
                'userpass2.Text="" 'drA1("Password")
                TBplan.Text = String.Concat(drA1("Years"), drA1("DistName"), drA1("PlanName"), drA1("Seq"), " _ ", drA1("OrgName"))
                PlanIDValue.Value = drA1("PlanID")
                RIDValue.Value = drA1("RID")

                'If Convert.ToString(drA1("RoleID")) <> "" Then RoleID=Convert.ToString(drA1("RoleID"))
                'If drA1("RoleID")=0 Then    '超級使用者例外處理
                '    Role.Items.Insert(0, New ListItem("超級使用者", 0))
                '    Common.SetListItem(Role, drA1("RoleID").ToString)
                '    Role.Enabled=False
                'Else
                '    '2005/2/3修改，將RoleID存入變數
                '    'Common.SetListItem(Role, drA1("RoleID").ToString)
                'End If

                iRoleID = If(Convert.ToString(drA1("RoleID")) <> "", Val(drA1("RoleID")), 99)
                txtname.Text = drA1("Name")
                IDNO.Text = TIMS.ChangeIDNO(drA1("IDNO").ToString)
                tr_IDNO.Visible = (IDNO.Text = "")

                telphone.Text = Convert.ToString(drA1("Phone"))
                email.Text = Convert.ToString(drA1("Email"))

                IsUsed.Checked = False
                If Convert.ToString(drA1("IsUsed")) = "Y" Then IsUsed.Checked = True

                '判斷是否勾選「帳號/密碼登入」
                Dim sLogWay2 As String = Convert.ToString(drA1("BACKUSE"))
                chkLoginWay2.Checked = If(sLogWay2 = "Y", True, False)

                '如果不是系統管理者則身分證欄位及角色欄位反白
                IDNO.Enabled = False
                DDL_Role.Enabled = False
                If sm.UserInfo.RoleID <= 1 Then '如果是系統管理者有刪除自然憑證序號的權限
                    IDNO.Enabled = True
                    DDL_Role.Enabled = True
                    update11.Visible = True
                    update11.Attributes("onclick") = "return confirm('這樣會清除使用者自然人憑證資料,\n確定要繼續清除?');"
                    Button2.Visible = True
                    Button2.Attributes("onclick") = "return confirm('這樣會清除使用者身分證與憑證序號資料,\n且狀態會設為不啟用,\n確定要繼續清除?');"
                End If
                StopDate.Text = Convert.ToString(drA1("StopDate"))
                lab_LastDATE.Text = Convert.ToString(drA1("LastDate"))
            End If

            'While selreader.Read() i += 1 End While selreader.Close()
            'Me.Role.Items.Clear()
            'If i=0 Then Common.MessageBox(Me, "查無該帳號預設計畫!!請執行計畫賦予修復該問題!!")

            Select Case sm.UserInfo.LID
                Case 0, 1 '署(局)、分署(中心)
                    '判斷選擇的機構是否為分署(中心)
                    's_ROLEID_RANG_COND 1:署／分署"A", "B", "C", "D", "E", "F", "G"/非1:委訓
                    Dim s_ROLEID_RANG_COND As String = If(Len(RIDValue.Value) = 1, "ROLEID >= 1", "ROLEID > 5")
                    Dim sql As String = ""
                    sql = String.Format(" SELECT NAME,ROLEID FROM ID_ROLE WHERE {0} ORDER BY ROLEID ", s_ROLEID_RANG_COND)
                    Dim dtROLE As DataTable = DbAccess.GetDataTable(sql, objconn)
                    With DDL_Role
                        'Dim vROLEID As String=dtROLE.Rows(0)("ROLEID")
                        .DataSource = dtROLE
                        .DataTextField = "Name"
                        .DataValueField = "RoleID"
                        .DataBind()
                        If Len(RIDValue.Value) = 1 Then
                            If DDL_Role.Items.FindByValue("") Is Nothing Then
                                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                            End If
                        End If
                    End With
                    Common.SetListItem(DDL_Role, Convert.ToString(iRoleID)) '"99"

                Case Else
                    '鎖定
                    DDL_Role.Items.Add(New ListItem("一般使用者", 99))
                    Common.SetListItem(DDL_Role, "99")
                    DDL_Role.Enabled = False

            End Select

            If iRoleID = 0 Then    '超級使用者例外處理
                If DDL_Role.Items.FindByValue("0") Is Nothing Then
                    DDL_Role.Items.Insert(0, New ListItem("超級使用者", "0"))
                    Common.SetListItem(DDL_Role, iRoleID.ToString())
                    DDL_Role.Enabled = False
                    'Else Common.SetListItem(DDL_Role, iRoleID.ToString())
                End If
            End If

        End If
    End Sub

    ''' <summary>'自動修正 帳號預設計畫 1次</summary>
    ''' <param name="drA1"></param>
    Private Sub UPDATE_ACCRWPLAN(ByRef drA1 As DataRow)
        If drA1 Is Nothing Then Return '查無資料離開

        Dim vACCOUNT As String = drA1("ACCOUNT").ToString().ToUpper()
        Dim vPLANID As String = drA1("PLANID")
        Dim vRID As String = drA1("RID")

        Dim u_parmas As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"ACCOUNT", vACCOUNT},
            {"PLANID", vPLANID},
            {"RID", vRID}
        }
        Dim sql As String = ""
        sql &= " UPDATE AUTH_ACCRWPLAN" & vbCrLf
        sql &= " SET CREATEBYACC='Y',MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        sql &= " WHERE UPPER(ACCOUNT)=@ACCOUNT AND PLANID=@PLANID AND RID=@RID" & vbCrLf
        DbAccess.ExecuteNonQuery(sql, objconn, u_parmas)
    End Sub

    ''' <summary>取得資料 (非第1次)</summary>
    Sub LoadData2()
        'If Role.SelectedValue <> "" Then vsRole=Role.SelectedValue
        'If userpass.TextMode=TextBoxMode.Password Then
        '    userpass.Attributes("value")=userpass.Text
        '    userpass2.Attributes("value")=userpass2.Text
        'End If
        '2005/1/5假如有選擇訓練機構，判斷可以用指定的角色---------------------------Start
        Dim rep_ROLEID As String = "99"
        If winreload.Value <> "1" Then Return

        '取得資料 (非第1次)
        Dim vs_DDL_Role As String = TIMS.GetListValue(DDL_Role) 'Role.SelectedValue
        'Dim dt As DataTable=Nothing
        Dim sql As String = ""
        '機構是委訓單位
        sql = "SELECT NAME,ROLEID FROM ID_ROLE WHERE ROLEID>5 ORDER BY ROLEID"
        '判斷選擇的機構是否為分署(中心)
        If Len(RIDValue.Value) = 1 Then
            '若機構是分署(中心)  "A", "B", "C", "D", "E", "F", "G"
            Select Case sm.UserInfo.RoleID '判斷登入的角色
                Case 0, 1 '系統管理者
                    rep_ROLEID = "1"
                Case 2   '一級以上
                    rep_ROLEID = "2"
                Case 3   '一級
                    rep_ROLEID = "3"
                Case 4   '二級
                    rep_ROLEID = "4"
                Case 5   '承辦人
                    rep_ROLEID = "5"
                Case 99  '一般使用者
                    rep_ROLEID = "99"
                Case Else '99  '一般使用者
                    rep_ROLEID = "99"
            End Select
            sql = String.Format("SELECT NAME,ROLEID FROM ID_ROLE WHERE ROLEID>={0} ORDER BY ROLEID", rep_ROLEID)
        End If

        Dim dtROLE As DataTable = DbAccess.GetDataTable(sql, objconn)
        With DDL_Role
            .SelectedIndex = -1
            .Items.Clear()
            If dtROLE.Rows.Count > 0 Then
                Dim vROLEID As String = dtROLE.Rows(0)("ROLEID")
                .DataSource = dtROLE
                .DataTextField = "NAME"
                .DataValueField = "ROLEID"
                .SelectedValue = vROLEID
                .DataBind()
            End If
            If rep_ROLEID <> "99" Then
                If DDL_Role.Items.FindByValue("") Is Nothing Then
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End If
            End If
        End With
        winreload.Value = ""
        Select Case rep_ROLEID
            Case "99"
                Common.SetListItem(DDL_Role, "99")
            Case Else
                If vs_DDL_Role <> "" Then Common.SetListItem(DDL_Role, vs_DDL_Role)
        End Select
        '2005/1/5假如有選擇訓練機構，判斷可以用指定的角色---------------------------End
    End Sub

    'Function Chk_AccountRegular() As Boolean
    '    Dim rst As Boolean=True 'true:驗證無誤
    '    Return rst
    'End Function

    '儲存
    Sub ChkSaveData1()
        Dim strErrmsg As String = ""
        txtname.Text = TIMS.ClearSQM(txtname.Text)
        If txtname.Text = "" Then strErrmsg &= cst_errorMsg34 & vbCrLf

        nameid.Text = TIMS.ClearSQM(nameid.Text)

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text)) '轉大寫
        If IDNO.Text = "" AndAlso tr_IDNO.Visible Then strErrmsg &= cst_errorMsg1 & vbCrLf

        email.Text = TIMS.ChangeEmail(TIMS.ClearSQM(email.Text)) '轉EMAIL
        If email.Text = "" Then strErrmsg &= cst_errorMsg31 & vbCrLf
        If email.Text <> "" AndAlso Not TIMS.CheckEmail(email.Text) Then strErrmsg &= cst_errorMsg32 & vbCrLf

        If strErrmsg <> "" Then
            Common.MessageBox(Me, strErrmsg)
            Exit Sub
        End If

        Dim sIDNO As String = IDNO.Text
        'Dim flagOK1 As Boolean=False 'false:證號檢核失敗

        '1:國民身分證 -檢查
        Dim flagIdno1 As Boolean = TIMS.CheckIDNO(sIDNO) '檢查身分證號
        '2:居留證 4:居留證2021 -檢查
        Dim flagPermit2 As Boolean = TIMS.CheckIDNO2(sIDNO, 2) '2:居留證
        Dim flagPermit4 As Boolean = TIMS.CheckIDNO2(sIDNO, 4) '4:居留證2021
        If Not flagIdno1 AndAlso Not flagPermit2 AndAlso Not flagPermit4 Then
            Common.MessageBox(Me, cst_errorMsg2)
            Exit Sub
        End If

        If nameid.Text = "" Then
            Common.MessageBox(Page, cst_errorMsg3)
            Exit Sub
        End If

        Dim rqAct As String = TIMS.ClearSQM(Request("act"))
        Dim sAccountEdit As String = "" '修改
        sAccountEdit = "" '新增
        Select Case rqAct'Convert.ToString(Request("act"))
            Case cst_rq_act_add '"add"
                '新增
                'https://jira.turbotech.com.tw/browse/TIMSC-237
                If IDNO.Text = TIMS.ChangeIDNO(nameid.Text) Then
                    Common.MessageBox(Page, cst_errorMsg4c)
                    Exit Sub
                End If
                HidsAccount.Value = nameid.Text '新增 請輸入帳號英數字,長度限定(5~15)!
                If Len(nameid.Text) < 5 Then
                    Common.MessageBox(Page, cst_errorMsg4a)
                    Exit Sub
                End If
                If Len(nameid.Text) > 15 Then
                    Common.MessageBox(Page, cst_errorMsg4b)
                    Exit Sub
                End If

            Case cst_rq_act_edit '"edit"
                '修改
                sAccountEdit = TIMS.ClearSQM(Request("account")) '修改
                If nameid.Text <> sAccountEdit Then
                    Common.MessageBox(Page, cst_errorMsg3)
                    Exit Sub
                End If
                HidsAccount.Value = nameid.Text '修改。

            Case Else
                'Dim rqAct As String=Convert.ToString(Request("act"))
                'rqAct=TIMS.ClearSQM(rqAct)
                Common.MessageBox(Page, rqAct & cst_errorMsg5)
                Exit Sub
        End Select

        If HidsAccount.Value = "" Then
            Common.MessageBox(Page, cst_errorMsg6)
            Exit Sub
        End If

        'SQL Injection @strCheckInput
        If TIMS.CheckInput(HidsAccount.Value) Then
            Common.MessageBox(Me, cst_errorMsg7)
            Exit Sub
        End If

        Dim sErrMsg As String = ""
        Call CheckData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        'Dim rqAct As String=TIMS.ClearSQM(Request("act"))
        Dim s1 As String = HidsAccount.Value
        Dim s2 As String = IDNO.Text
        Select Case rqAct'Convert.ToString(Request("act"))
            Case cst_rq_act_add '"add"
                '新增
                s1 = "" '新增為空白比對。
            Case cst_rq_act_edit '"edit"
                '修改
            Case Else
                '異常
        End Select

        Dim ReIDNOmsg As String = ""
        If Not CheckAccountIDNO3(s1, s2, ReIDNOmsg) Then
            Common.MessageBox(Me, ReIDNOmsg)
            Exit Sub
        End If

        If Not Me.IsUsed.Checked Then '執行停用功能
            Select Case Session(cst_sess_check_idno) '未驗身分證號
                Case 0
                    Session(cst_sess_check_idno) = 1 '不必再驗證
            End Select
        End If

        '帳號一律換回大寫
        'nameid.Text=UCase(nameid.Text)
        u_nameid = UCase(nameid.Text)
        Select Case Session(cst_sess_check_idno)
            Case 0
                Common.MessageBox(Page, cst_errorMsg8)
            Case 1
                Dim s_Parms_gp As New Hashtable
                s_Parms_gp.Clear()
                Dim sql_gp As String = ""
                Dim aGID As String = "" '使用群組先取得
                '判斷選擇的機構是否為分署(中心)
                If Len(RIDValue.Value) = 1 Then
                    '"A", "B", "C", "D", "E", "F", "G"
                    Dim v_DDL_Role As String = TIMS.GetListValue(DDL_Role)
                    Select Case v_DDL_Role 'Role.SelectedValue
                        Case "1", "2", "3", "4", "5", "99"
                            Select Case RIDValue.Value
                                Case "A" '署(局)的帳號功能
                                    s_Parms_gp.Add("GROLE", v_DDL_Role)
                                    sql_gp = "SELECT GID FROM AUTH_GROUPCONTRA WHERE GTYPE='0' AND GROLE=@GROLE" 'Role.SelectedValue & "'"
                                Case Else '分署(中心)的帳號功能
                                    s_Parms_gp.Add("GROLE", v_DDL_Role)
                                    sql_gp = "SELECT GID FROM AUTH_GROUPCONTRA WHERE GTYPE='1' AND GROLE=@GROLE" 'Role.SelectedValue & "'"
                            End Select

                        Case Else
                            '是否為SNOOPY (管理者)。
                            If chk_UserIsSupper Then
                                sql_gp = "SELECT GID FROM AUTH_GROUPCONTRA WHERE 1<>1"
                            End If
                            If flgROLEIDx0xLIDx0 Then
                                '非本人登入
                                If Not chk_UserIsSupper Then
                                    Common.MessageBox(Page, cst_errorMsg9b)
                                    Exit Sub
                                End If
                            Else
                                '檢核角色。
                                Common.MessageBox(Page, cst_errorMsg9)
                                Exit Sub
                            End If
                    End Select
                Else
                    '委訓的帳號功能
                    sql_gp = "SELECT GID FROM AUTH_GROUPCONTRA WHERE GTYPE='2'"
                End If

                '檢核TABLE 欄位存在與否
                Dim flag_COL_PXSSWAD As Boolean = TIMS.CHK_TACL_EXISTS(objconn, "AUTH_ACCOUNT", cst_PXSSWADX)

                Dim dt As DataTable
                dt = DbAccess.GetDataTable(sql_gp, objconn, s_Parms_gp)
                If dt.Rows.Count > 0 Then aGID = Convert.ToString(dt.Rows(0)("gid"))

                'If aGID="" Then
                '    Dim msg As String
                '    msg="系統群組資料有誤，請先確認「系統群組」資料!"
                '     Common.MessageBox(Page, msg)
                '    Exit Sub
                'End If

                Dim dLAST_LOGINDATE As Date = Now
                If sAccountEdit = "" Then
                    '新增動作
                    Dim i_Parms As New Hashtable
                    i_Parms.Add("ACCOUNT", u_nameid)
                    Dim sqlstr As String = ""
                    sqlstr = " SELECT 'x' FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT)=@ACCOUNT" '" & UCase(Me.nameid.Text) & "') "
                    Dim sqldr As DataRow = DbAccess.GetOneRow(sqlstr, objconn, i_Parms)
                    If sqldr IsNot Nothing Then
                        Common.MessageBox(Page, cst_errorMsg11)
                        'DbAccess.RollbackTrans(objTrans)
                        Exit Sub
                    End If

                Else
                    Dim u_Parms As New Hashtable
                    u_Parms.Add("ACCOUNT", UCase(sAccountEdit))
                    '修改動作'判斷重複字母不區分大小寫。
                    Dim sqlstr As String = ""
                    sqlstr = "SELECT COUNT(1) CNT,MAX(LAST_LOGINDATE) LAST_LOGINDATE FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT) =@ACCOUNT" ' UPPER('" & UCase(sAccountEdit) & "') "
                    Dim drU As DataRow = DbAccess.GetOneRow(sqlstr, objconn, u_Parms)
                    If drU Is Nothing Then
                        Common.MessageBox(Page, TIMS.cst_NODATAMsg2)
                        'DbAccess.RollbackTrans(objTrans)
                        Exit Sub
                    End If
                    Dim iCNT As Integer = Val(drU("CNT")) 'DbAccess.ExecuteScalar(sqlstr, objconn) '.GetOneRow(sqlstr, objconn)
                    If iCNT > 1 Then
                        Common.MessageBox(Page, cst_errorMsg12)
                        Exit Sub
                    End If
                    'Dim dLAST_LOGINDATE As Date=CDate(drU("LAST_LOGINDATE")) 
                    dLAST_LOGINDATE = CDate(drU("LAST_LOGINDATE"))
                End If

                IDNO.Text = TIMS.ChangeIDNO(IDNO.Text)
                hidOrgID.Value = TIMS.Get_OrgID(Me.RIDValue.Value, objconn)
                If Val(hidOrgID.Value) = 0 Then
                    Common.MessageBox(Page, cst_errorMsg24)
                    Exit Sub
                End If
                If Val(hidOrgID.Value) = -1 Then
                    Common.MessageBox(Page, cst_errorMsg24)
                    Exit Sub
                End If

                Dim s_DDL_Role As String = TIMS.GetListValue(DDL_Role) 'TIMS.ClearSQM(Role.SelectedValue)
                If s_DDL_Role = "" Then s_DDL_Role = "99"
                If chk_UserIsSupper Then s_DDL_Role = "0" '是否為SNOOPY (管理者)。 select * from id_role
                Select Case s_DDL_Role
                    Case "99"'s_DDL_Role="99"
                    Case "0", "1", "2", "3", "4", "5"
                    Case Else
                        Common.MessageBox(Page, cst_errorMsg9)
                        Exit Sub
                End Select

                'asfdjfdskljfkldasjf
                Dim Htb1 As New Hashtable
                Htb1.Add("sAccountEdit", sAccountEdit)
                Htb1.Add("s_DDL_Role", s_DDL_Role)
                Htb1.Add("aGID", aGID)
                Htb1.Add("dLAST_LOGINDATE", dLAST_LOGINDATE)
                Call SaveData1(Htb1)

        End Select

        Session.Remove(cst_sess_check_idno)
    End Sub

    Sub SaveData1(ByRef Htb1 As Hashtable)
        Dim sAccountEdit As String = TIMS.GetMyValue2(Htb1, "sAccountEdit")
        Dim s_DDL_Role As String = TIMS.GetMyValue2(Htb1, "s_DDL_Role")
        Dim aGID As String = TIMS.GetMyValue2(Htb1, "aGID")
        Dim dLAST_LOGINDATE As Date = TIMS.GetMyValue2(Htb1, "dLAST_LOGINDATE")

        Dim sqlstr As String = ""

        Dim s_parma As New Hashtable
        s_parma.Add("sAccountEdit", sAccountEdit)
        s_parma.Add("ACCOUNT", nameid.Text)
        s_parma.Add("ROLEID", s_DDL_Role)
        s_parma.Add("LID", Convert.ToString(Get_LID(RIDValue.Value)))
        s_parma.Add("NAME", txtname.Text)
        s_parma.Add("PHONE", telphone.Text)
        s_parma.Add("EMAIL", email.Text)
        s_parma.Add("ORGID", hidOrgID.Value)
        s_parma.Add("ISUSED", If(Me.IsUsed.Checked, "Y", "N"))
        s_parma.Add("IDNO", IDNO.Text)
        s_parma.Add("FuncPath", "/SYS/01/SYS_01_001_add.aspx")
        s_parma.Add("TargetTable", "AUTH_ACCOUNT")
        '(額外補充Log寫入功能，by:2018/07/30)
        Call ADD_SYS_TRANS_LOG(sm, s_parma, objconn)

        If sAccountEdit = "" Then
            u_nameid = UCase(nameid.Text)
            '新增動作
            '刪除帳號的功能、計畫權限(可能因為資料建立錯誤造成)
            Dim d_prams As New Hashtable
            d_prams.Clear()
            d_prams.Add("ACCOUNT", u_nameid)
            sqlstr = " DELETE AUTH_ACCRWFUN WHERE UPPER(ACCOUNT)=@ACCOUNT" 'UPPER('" & nameid.Text & "') "
            DbAccess.ExecuteNonQuery(sqlstr, objconn, d_prams)
            sqlstr = " DELETE AUTH_ACCRWPLAN WHERE UPPER(ACCOUNT)=@ACCOUNT" 'UPPER('" & nameid.Text & "') "
            DbAccess.ExecuteNonQuery(sqlstr, objconn, d_prams)

            Dim i_prams As New Hashtable
            i_prams.Clear()
            i_prams.Add("ACCOUNT", nameid.Text)
            i_prams.Add("ROLEID", s_DDL_Role) 'Me.Role.SelectedValue
            i_prams.Add("LID", Get_LID(RIDValue.Value))
            i_prams.Add("NAME", txtname.Text)

            i_prams.Add("PHONE", telphone.Text)
            i_prams.Add("EMAIL", email.Text)
            i_prams.Add("ORGID", Val(hidOrgID.Value)) 'If Not hidOrgID.Value.Equals("-1") Then sqldr("OrgID")=hidOrgID.Value
            i_prams.Add("ISUSED", If(IsUsed.Checked, "Y", "N")) '啟用

            i_prams.Add("LAST_LOGINDATE", If(IsUsed.Checked, Now(), dLAST_LOGINDATE))
            i_prams.Add("IDNO", TIMS.ChangeIDNO(IDNO.Text))
            i_prams.Add("STOPDATE", If(StopDate.Text <> "", TIMS.Cdate2(StopDate.Text), Convert.DBNull))
            i_prams.Add("BACKUSE", If(chkLoginWay2.Checked, "Y", Convert.DBNull)) '判斷是否勾選「帳號/密碼登入」

            i_prams.Add("MODIFYACCT", sm.UserInfo.UserID)
            '新增帳號時定義為該帳號首次登次登入時間
            sqlstr = "" & vbCrLf
            sqlstr &= " INSERT INTO AUTH_ACCOUNT(ACCOUNT,ROLEID,LID,NAME, PHONE,EMAIL,ORGID,ISUSED, IDNO" & vbCrLf
            sqlstr &= " ,LAST_LOGINDATE,STOPDATE,BACKUSE,MODIFYACCT,MODIFYDATE)" & vbCrLf
            sqlstr &= " VALUES (@ACCOUNT,@ROLEID,@LID,@NAME, @PHONE,@EMAIL,@ORGID,@ISUSED, @IDNO" & vbCrLf
            sqlstr &= " ,GETDATE(),@STOPDATE,@BACKUSE,@MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(sqlstr, objconn, i_prams)
            'If userpass.Text <> "" Then
            '    If flag_COL_PXSSWAD Then sqldr(cst_PXSSWADX)=userpass.Text
            '    sqldr("HASHPWD1")=TIMS.CreateHash(userpass.Text)
            'End If

            'Dim sql As String=""
            Dim i2_prams As New Hashtable
            i2_prams.Clear()
            i2_prams.Add("ACCOUNT", nameid.Text)
            i2_prams.Add("PLANID", Val(PlanIDValue.Value))
            i2_prams.Add("RID", RIDValue.Value)
            i2_prams.Add("CREATEBYACC", "Y")
            i2_prams.Add("MODIFYACCT", sm.UserInfo.UserID)
            sqlstr = "" & vbCrLf
            sqlstr &= " INSERT INTO AUTH_ACCRWPLAN(ACCOUNT ,PLANID ,RID ,CREATEBYACC ,MODIFYACCT ,MODIFYDATE)" & vbCrLf
            sqlstr &= " VALUES (@ACCOUNT ,@PLANID ,@RID ,@CREATEBYACC ,@MODIFYACCT ,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(sqlstr, objconn, i2_prams)

            '2010/06/09 建立群組
            If aGID <> "" Then Call TIMS.UPDATE_AUTH_GROUPACCT(Me, aGID, nameid.Text, PlanIDValue.Value, objconn) '新增帳號群組資料

            'Dim s_parma As New Hashtable
            s_parma.Clear()
            s_parma.Add("ACCOUNT", nameid.Text)
            s_parma.Add("SENDMAIL", email.Text)
            '/*密碼寄信種類-1:新設密碼-2:忘記密碼(修改密碼)-3:修改密碼 */ SENDMAILTYPE /SDMATYPE
            s_parma.Add("SDMATYPE", "1")
            s_parma.Add("MODIFYACCT", sm.UserInfo.UserID)
            Dim htSS As Hashtable = TIMS.INSERT_PXSSWARDHIS(objconn, s_parma)

            Dim flag_test_ENVC As Boolean = TIMS.CHK_IS_TEST_ENVC() '檢測為測試環境:true 正式環境為:false
            If flag_test_ENVC Then
                Dim xMybody As String = TIMS.GetMyValue2(htSS, "xMybody")
                If xMybody <> "" Then Response.Write(xMybody)
            End If

        Else
            '修改動作
            '帳號基本檔
            'Dim sqlstr As String=""
            sqlstr = "" & vbCrLf
            sqlstr &= " UPDATE AUTH_ACCOUNT" & vbCrLf
            sqlstr &= " SET ROLEID=@ROLEID" & vbCrLf
            sqlstr &= " ,LID=@LID" & vbCrLf
            sqlstr &= " ,NAME=@NAME" & vbCrLf
            sqlstr &= " ,PHONE=@PHONE" & vbCrLf

            sqlstr &= " ,EMAIL=@EMAIL" & vbCrLf
            sqlstr &= " ,ORGID=@ORGID" & vbCrLf
            sqlstr &= " ,IDNO=@IDNO" & vbCrLf
            sqlstr &= " ,ISUSED=@ISUSED" & vbCrLf
            sqlstr &= " ,LAST_LOGINDATE=@LAST_LOGINDATE" & vbCrLf

            sqlstr &= " ,STOPDATE=@STOPDATE" & vbCrLf
            sqlstr &= " ,BACKUSE=@BACKUSE" & vbCrLf
            sqlstr &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
            sqlstr &= " ,MODIFYDATE=GETDATE()" & vbCrLf
            sqlstr &= " WHERE 1=1" & vbCrLf
            sqlstr &= " AND ACCOUNT=@ACCOUNT" & vbCrLf

            Dim u_prams As New Hashtable
            'u_prams.Clear()
            u_prams.Add("ROLEID", s_DDL_Role) 'Me.Role.SelectedValue
            u_prams.Add("LID", Get_LID(RIDValue.Value))
            u_prams.Add("NAME", txtname.Text)
            u_prams.Add("PHONE", telphone.Text)

            u_prams.Add("EMAIL", email.Text)
            u_prams.Add("ORGID", Val(hidOrgID.Value)) 'If Not hidOrgID.Value.Equals("-1") Then sqldr("OrgID")=hidOrgID.Value
            u_prams.Add("IDNO", TIMS.ChangeIDNO(IDNO.Text))
            u_prams.Add("ISUSED", If(IsUsed.Checked, "Y", "N")) '啟用

            u_prams.Add("LAST_LOGINDATE", If(IsUsed.Checked, Now(), dLAST_LOGINDATE))
            u_prams.Add("STOPDATE", If(StopDate.Text <> "", TIMS.Cdate2(StopDate.Text), Convert.DBNull))
            u_prams.Add("BACKUSE", If(chkLoginWay2.Checked, "Y", Convert.DBNull)) '判斷是否勾選「帳號/密碼登入」
            u_prams.Add("MODIFYACCT", sm.UserInfo.UserID)

            u_prams.Add("ACCOUNT", nameid.Text)
            DbAccess.ExecuteNonQuery(sqlstr, objconn, u_prams)
            'If userpass.Text <> "" Then
            '    If flag_COL_PXSSWAD Then sqldr(cst_PXSSWADX)=userpass.Text
            '    sqldr("HASHPWD1")=TIMS.CreateHash(userpass.Text)
            'End If
        End If

        '若為停用刪除資料
        'If Not IsUsed.Checked Then Call DELETE_AUTH_ACCRWDIST() '刪除動作
        Call INSERT_AUTH_ACCRWDIST() '啟用確認 新增資料動作

        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, cst_errorMsg13)

        '    Dim strErrmsg As String=""
        '    strErrmsg &= TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
        '    strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
        '    'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
        '    Call TIMS.WriteTraceLog(Me, ex, strErrmsg)
        '    Exit Sub
        '    'Common.MessageBox(Me, ex.ToString)
        '    'Throw ex
        'End Try
        'Call TIMS.CloseDbConn(tConn)

        Dim rqMID As String = TIMS.Get_MRqID(Me)
        If sAccountEdit = "" Then
            Page.RegisterStartupScript("", "<script>if(confirm('新增成功!是否要繼續新增?')){location.href='SYS_01_001_add.aspx?ID=" & rqMID & "&act=add';}else{location.href='SYS_01_001.aspx?ID=" & rqMID & "';}</script>")
        Else
            Page.RegisterStartupScript("", "<script>alert('修改成功!');location.href='SYS_01_001.aspx?ID=" & rqMID & "';</script>")
        End If
    End Sub

    '刪除動作
    Sub DELETE_AUTH_ACCRWDIST(ByRef vACCOUNTnameid As String)
        Dim d_Param As Hashtable = New Hashtable
        d_Param.Add("ACCOUNT", vACCOUNTnameid)
        Dim d_sql As String = ""
        d_sql = "DELETE AUTH_ACCRWDIST WHERE ACCOUNT= @ACCOUNT"
        DbAccess.ExecuteNonQuery(d_sql, objconn, d_Param)
    End Sub

    '啟用確認 新增轄區資料動作
    Sub INSERT_AUTH_ACCRWDIST()
        Dim vACCOUNTnameid As String = nameid.Text
        If vACCOUNTnameid = "" Then Return

        Call DELETE_AUTH_ACCRWDIST(vACCOUNTnameid)

        Dim i_sql As String = ""
        i_sql = "" & vbCrLf
        i_sql &= " INSERT INTO AUTH_ACCRWDIST (ACCOUNT ,DISTID ,MODIFYACCT ,MODIFYDATE )" & vbCrLf
        i_sql &= " VALUES (@ACCOUNT ,@DISTID ,@MODIFYACCT ,GETDATE() )" & vbCrLf

        'Dim d_sql As String=""
        'd_sql="" & vbCrLf
        'd_sql &= " DELETE AUTH_ACCRWDIST" & vbCrLf
        'd_sql &= " WHERE 1=1" & vbCrLf
        'd_sql &= " AND ACCOUNT=@ACCOUNT AND DISTID=@DISTID" & vbCrLf

        Dim s_sql As String = ""
        s_sql = "" & vbCrLf
        s_sql &= " SELECT 1 FROM AUTH_ACCRWDIST" & vbCrLf
        s_sql &= " WHERE 1=1" & vbCrLf
        s_sql &= " AND ACCOUNT=@ACCOUNT AND DISTID=@DISTID" & vbCrLf
        Dim sCmd As New SqlCommand(s_sql, objconn)

        For i As Integer = 0 To cblDistID.Items.Count - 1
            Dim s_DISTID As String = cblDistID.Items(i).Value
            Dim flag_cblDIST As String = cblDistID.Items(i).Selected
            If flag_cblDIST Then
                Dim dt As New DataTable
                With sCmd
                    .Parameters.Clear()
                    .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = vACCOUNTnameid
                    .Parameters.Add("DISTID", SqlDbType.VarChar).Value = s_DISTID ' cblDistID.Items(i).Value
                    dt.Load(.ExecuteReader())
                End With
                If dt.Rows.Count = 0 Then
                    Dim i_Param As Hashtable = New Hashtable
                    i_Param.Add("ACCOUNT", vACCOUNTnameid)
                    i_Param.Add("DISTID", s_DISTID)
                    i_Param.Add("MODIFYACCT", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(i_sql, objconn, i_Param)
                End If
            End If
        Next
    End Sub

    ''' <summary>檢核資料1</summary>
    ''' <param name="sErrMsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef sErrMsg As String) As Boolean
        Dim rst As Boolean = False '異常:false
        sErrMsg = ""

        '表示該使用者 存存
        Dim flag_account_update As Boolean = False
        lab_LastDATE.Text = TIMS.ClearSQM(lab_LastDATE.Text)
        If lab_LastDATE.Text <> "" Then flag_account_update = True


        'nameid.Text=TIMS.ClearSQM(nameid.Text)
        'If nameid.Text="" Then
        '    sErrMsg &= "請輸入帳號!" & vbCrLf
        '    Return rst
        'End If
        txtname.Text = TIMS.ClearSQM(txtname.Text) '姓名
        nameid.Text = Trim(nameid.Text) '帳號
        If nameid.Text = "" Then
            sErrMsg &= cst_errorMsg6 & vbCrLf
            Return rst
        End If
        Dim V_nameid As String = nameid.Text
        nameid.Text = TIMS.ClearSQM(nameid.Text) '帳號
        If nameid.Text <> V_nameid Then
            sErrMsg &= cst_errorMsg7 & vbCrLf
            Return rst
        End If
        If TIMS.CheckIDNO(nameid.Text) Then
            sErrMsg &= cst_errorMsg4d & vbCrLf
            Return rst
        End If

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text)) '轉大寫
        If IDNO.Text = "" AndAlso tr_IDNO.Visible Then
            sErrMsg &= cst_errorMsg1 & vbCrLf
            Return rst
        End If
        Dim sIDNO As String = IDNO.Text

        '1:國民身分證 -檢查
        Dim flagIdno1 As Boolean = TIMS.CheckIDNO(sIDNO) '檢查身分證號
        '2:居留證 4:居留證2021 -檢查
        Dim flagPermit2 As Boolean = TIMS.CheckIDNO2(sIDNO, 2) '可檢查居留證號1
        Dim flagPermit4 As Boolean = TIMS.CheckIDNO2(sIDNO, 4) '可檢查居留證號2
        If Not flagIdno1 AndAlso Not flagPermit2 AndAlso Not flagPermit4 Then
            sErrMsg &= cst_errorMsg2 & vbCrLf
            Return rst
        End If

        'Dim flagOK1 As Boolean=False 'false:證號檢核失敗
        'Dim flagIDNO As Boolean=TIMS.CheckIDNO(sIDNO) '檢查身分證號
        'Dim flagIDNO2 As Boolean=TIMS.CheckIDNO2(sIDNO, 2) '可檢查居留證號
        'If flagIDNO OrElse flagIDNO2 Then flagOK1=True 'true:證號檢核OK
        'If Not flagOK1 Then
        '    sErrMsg &= cst_errorMsg2 & vbCrLf
        '    Return rst
        'End If

        Dim rqAct As String = TIMS.ClearSQM(Request("act"))
        Select Case rqAct 'Convert.ToString(Request("act"))
            Case cst_rq_act_add '"add"
                '新增
                HidsAccount.Value = nameid.Text '新增
                If Len(nameid.Text) < 5 Then
                    sErrMsg &= cst_errorMsg4a & vbCrLf
                    Return rst
                End If
                If Len(nameid.Text) > 15 Then
                    sErrMsg &= cst_errorMsg4b & vbCrLf
                    Return rst
                End If

                If nameid.Text <> "" Then
                    If Not TIMS.CheckABC(Left(nameid.Text, 1)) Then
                        sErrMsg &= cst_errorMsg17 & vbCrLf
                        Return rst
                    End If
                End If

            Case cst_rq_act_edit '"edit"
                '修改
                Dim sAccount As String = "" '修改
                sAccount = TIMS.ClearSQM(Request("account")) '修改
                If nameid.Text <> sAccount Then
                    sErrMsg &= cst_errorMsg3
                    Return rst
                End If
                HidsAccount.Value = nameid.Text '修改。

            Case Else
                'Dim rqAct As String=TIMS.ClearSQM(Request("act"))
                sErrMsg &= rqAct & cst_errorMsg5 & vbCrLf
                Return rst

        End Select
        If HidsAccount.Value = "" Then
            sErrMsg &= cst_errorMsg6 & vbCrLf
            Return rst
        End If
        'SQL Injection @strCheckInput
        If TIMS.CheckInput(HidsAccount.Value) Then
            sErrMsg &= cst_errorMsg7 & vbCrLf
            Return rst
        End If
        'TBplan
        '請選擇計畫階層
        TBplan.Text = TIMS.ClearSQM(TBplan.Text)
        If TBplan.Text = "" Then
            sErrMsg &= cst_errorMsg15 & vbCrLf
            Return rst
        End If
        'nameid
        'nameid.Text=TIMS.ClearSQM(nameid.Text)

        '(新增帳號時要檢核)
        If Not TIMS.CheckABC321acct(nameid.Text) AndAlso Not flag_account_update Then
            sErrMsg &= cst_errorMsg16 & vbCrLf
            Return rst
        End If
        If sErrMsg <> "" Then Return rst

        '請輸入帳號
        '帳號請輸入數字或英文字
        'userpass
        '請輸入密碼
        '密碼請輸入12~14碼(限定數字.英文)
        'https://jira.turbotech.com.tw/browse/TIMSC-237
        'flagOK1=False 'Dim flagOK1 As Boolean=False
        'Dim oUserpassTxt As String="" 'userpass.Text
        'Dim oUserpass2Txt As String="" 'userpass2.Text
        'Dim str_userpass As String=oUserpassTxt
        'Dim ok_userpass As String=TIMS.ClearSQM(oUserpassTxt)
        'If oUserpassTxt <> "" AndAlso str_userpass <> ok_userpass Then
        '    oUserpassTxt=TIMS.ClearSQM(oUserpassTxt)
        '    sErrMsg &= "密碼" & TIMS.cst_ErrorMsg10 & vbCrLf
        '    Return rst
        'End If
        'If oUserpassTxt <> "" Then
        '    If oUserpassTxt.Length >= 12 AndAlso oUserpassTxt.Length <= 14 Then flagOK1=True
        '    If Not flagOK1 Then sErrMsg &= cst_errorMsg18 & vbCrLf

        '    flagOK1=False
        '    If TIMS.ChkUpper(oUserpassTxt) Then flagOK1=True
        '    If Not flagOK1 Then sErrMsg &= cst_errorMsg19 & vbCrLf

        '    flagOK1=False
        '    If TIMS.ChkLower(oUserpassTxt) Then flagOK1=True
        '    If Not flagOK1 Then sErrMsg &= cst_errorMsg20 & vbCrLf

        '    flagOK1=False
        '    If TIMS.ChkNumber(oUserpassTxt) Then flagOK1=True
        '    If Not flagOK1 Then sErrMsg &= cst_errorMsg21 & vbCrLf

        '    If sErrMsg <> "" Then Return rst

        '    'userpass2
        '    '請輸入確認密碼
        '    If oUserpass2Txt="" Then
        '        sErrMsg &= cst_errorMsg22 & vbCrLf
        '        'Return rst
        '    End If
        '    '確認密碼輸入錯誤
        '    If oUserpassTxt <> oUserpass2Txt Then
        '        sErrMsg &= cst_errorMsg23 & vbCrLf
        '        'Return rst
        '    End If
        '    '確認密碼請輸入12~14碼(限定數字.英文)

        '    Dim V_userpass As String=oUserpassTxt
        '    oUserpassTxt=TIMS.ClearSQM(oUserpassTxt) '密碼
        '    If oUserpassTxt <> V_userpass Then
        '        sErrMsg &= cst_errorMsg7b & vbCrLf
        '        Return rst
        '    End If

        '    If TIMS.CheckIDNO(oUserpassTxt) Then
        '        sErrMsg &= cst_errorMsg4e & vbCrLf
        '        Return rst
        '    End If
        'End If
        'If sErrMsg <> "" Then Return rst
        Return True
    End Function

    '儲存(鈕)
    Private Sub btu_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btu_save.Click
        Call ChkSaveData1()
    End Sub

    '檢查是否有自然人憑證號
    Function CheckSerialNo(ByVal sAccount As String) As Boolean
        Dim Rst As Boolean = False
        If String.IsNullOrEmpty(sAccount) Then Return Rst

        Dim sParam As Hashtable = New Hashtable
        sParam.Add("ACCOUNT", sAccount)

        Dim Sql As String = " SELECT 'x' FROM AUTH_ACCOUNT WHERE 1=1 AND ISNULL(SERIALNO,' ') <> ' ' AND ACCOUNT =@ACCOUNT"
        Dim dr As DataRow = DbAccess.GetOneRow(Sql, objconn, sParam)
        If dr IsNot Nothing Then Rst = True
        Return Rst
    End Function

    '檢查是否有身分證號
    Function CheckAccountIDNO(ByVal sAccount As String) As Boolean
        Dim Rst As Boolean = False
        If String.IsNullOrEmpty(sAccount) Then Return Rst

        Dim sParam As Hashtable = New Hashtable
        sParam.Add("ACCOUNT", UCase(sAccount))

        Dim Sql As String = " SELECT 'x' FROM AUTH_ACCOUNT WHERE 1=1 AND ISNULL(IDNO,' ') <> ' ' AND UPPER(ACCOUNT)=@ACCOUNT"
        Dim dr As DataRow = DbAccess.GetOneRow(Sql, objconn, sParam)
        If dr IsNot Nothing Then Rst = True
        Return Rst
    End Function

    '驗證身分證號3
    Function CheckAccountIDNO3(ByRef sAccount As String, ByRef sIDNO As String, ByRef ReIDNOmsg As String) As Boolean
        Dim rst As Boolean = False '不可儲存 true:可以儲存
        If sIDNO = "" Then
            ReIDNOmsg = cst_NG_msg4 '"身分證號碼不可為空!!!"
            '身分證號 為空
            Return rst
        End If

        'SELECT * FROM AUTH_ACCOUNT 
        'WHERE UPPER(ACCOUNT ) IN (SELECT UPPER(ACCOUNT) X  FROM AUTH_ACCOUNT GROUP BY UPPER(ACCOUNT) HAVING COUNT(1)>1)
        'AND ISUSED='Y'
        'ORDER BY  UPPER(ACCOUNT )

        Dim myParam As Hashtable = New Hashtable
        myParam.Add("IDNO", UCase(sIDNO))
        If sAccount <> "" Then myParam.Add("ACCOUNT", UCase(sAccount))

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT 'X' FROM AUTH_ACCOUNT A" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND UPPER(a.IDNO)=@IDNO" & vbCrLf
        If sAccount <> "" Then sql &= " AND UPPER(a.ACCOUNT) <> @ACCOUNT"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, myParam)

        If dt.Rows.Count > 0 Then
            ReIDNOmsg = cst_NG_msg3 ' "身分證號碼已經存在!!!"
            Return rst
        End If
        rst = True
        Return rst
    End Function

    '驗證身分證號 
    'SELECT UPPER(ACCOUNT) X,COUNT(1) CNT FROM AUTH_ACCOUNT GROUP BY UPPER(ACCOUNT) HAVING COUNT(1)>1
    ''' <summary>驗證身分證號</summary>
    ''' <param name="sAccount"></param>
    ''' <param name="sIDNO"></param>
    ''' <param name="ReIDNOmsg"></param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Function CheckAccountIDNO2(ByRef sAccount As String, ByRef sIDNO As String, ByRef ReIDNOmsg As String) As Integer
        '未驗證身分證號 失敗為0 成功為1 回傳 sIDNO 回傳 ReIDNOmsg
        Dim check_idno As Integer = 0 '未驗證 失敗為0 成功為1
        If sIDNO = "" Then
            ReIDNOmsg = cst_errorMsg1 '身分證號 為空
            Return 0
        End If
        sIDNO = TIMS.ChangeIDNO(sIDNO) '轉換大寫

        Dim sParms As New Hashtable
        sParms.Clear()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.ACCOUNT, a.IDNO ,CASE WHEN ISNULL(A.SERIALNO,' ') != ' ' THEN 1 END SERIALNO" & vbCrLf
        sql &= " ,b.PLANID, b.RID, b.CREATEBYACC" & vbCrLf
        sql &= " ,concat(c.Years,d.Name,e.PlanName,c.Seq) PLANNAME" & vbCrLf
        sql &= " ,g.ORGNAME ,a.ISUSED" & vbCrLf
        sql &= " FROM AUTH_ACCOUNT A" & vbCrLf
        sql &= " LEFT JOIN ORG_ORGINFO g ON a.OrgID=g.OrgID" & vbCrLf
        sql &= " LEFT JOIN AUTH_ACCRWPLAN b ON a.Account=b.Account" & vbCrLf
        sql &= " LEFT JOIN AUTH_RELSHIP f ON f.RID=b.RID" & vbCrLf
        sql &= " LEFT JOIN ID_PLAN c ON b.PlanID=c.PlanID" & vbCrLf
        sql &= " LEFT JOIN ID_DISTRICT d ON c.DistID=d.DistID" & vbCrLf
        sql &= " LEFT JOIN KEY_PLAN e ON c.TPlanID=e.TPlanID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sIDNO <> "" Then sParms.Add("IDNO", sIDNO)
        If sIDNO <> "" Then sql &= " AND UPPER(a.IDNO)=@IDNO" & vbCrLf

        If Convert.ToString(sAccount) <> "" Then
            sAccount = TIMS.ClearSQM(sAccount)
            sParms.Add("ACCOUNT", UCase(sAccount))
            sql &= " AND UPPER(a.ACCOUNT) <>@ACCOUNT" & vbCrLf  '修改動作
        End If
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, sParms)

        Dim msg As String = ""
        ReIDNOmsg = ""
        'reidno.Visible=False

        Dim flagOK1 As Boolean = True  'true:成功 'false:證號檢核失敗
        '1:國民身分證 -檢查
        Dim flagIdno1 As Boolean = TIMS.CheckIDNO(sIDNO) '檢查身分證號
        '2:居留證 4:居留證2021 -檢查
        Dim flagPermit2 As Boolean = TIMS.CheckIDNO2(sIDNO, 2) '可檢查居留證號1
        Dim flagPermit4 As Boolean = TIMS.CheckIDNO2(sIDNO, 4) '可檢查居留證號2
        If Not flagIdno1 AndAlso Not flagPermit2 AndAlso Not flagPermit4 Then
            msg = cst_errorMsg2
            ReIDNOmsg = msg
            flagOK1 = False
        End If

        If flagOK1 Then  '一般驗証
            If dr IsNot Nothing Then
                msg = cst_errorMsg14 & vbCrLf
                If dr("PlanName").ToString <> "" Then
                    msg += "隸屬於：" & dr("PlanName").ToString & vbCrLf
                    'msg += "中心別：" & dr("OrgName").ToString & vbCrLf
                    msg += "分署別：" & dr("OrgName").ToString & vbCrLf
                    If dr("IsUsed") = "Y" Then
                        msg += "狀況：啟用中"
                    Else
                        msg += "狀況：停用中"
                    End If
                    msg += "身分證或居留證號碼" & sIDNO & "重覆"
                End If
                ReIDNOmsg = msg
            Else
                check_idno = 1
                msg = "此身分證或居留證號碼可使用!!!" & vbCrLf
                ReIDNOmsg = msg
            End If
        End If

        Select Case TIMS.Server_Path
            Case "DEMO"
                '測試環境身份證號相同
                msg = ";TIMSDEMO測試環境身份證號相同!!!"
                ReIDNOmsg += msg
                check_idno = 1
        End Select

#Region "(NOUSE)"

        ''Me.IDNO.Enabled=True '開放
        'If Me.IDNO.Enabled=True AndAlso check_idno=1 Then
        '    Me.IDNO.Enabled=False '鎖定
        'End If

#End Region
        Return check_idno
        'Session(cst_sess_check_idno)=check_idno
    End Function

    '取得LID
    Function Get_LID(ByVal RIDValue As String) As Integer
        Dim rst As Integer = 2
        rst = 2 '委訓其他單位
        If RIDValue <> "" AndAlso Len(RIDValue) = 1 Then
            If RIDValue = "A" Then
                rst = 0 '署(局)
            Else
                rst = 1 '分署(中心)
            End If
        End If
        Return rst
    End Function

    ''' <summary>清除 自然人憑證序號</summary>
    ''' <param name="s_ACCOUNT"></param>
    Sub UPDATE_CLEAR_SERIANLNO(ByVal s_ACCOUNT As String)
        Dim u_Parms As New Hashtable
        u_Parms.Add("ModifyAcct", sm.UserInfo.UserID)
        u_Parms.Add("ACCOUNT", s_ACCOUNT)

        Dim sql As String = ""
        sql = ""
        sql &= " UPDATE AUTH_ACCOUNT "
        sql &= " SET stop_Serialno=Serialno ,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() "
        sql &= " WHERE Serialno IS NOT NULL AND Serialno <>' ' AND ACCOUNT=@ACCOUNT "
        DbAccess.ExecuteNonQuery(sql, objconn, u_Parms)

        sql = ""
        sql &= " UPDATE AUTH_ACCOUNT "
        sql &= " SET Serialno=NULL ,ModifyAcct=@ModifyAcct ,ModifyDate=GETDATE() "
        sql &= " WHERE ACCOUNT=@ACCOUNT "
        DbAccess.ExecuteNonQuery(sql, objconn, u_Parms)

    End Sub

    ''' <summary>
    ''' 清除 自然人憑證序號
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub update11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles update11.Click
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        If nameid.Text = "" Then Exit Sub

        UPDATE_CLEAR_SERIANLNO(nameid.Text)

        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Page.RegisterStartupScript("", "<script>alert('清除自然人憑證序號成功!');location.href='SYS_01_001.aspx?ID=" & rqMID & "';</script>")
    End Sub

#Region "(No Use)"

    ''新增帳號群組資料
    'Public Shared Sub Update_Auth_GroupAcct(ByRef MyPage As Page, ByVal aGID As String, ByVal ACCOUNT As String, ByRef objconn As SqlConnection)
    '    Dim sql As String=""
    '    '試著刪除舊資料
    '    'da=New SqlDataAdapter
    '    sql="delete Auth_GroupAcct where gid= @gid and account= @account"
    '    Call TIMS.OpenDbConn(objconn)
    '    Dim oCmd As New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("gid", SqlDbType.VarChar).Value=aGID
    '        .Parameters.Add("account", SqlDbType.VarChar).Value=nameid.Text
    '        .ExecuteNonQuery()
    '    End With

    '    Dim TPlanID As String="" '取得TPlanID
    '    If PlanIDValue.Value <> "" Then TPlanID=TIMS.GetTPlanID(PlanIDValue.Value)

    '    '新增1筆資料
    '    sql="" & vbCrLf
    '    Sql &= " insert into Auth_GroupAcct(GID,ACCOUNT,MODIFYACCT,MODIFYDATE,GTPLANID)" & vbCrLf
    '    Sql &= " values(@GID,@ACCOUNT,@MODIFYACCT,getdate(),@GTPLANID)" & vbCrLf
    '    Call TIMS.OpenDbConn(objconn)
    '    oCmd=New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("GID", SqlDbType.VarChar).Value=aGID
    '        .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value=Convert.ToString(nameid.Text)
    '        .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value=sm.UserInfo.UserID
    '        .Parameters.Add("GTPLANID", SqlDbType.VarChar).Value=TPlanID
    '        .ExecuteNonQuery()
    '    End With
    'End Sub

    'Private Sub count2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles count2.Click
    '    Dim sql, str As String
    '    Dim sss, ss, i As Integer
    '    Dim msg As String
    '    Dim sqlAdapter, tmpAdapter As SqlDataAdapter
    '    Dim sqlTable, tmpTable As DataTable
    '    Dim sqldr, tmprow1, tmprow2, tmprow As DataRow
    '    Dim sqlstr, tmpstr1, tmpstr2, tmpstr As String
    '    sql="select * from AUTH_ACCOUNT where account='snoopy'"
    '    If TIMS.Get_SQLRecordCount(sql) > 0 Then
    '        sqlTable=DbAccess.GetDataTable(sql)
    '        For i=0 To sqlTable.Rows.Count - 1
    '            sqldr=sqlTable.Rows(i)
    '            count22.Text=sqldr("account")
    '            sqldr("serialno")="123123123"
    '            DbAccess.UpdateDataTable(sqlTable, sqlAdapter)
    '        Next
    '    Else
    '        count22.Text="查無資料"
    '        msg="無資料!!!" & vbCrLf
    '        Common.MessageBox(Page, msg)
    '    End If
    '    str="select * from AUTH_ACCOUNT"
    '    count33.Text=TIMS.Get_SQLRecordCount(str)
    'End Sub

#End Region

    ''' <summary>
    ''' 清除 身分證與憑證序號-UPDATE
    ''' </summary>
    ''' <param name="s_ACCOUNT"></param>
    Sub UPDATE_CLEAR_IDNO_SERIANLNO(ByVal s_ACCOUNT As String)
        s_ACCOUNT = TIMS.ClearSQM(s_ACCOUNT)
        Dim u_Parms As New Hashtable
        u_Parms.Add("ModifyAcct", sm.UserInfo.UserID)
        u_Parms.Add("ACCOUNT", s_ACCOUNT)

        Dim U_SQL1 As String = ""
        U_SQL1 &= " UPDATE AUTH_ACCOUNT "
        U_SQL1 &= " SET stop_Idno=Idno ,stop_Serialno=Serialno ,ModifyAcct =@ModifyAcct ,ModifyDate=GETDATE() "
        U_SQL1 &= " WHERE ISNULL(Idno,' ')<> ' ' AND ACCOUNT =@ACCOUNT"
        DbAccess.ExecuteNonQuery(U_SQL1, objconn, u_Parms)

        Dim U_SQL2 As String = ""
        U_SQL2 &= " UPDATE AUTH_ACCOUNT "
        U_SQL2 &= " SET Idno=NULL, Serialno=NULL, IsUsed='N' ,ModifyAcct =@ModifyAcct ,ModifyDate=GETDATE() "
        U_SQL2 &= " WHERE ACCOUNT =@ACCOUNT"
        DbAccess.ExecuteNonQuery(U_SQL2, objconn, u_Parms)

        '申請帳號暫存也要一併清除，才可再重新申請
        Dim U_SQL3 As String = ""
        U_SQL3 &= " UPDATE AUTH_ACCOUNTTEMP "
        U_SQL3 &= " SET Idno=NULL,Serialno=NULL,AuditStatus='N',AuditAcct=@ModifyAcct,AuditDate=GETDATE()"
        U_SQL3 &= " WHERE ACCOUNT=@ACCOUNT"
        DbAccess.ExecuteNonQuery(U_SQL3, objconn, u_Parms)
    End Sub

    ''' <summary>
    ''' 清除 身分證與憑證序號
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        If nameid.Text = "" Then Exit Sub

        '清除 身分證與憑證序號-UPDATE
        UPDATE_CLEAR_IDNO_SERIANLNO(nameid.Text)

        Dim rqMID As String = TIMS.Get_MRqID(Me)
        Page.RegisterStartupScript("", "<script>alert('清除身分證與憑證序號成功!');location.href='SYS_01_001.aspx?ID=" & rqMID & "';</script>")
    End Sub

    ''' <summary> 清除 帳號停用日 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub bt_clearDate_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles bt_clearDate.Click
        StopDate.Text = ""
    End Sub

    ''' <summary>檢查身分證</summary>
    ''' <param name="MyPage"></param>
    ''' <param name="IDNOtxt"></param>
    ''' <param name="Accounttxt"></param>
    ''' <param name="s_ReIDNOmsg"></param>
    Sub SUB_CHECK_IDNO(ByRef MyPage As Page, ByRef IDNOtxt As String, ByRef Accounttxt As String, ByRef s_ReIDNOmsg As String)
        Dim i_check_idno As Integer = 0 '未驗證身分證號 為0 成功為1
        Dim rqAct As String = TIMS.ClearSQM(MyPage.Request("act"))
        Dim rqACCOUNT As String = TIMS.ClearSQM(MyPage.Request("account"))

        IDNOtxt = TIMS.ClearSQM(IDNOtxt)
        Accounttxt = TIMS.ClearSQM(Accounttxt)
        'Dim ReIDNOmsg As String=""
        s_ReIDNOmsg = ""
        Dim sAccount As String = ""
        If rqACCOUNT <> "" Then
            sAccount = rqACCOUNT '修改動作
        Else
            If Accounttxt <> "" Then sAccount = Accounttxt '新增動作
        End If

        'act=add
        Select Case rqAct
            Case cst_rq_act_add
                i_check_idno = CheckAccountIDNO2("", IDNOtxt, s_ReIDNOmsg)  '驗證身分證號
            Case Else
                i_check_idno = CheckAccountIDNO2(sAccount, IDNOtxt, s_ReIDNOmsg)  '驗證身分證號 
        End Select

        Session(cst_sess_check_idno) = i_check_idno
    End Sub

    ''' <summary> 檢查身分證 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        reidno.Visible = False
        'Dim oReIDNOmsg As String=""
        'Call SUB_CHECK_IDNO(Me, IDNO.Text, nameid.Text, oReIDNOmsg)
        If oReIDNOmsg <> "" Then
            reidno.Visible = True
            reidno.Text = oReIDNOmsg
            'Common.MessageBox(Page, oReIDNOmsg)
        End If
    End Sub

    Protected Sub btn_back1_Click(sender As Object, e As EventArgs) Handles btn_back1.Click
        Dim url1 As String = String.Concat("SYS_01_001.aspx?ID=", TIMS.Get_MRqID(Me)) 'Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    ''' <summary>
    ''' 檢查帳號-server-帳號只能是英數字組合 myCheckID
    ''' </summary>
    Sub CheckAccountRepeat()
        '(用來取代原先"CheckID.aspx"頁面功能,by:20180926)
        Dim myMsg As String = ""
        'nameid.Text=TIMS.ClearSQM(nameid.Text)
        Dim v_nameidTxt As String = TIMS.ClearSQM(nameid.Text)
        If (v_nameidTxt <> nameid.Text) Then
            myMsg = cst_errorMsg7 '"請輸入帳號!!!"
            Common.MessageBox(Me, myMsg)
            Exit Sub
        End If
        If v_nameidTxt = "" Then
            myMsg = cst_NG_msg2 '"請輸入帳號!!!"
            Common.MessageBox(Me, myMsg)
            Exit Sub
        End If
        '帳號只能是英數字組合
        'nameid.Text=TIMS.ClearSQM(nameid.Text)
        If Not TIMS.CheckABC321acct(v_nameidTxt) Then
            myMsg = cst_errorMsg16b
            Common.MessageBox(Me, myMsg)
            Exit Sub
        End If

        Dim rqAct As String = TIMS.ClearSQM(Request("act"))
        Select Case rqAct 'Convert.ToString(Request("act"))
            Case cst_rq_act_add '"add"
                '新增-新規則
                If Not TIMS.CheckABC(Left(v_nameidTxt, 1)) Then
                    myMsg = cst_errorMsg17
                    Common.MessageBox(Me, myMsg)
                    Exit Sub
                End If
        End Select

        v_nameidTxt = UCase(v_nameidTxt) '轉大寫
        Dim sqlstr As String = ""
        sqlstr = "SELECT UPPER(ACCOUNT) ACCOUNT1 FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT)=@ACCOUNT "
        Dim sCmd As New SqlCommand(sqlstr, objconn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = v_nameidTxt
            dt1.Load(.ExecuteReader())
        End With

        Dim flagUseAccountName As Boolean = False '沒人使用
        If dt1.Rows.Count > 0 Then
            For Each dr1 As DataRow In dt1.Rows
                '轉換大寫後判斷
                If Convert.ToString(dr1("ACCOUNT1")) = v_nameidTxt Then
                    flagUseAccountName = True '有人使用
                    myMsg = cst_NG_msg1 '"該帳號已有人使用!!!"
                    Common.MessageBox(Me, myMsg)
                    Exit For
                End If
            Next
        End If

        If Not flagUseAccountName Then
            myMsg = cst_OK_msg1 '"您可以使用該帳號!!!"
            Common.MessageBox(Me, myMsg)
            Exit Sub
        End If
    End Sub

    ''' <summary>(額外補充Log寫入功能，by:2018/07/30)</summary>
    ''' <param name="sm"></param>
    ''' <param name="s_parma"></param>
    ''' <param name="oConn"></param>
    Public Shared Sub ADD_SYS_TRANS_LOG(ByRef sm As SessionModel, ByRef s_parma As Hashtable, ByRef oConn As SqlConnection)
        'Dim t_getSeqSql As String=" SELECT MAX(Seq) + 1 AS NewSeq FROM SYS_TRANS_LOG "
        'Dim t_dr As DataRow=DbAccess.GetOneRow(t_getSeqSql, objconn)
        'Dim t_Seq As Long=0
        'If Not t_dr Is Nothing Then t_Seq=Convert.ToInt64(Convert.ToString(t_dr("NewSeq")))
        Dim sAccountEdit As String = TIMS.GetMyValue2(s_parma, "sAccountEdit")
        Dim vACCOUNT As String = TIMS.GetMyValue2(s_parma, "ACCOUNT")
        Dim vROLEID As String = TIMS.GetMyValue2(s_parma, "ROLEID")
        Dim vLID As String = TIMS.GetMyValue2(s_parma, "LID")
        Dim vNAME As String = TIMS.GetMyValue2(s_parma, "NAME")
        Dim vPHONE As String = TIMS.GetMyValue2(s_parma, "PHONE")
        Dim vEMAIL As String = TIMS.GetMyValue2(s_parma, "EMAIL")
        Dim vORGID As String = TIMS.GetMyValue2(s_parma, "ORGID")
        Dim vISUSED As String = TIMS.GetMyValue2(s_parma, "ISUSED")
        Dim vIDNO As String = TIMS.GetMyValue2(s_parma, "IDNO")
        Dim vFuncPath As String = TIMS.GetMyValue2(s_parma, "FuncPath") '/SYS/01/SYS_01_001_add.aspx'
        Dim vTargetTable As String = TIMS.GetMyValue2(s_parma, "TargetTable") 'AUTH_ACCOUNT

        Dim vTransType As String = "Update"  '修改
        If sAccountEdit = "" Then vTransType = "Insert" '新增

        If sAccountEdit = "" Then
            '新增
            '==========
            Dim BeforeValues As String = ""
            BeforeValues &= "ACCOUNT=" & vACCOUNT
            BeforeValues &= ",ROLEID=" & vROLEID
            BeforeValues &= ",LID=" & vLID
            BeforeValues &= ",NAME=" & vNAME
            BeforeValues &= ",PHONE=" & vPHONE
            BeforeValues &= ",EMAIL=" & vEMAIL
            BeforeValues &= ",ORGID=" & vORGID
            BeforeValues &= ",ISUSED=" & vISUSED
            BeforeValues &= ",IDNO=" & vIDNO
            BeforeValues &= ",MODIFYACCT=" & sm.UserInfo.UserID
            BeforeValues &= ",MODIFYDATE=" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
            '==========
            Dim t_iSql As String = ""
            t_iSql = ""
            t_iSql &= " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
            t_iSql &= " VALUES(@SessionID, @TransTime, @FuncPath, @UserID, @TransType, @TargetTable, '', @BeforeValues, '') "
            Dim myParam As Hashtable = New Hashtable
            myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
            myParam.Add("FuncPath", vFuncPath)
            myParam.Add("TransType", vTransType)
            myParam.Add("TargetTable", vTargetTable)
            myParam.Add("BeforeValues", BeforeValues)
            myParam.Add("SessionID", sm.SessionID.ToString)
            myParam.Add("UserID", sm.UserInfo.UserID)
            Dim i_tCmd As New SqlCommand(t_iSql, oConn)
            DbAccess.HashParmsChange(i_tCmd, myParam)
            Dim i_rst As Integer = i_tCmd.ExecuteNonQuery()
            'DbAccess.ExecuteNonQuery(t_iSql, oConn, myParam)
        Else
            '修改
            Dim t_BeforeSql As String = " SELECT ACCOUNT,ROLEID,LID,NAME,PHONE,EMAIL,ORGID,ISUSED,IDNO,MODIFYACCT,MODIFYDATE FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT)='" & UCase(sAccountEdit) & "' "
            Dim t_BeforeTB As DataTable = DbAccess.GetDataTable(t_BeforeSql, oConn)
            Dim t_BeforeRow As DataRow = t_BeforeTB.Rows(0)
            '==========
            Dim BeforeValues As String = ""
            BeforeValues &= ("ACCOUNT=" & Convert.ToString(t_BeforeRow("ACCOUNT")))
            BeforeValues &= (",ROLEID=" & Convert.ToString(t_BeforeRow("ROLEID")))
            BeforeValues &= (",LID=" & Convert.ToString(t_BeforeRow("LID")))
            BeforeValues &= (",NAME=" & Convert.ToString(t_BeforeRow("NAME")))
            BeforeValues &= (",PHONE=" & Convert.ToString(t_BeforeRow("PHONE")))
            BeforeValues &= (",EMAIL=" & Convert.ToString(t_BeforeRow("PHONE")))
            BeforeValues &= (",ORGID=" & Convert.ToString(t_BeforeRow("ORGID")))
            BeforeValues &= (",ISUSED=" & Convert.ToString(t_BeforeRow("ISUSED")))
            BeforeValues &= (",IDNO=" & Convert.ToString(t_BeforeRow("IDNO")))
            BeforeValues &= (",MODIFYACCT=" & Convert.ToString(t_BeforeRow("MODIFYACCT")))
            BeforeValues &= (",MODIFYDATE=" & Convert.ToDateTime(Convert.ToString(t_BeforeRow("MODIFYDATE"))).ToString("yyyy-MM-dd HH:mm:ss.fff"))
            Dim AfterValues As String = ""
            AfterValues &= "ACCOUNT=" & vACCOUNT
            AfterValues &= ",ROLEID=" & vROLEID
            AfterValues &= ",LID=" & vLID
            AfterValues &= ",NAME=" & vNAME
            AfterValues &= ",PHONE=" & vPHONE
            AfterValues &= ",EMAIL=" & vEMAIL
            AfterValues &= ",ORGID=" & vORGID
            AfterValues &= ",ISUSED=" & vISUSED
            AfterValues &= ",IDNO=" & vIDNO
            AfterValues &= ",MODIFYACCT=" & sm.UserInfo.UserID
            AfterValues &= ",MODIFYDATE=" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Dim Conditions As String = ""
            Conditions &= "ACCOUNT=" & Convert.ToString(t_BeforeRow("ACCOUNT"))
            '==========
            Dim t_iSql As String = ""
            t_iSql += " INSERT INTO SYS_TRANS_LOG (SessionID, TransTime, FuncPath, UserID, TransType, TargetTable, Conditions, BeforeValues, AfterValues) "
            t_iSql += " VALUES(@SessionID ,@TransTime ,@FuncPath ,@UserID ,@TransType ,@TargetTable, @Conditions, @BeforeValues, @AfterValues) "
            Dim myParam As Hashtable = New Hashtable
            myParam.Add("TransTime", DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff"))
            myParam.Add("FuncPath", vFuncPath) '/SYS/01/SYS_01_001_add.aspx'
            myParam.Add("TransType", vTransType) 'Update'
            myParam.Add("TargetTable", vTargetTable) 'AUTH_ACCOUNT'
            myParam.Add("Conditions", Conditions) '" + Conditions + "'
            myParam.Add("BeforeValues", BeforeValues) '" + BeforeValues + "'
            myParam.Add("AfterValues", AfterValues)  '" + AfterValues + "'
            myParam.Add("SessionID", sm.SessionID.ToString)
            myParam.Add("UserID", sm.UserInfo.UserID)
            Dim i_tCmd As New SqlCommand(t_iSql, oConn)
            DbAccess.HashParmsChange(i_tCmd, myParam)
            Dim i_rst As Integer = i_tCmd.ExecuteNonQuery()
            'DbAccess.ExecuteNonQuery(t_iSql, oConn, myParam)
        End If
    End Sub

    ''' <summary>重啟帳號</summary>
    Sub UTL_BTNREUSED()
        Dim v_nameidTxt As String = TIMS.ClearSQM(nameid.Text)
        If v_nameidTxt = "" Then Exit Sub

        v_nameidTxt = UCase(v_nameidTxt) '轉大寫
        Dim u_parms As New Hashtable From {{"MODIFYACCT", sm.UserInfo.UserID}, {"ACCOUNT", v_nameidTxt}}
        Dim u_sql As String = ""
        u_sql &= " UPDATE AUTH_ACCOUNT" & vbCrLf
        u_sql &= " SET STOPDATE=DATEADD(YEAR,1,GETDATE()) ,ISUSED='Y',LAST_LOGINDATE=GETDATE()" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT)=@ACCOUNT" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        Common.MessageBox(Me, "帳號已重啟!!")
    End Sub

    Protected Sub BtnReUsed1_Click(sender As Object, e As EventArgs) Handles BtnReUsed1.Click
        UTL_BTNREUSED()
    End Sub

    ''' <summary> 資安檢核-1 </summary>
    Sub Critical_Issues_1()
        'openhttps.Value=TIMS.ClearSQM(openhttps.Value)
        g_parms = New Hashtable
        g_parms.Clear()
        g_parms.Add("1.txtUserId.Text", nameid.Text)
        g_parms.Add("2.txtUserIdno.Text", IDNO.Text)
        g_parms.Add("3.txtUserEMAIL.Text", email.Text)
        'g_parms.Add("4.txtVCode.Text", txtVCode.Text)

        sTxtIdUnchange = TIMS.ClearSQM(nameid.Text)
        sTxtIDNO = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))
        sTxtEMAIL = TIMS.ChangeEmail(TIMS.ClearSQM(email.Text))
        'sVCode=TIMS.ClearSQM(txtVCode.Text)

        nameid.Text = sTxtIdUnchange
        IDNO.Text = sTxtIDNO
        email.Text = sTxtEMAIL
        'txtVCode.Text=sVCode

        g_parms.Add("sTxtIdUnchg", sTxtIdUnchange)
        g_parms.Add("sTxtIDNO", sTxtIDNO)
        g_parms.Add("sTxtEMAIL", sTxtEMAIL)
        'g_parms.Add("sVCode", sVCode)
        '"https://ojtims.wda.gov.tw/emailChgPwd?GID={0}&OTK={1}&OTK2={2}"
        'TIMS.sUtl_404NOTFOUND(Page, objconn)
    End Sub

    ''' <summary> 寄送密碼函 </summary>
    Sub SendPWDletter()
        Call Critical_Issues_1()

        '未使用測試-正式者-檢核 true:攻擊異常達標
        If TIMS.Utl_ChkHISTORY1(Me, objconn, sTxtIdUnchange) Then Exit Sub

        Dim flag_test_ENVC As Boolean = TIMS.CHK_IS_TEST_ENVC() '檢測為測試環境:true 正式環境為:false
        'If sVCode="" OrElse txtUserId.Text="" OrElse txtUserIdno.Text="" OrElse txtUserEMAIL.Text="" Then
        '    '若不為空可接受字元檢核'防止駭客攻擊(紀錄) 
        '    Call TIMS.sUtl_SaveLoginData1(Me, objconn)
        'End If

        'Const cst_alert_msg_1b As String="查無此帳號!"
        'Const cst_alert_msg_3 As String="此帳號未設定EMAIL!"
        'sVCode=TIMS.ClearSQM(txtVCode.Text)
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        If nameid.Text = "" OrElse IDNO.Text = "" OrElse email.Text = "" Then
            sm.LastErrorMessage = cst_alert_msg_2c
            Return
        End If

        'AuthUtil.LoginLog(sTxtIdUnchange, False)

        Dim drAA As DataRow = TIMS.sUtl_GetAccount(nameid.Text, objconn, False)
        If drAA Is Nothing Then
            sm.LastErrorMessage = cst_alert_msg_1g
            Exit Sub
        End If
        Dim vEMAIL As String = TIMS.ChangeEmail(TIMS.ClearSQM(drAA("EMAIL")))
        If vEMAIL = "" Then
            sm.LastErrorMessage = cst_alert_msg_3
            Exit Sub
        End If
        If Not LCase(vEMAIL).Equals(LCase(sTxtEMAIL)) Then
            sm.LastErrorMessage = cst_alert_msg_99ot '"資料填寫有誤，請重新輸入，或洽系統管理員，謝謝!"
            Exit Sub
        End If

        Dim vIDNO As String = TIMS.ChangeIDNO(TIMS.ClearSQM(drAA("IDNO")))
        If Not vIDNO.Equals(sTxtIDNO) Then
            sm.LastErrorMessage = cst_alert_msg_99ot '"資料填寫有誤，請重新輸入，或洽系統管理員，謝謝!"
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

        'sm.LastResultMessage=String.Format("系統已寄發密碼重設通知函至您的E-Mail帳號! [{0}]", vEMAIL)
        sm.LastResultMessage = cst_alert_msg_98ot '"系統已寄發密碼重設通知函至您的E-Mail帳號! "

        If flag_test_ENVC Then
            Dim xMybody As String = TIMS.GetMyValue2(htSS, "xMybody")
            If xMybody <> "" Then Response.Write(xMybody)
        End If

        ''eforgetPwd
        ''Dim redirectUrl As String=ResolveUrl("~/login") 'String.Empty
        'Dim redirectUrl As String=ResolveUrl("~/eforgetPwd") 'String.Empty
        ''redirectUrl=ResolveUrl("~/login")
        'If Not IsNothing(redirectUrl) AndAlso Not String.IsNullOrEmpty(redirectUrl) Then
        '    '檢核成功, 導向首頁
        '    Response.Redirect(redirectUrl)
        '    If Not flag_test_ENVC Then Response.Redirect(redirectUrl)
        'End If

    End Sub

    ''' <summary>
    ''' 寄送密碼函
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnSendPXDEMAIL_Click(sender As Object, e As EventArgs) Handles BtnSendPXDEMAIL.Click
        Call SendPWDletter()
    End Sub

    ''' <summary>
    ''' 重設E-MAIL
    ''' </summary>
    Sub UTL_BTNREEMAIL()
        'email.Text=TIMS.ClearSQM(email.Text)
        email.Text = TIMS.ChangeEmail(TIMS.ClearSQM(email.Text)) '轉EMAIL
        If email.Text = "" Then
            Common.MessageBox(Me, cst_errorMsg31)
            Exit Sub
        End If
        If Not TIMS.CheckEmail(email.Text) Then
            Common.MessageBox(Me, cst_errorMsg32)
            Exit Sub
        End If

        Dim v_nameidTxt As String = TIMS.ClearSQM(nameid.Text)
        If v_nameidTxt = "" Then Exit Sub

        v_nameidTxt = UCase(v_nameidTxt) '轉大寫
        Dim u_sql As String = ""
        u_sql = "" & vbCrLf
        u_sql &= " UPDATE AUTH_ACCOUNT" & vbCrLf
        u_sql &= " SET EMAIL=@EMAIL ,LAST_LOGINDATE=GETDATE()" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " FROM AUTH_ACCOUNT WHERE UPPER(ACCOUNT)=@ACCOUNT" & vbCrLf
        Dim u_parms As New Hashtable
        u_parms.Clear()
        u_parms.Add("EMAIL", email.Text)
        u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_parms.Add("ACCOUNT", v_nameidTxt)
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)

        Common.MessageBox(Me, "E-MAIL已重新設定!!")
    End Sub

    ''' <summary>
    ''' 重設E-MAIL
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub BtnReEmail1_Click(sender As Object, e As EventArgs) Handles BtnReEmail1.Click
        UTL_BTNREEMAIL()
    End Sub

    ''' <summary> 帳號設為停用，並註記停用日期 </summary>
    ''' <param name="s_ACCOUNT"></param>
    Sub UPDATE_STOP_ACCOUNT(ByVal s_ACCOUNT As String)
        'Dim s_ACCOUNT As String=TIMS.ClearSQM(s_ACCOUNT)
        s_ACCOUNT = TIMS.ClearSQM(s_ACCOUNT)
        If s_ACCOUNT = "" Then Exit Sub

        s_ACCOUNT = UCase(s_ACCOUNT) '轉大寫
        Dim u_parms As New Hashtable
        u_parms.Clear()
        u_parms.Add("MODIFYACCT", sm.UserInfo.UserID)
        u_parms.Add("ACCOUNT", s_ACCOUNT)
        Dim u_sql As String = ""
        u_sql = "" & vbCrLf
        u_sql &= " UPDATE AUTH_ACCOUNT" & vbCrLf
        u_sql &= " SET ISUSED='N'" & vbCrLf
        u_sql &= " ,StopDate=GETDATE()" & vbCrLf
        'u_sql &= " ,LAST_LOGINDATE=GETDATE()" & vbCrLf
        u_sql &= " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE=GETDATE()" & vbCrLf
        u_sql &= " FROM AUTH_ACCOUNT" & vbCrLf
        u_sql &= " WHERE 1=1" & vbCrLf
        u_sql &= " AND ISUSED='Y'" & vbCrLf
        u_sql &= " AND UPPER(ACCOUNT)=@ACCOUNT" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, u_parms)
    End Sub

    Protected Sub BtnrResetPXD_Click(sender As Object, e As EventArgs) Handles BtnrResetPXD.Click

        nameid.Text = TIMS.ClearSQM(nameid.Text)
        Dim s_PXD1 As String = TIMS.RGetHashCode()

        Dim s_parma As New Hashtable
        s_parma.Clear()
        s_parma.Add("ACCOUNT", nameid.Text)
        s_parma.Add("HASHPWD1", TIMS.CreateHash(s_PXD1))
        s_parma.Add("PXSSENC1", TIMS.EncryptAes(s_PXD1))
        '/*SDMATYPE 密碼寄信種類-1:新設密碼-2:忘記密碼(修改密碼)-3:修改密碼 */ SENDMAILTYPE /SDMATYPE
        s_parma.Add("SDMATYPE", "4")
        s_parma.Add("MODIFYACCT", sm.UserInfo.UserID)
        '修改密碼-儲存 
        Dim htSS As Hashtable = TIMS.INSERT_PXSSWARDHIS(objconn, s_parma)

        'sm.LastResultMessage="密碼修改完成"
        Dim s_msg As String = String.Format("密碼修改完成!![{0}]", s_PXD1)
        Common.MessageBox(Me, s_msg)

    End Sub
End Class