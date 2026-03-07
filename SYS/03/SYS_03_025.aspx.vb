Partial Class SYS_03_025
    Inherits AuthBasePage

    Dim vs_itemname As String = ""
    Dim vs_orgid As String = ""
    Dim vs_txtaccount As String = ""
    Dim vs_txtname As String = ""
    Dim vs_planid As String = ""
    Dim vs_years As String = ""
    Dim vs_distid As String = ""
    Dim vs_mmname As String = ""


    Dim dtPlan As DataTable '(KEY_PLAN)

    Const cst_titleM1 As String = "系統預設"

    Dim chkEnableAll As CheckBox
    Dim trID As Integer = 0
    'Const cst_account_snoopy As String = "snoopy" '特殊權限使用者
    Const cst_account As String = "account" 'ViewState(cst_account) '帳號(登入者)
    Const cst_userid As String = "userid" 'ViewState(cst_userid) '所選擇要設定的帳號。
    Const cst_username As String = "username" 'ViewState(cst_userid) '所選擇要設定的帳號姓名。
    'Const cst_userGDistID As String = "userGDistID" 'ViewState(cst_userid) '所選擇要設定的帳號使用轄區。

    Const cst_mainmenu As String = "mainmenu" 'ViewState(cst_mainmenu)
    Const cst_vsgid As String = "gid" 'ViewState(cst_vsgid)

    'Dim oCmd As SqlCommand
    'arrFun = FunSort.Split(",")
    Dim arrFun As String()  '= {"TC", "SD", "CP", "TR", "CM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"} 'fun排列順序
    'Dim FunSort As String = System.Configuration.ConfigurationSettings.AppSettings("FunSort")
    'Const cst_DistID6 As String = "000,001,002,003,004,005,006"
    Dim vMsg As String = "" '暫存字

#Region "Sub"
    '初始化設定群組維護畫面
    Private Sub Clear_GroupItems()
        ViewState(cst_userid) = Nothing

        Call sUtl_NoShowAll()
        tb_Query.Visible = True
        'tb_Group.Visible = False
        'tb_GroupFun.Visible = False
        'tb_Function.Visible = False
    End Sub

    '初始化設定功能維護畫面
    Private Sub Clear_FunctionItems()
        ViewState(cst_mainmenu) = Nothing

        ddlFun.SelectedIndex = 0

        Call sUtl_NoShowAll()
        'tb_Query.Visible = False
        'tb_Group.Visible = False
        tb_GroupFun.Visible = True
        'tb_Function.Visible = False
    End Sub

    '顯示DropDownList資料
    Private Sub Show_DropDownList(ByVal strFlag As String, ByRef objDDL As DropDownList, ByVal textField As String, ByVal valueField As String)
        Dim sql As String = ""
        Call TIMS.OpenDbConn(objconn)
        'If conn.State.Closed Then conn.Open()
        Dim flag_PleaseChoose2 As Boolean = False
        Select Case strFlag
            Case "PlanYears"
                sql = "select distinct Years from ID_Plan where dbo.NVL(Years,' ')<>' ' order by Years DESC"

            Case "DistID"
                sql = "select DistID,Name from ID_District order by DistID Asc "

            Case "PlanID"
                sql = "" & vbCrLf
                sql &= " select distinct a.PlanID" & vbCrLf
                sql &= " ,a.TPlanID" & vbCrLf
                'sql &= " ,A.YEARS+C.NAME+B.PLANNAME+A.SEQ+NVL2(TRIM(A.SUBTITLE),'(' +A.SUBTITLE+ ')','') PLANNAME" & vbCrLf
                sql &= " ,cast(A.YEARS as varchar)+C.NAME+B.PLANNAME+cast(A.SEQ as varchar) + case when A.SUBTITLE is null or replace(A.SUBTITLE,' ','') = '' then '' else  '(' +A.SUBTITLE+ ')' end PLANNAME " & vbCrLf
                sql &= " from ID_Plan a " & vbCrLf
                sql &= " join Key_Plan b on b.TPlanID=a.TPlanID" & vbCrLf
                sql &= " join ID_District c on c.DistID=a.DistID" & vbCrLf
                sql &= " join Auth_AccRWPlan d on d.PlanID=a.PlanID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " AND a.Years='" & vs_years & "'" '依年度(登入者)
                sql &= " AND a.DistID='" & vs_distid & "'" '依轄區(登入者)

                'SNOOPY為全部
                If Convert.ToString(ViewState(cst_account)) <> "" Then
                    sql += "and d.Account='" & Convert.ToString(ViewState(cst_account)) & "' " '帳號(登入者)
                End If

                sql += "order by a.PlanID asc "
                'flag_PleaseChoose2 = True
            Case "OrgID"
                sql = ""
                sql &= " select distinct a.OrgID,a.OrgName "
                sql &= " from Org_OrgInfo a join Auth_Relship b on b.OrgID=a.OrgID "
                '如果不濾除該計畫沒有申請帳號的單位時，去除以下兩行即可
                sql &= " join Auth_Account c on c.OrgID=a.OrgID "
                sql &= " join Auth_AccRWPlan d on d.PlanID=b.PlanID and d.Account=c.Account "
                '=======================================================
                sql &= " where b.PlanID='" & Convert.ToString(vs_planid) & "' " '依登入計畫(登入者)
                sql &= " order by a.OrgName asc "
                'flag_PleaseChoose2 = True
        End Select

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If Convert.ToString(vs_planid) = "" AndAlso strFlag = "OrgID" Then
            Common.MessageBox(Me, "該年度無計畫可選擇!")
            Return
        End If
        Dim VF1 As String = If(dt.Rows.Count > 0, dt.Rows(0)(valueField), "")
        objDDL.Items.Clear()
        If dt.Rows.Count = 0 Then Return
        objDDL.DataSource = dt
        objDDL.DataTextField = textField
        objDDL.DataValueField = valueField

        If objDDL.Items.Count > 0 AndAlso objDDL.SelectedValue <> "" Then
            objDDL.SelectedIndex = -1
            If VF1 <> "" Then objDDL.SelectedValue = VF1
        End If
        objDDL.DataBind()

        If flag_PleaseChoose2 Then
            objDDL.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose2, ""))
        End If
        Common.SetListItem(objDDL, "")
    End Sub

    '重新載入機構DropDownList資料
    Private Sub Renew_ListOrgID(ByVal objDDL As DropDownList, ByVal strAccount As String, ByVal roleid As Integer)
        Dim Roles() As Integer = {0, 1}
        Dim sql As String = ""

        Select Case roleid
            Case 0
                sql = ""
                sql &= " select a.OrgID,a.OrgName from Org_OrgInfo a "
                sql &= " join Auth_Relship b on b.OrgID=a.OrgID where b.OrgLevel<=1 "
                sql &= " order by b.OrgLevel,b.DistID,b.RID,a.OrgID asc"
            Case 1
                sql = ""
                sql &= " select a.OrgID,a.OrgName from Org_OrgInfo a "
                sql &= " join Auth_Account b on b.OrgID=a.OrgID where b.Account= @account "
                sql &= " order by b.LID,b.RoleID,a.OrgID asc"
        End Select
        Dim oCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        If Array.IndexOf(Roles, roleid) <> -1 Then
            Dim dt As New DataTable
            With oCmd
                .Parameters.Clear()
                If roleid = 1 Then
                    .Parameters.Add("account", SqlDbType.VarChar).Value = strAccount
                End If
                dt.Load(.ExecuteReader())
            End With
            If dt.Rows.Count > 0 Then
                Dim intCnt As Integer = 0
                For Each dr As DataRow In dt.Rows
                    objDDL.Items.Insert(intCnt, New ListItem(Convert.ToString(dr("OrgName")), Convert.ToString(dr("OrgID"))))
                    intCnt += 1
                Next
            End If
        End If

    End Sub

    '查詢
    Sub sSearch1()

        vs_years = TIMS.ClearSQM(list_Years.SelectedValue)
        vs_distid = TIMS.ClearSQM(list_DistID.SelectedValue)
        vs_planid = TIMS.ClearSQM(list_PlanID.SelectedValue)
        vs_orgid = TIMS.GetListValue(list_OrgID)
        'vs_orgid = TIMS.ClearSQM(list_OrgID.SelectedValue)
        vs_txtaccount = TIMS.ClearSQM(txt_Account.Text)
        vs_txtname = TIMS.ClearSQM(txt_Name.Text)

        If vs_planid = "" Then
            Common.MessageBox(Me, "請選擇計畫代碼!")
            Exit Sub
        End If
        Hid_PLANID.Value = vs_planid
        Dim strArrGID As String = "" '記錄群組ID
        Dim strArrGName As String = "" '記錄群組名稱

        '查詢帳號資料
        Dim sql As String = ""
        sql = ""
        sql &= " select Account,Name,RoleID "
        'sql &= " ,cast(null as VARCHAR2(2000 CHAR)) GID "
        'sql &= " ,cast(null as NVARCHAR2(2000))  GName "
        sql &= " ,cast(null as VARCHAR(2000)) GID "
        sql &= " ,cast(null as NVARCHAR(2000)) GName "
        sql &= " from Auth_Account "
        sql &= " where 1=1 "
        sql &= " and IsUsed='Y' "
        sql &= " and Account in (select Account from Auth_AccRWPlan where PlanID= @PlanID) "
        '除 snoopy 外不可設定及看到 snoopy
        If sm.UserInfo.UserID <> str_superuser1 Then
            sql &= " and Account<>'" & str_superuser1 & "' "
        End If
        If vs_orgid <> "" Then
            sql &= " and OrgID= @OrgID "
        End If
        If vs_txtaccount <> "" Then
            sql &= " and Account like @Account "
        End If
        If vs_txtname <> "" Then
            sql &= " and Name like @Name "
        End If
        Dim oCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("PlanID", SqlDbType.Int).Value = vs_planid
            If vs_orgid <> "" Then
                .Parameters.Add("OrgID", SqlDbType.VarChar).Value = vs_orgid
            End If
            If vs_txtaccount <> "" Then
                .Parameters.Add("Account", SqlDbType.VarChar).Value = "%" & vs_txtaccount & "%"
            End If
            If vs_txtname <> "" Then
                .Parameters.Add("Name", SqlDbType.NVarChar).Value = "%" & vs_txtname & "%"
            End If
            'dt.Load(.ExecuteReader())
            dt = DbAccess.GetDataTable(oCmd.CommandText, objconn, oCmd.Parameters)
        End With

        '代入帳號群組資料
        sql = ""
        sql &= " select b.GID,convert(varchar(10),b.GType) GType,b.GName"
        sql &= " ,c.Name"
        sql &= " ,b.GDistID"
        sql &= " from Auth_GroupAcct a"
        sql &= " join Auth_Group b on b.GID=a.GID "
        sql &= " left join ID_District c on c.DistID=b.GDistID "
        sql &= " where a.Account= @Account "
        sql &= " order by b.GDistID,b.GType"
        Dim oCmd2 As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        For i As Integer = 0 To dt.Rows.Count - 1
            Dim dr As DataRow = dt.Rows(i)
            Dim dtChk As New DataTable
            With oCmd2
                .Parameters.Clear()
                .Parameters.Add("Account", SqlDbType.VarChar).Value = Convert.ToString(dr("account"))
                'dtChk.Load(.ExecuteReader())
                dtChk = DbAccess.GetDataTable(oCmd2.CommandText, objconn, oCmd2.Parameters)
            End With

            For Each drCkj As DataRow In dtChk.Rows
                Select Case Convert.ToString(drCkj("gtype"))
                    Case "0"
                        drCkj("gtype") = "署"'dtChk.Rows(j)("gtype") = "局"
                    Case "1"
                        drCkj("gtype") = "分署"'drCkj("gtype") = "中心"
                    Case "2"
                        drCkj("gtype") = "委外"
                    Case Else
                        drCkj("gtype") = ""
                End Select

                If Convert.ToString(drCkj("name")) = "" Then drCkj("name") = cst_titleM1
                If strArrGID <> "" Then strArrGID &= ","
                strArrGID &= Convert.ToString(drCkj("gid"))

                If strArrGName <> "" Then strArrGName &= "<br>"
                strArrGName &= "&nbsp;" & Convert.ToString(drCkj("gtype")) & "-" & Convert.ToString(drCkj("gname")) & "(" & drCkj("name") & ")"
            Next

            dr("gid") = strArrGID
            dr("gname") = strArrGName
            strArrGID = ""
            strArrGName = ""
        Next

        lab_Msg.Visible = True
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            lab_Msg.Visible = False
            DataGrid1.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If

    End Sub

    '顯示建檔單位下拉。(ddlGDist2) 配合 Load_Group(ByVal sGDistID As String)
    Sub Load_Group_ddlGDist2()
        Dim sql As String = ""
        If sm.UserInfo.UserID = str_superuser1 Then
            'snoopy專用
            'Dim sql As String = ""
            sql = "" & vbCrLf
            sql &= " select distinct dbo.NVL(b.GDistID,'X') GDistID" & vbCrLf
            sql &= " ,CASE WHEN b.GDistID IS NULL THEN '(系統預設)' ELSE ISNULL(d.Name,CONCAT('(',b.GDistID,')')) END GDistName" & vbCrLf
            sql &= " from Auth_Group b" & vbCrLf
            sql &= " LEFT JOIN ID_District d on d.distid=b.GDistID" & vbCrLf
            sql &= " where b.GValid='1' and b.GState not in ('D')" & vbCrLf
        Else
            If sm.UserInfo.RoleID = "0" Or sm.UserInfo.RoleID = "1" Then
                '系統管理者
                sql = "" & vbCrLf
                sql &= " select distinct dbo.NVL(b.GDistID,'X') GDistID " & vbCrLf
                sql &= " ,CASE WHEN b.GDistID IS NULL THEN '(系統預設)' ELSE ISNULL(d.Name,CONCAT('(',b.GDistID,')')) END GDistName" & vbCrLf
                sql &= " from Auth_Group b " & vbCrLf
                sql &= " left join ID_District d on d.distid=b.GDistID " & vbCrLf
                sql &= " where b.GValid='1' and b.GState not in ('D')" & vbCrLf
                sql &= " and (b.GDistID='" & sm.UserInfo.DistID & "' or b.GDistID is null)" & vbCrLf
            Else
                '一般使用者
                sql = "" & vbCrLf
                sql &= " select distinct dbo.NVL(b.GDistID,'X') GDistID" & vbCrLf
                sql &= " ,CASE WHEN b.GDistID IS NULL THEN '(系統預設)' ELSE ISNULL(d.Name,CONCAT('(',b.GDistID,')')) END GDistName" & vbCrLf
                sql &= " from Auth_GroupAcct a" & vbCrLf
                sql &= " join Auth_Group b on b.GID=a.GID" & vbCrLf
                sql &= " left join ID_District d on d.distid=b.GDistID" & vbCrLf
                sql &= " where b.GValid='1' and b.GState not in ('D')" & vbCrLf
                sql &= " and a.Account='" & sm.UserInfo.UserID & "'" & vbCrLf
            End If

            If sm.UserInfo.DistID <> "000" Then
                sql &= " and b.GType <>'0' " & vbCrLf
            Else
                sql &= " and b.GType not in ('1','2') " & vbCrLf
            End If
        End If
        sql += "order by 1 "

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        With ddlGDist2
            .DataSource = dt
            .DataTextField = "GDistName"
            .DataValueField = "GDistID"
            .DataBind()
            .Items.Insert(0, New ListItem("==請選擇==", "")) '初始值設0
        End With
    End Sub

    '代入群組list 代入建檔單位
    Private Sub Load_Group(ByVal sUse_Sch_Value As String)


        lab2UserID.Text = ViewState(cst_userid)
        lab2Name.Text = ViewState(cst_username)
        Dim sql As String = ""
        Call TIMS.OpenDbConn(objconn)

        If flgROLEIDx0xLIDx0 AndAlso ViewState(cst_userid) = str_superuser1 Then
            '登入者 是snoopy 且設定snoopy群組專用
            sql = "" & vbCrLf
            sql &= " select b.GID,b.GDistID,b.GType,b.GName,b.GNote " & vbCrLf
            sql &= " ,a.GTPLANID" & vbCrLf
            sql &= " ,a.GID AGID" & vbCrLf
            sql &= " from Auth_Group b " & vbCrLf '不限定。
            sql &= " left join (SELECT * FROM AUTH_GROUPACCT WHERE 1<>1) a  on b.GID=a.GID" & vbCrLf
            sql &= " where 1=1" & vbCrLf
            sql &= " and b.GValid='1'" & vbCrLf
            sql &= " and b.GState not in ('D')" & vbCrLf '不限定。

        Else
            'DIST 設定
            'Dim userGDistID As String = sm.UserInfo.DistID
            'If flgROLEIDx0xLIDx0 Then
            '    If ViewState(cst_userGDistID) <> userGDistID Then
            '        userGDistID = ViewState(cst_userGDistID)
            '    End If
            'End If

            Dim userGDistID As String = TIMS.Get_AccDistID(ViewState(cst_userid), objconn)

            Select Case Convert.ToString(sm.UserInfo.RoleID)
                Case "0", "1"
                    '系統管理者
                    sql = ""
                    sql &= " SELECT b.GID,b.GDistID,b.GType,b.GName,b.GNote" & vbCrLf
                    sql &= " ,a.GTPLANID" & vbCrLf
                    sql &= " ,a.GID AGID" & vbCrLf '該使用者已設定該群組
                    sql &= " FROM Auth_Group b" & vbCrLf
                    sql &= " LEFT JOIN Auth_GroupAcct a on b.GID=a.GID AND Account='" & ViewState(cst_userid) & "'" & vbCrLf
                    sql &= " where 1=1" & vbCrLf
                    sql &= " and b.GValid='1'" & vbCrLf
                    sql &= " and b.GState not in ('D')" & vbCrLf
                    If flgROLEIDx0xLIDx0 Then
                        '該帳號取得登入的轄區
                        sql &= " and (b.GDistID IN (" & userGDistID & ") OR b.GDistID is null)" & vbCrLf '限定轄區與系統預設。
                    Else
                        '該使用者目前登入的轄區
                        sql &= " and (b.GDistID ='" & sm.UserInfo.DistID & "' OR b.GDistID is null)" & vbCrLf '限定轄區與系統預設。
                    End If
                    'sql += " and (b.GDistID='" & sm.UserInfo.DistID & "' or b.GDistID is null)" & vbCrLf '限定轄區與系統預設。
                Case Else
                    '一般使用者
                    sql = ""
                    sql &= " SELECT DISTINCT b.GID,b.GDistID,b.GType,b.GName,b.GNote" & vbCrLf
                    sql &= " ,a.GTPLANID" & vbCrLf
                    sql &= " ,a.GID AGID" & vbCrLf '該使用者已設定該群組
                    sql &= " FROM Auth_GroupAcct a" & vbCrLf
                    sql &= " JOIN Auth_Group b on b.GID=a.GID" & vbCrLf
                    sql &= " where 1=1" & vbCrLf
                    sql &= " and b.GValid='1'" & vbCrLf
                    sql &= " and b.GState not in ('D')" & vbCrLf
                    sql &= " and a.Account='" & sm.UserInfo.UserID & "'" & vbCrLf '限定使用者。
            End Select

            If Not flgROLEIDx0xLIDx0 Then
                'GType: 使用單位:0:署(局),1:分署(中心),2:委訓
                If sm.UserInfo.DistID <> "000" Then
                    sql &= " and b.GType <>'0'" & vbCrLf '不可以用署(局)
                Else
                    sql &= " and b.GType not in ('1','2')" & vbCrLf '不可以用分署(中心)及委訓。
                End If
            End If

        End If

        If sUse_Sch_Value = "Y" Then
            Dim v_sGDIST2 As String = TIMS.GetListValue(ddlGDist2)
            Dim v_sGROUPTYPE2 As String = TIMS.GetListValue(ddlGroupType2)
            If v_sGDIST2 <> "" Then
                Select Case v_sGDIST2
                    Case "X" '(系統預設)'
                        sql &= " and b.GDistID IS NULL" & vbCrLf
                    Case Else
                        sql &= " and b.GDistID ='" & v_sGDIST2 & "'" & vbCrLf
                End Select
            End If
            If v_sGROUPTYPE2 <> "" Then sql &= " and b.GType='" & v_sGROUPTYPE2 & "'" & vbCrLf
        End If
        sql &= " ORDER BY b.GDistID,b.GType,b.GID asc" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        lab_Msg2.Visible = True
        btn_SaveGroup.Visible = False
        DataGrid2.Visible = False
        If dt.Rows.Count = 0 Then Return
        'If dt.Rows.Count > 0 Then End If

        lab_Msg2.Visible = False
        btn_SaveGroup.Visible = True
        DataGrid2.Visible = True

        DataGrid2.DataSource = dt
        DataGrid2.DataKeyField = "GID"
        DataGrid2.DataBind()

    End Sub

    '代入個人所有群組
    Sub Load_GroupFun()
        lab3UserID.Text = ViewState(cst_userid)
        lab3Name.Text = ViewState(cst_username)

        Dim sql As String = ""
        sql = ""
        sql &= " select a.GID,b.GDistID,b.GType,b.GName"
        sql &= " from Auth_GroupAcct a"
        sql &= " join Auth_Group b on b.GID=a.GID "
        sql &= " where 1=1"
        sql &= " and a.Account= @Account"
        sql &= " order by b.GDistID,b.GType,b.GID asc"
        Dim oCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("Account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
            dt.Load(.ExecuteReader())
        End With

        DataGrid3.DataSource = dt 'ds.Tables(0)
        DataGrid3.DataKeyField = "GID"
        DataGrid3.DataBind()

    End Sub

    '代入功能list
    Private Sub Load_Function(ByVal strGid As String)
        lab4UserID.Text = ViewState(cst_userid)
        lab4Name.Text = ViewState(cst_username)

        '組清單用DataTable
        Dim dt_fun As New DataTable
        dt_fun.Columns.Add(New DataColumn("gid"))
        dt_fun.Columns.Add(New DataColumn("funid"))
        dt_fun.Columns.Add(New DataColumn("name"))
        dt_fun.Columns.Add(New DataColumn("kind"))
        dt_fun.Columns.Add(New DataColumn("levels"))
        dt_fun.Columns.Add(New DataColumn("parent"))
        dt_fun.Columns.Add(New DataColumn("sort"))
        dt_fun.Columns.Add(New DataColumn("memo"))
        dt_fun.Columns.Add(New DataColumn("subs"))
        dt_fun.Columns.Add(New DataColumn("spage"))

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select a.funid" & vbCrLf
        sql &= " ,a.name,a.spage" & vbCrLf
        sql &= " ,a.kind,a.levels" & vbCrLf
        sql &= " ,(case a.levels when '0' then CONVERT(varchar, a.funid) else a.parent end) parent" & vbCrLf
        sql &= " ,a.sort,a.memo" & vbCrLf
        sql &= " ,(case a.levels when '0' then (select count(funid) from id_function where parent=a.funid) else 0 end) subs" & vbCrLf
        sql &= " ,b.gdid " & vbCrLf
        sql &= " from id_function a " & vbCrLf
        sql &= " left join auth_groupdfun b on b.funid=a.funid " & vbCrLf
        sql &= " 	and gid= @gid and account= @account " & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " 	and upper(a.valid)='Y' " & vbCrLf
        sql &= " 	and dbo.NVL(a.FState,' ') not in ('D') " & vbCrLf
        sql &= " 	and a.kind= @kind " & vbCrLf
        sql &= " 	and exists (" & vbCrLf
        sql &= " 		select 'x' from auth_groupfun x where x.gid= @gid and x.funid=a.funid" & vbCrLf
        sql &= " 	)" & vbCrLf
        sql &= " order by " & vbCrLf
        sql &= " 	a.kind,parent,a.levels,a.sort" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        For i As Integer = 0 To arrFun.Length - 1
            Dim dtS1 As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("gid", SqlDbType.VarChar).Value = strGid
                .Parameters.Add("account", SqlDbType.VarChar).Value = ViewState(cst_userid)

                If ddlFun.SelectedIndex <> 0 Then
                    .Parameters.Add("kind", SqlDbType.VarChar).Value = ddlFun.SelectedValue
                Else
                    .Parameters.Add("kind", SqlDbType.VarChar).Value = arrFun(i)
                End If
                dtS1.Load(.ExecuteReader())
                '.Fill(ds)
            End With

            For j As Integer = 0 To dtS1.Rows.Count - 1
                Dim dr As DataRow = dt_fun.NewRow
                dt_fun.Rows.Add(dr)
                dr("gid") = strGid
                dr("funid") = dtS1.Rows(j)("funid")
                dr("name") = dtS1.Rows(j)("name")
                dr("kind") = dtS1.Rows(j)("kind")
                dr("levels") = dtS1.Rows(j)("levels")
                dr("parent") = dtS1.Rows(j)("parent")
                dr("sort") = dtS1.Rows(j)("sort")
                dr("memo") = dtS1.Rows(j)("memo")
                dr("subs") = dtS1.Rows(j)("subs")
                dr("spage") = dtS1.Rows(j)("spage")
            Next
            dtS1.Rows.Clear()

            If ddlFun.SelectedIndex <> 0 Then
                Exit For
            End If
        Next

        '取得功能種類
        sql = "select distinct kind from ID_Function order by kind"
        Dim sCmdKd As New SqlCommand(sql, objconn)
        Dim dtKd As New DataTable
        With sCmdKd
            .Parameters.Clear()
            dtKd.Load(.ExecuteReader())
        End With

        'If Not ds.Tables("select_kind") Is Nothing Then
        '    If ds.Tables("select_kind").Rows.Count > 0 Then
        '        dt_kind = ds.Tables("select_kind")
        '    End If
        'End If

        For i As Integer = 0 To dtKd.Rows.Count - 1
            Dim dr As DataRow = dtKd.Rows(i)
            ViewState(dr("kind").ToString) = dt_fun.Select("kind='" & dr("kind").ToString & "'").Length
        Next

        lab_Msg4.Visible = True
        btn_SaveOption.Visible = False
        Datagrid4.Visible = False
        If dt_fun.Rows.Count > 0 Then
            lab_Msg4.Visible = False
            btn_SaveOption.Visible = True
            Datagrid4.Visible = True

            vs_mmname = ""
            Datagrid4.DataSource = dt_fun
            Datagrid4.DataKeyField = "FunID"
            Datagrid4.DataBind()
        End If

    End Sub
#End Region

#Region "Function"
    '功能類別對照 取得中文名稱
    'Private Function Get_MainMenuName(ByVal tmpCode As String) As String
    '    Dim rst As String = ""

    '    Select Case UCase(tmpCode)
    '        Case "TC"
    '            rst = "訓練機構管理"
    '        Case "SD"
    '            rst = "學員動態管理"
    '        Case "CP"
    '            rst = "查核/績效管理"
    '        Case "TR"
    '            rst = "訓練需求管理"
    '        Case "CM"
    '            rst = "訓練經費控管"
    '        Case "SYS"
    '            rst = "系統管理"
    '        Case "FAQ"
    '            rst = "問答集"
    '        Case "OB"
    '            rst = "委外訓練管理"
    '        Case "SE"
    '            rst = "技能檢定管理"
    '        Case "EXAM"
    '            rst = "甄試管理"
    '        Case "SV"
    '            rst = "問卷管理"
    '        Case "OO"
    '            rst = "其他系統"
    '        Case Else
    '            rst = tmpCode
    '    End Select

    '    Return rst
    'End Function

    '判斷帳號是否有設定群組(True=>是, False=>否)
    Private Function chk_Group(ByVal strAccount As String) As Boolean

        Dim bolRtn As Boolean = False
        Dim sql As String = ""
        sql = "select gid from Auth_GroupAcct where 1=1 and account= @account" 'and rownum<=1
        Dim oCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("account", SqlDbType.VarChar).Value = strAccount
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then bolRtn = True
        Return bolRtn
    End Function

    '判斷是否有群組功能資料(True=>是, False=>否)
    Function chk_GroupData(ByVal intGID As Integer) As Boolean
        Dim bolRtn As Boolean = False
        Dim sql As String = ""
        sql = "select gid from Auth_GroupAcct where gid= @gid and account= @account"
        Dim oCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("gid", SqlDbType.Int).Value = intGID
            .Parameters.Add("account", SqlDbType.VarChar).Value = ViewState(cst_userid)
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then bolRtn = True
        Return bolRtn
    End Function

    '判斷減項群組是否有設定(True=>是, False=>否)
    Function chk_UFun(ByVal strGID As String) As Boolean
        Dim bolRtn As Boolean = True

        Dim sql As String = ""
        sql = "select gdid from auth_groupdfun where account= @account and gid= @gid"
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("account", SqlDbType.VarChar).Value = ViewState(cst_userid)
            .Parameters.Add("gid", SqlDbType.VarChar).Value = strGID
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then bolRtn = False

        Return bolRtn
    End Function

    '判斷減項群組功能是否有設定(True=>是, False=>否)
    Function chk_UFunItem(ByVal strGID As String, ByVal strFunid As String) As Boolean
        Dim bolRtn As Boolean = True

        Dim sql As String = ""
        sql = "select gdid from auth_groupdfun where account= @account and gid= @gid and funid= @funid"
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("account", SqlDbType.VarChar).Value = ViewState(cst_userid)
            .Parameters.Add("gid", SqlDbType.VarChar).Value = strGID
            .Parameters.Add("funid", SqlDbType.VarChar).Value = strFunid
            dt.Load(.ExecuteReader())
        End With

        If dt.Rows.Count = 0 Then bolRtn = False

        Return bolRtn
    End Function
#End Region

    Dim str_superuser1 As String = "snoopy" '(預設)(吃管理者權限)
    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        'TIMS.TestDbConn(Me, OBJconn)
        '檢查Session是否存在 End
        arrFun = TIMS.c_FUNSORT.Split(",") 'arrFun = FunSort.Split(",")
        Dim sql As String = ""
        sql = "SELECT TPLANID,PLANNAME FROM KEY_PLAN ORDER BY TPLANID"
        dtPlan = DbAccess.GetDataTable(sql, objconn)

        '非 ROLEID=0 LID=0
        'Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。'ROLEID=0 LID=0
            str_superuser1 = CStr(sm.UserInfo.UserID)
        End If

        If Not IsPostBack Then
            vs_years = sm.UserInfo.Years
            vs_distid = sm.UserInfo.DistID
            vs_planid = sm.UserInfo.PlanID

            ddlFun = TIMS.Get_ddlFunction(ddlFun, 2)

            'snoopy管理者。(不鎖定帳號)
            If sm.UserInfo.UserID <> str_superuser1 Then
                list_DistID.Enabled = False
                ViewState(cst_account) = sm.UserInfo.UserID
            Else
                ViewState(cst_account) = String.Empty
            End If

            Show_DropDownList("PlanYears", list_Years, "Years", "Years")
            Show_DropDownList("DistID", list_DistID, "Name", "DistID")
            Show_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")

            list_Years.SelectedValue = sm.UserInfo.Years
            list_DistID.SelectedValue = sm.UserInfo.DistID
            list_PlanID.SelectedValue = sm.UserInfo.PlanID

            Show_DropDownList("OrgID", list_OrgID, "OrgName", "OrgID")
            Renew_ListOrgID(list_OrgID, sm.UserInfo.UserID, sm.UserInfo.RoleID)

            Call sUtl_NoShowAll()
            tb_Query.Visible = True
            'tb_Group.Visible = False
            'tb_GroupFun.Visible = False
            'tb_Function.Visible = False
            'btn_Query_Click(sender, e)
            DataGrid1.CurrentPageIndex = 0
            Call sSearch1()
        End If
    End Sub


    '儲存群組。 UPDATE Auth_GroupAcct 跟ACCOUNT
    Private Sub btn_SaveGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_SaveGroup.Click
        'ViewState(cst_userid)
        Dim intCnt As Integer = 0
        Dim sql As String = ""

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            '新增群組資料
            sql = ""
            sql &= " insert into Auth_GroupAcct(GID,Account,ModifyAcct,ModifyDate)"
            sql &= " values(@GID,@Account,@ModifyAcct,getdate())"
            Dim iCmd As New SqlCommand(sql, conn, trans)

            '刪除舊群組資料
            sql = " delete Auth_GroupAcct where Account= @Account AND GID= @GID"
            Dim dCmd As New SqlCommand(sql, conn, trans)
            '刪除舊有權限資料
            sql = " delete auth_accrwfun where account= @Account"
            Dim dCmd2 As New SqlCommand(sql, conn, trans)
            '群組未存在,將之前設定減項群組刪除
            sql = ""
            sql &= " delete auth_groupdfun where account= @Account"
            sql &= " and gid not in (select gid from auth_groupacct where account= @Account)"
            Dim dCmd3 As New SqlCommand(sql, conn, trans)

            '查詢 群組資料
            sql = " SELECT 'X' FROM Auth_GroupAcct where Account= @Account AND GID= @GID"
            Dim sCmd As New SqlCommand(sql, conn, trans)


            For Each itm As DataGridItem In DataGrid2.Items
                Dim chkGroupValid As CheckBox = itm.FindControl("chk_GroupValid")

                If chkGroupValid.Checked = True Then
                    Dim dt As New DataTable
                    With sCmd
                        .Parameters.Clear()
                        .Parameters.Add("GID", SqlDbType.Int).Value = DataGrid2.DataKeys.Item(itm.ItemIndex)
                        .Parameters.Add("Account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                        'dt.Load(.ExecuteReader())
                        dt = DbAccess.GetDataTable(sCmd.CommandText, trans, sCmd.Parameters)
                    End With
                    If dt.Rows.Count = 0 Then '沒資料可新增
                        With iCmd
                            .Parameters.Clear()
                            .Parameters.Add("GID", SqlDbType.Int).Value = DataGrid2.DataKeys.Item(itm.ItemIndex)
                            .Parameters.Add("Account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                            '.ExecuteNonQuery()
                            ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                            DbAccess.ExecuteNonQuery(iCmd.CommandText, trans, iCmd.Parameters)
                        End With
                    End If
                Else
                    '未選擇一律刪除。
                    With dCmd
                        .Parameters.Clear()
                        .Parameters.Add("GID", SqlDbType.Int).Value = DataGrid2.DataKeys.Item(itm.ItemIndex)
                        .Parameters.Add("Account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                        '.ExecuteNonQuery()
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(dCmd.CommandText, trans, dCmd.Parameters)
                    End With
                End If
            Next

            '刪除舊有權限資料
            With dCmd2
                .Parameters.Clear()
                .Parameters.Add("Account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                '.ExecuteNonQuery()
                ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                DbAccess.ExecuteNonQuery(dCmd2.CommandText, trans, dCmd2.Parameters)
            End With

            '群組未存在,將之前設定減項群組刪除
            With dCmd3
                .Parameters.Clear()
                .Parameters.Add("Account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                '.ExecuteNonQuery()
                ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                DbAccess.ExecuteNonQuery(dCmd3.CommandText, trans, dCmd3.Parameters)
            End With

            DbAccess.CommitTrans(trans)
            intCnt = 1

        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(conn)
            Exit Sub
        End Try
        Call TIMS.CloseDbConn(conn)

        If intCnt = 1 Then
            'btn_CancelGroup_Click(sender, e)
            Call sSearch1()
            Clear_GroupItems()
            Common.MessageBox(Me, "儲存成功!")
        End If
    End Sub

    Private Sub btn_CancelGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CancelGroup.Click
        Call sSearch1()
        Clear_GroupItems()
    End Sub

    '儲存個人功能。 UPDATE auth_groupDFun 減項功能儲存。跟ACCOUNT
    Private Sub btn_SaveOption_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_SaveOption.Click
        Dim intCnt As Integer = 0
        Dim sql As String = ""

        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            '新增勾選之功能權限減項
            sql = ""
            sql &= " insert into auth_groupDFun(gdid,account,gid,funid,modifyacct,modifydate)"
            sql &= " values(@gdid,@account,@gid,@funid,@modifyacct,getdate()) "
            Dim iCmd As New SqlCommand(sql, conn, trans)

            '刪除勾選之功能權限減項
            sql = " delete auth_groupDFun where account= @account and gid= @gid and funid= @funid"
            Dim dCmd As New SqlCommand(sql, conn, trans)
            Dim iGDID As Int64 = 0

            '新增勾選之功能權限減項
            For Each itm As DataGridItem In Datagrid4.Items
                Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
                If chkEnable.Checked = False Then
                    '先刪
                    With dCmd
                        .Parameters.Clear()
                        .Parameters.Add("account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                        .Parameters.Add("gid", SqlDbType.Int).Value = CInt(ViewState(cst_vsgid))
                        .Parameters.Add("funID", SqlDbType.Int).Value = CInt(Datagrid4.DataKeys.Item(itm.ItemIndex))
                        '.ExecuteNonQuery()
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(dCmd.CommandText, trans, dCmd.Parameters)
                    End With

                    '改由程式產生pk值
                    iGDID = DbAccess.GetNewId(objconn, "AUTH_GROUPDFUN_GDID_SEQ,AUTH_GROUPDFUN,GDID")
                    '後加
                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("gdid", SqlDbType.Int).Value = iGDID
                        .Parameters.Add("account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                        .Parameters.Add("gid", SqlDbType.Int).Value = CInt(ViewState(cst_vsgid))
                        .Parameters.Add("funID", SqlDbType.Int).Value = CInt(Datagrid4.DataKeys.Item(itm.ItemIndex))
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                        '.ExecuteNonQuery()
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(iCmd.CommandText, trans, iCmd.Parameters)
                    End With
                Else
                    '必刪
                    With dCmd
                        .Parameters.Clear()
                        .Parameters.Add("account", SqlDbType.NVarChar).Value = ViewState(cst_userid)
                        .Parameters.Add("gid", SqlDbType.Int).Value = CInt(ViewState(cst_vsgid))
                        .Parameters.Add("funID", SqlDbType.Int).Value = CInt(Datagrid4.DataKeys.Item(itm.ItemIndex))
                        '.ExecuteNonQuery()
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(dCmd.CommandText, trans, dCmd.Parameters)
                    End With
                End If
            Next

            DbAccess.CommitTrans(trans)
            'trans.Commit()
            intCnt = 1

        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
            DbAccess.RollbackTrans(trans)
            Call TIMS.CloseDbConn(conn)
            Exit Sub
        End Try
        Call TIMS.CloseDbConn(conn)

        If intCnt = 1 Then
            'btn_CancelOption2_Click(sender, e)
            Clear_FunctionItems()
            Load_GroupFun()
            Common.MessageBox(Me, "儲存成功!")
        End If
    End Sub

    Private Sub btn_CancelOption1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CancelOption1.Click
        ViewState(cst_userid) = Nothing

        Call sUtl_NoShowAll()
        tb_Query.Visible = True
        'tb_Group.Visible = False
        'tb_GroupFun.Visible = False
        'tb_Function.Visible = False
    End Sub

    Private Sub btn_CancelOption2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CancelOption2.Click
        Clear_FunctionItems()
        Load_GroupFun()
    End Sub

    Private Sub DataGrid1_PageIndexChanged(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        DataGrid1.CurrentPageIndex = e.NewPageIndex
        Call sSearch1()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        'Dim labAccount As Label = e.Item.FindControl("lab_Account")
        Select Case UCase(e.CommandName)
            Case "GROUP"
                Dim cmdArg As String = e.CommandArgument
                ViewState(cst_userid) = TIMS.GetMyValue(cmdArg, "UserID") 'labAccount.Text
                ViewState(cst_username) = TIMS.GetMyValue(cmdArg, "UserName") 'labAccount.Text
                'ViewState(cst_userGDistID) = TIMS.GetMyValue(cmdArg, "UserGDistID")

                Call Load_Group("")
                Call Load_Group_ddlGDist2()
                ddlGroupType2 = TIMS.GET_GQTYPE(ddlGroupType2)

                Call sUtl_NoShowAll()
                'tb_Query.Visible = False
                tb_Group.Visible = True
                'tb_GroupFun.Visible = False
                'tb_Function.Visible = False

            Case "OPTION"
                Dim cmdArg As String = e.CommandArgument
                ViewState(cst_userid) = TIMS.GetMyValue(cmdArg, "UserID") 'labAccount.Text
                ViewState(cst_username) = TIMS.GetMyValue(cmdArg, "UserName") 'labAccount.Text
                'ViewState(cst_userGDistID) = TIMS.GetMyValue(cmdArg, "UserGDistID")

                'ViewState(cst_userid) = labAccount.Text
                Call Load_GroupFun()

                Call sUtl_NoShowAll()
                tb_GroupFun.Visible = True
                'tb_Query.Visible = False
                'tb_Group.Visible = False
                'tb_Function.Visible = False
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem
                Dim labAccount As Label = e.Item.FindControl("lab_Account")
                Dim labName As Label = e.Item.FindControl("lab_Name")
                Dim labGroup As Label = e.Item.FindControl("lab_Group")
                Dim hideGID As HtmlInputHidden = e.Item.FindControl("hide_GID")
                Dim btnEditGroup As LinkButton = e.Item.FindControl("btn_EditGroup")
                Dim btnEditOption As LinkButton = e.Item.FindControl("btn_EditOption")

                e.Item.Cells(0).Text = DataGrid1.PageSize * DataGrid1.CurrentPageIndex + e.Item.ItemIndex + 1

                Dim cmdArg As String = ""
                Call TIMS.SetMyValue(cmdArg, "UserID", Convert.ToString(dr_Data("Account")))
                Call TIMS.SetMyValue(cmdArg, "UserName", Convert.ToString(dr_Data("Name")))
                'Call TIMS.SetMyValue(cmdArg, "UserGDistID", Convert.ToString(dr_Data("GDistID")))

                btnEditGroup.CommandArgument = cmdArg
                btnEditOption.CommandArgument = cmdArg

                labAccount.Text = Convert.ToString(dr_Data("Account"))

                labName.Text = Convert.ToString(dr_Data("Name"))

                labGroup.Text = Server.HtmlDecode(Convert.ToString(dr_Data("GName")))
                hideGID.Value = Convert.ToString(dr_Data("GID"))

                '當登入者角色權限小於找出的人時，不能進行設定
                If Convert.ToString(dr_Data("RoleID")) < sm.UserInfo.RoleID Then
                    btnEditGroup.Enabled = False
                    btnEditOption.Enabled = False
                End If

                If chk_Group(Convert.ToString(dr_Data("Account"))) = False Then
                    btnEditOption.ToolTip = "群組尚未設定"
                    btnEditOption.Enabled = False
                End If
        End Select
    End Sub

    Sub List_GroupTPlan(ByVal GID As String, ByVal sYears As String, ByVal sDistID As String, ByVal sUserID As String)
        If GID = "" Then Exit Sub

        Lab5UserID.Text = ViewState(cst_userid)
        Lab5Name.Text = ViewState(cst_username)

        HidGID.Value = GID
        LabGroupName.Text = TIMS.GET_GroupName(GID, objconn)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT DISTINCT a.TPlanID" & vbCrLf
        sql &= " ,b.PlanName" & vbCrLf
        sql &= " FROM ID_PLAN a " & vbCrLf
        sql &= " JOIN KEY_PLAN b on b.TPlanID=a.TPlanID" & vbCrLf
        sql &= " JOIN ID_DISTRICT c on c.DistID=a.DistID" & vbCrLf
        sql &= " JOIN AUTH_ACCRWPLAN d on d.PlanID=a.PlanID" & vbCrLf
        sql &= " where 1=1 " & vbCrLf
        sql &= " AND a.Years='" & sYears & "'" & vbCrLf
        sql &= " AND a.DISTID='" & sDistID & "'" & vbCrLf
        'sUserID:被設定者帳號 / sAccount:登入者帳號。
        If sUserID <> "" Then sql &= " AND D.ACCOUNT='" & sUserID & "'  " & vbCrLf
        sql &= " ORDER BY A.TPlanID ASC " & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        lab_Msg5.Visible = True
        Datagrid5.Visible = False
        If dt.Rows.Count > 0 Then
            lab_Msg5.Visible = False
            Datagrid5.Visible = True

            Datagrid5.DataSource = dt
            Datagrid5.DataBind()
        End If
    End Sub

    Private Sub DataGrid2_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid2.ItemCommand
        Select Case e.CommandName
            Case "PlanEdit"
                Dim cmdArg As String = e.CommandArgument
                Dim sGID As String = TIMS.GetMyValue(cmdArg, "GID")
                Dim sYears As String = TIMS.GetMyValue(cmdArg, "Years") '登入者 (年度)
                Dim sDistID As String = TIMS.GetMyValue(cmdArg, "DistID") '登入者
                Dim sAccount As String = TIMS.GetMyValue(cmdArg, "Account") '帳號(登入者)
                Dim sUserID As String = TIMS.GetMyValue(cmdArg, "UserID") '所選擇要設定的帳號。

                hid_GidTPlanIDs.Value = TIMS.GetMyValue(cmdArg, "GTPLANID")
                If Hid_PLANID.Value <> "" Then
                    Dim strSql As String = "SELECT DISTID,YEARS FROM ID_PLAN WHERE PLANID=@PLANID"
                    Dim parms As New Hashtable
                    parms.Add("PLANID", Hid_PLANID.Value)
                    Dim dtPlan As DataTable = TIMS.Get_KeyTable2(strSql, objconn, parms)
                    If dtPlan.Rows.Count > 0 Then
                        Dim dr1 As DataRow = dtPlan.Rows(0)
                        sDistID = Convert.ToString(dr1("DISTID")) '要設定的帳號-DISTID
                        sYears = Convert.ToString(dr1("YEARS")) '要設定的帳號-YEARS
                    End If
                End If
                Call List_GroupTPlan(sGID, sYears, sDistID, sUserID)

                Call sUtl_NoShowAll()
                tb_GroupTPlan.Visible = True
        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem
                Dim chkGroupValid As CheckBox = e.Item.FindControl("chk_GroupValid")
                Dim labDistName As Label = e.Item.FindControl("lab_DistName")
                Dim labGroupType As Label = e.Item.FindControl("lab_GroupType")
                Dim labGroupName As Label = e.Item.FindControl("lab_GroupName")
                Dim labGroupNote As Label = e.Item.FindControl("lab_GroupNote")
                Dim lplanName As Label = e.Item.FindControl("lplanName")
                Dim btnPlanEdit As LinkButton = e.Item.FindControl("btnPlanEdit")

                labDistName.Text = "(系統預設)"
                If Convert.ToString(dr_Data("GDistID")) <> "" Then labDistName.Text = TIMS.Get_DistName1(dr_Data("GDistID"))

                Dim s_GroupTypeN As String = "委訓"
                Select Case Convert.ToString(dr_Data("GType"))
                    Case "0"
                        s_GroupTypeN = "署"'labGroupType.Text = "局"
                    Case "1"
                        s_GroupTypeN = "分署" 'labGroupType.Text = "中心"
                        'Case "2" ' labGroupType.Text = "委訓"
                End Select
                labGroupType.Text = s_GroupTypeN

                labGroupName.Text += Convert.ToString(dr_Data("GName"))
                labGroupNote.Text = Convert.ToString(dr_Data("GNote"))

                '判斷是否有群組資料
                chkGroupValid.Checked = False
                If chk_GroupData(Convert.ToInt32(dr_Data("GID"))) Then chkGroupValid.Checked = True

                '找出群組計畫(中文)名稱。配合帳號的權限

                'lplanName.Text = TIMS.Get_GID_PlanName(Convert.ToString(dr_Data("GTPLANID")), Convert.ToString(dr_Data("Account")), dtPlan, objconn)
                If Convert.ToString(dr_Data("GTPLANID")) <> "" Then
                    lplanName.Text = TIMS.Get_GID_PlanName(Convert.ToString(dr_Data("GTPLANID")), dtPlan, objconn)
                End If

                Dim CmdArg As String = ""
                CmdArg = ""
                If Convert.ToString(dr_Data("GTPLANID")) <> "" Then
                    Call TIMS.SetMyValue(CmdArg, "GTPLANID", dr_Data("GTPLANID"))
                End If
                Call TIMS.SetMyValue(CmdArg, "GID", dr_Data("GID"))
                Call TIMS.SetMyValue(CmdArg, "Years", vs_years) '依年度(登入者)
                Call TIMS.SetMyValue(CmdArg, "DistID", vs_distid) '依轄區(登入者)
                'Call TIMS.SetMyValue(CmdArg, "DistID", dr_Data("GDistID")) '依轄區(要設定者)
                Call TIMS.SetMyValue(CmdArg, "Account", ViewState(cst_account)) '帳號(登入者)
                Call TIMS.SetMyValue(CmdArg, "UserID", ViewState(cst_userid)) '所選擇要設定的帳號。
                btnPlanEdit.CommandArgument = CmdArg

                '尚未設定計畫群組。
                If Not chkGroupValid.Checked Then
                    btnPlanEdit.CommandArgument = ""
                    btnPlanEdit.Attributes.Add("OnClick", "alert('無群組設定，請先儲存群組設定!!');return false;")
                End If
        End Select
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem
                Dim labUntName As Label = e.Item.FindControl("labUntName")
                Dim labGType As Label = e.Item.FindControl("labGType")
                Dim labGName As Label = e.Item.FindControl("labGName")
                Dim labMemo As Label = e.Item.FindControl("labMemo")
                Dim btnEdit As LinkButton = e.Item.FindControl("btnEdit")

                labUntName.Text = "(系統預設)"
                If Convert.ToString(dr_Data("GDistID")) <> "" Then
                    labUntName.Text = TIMS.Get_DistName1(dr_Data("GDistID"))
                End If

                Select Case Convert.ToString(dr_Data("GType"))
                    Case "0"
                        'labGType.Text = "局"
                        labGType.Text = "署"
                    Case "1"
                        'labGType.Text = "中心"
                        labGType.Text = "分署"
                    Case "2"
                        labGType.Text = "委訓"
                End Select

                labGName.Text = Convert.ToString(dr_Data("gname"))

                If chk_UFun(Convert.ToString(dr_Data("gid"))) Then
                    labMemo.Text = "功能減項已設定"
                Else
                    labMemo.Text = ""
                End If

                btnEdit.CommandArgument = Convert.ToString(dr_Data("gid"))
        End Select
    End Sub

    Private Sub DataGrid3_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid3.ItemCommand
        Select Case e.CommandName
            Case "Edit1"
                ViewState(cst_vsgid) = e.CommandArgument
                Load_Function(ViewState(cst_vsgid))

                sUtl_NoShowAll()
                'tb_Query.Visible = False
                'tb_Group.Visible = False
                'tb_GroupFun.Visible = False
                tb_Function.Visible = True
        End Select
    End Sub

    Private Sub DataGrid4_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                chkEnableAll = e.Item.FindControl("chk_EnableAll") '全選方塊

            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim txtFunID As TextBox = e.Item.FindControl("txtFunID") '記錄FunID
                Dim labMainMenu As Label = e.Item.FindControl("lab_MainMenu") '程式類型
                Dim labFunName As Label = e.Item.FindControl("lab_FunName") '選單名稱
                Dim chkEnable As CheckBox = e.Item.FindControl("chk_Enable") '選取方塊

                txtFunID.Text = Convert.ToString(drv("funid"))

                'labMainMenu.Text = Get_MainMenuName(Convert.ToString(drv("Kind")))
                Dim strkind As String = Convert.ToString(drv("Kind"))
                labMainMenu.Text = TIMS.Get_MainMenuName(strkind)

                If Convert.ToString(drv("levels")) = "1" Then
                    labFunName.Text = "&nbsp;&nbsp;&nbsp;" & Convert.ToString(drv("Name"))
                Else
                    labFunName.Text = Convert.ToString(drv("Name"))
                End If

                'snoopy特別權限
                If Convert.ToString(sm.UserInfo.UserID) = str_superuser1 Then
                    If Convert.ToString(drv("Spage")) <> "" Then
                        ''回上2層並給ID
                        Dim strUrl As String = "<A href=""../../" & Convert.ToString(drv("Spage")) & "?ID=" & Convert.ToString(drv("funid")) & """ target=""_blank"">" & "(" & Convert.ToString(drv("Spage")) & ")" & "</A>"
                        labFunName.Text += strUrl
                    End If
                End If

                If chk_UFunItem(Convert.ToString(drv("gid")), Convert.ToString(drv("funid"))) Then
                    chkEnable.Checked = False
                Else
                    chkEnable.Checked = True
                End If

                If Convert.ToString(vs_mmname) <> labMainMenu.Text Then '同類型選單的第一項
                    Dim subs As Integer = ViewState(Convert.ToString(drv("Kind"))) '依類型取得選單數量

                    vs_mmname = labMainMenu.Text
                    e.Item.Cells(0).RowSpan = subs '合併同類型選單
                    e.Item.Cells(0).BackColor = Color.FromArgb(241, 249, 252)
                    e.Item.Attributes.Add("id", Convert.ToString(drv("Kind")))

                    For i As Integer = 0 To e.Item.Cells.Count - 1
                        e.Item.Cells(i).Attributes.Add("id", Convert.ToString(drv("Kind")) & "td" & i)
                    Next

                    trID = 1

                Else '非同類型選單第一項的其他項
                    e.Item.Cells(0).Visible = False '隱藏功能類別欄位
                    e.Item.Attributes.Add("id", Convert.ToString(drv("Kind")) & trID)

                    trID += 1
                End If

                '設定主選單顏色
                If Convert.ToString(drv("Subs")) <> "0" Then
                    e.Item.BackColor = Color.FromArgb(235, 243, 254)
                End If

                '紀錄第一項的ClientID
                If e.Item.ItemIndex = 0 Then
                    vs_itemname = chkEnable.ClientID
                End If

            Case ListItemType.Footer
                '設定全選動作的JavaScript
                chkEnableAll.Attributes.Add("onclick", "Show_SelectAll('" & chkEnableAll.ClientID & "','" & Convert.ToString(vs_itemname) & "'," & Datagrid4.Items.Count & ")")
        End Select
    End Sub

    Sub Show_list_YearsDistID()
        vs_years = list_Years.SelectedValue
        vs_distid = list_DistID.SelectedValue

        '依年度 轄區 顯示可用計畫
        Show_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")

        '可用單位清除
        list_OrgID.Items.Clear()

        DataGrid1.CurrentPageIndex = 0
        Call sSearch1()
        '依計畫 顯示可用單位
        'Call Show_list_PlanID()
    End Sub

    Sub Show_list_PlanID()
        vs_planid = list_PlanID.SelectedValue

        '依計畫 顯示可用單位
        Show_DropDownList("OrgID", list_OrgID, "OrgName", "OrgID")
        Renew_ListOrgID(list_OrgID, sm.UserInfo.UserID, sm.UserInfo.RoleID)
    End Sub

    Private Sub ddlFun_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlFun.SelectedIndexChanged
        Load_Function(ViewState(cst_vsgid))
    End Sub

    Private Sub list_Years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_Years.SelectedIndexChanged
        '依年度 轄區 顯示可用計畫
        Call Show_list_YearsDistID()
    End Sub

    Private Sub list_DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_DistID.SelectedIndexChanged
        '依年度 轄區 顯示可用計畫
        Call Show_list_YearsDistID()
    End Sub

    Private Sub list_PlanID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_PlanID.SelectedIndexChanged
        '依計畫 顯示可用單位
        Call Show_list_PlanID()

        DataGrid1.CurrentPageIndex = 0
        Call sSearch1()
    End Sub

    Private Sub list_OrgID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_OrgID.SelectedIndexChanged
        DataGrid1.CurrentPageIndex = 0
        Call sSearch1()
    End Sub

    Protected Sub ddlGDist2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGDist2.SelectedIndexChanged
        Call Load_Group("Y")
        'tb_Query.Visible = False'tb_Group.Visible = True'tb_GroupFun.Visible = False'tb_Function.Visible = False    
    End Sub

    '取消所有顯示表格。
    Sub sUtl_NoShowAll()
        tb_Query.Visible = False
        tb_Group.Visible = False
        tb_GroupFun.Visible = False
        tb_Function.Visible = False
        tb_GroupTPlan.Visible = False
    End Sub

    '儲存
    Protected Sub btnSave5_Click(sender As Object, e As EventArgs) Handles btnSave5.Click
        If HidGID.Value = "" Then Exit Sub

        Dim gTPlanID As String = ""
        gTPlanID = ""
        'Dim NGgTPlanID As String = ""
        'NGgTPlanID = ""
        For Each itm As DataGridItem In Datagrid5.Items
            Dim ChkBox1 As CheckBox = itm.FindControl("ChkBox1")
            Dim hid_TPlanID As HtmlInputHidden = itm.FindControl("hid_TPlanID")
            If ChkBox1.Checked Then
                If gTPlanID <> "" Then gTPlanID &= ","
                gTPlanID &= hid_TPlanID.Value
            End If
            'If Not ChkBox1.Checked Then
            '    If NGgTPlanID <> "" Then NGgTPlanID &= ","
            '    NGgTPlanID &= hid_TPlanID.Value
            'End If
        Next

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE AUTH_GROUPACCT" & vbCrLf
        sql &= " SET GTPLANID =@GTPLANID" & vbCrLf
        sql &= " ,modifyacct =@modifyacct" & vbCrLf
        sql &= " ,modifydate =getdate()" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND GID =@GID" & vbCrLf
        sql &= " AND ACCOUNT =@ACCOUNT" & vbCrLf
        Dim oCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("GTPLANID", SqlDbType.VarChar).Value = gTPlanID '被設定者計畫。
            .Parameters.Add("modifyacct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

            .Parameters.Add("GID", SqlDbType.VarChar).Value = HidGID.Value '被設定者群組。
            .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = Lab5UserID.Text '被設定者帳號
            '.ExecuteNonQuery()
            ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
            DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)
        End With

        Call Load_Group("")
        Call sUtl_NoShowAll()
        tb_Group.Visible = True
    End Sub

    '取消
    Protected Sub btnCancel5_Click(sender As Object, e As EventArgs) Handles btnCancel5.Click
        Call sUtl_NoShowAll()
        tb_Group.Visible = True
    End Sub

    Private Sub Datagrid5_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid5.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim ChkBox1 As CheckBox = e.Item.FindControl("ChkBox1")
                Dim hid_TPlanID As HtmlInputHidden = e.Item.FindControl("hid_TPlanID")
                Dim lab_PlanName As Label = e.Item.FindControl("lab_PlanName")

                ChkBox1.Checked = False
                If hid_GidTPlanIDs.Value.IndexOf(drv("TPlanID")) > -1 Then
                    ChkBox1.Checked = True
                End If

                hid_TPlanID.Value = drv("TPlanID")
                lab_PlanName.Text = drv("PlanName")
        End Select
    End Sub

    ''' <summary> 查詢 </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Protected Sub btn_Query_Click(sender As Object, e As EventArgs) Handles btn_Query.Click
        DataGrid1.CurrentPageIndex = 0
        Call sSearch1()
    End Sub

    Protected Sub btn_Print_Click(sender As Object, e As EventArgs) Handles btn_Print.Click
        Const Cst_功能欄位 As Integer = 4
        Const Cst_FileName As String = "帳號群組資料"

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(Cst_功能欄位).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call sSearch1()

        'Dim sFileName As String = ""
        'sFileName = HttpUtility.UrlEncode(Cst_FileName & ".xls", System.Text.Encoding.UTF8)
        'sFileName = HttpUtility.UrlEncode(Cst_FileName & ".xls", System.Text.Encoding.ASCII)
        Dim sFileName As String = TIMS.ClearSQM(Cst_FileName & ".xls")

        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        'Response.Charset = "Big5" '設定字集
        'Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        Response.AddHeader("Content-Disposition", "attachment; filename=" & System.Web.HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8))
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False
        DataGrid1.Columns(Cst_功能欄位).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        Response.End()

        DataGrid1.AllowPaging = True
        DataGrid1.Columns(Cst_功能欄位).Visible = True
        'Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub DataGrid2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid2.SelectedIndexChanged

    End Sub

    Protected Sub ddlGroupType2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGroupType2.SelectedIndexChanged
        Call Load_Group("Y")
    End Sub
End Class

