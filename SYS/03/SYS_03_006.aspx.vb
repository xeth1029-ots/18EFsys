Partial Class SYS_03_006
    Inherits AuthBasePage

    Const cst_已結訓 As String = "C"
    Const cst_未結訓 As String = "O"
    Const cst_全部 As String = "A"

    Const cst_funid_262_e網報名審核 As String = "262" 'SD/01/SD_01_004.aspx
    Const cst_funid_70_報名登錄 As String = "70"
    Const cst_funid_76_甄試成績登錄 As String = "76"
    Const cst_funid_79_錄訓作業 As String = "79"
    Const cst_funid_82_學員參訓 As String = "82" 'SD/03/SD_03_001.aspx
    Const cst_funid_83_學員資料維護 As String = "83" 'SD/03/SD_03_002.aspx

    '班級結訓作業: 245
    '結訓成績登錄: 118
    '結訓學員資料卡登錄: 154
    '學員就業狀況作業: 208

    'e網報名審核: 262
    '報名登錄: 70
    '甄試成績登錄: 76
    '學員參訓: 82
    '學員資料維護: 83
    '錄訓作業: 79

    '異動Table : AUTH_RENDCLASS
    'SELECT UseAble ,COUNT(1) CNT FROM Auth_REndClass GROUP BY UseAble ORDER BY 1
    'MakeListItem: 從Sql查詢結果集合的第1欄位(Value)、第2欄位(Text)

    ''年度
    'Function GetTPlanIDYears(ByVal PlanID As String, ByRef TPlanID As String, ByRef Years As String) As DataRow
    '    Dim sql As String = "SELECT * FROM id_Plan Where PlanID ='" & PlanID & "'"
    '    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then
    '        TPlanID = dr("TPlanID")
    '        Years = dr("Years")
    '    End If
    '    Return dr
    'End Function

    '計畫
    Sub Makeplanlist(ByRef ddlobj As DropDownList, ByVal Years As String, ByVal DistID As String)
        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " select distinct a.PlanID, a.Years+b.Name+c.PlanName+a.seq PlanName, a.DistID " & vbCrLf
        Sql += " from ID_Plan a " & vbCrLf
        Sql += " JOIN ID_District b on a.DistID=b.DistID" & vbCrLf
        Sql += " JOIN Key_Plan c on a.TPlanID=c.TPlanID" & vbCrLf
        Sql += " JOIN Auth_AccRWPlan d on a.PlanID=d.PlanID" & vbCrLf
        Sql += " where 1=1" & vbCrLf
        Sql += " and a.years = '" & Years & "' " & vbCrLf
        Sql += " and a.DistID = '" & DistID & "'" & vbCrLf
        Sql += " order by 2 " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, Sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    '機構
    Sub MakeddlOrgName(ByRef ddlobj As DropDownList, ByVal PlanID As String, ByVal LID As String, ByVal RoleID As String, ByVal DistID As String)
        'Dim TPlanID As String = ""
        'Dim Years As String = ""
        'Dim dr As DataRow = GetTPlanIDYears(PlanID, TPlanID, Years)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT distinct oo.OrgID, oo.OrgName, c.OrgLevel, c.RID " & vbCrLf
        sql += " From Auth_Account a " & vbCrLf
        sql += " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        sql += " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        sql += " JOIN Org_Orginfo oo on oo.OrgID =c.OrgID" & vbCrLf
        sql += " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID " & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " and a.IsUsed='Y' " & vbCrLf
        sql += " and a.LID>='" & LID & "' " & vbCrLf
        sql += " and a.RoleID>='" & RoleID & "' " & vbCrLf
        sql += " and b.PlanID = '" & PlanID & "' " & vbCrLf
        sql += " and c.DistID = '" & DistID & "' " & vbCrLf
        'Sql += " and ip.TPlanID = '" & TPlanID & "' " & vbCrLf
        'Sql += " and ip.Years = '" & Years & "' " & vbCrLf
        sql += " order by oo.OrgName,c.OrgLevel, c.RID " & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, sql, objconn)
        'ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        ddlobj.Items.Insert(0, New ListItem("全部", ""))
    End Sub

    '帳號
    Sub MakeAccount(ByRef ddlobj As DropDownList, ByVal PlanID As String, ByVal LID As String, ByVal RoleID As String, ByVal DistID As String, ByVal OrgID As String)
        'Dim TPlanID As String = ""
        'Dim Years As String = ""
        'Dim dr As DataRow = GetTPlanIDYears(PlanID, TPlanID, Years)

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " SELECT distinct a.Account ,a.Name+'('+d.Name+')' sName" & vbCrLf
        Sql += " ,a.RoleID,a.LID" & vbCrLf
        Sql += " From Auth_Account a" & vbCrLf
        Sql += " JOIN Auth_AccRWPlan b ON a.Account=b.Account" & vbCrLf
        Sql += " JOIN Auth_Relship c ON b.RID=c.RID" & vbCrLf
        Sql += " LEFT JOIN id_plan ip ON ip.PlanID =c.PlanID" & vbCrLf
        Sql += " LEFT JOIN ID_Role d ON a.RoleID=d.RoleID" & vbCrLf
        Sql += " Where 1=1" & vbCrLf
        Sql += " and a.IsUsed='Y' " & vbCrLf
        Sql += " and a.LID>='" & LID & "' " & vbCrLf
        Sql += " and a.RoleID>='" & RoleID & "' " & vbCrLf
        Sql += " and b.PlanID = '" & PlanID & "' " & vbCrLf
        Sql += " and c.DistID = '" & DistID & "' " & vbCrLf
        'Sql += " and ip.TPlanID = '" & TPlanID & "' " & vbCrLf
        'Sql += " and ip.Years = '" & Years & "' " & vbCrLf
        If OrgID <> "" Then
            Sql += " and c.OrgID = '" & OrgID & "' " & vbCrLf
        End If
        Sql += " order by a.RoleID, sName" & vbCrLf

        ddlobj.Items.Clear()
        DbAccess.MakeListItem(ddlobj, Sql, objconn)
        ddlobj.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    End Sub

    ''取得目前 已結訓班級使用授權檔 流水號 最大值
    'Function Auth_REndClass_MaxNo() As Integer
    '    Dim MaxNo As Integer = 1
    '    Dim sql As String = "select max(RightID) max from Auth_REndClass "
    '    Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
    '    If Not dr Is Nothing Then
    '        If Not IsDBNull(dr("max")) Then
    '            MaxNo = CInt(dr("max")) + 1
    '        End If
    '    End If
    '    Return MaxNo
    'End Function

    '檢查帳號是否已賦于權限
    Private Function chkNoRecordInAuth_Rend(ByVal Account As String, ByVal OCID As String) As Boolean
        Dim NoRecord As Boolean = True

        For i As Integer = 0 To cb_SelFunID.Items.Count - 1
            If cb_SelFunID.Items(i).Selected Then
                Dim dt As DataTable
                Dim sql As String = ""
                Dim selstr As String = ""
                selstr = cb_SelFunID.Items(i).Value.ToString()

                sql = ""
                sql &= " select RightID,Years,OCID,account"
                sql &= " ,UseAble,EndDate,FunID"
                sql &= " from Auth_RendClass "
                sql += " where 1=1"
                sql += " and UseAble='Y' "
                sql += " and OCID='" & OCID & "'" '20090302  改每個班級只限一授于一個帳號
                dt = DbAccess.GetDataTable(sql, objconn)
                If dt.Rows.Count > 0 Then
                    NoRecord = False
                    Exit For
                End If
            End If
        Next
        Return NoRecord
    End Function

    Function Chkdata(ByRef MsgStr As String) As Boolean
        Dim rst As Boolean = True
        Dim ErrCount As Int16 = 0

        MsgStr = ""
        If Me.Account.SelectedValue = "" Then
            MsgStr += "請選擇【帳號】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.ReasonID.SelectedValue = "" Then
            MsgStr += "請選擇【補登資料原因】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.Reason.Text = "" Then
            MsgStr += "請填寫【補登資料原因簡述】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If TIMS.GetSelFunID(cb_SelFunID) = "" Then
            MsgStr += "請選擇【開放功能】！" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If EndDate.Text = "" Then
            MsgStr += "請選擇【結束日期】！" & vbCrLf
            ErrCount = ErrCount + 1
        End If

        If ErrCount > 0 Then
            rst = False
        End If

        Return rst
    End Function

    Function chk_search(ByRef MsgStr As String) As Boolean
        Dim rst As Boolean = True
        Dim ErrCount As Integer = 0

        MsgStr = ""
        If Me.yearlist.SelectedValue = "" Then
            MsgStr += "請選擇【年度】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.DistID.SelectedValue = "" Then
            MsgStr += "請選擇【轄區】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If
        If Me.planlist.SelectedValue = "" Then
            MsgStr += "請選擇【訓練計畫】!" & vbCrLf
            ErrCount = ErrCount + 1
        End If

        If ErrCount > 0 Then
            rst = False
        End If

        Return rst
    End Function

    '異動Table : Auth_REndClass
    'Dim FunDr As DataRow
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DG_ClassInfo

        If Not IsPostBack Then
            msg.Text = ""
            ReasonID = TIMS.Get_ReasonID(ReasonID, objconn)
            yearlist = TIMS.GetSyear(yearlist, 0, 0, False)
            DistID = TIMS.Get_DistID(DistID)
            '可使用的補登功能 存在 ID_Function (select * from ID_Function WHERE ReUse='Y')
            cb_SelFunID = TIMS.Get_FunIDReUse(cb_SelFunID, objconn, "")
            Common.SetListItem(yearlist, Now.Year)

            Reason_tr.Visible = False
            Account_tr.Visible = False
            Me.trOrgName.Visible = False
            '-----------------20090226 andy add
            Fun_tr.Visible = False
            '----------------
            PageControler1.Visible = False
        End If

        'rt_search.Attributes("onclick") = "javascript:return search()"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            '非署(局)之人員 若被賦予權限只有瀏覽權
        '            check_add.Value = 0
        '            check_mod.Value = 0
        '            check_del.Value = 0
        '            check_Sech.Value = 1
        '            rt_search.Enabled = True

        '            If sm.UserInfo.RID = "A" Then
        '                'check_Sech.Value = 0
        '                'rt_search.Enabled = False
        '                'If FunDr("Adds") = 1 OrElse FunDr("Mod") = 1 OrElse FunDr("Del") = 1 OrElse FunDr("Sech") = 1 Then
        '                '    check_Sech.Value = 1
        '                '    rt_search.Enabled = True
        '                'End If

        '                If FunDr("Adds") = 1 Then  '新增資料權
        '                    check_add.Value = 1
        '                End If

        '                If FunDr("Mod") = 1 Then  '修改資料權
        '                    check_mod.Value = 1
        '                End If

        '                If FunDr("Del") = 1 Then  '刪除資料權
        '                    check_del.Value = 1
        '                End If
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    Private Sub yearlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles yearlist.SelectedIndexChanged
        '若沒有選擇轄區帶入使用者登入轄區
        If Me.DistID.SelectedValue = "" Then
            Common.SetListItem(Me.DistID, sm.UserInfo.DistID)
        End If
        '計畫
        Call Makeplanlist(planlist, Me.yearlist.SelectedValue, Me.DistID.SelectedValue)

        tbSearch1.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"
    End Sub

    Private Sub DistID_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DistID.SelectedIndexChanged
        '計畫
        Call Makeplanlist(planlist, Me.yearlist.SelectedValue, Me.DistID.SelectedValue)
        'MakeddlOrgName(Me.ddlOrgName, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
        'MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
        tbSearch1.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"
    End Sub

    'SQL
    Function ShowDG_ClassInfo(ByVal dt As DataTable) As DataTable
        'Dim IsTPlan28 As Boolean = False
        'Dim vsTPlanID As String = ""
        'IsTPlan28 = False
        'vsTPlanID = TIMS.GetTPlanID(Me.planlist.SelectedValue)
        ''產業人才投資方案
        'If TIMS.Cst_TPlanID28AppPlan.IndexOf(vsTPlanID) > -1 Then
        '    IsTPlan28 = True
        'End If

        Dim sqlstr As String = ""
        sqlstr = "" & vbCrLf
        sqlstr &= " Select a.Years" & vbCrLf
        sqlstr &= " ,a.CyclType" & vbCrLf
        sqlstr &= " ,a.ClassNum" & vbCrLf
        sqlstr &= " ,b.ClassID" & vbCrLf
        sqlstr &= " ,a.PlanID" & vbCrLf
        sqlstr &= " ,a.OCID" & vbCrLf
        sqlstr &= " ,e.OrgName" & vbCrLf
        sqlstr &= " ,a.ClassCName+'(第'+ ISNULL(a.CyclType,'')+'期)' ClassCName" & vbCrLf
        sqlstr &= " ,g.TrainName" & vbCrLf
        sqlstr &= " ,a.STDate" & vbCrLf
        sqlstr &= " ,a.FTDate" & vbCrLf
        sqlstr &= " ,a.RID" & vbCrLf
        sqlstr &= " ,dbo.NVL(CONVERT(varchar, h.RightID),'XX') RightID" & vbCrLf
        sqlstr &= " ,dbo.NVL(h.NAME,' ') NAME" & vbCrLf
        sqlstr &= " ,dbo.NVL(h.ACCOUNT,' ') ACCOUNT" & vbCrLf
        sqlstr &= " ,(CASE WHEN dbo.NVL(h.ACCOUNT,'0') = '0' THEN '0' ELSE '1' END) Acnt" & vbCrLf
        sqlstr &= " ,h.EndDate" & vbCrLf
        sqlstr &= " ,h.Temp1" & vbCrLf
        sqlstr &= " , a.Years + '0' + ISNULL(b.ClassID2,b.ClassID)+ a.CyclType ClassID2" & vbCrLf
        sqlstr &= " From Class_ClassInfo a" & vbCrLf
        sqlstr &= " join id_plan ip on ip.PlanID =a.PlanID" & vbCrLf
        sqlstr &= " join ID_Class b on a.CLSID = b.CLSID" & vbCrLf
        sqlstr &= " JOIN ID_District c ON b.DistID = c.DistID" & vbCrLf
        sqlstr &= " JOIN Auth_Relship d on a.RID  = d.RID" & vbCrLf
        sqlstr &= " JOIN Org_OrgInfo e on d.OrgID = e.OrgID" & vbCrLf
        sqlstr &= " LEFT JOIN Key_TrainType g   on a.TMID  = g.TMID" & vbCrLf
        sqlstr &= " LEFT JOIN (" & vbCrLf
        sqlstr &= "     SELECT h1.RightID,h1.OCID,h2.ACCOUNT,h2.NAME,h1.EndDate " & vbCrLf
        sqlstr &= "     ,d1.Name+';開放FunID: '+h1.FunID  Temp1" & vbCrLf
        sqlstr &= " 	FROM Auth_REndClass h1" & vbCrLf
        sqlstr &= " 	join Auth_Account h2 ON h1.ACCOUNT = h2.ACCOUNT" & vbCrLf
        sqlstr &= " 	LEFT JOIN ID_KEYINREASON d1 on h1.ReasonID=d1.ReasonID" & vbCrLf
        sqlstr &= " 	where h1.UseAble = 'Y'" & vbCrLf
        sqlstr &= " ) h ON a.OCID = h.OCID" & vbCrLf
        sqlstr &= " Where 1=1 " & vbCrLf
        '20090617 andy  edit
        '--------------------
        sqlstr &= " and a.IsSuccess='Y'" & vbCrLf '是否轉入成功
        sqlstr &= " and a.NotOpen='N' " & vbCrLf  '不開班
        sqlstr &= " and ip.Years = @Years " & vbCrLf
        sqlstr &= " and ip.DistID = @DistID " & vbCrLf

        Dim parms As Hashtable = New Hashtable()
        parms.Add("Years", Me.yearlist.SelectedValue)
        parms.Add("DistID", Me.DistID.SelectedValue)

        If Me.planlist.SelectedValue <> "" Then
            sqlstr &= " and ip.PlanID = @PlanID " & vbCrLf
            parms.Add("PlanID", Me.planlist.SelectedValue)
        End If

        'If Me.ddlOrgName.SelectedValue <> "" Then
        '    sqlstr &= " and e.OrgID = '" & Me.ddlOrgName.SelectedValue & "' " & vbCrLf
        'End If
        ''產投計畫  e網報名審核匯入功能 -->改為不管是否結訓都帶出來
        'If IsTPlan28 = False Then
        '    sqlstr &= " and a.FTDate <= getdate() " & vbCrLf
        'Else
        '    sqlstr &= " and a.STDate<=getdate() " & vbCrLf '在訓
        'End If

        Select Case ClassRound.SelectedValue
            Case cst_已結訓
                sqlstr &= " and dbo.TRUNC_DATETIME(a.FTDate) <= dbo.TRUNC_DATETIME(getdate())" & vbCrLf
            Case cst_未結訓
                sqlstr &= " and dbo.TRUNC_DATETIME(a.FTDate) > dbo.TRUNC_DATETIME(getdate())" & vbCrLf
        End Select
        '-------------------
        If Me.schOrgName.Text <> "" Then
            sqlstr &= " and e.OrgName like @OrgName " & vbCrLf
            parms.Add("OrgName", "%" & Me.schOrgName.Text & "%")
        End If
        If Me.ClassName.Text <> "" Then
            sqlstr &= " and a.ClassCName like @ClassCName " & vbCrLf
            parms.Add("ClassCName", "%" & Me.ClassName.Text & "%")
        End If
        If Me.start_date.Text <> "" Then
            sqlstr += "and a.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", Me.start_date.Text)
        End If
        If Me.end_date.Text <> "" Then
            sqlstr &= " and a.STDate <= @STDate2" & vbCrLf
            parms.Add("STDate2", Me.end_date.Text)
        End If
        If Me.CyclType.Text <> "" Then
            If IsNumeric(Me.CyclType.Text) Then
                'If Int(Me.CyclType.Text) < 10 Then
                '    sqlstr &= " and a.CyclType = '0" & Int(Me.CyclType.Text) & "'" & vbCrLf
                'Else
                '    sqlstr &= " and a.CyclType = '" & Me.CyclType.Text & "'" & vbCrLf
                'End If
                sqlstr &= " and a.CyclType = @CyclType " & vbCrLf
                parms.Add("CyclType", IIf(Int(Me.CyclType.Text) < 10, "0", "") & Me.CyclType.Text)
            End If
        End If
        'Try
        '    dt = DbAccess.GetDataTable(Sqlstr)
        'Catch ex As Exception
        '    Common.RespWrite(Me, Sqlstr)
        'End Try
        dt = DbAccess.GetDataTable(sqlstr, objconn, parms)
        Return dt
    End Function

    '重新查詢。
    Sub dt_search()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        tbSearch1.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        'msg.Text = "查無資料!!"

        Dim MsgStr As String = ""
        MsgStr = ""
        If Not chk_search(MsgStr) Then
            Common.MessageBox(Me, MsgStr)
            Exit Sub
        End If

        Dim vsOrgName As String = ""
        Dim vsAccount As String = ""
        vsOrgName = ""
        vsAccount = ""
        If Me.ddlOrgName.SelectedValue <> "" Then
            vsOrgName = Me.ddlOrgName.SelectedValue
        End If
        If Me.Account.SelectedValue <> "" Then
            vsAccount = Me.Account.SelectedValue
        End If
        MakeddlOrgName(Me.ddlOrgName, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue)
        '帳號
        Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
        If vsOrgName <> "" Then
            Common.SetListItem(ddlOrgName, vsOrgName)
        End If
        If vsAccount <> "" Then
            Common.SetListItem(Account, vsAccount)
        End If

        Dim dt As DataTable = Nothing
        dt = ShowDG_ClassInfo(dt)

        tbSearch1.Visible = False
        DG_ClassInfo.Visible = False 'TPanel@DG_ClassInfo
        PageControler1.Visible = False 'TPanel@PageControler1
        Reason_tr.Visible = False
        Account_tr.Visible = False
        Me.trOrgName.Visible = False
        Label2.Visible = False
        '------------------20090226 andy add
        Fun_tr.Visible = False
        '----------------
        msg.Text = "查無資料!!"

        If dt.Rows.Count > 0 Then
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()

            tbSearch1.Visible = True
            DG_ClassInfo.Visible = True 'TPanel@DG_ClassInfo
            PageControler1.Visible = True 'TPanel@PageControler1
            Reason_tr.Visible = True
            Account_tr.Visible = True

            trOrgName.Visible = True
            Label2.Visible = True
            '------------------20090226 andy add
            Fun_tr.Visible = True
            '----------------
            msg.Text = ""
        End If
    End Sub

    'SAVE
    Private Sub DG_ClassInfo_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DG_ClassInfo.ItemCommand
        'Const Cst_選擇 As Integer = 0
        'Const Cst_OrgName As Integer = 1
        'Const Cst_ClassID2 As Integer = 2
        'Const Cst_STDate As Integer = 3 '開訓日期
        Const Cst_FTDate As Integer = 4 '結訓日期
        'Const Cst_ClassCName As Integer = 5
        'Const Cst_TrainName As Integer = 6
        Const Cst_OCID As Integer = 7 '班級代碼
        Const Cst_RightID As Integer = 8 '已結訓班級使用授權檔
        'Const Cst_Name As Integer = 9 '已授權給(授權帳號:姓名)

        Dim sql As String = ""
        Dim MsgStr As String = ""
        Dim strOCID As String = e.Item.Cells(Cst_OCID).Text
        Dim strRightID As String = e.Item.Cells(Cst_RightID).Text
        Dim strFDate As String = e.Item.Cells(Cst_FTDate).Text

        Select Case e.CommandName
            Case "Add"
                '20090617 andy  edit  只有「e網報名審核-匯入」功能不管結訓日期，開放其它功能要判斷該班結訓日期
                '---------------------  
                'Dim i As Integer = 0

                For i As Integer = 0 To cb_SelFunID.Items.Count - 1
                    If cb_SelFunID.Items(i).Selected = True Then
                        Dim flagCheckGo As Boolean = True  '需要檢核 判斷該班結訓日期
                        Select Case cb_SelFunID.Items(i).Value
                            Case cst_funid_262_e網報名審核, cst_funid_70_報名登錄, cst_funid_76_甄試成績登錄, cst_funid_79_錄訓作業
                                flagCheckGo = False '排除檢核
                            Case cst_funid_83_學員資料維護
                                flagCheckGo = False '排除檢核
                            Case cst_funid_82_學員參訓
                                flagCheckGo = False '排除檢核
                        End Select
                        Select Case cb_SelFunID.Items(i).Value
                            Case TIMS.cst_FunID_e網報名審核, TIMS.cst_FunID_學員資料維護
                                'e網報名審核: 262
                                '學員資料維護: 83
                                flagCheckGo = False '排除檢核 尚未結訓
                        End Select

                        If flagCheckGo AndAlso strFDate <> "" Then
                            If CDate(strFDate) > CDate(Date.Now.ToString("yyyy-MM-dd")) Then
                                Common.MessageBox(Me, "此班尚未結訓！")
                                Exit Sub
                            End If
                        End If
                    End If
                Next
                '---------------------
                '20090226 andy  edit
                '-------------------
                MsgStr = ""
                If Chkdata(MsgStr) = False Then
                    Common.MessageBox(Me, MsgStr)
                    Exit Sub
                End If
                If chkNoRecordInAuth_Rend(Account.SelectedValue.ToString(), strOCID) = False Then
                    Common.MessageBox(Me, "同一班級只限授權一個帳號！")
                    Exit Sub
                End If
                '-------------------
                sql = ""
                sql &= " INSERT INTO Auth_REndClass ("
                sql += " RightID,Years,PlanID,DistID,OCID,Account,CreateDate,UseAble,ModifyAcct,ModifyDate,Reason,ReasonID,FunID,EndDate "
                sql += " ) VALUES ("
                sql += " @RightID,@Years,@PlanID,@DistID,@OCID,@Account,getdate(),@UseAble,@ModifyAcct, getdate(),@Reason,@ReasonID,@FunID,@EndDate "
                sql += " ) "
                'sql = ""
                'sql &= " INSERT INTO Auth_REndClass "
                'sql += " (RightID,Years,PlanID,DistID,OCID,Account,CreateDate,UseAble,ModifyAcct,ModifyDate,Reason,ReasonID,FunID,EndDate) "
                'sql += " values(" & Auth_REndClass_MaxNo() & ",'" & Me.yearlist.SelectedValue & "','" & Me.planlist.SelectedValue & "' "
                'sql += " ,'" & Me.DistID.SelectedValue & "','" & strOCID & "','" & Me.Account.SelectedValue & "',getdate(),'Y' "
                'sql += " ,'" & sm.UserInfo.UserID & "',getdate(), @Reason ,'" & Me.ReasonID.SelectedValue & "', '" & chkSelFunID() & "', " & TIMS.to_date(Me.EndDate.Text) & " )"
                Dim iRightID As Integer = DbAccess.GetNewId(objconn, "AUTH_RENDCLASS_RIGHTID_SEQ,AUTH_RENDCLASS,RIGHTID")
                Call TIMS.OpenDbConn(objconn)
                Dim oCmd As New SqlCommand(sql, objconn)
                With oCmd
                    .Parameters.Clear()
                    .Parameters.Add("RightID", SqlDbType.Int).Value = iRightID 'Auth_REndClass_MaxNo()
                    .Parameters.Add("Years", SqlDbType.VarChar).Value = Me.yearlist.SelectedValue
                    .Parameters.Add("PlanID", SqlDbType.VarChar).Value = Me.planlist.SelectedValue
                    .Parameters.Add("DistID", SqlDbType.VarChar).Value = Me.DistID.SelectedValue
                    .Parameters.Add("OCID", SqlDbType.VarChar).Value = strOCID
                    .Parameters.Add("Account", SqlDbType.VarChar).Value = Me.Account.SelectedValue

                    .Parameters.Add("UseAble", SqlDbType.VarChar).Value = "Y"
                    .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = sm.UserInfo.UserID

                    .Parameters.Add("Reason", SqlDbType.NVarChar).Value = Me.Reason.Text
                    .Parameters.Add("ReasonID", SqlDbType.VarChar).Value = Me.ReasonID.SelectedValue
                    .Parameters.Add("FunID", SqlDbType.VarChar).Value = TIMS.GetSelFunID(cb_SelFunID)
                    .Parameters.Add("EndDate", SqlDbType.DateTime).Value = TIMS.Cdate2(Me.EndDate.Text)

                    '.ExecuteNonQuery()
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    DbAccess.ExecuteNonQuery(oCmd.CommandText, objconn, oCmd.Parameters)
                End With
                'DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "新增成功")

                Call dt_search()

            Case "Upd"
                '20090617 andy  edit  只有「e網報名審核-匯入」功能不管結訓日期，開放其它功能要判斷該班結訓日期
                '---------------------  
                'Dim i As Integer = 0
                For i As Integer = 0 To cb_SelFunID.Items.Count - 1
                    If cb_SelFunID.Items(i).Selected = True Then
                        Dim flagCheckGo As Boolean = True  '需要檢核 判斷該班結訓日期
                        Select Case cb_SelFunID.Items(i).Value
                            Case cst_funid_262_e網報名審核, cst_funid_70_報名登錄, cst_funid_76_甄試成績登錄, cst_funid_79_錄訓作業
                                flagCheckGo = False '排除檢核
                            Case cst_funid_83_學員資料維護
                                flagCheckGo = False '排除檢核
                            Case cst_funid_82_學員參訓
                                flagCheckGo = False '排除檢核
                        End Select
                        Select Case cb_SelFunID.Items(i).Value
                            Case TIMS.cst_FunID_e網報名審核, TIMS.cst_FunID_學員資料維護
                                'e網報名審核: 262
                                '學員資料維護: 83
                                flagCheckGo = False '排除檢核 尚未結訓
                        End Select

                        If flagCheckGo AndAlso strFDate <> "" Then
                            If DateDiff(DateInterval.Day, CDate(Now), CDate(strFDate)) > 0 Then
                                Common.MessageBox(Me, "此班尚未結訓！")
                                Exit Sub
                            End If
                        End If

                    End If
                Next
                '---------------------
                '20090226 andy  edit
                '-------------------
                If strRightID = "XX" Then '已結訓班級使用授權檔
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行修改動作！")
                    Exit Sub
                End If
                MsgStr = ""
                If Chkdata(MsgStr) = False Then
                    Common.MessageBox(Me, MsgStr)
                    Exit Sub
                End If
                '---------------

                Dim uSql As String = ""
                uSql = ""
                uSql &= " UPDATE Auth_REndClass"
                uSql &= " SET Account =@Account"
                uSql &= " ,Reason =@Reason"
                uSql &= " ,ReasonID =@ReasonID"
                uSql &= " ,ModifyAcct =@ModifyAcct"
                uSql &= " ,ModifyDate = getdate()"
                uSql &= " ,FunID = @FunID"
                uSql &= " ,EndDate= @EndDate"
                uSql &= " WHERE RightID= @RightID"
                uSql &= " And OCID= @OCID"
                Dim pParms As New Hashtable
                pParms.Clear()
                pParms.Add("Account", Me.Account.SelectedValue)
                pParms.Add("Reason", Me.Reason.Text)
                pParms.Add("ReasonID", Me.ReasonID.SelectedValue)
                pParms.Add("ModifyAcct", sm.UserInfo.UserID)
                pParms.Add("FunID", TIMS.GetSelFunID(cb_SelFunID))
                pParms.Add("EndDate", TIMS.Cdate3(Me.EndDate.Text))
                pParms.Add("RightID", strRightID)
                pParms.Add("OCID", strOCID)
                DbAccess.ExecuteNonQuery(uSql, objconn, pParms)
                Common.MessageBox(Me, "修改成功")

                '.ExecuteNonQuery()
                ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                'oCmd.CommandText
                'Call TIMS.OpenDbConn(objconn)
                'Dim oCmd As New SqlCommand(sql, objconn)
                'With oCmd
                'End With
                Call dt_search()

            Case "Del"
                If strRightID = "XX" Then
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行刪除動作！")
                    Exit Sub
                End If
                sql = ""
                sql &= " UPDATE Auth_REndClass "
                sql += " SET UseAble = 'N'"
                sql += " ,ModifyAcct = @ModifyAcct "
                sql += " ,ModifyDate = getdate() "
                sql += " WHERE RightID = @RightID "
                sql += " And OCID = @OCID "

                Dim parms As Hashtable = New Hashtable()
                parms.Clear()
                parms.Add("ModifyAcct", sm.UserInfo.UserID)
                parms.Add("RightID", TIMS.ClearSQM(strRightID))
                parms.Add("OCID", TIMS.ClearSQM(strOCID))
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
                Common.MessageBox(Me, "刪除成功")

                Call dt_search()

            Case "GetData"
                If strRightID = "XX" Then
                    Common.MessageBox(Me, "目前尚未針對本班級授權,無法進行取得動作！")
                    Exit Sub
                End If
                sql = "select * from Auth_REndClass WHERE RightID = @RightID "
                Dim parms As Hashtable = New Hashtable()
                parms.Clear()
                parms.Add("RightID", strRightID)
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
                If dt.Rows.Count > 0 Then
                    Dim dr As DataRow = dt.Rows(0)
                    Common.SetListItem(yearlist, dr("Years"))
                    Common.SetListItem(planlist, dr("PlanID"))
                    Common.SetListItem(DistID, dr("DistID"))

                    Common.SetListItem(Account, dr("Account"))
                    Common.SetListItem(ReasonID, dr("ReasonID"))
                    Reason.Text = Convert.ToString(dr("Reason"))
                    Me.EndDate.Text = TIMS.Cdate3(dr("EndDate"))
                    If Convert.ToString(dr("FunID")) <> "" Then
                        For i As Int16 = 0 To cb_SelFunID.Items.Count - 1
                            cb_SelFunID.Items(i).Selected = False
                            If Convert.ToString(dr("FunID")).IndexOf(cb_SelFunID.Items(i).Value) > -1 Then
                                cb_SelFunID.Items(i).Selected = True
                            End If
                        Next
                    End If
                End If

        End Select

    End Sub

    Private Sub DG_ClassInfo_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_ClassInfo.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem, ListItemType.EditItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim but1 As LinkButton = e.Item.FindControl("but1") '新增
                Dim but2 As LinkButton = e.Item.FindControl("but2") '修改
                Dim but3 As LinkButton = e.Item.FindControl("but3") '刪除
                Dim but4 As LinkButton = e.Item.FindControl("but4") '取得

                'Dim but4 As Button = e.Item.FindControl("but4") '查看
                'but4.CommandArgument = Convert.ToString(drv("RightID"))
                If Convert.ToString(drv("Temp1")) <> "" Then
                    TIMS.Tooltip(e.Item, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but1, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but2, Convert.ToString(drv("Temp1")))
                    TIMS.Tooltip(but3, Convert.ToString(drv("Temp1")))
                End If

                but3.Attributes("onclick") = "return confirm('確定要刪除此授權?');"
                'but3.Attributes("onclick") = "return ChkData();return confirm('確定要刪除此授權?');"   '20090226 andy edit
                '20090226 andy edit
                '---------------------
                'but1.Attributes("onclick") = "return ChkData();"
                'but2.Attributes("onclick") = "return ChkData();"
                '---------------------
        End Select
    End Sub

    '查詢
    Private Sub rt_search_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles rt_search.Click
        dt_search()
    End Sub

    '依計畫 查詢
    Private Sub planlist_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles planlist.SelectedIndexChanged
        dt_search()
    End Sub

    Private Sub ddlOrgName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlOrgName.SelectedIndexChanged
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        '帳號
        Call MakeAccount(Me.Account, Me.planlist.SelectedValue, sm.UserInfo.LID, sm.UserInfo.RoleID, Me.DistID.SelectedValue, Me.ddlOrgName.SelectedValue)
    End Sub

End Class