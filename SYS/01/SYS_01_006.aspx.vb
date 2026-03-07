Partial Class SYS_01_006
    Inherits AuthBasePage

    'SYS_EMAILCODE
    '上午 08:31:01    '排程名稱：dbt_240429    '執行位置：C:\batch\dbt_240429\dbt_240429.exe
    '上午 07:21:01    '排程名稱：dbt_241007    '執行位置：C:\batch\dbt_241007\dbt_241007.exe
    '上午 07:31:01    '排程名稱：dbt_251127    '執行位置：C:\batch\dbt_251127\dbt_251127.exe
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        PageControler1.PageDataGrid = DataGrid1
        If Not Me.Page.IsPostBack Then Call cCreate1()

    End Sub

    Sub cCreate1()
        Panel_Sch.Visible = True
        Panel_edit.Visible = False

        '預設排除停用
        CB_ISUSED_N_NOSHOW.Checked = True

        TIMS.Get_YearTPlan(ddlTPlanSCH, sm.UserInfo.Years, sm.UserInfo.TPlanID, TIMS.cst_NO, objconn)
        Common.SetListItem(ddlTPlanSCH, sm.UserInfo.TPlanID)
        ddlTPlanSCH.Enabled = False

        TIMS.Get_YearTPlan(ddlTPlan, sm.UserInfo.Years, sm.UserInfo.TPlanID, TIMS.cst_NO, objconn)
        Common.SetListItem(ddlTPlan, sm.UserInfo.TPlanID)
        ddlTPlan.Enabled = False

        '載入分署(中心)
        Dim dtDISTID As DataTable = TIMS.Get_DISTIDdt(objconn)
        ddl_DistID = TIMS.Get_DistID(ddl_DistID, dtDISTID)
        If sm.UserInfo.DistID <> "000" Then
            Dim liDistID As ListItem = ddl_DistID.Items.FindByValue("000")
            If liDistID IsNot Nothing Then ddl_DistID.Items.Remove(liDistID)
        End If
        'ddl_DistID.Items.Remove(ddl_DistID.Items.FindByValue(""))
        Try
            rbl_EMAILCODE = TIMS.Get_EMAILCODE(sm, rbl_EMAILCODE, objconn)
            If rbl_EMAILCODE.Items.Count = 0 Then labERR_EMAILCODE.Text = "(計畫選擇有誤)"
            If rbl_EMAILCODE.Items.Count > 1 Then rbl_EMAILCODE.Items.Insert(0, New ListItem("不區分", ""))
        Catch ex As Exception
            labERR_EMAILCODE.Text = "(計畫選擇有誤)"
        End Try

        ddl_DistID2 = TIMS.Get_DistID(ddl_DistID2, dtDISTID)
        If sm.UserInfo.DistID <> "000" Then
            Dim liDistID As ListItem = ddl_DistID2.Items.FindByValue("000")
            If liDistID IsNot Nothing Then ddl_DistID2.Items.Remove(liDistID)
        End If
        ddl_DistID2.Items.Remove(ddl_DistID2.Items.FindByValue(""))
        CBL_EMAILCODE2 = TIMS.Get_EMAILCODE(sm, CBL_EMAILCODE2, objconn)

        If sm.UserInfo.LID <> 0 Then
            Common.SetListItem(ddl_DistID, sm.UserInfo.DistID)
            Common.SetListItem(ddl_DistID2, sm.UserInfo.DistID)
            ddl_DistID.Enabled = (sm.UserInfo.LID = 0)
            ddl_DistID2.Enabled = (sm.UserInfo.LID = 0)
        End If
        rdo_role = TIMS.Get_rolelist(sm, rdo_role, objconn)
    End Sub

    Sub Search1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim vddl_DistID As String = TIMS.GetListValue(ddl_DistID)
        Dim v_rbl_EMAILCODE As String = TIMS.GetListValue(rbl_EMAILCODE)
        Dim v_RBL_FUNC_USE As String = TIMS.GetListValue(RBL_FUNC_USE)
        txtACCTNAME.Text = TIMS.ClearSQM(txtACCTNAME.Text)
        txtACCTID.Text = TIMS.ClearSQM(txtACCTID.Text)

        Dim PMS_1 As New Hashtable From {{"TPLANID", sm.UserInfo.TPlanID}}
        Dim sSql As String = ""
        sSql &= " SELECT dd.NAME DISTNAME,c.ECNAME" '轄區分署 '功能
        sSql &= " ,a.ACCOUNT,CONCAT(a.NAME,'(',r.NAME,')','[',a.ACCOUNT,']',CASE a.ISUSED WHEN 'N' THEN '-(已停用)' END) ACCTNAME" '寄發對象
        sSql &= " ,b.DISTID,b.ECSEQ,b.EASEQ,a.EMAIL"
        sSql &= " ,b.FGUSE,CASE b.FGUSE WHEN 'Y' THEN 'Y' ELSE 'N' END FGUSE_N" '啟用(Y/N) 
        sSql &= " FROM AUTH_ACCOUNT a" & vbCrLf
        sSql &= " JOIN ID_ROLE r on r.ROLEID=a.ROLEID" & vbCrLf
        sSql &= " JOIN AUTH_EMAILACCT b on b.ACCOUNT=a.ACCOUNT" & vbCrLf
        sSql &= " JOIN ID_DISTRICT dd on dd.DISTID=b.DISTID" & vbCrLf
        sSql &= " JOIN SYS_EMAILCODE c on c.ECSEQ=b.ECSEQ" & vbCrLf
        sSql &= " WHERE c.TPLANID=@TPLANID" & vbCrLf
        If CB_ISUSED_N_NOSHOW.Checked Then sSql &= " AND a.ISUSED='Y'" & vbCrLf

        Select Case v_RBL_FUNC_USE
            Case "Y", "N"
                PMS_1.Add("FGUSE", v_RBL_FUNC_USE)
                sSql &= " AND b.FGUSE=@FGUSE" & vbCrLf
        End Select

        If v_rbl_EMAILCODE <> "" Then
            PMS_1.Add("ECSEQ", Val(v_rbl_EMAILCODE))
            sSql &= " AND b.ECSEQ=@ECSEQ" & vbCrLf
            'Else sSql &= " AND 1<>1" & vbCrLf
        End If
        If vddl_DistID <> "" Then
            PMS_1.Add("DISTID", vddl_DistID)
            sSql &= " AND b.DISTID=@DISTID" & vbCrLf
        End If

        Dim ACCTNAME_lk As String = txtACCTNAME.Text
        If ACCTNAME_lk <> "" Then
            PMS_1.Add("ACCTNAME_lk", ACCTNAME_lk)
            sSql &= " AND a.NAME LIKE '%'+@ACCTNAME_lk+'%'" & vbCrLf
        End If
        Dim ACCTID_lk As String = txtACCTID.Text
        If ACCTID_lk <> "" Then
            PMS_1.Add("ACCOUNT_lk", ACCTID_lk)
            sSql &= " AND a.ACCOUNT LIKE '%'+@ACCOUNT_lk+'%'" & vbCrLf
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, PMS_1)

        DataGridTable.Visible = False
        msg1.Text = "查無資料"

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return

        DataGridTable.Visible = True
        msg1.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub
    Private Sub Utl_Clear1()
        'ddl_DistID2.Enabled = True
        'rdo_role.Enabled = True
        'DDL_ACCOUNT1.Enabled = True
        'bt_data_sch2.Visible = True
        'CBL_EMAILCODE2.Enabled = True

        Hid_EASEQ.Value = "" 'TIMS.GetMyValue(sCmdArg, "EASEQ")
        Hid_ECSEQ.Value = ""
        Hid_DISTID.Value = "" 'TIMS.GetMyValue(sCmdArg, "DISTID")
        Hid_ACCOUNT.Value = "" 'TIMS.GetMyValue(sCmdArg, "ACCOUNT")
        Hid_EMAIL.Value = ""

        TIMS.SetCblValue(CBL_EMAILCODE2, "")
    End Sub

    Protected Sub btn_Sch_Click(sender As Object, e As EventArgs) Handles btn_Sch.Click
        Utl_Clear1()

        Search1()
    End Sub

    Protected Sub btn_Add_Click(sender As Object, e As EventArgs) Handles btn_Add.Click
        ADD_NEWDATA()
    End Sub

    Private Sub ADD_NEWDATA()
        Panel_Sch.Visible = False
        Panel_edit.Visible = True

        Utl_Clear1()

        Common.SetListItem(ddl_DistID2, sm.UserInfo.DistID)
        ddl_DistID2.Enabled = (sm.UserInfo.LID = 0)

        'Dim v_rbl_EMAILCODE As String = TIMS.GetListValue(rbl_EMAILCODE)
        'Common.SetListItem(rbl_EMAILCODE2, v_rbl_EMAILCODE)
    End Sub


    Private Sub btn_data_search2(ByVal fg_reSch As Boolean)
        'fg_reSch TRUE : 重新查詢
        Dim v_ddl_DistID2 As String = TIMS.GetListValue(ddl_DistID2)
        Dim v_rdo_role As String = TIMS.GetListValue(rdo_role)

        Dim hparams As New Hashtable
        Dim sSql As String = ""
        sSql &= " SELECT a.ACCOUNT ,CONCAT(a.NAME,'(',r.NAME,')','[',a.ACCOUNT,']',CASE a.ISUSED WHEN 'N' THEN '-(已停用)' END) ACCTNAME" & vbCrLf
        sSql &= " ,a.ISUSED,a.EMAIL,oo.ORGNAME, oo.DISTNAME" & vbCrLf
        sSql &= " ,a.ROLEID,r.NAME ROLENAME" & vbCrLf ',e.EASEQ" & vbCrLf
        sSql &= " FROM AUTH_ACCOUNT a" & vbCrLf
        sSql &= " JOIN ID_ROLE r on r.ROLEID=a.ROLEID" & vbCrLf
        sSql &= " JOIN V_DISTRICT oo on oo.ORGID=a.ORGID" & vbCrLf
        'sSql &= " LEFT JOIN AUTH_EMAILACCT e on e.ACCOUNT=a.ACCOUNT " & vbCrLf
        sSql &= " WHERE a.EMAIL is not null and len(a.EMAIL)>1" & vbCrLf
        If CB_ISUSED_N_NOSHOW.Checked Then sSql &= " AND a.ISUSED='Y'" & vbCrLf
        If v_ddl_DistID2 <> "" Then
            hparams.Add("DISTID", v_ddl_DistID2)
            sSql &= " AND oo.DISTID=@DISTID" & vbCrLf
        Else
            sSql &= " AND 1<>1" & vbCrLf
        End If
        If v_rdo_role <> "" Then
            hparams.Add("ROLEID", v_rdo_role)
            sSql &= " AND a.ROLEID=@ROLEID" & vbCrLf
        End If
        'If v_ACCOUNT <> "" Then
        '    hparams.Add("ACCOUNT", v_ACCOUNT)
        '    sSql &= " AND a.ACCOUNT=@ACCOUNT" & vbCrLf
        'End If
        sSql &= " ORDER BY a.ROLEID,oo.DISTID,a.ACCOUNT" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, hparams)

        Dim s_drACCOUNT As String = ""
        Dim v_ACCOUNT_OLD1 As String = TIMS.GetListValue(DDL_ACCOUNT1)
        If fg_reSch AndAlso v_ACCOUNT_OLD1 <> "" Then
            Common.SetListItem(DDL_ACCOUNT1, v_ACCOUNT_OLD1)
        ElseIf Hid_ACCOUNT.Value <> "" Then
            Common.SetListItem(DDL_ACCOUNT1, Hid_ACCOUNT.Value)
        ElseIf (dt.Rows.Count > 0) Then
            s_drACCOUNT = Convert.ToString(dt.Rows(0)("ACCOUNT"))
            Common.SetListItem(DDL_ACCOUNT1, s_drACCOUNT)
        End If

        Dim fg_dthavedata As Boolean = (dt IsNot Nothing)
        Dim fg_dtrowcount As Integer = 0
        If fg_dthavedata Then fg_dtrowcount = dt.Rows.Count
        'Common.SetListItem(DDL_ACCOUNT1, "")
        'DDL_ACCOUNT1.SelectedValue = ""
        DDL_ACCOUNT1.SelectedIndex = -1
        DDL_ACCOUNT1.ClearSelection()
        DDL_ACCOUNT1.Items.Clear()
        If fg_dthavedata AndAlso dt.Rows.Count > 0 Then
            s_drACCOUNT = Convert.ToString(dt.Rows(0)("ACCOUNT"))
            Try
                DDL_ACCOUNT1.SelectedValue = s_drACCOUNT
                With DDL_ACCOUNT1
                    .DataSource = dt
                    .DataTextField = "ACCTNAME"
                    .DataValueField = "ACCOUNT"
                    .DataBind()
                End With
            Catch ex As Exception
                Dim strErrmsg As String = String.Concat("* ex.Message: ", ex.Message, vbCrLf)
                strErrmsg &= String.Concat("dt IsNot Nothing: ", fg_dthavedata, vbCrLf)
                strErrmsg &= String.Concat("dt.Rows.Count: ", fg_dtrowcount, vbCrLf)
                strErrmsg &= String.Concat("v_ACCOUNT_OLD1: ", v_ACCOUNT_OLD1, vbCrLf)
                strErrmsg &= String.Concat("Hid_ACCOUNT.Value: ", Hid_ACCOUNT.Value, vbCrLf)
                strErrmsg &= String.Concat("s_drACCOUNT: ", s_drACCOUNT, vbCrLf)

                Call TIMS.WriteTraceLog(strErrmsg, ex)
                Return
            End Try
        End If

        Dim v_ACCOUNT As String = TIMS.GetListValue(DDL_ACCOUNT1)
        If fg_reSch AndAlso v_ACCOUNT <> "" Then
            Common.SetListItem(DDL_ACCOUNT1, v_ACCOUNT)

            Dim hPMSA As New Hashtable From {{"ACCOUNT", v_ACCOUNT}}
            Dim sqlA As String = ""
            sqlA &= " SELECT ACCOUNT,ROLEID,LID,NAME,PHONE,EMAIL,ORGID,ISUSED"
            sqlA &= " FROM AUTH_ACCOUNT WHERE ACCOUNT=@ACCOUNT"
            Dim drA As DataRow = DbAccess.GetOneRow(sqlA, objconn, hPMSA)
            If drA Is Nothing Then
                TIMS.SetCblValue(CBL_EMAILCODE2, "")
                Common.MessageBox(Me, "查無有效帳號資料!")
                Return
            End If
            Dim hPMS2 As New Hashtable From {{"ACCOUNT", v_ACCOUNT}}
            Dim sql2 As String = "SELECT EASEQ,ECSEQ,DISTID,ACCOUNT,FGUSE FROM AUTH_EMAILACCT WHERE ACCOUNT=@ACCOUNT"
            Dim dt2 As DataTable = DbAccess.GetDataTable(sql2, objconn, hPMS2)
            If dt2 Is Nothing OrElse dt2.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無有效資料!")
                'Return
            End If
            Common.SetListItem(rdo_role, drA("ROLEID"))
            'CB_FGUSE.Checked = If(Convert.ToString(dr2("FGUSE")) = "Y", True, False)
            Dim vCBL2 As String = ""
            For Each dr2 As DataRow In dt2.Rows
                If Convert.ToString(dr2("FGUSE")) = "Y" Then
                    vCBL2 &= String.Concat(If(vCBL2 <> "", ",", ""), dr2("ECSEQ"))
                End If
            Next
            TIMS.SetCblValue(CBL_EMAILCODE2, vCBL2)
        End If

    End Sub

    Protected Sub bt_backoff_Click(sender As Object, e As EventArgs) Handles bt_backoff.Click
        Panel_Sch.Visible = True
        Panel_edit.Visible = False

        Utl_Clear1()

        Search1()
    End Sub

    Protected Sub bt_save_Click(sender As Object, e As EventArgs) Handles bt_save.Click
        Call SAVEDATA1()
    End Sub

    Function GET_DISTID_ACCOUNT(vACCOUNT As String) As String
        Dim v_ddl_DistID2 As String = TIMS.GetListValue(ddl_DistID2)
        Dim rst As String = If(v_ddl_DistID2 <> "", v_ddl_DistID2, sm.UserInfo.DistID)
        Dim vTPLANID As String = sm.UserInfo.TPlanID
        Dim vYEARS As String = Convert.ToString(sm.UserInfo.Years)
        Dim pms_1 As New Hashtable From {{"TPLANID", vTPLANID}, {"YEARS", vYEARS}, {"ACCOUNT", vACCOUNT}}
        Dim sql_1 As String = "SELECT DISTID FROM VIEW_LOGINACCOUNT WHERE TPLANID=@TPLANID AND YEARS=@YEARS AND ACCOUNT=@ACCOUNT"
        Dim dr As DataRow = DbAccess.GetOneRow(sql_1, objconn, pms_1)
        If dr IsNot Nothing Then rst = dr("DISTID")
        Return rst
    End Function

    Private Sub SAVEDATA1()
        'Dim vddl_DistID2 As String = TIMS.GetListValue(ddl_DistID2)
        Dim v_DDL_ACCOUNT1 As String = TIMS.GetListValue(DDL_ACCOUNT1)
        'Dim v_CBL_EMAILCODE2 As String = TIMS.GetCblValue(CBL_EMAILCODE2)
        If v_DDL_ACCOUNT1 = "" Then
            Common.MessageBox(Me, "請選擇 帳號!")
            Return
            'ElseIf v_CBL_EMAILCODE2 = "" Then
            '    Common.MessageBox(Me, "請勾選 功能1")
            '    Return
        End If
        'Dim vFGUSE As String = If(CB_FGUSE.Checked, "Y", "N")

        For Each item As ListItem In CBL_EMAILCODE2.Items
            Dim vECSEQ As String = TIMS.ClearSQM(item.Value)
            Dim vDISTID As String = Get_DISTID_ACCOUNT(v_DDL_ACCOUNT1)
            Dim vFGUSE As String = If(item.Selected, "Y", "N")
            Dim hPMS As New Hashtable From {{"ECSEQ", Val(vECSEQ)}, {"DISTID", vDISTID}, {"ACCOUNT", v_DDL_ACCOUNT1}, {"FGUSE", vFGUSE}}
            UPDATE_AUTH_EMAILACCT(hPMS)
        Next

        Utl_Clear1()

        Search1()
    End Sub

    Sub UPDATE_AUTH_EMAILACCT(ByRef hpms As Hashtable)
        Dim iECSEQ As Integer = Val(TIMS.GetMyValue2(hpms, "ECSEQ"))
        Dim sDISTID As String = TIMS.GetMyValue2(hpms, "DISTID")
        Dim sACCOUNT As String = TIMS.GetMyValue2(hpms, "ACCOUNT")
        Dim sFGUSE As String = TIMS.GetMyValue2(hpms, "FGUSE")

        Dim params As New Hashtable From {{"ECSEQ", iECSEQ}, {"ACCOUNT", sACCOUNT}}
        Dim sSql As String = " SELECT * FROM AUTH_EMAILACCT WHERE ECSEQ=@ECSEQ AND ACCOUNT=@ACCOUNT"
        Dim dt As DataTable = DbAccess.GetDataTable(sSql, objconn, params)

        If dt Is Nothing OrElse dt.Rows.Count = 0 Then
            Dim iEASEQ As Integer = DbAccess.GetNewId(objconn, "AUTH_EMAILACCT_EASEQ_SEQ,AUTH_EMAILACCT,EASEQ")
            Dim iParms As New Hashtable From {
                {"EASEQ", iEASEQ},
                {"ECSEQ", iECSEQ},
                {"DISTID", sDISTID},
                {"ACCOUNT", sACCOUNT},
                {"FGUSE", sFGUSE},
                {"MODIFYACCT", sm.UserInfo.UserID}
            }
            'iParms.Add("MODIFYDATE", MODIFYDATE)
            Dim isSql As String = ""
            isSql &= " INSERT INTO AUTH_EMAILACCT(EASEQ,ECSEQ,DISTID,ACCOUNT,FGUSE,MODIFYACCT,MODIFYDATE)" & vbCrLf
            isSql &= " VALUES(@EASEQ,@ECSEQ,@DISTID,@ACCOUNT,@FGUSE,@MODIFYACCT,GETDATE())" & vbCrLf
            DbAccess.ExecuteNonQuery(isSql, objconn, iParms)
        Else
            Dim iEASEQ As Integer = dt.Rows(0)("EASEQ")
            'uParms.Add("MODIFYDATE", MODIFYDATE)
            Dim uParms As New Hashtable From {
                {"DISTID", sDISTID},
                {"FGUSE", sFGUSE},
                {"MODIFYACCT", sm.UserInfo.UserID},
                {"EASEQ", iEASEQ},
                {"ECSEQ", iECSEQ},
                {"ACCOUNT", sACCOUNT}
            }
            Dim usSql As String = ""
            usSql &= " UPDATE AUTH_EMAILACCT" & vbCrLf
            usSql &= " SET DISTID=@DISTID,FGUSE=@FGUSE,MODIFYACCT=@MODIFYACCT,MODIFYDATE=GETDATE()" & vbCrLf
            usSql &= " WHERE EASEQ=@EASEQ AND ECSEQ=@ECSEQ AND ACCOUNT=@ACCOUNT" & vbCrLf
            DbAccess.ExecuteNonQuery(usSql, objconn, uParms)
        End If

    End Sub

    Private Sub DELETE_AUTH_EMAILACCT(ByRef hPMS As Hashtable)
        Dim iEASEQ As Integer = Val(TIMS.GetMyValue2(hPMS, "EASEQ"))
        Dim iECSEQ As Integer = Val(TIMS.GetMyValue2(hPMS, "ECSEQ"))
        Dim sDISTID As String = TIMS.GetMyValue2(hPMS, "DISTID")
        Dim sACCOUNT As String = TIMS.GetMyValue2(hPMS, "ACCOUNT")

        Dim pms_D As New Hashtable From {{"EASEQ", iEASEQ}, {"ECSEQ", iECSEQ}, {"DISTID", sDISTID}, {"ACCOUNT", sACCOUNT}}
        Dim sSql_D As String = " DELETE AUTH_EMAILACCT WHERE EASEQ=@EASEQ AND ECSEQ=@ECSEQ AND DISTID=@DISTID AND ACCOUNT=@ACCOUNT"
        DbAccess.ExecuteNonQuery(sSql_D, objconn, pms_D)
    End Sub

    Protected Sub bt_data_sch2_Click(sender As Object, e As EventArgs) Handles bt_data_sch2.Click
        Dim v_acct As String = TIMS.GetListValue(DDL_ACCOUNT1)
        If v_acct = "" Then
            Common.MessageBox(Me, "請選擇一組帳號!")
            Return
        End If

        btn_data_search2(True)
    End Sub

    Protected Sub ddl_DistID2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddl_DistID2.SelectedIndexChanged
        'If Hid_ACCOUNT.Value = "" Then btn_data_search2(False)
        btn_data_search2(False)
    End Sub
    Protected Sub rdo_role_SelectedIndexChanged(sender As Object, e As EventArgs) Handles rdo_role.SelectedIndexChanged
        'If Hid_ACCOUNT.Value = "" Then btn_data_search2(False)
        btn_data_search2(False)
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e Is Nothing OrElse e.CommandArgument = "" Then Return

        Utl_Clear1()

        Dim sCmdArg As String = e.CommandArgument
        Hid_EASEQ.Value = TIMS.GetMyValue(sCmdArg, "EASEQ")
        Hid_ECSEQ.Value = TIMS.GetMyValue(sCmdArg, "ECSEQ")
        Hid_DISTID.Value = TIMS.GetMyValue(sCmdArg, "DISTID")
        Hid_ACCOUNT.Value = TIMS.GetMyValue(sCmdArg, "ACCOUNT")
        Hid_EMAIL.Value = TIMS.GetMyValue(sCmdArg, "EMAIL")

        Select Case e.CommandName
            Case "btnUSED"
                Dim hPMS As New Hashtable From {{"ECSEQ", Val(Hid_ECSEQ.Value)}, {"DISTID", Hid_DISTID.Value}, {"ACCOUNT", Hid_ACCOUNT.Value}, {"FGUSE", "Y"}}
                UPDATE_AUTH_EMAILACCT(hPMS)
                Search1()

            Case "btnNOUSE"
                Dim hPMS As New Hashtable From {{"ECSEQ", Val(Hid_ECSEQ.Value)}, {"DISTID", Hid_DISTID.Value}, {"ACCOUNT", Hid_ACCOUNT.Value}, {"FGUSE", "N"}}
                UPDATE_AUTH_EMAILACCT(hPMS)
                Search1()

            Case "btnDELE"
                Dim hPMS As New Hashtable From {{"EASEQ", Val(Hid_EASEQ.Value)}, {"ECSEQ", Val(Hid_ECSEQ.Value)}, {"DISTID", Hid_DISTID.Value}, {"ACCOUNT", Hid_ACCOUNT.Value}}
                DELETE_AUTH_EMAILACCT(hPMS)
                Search1()

            Case "btnEMAIL"
                Hid_EMAIL.Value = TIMS.ClearSQM(Hid_EMAIL.Value)
                Dim S_EMAIL As String = If(Hid_EMAIL.Value <> "", Hid_EMAIL.Value, "(查無資料)")
                Common.MessageBox(Me, S_EMAIL)
                Return

        End Select
    End Sub
    Private Sub DataGrid1_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e)

                Dim drv As DataRowView = e.Item.DataItem
                Dim btnUSED As LinkButton = e.Item.FindControl("btnUSED") '啟用
                Dim btnNOUSE As LinkButton = e.Item.FindControl("btnNOUSE") '停用
                Dim btnDELE As LinkButton = e.Item.FindControl("btnDELE") '刪除
                Dim btnEMAIL As LinkButton = e.Item.FindControl("btnEMAIL") 'EMAIL查詢

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "EASEQ", drv("EASEQ"))
                TIMS.SetMyValue(sCmdArg, "ECSEQ", drv("ECSEQ"))
                TIMS.SetMyValue(sCmdArg, "DISTID", drv("DISTID"))
                TIMS.SetMyValue(sCmdArg, "ACCOUNT", drv("ACCOUNT"))
                TIMS.SetMyValue(sCmdArg, "EMAIL", TIMS.ChangeEmail(Convert.ToString(drv("EMAIL"))))
                'TIMS.SetMyValue(sCmdArg, "FGUSE", Convert.ToString(drv("FGUSE")))

                btnUSED.CommandArgument = sCmdArg
                btnNOUSE.CommandArgument = sCmdArg
                btnDELE.CommandArgument = sCmdArg
                btnEMAIL.CommandArgument = sCmdArg

                btnNOUSE.Visible = (Convert.ToString(drv("FGUSE")) = "Y")
                btnUSED.Visible = Not btnNOUSE.Visible
                btnDELE.Visible = (Convert.ToString(drv("FGUSE")) <> "")
                btnDELE.Attributes("onclick") = "return confirm('確定要刪除這一筆賦予資料?');"
        End Select
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
