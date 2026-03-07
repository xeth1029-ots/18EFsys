Partial Class SYS_01_003_chk
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁  
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'Dim RoleID As Integer
        Dim rqApplyType As String = TIMS.ClearSQM(Request("ApplyType"))
        Dim rqfun As String = TIMS.ClearSQM(Request("fun"))
        Dim rqaccount As String = TIMS.ClearSQM(Request("account"))
        Dim rqplanid As String = TIMS.ClearSQM(Request("planid"))
        Dim rqRID As String = TIMS.ClearSQM(Request("rid"))

        If Not Page.IsPostBack Then
            Me.ViewState("PageReload") = -1                         '計算Reload次數
            Me.Role = TIMS.Get_IDRole(Me.Role, TIMS.dtNothing(), 0, objconn)

            If rqApplyType = "1" Or rqfun = "V" Or rqfun = "D" Then
                'Dim sdr As SqlDataReader
                Dim selsql As String = ""
                selsql = "" & vbCrLf
                selsql &= " Select a.Account , d.Years, f.DistName, e.PlanName, d.Seq, c.OrgName, a.PlanID, a.RoleID, " & vbCrLf
                selsql &= " a.Result,a.Note,a.Name,a.IDNO,a.Phone,a.EMail,a.ApplyType,a.LID,a.OrgID,g.RID, g2.RID RID2 " & vbCrLf
                selsql &= " From Auth_Apply a " & vbCrLf
                selsql &= " JOIN org_orginfo c on a.OrgID = c.OrgID " & vbCrLf
                selsql &= " JOIN ID_Plan d on a.planid=d.planid " & vbCrLf
                selsql &= " JOIN Key_Plan e on d.TPlanID=e.TPlanID " & vbCrLf
                selsql &= " JOIN (SELECT DistID,Name as DistName FROM ID_District) f ON d.DistID=f.DistID " & vbCrLf
                selsql &= " JOIN Auth_Relship g2 ON a.orgid=g2.orgid " & vbCrLf
                selsql &= " left JOIN Auth_Relship g ON a.PlanID=g.PlanID and a.orgid=g.orgid " & vbCrLf
                selsql &= " WHERE a.Account='" & rqaccount & "' And a.PlanID=" & rqplanid & " And a.ApplyType='" & rqApplyType & "' " & vbCrLf
                Dim dt As DataTable = DbAccess.GetDataTable(selsql, objconn)
                If dt.Rows.Count = 0 Then
                    Dim url1 As String = "SYS_01_003.aspx?ID=" & Request("ID")
                    TIMS.Utl_Redirect(Me, objconn, url1)
                    Exit Sub
                End If
                Dim sDr As DataRow = dt.Rows(0)

                Me.nameid.Text = Convert.ToString(sDr("Account"))
                Me.nameid.ReadOnly = True
                Me.PlanIDValue.Value = Convert.ToString(sDr("PlanID"))
                Me.OrgID.Value = Convert.ToString(sDr("OrgID"))
                RIDValue.Value = If(Convert.ToString(sDr("RID")).ToString <> "", Convert.ToString(sDr("RID")), Convert.ToString(sDr("RID2")))
                Me.LIDValue.Value = Convert.ToString(sDr("LID"))
                Me.TBplan.Text = Convert.ToString(sDr("Years")) & Convert.ToString(sDr("DistName")) & Convert.ToString(sDr("PlanName")) & Convert.ToString(sDr("Seq")) & " _ " & Convert.ToString(sDr("OrgName"))
                Common.SetListItem(Role, Convert.ToString(sDr("RoleID")))
                Me.name.Text = Convert.ToString(sDr("Name"))
                Me.IDNO.Text = TIMS.ChangeIDNO(Convert.ToString(sDr("IDNO")))
                Me.telphone.Text = Convert.ToString(sDr("Phone"))
                Me.email.Text = Convert.ToString(sDr("Email"))
                ApplyType.Text = If(Convert.ToString(sDr("ApplyType")).ToString = "1", "申請帳號計劃", "註銷帳號計劃")
                Me.Result.Checked = If(Convert.ToString(sDr("Result")) = "Y", True, False)
                Me.Note.Text = Convert.ToString(sDr("Note"))
                'sdr = DbAccess.GetReader(selsql, objconn)

                Me.nameid.BorderStyle = BorderStyle.None
                Me.TBplan.BorderStyle = BorderStyle.None
                Me.name.BorderStyle = BorderStyle.None
                Me.IDNO.BorderStyle = BorderStyle.None
                Me.telphone.BorderStyle = BorderStyle.None
                Me.email.BorderStyle = BorderStyle.None
                Me.ApplyType.BorderStyle = BorderStyle.None

                If rqfun = "V" Then
                    btu_save.Visible = False
                    Me.Result.Enabled = False
                    Me.Note.ReadOnly = True
                    Me.Role.Enabled = False
                    Me.Note.BorderStyle = BorderStyle.None
                ElseIf rqfun = "D" Then
                    Me.Role.Enabled = False
                End If
                'sdr.Close()
                'objconn.Close()
            End If
        Else
            Me.ViewState("PageReload") -= 1
        End If

        back.Attributes("onclick") = "history.go(" & Me.ViewState("PageReload") & ");"

    End Sub

    Sub save_del()
        Dim rqApplyType As String = TIMS.ClearSQM(Request("ApplyType"))
        Dim rqfun As String = TIMS.ClearSQM(Request("fun"))
        Dim rqaccount As String = TIMS.ClearSQM(Request("account"))
        Dim rqplanid As String = TIMS.ClearSQM(Request("planid"))
        Dim rqRID As String = TIMS.ClearSQM(Request("rid"))

        Dim Errmsg As String = ""
        Note.Text = TIMS.ClearSQM(Note.Text)
        If Not Result.Checked AndAlso Note.Text = "" Then
            Errmsg += "不同意時，請務必填寫備註資料!!!" & vbCrLf
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            Dim da As SqlDataAdapter = Nothing
            Dim sqldr As DataRow = Nothing
            Dim sqlTable As DataTable = Nothing
            Dim sqlstr As String = ""
            If Me.Result.Checked Then
                'sqldr("Result") = "Y"
                '======刪除帳號計劃=======
                sqlstr = "DELETE AUTH_ACCRWPLAN WHERE Account='" & rqaccount & "'  and PlanID=" & rqplanid & " and RID='" & rqRID & "'"
                Dim delrows As Integer = DbAccess.ExecuteNonQuery(sqlstr, objTrans)
                If delrows >= 1 Then
                    '=====若無可使用計劃清除身分證號及自然人憑證號 記錄=======
                    sqlstr = "select count(1) cnt FROM AUTH_ACCRWPLAN WHERE account='" & rqaccount & "' "
                    sqlTable = DbAccess.GetDataTable(sqlstr, da, objTrans)
                    sqldr = sqlTable.Rows(0)
                    If CInt(sqldr("cnt")) = 0 Then
                        sqlstr = "" & vbCrLf
                        sqlstr += " UPDATE AUTH_ACCOUNT " & vbCrLf
                        sqlstr += " SET IDNO =NULL, SerialNo=NULL " & vbCrLf
                        sqlstr += " ,ModifyAcct='" & sm.UserInfo.UserID & "' " & vbCrLf
                        sqlstr += " ,ModifyDate=getdate() " & vbCrLf
                        sqlstr += " where Account='" & rqaccount & "' " & vbCrLf
                        delrows = DbAccess.ExecuteNonQuery(sqlstr, objTrans)
                    End If
                    '======審核帳號記錄=======
                    sqlstr = "select * from auth_apply where Account='" & rqaccount & "' and planid=" & rqplanid & " and applytype='" & rqApplyType & "' "
                    sqlTable = DbAccess.GetDataTable(sqlstr, da, objTrans)
                    sqldr = sqlTable.Rows(0)
                    'sqldr("RoleID") = Me.Role.SelectedValue
                    sqldr("Result") = If(Result.Checked, "Y", "N")
                    sqldr("Note") = Note.Text
                    sqldr("ModifyAcct") = sm.UserInfo.UserID
                    sqldr("ModifyDate") = Now()
                    DbAccess.UpdateDataTable(sqlTable, da, objTrans)

                    DbAccess.CommitTrans(objTrans)
                    Page.RegisterStartupScript("", "<script>alert('審核完成!繼續審核~');location.href='SYS_01_003.aspx?ID=" & Request("ID") & "';</script>")
                Else
                    DbAccess.CommitTrans(objTrans)
                    Page.RegisterStartupScript("", "<script>alert('刪除沒有動作，請確認審核資料~');location.href='SYS_01_003.aspx?ID=" & Request("ID") & "';</script>")
                End If
            Else

                '======審核帳號記錄=======
                sqlstr = "select * from auth_apply where account='" & rqaccount & "' and planid=" & rqplanid & " and applytype='" & rqApplyType & "' "
                sqlTable = DbAccess.GetDataTable(sqlstr, da, objTrans)
                sqldr = sqlTable.Rows(0)
                'sqldr("RoleID") = Me.Role.SelectedValue
                sqldr("Result") = If(Result.Checked, "Y", "N")
                sqldr("Note") = Note.Text
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now()
                DbAccess.UpdateDataTable(sqlTable, da, objTrans)

                DbAccess.CommitTrans(objTrans)
                Page.RegisterStartupScript("", "<script>alert('審核完成!繼續審核~');location.href='SYS_01_003.aspx?ID=" & Request("ID") & "';</script>")
            End If

        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            TIMS.CloseDbConn(objconn)
            Throw ex
        End Try
    End Sub

    Sub save_chk()
        Dim rqApplyType As String = TIMS.ClearSQM(Request("ApplyType"))
        Dim rqfun As String = TIMS.ClearSQM(Request("fun"))
        Dim rqaccount As String = TIMS.ClearSQM(Request("account"))
        Dim rqplanid As String = TIMS.ClearSQM(Request("planid"))
        Dim rqRID As String = TIMS.ClearSQM(Request("rid"))

        Dim flag_Auth_AccRWFun As Boolean = True
        IDNO.Text = TIMS.ChangeIDNO(UCase(TIMS.ClearSQM(IDNO.Text)))

        Dim Errmsg As String = ""
        Dim str_nameid As String = nameid.Text
        nameid.Text = TIMS.ClearSQM(nameid.Text)
        If str_nameid <> nameid.Text Then
            Errmsg = "帳號" & TIMS.cst_ErrorMsg10 & vbCrLf
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        End If

        Dim v_Role As String = TIMS.GetListValue(Role)
        If v_Role = "" Then
            Errmsg += "角色為必填資料!!!" & vbCrLf
        End If
        Note.Text = TIMS.ClearSQM(Note.Text)
        If Not Result.Checked AndAlso Note.Text = "" Then
            Errmsg += "不同意時，請務必填寫備註資料!!!" & vbCrLf
        End If
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'If (Me.Result.Checked = False) And Note.Text = "" Then
        '    Common.MessageBox(Page, "不同意時，請務必填寫備註資料!!!")
        '    Exit Sub
        'End If

        Dim objTrans As SqlTransaction = DbAccess.BeginTrans(objconn)
        Try
            '======審核帳號記錄=======
            Dim sqldr As DataRow = Nothing
            Dim da As SqlDataAdapter = Nothing
            Dim sqlTable As DataTable = Nothing
            Dim sqlstr As String = ""
            sqlstr = "SELECT * FROM AUTH_APPLY WHERE ACCOUNT='" & rqaccount & "' and planid=" & rqplanid & " and applytype='" & rqApplyType & "' "
            sqlTable = DbAccess.GetDataTable(sqlstr, da, objTrans)
            sqldr = sqlTable.Rows(0)
            sqldr("RoleID") = v_Role 'Me.Role.SelectedValue
            sqldr("Result") = If(Result.Checked, "Y", "N") '若同意則使用:Y'不同意則停用:N
            sqldr("Note") = Note.Text
            sqldr("ModifyAcct") = sm.UserInfo.UserID
            sqldr("ModifyDate") = Now()
            DbAccess.UpdateDataTable(sqlTable, da, objTrans)

            '======新增正式帳號=======
            nameid.Text = TIMS.ClearSQM(nameid.Text)
            sqlstr = "SELECT * FROM AUTH_ACCOUNT WHERE ACCOUNT ='" & nameid.Text & "' "
            sqlTable = DbAccess.GetDataTable(sqlstr, da, objTrans)

            If sqlTable.Rows.Count > 0 Then
                flag_Auth_AccRWFun = False '不須再建立功能權限
                sqldr = sqlTable.Rows(0)
                'sqldr("account") = Me.nameid.Text
                sqldr("RoleID") = v_Role 'Me.Role.SelectedValue
                'sqldr("LID") = Me.LIDValue.Value
                sqldr("Name") = Me.name.Text
                sqldr("IDNO") = TIMS.ChangeIDNO(Me.IDNO.Text)
                sqldr("Phone") = Me.telphone.Text
                sqldr("email") = Me.email.Text
                sqldr("OrgID") = Me.OrgID.Value
                sqldr("IsUsed") = If(Result.Checked, "Y", "N") '若同意則使用:Y'不同意則停用:N
                sqldr("Note") = Note.Text
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now()

            Else
                flag_Auth_AccRWFun = True '須建立功能權限
                sqldr = sqlTable.NewRow()
                sqldr("account") = Me.nameid.Text
                sqldr("RoleID") = v_Role 'Me.Role.SelectedValue
                sqldr("LID") = Me.LIDValue.Value
                sqldr("Name") = Me.name.Text
                sqldr("IDNO") = TIMS.ChangeIDNO(Me.IDNO.Text)
                sqldr("Phone") = Me.telphone.Text
                sqldr("email") = Me.email.Text
                sqldr("OrgID") = Me.OrgID.Value
                sqldr("IsUsed") = If(Result.Checked, "Y", "N") '若同意則使用:Y'不同意則停用:N
                sqldr("Note") = Note.Text
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now()
                sqlTable.Rows.Add(sqldr)
            End If
            DbAccess.UpdateDataTable(sqlTable, da, objTrans)

            If Me.Result.Checked Then '同意則建立計劃
                sqlstr = "SELECT * FROM AUTH_ACCRWPLAN WHERE 1<>1"
                sqlTable = DbAccess.GetDataTable(sqlstr, da, objTrans)
                sqldr = sqlTable.NewRow()
                sqldr("Account") = Me.nameid.Text
                sqldr("PlanID") = Me.PlanIDValue.Value
                sqldr("RID") = Me.RIDValue.Value
                sqldr("CreateByAcc") = "Y"
                sqldr("ModifyAcct") = sm.UserInfo.UserID
                sqldr("ModifyDate") = Now()
                sqlTable.Rows.Add(sqldr)
                DbAccess.UpdateDataTable(sqlTable, da, objTrans)

                If flag_Auth_AccRWFun Then
                    '2005/1/5建立帳號的功能權限-   Start

                    Dim dt As DataTable = Nothing
                    Dim sql As String = ""
                    If Len(RIDValue.Value) = 1 Then
                        Select Case RIDValue.Value
                            Case "A" '署(局)的帳號功能
                                sql = "SELECT * FROM Auth_LevelRoleFun WHERE LID=0 and RoleID='" & v_Role & "'" 'Role.SelectedValue & "'"
                            Case Else '分署(中心)的帳號功能
                                sql = "SELECT * FROM Auth_LevelRoleFun WHERE LID=1 and RoleID='" & v_Role & "'" 'Role.SelectedValue & "'"
                        End Select
                    Else
                        '委訓的帳號功能
                        sql = "SELECT * FROM AUTH_LEVELFUN WHERE LID=2"
                    End If

                    dt = DbAccess.GetDataTable(sql, da, objTrans)
                    If dt.Rows.Count <> 0 Then
                        sql = "SELECT * FROM AUTH_ACCRWFUN WHERE 1<>1"
                        Dim dr1 As DataRow = Nothing
                        Dim dt1 As DataTable = Nothing
                        dt1 = DbAccess.GetDataTable(sql, da, objTrans)
                        For Each dr As DataRow In dt.Rows

                            '2005/6/21--Melody,委訓的帳號功能,不新增計畫審核,計畫變更審核
                            Dim flag_LID_01 As Boolean = False '非'署(局)的帳號功能  '分署(中心)的帳號功能
                            If Len(RIDValue.Value) = 1 Then flag_LID_01 = True '署(局)的帳號功能  '分署(中心)的帳號功能

                            If Get_LID(Me.RIDValue.Value) = "2" Then
                                If dr("FunID") <> "65" And dr("FunID") <> "67" Then
                                    dr1 = dt1.NewRow
                                    dt1.Rows.Add(dr1)
                                    dr1("Account") = nameid.Text
                                    dr1("FunID") = dr("FunID")
                                    '署(局)的帳號功能  '分署(中心)的帳號功能 1:'委訓的帳號功能
                                    dr1("Adds") = If(flag_LID_01, Convert.ToString(dr("Adds")), "1")
                                    dr1("Mod") = If(flag_LID_01, Convert.ToString(dr("Mod")), "1")
                                    dr1("Del") = If(flag_LID_01, Convert.ToString(dr("Del")), "1")
                                    dr1("Sech") = If(flag_LID_01, Convert.ToString(dr("Sech")), "1")
                                    dr1("ModifyAcct") = sm.UserInfo.UserID
                                    dr1("ModifyDate") = Now
                                End If
                            Else
                                dr1 = dt1.NewRow
                                dt1.Rows.Add(dr1)
                                dr1("Account") = nameid.Text
                                dr1("FunID") = dr("FunID")
                                '署(局)的帳號功能  '分署(中心)的帳號功能 1:'委訓的帳號功能
                                dr1("Adds") = If(flag_LID_01, Convert.ToString(dr("Adds")), "1")
                                dr1("Mod") = If(flag_LID_01, Convert.ToString(dr("Mod")), "1")
                                dr1("Del") = If(flag_LID_01, Convert.ToString(dr("Del")), "1")
                                dr1("Sech") = If(flag_LID_01, Convert.ToString(dr("Sech")), "1")
                                dr1("ModifyAcct") = sm.UserInfo.UserID
                                dr1("ModifyDate") = Now
                            End If

                        Next

                        DbAccess.UpdateDataTable(dt1, da, objTrans)
                    End If

                    '2005/1/5建立帳號的功能權限-   End

                End If

            End If


            DbAccess.CommitTrans(objTrans)
            Page.RegisterStartupScript("", "<script>alert('審核完成!繼續審核~');location.href='SYS_01_003.aspx?ID=" & Request("ID") & "';</script>")

        Catch ex As Exception
            DbAccess.RollbackTrans(objTrans)
            TIMS.CloseDbConn(objconn)
            Throw ex
        End Try
    End Sub

    Private Sub btu_save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btu_save.Click
        Dim rqfun As String = TIMS.ClearSQM(Request("fun"))
        If rqfun = "D" Then
            save_del()
        Else
            save_chk()
        End If
    End Sub

    Function Get_LID(ByVal RIDValue As String) As Integer
        Dim rst As Integer = 2
        If RIDValue = "A" Then rst = 0
        If RIDValue <> "A" AndAlso Len(RIDValue) = 1 Then rst = 1
        Return rst
    End Function

End Class
