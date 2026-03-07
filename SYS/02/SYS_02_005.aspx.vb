Partial Class SYS_02_005
    Inherits AuthBasePage

    'Dim FunDr As DataRow
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Dim sql As String
            Dim dt As DataTable

            sql = "SELECT * FROM ID_Role WHERE RoleID IN (2,3,4,5) ORDER BY RoleID"
            dt = DbAccess.GetDataTable(sql, objconn)
            With Role
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "RoleID"
                .DataBind()
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End With


            sql = "" & vbCrLf
            sql += " SELECT a.Account" & vbCrLf
            sql += " ,a.RoleID" & vbCrLf
            sql += " ,a.Name+'('+c.Name+')' name " & vbCrLf
            sql += " FROM Auth_Account a " & vbCrLf
            sql += " join Auth_AccRWPlan b on a.Account=b.Account" & vbCrLf
            sql += " JOIN ID_Role c on a.RoleID=c.RoleID and  c.RoleID IN (2,3,4,5)" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and b.RID='" & sm.UserInfo.RID & "' " & vbCrLf
            sql += " and b.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql += " ORDER BY 1" & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)

            With BefAcct
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "Account"
                .DataBind()
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End With
            NewAcct.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
        End If

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
        '            If FunDr("Adds") = "1" Then
        '                Button1.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End        
    End Sub

    Private Sub Role_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Role.SelectedIndexChanged, BefAcct.SelectedIndexChanged
        If Role.SelectedIndex <> 0 And BefAcct.SelectedIndex <> 0 Then
            Dim sql As String
            Dim dt As DataTable
            Dim RoleID As Integer
            Select Case Role.SelectedValue
                Case 2
                    RoleID = 4
                Case 3
                    RoleID = 3
                Case 4
                    RoleID = 2
                Case 5
                    RoleID = 1
            End Select
            If RoleID.ToString <> "" Then
                sql = "SELECT   a.Account, a.Name "
                sql += "FROM     (SELECT Distinct Acct" & RoleID.ToString & " FROM Auth_AcctOrg) b, "
                sql += "         (SELECT * FROM Auth_Account WHERE Account<>'" & BefAcct.SelectedValue & "') a "
                sql += "WHERE    b.Acct" & RoleID.ToString & "=a.Account"

                dt = DbAccess.GetDataTable(sql, objconn)

                'NewAcct.Items.Clear()
                With NewAcct
                    .DataSource = dt
                    .DataTextField = "Name"
                    .DataValueField = "Account"
                    .DataBind()
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End With
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        If Role.SelectedIndex <> 0 And BefAcct.SelectedIndex <> 0 And NewAcct.SelectedIndex <> 0 Then
            Dim sql As String
            Dim dr As DataRow
            'Dim dt As DataTable
            Dim msg As String = ""

            '先抓出舊帳號的角色
            Dim OldRole As String
            sql = "SELECT RoleID FROM Auth_Account WHERE Account='" & BefAcct.SelectedValue & "'"
            dr = DbAccess.GetOneRow(sql, objconn)
            If dr Is Nothing Then
                msg += "查無舊帳號資料\n"
            Else
                Dim objTrans As SqlTransaction
                objTrans = DbAccess.BeginTrans(objconn)

                OldRole = dr("RoleID")
                sql = "Update Auth_Account SET RoleID='" & Role.SelectedValue & "' WHERE Account='" & BefAcct.SelectedValue & "'"
                Try
                    DbAccess.ExecuteNonQuery(sql, objTrans)
                Catch ex As Exception
                    msg += "舊帳號更新角色失敗\n"
                End Try

                sql = "Update Auth_Account SET RoleID='" & OldRole & "' WHERE Account='" & NewAcct.SelectedValue & "'"
                Try
                    DbAccess.ExecuteNonQuery(sql, objTrans)
                Catch ex As Exception
                    msg += "交換對象帳號更新角色失敗\n"
                End Try

                Dim RoleID As Integer
                Select Case Role.SelectedValue
                    Case 2
                        RoleID = 4
                    Case 3
                        RoleID = 3
                    Case 4
                        RoleID = 2
                    Case 5
                        RoleID = 1
                End Select

                sql = "UPDATE Auth_AcctOrg SET Acct" & RoleID & "='" & BefAcct.SelectedValue & "' WHERE Acct" & RoleID & "='" & NewAcct.SelectedValue & "'"
                Try
                    DbAccess.ExecuteNonQuery(sql, objTrans)
                Catch ex As Exception
                    msg += "舊帳號更新帳號組織失敗\n"
                End Try

                Select Case OldRole
                    Case 2
                        RoleID = 4
                    Case 3
                        RoleID = 3
                    Case 4
                        RoleID = 2
                    Case 5
                        RoleID = 1
                End Select
                sql = "UPDATE Auth_AcctOrg SET Acct" & RoleID & "='" & NewAcct.SelectedValue & "' WHERE Acct" & RoleID & "='" & BefAcct.SelectedValue & "'"
                Try
                    DbAccess.ExecuteNonQuery(sql, objTrans)
                Catch ex As Exception
                    msg += "交換更新帳號組織失敗\n"
                End Try

                Try
                    DbAccess.CommitTrans(objTrans)
                    msg += "交換成功"
                Catch ex As Exception
                    DbAccess.RollbackTrans(objTrans)
                    msg += "交換失敗"
                End Try
            End If

            Page.RegisterStartupScript("showmsg", "<script>alert('" & msg & "');location.href='SYS_02_005.aspx?ID=" & Request("ID") & "';</script>")
        End If
    End Sub
End Class
