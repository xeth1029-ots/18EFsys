Partial Class SYS_02_003
    Inherits AuthBasePage

    Dim objconn As SqlConnection
    'Dim FunDr As DataRow

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


        Dim sql As String
        Dim dt As DataTable
        'Dim dr As DataRow
        If Not IsPostBack Then
            sql = "SELECT * FROM ID_Role WHERE RoleID IN (2,3,4,5)"
            dt = DbAccess.GetDataTable(sql, objconn)
            With RoleName
                .DataSource = dt
                .DataTextField = "Name"
                .DataValueField = "RoleID"
                .DataBind()
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End With

            Button1.Enabled = False
            BefAcct.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
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

    Private Sub PlanName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PlanName.SelectedIndexChanged
        If Not PlanName.SelectedItem Is Nothing Then
            'Dim dt As DataTable
            'Dim drResult() As DataRow
            ''Dim dr As DataRow
            'Dim i As Integer
            'Dim sql As String

            Me.ViewState("PlanID") = PlanName.SelectedValue

            If Me.ViewState("dt") Is Nothing Then
                Button1.Enabled = False
            Else
                Button1.Enabled = True
                NewAcct.Items.Clear()
                NewAcct.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                Dim dt As DataTable = Me.ViewState("dt")
                Dim drResult() As DataRow = dt.Select("Account<>'" & BefAcct.SelectedItem.Value & "' and PlanID='" & PlanName.SelectedValue & "'")
                For i As Integer = 0 To drResult.Length - 1
                    NewAcct.Items.Add(New ListItem(drResult(i)("Name"), drResult(i)("Account")))
                Next
            End If
        End If
    End Sub

    Private Sub RoleName_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles RoleName.SelectedIndexChanged
        If RoleName.SelectedIndex <> 0 Then
            BefAcct.Items.Clear()                   '清除角色的下拉選單
            Button1.Enabled = False                 '把儲存按鈕的功能影藏

            Dim sql As String = ""
            Dim dt As DataTable = Nothing
            sql = ""
            sql += "SELECT a.Name,a.Account,b.RID FROM "
            sql += "(SELECT * FROM Auth_Account WHERE RoleID='" & RoleName.SelectedItem.Value & "') a join "
            sql += "(SELECT DISTINCT Account,RID FROM Auth_AccRWPlan WHERE RID='" & sm.UserInfo.RID & "') b on a.Account=b.Account "
            dt = DbAccess.GetDataTable(sql, objconn)

            sql = "SELECT a.Name,a.Account,b.RID,b.PlanID FROM "
            sql += "(SELECT * FROM Auth_Account WHERE RoleID='" & RoleName.SelectedItem.Value & "') a join "
            sql += "(SELECT DISTINCT Account,RID,PlanID FROM Auth_AccRWPlan WHERE RID='" & sm.UserInfo.RID & "') b on a.Account=b.Account "
            Me.ViewState("dt") = DbAccess.GetDataTable(sql, objconn)

            If dt.Rows.Count = 0 Then
                BefAcct.Items.Add(New ListItem("查無資料", ""))
            Else
                With BefAcct
                    .DataSource = dt
                    .DataTextField = "Name"
                    .DataValueField = "Account"
                    .DataBind()
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End With
            End If
        End If
    End Sub

    Private Sub BefAcct_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BefAcct.SelectedIndexChanged
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim i As Integer = 0

        PlanName.Items.Clear()
        If BefAcct.SelectedIndex <> 0 Then
            sql = ""
            sql += "SELECT   a.Years, a.Seq, b.PlanName, a.PlanID, c.RID "
            sql += "FROM     Key_Plan b, "
            sql += "         ID_Plan a, "
            sql += "         (SELECT * FROM Auth_AccRWPlan WHERE Account='" & BefAcct.SelectedItem.Value & "') c "
            sql += "WHERE    b.TPlanID=a.TPlanID "
            sql += "AND      c.PlanID=a.PlanID"
            dt = DbAccess.GetDataTable(sql, objconn)

            If dt.Rows.Count = 0 Then
                Me.ViewState("PlanID") = ""
                Common.MessageBox(Me, "查無計畫資料!")
            ElseIf dt.Rows.Count = 1 Then                   '該帳號只有一個計畫時，直接show出同計畫的帳號
                dr = dt.Rows(0)
                Me.ViewState("PlanID") = dr("PlanID")
                Dim dtResult As DataTable
                Dim drResult() As DataRow

                dtResult = Me.ViewState("dt")
                drResult = dtResult.Select("Account<>'" & BefAcct.SelectedItem.Value & "' and PlanID='" & dr("PlanID") & "'")
                NewAcct.Items.Clear()
                NewAcct.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                If drResult.Length <> 0 Then
                    Button1.Enabled = True
                    For i = 0 To drResult.Length - 1
                        NewAcct.Items.Add(New ListItem(drResult(i)("Name"), drResult(i)("Account")))
                    Next
                Else
                    Button1.Enabled = False
                End If
            Else                                            '填入該帳號的計畫到RadioButtonList
                For i = 0 To dt.Rows.Count - 1
                    dr = dt.Rows(i)
                    PlanName.Items.Add(New ListItem(dr("Years") & dr("PlanName") & dr("Seq"), dr("PlanID")))
                Next
                NewAcct.Items.Clear()
                NewAcct.Items.Add(New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                Button1.Enabled = False
            End If
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        Dim dtResult As DataTable = Nothing
        Dim drResult1() As DataRow = Nothing
        Dim drResult2() As DataRow = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim i As Integer = 0
        Dim j As Integer = 0
        If NewAcct.SelectedIndex <> 0 Then
            sql = "SELECT * FROM Auth_AcctOrg WHERE PlanID='" & Me.ViewState("PlanID") & "'"
            dtResult = DbAccess.GetDataTable(sql, objconn)

            sql = "SELECT * FROM Auth_AcctOrg WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, objconn)

            '檢查要新增的帳號有沒有被授予的部份
            Select Case RoleName.SelectedValue
                Case 2
                    drResult1 = dtResult.Select("Acct4='" & BefAcct.SelectedValue & "'")
                    drResult2 = dtResult.Select("Acct4='" & NewAcct.SelectedValue & "'")
                Case 3
                    drResult1 = dtResult.Select("Acct3='" & BefAcct.SelectedValue & "'")
                    drResult2 = dtResult.Select("Acct3='" & NewAcct.SelectedValue & "'")
                Case 4
                    drResult1 = dtResult.Select("Acct2='" & BefAcct.SelectedValue & "'")
                    drResult2 = dtResult.Select("Acct2='" & NewAcct.SelectedValue & "'")
                Case 5
                    drResult1 = dtResult.Select("Acct1='" & BefAcct.SelectedValue & "'")
                    drResult2 = dtResult.Select("Acct1='" & NewAcct.SelectedValue & "'")
            End Select

            For i = 0 To drResult1.Length - 1
                Dim Insert_Flag As Boolean = True
                'Insert_Flag = True
                For j = 0 To drResult2.Length - 1                           '檢查是否有相同的帳號組織資料
                    If RoleName.SelectedValue = 5 Then
                        If drResult1(i)("Acct3").ToString = drResult2(j)("Acct3").ToString Then
                            Insert_Flag = False
                        End If
                    End If
                    If RoleName.SelectedValue = 5 Or RoleName.SelectedValue = 4 Then
                        If drResult1(i)("Acct2").ToString = drResult2(j)("Acct2").ToString Then
                            Insert_Flag = False
                        End If
                    End If
                    If RoleName.SelectedValue = 5 Or RoleName.SelectedValue = 4 Or RoleName.SelectedValue = 3 Then
                        If drResult1(i)("Acct1") = drResult2(j)("Acct1") Then
                            Insert_Flag = False
                        End If
                    End If

                    If drResult1(i)("RID") = drResult2(j)("RID") Then
                        Insert_Flag = False
                    End If
                Next

                If Insert_Flag = True Then
                    dr = dt.NewRow
                    dt.Rows.Add(dr)

                    dr("PlanID") = drResult1(i)("PlanID")
                    dr("RID") = drResult1(i)("RID")
                    Select Case RoleName.SelectedValue
                        Case 2
                            dr("Acct4") = NewAcct.SelectedValue
                        Case 3
                            dr("Acct3") = NewAcct.SelectedValue
                            dr("Acct4") = drResult1(i)("Acct4")
                        Case 4
                            dr("Acct2") = If(String.IsNullOrEmpty(NewAcct.SelectedValue), " ", NewAcct.SelectedValue)
                            dr("Acct3") = drResult1(i)("Acct3")
                            dr("Acct4") = drResult1(i)("Acct4")
                        Case 5
                            dr("Acct1") = If(String.IsNullOrEmpty(NewAcct.SelectedValue), " ", NewAcct.SelectedValue)
                            dr("Acct2") = If(Convert.IsDBNull(drResult1(i)("Acct2")), " ", drResult1(i)("Acct2"))
                            dr("Acct3") = drResult1(i)("Acct3")
                            dr("Acct4") = drResult1(i)("Acct4")
                    End Select
                    dr("ModifyAcct") = sm.UserInfo.UserID
                    dr("ModifyDate") = Now
                End If
            Next

            Try
                DbAccess.UpdateDataTable(dt, da)
                Common.MessageBox(Me, "授予成功")
            Catch ex As Exception
                'Common.RespWrite(Me, ex)
                Common.MessageBox(Me, "授予失敗")
            End Try
        End If
    End Sub
End Class
