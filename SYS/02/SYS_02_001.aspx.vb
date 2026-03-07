Partial Class SYS_02_001
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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'Dim objreader As SqlDataReader
        'Dim objstr As String
        If Not Page.IsPostBack Then
            Call CreateItems1()
        End If

    End Sub

    Sub CreateItems1()
        If Convert.ToString(sm.UserInfo.LID) = "" Then Exit Sub
        '.Items.Add(New ListItem("局", "0"))
        '.Items.Add(New ListItem("中心", "1"))
        'Me.ShowLev.DataBind()
        With ShowLev
            .Items.Clear()
            .Items.Add(New ListItem("署", "0"))
            .Items.Add(New ListItem("分署", "1"))
        End With
        Common.SetListItem(ShowLev, sm.UserInfo.LID)

        'Dim objreader As SqlDataReader
        Dim objstr As String = ""
        objstr = "SELECT RoleID,Name FROM ID_ROLE WHERE ROLEID NOT IN (0,1,5,99) ORDER BY ROLEID"
        Dim dt As DataTable = DbAccess.GetDataTable(objstr, objconn)
        With ShowRole
            .DataSource = dt
            .DataTextField = "Name"
            .DataValueField = "RoleID"
            .DataBind()
        End With

        'objreader = DbAccess.GetReader(objstr, objconn)
        'Me.ShowRole.DataSource = objreader
        'Me.ShowRole.DataTextField = "Name"
        'Me.ShowRole.DataValueField = "RoleID"
        'Me.ShowRole.DataBind()
        'Me.ShowRole.Items.Insert(0, New ListItem("==請選擇==", ""))
        'objreader.Close()
        'objconn.Close()
    End Sub

    Private Sub ShowRole_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowRole.SelectedIndexChanged
        If Convert.ToString(sm.UserInfo.RID) = "" Then Exit Sub
        If Convert.ToString(sm.UserInfo.PlanID) = "" Then Exit Sub

        If ShowRole.SelectedValue <> "" Then
            Dim objreader As SqlDataReader
            Dim objstr As String = ""
            objstr = "" & vbCrLf
            objstr += " select b.Name,a.Account " & vbCrLf
            objstr += " from Auth_AccRwPlan a " & vbCrLf
            objstr += " join Auth_Account b on a.account=b.account and b.IsUsed='Y'" & vbCrLf
            objstr += " where b.RoleID = " & Me.ShowRole.SelectedValue & vbCrLf
            objstr += " and a.RID = '" & sm.UserInfo.RID & "' " & vbCrLf
            objstr += " and a.PlanID = " & sm.UserInfo.PlanID & vbCrLf
            objreader = DbAccess.GetReader(objstr, objconn)
            With ShowAcc
                .DataSource = objreader
                .DataTextField = "Name"
                .DataValueField = "Account"
                .DataBind()
                .Items.Insert(0, New ListItem("==請選擇==", ""))
            End With
            objreader.Close()
            Me.acclist.Items.Clear()
        End If
    End Sub

    Private Sub ShowAcc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ShowAcc.SelectedIndexChanged
        Dim objreader As SqlDataReader
        Dim objstr As String = ""
        objstr = "" & vbCrLf
        objstr += " select distinct c.Name,c.Account,e.PlanName " & vbCrLf
        objstr += " from Auth_AcctOrg a" & vbCrLf
        objstr += " join Auth_AccRWPlan b on a.PlanID=b.PlanID" & vbCrLf
        objstr += " join Auth_Account c on b.Account=c.Account" & vbCrLf
        objstr += " join ID_Plan d on d.PlanID=a.PlanID and d.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        objstr += " join Key_Plan e on e.TPlanID=d.TPlanID" & vbCrLf
        objstr += " where 1=1" & vbCrLf
        objstr += " and a.PlanID=" & sm.UserInfo.PlanID & vbCrLf
        objstr += " and b.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        Select Case Me.ShowRole.SelectedValue
            Case "4"
                objstr += " and a.Acct2 is null and a.Acct1=b.Account"
            Case "3"
                objstr += " and a.Acct3 is null and a.Acct2=b.Account"
            Case "2"
                objstr += " and a.Acct4 is null and a.Acct3=b.Account"
            Case Else
                Common.MessageBox(Me, "角色資料有誤，請重新選擇查詢!!")
                Exit Sub
        End Select

        objreader = DbAccess.GetReader(objstr, objconn)
        acclist.Items.Clear()
        While objreader.Read
            acclist.Items.Add(New ListItem(objreader("Name") & "(" & objreader("PlanName") & ")", objreader("Account")))
        End While
        objreader.Close()

        'Me.acclist.DataSource = objreader
        'Me.acclist.DataTextField = "Name"
        'Me.acclist.DataValueField = "Account"
        'Me.acclist.DataBind()
        'objconn.Close()
    End Sub

    Private Sub but_sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_sub.Click
        'Dim sqlAdapter As SqlDataAdapter
        'Dim sqlTable As DataTable
        'Dim sqldr_account As DataRow
        'Dim sqlstr As String
        'Dim i As Integer
        If Me.ShowRole.SelectedValue = "" Then
            Common.MessageBox(Me, "角色資料有誤，請重新選擇查詢!!")
            Exit Sub
        End If

        For i As Integer = 0 To Me.acclist.Items.Count - 1
            If Me.acclist.Items(i).Selected AndAlso Me.acclist.Items(i).Value <> "" Then
                'sqlstr = "select * from Auth_AcctOrg where PlanID=" & sm.UserInfo.PlanID & " and acct1='" & Me.acclist.Items(i).Value & "'"
                'Mod by Jack 04/11/19
                Dim sqlstr As String = ""
                sqlstr = "select * from Auth_AcctOrg where PlanID=" & sm.UserInfo.PlanID
                Select Case Me.ShowRole.SelectedValue
                    Case "4"
                        sqlstr += " and acct1='" & Me.acclist.Items(i).Value & "'"
                    Case "3"
                        sqlstr += " and acct2='" & Me.acclist.Items(i).Value & "'"
                    Case "2"
                        sqlstr += " and acct3='" & Me.acclist.Items(i).Value & "'"
                    Case Else
                        Common.MessageBox(Me, "角色資料有誤，請重新選擇查詢!!")
                        Exit Sub
                End Select
                'End Mod

                Dim sqlAdapter As SqlDataAdapter = Nothing
                Dim sqlTable As DataTable = Nothing
                sqlTable = DbAccess.GetDataTable(sqlstr, sqlAdapter, objconn)
                If sqlTable.Rows.Count > 0 Then
                    For Each sqldr_account As DataRow In sqlTable.Rows
                        Select Case Me.ShowRole.SelectedValue
                            Case "4"
                                sqldr_account("Acct2") = Me.ShowAcc.SelectedValue
                            Case "3"
                                sqldr_account("Acct3") = Me.ShowAcc.SelectedValue
                            Case "2"
                                sqldr_account("Acct4") = Me.ShowAcc.SelectedValue
                            Case Else
                                Common.MessageBox(Me, "角色資料有誤，請重新選擇查詢!!")
                                Exit Sub
                        End Select
                    Next
                    DbAccess.UpdateDataTable(sqlTable, sqlAdapter)
                End If

            End If
        Next

        TIMS.Utl_Redirect1(Me, "SYS_02_001.aspx") '重新載入
    End Sub
End Class
