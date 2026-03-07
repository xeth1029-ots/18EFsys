Public Class title
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    'List.Item( Oracle要大寫
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.Stop1(Me, objconn)
        If Not TIMS.OpenDbConn(objconn) Then Exit Sub

        '職業訓練中心 > 系統測試員 > 測試員
        Me.labID.Text = ""
        '2011北區職業訓練中心自辦職前訓練001
        Me.labPlan.Text = ""
        If sm.UserInfo.UserID Is Nothing Then Exit Sub
        If sm.UserInfo.PlanID Is Nothing Then Exit Sub

        Dim sql As String = ""
        Dim dt As New DataTable
        If Convert.ToString(sm.UserInfo.PlanID) <> "0" Then
            sql = "" & vbCrLf
            sql += " select " & vbCrLf
            sql += " b.name UserName" & vbCrLf
            sql += " ,d.name UserRole" & vbCrLf
            sql += " ,c.Years+e.Name+f.PlanName+c.seq+c.SubTitle UserPlan" & vbCrLf
            sql += " from Auth_AccRWPlan a" & vbCrLf
            sql += " join Auth_Account b on a.Account=b.Account" & vbCrLf
            sql += " join ID_Plan c on a.PlanID=c.PlanID" & vbCrLf
            sql += " join ID_Role d on b.RoleID=d.RoleID" & vbCrLf
            sql += " join ID_District e on c.DistID=e.DistID" & vbCrLf
            sql += " join Key_Plan f on c.TPlanID=f.TPlanID" & vbCrLf
            sql += " and a.Account=@Account" & vbCrLf
            sql += " and a.PlanID =@PlanID" & vbCrLf
            Dim cmd As New SqlCommand(sql, objconn)
            With cmd
                .Parameters.Clear()
                .Parameters.Add("Account", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("PlanID", SqlDbType.Int).Value = sm.UserInfo.PlanID
                dt.Load(.ExecuteReader())
            End With
        Else
            sql = "" & vbCrLf
            sql += " select b.name UserName" & vbCrLf
            sql += " ,c.name UserRole " & vbCrLf
            sql += " ,null UserPlan " & vbCrLf
            sql += " from Auth_AccRWPlan a" & vbCrLf
            sql += " join Auth_Account b on a.Account=b.Account" & vbCrLf
            sql += " join ID_Role c on b.RoleID=c.RoleID" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and a.Account=@Account" & vbCrLf
            sql += " and a.PlanID =@PlanID" & vbCrLf
            Dim cmd As New SqlCommand(sql, objconn)
            With cmd
                .Parameters.Clear()
                .Parameters.Add("Account", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                .Parameters.Add("PlanID", SqlDbType.Int).Value = sm.UserInfo.PlanID
                dt.Load(.ExecuteReader())
            End With
        End If

        '職業訓練中心 > 系統測試員 > 測試員
        Me.labID.Text = ""
        '2011北區職業訓練中心自辦職前訓練001
        Me.labPlan.Text = ""
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            Dim sOrgName As String = sm.UserInfo.OrgName
            Dim sUserRole As String = ""
            Dim sUserName As String = ""
            sUserRole = dr("UserRole")
            sUserName = dr("UserName")
            Me.labID.Text = sOrgName & " > " & sUserRole & " > " & sUserName
            If Convert.ToString(dr("UserPlan")) <> "" Then
                Me.labPlan.Text = Convert.ToString(dr("UserPlan"))
            End If
        End If

    End Sub

End Class