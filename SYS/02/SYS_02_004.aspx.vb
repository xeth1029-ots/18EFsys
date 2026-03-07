'Imports System.Data.SqlClient
'Imports System.Data
'Imports Turbo
Partial Class SYS_02_004
    Inherits AuthBasePage

    Dim dt As DataTable
    Dim sql As String = ""
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

        If Not Page.IsPostBack Then

            sql = "" & vbCrLf
            sql += " select distinct a.Acct1,c.Name " & vbCrLf
            sql += " from Auth_AcctOrg a" & vbCrLf
            sql += " join Auth_AccRWPlan b on a.Acct1=b.Account" & vbCrLf
            sql += " join Auth_Account c on b.Account=c.Account" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and b.RID='" & sm.UserInfo.RID & "' " & vbCrLf
            sql += " and c.IsUsed='Y'" & vbCrLf
            sql += " order by a.Acct1,c.Name " & vbCrLf
            dt = DbAccess.GetDataTable(sql, objconn)

            'da.SelectCommand.Parameters.Clear()
            'objreader = DbAccess.GetReader(objstr, objconn)
            Me.BefAcc.DataSource = dt ' objreader
            Me.BefAcc.DataTextField = "Name"
            Me.BefAcc.DataValueField = "Acct1"
            Me.BefAcc.DataBind()
            Me.BefAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
            'objreader.Close()
            'objconn.Close()

            Me.AftAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
        End If
    End Sub

    Private Sub BefAcc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BefAcc.SelectedIndexChanged
        sql = "" & vbCrLf
        sql += " select distinct " & vbCrLf
        sql += "  c.Years+d.Name+e.PlanName+c.Seq PlanName " & vbCrLf
        sql += " ,a.PlanID" & vbCrLf
        sql += " from Auth_AcctOrg a" & vbCrLf
        sql += " join Auth_AccRWPlan b on a.Acct1=b.Account" & vbCrLf
        sql += " join ID_Plan c on a.PlanID=c.PlanID" & vbCrLf
        sql += " join ID_District d on c.DistID=d.DistID" & vbCrLf
        sql += " join Key_Plan e on c.TPlanID=e.TPLanid" & vbCrLf
        sql += " where b.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql += " ORDER BY 1" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Me.Plan_lst.DataSource = dt
        Me.Plan_lst.DataTextField = "PlanName"
        Me.Plan_lst.DataValueField = "PlanID"
        Me.Plan_lst.DataBind()

        'objreader.Close()
        'objconn.Close()
    End Sub

    Private Sub Plan_lst_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Plan_lst.SelectedIndexChanged

        sql = "" & vbCrLf
        sql += " select distinct a.RID,c.OrgName " & vbCrLf
        sql += " from Auth_AcctOrg a" & vbCrLf
        sql += " join Auth_Relship b on a.RID=b.RID" & vbCrLf
        sql += " join Org_OrgInfo c on b.OrgID=c.OrgID" & vbCrLf
        sql += " where a.PlanID='" & Me.Plan_lst.SelectedValue & "'" & vbCrLf
        sql += " and a.Acct1 = '" & Me.BefAcc.SelectedValue & "'" & vbCrLf
        sql += " order by a.RID,c.OrgName" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Me.Org_lis.DataSource = dt
        Me.Org_lis.DataTextField = "OrgName"
        Me.Org_lis.DataValueField = "RID"
        Me.Org_lis.DataBind()

    End Sub

    Private Sub Org_lis_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Org_lis.SelectedIndexChanged

        sql = "" & vbCrLf
        sql += " select distinct a.Account,a.Name " & vbCrLf
        sql += " from Auth_Account a" & vbCrLf
        sql += " join Auth_AccRWPlan b on a.Account=b.Account" & vbCrLf
        sql += " where b.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        sql += " and b.PlanID=" & Me.Plan_lst.SelectedValue
        sql += " and a.RoleID = 5" & vbCrLf
        sql += " and a.Account <> '" & Me.BefAcc.SelectedValue & "'" & vbCrLf
        sql += " and a.IsUsed='Y'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        Me.AftAcc.DataSource = dt
        Me.AftAcc.DataTextField = "Name"
        Me.AftAcc.DataValueField = "Account"
        Me.AftAcc.DataBind()
        Me.AftAcc.Items.Insert(0, New ListItem("==請選擇==", ""))

    End Sub

    Private Sub btu_sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btu_sub.Click
        For Each orgitem As ListItem In Me.Org_lis.Items
            If orgitem.Selected = True Then
                'Dim objstr As String
                'objstr = "update Auth_AcctOrg set Acct1 = '" & Me.AftAcc.SelectedValue & "' where " & _
                '" Acct1 = '" & Me.BefAcc.SelectedValue & "' and PlanID=" & Me.Plan_lst.SelectedValue & _
                '" and RID = '" & orgitem.Value & "'"

                Dim objstr As String = ""
                objstr = "" & vbCrLf
                objstr += " update Auth_AcctOrg" & vbCrLf
                objstr += " set Acct1 = '" & Me.AftAcc.SelectedValue & "'" & vbCrLf
                objstr += " where 1=1" & vbCrLf
                objstr += " and Acct1 = '" & Me.BefAcc.SelectedValue & "'" & vbCrLf
                objstr += " and PlanID=" & Me.Plan_lst.SelectedValue
                objstr += " and RID = '" & orgitem.Value & "'" & vbCrLf
                DbAccess.ExecuteNonQuery(objstr, objconn)
            End If
        Next

        TIMS.Utl_Redirect1(Me, "SYS_02_004.aspx")
    End Sub
End Class
