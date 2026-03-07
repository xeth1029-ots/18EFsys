Partial Class SYS_02_002_
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

        Dim objreader As SqlDataReader
        Dim objstr As String
        Dim objarray As System.Array

        If Not Page.IsPostBack Then
            objstr = "select * from ID_Role where RoleID<>0 and RoleID<>1 and RoleID<>99 order by RoleID"
            objreader = DbAccess.GetReader(objstr, objconn)
            Me.BefLev.DataSource = objreader
            Me.BefLev.DataTextField = "Name"
            Me.BefLev.DataValueField = "RoleID"
            Me.BefLev.DataBind()
            Me.BefLev.Items.Insert(0, New ListItem("==請選擇==", ""))
            objarray = System.Array.CreateInstance(GetType(System.Web.UI.WebControls.ListItem), Me.BefLev.Items.Count)
            Me.BefLev.Items.CopyTo(objarray, 0)
            Me.AftLev.Items.AddRange(objarray)
            objreader.Close()

            Me.BefAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
            Me.AftAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
        End If
    End Sub

    Private Sub BefLev_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BefLev.SelectedIndexChanged
        Me.AftLev.SelectedValue = Me.BefLev.SelectedValue
        Call BefAcc_Data()
        Me.Plan_lst.Items.Clear()
    End Sub

    Private Sub AftLev_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AftLev.SelectedIndexChanged
        Me.BefLev.SelectedValue = Me.AftLev.SelectedValue
        Call BefAcc_Data()
        Me.Plan_lst.Items.Clear()
    End Sub

    Private Sub BefAcc_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BefAcc.SelectedIndexChanged
        Dim objreader As SqlDataReader
        Dim AcctWho As String = "Acct1"
        Select Case Me.BefLev.SelectedValue
            Case "5"
                AcctWho = "Acct1"
            Case "4"
                AcctWho = "Acct2"
            Case "3"
                AcctWho = "Acct3"
            Case "2"
                AcctWho = "Acct4"
        End Select

        Dim objstr As String = ""
        objstr = "" & vbCrLf
        objstr += " select distinct " & vbCrLf
        objstr += " b.Years+c.Name+d.PlanName+b.seq PlanName " & vbCrLf
        objstr += " ,b.PlanID"
        objstr += " from Auth_AcctOrg a" & vbCrLf
        objstr += " join Auth_AccRWPlan e on a." & AcctWho & "=e.Account" & vbCrLf
        objstr += " join ID_Plan b on e.PlanID=b.PlanID" & vbCrLf
        objstr += " join ID_District c on b.DistID=c.DistID" & vbCrLf
        objstr += " join Key_Plan d on b.TPlanID=d.TPlanID" & vbCrLf
        objstr += " where e.Account = '" & Me.BefAcc.SelectedValue & "'" & vbCrLf
        objstr += " ORDER BY 1" & vbCrLf
        objreader = DbAccess.GetReader(objstr, objconn)
        Me.Plan_lst.DataSource = objreader
        Me.Plan_lst.DataTextField = "PlanName"
        Me.Plan_lst.DataValueField = "PlanID"
        Me.Plan_lst.DataBind()
        objreader.Close()

    End Sub

    Private Sub BefAcc_Data()
        Dim objreader As SqlDataReader
        Dim AcctWho As String = "Acct1"
        Dim sRoleID As String = ""
        sRoleID = "'" & Me.BefLev.SelectedValue & "'"

        Select Case Me.BefLev.SelectedValue
            Case "5"
                AcctWho = "Acct1"
            Case "4"
                AcctWho = "Acct2"
            Case "3"
                AcctWho = "Acct3"
            Case "2"
                AcctWho = "Acct4"
            Case Else
                AcctWho = "Acct4"
                sRoleID = "-1"
        End Select

        Dim objstr As String = ""
        objstr = "" & vbCrLf
        objstr += " select distinct c.Account,c.Name " & vbCrLf
        objstr += " from Auth_AcctOrg a" & vbCrLf
        objstr += " join Auth_AccRWPlan b on a." & AcctWho & "=b.Account and a.PlanID=b.PlanID" & vbCrLf
        objstr += " join Auth_Account c on b.Account=c.Account" & vbCrLf
        objstr += " where b.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        objstr += " and c.RoleID = " & sRoleID & vbCrLf
        objstr += " and c.IsUsed='Y'" & vbCrLf
        objstr += " ORDER BY 1" & vbCrLf
        objreader = DbAccess.GetReader(objstr, objconn)
        Me.BefAcc.DataSource = objreader
        Me.BefAcc.DataTextField = "Name"
        Me.BefAcc.DataValueField = "Account"
        Me.BefAcc.DataBind()
        Me.BefAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
        objreader.Close()

        Me.AftAcc.Items.Clear()
        Me.AftAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
    End Sub

    Private Sub btu_sub_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btu_sub.Click
        Dim objAdapter As SqlDataAdapter = Nothing
        Dim objTable As DataTable = Nothing
        'Dim objdr As DataRow=nothing
        Dim objstr As String = ""
        'Dim CID As Integer


        Select Case Me.BefLev.SelectedValue
            Case "5"
                objstr = "select * from Auth_AcctOrg where PlanID = " & Me.Plan_lst.SelectedValue & " and acct1='" & Me.BefAcc.SelectedValue & "'"
                objTable = DbAccess.GetDataTable(objstr, objAdapter, objconn)
                For Each objdr As DataRow In objTable.Rows
                    objdr("Acct1") = Me.AftAcc.SelectedValue
                    DbAccess.UpdateDataTable(objTable, objAdapter)
                Next
            Case "4"
                objstr = "select * from Auth_AcctOrg where PlanID = " & Me.Plan_lst.SelectedValue & " and acct2='" & Me.BefAcc.SelectedValue & "'"
                objTable = DbAccess.GetDataTable(objstr, objAdapter, objconn)
                For Each objdr As DataRow In objTable.Rows
                    objdr("Acct2") = Me.AftAcc.SelectedValue
                    DbAccess.UpdateDataTable(objTable, objAdapter)
                Next
            Case "3"
                objstr = "select * from Auth_AcctOrg where PlanID = " & Me.Plan_lst.SelectedValue & " and acct3='" & Me.BefAcc.SelectedValue & "'"
                objTable = DbAccess.GetDataTable(objstr, objAdapter, objconn)
                For Each objdr As DataRow In objTable.Rows
                    objdr("Acct3") = Me.AftAcc.SelectedValue
                    DbAccess.UpdateDataTable(objTable, objAdapter)
                Next
            Case "2"
                objstr = "select * from Auth_AcctOrg where PlanID = " & Me.Plan_lst.SelectedValue & " and acct4='" & Me.BefAcc.SelectedValue & "'"
                objTable = DbAccess.GetDataTable(objstr, objAdapter, objconn)
                For Each objdr As DataRow In objTable.Rows
                    objdr("Acct4") = Me.AftAcc.SelectedValue
                    DbAccess.UpdateDataTable(objTable, objAdapter)
                Next
        End Select

        Dim url1 As String = "SYS_02_002.aspx?ID=" & Request("ID")
        TIMS.Utl_Redirect(Me, objconn, url1)
        'Response.Redirect("SYS_02_002.aspx")
    End Sub

    Private Sub Plan_lst_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Plan_lst.SelectedIndexChanged
        Dim objreader As SqlDataReader
        Dim objstr As String = ""
        objstr = "" & vbCrLf
        objstr += " select a.Account,b.Name " & vbCrLf
        objstr += " from Auth_AccRWPlan a " & vbCrLf
        objstr += " join Auth_Account b on a.Account=b.Account " & vbCrLf
        objstr += " where a.PlanID=" & Me.Plan_lst.SelectedValue & vbCrLf
        objstr += " and a.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        objstr += " and b.RoleID=" & Me.BefLev.SelectedValue & vbCrLf
        objstr += " and a.Account not in ('" & Me.BefAcc.SelectedValue & "') " & vbCrLf
        objstr += " and b.IsUsed='Y'" & vbCrLf
        objstr += " ORDER BY 1" & vbCrLf
        objreader = DbAccess.GetReader(objstr, objconn)
        Me.AftAcc.DataSource = objreader
        Me.AftAcc.DataTextField = "Name"
        Me.AftAcc.DataValueField = "Account"
        Me.AftAcc.DataBind()
        Me.AftAcc.Items.Insert(0, New ListItem("==請選擇==", ""))
        objreader.Close()

    End Sub
End Class
