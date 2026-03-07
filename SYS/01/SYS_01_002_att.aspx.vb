Partial Class SYS_01_002_att
    Inherits AuthBasePage

    'Dim objconn As SqlConnection
    '更動的TABLE: AUTH_ACCTORG
    '一切都是為了 RID

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        msg2.Text = ""
        'rqAN=TIMS.ClearSQM(Request("AN"))
        Dim i_RoleId As Integer = 999
        If Request("AN") = "" Then '採直接呼叫，確認使用者層級
            Button1.Visible = False
            i_RoleId = TIMS.Get_RoleID(sm.UserInfo.UserID, objconn)
            Select Case i_RoleId
                Case 4, 5
                    Me.ViewState("RoleID") = i_RoleId 'strRoleId
                    Me.ViewState("AN") = sm.UserInfo.UserID
                    Me.ViewState("RID") = sm.UserInfo.RID
                    Me.ViewState("PID") = sm.UserInfo.PlanID
                Case Else
                    'i_RoleId 0:超級使用者,1:系統管理者,2:一級以上,3:一級,4:二級,5:承辦人,99:一般使用者
                    Common.MessageBox(Me, "限定權限 二級 與 承辦人 可做委訓單位歸屬!!")
                    Exit Sub
            End Select
        Else
            '上層呼叫，塞入必要之參數
            Button1.Visible = True
            Me.ViewState("RoleID") = Val(TIMS.ClearSQM(Request("RoleID")))
            Me.ViewState("AN") = TIMS.ClearSQM(Request("AN"))
            Me.ViewState("RID") = TIMS.ClearSQM(Request("RID"))
            Me.ViewState("PID") = TIMS.ClearSQM(Request("PID"))
        End If

        If Not Me.IsPostBack Then
            Create1()
        End If
    End Sub

    Sub Create1()
        msg2.Text = "沒有可以歸屬的單位"
        CheckBox1.Visible = False
        but_add.Enabled = False

        Dim dt As DataTable = TIMS.GET_OrgListdt(Me.ViewState("PID"), Me.ViewState("RID"), objconn)
        If dt.Rows.Count > 0 Then
            msg2.Text = ""
            'CheckBox1.Visible = True
            but_add.Enabled = True

            Me.OrgList.DataSource = dt
            Me.OrgList.DataTextField = "OrgName"
            Me.OrgList.DataValueField = "RID"
            Me.OrgList.DataBind()

            '顯示-查詢-使用者 已設定的資訊
            UTL_ShowORGLIST1(Val(ViewState("RoleID")), ViewState("PID"), ViewState("AN"), OrgList, objconn)

            CheckBox1.Visible = True 'SelectAll
            CheckBox1.Attributes("onclick") = "SelectAll(this.checked," & OrgList.Items.Count & ");"
        End If
    End Sub

    ''' <summary>
    ''' 查詢目前使用者 已設定的資訊
    ''' </summary>
    ''' <param name="iROLEID"></param>
    ''' <param name="PLANID"></param>
    ''' <param name="ACCT"></param>
    ''' <param name="OrgList"></param>
    ''' <param name="oConn"></param>
    Public Shared Sub UTL_ShowORGLIST1(ByVal iROLEID As Integer, ByVal PLANID As String, ByVal ACCT As String, ByRef OrgList As CheckBoxList, ByRef oConn As SqlConnection)
        '0:超級使用者,1:系統管理者,2:一級以上,3:一級,4:二級,5:承辦人,99:一般使用者
        'sql = "SELECT a.RID FROM "
        If ACCT = "" Then Return
        Dim sql As String = ""
        If iROLEID = 4 Then
            'sql = "SELECT RID,Acct1,Acct2,Acct3,Acct4 F..ROM Auth_AcctOrg WHERE PlanID=" & PLANID & " and Acct2='" & ACCT & "'"
            sql = "SELECT DISTINCT RID FROM AUTH_ACCTORG WHERE PlanID=" & PLANID & " and Acct2='" & ACCT & "'"
        Else
            'sql = "SELECT RID,Acct1,Acct2,Acct3,Acct4 FROM Auth_AcctOrg WHERE PlanID=" & PLANID & " and Acct1='" & ACCT & "'"
            sql = "SELECT DISTINCT RID FROM AUTH_ACCTORG WHERE PlanID=" & PLANID & " and Acct1='" & ACCT & "'"
        End If
        'sql += "Join (SELECT * FROM Auth_Relship where PlanID = " & Me.Request("PID") & ") b on a.RID=b.RID "
        'sql += "Join (SELECT * FROM Org_OrgInfo) c ON b.OrgID=c.OrgID "
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn)
        If dt.Rows.Count = 0 Then Return

        For Each item As ListItem In OrgList.Items
            For Each dr As DataRow In dt.Rows
                If item.Value = dr("RID").ToString() Then
                    item.Selected = True
                End If
            Next
        Next
    End Sub

    Private Sub DataList1_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataListItemEventArgs)
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim MyLabel As Label = e.Item.FindControl("Label1")
                Dim drv As DataRowView = e.Item.DataItem

                MyLabel.Text = drv("OrgName").ToString
        End Select

    End Sub

    ''' <summary>
    ''' 回上頁
    ''' </summary>
    Sub goback1()
        'Response.Redirect("SYS_01_002.aspx?ID=" & Request("ID"))
        Dim rqID As String = TIMS.ClearSQM(Request("ID"))
        Dim rqAN As String = TIMS.ClearSQM(Request("AN"))
        Dim rqRID As String = TIMS.ClearSQM(Request("RID"))
        Dim rqYears As String = TIMS.ClearSQM(Request("Years"))
        If rqAN <> "" Then
            '若為上層呼叫，可儲存後返回上層
            Common.AddClientScript(Page, "location.href='SYS_01_002.aspx?ID=" & rqID & "&AN=" & rqAN & "&RID=" & rqRID & "&Years=" & rqYears & "';")
        End If
    End Sub

    '回上頁
    Private Sub Button1_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.ServerClick
        '回上頁
        goback1()
    End Sub

    '儲存
    Private Sub But_add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles but_add.Click
        Dim errMsg As String = ""
        Dim flag_SAVEOK As Boolean = True
        Dim v_RoleID As String = If(ViewState("RoleID"), "")
        Dim v_PID As String = If(ViewState("PID"), "")
        Dim v_AN As String = If(ViewState("AN"), "")
        v_RoleID = TIMS.ClearSQM(v_RoleID)
        v_PID = TIMS.ClearSQM(v_PID)
        v_AN = TIMS.ClearSQM(v_AN)
        If (v_RoleID = "" OrElse v_PID = "" OrElse v_AN = "") Then
            errMsg = "儲存參數有誤，請重新查詢設定！"
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If
        flag_SAVEOK = TIMS.Update_AUTH_ACCTORG(Me, v_RoleID, CInt(v_PID), v_AN, OrgList, errMsg, objconn)
        If flag_SAVEOK Then
            Common.AddClientScript(Page, "alert('歸屬成功!!!');")
        Else
            '失敗
            Common.MessageBox(Me, errMsg) '
            Exit Sub
        End If

        '回上頁
        goback1()
    End Sub

End Class
