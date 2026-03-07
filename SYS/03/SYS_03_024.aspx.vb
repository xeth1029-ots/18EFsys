Partial Class SYS_03_024
    Inherits AuthBasePage

    'Dim conn As SqlConnection = DbAccess.GetConnection()
    Dim sql As String = ""

    '#Region "Sub"
    '#End Region

    '代入DropDownList資料
    Private Sub Set_DropDownList(ByVal strFlag As String, ByVal obj As DropDownList, _
                                 ByVal textField As String, ByVal valueField As String)
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet


        Select Case strFlag
            Case "PlanYears"
                sql = "select distinct Years from ID_Plan where ISNULL(Years,' ')<>' ' order by Years"
            Case "DistID"
                sql = "select DistID,Name from ID_District order by DistID Asc "
            Case "PlanID"
                'Dim sql As String = ""
                sql = "" & vbCrLf
                sql &= " select distinct a.PlanID,a.Years+c.Name+b.PlanName+a.Seq" & vbCrLf
                sql &= " +case when ISNULL(a.SubTitle,'')='' then '' else '('+CONVERT(varchar, a.SubTitle)+')' end PlanName" & vbCrLf
                sql &= " from ID_Plan a" & vbCrLf
                sql &= " join Key_Plan b on b.TPlanID=a.TPlanID" & vbCrLf
                sql &= " join ID_District c on c.DistID=a.DistID" & vbCrLf
                sql &= " join Auth_AccRWPlan d on d.PlanID=a.PlanID" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql += " and a.Years='" & Me.ViewState("years") & "'" & vbCrLf
                sql += " and a.DistID='" & Me.ViewState("distid") & "'" & vbCrLf
                If Me.ViewState("account") <> "" Then '登入者限定。
                    sql += " and d.Account='" & Me.ViewState("account") & "' "
                End If
                sql &= " order by a.PlanID asc" & vbCrLf

            Case "OrgID"
                sql = ""
                sql &= " select distinct a.OrgID,a.OrgName " & vbCrLf
                sql += " from Org_OrgInfo a" & vbCrLf
                sql += " join Auth_Relship b on b.OrgID=a.OrgID " & vbCrLf
                '如果不濾除該計畫沒有申請帳號的單位時，去除以下兩行即可=================================
                sql += " join Auth_Account c on c.OrgID=a.OrgID " & vbCrLf
                sql += " join Auth_AccRWPlan d on d.PlanID=b.PlanID and d.Account=c.Account " & vbCrLf
                '=======================================================================================
                sql += " where 1=1" & vbCrLf
                sql += " and b.PlanID=" & Me.ViewState("planid") & "" & vbCrLf
                sql += " order by a.OrgName asc "
        End Select

        'conn.Open()
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        'With sda
        '    .SelectCommand = New SqlCommand(sql, objconn)
        '    .Fill(ds)
        'End With
        With obj
            .DataSource = dt 'ds.Tables(0)
            .DataTextField = textField
            .DataValueField = valueField
            .DataBind()
        End With

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        '    conn.Close()
        '    If Not sda Is Nothing Then sda.Dispose()
        '    If Not ds Is Nothing Then ds.Dispose()
        'End Try
    End Sub

    '代入user機構單位
    Private Sub Renew_ListOrgID(ByVal obj As DropDownList, ByVal account As String, ByVal roleid As Integer)
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim dr As DataRow = Nothing
        Dim Roles() As Integer = {0, 1}

        'conn.Open()
        Call TIMS.OpenDbConn(objconn)
        Select Case roleid
            Case 0
                sql = ""
                sql &= " Select a.OrgID,a.OrgName from Org_OrgInfo a join Auth_Relship b On b.OrgID=a.OrgID "
                sql += " where b.OrgLevel<=1 order by a.OrgID asc"
            Case 1
                sql = ""
                sql &= "Select a.OrgID,a.OrgName from Org_OrgInfo a join Auth_Account b On b.OrgID=a.OrgID "
                sql += "where b.Account= @account "
        End Select

        If Array.IndexOf(Roles, roleid) <> -1 Then
            With sda
                .SelectCommand = New SqlCommand(sql, objconn)
                .SelectCommand.Parameters.Clear()
                If roleid = 1 Then
                    .SelectCommand.Parameters.Add("account", SqlDbType.VarChar).Value = account
                End If
                .Fill(ds)
            End With

            If Not ds.Tables(0) Is Nothing Then
                If ds.Tables(0).Rows.Count > 0 Then
                    For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                        dr = ds.Tables(0).Rows(i)
                        obj.Items.Insert(i, New ListItem(Convert.ToString(dr("OrgName")), Convert.ToString(dr("OrgID"))))
                    Next
                End If
            End If
        End If

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        '    'conn.Close()
        '    'If Not sda Is Nothing Then sda.Dispose()
        '    'If Not ds Is Nothing Then ds.Dispose()
        'End Try
    End Sub

    '查詢
    Private Sub search()
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet

        'Try
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        '    conn.Close()
        '    If Not sda Is Nothing Then
        '        sda.Dispose()
        '    End If
        '    If Not ds Is Nothing Then
        '        ds.Dispose()
        '    End If
        'End Try

        'conn.Open()
        Call TIMS.OpenDbConn(objconn)
        If flgROLEIDx0xLIDx0 Then
            'snoopy專用
            sql = "Select * from Auth_Group b where GValid='1' and b.GState not in ('D') "
        End If
        If Not flgROLEIDx0xLIDx0 Then
            Select Case Convert.ToString(sm.UserInfo.RoleID)
                Case "0", "1"
                    '系統管理者
                    sql = ""
                    sql &= " select b.GID,b.GDistID,b.GType,b.GName,b.GNote"
                    sql &= " from Auth_Group b"
                    sql += " where 1=1"
                    sql += " and b.GValid='1'"
                    sql += " and b.GState not in('D')"
                    sql += " and (b.GDistID='" & sm.UserInfo.DistID & "' or b.GDistID is null) "
                Case Else
                    '一般使用者
                    sql = ""
                    sql &= " select distinct b.GID,b.GDistID,b.GType,b.GName,b.GNote"
                    sql += " from Auth_GroupAcct a "
                    sql += " join Auth_Group b on b.GID=a.GID"
                    sql += " where 1=1"
                    sql &= " and b.GValid='1'"
                    sql &= " and b.GState not in('D')"
                    sql += " and (b.GDistID='" & sm.UserInfo.DistID & "' or b.GDistID is null)"
                    sql += " and a.Account='" & sm.UserInfo.UserID & "'"
            End Select

            If sm.UserInfo.DistID <> "000" Then
                sql += " and b.GType <>'0' "
            Else
                sql += " and b.GType not in ('1','2') "
            End If
        End If
        'If sm.UserInfo.UserID = "snoopy" Then
        'Else
        'End If
        sql += " order by b.GDistID,b.GType,b.GID asc"

        With sda
            .SelectCommand = New SqlCommand(sql, objconn)
            .Fill(ds)
        End With

        If ds.Tables(0).Rows.Count > 0 Then
            DataGrid3.Visible = True
            lab_Msg2.Visible = False
            tr_btn.Visible = True

            DataGrid3.DataSource = ds.Tables(0)
            DataGrid3.DataKeyField = "GID"
            DataGrid3.DataBind()
        Else
            DataGrid3.Visible = False
            lab_Msg2.Visible = True
            tr_btn.Visible = False
        End If
    End Sub

    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        '非 ROLEID=0 LID=0
        'Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
        flgROLEIDx0xLIDx0 = False
        '如果是系統管理者開啟功能。
        If TIMS.IsSuperUser(Me, 1) Then
            'ROLEID=0 LID=0
            flgROLEIDx0xLIDx0 = True '判斷登入者的權限。
        End If

        If Not IsPostBack Then
            Me.ViewState("years") = sm.UserInfo.Years
            Me.ViewState("distid") = sm.UserInfo.DistID
            Me.ViewState("planid") = sm.UserInfo.PlanID
            Me.ViewState("account") = String.Empty
            If Not flgROLEIDx0xLIDx0 Then
                list_DistID.Enabled = False
                Me.ViewState("account") = sm.UserInfo.UserID
            End If
            'If sm.UserInfo.UserID <> "snoopy" Then
            'Else
            'End If

            Set_DropDownList("PlanYears", list_Years, "Years", "Years")
            Set_DropDownList("DistID", list_DistID, "Name", "DistID")
            Set_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")

            list_Years.SelectedValue = sm.UserInfo.Years
            list_DistID.SelectedValue = sm.UserInfo.DistID
            list_PlanID.SelectedValue = sm.UserInfo.PlanID

            Set_DropDownList("OrgID", list_OrgID, "OrgName", "OrgID")
            Renew_ListOrgID(list_OrgID, sm.UserInfo.UserID, sm.UserInfo.RoleID)

            tr_btn.Visible = False
            rblType.SelectedValue = "0"
            Call search()
        End If
    End Sub

    '儲存。
    Private Sub btn_SaveGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_SaveGroup.Click
        If Convert.ToString(Me.ViewState("planid")) = "" Then
            Common.MessageBox(Me, "請選擇計畫代碼!")
            Exit Sub
        End If
        If Val(list_PlanID.SelectedValue) = 0 Then
            Common.MessageBox(Me, "請選擇計畫代碼!")
            Exit Sub
        End If
        If Val(list_OrgID.SelectedValue) = 0 Then
            Common.MessageBox(Me, "請選擇訓練單位!")
            Exit Sub
        End If

        '取帳號與計畫檔(依單位)
        sql = ""
        sql &= " select a.account"
        sql &= " from auth_accrwplan a"
        sql &= " join auth_account b on b.account=a.account "
        sql += " where 1=1"
        sql &= " and a.PlanID= @planid"
        sql &= " and b.OrgID= @orgid "
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Clear()
            .Parameters.Add("planid", SqlDbType.Int).Value = Val(list_PlanID.SelectedValue)
            .Parameters.Add("orgid", SqlDbType.Int).Value = Val(list_OrgID.SelectedValue)
            dt.Load(.ExecuteReader())
        End With

        Dim TPLANID As String = ""
        TPLANID = TIMS.GetTPlanID(list_PlanID.SelectedValue, objconn)

        sql = " select gid from Auth_GroupAcct where gid= @GID and account= @ACCOUNT"
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " insert into Auth_GroupAcct(GID,ACCOUNT,MODIFYACCT,MODIFYDATE,GTPLANID)" & vbCrLf
        sql += " values(@GID,@ACCOUNT,@MODIFYACCT,getdate(),@GTPLANID)" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " UPDATE Auth_GroupAcct" & vbCrLf
        sql += " SET GTPLANID=@GTPLANID" & vbCrLf
        sql += " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND GID=@GID" & vbCrLf
        sql += " AND ACCOUNT=@ACCOUNT" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = " delete Auth_GroupAcct where gid= @GID and account= @ACCOUNT"
        Dim dCmd As New SqlCommand(sql, objconn)

        '刪除 舊的權限設定(指定Account)
        sql = "delete auth_accrwfun where account= @ACCOUNT"
        Dim d2Cmd As New SqlCommand(sql, objconn)

        '依帳號循環。
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim dr1 As DataRow = dt.Rows(i)
            '依群組循環。
            For Each itm As DataGridItem In DataGrid3.Items
                'DataGrid1.DataKeys.Item(itm.ItemIndex) GID
                Dim GID As String = DataGrid3.DataKeys.Item(itm.ItemIndex)
                Dim chkGroupValid As CheckBox = itm.FindControl("chk_GroupValid") '有勾選。
                If chkGroupValid.Checked = True Then
                    If rblType.SelectedValue = "0" Then
                        '新增群組
                        '判斷帳號是否在群組裡(有->不動作,無->新增帳號群組)
                        Dim dtG As New DataTable
                        With sCmd
                            .Parameters.Clear()
                            .Parameters.Add("GID", SqlDbType.VarChar).Value = GID
                            .Parameters.Add("ACCOUNT", SqlDbType.VarChar).Value = dr1("account")
                            dtG.Load(.ExecuteReader())
                        End With

                        If dtG.Rows.Count = 0 Then
                            '查無資料時新增。
                            With iCmd
                                .Parameters.Clear()
                                .Parameters.Add("GID", SqlDbType.VarChar).Value = GID
                                .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                .Parameters.Add("GTPLANID", SqlDbType.VarChar).Value = TPLANID
                                .ExecuteNonQuery()
                            End With
                        Else
                            '有資料時 UPDATE
                            With uCmd
                                .Parameters.Clear()
                                .Parameters.Add("GTPLANID", SqlDbType.VarChar).Value = TPLANID
                                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                .Parameters.Add("GID", SqlDbType.VarChar).Value = GID
                                .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                                .ExecuteNonQuery()
                            End With
                        End If
                    Else
                        '刪除群組
                        With dCmd
                            .Parameters.Clear()
                            .Parameters.Add("GID", SqlDbType.VarChar).Value = GID
                            .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                            .ExecuteNonQuery()
                        End With
                    End If
                End If
            Next

            '刪除 舊的權限設定(指定Account)
            With d2Cmd
                .Parameters.Clear()
                .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                .ExecuteNonQuery()
            End With
        Next

    End Sub

    Private Sub btn_CancelGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CancelGroup.Click
        tr_btn.Visible = False
        DataGrid3.Visible = False
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem
                Dim labDistName As Label = e.Item.FindControl("lab_DistName")
                Dim labTypeName As Label = e.Item.FindControl("lab_TypeName")
                Dim labGroupName As Label = e.Item.FindControl("lab_GroupName")
                Dim labGroupNote As Label = e.Item.FindControl("lab_GroupNote")

                labDistName.Text = "(系統預設)"
                If Convert.ToString(dr_Data("GDistID")) <> "" Then
                    labDistName.Text = TIMS.Get_DistName1(dr_Data("GDistID"))
                End If

                'Select Case Convert.ToString(dr_Data("GDistID"))
                '    Case "000"
                '        labDistName.Text = "職訓局"
                '    Case "001"
                '        labDistName.Text = "北區職業訓練中心"
                '    Case "002"
                '        labDistName.Text = "泰山職業訓練中心"
                '    Case "003"
                '        labDistName.Text = "桃園職業訓練中心"
                '    Case "004"
                '        labDistName.Text = "中區職業訓練中心"
                '    Case "005"
                '        labDistName.Text = "台南職業訓練中心"
                '    Case "006"
                '        labDistName.Text = "南區職業訓練中心"
                '    Case Else
                '        labDistName.Text = "(系統預設)"
                'End Select

                Select Case Convert.ToString(dr_Data("GType"))
                    Case "0"
                        'labTypeName.Text = "局"
                        labTypeName.Text = "署"
                    Case "1"
                        'labTypeName.Text = "中心"
                        labTypeName.Text = "分署"
                    Case "2"
                        labTypeName.Text = "委訓"
                End Select

                labGroupName.Text += Convert.ToString(dr_Data("GName"))
                labGroupNote.Text = Convert.ToString(dr_Data("GNote"))
        End Select
    End Sub

    Private Sub list_Years_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_Years.SelectedIndexChanged
        Me.ViewState("years") = list_Years.SelectedValue

        Set_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")
        list_PlanID_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub list_DistID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_DistID.SelectedIndexChanged
        Me.ViewState("distid") = list_DistID.SelectedValue

        Set_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")
        list_PlanID_SelectedIndexChanged(sender, e)
    End Sub

    Private Sub list_PlanID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_PlanID.SelectedIndexChanged
        Me.ViewState("planid") = list_PlanID.SelectedValue

        If Me.ViewState("planid") <> "" Then
            Set_DropDownList("OrgID", list_OrgID, "OrgName", "OrgID")
            Renew_ListOrgID(list_OrgID, sm.UserInfo.UserID, sm.UserInfo.RoleID)
        Else
            Common.MessageBox(Me, "該年度無計畫可選擇!")
        End If
    End Sub

    Private Sub list_OrgID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_OrgID.SelectedIndexChanged
        search()
    End Sub
End Class

