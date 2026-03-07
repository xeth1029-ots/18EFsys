Partial Class SYS_03_023
    Inherits AuthBasePage

    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    'Dim sql As String = ""
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
            Me.ViewState("account") = String.Empty
            If Not flgROLEIDx0xLIDx0 Then
                list_DistID.Enabled = False
                Me.ViewState("account") = sm.UserInfo.UserID
            End If

            Set_DropDownList("PlanYears", list_Years, "Years", "Years")
            Set_DropDownList("DistID", list_DistID, "Name", "DistID")
            Set_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")

            list_Years.SelectedValue = sm.UserInfo.Years
            list_DistID.SelectedValue = sm.UserInfo.DistID
            list_PlanID.SelectedValue = sm.UserInfo.PlanID

            tr_btn.Visible = False
            rblType.SelectedValue = "0"
            search()
        End If
    End Sub

    '代入DropDownList資料
    Private Sub Set_DropDownList(ByVal strFlag As String, ByVal obj As DropDownList, ByVal textField As String, ByVal valueField As String)
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim sql As String = ""
        Select Case strFlag
            Case "PlanYears"
                sql = "select distinct Years from ID_Plan where ISNULL(Years,' ')<>' ' order by Years"
            Case "DistID"
                sql = "select DistID,Name from ID_District order by DistID Asc "
            Case "PlanID"
                sql = "" & vbCrLf
                sql += " select distinct a.PlanID,a.Years+c.Name+b.PlanName+a.Seq" & vbCrLf
                'sql += " +nvl2(trim(a.SubTitle), '('+a.SubTitle+')' ,'') PlanName" & vbCrLf
                sql += " +case when ISNULL(a.SubTitle,' ')<>' ' then '('+CONVERT(varchar, a.SubTitle)+')' else '' end PlanName" & vbCrLf
                sql += " from ID_Plan a " & vbCrLf
                sql += " join Key_Plan b on b.TPlanID=a.TPlanID" & vbCrLf
                sql += " join ID_District c on c.DistID=a.DistID" & vbCrLf
                sql += " join Auth_AccRWPlan d on d.PlanID=a.PlanID" & vbCrLf
                sql += " where 1=1" & vbCrLf
                sql += " and a.Years='" & Me.ViewState("years") & "'" & vbCrLf
                sql += " and a.DistID='" & Me.ViewState("distid") & "'" & vbCrLf
                If Me.ViewState("account") <> "" Then
                    sql += " and d.Account='" & Me.ViewState("account") & "'" & vbCrLf
                End If
                sql += " order by a.PlanID asc " & vbCrLf
        End Select

        With sda
            .SelectCommand = New SqlCommand(sql, objconn)
            .Fill(ds)
        End With

        With obj
            .DataSource = ds.Tables(0)
            .DataTextField = textField
            .DataValueField = valueField
            .DataBind()

            If strFlag = "PlanID" Then
                .Items.Insert(0, New ListItem("請選擇", ""))
            End If
        End With

        'Try
        '    conn.Open()
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
    End Sub

    '查詢
    Private Sub search()
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet

        'Try
        '    conn.Open()
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
        Dim sql As String = ""
        If flgROLEIDx0xLIDx0 Then
            'snoopy專用
            sql = "select * from Auth_Group b where b.GValid='1' and b.GState not in ('D')"
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
                    sql += " and b.GState not in ('D')"
                    sql += " and (b.GDistID='" & sm.UserInfo.DistID & "' or b.GDistID is null) "
                Case Else
                    '一般使用者
                    sql = ""
                    sql &= " select distinct b.GID,b.GDistID,b.GType,b.GName,b.GNote"
                    sql += " from AUTH_GROUPACCT a "
                    sql += " join Auth_Group b on b.GID=a.GID"
                    sql += " where 1=1"
                    sql &= " and b.GValid='1'"
                    sql += " and b.GState not in ('D')"
                    sql += " and (b.GDistID='" & sm.UserInfo.DistID & "' or b.GDistID is null)"
                    sql += " and a.Account='" & sm.UserInfo.UserID & "'"
            End Select

            If sm.UserInfo.DistID <> "000" Then
                sql += "and b.GType <>'0' "
            Else
                sql += "and b.GType not in ('1','2') "
            End If
        End If

        'If sm.UserInfo.UserID = "snoopy" Then
        'Else
        'End If
        sql += "order by b.GDistID,b.GType,b.GID asc"

        With sda
            .SelectCommand = New SqlCommand(sql, objconn)
            .Fill(ds)
        End With

        If ds.Tables(0).Rows.Count > 0 Then
            tr_btn.Visible = True
            DataGrid1.Visible = True
            lab_Msg2.Visible = False

            DataGrid1.DataSource = ds.Tables(0)
            DataGrid1.DataKeyField = "GID"
            DataGrid1.DataBind()
        Else
            tr_btn.Visible = False
            DataGrid1.Visible = False
            lab_Msg2.Visible = True
        End If
    End Sub

    '儲存。
    Private Sub btn_SaveGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_SaveGroup.Click
        Call sSaveDatat1()
    End Sub

    Sub sSaveDatat1()
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        'Dim trans As SqlTransaction = Nothing
        'conn.Open()
        'trans = conn.BeginTransaction()

        '取得帳號 (依計畫檔)
        Dim sql As String = ""
        sql = "select account from auth_accrwplan where planid= @planid "
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("planid", SqlDbType.Int).Value = list_PlanID.SelectedValue
            dt.Load(.ExecuteReader())
        End With

        Dim TPLANID As String = ""
        If list_PlanID.SelectedValue <> "" Then TPLANID = TIMS.GetTPlanID(list_PlanID.SelectedValue, objconn)

        sql = " select gid from AUTH_GROUPACCT where gid= @GID and account= @ACCOUNT"
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " insert into AUTH_GROUPACCT(GID,ACCOUNT,MODIFYACCT,MODIFYDATE,GTPLANID)" & vbCrLf
        sql += " values(@GID,@ACCOUNT,@MODIFYACCT,getdate(),@GTPLANID)" & vbCrLf
        Dim iCmd As New SqlCommand(sql, objconn)

        sql = "" & vbCrLf
        sql += " UPDATE AUTH_GROUPACCT" & vbCrLf
        sql += " SET GTPLANID=@GTPLANID" & vbCrLf
        sql += " ,MODIFYACCT=@MODIFYACCT" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND GID=@GID" & vbCrLf
        sql += " AND ACCOUNT=@ACCOUNT" & vbCrLf
        Dim uCmd As New SqlCommand(sql, objconn)

        sql = " delete AUTH_GROUPACCT where gid= @GID and account= @ACCOUNT"
        Dim dCmd As New SqlCommand(sql, objconn)

        '刪除 舊的權限設定(指定Account)
        sql = "delete auth_accrwfun where account= @ACCOUNT"
        Dim d2Cmd As New SqlCommand(sql, objconn)
        'With sda
        '    .SelectCommand = New SqlCommand(sql, objconn)
        '    .SelectCommand.Parameters.Clear()
        '    .SelectCommand.Parameters.Add("planid", SqlDbType.Int).Value = list_PlanID.SelectedValue
        '    .Fill(ds)
        'End With
        '依帳號循環。
        For i As Integer = 0 To dt.Rows.Count - 1
            Dim dr1 As DataRow = dt.Rows(i)
            '依群組循環。
            For Each itm As DataGridItem In DataGrid1.Items
                'DataGrid1.DataKeys.Item(itm.ItemIndex) GID
                Dim GID As String = DataGrid1.DataKeys.Item(itm.ItemIndex)
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
                            '查無資料時新增。-AUTH_GROUPACCT
                            With iCmd
                                .Parameters.Clear()
                                .Parameters.Add("GID", SqlDbType.VarChar).Value = GID
                                .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                                .Parameters.Add("MODIFYACCT", SqlDbType.VarChar).Value = sm.UserInfo.UserID
                                .Parameters.Add("GTPLANID", SqlDbType.VarChar).Value = TPLANID
                                .ExecuteNonQuery()
                            End With
                        Else
                            '有資料時 UPDATE-AUTH_GROUPACCT
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
                        '刪除群組-AUTH_GROUPACCT
                        With dCmd
                            .Parameters.Clear()
                            .Parameters.Add("GID", SqlDbType.VarChar).Value = GID
                            .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                            .ExecuteNonQuery()
                        End With
                    End If
                End If
            Next

            '刪除 舊的權限設定(指定Account)-AUTH_ACCRWFUN
            With d2Cmd
                .Parameters.Clear()
                .Parameters.Add("ACCOUNT", SqlDbType.NVarChar).Value = dr1("account")
                .ExecuteNonQuery()
            End With
        Next

        'trans.Commit()
        Common.MessageBox(Me, "儲存成功!")
        'Try


        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        '    trans.Rollback()
        'Finally
        '    conn.Close()
        '    If Not sda Is Nothing Then
        '        sda.Dispose()
        '    End If
        '    If Not ds Is Nothing Then
        '        ds.Dispose()
        '    End If
        'End Try

    End Sub

    Private Sub btn_CancelGroup_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_CancelGroup.Click
        tr_btn.Visible = False
        DataGrid1.Visible = False

        list_PlanID.SelectedValue = ""
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
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

        tr_btn.Visible = False
        DataGrid1.Visible = False
    End Sub

    Private Sub list_DistID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_DistID.SelectedIndexChanged
        Me.ViewState("distid") = list_DistID.SelectedValue
        Set_DropDownList("PlanID", list_PlanID, "PlanName", "PlanID")

        tr_btn.Visible = False
        DataGrid1.Visible = False
    End Sub

    Private Sub list_PlanID_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_PlanID.SelectedIndexChanged
        If list_PlanID.SelectedValue <> "" Then
            search()
        Else
            tr_btn.Visible = False
            DataGrid1.Visible = False
        End If
    End Sub
End Class


