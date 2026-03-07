Partial Class SYS_03_022
    Inherits AuthBasePage

    'select * from AUTH_GROUP where 1=1 AND dbo.TRUNC_DATETIME(MODIFYDATE)=dbo.TRUNC_DATETIME(getdate())
    'select * from AUTH_GROUPFUN where 1=1 AND dbo.TRUNC_DATETIME(MODIFYDATE)=dbo.TRUNC_DATETIME(getdate())
    'select * from AUTH_GROUP where 1=1、AUTH_GROUPFUN

    Dim vMsg As String = ""
    Dim trID As Integer = 0
    Const cst_tab2 As String = "&nbsp;&nbsp;&nbsp;　"
    Dim arrFun As String()  '= {"TC", "SD", "CP", "TR", "CM", "OB", "SE", "EXAM", "SV", "SYS", "FAQ", "OO"} 'fun排列順序
    'Dim FunSort As String = System.Configuration.ConfigurationSettings.AppSettings("FunSort")

#Region "NO USE"
    'select * from AUTH_GROUP where gname like '%計畫承辦人(北區)%'
    'with wc1 as (SELECT * FROM AUTH_GROUPFUN where gid ='16')
    ', wc2 as (
    'SELECT * FROM AUTH_GROUPFUN 
    'AS OF TIMESTAMP TO_TIMESTAMP('2016/08/17 19:00:25', 'YYYY/MM/DD hh24:mi:ss') 
    'where 1=1
    'and gid ='16'
    ')
    'SELECT * FROM  wc2 m 
    'where 1=1 
    'and not exists (select 'x' from wc1 x where x.funid=m.funid )


    ''代入資料
    'Private Sub LoadData()
    '    Dim sda As New SqlDataAdapter
    '    Dim ds As New DataSet
    '    Dim dr As DataRow = Nothing

    '    Call TIMS.OpenDbConn(objconn)
    '    Try
    '        sql = "SELECT * FROM AUTH_GROUP WHERE GID= @gid"
    '        With sda
    '            .SelectCommand = New SqlCommand(sql, objconn)
    '            .SelectCommand.Parameters.Clear()
    '            .SelectCommand.Parameters.Add("gid", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
    '            .Fill(ds)
    '        End With

    '        dr = ds.Tables(0).Rows(0)

    '        If ddlDistID.Enabled = True Then ddlDistID.SelectedValue = Convert.ToString(dr("gdistid"))
    '        If Convert.ToString(dr("gdistid")) = "" Then
    '            hidIsSys.Value = "Y" '系統取出。
    '        Else
    '            hidIsSys.Value = "N"
    '        End If
    '        ddlGtype.SelectedValue = Convert.ToString(dr("gtype"))
    '        txt_GroupName.Text = Convert.ToString(dr("gname"))
    '        txt_GroupNote.Text = Convert.ToString(dr("gnote"))

    '        If Convert.ToString(dr("gvalid")) Then
    '            chk_Valid.Checked = True
    '        Else
    '            chk_Valid.Checked = False
    '        End If
    '        'If conn.State = ConnectionState.Open Then conn.Close()
    '    Catch ex As Exception
    '        Common.MessageBox(Me, ex.ToString)
    '    Finally
    '        'conn.Close()
    '        If Not sda Is Nothing Then sda.Dispose()
    '        If Not ds Is Nothing Then ds.Dispose()

    '    End Try
    'End Sub

#End Region

    Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        arrFun = TIMS.c_FUNSORT.Split(",") 'arrFun = FunSort.Split(",")

        '非 ROLEID=0 LID=0 'Dim flgROLEIDx0xLIDx0 As Boolean = False '判斷登入者的權限。'如果是系統管理者開啟功能。 'ROLEID=0 LID=0
        flgROLEIDx0xLIDx0 = If(TIMS.IsSuperUser(Me, 1), True, False) '判斷登入者的權限。

        If Not IsPostBack Then
            btn_Save.Attributes.Add("onClick", "return Check_Data();")
            Call ClsVisible1(1)

            list_MainMenu = TIMS.Get_ddlFunction(list_MainMenu, 2)

            Set_DropDownList("DistID", ddlQDistID, "Name", "DistID")
            Set_DropDownList("DistID", ddlDistID, "Name", "DistID")
            Set_DropDownList("DistID", ddlDistID3, "Name", "DistID")

            If Not flgROLEIDx0xLIDx0 Then
                ddlQDistID.Enabled = False
                ddlDistID.Enabled = False

                ddlQDistID.SelectedValue = sm.UserInfo.DistID
                ddlDistID.SelectedValue = sm.UserInfo.DistID

                If sm.UserInfo.DistID <> "000" Then
                    ddlQType.Items.RemoveAt(1)
                    ddlGtype.Items.RemoveAt(1)
                Else
                    ddlQType.Enabled = False
                    ddlGtype.Enabled = False

                    ddlQType.SelectedValue = "0"
                    ddlGtype.SelectedValue = "0"
                End If
            End If

            Call search1()
        End If
    End Sub

    'Visible
    Sub ClsVisible1(ByVal iType As Integer)
        tb_Query.Visible = False
        tb_List.Visible = False
        Panel_Exe1.Visible = False
        tb_Edit.Visible = False
        Select Case iType
            Case 1 '查詢功能
                tb_Query.Visible = True
                tb_List.Visible = True
            Case 2 '編輯功能2
                tb_Edit.Visible = True
            Case 3 '編輯功能3
                Panel_Exe1.Visible = True
        End Select
    End Sub

    '查詢
    Private Sub btn_Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Query.Click
        Call search1()
    End Sub

    '新增
    Private Sub btn_Add_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Add.Click
        Call Clear_Items()

        Me.ViewState("mainmenu") = Nothing
        Me.ViewState("act") = "add"

        Call Get_List3()

        Call ClsVisible1(2)
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        txt_GroupName.Text = TIMS.ClearSQM(txt_GroupName.Text)
        txt_GroupNote.Text = TIMS.ClearSQM(txt_GroupNote.Text)
        If ddlDistID.SelectedValue = "" Then
            'If sm.UserInfo.UserID <> "snoopy" Then Errmsg += "請選擇建檔單位" & vbCrLf
            If Not flgROLEIDx0xLIDx0 Then Errmsg += "請選擇建檔單位" & vbCrLf
        End If
        If ddlGtype.SelectedValue = "" Then Errmsg += "請選擇群組階層" & vbCrLf
        If txt_GroupName.Text = "" Then Errmsg += "請輸入群組名稱" & vbCrLf
        If txt_GroupNote.Text = "" Then Errmsg += "請輸入群組備註" & vbCrLf
        If Errmsg <> "" Then
            Rst = False
            Return Rst
        End If

        ViewState("txtgroupname") = Trim(txt_GroupName.Text)
        ViewState("txtgroupnote") = Trim(txt_GroupNote.Text)
        ViewState("chkvalid") = If(chk_Valid.Checked, "1", "0") '1:啟用/0:停用

        Dim sql As String = ""
        sql = "select * from AUTH_GROUP where GDistID= @GDistID and GType= @GType and GName= @GName and GID!= @GID"
        Dim sCmd As New SqlCommand(sql, objconn)

        sql = "select * from AUTH_GROUP where GDistID= @GDistID and GType= @GType and GName= @GName"
        Dim sCmd2 As New SqlCommand(sql, objconn)

        If Errmsg = "" Then
            Try
                Select Case Convert.ToString(Me.ViewState("act")) '編輯號碼
                    Case "update" '修改儲存
                        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
                        'da.SelectCommand.Parameters.Clear()
                        Dim dt As New DataTable
                        With sCmd
                            .Parameters.Clear()
                            .Parameters.Add("GDistID", SqlDbType.VarChar).Value = If(ddlDistID.SelectedValue <> "", ddlDistID.SelectedValue, Convert.DBNull)
                            .Parameters.Add("GType", SqlDbType.VarChar).Value = ddlGtype.SelectedValue
                            .Parameters.Add("GName", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupname")))
                            .Parameters.Add("GID", SqlDbType.Int).Value = CInt(hide_GID.Value)
                            'dt.Load(.ExecuteReader())
                            dt = DbAccess.GetDataTable(sCmd.CommandText, objconn, sCmd.Parameters)
                        End With

                        If dt.Rows.Count > 0 Then
                            Errmsg += "建檔單位/群組階層/群組名稱,已經存在,請確認選擇與輸入" & vbCrLf
                        End If

                    Case "add" '新增儲存
                        'Dim dt As New DataTable
                        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
                        'da.SelectCommand.Parameters.Clear()
                        Dim dt As New DataTable
                        With sCmd2
                            .Parameters.Clear()
                            .Parameters.Add("GDistID", SqlDbType.VarChar).Value = If(ddlDistID.SelectedValue <> "", ddlDistID.SelectedValue, Convert.DBNull)
                            .Parameters.Add("GType", SqlDbType.VarChar).Value = ddlGtype.SelectedValue
                            .Parameters.Add("GName", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupname")))
                            'dt.Load(.ExecuteReader())
                            dt = DbAccess.GetDataTable(sCmd2.CommandText, objconn, sCmd2.Parameters)
                        End With

                        If dt.Rows.Count > 0 Then
                            Errmsg += "建檔單位/群組階層/群組名稱,已經存在,請重新選擇與輸入" & vbCrLf
                        End If
                End Select
            Catch ex As Exception
                Errmsg += "檢查 建檔單位/群組階層/群組名稱 是否重覆,發生錯誤" & vbCrLf
                Errmsg += ex.ToString & vbCrLf
            End Try

        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    ''' <summary>
    ''' 儲存
    ''' </summary>
    Sub SaveData1()
        'Dim sda As New SqlDataAdapter
        Dim intGID As Int32 = 0 '新增時記錄群組主檔ID
        Dim intCnt As Integer = 1 '判斷儲存是否成功(0=>否,1=>是)

        Me.ViewState("txtgroupname") = Trim(txt_GroupName.Text)
        Me.ViewState("txtgroupnote") = Trim(txt_GroupNote.Text)
        ViewState("chkvalid") = If(chk_Valid.Checked, "1", "0") '1:啟用/0:停用

        Call TIMS.OpenDbConn(objconn)
        Try
            Select Case Convert.ToString(Me.ViewState("act"))
                Case "update" '修改儲存
                    If IsNumeric(hide_GID.Value) = True Then
                        '修改群組主檔
                        Dim sql As String = ""
                        sql &= " update AUTH_GROUP"
                        sql &= " set GDistID= @GDistID ,GType= @GType ,GName= @GName ,GNote= @GNote"
                        sql &= " ,GValid= @GValid" '1:啟用/0:停用
                        sql &= " ,ModifyAcct= @ModifyAcct ,ModifyDate=getdate() ,GState='U'"
                        sql &= " where GID= @GID"
                        Dim uCmd As New SqlCommand(sql, objconn)
                        With uCmd
                            .Parameters.Clear()
                            If hidIsSys.Value = "Y" Then
                                .Parameters.Add("GDistID", SqlDbType.NVarChar).Value = Convert.DBNull
                            Else
                                .Parameters.Add("GDistID", SqlDbType.NVarChar).Value = ddlDistID.SelectedValue
                            End If
                            .Parameters.Add("GType", SqlDbType.NVarChar).Value = ddlGtype.SelectedValue
                            .Parameters.Add("GName", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupname")))
                            .Parameters.Add("GNote", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupnote")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupnote")))
                            .Parameters.Add("GValid", SqlDbType.Decimal, 1).Value = Convert.ToDecimal(Me.ViewState("chkvalid")) '1:啟用/0:停用
                            .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                            .Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                            'If conn.State = ConnectionState.Closed Then conn.Open()
                            ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                            '.ExecuteNonQuery()
                            DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                        End With

                        'With sda
                        '    .UpdateCommand = New SqlCommand(sql, objconn)
                        '    .UpdateCommand.Parameters.Clear()
                        '    If hidIsSys.Value = "Y" Then
                        '        .UpdateCommand.Parameters.Add("GDistID", SqlDbType.NVarChar).Value = Convert.DBNull
                        '    Else
                        '        .UpdateCommand.Parameters.Add("GDistID", SqlDbType.NVarChar).Value = ddlDistID.SelectedValue
                        '    End If
                        '    .UpdateCommand.Parameters.Add("GType", SqlDbType.NVarChar).Value = ddlGtype.SelectedValue
                        '    .UpdateCommand.Parameters.Add("GName", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupname")))
                        '    .UpdateCommand.Parameters.Add("GNote", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupnote")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupnote")))
                        '    .UpdateCommand.Parameters.Add("GValid", SqlDbType.Decimal, 1).Value = Convert.ToBoolean(Me.ViewState("chkvalid"))
                        '    .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        '    .UpdateCommand.Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                        '    'If conn.State = ConnectionState.Closed Then conn.Open()
                        '    .UpdateCommand.ExecuteNonQuery()
                        'End With

                        '刪除現有的群組明細資料(分all & 單一)
                        'sql = ""
                        'sql &= " delete AUTH_GROUPFUN where GID= @GID "
                        'If list_MainMenu.SelectedIndex <> 0 Then
                        '    sql &= " and funid in (select funid from id_function where kind='" & list_MainMenu.SelectedValue & "')"
                        'End If
                        'Dim dCmd As New SqlCommand(sql, objconn)
                        'With dCmd
                        '    .Parameters.Clear()
                        '    .Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                        '    'If conn.State = ConnectionState.Closed Then conn.Open()
                        '    .ExecuteNonQuery()
                        'End With

                        'With sda
                        '    .DeleteCommand = New SqlCommand(sql, objconn)
                        '    .DeleteCommand.Parameters.Clear()
                        '    .DeleteCommand.Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                        '    'If conn.State = ConnectionState.Closed Then conn.Open()
                        '    .DeleteCommand.ExecuteNonQuery()
                        'End With

                        '新增修改過後之群組明資料
                        sql = ""
                        sql &= " insert into AUTH_GROUPFUN(GID,FunID,ModifyAcct,ModifyDate)"
                        sql &= " values(@GID,@FunID,@ModifyAcct,getdate())"
                        Dim iCmd3 As New SqlCommand(sql, objconn)

                        sql = ""
                        sql &= " select 'x' FROM AUTH_GROUPFUN "
                        sql &= " where gid =@GID and funid =@FunID"
                        Dim sCmd3 As New SqlCommand(sql, objconn)

                        sql = ""
                        sql &= " delete AUTH_GROUPFUN"
                        sql &= " where gid =@GID and funid =@FunID"
                        Dim dCmd3 As New SqlCommand(sql, objconn)

                        For Each itm As DataGridItem In DataGrid2.Items
                            Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
                            If chkEnable.Checked = True Then
                                Dim dt3 As New DataTable
                                With sCmd3
                                    .Parameters.Clear()
                                    .Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                                    .Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid2.DataKeys.Item(itm.ItemIndex))
                                    'dt3.Load(.ExecuteReader())
                                    dt3 = DbAccess.GetDataTable(sCmd3.CommandText, objconn, sCmd3.Parameters)
                                End With
                                If dt3.Rows.Count = 0 Then
                                    With iCmd3
                                        .Parameters.Clear()
                                        .Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                                        .Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid2.DataKeys.Item(itm.ItemIndex))
                                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                                        '.ExecuteNonQuery()
                                        DbAccess.ExecuteNonQuery(iCmd3.CommandText, objconn, iCmd3.Parameters)
                                    End With
                                End If
                            Else
                                With dCmd3
                                    .Parameters.Clear()
                                    .Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                                    .Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid2.DataKeys.Item(itm.ItemIndex))
                                    '.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                                    '.ExecuteNonQuery()
                                    DbAccess.ExecuteNonQuery(dCmd3.CommandText, objconn, dCmd3.Parameters)
                                End With
                            End If
                        Next

                        'With sda
                        '    .InsertCommand = New SqlCommand(sql, objconn)
                        '    For Each itm As DataGridItem In DataGrid2.Items
                        '        Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
                        '        If chkEnable.Checked = True Then
                        '            .InsertCommand.Parameters.Clear()
                        '            .InsertCommand.Parameters.Add("GID", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
                        '            .InsertCommand.Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid2.DataKeys.Item(itm.ItemIndex))
                        '            .InsertCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        '            'If conn.State = ConnectionState.Closed Then conn.Open()
                        '            .InsertCommand.ExecuteNonQuery()
                        '        End If
                        '    Next
                        'End With

                    Else
                        intCnt = 0
                        Common.MessageBox(Me, "儲存失敗(ID lost)。")
                    End If

                Case "add" '新增儲存
                    Dim objGDistID As Object = ddlDistID.SelectedValue
                    If hidIsSys.Value = "Y" Then
                        If sm.UserInfo.RoleID = "1" Then
                            objGDistID = ddlDistID.SelectedValue
                        Else
                            objGDistID = Convert.DBNull
                        End If
                    End If

                    '新增群組主檔
                    Dim sql As String = ""
                    sql &= " insert into AUTH_GROUP(GID,GDistID,GType,GName,GNote,GValid,CreateAcct,ModifyAcct,ModifyDate)"
                    sql &= " VALUES (@GID,@GDistID, @GType, @GName, @GNote, @GValid, @CreateAcct, @ModifyAcct,current_timestamp)"
                    Dim iCmd As New SqlCommand(sql, objconn)
                    intGID = DbAccess.GetNewId(objconn, "AUTH_GROUP_GID_SEQ,AUTH_GROUP,GID")
                    With iCmd
                        .Parameters.Clear()
                        .Parameters.Add("GID", SqlDbType.Int).Value = intGID
                        .Parameters.Add("GDistID", SqlDbType.NVarChar).Value = objGDistID
                        .Parameters.Add("GType", SqlDbType.NVarChar).Value = ddlGtype.SelectedValue
                        .Parameters.Add("GName", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupname")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupname")))
                        .Parameters.Add("GNote", SqlDbType.NVarChar).Value = If(Convert.ToString(Me.ViewState("txtgroupnote")) = "", Convert.DBNull, Convert.ToString(Me.ViewState("txtgroupnote")))
                        .Parameters.Add("GValid", SqlDbType.Decimal, 1).Value = Convert.ToDecimal(Me.ViewState("chkvalid")) '1:啟用/0:停用
                        .Parameters.Add("CreateAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                        'If conn.State = ConnectionState.Closed Then conn.Open()
                        ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                        '.ExecuteNonQuery()
                        DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, iCmd.Parameters)
                    End With

                    'sql = "select AUTH_GROUP_GID_SEQ.CURRVAL "
                    '.InsertCommand = New SqlCommand(sql, objconn)
                    '.InsertCommand.Parameters.Clear()
                    'intGID = Convert.ToInt32(.InsertCommand.ExecuteScalar())


                    '新增群組明資料
                    sql = ""
                    sql &= " insert into AUTH_GROUPFUN(GID,FunID,ModifyAcct,ModifyDate)"
                    sql &= " values( @GID, @FunID, @ModifyAcct,current_timestamp)"
                    Dim iCmd2 As New SqlCommand(sql, objconn)

                    For Each itm As DataGridItem In DataGrid2.Items
                        Dim chkEnable As CheckBox = itm.FindControl("chk_Enable")
                        If chkEnable.Checked = True Then
                            With iCmd2
                                .Parameters.Clear()
                                .Parameters.Add("GID", SqlDbType.Int).Value = intGID
                                .Parameters.Add("FunID", SqlDbType.Int).Value = CInt(DataGrid2.DataKeys.Item(itm.ItemIndex))
                                .Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                                'If conn.State = ConnectionState.Closed Then conn.Open()
                                ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                                '.ExecuteNonQuery()
                                DbAccess.ExecuteNonQuery(iCmd2.CommandText, objconn, iCmd2.Parameters)
                            End With
                        End If
                    Next

            End Select
            'If conn.State = ConnectionState.Open Then conn.Close()

        Catch ex As Exception
            intCnt = 0
            Common.MessageBox(Me, ex.ToString)
            'Finally
            'If conn.State = ConnectionState.Open Then conn.Close()
            'If Not sda Is Nothing Then sda.Dispose()
        End Try

        If intCnt = 1 Then
            Call search1()

            Call Clear_Items()
            Common.MessageBox(Me, "儲存成功!")
            Call ClsVisible1(1)
        End If
    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btn_Save_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save.Click, btn_Save_2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        SaveData1()
    End Sub

    '回上一頁
    Private Sub btn_LoadBack_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_LoadBack.Click, btn_LoadBack_2.Click
        Call Clear_Items()
        Call ClsVisible1(1)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As System.Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sCmdArg As String = e.CommandArgument
        If sCmdArg = "" Then Exit Sub
        hide_GID.Value = TIMS.GetMyValue(sCmdArg, "GID")
        If hide_GID.Value = "" Then Exit Sub

        Select Case UCase(e.CommandName)
            Case "EXE1"
                Call ClsVisible1(3)
                Call LoadData3(hide_GID.Value, 3)
                Call Search3(hide_GID.Value)

            Case "EDIT"
                'hide_GID.Value = DataGrid1.DataKeys.Item(e.Item.ItemIndex)
                Me.ViewState("act") = "update"
                Call ClsVisible1(2)
                'tb_Query.Visible = False
                'tb_List.Visible = False
                'tb_Edit.Visible = True
                Call LoadData3(hide_GID.Value, 1)
                Call Get_List3()

            Case "COPY"
                'hide_GID.Value = DataGrid1.DataKeys.Item(e.Item.ItemIndex)
                Me.ViewState("act") = "update"
                Call ClsVisible1(2)
                'tb_Query.Visible = False
                'tb_List.Visible = False
                'tb_Edit.Visible = True
                Call LoadData3(hide_GID.Value, 1)
                Call Get_List3()

                hide_GID.Value = ""
                Me.ViewState("act") = "add"
                txt_GroupName.Text = ""

            Case "DEL"
                'Dim sda As New SqlDataAdapter
                Dim intCnt As Integer = 0
                Call TIMS.OpenDbConn(objconn)
                Dim sql As String = ""
                sql = " UPDATE AUTH_GROUP SET GSTATE= @GState where GID= @GID"
                Dim uCmd As New SqlCommand(sql, objconn)
                Call TIMS.OpenDbConn(objconn)
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("GState", SqlDbType.VarChar).Value = "D"
                    .Parameters.Add("GID", SqlDbType.Int).Value = CInt(hide_GID.Value) 'CInt(DataGrid1.DataKeys.Item(e.Item.ItemIndex))
                    'If conn.State = ConnectionState.Closed Then conn.Open()
                    'intCnt = .ExecuteNonQuery()
                    ' 為了統一寫入資料異動記錄, 一律改為呼叫 DBAccess.ExecuteNonQuery()
                    '.ExecuteNonQuery()
                    intCnt = DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, uCmd.Parameters)
                End With
                If intCnt = 0 Then
                    Common.MessageBox(Me, "沒有資料更新!")
                    Exit Sub
                End If
                Call search1()

                Call ClsVisible1(1)
                'tb_List.Visible = True
                Common.MessageBox(Me, "刪除成功!")

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem
                Dim labSNo As Label = e.Item.FindControl("lab_SNo")
                Dim lab_GroupDistID As Label = e.Item.FindControl("lab_GroupDistID")
                Dim lab_GroupType As Label = e.Item.FindControl("lab_GroupType")
                Dim labGroupName As Label = e.Item.FindControl("lab_GroupName")
                Dim lab_GroupCUsr As Label = e.Item.FindControl("lab_GroupCUsr")
                Dim lab_GroupMUsr As Label = e.Item.FindControl("lab_GroupMUsr")
                Dim labGroupNote As Label = e.Item.FindControl("lab_GroupNote")
                Dim labEnable As Label = e.Item.FindControl("lab_Enable")
                Dim btnEdit As Button = e.Item.FindControl("btn_Edit")
                Dim btnCopy As Button = e.Item.FindControl("btn_Copy")
                Dim btnDel As Button = e.Item.FindControl("btn_Del")
                '按下此按鈕，則另開視窗顯示有此群組權限的使用者，顯示的資訊有單位與使用者姓名與帳號。
                Dim btnExe1 As Button = e.Item.FindControl("btnExe1")

                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                lab_GroupDistID.Text = "(系統預設)"
                If Convert.ToString(dr_Data("GDistID")) <> "" Then
                    lab_GroupDistID.Text = TIMS.Get_DistName1(dr_Data("GDistID"))
                End If

                Select Case Convert.ToString(dr_Data("GType"))
                    Case "0"
                        lab_GroupType.Text = "署"
                    Case "1"
                        lab_GroupType.Text = "分署"
                    Case "2"
                        lab_GroupType.Text = "委訓"
                End Select

                labGroupName.Text = Convert.ToString(dr_Data("GName"))
                lab_GroupCUsr.Text = Convert.ToString(dr_Data("cname"))
                lab_GroupMUsr.Text = Convert.ToString(dr_Data("mname"))
                labGroupNote.Text = Convert.ToString(dr_Data("GNote"))
                labEnable.Text = If(Convert.ToBoolean(dr_Data("GValid")) = True, "是", "否")

                btnEdit.Attributes.Add("style", "cursor@hand")
                btnCopy.Attributes.Add("style", "cursor@hand")
                btnDel.Attributes.Add("style", "cursor@hand")
                btnExe1.Attributes.Add("style", "cursor@hand")

                Dim flagCanEdtDel As Boolean = True
                If Convert.ToString(dr_Data("GDistID")) = "" _
                    AndAlso Not flgROLEIDx0xLIDx0 Then
                    flagCanEdtDel = False

                    vMsg = "系統權限"
                    btnEdit.Enabled = False
                    btnDel.Enabled = False
                    TIMS.Tooltip(btnEdit, vMsg)
                    TIMS.Tooltip(btnDel, vMsg)
                    'btnCopy.Enabled = False
                    'btnExe1.Enabled = False
                End If

                If flagCanEdtDel Then
                    If Convert.ToString(dr_Data("GID")) <> "" Then
                        TIMS.Tooltip(btnEdit, "GID:" & dr_Data("GID"))
                    End If
                    If chkGroup(Convert.ToInt32(dr_Data("gid"))) Then
                        btnDel.Enabled = False
                        vMsg = "該群組已附予帳號,故無法刪除!"
                        TIMS.Tooltip(btnDel, vMsg)
                    Else
                        btnDel.Attributes.Add("style", "cursor@hand")
                        btnDel.Attributes.Add("onclick", "return confirm('確認將【" & labGroupName.Text & "】刪除?');")
                    End If
                End If

                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "GID", Convert.ToString(dr_Data("GID")))
                btnEdit.CommandArgument = sCmdArg
                btnCopy.CommandArgument = sCmdArg
                btnDel.CommandArgument = sCmdArg
                btnExe1.CommandArgument = sCmdArg

        End Select
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim chkEnableAll As CheckBox = e.Item.FindControl("chk_EnableAll") '選用全選方塊
                Dim chkSchAll As CheckBox = e.Item.FindControl("chk_SchAll") '查詢全選方塊
                Dim chkEditAll As CheckBox = e.Item.FindControl("chk_EditAll") '維護全選方塊
                Dim chkPrtAll As CheckBox = e.Item.FindControl("chk_PrtAll") '列印全選方塊

                chkEnableAll.Attributes.Add("onclick", "Show_SelectAll('" & chkEnableAll.ClientID & "','" & Convert.ToString(Me.ViewState("itemname")) & "')")
                chkSchAll.Attributes.Add("onclick", "Show_SelectAll('" & chkSchAll.ClientID & "','" & Convert.ToString(Me.ViewState("itemname")) & "')")
                chkEditAll.Attributes.Add("onclick", "Show_SelectAll('" & chkEditAll.ClientID & "','" & Convert.ToString(Me.ViewState("itemname")) & "')")
                chkPrtAll.Attributes.Add("onclick", "Show_SelectAll('" & chkPrtAll.ClientID & "','" & Convert.ToString(Me.ViewState("itemname")) & "')")

                chkEnableAll.Attributes.Add("onclick", "Show_SelectAll('DataGrid2','" & chkEnableAll.ClientID & "',3)")
                chkSchAll.Attributes.Add("onclick", "Show_SelectAll('DataGrid2','" & chkSchAll.ClientID & "',4)")
                chkEditAll.Attributes.Add("onclick", "Show_SelectAll('DataGrid2','" & chkEditAll.ClientID & "',5)")
                chkPrtAll.Attributes.Add("onclick", "Show_SelectAll('DataGrid2','" & chkPrtAll.ClientID & "',6)")

            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim txtFunID As TextBox = e.Item.FindControl("txtFunID") '記錄FunID
                Dim labMainMenu As Label = e.Item.FindControl("lab_MainMenu") '程式類型
                Dim labFunName As Label = e.Item.FindControl("lab_FunName") '選單名稱
                Dim chkEnable As CheckBox = e.Item.FindControl("chk_Enable") '選用方塊
                Dim chkSch As CheckBox = e.Item.FindControl("chk_Sch") '查詢方塊
                Dim chkEdit As CheckBox = e.Item.FindControl("chk_Edit") '維護方塊
                Dim chkPrt As CheckBox = e.Item.FindControl("chk_Prt") '列印方塊

                txtFunID.Text = Convert.ToString(drv("funid"))

                'labMainMenu.Text = Get_MainMenuName(Convert.ToString(drv("Kind")))
                Dim strkind As String = Convert.ToString(drv("Kind"))
                labMainMenu.Text = TIMS.Get_MainMenuName(strkind)

                If Convert.ToString(drv("levels")) = "1" Then
                    labFunName.Text = "&nbsp;&nbsp;&nbsp;" & "&nbsp;&nbsp;&nbsp;" & Convert.ToString(drv("Name"))
                Else
                    labFunName.Text = Convert.ToString(drv("Name"))

                    chkSch.Visible = False
                    chkEdit.Visible = False
                    chkPrt.Visible = False
                End If

                'snoopy特別權限
                If flgROLEIDx0xLIDx0 Then
                    '連線系統本身功能
                    Dim sPage As String = Convert.ToString(drv("Spage"))
                    If sPage <> "" Then
                        Dim strUrl As String = ""
                        '含有http字眼另開視窗
                        Const strUrl_f1 As String = "<a href=""../../{0}?ID={1}"" target=""_self"">({0})</a>"
                        Const strUrl_f2 As String = "<a href=""{0}"" target=""_blank"">({1})</a>"
                        Dim flag_new_windows As Boolean = TIMS.CHK_FUNCSPAGE2(sm, sPage)
                        If flag_new_windows Then
                            Dim sPage1 As String = sPage
                            If sPage1.Length > 40 Then sPage1 = TIMS.ClearSQM(sPage1.Substring(0, 40)) & "..."
                            strUrl = String.Format(strUrl_f2, sPage, sPage1)
                        Else
                            '回上2層並給ID
                            If sPage.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase) Then sPage = sPage.Substring(0, sPage.Length - 5)
                            strUrl = String.Format(strUrl_f1, sPage, Convert.ToString(drv("funid")))
                        End If
                        labFunName.Text &= strUrl
                    End If
                End If

                TIMS.Tooltip(labFunName, Convert.ToString(drv("funid")))

                If Me.ViewState("act") = "update" Then
                    Dim dt As DataTable = chkData(Convert.ToInt32(drv("funid")))

                    If dt.Rows.Count > 0 Then
                        chkEnable.Checked = True

                        If Convert.ToString(dt.Rows(0)("sech")) <> "" Then chkSch.Checked = True
                        If Convert.ToString(dt.Rows(0)("adds")) <> "" Then chkEdit.Checked = True
                        If Convert.ToString(dt.Rows(0)("prnt")) <> "" Then chkPrt.Checked = True
                    End If
                End If

                If Not (drv("Kind") = "TC" Or drv("Kind") = "SD") Then
                    chkSch.Enabled = False
                    chkEdit.Enabled = False
                    chkPrt.Enabled = False
                End If

                If Convert.ToString(Me.ViewState("mmname")) <> labMainMenu.Text Then  '同類型選單的第一項
                    Dim subs As Integer = Me.ViewState(Convert.ToString(drv("Kind"))) '依類型取得選單數量

                    Me.ViewState("mmname") = labMainMenu.Text
                    e.Item.Cells(0).RowSpan = subs '合併同類型選單
                    e.Item.Cells(0).BackColor = Color.FromArgb(241, 249, 252)
                    e.Item.Attributes.Add("id", Convert.ToString(drv("Kind")))

                    For i As Integer = 0 To e.Item.Cells.Count - 1
                        e.Item.Cells(i).Attributes.Add("id", Convert.ToString(drv("Kind")) & "td" & i)
                    Next

                    trID = 1

                Else '非同類型選單第一項的其他項
                    e.Item.Cells(0).Visible = False '隱藏功能類別欄位
                    e.Item.Attributes.Add("id", Convert.ToString(drv("Kind")) & trID)

                    trID += 1
                End If

                '設定主選單顏色
                If Convert.ToString(drv("Subs")) <> "0" Then
                    e.Item.BackColor = Color.FromArgb(235, 243, 254)
                End If

                '紀錄第一項的ClientID
                If e.Item.ItemIndex = 0 Then
                    Me.ViewState("itemname") = chkEnable.ClientID
                End If

                chkEnable.Attributes.Add("onclick", "Show_Select('enable', " & e.Item.ItemIndex + 1 & ")")
                chkSch.Attributes.Add("onclick", "Show_Select('sch', " & e.Item.ItemIndex + 1 & ")")
                chkEdit.Attributes.Add("onclick", "Show_Select('edit', " & e.Item.ItemIndex + 1 & ")")
                chkPrt.Attributes.Add("onclick", "Show_Select('prt', " & e.Item.ItemIndex + 1 & ")")

        End Select
    End Sub

    'Private Sub list_MainMenu_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles list_MainMenu.SelectedIndexChanged
    '    Call Get_List3()
    'End Sub

    '查詢 被賦予者
    Sub Search3(ByVal vGID As String)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        Dim sql As String = ""
        sql &= " select aa.name" & vbCrLf
        sql &= " ,aa.account" & vbCrLf
        sql &= " ,oo.orgname" & vbCrLf
        sql &= " ,aa.ISUSED" & vbCrLf
        sql &= " from auth_groupacct gp" & vbCrLf
        sql &= " join AUTH_GROUP ag on ag.gid =gp.gid" & vbCrLf
        sql &= " join auth_account aa on aa.account =gp.account" & vbCrLf
        sql &= " join org_orginfo oo on oo.orgid =aa.orgid" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and ag.GState<>'D'" & vbCrLf
        sql &= " and gp.gid=@gid " & vbCrLf
        sql &= " and aa.roleid >= @roleid " & vbCrLf
        sql &= " and aa.lid >= @lid" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql &= " and ag.GDistID = @DistID" & vbCrLf
        End If
        Dim sCmd As New SqlCommand(sql, objconn)
        TIMS.OpenDbConn(objconn)

        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("gid", SqlDbType.VarChar).Value = vGID
            .Parameters.Add("roleid", SqlDbType.Int).Value = Val(sm.UserInfo.RoleID)
            .Parameters.Add("lid", SqlDbType.Int).Value = Val(sm.UserInfo.LID)
            If sm.UserInfo.DistID <> "000" Then
                .Parameters.Add("DistID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
            End If
            dt.Load(.ExecuteReader())
        End With

        lab_Msg3.Visible = True
        lab_Msg3.Text = "查無資料" '.Visible = True
        DataGrid3.Visible = False
        If dt.Rows.Count > 0 Then
            lab_Msg3.Text = "" 'Visible = False
            DataGrid3.Visible = True

            DataGrid3.DataSource = dt
            'DataGrid1.DataKeyField = "GID"
            DataGrid3.DataBind()
        End If

    End Sub

    '回上一頁
    Protected Sub btn_LoadBack3_Click(sender As Object, e As EventArgs) Handles btn_LoadBack3.Click
        Call Clear_Items()
        Call ClsVisible1(1)
    End Sub

    Private Sub DataGrid3_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                e.Item.Cells(0).Text = Convert.ToString(e.Item.ItemIndex + 1)
                'ForeColor="#CCCCFF"
                Dim htmlColor1 As String = "#CCCCFF"
                If Convert.ToString(drv("ISUSED")) = "N" Then
                    For i As Integer = 1 To 3
                        e.Item.Cells(i).ForeColor = ColorTranslator.FromHtml(htmlColor1)
                        TIMS.Tooltip(e.Item.Cells(i), "權限停用")
                    Next

                End If
        End Select
    End Sub


    '代入資料
    Private Sub LoadData3(ByVal vGID As String, ByVal iType As Integer)
        If vGID = "" Then Exit Sub
        'iType:1.編輯 3.顯示(被賦予者)
        Dim sql As String = ""
        sql = "SELECT * FROM AUTH_GROUP WHERE GID= @gid"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("gid", SqlDbType.Int).Value = Val(vGID)
            dt.Load(.ExecuteReader())
        End With

        hidIsSys.Value = "N"
        If dt.Rows.Count > 0 Then
            Dim dr As DataRow = dt.Rows(0)
            hidIsSys.Value = "N"
            If Convert.ToString(dr("gdistid")) = "" Then
                hidIsSys.Value = "Y" '系統取出。
            End If
            Select Case iType
                Case 1
                    'If ddlDistID.Enabled = True Then ddlDistID.SelectedValue = Convert.ToString(dr("gdistid"))
                    If ddlDistID.Enabled Then
                        Common.SetListItem(ddlDistID, Convert.ToString(dr("gdistid")))
                    End If
                    'ddlGtype.SelectedValue = Convert.ToString(dr("gtype"))
                    Common.SetListItem(ddlGtype, Convert.ToString(dr("gtype")))
                    txt_GroupName.Text = Convert.ToString(dr("gname"))
                    txt_GroupNote.Text = Convert.ToString(dr("gnote"))
                    chk_Valid.Checked = If(Convert.ToString(dr("GVALID")) = "1", True, False)'1:啟用/0:停用
                Case 3
                    'If ddlDistID.Enabled = True Then ddlDistID.SelectedValue = Convert.ToString(dr("gdistid"))
                    If ddlDistID3.Enabled Then
                        Common.SetListItem(ddlDistID3, Convert.ToString(dr("gdistid")))
                    End If
                    Common.SetListItem(ddlGtype3, Convert.ToString(dr("gtype")))
                    'ddlGtype3.SelectedValue = Convert.ToString(dr("gtype"))
                    txt_GroupName3.Text = Convert.ToString(dr("gname"))
                    txt_GroupNote3.Text = Convert.ToString(dr("gnote"))
                    'chk_Valid3.Checked = False
                    'If Convert.ToString(dr("gvalid")) Then
                    '    chk_Valid3.Checked = True
                    'End If
            End Select
        End If
    End Sub

    '維護頁面初始化
    Private Sub Clear_Items()
        hidIsSys.Value = "N"

        If ddlDistID.Enabled = True Then ddlDistID.SelectedValue = ""
        If ddlGtype.Enabled = True Then ddlGtype.SelectedValue = ""

        txt_GroupName.Text = ""
        txt_GroupNote.Text = ""
        chk_Valid.Checked = True '1:啟用/0:停用

        list_MainMenu.SelectedIndex = 0
        txtFunName.Text = ""

        hide_GID.Value = ""
        Me.ViewState("mainmenu") = Nothing
    End Sub

    '功能類別對照 取得中文名稱
    'Private Function Get_MainMenuName(ByVal tmpCode As String) As String
    '    Dim rst As String = ""

    '    Select Case UCase(tmpCode)
    '        Case "TC"
    '            rst = "訓練機構管理"
    '        Case "SD"
    '            rst = "學員動態管理"
    '        Case "CP"
    '            rst = "查核/績效管理"
    '        Case "TR"
    '            rst = "訓練需求管理"
    '        Case "CM"
    '            rst = "訓練經費控管"
    '        Case "SYS"
    '            rst = "系統管理"
    '        Case "FAQ"
    '            rst = "問答集"
    '        Case "OB"
    '            rst = "委外訓練管理"
    '        Case "SE"
    '            rst = "技能檢定管理"
    '        Case "EXAM"
    '            rst = "甄試管理"
    '        Case "SV"
    '            rst = "問卷管理"
    '        Case "OO"
    '            rst = "其他系統"
    '        Case Else
    '            rst = tmpCode
    '    End Select

    '    Return rst
    'End Function

    '判斷功能類別是否勾選情況
    Private Function chkData(ByVal intFunID As Integer) As DataTable
        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        Dim sql As String = ""
        sql = "SELECT SECH,ADDS,PRNT FROM AUTH_GROUPFUN  WHERE GID= @gid and funid= @funid"
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("gid", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
            .Parameters.Add("funid", SqlDbType.Int).Value = intFunID
            dt.Load(.ExecuteReader())
        End With
        Return dt

        'Try
        '    sql = "SELECT SECH,ADDS,PRNT FROM AUTH_GROUPFUN  WHERE GID= @gid and funid= @funid"
        '    With sda
        '        .SelectCommand = New SqlCommand(sql, objconn)
        '        .SelectCommand.Parameters.Clear()
        '        .SelectCommand.Parameters.Add("gid", SqlDbType.Int).Value = Convert.ToInt32(hide_GID.Value)
        '        .SelectCommand.Parameters.Add("funid", SqlDbType.Int).Value = intFunID
        '        .Fill(ds, "select")
        '    End With
        '    'If conn.State = ConnectionState.Open Then conn.Close()
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        '    'conn.Close()
        '    If Not sda Is Nothing Then sda.Dispose()
        '    If Not ds Is Nothing Then ds.Dispose()
        'End Try

        'Return ds.Tables("select")
    End Function

    '判斷該群組是否被附予帳號(true=>是,false=>否)
    Private Function chkGroup(ByVal intGID As Int32) As Boolean
        Dim rst As Boolean = False
        Dim sql As String = ""
        sql = "SELECT GID FROM AUTH_GROUPACCT WHERE GID= @gid"
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("gid", SqlDbType.Int).Value = intGID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then rst = True
        Return rst

        'Dim sda As New SqlDataAdapter
        'Dim ds As New DataSet
        'Try
        '    sql = "SELECT GID FROM AUTH_GROUPACCT WHERE GID= @gid"
        '    With sda
        '        .SelectCommand = New SqlCommand(sql, objconn)
        '        .SelectCommand.Parameters.Clear()
        '        .SelectCommand.Parameters.Add("gid", SqlDbType.Int).Value = intGID
        '        .Fill(ds)
        '    End With
        '    'If conn.State = ConnectionState.Open Then conn.Close()
        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'Finally
        '    'conn.Close()
        '    If Not sda Is Nothing Then sda.Dispose()
        '    If Not ds Is Nothing Then ds.Dispose()
        'End Try

        'If ds.Tables(0).Rows.Count > 0 Then
        '    Return True
        'Else
        '    Return False
        'End If

    End Function

    '代入DropDownList資料
    Private Sub Set_DropDownList(ByVal strFlag As String, ByVal obj As DropDownList, ByVal textField As String, ByVal valueField As String)
        Call TIMS.OpenDbConn(objconn)
        Dim sql As String = ""
        sql = "SELECT DISTID,NAME FROM ID_DISTRICT ORDER BY DISTID ASC "
        Dim sCmd As New SqlCommand(sql, objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        With obj
            .DataSource = dt 'ds.Tables(0)
            .DataTextField = textField
            .DataValueField = valueField
            .DataBind()
            .Items.Insert(0, New ListItem("請選擇", ""))
        End With
    End Sub

    '查詢
    Private Sub search1()
        'Try

        'Catch ex As Exception
        '    Common.MessageBox(Me, ex.ToString)
        'End Try

        'Dim dt As New DataTable
        'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        'da.SelectCommand.Parameters.Clear()

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.GID" & vbCrLf
        sql &= " ,a.GDISTID" & vbCrLf
        sql &= " ,a.GTYPE" & vbCrLf
        sql &= " ,a.GNAME" & vbCrLf
        sql &= " ,a.GNOTE" & vbCrLf
        sql &= " ,a.GVALID" & vbCrLf
        sql &= " ,a.GSTATE" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        'sql &= " ,a.MODIFYDATE" & vbCrLf
        'sql += " select a.*,cr.name cname,mo.name mname " & vbCrLf
        sql &= " ,cr.name cname" & vbCrLf
        sql &= " ,mo.name mname " & vbCrLf
        sql &= " from AUTH_GROUP a " & vbCrLf
        sql &= " left join Auth_Account cr on cr.Account=a.CreateAcct " & vbCrLf
        sql &= " left join Auth_Account mo on mo.Account=a.ModifyAcct " & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.GState<>'D' " & vbCrLf

        txt_QGroupName.Text = TIMS.ClearSQM(txt_QGroupName.Text)
        Dim v_ddlQDistID As String = TIMS.GetListValue(ddlQDistID)
        Dim v_ddlQType As String = TIMS.GetListValue(ddlQType)
        Dim v_rdoQVailid As String = TIMS.GetListValue(rdoQVailid) '1:啟用/0:停用

        Dim parms As New Hashtable
        parms.Clear()
        If v_ddlQDistID <> "" Then
            sql &= " and (a.gdistid= @gdistid or a.gdistid is null) " & vbCrLf
            parms.Add("gdistid", v_ddlQDistID)
            'da.SelectCommand.Parameters.Add("gdistid", SqlDbType.VarChar).Value = ddlQDistID.SelectedValue
        End If
        If txt_QGroupName.Text <> "" Then
            sql &= " and a.gname like @gname " & vbCrLf
            parms.Add("gname", "%" & txt_QGroupName.Text & "%")
            'da.SelectCommand.Parameters.Add("gname", SqlDbType.VarChar).Value = "%" & txt_QGroupName.Text & "%"
        End If
        If v_ddlQType <> "" Then
            sql &= " and a.gtype= @gtype " & vbCrLf
            parms.Add("gtype", v_ddlQType)
            'da.SelectCommand.Parameters.Add("gtype", SqlDbType.VarChar).Value = ddlQType.SelectedValue
        End If
        If v_rdoQVailid <> "" Then '1:啟用/0:停用
            sql &= " and a.gvalid= @gvalid " & vbCrLf
            parms.Add("gvalid", v_rdoQVailid)
            'da.SelectCommand.Parameters.Add("gvalid", SqlDbType.VarChar).Value = rdoQVailid.SelectedValue
        End If
        '管理者特殊處理
        'If sm.UserInfo.UserID <> "snoopy" Then
        'End If
        If Not flgROLEIDx0xLIDx0 Then
            If sm.UserInfo.DistID <> "000" Then
                sql &= " and a.gtype <>'0' "
            Else
                sql &= " and a.gtype not in ('1','2') "
            End If
        End If

        sql &= " order by a.gdistid,a.gtype,a.gid asc" & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        'TIMS.Fill(sql, da, dt)

        lab_Msg.Visible = True
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            lab_Msg.Visible = False
            DataGrid1.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "GID"
            DataGrid1.DataBind()
        End If

    End Sub

    '顯示功能list
    Private Sub Get_List3()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        txtFunName.Text = TIMS.ClearSQM(txtFunName.Text)

        '組清單用DataTable
        Dim tmpDT As New DataTable
        tmpDT.Columns.Add(New DataColumn("funid")) '功能流水號
        tmpDT.Columns.Add(New DataColumn("name")) '功能名稱
        tmpDT.Columns.Add(New DataColumn("spage")) '程式名稱
        tmpDT.Columns.Add(New DataColumn("kind")) '功能類別
        tmpDT.Columns.Add(New DataColumn("levels")) '顯示階層
        tmpDT.Columns.Add(New DataColumn("parent")) '目錄
        tmpDT.Columns.Add(New DataColumn("sort")) '排序
        tmpDT.Columns.Add(New DataColumn("memo")) '備註
        tmpDT.Columns.Add(New DataColumn("subs")) '

        '取得功能清單
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select a.funid" & vbCrLf
        sql &= " ,a.name" & vbCrLf
        sql &= " ,a.spage" & vbCrLf
        sql &= " ,a.kind" & vbCrLf
        sql &= " ,a.levels" & vbCrLf
        sql &= " ,dbo.DECODE(CONVERT(varchar, a.levels),'0',CONVERT(varchar, a.funid),CONVERT(varchar, a.parent)) parent" & vbCrLf
        sql &= " ,a.sort" & vbCrLf
        sql &= " ,a.memo" & vbCrLf
        sql &= " ,(case CONVERT(varchar, a.levels) when '0' then (select count(x.funid) from id_function x where x.parent=a.funid) else 0 end) subs" & vbCrLf
        sql &= " ,a.PSORT" & vbCrLf
        sql &= " from V_FUNCTION a" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.valid='Y' " & vbCrLf
        sql &= " and ISNULL(a.FState,' ') not in ('D') " & vbCrLf
        sql &= " and a.kind= @kind" & vbCrLf
        'If Convert.ToString(sm.UserInfo.UserID) <> "snoopy" Then
        '    sql &= " and a.funid in (select distinct funid from AUTH_GROUPacct a   "
        '    sql &= " join AUTH_GROUPFUN b on b.gid=a.gid where a.account='" & sm.UserInfo.UserID & "') "
        '    sql &= " and a.funid not in (select funid from AUTH_GROUPdfun where account='" & sm.UserInfo.UserID & "') "
        'End If
        If txtFunName.Text <> "" Then
            sql &= " and (1!=1" & vbCrLf
            sql &= " OR a.name like '%'+@txtFunName+'%'" & vbCrLf
            sql &= " OR a.spage like '%'+@txtFunName+'%'" & vbCrLf
            sql &= " ) " & vbCrLf
        End If
        If Not flgROLEIDx0xLIDx0 Then
            sql &= " and a.funid in (select distinct funid from AUTH_GROUPacct a   "
            sql &= " join AUTH_GROUPFUN b on b.gid=a.gid where a.account='" & sm.UserInfo.UserID & "') "
            sql &= " and a.funid not in (select funid from AUTH_GROUPdfun where account='" & sm.UserInfo.UserID & "') "
        End If

        'sql += " order by a.kind,dbo.DECODE(CONVERT(varchar, a.levels),'0',CONVERT(varchar, a.funid),CONVERT(varchar, a.parent)) ,a.levels,a.sort"
        sql &= " order by a.kind,a.PSORT ,a.levels,a.sort"
        Dim sCmd3 As New SqlCommand(sql, objconn)

        '取得功能種類
        sql = "SELECT DISTINCT KIND FROM ID_FUNCTION ORDER BY KIND"
        Dim sCmd3k As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        If list_MainMenu.SelectedIndex <> 0 Then
            Dim dt3 As New DataTable
            With sCmd3
                .Parameters.Clear()
                .Parameters.Add("kind", SqlDbType.VarChar).Value = list_MainMenu.SelectedValue
                If txtFunName.Text <> "" Then
                    .Parameters.Add("txtFunName", SqlDbType.VarChar).Value = txtFunName.Text
                End If
                dt3.Load(.ExecuteReader())
            End With
            Call sUtl_dt3toTmpDt3(dt3, tmpDT)
        Else
            For i As Integer = 0 To arrFun.Length - 1
                Dim dt3 As New DataTable
                With sCmd3
                    .Parameters.Clear()
                    .Parameters.Add("kind", SqlDbType.VarChar).Value = arrFun(i)
                    If txtFunName.Text <> "" Then
                        .Parameters.Add("txtFunName", SqlDbType.VarChar).Value = txtFunName.Text
                    End If
                    dt3.Load(.ExecuteReader())
                End With
                Call sUtl_dt3toTmpDt3(dt3, tmpDT)
            Next
        End If

        '取得功能種類
        Dim dt3k As New DataTable
        With sCmd3k
            .Parameters.Clear()
            dt3k.Load(.ExecuteReader())
        End With

        For i As Integer = 0 To dt3k.Rows.Count - 1
            Dim dr3k As DataRow = dt3k.Rows(i)
            If Convert.ToString(dr3k("kind")) <> "" Then
                Dim ff3 As String = "kind='" & Convert.ToString(dr3k("kind")) & "'"
                Me.ViewState(Convert.ToString(dr3k("kind"))) = tmpDT.Select(ff3).Length
            End If
        Next

        Me.ViewState("mmname") = ""

        DataGrid2.DataSource = tmpDT
        DataGrid2.DataKeyField = "FunID"
        DataGrid2.DataBind()
    End Sub

    '將dt3資料 塞進 tmpDT
    Sub sUtl_dt3toTmpDt3(ByRef dt3 As DataTable, ByRef tmpDT As DataTable)
        For j As Integer = 0 To dt3.Rows.Count - 1
            Dim dr As DataRow = tmpDT.NewRow
            tmpDT.Rows.Add(dr)

            dr("funid") = dt3.Rows(j)("funid")
            dr("name") = dt3.Rows(j)("name")
            dr("spage") = dt3.Rows(j)("spage")
            dr("kind") = dt3.Rows(j)("kind")
            dr("levels") = dt3.Rows(j)("levels")
            dr("parent") = dt3.Rows(j)("parent")
            dr("sort") = dt3.Rows(j)("sort")
            dr("memo") = dt3.Rows(j)("memo")
            dr("subs") = dt3.Rows(j)("subs")
        Next
    End Sub

    '查詢3
    Protected Sub btnSearch3_Click(sender As Object, e As EventArgs) Handles btnSearch3.Click
        Call Get_List3()
    End Sub

    Protected Sub DataGrid3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid3.SelectedIndexChanged

    End Sub
End Class


