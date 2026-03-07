Partial Class PS_01_001
    Inherits AuthBasePage

    ''PS_COMMONREPORT

    Const cst_errmsg2 As String = "該使用者/計畫無此功能，請重新選擇計畫!!"
    Const cst_tab2 As String = "&nbsp;&nbsp;&nbsp;　"
    Const cst_t請選擇 As String = "==請選擇=="
    Const cst_v請選擇 As String = "==請選擇=="
    Const cst_請選擇3 As String = TIMS.cst_ddl_PleaseChoose3

    Dim dtPSDATA As DataTable
    Dim dt_group As DataTable
    Dim dt_group_F1 As DataTable

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

        If Not IsPostBack Then
            list_MainMenu2 = TIMS.Get_ddlFunction(list_MainMenu2, 2)
            '查詢 SQL
            Call sSearch1("")
        End If

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim lab_MainMenu2 As Label = e.Item.FindControl("lab_MainMenu2")
                Dim lab_MainMenu3 As Label = e.Item.FindControl("lab_MainMenu3")
                Dim lab_FunName As Label = e.Item.FindControl("lab_FunName")
                Dim labMemo As Label = e.Item.FindControl("labMemo")
                Dim chk_Enable As CheckBox = e.Item.FindControl("chk_Enable") '選取方塊
                'Dim txtFunID As TextBox = e.Item.FindControl("txtFunID") '記錄FunID
                Dim Hid_FunID As HiddenField = e.Item.FindControl("Hid_FunID")

                e.Item.Cells(1).ToolTip = drv("FunID")
                Hid_FunID.Value = Convert.ToString(drv("FunID"))
                lab_MainMenu2.Text = TIMS.Get_MainMenuName(Convert.ToString(drv("Kind")))
                Dim FF3 As String = "FUNID=" & Convert.ToString(drv("PARENT"))
                If dt_group_F1.Select(FF3).Length > 0 Then
                    lab_MainMenu3.Text = Convert.ToString(dt_group_F1.Select(FF3)(0)("NAME"))
                End If
                labMemo.Text = Convert.ToString(drv("Memo"))

                lab_FunName.Text = Convert.ToString(drv("Name"))
                If Convert.ToString(drv("levels")) = "1" Then
                    lab_FunName.Text = cst_tab2 & Convert.ToString(drv("Name"))
                End If

                chk_Enable.Checked = ChkData1(dtPSDATA, Convert.ToInt32(drv("funid")))

        End Select
    End Sub


    ''' <summary>
    ''' 判斷計畫中功能類別是否勾選(true=>勾選,false=>否)
    ''' </summary>
    ''' <param name="dtPSDATA"></param>
    ''' <param name="intFunID"></param>
    ''' <returns></returns>
    Function ChkData1(ByRef dtPSDATA As DataTable, ByVal intFunID As Integer) As Boolean
        Dim Rst As Boolean = False
        Dim ff As String = ""
        ff = " FUNID=" & intFunID
        If dtPSDATA.Select(ff).Length > 0 Then Rst = True
        Return Rst
    End Function


    ''' <summary>
    ''' 查詢 SQL / list_MainMenu3 設定
    ''' </summary>
    ''' <param name="sMenu3Val"></param>
    Sub sSearch1(ByVal sMenu3Val As String)
        dtPSDATA = TIMS.Get_PSDATAdt(sm, objconn)

        dt_group = TIMS.sGet_CanUseSchDt(objconn)
        dt_group_F1 = TIMS.dv2dt(dt_group.DefaultView)
        Dim blnGroupF As Boolean = True '查詢有資料。
        If dt_group Is Nothing Then
            blnGroupF = False '查詢無資料。
        Else
            dt_group.DefaultView.RowFilter = "Valid='Y' AND ISREPORT='Y'"
            dt_group.DefaultView.Sort = "KIND,LEVELS,PARENT,NEWSORT"
            dt_group = TIMS.dv2dt(dt_group.DefaultView)
            If dt_group.Rows.Count = 0 Then blnGroupF = False '查詢無資料。
        End If
        If Not blnGroupF Then
            sm.LastErrorMessage = cst_errmsg2
            Exit Sub
        End If


        Try
            'list_MainMenu3 設定
            Dim FF3 As String = ""
            If sMenu3Val = "" Then
                With list_MainMenu3
                    .Items.Clear()
                    .Items.Insert(0, New ListItem("無", ""))
                End With

                If list_MainMenu2.SelectedValue <> "" AndAlso list_MainMenu2.SelectedIndex <> 0 Then
                    With list_MainMenu3
                        '.Items.Clear()
                        dt_group.DefaultView.RowFilter = "KIND = '" & list_MainMenu2.SelectedValue & "' AND Valid='Y' AND ISREPORT='Y' AND SPAGE IS NULL"
                        dt_group.DefaultView.Sort = "KIND,LEVELS,PARENT,NEWSORT"
                        Dim dtM3 As DataTable = TIMS.dv2dt(dt_group.DefaultView)
                        If dtM3.Rows.Count > 0 Then
                            .DataSource = dtM3
                            .DataValueField = "FUNID"
                            .DataTextField = "NAME"
                            .DataBind()
                            .Items.Insert(0, New ListItem("全部", "ALL"))
                        Else
                        End If
                    End With
                End If
            End If

            DataGrid1.Visible = False
            If list_MainMenu2.SelectedValue = "" Then
                If dt_group.Rows.Count > 0 Then
                    DataGrid1.Visible = True
                    DataGrid1.DataSource = dt_group
                    DataGrid1.DataBind()
                Else
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
            Else
                Select Case list_MainMenu3.SelectedValue
                    Case "", "ALL"
                        dt_group.DefaultView.RowFilter = "KIND = '" & list_MainMenu2.SelectedValue & "' AND Valid='Y' AND ISREPORT='Y' AND SPAGE IS NOT NULL"
                        dt_group.DefaultView.Sort = "KIND,LEVELS,PARENT,NEWSORT"
                    Case Else
                        dt_group.DefaultView.RowFilter = "KIND = '" & list_MainMenu2.SelectedValue & "' AND [Parent]='" & list_MainMenu3.SelectedValue & "' AND Valid='Y' AND ISREPORT='Y' AND SPAGE IS NOT NULL AND LEVELS<>0"
                        dt_group.DefaultView.Sort = "KIND,LEVELS,PARENT,NEWSORT"
                End Select
                Dim dt_group_FF As DataTable = TIMS.dv2dt(dt_group.DefaultView)
                If dt_group_FF.Rows.Count > 0 Then
                    DataGrid1.Visible = True
                    DataGrid1.DataSource = dt_group_FF
                    DataGrid1.DataBind()
                Else
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If
            End If

        Catch ex As Exception
            Common.MessageBox(Me, ex.ToString)
        End Try

    End Sub

    ''' <summary>
    ''' 儲存
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub btn_Save2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save2.Click

        Dim sql As String = ""
        sql = "SELECT 1 FROM PS_COMMONREPORT where ACCOUNT =@ACCOUNT and TPlanID=@TPlanID and FunID=@FunID"
        Dim sCmd As New SqlCommand(sql, objconn)
        sql = ""
        sql &= " INSERT INTO PS_COMMONREPORT(PSRID,ACCOUNT,TPlanID,FunID,SORT,ISUSED,ModifyAcct,ModifyDate) "
        sql &= " values(@PSRID,@ACCOUNT,@TPlanID,@FunID,@SORT,@ISUSED,@ModifyAcct, GETDATE()) "
        Dim iCmd As New SqlCommand(sql, objconn)
        sql = ""
        sql &= " UPDATE PS_COMMONREPORT"
        sql &= " SET SORT = @SORT, ISUSED = @ISUSED, ModifyAcct= @ModifyAcct, MODIFYDATE=GETDATE()"
        sql &= " where ACCOUNT =@ACCOUNT and TPlanID=@TPlanID and FunID=@FunID"
        Dim uCmd As New SqlCommand(sql, objconn)
        sql = ""
        sql &= " DELETE PS_COMMONREPORT"
        sql &= " where ACCOUNT =@ACCOUNT and TPlanID=@TPlanID and FunID=@FunID"
        Dim dCmd As New SqlCommand(sql, objconn)

        Dim parms As New Hashtable

        Dim iRow As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim chkEnable As CheckBox = eItem.FindControl("chk_Enable")
            Dim Hid_FunID As HiddenField = eItem.FindControl("Hid_FunID")
            parms.Clear()
            parms.Add("ACCOUNT", sm.UserInfo.UserID)
            parms.Add("TPlanID", sm.UserInfo.TPlanID)
            parms.Add("FunID", Val(Hid_FunID.Value))
            Dim dt1 As DataTable = DbAccess.GetDataTable(sCmd.CommandText, objconn, parms)

            If chkEnable.Checked Then
                If dt1.Rows.Count = 0 Then
                    iRow += 1
                    '無資料新增
                    Dim iPSRID As Integer = DbAccess.GetNewId(objconn, "PS_COMMONREPORT_PSRID_SEQ,PS_COMMONREPORT,PSRID")
                    parms.Clear()
                    parms.Add("PSRID", iPSRID)
                    parms.Add("ACCOUNT", sm.UserInfo.UserID)
                    parms.Add("TPlanID", sm.UserInfo.TPlanID)
                    parms.Add("FunID", Val(Hid_FunID.Value))
                    parms.Add("SORT", Convert.DBNull)
                    parms.Add("ISUSED", "Y")
                    parms.Add("ModifyAcct", sm.UserInfo.UserID)
                    DbAccess.ExecuteNonQuery(iCmd.CommandText, objconn, parms)
                Else
                    iRow += 1
                    '存在即update
                    parms.Clear()
                    parms.Add("SORT", Convert.DBNull)
                    parms.Add("ISUSED", "Y")
                    parms.Add("ModifyAcct", sm.UserInfo.UserID)
                    parms.Add("ACCOUNT", sm.UserInfo.UserID)
                    parms.Add("TPlanID", sm.UserInfo.TPlanID)
                    parms.Add("FunID", Val(Hid_FunID.Value))
                    DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, parms)
                End If
            Else
                If dt1.Rows.Count <> 0 Then
                    iRow += 1
                    '存在即update
                    parms.Clear()
                    parms.Add("SORT", Convert.DBNull)
                    parms.Add("ISUSED", Convert.DBNull)
                    parms.Add("ModifyAcct", sm.UserInfo.UserID)
                    parms.Add("ACCOUNT", sm.UserInfo.UserID)
                    parms.Add("TPlanID", sm.UserInfo.TPlanID)
                    parms.Add("FunID", Val(Hid_FunID.Value))
                    DbAccess.ExecuteNonQuery(uCmd.CommandText, objconn, parms)
                    'parms.Clear()
                    'parms.Add("ACCOUNT", sm.UserInfo.UserID)
                    'parms.Add("TPlanID", sm.UserInfo.TPlanID)
                    'parms.Add("FunID", Val(Hid_FunID.Value))
                    'DbAccess.ExecuteNonQuery(dCmd.CommandText, objconn, parms)
                End If
            End If
        Next

        If iRow = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Common.MessageBox(Me, "儲存成功!")

    End Sub

    Private Sub list_MainMenu2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_MainMenu2.SelectedIndexChanged
        Call sSearch1("")
    End Sub

    Private Sub list_MainMenu3_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles list_MainMenu3.SelectedIndexChanged
        Call sSearch1(list_MainMenu3.SelectedValue)
    End Sub

End Class


