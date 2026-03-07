Partial Class SYS_06_004
    Inherits AuthBasePage

    Const cst_addnew As String = "addnew" '新增
    Const cst_update As String = "update" '修改

    'btn
    Const cst_view As String = "view" '檢視
    Const cst_edit As String = "edit" '編輯修改

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
        PageControler1.PageDataGrid = Me.DataGrid1

        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '    Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        'Else
        '   'Dim FunDr As DataRow
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")
        '    FunDr = FunDrArray(0)
        '    If FunDr("Sech") = 1 Then
        '        btnQuery.Enabled = True
        '    Else
        '        btnQuery.Enabled = False
        '        TIMS.Tooltip(btnQuery, "沒有搜尋權限")
        '    End If
        'End If

        If Not Me.IsPostBack Then
            Call Create1()

            DataGridTable1.Visible = False '預設搜尋資料不顯示

            panelSearch.Visible = True '搜尋功能啟動
            PanelEdit1.Visible = False '修改功能關閉
        End If
    End Sub

    Sub Create1()
        '取得訓練計畫
        'TPlan = TIMS.Get_TPlan(TPlan, , 1)
        Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn, 1)

        chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
        chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
        chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"

    End Sub

    Sub sUtl_ClearPanelEdit1()
        hidslid.Value = ""
        btnSave1.Visible = True '顯示 儲存鈕

        STDate.Text = ""
        PYears.Text = ""
        PMoneys.Text = ""

        Call sUtl_ClsTPlanSelected()
        'For i As Integer = 0 To cblobjecttype.Items.Count - 1
        '    cblobjecttype.Items(i).Selected = False
        'Next
    End Sub

    '顯示修改區
    Sub sUtl_ShowPanelEdit1(ByVal sSearchW As String)
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        '修改
        'sSearchW = e.CommandArgument
        Me.hidslid.Value = TIMS.GetMyValue(sSearchW, "slid")
        'btnSave1
        btnSave1.Visible = True '顯示 儲存鈕
        If TIMS.GetMyValue(sSearchW, "btn") = cst_view Then
            btnSave1.Visible = False '隱藏
        End If

        Dim dr As DataRow = search2(Me.hidslid.Value) '搜尋顯示
        If Not dr Is Nothing Then
            STDate.Text = dr("STDate").ToString
            PYears.Text = dr("PYears").ToString
            PMoneys.Text = dr("PMoneys").ToString
            If dr("isUsed").ToString <> "" Then
                Common.SetListItem(rblisUsed, dr("isUsed").ToString)
            End If

            Call sUtl_ShowTPlanSelected(Me.hidslid.Value)
        End If
    End Sub

    '設定新增初值
    Sub sUtl_InsertPanelEdit1()
        'Me.hidslid.Value = ""
        '搜尋->新增
        If STDate_s.Text <> "" Then
            STDate.Text = STDate_s.Text
        End If
        '搜尋->新增
        If rblisUsed_s.SelectedValue <> "" Then
            'rblisUsed_s.SelectedValue 
            Common.SetListItem(rblisUsed, rblisUsed_s.SelectedValue)
        End If

    End Sub

    '設定 ViewState Value
    Sub sUtl_ViewStateValue()

        '發送範圍 (全系統 / 指定計畫)
        Me.ViewState("STDate") = ""
        Me.ViewState("isUsed") = ""
        If STDate_s.Text <> "" Then
            Me.ViewState("STDate") = STDate_s.Text
        End If
        If rblisUsed_s.SelectedValue <> "" Then
            Me.ViewState("isUsed") = rblisUsed_s.SelectedValue
        End If

    End Sub

    'SQL
    Sub search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " SELECT " & vbCrLf
        Sql += " a.slid  /*PK*/ " & vbCrLf
        Sql += " ,CONVERT(varchar, a.STDate, 111) STDate" & vbCrLf
        Sql += " ,a.isUsed" & vbCrLf
        Sql += " ,a.PYears" & vbCrLf
        Sql += " ,a.PMoneys" & vbCrLf
        'Sql += " ,ModifyAcct" & vbCrLf
        'Sql += " ,ModifyDate" & vbCrLf
        Sql += " FROM Sys_Supply a" & vbCrLf
        Sql += " WHERE 1=1" & vbCrLf
        If Me.ViewState("STDate") <> "" Then
            Sql += " and a.STDate = " & TIMS.To_date(Me.ViewState("STDate")) & vbCrLf
        End If
        If Me.ViewState("isUsed") <> "" Then
            Sql += " and a.isUsed = '" & Me.ViewState("isUsed") & "'" & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(Sql, objconn)

        DataGridTable1.Visible = False
        msg.Text = "查無資料"
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            DataGridTable1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    'SQL return datarow
    Function search2(ByVal slid As String) As DataRow
        Dim rst As DataRow = Nothing

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " SELECT " & vbCrLf
        Sql += " a.slid  /*PK*/ " & vbCrLf
        Sql += " ,CONVERT(varchar, a.STDate, 111) STDate" & vbCrLf
        Sql += " ,a.isUsed" & vbCrLf
        Sql += " ,a.PYears" & vbCrLf
        Sql += " ,a.PMoneys" & vbCrLf
        Sql += " FROM Sys_Supply a" & vbCrLf
        Sql += " WHERE 1=1" & vbCrLf
        Sql += " and a.slid  = '" & slid & "'" & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(Sql, objconn)
        If dt.Rows.Count > 0 Then rst = dt.Rows(0)

        Return rst
    End Function

    '儲存 動作
    Function sUtl_SaveData1(ByVal sSearchW As String) As Integer
        Dim rst As Integer = 0 '0:異常儲存 1:正常儲存

        Dim sType As String  '儲存狀態 (cst_addnew / cst_update )
        sType = cst_update '修改
        If hidslid.Value = "" Then
            sType = cst_addnew '新增
        End If
        Dim sql As String = ""
        '確認
        'sSearchW = e.CommandArgument
        Dim sTPlanID As String = TIMS.GetMyValue(sSearchW, "TPlanID")

        Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
        Select Case sType
            Case cst_addnew '新增
                sql = "" & vbCrLf
                sql += "  /* IDENTITY(1,1): slid */ " & vbCrLf
                sql += " INSERT INTO Sys_Supply(" & vbCrLf
                sql += " STDate" & vbCrLf
                sql += "  ,isUsed" & vbCrLf
                sql += "  ,PYears" & vbCrLf
                sql += "  ,PMoneys" & vbCrLf
                sql += "  ,ModifyAcct" & vbCrLf
                sql += "  ,ModifyDate" & vbCrLf
                sql += " ) VALUES (" & vbCrLf
                sql += " @STDate" & vbCrLf
                sql += " ,@isUsed" & vbCrLf
                sql += " ,@PYears" & vbCrLf
                sql += " ,@PMoneys" & vbCrLf
                sql += " ,@ModifyAcct" & vbCrLf
                sql += " ,getdate() " & vbCrLf
                sql += " ) " & vbCrLf
                'sql += " select @@identity as slid" & vbCrLf
                da.SelectCommand.CommandText = sql
                da.SelectCommand.Parameters.Clear()
                da.SelectCommand.Parameters.Add("STDate", SqlDbType.VarChar).Value = IIf(STDate.Text <> "", STDate.Text, "")
                da.SelectCommand.Parameters.Add("isUsed", SqlDbType.VarChar).Value = rblisUsed.SelectedValue
                da.SelectCommand.Parameters.Add("PYears", SqlDbType.VarChar).Value = CInt(PYears.Text)
                da.SelectCommand.Parameters.Add("PMoneys", SqlDbType.VarChar).Value = CInt(PMoneys.Text)
                da.SelectCommand.Parameters.Add("ModifyAcct", SqlDbType.NVarChar).Value = sm.UserInfo.UserID

                If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                da.SelectCommand.ExecuteNonQuery()

                hidslid.Value = DbAccess.GetNewId(objconn, "SYS_SUPPLY_SLID_SEQ")
                'hidslid.Value = da.SelectCommand.ExecuteScalar() '.ExecuteNonQuery()

                If sTPlanID <> "" Then
                    sql = "" & vbCrLf
                    sql += " INSERT INTO Sys_SupplyList( slid,TPlanID ) " & vbCrLf
                    sql += " SELECT " & hidslid.Value & " slid,TPlanID  " & vbCrLf
                    sql += " FROM KEY_PLAN" & vbCrLf
                    sql += " WHERE TPlanID IN (" & sTPlanID & ")" & vbCrLf
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                    da.SelectCommand.ExecuteNonQuery()
                End If

                rst = 1 '儲存成功

            Case Else 'cst_update '修改
                sql = "" & vbCrLf
                sql += " UPDATE Sys_Supply SET" & vbCrLf
                sql += " STDate= @STDate" & vbCrLf
                sql += "  ,isUsed= @isUsed" & vbCrLf
                sql += "  ,PYears= @PYears" & vbCrLf
                sql += "  ,PMoneys= @PMoneys" & vbCrLf
                sql += "  ,ModifyAcct= @ModifyAcct" & vbCrLf
                sql += "  ,ModifyDate= getdate()" & vbCrLf
                sql += " WHERE 1=1 " & vbCrLf
                sql += " AND slid= @slid" & vbCrLf

                da.SelectCommand.CommandText = sql
                da.SelectCommand.Parameters.Clear()
                da.SelectCommand.Parameters.Add("slid", SqlDbType.VarChar).Value = Me.hidslid.Value
                da.SelectCommand.Parameters.Add("STDate", SqlDbType.DateTime).Value = IIf(STDate.Text <> "", TIMS.Cdate2(STDate.Text), "")
                da.SelectCommand.Parameters.Add("isUsed", SqlDbType.VarChar).Value = rblisUsed.SelectedValue
                da.SelectCommand.Parameters.Add("PYears", SqlDbType.VarChar).Value = CInt(PYears.Text)
                da.SelectCommand.Parameters.Add("PMoneys", SqlDbType.VarChar).Value = CInt(PMoneys.Text)
                da.SelectCommand.Parameters.Add("ModifyAcct", SqlDbType.NVarChar).Value = sm.UserInfo.UserID

                If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                da.SelectCommand.ExecuteNonQuery()

                If sTPlanID <> "" Then
                    sql = "" & vbCrLf
                    sql += " DELETE Sys_SupplyList WHERE slid =" & hidslid.Value & vbCrLf '刪除動作
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                    da.SelectCommand.ExecuteNonQuery()

                    sql = " INSERT INTO Sys_SupplyList( slid,TPlanID ) " & vbCrLf
                    sql += " SELECT " & hidslid.Value & " slid,TPlanID  " & vbCrLf
                    sql += " FROM KEY_PLAN" & vbCrLf
                    sql += " WHERE TPlanID IN (" & sTPlanID & ")" & vbCrLf
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                    da.SelectCommand.ExecuteNonQuery()
                Else
                    sql = "" & vbCrLf
                    sql += " DELETE Sys_SupplyList WHERE slid =" & hidslid.Value & vbCrLf '刪除動作
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                    da.SelectCommand.ExecuteNonQuery()
                End If

                rst = 1 '儲存成功

        End Select

        Return rst
    End Function

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Call sUtl_ViewStateValue()  '設定 ViewState Value

        Call search1()
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sSearchW As String = ""

        Select Case e.CommandName
            Case "View1"
                '修改
                sSearchW = e.CommandArgument
                Me.hidslid.Value = TIMS.GetMyValue(sSearchW, "slid")

                Call sUtl_ClearPanelEdit1()

                Call sUtl_ShowPanelEdit1(sSearchW)

            Case "Edit1"
                '修改
                sSearchW = e.CommandArgument
                Me.hidslid.Value = TIMS.GetMyValue(sSearchW, "slid")

                Call sUtl_ClearPanelEdit1()

                Call sUtl_ShowPanelEdit1(sSearchW)

            Case "Delete1"
                '刪除
                sSearchW = e.CommandArgument
                Me.hidslid.Value = TIMS.GetMyValue(sSearchW, "slid")
                Try
                    '刪除sql
                    Dim sql As String = ""
                    sql = ""
                    sql += " DELETE Sys_Supply " & vbCrLf
                    sql += " WHERE 1=1 " & vbCrLf
                    sql += " AND slid= @slid" & vbCrLf

                    Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    da.SelectCommand.Parameters.Add("slid", SqlDbType.VarChar).Value = hidslid.Value
                    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                    da.SelectCommand.ExecuteNonQuery()

                    Call sUtl_ClearPanelEdit1()

                    Call search1()
                Catch ex As Exception
                    Common.MessageBox(Me, ex.ToString)
                End Try

        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem

                Dim lbPYMs As Label = e.Item.FindControl("lbPYMs")
                Dim lbPlanNames As Label = e.Item.FindControl("lbPlanNames")

                Dim btnView1 As LinkButton = e.Item.FindControl("btnView1")
                Dim btnEdit1 As LinkButton = e.Item.FindControl("btnEdit1")
                Dim btnDelete1 As LinkButton = e.Item.FindControl("btnDelete1")

                '序號
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex

                '顯示文字
                lbPYMs.Text = drv("PYears").ToString & "年 " & drv("PMoneys").ToString & "萬 "

                lbPlanNames.Text = sUtl_ShowPlanName(drv("slid"))

                Dim cmdArg As String = ""
                cmdArg = ""
                cmdArg &= "&slid=" & Convert.ToString(drv("slid"))

                btnView1.CommandArgument = cmdArg & "&btn=" & cst_view
                btnEdit1.CommandArgument = cmdArg & "&btn=" & cst_edit
                btnDelete1.CommandArgument = cmdArg
                btnDelete1.Attributes("onclick") = "return confirm('此動作會刪除此筆資料，是否確定刪除?');"

        End Select
    End Sub

    Sub sUtl_ClsTPlanSelected()
        For Each objitem As ListItem In chkTPlanID0.Items '大計畫
            objitem.Selected = False
        Next
        For Each objitem As ListItem In chkTPlanID1.Items '大計畫
            objitem.Selected = False
        Next
        For Each objitem As ListItem In chkTPlanIDX.Items '大計畫
            objitem.Selected = False
        Next
    End Sub

    Sub sUtl_ShowTPlanSelected(ByVal slid As String)
        'Dim rst As String
        Dim sql As String
        Dim dt As DataTable
        sql = "SELECT slid ,TPlanID FROM Sys_SupplyList   WHERE slid =" & slid
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each objitem As ListItem In chkTPlanID0.Items '大計畫
            If dt.Select("TPlanID='" & objitem.Value & "'").Length > 0 Then
                objitem.Selected = True
            End If
        Next
        For Each objitem As ListItem In chkTPlanID1.Items '大計畫
            If dt.Select("TPlanID='" & objitem.Value & "'").Length > 0 Then
                objitem.Selected = True
            End If
        Next
        For Each objitem As ListItem In chkTPlanIDX.Items '大計畫
            If dt.Select("TPlanID='" & objitem.Value & "'").Length > 0 Then
                objitem.Selected = True
            End If
        Next
    End Sub

    Function sUtl_ShowPlanName(ByVal slid As String) As String
        Dim rst As String = ""
        Dim sql As String = ""
        Dim dt As DataTable
        sql = "SELECT slid ,TPlanID FROM Sys_SupplyList   WHERE slid =" & slid
        dt = DbAccess.GetDataTable(sql, objconn)
        For Each objitem As ListItem In chkTPlanID0.Items '大計畫
            If dt.Select("TPlanID='" & objitem.Value & "'").Length > 0 Then
                If rst <> "" Then rst += ","
                rst += objitem.Text
            End If
        Next
        For Each objitem As ListItem In chkTPlanID1.Items '大計畫
            If dt.Select("TPlanID='" & objitem.Value & "'").Length > 0 Then
                If rst <> "" Then rst += ","
                rst += objitem.Text
            End If
        Next
        For Each objitem As ListItem In chkTPlanIDX.Items '大計畫
            If dt.Select("TPlanID='" & objitem.Value & "'").Length > 0 Then
                If rst <> "" Then rst += ","
                rst += objitem.Text
            End If
        Next
        Return rst
    End Function

    '新增
    Private Sub btnAddnew_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAddnew.Click
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        Call sUtl_ClearPanelEdit1()

        Call sUtl_InsertPanelEdit1()
    End Sub

    '回上一頁
    Private Sub btnBack1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBack1.Click
        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉
    End Sub

    '儲存
    Private Sub btnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSave1.Click
        Dim TPlanID As String = ""
        TPlanID = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX, 0)

        ''-------------------------------儲存前檢核-------------------------------
        'Dim sErrmsg As String = ""
        'sErrmsg = ""
        'If txtSubject.Text = "" Then
        '    sErrmsg += "請輸入主題" & vbCrLf
        'End If
        'If SendDate.Text = "" Then
        '    sErrmsg += "請輸入發送日期" & vbCrLf
        'End If
        'If PlanID = "" Then
        '    sErrmsg += "請選擇發送範圍之計畫" & vbCrLf
        'End If
        'If objecttype = "" Then
        '    sErrmsg += "請選擇發送對象" & vbCrLf
        'End If
        'If txtcontents.Text = "" Then
        '    sErrmsg += "請輸入內容" & vbCrLf
        'End If
        'If sErrmsg <> "" Then
        '    Common.MessageBox(Me, sErrmsg)
        '    Exit Sub
        'End If
        ''-------------------------------儲存前檢核-------------------------------

        Dim cmdArg As String = ""
        cmdArg = ""
        cmdArg &= "&TPlanID=" & TPlanID
        'cmdArg &= "&objecttype=" & objecttype 'sUtl_GetobjectType(cblobjecttype, cst_save)

        If sUtl_SaveData1(cmdArg) = 1 Then
            '正常儲存
            Call search1()
        End If

    End Sub


End Class
