'Imports System.Reflection

Public Class SYS_06_005
    Inherits AuthBasePage

    'SELECT * FROM Key_BusinessKeys--'關鍵字。
    'select * from v_Key_BusinessKeys where years ='2013' --產業別 與 關鍵字對照表。
    'select * from KEY_BUSINESSKEYS where DEPID IN ('06','07','08') --'關鍵字。
    'select * from Key_Depot where years ='2013' --2013年度產業。
    'select * from KEY_BUSINESS where DEPID IN ('06','07','08') --各產業 產業別。
    ' select * from v_Depot06 ORDER BY KID --2013-六大新興產業
    ' select * from v_Depot10 ORDER BY KID --2013-十大重點服務業
    ' select * from v_Depot04 ORDER BY KID --2013-四大新興智慧型產業

    Const cst_DepID06 As String = "07" 'v_Depot06
    Const cst_DepID10 As String = "08" 'v_Depot10
    Const cst_DepID04 As String = "06" 'v_Depot04

    Const cst_multiaddnew As String = "multiaddnew" '多筆新增
    Const cst_addnew As String = "addnew" '新增
    Const cst_update As String = "update" '修改

    'btn
    'Const cst_view As String = "view" '檢視
    Const cst_edit As String = "edit" '編輯修改

    Const cst_ItemEdit1 As String = "Edit1"  '編輯修改
    Const cst_ItemDelete1 As String = "Delete1"  '編輯刪除

    Const cst_DelOld_Y As String = "Y"
    Const cst_DelOld_N As String = "N"

    '"Edit1"
    Dim objconn As SqlConnection

    'Private m_MyAssembly As [Assembly]

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
        '    btnQuery.Enabled = False
        '    If FunDrArray.Length > 0 Then
        '        FunDr = FunDrArray(0)
        '        If FunDr("Sech") = 1 Then btnQuery.Enabled = True
        '    End If
        '    If Not btnQuery.Enabled Then
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
        '新年度2013
        'Get_KeyBusiness(KID_6, cst_DepID06)
        'Get_KeyBusiness(KID_10, cst_DepID10)
        'Get_KeyBusiness(KID_4, cst_DepID04)

        ddlKID04.Items.Clear()
        ddlKID06.Items.Clear()
        ddlKID10.Items.Clear()

        Const Cst_EmptySelValue As String = "==無==" 'TIMS.cst_ddl_PleaseChoose3
        Dim sql As String = ""
        sql = "SELECT KID,KNAME FROM V_DEPOT04 ORDER BY KID"
        DbAccess.MakeListItem(ddlKID04, sql, objconn)
        ddlKID04.Items.Insert(0, New ListItem(Cst_EmptySelValue, ""))

        sql = "SELECT KID,KNAME FROM V_DEPOT06 ORDER BY KID"
        DbAccess.MakeListItem(ddlKID06, sql, objconn)
        ddlKID06.Items.Insert(0, New ListItem(Cst_EmptySelValue, ""))

        sql = "SELECT KID,KNAME FROM V_DEPOT10 ORDER BY KID"
        DbAccess.MakeListItem(ddlKID10, sql, objconn)
        ddlKID10.Items.Insert(0, New ListItem(Cst_EmptySelValue, ""))

    End Sub

#Region "NO USE"
    ''產業別鍵詞
    'Function Get_KeyBusiness(ByVal obj As ListControl, ByVal DepID As String) As ListControl
    '    Dim dt As DataTable = Nothing

    '    Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
    '    Dim Sql As String = ""
    '    Sql = "" & vbCrLf
    '    sql &= " select a.KNAME ,a.KID ,a.SeqNo ,a.DepID,a.Status" & vbCrLf
    '    sql &= " from Key_Business a" & vbCrLf
    '    sql &= " join Key_Depot b on b.Depid=a.DepID" & vbCrLf
    '    sql &= " where 1=1" & vbCrLf
    '    sql &= " AND a.Status is NULL " & vbCrLf
    '    da.SelectCommand.Parameters.Clear()
    '    If DepID <> "" Then
    '        sql &= " AND a.DepID= @DepID "
    '        da.SelectCommand.Parameters.Add("@DepID", SqlDbType.VarChar).Value = DepID
    '    End If
    '    TIMS.Fill(Sql, da, dt)
    '    With obj
    '        .Items.Clear()
    '        .DataSource = dt
    '        .DataTextField = "KNAME"
    '        .DataValueField = "KID" '"SeqNo"
    '        .DataBind()
    '        If TypeOf obj Is DropDownList Then
    '            .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
    '        End If
    '        If TypeOf obj Is CheckBoxList Then
    '            .Items.Insert(0, New ListItem("全部", ""))
    '        End If
    '    End With
    '    Return obj
    'End Function
#End Region

    Sub sUtl_ClearPanelEdit1()
        cb_delKID.Visible = False
        hidKeysID.Value = ""

        ddlKID06.SelectedIndex = -1
        ddlKID10.SelectedIndex = -1
        ddlKID04.SelectedIndex = -1
        tKeyNAME.Text = ""

    End Sub

    '顯示修改區
    Sub sUtl_ShowPanelEdit1(ByVal sSearchW As String)
        panelSearch.Visible = False '搜尋功能關閉
        PanelEdit1.Visible = True '修改功能啟動

        '修改
        'sSearchW = e.CommandArgument
        Me.hidKeysID.Value = TIMS.GetMyValue(sSearchW, "KeysID")

        Dim dr As DataRow = search2(Me.hidKeysID.Value) '搜尋顯示
        If Not dr Is Nothing Then
            '06	四大新興智慧型產業	2013
            '07	六大新興產業	2013
            '08	十大重點服務業	2013
            If Convert.ToString(dr("KID")) <> "" Then
                Select Case dr("DepID").ToString
                    Case cst_DepID06
                        Common.SetListItem(ddlKID06, Convert.ToString(dr("KID")))
                    Case cst_DepID10
                        Common.SetListItem(ddlKID10, Convert.ToString(dr("KID")))
                    Case cst_DepID04
                        Common.SetListItem(ddlKID04, Convert.ToString(dr("KID")))
                End Select
            End If
            tKeyNAME.Text = dr("KeyNAME").ToString
        End If
    End Sub

    '設定新增初值
    Sub sUtl_InsertPanelEdit1()
        cb_delKID.Visible = True
        'Me.hidKeysID.Value = ""
        '搜尋->新增
        If Trim(Me.tKeyNAME_s.Text) <> "" Then
            Me.tKeyNAME.Text = Me.tKeyNAME_s.Text
        End If

    End Sub

    '設定 ViewState Value
    Sub sUtl_ViewStateValue()
        Me.ViewState("KeyNAME") = ""
        Me.ViewState("KNAME") = ""
        Me.ViewState("rblDep") = ""

        If Trim(Me.tKeyNAME_s.Text) <> "" Then
            Me.ViewState("KeyNAME") = Trim(Me.tKeyNAME_s.Text.Replace("'", "''"))
        End If
        'tKNAME_s
        If Trim(Me.tKNAME_s.Text) <> "" Then
            Me.ViewState("KNAME") = Trim(Me.tKNAME_s.Text.Replace("'", "''"))
        End If
        Select Case rblDep_s.SelectedValue
            Case cst_DepID06, cst_DepID10, cst_DepID04
                Me.ViewState("rblDep") = rblDep_s.SelectedValue
        End Select
    End Sub

    'SQL
    Sub search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        panelSearch.Visible = True '搜尋功能啟動
        PanelEdit1.Visible = False '修改功能關閉

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql &= " select  k.KeysID" & vbCrLf
        Sql &= " ,k.KeyNAME" & vbCrLf
        Sql &= " ,k.DepID" & vbCrLf
        Sql &= " ,d.DNAME" & vbCrLf
        Sql &= " ,k.KID" & vbCrLf
        Sql &= " ,m.KNAME" & vbCrLf
        Sql &= " from Key_BusinessKeys k " & vbCrLf
        Sql &= " LEFT JOIN Key_Depot d  on d.DepID =k.DepID " & vbCrLf
        Sql &= " LEFT JOIN Key_Business m  on m.KID=k.KID and m.DepID=k.DepID" & vbCrLf
        Sql &= " WHERE 1=1" & vbCrLf
        If Me.ViewState("KeyNAME") <> "" Then
            Sql &= " and k.KeyNAME like '%" & Me.ViewState("KeyNAME") & "%'" & vbCrLf
        End If
        If Me.ViewState("KNAME") <> "" Then
            Sql &= " and m.KNAME like '%" & Me.ViewState("KNAME") & "%'" & vbCrLf
        End If
        If Me.ViewState("rblDep") <> "" Then
            Sql &= " and k.DepID= '" & Me.ViewState("rblDep") & "'" & vbCrLf
        End If
        Sql &= " ORDER BY k.DepID,k.KID,k.KeyNAME" & vbCrLf

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
    Function search2(ByVal KeysID As String) As DataRow
        Dim rst As DataRow = Nothing

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql &= " SELECT " & vbCrLf
        Sql &= " a.KeysID  /*PK*/ " & vbCrLf
        Sql &= " ,a.DepID" & vbCrLf
        Sql &= " ,a.KID" & vbCrLf
        Sql &= " ,a.KeyNAME" & vbCrLf
        Sql &= " FROM Key_BusinessKeys a" & vbCrLf
        Sql &= " WHERE 1=1" & vbCrLf
        Sql &= " and a.KeysID  = '" & KeysID & "'" & vbCrLf

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
        If hidKeysID.Value = "" Then
            sType = cst_addnew '新增
        End If
        Dim sql As String = ""

        '確認
        'sSearchW = e.CommandArgument
        Dim sDepID As String = TIMS.GetMyValue(sSearchW, "DepID")
        Dim sKID As String = TIMS.GetMyValue(sSearchW, "KID")
        Dim sDelOld As String = TIMS.GetMyValue(sSearchW, "DelOld")

        tKeyNAME.Text = TIMS.ClearSQM(tKeyNAME.Text)
        tKeyNAME.Text = Trim(tKeyNAME.Text.Replace("、", ",")) '大寫頓號轉換
        tKeyNAME.Text = Trim(tKeyNAME.Text.Replace(";", ",")) '分號轉換
        tKeyNAME.Text = Trim(tKeyNAME.Text.Replace("；", ",")) '大寫分號轉換
        Dim multidata_flag As Boolean = False '單筆資料
        Dim tKeyNAMEs As String()
        If tKeyNAME.Text.ToString.IndexOf(",") > -1 Then
            '多筆資料
            sType = cst_multiaddnew '多筆新增
            multidata_flag = True
            tKeyNAMEs = Split(tKeyNAME.Text.ToString, ",")
        Else
            tKeyNAMEs = Nothing
        End If

        Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)

        '刪除用 (可刪除多筆)
        Dim daDel As SqlDataAdapter = TIMS.GetOneDA(objconn)
        If sDelOld = cst_DelOld_N Then
            sql = "delete Key_BusinessKeys where DepID= @DepID AND KID= @KID AND KeyNAME= @KeyNAME "
        Else
            sql = "delete Key_BusinessKeys where DepID= @DepID AND KID= @KID"
        End If
        daDel.SelectCommand.CommandText = sql

        Select Case sType
            Case cst_multiaddnew '多筆新增
                sql = "" & vbCrLf
                'Sql += "  /* IDENTITY(1,1): KeysID */ " & vbCrLf
                sql &= " INSERT INTO Key_BusinessKeys(" & vbCrLf
                sql &= " DepID" & vbCrLf
                sql &= "  ,KID" & vbCrLf
                sql &= "  ,KeyNAME" & vbCrLf
                sql &= " ) VALUES (" & vbCrLf
                sql &= " @DepID" & vbCrLf
                sql &= "  ,@KID" & vbCrLf
                sql &= "  ,@KeyNAME" & vbCrLf
                sql &= " ) " & vbCrLf
                'sql += " select @@identity as KeysID" & vbCrLf

                If sDelOld = cst_DelOld_Y Then
                    '刪除用
                    daDel.SelectCommand.Parameters.Clear()
                    daDel.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                    daDel.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                    If daDel.SelectCommand.Connection.State = ConnectionState.Closed Then daDel.SelectCommand.Connection.Open()
                    Call daDel.SelectCommand.ExecuteNonQuery() '.ExecuteNonQuery()
                End If

                If Not tKeyNAMEs Is Nothing Then
                    For i As Integer = 0 To tKeyNAMEs.Length - 1
                        If Trim(tKeyNAMEs(i)) <> "" Then
                            If sDelOld = cst_DelOld_N Then
                                '刪除用
                                daDel.SelectCommand.Parameters.Clear()
                                daDel.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                                daDel.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                                daDel.SelectCommand.Parameters.Add("KeyNAME", SqlDbType.VarChar).Value = Trim(tKeyNAMEs(i))
                                If daDel.SelectCommand.Connection.State = ConnectionState.Closed Then daDel.SelectCommand.Connection.Open()
                                Call daDel.SelectCommand.ExecuteNonQuery() '.ExecuteNonQuery()
                            End If

                            da.SelectCommand.CommandText = sql
                            da.SelectCommand.Parameters.Clear()
                            da.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                            da.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                            da.SelectCommand.Parameters.Add("KeyNAME", SqlDbType.VarChar).Value = Trim(tKeyNAMEs(i))

                            If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                            da.SelectCommand.ExecuteNonQuery()

                            hidKeysID.Value = DbAccess.GetNewId(da.SelectCommand.Connection, "KEY_BUSINESSKEYS_KEYSID_SEQ") 'da.SelectCommand.ExecuteScalar() '.ExecuteNonQuery()
                        End If
                    Next
                End If

                rst = 1 '儲存成功

            Case cst_addnew '新增
                sql = "" & vbCrLf
                'Sql += "  /* IDENTITY(1,1): KeysID */ " & vbCrLf
                sql &= " INSERT INTO Key_BusinessKeys(" & vbCrLf
                sql &= " DepID" & vbCrLf
                sql &= "  ,KID" & vbCrLf
                sql &= "  ,KeyNAME" & vbCrLf
                sql &= " ) VALUES (" & vbCrLf
                sql &= " @DepID" & vbCrLf
                sql &= "  ,@KID" & vbCrLf
                sql &= "  ,@KeyNAME" & vbCrLf
                sql &= " ) " & vbCrLf
                'sql += " select @@identity as KeysID" & vbCrLf

                If Trim(tKeyNAME.Text) <> "" Then
                    If sDelOld = cst_DelOld_N Then
                        '刪除用
                        daDel.SelectCommand.Parameters.Clear()
                        daDel.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                        daDel.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                        daDel.SelectCommand.Parameters.Add("KeyNAME", SqlDbType.VarChar).Value = Trim(tKeyNAME.Text)
                    Else
                        '刪除用
                        daDel.SelectCommand.Parameters.Clear()
                        daDel.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                        daDel.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                    End If
                    If daDel.SelectCommand.Connection.State = ConnectionState.Closed Then daDel.SelectCommand.Connection.Open()
                    Call daDel.SelectCommand.ExecuteNonQuery() '.ExecuteNonQuery()

                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    da.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                    da.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                    da.SelectCommand.Parameters.Add("KeyNAME", SqlDbType.VarChar).Value = Trim(tKeyNAME.Text)

                    If da.SelectCommand.Connection.State = ConnectionState.Closed Then da.SelectCommand.Connection.Open()
                    da.SelectCommand.ExecuteNonQuery()
                    'hidKeysID.Value = da.SelectCommand.ExecuteScalar() '.ExecuteNonQuery()
                    hidKeysID.Value = DbAccess.GetNewId(da.SelectCommand.Connection, "KEY_BUSINESSKEYS_KEYSID_SEQ")

                End If

                rst = 1 '儲存成功

            Case Else 'cst_update '修改
                sql = "" & vbCrLf
                sql &= " UPDATE Key_BusinessKeys SET" & vbCrLf
                sql &= " DepID= @DepID" & vbCrLf
                sql &= " ,KID= @KID" & vbCrLf
                sql &= " ,KeyNAME= @KeyNAME" & vbCrLf
                sql &= " WHERE 1=1 " & vbCrLf
                sql &= " AND KeysID= @KeysID" & vbCrLf

                If Trim(tKeyNAME.Text) <> "" Then
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    da.SelectCommand.Parameters.Add("DepID", SqlDbType.VarChar).Value = sDepID
                    da.SelectCommand.Parameters.Add("KID", SqlDbType.VarChar).Value = sKID
                    da.SelectCommand.Parameters.Add("KeyNAME", SqlDbType.VarChar).Value = Trim(tKeyNAME.Text)
                    da.SelectCommand.Parameters.Add("KeysID", SqlDbType.VarChar).Value = hidKeysID.Value 'WHERE

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
            Case cst_ItemEdit1
                '修改
                sSearchW = e.CommandArgument
                Me.hidKeysID.Value = TIMS.GetMyValue(sSearchW, "KeysID")

                Call sUtl_ClearPanelEdit1()

                Call sUtl_ShowPanelEdit1(sSearchW)

            Case cst_ItemDelete1
                '刪除
                sSearchW = e.CommandArgument
                Me.hidKeysID.Value = TIMS.GetMyValue(sSearchW, "KeysID")
                Try
                    '刪除sql
                    Dim sql As String = ""
                    sql = ""
                    sql &= " DELETE Key_BusinessKeys " & vbCrLf
                    sql &= " WHERE 1=1 " & vbCrLf
                    sql &= " AND KeysID= @KeysID" & vbCrLf

                    Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
                    da.SelectCommand.CommandText = sql
                    da.SelectCommand.Parameters.Clear()
                    da.SelectCommand.Parameters.Add("KeysID", SqlDbType.VarChar).Value = hidKeysID.Value

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

                Dim lbDNAME As Label = e.Item.FindControl("lbDNAME")
                Dim lbKNAME As Label = e.Item.FindControl("lbKNAME")
                Dim lbKeyNAME As Label = e.Item.FindControl("lbKeyNAME")
                Dim btnEdit1 As LinkButton = e.Item.FindControl("btnEdit1")
                Dim btnDelete1 As LinkButton = e.Item.FindControl("btnDelete1")

                '序號
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + sender.PageSize * sender.CurrentPageIndex

                lbDNAME.Text = Convert.ToString(drv("DNAME"))
                lbKNAME.Text = Convert.ToString(drv("KNAME"))
                lbKeyNAME.Text = Convert.ToString(drv("KeyNAME"))

                Dim cmdArg As String = ""
                cmdArg = ""
                cmdArg &= "&KeysID=" & Convert.ToString(drv("KeysID"))

                'btnView1.CommandArgument = cmdArg & "&btn=" & cst_view
                btnEdit1.CommandArgument = cmdArg & "&btn=" & cst_edit
                btnDelete1.CommandArgument = cmdArg
                btnDelete1.Attributes("onclick") = "return confirm('此動作會刪除此筆資料，是否確定刪除?');"

        End Select
    End Sub

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
        Dim DepID As String = ""
        Dim KID As String = ""
        '-------------------------------儲存前檢核-------------------------------
        Dim sErrmsg As String = ""
        sErrmsg = ""
        Dim i As Integer = 0
        i = 0
        If ddlKID06.SelectedValue <> "" Then
            DepID = cst_DepID06
            KID = ddlKID06.SelectedValue
            i += 1
        End If
        If ddlKID10.SelectedValue <> "" Then
            DepID = cst_DepID10
            KID = ddlKID10.SelectedValue
            i += 1
        End If
        If ddlKID04.SelectedValue <> "" Then
            DepID = cst_DepID04
            KID = ddlKID04.SelectedValue
            i += 1
        End If
        If Me.tKeyNAME.Text = "" Then
            sErrmsg += "請輸入有效關鍵字" & vbCrLf
        End If

        If i <> 1 Then
            sErrmsg += "只能選擇 1項產業" & vbCrLf
        End If
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If
        ''-------------------------------儲存前檢核-------------------------------

        '刪除舊資
        Dim sDelOld As String = cst_DelOld_N '"N"
        If Me.cb_delKID.Checked Then sDelOld = cst_DelOld_Y '"Y"

        Dim cmdArg As String = ""
        cmdArg = ""
        cmdArg &= "&DepID=" & DepID
        cmdArg &= "&KID=" & KID
        cmdArg &= "&DelOld=" & sDelOld

        If sUtl_SaveData1(cmdArg) = 1 Then
            '正常儲存
            Call search1()
        End If

    End Sub

    ''Return the assembly's Title attribute value.
    'Private Function GetAssemblyTitle() As String
    '    m_MyAssembly = [Assembly].GetExecutingAssembly()

    '    If m_MyAssembly.IsDefined(GetType(AssemblyTitleAttribute), _
    '        True) Then
    '        Dim attr As Attribute = _
    '            Attribute.GetCustomAttribute(m_MyAssembly, _
    '                GetType(AssemblyTitleAttribute))
    '        Dim title_attr As AssemblyTitleAttribute = _
    '            DirectCast(attr, AssemblyTitleAttribute)
    '        Return title_attr.Title
    '    Else
    '        Return ""
    '    End If
    'End Function

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出Excel
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了
        DataGrid1.Columns(4).Visible = False

        Call sUtl_ViewStateValue()  '設定 ViewState Value

        Call search1()

        Dim txtTitle As String = "產業關鍵字設定"
        'Dim txtTitle As String = GetAssemblyTitle()

        Dim sFileName As String = txtTitle & ".xls" '"離退訓人數統計表.xls"
        sFileName = HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8)
        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        'DataGrid1.Visible = False

    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
