Public Class SYS_03_036
    Inherits AuthBasePage

    'Const cst_title1 As String = "功能使用盤點統計資料"
    'Const cst_title2 As String = "功能使用盤點明細資料"
    Const cst_sys03036Scope As String = "sys03036Scope"
    Dim ss_sys03036Scope As String = "" '儲存查詢值。

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call sUtl_Show(1)
            Call Create1()
        End If

    End Sub

    Sub Create1()

    End Sub

    '顯示
    Sub sUtl_Show(ByVal iType As Integer)
        SchTable.Visible = False
        DataTable1.Visible = False
        DataTable2.Visible = False
        Select Case iType
            Case 1
                SchTable.Visible = True
                DataTable1.Visible = True
            Case 2
                DataTable2.Visible = True
        End Select
    End Sub

    '查詢
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        Call sUtl_Show(1)
        Call Show_DataGrid1(0)
    End Sub

    '查詢1
    Sub Show_DataGrid1(ByVal tmpPage As Integer)
        '儲存查詢值。
        Me.ViewState(cst_sys03036Scope) = ""
        ss_sys03036Scope = ""
        TIMS.SetMyValue(ss_sys03036Scope, "FunName", txtFunName.Text)
        Me.ViewState(cst_sys03036Scope) = ss_sys03036Scope

        '整理輸入值值
        txtFunName.Text = TIMS.ClearSQM(txtFunName.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT f.funid" & vbCrLf
        sql &= " ,f.name funname" & vbCrLf
        sql &= " ,replace(replace(f.funpath2,'//','/'),'/','>>') funpath" & vbCrLf
        sql &= " ,f.spage" & vbCrLf
        sql &= " ,f.kind" & vbCrLf
        sql &= " ,f.levels" & vbCrLf
        sql &= " ,f.memo" & vbCrLf
        sql &= " FROM VIEW_FUNCTION f" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If txtFunName.Text <> "" Then
            sql += " AND f.name like '%'+@FunName+'%'" & vbCrLf
        Else
            '未輸入有效值，查無資料
            sql += " and 1<>1" & vbCrLf
        End If
        sql &= " ORDER BY f.kind,f.levels,f.funid" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            If txtFunName.Text <> "" Then
                .Parameters.Add("FunName", SqlDbType.VarChar).Value = txtFunName.Text
            End If
            dt.Load(.ExecuteReader())
        End With

        lab_Msg1.Visible = True
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            lab_Msg1.Visible = False
            DataGrid1.Visible = True

            DataGrid1.DataSource = dt
            DataGrid1.CurrentPageIndex = tmpPage
            DataGrid1.DataBind()
        End If

    End Sub

    '查詢2
    Sub Show_DataGrid2(ByVal ifunid As Integer, ByVal ssScope As String, ByVal tmpPage As Integer)
        'ifunid:若為0 則是顯示全部功能範圍。
        'If txtFunName.Text <> "" Then txtFunName.Text = Trim(txtFunName.Text)
        'Const cst_funPath As Integer = 1
        'Const cst_funName As Integer = 2

        'Dim srblScope As String = TIMS.GetMyValue(ssScope, "rblScope")

        'Dim SYM1 As String = TIMS.GetMyValue(sys03035Scope, "SYM1")
        'Dim SYM2 As String = TIMS.GetMyValue(sys03035Scope, "SYM2")
        'Dim sddlY1 As String = TIMS.GetMyValue(sys03035Scope, "ddlY1")
        'Dim sddlY2 As String = TIMS.GetMyValue(sys03035Scope, "ddlY2")
        'Dim sMDATE1 As String = TIMS.GetMyValue(sys03035Scope, "MDATE1")
        'Dim sMDATE2 As String = TIMS.GetMyValue(sys03035Scope, "MDATE2")

        Dim sFunName As String = TIMS.GetMyValue(ssScope, "FunName")

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " with wgf as (" & vbCrLf
        sql &= " select gid from auth_groupfun where funid =@funid" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT a.GID" & vbCrLf
        sql &= " ,a.GDISTID" & vbCrLf
        sql &= " ,a.GTYPE" & vbCrLf
        sql &= " ,a.GNAME" & vbCrLf
        sql &= " ,a.GNOTE" & vbCrLf
        sql &= " ,a.GVALID" & vbCrLf
        sql &= " ,a.GSTATE" & vbCrLf
        sql &= " ,a.CREATEACCT" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,cr.name cname" & vbCrLf
        sql &= " ,mo.name mname" & vbCrLf
        sql &= " from AUTH_GROUP a" & vbCrLf
        sql &= " join wgf ON wgf.GID=a.GID " & vbCrLf
        sql &= " left join Auth_Account cr on cr.Account=a.CreateAcct" & vbCrLf
        sql &= " left join Auth_Account mo on mo.Account=a.ModifyAcct" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and a.GState<>'D'" & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql &= " AND a.GDISTID=@GDISTID" & vbCrLf
        End If
        sql &= " ORDER BY a.GDISTID,a.GTYPE" & vbCrLf
        Dim oCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("funid", SqlDbType.Int).Value = ifunid
            If sm.UserInfo.DistID <> "000" Then
                .Parameters.Add("GDISTID", SqlDbType.VarChar).Value = sm.UserInfo.DistID
                'sql &= " AND a.GDistID=@GDistID" & vbCrLf
            End If
            dt.Load(.ExecuteReader())
        End With

        'btnExport1.Visible = False
        'btnExport2.Visible = False

        lab_Msg2.Visible = True
        DataGrid2.Visible = False
        If dt.Rows.Count > 0 Then
            'btnExport1.Visible = True
            'btnExport2.Visible = True
            lab_Msg2.Visible = False
            DataGrid2.Visible = True

            'DataGrid2.Columns(cst_funPath).Visible = False
            'DataGrid2.Columns(cst_funName).Visible = False
            'If ifunid = 0 Then
            '    DataGrid2.Columns(cst_funPath).Visible = True
            '    DataGrid2.Columns(cst_funName).Visible = True
            'End If

            DataGrid2.DataSource = dt
            DataGrid2.CurrentPageIndex = tmpPage
            DataGrid2.DataBind()
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Select Case e.CommandName
            Case "ListData1"
                Dim cmdArg As String = e.CommandArgument
                Hidfunid.Value = TIMS.GetMyValue(cmdArg, "funid")
                lFunPath.Text = TIMS.GetMyValue(cmdArg, "funpath")
                lFunName.Text = TIMS.GetMyValue(cmdArg, "funName")

                Call sUtl_Show(2)
                Show_DataGrid2(Val(Hidfunid.Value), Me.ViewState(cst_sys03036Scope), 0)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim dr_Data As DataRowView = e.Item.DataItem

                'Dim labFunID As Label = e.Item.FindControl("labFunID")
                Dim labFunPath As Label = e.Item.FindControl("labFunPath")
                Dim labFunName As Label = e.Item.FindControl("labFunName")
                'Dim labCount As Label = e.Item.FindControl("labCount")
                Dim btnListData1 As LinkButton = e.Item.FindControl("btnListData1")

                e.Item.Cells(0).Text = sender.PageSize * sender.CurrentPageIndex + e.Item.ItemIndex + 1

                Dim cmdArg As String = ""
                Call TIMS.SetMyValue(cmdArg, "funid", Convert.ToString(dr_Data("FunID")))
                Call TIMS.SetMyValue(cmdArg, "funpath", Convert.ToString(dr_Data("FunPath")))
                Call TIMS.SetMyValue(cmdArg, "funName", Convert.ToString(dr_Data("FunName")))
                'Call TIMS.SetMyValue(cmdArg, "Scope", Me.rblScope.SelectedValue)
                btnListData1.CommandArgument = cmdArg

                'labFunID.Text = Convert.ToString(dr_Data("FunID"))
                'labFunID.Text = Convert.ToString(sender.PageSize * sender.CurrentPageIndex + e.Item.ItemIndex + 1)
                labFunPath.Text = Convert.ToString(dr_Data("FunPath"))
                labFunName.Text = Convert.ToString(dr_Data("funname"))
                'labCount.Text = Convert.ToString(dr_Data("count1"))

        End Select
    End Sub

    Private Sub DataGrid1_PageIndexChanged(source As Object, e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid1.PageIndexChanged
        Show_DataGrid1(e.NewPageIndex)
    End Sub

    Private Sub DataGrid2_ItemDataBound(sender As Object, e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
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

                labSNo.Text = Convert.ToString(e.Item.ItemIndex + 1)

                lab_GroupDistID.Text = "(系統預設)"
                If Convert.ToString(dr_Data("GDistID")) <> "" Then
                    lab_GroupDistID.Text = TIMS.Get_DistName1(dr_Data("GDistID"))
                End If

                Select Case Convert.ToString(dr_Data("GType"))
                    Case "0"
                        'lab_GroupType.Text = "局"
                        lab_GroupType.Text = "署"
                    Case "1"
                        'lab_GroupType.Text = "中心"
                        lab_GroupType.Text = "分署"
                    Case "2"
                        lab_GroupType.Text = "委訓"
                End Select

                labGroupName.Text = Convert.ToString(dr_Data("GName"))
                lab_GroupCUsr.Text = Convert.ToString(dr_Data("cname"))
                lab_GroupMUsr.Text = Convert.ToString(dr_Data("mname"))
                labGroupNote.Text = Convert.ToString(dr_Data("GNote"))
                labEnable.Text = IIf(Convert.ToString(dr_Data("GValid")) = "1", "是", "否")

        End Select
    End Sub

    Private Sub DataGrid2_PageIndexChanged(source As Object, e As System.Web.UI.WebControls.DataGridPageChangedEventArgs) Handles DataGrid2.PageIndexChanged
        Show_DataGrid2(Val(Hidfunid.Value), Me.ViewState(cst_sys03036Scope), e.NewPageIndex)
    End Sub

    Protected Sub btnBack2_Click(sender As Object, e As EventArgs) Handles btnBack2.Click
        Call sUtl_Show(1)
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

End Class