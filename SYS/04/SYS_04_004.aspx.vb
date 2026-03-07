Partial Class SYS_04_004
    Inherits AuthBasePage

    Dim dtKeyBudget As DataTable = Nothing
    Dim dtPlanBudget As DataTable = Nothing
    'Dim FunDr As DataRow = Nothing
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn) '開啟連線

        If Not IsPostBack Then
            Call cCreate1()
        End If

        Syear.Attributes("onchange") = "DropChange();"
        TPlanID.Attributes("onchange") = "DropChange();"
        Button1.Attributes("onclick") = "return check_search();"

    End Sub

    Sub cCreate1()
        'ddlDISTID = TIMS.Get_DistID(ddlDISTID, Nothing, objconn)
        'Hid_DistID.Value = Convert.ToString(sm.UserInfo.DistID)
        'Common.SetListItem(ddlDISTID, sm.UserInfo.DistID)
        'ddlDISTID.Enabled = False
        'Select Case sm.UserInfo.LID
        '    Case 0
        '        ddlDISTID.Enabled = True
        'End Select

        Syear = TIMS.GetSyear(Syear)
        TPlanID = TIMS.Get_TPlan(TPlanID)
        Common.SetListItem(Syear, Now.Year)
        DataGridTable.Style.Item("display") = "none"

    End Sub

    Sub sSearch1()
        Dim sql As String = ""
        sql = "SELECT BUDID,BUDNAME FROM KEY_BUDGET ORDER BY BUDID"
        dtKeyBudget = DbAccess.GetDataTable(sql, objconn)

        'TPLANID
        Dim v_SYEAR As String = TIMS.GetListValue(Syear) '.SelectedValue
        sql = "SELECT * FROM PLAN_BUDGET WHERE SYEAR='" & v_SYEAR & "'"
        dtPlanBudget = DbAccess.GetDataTable(sql, objconn)

        Dim v_TPlanID As String = TIMS.GetListValue(TPlanID) '.SelectedValue
        sql = " SELECT TPLANID,PLANNAME FROM KEY_PLAN WHERE 1=1" & vbCrLf
        If v_TPlanID <> "" Then sql += " and TPlanID='" & v_TPlanID & "'"
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Style.Item("display") = "none"
        If dt.Rows.Count > 0 Then
            DataGridTable.Style.Item("display") = ""

            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "TPlanID"
            DataGrid1.DataBind()
        End If
    End Sub

    ''' <summary>
    ''' 查詢
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call sSearch1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Checkbox1.Attributes("onclick") = "SelectAll(this.checked)"

            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Hid_TPlanID As HiddenField = e.Item.FindControl("Hid_TPlanID")
                Dim BudID As CheckBoxList = e.Item.FindControl("BudID")
                Hid_TPlanID.Value = Convert.ToString(drv("TPlanID"))
                With BudID
                    .DataSource = dtKeyBudget
                    .DataTextField = "BudName"
                    .DataValueField = "BudID"
                    .DataBind()
                End With

                Dim ff As String = "TPLANID='" & drv("TPLANID") & "'"
                If dtPlanBudget.Select(ff).Length > 0 Then
                    For Each dr As DataRow In dtPlanBudget.Select(ff)
                        For Each item As ListItem In BudID.Items
                            If item.Value = dr("BudID").ToString Then
                                item.Selected = True
                                Exit For
                            End If
                        Next
                    Next
                End If
        End Select

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim s_sql As String = ""
        s_sql &= " SELECT 'x' FROM PLAN_BUDGET" & vbCrLf
        s_sql &= " WHERE SYEAR=@SYEAR AND TPLANID=@TPLANID AND BUDID=@BUDID" & vbCrLf

        Dim v_SYEAR As String = TIMS.GetListValue(Syear) '.SelectedValue

        Dim i_UPROW As Integer = 0
        For Each eItem As DataGridItem In DataGrid1.Items
            'Dim Hid_SYEAR As HiddenField = eItem.FindControl("Hid_SYEAR")
            Dim Checkbox1c As HtmlInputCheckBox = eItem.FindControl("Checkbox1c")
            Dim Hid_TPlanID As HiddenField = eItem.FindControl("Hid_TPlanID")
            Dim BudID As CheckBoxList = eItem.FindControl("BudID")
            Hid_TPlanID.Value = TIMS.ClearSQM(Hid_TPlanID.Value)
            If Hid_TPlanID.Value = "" Then Exit For

            For Each bud_item1 As ListItem In BudID.Items
                bud_item1.Value = TIMS.ClearSQM(bud_item1.Value)
                If bud_item1.Value <> "" AndAlso Checkbox1c.Checked Then

                    If bud_item1.Selected Then
                        Dim dt As New DataTable
                        Using CMD_s1 As New SqlCommand(s_sql, objconn)
                            With CMD_s1
                                .Parameters.Add("SYEAR", SqlDbType.VarChar).Value = v_SYEAR
                                .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = Hid_TPlanID.Value
                                .Parameters.Add("BUDID", SqlDbType.VarChar).Value = bud_item1.Value
                                dt.Load(.ExecuteReader())
                            End With
                        End Using
                        If TIMS.dtNODATA(dt) Then
                            i_UPROW += 1
                            'INSERT
                            Dim i_sql As String = ""
                            i_sql &= " INSERT INTO PLAN_BUDGET(PBID,SYEAR,TPLANID,BUDID,MODIFYACCT,MODIFYDATE)" & vbCrLf
                            i_sql &= " VALUES(@PBID,@SYEAR,@TPLANID,@BUDID,@MODIFYACCT,GETDATE())" & vbCrLf
                            Dim iPBID As Integer = DbAccess.GetNewId(objconn, "PLAN_BUDGET_PBID_SEQ,PLAN_BUDGET,PBID")
                            Dim parms As New Hashtable From {{"PBID", iPBID}, {"SYEAR", v_SYEAR}, {"TPLANID", Hid_TPlanID.Value}, {"BUDID", bud_item1.Value}, {"MODIFYACCT", sm.UserInfo.UserID}}
                            DbAccess.ExecuteNonQuery(i_sql, objconn, parms)
                        End If
                    Else
                        Dim dt As New DataTable
                        Using CMD_s1 As New SqlCommand(s_sql, objconn)
                            With CMD_s1
                                .Parameters.Add("SYEAR", SqlDbType.VarChar).Value = v_SYEAR
                                .Parameters.Add("TPLANID", SqlDbType.VarChar).Value = Hid_TPlanID.Value
                                .Parameters.Add("BUDID", SqlDbType.VarChar).Value = bud_item1.Value
                                dt.Load(.ExecuteReader())
                            End With
                        End Using
                        If TIMS.dtHaveDATA(dt) Then
                            i_UPROW += 1
                            'DELETE
                            Dim d_sql As String = ""
                            d_sql &= " DELETE PLAN_BUDGET" & vbCrLf
                            d_sql &= " WHERE SYEAR=@SYEAR AND TPLANID=@TPLANID AND BUDID=@BUDID" & vbCrLf
                            Dim parms As New Hashtable From {{"SYEAR", v_SYEAR}, {"TPLANID", Hid_TPlanID.Value}, {"BUDID", bud_item1.Value}}
                            DbAccess.ExecuteNonQuery(d_sql, objconn, parms)
                        End If
                    End If

                End If

            Next
        Next

        'DbAccess.UpdateDataTable(dt, da)
        If i_UPROW > 0 Then
            Common.MessageBox(Me, String.Concat("儲存成功!(", i_UPROW, ")"))
        Else
            Common.MessageBox(Me, "無資料變動!")
        End If
    End Sub
End Class
