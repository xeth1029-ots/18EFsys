Partial Class OB_01_003
    Inherits AuthBasePage

    Const Cst_MTDate As Integer = 5

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            ddl_years = TIMS.GetSyear(ddl_years, Year(Now) - 1, Year(Now) + 3, True)
            '因有傳入值 yearlist.SelectedValue.ToString 故放此位置，才可讀到值

            DataGridTable.Visible = False
        End If
    End Sub

    Sub Query()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select a.* , ot.TenderCName, ot.years" & vbCrLf
        sql += " from OB_Meeting a " & vbCrLf
        sql += " join OB_tender ot on ot.tsn=a.tsn " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        If sm.UserInfo.DistID <> "000" Then
            sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If

        If Me.ddl_years.SelectedValue <> "" Then
            sql += " AND ot.years='" & Me.ddl_years.SelectedValue & "'" & vbCrLf
        End If

        If Me.ddlTenderCName.SelectedValue <> "" Then
            sql += " AND a.tsn=" & Me.ddlTenderCName.SelectedValue & vbCrLf
        End If

        MTSubject.Text = TIMS.ClearSQM(MTSubject.Text)
        If MTSubject.Text.Trim <> "" Then
            MTSubject.Text = MTSubject.Text.Trim
            sql += " AND a.MTSubject like '%" & MTSubject.Text & "%'" & vbCrLf
        End If

        MTPlace.Text = TIMS.ClearSQM(MTPlace.Text)
        If MTPlace.Text.Trim <> "" Then
            MTPlace.Text = MTPlace.Text.Trim
            sql += " AND a.MTPlace like '%" & MTPlace.Text & "%'" & vbCrLf
        End If

        MTDate.Text = TIMS.ClearSQM(MTDate.Text)
        If MTDate.Text.Trim <> "" Then
            MTDate.Text = FormatDateTime(CDate(MTDate.Text.Trim), DateFormat.ShortDate)
            sql += " AND a.MTDate=convert(datetime, '" & MTDate.Text & "', 111) " & vbCrLf
        End If

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        Else
            DataGridTable.Visible = False
            msg.Text = "查無資料!!"
        End If

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        If Page.IsValid Then
            Query()
        End If
    End Sub

    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click
        TIMS.CloseDbConn(objconn)
        TIMS.Utl_Redirect1(Me, "OB_01_003_add.aspx?ID=" & Request("ID") & "&Action=ADD")
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim btn_edit As Button = e.Item.FindControl("btn_edit")
                Dim btn_del As Button = e.Item.FindControl("btn_del")
                btn_edit.CommandArgument = drv("mtsn")
                btn_del.CommandArgument = drv("mtsn")
                btn_del.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                e.Item.Cells(Cst_MTDate).Text = FormatDateTime(drv("MTDate"), DateFormat.ShortDate)

        End Select
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        Dim sql As String = ""
        Select Case e.CommandName
            Case "del"
                sql = "delete OB_Meeting where mtsn='" & e.CommandArgument & "'"
                DbAccess.ExecuteNonQuery(sql, objconn)
                Common.MessageBox(Me, "刪除成功！")
                Call Query()

            Case "edit"
                TIMS.Utl_Redirect1(Me, "OB_01_003_add.aspx?ID=" & Request("ID") & "&Action=EDIT&mtsn=" & e.CommandArgument)

        End Select
    End Sub

    Private Sub Page_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles MyBase.PreRender
        Dim strMessage As String = ""
        For Each obj As WebControls.BaseValidator In Page.Validators
            If obj.IsValid = False Then
                strMessage &= obj.ErrorMessage & vbCrLf
            End If
        Next

        If strMessage <> "" Then
            Common.MessageBox(Page, strMessage)
        End If
    End Sub

    Private Sub CustomValidator1_ServerValidate(ByVal source As System.Object, ByVal args As System.Web.UI.WebControls.ServerValidateEventArgs) Handles CustomValidator1.ServerValidate
        source.ErrorMessage = ""

        MTDate.Text = TIMS.ClearSQM(MTDate.Text)
        If MTDate.Text.Trim <> "" Then
            MTDate.Text = MTDate.Text.Trim
            If Not IsDate(MTDate.Text) Then
                source.ErrorMessage &= "會議日期格式有誤" & vbCrLf
            Else
                MTDate.Text = FormatDateTime(CDate(MTDate.Text.Trim), DateFormat.ShortDate)
            End If
        End If

        If source.ErrorMessage = "" Then args.IsValid = True Else args.IsValid = False

    End Sub

    Private Sub ddl_years_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddl_years.SelectedIndexChanged
        TIMS.Get_TenderCName(ddlTenderCName, sender.SelectedValue, objconn)
    End Sub
End Class
