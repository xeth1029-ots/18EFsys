Partial Class OB_01_001_mem
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)

        If Not IsPostBack Then

            '因有傳入值 yearlist.SelectedValue.ToString 故放此位置，才可讀到值
            DataGridTable.Visible = False
            'btn_back.Attributes.Add("onclick", "window.close();return false;")

            ViewState("TSN") = TIMS.ClearSQM(Request("TSN"))
            ViewState("Action") = TIMS.ClearSQM(Request("Action"))
            If ViewState("TSN") <> "" And ViewState("Action") = "MEM" Then
                Query()
            End If

        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        Page.RegisterStartupScript("ClearOrg1", "<script>ClearOrg();</script>")
    End Sub


    Sub Query(Optional ByVal sql As String = "")

        Dim dt As DataTable
        If sql <> "" Then
        Else
            sql = "" & vbCrLf
            sql += " select b.OrgName, a.*, c.Msn as TMsn, c.leader" & vbCrLf
            sql += " from OB_Member a " & vbCrLf
            sql += " JOIN OB_Org b on a.OrgSN=b.OrgSN " & vbCrLf
            sql += " LEFT JOIN (SELECT * FROM OB_Tmember WHERE TSN = '" & ViewState("TSN") & "' " & vbCrLf
            sql += "    ) c on a.msn=c.msn " & vbCrLf
            sql += " WHERE 1=1 " & vbCrLf
            If sm.UserInfo.DistID <> "000" Then
                sql += " AND a.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
            End If

            Select Case True
                Case Radiobutton1.Checked
                    sql += " AND c.TSN ='" & ViewState("TSN") & "' " & vbCrLf
                Case Radiobutton2.Checked
                    sql += " AND c.TSN IS NULL " & vbCrLf
            End Select


            If center.Text.Trim <> "" Then
                center.Text = center.Text.Trim
                sql += " AND b.OrgName like '%" & center.Text & "%'" & vbCrLf
            End If

            If orgid_value.Value <> "" Then
                sql += " AND a.OrgID='" & orgid_value.Value & "'" & vbCrLf
            End If

            If DeptName.Text.Trim <> "" Then
                DeptName.Text = DeptName.Text.Trim()
                sql += " AND b.DeptName like '%" & DeptName.Text & "%'" & vbCrLf
            End If

            Select Case Me.rblQualified.SelectedValue
                Case "N", "B", "A"
                    sql += " AND a.Qualified ='" & Me.rblQualified.SelectedValue & "'" & vbCrLf
            End Select

        End If

        dt = DbAccess.GetDataTable(sql)
        If dt.Rows.Count > 0 Then
            DataGridTable.Visible = True
            msg.Text = ""
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
            'PageControler1.PageDataTable = dt
            'PageControler1.ControlerLoad()
        Else
            DataGridTable.Visible = False
            msg.Text = "查無資料!!"
        End If

    End Sub

    Private Sub btnQuery_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuery.Click
        Query()
    End Sub


    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidmsn As HtmlInputHidden = e.Item.FindControl("hidmsn")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim Checkbox2 As HtmlInputCheckBox = e.Item.FindControl("Checkbox2")
                hidmsn.Value = drv("msn")

                If drv("TMsn").ToString <> "" Then
                    Checkbox1.Checked = True
                End If
                Select Case drv("leader").ToString
                    Case "Y"
                        Checkbox2.Checked = True
                    Case Else
                        Checkbox2.Checked = False
                End Select
        End Select
    End Sub

    Private Sub btn_back_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_back.Click
        TIMS.Utl_Redirect1(Me, "OB_01_001.aspx?ID=" & Request("ID"))
    End Sub

    Private Sub btn_Save2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btn_Save2.Click
        Dim Errmsg As String = ""

        Check_OB_Tmember(Errmsg)

        If Errmsg <> "" Then
            Common.MessageBox(Me, Errmsg)
            Exit Sub
        Else
            Insert_OB_Tmember()
        End If

    End Sub

    Sub Check_OB_Tmember(ByRef Errmsg As String)
        Errmsg = ""

        Dim intRows As Integer = 0
        Dim intCheckbox2 As Integer = 0

        For Each Item As DataGridItem In DataGrid1.Items
            Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
            Dim Checkbox2 As HtmlInputCheckBox = Item.FindControl("Checkbox2")
            If Checkbox1.Checked Then
                intRows += 1
                If Checkbox2.Checked Then
                    intCheckbox2 += 1
                    If intCheckbox2 > 1 Then
                        Errmsg += "只能有一位是組長" & vbCrLf
                        Exit For
                    End If

                End If
            Else
                If Checkbox2.Checked Then
                    Errmsg += "組長必須是成員之一" & vbCrLf
                    Exit For
                End If
            End If
        Next

        If intRows > 0 And intCheckbox2 = 0 Then
            Errmsg += "必須有一位是組長" & vbCrLf
        End If

    End Sub

    Sub Insert_OB_Tmember()
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing

        Dim sql As String = ""
        Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            sql = "SELECT * FROM OB_Tmember WHERE tsn='" & ViewState("TSN") & "'"
            dt = DbAccess.GetDataTable(sql, da, trans)
            For Each Item As DataGridItem In DataGrid1.Items
                Dim hidmsn As HtmlInputHidden = Item.FindControl("hidmsn")
                Dim Checkbox1 As HtmlInputCheckBox = Item.FindControl("Checkbox1")
                Dim Checkbox2 As HtmlInputCheckBox = Item.FindControl("Checkbox2")
                If Checkbox1.Checked Then
                    If dt.Select("msn='" & hidmsn.Value & "'").Length > 0 Then
                        dr = dt.Select("msn='" & hidmsn.Value & "'")(0)
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyTime") = Now
                    Else
                        dr = dt.NewRow
                        dt.Rows.Add(dr)
                        dr("msn") = hidmsn.Value
                        dr("tsn") = ViewState("TSN")
                        dr("CreateAcct") = sm.UserInfo.UserID
                        dr("CreateTime") = Now
                        dr("ModifyAcct") = sm.UserInfo.UserID
                        dr("ModifyTime") = Now
                    End If
                    If Checkbox2.Checked Then
                        dr("leader") = "Y"
                    Else
                        dr("leader") = "N"
                    End If
                Else
                    If dt.Select("msn='" & hidmsn.Value & "'").Length > 0 Then
                        dt.Select("msn='" & hidmsn.Value & "'")(0).Delete()
                    End If
                End If
            Next

            DbAccess.UpdateDataTable(dt, da, trans)
            DbAccess.CommitTrans(trans)

            Common.MessageBox(Me, "儲存成功")
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try

    End Sub
End Class
