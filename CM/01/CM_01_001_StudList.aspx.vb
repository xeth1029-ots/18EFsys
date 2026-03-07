Partial Class CM_01_001_StudList
    Inherits AuthBasePage

    Dim Key_Identity As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        Key_Identity = TIMS.Get_KeyTable("Key_Identity", "", objconn)
        ListMode1.Visible = False
        ListMode2.Visible = False
        ListMode3.Visible = False

        Select Case Request("Mode")
            Case "3"
                ListMode1.Visible = True
            Case "4"
                ListMode2.Visible = True
            Case "5"
                ListMode3.Visible = True
        End Select
        If Not IsPostBack Then
            Create()
        End If
        Button2.Attributes("onclick") = "return CheckData3();"

    End Sub

    Sub Create()
        Dim sql As String
        Dim dt As DataTable

        Select Case Request("Mode")
            Case "3"
                sql = ""
                sql &= " SELECT a.SOCID,a.StudentID,a.BudgetID,a.PMode,b.Name FROM "
                sql += " (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "') a "
                sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID "
                sql += " Order By a.StudentID"

                dt = DbAccess.GetDataTable(sql)
                If dt.Rows.Count = 0 Then
                Else
                    DataGrid1.DataSource = dt
                    DataGrid1.DataKeyField = "SOCID"
                    DataGrid1.DataBind()
                End If
            Case "4"
                sql = ""
                sql &= " SELECT a.SOCID,a.StudentID,a.BudgetID,a.RelClass_Unit,b.Name FROM "
                sql += " (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "') a "
                sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID "
                sql += " Order By a.StudentID"

                dt = DbAccess.GetDataTable(sql)
                If dt.Rows.Count = 0 Then
                Else
                    DataGrid2.DataSource = dt
                    DataGrid2.DataKeyField = "SOCID"
                    DataGrid2.DataBind()
                End If
            Case "5"
                sql = ""
                sql &= " SELECT a.SOCID,a.StudentID,a.IdentityID,a.MIdentityID,c.Name,b.SUBID FROM "
                sql += " (SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "') a "
                sql += " LEFT JOIN Stud_SubsidyResult b ON a.SOCID=b.SOCID "
                sql += " JOIN Stud_StudentInfo c ON a.SID=c.SID "
                sql += " Order By a.StudentID"

                dt = DbAccess.GetDataTable(sql)
                If dt.Rows.Count = 0 Then
                Else
                    DataGrid3.DataSource = dt
                    DataGrid3.DataKeyField = "SOCID"
                    DataGrid3.DataBind()
                End If
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Dim Radio2 As HtmlInputRadioButton = e.Item.FindControl("Radio2")
                Dim Radio3 As HtmlInputRadioButton = e.Item.FindControl("Radio3")
                Dim Radio4 As HtmlInputRadioButton = e.Item.FindControl("Radio4")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                e.Item.Cells(0).Text = Right(drv("StudentID"), 2)
                Select Case drv("BudgetID").ToString
                    Case "02"
                        Select Case drv("PMode").ToString
                            Case "1"
                                Radio2.Checked = True
                            Case "2"
                                Radio1.Checked = True
                        End Select
                    Case "03"
                        Select Case drv("PMode").ToString
                            Case "1"
                                Radio4.Checked = True
                            Case "2"
                                Radio3.Checked = True
                        End Select
                End Select
        End Select
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        Dim Total1 As Integer
        Dim Total2 As Integer
        Dim Total3 As Integer
        Dim Total4 As Integer
        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn, True)

        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "'"
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For Each item As DataGridItem In DataGrid1.Items
            Dim Radio1 As HtmlInputRadioButton = item.FindControl("Radio1")
            Dim Radio2 As HtmlInputRadioButton = item.FindControl("Radio2")
            Dim Radio3 As HtmlInputRadioButton = item.FindControl("Radio3")
            Dim Radio4 As HtmlInputRadioButton = item.FindControl("Radio4")

            dr = dt.Select("SOCID='" & DataGrid1.DataKeys(item.ItemIndex) & "'")(0)
            If Radio2.Checked Then
                dr("BudgetID") = "02"
                dr("PMode") = "1"
                Total3 += 1
            End If
            If Radio1.Checked Then
                dr("BudgetID") = "02"
                dr("PMode") = "2"
                Total4 += 1
            End If
            If Radio4.Checked Then
                dr("BudgetID") = "03"
                dr("PMode") = "1"
                Total1 += 1
            End If
            If Radio3.Checked Then
                dr("BudgetID") = "03"
                dr("PMode") = "2"
                Total2 += 1
            End If

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
        Next

        DbAccess.UpdateDataTable(dt, da)
        'If conn.State = ConnectionState.Open Then conn.Close()
        Common.RespWrite(Me, "<script>opener.document.getElementById('Num1').value=" & Total1 & ";opener.document.getElementById('Num2').value=" & Total2 & ";opener.document.getElementById('Num3').value=" & Total3 & ";opener.document.getElementById('Num4').value=" & Total4 & ";window.close();</script>")
    End Sub

    Private Sub DataGrid3_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid3.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim Correct As HtmlInputCheckBox = e.Item.FindControl("Correct")
                Dim MIdentityID1 As HtmlInputHidden = e.Item.FindControl("MIdentityID1")
                Dim MIdentityID As DropDownList = e.Item.FindControl("MIdentityID")

                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = "CM_TD2"
                End If

                e.Item.Cells(0).Text = Right(drv("StudentID"), 2)
                Select Case drv("MIdentityID").ToString
                    Case "02", "03", "04", "05", "06", "07", "08", "09", "10", "13", "14", "17", "18"
                        e.Item.Cells(3).Text = "●"
                    Case ""
                    Case Else
                        e.Item.Cells(2).Text = "●"
                End Select

                Correct.Attributes("onclick") = "document.getElementById('" & MIdentityID.ClientID & "').disabled=!this.checked"
                MIdentityID1.Value = drv("MIdentityID").ToString
                MIdentityID.Enabled = False
                MIdentityID.Items.Add(New ListItem("請選擇", ""))
                For i As Integer = 0 To Split(drv("IdentityID"), ",").Length - 1
                    MIdentityID.Items.Add(New ListItem(Key_Identity.Select("IdentityID='" & Split(drv("IdentityID"), ",")(i) & "'")(0)("Name"), Split(drv("IdentityID"), ",")(i)))
                Next
                Common.SetListItem(MIdentityID, drv("MIdentityID"))
        End Select
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        Dim Total1 As Integer
        Dim Total2 As Integer
        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn, True)

        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "'"
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For Each item As DataGridItem In DataGrid3.Items
            Dim Correct As HtmlInputCheckBox = item.FindControl("Correct")
            Dim MIdentityID1 As HtmlInputHidden = item.FindControl("MIdentityID1")
            Dim MIdentityID As DropDownList = item.FindControl("MIdentityID")

            If Correct.Checked Then
                dr = dt.Select("SOCID='" & DataGrid3.DataKeys(item.ItemIndex) & "'")(0)

                dr("MIdentityID") = MIdentityID.SelectedValue
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now

                Select Case MIdentityID.SelectedValue
                    Case "02", "03", "04", "05", "06", "07", "08", "09", "10", "13", "14", "17", "18"
                        Total2 += 1
                    Case ""
                    Case Else
                        Total1 += 1
                End Select
            Else
                Select Case MIdentityID1.Value
                    Case "02", "03", "04", "05", "06", "07", "08", "09", "10", "13", "14", "17", "18"
                        Total2 += 1
                    Case ""
                    Case Else
                        Total1 += 1
                End Select
            End If
        Next

        DbAccess.UpdateDataTable(dt, da)
        'If conn.State = ConnectionState.Open Then conn.Close()
        Common.RespWrite(Me, "<script>opener.document.getElementById('GNum').value=" & Total1 & ";opener.document.getElementById('SNum').value=" & Total2 & ";window.close();</script>")
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "CM_TD1"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim BudgetID As RadioButtonList = e.Item.FindControl("BudgetID")
                Dim RelClass_Unit1 As HtmlInputCheckBox = e.Item.FindControl("RelClass_Unit1")
                Dim RelClass_Unit2 As HtmlInputCheckBox = e.Item.FindControl("RelClass_Unit2")
                Dim RelClass_Unit3 As HtmlInputCheckBox = e.Item.FindControl("RelClass_Unit3")

                e.Item.Cells(0).Text = Right(drv("StudentID"), 2)
                Common.SetListItem(BudgetID, drv("BudgetID"))
                If drv("RelClass_Unit").ToString.Length = 3 Then
                    If drv("RelClass_Unit").ToString.Chars(0) = "1" Then
                        RelClass_Unit1.Checked = True
                    End If
                    If drv("RelClass_Unit").ToString.Chars(1) = "1" Then
                        RelClass_Unit2.Checked = True
                    End If
                    If drv("RelClass_Unit").ToString.Chars(2) = "1" Then
                        RelClass_Unit3.Checked = True
                    End If
                End If
        End Select
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim da As SqlDataAdapter = Nothing
        'Dim Total1 As Integer
        'Dim Total2 As Integer

        'Dim conn As SqlConnection
        'TIMS.TestDbConn(Me, conn, True)
        sql = "SELECT * FROM Class_StudentsOfClass WHERE OCID='" & Request("OCID") & "'"
        '2006/03/28 add conn by matt
        dt = DbAccess.GetDataTable(sql, da, objconn)

        For Each item As DataGridItem In DataGrid2.Items
            Dim BudgetID As RadioButtonList = item.FindControl("BudgetID")
            Dim RelClass_Unit1 As HtmlInputCheckBox = item.FindControl("RelClass_Unit1")
            Dim RelClass_Unit2 As HtmlInputCheckBox = item.FindControl("RelClass_Unit2")
            Dim RelClass_Unit3 As HtmlInputCheckBox = item.FindControl("RelClass_Unit3")
            Dim RelClass_Unit As String

            dr = dt.Select("SOCID='" & DataGrid2.DataKeys(item.ItemIndex) & "'")(0)
            dr("BudgetID") = BudgetID.SelectedValue
            If RelClass_Unit1.Checked = True Then
                RelClass_Unit = 1
            Else
                RelClass_Unit = 0
            End If
            If RelClass_Unit2.Checked = True Then
                RelClass_Unit = RelClass_Unit & 1
            Else
                RelClass_Unit = RelClass_Unit & 0
            End If
            If RelClass_Unit3.Checked = True Then
                RelClass_Unit = RelClass_Unit & 1
            Else
                RelClass_Unit = RelClass_Unit & 0
            End If
            dr("RelClass_Unit") = RelClass_Unit

            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
        Next

        DbAccess.UpdateDataTable(dt, da)
        'If conn.State = ConnectionState.Open Then conn.Close()
        Common.RespWrite(Me, "<script>opener.document.getElementById('Button18').click();window.close();</script>")
    End Sub
End Class
