Partial Class SD_05_009_List
    Inherits AuthBasePage

    'Dim sqldr As DataRow
    Dim intSum As Integer = 0
    Dim IntAllSum As Integer = 0

    Dim ProcessType As String = ""

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        objconn = DbAccess.GetConnection()
        ProcessType = Request("ProcessType")

        'Dim i
        If Not IsPostBack Then
            years = TIMS.GetSyear(years)
            months.Items.Add(New ListItem("==請選擇==", 0))
            For i As Integer = 1 To 12
                months.Items.Add(i)
            Next
            'If ProcessType = "Back" Then
            '    OCIDValue1.Value = Request("OCID")
            '    months.SelectedValue = Convert.ToDouble(Request("Month"))
            '    years.SelectedValue = Convert.ToDouble(Request("Year"))

            '    Button1_Click(sender, e)
            'End If

        End If
        Button1.Attributes("onclick") = "javascript:return print();"
        count_Button.Attributes("onclick") = "javascript:return check();"
        Button2.Attributes("onclick") = "history.go(-1);return false;"

        If openwin.Value = "1" Then
            Dim Myteacher As String = ""
            openwin.Value = 0
            If Not RB_Teacher_List.SelectedItem Is Nothing Then
                Myteacher = RB_Teacher_List.SelectedValue
            End If
            Button1_Click(sender, e)
            Common.SetListItem(RB_Teacher_List, Myteacher)
            count_Button_Click(sender, e)
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select DISTINCT a.TechID,b.TeachCName " & vbCrLf
        sql += " from Teach_PayHour a  " & vbCrLf
        sql += " join Teach_TeacherInfo b  on a.TechID=b.TechID" & vbCrLf
        sql += " Where  1=1" & vbCrLf
        sql += " AND a.OCID='" & OCIDValue1.Value & "' " & vbCrLf
        sql += " and DATEPART(YEAR, a.TeachDate)='" & years.SelectedValue & "' " & vbCrLf
        sql += " and DATEPART(MONTH, a.TeachDate)='" & months.SelectedValue & "' " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無資料!!"
        Me.Panel1.Visible = False
        Me.Panel.Visible = True
        Me.Panel2.Visible = False
        msg2.Text = ""
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            Me.Panel1.Visible = False
            Me.Panel.Visible = True
            Me.Panel2.Visible = False
            msg2.Text = ""

            Me.RB_Teacher_List.DataSource = dt
            Me.RB_Teacher_List.DataTextField = "TeachCName"
            Me.RB_Teacher_List.DataValueField = "TechID"
            Me.RB_Teacher_List.DataBind()
        End If

    End Sub

    Private Sub count_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles count_Button.Click
        Me.TechID_str.Value = RB_Teacher_List.SelectedValue
        Me.Year_str.Value = years.SelectedValue
        Me.Month_str.Value = months.SelectedValue
        Me.Add_Button.Attributes.Add("onclick", "javascript:wopen('SD_05_009_detail_add.aspx?ocid='+ document.getElementById('OCIDValue1').value+'&Month='+document.getElementById('Month_str').value+'&Year='+document.getElementById('Year_str').value+'&TechID='+document.getElementById('TechID_str').value,'新增講師上課明細',350,250,0);document.form1.openwin.value=1;")
        'Me.Add_Button.Attributes.Add("onclick", "javascript:wopen('SD_05_009_detail_add.aspx','新增講師上課明細',350,250,0);document.form1.openwin.value=1;")
        LoadData()
    End Sub

    Sub LoadData()
        Dim Teacher_list As String = "select a.TechID,b.TeachCName,a.UnitPrice from Teach_PayHour a join Teach_TeacherInfo b  on a.TechID=b.TechID  where a.OCID='" & OCIDValue1.Value & "' and a.TechID='" & RB_Teacher_List.SelectedValue & "' and Year(a.TeachDate)='" & years.SelectedValue & "' and MONTH(a.TeachDate)='" & months.SelectedValue & "' group by a.UnitPrice,a.TechID,b.TeachCName"
        Dim TeaList As DataTable = DbAccess.GetDataTable(Teacher_list, objconn)
        DG_Teacher.DataSource = TeaList
        DG_Teacher.DataBind()

        If TeaList.Rows.Count = 0 Then
            Me.Panel2.Visible = False
            msg2.Text = "查無資料!!"
            Me.Panel1.Visible = True
        Else
            Me.Panel1.Visible = False
            Me.Panel2.Visible = True
            msg2.Text = ""
            Dim sqldr As DataRow = TeaList.Rows(0)
            Me.Label1.Text = sqldr("TeachCName") + "，上課明細表 "
        End If
    End Sub

    Private Sub DG_Teacher_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DG_Teacher.ItemDataBound
        Dim dr As DataRowView = e.Item.DataItem
        Dim dg As DataGrid
        Dim objPrice As TextBox
        Dim objLabel As Label
        Select Case e.Item.ItemType
            Case ListItemType.Header
                IntAllSum = 0

            Case ListItemType.Item, ListItemType.AlternatingItem

                dg = e.Item.FindControl("DG_Prices")
                objPrice = e.Item.FindControl("Text_Price")
                objPrice.Text = dr("UnitPrice")
                Dim teacher_sql As String
                teacher_sql = "select a.Teach_Pay_ID,a.TechID,b.TeachCName,a.UnitHour,a.UnitPrice,a.TeachDate  from Teach_PayHour a  join Teach_TeacherInfo b  on a.TechID=b.TechID where a.OCID='" & OCIDValue1.Value & "' and a.TechID='" & RB_Teacher_List.SelectedValue & "' and Year(a.TeachDate)='" & years.SelectedValue & "' and MONTH(a.TeachDate)='" & months.SelectedValue & "' and UnitPrice='" & dr("UnitPrice") & "' order by a.TeachDate"
                Dim dthours As DataTable = DbAccess.GetDataTable(teacher_sql, objconn)
                dg.DataSource = dthours
                dg.DataBind()

            Case ListItemType.Footer
                objLabel = e.Item.FindControl("lblAllSum")
                objLabel.Text = IntAllSum


        End Select
    End Sub

    Public Sub DG_Prices_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs)
        Dim dr As DataRowView = e.Item.DataItem
        Dim objTextbox As TextBox
        Dim objLabel As Label
        Dim objLink As HtmlAnchor

        Select Case e.Item.ItemType
            Case ListItemType.Header
                intSum = 0

            Case ListItemType.Item, ListItemType.AlternatingItem
                objTextbox = e.Item.FindControl("Text_Date")
                objLink = e.Item.FindControl("linkDate")

                objLink.HRef = "javascript:show_calendar('" & objTextbox.ClientID & "','','','CY/MM/DD');"
                intSum += dr("Unitprice") * dr("UnitHour")

            Case ListItemType.Footer
                objLabel = e.Item.FindControl("lblSum")
                objLabel.Text = intSum
                IntAllSum += intSum

        End Select
    End Sub

    Public Sub DG_Prices_DeleteCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs)
        Dim objDate As TextBox
        Dim strPrice As String

        'TechID = Request("TechID")
        'ocid = Request("OCID")
        Dim sql_del As String
        objDate = e.Item.FindControl("Text_Date")
        strPrice = e.Item.Cells(3).Text

        sql_del = "delete Teach_PayHour where  OCID='" & OCIDValue1.Value & "'  and TechID='" & RB_Teacher_List.SelectedValue & "' and TeachDate='" & objDate.Text & "' and UnitPrice='" & strPrice & "' "
        DbAccess.ExecuteNonQuery(sql_del, objconn)
        Common.MessageBox(Page, "講師鐘點費刪除成功!!")
        LoadData()
    End Sub

    Private Sub save_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles save_Button.Click
        'Dim MainItem As DataGridItem
        'Dim SubItem As DataGridItem
        'Dim dgPrice As DataGrid
        'Dim sqlAdapter As SqlDataAdapter
        'Dim objPrice As TextBox
        'Dim objdate As TextBox
        'Dim objhour As TextBox
        'Dim Oldstr, newdatestr As String
        'Dim rows() As DataRow
        'Dim teach_Pay_id As String
        'Dim dr As DataRow
        'Dim drow As DataRow

        'Dim strMessage_date As String = ""
        'Dim DuplRows As New ArrayList
        'Dim i As Integer
        Dim ff3 As String = ""
        Dim DuplRows As New ArrayList
        Dim strMessage1 As String = ""
        Dim strMessage_date As String = ""

        Dim Oldstr As String = ""
        Oldstr = "select *  from Teach_PayHour  where OCID='" & OCIDValue1.Value & "' and TechID='" & RB_Teacher_List.SelectedValue & "' and Year(TeachDate)='" & years.SelectedValue & "' and MONTH(TeachDate)='" & months.SelectedValue & "'"
        Dim sqlAdapter As SqlDataAdapter = Nothing
        Dim dtPrice As DataTable = Nothing
        dtPrice = DbAccess.GetDataTable(Oldstr, sqlAdapter, objconn) '未更動前的DataTable
        For Each MainItem As DataGridItem In Me.DG_Teacher.Items
            Dim dgPrice As DataGrid = MainItem.FindControl("DG_Prices")
            Dim objPrice As TextBox = MainItem.FindControl("Text_price")

            For Each SubItem As DataGridItem In dgPrice.Items
                Dim strOldDate As String = ""
                Dim strOldUnitPrice As String = ""
                Dim newdatestr As String = ""
                Dim teach_Pay_id As String = ""
                Dim objdate As TextBox = SubItem.FindControl("Text_Date")
                Dim objhour As TextBox = SubItem.FindControl("Text_hour")
                newdatestr = Convert.ToDateTime(objdate.Text)
                teach_Pay_id = SubItem.Cells(5).Text
                strOldDate = Convert.ToDateTime(SubItem.Cells(4).Text)
                strOldUnitPrice = SubItem.Cells(3).Text
                ff3 = "OCID='" & OCIDValue1.Value & "' and TechID='" & RB_Teacher_List.SelectedValue & "' and TeachDate='" & strOldDate & "' and UnitPrice='" & strOldUnitPrice & "'"
                Dim rows() As DataRow = dtPrice.Select(ff3)
                If rows.Length = 0 Then Exit For
                Dim dr As DataRow = rows(0)
                dr("UnitPrice") = objPrice.Text
                dr("UnitHour") = objhour.Text
                dr("TeachDate") = objdate.Text

                If newdatestr.Substring(0, 7) <> strOldDate.Substring(0, 7) Then
                    strMessage_date = "日期請選擇" & strOldDate.Substring(0, 7) & "區間"
                End If

                ff3 = "OCID='" & OCIDValue1.Value & "' and TechID='" & RB_Teacher_List.SelectedValue & "' and TeachDate='" & dr("TeachDate") & "' and UnitPrice='" & dr("UnitPrice") & "'"
                rows = dtPrice.Select(ff3)

                If rows.Length > 1 Then '檢查是否重複
                    DuplRows.Add(dr)
                    For Each drow As DataRow In DuplRows
                        If drow("TeachDate") <> dr("TeachDate") Or strMessage1 = "" Then
                            strMessage1 &= "," & dr("TeachDate")
                        End If
                    Next
                End If
            Next
        Next

        If strMessage_date <> "" Then
            Common.MessageBox(Page, strMessage_date)
            Exit Sub
        End If

        If strMessage1 <> "" And Request("save_type") Is Nothing Then '確認的錯誤訊息
            Dim totalstring As String = strMessage1.Substring(1)
            Page.RegisterHiddenField("save_type", "")
            Common.AddClientScript(Page, "if (window.confirm('" & totalstring & "鍾點費重複,是否合併此筆資料?')) {")
            Common.AddClientScript(Page, "  form1.save_type.value='replace';")
            Common.AddClientScript(Page, "  form1.save_Button.click();")
            Common.AddClientScript(Page, " } else { ")
            Common.AddClientScript(Page, "  form1.save_type.value='stop';")
            Common.AddClientScript(Page, "  form1.save_Button.click();")
            Common.AddClientScript(Page, "}")
            Exit Sub
        End If

        If Request("save_type") = "replace" Then '若點選合併
            For Each dr As DataRow In DuplRows
                If dr.RowState <> DataRowState.Deleted Then
                    Dim rows() As DataRow = dtPrice.Select("OCID='" & OCIDValue1.Value & "' and TechID='" & RB_Teacher_List.SelectedValue & "' and TeachDate='" & dr("TeachDate") & "' and UnitPrice='" & dr("UnitPrice") & "'")
                    For i As Integer = 1 To rows.Length - 1
                        rows(0)("UnitHour") += rows(i)("UnitHour")
                        rows(i).Delete()
                    Next
                End If
            Next
        ElseIf Request("save_type") = "stop" Then '點選取消
            For Each dr As DataRow In DuplRows
                If dr.RowState <> DataRowState.Detached Then
                    Dim rows() As DataRow = dtPrice.Select("OCID='" & OCIDValue1.Value & "' and TechID='" & RB_Teacher_List.SelectedValue & "' and TeachDate='" & dr("TeachDate") & "' and UnitPrice='" & dr("UnitPrice") & "'")
                    For i As Integer = 0 To rows.Length - 1
                        dtPrice.Rows.Remove(rows(i))
                    Next
                End If
            Next
        End If

        DbAccess.UpdateDataTable(dtPrice, sqlAdapter)
        '更新成功訊息
        Common.MessageBox(Page, "講師鍾點費更新成功!!")
        LoadData()
    End Sub

End Class
