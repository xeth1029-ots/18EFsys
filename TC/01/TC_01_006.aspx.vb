Partial Class TC_01_006
    Inherits AuthBasePage

    'Dim FunDr As DataRow
    Dim dt2 As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        msg.Text = ""
        Button1.Attributes("onclick") = "javascript:return chk();"

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = "1" Then
        '            Button2.Enabled = True
        '        Else
        '            Button2.Enabled = False
        '        End If
        '        If FunDr("Sech") = "1" Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End       
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '取得表單上的輸入欄位

        Dim kindStr As String = ""
        Dim cateStr As String = ""
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT * FROM ID_KindOfTeacher WHERE 1=1 "
        '& kindStr & cateStr
        Dim name As String = TextBox1.Text
        If name <> "" Then
            sql &= " AND KindName LIKE '%" & name & "%' "
        End If
        If DropDownList1.SelectedValue <> 0 AndAlso DropDownList1.SelectedValue <> "" Then
            Dim kindengage As String = DropDownList1.SelectedValue
            sql &= " AND KindEngage = " & kindengage & " "
        End If
        If DropDownList2.SelectedValue <> 0 AndAlso DropDownList2.SelectedValue <> "" Then
            Dim catekind As String = DropDownList2.SelectedValue
            sql &= " AND CateKind = " & catekind & " "
        End If
        '假如有輸入基本時數，加入SQL的參數search
        Dim minhour As String = Trim(TextBox2.Text)
        Dim maxhour As String = Trim(TextBox3.Text)
        If minhour <> "" AndAlso maxhour <> "" Then
            '不為空轉換為數值。
            sql &= " AND BaseHours BETWEEN " & Val(minhour) & " AND " & Val(maxhour) & " "
        End If
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        '取出Teach_TeacherInfo中的所有資料
        'select * from Teach_TeacherInfo where rownum <=10

        'Dim RID1 As String = sm.UserInfo.RID.ToString.Substring(0, 1)
        'sql = "SELECT DISTINCT KindID FROM TEACH_TEACHERINFO WHERE dbo.SUBSTR(RID,1,1) ='" & RID1 & "'"
        sql = "SELECT DISTINCT KindID FROM TEACH_TEACHERINFO WHERE KindID IS NOT NULL"
        dt2 = DbAccess.GetDataTable(sql, objconn)

        'Dim da As New SqlDataAdapter(sql, objconn)
        'da.Fill(ds, "table1")
        'dv.Table = ds.Tables("table1")
        'da = New SqlDataAdapter(sql, objconn)
        'da.Fill(ds, "table2")

        DataGrid1.Visible = False
        msg.Text = "尚無資料!!"

        If dt.Rows.Count > 0 Then
            DataGrid1.Visible = True
            msg.Text = ""
            DataGrid1.DataSource = dt
            DataGrid1.DataKeyField = "KindID"
            DataGrid1.DataBind()
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        'Response.Redirect("TC_01_006_add.aspx?ID=" & Request("ID") & "")
        '20100208 按新增時代查詢之 師資別名稱
        'Response.Redirect("TC_01_006_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&KindName=" & TextBox1.Text & "")
        TextBox1.Text = TIMS.ClearSQM(TextBox1.Text)
        Dim url1 As String = "TC_01_006_add.aspx?ProcessType=Insert&ID=" & Request("ID") & "&KindName=" & TextBox1.Text & ""
        Call TIMS.Utl_Redirect(Me, objconn, url1)
    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Select Case e.CommandName
            Case "del"
                Dim sKindID As String = TIMS.GetMyValue(e.CommandArgument, "KindID")
                Dim sql As String = "DELETE ID_KINDOFTEACHER WHERE KindID = @KindID "
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("KindID", sKindID)
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
                msg.Text = "刪除成功!"
                DataGrid1.DataBind()
            Case "edit"
                'Response.Redirect("TC_01_006_edit.aspx?ID=" & Request("ID") & "&&serial=" & e.CommandArgument)
                Dim sKindID As String = TIMS.GetMyValue(e.CommandArgument, "KindID")
                Dim url1 As String = "TC_01_006_edit.aspx?ID=" & Request("ID") & "&serial=" & sKindID
                Call TIMS.Utl_Redirect(Me, objconn, url1)
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                'e.Item.Cells(0).Text = e.Item.ItemIndex + 1
                e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + (DataGrid1.CurrentPageIndex * DataGrid1.PageSize)

                Dim drv As DataRowView = e.Item.DataItem

                Dim vKindID As String = Convert.ToString(drv("KindID"))
                Dim vKindName As String = Convert.ToString(drv("KindName"))

                '將按鈕增加屬性
                Dim lbtDel As LinkButton = e.Item.FindControl("lbtDel")
                Dim lbtEdit As LinkButton = e.Item.FindControl("lbtEdit")

                '檢查Teach_TeacherInfo中的關聯性()
                Dim row() As DataRow = dt2.Select("KindID='" & vKindID & "'")
                If row.Length > 0 Then
                    lbtDel.Attributes("onclick") = "javascript:window.alert('此師資類型(" & vKindName & ")下有老師，\n請先刪除此類型(" & vKindName & ")的老師再進行刪除!');return false;"
                Else
                    lbtDel.Attributes("onclick") = "javascript:return confirm('確定要刪除此類型(" & vKindName & ")資料嗎?')"
                End If
                Dim cmdArg As String = ""
                TIMS.SetMyValue(cmdArg, "KindID", Convert.ToString(drv("KindID")))
                lbtDel.CommandArgument = cmdArg 'DataGrid1.DataKeys(e.Item.ItemIndex)
                lbtEdit.CommandArgument = cmdArg 'DataGrid1.DataKeys(e.Item.ItemIndex)

                Dim KindEngageTxt As String = "外聘"
                If Convert.ToString(drv("KindEngage")) = "1" Then KindEngageTxt = "內聘"
                e.Item.Cells(1).Text = KindEngageTxt

                Dim CateKindTxt As String = "外聘"
                Select Case Convert.ToString(drv("CateKind"))
                    Case "1"
                        CateKindTxt = "訓練師類"
                    Case "2"
                        CateKindTxt = "行政人員類"
                End Select
                e.Item.Cells(2).Text = CateKindTxt
        End Select
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs)

    End Sub
End Class