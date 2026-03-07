Partial Class SD_06_004
    Inherits AuthBasePage

    '查詢
    Private Function search() As Boolean
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim dr As DataRow = Nothing

        Dim bolRtn As Boolean = False

        '異動別,陣列值為"Y" 則表示有該判斷狀態 (陣列: 0=>加保, 1=>退保, 2=>無異動)
        Dim strType() As String = {"N", "N", "N"}
        Dim datApply As DateTime = Nothing '加保日期
        Dim datOut As DateTime = Nothing '退保日期

        Try
            'conn.Open()

            '查詢學員資料
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql += " select a.ocid,a.classcname,b.socid,c.sid,c.name,a.stdate,a.ftdate,c.idno,c.birthday,d.budname" & vbCrLf
            sql += " ,e.applyinsurance,e.dropoutinsurance,'' type,'' applydate,'' outdate" & vbCrLf
            sql += " from class_classinfo a" & vbCrLf
            sql += " join class_studentsofclass b on b.ocid=a.ocid" & vbCrLf
            sql += " join stud_studentinfo c on c.sid=b.sid" & vbCrLf
            sql += " left join key_budget d on d.budid=b.budgetid" & vbCrLf
            sql += " join stud_insurance e on e.socid=b.socid" & vbCrLf
            sql += " where planid= @planid "
            sql += " and (e.applyinsurance between @sdate and @edate or e.dropoutinsurance between @sdate and @edate) "
            sql += " order by a.ocid,c.sid"

            With sda
                .SelectCommand = New SqlCommand(sql, objconn)
                .SelectCommand.Parameters.Clear()
                .SelectCommand.Parameters.Add("planid", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
                .SelectCommand.Parameters.Add("sdate", SqlDbType.VarChar).Value = txtSDate.Text
                .SelectCommand.Parameters.Add("edate", SqlDbType.VarChar).Value = txtEDate.Text
                .Fill(ds)
            End With

            '依查詢日期起迄判斷 異動別 & 加退保日顯示
            For i As Integer = 0 To ds.Tables(0).Rows.Count - 1
                dr = ds.Tables(0).Rows(i)

                strType(0) = "N" : strType(1) = "N" : strType(2) = "N"
                datApply = Nothing
                datOut = Nothing

                If Convert.ToString(dr("applyinsurance")) <> "" Then datApply = Convert.ToDateTime(dr("applyinsurance"))
                If Convert.ToString(dr("dropoutinsurance")) <> "" Then datOut = Convert.ToDateTime(dr("dropoutinsurance"))

                '加保日 = 查詢起日, 異動別=>加保
                If datApply = Convert.ToDateTime(txtSDate.Text) Then
                    strType(0) = "Y"
                End If

                '加保日 > 查詢起日, 異動別=>無異動
                If datApply > Convert.ToDateTime(txtSDate.Text) Then
                    strType(2) = "Y"
                End If

                '退保日 界於 查詢起迄日, 異動別=>退保
                If datOut >= Convert.ToDateTime(txtSDate.Text) And datOut <= Convert.ToDateTime(txtEDate.Text) Then
                    strType(1) = "Y"
                End If

                Select Case strType(0) & strType(1) & strType(2)
                    Case "YYY", "YYN"
                        dr("type") = "0,1"
                        dr("applydate") = dr("applyinsurance")
                        dr("outdate") = dr("dropoutinsurance")
                    Case "YNN", "YNY"
                        dr("type") = "0"
                        dr("applydate") = dr("applyinsurance")
                    Case "NYY", "NYN"
                        dr("type") = "1"
                        dr("outdate") = dr("dropoutinsurance")
                    Case Else
                        dr("type") = "2"
                End Select
            Next
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            '依查詢條件(狀態), 重組資料
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            If rblType.SelectedValue <> "" Then
                For i As Integer = ds.Tables(0).Rows.Count - 1 To 0 Step -1
                    dr = ds.Tables(0).Rows(i)

                    If Convert.ToString(dr("type")).IndexOf(rblType.SelectedValue) < 0 Then
                        ds.Tables(0).Rows(i).Delete()
                    End If
                Next
            End If
            ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

            If ds.Tables(0).Rows.Count > 0 Then
                bolRtn = True
                labMsg.Visible = False

                DataGrid1.DataSource = ds.Tables(0)
                DataGrid1.DataBind()
                DataGrid1.Visible = True
            Else
                bolRtn = False

                labMsg.Visible = True
                DataGrid1.Visible = False
            End If

            'conn.Close()
            If Not sda Is Nothing Then sda.Dispose()
            If Not ds Is Nothing Then ds.Dispose()
        Catch ex As Exception
            Common.MessageBox(Me, "系統錯誤:" & ex.ToString)
        End Try

        Return bolRtn
    End Function

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            rblType.SelectedValue = ""
            div1.Visible = False

            btnExport.Attributes.Add("onclick", "return chkExport();")
        End If
    End Sub

    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        If search() Then
            div1.Visible = True

            labPDataArea.Text = txtSDate.Text & "~" & txtEDate.Text
            labPType.Text = rblType.SelectedItem.Text

            Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")

            '提示使用者是否要儲存檔案
            Dim sFileName As String = ""
            sFileName = HttpUtility.UrlEncode("投保資料.xls", System.Text.Encoding.UTF8)

            Response.AddHeader("content-disposition", "attachment; filename=" & sFileName)

            '文件內容指定為excel
            Response.ContentType = "application/ms-excel;charset=utf-8"

            '繪出要輸出的html內容
            Dim strContent As New System.Text.StringBuilder
            Dim stringWrite As New System.IO.StringWriter(strContent)
            Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)

            div1.RenderControl(htmlWrite)
            strContent.Replace("<html>", "")
            strContent.Replace("</html>", "")
            strContent.Replace("<a", "<span")
            strContent.Replace("</a>", "</span>")
            strContent.Replace("<input", "<span")
            Common.RespWrite(Me, "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>")

            ''套CSS值
            Common.RespWrite(Me, "<style>")
            Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
            Common.RespWrite(Me, "</style>")

            Common.RespWrite(Me, strContent)
            Common.RespWrite(Me, "</html>")

            '結束程式執行
            Response.End()
        End If
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.EditItem, ListItemType.Item, ListItemType.SelectedItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim labClassName As Label = e.Item.FindControl("labClassName")
                Dim labStdName As Label = e.Item.FindControl("labStdName")
                Dim labSTDate As Label = e.Item.FindControl("labSTDate")
                Dim labETDate As Label = e.Item.FindControl("labETDate")
                Dim labIDNO As Label = e.Item.FindControl("labIDNO")
                Dim labBirth As Label = e.Item.FindControl("labBirth")
                Dim labBud As Label = e.Item.FindControl("labBud")
                Dim labType As Label = e.Item.FindControl("labType")
                Dim labAppDate As Label = e.Item.FindControl("labAppDate")
                Dim labOutDate As Label = e.Item.FindControl("labOutDate")

                labClassName.Text = Convert.ToString(drv("classcname"))
                labStdName.Text = Convert.ToString(drv("name"))

                If Convert.ToString(drv("stdate")) <> "" Then
                    labSTDate.Text = Convert.ToDateTime(drv("stdate")).ToString("yyyy/MM/dd")
                End If

                If Convert.ToString(drv("ftdate")) <> "" Then
                    labETDate.Text = Convert.ToDateTime(drv("ftdate")).ToString("yyyy/MM/dd")
                End If

                labIDNO.Text = Convert.ToString(drv("idno"))

                If Convert.ToString(drv("birthday")) <> "" Then
                    labBirth.Text = Convert.ToDateTime(drv("birthday")).ToString("yyyy/MM/dd")
                End If

                labBud.Text = Convert.ToString(drv("budname"))

                Select Case Convert.ToString(drv("type"))
                    Case "0,1"
                        labType.Text = "加/退保"
                    Case "0"
                        labType.Text = "加保"
                    Case "1"
                        labType.Text = "退保"
                    Case "2"
                        labType.Text = "無異動"
                End Select

                If Convert.ToString(drv("applydate")) <> "" Then
                    labAppDate.Text = Convert.ToDateTime(drv("applydate")).ToString("yyyy/MM/dd")
                End If

                If Convert.ToString(drv("outdate")) <> "" Then
                    labOutDate.Text = Convert.ToDateTime(drv("outdate")).ToString("yyyy/MM/dd")
                End If
        End Select
    End Sub

End Class
