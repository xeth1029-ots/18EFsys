Partial Class JobCode
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        If Me.todo.Value = "" Then
            ShowMainTable()
        Else
            '顯示所選的大項名稱
            Me.lblMainName.Text = Me.main_name.Value
            ShowDetailTable()
        End If
    End Sub

    Sub ShowMainTable()
#Region "(No Use)"

        'Dim strSql As String = ""
        'Dim daTable As SqlDataAdapter = Nothing
        'Dim objTable As New DataTable
        'Dim dr As DataRow = Nothing
        'Dim myrow As HtmlTableRow = Nothing
        'Dim mycell As HtmlTableCell = Nothing
        'Dim i As Integer = 0
        'Dim intShowTableRowCount As Integer = 0
        'strSql = "SELECT * FROM Key_Job WHERE Len(@JobType)=1 ORDER BY JobType"
        'daTable = New SqlDataAdapter(strSql, objconn)
        'daTable.Fill(objTable)

#End Region
        Dim strSql As String = ""
        strSql = " SELECT * FROM Key_Job WHERE LEN(@JobType) = 1 ORDER BY JobType "
        Dim sCmd As New SqlCommand(strSql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            dt.Load(.ExecuteReader())
        End With
        Me.MainBlock.Visible = True
        Me.SubBlock.Visible = False
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0
        For Each dr As DataRow In dt.Rows
            '每列顯示2筆資料
            If i Mod 2 = 0 Then
                intShowTableRowCount += 1
                myrow = New HtmlTableRow
                Me.MainList.Rows.Add(myrow)
                '設定顯示列的背景顏色
                If intShowTableRowCount Mod 2 = 0 Then
                    myrow.BgColor = "#ddddff"
                Else
                    myrow.BgColor = "#ccccff"
                End If
            End If
            mycell = New HtmlTableCell
            mycell.InnerHtml = String.Format("<a href=""javascript:showDetailTable('{0}','{1}');"">【{1}】{0}</a>", dr("Name"), dr("JobType"))
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub

    Sub ShowDetailTable()
#Region "(No Use)"

        'Dim strSql As String = ""
        'Dim objTable As New DataTable
        'Dim daTable As SqlDataAdapter = Nothing
        'Dim dr As DataRow = Nothing
        'Dim myrow As HtmlTableRow = Nothing
        'Dim mycell As HtmlTableCell = Nothing
        'Dim i As Integer = 0
        'Dim intShowTableRowCount As Integer = 0

#End Region
        main_id.Value = TIMS.ClearSQM(main_id.Value)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.JOBID ,a.JOBTYPE ,a.JOBNO ,a.NAME " & vbCrLf
        sql &= " FROM KEY_JOB a " & vbCrLf
        sql &= " WHERE 1=1 AND LEN(a.JobNo) = 4 AND a.JobNo LIKE @JobNo + '%' " & vbCrLf
        sql &= " ORDER BY a.JobNo "
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("JobNo", SqlDbType.VarChar).Value = main_id.Value
            dt.Load(.ExecuteReader())
        End With
#Region "(No Use)"

        'Call CloseDbConn(conn)
        'If dt.Rows.Count > 0 Then Rst = Convert.ToString(dt.Rows(0)("?"))
        'strSql = "SELECT * FROM Key_Job" & _
        '         sql &= " " & vbCrLf
        'daTable = New SqlDataAdapter(strSql, objconn)
        'daTable.Fill(objTable)

#End Region
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        Me.MainBlock.Visible = False
        Me.SubBlock.Visible = True
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0
        For Each dr As DataRow In dt.Rows
            '每列顯示2筆資料
            If i Mod 2 = 0 Then
                intShowTableRowCount += 1
                myrow = New HtmlTableRow
                Me.SubList.Rows.Add(myrow)
                '設定顯示列的背景顏色
                If intShowTableRowCount Mod 2 = 0 Then
                    myrow.BgColor = "#ddddff"
                Else
                    myrow.BgColor = "#ccccff"
                End If
            End If
            mycell = New HtmlTableCell
            mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}','{2}','{3}');"">【{2}】{3}</a>", Me.main_id.Value, Me.main_name.Value, dr("JobNo"), dr("Name"))
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub
End Class