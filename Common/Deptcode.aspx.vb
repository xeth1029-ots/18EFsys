Partial Class Deptcode
    Inherits AuthBasePage

    '科系所代碼查詢

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        TIMS.OpenDbConn(Me, objconn)
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Me.todo.Value = "" Then
            ShowMainTable()
        Else
            Me.lblMainName.Text = Me.main_name.Value  '顯示所選的大項名稱
            ShowDetailTable()
        End If
    End Sub

    Sub ShowMainTable()
        'Dim strSql As String = ""
        Dim objTable As New DataTable
        Dim daTable As SqlDataAdapter = Nothing
        Dim dr As DataRow = Nothing
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.DEPID ,a.LEVELS ,a.NAME " & vbCrLf
        sql &= " FROM KEY_DEPARTMENT a " & vbCrLf
        sql &= " WHERE 1=1 AND a.Levels = '1' " & vbCrLf
        sql &= " ORDER BY a.DEPID " & vbCrLf
        'strSql = "SELECT * FROM Key_Department WHERE Levels='1' ORDER BY DepID"
        daTable = New SqlDataAdapter(sql, objconn)
        daTable.Fill(objTable)
        Me.MainBlock.Visible = True
        Me.SubBlock.Visible = False

        For Each dr In objTable.Rows
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
            mycell.InnerHtml = String.Format("<a href=""javascript:showDetailTable('{0}','{1}');"">【{1}】{0}</a>", dr("Name"), dr("DepID"))
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub

    Sub ShowDetailTable()
#Region "(No Use)"

        'Dim strSql As String
        'Dim daTable As SqlDataAdapter
        'Dim objTable As New DataTable
        'Dim dr As DataRow
        'Dim myrow As HtmlTableRow
        'Dim mycell As HtmlTableCell
        'Dim i As Integer = 0
        'Dim intShowTableRowCount As Integer = 0
        'strSql = "SELECT * FROM Key_Department" & _
        '         "  WHERE Levels='2' AND DepID In (SELECT DepID FROM Key_Department WHERE Levels='2' AND DepID like '" & Me.main_id.Value & "__')" & _
        '         "  ORDER BY DepID"
        'daTable = New SqlDataAdapter(strSql, objconn)
        'daTable.Fill(objTable)

#End Region
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WKD AS (" & vbCrLf
        sql &= "    SELECT DepID FROM KEY_DEPARTMENT WHERE Levels = '2' AND DepID LIKE @DepID + '__' " & vbCrLf
        sql &= " ) " & vbCrLf
        sql &= " SELECT a.DEPID ,a.LEVELS ,a.NAME " & vbCrLf
        sql &= " FROM KEY_DEPARTMENT a " & vbCrLf
        sql &= " JOIN WKD on wkd.DepID = a.DepID " & vbCrLf
        sql &= " WHERE 1=1 AND a.Levels = '2' " & vbCrLf
        sql &= " ORDER BY a.DEPID " & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim oDt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("DepID", SqlDbType.VarChar).Value = Me.main_id.Value
            oDt.Load(.ExecuteReader())
        End With

        If oDt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Me.MainBlock.Visible = False
        Me.SubBlock.Visible = True
        Dim intShowTableRowCount As Integer = 0
        Dim i As Integer = 0
        Dim myrow As HtmlTableRow = Nothing

        For Each dr As DataRow In oDt.Rows
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
            Dim mycell As HtmlTableCell = New HtmlTableCell
            mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}','{2}','{3}');"">【{2}】{3}</a>", Me.main_id.Value, Me.main_name.Value, dr("DepID"), dr("Name"))
            mycell.Width = "50%"
            If Not myrow Is Nothing Then
                myrow.Cells.Add(mycell)
            End If
            i += 1
        Next
    End Sub
End Class