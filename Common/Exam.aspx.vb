Partial Class Exam
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        TIMS.OpenDbConn(Me, objconn)
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        ShowMainTable()
        'If Me.todo.Value = "" Then
        '    ShowMainTable()
        'Else
        '    '顯示所選的大項名稱
        '    Me.lblMainName.Text = Me.main_name.Value
        '    ShowDetailTable()
        'End If
    End Sub

    Sub ShowMainTable()
        Dim strSql As String = ""
        Dim daTable As SqlDataAdapter = Nothing
        Dim objTable As New DataTable
        Dim dr As DataRow = Nothing
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0

        strSql = " SELECT * FROM Key_Exam ORDER BY ExamID "
        daTable = New SqlDataAdapter(strSql, objconn)
        daTable.Fill(objTable)
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
            mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}');"">【{0}】{1}</a>", dr("ExamID"), dr("Name"))
            'mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}','{2}','{3}');"">【{2}】{3}</a>", Me.main_id.Value, Me.main_name.Value, dr("DepID"), dr("Name"))
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub

#Region "(No Use)"

    'Sub ShowDetailTable()
    '    Dim strSql As String
    '    Dim daTable As SqlDataAdapter
    '    Dim objTable As New DataTable
    '    Dim dr As DataRow
    '    Dim myrow As HtmlTableRow
    '    Dim mycell As HtmlTableCell
    '    Dim i As Integer = 0
    '    Dim intShowTableRowCount As Integer = 0

    '    strSql = "SELECT * FROM Key_Department" & _
    '             "  WHERE Levels='2' AND DepID=(SELECT DepID FROM Key_Department WHERE Levels='1' AND DepID like '" & Me.main_id.Value & "__')" & _
    '             "  ORDER BY DepID"
    '    daTable = New SqlDataAdapter(strSql, Me.objConnection)
    '    daTable.Fill(objTable)

    '    Me.MainBlock.Visible = False
    '    Me.SubBlock.Visible = True

    '    For Each dr In objTable.Rows
    '        '每列顯示2筆資料
    '        If i Mod 2 = 0 Then
    '            intShowTableRowCount += 1
    '            myrow = New HtmlTableRow
    '            Me.SubList.Rows.Add(myrow)

    '            '設定顯示列的背景顏色
    '            If intShowTableRowCount Mod 2 = 0 Then
    '                myrow.BgColor = "#ddddff"
    '            Else
    '                myrow.BgColor = "#ccccff"
    '            End If
    '        End If
    '        mycell = New HtmlTableCell
    '        mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}','{2}','{3}');"">【{2}】{3}</a>", Me.main_id.Value, Me.main_name.Value, dr("DepID"), dr("Name"))
    '        mycell.Width = "50%"
    '        myrow.Cells.Add(mycell)

    '        i += 1
    '    Next
    'End Sub

#End Region
End Class