Partial Class ProSkill
    Inherits AuthBasePage

    Dim Key_ProSkill As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Key_ProSkill = TIMS.Get_KeyTable("Key_ProSkill", "", objconn)

        If SecID.Value <> "" Then
            ShowThirdTable()
        ElseIf MainID.Value <> "" Then
            ShowDetailTable()
        Else
            ShowMainTable()
        End If
    End Sub

    Sub ShowMainTable()
        'Dim strSql As String = ""
        'Dim daTable As SqlDataAdapter
        Dim objTable As New DataTable
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0

        Dim strSql As String = ""
        strSql = " SELECT * FROM Key_ProSkill WHERE Levels = 1 "
        objTable = DbAccess.GetDataTable(strSql, objconn)
        'daTable = New SqlDataAdapter(strSql, objconn)
        'daTable.Fill(objTable)
        Me.MainBlock.Style("display") = "inline"
        Me.SubBlock.Style("display") = "none"
        Me.ThirdBlock.Style("display") = "none"

        For Each dr As DataRow In objTable.Rows
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
            mycell.InnerHtml = String.Format("<a href=""javascript:showDetailTable('{0}','{1}','{2}');"">【{1}】{2}</a>", dr("KPID"), dr("ProID"), dr("ProName"))
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub

    Sub ShowDetailTable()
        Dim sql As String
        'Dim daTable As SqlDataAdapter
        Dim objTable As New DataTable
        Dim dr As DataRow = Nothing
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0

        sql = " SELECT * FROM Key_ProSkill WHERE Parent = '" & Me.MainID.Value & "' "
        objTable = DbAccess.GetDataTable(sql)

        Me.MainBlock.Style("display") = "none"
        Me.SubBlock.Style("display") = "inline"
        Me.ThirdBlock.Style("display") = "none"
        MainNum.Text = HidMainNum.Value
        MainName.Text = HidMainName.Value

        For Each dr In objTable.Rows
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
            If Key_ProSkill.Select("[Parent]='" & dr("KPID") & "'").Length = 0 Then
                mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}','{2}');"">【{1}】{2}</a>", dr("KPID"), dr("ProID"), dr("ProName"))
            Else
                mycell.InnerHtml = String.Format("<a href=""javascript:showThridTable('{0}','{1}','{2}');"">【{1}】{2}↓</a>", dr("KPID"), dr("ProID"), dr("ProName"))
            End If
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub

    Sub ShowThirdTable()
        Me.MainBlock.Style("display") = "none"
        Me.SubBlock.Style("display") = "none"
        Me.ThirdBlock.Style("display") = "inline"
        SecNum.Text = HidSecNum.Value
        SecName.Text = HidSecName.Value

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim i As Integer = 0
        Dim intShowTableRowCount As Integer = 0
        Dim myrow As HtmlTableRow = Nothing
        Dim mycell As HtmlTableCell = Nothing
        sql = " SELECT * FROM Key_ProSkill WHERE Parent = '" & Me.SecID.Value & "' "
        dt = DbAccess.GetDataTable(sql)

        For Each dr In dt.Rows
            '每列顯示2筆資料
            If i Mod 2 = 0 Then
                intShowTableRowCount += 1
                myrow = New HtmlTableRow
                Me.ThirdList.Rows.Add(myrow)
                '設定顯示列的背景顏色
                If intShowTableRowCount Mod 2 = 0 Then
                    myrow.BgColor = "#ddddff"
                Else
                    myrow.BgColor = "#ccccff"
                End If
            End If
            mycell = New HtmlTableCell
            mycell.InnerHtml = String.Format("<a href=""javascript:return_value('{0}','{1}','{2}');"">【{1}】{2}</a>", dr("KPID"), dr("ProID"), dr("ProName"))
            mycell.Width = "50%"
            myrow.Cells.Add(mycell)
            i += 1
        Next
    End Sub
End Class