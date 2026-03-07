Partial Class CM_03_002
    Inherits AuthBasePage

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        If Not IsPostBack Then
            CreateItem()
            ShowDataTable.Visible = False
            Common.SetListItem(Syear, sm.UserInfo.Years)
        End If

        Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        Syear = TIMS.GetSyear(Syear)
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        TPlanID.Items(0).Selected = True
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim MyRow As TableRow
        Dim MyCell As TableCell
        Dim SearchStr As String = ""

        If Syear.SelectedIndex <> 0 Then
            SearchStr += " and Years='" & Right(Syear.SelectedValue, 2) & "'"
        End If
        If TPlanID.SelectedIndex <> 0 Then
            SearchStr += " and PlanID IN (SELECT PlanID From ID_Plan WHERE TPlanID='" & TPlanID.SelectedValue & "')"
        End If

        sql = "SELECT a.Name as DistName"
        For i As Integer = 1 To 12
            sql += ",dbo.NVL(bb" & i & ".StudCount" & i & ",0) as StudCount" & i & ",dbo.NVL(bb" & i & ".CancelCost" & i & ",0) as CancelCost" & i & " "
        Next
        sql += " FROM ID_District a "

        For i As Integer = 1 To 12
            sql += "LEFT JOIN (SELECT c" & i & ".DistID,Sum(d" & i & ".StudCount" & i & ") as StudCount" & i & ",Sum(e" & i & ".CancelCost" & i & ") as CancelCost" & i & " FROM "
            sql += "(SELECT * FROM Class_ClassInfo WHERE CONVERT(numeric, DATEPART(MONTH, FTDate))=" & i & SearchStr & ") b" & i & " "
            sql += "JOIN ID_Plan c" & i & " ON b" & i & ".PlanID=c" & i & ".PlanID "
            sql += "LEFT JOIN (SELECT OCID,Count(*) as StudCount" & i & " FROM Class_StudentsOfClass WHERE StudStatus='5' Group By OCID) d" & i & " ON b" & i & ".OCID=d" & i & ".OCID "
            sql += "LEFT JOIN (SELECT OCID,Sum(CancelCost) as CancelCost" & i & " FROM Budget_ClassCancel Group By OCID) e" & i & " ON b" & i & ".OCID=e" & i & ".OCID "
            sql += " Group By c" & i & ".DistID) bb" & i & " ON a.DistID=bb" & i & ".DistID "
        Next
        sql += "Order By a.DistID"

        dt = DbAccess.GetDataTable(sql)

        ShowDataTable.Rows.Clear()

        '建立表頭-----------   Start
        MyRow = New TableRow
        MyRow.CssClass = "CM_TR1"

        MyCell = New TableCell
        MyCell.Text = "月份別"
        MyCell.RowSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "一月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "二月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "三月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "四月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "五月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "六月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "總計"
        MyCell.RowSpan = 2
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        ShowDataTable.Rows.Add(MyRow)

        MyRow = New TableRow
        MyRow.CssClass = "CM_TR1"

        MyCell = New TableCell
        MyCell.Text = "七月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "八月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "九月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "十月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "十一月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "十二月份"
        MyCell.ColumnSpan = 2
        MyRow.Cells.Add(MyCell)

        ShowDataTable.Rows.Add(MyRow)

        MyRow = New TableRow
        MyRow.CssClass = "CM_TR1"

        MyCell = New TableCell
        MyCell.Text = "轄區中心"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.Text = "人數"
        MyRow.Cells.Add(MyCell)
        MyCell = New TableCell
        MyCell.Text = "金額"
        MyRow.Cells.Add(MyCell)

        ShowDataTable.Rows.Add(MyRow)
        '建立表頭-----------   End

        '建立資料-----------   Start
        For Each dr In dt.Rows
            MyRow = New TableRow
            MyRow.CssClass = "CM_TR2"

            MyCell = New TableCell
            MyCell.Text = dr("DistName")
            MyCell.RowSpan = 2
            MyRow.Cells.Add(MyCell)

            'ShowDataTable.Rows.Add(MyRow)

            For i As Integer = 1 To 12
                If i = 7 Then
                    MyRow = New TableRow
                    MyRow.CssClass = "CM_TR2"
                End If
                MyCell = New TableCell
                MyCell.Text = dr("StudCount" & i)
                MyRow.Cells.Add(MyCell)

                MyCell = New TableCell
                MyCell.Text = dr("CancelCost" & i)
                MyRow.Cells.Add(MyCell)

                If i = 6 Or i = 12 Then
                    If i = 6 Then
                        MyCell = New TableCell
                        MyCell.RowSpan = 2
                        MyCell.Text = 0
                        For j As Integer = 1 To 12
                            MyCell.Text += dr("StudCount" & j)
                        Next
                        MyRow.Cells.Add(MyCell)

                        MyCell = New TableCell
                        MyCell.RowSpan = 2
                        MyCell.Text = 0
                        For j As Integer = 1 To 12
                            MyCell.Text += dr("CancelCost" & j)
                        Next
                        MyRow.Cells.Add(MyCell)
                    End If

                    ShowDataTable.Rows.Add(MyRow)
                End If
            Next
        Next
        '建立資料-----------   End


        ShowDataTable.Visible = True
    End Sub
End Class
