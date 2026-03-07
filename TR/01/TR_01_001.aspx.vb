Partial Class TR_01_001
    Inherits AuthBasePage

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
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            SYear = TIMS.GetSyear(SYear)
            FYear = TIMS.GetSyear(FYear)
            For i As Integer = 1 To 12
                SMonth.Items.Add(i)
                FMonth.Items.Add(i)
            Next

            Common.SetListItem(SYear, Now.Year)
            Common.SetListItem(FYear, Now.Year)
            'Common.SetListItem(SMonth, Now.Year)
            Common.SetListItem(FMonth, Now.Month)

            If Not SYear.Items.FindByValue("2003") Is Nothing Then
                SYear.Items.Remove(SYear.Items.FindByValue("2003"))
            End If
            If Not SYear.Items.FindByValue("2004") Is Nothing Then
                SYear.Items.Remove(SYear.Items.FindByValue("2004"))
            End If
            If Not FYear.Items.FindByValue("2003") Is Nothing Then
                FYear.Items.Remove(FYear.Items.FindByValue("2003"))
            End If
            If Not FYear.Items.FindByValue("2004") Is Nothing Then
                FYear.Items.Remove(FYear.Items.FindByValue("2004"))
            End If
        End If

        TicketMode.Attributes("onchange") = "TicketChange();"
        Button1.Attributes("onclick") = "return check_data();"
        Page.RegisterStartupScript("TicketChange", "<script>TicketChange();</script>")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String
        Dim dt As DataTable
        Dim dr As DataRow
        Dim RGSTN As String
        Dim RGSTNRound As String
        Dim SDate As Date
        Dim EDate As Date
        Dim TempDate As Date
        'Dim TempEdate As Date
        Dim i As Integer
        Dim SearchTableName As String = ""
        Dim TicketType As String = ""
        'Dim Status As String
        Dim indexTR As Integer = 1
        'RGSTNRound = "(Station_Scheme_ID='A' and Station_Unit_ID='31' and Station_ID='000')"
        'RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='32' and Station_ID='000')"
        'RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='33' and Station_ID='000')"
        'RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='34' and Station_ID='000')"
        'RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='35' and Station_ID='000')"
        'RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='41' and Station_ID='000')"
        'RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='51' and Station_ID='000')"

        RGSTNRound = "(Station_Scheme_ID='A' and Station_ID='000' and Station_Unit_ID IN ('31','32','33','34','35','41','51'))"

        SDate = CDate(SYear.SelectedValue & "/" & SMonth.SelectedValue & "/1")
        EDate = CDate(FYear.SelectedValue & "/" & FMonth.SelectedValue & "/1")

        RGSTN = " and (CREATE_RGSTN Like 'A31%' or CREATE_RGSTN Like 'A32%' or CREATE_RGSTN Like 'A33%' or CREATE_RGSTN Like 'A34%' or CREATE_RGSTN Like 'A35%' or CREATE_RGSTN Like 'A41%' or CREATE_RGSTN Like 'A51%')"

        Select Case TicketMode.SelectedIndex
            Case 1
                SearchTableName = "Adp_TRNData"
                If TICKET_TYPE.SelectedIndex <> 0 Then
                    TicketType = " and TICKET_TYPE='" & TICKET_TYPE.SelectedValue & "'"
                End If
            Case 2
                SearchTableName = "Adp_DGTRNData"
                TicketType = ""
            Case 3
                SearchTableName = "Adp_GOVTRNData"
                TicketType = ""
        End Select

        i = 1
        TempDate = SDate

        sql = "SELECT a.Station_Name "
        While TempDate <= EDate
            sql += ",dbo.NVL(AllTicket" & i & ",0) as AllTicket" & i & " "
            sql += ",dbo.NVL(Enter" & i & ",0) as Enter" & i & " "
            sql += ",dbo.NVL(InJoin" & i & ",0) as InJoin" & i & " "
            i += 1
            TempDate = TempDate.AddMonths(1)
        End While
        sql += " FROM "
        sql += " (SELECT Station_Name as Station_Name,Station_Scheme_ID+Station_Unit_ID  RGSTN FROM Adp_WorkStation WHERE " & RGSTNRound & ") a "
        i = 1
        TempDate = SDate
        While TempDate <= EDate
            sql += " LEFT JOIN (SELECT dbo.SUBSTR2(CREATE_RGSTN, 3) AS CREATE_RGSTN,count(*) as AllTicket" & i & ", SUM(CASE WHEN TransToTIMS = 'Y' THEN 1 ELSE 0 END) as Enter" & i & ", SUM(CASE WHEN dbo.NVL(SOCID,0) > 0 THEN 1 ELSE 0 END) as InJoin" & i
            sql += " FROM " & SearchTableName & " WHERE TICKET_STATE='1'" & RGSTN & " and APPLY_DATE >= " & TIMS.to_date(FormatDateTime(TempDate, 2)) & " and APPLY_DATE < " & TIMS.to_date(FormatDateTime(DateAdd(DateInterval.Month, 1, TempDate), 2)) & " " & TicketType & " Group By dbo.SUBSTR2(CREATE_RGSTN, 3)) b" & i & " ON a.RGSTN=b" & i & ".CREATE_RGSTN "

            'sql += "LEFT JOIN (SELECT LEFT(CREATE_RGSTN,3) AS CREATE_RGSTN,count(*) as AllTicket" & i & " FROM " & SearchTableName & " WHERE TICKET_STATE='1'" & RGSTN & " and APPLY_DATE>='" & FormatDateTime(TempDate, 2).ToString & "' and APPLY_DATE<'" & FormatDateTime(DateAdd(DateInterval.Month, 1, TempDate), 2).ToString & "'" & TicketType & " Group By LEFT(CREATE_RGSTN,3)) b" & i & " ON a.RGSTN=b" & i & ".CREATE_RGSTN "
            'sql += "LEFT JOIN (SELECT LEFT(CREATE_RGSTN,3) AS CREATE_RGSTN,count(*) as Enter" & i & " FROM " & SearchTableName & " WHERE TICKET_STATE='1'" & RGSTN & " and APPLY_DATE>='" & FormatDateTime(TempDate, 2).ToString & "' and APPLY_DATE<'" & FormatDateTime(DateAdd(DateInterval.Month, 1, TempDate), 2).ToString & "' and TransToTIMS='Y'" & TicketType & " Group By LEFT(CREATE_RGSTN,3)) c" & i & " ON a.RGSTN=c" & i & ".CREATE_RGSTN "
            'sql += "LEFT JOIN (SELECT LEFT(CREATE_RGSTN,3) AS CREATE_RGSTN,count(*) as InJoin" & i & " FROM " & SearchTableName & " WHERE TICKET_STATE='1'" & RGSTN & " and APPLY_DATE>='" & FormatDateTime(TempDate, 2).ToString & "' and APPLY_DATE<'" & FormatDateTime(DateAdd(DateInterval.Month, 1, TempDate), 2).ToString & "' and SOCID IS NOT NULL" & TicketType & " Group By LEFT(CREATE_RGSTN,3)) d" & i & " ON a.RGSTN=d" & i & ".CREATE_RGSTN "
            i += 1
            TempDate = TempDate.AddMonths(1)
        End While

        'MsgBox(sql)

        'sql = "SELECT b.公立就業服務機構名稱,"
        'While TempDate <= EDate
        '    sql += "ISNULL(a" & i & ".[" & TempDate.Year & "年" & TempDate.Month & "月份],0) as [" & TempDate.Year & "年" & TempDate.Month & "月份],"

        '    TempDate = TempDate.AddMonths(1)
        '    i += 1
        'End While
        'i = 1
        'TempDate = SDate
        'sql += "ISNULL(a.合計,0) as 合計 FROM "
        'sql += "(SELECT Station_Name as '公立就業服務機構名稱',Station_Scheme_ID+Station_Unit_ID as RGSTN FROM Adp_WorkStation WHERE " & RGSTNRound & ") b "
        'sql += "LEFT JOIN (SELECT LEFT(CREATE_RGSTN,3) AS CREATE_RGSTN,count(*) as 合計 FROM " & SearchTableName & " WHERE APPLY_DATE>='2005/5/31'" & RGSTN & " and APPLY_DATE>='" & FormatDateTime(SDate, 2).ToString & "' and APPLY_DATE<'" & FormatDateTime(DateAdd(DateInterval.Month, 1, EDate), 2).ToString & "'" & TicketType & Status & " Group By LEFT(CREATE_RGSTN,3)) a ON a.CREATE_RGSTN=b.RGSTN "
        'While TempDate <= EDate
        '    sql += "LEFT JOIN (SELECT LEFT(CREATE_RGSTN,3) AS CREATE_RGSTN,count(*) as '" & TempDate.Year & "年" & TempDate.Month & "月份' FROM " & SearchTableName & " WHERE APPLY_DATE>='2005/5/31'" & RGSTN & " and APPLY_DATE>='" & FormatDateTime(TempDate, 2).ToString & "' and APPLY_DATE<'" & FormatDateTime(DateAdd(DateInterval.Month, 1, TempDate), 2).ToString & "'" & TicketType & Status & " Group By LEFT(CREATE_RGSTN,3)) a" & i & " ON b.RGSTN=a" & i & ".CREATE_RGSTN "

        '    i += 1
        '    TempDate = TempDate.AddMonths(1)
        'End While
        dt = DbAccess.GetDataTable(sql, objconn)

        Dim MyCell As TableCell
        Dim MyRow As TableRow

        '建立表頭
        MyRow = New TableRow
        MyRow.CssClass = "TR_TR1"

        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.RowSpan = 2
        MyCell.Text = "公立就業服務機構名稱"
        MyRow.Cells.Add(MyCell)

        TempDate = SDate
        While TempDate <= EDate
            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.ColumnSpan = 3
            MyCell.Text = TempDate.Year & "年" & TempDate.Month & "月"
            MyRow.Cells.Add(MyCell)

            TempDate = TempDate.AddMonths(1)
        End While

        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyCell.ColumnSpan = 3
        MyCell.Text = "合計"
        MyRow.Cells.Add(MyCell)

        RecordTable.Rows.Add(MyRow)

        MyRow = New TableRow
        MyRow.CssClass = "TR_TR1"
        i = 1
        TempDate = SDate
        While TempDate <= EDate.AddMonths(1)
            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.Text = "開券(單)人數"
            MyRow.Cells.Add(MyCell)

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.Text = "報名人數"
            MyRow.Cells.Add(MyCell)

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.Text = "開訓人數"
            MyRow.Cells.Add(MyCell)

            TempDate = TempDate.AddMonths(1)
        End While

        RecordTable.Rows.Add(MyRow)

        '建立資料
        Dim AllTicket As Integer
        Dim Enter As Integer
        Dim InJoin As Integer
        For Each dr In dt.Rows
            AllTicket = 0
            Enter = 0
            InJoin = 0
            MyRow = New TableRow
            MyRow.CssClass = "TR_TR2"

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.Text = dr("Station_Name")
            MyRow.Cells.Add(MyCell)

            i = 1
            TempDate = SDate
            While TempDate <= EDate
                MyCell = New TableCell
                MyCell.BorderWidth = Unit.Pixel(1)
                MyCell.HorizontalAlign = HorizontalAlign.Center
                MyCell.Text = dr("AllTicket" & i)
                MyRow.Cells.Add(MyCell)

                MyCell = New TableCell
                MyCell.BorderWidth = Unit.Pixel(1)
                MyCell.HorizontalAlign = HorizontalAlign.Center
                MyCell.Text = dr("Enter" & i)
                MyRow.Cells.Add(MyCell)

                MyCell = New TableCell
                MyCell.BorderWidth = Unit.Pixel(1)
                MyCell.HorizontalAlign = HorizontalAlign.Center
                MyCell.Text = dr("InJoin" & i)
                MyRow.Cells.Add(MyCell)

                AllTicket += dr("AllTicket" & i)
                Enter += dr("Enter" & i)
                InJoin += dr("InJoin" & i)

                i += 1
                TempDate = TempDate.AddMonths(1)
            End While

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.Text = AllTicket
            MyRow.Cells.Add(MyCell)

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.Text = Enter
            MyRow.Cells.Add(MyCell)

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            MyCell.Text = InJoin
            MyRow.Cells.Add(MyCell)

            RecordTable.Rows.Add(MyRow)
        Next


        '表尾
        Dim total As Integer
        MyRow = New TableRow
        MyRow.CssClass = "TR_TR1"

        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.Text = "合計"
        MyRow.Cells.Add(MyCell)
        For i = 1 To dt.Columns.Count - 1
            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.HorizontalAlign = HorizontalAlign.Center
            total = 0

            For Each dr In dt.Rows
                total += dr(i)
            Next
            MyCell.Text = total

            MyRow.Cells.Add(MyCell)
        Next

        AllTicket = 0
        Enter = 0
        InJoin = 0
        For i = 2 To RecordTable.Rows.Count - 1
            Dim TRow As TableRow
            TRow = RecordTable.Rows(i)
            AllTicket += TRow.Cells(TRow.Cells.Count - 3).Text
            Enter += TRow.Cells(TRow.Cells.Count - 2).Text
            InJoin += TRow.Cells(TRow.Cells.Count - 1).Text
        Next

        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyCell.Text = AllTicket
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyCell.Text = Enter
        MyRow.Cells.Add(MyCell)

        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.HorizontalAlign = HorizontalAlign.Center
        MyCell.Text = InJoin
        MyRow.Cells.Add(MyCell)

        RecordTable.Rows.Add(MyRow)
        Exit Sub

        For i = 0 To dt.Columns.Count - 1
            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)

            MyCell.Text = dt.Columns(i).ColumnName
            MyRow.Cells.Add(MyCell)
        Next

        RecordTable.Rows.Add(MyRow)

        For Each dr In dt.Rows
            MyRow = New TableRow

            For i = 0 To dt.Columns.Count - 1
                MyCell = New TableCell
                MyCell.BorderWidth = Unit.Pixel(1)
                MyCell.Text = dr(i).ToString
                MyRow.Cells.Add(MyCell)
            Next

            If indexTR Mod 2 <> 0 Then
                MyRow.CssClass = "TR_TR2"
            End If
            indexTR += 1
            RecordTable.Rows.Add(MyRow)
        Next

        MyRow = New TableRow
        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.Text = "合計"
        MyRow.Cells.Add(MyCell)
        'Dim total As Integer
        For i = 1 To dt.Columns.Count - 1
            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            total = 0

            For Each dr In dt.Rows
                total += dr(i)
            Next
            MyCell.Text = total

            If indexTR Mod 2 <> 0 Then
                MyRow.CssClass = "TR_TR2"
            End If
            indexTR += 1
            MyRow.Cells.Add(MyCell)
        Next
        MyRow.CssClass = "TR_TR1"
        RecordTable.Rows.Add(MyRow)
    End Sub

End Class
