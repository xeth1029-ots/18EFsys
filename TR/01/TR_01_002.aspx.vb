Partial Class TR_01_002
    Inherits AuthBasePage

    ''Dim FunDr As DataRow
    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me) '☆
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


            Dim RGSTNRound As String
            RGSTNRound = " (Station_Scheme_ID='A' and Station_Unit_ID='31' and Station_ID='000')"
            RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='32' and Station_ID='000')"
            RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='33' and Station_ID='000')"
            RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='34' and Station_ID='000')"
            RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='35' and Station_ID='000')"
            RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='41' and Station_ID='000')"
            RGSTNRound += " or (Station_Scheme_ID='A' and Station_Unit_ID='51' and Station_ID='000')"

            Dim dt As DataTable
            Dim sql As String
            sql = "SELECT Station_Name,Station_Scheme_ID+Station_Unit_ID as Station_ID FROM Adp_WorkStation WHERE " & RGSTNRound
            dt = DbAccess.GetDataTable(sql, objconn)

            With Station
                .DataSource = dt
                .DataTextField = "Station_Name"
                .DataValueField = "Station_ID"
                .DataBind()
                .Items.Insert(0, New ListItem("全部", ""))
                .Items(0).Selected = True
            End With
        End If

        TicketMode.Attributes("onchange") = "TicketChange();"
        Button1.Attributes("onclick") = "return check_data();"
        Page.RegisterStartupScript("TicketChange", "<script>TicketChange();</script>")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        Select Case TicketMode.SelectedIndex
            Case 1, 2, 3
            Case Else
                Common.MessageBox(Me, "請先選擇，券別種類!")
                Exit Sub
        End Select

        Dim SDate As Date
        Dim FDate As Date
        SDate = CDate(SYear.SelectedValue & "/" & SMonth.SelectedValue & "/1")
        FDate = CDate(FYear.SelectedValue & "/" & FMonth.SelectedValue & "/1")

        Dim sql As String = "SELECT * FROM Key_Identity WHERE 1<>1"
        Dim SearchTableName As String = ""
        Dim TicketType As String = ""
        Select Case TicketMode.SelectedIndex
            Case 1
                SearchTableName = "Adp_TRNData"
                If TICKET_TYPE.SelectedIndex <> 0 AndAlso TICKET_TYPE.SelectedValue <> "" Then
                    '1:甲式 2:乙式
                    TicketType = " and TICKET_TYPE='" & TICKET_TYPE.SelectedValue & "'"
                End If

                sql = "SELECT * FROM Key_Identity WHERE IdentityID IN ('02','03','04','05','06','07','09','10','13','14','17','18')"
            Case 2
                SearchTableName = "Adp_DGTRNData"
                TicketType = ""

                sql = "SELECT Share_ID IdentityID,Share_Name Name FROM Adp_ShareSource WHERE Share_Type='301'"
            Case 3
                SearchTableName = "Adp_GOVTRNData"
                TicketType = ""

                sql = "SELECT * FROM Key_Identity WHERE IdentityID='02'"
        End Select
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)

        Dim StationStr As String = ""
        If Station.SelectedIndex <> 0 AndAlso Station.SelectedValue <> "" Then
            StationStr += " and CREATE_RGSTN like '" & Station.SelectedValue & "%'"
        End If

        sql = ""
        sql &= " SELECT *"
        sql &= " FROM ("
        sql &= "    SELECT *"
        sql &= "    FROM " & SearchTableName
        sql &= "    WHERE 1=1"
        sql &= "    and TICKET_STATE='1'" '有效'(1：有效；0：註銷)
        If TicketType <> "" Then sql &= TicketType
        If StationStr <> "" Then sql &= StationStr
        sql &= "    ) a"
        Dim dt1 As DataTable
        dt1 = DbAccess.GetDataTable(sql, objconn)

        Dim MyCell As TableCell
        Dim MyRow As TableRow

        '鍵立表頭
        Dim TempDate As Date
        TempDate = SDate
        MyRow = New TableRow
        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.Text = ""
        MyRow.Cells.Add(MyCell)
        While TempDate <= FDate
            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.Text = TempDate.Year & "年" & TempDate.Month & "月份"
            MyRow.Cells.Add(MyCell)

            TempDate = TempDate.AddMonths(1)
        End While
        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.Text = "合計"
        MyRow.Cells.Add(MyCell)
        MyRow.CssClass = "TR_TR1"
        RecordTable.Rows.Add(MyRow)

        Dim ff3 As String = ""
        Dim indexTR As Integer = 1
        For Each dr As DataRow In dt.Rows
            TempDate = SDate
            MyRow = New TableRow
            Dim total As Integer = 0

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.Text = dr("Name").ToString
            MyRow.Cells.Add(MyCell)

            While TempDate <= FDate
                MyCell = New TableCell
                MyCell.BorderWidth = Unit.Pixel(1)
                Dim MonthTotal As Integer

                Select Case TicketMode.SelectedIndex
                    Case 1
                        'MonthTotal = dt1.Select("IdentityID Like '%" & dr("IdentityID") & "%' and APPLY_DATE>='" & FormatDateTime(TempDate, 2) & "' and APPLY_DATE<'" & FormatDateTime(TempDate.AddMonths(1), 2) & "'").Length
                        ff3 = "OBJECT_TYPE='" & dr("IdentityID") & "' and APPLY_DATE>='" & FormatDateTime(TempDate, 2) & "' and APPLY_DATE<'" & FormatDateTime(TempDate.AddMonths(1), 2) & "'"
                        MonthTotal = dt1.Select(ff3).Length
                    Case 2
                        ff3 = "OBJECT_TYPE='" & dr("IdentityID") & "' and APPLY_DATE>='" & FormatDateTime(TempDate, 2) & "' and APPLY_DATE<'" & FormatDateTime(TempDate.AddMonths(1), 2) & "'"
                        MonthTotal = dt1.Select(ff3).Length
                        'OBJECT_TYPE
                    Case 3
                        ff3 = "APPLY_DATE>='" & FormatDateTime(TempDate, 2) & "' and APPLY_DATE<'" & FormatDateTime(TempDate.AddMonths(1), 2) & "'"
                        MonthTotal = dt1.Select(ff3).Length
                End Select

                MyCell.Text = MonthTotal
                total += MonthTotal
                MyRow.Cells.Add(MyCell)

                TempDate = TempDate.AddMonths(1)
            End While

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.Text = total
            MyRow.Cells.Add(MyCell)

            'i += 1
            If indexTR Mod 2 <> 0 Then
                MyRow.CssClass = "TR_TR2"
            End If
            indexTR += 1
            RecordTable.Rows.Add(MyRow)
        Next

        '建立表尾
        MyRow = New TableRow
        MyCell = New TableCell
        MyCell.BorderWidth = Unit.Pixel(1)
        MyCell.Text = "合計"
        MyRow.Cells.Add(MyCell)

        'Dim total As Integer = 0
        For i As Integer = 1 To RecordTable.Rows(0).Cells.Count - 1
            Dim total As Integer = 0
            For j As Integer = 1 To RecordTable.Rows.Count - 1
                If IsNumeric(RecordTable.Rows(j).Cells(i).Text) Then
                    total += Int(RecordTable.Rows(j).Cells(i).Text)
                End If
            Next

            MyCell = New TableCell
            MyCell.BorderWidth = Unit.Pixel(1)
            MyCell.Text = total
            MyRow.Cells.Add(MyCell)
        Next
        MyRow.CssClass = "TR_TR1"
        RecordTable.Rows.Add(MyRow)
    End Sub
End Class
