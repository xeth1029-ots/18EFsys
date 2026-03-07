Partial Class TC_03_oper
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Button1.Attributes("onclick") = "javascript:return chkDataFormat();"
        End If
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        Select Case RadioButtonList1.SelectedValue
            Case "1", "2"
            Case Else
                Errmsg += "請先選擇要換算的總類" & vbCrLf
        End Select

        oneday.Text = TIMS.ClearSQM(oneday.Text)
        If oneday.Text <> "" Then
            Try
                If IsNumeric(oneday.Text) Then
                    oneday.Text = CInt(oneday.Text)
                Else
                    Errmsg += "一天時數必須為正整數" & vbCrLf
                End If
            Catch ex As Exception
                Errmsg += "一天時數必須為正整數" & vbCrLf
            End Try
        Else
            Errmsg += "請輸入一天時數" & vbCrLf
        End If

        If Errmsg = "" Then
            start_date.Text = TIMS.ClearSQM(start_date.Text)
            If start_date.Text <> "" Then
                'start_date.Text = Trim(start_date.Text)
                If Not TIMS.IsDate1(start_date.Text) Then Errmsg += "開訓日 日期格式有誤" & vbCrLf
                If Errmsg = "" Then start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            Else
                start_date.Text = ""
                Errmsg += "開訓日 日期為必填" & vbCrLf
            End If

            end_date.Text = TIMS.ClearSQM(end_date.Text)
            If end_date.Text <> "" Then
                'end_date.Text = Trim(end_date.Text)
                If Not TIMS.IsDate1(end_date.Text) Then Errmsg += "結訓日 日期格式有誤" & vbCrLf
                If Errmsg = "" Then end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            End If

            hours.Text = TIMS.ClearSQM(hours.Text)
            If hours.Text <> "" Then
                'hours.Text = hours.Text.Trim
                Try
                    If IsNumeric(hours.Text) Then
                        hours.Text = CInt(hours.Text)
                    Else
                        Errmsg += "時數必須為正整數" & vbCrLf
                    End If
                Catch ex As Exception
                    Errmsg += "時數必須為正整數" & vbCrLf
                End Try
            End If

            If Errmsg = "" Then
                Select Case RadioButtonList1.SelectedValue
                    Case "1"
                        If hours.Text = "" Then
                            'hours.Text = ""
                            Errmsg += "「換算種類 依結訓日」 時數 為必填" & vbCrLf
                        End If
                    Case "2"
                        If end_date.Text = "" Then
                            'end_date.Text = ""
                            Errmsg += "「換算種類 依時數」 結訓日 日期為必填" & vbCrLf
                        End If
                End Select
            End If
        End If

        If Errmsg = "" Then
            If start_date.Text.ToString <> "" AndAlso end_date.Text.ToString <> "" Then
                If CDate(start_date.Text) > CDate(end_date.Text) Then Errmsg += "【開訓日】的不得大於【結訓日】!!" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '換算
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim sql As String = ""
        Dim dt As DataTable
        Dim dr As DataRow
        Dim new_date As DateTime
        Dim AllDays As Integer = 0          '增加的天數

        Select Case RadioButtonList1.SelectedValue
            Case "1" '計算 結訓日 日期
                sql = " SELECT * FROM Sys_Holiday WHERE HolDate >= " & TIMS.To_date(start_date.Text) & " AND RID = '" & sm.UserInfo.RID & "' ORDER BY HolDate "
                dt = DbAccess.GetDataTable(sql, objconn)

                AllDays = Int(hours.Text) \ Int(oneday.Text)
                If CInt(hours.Text) Mod oneday.Text = 0 Then AllDays -= 1 '因為開訓日也算時數,所以-1天

                new_date = CDate(start_date.Text)
                For i As Integer = 1 To AllDays
                    new_date = DateAdd(DateInterval.Day, 1, new_date)
                    If new_date.DayOfWeek = DayOfWeek.Sunday Or new_date.DayOfWeek = DayOfWeek.Saturday Then '如果是星期6就+2天變星期一,不會遇到星期日
                        new_date = DateAdd(DateInterval.Day, 2, new_date)
                        'Else
                    End If
                    For Each dr In dt.Rows
                        If CDate(dr("HolDate")) = new_date Then new_date = DateAdd(DateInterval.Day, 1, new_date) '如果在參數管理->行事曆功能有排日期時,要排除
                    Next
                Next
                end_date.Text = TIMS.Cdate3(new_date)
            Case "2" '計算時數
                Dim SDate As Date = CDate(start_date.Text)
                Dim EDate As Date = CDate(end_date.Text)

                sql = " SELECT * FROM Sys_Holiday WHERE HolDate >= " & TIMS.To_date(start_date.Text) & " AND HolDate <= " & TIMS.To_date(end_date.Text) & " AND RID = '" & sm.UserInfo.RID & "' ORDER BY HolDate "
                dt = DbAccess.GetDataTable(sql, objconn)

                While SDate <= EDate
                    If SDate.DayOfWeek <> DayOfWeek.Saturday And SDate.DayOfWeek <> DayOfWeek.Sunday Then
                        AllDays += 1
                        For Each dr In dt.Rows
                            If dr("HolDate") = SDate Then AllDays -= 1
                        Next
                    End If
                    SDate = DateAdd(DateInterval.Day, 1, SDate)
                End While

                'Dim AllDays = DateDiff(DateInterval.Day, CDate(start_date.Text), CDate(end_date.Text))
                hours.Text = AllDays * Int(oneday.Text)
        End Select
    End Sub
End Class