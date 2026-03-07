Partial Class TR_04_012
    Inherits AuthBasePage

    Const cst_css_TR_04002_TR As String = "TR_04002_TR"
    Const cst_strWhere1 As String = "TPLANID NOT IN ('06','07','15')"

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

        If Not IsPostBack Then
            STDate2.Text = Now.Date
            TPlanID = TIMS.Get_TPlan(TPlanID, , 1, , cst_strWhere1, objconn)

            Button1.Attributes("onclick") = "return search();"
            SelectAllItem.Attributes("onclick") = "SelectAll(this.checked);"

            If CInt(Me.sm.UserInfo.Years) >= 2010 Then
                Range1.Text = "23"
                Range2.Text = "24"
                Range3.Text = "51"
                Range4.Text = "52"
            Else
                Range1.Text = "26"
                Range2.Text = "27"
                Range3.Text = "52"
                Range4.Text = "53"
            End If
        End If

    End Sub

    '查詢SQL
    Sub search1()
        'Dim SearchStr As String
        'Dim dr As DataRow
        'Dim conn As New SqlConnection
        Dim TPlanIDStr As String = ""
        Dim sql As String = ""
        Dim dt As New DataTable
        'Dim conn As SqlConnection = DbAccess.GetConnection()
        Try
            'Call TIMS.OpenDbConn(conn)
            'conn.Open()
            TPlanIDStr = ""
            For Each item As ListItem In TPlanID.Items
                If item.Selected = True Then
                    If TPlanIDStr <> "" Then TPlanIDStr += ","
                    TPlanIDStr += "'" & item.Value & "'"
                End If
            Next

            '資料庫連線---   Start
            sql = "" & vbCrLf
            sql &= " SELECT a.SOCID" & vbCrLf
            sql &= " ,a.StudStatus" & vbCrLf
            'sql &= " ,a.RTReasonID" & vbCrLf
            sql &= " ,b.FTDate" & vbCrLf
            sql &= " ,dbo.NVL(c.IsGetJob,0) IsGetJob" & vbCrLf
            sql &= " ,c.Mode_" & vbCrLf
            sql &= " ,d.LostJob" & vbCrLf
            sql &= " ,j.DistID" & vbCrLf
            sql &= " ,j.Name DistName " & vbCrLf
            sql &= " ,c.BusGNO" & vbCrLf
            sql &= " ,a.ActNo" & vbCrLf
            sql &= " ,c.PUBLICRESCUE" & vbCrLf
            sql &= " ,a.WkAheadOfSch" & vbCrLf
            sql &= " ,CASE WHEN a.WkAheadOfSch='Y' and a.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='1' THEN 'Y' END WINumS1"
            sql &= " ,CASE WHEN a.WkAheadOfSch='Y' and a.StudStatus in (2,3) AND sg9.mode_=2 and dbo.NVL(sg9.SureItem,'3')='2' THEN 'Y' END WINumS2"
            sql &= " ,CASE WHEN a.WkAheadOfSch='Y' and a.StudStatus in (2,3) AND sg9.mode_=1 and dbo.NVL(sg9.SureItem,'3')='3' THEN 'Y' END WINumS3"
            'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
            '    sql &= ",c.BusGNO,a.ActNo "
            'End If
            sql += " FROM Class_StudentsOfClass a " & vbCrLf
            sql += " JOIN Class_ClassInfo b ON a.OCID=b.OCID " & vbCrLf
            sql += " JOIN Stud_StudentInfo h ON a.SID=h.SID " & vbCrLf
            sql += " JOIN ID_Plan i ON b.PlanID=i.PlanID " & vbCrLf
            sql += " JOIN ID_District j ON i.DistID=j.DistID " & vbCrLf
            sql += " LEFT JOIN STUD_GETJOBSTATE3 c ON c.SOCID =a.SOCID and c.CPoint=1" & vbCrLf
            sql += " LEFT JOIN STUD_GETJOBSTATE3 sg9 ON sg9.socid =a.socid and sg9.CPoint=9" & vbCrLf

            sql += " LEFT JOIN STUD_LOSTJOBWEEK d ON a.SOCID=d.SOCID" & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            'Class_ClassInfo b
            sql += " and b.IsSuccess='Y'" & vbCrLf
            sql += " and b.NotOpen='N'" & vbCrLf
            '3個月前的結訓班級。
            sql += " and b.FTDate< DATEADD(month, -3, dbo.TRUNC_DATETIME(getdate()))" & vbCrLf
            If STDate1.Text <> "" Then
                sql += " and b.STDate>= " & TIMS.To_date(STDate1.Text) & vbCrLf
            End If
            If STDate2.Text <> "" Then
                sql += " and b.STDate<= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'"
            End If
            If FTDate1.Text <> "" Then
                sql += " and b.FTDate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'"
            End If
            If FTDate2.Text <> "" Then
                sql += " and b.FTDate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'"
            End If
            If TPlanIDStr <> "" Then
                sql += " and i.TPlanID IN (" & TPlanIDStr & ")" & vbCrLf
            End If
            '排除 TPlanID: 06 07 15
            sql += " and i.TPlanID NOT IN ('06','07','15')" & vbCrLf

            sql += " Order By i.DistID " & vbCrLf
            Dim da As New SqlDataAdapter
            With da
                .SelectCommand = New SqlCommand(sql, objconn)
                .SelectCommand.CommandTimeout = 100
                .Fill(dt)
            End With
            'dt = DbAccess.GetDataTable(sql)
            If dt.Rows.Count > 0 Then
                Call create_list(dt, TPlanIDStr)
            End If
        Catch ex As Exception
            Me.Page.RegisterStartupScript("Errmsg", "<script>alert('【發生錯誤】:\n" & ex.ToString.Replace("'", "\'").Replace(Convert.ToChar(10), "\n").Replace(Convert.ToChar(13), "") & "');</script>")
            'Finally
            'If Not da Is Nothing Then da.Dispose()
            'If Not dt Is Nothing Then dt.Dispose()
        End Try
        'Call TIMS.CloseDbConn(conn)
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call search1()
    End Sub

    Sub create_list(ByVal dt As DataTable, ByVal TPlanIDStr As String)
        Dim MyCell As TableCell = Nothing
        Dim MyRow As TableRow = Nothing
        Dim LostRange1 As String = "LostJob<=" & Range1.Text & " and LostJob>=0"
        Dim LostRange2 As String = "LostJob>=" & Range2.Text & " and LostJob<=" & Range3.Text & ""
        Dim LostRange3 As String = "LostJob>=" & Range4.Text & ""

        CreateRow(ShowDataTable, MyRow)
        CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 2, , cst_css_TR_04002_TR)
        'CreateCell(MyRow, MyCell, "訓前已加保者", , , cst_css_TR_04002_TR)
        'MyCell.Width = Unit.Pixel(90)
        CreateCell(MyRow, MyCell, "無加退保紀錄", , , cst_css_TR_04002_TR)
        MyCell.Width = Unit.Pixel(90)
        CreateCell(MyRow, MyCell, Range1.Text & "週(含)以下", , , cst_css_TR_04002_TR)
        MyCell.Width = Unit.Pixel(90)
        CreateCell(MyRow, MyCell, Range2.Text & "週至" & Range3.Text & "週", , , cst_css_TR_04002_TR)
        MyCell.Width = Unit.Pixel(90)
        CreateCell(MyRow, MyCell, Range4.Text & "週(含)以上", , , cst_css_TR_04002_TR)
        MyCell.Width = Unit.Pixel(90)
        CreateCell(MyRow, MyCell, "合計", , , cst_css_TR_04002_TR)
        MyCell.Width = Unit.Pixel(90)

        'j:
        Const cst_結訓p As Integer = 1 '結訓人數
        Const cst_提前就業p As Integer = 2 '提前就業人數 '1:雇主切結 2:學員切結 3:勞保勾稽
        Const cst_提前就業p1 As Integer = 3 '提前就業人數-雇主切結
        Const cst_提前就業p2 As Integer = 4 '提前就業人數-學員切結
        Const cst_提前就業p3 As Integer = 5 '提前就業人數-勞保勾稽
        Const cst_未就業p As Integer = 6 '未就業
        Const cst_不就業p As Integer = 7 '不就業
        Const cst_人判就業p As Integer = 8 '人工判定<BR>就業人數
        Const cst_系判就業p As Integer = 9 '系統判定<BR>就業人數
        Const cst_公法救助就業p As Integer = 10 '公法上救助對象之就業人數
        Const cst_就業率1 As Integer = 11 '就業率1
        Const cst_就業率2 As Integer = 12 '就業率2
        Const Cst_max列數 As Integer = 12

        For i As Integer = 0 To 6 '轄區迴圈
            Dim FinPeo As New ArrayList          '結訓人數
            Dim RejPeo As New ArrayList          '提前就業人數
            Dim RejPeo1 As New ArrayList          '提前就業人數1
            Dim RejPeo2 As New ArrayList          '提前就業人數2
            Dim RejPeo3 As New ArrayList          '提前就業人數3

            Dim NotInWork As New ArrayList       '未就業人數
            Dim TWork As New ArrayList           '不就業人數
            Dim InWorkByPeo As New ArrayList     '就業人數(人工)
            Dim InWork As New ArrayList          '就業人數
            Dim BeforeInWork As New ArrayList    '訓前一個月已加保
            Dim PUWork As New ArrayList         '公法上救助對象之就業人數

            For j As Integer = 1 To Cst_max列數
                CreateRow(ShowDataTable, MyRow)
                If j = 1 Then
                    Dim myValue As String = ""
                    Select Case i
                        Case 0
                            myValue = TIMS.Get_DistName2("000") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                        Case 1
                            myValue = TIMS.Get_DistName2("001") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                        Case 2
                            myValue = TIMS.Get_DistName2("002") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                        Case 3
                            myValue = TIMS.Get_DistName2("003") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                        Case 4
                            myValue = TIMS.Get_DistName2("004") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                        Case 5
                            myValue = TIMS.Get_DistName2("005") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                        Case 6
                            myValue = TIMS.Get_DistName2("006") : CreateCell(MyRow, MyCell, myValue, , Cst_max列數, cst_css_TR_04002_TR)
                    End Select
                    MyCell.Width = Unit.Pixel(90)
                End If

                Select Case j
                    Case cst_結訓p '結訓人數'1
                        CreateCell(MyRow, MyCell, "結訓人數")
                        MyCell.ToolTip = "過結訓日班級，且學員非離訓、退訓狀態的人數"
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        'FinPeo.Add(dt.Select("LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        FinPeo.Add(dt.Select("LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        FinPeo.Add(dt.Select(LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        FinPeo.Add(dt.Select(LostRange2 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        FinPeo.Add(dt.Select(LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        FinPeo.Add(dt.Select("StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, FinPeo(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_提前就業p '提前就業人數(人)
                        CreateCell(MyRow, MyCell, "提前就業人數(人)")
                        MyCell.ToolTip = "學員離訓、退訓的原因為提前就業的人數的人數"
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        RejPeo.Add(dt.Select("LostJob IS NULL and StudStatus IN (2,3) and WkAheadOfSch='Y' and DistID='00" & i & "'").Length)
                        RejPeo.Add(dt.Select(LostRange1 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' and DistID='00" & i & "'").Length)
                        RejPeo.Add(dt.Select(LostRange2 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' and DistID='00" & i & "'").Length)
                        RejPeo.Add(dt.Select(LostRange3 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' and DistID='00" & i & "'").Length)
                        RejPeo.Add(dt.Select("StudStatus IN (2,3) and WkAheadOfSch='Y' and DistID='00" & i & "'").Length)

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, RejPeo(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_提前就業p1 '提前就業人數-雇主切結
                        CreateCell(MyRow, MyCell, "提前就業人數<BR>雇主切結(人)")
                        MyCell.ToolTip = "學員離訓、退訓的原因為提前就業的人數(雇主切結)"
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        RejPeo1.Add(dt.Select("LostJob IS NULL and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS1='Y' and DistID='00" & i & "'").Length)
                        RejPeo1.Add(dt.Select(LostRange1 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS1='Y' and DistID='00" & i & "'").Length)
                        RejPeo1.Add(dt.Select(LostRange2 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS1='Y' and DistID='00" & i & "'").Length)
                        RejPeo1.Add(dt.Select(LostRange3 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS1='Y' and DistID='00" & i & "'").Length)
                        RejPeo1.Add(dt.Select("StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS1='Y' and DistID='00" & i & "'").Length)

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, RejPeo1(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_提前就業p2 '提前就業人數-學員切結
                        CreateCell(MyRow, MyCell, "提前就業人數<BR>學員切結(人)")
                        MyCell.ToolTip = "學員離訓、退訓的原因為提前就業的人數(學員切結)"
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        RejPeo2.Add(dt.Select("LostJob IS NULL and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS2='Y' and DistID='00" & i & "'").Length)
                        RejPeo2.Add(dt.Select(LostRange1 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS2='Y' and DistID='00" & i & "'").Length)
                        RejPeo2.Add(dt.Select(LostRange2 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS2='Y' and DistID='00" & i & "'").Length)
                        RejPeo2.Add(dt.Select(LostRange3 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS2='Y' and DistID='00" & i & "'").Length)
                        RejPeo2.Add(dt.Select("StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS2='Y' and DistID='00" & i & "'").Length)

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, RejPeo2(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_提前就業p3 '提前就業人數-勞保勾稽
                        CreateCell(MyRow, MyCell, "提前就業人數<BR>勞保勾稽(人)")
                        MyCell.ToolTip = "學員離訓、退訓的原因為提前就業的人數(勞保勾稽)"
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        RejPeo3.Add(dt.Select("LostJob IS NULL and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS3='Y' and DistID='00" & i & "'").Length)
                        RejPeo3.Add(dt.Select(LostRange1 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS3='Y' and DistID='00" & i & "'").Length)
                        RejPeo3.Add(dt.Select(LostRange2 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS3='Y' and DistID='00" & i & "'").Length)
                        RejPeo3.Add(dt.Select(LostRange3 & " and StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS3='Y' and DistID='00" & i & "'").Length)
                        RejPeo3.Add(dt.Select("StudStatus IN (2,3) and WkAheadOfSch='Y' AND WINumS3='Y' and DistID='00" & i & "'").Length)

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, RejPeo3(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next


                    Case cst_未就業p '未就業
                        CreateCell(MyRow, MyCell, "未就業")
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫
                                MyCell.ToolTip = "學員尚未就業之人數的人數,或者投保單位並非指定的訓練單位之人數"
                            Case Else
                                MyCell.ToolTip = "學員尚未就業之人數的人數"
                        End Select
                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        'IsGetJob@NUMBER(10,0)
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫
                                'NotInWork.Add(dt.Select("(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and " & LostRange2 & " and LostJob<=" & Range3.Text & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                            Case Else
                                'NotInWork.Add(dt.Select("IsGetJob='0' and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("IsGetJob='0' and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("IsGetJob='0' and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("IsGetJob='0' and " & LostRange2 & " and LostJob<=" & Range3.Text & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("IsGetJob='0' and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                NotInWork.Add(dt.Select("IsGetJob='0' and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        End Select

                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, NotInWork(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_不就業p '不就業
                        CreateCell(MyRow, MyCell, "不就業")
                        MyCell.ToolTip = "學員選擇不願就業的人數(可能升學、出國等等原因)"
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        'TWork.Add(dt.Select("IsGetJob='2' and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        TWork.Add(dt.Select("IsGetJob='2' and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        TWork.Add(dt.Select("IsGetJob='2' and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        TWork.Add(dt.Select("IsGetJob='2' and " & LostRange2 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        TWork.Add(dt.Select("IsGetJob='2' and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        TWork.Add(dt.Select("IsGetJob='2' and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, TWork(m))
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_人判就業p '人工判定<BR>就業人數5
                        CreateCell(MyRow, MyCell, "人工判定<BR>就業人數")
                        If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                            MyCell.ToolTip = "人工判定判定學員就業,且投保單位為指定機構之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數+提前就業人數-未就業-不就業"
                        Else
                            MyCell.ToolTip = "人工判定判定學員就業之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數+提前就業人數-未就業-不就業"
                        End If
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        MyCell.Font.Bold = True
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        'Mode_: NUMBER(10,0)
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫

                                'InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange2 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                            Case Else
                                'InWorkByPeo.Add(dt.Select("Mode_=2 and IsGetJob='1' and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and IsGetJob='1' and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and IsGetJob='1' and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and IsGetJob='1' and " & LostRange2 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and IsGetJob='1' and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                InWorkByPeo.Add(dt.Select("Mode_=2 and IsGetJob='1' and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        End Select
                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If

                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, InWorkByPeo(m))
                            MyCell.Font.Bold = True
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_系判就業p '系統判定<BR>就業人數
                        CreateCell(MyRow, MyCell, "系統判定<BR>就業人數")
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫
                                MyCell.ToolTip = "系統自動判定判定學員就業,且投保單位為指定機構之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數+提前就業人數-未就業-不就業"
                            Case Else
                                MyCell.ToolTip = "系統自動判定判定學員就業之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數+提前就業人數-未就業-不就業"
                        End Select

                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        MyCell.Font.Bold = True
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        For m As Integer = 0 To 4
                            Select Case TPlanIDStr
                                Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                    '23	訓用合一
                                    '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                    '41	推動營造業事業單位辦理職前培訓計畫
                                    InWork.Add(FinPeo(m) - NotInWork(m) - TWork(m) - InWorkByPeo(m))
                                Case Else
                                    InWork.Add(FinPeo(m) - NotInWork(m) - TWork(m) - InWorkByPeo(m))
                            End Select

                            'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                            'Else
                            'End If

                            CreateCell(MyRow, MyCell, InWork(m))
                            MyCell.Font.Bold = True
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next

                    Case cst_公法救助就業p '公法上救助對象之就業人數
                        CreateCell(MyRow, MyCell, "公法救助對象<BR>就業人數")
                        If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                            MyCell.ToolTip = "公法上救助對象之就業人數,且投保單位為指定機構之人數" & vbCrLf
                        Else
                            MyCell.ToolTip = "公法上救助對象之就業人數" & vbCrLf
                        End If
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Width = Unit.Pixel(90)
                        MyCell.Font.Bold = True
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        'Mode_: NUMBER(10,0)
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫
                                'InWorkByPeo.Add(dt.Select("Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and " & LostRange2 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                            Case Else
                                'PUWork.Add(dt.Select("Mode_=2 and IsGetJob='1' and LostJob=-1 and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and LostJob IS NULL and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and " & LostRange1 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and " & LostRange2 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and " & LostRange3 & " and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                                PUWork.Add(dt.Select("PUBLICRESCUE='Y' and IsGetJob='1' and StudStatus Not IN (2,3) and DistID='00" & i & "'").Length)
                        End Select
                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If
                        For m As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, PUWork(m))
                            MyCell.Font.Bold = True
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next


                    Case cst_就業率1 '就業率1'7
                        CreateCell(MyRow, MyCell, "就業率1")
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫
                                MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數-不就業+提前就業人數)"
                            Case Else
                                MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數-不就業+提前就業人數)"
                        End Select
                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If
                        MyCell.Width = Unit.Pixel(90)
                        MyCell.Font.Bold = True
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        For m As Integer = 0 To 4
                            If FinPeo(m) + RejPeo(m) = 0 Then
                                CreateCell(MyRow, MyCell, "0%")
                            ElseIf FinPeo(m) - TWork(m) = 0 Then
                                CreateCell(MyRow, MyCell, "0%")
                            Else
                                Select Case TPlanIDStr
                                    Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                        '23	訓用合一
                                        '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                        '41	推動營造業事業單位辦理職前培訓計畫
                                        CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) - TWork(m) + RejPeo(m)) * 100), 2) & "%")
                                    Case Else
                                        'CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m)) / (FinPeo(m) + RejPeo(m) - TWork(m)) * 100), 2) & "%")
                                        CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) - TWork(m) + RejPeo(m)) * 100), 2) & "%")
                                End Select

                                'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                                'Else
                                'End If
                            End If

                            MyCell.Font.Bold = True
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next
                    Case cst_就業率2 '就業率2
                        CreateCell(MyRow, MyCell, "就業率2")
                        MyCell.Width = Unit.Pixel(90)
                        MyCell.Width = Unit.Pixel(90)
                        Select Case TPlanIDStr
                            Case "'23'", "'34'", "'41'" '單1計畫選擇。
                                '23	訓用合一
                                '34	推動事業單位辦理職前培訓計畫(原與企業合作辦理職前訓練)
                                '41	推動營造業事業單位辦理職前培訓計畫
                                MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數)"
                            Case Else
                                MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數)"
                        End Select

                        'If TPlanIDStr = "'23'" Or TPlanIDStr = "'34'" Or TPlanIDStr = "'41'" Then
                        'Else
                        'End If
                        MyCell.Style("CURSOR") = "help"
                        MyCell.Font.Bold = True
                        If i Mod 2 = 1 Then
                            MyCell.CssClass = "TR_04002_TD2"
                        End If

                        For m As Integer = 0 To 4
                            If FinPeo(m) + RejPeo(m) = 0 Then
                                CreateCell(MyRow, MyCell, "0%")
                            Else
                                ' If TPlanIDStr = "'23'" Or "'34'" Then
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) + RejPeo(m))) * 100, 2) & "%")
                                'Else
                                ' CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m)) / (FinPeo(m) + RejPeo(m))) * 100, 2) & "%")
                                'End If
                            End If
                            MyCell.Font.Bold = True
                            If i Mod 2 = 1 Then
                                MyCell.CssClass = "TR_04002_TD2"
                            End If
                        Next
                End Select
            Next
        Next
        Call CreateRow(ShowDataTable, MyRow)
        Call CreateCell(MyRow, MyCell, "備註：提前就業人數：學員實際參訓時數達總訓練時數1/2以上，經分署專案核定免負擔退訓賠償費用者。", 7, , "TR_04002_TD2")

    End Sub

    Sub CreateCell(ByRef MyRow As TableRow, ByRef MyCell As TableCell, Optional ByVal MyText As String = "", Optional ByVal ColumnSpan As Integer = 1, Optional ByVal RowSpan As Integer = 1, Optional ByVal CssClass As String = "TR_04002_TD")
        MyCell = New TableCell
        MyCell.Text = MyText
        MyCell.ColumnSpan = ColumnSpan
        MyCell.RowSpan = RowSpan
        MyCell.CssClass = CssClass
        MyRow.Cells.Add(MyCell)
    End Sub

    Sub CreateRow(ByRef MyTable As Table, ByRef MyRow As TableRow)
        MyRow = New TableRow
        MyRow.CssClass = cst_css_TR_04002_TR
        MyTable.Rows.Add(MyRow)
    End Sub
End Class
