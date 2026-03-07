Partial Class SD_05_002_R_Rpt
    Inherits AuthBasePage

    'stud_turnout / KEY_LEAVE
    Dim dtLeave As DataTable = Nothing
    Dim s_ORGNAME As String = ""
    Dim sql As String = ""

    'Dim flagYear2017 As Boolean = False
    'Const cst_leaveid_set_1 As String = "'01','02','03','04','05','06','07','08'"
    Const cst_print_style_x1 As String = "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none"

    Dim objconn As SqlConnection

    'Function Get_dtLeave() As DataTable
    '    Dim rst As DataTable = Nothing
    '    sql = "" & vbCrLf
    '    sql &= " select leaveid" & vbCrLf
    '    sql &= " ,Name " & vbCrLf
    '    sql &= " ,0 intSum" & vbCrLf
    '    sql &= " FROM dbo.KEY_LEAVE WITH(NOLOCK)" & vbCrLf
    '    sql &= " where 1=1" & vbCrLf
    '    sql &= " and leaveid in (" & cst_leaveid_set_1 & ")" & vbCrLf
    '    sql &= " order by leaveid " & vbCrLf
    '    If flagYear2017 Then
    '        sql = "" & vbCrLf
    '        sql &= " select leaveid" & vbCrLf
    '        sql &= " ,Name " & vbCrLf
    '        sql &= " ,0 intSum" & vbCrLf
    '        sql &= " FROM dbo.KEY_LEAVE WITH(NOLOCK)" & vbCrLf
    '        sql &= " where 1=1" & vbCrLf
    '        sql &= " and NOUSE IS NULL" & vbCrLf
    '        sql &= " order by LEAVESORT " & vbCrLf
    '    End If
    '    rst = DbAccess.GetDataTable(sql, objconn)
    '    Return rst
    'End Function

    '建置表格
    Sub crtTable()
        dtLeave = TIMS.Get_dtLEAVE(objconn) 'DbAccess.GetDataTable(sql, objconn)

        hidprtPageSize.Value = TIMS.ClearSQM(hidprtPageSize.Value)
        Dim int_ColSpn8 As Integer = dtLeave.Rows.Count
        Dim dt As DataTable = cntData()
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Common.MessageBox2(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim dr As DataRow = Nothing
        'Dim intRowCnt As Integer = 54 '每頁筆數
        Dim intRowCnt As Integer = 26 '每頁筆數
        Dim intPageCnt As Integer = 0  '頁數控制
        Dim intCnt As Integer = 0 '報表處理筆數
        If hidprtPageSize.Value <> "" AndAlso TIMS.IsNumeric2(hidprtPageSize.Value) Then intRowCnt = Val(hidprtPageSize.Value) '每頁筆數

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl

        Dim strTitleStyle As String = "font-size:14pt;font-family:DFKai-SB"
        Dim strCellStyle As String = "font-size:10pt;font-family:DFKai-SB"

        intPageCnt = dt.Rows.Count / intRowCnt
        If intPageCnt * intRowCnt < dt.Rows.Count Then intPageCnt += 1

        For i As Integer = 1 To intPageCnt
            '加背景圖的div
            nl = New HtmlGenericControl
            div_print_content.Controls.Add(nl)
            'If i = 1 Then
            '    'nl.InnerHtml = "<div style='position:absolute;z-index:-1; margin:%;padding:0;left:0px;top:0px;'>" & vbCrLf
            '    nl.InnerHtml = "<div style='position:absolute;z-index:-1; margin:%;padding:20px;left:0px;top:0px;'>" & vbCrLf
            'Else
            '    nl.InnerHtml = "<div style='position:absolute;z-index:-1; margin:%;padding:20px;'>" & vbCrLf
            'End If

            'nl.InnerHtml += "<img src='../../images/rptpic/temple/TIMS_1.jpg' height='942' />" & vbCrLf
            ''nl.InnerHtml += "<img src='../../images/rptpic/temple/TIMS_2.jpg' height='1030' />" & vbCrLf
            'nl.InnerHtml += "</div>" & vbCrLf

            '表首
            nt = New HtmlTable
            nt.Attributes.Add("style", cst_print_style_x1)
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", "0")
            div_print_content.Controls.Add(nt)

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strTitleStyle)
            nc.InnerHtml = s_ORGNAME

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strTitleStyle)
            Dim sTPlanName As String = TIMS.GetTPlanName(hidTPlanID.Value, objconn)
            'nc.InnerHtml = getName("plan") & "　" & "出缺勤明細表"
            nc.InnerHtml = sTPlanName & "　" & "出缺勤明細表" & "　"

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "查詢日期：" & hidSDate.Value & "~" & hidEDate.Value & "　"

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("width", "50%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "　列印者：" & TIMS.Get_ACCNAME(hidUserID.Value, objconn) 'hidUserID.Value  getName("user")

            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("width", "50%")
            nc.Attributes.Add("align", "right")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "列印日期：" & Now.ToString("yyyy/MM/dd") & "　"

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "　　頁數：" & i.ToString & " / " + intPageCnt.ToString

            '表頭
            nt = New HtmlTable
            nt.Attributes.Add("style", cst_print_style_x1)
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", 1)
            div_print_content.Controls.Add(nt)

            '表頭1
            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("rowspan", "2")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("width", "14%")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "學號/日期"

            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("rowspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "姓　名"

            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("rowspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("width", "9%")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "總時數"

            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("rowspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("width", "10%")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "扣除公喪假<br>總時數" '扣除公喪假總時數

            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", int_ColSpn8)
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "時間/假別"

            '表頭2
            nr = New HtmlTableRow : nt.Controls.Add(nr)
            For Each drR1 As DataRow In dtLeave.Rows
                nc = New HtmlTableCell : nr.Controls.Add(nc)
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("style", strCellStyle)
                nc.InnerHtml = Convert.ToString(drR1("Name"))
            Next

            'For j As Integer = 1 To 8
            '    Select Case j
            '        Case 1
            '            nc.InnerHtml = "病假"
            '        Case 2
            '            nc.InnerHtml = "事假"
            '        Case 3
            '            nc.InnerHtml = "公假"
            '        Case 4
            '            nc.InnerHtml = "曠課"
            '        Case 5
            '            nc.InnerHtml = "喪假"
            '        Case 6
            '            nc.InnerHtml = "遲到"
            '        Case 7
            '            nc.InnerHtml = "婚假"
            '        Case 8
            '            nc.InnerHtml = "陪產假"
            '    End Select
            'Next

            '學員資料明細
            Do While intRowCnt <> intCnt And dt.Rows.Count <> 0
                dr = dt.Rows(0) : intCnt += 1

                nr = New HtmlTableRow : nt.Controls.Add(nr)

                Select Case Convert.ToString(dr("type"))
                    Case "class"
                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = "班別:"

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", 3 + int_ColSpn8 + 2)
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("class"))

                    Case "std"
                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("stdid"))

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("stdname"))

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "2")
                        nc.Attributes.Add("align", "right")
                        nc.Attributes.Add("style", strCellStyle)
                        Dim sNcTxt As String = ""
                        If Convert.ToString(dr("lvdate")) <> "" Then sNcTxt = "離退訓日期:"
                        nc.InnerHtml = sNcTxt '"離退訓日期:"

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", int_ColSpn8)
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("lvdate"))

                    Case "leave"
                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("hdate"))

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "3")
                        nc.Attributes.Add("align", "left")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = " "

                        For Each drR1 As DataRow In dtLeave.Rows
                            nc = New HtmlTableCell : nr.Controls.Add(nc)
                            nc.Attributes.Add("colspan", "1")
                            nc.Attributes.Add("align", "center")
                            nc.Attributes.Add("style", strCellStyle)
                            nc.InnerHtml = Convert.ToString(dr("lv" & drR1("leaveid")))
                        Next

                        'For j As Integer = 1 To 8
                        '    nc = New HtmlTableCell : nr.Controls.Add(nc)
                        '    nc.Attributes.Add("colspan", "1")
                        '    nc.Attributes.Add("align", "center")
                        '    nc.Attributes.Add("style", strCellStyle)
                        '    nc.InnerHtml = Convert.ToString(dr("lv0" & j.ToString))
                        'Next

                    Case "total"
                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "2")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = "小計:"

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("total")) '總時數

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("vtotal")) '扣除公喪假總時數

                        For Each drR1 As DataRow In dtLeave.Rows
                            nc = New HtmlTableCell : nr.Controls.Add(nc)
                            nc.Attributes.Add("colspan", "1")
                            nc.Attributes.Add("align", "center")
                            nc.Attributes.Add("style", strCellStyle)
                            nc.InnerHtml = Convert.ToString(dr("lv" & drR1("leaveid")))
                        Next

                        'For j As Integer = 1 To 8
                        '    nc = New HtmlTableCell : nr.Controls.Add(nc)
                        '    nc.Attributes.Add("colspan", "1")
                        '    nc.Attributes.Add("align", "center")
                        '    nc.Attributes.Add("style", strCellStyle)
                        '    nc.InnerHtml = Convert.ToString(dr("lv0" & j.ToString))
                        'Next
                End Select

                dt.Rows.RemoveAt(0)

                'If intRowCnt = intCnt Then
                '    '加入換頁且重新再開始
                '    nl = New HtmlGenericControl
                '    'nl.InnerHtml = "<p style='line-height:2px;margin:%;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'>"
                '    nl.InnerHtml = "<p style='line-height:20px;margin:%;mso-pagination:widow-orphan;'>"
                '    nl.InnerHtml &= "    <br clear=all style='mso-special-character:line-break;page-break-before:always'>"
                '    nl.InnerHtml &= "</p>"
                '    'nl.InnerHtml = "<P style='page-break-after@always'></P>"
                '    div_print_content.Controls.Add(nl)
                'End If
            Loop

            intCnt = 0
        Next
    End Sub


#Region "Function"
    ''' <summary>'整理成報表用格式(取得資料)</summary>
    ''' <returns></returns>
    Function cntData() As DataTable
        Dim dtBase As DataTable = getData("base") '班級學員資料

        Dim dtHours As DataTable = Nothing '學員請假資料
        Dim dtTmp As New DataTable '報表用資料

        Dim dr As DataRow = Nothing
        Dim drTmp As DataRow = Nothing

        Dim strOCID As String = "" '紀錄課程代碼
        Dim strBefLvDate As String = "" '前一筆請假日
        'Dim intSum01 As Integer = 0 '病假總時數
        'Dim intSum02 As Integer = 0 '事假時數
        'Dim intSum03 As Integer = 0 '公假時數
        'Dim intSum04 As Integer = 0 '曠課時數
        'Dim intSum05 As Integer = 0 '喪假時數
        'Dim intSum06 As Integer = 0 '遲假時數
        'Dim intSum07 As Integer = 0 '婚假時數
        'Dim intSum08 As Integer = 0 '陪產假時數
        Dim intTotal As Integer = 0 '總時數
        Dim intVTotal As Integer = 0 '扣除公喪假總時數

        dtTmp.Columns.Add(New DataColumn("type")) '資料型態
        dtTmp.Columns.Add(New DataColumn("ocid")) '學員姓名
        dtTmp.Columns.Add(New DataColumn("ORGNAME")) 'ORGNAME
        dtTmp.Columns.Add(New DataColumn("class")) '班別
        dtTmp.Columns.Add(New DataColumn("stdid")) '學號
        dtTmp.Columns.Add(New DataColumn("stdname")) '學員姓名
        dtTmp.Columns.Add(New DataColumn("lvdate")) '離退日期
        dtTmp.Columns.Add(New DataColumn("hdate")) '請假日期

        For Each drR1 As DataRow In dtLeave.Rows
            dtTmp.Columns.Add(New DataColumn("lv" & drR1("leaveid"))) '病假(總)時數
        Next
        'dtTmp.Columns.Add(New DataColumn("lv01")) '病假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv02")) '事假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv03")) '公假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv04")) '曠課(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv05")) '喪假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv06")) '遲假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv07")) '婚假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv08")) '陪產假(總)時數
        'dtTmp.Columns.Add(New DataColumn("lv11")) '生理假(總)時數

        dtTmp.Columns.Add(New DataColumn("total")) '小計時數
        dtTmp.Columns.Add(New DataColumn("vtotal")) '扣除公喪假總時數

        For i As Integer = 0 To dtBase.Rows.Count - 1
            dr = dtBase.Rows(i)
            dtHours = getData("hours", dr("socid"))

            '第一筆 班別資料(同班別時, 班別只顯示一次)
            Dim ff3_OCID As String = String.Concat("'", dr("ocid"), "'")
            If strOCID.IndexOf(ff3_OCID) = -1 Then
                drTmp = dtTmp.NewRow
                dtTmp.Rows.Add(drTmp)

                drTmp("type") = "class"
                drTmp("ocid") = Convert.ToString(dr("ocid"))
                s_ORGNAME = Convert.ToString(dr("ORGNAME"))
                drTmp("ORGNAME") = Convert.ToString(dr("ORGNAME"))
                drTmp("class") = Convert.ToString(dr("CLASSCNAME2"))

                'If Convert.ToString(dr("cycltype")) <> "" Then
                '    drTmp("class") += "第" & Convert.ToString(dr("cycltype")) & "期"
                'End If
            End If

            '第二筆 學號、姓名、離退訓日期
            drTmp = dtTmp.NewRow
            dtTmp.Rows.Add(drTmp)

            drTmp("type") = "std"
            drTmp("stdid") = Convert.ToString(dr("studentid"))
            drTmp("stdname") = Convert.ToString(dr("name")) & chkStdName(dr("thours"), dtHours)
            'intTHours:班級訓練時數

            If Convert.ToString(dr("rejecttdate1")) <> "" Then drTmp("lvdate") = Convert.ToDateTime(dr("rejecttdate1")).ToString("yyyy/MM/dd")
            If Convert.ToString(dr("rejecttdate2")) <> "" Then drTmp("lvdate") = Convert.ToDateTime(dr("rejecttdate2")).ToString("yyyy/MM/dd")

            '第N筆 請假資料
            For j As Integer = 0 To dtHours.Rows.Count - 1

                '若請假日期相同, 則計算上筆資料
                If strBefLvDate <> Convert.ToDateTime(dtHours.Rows(j)("leavedate")).ToString("yyyy/MM/dd") Then
                    drTmp = dtTmp.NewRow
                    dtTmp.Rows.Add(drTmp)

                    drTmp("type") = "leave"
                    drTmp("hdate") = Convert.ToDateTime(dtHours.Rows(j)("leavedate")).ToString("yyyy/MM/dd")
                    For Each drR1 As DataRow In dtLeave.Rows
                        drTmp("lv" & drR1("leaveid")) = "0"
                    Next

                    'drTmp("lv01") = "0"
                    'drTmp("lv02") = "0"
                    'drTmp("lv03") = "0"
                    'drTmp("lv04") = "0"
                    'drTmp("lv05") = "0"
                    'drTmp("lv06") = "0"
                    'drTmp("lv07") = "0"
                    'drTmp("lv08") = "0"
                Else
                    drTmp = dtTmp.Rows(dtTmp.Rows.Count - 1)
                End If

                '計算時數
                Dim ff3 As String = "leaveid=" & Convert.ToString(dtHours.Rows(j)("leaveid"))
                If dtLeave.Select(ff3).Length > 0 Then
                    Dim drR3 As DataRow = dtLeave.Select(ff3)(0)
                    Dim filedN As String = "lv" & drR3("leaveid")
                    drTmp(filedN) = Convert.ToInt32(drTmp(filedN)) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                    'intSum01 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                    drR3("intSum") = Val(drR3("intSum")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))

                    '扣除公喪假總時數 03:公假 05:喪假 11:生理假
                    '扣除公喪假/生理假總時數 03:公假 05:喪假 11:生理假
                    Select Case Convert.ToString(dtHours.Rows(j)("leaveid"))
                        Case "03", "05", "11"
                        Case Else
                            '扣除公喪假總時數 03:公假 05:喪假 11:生理假
                            intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                    End Select
                End If

                ''計算時數
                'Select Case Convert.ToString(dtHours.Rows(j)("leaveid"))
                '    Case "01"
                '        drTmp("lv01") = Convert.ToInt32(drTmp("lv01")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum01 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "02"
                '        drTmp("lv02") = Convert.ToInt32(drTmp("lv02")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum02 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "03"
                '        drTmp("lv03") = Convert.ToInt32(drTmp("lv03")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum03 += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "04"
                '        drTmp("lv04") = Convert.ToInt32(drTmp("lv04")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum04 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "05"
                '        drTmp("lv05") = Convert.ToInt32(drTmp("lv05")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum05 += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "06"
                '        drTmp("lv06") = Convert.ToInt32(drTmp("lv06")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum06 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "07"
                '        drTmp("lv07") = Convert.ToInt32(drTmp("lv07")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum07 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))

                '    Case "08"
                '        drTmp("lv08") = Convert.ToInt32(drTmp("lv08")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intSum08 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                '        intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                'End Select

                intTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                strBefLvDate = Convert.ToDateTime(dtHours.Rows(j)("leavedate")).ToString("yyyy/MM/dd")
            Next

            '最後一筆 小計
            drTmp = dtTmp.NewRow
            dtTmp.Rows.Add(drTmp)

            drTmp("type") = "total"
            '計算時數
            For Each drR1 As DataRow In dtLeave.Rows
                Dim filedN As String = "lv" & drR1("leaveid")
                drTmp(filedN) = Val(drR1("intSum"))
            Next

            'drTmp("lv01") = intSum01
            'drTmp("lv02") = intSum02
            'drTmp("lv03") = intSum03
            'drTmp("lv04") = intSum04
            'drTmp("lv05") = intSum05
            'drTmp("lv06") = intSum06
            'drTmp("lv07") = intSum07
            'drTmp("lv08") = intSum08
            drTmp("total") = intTotal
            drTmp("vtotal") = intVTotal '扣除公喪假總時數 03:公假 05:喪假 11:生理假

            If strOCID <> "" Then strOCID &= ","
            strOCID &= ff3_OCID 'Convert.ToString(dr("ocid"))

            dtHours.Rows.Clear()
            strBefLvDate = ""
            '計算時數
            For Each drR1 As DataRow In dtLeave.Rows
                'Dim filedN As String = "lv" & drR1("leaveid")
                'drTmp(filedN) = Val(drR1("intSum"))
                drR1("intSum") = 0
            Next

            'intSum01 = 0
            'intSum02 = 0
            'intSum03 = 0
            'intSum04 = 0
            'intSum05 = 0
            'intSum06 = 0
            'intSum07 = 0
            'intSum08 = 0
            intTotal = 0
            intVTotal = 0 '扣除公喪假總時數 03:公假 05:喪假 11:生理假
        Next

        Return dtTmp
    End Function

    '取得資料
    Function getData(ByVal strFlag As String, Optional ByVal strSocid As String = "") As DataTable
        Dim parms As Hashtable = New Hashtable()
        parms.Clear()

        Select Case strFlag
            Case "plan"
                sql = "SELECT PLANNAME FROM dbo.KEY_PLAN WITH(NOLOCK) WHERE TPLAID=@TPLAID"
                parms.Add("TPLAID", hidTPlanID.Value)
            Case "base" '班級學員資料
                sql = ""
                sql &= " SELECT s.OCID"
                sql &= " ,s.SOCID"
                sql &= " ,concat('勞動部',s.ORGNAME) ORGNAME" & vbCrLf
                sql &= " ,s.CLASSCNAME2" & vbCrLf
                sql &= " ,s.CLASSCNAME"
                sql &= " ,s.CYCLTYPE"
                sql &= " ,s.THOURS" '班級訓練時數
                sql &= " ,s.NAME"
                sql &= " ,s.STUDENTID"
                sql &= " ,s.REJECTTDATE1"
                sql &= " ,s.REJECTTDATE2 "
                sql &= " FROM dbo.V_STUDENTINFO s"
                sql &= " WHERE 1=1 "
                If hidOCID.Value <> "" Then
                    sql &= " and s.ocid=@ocid "
                    parms.Add("ocid", hidOCID.Value)
                End If
                If hidRID.Value <> "" Then
                    sql &= " and s.rid=@rid "
                    parms.Add("rid", hidRID.Value)
                End If
                If hidTPlanID.Value <> "" Then
                    sql &= " and s.tplanid=@tplanid "
                    parms.Add("tplanid", hidTPlanID.Value)
                End If
                If hidSDate.Value <> "" OrElse hidEDate.Value <> "" Then
                    sql &= " and exists ("
                    sql &= "  select 'x' "
                    sql &= "  from stud_turnout x "
                    sql &= "  where 1=1"
                    If hidSDate.Value <> "" Then
                        'sql &= "  and x.leavedate>= " & TIMS.to_date(hidSDate.Value)
                        sql &= "  and x.leavedate >= @leavedate1 "
                        parms.Add("leavedate1", hidSDate.Value)
                    End If
                    If hidEDate.Value <> "" Then
                        'sql &= "  and x.leavedate<=" & TIMS.to_date(hidEDate.Value)
                        sql &= "  and x.leavedate <= @leavedate2 "
                        parms.Add("leavedate2", hidEDate.Value)
                    End If
                    sql &= "  and x.socid =s.socid"
                    sql &= " ) "
                End If
                sql &= " order by s.ocid,s.studentid"

            Case "hours" '請假時數
                sql = ""
                sql &= " select socid,LeaveDate,leaveid"
                'sql &= " ,dbo.DECODE(c1,'Y',1,0)+dbo.DECODE(c2,'Y',1,0)+dbo.DECODE(c3,'Y',1,0)+dbo.DECODE(c4,'Y',1,0)"
                'sql &= " +dbo.DECODE(c5,'Y',1,0)+dbo.DECODE(c6,'Y',1,0)+dbo.DECODE(c7,'Y',1,0)+dbo.DECODE(c8,'Y',1,0)"
                'sql &= " +dbo.DECODE(c9,'Y',1,0)+dbo.DECODE(c10,'Y',1,0)+dbo.DECODE(c11,'Y',1,0)+dbo.DECODE(c12,'Y',1,0) hours1 "
                sql &= " ,dbo.FN_GET_HOURS1(SOCID,LEAVEDATE,SEQNO) HOURS1 "
                sql &= " from stud_turnout "
                sql &= " where 1=1"
                If hidSDate.Value <> "" Then
                    'sql &= " and LeaveDate>= " & TIMS.to_date(hidSDate.Value)
                    sql &= " and LeaveDate >= @LeaveDate1 "
                    parms.Add("LeaveDate1", hidSDate.Value)
                End If
                If hidEDate.Value <> "" Then
                    'sql &= " and LeaveDate<= " & TIMS.to_date(hidEDate.Value)
                    sql &= " and LeaveDate <= @LeaveDate2 "
                    parms.Add("LeaveDate2", hidEDate.Value)
                End If
                sql &= " and socid=@socid "
                sql &= " order by leavedate"

                parms.Add("socid", strSocid)
        End Select

        Return DbAccess.GetDataTable(sql, objconn, parms)
    End Function


    '判斷學員請假時數是否超過1/5、1/15, 回傳文字
    Function chkStdName(ByVal intTHours As Double, ByVal dt As DataTable) As String
        Dim strRtn As String = ""
        Dim intTotal As Double = 0 '請假時數
        'intTHours:班級訓練時數
        For i As Integer = 0 To dt.Rows.Count - 1
            If IsNumeric(dt.Rows(i)("hours1")) Then
                intTotal += Convert.ToInt32(dt.Rows(i)("hours1"))
            End If
        Next

        '是否為百分比 1:依比例/2:依百分比
        Dim r4rbl As String = TIMS.GetGlobalVar4rbl(Me, "", objconn)
        Dim x2 As Double = TIMS.GetGlobalVar4(Me, "2", objconn)
        Dim x1 As Double = TIMS.GetGlobalVar4(Me, "1", objconn)
        Dim iTest As Double = TIMS.ROUND((intTotal / intTHours), 4)
        Select Case r4rbl
            Case "1" '是否為百分比 1:依比例/2:依百分比
                Dim sX2 As String = TIMS.GetGlobalVar(Me, "4", "2", objconn)
                Dim sX1 As String = TIMS.GetGlobalVar(Me, "4", "1", objconn)
                If x2 < iTest Then
                    strRtn = "(請假總時數已超過" & sX2 & ")"
                ElseIf x1 < iTest Then
                    strRtn = "(請假總時數已超過" & sX1 & ")"
                End If
            Case Else '"2"'是否為百分比 1:依比例/2:依百分比
                Dim sX2pp As String = TIMS.GetGlobalVar4rbl(Me, "2", objconn)
                Dim sX1pp As String = TIMS.GetGlobalVar4rbl(Me, "1", objconn)
                If x2 < iTest Then
                    strRtn = "(請假總時數已超過" & sX2pp & ")"
                ElseIf x1 < iTest Then
                    strRtn = "(請假總時數已超過" & sX1pp & ")"
                End If
        End Select
        'If intTHours / intTotal <= 5 Then
        '    strRtn = "(請假總時數已超過1/5)"
        'ElseIf intTHours / intTotal <= 15 Then
        '    strRtn = "(請假總時數已超過1/15)"
        'End If
        Return strRtn
    End Function
#End Region

#Region "NO USE"
    ''取得參數名稱
    'Private Function getName(ByVal strFlag As String) As String
    '    Dim dt As New DataTable
    '    Dim strRtn As String = ""
    '    Select Case strFlag
    '        Case "plan"
    '            sql = "select planname from key_plan where tplanid='" & hidTPlanID.Value & "'"
    '        Case "user"
    '            sql = "select name from auth_account where account='" & hidUserID.Value & "'"
    '    End Select
    '    dt = DbAccess.GetDataTable(sql, objconn)
    '    If dt.Rows.Count > 0 Then strRtn = Convert.ToString(dt.Rows(0)(0))
    '    Return strRtn
    'End Function
#End Region

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call DbAccess.Close(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    '若SESSION 異常
        '    Call TIMS.sUtl_404NOTFOUND(Me, objconn)
        '    Exit Sub
        'End If
        'Dim flagYear2017 As Boolean = False
        'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)

        If Not IsPostBack Then
            hidOCID.Value = Convert.ToString(Request("OCID"))
            hidTMID.Value = Convert.ToString(Request("TMID")) 'TrainID/TMID
            hidTPlanID.Value = Convert.ToString(Request("TPlanID"))
            hidRID.Value = Convert.ToString(Request("RID"))
            hidSDate.Value = Convert.ToString(Request("start_date"))
            hidEDate.Value = Convert.ToString(Request("end_date"))
            hidUserID.Value = Convert.ToString(Request("UserID"))
            hidprtPageSize.Value = Convert.ToString(Request("prtPageSize"))

            btnCancel.Attributes.Add("onclick", "window.close();")

            crtTable()
        End If
    End Sub

    '匯出 PDF
    Sub Export_PDF1()
        Dim YMDSTR1x As String = DateTime.Now.ToString("ssHHddMMyyyymmss")
        Dim strFileName As String = String.Concat(YMDSTR1x, ".pdf")
        Dim s_Charset As String = TIMS.cst_Charset_UTF8 '"UTF-8" 'default

        Response.Clear()
        'MyPage.Response.ClearHeaders()
        'MyPage.Response.Charset = "UTF-8"
        'MyPage.Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        Response.Charset = s_Charset
        Response.ContentEncoding = System.Text.Encoding.GetEncoding(s_Charset)
        Response.ContentType = "application/pdf" 'PDF
        Response.AppendHeader("Content-Disposition", String.Concat("attachment; filename=", strFileName))

        Dim sw As New System.IO.StringWriter
        Dim htw As New HtmlTextWriter(sw)
        div_print_content.RenderControl(htw)
        Dim strHtml As String = sw.ToString().Replace("<div>", "").Replace("</div>", "")

        Using stream As New System.IO.MemoryStream
            Dim pdf As HiQPdf.HtmlToPdf = New HiQPdf.HtmlToPdf With {
                .SerialNumber = HiQPdf_SerialNumber '"/7eWrq+b-mbOWnY2e-jYbOz9HP-387fy9/G-yM7fzM7R-zs3RxsbG-xg=="   '「HiQPdf」的SerialNumber
                }
            pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Portrait   '紙張直向
            'pdf.Document.PageOrientation = HiQPdf.PdfPageOrientation.Landscape   '紙張橫向
            pdf.ConvertHtmlToStream(strHtml, Nothing, stream)

            Response.BinaryWrite(stream.ToArray()) '輸出PDF檔案。
        End Using
    End Sub

    '列印PDF
    'Private Sub btnPrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrt.Click
    '    crtTable()

    '    trBtn.Attributes.Remove("style")
    '    trBtn.Attributes.Add("style", "display:none")

    '    Call Export_PDF1()
    'End Sub

    '匯出 匯出明細
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        'dtLeave = Get_dtLeave() 'DbAccess.GetDataTable(sql, objconn)
        'dtLeave = TIMS.GET_LEAVEdt("", objconn)
        dtLeave = TIMS.Get_dtLEAVE(objconn)
        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select cs.OCID" & vbCrLf
        sql &= " ,cs.SOCID" & vbCrLf
        sql &= " ,cs.CLASSCNAME2" & vbCrLf
        sql &= " ,cs.NAME" & vbCrLf
        sql &= " ,cs.studentid" & vbCrLf
        sql &= " ,CONVERT(varchar, cs.rejecttdate1, 111) rejecttdate1" & vbCrLf
        sql &= " ,CONVERT(varchar, cs.rejecttdate2, 111) rejecttdate2" & vbCrLf
        sql &= " ,CONVERT(varchar, tt.LeaveDate, 111) LeaveDate" & vbCrLf
        sql &= " ,tt.leaveid" & vbCrLf
        sql &= " ,k1.name leaveName" & vbCrLf
        For Each dr1 As DataRow In dtLeave.Rows
            sql &= " ,ISNULL(CASE WHEN tt.leaveid='" & dr1("leaveid") & "' THEN tt.Hours END,0) leaveHours" & dr1("leaveid") & vbCrLf
        Next
        sql &= " ,isnull(tt.Hours,0) Hours" & vbCrLf
        sql &= " FROM STUD_TURNOUT tt" & vbCrLf
        sql &= " JOIN V_STUDENTINFO cs on cs.socid =tt.socid" & vbCrLf
        sql &= " LEFT JOIN KEY_LEAVE k1 on k1.leaveid=tt.leaveid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

        If hidOCID.Value <> "" Then
            sql &= " and cs.ocid=@ocid " & vbCrLf
            parms.Add("ocid", hidOCID.Value)
        End If
        If hidRID.Value <> "" Then
            sql &= " and cs.rid=@rid " & vbCrLf
            parms.Add("rid", hidRID.Value)
        End If
        If hidTPlanID.Value <> "" Then
            sql &= " and cs.tplanid=@tplanid " & vbCrLf
            parms.Add("tplanid", hidTPlanID.Value)
        End If
        If hidSDate.Value <> "" Then
            'sql &= "  and tt.leavedate>= " & TIMS.to_date(hidSDate.Value)
            sql &= "  and tt.leavedate >= @leavedate1 " & vbCrLf
            parms.Add("leavedate1", hidSDate.Value)
        End If
        If hidEDate.Value <> "" Then
            'sql &= "  and tt.leavedate<=" & TIMS.to_date(hidEDate.Value)
            sql &= "  and tt.leavedate <= @leavedate2 " & vbCrLf
            parms.Add("leavedate2", hidEDate.Value)
        End If
        'sql += " and ip.tplanid ='02'" & vbCrLf
        'sql += " and ip.distid ='001'" & vbCrLf
        'sql += " and ip.years ='2014'" & vbCrLf
        'sql += " and cc.ocid =57498" & vbCrLf
        sql &= " order by cs.studentid,tt.LeaveDate" & vbCrLf

        Dim dt As DataTable = Nothing
        dt = DbAccess.GetDataTable(sql, objconn, parms)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If
        Call ExpReport1(dt)

    End Sub

    '匯出 匯出明細 Response 
    Sub ExpReport1(ByRef dt As DataTable)

        Dim strTitle1 As String = "" '匯出表頭名稱
        strTitle1 = "出缺勤明細表"

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=utf-8>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        '套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        'mso-number-format:"0" 
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""
        '建立抬頭
        '第1行

        ExportStr = "<tr>" & vbCrLf
        ExportStr &= "<td>班別</td>" & vbTab
        ExportStr &= "<td>學號</td>" & vbTab '訓練機構
        ExportStr &= "<td>日期</td>" & vbTab '
        ExportStr &= "<td>姓名</td>" & vbTab '
        '假別
        For Each dr As DataRow In dtLeave.Rows
            ExportStr &= "<td>" & dr("Name") & "</td>" & vbTab '
        Next
        ExportStr &= "<td>離退訓日期</td>" & vbTab '
        ExportStr &= "<td>請假時數</td>" & vbTab
        ExportStr += "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        For Each dr As DataRow In dt.Rows
            'For Each dr As DataRow In dt.Rows
            '建立資料面

            ExportStr = "<tr>" & vbCrLf
            ExportStr &= "<td>" & Convert.ToString(dr("CLASSCNAME2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("studentid")) & "</td>" & vbTab '訓練機構
            ExportStr &= "<td>" & Convert.ToString(dr("LeaveDate")) & "</td>" & vbTab '
            ExportStr &= "<td>" & Convert.ToString(dr("name")) & "</td>" & vbTab '
            '假別
            For Each dr1 As DataRow In dtLeave.Rows
                Dim s_COLN_LEAVE As String = String.Concat("leaveHours", dr1("leaveid"))
                ExportStr &= "<td>" & Convert.ToString(dr(s_COLN_LEAVE)) & "</td>"
            Next
            Dim rejecttdate As String = If(Convert.ToString(dr("rejecttdate1")) <> "", Convert.ToString(dr("rejecttdate1")), If(Convert.ToString(dr("rejecttdate2")) <> "", Convert.ToString(dr("rejecttdate2")), " "))
            ExportStr &= String.Format("<td>{0}</td>", rejecttdate) & vbTab '
            ExportStr &= "<td>" & Convert.ToString(dr("Hours")) & "</td>" & vbTab
            ExportStr += "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")
        Call TIMS.CloseDbConn(objconn)
        Response.End()

    End Sub
End Class
