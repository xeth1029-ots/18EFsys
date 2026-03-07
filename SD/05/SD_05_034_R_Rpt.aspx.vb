Partial Class SD_05_034_R_Rpt
    Inherits AuthBasePage

    Const cst_plan As String = "plan" '計畫範圍
    Const cst_baseStud As String = "baseStud" '班級學員資料
    Const cst_baseStudExp As String = "baseStudExp" '班級學員資料(匯出明細)
    Const cst_hours As String = "hours" '學員請假時數
    Const cst_baseleave As String = "baseleave" '請假代碼回傳
    Dim dtLeave As DataTable = Nothing
    Dim sql As String = ""

    'Dim flagYear2017 As Boolean = False
    'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)
    'Dim flag_test1 As Boolean = False '測試環境為true/正式為false

    Dim objconn As SqlConnection

    'Function Get_dtLeave() As DataTable
    '    Dim rst As DataTable = Nothing
    '    sql = "" & vbCrLf
    '    sql &= " select leaveid" & vbCrLf
    '    sql &= " ,Name " & vbCrLf
    '    sql &= " ,0 as intSum" & vbCrLf
    '    sql &= " from Key_Leave " & vbCrLf
    '    sql &= " where 1=1" & vbCrLf
    '    sql &= " and leaveid in ('01','02','03','04','05','06','07','08')" & vbCrLf
    '    sql &= " order by leaveid " & vbCrLf
    '    If flagYear2017 Then
    '        sql = "" & vbCrLf
    '        sql &= " select leaveid" & vbCrLf
    '        sql &= " ,Name " & vbCrLf
    '        sql &= " ,0 as intSum" & vbCrLf
    '        sql &= " FROM KEY_LEAVE " & vbCrLf
    '        sql &= " where 1=1" & vbCrLf
    '        sql &= " and NOUSE IS NULL" & vbCrLf
    '        sql &= " order by LEAVESORT " & vbCrLf
    '    End If
    '    rst = DbAccess.GetDataTable(sql, objconn)
    '    Return rst
    'End Function

#Region "TABLE PROCESS"
    '建置表格
    Function crtTable() As Boolean
        'dtLeave = TIMS.GET_LEAVEdt("", objconn) 'DbAccess.GetDataTable(sql, objconn)
        dtLeave = TIMS.Get_dtLEAVE(objconn)
        Dim int_ColSpn8 As Integer = dtLeave.Rows.Count
        Dim rst As Boolean = True '正常列印報表,false:異常。
        Dim dt As DataTable = cntData()
        If dt.Rows.Count = 0 Then
            Return False
            'Common.MessageBox(Me, "查無資料!!")
            'Exit Function
        End If

        Dim dr As DataRow = Nothing
        Dim intPageCnt As Integer = 0  '頁數控制
        Dim intRowCnt As Integer = 54 '每頁筆數
        Dim intCnt As Integer = 0 '報表處理筆數

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl

        Dim strTitleStyle As String = "font-size:14pt;font-family:DFKai-SB"
        Dim strCellStyle As String = "font-size:10pt;font-family:DFKai-SB"

        intPageCnt = dt.Rows.Count / intRowCnt
        If intPageCnt * intRowCnt < dt.Rows.Count Then
            intPageCnt += 1
        End If

        For i As Integer = 1 To intPageCnt
            '加背景圖的div
            nl = New HtmlGenericControl
            div_print.Controls.Add(nl)
            'nl.InnerHtml = "<div style='position:fixed;z-index:-1; margin:0;padding:0;left:0px;top:0px;'>" & vbCrLf
            nl.InnerHtml = "<div style='position:absolute;z-index:-1; margin:0;padding:0;left:0px;top:0px;'>" & vbCrLf
            nl.InnerHtml += "   <img src='../../images/rptpic/temple/TIMS_1.jpg' height='1030' />" & vbCrLf
            nl.InnerHtml += "</div>" & vbCrLf

            '表首
            nt = New HtmlTable
            nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", "0")
            div_print.Controls.Add(nt)

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strTitleStyle)
            nc.InnerHtml = sm.UserInfo.OrgName

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strTitleStyle)
            'Dim sTPlanName As String = TIMS.GetTPlanName(hidTPlanID.Value)
            'nc.InnerHtml = sTPlanName & "　" & "屆退官兵出缺勤明細表"
            nc.InnerHtml = "屆退官兵出缺勤明細表"

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "查詢日期：" & hidSDate.Value & "~" & hidEDate.Value

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
            nc.InnerHtml = "列印日期：" & Now.ToString("yyyy/MM/dd")

            nr = New HtmlTableRow : nt.Controls.Add(nr)
            nc = New HtmlTableCell : nr.Controls.Add(nc)
            nc.Attributes.Add("colspan", "2")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", strCellStyle)
            nc.InnerHtml = "　　頁數：" & i.ToString & " / " + intPageCnt.ToString

            '表頭
            nt = New HtmlTable
            nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
            nt.Attributes.Add("align", "center")
            nt.Attributes.Add("border", 1)
            div_print.Controls.Add(nt)

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
            nc.InnerHtml = "姓　　名"

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
            nc.InnerHtml = "扣除公喪假<br>總時數"

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
                nc.Attributes.Add("width", "7%")
                nc.Attributes.Add("style", strCellStyle)
                nc.InnerHtml = Convert.ToString(drR1("Name"))
            Next

            'For j As Integer = 1 To 8
            '    nc = New HtmlTableCell : nr.Controls.Add(nc)
            '    nc.Attributes.Add("align", "center")
            '    nc.Attributes.Add("width", "7%")
            '    nc.Attributes.Add("style", strCellStyle)
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
                        nc.Attributes.Add("colspan", 3 + int_ColSpn8)
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
                        nc.InnerHtml = Convert.ToString(dr("total"))

                        nc = New HtmlTableCell : nr.Controls.Add(nc)
                        nc.Attributes.Add("colspan", "1")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strCellStyle)
                        nc.InnerHtml = Convert.ToString(dr("vtotal"))

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

                If intRowCnt = intCnt Then
                    '加入換頁且重新再開始
                    nl = New HtmlGenericControl
                    div_print.Controls.Add(nl)
                    nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'>"
                    nl.InnerHtml += "    <br clear=all style='mso-special-character:line-break;page-break-before:always'>"
                    nl.InnerHtml += "</p>"
                End If
            Loop

            intCnt = 0
        Next
        Return rst
    End Function

    '整理成報表用格式
    Function cntData() As DataTable

        Dim dt As DataTable = getData(cst_baseStud, "") '班級學員資料
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
        Dim intVTotal As Integer = 0 '公喪假總時數

        dtTmp.Columns.Add(New DataColumn("type")) '資料型態
        dtTmp.Columns.Add(New DataColumn("ocid")) '學員姓名
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
        dtTmp.Columns.Add(New DataColumn("total")) '小計時數
        dtTmp.Columns.Add(New DataColumn("vtotal")) '扣除公喪假總時數

        For i As Integer = 0 To dt.Rows.Count - 1
            dr = dt.Rows(i)
            dtHours = getData(cst_hours, dr("socid"))

            '第一筆 班別資料(同班別時, 班別只顯示一次)
            If strOCID.IndexOf(Convert.ToString(dr("ocid"))) < 0 Then
                drTmp = dtTmp.NewRow
                dtTmp.Rows.Add(drTmp)
                drTmp("type") = "class"
                drTmp("ocid") = Convert.ToString(dr("ocid"))
                drTmp("class") = Convert.ToString(dr("classcname2"))
                'If Convert.ToString(dr("cycltype")) <> "" Then
                '    drTmp("class") += "第" & Convert.ToString(dr("cycltype")) & "期"
                'End If
            End If

            '第二筆 學號、姓名、離退訓日期
            drTmp = dtTmp.NewRow
            dtTmp.Rows.Add(drTmp)

            drTmp("type") = "std"
            drTmp("stdid") = Convert.ToString(dr("studentid"))
            drTmp("stdname") = Convert.ToString(dr("name")) & chkStdName(Convert.ToInt32(dr("thours")), dtHours)

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

                Dim ff3 As String = "leaveid=" & Convert.ToString(dtHours.Rows(j)("leaveid"))
                If dtLeave.Select(ff3).Length > 0 Then
                    Dim drR3 As DataRow = dtLeave.Select(ff3)(0)
                    Dim filedN As String = "lv" & drR3("leaveid")
                    drTmp(filedN) = Convert.ToInt32(drTmp(filedN)) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                    'intSum01 += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                    drR3("intSum") = Val(drR3("intSum")) + Convert.ToInt32(dtHours.Rows(j)("hours1"))
                    intVTotal += Convert.ToInt32(dtHours.Rows(j)("hours1"))
                End If

                '計算時數
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
            drTmp("vtotal") = intVTotal

            strOCID += Convert.ToString(dr("ocid"))
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
            intVTotal = 0
        Next

        Return dtTmp
    End Function

    '判斷學員請假時數是否超過1/5、1/15, 回傳文字
    Private Function chkStdName(ByVal intTHours As String, ByVal dt As DataTable) As String
        Dim strRtn As String = ""
        Dim intTotal As Integer = 0

        For i As Integer = 0 To dt.Rows.Count - 1
            If IsNumeric(dt.Rows(i)("hours1")) Then
                intTotal += Convert.ToInt32(dt.Rows(i)("hours1"))
            End If
        Next

        If intTHours / intTotal <= 5 Then
            strRtn = "(請假總時數已超過1/5)"
        ElseIf intTHours / intTotal <= 15 Then
            strRtn = "(請假總時數已超過1/15)"
        End If

        Return strRtn
    End Function
#End Region

    '取得資料
    Function getData(ByVal strFlag As String, ByVal strSocid As String) As DataTable
        Dim dt1 As DataTable = Nothing
        'Optional ByVal strSocid As String = ""
        Select Case strFlag
            Case cst_plan '"plan"
                sql = "SELECT PLANNAME FROM KEY_PLAN WHERE TPLANID in (" & hidTPlanID.Value & ") ORDER BY TPLANID"
            Case cst_baseStud '"base" '班級學員資料
                'Dim sql As String = ""
                Dim ssYears As String = Convert.ToString("Years")
                If ssYears = "" Then ssYears = Convert.ToString(Now.Year)

                sql = "" & vbCrLf
                sql &= " WITH WC1 AS (" & vbCrLf
                sql &= " select s.ocid" & vbCrLf
                sql &= " ,s.socid" & vbCrLf
                sql &= " ,s.classcname" & vbCrLf
                sql &= " ,s.cycltype" & vbCrLf
                sql &= " ,s.thours" & vbCrLf
                sql &= " ,s.name" & vbCrLf
                sql &= " ,s.studentid" & vbCrLf
                sql &= " ,s.rejecttdate1" & vbCrLf
                sql &= " ,s.rejecttdate2" & vbCrLf
                sql &= " ,s.tplanid,s.distid,s.planid,s.comidno,s.studid" & vbCrLf
                sql &= " FROM V_STUDENTINFO s" & vbCrLf
                sql &= " WHERE 1=1" & vbCrLf
                sql &= " and s.MIDENTITYID ='12'" & vbCrLf
                sql &= " AND S.YEARS ='" & ssYears & "'" & vbCrLf
                If hidRID.Value <> "" Then
                    sql &= " and s.rid='" & hidRID.Value & "'" & vbCrLf
                End If
                If hidTPlanID.Value <> "" Then
                    sql &= " and s.tplanid in (" & hidTPlanID.Value & ")" & vbCrLf
                End If
                sql &= " )" & vbCrLf
                sql &= " ,WST1 AS (" & vbCrLf
                sql &= " select DISTINCT x.socid" & vbCrLf
                sql &= " FROM WC1 c" & vbCrLf
                sql &= " JOIN STUD_TURNOUT x on x.socid =c.socid" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                'sql &= " and x.leavedate >= convert(datetime, '2017/03/01', 111)" & vbCrLf
                'sql &= " and x.leavedate <=convert(datetime, '2017/03/31', 111)" & vbCrLf
                If hidSDate.Value <> "" Then
                    sql += " and x.leavedate >= " & TIMS.To_date(hidSDate.Value) & vbCrLf
                End If
                If hidEDate.Value <> "" Then
                    sql += " and x.leavedate <=" & TIMS.To_date(hidEDate.Value) & vbCrLf
                End If
                sql &= " )" & vbCrLf
                sql &= " select s.ocid" & vbCrLf
                sql &= " ,s.socid" & vbCrLf
                sql &= " ,s.classcname" & vbCrLf
                sql &= " ,s.cycltype" & vbCrLf
                sql &= " ,dbo.FN_GET_CLASSCNAME(s.CLASSCNAME,s.CYCLTYPE) CLASSCNAME2 " & vbCrLf
                sql &= " ,s.thours" & vbCrLf
                sql &= " ,s.name" & vbCrLf
                sql &= " ,s.studentid" & vbCrLf
                sql &= " ,dbo.FN_CSTUDID2(s.studentid) STUDID" & vbCrLf
                sql &= " ,s.rejecttdate1" & vbCrLf
                sql &= " ,s.rejecttdate2" & vbCrLf
                sql &= " ,s.tplanid,s.distid,s.planid,s.comidno,s.studid" & vbCrLf
                sql &= " FROM WC1 s" & vbCrLf
                sql &= " JOIN WST1 t on t.socid =s.socid " & vbCrLf
                sql &= " ORDER BY s.tplanid,s.distid,s.planid,s.comidno,s.OCID,s.studid" & vbCrLf
                'K221769640
            Case cst_baseStudExp
                sql = "" & vbCrLf
                sql &= " select ss.ocid" & vbCrLf
                sql &= " ,ss.socid" & vbCrLf
                sql &= " ,dbo.FN_GET_CLASSCNAME(ss.CLASSCNAME,ss.CYCLTYPE) CLASSCNAME2 " & vbCrLf
                sql &= " ,ss.name" & vbCrLf
                sql &= " ,ss.studentid" & vbCrLf
                sql &= " ,CONVERT(varchar, ss.rejecttdate1, 111) rejecttdate1" & vbCrLf
                sql &= " ,CONVERT(varchar, ss.rejecttdate2, 111) rejecttdate2" & vbCrLf
                sql &= " ,CONVERT(varchar, tt.LeaveDate, 111) LeaveDate" & vbCrLf
                sql &= " ,tt.leaveid" & vbCrLf
                sql &= " ,k1.name leaveName" & vbCrLf

                For Each dr1 As DataRow In dtLeave.Rows
                    sql &= String.Concat(" ,ISNULL(CASE tt.LEAVEID WHEN '", dr1("leaveid"), "' THEN tt.Hours END,0) leaveHours", dr1("leaveid")) & vbCrLf
                Next
                sql &= " ,dbo.NVL(tt.Hours,0) Hours" & vbCrLf
                sql &= " FROM STUD_TURNOUT tt" & vbCrLf
                sql &= " JOIN V_STUDENTINFO ss on ss.socid =tt.socid" & vbCrLf
                sql &= " JOIN KEY_LEAVE k1 on k1.leaveid=tt.leaveid" & vbCrLf
                sql &= " where 1=1" & vbCrLf
                sql &= " and ss.MIDENTITYID ='12'" & vbCrLf
                'sql &= " and ss.rid='B'" & vbCrLf
                'sql &= " and ss.tplanid in ('02','14','64','65')" & vbCrLf
                If hidRID.Value <> "" Then sql += " and ss.rid='" & hidRID.Value & "' " & vbCrLf
                If hidTPlanID.Value <> "" Then sql += " and ss.tplanid in (" & hidTPlanID.Value & ")" & vbCrLf
                If hidSDate.Value <> "" Then
                    sql &= "  and tt.leavedate>= " & TIMS.To_date(hidSDate.Value) & vbCrLf
                End If
                If hidEDate.Value <> "" Then
                    sql &= "  and tt.leavedate<=" & TIMS.To_date(hidEDate.Value) & vbCrLf
                End If
                sql &= " order by ss.studentid,tt.LeaveDate" & vbCrLf
            Case cst_hours '"hours" '請假時數
                sql = ""
                sql &= " select socid,LeaveDate,leaveid" & vbCrLf
                sql += " ,dbo.DECODE(c1,'Y',1,0)+dbo.DECODE(c2,'Y',1,0)+dbo.DECODE(c3,'Y',1,0)+dbo.DECODE(c4,'Y',1,0)" & vbCrLf
                sql += " +dbo.DECODE(c5,'Y',1,0)+dbo.DECODE(c6,'Y',1,0)+dbo.DECODE(c7,'Y',1,0)+dbo.DECODE(c8,'Y',1,0)" & vbCrLf
                sql += " +dbo.DECODE(c9,'Y',1,0)+dbo.DECODE(c10,'Y',1,0)+dbo.DECODE(c11,'Y',1,0)+dbo.DECODE(c12,'Y',1,0) hours1 " & vbCrLf
                sql += " FROM STUD_TURNOUT " & vbCrLf
                sql += " where 1=1" & vbCrLf
                If hidSDate.Value <> "" Then
                    sql += " and LeaveDate>= " & TIMS.To_date(hidSDate.Value) & vbCrLf
                End If
                If hidEDate.Value <> "" Then
                    sql += " and LeaveDate<= " & TIMS.To_date(hidEDate.Value) & vbCrLf
                End If
                sql += " and socid =" & strSocid & vbCrLf
                sql += " order by LeaveDate,leaveid" & vbCrLf
            Case cst_baseleave '"baseleave"
                dt1 = TIMS.Get_dtLEAVE(objconn)
                Return dt1
        End Select
        If sql = "" Then Return dt1
        dt1 = DbAccess.GetDataTable(sql, objconn)
        Return dt1
    End Function

    '匯出明細
    'Function Search1dt() As DataTable
    '    Dim rst As DataTable = getData(cst_baseStudExp) '班級學員資料(匯出明細)
    '    Return rst
    'End Function

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

        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then
        '    '若SESSION 異常
        '    Call TIMS.sUtl_404NOTFOUND(Me, objconn)
        '    Exit Sub
        'End If
        'Dim flagYear2017 As Boolean = False
        'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)
        'If TIMS.sUtl_ChkTest() Then flag_test1 = True '測試環境為true

        If Not IsPostBack Then
            'hidOCID.Value = Convert.ToString(Request("OCID"))
            'hidTMID.Value = Convert.ToString(Request("TMID")) 'TrainID/TMID
            hidTPlanID.Value = Convert.ToString(Request("TPlanID"))
            hidTPlanID.Value = TIMS.Get_SplitValeu1(hidTPlanID.Value)

            hidRID.Value = Convert.ToString(Request("RID"))
            hidSDate.Value = Convert.ToString(Request("STDate1"))
            hidEDate.Value = Convert.ToString(Request("STDate2"))
            hidUserID.Value = Convert.ToString(Request("UserID"))

            btnCancel.Attributes.Add("onclick", "window.close();")

            'Call crtTable()
            If Not crtTable() Then
                Common.MessageBox(Me, "查無資料!!")
                Exit Sub
            End If
        End If
    End Sub

    '列印
    Private Sub btnPrt_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrt.Click
        'Call crtTable()
        If Not crtTable() Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        trBtn.Attributes.Remove("style")
        trBtn.Attributes.Add("style", "display:none")

        Dim strScript As String = ""
        strScript = "<script language=""javascript"">window.print();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        Return

        'strScript = "<script language=""javascript"">" + vbCrLf
        'strScript += "  if (factory.object) {"
        'strScript += "      factory.printing.header = """";"
        'strScript += "      factory.printing.footer = """";"
        'strScript += "      factory.printing.leftMargin = 5; "
        'strScript += "      factory.printing.topMargin = 10; "
        'strScript += "      factory.printing.rightMargin = 5; "
        'strScript += "      factory.printing.bottomMargin = 10; "
        'strScript += "      factory.printing.portrait = " + TIMS.c_true + ";"
        'strScript += "      factory.printing.Print(true);"
        'strScript += "      window.close();"
        'strScript += "  }"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        'Return
    End Sub

    '匯出(匯出明細)
    Protected Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        'dtLeave = getData(cst_baseleave)
        'dtLeave = Get_dtLeave() 'DbAccess.GetDataTable(sql, objconn)
        dtLeave = TIMS.Get_dtLEAVE(objconn)
        'Dim dt As DataTable = Search1dt()
        Dim dt As DataTable = getData(cst_baseStudExp, "") '班級學員資料(匯出明細)
        'Dim dt As DataTable = Nothing
        'dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!!")
            Exit Sub
        End If

        Call ExpReport1(dt)
    End Sub

    '匯出 Response 
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
            ExportStr &= "<td>" & Convert.ToString(dr("classcname2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("studentid")) & "</td>" & vbTab '訓練機構
            ExportStr &= "<td>" & Convert.ToString(dr("LeaveDate")) & "</td>" & vbTab '
            ExportStr &= "<td>" & Convert.ToString(dr("name")) & "</td>" & vbTab '
            '假別
            For Each dr1 As DataRow In dtLeave.Rows
                ExportStr &= "<td>" & Convert.ToString(dr("leaveHours" & dr1("leaveid"))) & "</td>" & vbTab '
            Next
            Dim rejecttdate As String = " "
            If Convert.ToString(dr("rejecttdate1")) <> "" Then rejecttdate = Convert.ToString(dr("rejecttdate1"))
            If Convert.ToString(dr("rejecttdate2")) <> "" Then rejecttdate = Convert.ToString(dr("rejecttdate2"))
            ExportStr &= "<td>" & rejecttdate & "</td>" & vbTab '
            ExportStr &= "<td>" & Convert.ToString(dr("Hours")) & "</td>" & vbTab
            ExportStr += "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

        Response.End()

    End Sub
End Class
