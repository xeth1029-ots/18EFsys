Partial Class TR_04_010
    Inherits AuthBasePage

    Const cst_sExp1 As String = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數-不就業+提前就業人數)"
    Const cst_sExp2 As String = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數)"
    Const cst_sExp3 As String = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數-在職者人數)"

#Region "Function2"
    'Dim sOutText As String = ""
    'Sub Show_OutText(ByVal sOutText As String)
    '    Range1.Text = TIMS.GetMyValue(sOutText, "Range1")
    '    Range2.Text = TIMS.GetMyValue(sOutText, "Range2")
    '    Range3.Text = TIMS.GetMyValue(sOutText, "Range3")
    '    Range4.Text = TIMS.GetMyValue(sOutText, "Range4")
    'End Sub
#End Region

    'TR_04002_TD2
    Const cst_css_TR_04002_TD2 As String = "TR_04002_TD2"
    Const cst_css_TR_04002_TR As String = "TR_04002_TR"

    Dim sInputText As String = ""
    Const Cst_CommandTimeout As Integer = 30 '1000
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
            DistID = TIMS.Get_DistID(DistID)

            '排除計畫
            '06:在職進修訓練 
            '07:接受企業委託訓練
            '15:學習券
            TPlanID = TIMS.Get_TPlan(TPlanID, , 1, , "TPlanID not in ('06','07','15') ")
            Page.RegisterStartupScript("", "<script>GetMode();</script>")

            Session("TitleTable") = Nothing
            Session("SearchStr") = Nothing

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

            Common.SetListItem(DistID, sm.UserInfo.DistID)
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)
        End If

        If RIDValue.Value = "" Then
            OCID.Items.Clear()
            OCID.Items.Add("請選擇機構")
        End If

        OCID.Attributes("onchange") = "if(this.selectedIndex!=0){document.form1.OCIDValue.value=this.value;}else{document.form1.OCIDValue.value='';}"
        DistID.Attributes("onchange") = "GetMode();"
        TPlanID.Attributes("onchange") = "GetMode();"
        Button1.Attributes("onclick") = "return search();"
        Button2.Style("display") = "none"

        ''選訓練機構
        'Button3.Attributes("onclick") = "javascript:wopen('../../Common/MainOrg.aspx?DistID='+document.form1.DistID.value+'&amp;TPlanID='+document.form1.TPlanID.value,'訓練機構',400,400,1)"
        '列印就業率
        Button4.Attributes("onclick") = "window.open('TR_04_010_R.aspx?Mode=1','print1','resizable=1,scrollbars=1,status=1')"
        '列印參考資料
        Button5.Attributes("onclick") = "window.open('TR_04_010_R.aspx?Mode=2','print2','resizable=1,scrollbars=1,status=1')"

        MenuTable.Visible = False
        Button4.Disabled = True
        Button5.Disabled = True

    End Sub

    '查詢1
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Dim sql As String
        'Dim dt As DataTable
        Dim dr As DataRow
        Dim SearchStr As String
        Dim TitleTable As New DataTable

        '建立TitleTable的欄位---------------------Start
        TitleTable.Columns.Add(New DataColumn("STDate1")) '開訓期間
        TitleTable.Columns.Add(New DataColumn("STDate2"))
        TitleTable.Columns.Add(New DataColumn("FTDate1")) '結訓期間
        TitleTable.Columns.Add(New DataColumn("FTDate2"))
        TitleTable.Columns.Add(New DataColumn("DistID")) '轄區
        TitleTable.Columns.Add(New DataColumn("TPlanID")) '訓練計畫
        TitleTable.Columns.Add(New DataColumn("PlanID"))
        TitleTable.Columns.Add(New DataColumn("RIDValue"))
        TitleTable.Columns.Add(New DataColumn("OCIDValue"))
        TitleTable.Columns.Add(New DataColumn("Range1"))
        TitleTable.Columns.Add(New DataColumn("Range2"))
        TitleTable.Columns.Add(New DataColumn("Range3"))
        TitleTable.Columns.Add(New DataColumn("Range4"))

        dr = TitleTable.NewRow
        TitleTable.Rows.Add(dr)
        '建立TitleTable的欄位---------------------End

        MenuTable.Visible = True
        Button4.Disabled = False
        Button5.Disabled = False

        ShowDataTable.Rows.Clear()
        ShowDataTable2.Rows.Clear()
        ShowDataTable3.Rows.Clear()

        Dim str_DefaultDate1 As String
        Dim str_NowDate1 As String
        str_DefaultDate1 = Common.FormatDate(CDate("2005/1/1"))
        str_NowDate1 = Common.FormatDate(Date.Today)

        SearchStr = ""
        If OCIDValue.Value = "" Then
            If STDate1.Text <> "" Then
                SearchStr += " and cc.STDate>= " & TIMS.To_date(STDate1.Text) & vbCrLf
                dr("STDate1") = Common.FormatDate(STDate1.Text)
            Else
                dr("STDate1") = str_DefaultDate1
            End If
            If STDate2.Text <> "" Then
                SearchStr += " and cc.STDate<= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
                dr("STDate2") = Common.FormatDate(STDate2.Text)
            Else
                dr("STDate2") = str_NowDate1
            End If
            If FTDate1.Text <> "" Then
                SearchStr += " and cc.FTDate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
                dr("FTDate1") = Common.FormatDate(FTDate1.Text)
            Else
                dr("FTDate1") = str_DefaultDate1
            End If
            If FTDate2.Text <> "" Then
                SearchStr += " and cc.FTDate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
                dr("FTDate2") = Common.FormatDate(FTDate2.Text) 'FTDate2.Text
            Else
                dr("FTDate2") = str_NowDate1
            End If
        Else
            dr("STDate1") = str_DefaultDate1
            dr("STDate2") = str_NowDate1
            dr("FTDate1") = str_DefaultDate1
            dr("FTDate2") = str_NowDate1
        End If

        If DistID.SelectedIndex <> 0 Then
            SearchStr += " and ip.DistID='" & DistID.SelectedValue & "'" & vbCrLf
            dr("DistID") = DistID.SelectedValue
        End If
        If TPlanID.SelectedIndex <> 0 Then
            SearchStr += " and ip.TPlanID='" & TPlanID.SelectedValue & "'" & vbCrLf
            dr("TPlanID") = TPlanID.SelectedValue
        End If
        If PlanID.Value <> "" Then
            'SearchStr += " and PlanID='" & PlanID.Value & "'"
            SearchStr += " and ip.PlanID='" & PlanID.Value & "'" & vbCrLf
            dr("PlanID") = PlanID.Value
        End If
        If RIDValue.Value <> "" Then
            'SearchStr += " and RID='" & RIDValue.Value & "'"
            SearchStr += " and cc.RID='" & RIDValue.Value & "'" & vbCrLf
            dr("RIDValue") = RIDValue.Value
        End If
        If OCIDValue.Value <> "" Then
            'SearchStr += " and OCID='" & OCIDValue.Value & "'"
            SearchStr += " and cc.OCID='" & OCIDValue.Value & "'" & vbCrLf
            dr("OCIDValue") = OCIDValue.Value
        End If
        '排除計畫
        '06:在職進修訓練 
        '07:接受企業委託訓練
        '15:學習券
        SearchStr += " and ip.TPlanID Not IN ('06','07','15')" & vbCrLf
        '不列入就業率統計 join cc
        SearchStr += " and NOT EXISTS (SELECT 'x' FROM Class_NoWorkRate cnw WHERE cnw.OCID=cc.OCID)" & vbCrLf

        dr("Range1") = Range1.Text
        dr("Range2") = Range2.Text
        dr("Range3") = Range3.Text
        dr("Range4") = Range4.Text

        Session("TitleTable") = TitleTable
        Session("SearchStr") = SearchStr 'CreateData(sql)

        Dim MyCell As TableCell = Nothing
        Dim MyRow As TableRow = Nothing
        'Dim TotalArray As New ArrayList
        'Dim TotalStudent As Integer = 0         '結訓總人數
        'Dim InWorkEarly As New ArrayList
        'Dim InWorkEarlyTotal As Integer = 0     '提前就業總人數

        sInputText = ""
        sInputText &= "&TPlanID=" & Me.TPlanID.SelectedValue
        sInputText &= "&Range1=" & Range1.Text
        sInputText &= "&Range2=" & Range2.Text
        sInputText &= "&Range3=" & Range3.Text
        sInputText &= "&Range4=" & Range4.Text

        '建立表頭(第一行)
        '建立就業率
        CreateRow(ShowDataTable, MyRow)
        CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , cst_css_TR_04002_TR)
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

        For j As Integer = 1 To 3
            If Not CreateData(3, j, SearchStr, False, Me, objconn, sInputText) Then
                Exit Sub
            End If
            'Call Show_OutText(sOutText)
        Next

        '建立參考資料
        CreateRow(ShowDataTable2, MyRow)
        CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , cst_css_TR_04002_TR)
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


        CreateData(1, 1, SearchStr, False, Me, objconn, sInputText)
        CreateData(2, 1, SearchStr, False, Me, objconn, sInputText)
        CreateData(1, 2, SearchStr, False, Me, objconn, sInputText)
        CreateData(2, 2, SearchStr, False, Me, objconn, sInputText)
        CreateData(4, 2, SearchStr, False, Me, objconn, sInputText)

        CreateData(1, 3, SearchStr, False, Me, objconn, sInputText)
        CreateData(2, 3, SearchStr, False, Me, objconn, sInputText)
        CreateData(4, 3, SearchStr, False, Me, objconn, sInputText)
        'Call Show_OutText(sOutText)

        ShowMode.Value = 1
        ShowDataTable.Style("display") = "inline"
        ShowDataTable2.Style("display") = "none"
        ShowDataTable3.Style("display") = "none"

        Select Case TPlanID.SelectedValue
            Case "23" ''訓用合一要額外顯示資料
                TranId1.Visible = True
                TranId2.Visible = True '訓用合一要額外顯示資料
                TranId3.Visible = True
                TranId4.Visible = False
                TranId5.Visible = False
                TranId6.Visible = False
                TranId7.Visible = False

                CreateRow(ShowDataTable3, MyRow)
                CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , cst_css_TR_04002_TR)
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

                sInputText = ""
                sInputText &= "&TPlanID=" & Me.TPlanID.SelectedValue
                sInputText &= "&Range1=" & Range1.Text
                sInputText &= "&Range2=" & Range2.Text
                sInputText &= "&Range3=" & Range3.Text
                sInputText &= "&Range4=" & Range4.Text

                CreateData(3, 0, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 1, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 2, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 3, SearchStr, True, Me, objconn, sInputText)
                'Call Show_OutText(sOutText)

            Case "34" '與企業辦理合作辦訓
                'MenuTable
                TranId1.Visible = True
                TranId2.Visible = False
                TranId3.Visible = False
                TranId4.Visible = True '與企業合作辦訓專用
                TranId5.Visible = True
                TranId6.Visible = False
                TranId7.Visible = False

                CreateRow(ShowDataTable3, MyRow)
                CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , cst_css_TR_04002_TR)
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

                '因為與企業辦訓為一般計畫,學員就業狀況存在學員就業狀況檔1、學員就業狀況檔1
                CreateData(3, 0, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 1, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 2, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 3, SearchStr, True, Me, objconn, sInputText)
                'Call Show_OutText(sOutText)

            Case "41" '推動營造業事業單位辦理職前培訓計畫
                TranId1.Visible = True
                TranId2.Visible = False
                TranId3.Visible = False
                TranId4.Visible = False
                TranId5.Visible = False
                TranId6.Visible = True '推動營造業事業單位辦理職前培訓計畫
                TranId7.Visible = True

                CreateRow(ShowDataTable3, MyRow)
                CreateCell(MyRow, MyCell, "查核時點\參訓前失業週數", 3, , cst_css_TR_04002_TR)
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

                sInputText = ""
                sInputText &= "&TPlanID=" & Me.TPlanID.SelectedValue
                sInputText &= "&Range1=" & Range1.Text
                sInputText &= "&Range2=" & Range2.Text
                sInputText &= "&Range3=" & Range3.Text
                sInputText &= "&Range4=" & Range4.Text

                CreateData(3, 0, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 1, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 2, SearchStr, True, Me, objconn, sInputText)
                CreateData(3, 3, SearchStr, True, Me, objconn, sInputText)
                'Call Show_OutText(sOutText)

            Case Else
                TranId1.Visible = False
                TranId2.Visible = False
                TranId3.Visible = False
                TranId4.Visible = False
                TranId5.Visible = False
                TranId6.Visible = False
                TranId7.Visible = False

        End Select

    End Sub

#Region "NOUSE"
    ''sql = "SELECT a.SOCID,a.StudStatus,a.RTReasonID,b.FTDate,dbo.NVL(c.IsGetJob,0) as IsGetJob,c.Mode,c.SureItem,d.LostJob "
    ''If Train = True Then
    ''    sql += ",c.BusGNO,a.ActNo "
    ''End If
    ''sql += "FROM "
    ''sql += "(SELECT * FROM Class_StudentsOfClass) a "
    ''If Train = True Then
    ''    sql += "JOIN (SELECT * FROM Class_ClassInfo WHERE IsSuccess='Y' and NotOpen='N' and FTDate<getdate()" & SearchStr & ") b ON a.OCID=b.OCID "
    ''Else
    ''    sql += "JOIN (SELECT * FROM Class_ClassInfo WHERE IsSuccess='Y' and NotOpen='N' and FTDate<'" & FormatDateTime(TIMS.DateAdd1(DateInterval.Month, -3, Now.Date), 2) & "'" & SearchStr & ") b ON a.OCID=b.OCID "
    ''End If
    ''sql += "LEFT JOIN (SELECT * FROM Stud_GetJobState" & i & " WHERE CPoint=" & j & ") c ON a.SOCID=c.SOCID "
    ''sql += "LEFT JOIN Stud_LostJobWeek d ON a.SOCID=d.SOCID "
    ''sql += "JOIN Stud_StudentInfo h ON a.SID=h.SID "

    'sql = "" & vbCrLf
    'sql += " SELECT " & vbCrLf
    'sql += " a.SOCID,a.StudStatus" & vbCrLf
    'sql += " ,a.RTReasonID" & vbCrLf
    'sql += " ,b.FTDate" & vbCrLf
    'sql += " ,dbo.NVL(c.IsGetJob,0) as IsGetJob" & vbCrLf
    'sql += " ,c.Mode,c.SureItem" & vbCrLf
    'sql += " ,d.LostJob, a.WkAheadOfSch,S9. A_hours " & vbCrLf
    'If Train = True Then
    '    sql += ",c.BusGNO,a.ActNo " & vbCrLf
    'End If
    'sql += " FROM (SELECT OCID, FTDate " & vbCrLf
    'sql += " 	FROM Class_ClassInfo cc " & vbCrLf
    'sql += " 	WHERE IsSuccess='Y' and NotOpen='N' " & vbCrLf
    'If Train = True Then
    '    sql += " and FTDate<getdate() " & SearchStr & vbCrLf
    'Else
    '    sql += " and FTDate<'" & FormatDateTime(TIMS.DateAdd1(DateInterval.Month, -3, Now.Date), 2) & "' " & SearchStr & vbCrLf
    'End If
    ''sql += " 	and PlanID Not IN (SELECT PlanID From ID_Plan WHERE TPlanID IN ('06','07','15'))) b" & vbCrLf
    'sql += " ) b" & vbCrLf
    'sql += " JOIN (SELECT OCID, SOCID, SID, StudStatus, RTReasonID, ActNo, WkAheadOfSch " & vbCrLf
    'sql += " 	FROM Class_StudentsOfClass) a ON a.OCID=b.OCID " & vbCrLf
    'sql += " JOIN Stud_StudentInfo h ON a.SID=h.SID " & vbCrLf
    'sql += " LEFT JOIN (SELECT DISTINCT SOCID, IsGetJob, Mode, SureItem, BusGNO " & vbCrLf
    'sql += " 	FROM Stud_GetJobState" & i & " " & vbCrLf
    'sql += " 	WHERE CPoint=" & j & ") c ON a.SOCID=c.SOCID " & vbCrLf
    'sql += " LEFT JOIN Stud_LostJobWeek d ON a.SOCID=d.SOCID " & vbCrLf
    'sql += " LEFT JOIN  ( SELECT  b.socid," & vbCrLf
    'sql += " sum(case when  (convert(float,a.THOURS)/2) >= total_hours then 1 else 0 end) as  A_hours " & vbCrLf  '/*離退訓者,上課時數已達1/2人數*/
    'sql += " FROM (select ocid,ftdate,THOURS from Class_ClassInfo cc where IsSuccess='Y' and NotOpen='N'" & vbCrLf
    'If Train = True Then
    '    sql += " and FTDate<getdate() " & SearchStr & vbCrLf
    'Else
    '    sql += " and FTDate<'" & FormatDateTime(TIMS.DateAdd1(DateInterval.Month, -3, Now.Date), 2) & "' " & SearchStr & vbCrLf
    'End If
    'sql += " ) a" & vbCrLf
    'sql += " join Class_StudentsOfClass b on a.OCID=b.OCID " & vbCrLf
    'sql += " left join( Select a.ocid,b.socid," & vbCrLf
    'sql += " (sum(dbo.NVL(c.hours,0))+ (case when b.studstatus=2  then dbo.fn_get_empdate (a.ocid,(rejecttdate1),a.FTDATE)" & vbCrLf
    'sql += " when b.studstatus=3 then dbo.fn_get_empdate (a.ocid,(rejecttdate2),a.FTDATE)  else 0 end)) as total_hours" & vbCrLf
    'sql += " From (select ocid,ftdate,THOURS from Class_ClassInfo cc where IsSuccess='Y' and NotOpen='N'" & vbCrLf
    'If Train = True Then
    '    sql += " and FTDate<getdate() " & SearchStr & vbCrLf
    'Else
    '    sql += " and FTDate<'" & FormatDateTime(TIMS.DateAdd1(DateInterval.Month, -3, Now.Date), 2) & "' " & SearchStr & vbCrLf
    'End If
    'sql += " ) a join Class_StudentsOfClass b on a.OCID=b.OCID" & vbCrLf
    'sql += " left join  stud_Turnout c on  b.socid=c.socid" & vbCrLf
    'sql += "  WHERE 1=1 and b.StudStatus in (2,3) " & vbCrLf
    'sql += " group by a.ocid,b.socid,a.THOURS,a.ftdate,b.rejecttdate1,b.studstatus,b.rejecttdate2,a.ftdate" & vbCrLf
    'sql += " ) c on a.ocid=c.ocid and b.socid=c.socid" & vbCrLf
    'sql += " WHERE 1=1  group by a.ocid,b.socid" & vbCrLf
    'sql += " ) as S9 on a.socid=S9.socid "


    'sql += " LEFT JOIN (SELECT DISTINCT s3.SOCID, s3.IsGetJob, s3.Mode, s3.SureItem, s3.BusGNO " & vbCrLf
    'sql += " FROM  ID_Plan ip join class_classinfo cc  on cc.Planid = ip.Planid" & vbCrLf
    'sql += " join Class_StudentsOfClass cs on cc.ocid = cs.ocid" & vbCrLf
    'sql += " join Stud_GetJobState" & i & " s3  on cs.socid = s3.socid" & vbCrLf
    'sql += " WHERE CPoint=" & j & SearchStr & vbCrLf
    'sql += " ) c ON a.SOCID= c.SOCID LEFT JOIN Stud_LostJobWeek d ON a.SOCID=d.SOCID " & vbCrLf
#End Region

    '取得顯示Table 
    Public Shared Function GetShowDataTable(ByRef MyPage As Page, ByVal sType As Integer) As System.Web.UI.WebControls.Table
        Dim Rstobj As System.Web.UI.WebControls.Table = Nothing
        Select Case sType
            Case 1
                Rstobj = MyPage.FindControl("ShowDataTable")
            Case 2
                Rstobj = MyPage.FindControl("ShowDataTable2")
            Case 3
                Rstobj = MyPage.FindControl("ShowDataTable3")
        End Select
        If Rstobj Is Nothing Then
            Rstobj = MyPage.FindControl("ShowDataTable")
        End If
        Return Rstobj
    End Function

    '含查詢 SQL
    Public Shared Function CreateData(ByVal i As Integer, _
                                      ByVal j As Integer, _
                                      ByVal SearchStr As String, _
                                      ByVal Train As Boolean, _
                                      ByRef MyPage As Page, _
                                      ByVal tConn As SqlConnection, _
                                      ByRef sInputText As String) As Boolean
        'Public Shared 
        '= False
        'i:'1, 2, 4 '3 橫向設定
        'j:縱向設定

        'SearchStr: 查詢sql
        '含 BusGNO:勞保證字號 ActNo:投保證號 比對使用
        ',TPlanID As System.Web.UI.WebControls.DropDownList
        'CreateData = True
        Dim rst As Boolean = True
        Dim sql As String = ""
        Dim dt As New DataTable
        'Dim dr As DataRow
        Dim MyCell As TableCell = Nothing
        Dim MyRow As TableRow = Nothing

        '2010/12/14 改sql 語法改善執行效率
        sql = "" & vbCrLf
        sql += " select a.socid" & vbCrLf
        sql += " ,a.studstatus" & vbCrLf
        sql += " ,a.rtreasonid" & vbCrLf
        sql += " ,dbo.NVL(c.IsGetJob,0) IsGetJob" & vbCrLf
        sql += " ,CONVERT(varchar, a.FTDate, 111) FTDate" & vbCrLf
        sql += " ,c.Mode_" & vbCrLf
        sql += " ,c.SureItem" & vbCrLf
        sql += " ,d.LostJob" & vbCrLf '失業週數
        'WorkSuppIdent 是否為在職者補助身份
        sql += " ,a.WorkSuppIdent" & vbCrLf
        'WkAheadOfSch 提前就業
        sql += " ,a.WkAheadOfSch" & vbCrLf
        'PUBLICRESCUE 公法救助。
        'CASE WHEN cs.StudStatus not in (2,3) and sg3.PUBLICRESCUE='Y' AND sg3.SOCID IS NOT NULL THEN 1 END 
        sql += " ,c.PUBLICRESCUE" & vbCrLf
        'sql += " ,case when a.THOURS/2 >= a.total_hours then 1 else 0 end A_hours " & vbCrLf
        If Train = True Then
            sql += ",c.BusGNO,a.ActNo " & vbCrLf
        End If
        sql += " FROM (" & vbCrLf
        sql += " select cs.SOCID" & vbCrLf
        sql += " ,cs.StudStatus" & vbCrLf
        sql += " ,cs.RTReasonID" & vbCrLf
        sql += " ,cc.FTDate" & vbCrLf
        sql += " ,cc.THOURS" & vbCrLf
        sql += " ,cs.ActNo" & vbCrLf
        'WorkSuppIdent 是否為在職者補助身份
        sql += " ,cs.WorkSuppIdent" & vbCrLf
        'WkAheadOfSch 提前就業
        sql += " ,cs.WkAheadOfSch" & vbCrLf
        'sql += " ,case when cs.studstatus=2  then dbo.NVL(st.hours,0)+ dbo.fn_get_empdate (cc.ocid,(cs.rejecttdate1),cc.FTDATE)" & vbCrLf
        'sql += "       when cs.studstatus=3 then  dbo.NVL(st.hours,0)+ dbo.fn_get_empdate (cc.ocid,(cs.rejecttdate2),cc.FTDATE)" & vbCrLf
        'sql += "       else 0 end  total_hours" & vbCrLf
        sql += " FROM ID_PLAN ip " & vbCrLf
        sql += " JOIN CLASS_CLASSINFO cc on cc.Planid = ip.Planid" & vbCrLf
        sql += " JOIN CLASS_STUDENTSOFCLASS cs on cc.ocid = cs.ocid" & vbCrLf
        'sql += " left join ( select socid,sum(dbo.NVL(hours,0)) hours from stud_Turnout st1 group by socid) st on cs.socid=st.socid" & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += " AND cc.IsSuccess='Y'" & vbCrLf
        sql += " AND cc.NotOpen='N'" & vbCrLf
        If Train = True Then
            sql += " and cc.FTDate< dbo.TRUNC_DATETIME(getdate())" & vbCrLf
        Else
            sql += " and cc.FTDate< DATEADD(month, -3, dbo.TRUNC_DATETIME(getdate()))" & vbCrLf
        End If
        sql += SearchStr & vbCrLf

        sql += " ) a" & vbCrLf 'Class_StudentsOfClass

        'fix ORA-01747: invalid user.table.column, table.column, or column specification 有欄位定義成 db 關鍵字，轉db後會自動被加 _ (底線)
        'sql += " LEFT JOIN (SELECT DISTINCT SOCID, IsGetJob, Mode_, SureItem, BusGNO " & vbCrLf
        'sql += " 	FROM Stud_GetJobState" & i & " " & vbCrLf
        'sql += " 	WHERE CPoint=" & j & ") c ON a.SOCID=c.SOCID " & vbCrLf
        sql += " LEFT JOIN Stud_GetJobState" & i & " c ON a.SOCID=c.SOCID and c.CPoint=" & j & "" & vbCrLf
        sql += " LEFT JOIN Stud_LostJobWeek d ON a.SOCID=d.SOCID " & vbCrLf

        Try
            Call TIMS.OpenDbConn(tConn)
            Dim da As New SqlDataAdapter
            With da
                .SelectCommand = New SqlCommand(sql, tConn)
                .SelectCommand.CommandTimeout = Cst_CommandTimeout
                .Fill(dt)
            End With
            'If conn.State = ConnectionState.Open Then conn.Close()
        Catch ex As Exception
            Call TIMS.CloseDbConn(tConn)
            rst = False 'CreateData = False
            Common.MessageBox(MyPage, "資料庫效能異常，請重新查詢")
            Common.MessageBox(MyPage, ex.ToString)
            Return rst 'Exit Function
        End Try

        Dim FinPeo As New ArrayList         '結訓人數
        Dim WorkPeo As New ArrayList        '在職者人數
        Dim RejPeo As New ArrayList         '提前就業人數

        Dim NotInWork As New ArrayList      '未就業人數
        Dim TWork As New ArrayList          '不就業人數

        Dim InWorkByPeo As New ArrayList    '就業人數
        Dim InWorkByPeo1 As New ArrayList   '就業人數(人工)-雇主切結
        Dim InWorkByPeo2 As New ArrayList   '就業人數(人工)-學員切結

        Dim InWork As New ArrayList         '系統判定 就業人數
        Dim PUWork As New ArrayList         '公法上救助對象之就業人數

        'Dim BeforeInWork As New ArrayList   '訓前一個月已加保

        Dim sBorderWidth As String = TIMS.GetMyValue(sInputText, "BorderWidth")
        Dim iBorderWidth As Integer = 0
        If sBorderWidth <> "" Then iBorderWidth = sBorderWidth

        Dim sTPlanID As String = TIMS.GetMyValue(sInputText, "TPlanID")
        Dim sRange1 As String = TIMS.GetMyValue(sInputText, "Range1")
        Dim sRange2 As String = TIMS.GetMyValue(sInputText, "Range2")
        Dim sRange3 As String = TIMS.GetMyValue(sInputText, "Range3")
        Dim sRange4 As String = TIMS.GetMyValue(sInputText, "Range4")

        Dim LostRange1 As String = "LostJob<=" & sRange1 & " and LostJob>=0"
        Dim LostRange2 As String = "LostJob>=" & sRange2 & " and LostJob<=" & sRange3 & ""
        Dim LostRange3 As String = "LostJob>=" & sRange4 & ""

        'k:結訓人數(人),提前就業人數(人),未就業,不就業,人工判定-就業人數,人工判定-就業人數-雇主切結 
        ',人工判定-就業人數-學員切結 ,系統判定-就業人數 ,就業率1,就業率2 ,11:就業率3
        Const cst_結訓p As Integer = 1
        Const cst_提前就業p As Integer = 2
        Const cst_未就業p As Integer = 3
        Const cst_不就業p As Integer = 4
        Const cst_人判就業p As Integer = 5
        Const cst_人判就業雇主p As Integer = 6
        Const cst_人判就業學員p As Integer = 7
        Const cst_系判就業p As Integer = 8
        Const cst_公法救助就業p As Integer = 9 '公法上救助對象之就業人數
        Const cst_就業率1 As Integer = 10
        Const cst_就業率2 As Integer = 11
        Const cst_就業率3 As Integer = 12
        Const Cst_max列數 As Integer = 12

        For k As Integer = 1 To Cst_max列數
            'For k As Integer = 1 To 8
            Select Case i
                Case 1, 2, 4
                    CreateRow(GetShowDataTable(MyPage, 2), MyRow)
                Case 3
                    If Train = True Then
                        If k <> 2 Then
                            CreateRow(GetShowDataTable(MyPage, 3), MyRow)
                        End If
                    Else
                        CreateRow(GetShowDataTable(MyPage, 1), MyRow)
                    End If
            End Select

            Select Case k
                Case cst_結訓p
                    If i = 3 Then
                        If Train = True Then
                            If j = 0 Then
                                Select Case sTPlanID
                                    Case "23"
                                        CreateCell(MyRow, MyCell, "訓用合一專用", 1, (Cst_max列數 - 1) * 4, cst_css_TR_04002_TR, iBorderWidth)
                                    Case "34"
                                        CreateCell(MyRow, MyCell, "與企業合作辦訓專用", 1, (Cst_max列數 - 1) * 4, cst_css_TR_04002_TR, iBorderWidth)
                                    Case "41"
                                        CreateCell(MyRow, MyCell, "推動營造業事業單位辦理職前培訓專用", 1, (Cst_max列數 - 1) * 4, cst_css_TR_04002_TR, iBorderWidth)
                                End Select

                                MyCell.Width = Unit.Pixel(20)
                                CreateCell(MyRow, MyCell, "結訓時就業人數", 1, Cst_max列數 - 1, cst_css_TR_04002_TR, iBorderWidth)
                            Else
                                CreateCell(MyRow, MyCell, "訓後" & j * 3 & "個月內<BR>就業人數", 1, Cst_max列數 - 1, cst_css_TR_04002_TR, iBorderWidth)
                            End If
                            MyCell.Width = Unit.Pixel(90)
                        Else
                            If j = 1 Then
                                CreateCell(MyRow, MyCell, "持續追蹤", 1, Cst_max列數 * 3, cst_css_TR_04002_TR, iBorderWidth)
                                MyCell.Width = Unit.Pixel(20)
                            End If
                            CreateCell(MyRow, MyCell, "訓後" & j * 3 & "個月內<BR>就業人數", 1, Cst_max列數, cst_css_TR_04002_TR, iBorderWidth)
                            MyCell.Width = Unit.Pixel(90)
                        End If
                    Else
                        Select Case j
                            Case 1
                                If i = 1 Then
                                    CreateCell(MyRow, MyCell, "訓後三個月", 1, Cst_max列數 * 2)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                    MyCell.Width = Unit.Pixel(20)
                                    CreateCell(MyRow, MyCell, "累計14天加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                ElseIf i = 2 Then
                                    CreateCell(MyRow, MyCell, "連續14天加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                End If
                            Case 2
                                If i = 1 Then
                                    CreateCell(MyRow, MyCell, "訓後六個月", 1, Cst_max列數 * 3)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                    MyCell.Width = Unit.Pixel(20)
                                    CreateCell(MyRow, MyCell, "累計90天加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                ElseIf i = 2 Then
                                    CreateCell(MyRow, MyCell, "連續90天加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                ElseIf i = 4 Then
                                    CreateCell(MyRow, MyCell, "曾加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                End If
                            Case 3
                                If i = 1 Then
                                    CreateCell(MyRow, MyCell, "訓後九個月", 1, Cst_max列數 * 3)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                    MyCell.Width = Unit.Pixel(20)
                                    CreateCell(MyRow, MyCell, "累計180天加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                ElseIf i = 2 Then
                                    CreateCell(MyRow, MyCell, "連續180天加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                ElseIf i = 4 Then
                                    CreateCell(MyRow, MyCell, "曾加保者", 1, Cst_max列數)
                                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                                    MyCell.CssClass = cst_css_TR_04002_TR
                                End If
                        End Select
                    End If

                    CreateCell(MyRow, MyCell, "結訓人數(人)")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    MyCell.ToolTip = "結訓後" & j * 3 & "個月之班級，且超過結訓日班級，學員非離訓、退訓狀態的人數"
                    MyCell.Style("CURSOR") = "help"
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    '結訓人數
                    'FinPeo.Add(dt.Select("LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)
                    FinPeo.Add(New DataView(dt, "LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    FinPeo.Add(New DataView(dt, LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    FinPeo.Add(New DataView(dt, LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    FinPeo.Add(New DataView(dt, LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    FinPeo.Add(New DataView(dt, "FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                    '在職者人數
                    WorkPeo.Add(New DataView(dt, "LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3) and WorkSuppIdent='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                    WorkPeo.Add(New DataView(dt, LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3) and WorkSuppIdent='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                    WorkPeo.Add(New DataView(dt, LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3) and WorkSuppIdent='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                    WorkPeo.Add(New DataView(dt, LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3) and WorkSuppIdent='Y'", Nothing, DataViewRowState.CurrentRows).Count)
                    WorkPeo.Add(New DataView(dt, "FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3) and WorkSuppIdent='Y'", Nothing, DataViewRowState.CurrentRows).Count)

                    For l As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, FinPeo(l))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_提前就業p '提前就業人數
                    If Train = True Then
                        RejPeo.Add(0)
                        'RejPeo.Add(0)
                        RejPeo.Add(0)
                        RejPeo.Add(0)
                        RejPeo.Add(0)
                        RejPeo.Add(0)
                    Else
                        CreateCell(MyRow, MyCell, "提前就業人數(人)")
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Width = Unit.Pixel(90)
                        MyCell.ToolTip = "學員離訓、退訓的原因為提前就業的人數的人數"
                        MyCell.Style("CURSOR") = "help"
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If

                        '提前就業人數
                        'RejPeo.Add(dt.Select("LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02'").Length)
                        RejPeo.Add(New DataView(dt, "LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02' and WkAheadOfSch = 'Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        RejPeo.Add(New DataView(dt, LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02' and WkAheadOfSch = 'Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        RejPeo.Add(New DataView(dt, LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02' and WkAheadOfSch = 'Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        RejPeo.Add(New DataView(dt, LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02' and WkAheadOfSch = 'Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        RejPeo.Add(New DataView(dt, "FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02' and WkAheadOfSch = 'Y'", Nothing, DataViewRowState.CurrentRows).Count)
                        'RejPeo.Add(dt.Select("LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02'").Length)
                        'RejPeo.Add(dt.Select(LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02'").Length)
                        'RejPeo.Add(dt.Select(LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02'").Length)
                        'RejPeo.Add(dt.Select(LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02'").Length)
                        'RejPeo.Add(dt.Select("FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus IN (2,3) and RTReasonID='02'").Length)

                        For l As Integer = 0 To 4
                            CreateCell(MyRow, MyCell, RejPeo(l))
                            MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                            If i = 2 Then
                                MyCell.CssClass = cst_css_TR_04002_TD2
                            End If
                        Next
                    End If

                Case cst_未就業p '未就業
                    CreateCell(MyRow, MyCell, "未就業")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    If Train = True Then
                        MyCell.ToolTip = "學員尚未就業之人數的人數,或者投保單位並非指定的訓練單位之人數"
                    Else
                        MyCell.ToolTip = "學員尚未就業之人數的人數"
                    End If
                    MyCell.Style("CURSOR") = "help"
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    '未就業人數
                    If Train = True Then
                        '判斷學員就業狀況檔的勞保證字號是否等於班級學員檔的保險證號
                        NotInWork.Add(New DataView(dt, "(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and " & LostRange2 & " and LostJob<=" & sRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "(IsGetJob='0' or (IsGetJob='1' and (BusGNO<>ActNo or ActNo IS NULL))) and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    Else
                        'NotInWork.Add(dt.Select("IsGetJob='0' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)
                        NotInWork.Add(New DataView(dt, "IsGetJob='0' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "IsGetJob='0' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "IsGetJob='0' and " & LostRange2 & " and LostJob<=" & sRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "IsGetJob='0' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        NotInWork.Add(New DataView(dt, "IsGetJob='0' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    End If

                    For l As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, NotInWork(l))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_不就業p '不就業
                    CreateCell(MyRow, MyCell, "不就業")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    MyCell.ToolTip = "學員選擇不願就業的人數(可能升學、出國等等原因)"
                    MyCell.Style("CURSOR") = "help"
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    '不就業人數
                    'TWork.Add(dt.Select("IsGetJob='2' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)
                    TWork.Add(New DataView(dt, "IsGetJob='2' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    TWork.Add(New DataView(dt, "IsGetJob='2' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    TWork.Add(New DataView(dt, "IsGetJob='2' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    TWork.Add(New DataView(dt, "IsGetJob='2' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    TWork.Add(New DataView(dt, "IsGetJob='2' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                    For l As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, TWork(l))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_人判就業p '人工判定-就業人數
                    CreateCell(MyRow, MyCell, "人工判定<BR>就業人數")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    If Train = True Then
                        MyCell.ToolTip = "人工判定判定學員就業,且投保單位為指定機構之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    Else
                        MyCell.ToolTip = "人工判定判定學員就業之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    End If
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    If Train = True Then
                        'InWorkByPeo.Add(dt.Select("Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)

                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    Else
                        'InWorkByPeo.Add(dt.Select("Mode=2 and IsGetJob='1' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)

                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    End If

                    For l As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, InWorkByPeo(l))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_人判就業雇主p '人工判定-就業人數-雇主切結
                    CreateCell(MyRow, MyCell, "人工判定<BR>就業人數-雇主切結")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    If Train = True Then
                        MyCell.ToolTip = "人工判定判定學員就業,且投保單位為指定機構之人數,切結對象為雇主切結"
                        'MyCell.ToolTip = "人工判定判定學員就業,且投保單位為指定機構之人數,切結對象為雇主切結" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    Else
                        MyCell.ToolTip = "人工判定判定學員就業之人數,切結對象為雇主切結"
                        'MyCell.ToolTip = "人工判定判定學員就業之人數,切結對象為雇主切結" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    End If
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    '就業人數(人工)-雇主切結
                    If Train = True Then
                        'InWorkByPeo.Add(dt.Select("Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)

                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    Else
                        'InWorkByPeo.Add(dt.Select("Mode=2 and IsGetJob='1' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)

                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo1.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '1' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '1' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '1' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '1' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo1.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '1' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    End If

                    For l As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, InWorkByPeo1(l))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_人判就業學員p '人工判定-就業人數-學員切結
                    CreateCell(MyRow, MyCell, "人工判定<BR>就業人數-學員切結")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    If Train = True Then
                        MyCell.ToolTip = "人工判定判定學員就業,且投保單位為指定機構之人數,切結對象為學員切結"
                        'MyCell.ToolTip = "人工判定判定學員就業,且投保單位為指定機構之人數,切結對象為學員切結" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    Else
                        MyCell.ToolTip = "人工判定判定學員就業之人數,切結對象為學員切結"
                        'MyCell.ToolTip = "人工判定判定學員就業之人數,切結對象為學員切結" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    End If
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    '就業人數(人工)-學員切結
                    If Train = True Then
                        'InWorkByPeo.Add(dt.Select("Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)

                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and BusGNO=ActNo and ActNo IS Not NULL and IsGetJob='1' and SureItem = '2' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    Else
                        'InWorkByPeo.Add(dt.Select("Mode=2 and IsGetJob='1' and LostJob=-1 and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)").Length)

                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '2' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '2' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '2' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '2' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        'InWorkByPeo2.Add(New DataView(dt, "Mode=2 and IsGetJob='1' and SureItem = '2' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)

                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '2' and LostJob IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '2' and " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '2' and " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '2' and " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        InWorkByPeo2.Add(New DataView(dt, "Mode_=2 and IsGetJob='1' and SureItem = '2' and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' and StudStatus Not IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    End If

                    For l As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, InWorkByPeo2(l))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_系判就業p '系統判定-就業人數
                    CreateCell(MyRow, MyCell, "系統判定<BR>就業人數")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    If Train = True Then
                        MyCell.ToolTip = "系統自動判定判定學員就業,且投保單位為指定機構之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    Else
                        MyCell.ToolTip = "系統自動判定判定學員就業之人數" & vbCrLf & vbCrLf & "[公式]" & vbCrLf & "人工判定就業人數+就業人數=結訓人數-未就業-不就業"
                    End If
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    '就業人數
                    If Train = True Then
                        For il As Integer = 0 To 4
                            InWork.Add(FinPeo(il) - NotInWork(il) - TWork(il) - InWorkByPeo(il))
                        Next
                    Else
                        For il As Integer = 0 To 4
                            'InWork.Add(FinPeo(l) + RejPeo(l) - NotInWork(l) - TWork(l) - InWorkByPeo(l))
                            InWork.Add(FinPeo(il) - NotInWork(il) - TWork(il) - InWorkByPeo(il))
                        Next
                    End If

                    For il As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, InWork(il))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_公法救助就業p '公法救助對象就業人數 
                    CreateCell(MyRow, MyCell, "公法救助對象<BR>就業人數")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    If Train = True Then
                        MyCell.ToolTip = "公法上救助對象之就業人數,且投保單位為指定機構之人數" & vbCrLf
                    Else
                        MyCell.ToolTip = "公法上救助對象之就業人數" & vbCrLf
                    End If
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If
                    If Train = True Then
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND LOSTJOB IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    Else
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND LOSTJOB IS NULL and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND " & LostRange1 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND " & LostRange2 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND " & LostRange3 & " and FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                        PUWork.Add(New DataView(dt, "PUBLICRESCUE='Y' AND IsGetJob='1' AND FTDate<'" & TIMS.DateAdd1(DateInterval.Month, -j * 3, Now.Date) & "' AND STUDSTATUS NOT IN (2,3)", Nothing, DataViewRowState.CurrentRows).Count)
                    End If

                    For il As Integer = 0 To 4
                        CreateCell(MyRow, MyCell, PUWork(il))
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_就業率1 '就業率1
                    CreateCell(MyRow, MyCell, "就業率1")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    'If Train = True Then
                    '    MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數-不就業+提前就業人數)"
                    'Else
                    '    MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數-不就業+提前就業人數)"
                    'End If
                    MyCell.ToolTip = cst_sExp1 '"[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數-不就業+提前就業人數)"
                    MyCell.Width = Unit.Pixel(90)

                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    For m As Integer = 0 To 4
                        If FinPeo(m) + RejPeo(m) = 0 Then
                            CreateCell(MyRow, MyCell, "0%")
                            MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        Else
                            If Train = True Then
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) - TWork(m) + RejPeo(m))) * 100, 2) & "%")
                            Else
                                'CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m)) / (FinPeo(m) + RejPeo(m) - TWork(m))) * 100, 2) & "%")
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) - TWork(m) + RejPeo(m))) * 100, 2) & "%")
                            End If
                            MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        End If

                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next
                Case cst_就業率2 '就業率2
                    CreateCell(MyRow, MyCell, "就業率2")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)
                    'If Train = True Then
                    '    MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數)"
                    'Else
                    '    MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數)"
                    'End If
                    MyCell.ToolTip = cst_sExp2 '"[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數)"
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If

                    For m As Integer = 0 To 4
                        If FinPeo(m) + RejPeo(m) = 0 Then
                            CreateCell(MyRow, MyCell, "0%")
                        Else
                            If Train = True Then
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) + RejPeo(m))) * 100, 2) & "%")
                            Else
                                'CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m)) / (FinPeo(m) + RejPeo(m))) * 100, 2) & "%")
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) + RejPeo(m))) * 100, 2) & "%")
                            End If
                        End If
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next

                Case cst_就業率3 '就業率3
                    CreateCell(MyRow, MyCell, "就業率3")
                    MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                    MyCell.Width = Unit.Pixel(90)

                    MyCell.ToolTip = cst_sExp3 '"[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數)/(結訓人數+提前就業人數-在職者人數)"
                    'If Train = True Then
                    '    MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數-在職者人數)/(結訓人數+提前就業人數-在職者人數)"
                    'Else
                    '    MyCell.ToolTip = "[公式]" & vbCrLf & "(人工判定就業人數+系統判定就業人數+提前就業人數-在職者人數)/(結訓人數+提前就業人數-在職者人數)"
                    'End If
                    MyCell.Style("CURSOR") = "help"
                    MyCell.Font.Bold = True
                    If i = 2 Then
                        MyCell.CssClass = cst_css_TR_04002_TD2
                    End If


                    For m As Integer = 0 To 4
                        If FinPeo(m) + RejPeo(m) = 0 Then
                            CreateCell(MyRow, MyCell, "0%")
                        Else
                            If Train = True Then
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) + RejPeo(m) - WorkPeo(m))) * 100, 2) & "%")
                            Else
                                'CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m)) / (FinPeo(m) + RejPeo(m))) * 100, 2) & "%")
                                CreateCell(MyRow, MyCell, Math.Round(CSng((InWork(m) + InWorkByPeo(m) + RejPeo(m)) / (FinPeo(m) + RejPeo(m) - WorkPeo(m))) * 100, 2) & "%")
                            End If
                        End If
                        MyCell.BorderWidth = Unit.Pixel(iBorderWidth)
                        MyCell.Font.Bold = True
                        If i = 2 Then
                            MyCell.CssClass = cst_css_TR_04002_TD2
                        End If
                    Next
            End Select
        Next

        dt.Dispose()
        Return rst
    End Function

    '建立TableCell
    Public Shared Sub CreateCell(ByRef MyRow As TableRow, ByRef MyCell As TableCell, _
        Optional ByVal MyText As String = "", Optional ByVal ColumnSpan As Integer = 1, Optional ByVal RowSpan As Integer = 1, _
        Optional ByVal CssClass As String = "TR_04002_TD", Optional ByVal BorderWidth As Integer = 0)
        MyCell = New TableCell
        MyCell.Text = MyText
        MyCell.ColumnSpan = ColumnSpan
        MyCell.RowSpan = RowSpan
        MyCell.CssClass = CssClass
        If BorderWidth > 0 Then MyCell.BorderWidth = Unit.Pixel(BorderWidth) '有線()
        MyRow.Cells.Add(MyCell)
    End Sub

    '建立TableRow
    Public Shared Sub CreateRow(ByRef MyTable As Table, ByRef MyRow As TableRow)
        MyRow = New TableRow
        MyRow.CssClass = cst_css_TR_04002_TR
        If Not MyTable Is Nothing Then
            MyTable.Rows.Add(MyRow)
        End If
    End Sub

    '查詢 SQL
    Sub Search1()
        If RIDValue.Value = "" Then
            OCID.Items.Clear()
        Else
            OCID.Items.Clear()

            'Dim da As SqlDataAdapter = TIMS.GetOneDA(objconn)
            Dim sql As String = ""
            sql = "" & vbCrLf
            sql += " SELECT cc.OCID" & vbCrLf
            sql += " ,cc.ClassCName" & vbCrLf
            sql += " ,cc.CyclType" & vbCrLf
            sql += " ,cc.STDate" & vbCrLf
            sql += " ,cc.FTDate " & vbCrLf
            sql += " FROM Class_ClassInfo cc" & vbCrLf
            sql += " JOIN ID_Plan ip on ip.planid =cc.planid " & vbCrLf
            sql += " WHERE 1=1" & vbCrLf
            sql += " and ip.TPlanID= @TPlanID" & vbCrLf
            sql += " and ip.distid= @distid" & vbCrLf
            sql += " and ip.PlanID= @PlanID " & vbCrLf
            sql += " and cc.RID= @RID " & vbCrLf
            '不列入就業率統計
            sql += " AND NOT EXISTS (SELECT 'x' FROM Class_NoWorkRate cnw WHERE cnw.OCID=cc.OCID)" & vbCrLf
            '班別關鍵字
            ViewState("sInputCLSNAME") = TIMS.ClearSQM(ViewState("sInputCLSNAME"))
            If Me.ViewState("sInputCLSNAME") <> "" Then
                Dim sInputCLSNAME As String = UCase(Me.ViewState("sInputCLSNAME"))
                'Sql += " AND cc.ClassCname like N'%" & Me.ViewState("sInputCLSNAME") & "%'" & vbCrLf
                'regexp_like(classcname,'java','i')
                'Sql += " AND regexp_like(cc.ClassCname,N'" & Me.ViewState("sInputCLSNAME") & "','i')" & vbCrLf
                sql += " AND UPPER(cc.ClassCname) like N'%" & sInputCLSNAME & "%' " & vbCrLf
            End If
            Dim sCmd As New SqlCommand(sql, objconn)

            TIMS.OpenDbConn(objconn)
            Dim dt As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID.SelectedValue
                .Parameters.Add("distid", SqlDbType.VarChar).Value = DistID.SelectedValue
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = PlanID.Value
                .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
                dt.Load(.ExecuteReader())
            End With
            'da.SelectCommand.Parameters.Clear()
            'da.SelectCommand.Parameters.Add("TPlanID", SqlDbType.VarChar).Value = TPlanID.SelectedValue
            'da.SelectCommand.Parameters.Add("distid", SqlDbType.VarChar).Value = DistID.SelectedValue
            'da.SelectCommand.Parameters.Add("PlanID", SqlDbType.VarChar).Value = PlanID.Value
            'da.SelectCommand.Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
            'TIMS.Fill(sql, da, dt)
            dt.DefaultView.Sort = "STDate"
            If dt.Rows.Count = 0 Then
                OCID.Items.Insert(0, New ListItem("此計畫、機構底下沒有任何班級", ""))
            Else
                For Each dr As DataRow In dt.DefaultView.Table.Rows
                    Dim ClassName As String = dr("ClassCName").ToString
                    If Int(dr("CyclType")) <> 0 Then
                        ClassName += "第" & Int(dr("CyclType")) & "期"
                    End If
                    If dr("STDate").ToString <> "" And dr("FTDate").ToString <> "" Then
                        ClassName += "(" & FormatDateTime(dr("STDate"), 2) & "~" & FormatDateTime(dr("FTDate"), 2) & ")"
                    End If

                    OCID.Items.Add(New ListItem(ClassName, dr("OCID")))
                Next
                OCID.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End If
        End If
    End Sub

    '查詢 (依班別)
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Call Search1()
    End Sub

    'SERVER端 檢查
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If DistID.SelectedValue = "" Then
            'Errmsg += "請選擇 轄區中心" & vbCrLf
            Errmsg += "請選擇 轄區分署" & vbCrLf
        End If
        If TPlanID.SelectedValue = "" Then
            Errmsg += "請選擇 訓練計畫" & vbCrLf
        End If
        If RIDValue.Value = "" OrElse PlanID.Value = "" Then
            Errmsg += "請選擇 訓練機構" & vbCrLf
        End If
        If Me.ViewState("sInputCLSNAME") = "" Then
            Errmsg += "請輸入 班別關鍵字" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢2 (依班別關鍵字)
    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        Dim rtnPath As String = Request.FilePath
        Me.ViewState("sInputCLSNAME") = ""
        If classname.Text.Trim <> "" Then
            classname.Text = classname.Text.Trim
            Me.ViewState("sInputCLSNAME") = classname.Text
            If TIMS.CheckInput(Me.ViewState("sInputCLSNAME")) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2, rtnPath)
                Exit Sub
            End If
        End If

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call Search1()
    End Sub

End Class

