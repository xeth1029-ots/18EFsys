Public Class CP_02_002_R_Prt
    Inherits AuthBasePage

    Dim dataDt As DataTable = Nothing
    Dim ClassDataDt As DataTable = Nothing

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Dim start_date As String = ""
        Dim end_date As String = ""
        Dim OC_TYPE As String = ""
        Dim TPlan As String = ""
        Dim SYMD As String = ""
        Dim EYMD As String = ""

        start_date = Request("start_date")
        end_date = Request("end_date")
        OC_TYPE = Request("OC_TYPE")
        TPlan = Request("TPlan")
        SYMD = Request("SYMD")
        EYMD = Request("EYMD")

        PrintDiv(start_date, end_date, OC_TYPE, TPlan, SYMD, EYMD)

    End Sub

    Private Sub PrintDiv(ByVal start_date As String, ByVal end_date As String, ByVal OC_TYPE As String, ByVal TPlan As String, ByVal SYMD As String, ByVal EYMD As String)

        Dim station_Dt As DataTable = Nothing   '分署
        Dim class_Dt As DataTable = Nothing '所有班級
        'Dim tmpDt_return As DataRow()    '分署(中心)的select結果
        'Dim tmpDt_return1 As DataRow()   '班次的select結果
        Dim resDt As DataTable = Nothing
        Dim strSQL_Station As String = ""
        Dim strSQL As String = ""
        Dim tmpDT As New DataTable
        'Dim tmpDR As DataRow
        'Dim tmpObj As Object

        Dim TT As DataTable = Nothing
        Dim intT1, intT2 As Integer
        Dim intTmp As Integer = 0
        Dim PageCount As Integer = 0
        Dim rsCursor As Integer = 0
        Dim intPageRecord As Integer = 23 '每頁列印幾筆(從0開始)

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl
        Dim strStyle As String = "font-size:12pt;font-family:DFKai-SB"
        Dim strAlign As String = ""

        'Select 全部資料 
        strSQL = "Select I.StatID,I.StatName,K.Trainice, "
        strSQL += "Count( case when T2.Sex='1' then 1 end ) as Sex_M, "
        strSQL += "Count( case when T2.Sex='2' then 1 end ) as Sex_F, "
        strSQL += "Count( case when (T2.age_desc<15) then 1 end ) as Age_0015,"
        strSQL += "Count( case when T2.age_desc>=15 and T2.age_desc <=24 then 1 end ) as Age_1524,"
        strSQL += "Count( case when T2.age_desc>=25 and T2.age_desc <=34 then 1 end ) as Age_2534,"
        strSQL += "Count( case when T2.age_desc>=35 and T2.age_desc <=44 then 1 end ) as Age_3544,"
        strSQL += "Count( case when T2.age_desc>=45 and T2.age_desc <=54 then 1 end ) as Age_4554,"
        strSQL += "Count( case when T2.age_desc>=55 then 1 end ) as Age_55,"
        strSQL += "Count( case when T2.DegreeID='01' then 1 end ) as DegreeID_01,"
        strSQL += "Count( case when T2.DegreeID='02' then 1 end ) as DegreeID_02,"
        strSQL += "Count( case when T2.DegreeID='03' then 1 end ) as DegreeID_03,"
        strSQL += "Count( case when T2.DegreeID='04' then 1 end ) as DegreeID_04,"
        strSQL += "Count( case when T2.DegreeID='05' then 1 end ) as DegreeID_05,"
        strSQL += "Count( case when T2.DegreeID='06' then 1 end ) as DegreeID_06,"
        strSQL += "Count( case when T2.Trainice='1' then 1 end ) as Trainice_1,"
        strSQL += "Count( case when T2.Trainice='2' then 1 end ) as Trainice_2 "
        strSQL += "From ID_StatistDist I  cross join (select 1 as Trainice  union select 2 ) K "
        strSQL += "Left Join "
        strSQL += "( Select  B.UnitCode,B.Trainice,B.ResultDate,C.Sex,(CONVERT(numeric, datepart(year,getdate()))- C.BirthYear) as age_desc,C.DegreeID "
        strSQL += "From  Stud_DataLid  B  "
        strSQL += " Left Join Stud_ResultStudData C on B.DLID=C.DLID "
        strSQL += "Where 1=1 "

        If start_date <> "" Then
            strSQL += "And B.ResultDate >=to_date('" + start_date + "','YYYY/MM/DD') "
        End If
        If end_date <> "" Then
            strSQL += "And B.ResultDate <=to_date('" + end_date + " 23:59:59','YYYY/MM/DD hh24:mi:ss') "
        End If
        If TPlan <> "" Then
            strSQL += "And B.TPlanID in (" + TPlan + ") "
        End If
        strSQL += ") T2  on I.StatID=T2.UnitCode and K.Trainice=T2.Trainice "
        strSQL += "Where 1=1 "
        If OC_TYPE <> "" Then
            strSQL += " and I.Type='" + OC_TYPE + "'"
        End If
        strSQL += "Group by I.StatID,I.StatName,K.Trainice "
        strSQL += "order by I.StatID,K.Trainice "

        dataDt = DbAccess.GetDataTable(strSQL, objconn)

        station_Dt = dataDt.DefaultView.ToTable(True, "StatName", "StatID")

        strSQL = "Select I.StatID,I.StatName,T2.Trainice "
        strSQL += "From ID_StatistDist I "
        strSQL += " Join "
        strSQL += "( Select  B.UnitCode,B.Trainice "
        strSQL += "From  Stud_DataLid B "
        strSQL += "Where 1=1 "
        If start_date <> "" Then
            strSQL += "And B.ResultDate >=to_date('" + start_date + "','YYYY/MM/DD') "
        End If
        If end_date <> "" Then
            strSQL += "And B.ResultDate <=to_date('" + end_date + "','YYYY/MM/DD') "
        End If
        If TPlan <> "" Then
            strSQL += "And B.TPlanID in (" + TPlan + ") "
        End If
        strSQL += ") T2  on I.StatID=T2.UnitCode "
        strSQL += "Where 1=1 "
        If OC_TYPE <> "" Then
            strSQL += " and I.Type='" + OC_TYPE + "'"
        End If
        strSQL += "order by I.StatID "

        ClassDataDt = DbAccess.GetDataTable(strSQL, objconn)

        '結果報表
        tmpDT.Columns.Add(New DataColumn("StatID"))    '站別
        tmpDT.Columns.Add(New DataColumn("Item"))    '項目別
        tmpDT.Columns.Add(New DataColumn("ClassNum"))    '班次
        tmpDT.Columns.Add(New DataColumn("Total"))    'Total
        tmpDT.Columns.Add(New DataColumn("M"))    '男生
        tmpDT.Columns.Add(New DataColumn("F"))    '男生
        tmpDT.Columns.Add(New DataColumn("Age_0015"))    '未滿15歲
        tmpDT.Columns.Add(New DataColumn("Age_1524"))    '15~24
        tmpDT.Columns.Add(New DataColumn("Age_2534"))    '25~34
        tmpDT.Columns.Add(New DataColumn("Age_3544"))    '35~44
        tmpDT.Columns.Add(New DataColumn("Age_4554"))    '45~54
        tmpDT.Columns.Add(New DataColumn("Age_55"))    '55~
        tmpDT.Columns.Add(New DataColumn("DegreeID_01"))    '國中
        tmpDT.Columns.Add(New DataColumn("DegreeID_02"))    '高中
        tmpDT.Columns.Add(New DataColumn("DegreeID_03"))    '專科
        tmpDT.Columns.Add(New DataColumn("DegreeID_04"))    '大學
        tmpDT.Columns.Add(New DataColumn("DegreeID_05"))    '碩士
        tmpDT.Columns.Add(New DataColumn("DegreeID_06"))    '博士
        'Column Header
        'tmpDR = tmpDT.NewRow
        'tmpDT.Rows.Add(tmpDR)
        'tmpDR("Item") = "項目別"
        'tmpDR("ClassNum") = "班次"
        'tmpDR("Total") = "人數"
        'tmpDR("M") = "男"
        'tmpDR("F") = "女"
        'tmpDR("Age_0015") = "未滿15歲"
        'tmpDR("Age_1524") = "15-歲24歲"
        'tmpDR("Age_2534") = "25歲-34歲"
        'tmpDR("Age_3544") = "35歲-44歲"
        'tmpDR("Age_4554") = "45歲-54歲"
        'tmpDR("Age_55") = "55歲以上"
        'tmpDR("DegreeID_01") = "國中(含以下)"
        'tmpDR("DegreeID_02") = "高中/職"
        'tmpDR("DegreeID_03") = "專科"
        'tmpDR("DegreeID_04") = "大學"
        'tmpDR("DegreeID_05") = "碩士"
        'tmpDR("DegreeID_06") = "博士"
        For i As Integer = 0 To station_Dt.Rows.Count - 1
            intT1 = 0
            intT2 = 0

            '分署(中心)【Total】

            tmpDT = AddDR(tmpDT, station_Dt.Rows(i)("StatID").ToString, station_Dt.Rows(i)("StatName").ToString, "0")

            '職前
            tmpDT = AddDR(tmpDT, station_Dt.Rows(i)("StatID").ToString, station_Dt.Rows(i)("StatName").ToString, "1")

            '進修
            tmpDT = AddDR(tmpDT, station_Dt.Rows(i)("StatID").ToString, station_Dt.Rows(i)("StatName").ToString, "2")

        Next



        intTmp = tmpDT.Rows.Count
        If (intTmp Mod 10) = 0 Then
            PageCount = (intTmp / intPageRecord) - 1
        Else
            PageCount = intTmp / intPageRecord
        End If

        If tmpDT.Rows.Count > 0 Then
            For i As Integer = 0 To PageCount
                '表頭
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = "公立職訓機構結訓人數按訓練機構及訓練性質分"

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                nc.InnerHtml = SYMD + " ~ " + EYMD

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "right")
                nc.Attributes.Add("style", "font-size:12pt;font-family:DFKai-SB")
                nc.InnerHtml = "單位：人"
                'Column Header
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "2")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("rowspan", "2")
                nc.InnerHtml = "項 目 別"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "2")
                nc.InnerHtml = "總計"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "2")
                nc.InnerHtml = "性別"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "30%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "6")
                nc.InnerHtml = "年齡"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "30%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "6")
                nc.InnerHtml = "教育程度"

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "班次"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "人數"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "男"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "女"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "未滿15歲"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "15歲-24歲"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "25歲-34歲"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "35歲-44歲"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "45歲-54歲"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "55歲以上"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "國中(含以下)"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "高中/職"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "專科"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "大學"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "碩士"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "5%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "博士"

                '報表內容
                Dim strTmp As String = ""
                For j As Integer = 0 To intPageRecord
                    If rsCursor >= tmpDT.Rows.Count Then
                        GoTo [CONTINUE]
                    End If

                    nr = New HtmlTableRow
                    nt.Controls.Add(nr)

                    For m As Integer = 1 To tmpDT.Columns.Count - 1
                        strTmp = ""
                        strStyle = "font-size:12pt;font-family:DFKai-SB"
                        If m = 1 Then
                            strAlign = "left"

                            Select Case rsCursor Mod 3
                                Case 0
                                    strTmp = tmpDT.Rows(rsCursor)(m).ToString
                                    strStyle = "font-size:12pt;font-family:DFKai-SB;font-weight: bold;"
                                Case Else
                                    strTmp = tmpDT.Rows(rsCursor)(m).ToString
                            End Select
                            strTmp = tmpDT.Rows(rsCursor)(m).ToString

                        Else
                            strAlign = "right"
                            If tmpDT.Rows(rsCursor)(m).ToString = "" Then
                                strTmp += "0"
                            Else
                                strTmp += Convert.ToInt64(tmpDT.Rows(rsCursor)(m)).ToString("#,#0")
                            End If
                            strStyle = "font-size:10pt;font-family:DFKai-SB"
                        End If




                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        nc.Attributes.Add("align", strAlign)
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = strTmp
                    Next
                    rsCursor += 1
                Next

[CONTINUE]:
                '表尾
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "列印日期：" + Now().ToShortDateString()

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "頁數：" + (i + 1).ToString + " / " + (PageCount).ToString

                If rsCursor + 1 > tmpDT.Rows.Count Then
                    GoTo out
                End If
                '換頁列印
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
            Next

out:
        End If

    End Sub

    Private Function AddDR(ByVal tDt As DataTable, ByVal StatID As String, ByVal StatName As String, ByVal Traince As String) As DataTable
        Dim tmpDT As New DataTable
        Dim tmpDR As DataRow
        Dim tmpObj As Object

        Dim intT1, intT2 As Integer
        Dim strTrance As String = ""

        tmpDT = tDt

        Select Case Traince
            Case "0"
                strTrance = ""
            Case "1"
                strTrance = " and Trainice='" + Traince + "'"
                StatName = "　職前"
            Case "2"
                strTrance = " and Trainice='" + Traince + "'"
                StatName = "　進修"
            Case Else
                strTrance = ""
        End Select

        tmpDR = tmpDT.NewRow
        tmpDT.Rows.Add(tmpDR)
        tmpDR("Item") = StatName

        tmpObj = ClassDataDt.Compute("Count(StatID)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("ClassNum") = tmpObj.ToString

        tmpObj = dataDt.Compute("sum(Sex_M)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("M") = tmpObj.ToString
        intT1 = Convert.ToInt32(tmpObj.ToString)
        tmpObj = dataDt.Compute("sum(Sex_F)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("F") = tmpObj.ToString
        intT2 = Convert.ToInt32(tmpObj.ToString)
        tmpDR("Total") = intT1 + intT2

        tmpObj = dataDt.Compute("sum(Age_0015)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("Age_0015") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(Age_1524)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("Age_1524") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(Age_2534)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("Age_2534") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(Age_3544)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("Age_3544") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(Age_4554)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("Age_4554") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(Age_55)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("Age_55") = tmpObj.ToString

        tmpObj = dataDt.Compute("sum(DegreeID_01)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("DegreeID_01") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(DegreeID_02)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("DegreeID_02") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(DegreeID_03)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("DegreeID_03") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(DegreeID_04)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("DegreeID_04") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(DegreeID_05)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("DegreeID_05") = tmpObj.ToString
        tmpObj = dataDt.Compute("sum(DegreeID_06)", "StatID='" + StatID + "'" + strTrance)
        tmpDR("DegreeID_06") = tmpObj.ToString

        Return tmpDT

    End Function

End Class