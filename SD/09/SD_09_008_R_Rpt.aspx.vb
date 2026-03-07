'Imports System.Data
'Imports Oracle.DataAccess.Client
'Imports System.Xml.Serialization

<Serializable()>
Partial Class SD_09_008_R_Rpt
    Inherits AuthBasePage

    'Dim arrItem(113, 113) As String
    'Dim arrSpanRow(114) As String
    'Dim sourceDT As DataTable = Nothing
    'Dim PName As String = ""
    'Dim plankind As String = ""

    <Serializable()>
    Public Class myrow
        Public StudentID As String
        Public name As String
        Public ItemVar1 As String
        Public TechPoint As String
        Public RemedPoint As String
        Public MinusLeave As String
        Public MinusSanction As String
        Public total_field As String
        Public field_hour As List(Of Integer)
    End Class

    'Dim SYMD As String = "102年01月"
    'Dim EYMD As String = "102年08月"

    Dim tDt_main As DataTable
    Dim field1 As String = "field1"
    Dim orgName As String = "orgName"
    Dim PlanName As String = "PlanName"
    Dim ClassCName As String = "ClassCName"

    Dim col As List(Of String) = New List(Of String)
    Dim row As List(Of myrow) = New List(Of myrow)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        'Dim Years As String = "2012"
        'Dim OCID As String = "41640"
        'Dim CJOB_UNKEY As String = "363"
        Dim Years As String = Request("Years")
        Dim OCID As String = Request("OCID")
        Dim CJOB_UNKEY As String = Request("CJOB_UNKEY")
        Years = TIMS.ClearSQM(Years)
        OCID = TIMS.ClearSQM(OCID)
        CJOB_UNKEY = TIMS.ClearSQM(CJOB_UNKEY)

        If Years = "" AndAlso OCID = "" AndAlso CJOB_UNKEY = "" Then
            Exit Sub
        End If

        Dim tDt As New DataTable

        'STDate = Request("STDate")
        'STDate2 = Request("STDate2")
        'DistID = Request("DistID")
        'DistName = Request("DistName")
        'title = Request("title")
        Try
            tDt_main = db_main(Years, OCID, CJOB_UNKEY)
        Catch ex As Exception

            Dim strErrmsg As String = ""
            strErrmsg += "Years: " & Years & vbCrLf
            strErrmsg += "OCID: " & OCID & vbCrLf
            strErrmsg += "CJOB_UNKEY: " & CJOB_UNKEY & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
        End Try

        'Dim i As Integer
        'Dim j As Integer
        'Dim k As Integer
        'Dim l As Integer
        'Dim m As Integer

        Dim b As Boolean = False
        'Dim r As myrow
        Dim str As String = ""
        Dim str2 As String = ""
        Dim str3 As String = ""
        Dim str4 As String = ""
        Dim str5 As String = ""
        Dim str6 As String = ""
        Dim str7 As String = ""
        Dim str8 As String = ""
        Dim str9 As String = ""
        'Dim cv As Integer

        For i As Integer = 0 To tDt_main.Rows.Count - 1
            orgName = tDt_main.Rows(i).Item("orgName").ToString
            PlanName = tDt_main.Rows(i).Item("PlanName").ToString
            ClassCName = tDt_main.Rows(i).Item("ClassCName").ToString
            field1 = tDt_main.Rows(i).Item("field1").ToString
            str = tDt_main.Rows(i).Item("key_name").ToString
            b = False
            For j As Integer = 0 To col.Count - 1
                If str = col.Item(j) Then
                    b = True
                    Exit For
                End If
            Next
            If Not b Then
                If str = "" Then
                    col.Insert(0, New String(str))
                Else
                    col.Add(New String(str))
                End If
            End If
        Next

        For i As Integer = 0 To tDt_main.Rows.Count - 1
            str = tDt_main.Rows(i).Item("name").ToString
            str2 = tDt_main.Rows(i).Item("StudentID").ToString
            str3 = Double.Parse(tDt_main.Rows(i).Item("ItemVar1").ToString).ToString("0.00")
            str4 = Double.Parse(tDt_main.Rows(i).Item("TechPoint").ToString).ToString("0.00")
            str5 = Double.Parse(tDt_main.Rows(i).Item("RemedPoint").ToString).ToString("0.00")
            str6 = tDt_main.Rows(i).Item("MinusLeave").ToString
            str7 = tDt_main.Rows(i).Item("MinusSanction").ToString
            str8 = tDt_main.Rows(i).Item("total_field").ToString
            Dim cv As Integer = 0
            If tDt_main.Rows(i).Item("field_hour").ToString <> "" Then
                cv = tDt_main.Rows(i).Item("field_hour").ToString
            End If
            str9 = tDt_main.Rows(i).Item("key_name").ToString

            b = False
            For j As Integer = 0 To row.Count - 1
                If str = row.Item(j).name Then
                    b = True
                    Exit For
                End If
            Next
            If Not b Then
                Dim r As myrow = New myrow
                r.name = str
                r.StudentID = str2
                r.ItemVar1 = str3
                r.TechPoint = str4
                r.RemedPoint = str5
                r.MinusLeave = str6
                r.MinusSanction = str7
                r.total_field = str8

                r.field_hour = New List(Of Integer)
                For j As Integer = 0 To col.Count - 1
                    r.field_hour.Add(New Integer)
                    If str9 = col.Item(j) Then
                        r.field_hour.Item(j) += cv
                    End If
                Next

                row.Add(r)
            End If
        Next

        PrintDiv(tDt_main, "AAAA", "150", "150", 40, "10", "V")
        Call exc_Print(TIMS.c_false)
    End Sub

    Function db_main(ByVal Years As String, ByVal OCID As String, ByVal CJOB_UNKEY As String) As DataTable
        Dim resDt As DataTable = Nothing

        If Years = "" AndAlso OCID = "" AndAlso CJOB_UNKEY = "" Then Return resDt

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select  total_item.Years+'年度'+f.PlanName+'　操行成績明細表' PlanName" & vbCrLf
        sql += " ,total_item.ClassCName" & vbCrLf
        sql += " ,total_item.key_name" & vbCrLf
        sql += " ,total_item.StudentID" & vbCrLf
        sql += " ,total_item.Name" & vbCrLf
        sql += " ,total_item.field1" & vbCrLf
        sql += " ,total_item.field_hour" & vbCrLf
        sql += " ,total_item.BehaviorResult" & vbCrLf
        sql += " ,dbo.NVL(CONVERT(numeric, replace(i.TechPoint, ' ','0')), 0) TechPoint" & vbCrLf
        sql += " ,dbo.NVL(CONVERT(numeric, replace(i.RemedPoint, ' ','0')), 0) RemedPoint" & vbCrLf
        sql += " ,dbo.NVL(i.MinusLeave,0) as MinusLeave" & vbCrLf
        sql += " ,dbo.NVL(i.MinusSanction,0) as MinusSanction" & vbCrLf
        sql += " ,h.orgName" & vbCrLf
        sql += " ,ss.ItemVar1 " & vbCrLf
        sql += " + dbo.NVL(CONVERT(numeric, replace(i.TechPoint, ' ','0')), 0) " & vbCrLf
        sql += " + dbo.NVL(CONVERT(numeric, replace(i.RemedPoint, ' ','0')), 0) " & vbCrLf
        sql += " - dbo.NVL(i.MinusLeave,0)" & vbCrLf
        sql += " + dbo.NVL(i.MinusSanction,0) total_field" & vbCrLf
        sql += " ,ss.ItemVar1 " & vbCrLf
        sql += " FROM (" & vbCrLf ' /*出缺席資料*/  
        sql += "   SELECT  c.PlanID,substr(a.StudentID, -2) as StudentID" & vbCrLf
        sql += "   ,c.ClassCName" & vbCrLf
        sql += "   +case when dbo.NVL(c.levelType,'0') !='0' then '第'+c.levelType+'階段' end " & vbCrLf
        sql += "   +'第'+c.CyclType+'期' as ClassCName" & vbCrLf
        sql += "   ,d.Name" & vbCrLf
        sql += "   ,a.OCID" & vbCrLf
        sql += "   ,CONVERT(varchar, key_leave.name) as key_name" & vbCrLf
        sql += "   ,a.BehaviorResult" & vbCrLf
        sql += "   ,CONVERT(varchar, b.hours) as field_hour" & vbCrLf
        sql += "   ,'假別' as field1" & vbCrLf
        sql += "   ,g.orgid" & vbCrLf
        sql += "   ,e.TPlanID" & vbCrLf
        sql += "   ,e.Years" & vbCrLf
        sql += "   ,key_leave.SOCID   " & vbCrLf

        sql += "   FROM   (" & vbCrLf
        sql += "   SELECT * " & vbCrLf
        sql += "   FROM (SELECT DISTINCT  st.SOCID  FROM Stud_Turnout st ) A " & vbCrLf
        sql += "   cross JOIN ( SELECT leaveid,name FROM key_leave ) B ) key_leave    " & vbCrLf
        sql += "   LEFT JOIN stud_Turnout b on key_leave.SOCID=b.SOCID AND b.leaveid=key_leave.leaveid    " & vbCrLf
        sql += "   JOIN Class_studentsofclass a on a.SOCID=key_leave.SOCID    " & vbCrLf
        sql += "   JOIN Class_ClassInfo c on c.OCID=a.OCID " & vbCrLf
        sql += "   JOIN Stud_StudentInfo d on d.SID=a.SID    " & vbCrLf
        sql += "   JOIN Auth_Relship g on g.RID=c.RID " & vbCrLf
        sql += "   JOIN ID_Plan e on c.PlanID=e.PlanID    " & vbCrLf
        sql += "   WHERE 1=1 " & vbCrLf
        If Years <> "" Then
            sql += "    and e.Years='" + Years + "'" & vbCrLf
        End If
        If OCID <> "" Then
            sql += "    and a.OCID='" + OCID + "'" & vbCrLf
        End If
        If CJOB_UNKEY <> "" Then
            sql += "    and c.CJOB_UNKEY = '" + CJOB_UNKEY + "'" & vbCrLf
        End If
        sql += " UNION  all  /*獎懲資料*/   " & vbCrLf
        sql += "   SELECT  c.PlanID" & vbCrLf
        sql += "   ,substr(a.StudentID, -2) as StudentID" & vbCrLf
        sql += "   ,c.ClassCName" & vbCrLf
        sql += "   + case when (c.levelType is null or c.levelType='0') then ''   else '第'+c.levelType+'階段' end    " & vbCrLf
        sql += "   +case when (c.CyclType is null or c.CyclType='00') then ''   else '第'+c.CyclType+'期' end as ClassCName" & vbCrLf
        sql += "   ,d.Name" & vbCrLf
        sql += "   ,a.OCID" & vbCrLf
        sql += "   ,CONVERT(varchar, key_San.name) as key_name" & vbCrLf
        sql += "   ,a.BehaviorResult" & vbCrLf
        sql += "   ,CONVERT(varchar, b.Times) as field_hour" & vbCrLf
        sql += "   ,'獎懲' as field1" & vbCrLf
        sql += "   ,g.orgid" & vbCrLf
        sql += "   ,e.TPlanID" & vbCrLf
        sql += "   ,e.Years" & vbCrLf
        sql += "   ,key_San.SOCID   " & vbCrLf
        sql += "   FROM   (" & vbCrLf
        sql += "   SELECT * " & vbCrLf
        sql += "   FROM    (SELECT DISTINCT  SOCID FROM Stud_Sanction) A " & vbCrLf
        sql += "   cross JOIN (SELECT SanID,name FROM key_Sanction ) B )  key_San    " & vbCrLf
        sql += "   LEFT JOIN Stud_Sanction b on key_San.SOCID=b.SOCID " & vbCrLf
        sql += "   AND b.SanID=key_San.SanID    " & vbCrLf
        sql += "   JOIN Class_studentsofclass a on  a.SOCID=key_San.SOCID    " & vbCrLf
        sql += "   JOIN Class_ClassInfo c on c.OCID=a.OCID " & vbCrLf
        sql += "   JOIN Stud_StudentInfo d on d.SID=a.SID    " & vbCrLf
        sql += "   JOIN Auth_Relship g on g.RID=c.RID    " & vbCrLf
        sql += "   JOIN ID_Plan e on c.PlanID=e.PlanID     " & vbCrLf
        sql += "   WHERE 1=1    " & vbCrLf
        If Years <> "" Then
            sql += "    and e.Years='" + Years + "'" & vbCrLf
        End If
        If OCID <> "" Then
            sql += "    and a.OCID='" + OCID + "'" & vbCrLf
        End If
        If CJOB_UNKEY <> "" Then
            sql += "    and c.CJOB_UNKEY = '" + CJOB_UNKEY + "'" & vbCrLf
        End If
        sql += " UNION all /*沒有出缺席與獎懲資料*/   " & vbCrLf
        sql += "   SELECT  c.PlanID" & vbCrLf
        sql += "   ,substr(a.StudentID, -2) as StudentID" & vbCrLf
        sql += "   ,c.ClassCName+   case when (c.levelType is null or c.levelType='0') then ''   else '第'+c.levelType+'階段' end    +case when (c.CyclType is null or c.CyclType='00') then ''   else '第'+c.CyclType+'期' end as ClassCName" & vbCrLf
        sql += "   ,d.Name" & vbCrLf
        sql += "   ,a.OCID" & vbCrLf
        sql += "   ,'' as key_name" & vbCrLf
        sql += "   , a.BehaviorResult" & vbCrLf
        sql += "   ,'' as field_hour" & vbCrLf
        sql += "   ,'' as field1" & vbCrLf
        sql += "   ,g.orgid" & vbCrLf
        sql += "   ,e.TPlanID" & vbCrLf
        sql += "   ,e.Years" & vbCrLf
        sql += "   ,a.SOCID   " & vbCrLf
        sql += "   FROM   (" & vbCrLf
        sql += "   SELECT * " & vbCrLf
        sql += "   FROM Class_StudentsOfClass " & vbCrLf
        sql += "   WHERE SOCID not in   (" & vbCrLf
        sql += "   SELECT DISTINCT SOCID FROM Stud_Sanction " & vbCrLf
        sql += "   UNION " & vbCrLf
        sql += "   SELECT DISTINCT SOCID FROM Stud_Turnout))  a    " & vbCrLf
        sql += "   JOIN Class_ClassInfo c on c.OCID=a.OCID     " & vbCrLf
        sql += "   JOIN Stud_StudentInfo d on d.SID=a.SID    " & vbCrLf
        sql += "   JOIN Auth_Relship g on g.RID=c.RID    " & vbCrLf
        sql += "   JOIN ID_Plan e on c.PlanID=e.PlanID     " & vbCrLf
        sql += "   WHERE 1=1    " & vbCrLf
        If Years <> "" Then
            sql += "    and e.Years='" + Years + "'" & vbCrLf
        End If
        If OCID <> "" Then
            sql += "    and a.OCID='" + OCID + "'" & vbCrLf
        End If
        If CJOB_UNKEY <> "" Then
            sql += "    and c.CJOB_UNKEY = '" + CJOB_UNKEY + "'" & vbCrLf
        End If
        sql += " ) total_item  " & vbCrLf
        sql += " JOIN Key_Plan f on f.TPlanID=total_item.TPlanID  " & vbCrLf
        sql += " JOIN Org_orginfo h on h.orgid=total_item.orgid  " & vbCrLf
        sql += " LEFT JOIN Stud_Conduct i on i.SOCID=total_item.SOCID   " & vbCrLf
        sql += " JOIN (" & vbCrLf
        sql += "   SELECT ip.PlanID" & vbCrLf
        sql += "   ,CONVERT(numeric, dbo.NVL(sg.ItemVar1,'0') ) ItemVar1" & vbCrLf
        sql += "   FROM Sys_GlobalVar sg   " & vbCrLf
        sql += "   JOIN ID_Plan ip ON ip.TPlanID = sg.TPlanID " & vbCrLf
        sql += "   AND ip.DistID = sg.DistID  " & vbCrLf
        sql += "   WHERE (sg.GVID = '3')" & vbCrLf
        sql += " ) ss ON ss.PlanID =  total_item.PlanID  " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        If Years <> "" Then
            sql += " and total_item.Years='" + Years + "'" & vbCrLf
        End If
        If OCID <> "" Then
            sql += " and total_item.OCID='" + OCID + "'" & vbCrLf
        End If

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Sub PrintDiv(ByVal dt As DataTable, ByVal selRpt As String, ByVal Field1_width As String, ByVal Field2_width As String, ByVal RCount As Integer, ByVal font_size As String, ByVal portrait As String)
        'dt:要顯示的資料,selRpt:,Field1_width:標題題目的寬度,Field2_width:標題題目的寬度,RCount:每頁筆數,font_size:內容字型大小,portrait:直式/橫式

        Dim tmpDT As New DataTable
        'Dim tmpDR As DataRow
        'Dim tmpObj As Object
        Dim sql As String = ""
        Dim PageCount As Int32 = 0  'Pages
        Dim ReportCount As Integer = RCount '每頁筆數
        Dim ColCount As Integer = 0
        Dim intTmp As Integer = 0
        Dim rsCursor As Integer = 0   '報表內容列印的NO
        Dim intPageRecord As Integer = RCount '每頁列印幾筆

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl
        Dim strStyle As String = "font-size:" + font_size + "pt;font-family:DFKai-SB"
        Dim strStyle2 As String = "font-size:14pt;font-family:DFKai-SB"
        Dim int_width As Integer
        Dim strWatermarkImg As String
        Dim strWatermarkDiv As String
        Dim intWatermarkTop As Integer

        tmpDT = dt
        ColCount = dt.Columns.Count

        intTmp = tmpDT.Rows.Count
        PageCount = 1
        'If (intTmp Mod ReportCount) = 0 Then
        '    PageCount = (intTmp / ReportCount) - 1
        'Else
        '    PageCount = intTmp / ReportCount
        'End If

        '表格寬度的設定
        'If portrait = "H" Then
        '    int_width = Int((550 - Field1_width - Field2_width) / 19)
        '    strWatermarkImg = "TIMS_1.jpg"
        'Else
        int_width = 100 'Int((820 - Field1_width - Field2_width) / 19)
        strWatermarkImg = "TIMS_1.jpg"
        'End If

        If dt.Rows.Count > 0 Then
            'For i As Integer = 0 To PageCount
            '加背景圖的div
            If portrait = "H" Then
                intWatermarkTop = 0 * 800
            Else
                intWatermarkTop = 0 * 550
            End If
            strWatermarkDiv = "<div style='position:absolute;z-index:-1; margin:0;padding:0;left:0px;top: " + intWatermarkTop.ToString + "px;'><img src='../../images/rptpic/temple/" + strWatermarkImg + "' /></div>"
            nl = New HtmlGenericControl
            div_print.Controls.Add(nl)
            nl.InnerHtml = strWatermarkDiv

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
            nc.Attributes.Add("colspan", "3")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = ""

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "100%")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("colspan", "3")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = orgName

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "100%")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("colspan", "3")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = PlanName

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "5%")
            nc.Attributes.Add("align", "right")
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = "列印日期:"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "5%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = Now().ToShortDateString()

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "90%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = ""

            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "5%")
            nc.Attributes.Add("align", "right")
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = "頁數："

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "5%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = (0 + 1).ToString + " / " + PageCount.ToString

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "90%")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("colspan", "1")
            nc.Attributes.Add("style", strStyle2)
            nc.InnerHtml = ""

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
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "學號"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "100")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "學生姓名"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "操行底分"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "導師+-"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "輔導課+-"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "假別"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "獎懲"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("rowspan", "2")
            nc.InnerHtml = "總分"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            'nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("colspan", col.Count)
            nc.InnerHtml = field1

            Dim i As Integer = 0
            nr = New HtmlTableRow
            nt.Controls.Add(nr)
            For i = 0 To col.Count - 1
                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = col.Item(i)
            Next


            nr = New HtmlTableRow
            nt.Controls.Add(nr)

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            'nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "center")
            nc.Attributes.Add("style", strStyle)
            nc.InnerHtml = "班別:"

            nc = New HtmlTableCell
            nr.Controls.Add(nc)
            'nc.Attributes.Add("width", "40")
            nc.Attributes.Add("align", "left")
            nc.Attributes.Add("style", strStyle)
            nc.Attributes.Add("colspan", col.Count + 7)
            nc.InnerHtml = ClassCName

            For i = 0 To row.Count - 1
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).StudentID

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).name

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).ItemVar1

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).TechPoint

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).RemedPoint

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).MinusLeave

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).MinusSanction

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = row.Item(i).total_field

                For j As Integer = 0 To row.Item(i).field_hour.Count - 1
                    nc = New HtmlTableCell
                    nr.Controls.Add(nc)
                    'nc.Attributes.Add("width", "40")
                    nc.Attributes.Add("align", "right")
                    nc.Attributes.Add("style", strStyle)
                    nc.InnerHtml = row.Item(i).field_hour.Item(j)
                Next
            Next

[CONTINUE]:
            '表尾
            'If rsCursor + 1 > tmpDT.Rows.Count Then
            '    GoTo out
            'End If
            '換頁列印
            nl = New HtmlGenericControl
            div_print.Controls.Add(nl)
            nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"
            'Next
out:
        End If
    End Sub

    Private Sub exc_Print(ByVal portrait As String)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">window.print();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        Return

        'Dim strScript As String = ""
        'strScript = "<script language=""javascript"">" + vbCrLf
        ''strScript = "function print() {"
        'strScript += "if (!factory.object) {"
        ''strScript += "return"
        'strScript += "} else {"
        'strScript += "factory.printing.header = """";"
        'strScript += "factory.printing.footer = """";"
        'strScript += "factory.printing.leftMargin = 5; "
        'strScript += "factory.printing.topMargin = 10; "
        'strScript += "factory.printing.rightMargin = 5; "
        'strScript += "factory.printing.bottomMargin = 10; "
        'strScript += "factory.printing.portrait = " + portrait + ";"
        'strScript += "factory.printing.Print(true);"
        'strScript += "window.close();"
        'strScript += "}"
        ''strScript += "}"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
    End Sub

End Class
