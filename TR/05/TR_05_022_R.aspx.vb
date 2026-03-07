Public Class TR_05_022_R
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            Call CreateItem()
        End If

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        '選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If
        Export1.Attributes("onclick") = "return search();"
        'Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()

        'Dim dt As DataTable
        'Dim dr As DataRow
        'Dim sqlstr As String
        Syear = TIMS.GetSyear(Syear) '年度
        Common.SetListItem(Syear, Now.Year)

        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) 'Handles Button1.Click
    '    Dim msg As String = ""

    '    If Me.STDate1.Text = "" And Me.STDate2.Text = "" Then
    '        If Me.FTDate1.Text = "" And Me.FTDate2.Text = "" Then
    '            If Syear.SelectedValue = "" Then
    '                msg += "年度、開訓日期、結訓日期擇一為查詢條件!!" & vbCrLf
    '            End If
    '        End If
    '    End If
    '    If msg <> "" Then
    '        Common.MessageBox(Me, msg)
    '        Exit Sub
    '    End If

    '    '選擇轄區 '報表要用的轄區參數
    '    Dim DistID1 As String
    '    Dim DistName As String
    '    DistID1 = ""
    '    DistName = ""
    '    For i As Integer = 1 To Me.DistID.Items.Count - 1
    '        If Me.DistID.Items(i).Selected Then
    '            If DistID1 <> "" Then DistID1 = ","
    '            DistID1 &= Convert.ToString("\'" & Me.DistID.Items(i).Value & "\'")
    '            If DistName <> "" Then DistName = ","
    '            DistName &= Convert.ToString("\'" & Me.DistID.Items(i).Text & "\'")
    '        End If
    '    Next

    '    '報表要用的訓練計畫參數
    '    Dim TPlanID1 As String
    '    Dim TPlanName As String
    '    TPlanID1 = ""
    '    TPlanName = ""
    '    For i As Integer = 1 To Me.TPlanID.Items.Count - 1
    '        If Me.TPlanID.Items(i).Selected Then
    '            If TPlanID1 <> "" Then TPlanID1 = ","
    '            TPlanID1 &= Convert.ToString("\'" & Me.TPlanID.Items(i).Value & "\'")
    '            If TPlanName <> "" Then TPlanName = ","
    '            TPlanName &= Convert.ToString("\'" & Me.TPlanID.Items(i).Text & "\'")
    '        End If
    '    Next

    '    Dim Years As String = Syear.SelectedValue
    '    If Syear.SelectedValue = "" Then
    '        Years = sm.UserInfo.Years
    '    End If

    '    Dim MyValue As String = ""
    '    MyValue = "jkl=jlk"
    '    MyValue += "&Years=" & Years
    '    MyValue += "&STDate1=" & STDate1.Text
    '    MyValue += "&STDate2=" & STDate2.Text
    '    MyValue += "&FTDate1=" & FTDate1.Text
    '    MyValue += "&FTDate2=" & FTDate2.Text
    '    MyValue += "&DistID=" & DistID1
    '    MyValue += "&DistName=" & DistName
    '    MyValue += "&TPlanID=" & TPlanID1
    '    MyValue += "&PlanName=" & TPlanName

    '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report2011", "TR_05_022_R", MyValue)
    'End Sub

    Protected Sub Export1_Click(sender As Object, e As EventArgs) Handles Export1.Click
        Dim dt As DataTable
        dt = LoadData1()
        Call ExpReport1(dt)
    End Sub

    Function LoadData1() As DataTable
        Dim rst As DataTable
        '報表要用的轄區參數,1:為起始位置
        Dim DistID1 As String = ""
        DistID1 = TIMS.GetCheckBoxListRptVal(DistID, 1)
        '報表要用的訓練計畫參數,1:為起始位置
        Dim TPlanID1 As String = ""
        TPlanID1 = TIMS.GetCheckBoxListRptVal(TPlanID, 1)

        Dim strSchpp As String = ""
        Dim strSchcc As String = ""
        If Syear.SelectedValue <> "" Then
            strSchpp += " and ip.years ='" & Syear.SelectedValue & "'" & vbCrLf
            strSchcc += " and cc.years ='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If DistID1 <> "" Then
            '轉換sql查詢使用
            strSchpp += " and ip.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
            strSchcc += " and cc.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            '轉換sql查詢使用
            strSchpp += " and ip.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
            strSchcc += " and cc.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If STDate1.Text <> "" Then
            strSchpp += " and pp.stdate>= " & TIMS.To_date(STDate1.Text) & vbCrLf
            strSchcc += " and cc.stdate>= " & TIMS.To_date(STDate1.Text) & vbCrLf '" & STDate1.Text & "'" & vbCrLf
        End If
        If STDate2.Text <> "" Then
            strSchpp += " and pp.stdate<= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
            strSchcc += " and cc.stdate<= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
        End If
        If FTDate1.Text <> "" Then
            strSchpp += " and pp.fddate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
            strSchcc += " and cc.ftdate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            strSchpp += " and pp.fddate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
            strSchcc += " and cc.ftdate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
        End If

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select k1.tplanid" & vbCrLf
        sql += " ,k1.planname" & vbCrLf
        sql += " ,g1.PNum" & vbCrLf
        sql += " ,g2.CNum" & vbCrLf
        sql += " ,dbo.NVL(g3.A02,0) A02" & vbCrLf
        sql += " ,dbo.NVL(g3.A03,0) A03" & vbCrLf
        sql += " ,dbo.NVL(g3.A01,0) A01" & vbCrLf
        sql += " ,dbo.NVL(g3.A97,0) A97" & vbCrLf
        sql += " ,dbo.NVL(g3.ACNT,0) ACNT" & vbCrLf

        sql += " ,dbo.NVL(g3.B02,0) B02" & vbCrLf
        sql += " ,dbo.NVL(g3.B03,0) B03" & vbCrLf
        sql += " ,dbo.NVL(g3.B01,0) B01" & vbCrLf
        sql += " ,dbo.NVL(g3.B97,0) B97" & vbCrLf
        sql += " ,dbo.NVL(g3.BCNT,0) BCNT" & vbCrLf

        sql += " ,dbo.NVL(g3.C02,0) C02" & vbCrLf
        sql += " ,dbo.NVL(g3.C03,0) C03" & vbCrLf
        sql += " ,dbo.NVL(g3.C01,0) C01" & vbCrLf
        sql += " ,dbo.NVL(g3.C97,0) C97" & vbCrLf
        sql += " ,dbo.NVL(g3.CCNT,0) CCNT" & vbCrLf

        sql += " from key_plan k1" & vbCrLf
        sql += " join (" & vbCrLf
        sql += " 	SELECT ip.tplanid ,count(*) PNum" & vbCrLf
        sql += " 	from plan_planinfo pp" & vbCrLf
        sql += " 	join view_plan ip on ip.planid =pp.planid " & vbCrLf
        sql += " 	where 1=1" & vbCrLf
        sql += " 	and pp.IsApprPaper='Y'" & vbCrLf
        sql += " 	and pp.AppliedResult='Y'" & vbCrLf
        sql += strSchpp
        sql += " 	group by ip.tplanid" & vbCrLf
        sql += " ) g1 on g1.tplanid=k1.tplanid" & vbCrLf
        sql += " join (" & vbCrLf
        sql += " 	SELECT cc.tplanid ,count(*) CNum" & vbCrLf
        sql += " 	from view2 cc" & vbCrLf
        sql += " 	where 1=1" & vbCrLf
        sql += strSchcc
        sql += " 	group by cc.tplanid" & vbCrLf
        sql += " ) g2 on g2.tplanid=k1.tplanid" & vbCrLf
        sql += " join (" & vbCrLf
        sql += " 	SELECT cc.tplanid " & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='02' then 1 end) A02" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='03' then 1 end) A03" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='01' then 1 end) A01" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='97' then 1 end) A97" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID in ('02','03','01','97') then 1 end) ACnt" & vbCrLf

        sql += " 	,sum(case when cs.BudgetID='02' and cs.studstatus not in (2,3) then 1 end) B02" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='03' and cs.studstatus not in (2,3) then 1 end) B03" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='01' and cs.studstatus not in (2,3) then 1 end) B01" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='97' and cs.studstatus not in (2,3) then 1 end) B97" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID in ('02','03','01','97') and cs.studstatus not in (2,3) then 1 end) BCnt" & vbCrLf

        sql += " 	,sum(case when cs.BudgetID='02' and cs.studstatus not in (2,3) and sg3.socid is not null then 1 end) C02" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='03' and cs.studstatus not in (2,3) and sg3.socid is not null then 1 end) C03" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='01' and cs.studstatus not in (2,3) and sg3.socid is not null then 1 end) C01" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID='97' and cs.studstatus not in (2,3) and sg3.socid is not null then 1 end) C97" & vbCrLf
        sql += " 	,sum(case when cs.BudgetID in ('02','03','01','97') and cs.studstatus not in (2,3) and sg3.socid is not null then 1 end) CCnt" & vbCrLf

        sql += " 	from view2 cc" & vbCrLf
        sql += " 	join class_studentsofclass cs on cs.ocid=cc.ocid " & vbCrLf
        sql += " 	join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        sql += " 	join stud_subdata ss2 on ss2.sid =cs.sid" & vbCrLf
        sql += " 	join view_budget bb on bb.budid=cs.BudgetID" & vbCrLf
        sql += " 	left join Stud_GetJobState3 sg3 on sg3.CPoint=1 and sg3.socid =cs.socid and sg3.IsGetJob='1'" & vbCrLf
        sql += " 	where 1=1" & vbCrLf
        sql += strSchcc
        sql += " 	group by cc.tplanid" & vbCrLf
        sql += " ) g3 on g3.tplanid=k1.tplanid" & vbCrLf
        'sql += " where 1=1" & vbCrLf

        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    Sub ExpReport1(ByRef dt As DataTable)
        Dim strTitle1 As String = "各項訓練計畫辦理情形統計月報表"

        Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(strTitle1, System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/ms-excel"
        'Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        Common.RespWrite(Me, "<html>")
        Common.RespWrite(Me, "<head>")
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=BIG5>")
        '<head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, ".noDecFormat{mso-number-format:""0"";}")
        ''mso-number-format:"0" 
        'Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, "</head>")

        Common.RespWrite(Me, "<body>")
        Common.RespWrite(Me, "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        Dim ExportStr As String = ""

        '建立抬頭
        '第1行
        ExportStr = ""
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td rowspan=""2"">序號</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">計畫</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">核定開訓班數</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">已開訓班數</td>" & vbTab
        ExportStr &= "<td colspan=""5"">開訓人數(依預算別)</td>" & vbTab
        ExportStr &= "<td colspan=""5"">結訓人數(依預算別)</td>" & vbTab
        ExportStr &= "<td colspan=""5"">就業人數(依預算別)</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        '第2行
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>就安</td>" & vbTab
        ExportStr &= "<td>就保</td>" & vbTab
        ExportStr &= "<td>公務</td>" & vbTab
        ExportStr &= "<td>協助<br />(ECFA)</td>" & vbTab
        ExportStr &= "<td>合計</td>" & vbTab
        ExportStr &= "<td>就安</td>" & vbTab
        ExportStr &= "<td>就保</td>" & vbTab
        ExportStr &= "<td>公務</td>" & vbTab
        ExportStr &= "<td>協助<br />(ECFA)</td>" & vbTab
        ExportStr &= "<td>合計</td>" & vbTab
        ExportStr &= "<td>就安</td>" & vbTab
        ExportStr &= "<td>就保</td>" & vbTab
        ExportStr &= "<td>公務</td>" & vbTab
        ExportStr &= "<td>協助<br />(ECFA)</td>" & vbTab
        ExportStr &= "<td>合計</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            ExportStr &= "<td>" & Convert.ToString(iSeqno) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("planname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("PNum")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CNum")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("A02")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A03")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A01")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A97")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ACNT")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("B02")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B03")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B01")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B97")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("BCNT")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("C02")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C03")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C01")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C97")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("CCNT")) & "</td>" & vbTab
            ExportStr &= "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

        Response.End()
    End Sub
End Class