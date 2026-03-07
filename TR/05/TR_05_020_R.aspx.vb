Public Class TR_05_020_R
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
            CreateItem()
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

#Region "NO USE"
    'Protected Sub Button1_Click(sender As Object, e As EventArgs) 'Handles Button1.Click
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

    '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report2011", "TR_05_020_R", MyValue)

    'End Sub
#End Region

    Private Sub Export1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Export1.Click
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

        Dim sql As String = ""
        sql = "" & vbCrLf
        'sql += " /*33低收入戶 07生活失助戶 a開b結c就*/" & vbCrLf
        sql = "" & vbCrLf
        sql += " SELECT cc.orgname" & vbCrLf
        sql += " ,cc.classcname" & vbCrLf
        sql += " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql += " ,CONVERT(varchar, cc.ftdate, 111) ftdate" & vbCrLf
        sql += " ,cc.thours" & vbCrLf
        sql += " ,cc.ocid " & vbCrLf
        'sql += " ,dbo.NVL(g2.EnterCnt,0) EnterCnt" & vbCrLf
        sql += " ,dbo.NVL(g.ID33a,0) ID33a" & vbCrLf
        sql += " ,dbo.NVL(g.ID33b,0) ID33b" & vbCrLf
        sql += " ,dbo.NVL(g.ID33c,0) ID33c" & vbCrLf
        sql += " ,dbo.NVL(g.ID07a,0) ID07a" & vbCrLf
        sql += " ,dbo.NVL(g.ID07b,0) ID07b" & vbCrLf
        sql += " ,dbo.NVL(g.ID07c,0) ID07c" & vbCrLf
        sql += " FROM VIEW2 cc" & vbCrLf
        sql += " join (" & vbCrLf
        sql += " 	select cs.OCID " & vbCrLf
        sql += " 	,sum (CASE when cs.MIdentityID='33' then 1 end ) ID33a" & vbCrLf
        sql += " 	,sum (CASE when cs.MIdentityID='07' then 1 end ) ID07a" & vbCrLf
        sql += " 	,sum (CASE when cs.MIdentityID='33' and cs.studstatus not in (2,3) then 1 end ) ID33b" & vbCrLf
        sql += " 	,sum (CASE when cs.MIdentityID='07' and cs.studstatus not in (2,3) then 1 end ) ID07b" & vbCrLf
        sql += " 	,sum (CASE when cs.MIdentityID='33' and cs.studstatus not in (2,3) and j3.socid is not null then 1 end ) ID33c" & vbCrLf
        sql += " 	,sum (CASE when cs.MIdentityID='07' and cs.studstatus not in (2,3) and j3.socid is not null then 1 end ) ID07c" & vbCrLf
        sql += " 	FROM class_studentsofclass cs" & vbCrLf
        sql += " 	left join Stud_GetJobState3 j3 on j3.socid =cs.socid and j3.cpoint =1  and j3.IsGetJob =1" & vbCrLf
        sql += " 	group by cs.OCID " & vbCrLf
        sql += " ) g on g.ocid =cc.ocid " & vbCrLf
        'sql += " join (" & vbCrLf
        'sql += " 	select OCID1 OCID " & vbCrLf
        'sql += " 	,count(*) EnterCnt " & vbCrLf
        'sql += " 	FROM v_entertype1 cs" & vbCrLf
        'sql += " 	group by OCID1 " & vbCrLf
        'sql += " ) g2 on g2.ocid =cc.ocid " & vbCrLf

        sql += " where 1=1" & vbCrLf
        If Syear.SelectedValue <> "" Then
            sql += " and cc.years ='" & Syear.SelectedValue & "'" & vbCrLf
        End If
        If DistID1 <> "" Then
            '轉換sql查詢使用
            sql += " and cc.DistID IN (" & DistID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If TPlanID1 <> "" Then
            '轉換sql查詢使用
            sql += " and cc.TPlanID IN (" & TPlanID1.Replace("\'", "'") & ")" & vbCrLf
        End If
        If STDate1.Text <> "" Then
            sql += " and cc.stdate>= " & TIMS.To_date(STDate1.Text) & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql += " and cc.stdate<= " & TIMS.To_date(STDate2.Text) & vbCrLf '" & STDate2.Text & "'" & vbCrLf
        End If
        If FTDate1.Text <> "" Then
            sql += " and cc.ftdate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql += " and cc.ftdate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
        End If

        rst = DbAccess.GetDataTable(sql, objconn)

        Return rst
    End Function

    Sub ExpReport1(ByRef dt As DataTable)
        Dim strTitle1 As String = "低收與中低職訓措施辦理情形"

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
        ExportStr &= "<td rowspan=""2"">訓練機構</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">班別名稱</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">訓練起迄</td>" & vbTab
        ExportStr &= "<td rowspan=""2"">時數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">低收入戶</td>" & vbTab
        ExportStr &= "<td colspan=""4"">中收入戶</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf

        '第2行
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>報名人數</td>" & vbTab
        ExportStr &= "<td>開訓人數</td>" & vbTab
        ExportStr &= "<td>結訓人數</td>" & vbTab
        ExportStr &= "<td>就業人數</td>" & vbTab
        ExportStr &= "<td>報名人數</td>" & vbTab
        ExportStr &= "<td>開訓人數</td>" & vbTab
        ExportStr &= "<td>結訓人數</td>" & vbTab
        ExportStr &= "<td>就業人數</td>" & vbTab
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
            ExportStr &= "<td>" & Convert.ToString(dr("orgname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("classcname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("stdate")) & "~" & Convert.ToString(dr("ftdate")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("thours")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID33a")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID33a")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID33b")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID33c")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("ID07a")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID07a")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID07b")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("ID07c")) & "</td>" & vbTab

            ExportStr &= "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next

        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

        Response.End()
    End Sub

End Class

