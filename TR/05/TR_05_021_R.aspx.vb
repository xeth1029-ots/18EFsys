Public Class TR_05_021_R
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

#Region "NO USE"
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

    '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report2011", "TR_05_021_R", MyValue)
    'End Sub

#End Region

    '匯出
    Protected Sub Export1_Click(sender As Object, e As EventArgs) Handles Export1.Click
        Dim dt As DataTable
        dt = LoadData1()
        Call ExpReport1(dt)
    End Sub

    '匯出(SQL)
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
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select cc.orgname" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,cc.stdate" & vbCrLf
        sql &= " ,cc.FTDate" & vbCrLf
        sql &= " ,cc.thours" & vbCrLf
        sql &= " ,cc.ctname" & vbCrLf
        sql &= " ,iz.zipname" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " JOIN ID_ZIP iz on iz.zipcode=cc.taddresszip" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        'sql &= " and cc.years ='2016'" & vbCrLf
        'sql &= " and cc.DistID IN ('001')" & vbCrLf
        'sql &= " and cc.TPlanID IN ('01','02')" & vbCrLf
        'sql &= " AND cc.FTDate<= convert(datetime, '2016/12/30', 111)" & vbCrLf
        'sql &= " AND ROWNUM <=10" & vbCrLf
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
            sql += " AND cc.FTDate>= " & TIMS.To_date(FTDate1.Text) & vbCrLf '" & FTDate1.Text & "'" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql += " AND cc.FTDate<= " & TIMS.To_date(FTDate2.Text) & vbCrLf '" & FTDate2.Text & "'" & vbCrLf
        End If
        sql &= " )" & vbCrLf

        sql &= " select cc.orgname" & vbCrLf
        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.stdate, 111) stdate" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.FTDate, 111) ftdate" & vbCrLf
        sql &= " ,cc.thours" & vbCrLf
        sql &= " ,cc.ctname" & vbCrLf
        sql &= " ,cc.zipname" & vbCrLf
        sql &= " ,cc.ocid" & vbCrLf
        sql &= " ,dbo.NVL(A1,0) A1" & vbCrLf
        sql &= " ,dbo.NVL(A2,0) A2" & vbCrLf
        sql &= " ,dbo.NVL(A3,0) A3" & vbCrLf
        sql &= " ,dbo.NVL(A4,0) A4" & vbCrLf
        sql &= " ,dbo.NVL(B1,0) B1" & vbCrLf
        sql &= " ,dbo.NVL(B2,0) B2" & vbCrLf
        sql &= " ,dbo.NVL(B3,0) B3" & vbCrLf
        sql &= " ,dbo.NVL(B4,0) B4" & vbCrLf
        sql &= " ,dbo.NVL(C1,0) C1" & vbCrLf
        sql &= " ,dbo.NVL(C2,0) C2" & vbCrLf
        sql &= " ,dbo.NVL(C3,0) C3" & vbCrLf
        sql &= " ,dbo.NVL(C4,0) C4" & vbCrLf
        sql &= " ,dbo.NVL(D1,0) D1" & vbCrLf
        sql &= " ,dbo.NVL(D2,0) D2" & vbCrLf
        sql &= " ,dbo.NVL(D3,0) D3" & vbCrLf
        sql &= " ,dbo.NVL(D4,0) D4" & vbCrLf
        sql &= " ,dbo.NVL(P1,0) P1" & vbCrLf
        sql &= " ,dbo.NVL(P2,0) P2" & vbCrLf
        sql &= " ,dbo.NVL(P3,0) P3" & vbCrLf
        sql &= " ,dbo.NVL(P4,0) P4" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        sql &= " LEFT join (" & vbCrLf
        sql &= " select cc.ocid" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID !='05' then 1 end) A1" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID !='05' then 1 end) A2" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID ='05' then 1 end) A3" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID ='05' then 1 end) A4" & vbCrLf

        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) then 1 end) B1" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) then 1 end) B2" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) then 1 end) B3" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) then 1 end) B4" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.socid is not null then 1 end) C1" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.socid is not null then 1 end) C2" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.socid is not null then 1 end) C3" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.socid is not null then 1 end) C4" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.JOBRELATE='Y' then 1 end) D1" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.JOBRELATE='Y' then 1 end) D2" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.JOBRELATE='Y' then 1 end) D3" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.JOBRELATE='Y' then 1 end) D4" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.PUBLICRESCUE='Y' then 1 end) P1" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID !='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.PUBLICRESCUE='Y' then 1 end) P2" & vbCrLf
        sql &= " ,SUM(case when ss.sex='M' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) and sg3.PUBLICRESCUE='Y' then 1 end) P3" & vbCrLf
        sql &= " ,SUM(case when ss.sex='F' AND cs.MIdentityID ='05' and cs.studstatus not in (2,3) AND cc.FTDate<=dbo.TRUNC_DATETIME(getdate()) AND sg3.PUBLICRESCUE='Y' then 1 end) P4" & vbCrLf
        sql &= " from WC1 cc" & vbCrLf
        sql &= " JOIN class_studentsofclass cs on cc.ocid =cs.ocid" & vbCrLf
        sql &= " join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        sql &= " join stud_subdata ss2 on ss2.sid =cs.sid" & vbCrLf
        sql &= " left join Stud_GetJobState3 sg3 on sg3.CPoint=1 and sg3.socid =cs.socid and sg3.IsGetJob='1'" & vbCrLf
        'sql += " 	left join Stud_GetJobState3 sg9 on sg9.socid =cs.socid and sg9.CPoint= 9" & vbCrLf
        sql &= " group by cc.ocid" & vbCrLf
        sql &= " ) g on g.ocid =cc.ocid" & vbCrLf

        rst = DbAccess.GetDataTable(sql, objconn)
        Return rst
    End Function

    '匯出(OUTPUT XLS)
    Sub ExpReport1(ByRef dt As DataTable)
        Dim strTitle1 As String = "原住民職訓措施辦理情形"

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
        ExportStr &= "<td rowspan=""3"">序號</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">訓練機構</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">班別名稱</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">訓練起迄</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">時數</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">縣市</td>" & vbTab
        ExportStr &= "<td rowspan=""3"">鄉、鎮(區)</td>" & vbTab
        ExportStr &= "<td colspan=""4"">參訓人數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">結訓人數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">就業人數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">就業關聯人數</td>" & vbTab
        ExportStr &= "<td colspan=""4"">公法救助</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        '第2行
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td colspan=""2"">非原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">非原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">非原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">非原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">非原住民</td>" & vbTab
        ExportStr &= "<td colspan=""2"">原住民</td>" & vbTab
        ExportStr &= "</tr>" & vbCrLf
        '第3行
        ExportStr &= "<tr>" & vbCrLf
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
        ExportStr &= "<td>男性</td>" & vbTab
        ExportStr &= "<td>女性</td>" & vbTab
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
            ExportStr &= "<td>" & Convert.ToString(dr("ctname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("zipname")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("A4")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("B1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("B4")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("C1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("C4")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("D1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("D4")) & "</td>" & vbTab

            ExportStr &= "<td>" & Convert.ToString(dr("P1")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("P2")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("P3")) & "</td>" & vbTab
            ExportStr &= "<td>" & Convert.ToString(dr("P4")) & "</td>" & vbTab

            ExportStr &= "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

        TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub
End Class