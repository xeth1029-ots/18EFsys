Partial Class SD_06_002_R
    Inherits AuthBasePage

#Region "Sub"

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)
    End Sub

    '建置匯出資料 (strType 1=>加保, 2=>退保)
    Private Sub crtTable(ByVal strType As String, ByVal dt As DataTable, ByVal strPrt As String)
#Region "建置匯出資料"

        'strPrt xls@Excel檔
        'Dim conn As SqlConnection = DbAccess.GetConnection()
        Dim sda As New SqlDataAdapter
        Dim ds As New DataSet
        Dim sql As String = ""
        Dim tmpDT As New DataTable
        Dim tmpDR As DataRow = Nothing
        Dim dr As DataRow = Nothing

        '建置tmpDT
        Select Case strType
            Case "1"
                tmpDT.Columns.Add(New DataColumn("0")) '異動別
                tmpDT.Columns.Add(New DataColumn("1")) '格式別
                tmpDT.Columns.Add(New DataColumn("2")) '保險證號
                tmpDT.Columns.Add(New DataColumn("3")) '保險證號檢查碼
                tmpDT.Columns.Add(New DataColumn("4")) '被保險人外籍
                tmpDT.Columns.Add(New DataColumn("5")) '被保險人姓名
                tmpDT.Columns.Add(New DataColumn("6")) '被保險人身分證號
                tmpDT.Columns.Add(New DataColumn("7")) '被保險人出生日期
                tmpDT.Columns.Add(New DataColumn("8")) '勞保投保薪資/勞退月提繳工資
                tmpDT.Columns.Add(New DataColumn("9")) '特殊身分別
                tmpDT.Columns.Add(New DataColumn("10")) '勞基法特殊身分別
                tmpDT.Columns.Add(New DataColumn("11")) '已領取社會保險給付種類
                tmpDT.Columns.Add(New DataColumn("12")) '被保險人性別
                tmpDT.Columns.Add(New DataColumn("13")) '提繳身分別
                tmpDT.Columns.Add(New DataColumn("14")) '雇主提繳率
                tmpDT.Columns.Add(New DataColumn("15")) '個人自願提繳率
                tmpDT.Columns.Add(New DataColumn("16")) '勞退提繳日期
                'tmpDR = tmpDT.NewRow
                'tmpDT.Rows.Add(tmpDR)
                'tmpDR("0") = "異動別"
                'tmpDR("1") = "格式別"
                'tmpDR("2") = "保險證號"
                'tmpDR("3") = "保險證號檢查碼"
                'tmpDR("4") = "被保險人外籍"
                'tmpDR("5") = "被保險人姓名"
                'tmpDR("6") = "被保險人身分證號"
                'tmpDR("7") = "被保險人出生日期"
                'tmpDR("8") = "勞保投保薪資/勞退月提繳工資"
                'tmpDR("9") = "特殊身分別"
                'tmpDR("10") = "勞基法特殊身分別"
                'tmpDR("11") = "已領取社會保險給付種類"
                'tmpDR("12") = "被保險人性別"
                'tmpDR("13") = "提繳身分別"
                'tmpDR("14") = "雇主提繳率"
                'tmpDR("15") = "個人自願提繳率"
                'tmpDR("16") = "勞退提繳日期"
            Case "2"
                tmpDT.Columns.Add(New DataColumn("0")) '異動別
                tmpDT.Columns.Add(New DataColumn("1")) '保險證號
                tmpDT.Columns.Add(New DataColumn("2")) '保險證號檢查碼
                tmpDT.Columns.Add(New DataColumn("3")) '被保險人外籍
                tmpDT.Columns.Add(New DataColumn("4")) '被保險人姓名
                tmpDT.Columns.Add(New DataColumn("5")) '被保險人身分證號
                tmpDT.Columns.Add(New DataColumn("6")) '被保險人居留證號碼
                tmpDT.Columns.Add(New DataColumn("7")) '被保險人出生日期
                'tmpDR = tmpDT.NewRow
                'tmpDT.Rows.Add(tmpDR)
                'tmpDR("0") = "異動別"
                'tmpDR("1") = "格式別"
                'tmpDR("2") = "保險證號檢查碼"
                'tmpDR("3") = "被保險人外籍"
                'tmpDR("4") = "被保險人姓名"
                'tmpDR("5") = "被保險人身分證號"
                'tmpDR("6") = "被保險人居留證號碼"
                'tmpDR("7") = "被保險人出生日期"
        End Select

        '建置資料
        Try
            'conn.Open()
            Call TIMS.OpenDbConn(objconn)
            Select Case strType
                Case "1"
                    sql = " SELECT CASE passportno WHEN 2 THEN 'Y' ELSE '' END flag, CASE passportno WHEN 2 THEN sex ELSE '' END sex FROM stud_studentinfo WHERE UPPER(idno) = UPPER(@idno) "
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dr = dt.Rows(i)
                        With sda
                            .SelectCommand = New SqlCommand(sql, objconn)
                            .SelectCommand.Parameters.Clear()
                            .SelectCommand.Parameters.Add("idno", SqlDbType.VarChar).Value = TIMS.ChangeIDNO(dr("idno"))
                            .Fill(ds)
                        End With
                        tmpDR = tmpDT.NewRow
                        tmpDT.Rows.Add(tmpDR)
                        tmpDR("0") = "4"
                        tmpDR("1") = "1"
                        tmpDR("2") = Mid(dr("actno"), 1, 8)
                        tmpDR("3") = Mid(dr("actno"), 9, 1)
                        tmpDR("4") = Convert.ToString(ds.Tables(0).Rows(0)("flag"))
                        tmpDR("5") = dr("name")
                        tmpDR("6") = dr("idno")
                        tmpDR("7") = (Convert.ToInt16(Split(dr("birthday"), "/")(0)) - 1911).ToString.PadLeft(3, "0") & Split(dr("birthday"), "/")(1).PadLeft(2, "0") & Split(dr("birthday"), "/")(2).PadLeft(2, "0")
                        tmpDR("8") = dr("insuresalary")
                        tmpDR("9") = ""
                        tmpDR("10") = ""
                        tmpDR("11") = ""
                        tmpDR("12") = ds.Tables(0).Rows(0)("sex")
                        tmpDR("13") = ""
                        tmpDR("14") = ""
                        tmpDR("15") = ""
                        tmpDR("16") = ""
                        ds.Tables(0).Rows.Clear()
                    Next
                Case "2"
                    sql = " SELECT ppno, CASE passportno WHEN 2 THEN 'Y' ELSE '' END flag, CASE passportno WHEN 2 THEN sex ELSE '' END sex FROM stud_studentinfo WHERE UPPER(idno) = UPPER(@idno) "
                    For i As Integer = 0 To dt.Rows.Count - 1
                        dr = dt.Rows(i)
                        With sda
                            .SelectCommand = New SqlCommand(sql, objconn)
                            .SelectCommand.Parameters.Clear()
                            .SelectCommand.Parameters.Add("idno", SqlDbType.VarChar).Value = TIMS.ChangeIDNO(dr("idno"))
                            .Fill(ds)
                        End With
                        tmpDR = tmpDT.NewRow
                        tmpDT.Rows.Add(tmpDR)
                        tmpDR("0") = "2"
                        tmpDR("1") = Mid(dr("actno"), 1, 8)
                        tmpDR("2") = Mid(dr("actno"), 9, 1)
                        tmpDR("3") = Convert.ToString(ds.Tables(0).Rows(0)("flag"))
                        tmpDR("4") = dr("name")
                        If Convert.ToString(ds.Tables(0).Rows(0)("ppno")) = "2" Then
                            tmpDR("5") = ""
                            tmpDR("6") = dr("idno")
                        Else
                            tmpDR("5") = dr("idno")
                            tmpDR("6") = ""
                        End If
                        tmpDR("7") = (Convert.ToInt16(Split(dr("birthday"), "/")(0)) - 1911).ToString.PadLeft(3, "0") & Split(dr("birthday"), "/")(1).PadLeft(2, "0") & Split(dr("birthday"), "/")(2).PadLeft(2, "0")
                        ds.Tables(0).Rows.Clear()
                    Next
            End Select
            'conn.Close()
            If Not sda Is Nothing Then sda.Dispose()
            If Not ds Is Nothing Then ds.Dispose()
        Catch ex As Exception
            Common.MessageBox(Me, "系統錯誤:" & ex.ToString)
        End Try
        Select Case strPrt
            Case "txt"
                Call exportTXT(strType, tmpDT)
            Case "xls"
                Call exportXLS(strType, tmpDT)
        End Select

#End Region
    End Sub

    '匯出txt
    Private Sub exportTXT(ByVal strType As String, ByVal dt As DataTable)
#Region "匯出txt"

        Dim strValue As String = ""

        Select Case strType
            Case "1"
                If OCID1.Text <> "" Then
                    Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(OCID1.Text, System.Text.Encoding.UTF8) & ".txt")
                Else
                    Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("studentAct", System.Text.Encoding.UTF8) & ".txt")
                End If

                Response.ContentType = "Application/octet-stream"
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

                Common.RespWrite(Me, "異動別,格式別,保險證號,保險證號檢查碼,被保險人外籍,被保險人姓名(外籍含全名),被保險人身分證號,被保險人出生日期,月實際工資(勞保、勞退用),特殊身分別,勞基法特殊身分別,已領取社會保險給付種類,被保險人性別,提繳身分別,雇主提繳率(%),個人自願提繳率(%),勞退提繳日期(與勞保加保日期不同時才需輸入)")
                Common.RespWrite(Me, vbCrLf)

                For Each dr As DataRow In dt.Rows
                    strValue = ""
                    For i As Integer = 0 To dt.Columns.Count - 1
                        If strValue <> "" Then strValue += ","
                        strValue += Convert.ToString(dr(i))
                    Next
                    Common.RespWrite(Me, strValue)
                    Common.RespWrite(Me, vbCrLf)
                Next

                Response.End()

            Case "2"
                Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode(OCID1.Text, System.Text.Encoding.UTF8) & ".txt")
                Response.ContentType = "Application/octet-stream"
                Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

                Common.RespWrite(Me, "異動別,保險證號,保險證號檢查碼,被保險人外籍,被保險人姓名（外籍含全名）,被保險人身分證號(或護照號碼),被保險人居留證號碼,被保險人出生日期")
                Common.RespWrite(Me, vbCrLf)

                For Each dr As DataRow In dt.Rows
                    strValue = ""

                    For i As Integer = 0 To dt.Columns.Count - 1
                        If strValue = "" Then
                            strValue = Convert.ToString(dr(i))
                        Else
                            strValue += "," & Convert.ToString(dr(i))
                        End If
                    Next

                    Common.RespWrite(Me, strValue)
                    Common.RespWrite(Me, vbCrLf)
                Next

                Response.End()
        End Select

#End Region
    End Sub

    '匯出excel
    Private Sub exportXLS(ByVal strType As String, ByVal dt As DataTable)
#Region "匯出excel"

        Dim myCell As TableCell = Nothing
        Dim myRow As TableRow = Nothing
        Dim sKind As String = "border:0.5pt solid #000000"

        For i As Integer = 0 To dt.Rows.Count - 1
            myRow = New TableRow

            For j As Integer = 0 To dt.Columns.Count - 1
                myCell = New TableCell
                myCell.Attributes.Add("style", sKind)
                myCell.HorizontalAlign = HorizontalAlign.Center
                myCell.Attributes("nowrap") = "nowrap"
                myCell.Text = Convert.ToString(dt.Rows(i)(j))

                myRow.Cells.Add(myCell)
            Next

            tbRpt.Rows.Add(myRow)
        Next

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8")

        '提示使用者是否要儲存檔案
        Dim sFileName As String = ""

        If strType = "1" Then
            sFileName = HttpUtility.UrlEncode("加保資料.xls", System.Text.Encoding.UTF8)
        Else
            sFileName = HttpUtility.UrlEncode("退保資料.xls", System.Text.Encoding.UTF8)
        End If

        Response.AddHeader("content-disposition", "attachment; filename=" & sFileName)

        '文件內容指定為excel
        'Response.ContentType = "application/ms-excel;charset=utf-8"
        Response.ContentType = "application/vnd.ms-excel"

        '繪出要輸出的html內容
        Dim strContent As New System.Text.StringBuilder
        Dim stringWrite As New System.IO.StringWriter(strContent)
        Dim htmlWrite As New System.Web.UI.HtmlTextWriter(stringWrite)

        Div1.RenderControl(htmlWrite)
        strContent.Replace("<html>", "")
        strContent.Replace("</html>", "")
        strContent.Replace("<a", "<span")
        strContent.Replace("</a>", "</span>")
        strContent.Replace("<input", "<span")
        Common.RespWrite(Me, "<html><head><meta http-equiv='Content-Type' content='text/html; charset=utf-8'></head>")

        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")
        Common.RespWrite(Me, strContent)
        Common.RespWrite(Me, "</html>")

        '結束程式執行
        Response.End()

#End Region
    End Sub

#End Region

#Region "Function"

    '判斷條件
    Function CheckData1(ByRef Errmsg As String) As Boolean
#Region "判斷條件"

        Dim Rst As Boolean = True
        Errmsg = ""
        Dim sErrMsg1 As String
        sErrMsg1 = "[結訓日期]開始 "
        Call CheckTextBoxDate1(FTDate1, sErrMsg1, Errmsg)
        sErrMsg1 = "[結訓日期]結束 "
        Call CheckTextBoxDate1(FTDate2, sErrMsg1, Errmsg)
        sErrMsg1 = "[加保日期]開始 "
        Call CheckTextBoxDate1(ApplyD1, sErrMsg1, Errmsg)
        sErrMsg1 = "[加保日期]結束 "
        Call CheckTextBoxDate1(ApplyD2, sErrMsg1, Errmsg)
        sErrMsg1 = "[退保日期]開始 "
        Call CheckTextBoxDate1(DropoutD1, sErrMsg1, Errmsg)
        sErrMsg1 = "[退保日期]結束 "
        Call CheckTextBoxDate1(DropoutD2, sErrMsg1, Errmsg)
        If Errmsg <> "" Then Rst = False
        Return Rst

#End Region
    End Function

    '依判斷條件回傳告警內容
    Public Shared Sub CheckTextBoxDate1(ByRef objTextBoxDate1 As TextBox, ByRef sErrMsg1 As String, ByRef Errmsg As String)
#Region "依判斷條件回傳告警內容"

        '累加 Errmsg
        If Trim(objTextBoxDate1.Text) <> "" Then
            objTextBoxDate1.Text = Trim(objTextBoxDate1.Text)

            If Not TIMS.IsDate1(objTextBoxDate1.Text) Then
                Errmsg += sErrMsg1 & "應為日期 格式有誤!" & vbCrLf
            Else
                Try
                    objTextBoxDate1.Text = CDate(objTextBoxDate1.Text).ToString("yyyy/MM/dd")
                Catch ex As Exception
                    Errmsg += sErrMsg1 & "應為日期 格式有誤!" & vbCrLf
                End Try
            End If
        Else
            objTextBoxDate1.Text = ""
        End If

#End Region
    End Sub

    '查詢語法組合送出
    Function Search_Query(Optional ByVal SType As Integer = 1) As String
        'SType=1 匯出加保, SType=2 匯出退保
        Dim Sql As String = ""

        Sql = "" & vbCrLf
        Sql += " SELECT d.OCID, b.Name, b.IDNO, b.Birthday, c.InsureSalary, g.ActNo " & vbCrLf
        Sql += " FROM Class_StudentsOfClass a " & vbCrLf
        Sql += " JOIN Stud_StudentInfo b ON a.SID = b.SID " & vbCrLf
        Sql += " JOIN Stud_Insurance c ON a.SOCID = c.SOCID " & vbCrLf
        Sql += " JOIN Class_ClassInfo d ON d.OCID = a.OCID " & vbCrLf
        Sql += " JOIN Auth_Relship e ON e.RID = d.RID " & vbCrLf
        Sql += " JOIN Org_OrgInfo f ON e.OrgID = f.OrgID " & vbCrLf
        Sql += " JOIN Org_OrgPlanInfo g ON e.RSID = g.RSID " & vbCrLf
        Sql += " WHERE 1=1 " & vbCrLf

        If SType = 2 Then
            '增加判斷避免無加退保資料的人也一並列出
            Sql += " AND c.insuresalary is not null" & vbCrLf
        End If

        If sm.UserInfo.RID = "A" Then
            Sql += " AND d.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID = '" & sm.UserInfo.TPlanID & "' AND Years = '" & sm.UserInfo.Years & "') " & vbCrLf
        Else
            Sql += " AND d.PlanID = '" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If

        If Me.OCIDValue1.Value <> "" Then
            Sql += " AND a.OCID = '" & Me.OCIDValue1.Value & "' " & vbCrLf
            Sql += " AND d.OCID = '" & Me.OCIDValue1.Value & "' " & vbCrLf
        End If

        If Me.cjobValue.Value <> "" Then
            Sql += " AND d.CJOB_UNKEY = '" & Me.cjobValue.Value & "' " & vbCrLf
        End If

        If FTDate1.Text <> "" Then
            Sql += " AND d.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf
        End If

        If FTDate2.Text <> "" Then
            Sql += " AND d.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf
        End If

        If ApplyD1.Text <> "" Then
            Sql += " AND c.ApplyInsurance >= " & TIMS.To_date(ApplyD1.Text) & vbCrLf
        End If

        If ApplyD2.Text <> "" Then
            Sql += " AND c.ApplyInsurance <= " & TIMS.To_date(ApplyD2.Text) & vbCrLf
        End If

        If DropoutD1.Text <> "" Then
            Sql += " AND c.DropoutInsurance >= " & TIMS.To_date(DropoutD1.Text) & vbCrLf
        End If

        If DropoutD2.Text <> "" Then
            Sql += " AND c.DropoutInsurance <= " & TIMS.To_date(DropoutD2.Text) & vbCrLf
        End If
        '是否為在職者補助身分 46:補助辦理保母職業訓練'47:補助辦理照顧服務員職業訓練
        If TIMS.Cst_TPlanID46AppPlan5.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            '排除在職者補助身分
            Sql += " AND ( ISNULL(a.WorkSuppIdent,' ') !='Y') " & vbCrLf
        End If

        'Sql += " ORDER BY d.OCID, SUBSTRING(, 0, Len(a.StudentID)) ASC " & vbCrLf
        Sql += " ORDER BY d.OCID, a.StudentID ASC " & vbCrLf

        Return Sql
    End Function

#End Region

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
#Region "Page_Load"

        '檢查Session是否存在
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not Page.IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            Dim sql As String = ""
            Dim dt As DataTable = DbAccess.GetDataTable("SELECT * FROM Key_Plan", objconn)
            Dim dr As DataRow = Nothing

            With TPlan
                .DataSource = dt
                .DataTextField = "PlanName"
                .DataValueField = "TPlanID"
                .DataBind()
                .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
            End With

            dr = DbAccess.GetOneRow("SELECT * FROM ID_Plan WHERE PlanID = '" & sm.UserInfo.PlanID & "' ", objconn)

            If Not dr Is Nothing Then Common.SetListItem(TPlan, dr("TplanID"))

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
            Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

            Button1.Attributes("onclick") = "javascript:return print();"
            'Button3.Attributes("onclick") = "javascript:return searchcheck();"
            'Button4.Attributes("onclick") = "javascript:return searchcheck();"
            btnExport1.Attributes("onclick") = "javascript:return searchcheck();"
            btnExport2.Attributes("onclick") = "javascript:return searchcheck();"
            'Button5_Click(sender, e)
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            'center.Attributes("onclick") = "showObj('HistoryList2');showFrame('inline');"
            center.Attributes("onclick") = "showObj('HistoryList2');showFrame('');"
            center.Style("CURSOR") = "hand"
        End If
        HistoryRID.Attributes("onclick") = "showFrame('none');"
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

#End Region
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
#Region "Button1_Click"

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim MyValue As String = ""
        MyValue = "OCID=" & Me.OCIDValue1.Value
        MyValue += "&TMID=" & Me.TMIDValue1.Value
        MyValue += "&TPlanID=" & Me.TPlan.SelectedValue
        MyValue += "&Relship=" & RelShip
        MyValue += "&CJOB_UNKEY=" & Me.cjobValue.Value
        MyValue += "&FTDate1=" & Me.FTDate1.Text
        MyValue += "&FTDate2=" & Me.FTDate2.Text
        MyValue += "&ApplyD1=" & Me.ApplyD1.Text
        MyValue += "&ApplyD2=" & Me.ApplyD2.Text
        MyValue += "&DropoutD1=" & Me.DropoutD1.Text
        MyValue += "&DropoutD2=" & Me.DropoutD2.Text

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Member", "insurance_list", MyValue)

#End Region
    End Sub

    '加保xls
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
#Region "加保xls"

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
        Else
            Dim dt As DataTable = DbAccess.GetDataTable(Search_Query(1), objconn)
            If dt.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無學員加保資料")
                lblMsg.Text = "查無學員加保資料"
            Else
                Call crtTable("1", dt, "xls")
                lblMsg.Text = ""
            End If
        End If

#End Region
    End Sub

    '退保xls
    Private Sub btnExport2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport2.Click
#Region "退保xls"

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
        Else
            Dim dt As DataTable = DbAccess.GetDataTable(Search_Query(2), objconn)
            If dt.Rows.Count = 0 Then
                Common.MessageBox(Me, "查無學員退保資料")
                lblMsg.Text = "查無學員退保資料"
            Else
                crtTable("2", dt, "xls")
                lblMsg.Text = ""
            End If
        End If

#End Region
    End Sub
End Class