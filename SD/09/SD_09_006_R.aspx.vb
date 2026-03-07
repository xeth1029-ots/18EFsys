Partial Class SD_09_006_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SchoolBegins_List"
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        '檢查Session是否存在 End
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
        End If

        Button1.Attributes("onclick") = "javascript:return print();"

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim PlanID As String = sm.UserInfo.PlanID
        Dim TPlanID As String = sm.UserInfo.TPlanID
        If sm.UserInfo.RID = "A" Then
            PlanID = ""
            TPlanID = sm.UserInfo.TPlanID
        End If

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        start_date2.Text = TIMS.Cdate3(start_date2.Text)
        end_date2.Text = TIMS.Cdate3(end_date2.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)

        Dim MyValue1 As String = "k=r"
        TIMS.SetMyValue(MyValue1, "TPlanID", TPlanID)
        TIMS.SetMyValue(MyValue1, "PlanID", PlanID)

        TIMS.SetMyValue(MyValue1, "start_date", start_date.Text)
        TIMS.SetMyValue(MyValue1, "end_date", end_date.Text)
        TIMS.SetMyValue(MyValue1, "start_date2", start_date2.Text)
        TIMS.SetMyValue(MyValue1, "end_date2", end_date2.Text)
        TIMS.SetMyValue(MyValue1, "CJOB_UNKEY", cjobValue.Value)
        If Me.RadioButton1.Checked = True Then
            TIMS.SetMyValue(MyValue1, "Parameter", RIDValue.Value)
        Else
            TIMS.SetMyValue(MyValue1, "RID", RIDValue.Value)
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue1)
    End Sub

    Function Utl_GetDataTable() As DataTable
        Dim PlanID As String = sm.UserInfo.PlanID
        Dim TPlanID As String = sm.UserInfo.TPlanID
        If sm.UserInfo.RID = "A" Then
            PlanID = ""
            TPlanID = sm.UserInfo.TPlanID
        End If
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        start_date2.Text = TIMS.Cdate3(start_date2.Text)
        end_date2.Text = TIMS.Cdate3(end_date2.Text)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)

        Dim parms As New Hashtable
        parms.Clear()
        If PlanID <> "" Then parms.Add("PlanID", PlanID)
        If TPlanID <> "" Then parms.Add("TPlanID", TPlanID)
        If start_date.Text <> "" Then parms.Add("STDate1", TIMS.Cdate2(start_date.Text))
        If end_date.Text <> "" Then parms.Add("STDate2", TIMS.Cdate2(end_date.Text))
        If start_date2.Text <> "" Then parms.Add("FTDate1", TIMS.Cdate2(start_date2.Text))
        If end_date2.Text <> "" Then parms.Add("FTDate2", TIMS.Cdate2(end_date2.Text))
        If cjobValue.Value <> "" Then parms.Add("CJOB_UNKEY", cjobValue.Value)
        parms.Add("RID", RIDValue.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= "  SELECT cc.OCID" & vbCrLf
        sql &= "  ,cc.CLASSCNAME2" & vbCrLf
        sql &= "  ,cc.THOURS" & vbCrLf
        sql &= "  ,cc.STDATE" & vbCrLf
        sql &= "  ,cc.FTDATE" & vbCrLf
        sql &= "  ,cc.ORGNAME" & vbCrLf
        sql &= "  ,cc.TPLANID" & vbCrLf
        sql &= "  ,cc.PLANID" & vbCrLf
        sql &= "  ,cc.YEARS" & vbCrLf
        sql &= "  ,cc.ORGID" & vbCrLf
        sql &= "  ,cc.RID" & vbCrLf
        sql &= "  ,cc.TaddressZip" & vbCrLf
        sql &= "  ,cc.TADDRESSZIP6W" & vbCrLf
        sql &= "  ,cc.TAddress" & vbCrLf
        sql &= "  ,cc.PLANNAME" & vbCrLf
        sql &= "  FROM dbo.VIEW2 cc" & vbCrLf
        sql &= "  where 1=1" & vbCrLf
        If PlanID <> "" Then sql &= " AND cc.PlanID = @PlanID" & vbCrLf
        If TPlanID <> "" Then sql &= " AND cc.TPlanID = @TPlanID" & vbCrLf
        If start_date.Text <> "" Then sql &= " AND cc.STDate >= @STDate1" & vbCrLf
        If end_date.Text <> "" Then sql &= " AND cc.STDate <= @STDate2" & vbCrLf
        If start_date2.Text <> "" Then sql &= " AND cc.FTDate >= @FTDate1" & vbCrLf
        If end_date2.Text <> "" Then sql &= " AND cc.FTDate <= @FTDate2" & vbCrLf
        If cjobValue.Value <> "" Then sql &= " AND cc.CJOB_UNKEY = @CJOB_UNKEY" & vbCrLf
        If Me.RadioButton1.Checked = True Then
            sql &= " AND cc.RID LIKE @RID+'%'" & vbCrLf
        Else
            sql &= " AND cc.RID=@RID" & vbCrLf
        End If
        sql &= " )" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY a.STDate ASC) ROWNUM" & vbCrLf
        sql &= " ,a.CLASSCNAME2" & vbCrLf
        'sql &= " ,a.CLASSCNAME2 CLASSCNAME" & vbCrLf
        'sql &= " ,CONVERT(varchar,a.STDATE,111) STDATE" & vbCrLf
        'sql &= " ,CONVERT(varchar,a.FTDATE,111) FTDATE" & vbCrLf
        sql &= " ,FORMAT(a.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf '開訓日期
        sql &= " ,FORMAT(a.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf '結訓日期
        '上課起迄日
        'sql &= " ,concat(FORMAT(a.STDATE,'yyyy/MM/dd'),'~',FORMAT(a.FTDATE,'yyyy/MM/dd')) SFTDATE" & vbCrLf
        sql &= " ,a.ORGNAME" & vbCrLf
        sql &= " ,concat(dbo.FN_GET_ZIPCODE(a.TaddressZip,a.TADDRESSZIP6W),a.TAddress) TADD" & vbCrLf
        sql &= " ,o.PHONE" & vbCrLf
        sql &= " ,a.THOURS" & vbCrLf
        sql &= " ,a.TPLANID" & vbCrLf
        'sql &= " ,CONVERT(VARCHAR, a.YEARS) + '年度' + a.PLANNAME + '　開課一覽表' AS PLANNAME" & vbCrLf
        sql &= " ,a.PLANID" & vbCrLf
        sql &= " ,ISNULL(cs.IN_CLASS, 0) IN_CLASS" & vbCrLf
        sql &= " ,ISNULL(cs.TOTAL, 0) TOTAL" & vbCrLf
        sql &= " ,ISNULL(cs.END_TOTAL, 0) END_TOTAL" & vbCrLf
        sql &= " ,ISNULL(cs.return_total, 0) RETURN_TOTAL" & vbCrLf
        sql &= " ,'' REMARKS" & vbCrLf
        sql &= " FROM WC1 a" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP d ON d.RID = a.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGPLANINFO O ON O.RSID=d.RSID" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= "  SELECT cs.OCID" & vbCrLf
        sql &= "  ,COUNT(1) TOTAL" & vbCrLf
        sql &= "  ,COUNT(CASE WHEN cs.StudStatus = 1 THEN 1 END ) IN_CLASS" & vbCrLf
        sql &= "  ,COUNT(CASE WHEN cs.StudStatus = 5 THEN 1 END ) END_TOTAL" & vbCrLf
        sql &= "  ,COUNT(CASE WHEN cs.StudStatus IN (2,3) THEN 1 END ) RETURN_TOTAL" & vbCrLf
        sql &= "  FROM WC1 cc" & vbCrLf
        sql &= "  JOIN dbo.CLASS_STUDENTSOFCLASS cs ON cs.OCID = cc.OCID" & vbCrLf
        sql &= "  JOIN ID_Plan ip on ip.PlanID = cc.PlanID" & vbCrLf
        sql &= "  GROUP BY cs.OCID" & vbCrLf
        sql &= " ) cs ON cs.OCID = a.OCID" & vbCrLf
        sql &= " ORDER BY a.STDate" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    Sub Utl_EXPORT1()
        Dim oTable As DataTable = Utl_GetDataTable()

        msg1.Text = "查無資料!!"
        If oTable Is Nothing Then Return
        If oTable.Rows.Count = 0 Then Return

        msg1.Text = ""

        '委員,應出席次數,實際出席次數,出席率,委員,應出席次數,實際出席次數,出席率

        Const s_title1 As String = "項次,訓練機構,班別,上課地點,電話,訓練時數,開訓日期,結訓日期,開訓人數,結訓人數,在訓人數,退訓人數,備註"
        Const s_data1 As String = "ROWNUM,ORGNAME,CLASSCNAME2,TADD,PHONE,THOURS,STDATE,FTDATE,TOTAL,END_TOTAL,IN_CLASS,RETURN_TOTAL,REMARKS"
        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")
        Dim s_colspan As String = As_title1.Length.ToString()

        Dim sFileName1 As String = "開課一覽表" & TIMS.GetDateNo2()
        Dim s_titleA1 As String = "開課一覽表"
        '套CSS值 'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim ExportStr As String '建立輸出文字
        Dim sbHTML As System.Text.StringBuilder = New System.Text.StringBuilder

        'Dim strHTML As String = ""
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'titleA1
        ExportStr = String.Format("<tr><td colspan=""{0}"">{1}</td></tr>", s_colspan, s_titleA1) '& vbTab
        sbHTML.Append(ExportStr)
        'title1
        ExportStr = "<tr>"
        For Each s_T1 As String In As_title1
            ExportStr &= "<td>" & s_T1 & "</td>"   '& vbTab
        Next
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        '建立資料面
        Dim i_num As Integer = 0
        For Each oDr1 As DataRow In oTable.Rows
            i_num += 1
            ExportStr = "<tr>"
            For Each s_D1 As String In As_data1
                ExportStr &= "<td>" & TIMS.ClearSQM(oDr1(s_D1)) & "</td>"
            Next
            ExportStr &= "</tr>"
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")
        oTable = Nothing

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    Protected Sub btnExp1_Click(sender As Object, e As EventArgs) Handles btnExp1.Click
        Call Utl_EXPORT1()
    End Sub
End Class
