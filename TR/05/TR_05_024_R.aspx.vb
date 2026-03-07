Partial Class TR_05_024_R
    Inherits AuthBasePage

    'Const cst_printFN1 As String = "TR_05_024_R"

    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        If Not IsPostBack Then
            CreateItem()
        End If

    End Sub

    Sub CreateItem()
        'Button1.Attributes("onclick") = "return search();"
        SYearList = TIMS.GetSyear(SYearList)
        Common.SetListItem(SYearList, sm.UserInfo.Years)

        DistIDList = TIMS.Get_DistID(DistIDList) '轄區
        DistIDList.Items.Insert(0, New ListItem("全部", ""))

        TPlanIDList = TIMS.Get_TPlan(TPlanIDList, , 1, "Y")

        'https://jira.turbotech.com.tw/browse/TIMSC-157
        DistIDList.Enabled = True
        If CStr(sm.UserInfo.DistID) <> "000" Then
            Common.SetListItem(DistIDList, CStr(sm.UserInfo.DistID))
            DistIDList.Enabled = False
        End If

        Dim flagLID23 As Boolean = TIMS.Chk_Relship23(Me, objconn)
        '當分署(中心)使用者使用時,轄區應該都要鎖死該轄區,不可選擇其它轄區
        HIDOrgID.Value = ""
        Select Case sm.UserInfo.LID '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
            Case "0"
                '是本署  '完全不鎖定
            Case "1"
                '是分署
                Common.SetListItem(DistIDList, CStr(sm.UserInfo.DistID))
                DistIDList.Enabled = False
            Case "2"
                '是補助地方政府 / 一般機構
                Common.SetListItem(DistIDList, CStr(sm.UserInfo.DistID))
                DistIDList.Enabled = False
                Common.SetListItem(SYearList, sm.UserInfo.Years)
                SYearList.Enabled = False '年度

                TIMS.SetCblValue(TPlanIDList, sm.UserInfo.TPlanID)
                TPlanIDList.Enabled = False '計畫
                TPlanIDList.Style.Item("display") = "none"
                HIDOrgID.Value = sm.UserInfo.OrgID
            Case Else
                '是補助地方政府
                Common.SetListItem(DistIDList, CStr(sm.UserInfo.DistID))
                DistIDList.Enabled = False
                Common.SetListItem(SYearList, sm.UserInfo.Years)
                SYearList.Enabled = False '年度
                TIMS.SetCblValue(TPlanIDList, sm.UserInfo.TPlanID)
                TPlanIDList.Enabled = False '計畫
                TPlanIDList.Style.Item("display") = "none"
                HIDOrgID.Value = sm.UserInfo.OrgID
                'DistrictList.Style.Item("display") = "none"
        End Select

    End Sub

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    Dim stitle As String = ""
    '    Dim etitle As String = ""
    '    If STDate1.Text <> "" Or STDate2.Text <> "" Then
    '        stitle = STDate1.Text + " ~ " + STDate2.Text
    '    End If
    '    If FTDate1.Text <> "" Or FTDate2.Text <> "" Then
    '        etitle = FTDate1.Text + " ~ " + FTDate2.Text
    '    End If
    '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, "Years1=" & Syear.SelectedValue & "&STDate1=" & STDate1.Text & "&STDate2=" & STDate2.Text & "&FTDate1=" & FTDate1.Text & "&FTDate2=" & FTDate2.Text & "&stitle=" & stitle & "&etitle=" & etitle)
    'End Sub

    Function GetExport2dt() As DataTable
        Dim dt As New DataTable
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (SELECT RTReasonID,Reason FROM Key_RejectTReason)" & vbCrLf
        sql &= " ,WC2 AS (SELECT CODE_ID,CODE_CNAME FROM SYS_SHAREDCODE where 1=1 AND CODE_KIND='SD_12_008_LB1')" & vbCrLf

        sql &= " select cc.years" & vbCrLf
        sql &= " ,cc.planname" & vbCrLf
        sql &= " ,cc.distname" & vbCrLf
        sql &= " ,cc.orgname2" & vbCrLf
        sql &= " ,cc.orgname" & vbCrLf
        sql &= " ,cc.ORGTYPENAME" & vbCrLf

        sql &= " ,cc.classcname" & vbCrLf
        sql &= " ,tt.busname" & vbCrLf
        sql &= " ,tt.jobname" & vbCrLf
        sql &= " ,tt.trainname" & vbCrLf
        sql &= " ,cc.taddresszip" & vbCrLf
        sql &= " ,cc.ZName" & vbCrLf
        sql &= " ,cc.CTName" & vbCrLf
        sql &= " ,CC.THOURS" & vbCrLf
        sql &= " ,cc.tnum" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.STDATE, 111) STDATE" & vbCrLf
        sql &= " ,CONVERT(varchar, cc.FTDATE, 111) FTDATE" & vbCrLf
        'sql &= " --學員姓名、性別、年齡、學歷、主要身分別、戶籍地縣市、通訊地縣市、預算別、是否為在職者、" & vbCrLf
        sql &= " ,cs.name" & vbCrLf
        sql &= " ,cs.sex2" & vbCrLf
        sql &= " ,cs.yearsold" & vbCrLf
        sql &= " ,cs.degreename" & vbCrLf '教育程度/學歷
        sql &= " ,cs.miname" & vbCrLf
        sql &= " ,cs.ZIPNAME2" & vbCrLf
        sql &= " ,cs.ZIPNAME" & vbCrLf
        sql &= " ,cs.BUDGETIDN" & vbCrLf
        sql &= " ,cs.WORKSUPPIDENT" & vbCrLf
        'sql &= " --是否申請生活津貼、離退訓-離訓原因、離退訓-是否為適應期內離退訓、離退訓-退訓原因" & vbCrLf
        'sql &= " --、是否結訓、" & vbCrLf
        sql &= " ,CASE WHEN c.socid is not null then 'Y' END SubSidyY" & vbCrLf
        sql &= " ,CASE WHEN cs.REJECTTDATE1 IS NOT NULL THEN r1.Reason end Reason1" & vbCrLf
        sql &= " ,CASE WHEN cs.RejectDayIn14='Y' THEN 'Y' END RejectDayIn14" & vbCrLf
        sql &= " ,CASE WHEN cs.REJECTTDATE2 IS NOT NULL THEN r1.Reason end Reason2" & vbCrLf
        sql &= " ,cs.STUDSTATUS2" & vbCrLf
        'sql &= " --是否就業、就業狀態、是否為公法救助就業、就業單位名稱、到職日、是否有就業關連性、" & vbCrLf
        'sql &= " --就業關連性原因1、" & vbCrLf
        'sql &= " --就業關連性原因2、就業長度是否超過1個月以上" & vbCrLf
        sql &= " ,case when sg3.IsGetJob=1 then 'Y' END  IsGetJob" & vbCrLf
        sql &= " ,dbo.DECODE9(sg3.IsGetJob,0,'未就業',1,'就業',2,'不就業', '不明') JobStatus" & vbCrLf
        sql &= " ,sg3.PUBLICRESCUE" & vbCrLf
        sql &= " ,sg3.BUSNAME COMPNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, sg3.MDATE, 111) MDATE" & vbCrLf
        sql &= " ,sg3.JOBRELATE" & vbCrLf
        sql &= " ,CASE WHEN SG3.JOBRELATE_Y LIKE '%01%' THEN 'Y' END JOBRELATE_Y1" & vbCrLf
        sql &= " ,CASE WHEN SG3.JOBRELATE_Y LIKE '%02%' THEN 'Y' END JOBRELATE_Y2" & vbCrLf
        sql &= " ,case when gd.DAYS2>=30 then '是' end DAYS2" & vbCrLf
        'sql &= " --,cs.r.re" & vbCrLf
        sql &= " FROM VIEW2 cc" & vbCrLf
        sql &= " JOIN VIEW_TRAINTYPE tt on tt.tmid=cc.tmid" & vbCrLf
        sql &= " JOIN V_STUDENTINFO cs on cs.ocid =cc.ocid" & vbCrLf
        sql &= " LEFT JOIN SUB_SUBSIDYAPPLY c on c.socid=cs.socid" & vbCrLf
        sql &= " LEFT JOIN WC1 r1 on r1.RTReasonID=cs.RTReasonID" & vbCrLf
        sql &= " LEFT JOIN STUD_GETJOBSTATE3 sg3 on sg3.socid =cs.socid and sg3.CPoint=1" & vbCrLf
        sql &= " LEFT JOIN STUD_GETJOBBYDAYS gd on gd.socid =cs.socid AND gd.OCID=cc.ocid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        'sql &= " and cc.tplanid ='17'" & vbCrLf
        'sql &= " and cc.years ='2016'" & vbCrLf
        Dim vTPLANID As String = TIMS.GetCblValue(TPlanIDList)
        vTPLANID = TIMS.CombiSQM2IN(vTPLANID)
        Dim vDISTID As String = TIMS.GetCblValue(DistIDList)
        vDISTID = TIMS.CombiSQM2IN(vDISTID)

        HIDOrgID.Value = TIMS.ClearSQM(HIDOrgID.Value)
        If HIDOrgID.Value <> "" Then
            '是補助地方政府
            sql &= " and cc.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " and cc.RID2='" & sm.UserInfo.RID & "'" & vbCrLf
            sql &= " and cc.OrgID2='" & HIDOrgID.Value & "'" & vbCrLf
        End If

        If vTPLANID <> "" Then sql &= " and cc.TPLANID IN (" & vTPLANID & ")" & vbCrLf
        If vDISTID <> "" Then sql &= " and cc.DISTID IN (" & vDISTID & ")" & vbCrLf
        sql &= " and cc.YEARS =@YEARS" & vbCrLf

        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        FTDate1.Text = TIMS.ClearSQM(FTDate1.Text)
        FTDate2.Text = TIMS.ClearSQM(FTDate2.Text)
        STDate1.Text = TIMS.Cdate3(STDate1.Text)
        STDate2.Text = TIMS.Cdate3(STDate2.Text)
        FTDate1.Text = TIMS.Cdate3(FTDate1.Text)
        FTDate2.Text = TIMS.Cdate3(FTDate2.Text)
        If STDate1.Text <> "" Then sql &= " and cc.STDATE >=@STDate1" & vbCrLf
        If STDate2.Text <> "" Then sql &= " and cc.STDATE <=@STDate2" & vbCrLf
        If FTDate1.Text <> "" Then sql &= " and cc.FTDATE >=@FTDate1" & vbCrLf
        If FTDate2.Text <> "" Then sql &= " and cc.FTDATE <=@FTDate2" & vbCrLf

        'sql &= " and rownum <=10" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("YEARS", SqlDbType.VarChar).Value = SYearList.SelectedValue
            If STDate1.Text <> "" Then
                .Parameters.Add("STDate1", SqlDbType.DateTime).Value = TIMS.Cdate2(STDate1.Text)
            End If
            If STDate2.Text <> "" Then
                .Parameters.Add("STDate2", SqlDbType.DateTime).Value = TIMS.Cdate2(STDate2.Text)
            End If
            If FTDate1.Text <> "" Then
                .Parameters.Add("FTDate1", SqlDbType.DateTime).Value = TIMS.Cdate2(FTDate1.Text)
            End If
            If FTDate2.Text <> "" Then
                .Parameters.Add("FTDate2", SqlDbType.DateTime).Value = TIMS.Cdate2(FTDate2.Text)
            End If
            dt.Load(.ExecuteReader())
        End With
        'Call CloseDbConn(conn)
        'If dt.Rows.Count > 0 Then Rst = Convert.ToString(dt.Rows(0)("?"))
        Return dt
    End Function

    Sub ExpReport2(ByRef dt As DataTable)
        'If dt.Rows.Count = 0 Then
        'End If
        Dim strTitle1 As String = "地方政府績效班級明細表"

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
        ExportStr &= "<td>年度" & "</td>"
        ExportStr &= "<td>計畫名稱" & "</td>"
        ExportStr &= "<td>轄區" & "</td>"
        ExportStr &= "<td>地方政府" & "</td>"
        ExportStr &= "<td>訓練單位" & "</td>"
        ExportStr &= "<td>單位屬性" & "</td>"
        ExportStr &= "<td>班級名稱" & "</td>"
        ExportStr &= "<td>訓練職類-大類" & "</td>"
        ExportStr &= "<td>訓練職類-中類" & "</td>"
        ExportStr &= "<td>訓練職類-小類" & "</td>"
        ExportStr &= "<td>辦訓地點(縣市鄉鎮區)-郵遞區號" & "</td>"
        ExportStr &= "<td>辦訓地點(縣市鄉鎮區)-中文名稱" & "</td>"
        ExportStr &= "<td>單位立案縣市" & "</td>" '單位所在地縣市/單位立案縣市
        ExportStr &= "<td>訓練時數" & "</td>"
        ExportStr &= "<td>招生人數" & "</td>"
        ExportStr &= "<td>開訓日期" & "</td>"
        ExportStr &= "<td>結訓日期" & "</td>"
        ExportStr &= "<td>學員姓名" & "</td>"
        ExportStr &= "<td>性別" & "</td>"
        ExportStr &= "<td>年齡" & "</td>"
        ExportStr &= "<td>教育程度" & "</td>" '學歷/教育程度
        ExportStr &= "<td>主要身分別" & "</td>"
        ExportStr &= "<td>戶籍地縣市" & "</td>"
        ExportStr &= "<td>通訊地縣市" & "</td>"
        ExportStr &= "<td>預算別" & "</td>"
        ExportStr &= "<td>是否為在職者" & "</td>"
        ExportStr &= "<td>是否申請職訓生活津貼" & "</td>" '是否申請生活津貼/是否申請職訓生活津貼

        ExportStr &= "<td>是否結訓" & "</td>" '是否結訓
        ExportStr &= "<td>離退訓-是否為適應期內離退訓" & "</td>" '離退訓-是否為適應期內離訓
        ExportStr &= "<td>離退訓-離訓原因" & "</td>" '離退訓-離訓原因
        ExportStr &= "<td>離退訓-退訓原因" & "</td>" '離退訓-退訓原因

        ExportStr &= "<td>是否就業" & "</td>"
        ExportStr &= "<td>就業狀態" & "</td>"
        ExportStr &= "<td>是否為公法救助就業" & "</td>"
        ExportStr &= "<td>就業單位名稱" & "</td>"
        ExportStr &= "<td>到職日" & "</td>"

        ExportStr &= "<td>是否有就業關連性" & "</td>"
        ExportStr &= "<td>就業關連性原因1" & "</td>"
        ExportStr &= "<td>就業關連性原因2" & "</td>"
        ExportStr &= "<td>就業長度是否超過1個月" & "</td>"
        ExportStr &= "</tr>" & vbCrLf
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

        Dim iSeqno As Integer = 0
        For Each dr As DataRow In dt.Rows
            '序號+1
            iSeqno += 1
            '建立資料面
            ExportStr = ""
            ExportStr &= "<tr>" & vbCrLf
            'ExportStr &= "<td>" & Convert.ToString(iSeqno) & "</td>" 
            ExportStr &= "<td>" & Convert.ToString(dr("YEARS")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("PLANNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("DISTNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ORGNAME2")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ORGNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ORGTYPENAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("CLASSCNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("BUSNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("JOBNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("TRAINNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("TADDRESSZIP")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ZNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("CTNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("THOURS")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("TNUM")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("STDATE")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("FTDATE")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("NAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("SEX2")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("YEARSOLD")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("DEGREENAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("MINAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ZIPNAME2")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("ZIPNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("BUDGETIDN")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("WORKSUPPIDENT")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("SUBSIDYY")) & "</td>"

            ExportStr &= "<td>" & Convert.ToString(dr("STUDSTATUS2")) & "</td>" '是否結訓
            ExportStr &= "<td>" & Convert.ToString(dr("REJECTDAYIN14")) & "</td>" '離退訓-是否為適應期內離訓
            ExportStr &= "<td>" & Convert.ToString(dr("REASON1")) & "</td>" '離退訓-離訓原因
            ExportStr &= "<td>" & Convert.ToString(dr("REASON2")) & "</td>" '離退訓-退訓原因

            ExportStr &= "<td>" & Convert.ToString(dr("ISGETJOB")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("JOBSTATUS")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("PUBLICRESCUE")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("COMPNAME")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("MDATE")) & "</td>"

            ExportStr &= "<td>" & Convert.ToString(dr("JOBRELATE")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("JOBRELATE_Y1")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("JOBRELATE_Y2")) & "</td>"
            ExportStr &= "<td>" & Convert.ToString(dr("DAYS2")) & "</td>"

            ExportStr &= "</tr>" & vbCrLf
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        Next
        Common.RespWrite(Me, "</table>")
        Common.RespWrite(Me, "</body>")

        TIMS.CloseDbConn(objconn)
        Response.End()
    End Sub

    Protected Sub btnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click
        Dim dt As DataTable = GetExport2dt()
        Call ExpReport2(dt)
    End Sub

    'https://jira.turbotech.com.tw/browse/TIMSC-157
    '匯出按鈕名稱為”匯出班級統計表”、”匯出班級明細表”，
    '因業務單位尚需於地方政府聯繫會議討論，待獲得共識後，再行辦理，
    '故先行處理匯出班級明細表。
    '2017/08/23 17:10
    'Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
    '    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
    'End Sub
End Class
