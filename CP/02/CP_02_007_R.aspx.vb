Partial Class CP_02_007_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "CP_02_007_R" 'OLD
    Const cst_printFN2 As String = "CP_02_007_R2" 'NEW

    'CP_02_007_R2
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
            Call Create1()
        End If
    End Sub

    Sub Create1()
        DistID = TIMS.Get_DistID(DistID)
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, , , objconn)

        DistID.Enabled = False
        TPlanID.Enabled = False
        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
            TPlanID.Enabled = True
        End If

        Common.SetListItem(DistID, sm.UserInfo.DistID)
        Common.SetListItem(TPlanID, sm.UserInfo.TPlanID)

        '顯示年度下拉選單1
        syear = TIMS.GetSyear(syear)
        '顯示月份下拉選單1
        smonth = TIMS.GetSmonth(smonth)

        Common.SetListItem(syear, Now.Year)
        Common.SetListItem(smonth, Now.Month)

        Button1.Attributes("onclick") = "javascript:return print();"
    End Sub

    Sub CheckData1(ByRef sErrmsg As String)
        sErrmsg = ""

        Dim v_syear As String = TIMS.GetListValue(syear) '統計月份-年度
        Dim v_smonth As String = TIMS.GetListValue(smonth) '統計月份-月份
        Dim v_DistID As String = TIMS.GetListValue(DistID) '轄區分署
        Dim v_TPlanID As String = TIMS.GetListValue(TPlanID) '訓練計畫

        If v_syear = "" Then sErrmsg &= "請選擇 統計月份-年度，必選不可為空" & vbCrLf
        If v_smonth = "" Then sErrmsg &= "請選擇 統計月份-月份，必選不可為空" & vbCrLf
        If v_DistID = "" Then sErrmsg &= "請選擇 轄區分署，必選不可為空" & vbCrLf
        If v_TPlanID = "" Then sErrmsg &= "請選擇 訓練計畫，必選不可為空" & vbCrLf

        Return
    End Sub


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Const Cst_功能欄位 As Integer = 14
        Dim sErrmsg As String = ""
        CheckData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If


        Dim v_syear As String = TIMS.GetListValue(syear) '統計月份-年度
        Dim v_smonth As String = TIMS.GetListValue(smonth) '統計月份-月份
        Dim v_DistID As String = TIMS.GetListValue(DistID) '轄區分署
        Dim v_TPlanID As String = TIMS.GetListValue(TPlanID) '訓練計畫

        Dim vsMyValue As String = ""
        vsMyValue = ""
        vsMyValue &= String.Format("&DistID={0}", v_DistID)  '轄區
        vsMyValue &= String.Format("&TPlanID={0}", v_TPlanID)  '計畫
        vsMyValue &= String.Format("&year={0}", v_syear)  '年度
        vsMyValue &= String.Format("&smonth={0}", v_smonth)  '使用個位數傳送 月份

        'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "list", "student_month", Me.vsMyValue)
        Dim flag_R_2012 As Boolean = False
        If sm.UserInfo.Years >= "2012" Then flag_R_2012 = True '新報表

        Dim sfileName As String = cst_printFN1
        'student_month更名為CP_02_007_R
        If flag_R_2012 Then sfileName = cst_printFN2

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfileName, vsMyValue)

    End Sub

    Function Utl_GetDataTable() As DataTable

        Dim v_syear As String = TIMS.GetListValue(syear) '統計月份-年度
        Dim v_smonth As String = TIMS.GetListValue(smonth) '統計月份-月份
        Dim v_DistID As String = TIMS.GetListValue(DistID) '轄區分署
        Dim v_TPlanID As String = TIMS.GetListValue(TPlanID) '訓練計畫

        Dim v_TM1 As String = TIMS.AddZero(v_smonth, 2) 'If(Val(v_smonth) < 10, "0" & v_smonth, v_smonth)
        Dim v_TYM As String = String.Format("{0}{1}", v_syear, v_TM1)
        Dim parms As New Hashtable
        parms.Clear()
        parms.Add("TYM", v_TYM)
        parms.Add("SYEAR", v_syear)
        parms.Add("SMONTH", v_TM1)
        parms.Add("DistID", v_DistID)
        parms.Add("TPlanID", v_TPlanID)

        Dim sql As String = ""
        sql = "" & vbCrLf
        'sql &= " WITH WP1 AS (" & vbCrLf
        'sql &= " SELECT format(GETDATE(),'yyyyMM') TYM , format(GETDATE(),'yyyy') TY1 ,format(GETDATE(),'MM') TM1" & vbCrLf
        'sql &= " ,@TPlanID TPLANID,'001' DISTID )" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " select cc.ORGNAME" & vbCrLf
        sql &= " ,cc.CLASSCNAME2 CLASSCNAME" & vbCrLf
        sql &= " ,cc.THOURS" & vbCrLf
        sql &= " ,cc.STDATE" & vbCrLf
        sql &= " ,cc.FTDATE" & vbCrLf
        sql &= " ,cc.TPROPERTYID" & vbCrLf
        sql &= " ,cc.TPLANID" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.TMID" & vbCrLf
        sql &= " ,cc.TRAINID" & vbCrLf
        sql &= " ,cc.TRAINNAME" & vbCrLf
        sql &= " ,cc.PLANID" & vbCrLf
        sql &= " ,cc.RID" & vbCrLf
        sql &= " ,cc.DISTNAME" & vbCrLf
        sql &= " ,cc.TNUM" & vbCrLf
        sql &= " ,cc.YEARS" & vbCrLf
        sql &= " ,cc.PLANNAME" & vbCrLf
        sql &= " ,cc.ORGID" & vbCrLf
        sql &= " ,FORMAT(cc.STDate,'yyyyMM') STDateYM" & vbCrLf
        sql &= " ,FORMAT(cc.FTDate,'yyyyMM') FTDateYM" & vbCrLf
        sql &= " FROM dbo.VIEW2 cc" & vbCrLf
        'sql &= " CROSS JOIN WP1 p" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND FORMAT(cc.STDate,'yyyyMM') <= @TYM" & vbCrLf
        sql &= " AND FORMAT(cc.FTDate,'yyyyMM') >= @TYM" & vbCrLf
        sql &= " AND cc.TPLANID=@TPlanID" & vbCrLf
        sql &= " AND cc.DISTID=@DistID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select cc.ORGNAME" & vbCrLf
        sql &= " ,cc.CLASSCNAME" & vbCrLf
        sql &= " ,cc.THOURS" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(cc.FTDATE,'yyyy/MM/dd') FTDATE" & vbCrLf
        'sql &= " ,cc.FTDATE" & vbCrLf
        sql &= " ,cc.TPROPERTYID" & vbCrLf
        sql &= " ,cc.TPLANID" & vbCrLf
        sql &= " ,cc.OCID" & vbCrLf
        sql &= " ,cc.TMID" & vbCrLf
        sql &= " ,c2.COMPANYNAME" & vbCrLf
        sql &= " ,'['+cc.TrainID+']'+cc.TrainName TRAINNAME" & vbCrLf
        sql &= " ,cc.PLANID" & vbCrLf
        sql &= " ,cc.RID" & vbCrLf
        sql &= " ,cc.DISTNAME" & vbCrLf
        sql &= " ,cc.TNUM" & vbCrLf
        'sql &= " ,CONCAT(cc.Years,'年度',cc.PlanName ) PLANNAME" & vbCrLf
        'sql &= " ,CONCAT(p.TY1,'年',p.TM1,'月','　受訓學員異動月報表') TITLE2" & vbCrLf
        sql &= " ,cc.ORGID" & vbCrLf
        sql &= " ,ISNULL(s.S_TOTAL ,0) S_TOTAL" & vbCrLf
        sql &= " ,ISNULL(t.total,0) AS TOTAL" & vbCrLf
        sql &= " ,ISNULL(t.add_total,0) AS ADD_TOTAL" & vbCrLf
        sql &= " ,ISNULL(t.rej_total,0) AS REJ_TOTAL" & vbCrLf
        sql &= " ,ISNULL(t.m_total,0) AS M_TOTAL" & vbCrLf
        sql &= " ,ISNULL(t.total2,0) AS TOTAL2" & vbCrLf
        sql &= " FROM WC1 cc" & vbCrLf
        'sql &= " CROSS JOIN WP1 p" & vbCrLf
        sql &= " join dbo.CLASS_CLASSINFO c2 on c2.OCID =cc.OCID" & vbCrLf
        sql &= " LEFT JOIN (" & vbCrLf
        sql &= " 	select cc.OCID" & vbCrLf
        sql &= " 	,count(1) S_TOTAL /*報名人數*/" & vbCrLf
        sql &= " 	from WC1 cc" & vbCrLf
        sql &= " 	JOIN dbo.STUD_ENTERTYPE st ON st.OCID1=cc.OCID" & vbCrLf
        sql &= " 	group by cc.ocid" & vbCrLf
        sql &= " ) s on s.OCID =cc.ocid" & vbCrLf
        sql &= " left join (" & vbCrLf
        sql &= " 	select cc.ocid" & vbCrLf
        sql &= " 	,count(case when CS.ENTERDATE-(CC.STDATE+7) <=0 then 1 end) TOTAL /*開訓人數*/" & vbCrLf
        sql &= " 	,count(case when cs.RejectSOCID IS NOT NULL OR  (CC.STDATE+7)-CS.ENTERDATE <=0 then 1 end) ADD_TOTAL /*累計增補人數*/" & vbCrLf
        sql &= " 	,count(case when cs.StudStatus in (2,3) then 1 end) REJ_TOTAL /*累計離退人數*/" & vbCrLf
        sql &= " 	,count(case when cs.StudStatus not in (2,3) and FORMAT(cc.FTDate,'yyyyMM') > @TYM then 1 end) M_TOTAL /*本月底在訓人數*/" & vbCrLf
        sql &= " 	,count(case when cs.StudStatus not in (2,3) and FORMAT(cc.FTDate,'yyyyMM') = @TYM then 1 end) TOTAL2  /*當月結訓人數*/" & vbCrLf
        sql &= " 	FROM WC1 cc" & vbCrLf
        sql &= " 	JOIN dbo.CLASS_STUDENTSOFCLASS cs on cs.OCID =cc.OCID" & vbCrLf
        sql &= " 	join dbo.STUD_STUDENTINFO ss on ss.sid =cs.sid" & vbCrLf
        'sql &= " 	CROSS JOIN WP1 p" & vbCrLf
        sql &= " 	group by cc.ocid" & vbCrLf
        sql &= " ) t on t.ocid =cc.ocid" & vbCrLf
        sql &= " ORDER BY cc.orgid,cc.OrgName,cc.FTDate,cc.STDate" & vbCrLf

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        Return dt
    End Function

    Sub Utl_Export1()
        Dim oTable As DataTable = Utl_GetDataTable()
        If oTable Is Nothing Then
            msg1.Text = "查無資料!!"
            Return
        End If
        If oTable.Rows.Count = 0 Then
            msg1.Text = "查無資料!!"
            Return
        End If
        msg1.Text = ""

        Dim v_syear As String = TIMS.GetListValue(syear) '統計月份-年度
        Dim v_smonth As String = TIMS.GetListValue(smonth) '統計月份-月份
        Dim v_DistID As String = TIMS.GetListValue(DistID) '轄區分署
        Dim v_TPlanID As String = TIMS.GetListValue(TPlanID) '訓練計畫
        Dim t_TPlanID As String = TIMS.GetListText(TPlanID) '訓練計畫 txt

        'Dim vsMyValue As String = ""
        'vsMyValue = ""
        'vsMyValue &= String.Format("&DistID={0}", v_DistID)  '轄區
        'vsMyValue &= String.Format("&TPlanID={0}", v_TPlanID)  '計畫
        'vsMyValue &= String.Format("&year={0}", v_syear)  '年度
        'vsMyValue &= String.Format("&smonth={0}", v_smonth)  '使用個位數傳送 月份

        Dim s_Title1 As String = "訓練機構,班別名稱,訓練職類,訓練時數,開訓日期,結訓日期,招生人數,報名人數,開訓人數,累計增補人數,累計離退人數,本月底在訓人數,結訓人數"
        Dim s_dataCol1 As String = "ORGNAME,CLASSCNAME,TRAINNAME,THOURS,STDATE,FTDATE,TNUM,S_TOTAL,TOTAL,ADD_TOTAL,REJ_TOTAL,M_TOTAL,TOTAL2"
        Dim s_noDec1 As String = "TNUM,S_TOTAL,TOTAL,ADD_TOTAL,REJ_TOTAL,M_TOTAL,TOTAL2"
        Dim As_Title1 As String() = s_Title1.Split(",")
        Dim As_dataCol1 As String() = s_dataCol1.Split(",")
        Dim As_noDec1 As String() = s_noDec1.Split(",")
        Dim i_colspan As Integer = As_Title1.Length

        Dim sFileName1 As String = "學員異動月報表" & TIMS.GetDateNo2()
        Dim s_titleA1 As String = String.Format("{0} {1}年{2}月 受訓學員異動月報表", t_TPlanID, v_syear, v_smonth)
        '套CSS值 'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Const cst_ColFormat1 As String = "<td>{0}</td>"
        Const cst_ColFormat2 As String = "<td class=""noDecFormat"">{0}</td>" '(純數字)

        Dim ExportStr As String '建立輸出文字
        Dim sbHTML As System.Text.StringBuilder = New System.Text.StringBuilder

        'Dim strHTML As String = ""
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")
        'titleA1
        ExportStr = String.Format("<tr><td colspan=""{0}"">{1}</td></tr>", i_colspan, s_titleA1) '& vbTab
        sbHTML.Append(ExportStr)
        'title1
        ExportStr = "<tr>"
        For Each s_T1 As String In As_Title1
            ExportStr &= String.Format("<td>{0}</td>", s_T1)
        Next
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        '建立資料面
        Dim i_num As Integer = 0
        For Each oDr1 As DataRow In oTable.Rows
            i_num += 1
            ExportStr = "<tr>"
            For Each s_D1 As String In As_dataCol1

                Dim flag_noDec1 As Boolean = TIMS.FindValue1(As_noDec1, s_D1) '(純數字)
                Dim s_ColoumFMT2 As String = cst_ColFormat1
                If flag_noDec1 Then s_ColoumFMT2 = cst_ColFormat2 '(純數字)
                ExportStr &= String.Format(s_ColoumFMT2, TIMS.ClearSQM(oDr1(s_D1))) & vbTab

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

        TIMS.Utl_RespWriteEnd(Me, objconn, "")
    End Sub

    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        'Const Cst_功能欄位 As Integer = 14
        Dim sErrmsg As String = ""
        CheckData1(sErrmsg)
        If sErrmsg <> "" Then
            Common.MessageBox(Me, sErrmsg)
            Exit Sub
        End If

        Call Utl_Export1()
    End Sub
End Class
