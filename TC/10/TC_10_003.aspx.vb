Partial Class TC_10_003
    Inherits AuthBasePage

    'VIEW_MEETEXAM 
    Const cst_printFN1 As String = "TC_10_003_R"
    'Const cst_ADD1 As String = "ADD1" '新增
    'Const cst_UPD1 As String = "UPD1" '修改
    'Const cst_DEL1 As String = "DEL1" '刪除
    'Const cst_EDIT3 As String = "EDIT3" '管理出席名單 'BTNEDIT3
    'Const cst_VIEW1 As String = "VIEW1" '查看出席名單

    'Dim ff3 As String = ""
    'Const Cst_EXAMINERpkName As String = "EMSEQ"

    ''Dim BTNUPD1 As Button = e.Item.FindControl("BTNUPD1") '修改
    ''Dim BTNDEL1 As Button = e.Item.FindControl("BTNDEL1") '刪除
    ''Dim BTNEDIT3 As Button = e.Item.FindControl("BTNEDIT3") '管理出席名單
    ''Dim BTNVIEW1 As Button = e.Item.FindControl("BTNVIEW1") '查看出席名單-分署

    Dim dtDist As DataTable = Nothing 'TIMS.Get_DistIDdt(objconn)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        trRBListExpType.Visible = False
        BtnExport1.Visible = False

        dtDist = TIMS.Get_DISTIDdt(objconn) 'Dim dtDist As DataTable = TIMS.Get_DistIDdt(objconn)

        If Not IsPostBack Then cCreate1()

    End Sub

    ''' <summary>
    ''' 載入1
    ''' </summary>
    Sub cCreate1()
        msg1.Text = ""

        '選擇全部別
        cblACCEPTSTAGE_sch.Items.Insert(0, New ListItem("全部", ""))
        cblACCEPTSTAGE_sch.Attributes("onclick") = "SelectAll('cblACCEPTSTAGE_sch','cblACCEPTSTAGE_sch_List');"

        Dim iSYears As Integer = 2019
        Dim iEYears As Integer = (Now.Year + 1)

        '查詢
        cblDISTID_SCH = TIMS.Get_DistID(cblDISTID_SCH, dtDist) '轄區分署
        cblDISTID_SCH.Items.Insert(0, New ListItem("全部", ""))
        cblDISTID_SCH.Attributes("onclick") = "SelectAll('cblDISTID_SCH','cblDISTID_SCH_List');"

        ddlMYEARS1_SCH = TIMS.GetSyear(ddlMYEARS1_SCH, iSYears, iEYears, True)  '年度
        ddlMYEARS2_SCH = TIMS.GetSyear(ddlMYEARS2_SCH, iSYears, iEYears, True)  '年度
        'Common.SetListItem(ddlDISTID_SCH, sm.UserInfo.DistID) '轄區分署(查詢)
        Common.SetListItem(ddlMYEARS1_SCH, sm.UserInfo.Years) '年度(查詢)
        Common.SetListItem(ddlMYEARS2_SCH, sm.UserInfo.Years) '年度(查詢)

        'SHOW_PANEL(0)
        Dim flagS1 As Boolean = TIMS.IsSuperUser(Me, 1) '是否為(後台)系統管理者 
        trRBListExpType.Visible = If(flagS1, True, False)
        BtnExport1.Visible = If(flagS1, True, False)

    End Sub

    ''' <summary>
    ''' 列印
    ''' </summary>
    Sub Utl_Print1()
        msg1.Text = ""
        Dim v_DISTID As String = TIMS.GetCblValue(cblDISTID_SCH) '轄區分署
        Dim v_MYEARS1 As String = TIMS.GetListValue(ddlMYEARS1_SCH) '年度區間1
        Dim v_MYEARS2 As String = TIMS.GetListValue(ddlMYEARS2_SCH) '年度區間2
        If TIMS.ChkYearErr3(v_MYEARS1, v_MYEARS2) Then
            Dim T_MYEAR As String = v_MYEARS1
            v_MYEARS1 = v_MYEARS2
            v_MYEARS2 = T_MYEAR
            Common.SetListItem(ddlMYEARS1_SCH, v_MYEARS1)
            Common.SetListItem(ddlMYEARS2_SCH, v_MYEARS2)
        End If
        'rblORGPLANKIND_sch.Text = "" '計畫別 '產業人才投資計畫 // 提升勞工自主學習計畫
        Dim v_ORGPLANKIND As String = TIMS.GetListValue(rblORGPLANKIND_sch)
        'rblCATEGORY_SCH.Text = "" '審查會議類別 '1:轄區/2:跨區
        Dim v_CATEGORY As String = TIMS.GetListValue(rblCATEGORY_SCH)
        'cblACCEPTSTAGE_sch.Text = "" '受理階段
        Dim v_ACCEPTSTAGE As String = TIMS.GetCblValue(cblACCEPTSTAGE_sch)

        Dim s_ERRMSG As String = ""
        If v_MYEARS1 = "" Then s_ERRMSG &= "年度區間 請選擇 起始年度" & vbCrLf
        If v_MYEARS2 = "" Then s_ERRMSG &= "年度區間 請選擇 迄止年度" & vbCrLf
        If s_ERRMSG = "" AndAlso v_MYEARS2 < v_MYEARS1 Then s_ERRMSG &= "年度區間 請選擇 起始~迄止年度排序有誤" & vbCrLf
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, s_ERRMSG)
            Return
        End If

        Dim v_ROCY1 As String = CStr(Val(v_MYEARS1) - 1911) '年度區間1
        Dim v_ROCY2 As String = CStr(Val(v_MYEARS2) - 1911) '年度區間2

        'http://192.168.0.76:8080/ReportServer3/report?RptID=TC_10_003_R&DISTID=000,001,003,004,005,006&MYEARS1=2020&MYEARS2=2022&ORGPLANKIND=G&CATEGORY=1&ACCEPTSTAGE=A1,A2,B1,B2,C1,C2&UserID=snoopy
        'http://192.168.0.76:8080/ReportServer3/report?RptID=TC_10_003_R&DISTID=000,001,003,004,005,006&MYEARS1=2020&MYEARS2=2022&ORGPLANKIND=&CATEGORY=&ACCEPTSTAGE=A1,A2,B1,B2,C1,C2&UserID=snoopy
        'http://192.168.0.76:8080/ReportServer3/report?RptID=TC_10_003_R&DISTID=&MYEARS1=2021&MYEARS2=2021&ORGPLANKIND=&CATEGORY=&ACCEPTSTAGE=&UserID=snoopy
        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "DISTID", v_DISTID) 'where
        TIMS.SetMyValue(MyValue, "MYEARS1", v_MYEARS1)
        TIMS.SetMyValue(MyValue, "MYEARS2", v_MYEARS2)
        TIMS.SetMyValue(MyValue, "ROCY1", v_ROCY1)
        TIMS.SetMyValue(MyValue, "ROCY2", v_ROCY2)
        TIMS.SetMyValue(MyValue, "ORGPLANKIND", v_ORGPLANKIND)
        TIMS.SetMyValue(MyValue, "CATEGORY", v_CATEGORY)
        TIMS.SetMyValue(MyValue, "ACCEPTSTAGE", v_ACCEPTSTAGE) 'where
        'MyValue = "RID=" & RIDValue.Value & "&Years=" & syear.SelectedValue & "&TPlanID=" & sm.UserInfo.TPlanID
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    ''' <summary>
    ''' 匯出
    ''' </summary>
    Sub Utl_EXPORT1()
        msg1.Text = ""
        '轄區分署 (where)
        'Dim v_DISTID As String = TIMS.GetCblValue(cblDISTID_SCH) 
        Dim in_DISTID As String = TIMS.CombiSQM2IN(TIMS.GetCblValue(cblDISTID_SCH))

        Dim v_MYEARS1 As String = TIMS.GetListValue(ddlMYEARS1_SCH) '年度區間1
        Dim v_MYEARS2 As String = TIMS.GetListValue(ddlMYEARS2_SCH) '年度區間2
        'rblORGPLANKIND_sch.Text = "" '計畫別 '產業人才投資計畫 // 提升勞工自主學習計畫
        Dim v_ORGPLANKIND As String = TIMS.GetListValue(rblORGPLANKIND_sch)
        'rblCATEGORY_SCH.Text = "" '審查會議類別 '1:轄區/2:跨區
        Dim v_CATEGORY As String = TIMS.GetListValue(rblCATEGORY_SCH)

        'cblACCEPTSTAGE_sch.Text = "" '受理階段 (where)
        'Dim v_ACCEPTSTAGE As String = TIMS.GetCblValue(cblACCEPTSTAGE_sch)
        Dim in_ACCEPTSTAGE As String = TIMS.CombiSQM2IN(TIMS.GetCblValue(cblACCEPTSTAGE_sch))

        Dim s_ERRMSG As String = ""
        If v_MYEARS1 = "" Then s_ERRMSG &= "年度區間 請選擇 起始年度" & vbCrLf
        If v_MYEARS2 = "" Then s_ERRMSG &= "年度區間 請選擇 迄止年度" & vbCrLf
        If s_ERRMSG = "" AndAlso v_MYEARS2 < v_MYEARS1 Then s_ERRMSG &= "年度區間 請選擇 起始~迄止年度排序有誤" & vbCrLf

        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, s_ERRMSG)
            Return
        End If

        Dim parms As Hashtable = New Hashtable
        Dim sql As String = ""
        sql &= " WITH WC1 AS ( SELECT a.MBRNAME" & vbCrLf
        sql &= "  ,a.EMSEQ" & vbCrLf
        sql &= "  ,a.DISTID" & vbCrLf
        sql &= "  ,a.MYEARS" & vbCrLf
        sql &= "  ,a.ORGPLANKIND" & vbCrLf
        sql &= "  ,a.CATEGORY" & vbCrLf
        sql &= "  ,a.ACCEPTSTAGE" & vbCrLf
        sql &= "  ,a.ATTEND" & vbCrLf
        sql &= "  ,a.NOTINABS" & vbCrLf
        sql &= "  FROM dbo.VIEW_MEETEXAM a" & vbCrLf
        sql &= "  WHERE 1=1" & vbCrLf
        If (in_DISTID <> "") Then sql &= String.Format(" AND a.DISTID IN ({0})", in_DISTID) & vbCrLf
        If (v_MYEARS1 <> "") Then sql &= String.Format(" AND a.MYEARS >='{0}'", v_MYEARS1) & vbCrLf
        If (v_MYEARS2 <> "") Then sql &= String.Format(" AND a.MYEARS <='{0}'", v_MYEARS2) & vbCrLf
        If (v_ORGPLANKIND <> "") Then sql &= String.Format(" AND a.ORGPLANKIND ='{0}'", v_ORGPLANKIND) & vbCrLf
        If (v_CATEGORY <> "") Then sql &= String.Format(" AND a.CATEGORY ='{0}'", v_CATEGORY) & vbCrLf
        If (in_ACCEPTSTAGE <> "") Then sql &= String.Format(" AND a.ACCEPTSTAGE IN ({0})", in_ACCEPTSTAGE) & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WC1B AS ( SELECT a.MBRNAME,a.EMSEQ" & vbCrLf
        sql &= " 	,COUNT(1) CNT1" & vbCrLf '應出席次數
        sql &= " 	,COUNT(CASE WHEN a.ATTEND='Y' THEN 1 END) CNT2" & vbCrLf '實際出席次數
        sql &= " 	,COUNT(CASE WHEN a.NOTINABS='Y' THEN 1 END) CNT3" & vbCrLf '不列入缺席
        sql &= " 	,COUNT(CASE WHEN a.ATTEND='Y' OR a.NOTINABS='Y' THEN 1 END) CNT4" & vbCrLf '實際出席次數+不列入缺席(算有出席次數)
        sql &= " 	FROM WC1 a" & vbCrLf
        sql &= " 	GROUP BY a.MBRNAME,a.EMSEQ )" & vbCrLf

        sql &= " ,WC2 AS ( SELECT ROW_NUMBER() OVER(ORDER BY ROUND(CAST(c1.CNT4 AS FLOAT)/c1.CNT1*100,2) DESC,c1.MBRNAME ASC) AS ROWX" & vbCrLf
        sql &= " 	,c1.MBRNAME" & vbCrLf
        sql &= " 	,c1.EMSEQ" & vbCrLf
        sql &= " 	,c1.CNT1,c1.CNT2,c1.CNT3,c1.CNT4" & vbCrLf
        sql &= "    ,CASE WHEN c1.CNT3>0 then concat('(假',c1.CNT3,')') end TXT3" & vbCrLf
        sql &= " 	FROM WC1B c1 )" & vbCrLf
        sql &= " ,WC2A AS (SELECT ROW_NUMBER() OVER(ORDER BY ROWX ASC) AS ROWY,ROWX,MBRNAME,EMSEQ,CNT1,CNT2,CNT3,CNT4,TXT3 FROM WC2 WHERE (ROWX % 2=1))" & vbCrLf
        sql &= " ,WC2B AS (SELECT ROW_NUMBER() OVER(ORDER BY ROWX ASC) AS ROWY,ROWX,MBRNAME,EMSEQ,CNT1,CNT2,CNT3,CNT4,TXT3 FROM WC2 WHERE (ROWX % 2=0))" & vbCrLf

        sql &= " SELECT A.ROWY" & vbCrLf
        sql &= " ,A.MBRNAME AMBRNAME,A.CNT1 ACNT1,A.CNT2 ACNT2,A.CNT3 ACNT3" & vbCrLf
        sql &= " ,CASE WHEN A.CNT4 IS NOT NULL THEN CONCAT(ROUND(CAST(A.CNT4 AS FLOAT)/A.CNT1*100,2),'%') END AATTERATE" & vbCrLf
        sql &= " ,CONCAT(A.CNT2,A.TXT3) ATXT3" & vbCrLf
        sql &= " ,B.MBRNAME BMBRNAME,B.CNT1 BCNT1,B.CNT2 BCNT2,B.CNT3 BCNT3" & vbCrLf
        sql &= " ,CASE WHEN B.CNT4 IS NOT NULL THEN CONCAT(ROUND(CAST(B.CNT4 AS FLOAT)/B.CNT1*100,2),'%') END BATTERATE" & vbCrLf
        sql &= " ,CONCAT(B.CNT2,B.TXT3) BTXT3" & vbCrLf
        'sql &= " ,CONCAT(C.ROCY1,'年~',C.ROCY2,'年　產業人才投資方案課程審查委員出席統計總表')  TOPNAME2" & vbCrLf
        sql &= " FROM WC2A A" & vbCrLf
        'sql &= " CROSS JOIN WC3 C" & vbCrLf
        sql &= " LEFT JOIN WC2B B ON A.ROWY=B.ROWY" & vbCrLf

        Dim objtable As DataTable
        objtable = DbAccess.GetDataTable(sql, objconn, parms)
        If objtable Is Nothing Then
            msg1.Text = "查無資料!!"
            Return
        End If
        If objtable.Rows.Count = 0 Then
            msg1.Text = "查無資料!!"
            Return
        End If
        msg1.Text = ""

        '委員,應出席次數,實際出席次數,出席率,委員,應出席次數,實際出席次數,出席率

        Const s_title1 As String = "委員,應出席次數,實際出席次數,出席率,委員,應出席次數,實際出席次數,出席率"
        Const s_data1 As String = "aMBRNAME,aCNT1,aTXT3,aAtteRate,bMBRNAME,bCNT1,bTXT3,bAtteRate"
        Dim As_title1() As String = s_title1.Split(",")
        Dim As_data1() As String = s_data1.Split(",")
        Dim s_colspan As String = As_title1.Length.ToString()

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("OrgPlanInfo", System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim sFileName1 As String = "審查委員出席統計總表" & TIMS.GetDateNo2()
        Dim Roc_MYEARS1 As String = Val(v_MYEARS1) - 1911
        Dim Roc_MYEARS2 As String = Val(v_MYEARS2) - 1911
        '107年~109年產業人才投資方案課程審查委員出席統計總表
        Dim s_titleA1 As String = String.Format("{0}年~{1}年產業人才投資方案課程審查委員出席統計總表", Roc_MYEARS1, Roc_MYEARS2)

        '套CSS值
        'mso-number-format:"0" 
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
            ExportStr &= String.Format("<td>{0}</td>", s_T1) '& vbTab
        Next
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        '建立資料面
        Dim i_num As Integer = 0
        For Each oDr1 As DataRow In objtable.Rows
            i_num += 1
            ExportStr = "<tr>"
            For Each s_D1 As String In As_data1
                ExportStr &= String.Format("<td>{0}</td>", TIMS.ClearSQM(oDr1(s_D1)))
            Next
            ExportStr &= "</tr>"
            sbHTML.Append(ExportStr)
        Next
        sbHTML.Append("</table>")
        sbHTML.Append("</div>")
        objtable = Nothing

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)

        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    '匯出
    Protected Sub BtnExport1_Click(sender As Object, e As EventArgs) Handles BtnExport1.Click
        Utl_EXPORT1()
    End Sub

    '列印
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        Utl_Print1()
    End Sub

    '列印
    'Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
    '    Call Utl_Print1()
    'End Sub
End Class

