Partial Class TR_05_008_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "TR_05_008_R"

    'TR_05_008_R
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        Button3.Style("display") = "none"

        If Not IsPostBack Then
            CreateItem()
        End If
        'Button4.Visible = False
        'Button1.Attributes("onclick") = "return search();"
    End Sub

    Sub CreateItem()
        DistID.Attributes("onclick") = "ClearData();"
        TPlanID.Attributes("onclick") = "ClearData();"

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"

        ''選擇全部訓練計畫
        TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

        FTDate1.Text = TIMS.Cdate3(Now.Year.ToString() & "/1/1")
        FTDate2.Text = TIMS.Cdate3(Now.Date)

        DistID = TIMS.Get_DistID(DistID) '轄區
        DistID.Items.Insert(0, New ListItem("全部", ""))

        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
        Call TIMS.SetCblValue(TPlanID, sm.UserInfo.TPlanID)

        'FTDate2.Text = Now.Date
        OCID.Style("display") = "none"
        msg.Text = TIMS.cst_NODATAMsg11
    End Sub

    Private Sub Button2_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.ServerClick

        Dim N As Integer = 0
        Dim N1 As Integer = 0
        Dim DistID1 As String = ""
        'DistID1 = ""
        'N = 0   '預設 N =0 表示沒有勾選轄區選項
        For i As Integer = 1 To Me.DistID.Items.Count - 1
            If Me.DistID.Items(i).Selected Then '假如有勾選
                N = N + 1  '計算轄區勾選選項的數目
                If N = 1 Then '如果是勾選一個選項
                    DistID1 = Convert.ToString(Me.DistID.Items(i).Value) '取得選項的值
                End If
                If N = 2 Then '如果轄區勾選選項的數目=2
                    Common.MessageBox(Me, "只能選擇一個轄區")
                    DistID1 = ""
                    Exit For
                End If
            End If
        Next

        If N = 0 Then '如果轄區選項沒有選
            Common.MessageBox(Me, "請選擇轄區")
        End If

        Dim TPlanID1 As String = ""
        'TPlanID1 = ""
        N1 = 0 '預設 N1 =0 表示沒有勾選計畫選項
        For j As Integer = 1 To Me.TPlanID.Items.Count - 1

            If Me.TPlanID.Items(j).Selected Then '假如有勾選
                N1 = N1 + 1 '計算計畫勾選選項的數目
                If N1 = 1 Then '如果是勾選一個選項
                    TPlanID1 = Convert.ToString(Me.TPlanID.Items(j).Value) '取得選項的值
                End If
                If N1 = 2 Then '如果計畫勾選選項的數目=2
                    Common.MessageBox(Me, "只能選擇一個計畫")
                    TPlanID1 = ""
                    Exit For
                End If

            End If
        Next

        If N = 0 Then '如果計畫選項沒有選
            Common.MessageBox(Me, "請選擇計畫")
            Exit Sub
        End If
        If DistID1 = "" Then
            Common.MessageBox(Me, "請選擇轄區")
            Exit Sub
        End If
        If TPlanID1 = "" Then
            Common.MessageBox(Me, "請選擇計畫")
            Exit Sub
        End If

        Dim strScript1 As String
        strScript1 = "<script language=""javascript"">" + vbCrLf
        strScript1 += "wopen('../../Common/MainOrg.aspx?DistID=' + '" & DistID1 & "' + '&TPlanID=' + '" & TPlanID1 & "'  + '&BtnName=Button3','查詢機構',400,400,1);"
        strScript1 += "</script>"
        Page.RegisterStartupScript("", strScript1)

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click

        Dim dt As DataTable
        'Dim dr As DataRow
        'msg.Text = ""
        msg.Text = "查無此機構底下的班級"
        OCID.Items.Clear()
        OCID.Style("display") = "none"

        PlanID.Value = TIMS.ClearSQM(PlanID.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If PlanID.Value = "" Then Exit Sub
        If RIDValue.Value = "" Then Exit Sub

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " FROM dbo.CLASS_CLASSINFO a" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND a.NotOpen='N'" & vbCrLf
        sql &= " and a.IsSuccess='Y'" & vbCrLf
        sql &= " AND a.PlanID='" & PlanID.Value & "'" & vbCrLf
        sql &= " and a.RID='" & RIDValue.Value & "'" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)

        msg.Text = "查無此機構底下的班級"
        OCID.Style("display") = "none"
        If dt.Rows.Count = 0 Then Exit Sub

        msg.Text = ""
        OCID.Style("display") = ""
        OCID.Items.Clear()
        OCID.Items.Add(New ListItem("全選", "%"))
        For Each dr As DataRow In dt.Rows
            Dim v_ClassName As String = Convert.ToString(dr("CLASSCNAME"))
            Dim v_OCID As String = Convert.ToString(dr("OCID"))
            If v_ClassName <> "" AndAlso v_OCID <> "" Then
                OCID.Items.Add(New ListItem(v_ClassName, v_OCID))
            End If
        Next
        OCID.Style("display") = "inline"
    End Sub


    Sub print_Report1()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)

        '要用的轄區參數
        Dim DistID1 As String = TIMS.GetCblValue(DistID)
        '要用的訓練計畫參數
        Dim TPlanID1 As String = TIMS.GetCblValue(TPlanID)
        '勾選班級後會省略結訓日期的條件
        If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
            FTDate1.Text = ""
            FTDate2.Text = ""
        End If
        FTDate1.Text = TIMS.Cdate3(FTDate1.Text)
        FTDate2.Text = TIMS.Cdate3(FTDate2.Text)

        Dim OCIDStr As String = TIMS.GetCblValue(OCID)

        Dim s_MIdentityID As String = ""
        s_MIdentityID = TIMS.Cst_Identity28_2019_11
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then s_MIdentityID = TIMS.Cst_Identity06_2019_11

        Dim MyValue As String = ""
        MyValue = ""
        MyValue += "&start_date=" & start_date.Text
        MyValue += "&end_date=" & end_date.Text
        MyValue += "&SFTDate=" & Me.FTDate1.Text
        MyValue += "&FFTDate=" & Me.FTDate2.Text
        MyValue += "&MIdentityID=" & Replace(s_MIdentityID, "'", "")
        If DistID1 <> "" Then MyValue += "&DistID=" & DistID1
        If TPlanID1 <> "" Then MyValue += "&TPlanID=" & TPlanID1
        If OCIDStr <> "" Then MyValue += "&OCID=" & OCIDStr
        If RIDValue.Value <> "" Then MyValue += "&RID=" & RIDValue.Value
        'MyValue += "&ClassCName=" & Server.UrlEncode(TIMS.ClearSQM(OCIDName))
        MyValue += "&OrgName=" & Server.UrlEncode(TIMS.ClearSQM(center.Text))
        'MyValue += "&DistName=" & Server.UrlEncode(DistName)
        'MyValue += "&PlanName=" & Server.UrlEncode(TPlanName)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    Function Get_DATATABLE1(ByRef objtable As DataTable) As DataTable
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        PlanID.Value = TIMS.ClearSQM(PlanID.Value)

        '要用的轄區參數
        Dim DistID1 As String = TIMS.GetCblValueIn(DistID)
        '要用的訓練計畫參數
        Dim TPlanID1 As String = TIMS.GetCblValueIn(TPlanID)
        '勾選班級後會省略結訓日期的條件
        If (RIDValue.Value <> "") OrElse (PlanID.Value <> "") Then
            FTDate1.Text = ""
            FTDate2.Text = ""
        End If
        FTDate1.Text = TIMS.Cdate3(FTDate1.Text)
        FTDate2.Text = TIMS.Cdate3(FTDate2.Text)

        Dim OCIDStr As String = TIMS.GetCblValue(OCID)

        Dim s_MIdentityID As String = ""
        s_MIdentityID = TIMS.Cst_Identity28_2019_11
        If TIMS.Cst_TPlanID06.IndexOf(sm.UserInfo.TPlanID) > -1 Then s_MIdentityID = TIMS.Cst_Identity06_2019_11

        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)

        'Dim MyValue As String = ""
        'MyValue = ""
        'MyValue += "&start_date=" & start_date.Text
        'MyValue += "&end_date=" & end_date.Text
        'MyValue += "&SFTDate=" & FTDate1.Text
        'MyValue += "&FFTDate=" & FTDate2.Text
        'MyValue += "&MIdentityID=" & Replace(s_MIdentityID, "'", "")
        'If DistID1 <> "" Then MyValue &= "&DistID=" & DistID1
        'If TPlanID1 <> "" Then MyValue &= "&TPlanID=" & TPlanID1
        'If OCIDStr <> "" Then MyValue &= "&OCID=" & OCIDStr
        'If RIDValue.Value <> "" Then MyValue &= "&RID=" & RIDValue.Value

        Dim parms As New Hashtable
        parms.Clear()
        If start_date.Text <> "" Then parms.Add("STDate1", start_date.Text)
        If end_date.Text <> "" Then parms.Add("STDate2", end_date.Text)
        If FTDate1.Text <> "" Then parms.Add("FTDate1", FTDate1.Text)
        If FTDate2.Text <> "" Then parms.Add("FTDate2", FTDate2.Text)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WC1 AS (" & vbCrLf
        sql &= " 	SELECT CC.OCID" & vbCrLf
        sql &= " 	,CC.CLASSCNAME2" & vbCrLf
        sql &= " 	,CC.YEARS,CC.TPLANID" & vbCrLf
        sql &= " 	,CC.PLANID,CC.DISTID" & vbCrLf
        sql &= " 	,CC.RID,CC.STDATE,CC.FTDATE" & vbCrLf
        sql &= " 	,CC.NOTOPEN,CC.ISSUCCESS" & vbCrLf
        sql &= " 	from dbo.VIEW2 CC" & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        'sql &= " 	AND CC.YEARS ='2018' AND CC.TPLANID='06' AND CC.DISTID='001'" & vbCrLf
        If start_date.Text <> "" Then sql &= " AND CC.STDate >=@STDate1" & vbCrLf
        If end_date.Text <> "" Then sql &= " AND CC.STDate <=@STDate2" & vbCrLf
        If FTDate1.Text <> "" Then sql &= " AND CC.FTDate >=@FTDate1" & vbCrLf
        If FTDate2.Text <> "" Then sql &= " AND CC.FTDate <=@FTDate2" & vbCrLf
        If DistID1 <> "" Then sql &= String.Format(" AND CC.DISTID IN ({0})", DistID1) & vbCrLf
        If TPlanID1 <> "" Then sql &= String.Format(" AND CC.TPLANID IN ({0})", TPlanID1) & vbCrLf
        If OCIDStr <> "" Then sql &= String.Format(" AND CC.OCID IN ({0})", OCIDStr) & vbCrLf
        If RIDValue.Value <> "" Then sql &= String.Format(" AND CC.RID='{0}'", RIDValue.Value) & vbCrLf
        sql &= " 	AND CC.NOTOPEN ='N'" & vbCrLf
        sql &= " 	AND CC.ISSUCCESS = 'Y'" & vbCrLf
        sql &= " )	" & vbCrLf

        sql &= " ,WE1B AS (" & vbCrLf
        sql &= " 	SELECT dbo.FN_GET_MIDENTITYID3(B.MIDENTITYID,B.IDENTITYID) IDENTITYID" & vbCrLf
        sql &= " 	FROM WC1 cc" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTYPE b on b.ocid1=cc.ocid" & vbCrLf
        sql &= " 	JOIN STUD_ENTERTEMP a on a.SETID=b.SETID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WE1 AS (" & vbCrLf
        sql &= " 	SELECT B.IDENTITYID" & vbCrLf
        sql &= " 	,COUNT(1) STUDETNUM" & vbCrLf
        sql &= " 	FROM WE1B B" & vbCrLf
        sql &= " 	GROUP BY B.IDENTITYID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WS1 AS (" & vbCrLf
        sql &= " 	select cs.MIDENTITYID IdentityID" & vbCrLf
        sql &= " 	,COUNT(1) OPENSTUD /*開訓*/" & vbCrLf
        sql &= " 	,COUNT(case when cs.StudStatus not in (2,3) and cc.FTDAte< GETDATE() then 1 END) CLOSESTUD /*結訓*/" & vbCrLf
        sql &= " 	From WC1 cc" & vbCrLf
        sql &= " 	join dbo.CLASS_STUDENTSOFCLASS cs on cc.OCID=cs.OCID" & vbCrLf
        sql &= " 	join dbo.KEY_IDENTITY ky on ky.IdentityID = cs.MIdentityID" & vbCrLf
        sql &= " 	where 1=1" & vbCrLf
        sql &= " 	and cs.MakeSOCID is null" & vbCrLf
        sql &= " 	GROUP BY cs.MIDENTITYID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT kd.IDENTITYID" & vbCrLf
        sql &= " ,kd.NAME MIDNAME" & vbCrLf
        sql &= " ,ISNULL(g1.STUDETNUM,0) STUDETNUM" & vbCrLf
        sql &= " ,ISNULL(g2.OPENSTUD,0) OPENSTUD" & vbCrLf
        sql &= " ,ISNULL(g2.CLOSESTUD,0) CLOSESTUD" & vbCrLf
        sql &= " From dbo.KEY_IDENTITY kd" & vbCrLf
        sql &= " left join WE1 g1 on g1.IdentityID=kd.IdentityID" & vbCrLf
        sql &= " left join WS1 g2 on g2.IdentityID=kd.IdentityID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= String.Format(" AND kd.IDENTITYID IN ({0})", s_MIdentityID) & vbCrLf
        sql &= " ORDER BY kd.IDENTITYID" & vbCrLf
        objtable = DbAccess.GetDataTable(sql, objconn, parms)
        Return objtable
    End Function

    Sub Utl_Export1()

        'DataGrid1.AllowPaging = False
        'DataGrid1.EnableViewState = False  '把ViewState給關了
        'Call Search1()

        Dim objtable As DataTable = Nothing
        objtable = Get_DATATABLE1(objtable)
        If objtable Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
            Return 'Exit Sub
        End If
        If objtable.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Return 'Exit Sub
        End If
        Dim s_log1 As String = ""
        s_log1 = String.Format("Total rows={0}, columns={1}", objtable.Rows.Count, objtable.Columns.Count)
        TIMS.LOG.Debug(s_log1)
        msg.Text = ""

        Const sFileName1 As String = "訓練計畫特定對象人數統計表"

        Dim i_STUDETNUM As Integer = 0 '報名人數
        Dim i_OPENSTUD As Integer = 0 '開訓人數
        Dim i_CLOSESTUD As Integer = 0 '結訓人數

        'mso-number-format:"0" 
        Dim strSTYLE As String = ""
        strSTYLE &= ("<style>")
        strSTYLE &= ("td{mso-number-format:""\@"";}")
        strSTYLE &= (".noDecFormat{mso-number-format:""0"";}")
        strSTYLE &= ("</style>")

        Dim sbHTML As New System.Text.StringBuilder()
        'Dim strHTML As String = ""
        sbHTML.Append("<div>")
        sbHTML.Append("<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">")

        '主要參訓身分、報名人數、開訓人數、結訓人數
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= "<td>主要參訓身分</td>"
        ExportStr &= "<td>報名人數</td>"
        ExportStr &= "<td>開訓人數</td>"
        ExportStr &= "<td>結訓人數</td>"
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        '建立資料面
        ExportStr = ""
        For Each dr As DataRow In objtable.Rows
            If (Convert.ToString(dr("STUDETNUM")) <> "") Then i_STUDETNUM += Val(dr("STUDETNUM"))
            If (Convert.ToString(dr("OPENSTUD")) <> "") Then i_OPENSTUD += Val(dr("OPENSTUD"))
            If (Convert.ToString(dr("CLOSESTUD")) <> "") Then i_CLOSESTUD += Val(dr("CLOSESTUD"))
            ExportStr = ""
            ExportStr &= "<tr>"
            ExportStr &= String.Format("<td>{0}</td>", Convert.ToString(dr("MIDNAME"))) '主要參訓身分
            ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", Convert.ToString(dr("STUDETNUM"))) '報名人數
            ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", Convert.ToString(dr("OPENSTUD"))) '開訓人數
            ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", Convert.ToString(dr("CLOSESTUD"))) '結訓人數
            ExportStr &= "</tr>"
            sbHTML.Append(ExportStr)
        Next

        ExportStr = ""
        ExportStr &= "<tr>"
        ExportStr &= String.Format("<td>{0}</td>", "合計") '主要參訓身分
        ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", i_STUDETNUM) '報名人數
        ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", i_OPENSTUD) '開訓人數
        ExportStr &= String.Format("<td class=""noDecFormat"">{0}</td>", i_CLOSESTUD) '結訓人數
        ExportStr &= "</tr>"
        sbHTML.Append(ExportStr)

        'Dim strHTML As String = ""
        'DataGrid1.AllowPaging = False
        'DataGrid1.Columns(Cst_功能欄位).Visible = False
        'DataGrid1.EnableViewState = False  '把ViewState給關了

        'Dim objStringWriter As New System.IO.StringWriter
        'Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        'Div1.RenderControl(objHtmlTextWriter)
        'strHTML &= (TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", sbHTML.ToString())
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        'DataGrid1.AllowPaging = True
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()

        'DataGrid1.Columns(Cst_功能欄位).Visible = True
        'Call TIMS.CloseDbConn(objconn)
    End Sub

    '匯出 
    Protected Sub Export1_Click(sender As Object, e As EventArgs) Handles Export1.Click
        'Const Cst_功能欄位 As Integer = 14
        'Dim sErrmsg As String = ""
        'CheckData1(sErrmsg)
        'If sErrmsg <> "" Then
        '    Common.MessageBox(Me, sErrmsg)
        '    Exit Sub
        'End If

        Call Utl_Export1()
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Call print_Report1()
    End Sub
End Class
