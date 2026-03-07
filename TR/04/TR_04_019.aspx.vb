Partial Class TR_04_019
    Inherits AuthBasePage

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
        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not IsPostBack Then
            msg.Text = ""
            PageControler1.Visible = False
            DataGrid1.Visible = False

            '年度
            Syear = TIMS.GetSyear(Syear)
            Syear.Enabled = True
            Common.SetListItem(Syear, sm.UserInfo.Years)
            If sm.UserInfo.DistID <> "000" Then
                Syear.Enabled = False
            End If

            'DistID.SelectedValue = sm.UserInfo.DistID
            DistID = TIMS.Get_DistID(DistID)
            DistID.Enabled = True
            If sm.UserInfo.DistID <> "000" Then
                DistID.Enabled = False
                'DistID.SelectedValue = sm.UserInfo.DistID
                Common.SetListItem(DistID, sm.UserInfo.DistID)
            End If

            TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y")
            ''預算來源
            'BudgetList = TIMS.Get_Budget(BudgetList, 3)

            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"

            btnSearch.Attributes("onclick") = "return chkSearch();"
            btnExport1.Attributes("onclick") = "return chkSearch();"
        End If

    End Sub

    '查詢
    Sub Search1()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim j As Integer = 0
        Dim itemPlan As String = ""
        Dim itemIsGetJob As String = ""
        j = 0
        itemPlan = ""
        For i As Integer = 1 To Me.TPlanID.Items.Count - 1
            Dim objitem As ListItem = Me.TPlanID.Items(i)
            If objitem.Selected = True Then
                j += 1
                If itemPlan <> "" Then itemPlan += ","
                itemPlan += "'" & objitem.Value & "'"
            End If
        Next
        If j = Me.TPlanID.Items.Count - 1 Then itemPlan = ""

        j = 0
        For i As Integer = 0 To IsGetJob.Items.Count - 1
            Dim objitem As ListItem = Me.IsGetJob.Items(i)
            If objitem.Selected = True Then
                j += 1
                If itemIsGetJob <> "" Then itemIsGetJob += ","
                itemIsGetJob += "'" & objitem.Value & "'"
            End If
        Next
        If j = Me.IsGetJob.Items.Count Then itemIsGetJob = ""


        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " select " & vbCrLf
        sql += " vp.planname 訓練計畫" & vbCrLf
        sql += " ,vp.distname 轄區" & vbCrLf
        sql += " ,oo.OrgName 訓練機構" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) ""班級名稱""" & vbCrLf
        sql += " ,CONVERT(varchar, cc.stdate, 111)+'~'+CONVERT(varchar, cc.ftdate, 111)  訓練起迄" & vbCrLf
        sql += " ,ss.name 姓名" & vbCrLf
        sql += " ,ss.idno 身分證字號" & vbCrLf
        'Sql += " ---,datediff(year,ss.birthday,cc.stdate) 年齡" & vbCrLf
        sql += " ,CONVERT(varchar, ss.birthday, 111)  出生日期" & vbCrLf
        sql += " ,ss2.PhoneD 聯絡電話1 ,ss2.PhoneN 聯絡電話2" & vbCrLf
        sql += " ,ss2.cellPhone 行動電話" & vbCrLf
        'Sql += " --,ss2.PhoneD 聯絡電話日 ,ss2.PhoneN 聯絡電話夜 ,ss2.cellPhone 行動電話" & vbCrLf
        sql += " ,vz.ZipName 居住地" & vbCrLf
        sql += " ,ss2.Address 通訊地址" & vbCrLf
        'Sql += " , case when " & vbCrLf
        'Sql += "  dbo.NVL(cs.EnterChannel ,0) =4 " & vbCrLf
        'Sql += "  or ggg.cnt>0 then '有' else '' end 開立推介單" & vbCrLf
        'Sql += " -- ,sg3.BusName 就業公司 ,sg3.BusTel 公司聯絡電話" & vbCrLf
        'Sql += " ,case when sg3.IsGetJob='0' then '0.未就業'when sg3.IsGetJob='1' then '1.就業'when sg3.IsGetJob='2' then '2.不就業' else '99.未填寫' end 職訓就業狀況" & vbCrLf
        sql += " ,dbo.DECODE8(sg3.IsGetJob,0,'0.未就業',1,'1.就業',2,'2.不就業','99.未填寫') 職訓就業狀況" & vbCrLf
        'Sql += " -----select count(*) cnt" & vbCrLf
        sql += " from view_plan vp " & vbCrLf
        sql += " join key_plan kp  on kp.TPlanID =vp.TPlanID " & vbCrLf
        sql += " join class_classinfo cc  on vp.planid=cc.planid" & vbCrLf
        sql += " join plan_planinfo pp  on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        sql += " join org_orginfo oo on oo.comidno =cc.comidno " & vbCrLf
        sql += " join Class_StudentsOfClass cs  on cs.ocid =cc.ocid" & vbCrLf
        sql += " join stud_studentinfo ss  on ss.SID =cs.SID" & vbCrLf
        sql += " join Stud_SubData ss2  on ss2.sid=ss.sid" & vbCrLf
        'Sql += " join Stud_GetJobState3 sg3 on sg3.CPoint=1 and sg3.socid =cs.socid and sg3.IsGetJob='1'--採已就業名單" & vbCrLf
        sql += " left join view_zipname vz  on vz.ZipCode=ss2.ZipCode1" & vbCrLf
        sql += " left join Stud_GetJobState3 sg3  on sg3.CPoint=1 and sg3.socid =cs.socid" & vbCrLf
        'Sql += " left join (" & vbCrLf
        'Sql += "  select gg.idno, gg.trn_class, count(*) cnt " & vbCrLf
        'Sql += "  from Adp_GOVTRNData gg " & vbCrLf
        'Sql += "  group by gg.idno, gg.trn_class" & vbCrLf
        'Sql += " ) ggg on ggg.idno =ss.idno and ggg.trn_class=cc.ocid" & vbCrLf
        sql += " where 1=1" & vbCrLf
        sql += " and cc.IsSuccess='Y'" & vbCrLf
        sql += " and cc.NotOpen='N'" & vbCrLf
        sql += " AND cs.StudStatus NOT IN (2,3)" & vbCrLf

        '依登入年度
        'Sql += " and vp.years ='" & sm.UserInfo.Years & "'" & vbCrLf
        '依選擇年度
        If Me.Syear.SelectedValue <> "" Then
            sql += " and vp.years ='" & Me.Syear.SelectedValue & "'" & vbCrLf
        End If
        If Me.DistID.SelectedValue <> "" Then
            sql += " and vp.distid ='" & Me.DistID.SelectedValue & "'" & vbCrLf
        End If
        '結訓日期
        If Me.FTDate1.Text <> "" Then
            sql += " and cc.ftdate>= " & TIMS.To_date(Me.FTDate1.Text) & vbCrLf
        End If
        If Me.FTDate2.Text <> "" Then
            sql += " and cc.ftdate<= " & TIMS.To_date(Me.FTDate2.Text) & vbCrLf '" & Me.FTDate2.Text & "'" & vbCrLf
        End If
        '選擇計畫
        If itemPlan <> "" Then
            sql += " and vp.TPlanID IN (" & itemPlan & ")" & vbCrLf
        End If
        '就業狀況
        If itemIsGetJob <> "" Then
            sql += " and dbo.NVL(sg3.IsGetJob,'0') IN (" & itemIsGetJob & ")" & vbCrLf
        End If

        Dim dt As DataTable
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            'Common.RespWrite(Me, Sql)
            'Common.RespWrite(Me, ex.ToString)
            Common.MessageBox(Me, ex.ToString)
            Exit Sub
        End Try

        'Table4.Style("display") = "inline"
        'Print.Visible = False
        'btnExport1.Visible = False
        msg.Text = "查無資料"
        PageControler1.Visible = False
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True

            'Print.Visible = True
            'btnExport1.Visible = True
            'PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.Sort = "訓練計畫,轄區,訓練機構,班級名稱,訓練起迄"
            'PageControler1.SqlString = Sql
            PageControler1.ControlerLoad()
        End If
        'Else
        'Common.MessageBox(Me, "查無資料")
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Me.Syear.SelectedValue = "" Then
            Errmsg += "請選擇年度" & vbCrLf
        End If

        If Me.DistID.SelectedValue = "" Then
            'Errmsg += "請選擇轄區中心" & vbCrLf
            Errmsg += "請選擇轄區分署" & vbCrLf
        End If

        If Trim(FTDate1.Text) <> "" Then
            FTDate1.Text = Trim(FTDate1.Text)
            If Not TIMS.IsDate1(FTDate1.Text) Then
                Errmsg += "結訓期間 的起始日不是正確的日期格式" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate1.Text = CDate(FTDate1.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate1.Text = ""
        End If

        If Trim(FTDate2.Text) <> "" Then
            FTDate2.Text = Trim(FTDate2.Text)
            If Not TIMS.IsDate1(FTDate2.Text) Then
                Errmsg += "結訓期間 的迄止日不是正確的日期格式" & vbCrLf
            End If
            If Errmsg = "" Then
                FTDate2.Text = CDate(FTDate2.Text).ToString("yyyy/MM/dd")
            End If
        Else
            FTDate2.Text = ""
        End If

        If Errmsg = "" Then
            If Me.FTDate1.Text <> "" AndAlso Me.FTDate2.Text <> "" Then
                If DateDiff(DateInterval.Day, CDate(FTDate1.Text), CDate(FTDate2.Text)) < 0 Then
                    Errmsg += "結訓期間 日期起迄有誤，迄日需大於起日" & vbCrLf
                End If
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call Search1()
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Private Sub btnExport1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1()

        If DataGrid1.Visible = False OrElse msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If

        Dim sFileName As String = ""
        sFileName = HttpUtility.UrlEncode("就服輔導就業名單.xls", System.Text.Encoding.UTF8)

        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集

        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)

        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType = "application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"

        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")

        ''套CSS值
        Common.RespWrite(Me, "<style>")
        Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False
        'DataGrid1.Columns(8).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid1.AllowPaging = True
        'DataGrid1.Columns(8).Visible = True

    End Sub

End Class
