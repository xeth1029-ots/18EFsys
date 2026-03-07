Partial Class TR_05_014_R
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

            '關鍵字詞建立
            Call CreateItem()

            '選擇全部轄區
            DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
            '選擇全部縣市
            CTID.Attributes("onclick") = "SelectAll('CTID','CTIDHidden');"
            '選擇全部訓練計畫
            TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
            If sm.UserInfo.DistID <> "000" Then
                Common.SetListItem(DistID, sm.UserInfo.DistID) '轄區
            End If
            Common.SetListItem(TPlanID, sm.UserInfo.TPlanID) '計畫

            '登入年度 取得有效年月
            Call Get_OldYearMonth()

            '就業日期區間
            Common.SetListItem(cbl_JOBMDATE_MM, "1") '1月
            'If TIMS.sUtl_ChkTest() Then
            '    Common.SetListItem(cbl_JOBMDATE_MM, "3") '3月
            'End If

            If sm.UserInfo.DistID <> "000" Then
                DistID.Enabled = False
            End If

            btnSearch.Attributes("onclick") = "return chkSearch();"
            'btnExport.Attributes("onclick") = "return chkSearch();"
        End If
    End Sub

    '關鍵字詞建立
    Sub CreateItem()
        ''轄區別
        'DistID = TIMS.Get_DistID(DistID)
        'DistID.Items.Insert(0, New ListItem("全部", ""))
        '轄區別
        DistID = TIMS.Get_DistID(DistID)
        DistID.Items.Remove(DistID.Items.FindByValue(""))
        DistID.Items.Insert(0, New ListItem("全部", ""))
        '縣市
        Dim dt As DataTable
        Dim sql As String = ""
        sql = "SELECT a.CTID,a.CTName FROM ID_City a ORDER BY a.CTID"
        dt = DbAccess.GetDataTable(sql, objconn)
        With CTID
            .DataSource = dt
            .DataTextField = "CTName"
            .DataValueField = "CTID"
            .DataBind()
        End With
        CTID.Items.Insert(0, New ListItem("全部", ""))
        '計畫別
        TPlanID = TIMS.Get_TPlan(TPlanID, , 1, "Y", "TPlanID not in ('28','15','36','54')")
        '年度 '計畫年度
        Syear = TIMS.GetSyear(Syear)
        '月份
        ddlMonths.Items.Clear()
        For i As Integer = 1 To 12
            ddlMonths.Items.Add(New ListItem(i & "月份", i))
        Next
        ddlMonths.Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))

        '就業日期區間
        cbl_JOBMDATE_MM.Items.Clear()
        'cbl_JOBMDATE_MM.Items.Add(New ListItem("當月", 0))
        cbl_JOBMDATE_MM.Items.Add(New ListItem("1個月", 1))
        cbl_JOBMDATE_MM.Items.Add(New ListItem("2個月", 2))
        'cbl_JOBMDATE_MM.Items.Add(New ListItem("3個月", 3))

    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        'If Me.DistID.SelectedValue = "" Then
        '    Errmsg += "請選擇轄區" & vbCrLf
        'End If
        'If Me.CTID.SelectedValue = "" Then
        '    Errmsg += "請選擇縣市" & vbCrLf
        'End If
        'If Me.TPlanID.SelectedValue = "" Then
        '    Errmsg += "請選擇訓練計畫" & vbCrLf
        'End If
        If Me.Syear.SelectedValue = "" Then
            Errmsg += "請選擇結訓年度" & vbCrLf
        End If
        If Me.ddlMonths.SelectedValue = "" Then
            Errmsg += "請選擇結訓月份" & vbCrLf
        End If
        'If Me.cbl_JOBMDATE_MM.SelectedValue = "" Then
        '    Errmsg += "請選擇就業區間" & vbCrLf
        'End If

        Dim j As Integer = 0
        Dim CBLobj As CheckBoxList
        j = 0
        CBLobj = DistID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇轄區" & vbCrLf
        j = 0
        CBLobj = CTID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇縣市" & vbCrLf
        j = 0
        CBLobj = TPlanID
        For i As Integer = 1 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇訓練計畫" & vbCrLf
        j = 0
        CBLobj = cbl_JOBMDATE_MM
        For i As Integer = 0 To CBLobj.Items.Count - 1
            Dim objitem As ListItem = CBLobj.Items(i)
            If objitem.Selected = True Then
                j += 1
                Exit For
            End If
        Next
        If j = 0 Then Errmsg += "請選擇就業區間" & vbCrLf

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '登入年度 取得有效年月
    Sub Get_OldYearMonth()
        Dim tmpDate As Date
        Dim sDate1 As String
        Dim rYears As String = sm.UserInfo.Years
        Dim rMonth As String = Now.Month
        sDate1 = CDate(sm.UserInfo.Years & "/" & CStr(Now.Month) & "/01").ToString("yyyy/MM/dd")

        tmpDate = DateAdd(DateInterval.Month, -2, CDate(sDate1)) '回2個月前
        rYears = CStr(tmpDate.Year)
        rMonth = CStr(tmpDate.Month)

        Common.SetListItem(Syear, rYears)  '計畫年度
        Common.SetListItem(ddlMonths, rMonth) '當月
    End Sub

    '登入年度 取得有效年月
    Sub Get_OkYearMonth2(ByVal sDate1 As String, ByRef rJobMdateM1 As String, ByRef rJobMdateM2 As String)
        Dim tmpDate As Date
        Dim sYears1 As String = ""
        Dim sMonth1 As String = ""
        If sDate1 <> "" Then
            sYears1 = Mid(sDate1, 1, 4)
            sMonth1 = Mid(sDate1, 5, 2)
            tmpDate = CDate(sYears1 & "/" & sMonth1 & "/01").ToString("yyyy/MM/dd")
            rJobMdateM1 = DateAdd(DateInterval.Month, +1, tmpDate).ToString("yyyyMM") '1個月後
            rJobMdateM2 = DateAdd(DateInterval.Month, +2, tmpDate).ToString("yyyyMM") '2個月後
        End If
    End Sub

    '統計 SQL 查詢
    Sub Search1()

        Dim tmpStr As String = ""
        Dim itemDist As String = "" '轄區
        Dim itemCTID As String = "" '縣市
        Dim itemTPlanID As String = "" '計畫
        Dim itemJOBMDATE As String = "" '就業區間

        Dim CBLobj As CheckBoxList
        tmpStr = ""
        CBLobj = DistID
        For Each objitem As ListItem In CBLobj.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "'" & objitem.Value & "'"
            End If
        Next
        If tmpStr <> "" Then itemDist = tmpStr

        tmpStr = ""
        CBLobj = CTID
        For Each objitem As ListItem In CBLobj.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "'" & objitem.Value & "'"
            End If
        Next
        If tmpStr <> "" Then itemCTID = tmpStr

        tmpStr = ""
        CBLobj = TPlanID
        For Each objitem As ListItem In CBLobj.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If tmpStr <> "" Then tmpStr += ","
                tmpStr += "'" & objitem.Value & "'"
            End If
        Next
        If tmpStr <> "" Then itemTPlanID = tmpStr

        Dim sJobflag1 As Boolean = False
        Dim sJobflag2 As Boolean = False
        sJobflag1 = False
        sJobflag2 = False
        CBLobj = cbl_JOBMDATE_MM
        For Each objitem As ListItem In CBLobj.Items
            If objitem.Selected = True Then
                Select Case objitem.Value
                    Case "1"
                        sJobflag1 = True
                    Case "2"
                        sJobflag2 = True
                End Select
            End If
        Next

        Dim sFTDate1 As String = ""
        Dim sJobMdateM1 As String = ""
        Dim sJobMdateM2 As String = ""
        sFTDate1 = Me.Syear.SelectedValue & Right("0" & Me.ddlMonths.SelectedValue, 2)
        '取得有效1個月後與2個月後
        Call Get_OkYearMonth2(sFTDate1, sJobMdateM1, sJobMdateM2)

        Dim modeSql As String = ""
        'modeSql = " =1"
        modeSql = " in (1,2)"
        Select Case rblJobMode.SelectedValue
            Case "1"
                modeSql = " =1"
            Case "2"
                modeSql = " =2"
        End Select

        Dim Sql As String = ""
        Sql = "" & vbCrLf
        Sql += " select ip.Years  年度 " & vbCrLf
        Sql += " ,ip.distname  轄區 " & vbCrLf
        Sql += " ,ip.planname  計畫名稱 " & vbCrLf
        Sql += " ,oo.orgname  訓練機構 " & vbCrLf
        'Sql += " --,ip.tplanid '訓練計畫代碼'" & vbCrLf
        'Sql += " --,ip.planname+'('+ip.seq+')' '訓練計畫'" & vbCrLf
        Sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) 班級名稱" & vbCrLf
        Sql += " ,vz.CTName as  訓練縣市 " & vbCrLf
        'Sql += " --,convert(varchar,cc.stdate,111) '開訓日期'" & vbCrLf
        Sql += " ,CONVERT(varchar, cc.ftdate, 111)  結訓日期  " & vbCrLf
        'Sql += " --,gcs.clsNum '開訓人數'" & vbCrLf
        'Sql += " --,gcs.closeNum   '結訓人數'" & vbCrLf
        'Sql += " --,gcs.jobNum1 '就業人數1'" & vbCrLf
        Sql += " ,ss.name  姓名  " & vbCrLf
        Sql += " ,ss.idno  身分證號  " & vbCrLf
        Sql += " ,kd.Name  身分別 " & vbCrLf
        If sJobflag1 Then Sql += " ,CONVERT(varchar, sg1.MDate, 111)  ""加保日期(1)"" " & vbCrLf
        If sJobflag2 Then Sql += " ,CONVERT(varchar, sg2.MDate, 111)  ""加保日期(2)"" " & vbCrLf

        Sql += " from view_plan ip" & vbCrLf
        Sql += " JOIN class_classinfo cc on cc.planid =ip.planid" & vbCrLf
        Sql += " join plan_planinfo pp on pp.planid=cc.planid and pp.comidno=cc.comidno and pp.seqno=cc.seqno" & vbCrLf
        Sql += " join AUTH_RELSHIP aa on aa.RID=cc.RID" & vbCrLf
        Sql += " join view_traintype tt on tt.tmid=cc.tmid" & vbCrLf
        Sql += " join org_orginfo oo on oo.comidno =cc.comidno" & vbCrLf
        Sql += " join view_zipname vz on vz.ZipCode=cc.TaddressZip" & vbCrLf
        Sql += " join class_studentsofclass cs on cs.ocid =cc.ocid and cs.StudStatus not in (2,3) " & vbCrLf
        Sql += " 	 and cs.closedate<=getdate() and cc.ftdate<=getdate() " & vbCrLf
        Sql += " join stud_studentinfo ss on ss.sid =cs.sid" & vbCrLf
        Sql += " join Key_Identity kd on kd.IdentityID =cs.mIdentityid  " & vbCrLf
        If sJobflag1 Then
            Sql += " left join Stud_GetJobState3 sg1 on sg1.CPoint=1 and sg1.socid =cs.socid and sg1.Mode_" & modeSql & vbCrLf
            Sql += " and convert(varchar(6), sg1.MDate, 112) ='" & sJobMdateM1 & "'" & vbCrLf
        End If
        If sJobflag2 Then
            Sql += " left join Stud_GetJobState3 sg2 on sg2.CPoint=1 and sg2.socid =cs.socid and sg2.Mode_" & modeSql & vbCrLf
            Sql += " and convert(varchar(6), sg2.MDate, 112) ='" & sJobMdateM2 & "'" & vbCrLf
        End If

        Sql += " WHERE 1=1" & vbCrLf
        'Sql += " --AND ip.planname like '%職前%'" & vbCrLf
        'Sql += " --AND ip.planname like '%自辦職前訓練%'" & vbCrLf
        Sql += " and ip.TPlanID not in ('28','15','36','54') " & vbCrLf
        'Sql += " --and cc.TPropertyID<>'2' --2委託訓練 " & vbCrLf
        'Sql += " --and cc.TPropertyID<>'1' --1在職(進修)" & vbCrLf
        Sql += " and cc.TPropertyID='0'--0職前" & vbCrLf
        Sql += " and cc.NotOpen='N' -- '排除不開班" & vbCrLf
        Sql += " and cc.IsSuccess='Y'-- '轉入成功資料" & vbCrLf
        Sql += " and cc.Evta_NoShow is null " & vbCrLf
        'Sql += " and convert(varchar(6),cc.ftdate,112) =''" & vbCrLf

        If itemDist <> "" Then
            Sql += " and ip.distid IN (" & itemDist & ")" & vbCrLf
        End If
        If itemCTID <> "" Then
            Sql += " and vz.CTID IN (" & itemCTID & ")" & vbCrLf
        End If
        If itemTPlanID <> "" Then
            Sql += " and ip.TPlanID IN (" & itemTPlanID & ")" & vbCrLf
        End If

        Sql += " and convert(varchar(6), cc.ftdate, 112) ='" & sFTDate1 & "'" & vbCrLf

        Dim dt As DataTable
        Try
            dt = DbAccess.GetDataTable(Sql, objconn)
        Catch ex As Exception
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(sql))
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
            'dt.DefaultView
            msg.Text = ""
            PageControler1.Visible = True
            DataGrid1.Visible = True

            PageControler1.PageDataTable = dt
            'PageControler1.Sort = "DistID "
            PageControler1.ControlerLoad()
        End If
        'Else
        'Common.MessageBox(Me, "查無資料")
    End Sub

    '匯出 明細資料
    Sub sExport1()

        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        DataGrid1.AllowPaging = False '關閉分頁
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1()

        If DataGrid1.Visible = False OrElse msg.Text <> "" Then
            Common.MessageBox(Page, msg.Text)
            Exit Sub
        End If

        Dim sFileName As String = ""
        '勞保勾稽查詢
        sFileName = HttpUtility.UrlEncode("勞保勾稽查詢.xls", System.Text.Encoding.UTF8)

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

        DataGrid1.AllowPaging = False '關閉分頁
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)

        Common.RespWrite(Me, TIMS.sUtl_AntiXss(Convert.ToString(objStringWriter)))
        Response.End()

        DataGrid1.AllowPaging = True '開啟分頁
    End Sub

    '查詢 明細資料
    Private Sub btnSearch_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSearch.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        '查詢
        Call Search1()
    End Sub

    '覆寫，不執行 MyBase.VerifyRenderingInServerForm 方法，解決執行 RenderControl 產生的錯誤   
    Public Overrides Sub VerifyRenderingInServerForm(ByVal Control As System.Web.UI.Control)

    End Sub

    '匯出
    Private Sub btnExport_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnExport.Click
        Call sExport1()
    End Sub
End Class
