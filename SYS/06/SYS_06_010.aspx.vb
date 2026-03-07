Partial Class SYS_06_010
    Inherits AuthBasePage

    'SYS_TRANS_LOG
    '"排除關鍵字為『PASSWORD』內容"
    Const cst_PASSWORD_word As String = "PASSWORD"

    Const cst_myMode_新增 As String = "新增"
    Const cst_myMode_修改 As String = "修改"

    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '檢查Session是否存在 Start
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1 '分頁設定

        '處理[分頁設定元件]出現的時機
        PageControler1.Visible = False
        If PageControler1.PageDataGrid.Items.Count > 0 Then PageControler1.Visible = True

        If Not IsPostBack Then Call sCreate1() '頁面初始化
    End Sub

    '頁面初始化
    Sub sCreate1()
        ddlType1 = TIMS.Get_ddlFunction(ddlType1, 2)

        qDATE1.Text = ""
        qDATE2.Text = ""
        'ddlType1.SelectedIndex=0
        'ddlType2.SelectedIndex=0
        ddlType1.SelectedIndex = -1
        Common.SetListItem(ddlType1, "")
        ddlType2.SelectedIndex = -1
        Common.SetListItem(ddlType2, "")

        qFuncName.Text = ""
        'ddlType3.SelectedIndex=0
        ddlType3.SelectedIndex = -1
        Common.SetListItem(ddlType3, "")
        qAcc.Text = ""
        qDATE1.Text = TIMS.Cdate17(DateAdd(DateInterval.Day, -3, Date.Today))
        qDATE2.Text = TIMS.Cdate17(Date.Today)
    End Sub

    Protected Sub ddlType1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ddlType1.SelectedIndexChanged
        getFuncItem()
    End Sub

    '下拉式選單連動「功能項目類別」
    Sub getFuncItem()
        'ByVal sMenuVal As String
        'Optional ByVal sMenuVal As String=""
        Dim v_ddlType1 As String = TIMS.GetListValue(ddlType1)
        If v_ddlType1 <> "" AndAlso ddlType1.SelectedIndex <> -1 Then
            ddlType2.Items.Clear()

            'DbAccess.Open(objconn) '.OpenDbConn(objconn)
            Dim Sql As String = ""
            Sql &= " SELECT a.funid, a.name, a.spage, a.kind, a.levels, a.parent" & vbCrLf
            Sql &= " ,ISNULL(p.sort, a.sort) psort, a.sort, a.memo, a.valid" & vbCrLf
            Sql &= " ,(CASE a.levels WHEN '0' THEN (SELECT COUNT(funid) cnt FROM id_function WHERE parent=a.funid) ELSE 0 END) AS subs" & vbCrLf
            Sql &= " FROM ID_FUNCTION a" & vbCrLf
            Sql &= " LEFT JOIN ID_FUNCTION p ON p.funid=a.parent" & vbCrLf
            Sql &= " WHERE ISNULL(a.FState, ' ') NOT IN ('D')" & vbCrLf
            Sql &= " AND a.spage IS NULL AND a.levels='0' AND a.kind=@kind" & vbCrLf
            Sql &= " ORDER BY a.kind, psort, a.levels, a.sort" & vbCrLf
            Dim sCmd2 As New SqlCommand(Sql, objconn)
            Dim dt2 As New DataTable '= Nothing
            With sCmd2
                .Parameters.Clear()
                .Parameters.Add("kind", SqlDbType.VarChar).Value = v_ddlType1 ' ddlType1.SelectedValue
                dt2.Load(.ExecuteReader())
            End With
            With ddlType2
                .DataSource = dt2
                .DataValueField = "funid"
                .DataTextField = "name"
                .DataBind()
            End With

            ddlType2.Items.Insert(0, New ListItem("無", "0"))
            ddlType2.Items.Insert(0, New ListItem("全部", ""))
        End If
    End Sub

    '準備進行資料查詢作業
    Protected Sub bt_search_Click(sender As Object, e As EventArgs) Handles bt_search.Click
        '「顯示列數」相關設定
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        Call sSearch1() '進行資料查詢作業
    End Sub

    '資料查詢
    Sub sSearch1()
        qKeyWord1.Text = TIMS.ClearSQM(qKeyWord1.Text)
        qtablename.Text = TIMS.ClearSQM(qtablename.Text)
        If qtablename.Text <> "" Then qtablename.Text = UCase(qtablename.Text)
        qDATE1.Text = TIMS.ClearSQM(qDATE1.Text)
        qDATE2.Text = TIMS.ClearSQM(qDATE2.Text)
        qFuncName.Text = TIMS.ClearSQM(qFuncName.Text)
        qAcc.Text = TIMS.ClearSQM(qAcc.Text) '帳號換大寫
        Dim s_UQACC_lk As String = UCase(qAcc.Text) '帳號換大寫

        '沒有輸入值時，帶入當日資訊
        If qDATE1.Text = "" Then
            qDATE1.Text = If(flag_ROC, TIMS.Cdate17(Now.Date.ToString("yyyy/MM/dd")), TIMS.Cdate3(Now.Date.ToString("yyyy/MM/dd")))
        End If
        If qDATE2.Text = "" Then
            qDATE2.Text = If(flag_ROC, TIMS.Cdate17(Now.Date.ToString("yyyy/MM/dd")), TIMS.Cdate3(Now.Date.ToString("yyyy/MM/dd")))
        End If

        Dim myqDate1 As String = If(flag_ROC, TIMS.Cdate18(qDATE1.Text), TIMS.Cdate3(qDATE1.Text)).Replace("/", "-")  'edit，by:20181019
        Dim myqDate2 As String = If(flag_ROC, TIMS.Cdate18(qDATE2.Text), TIMS.Cdate3(qDATE2.Text)).Replace("/", "-")  'edit，by:20181019
        Dim myType1 As String = TIMS.GetListValue(ddlType1) '.SelectedValue.Trim
        Dim myType2 As String = TIMS.GetListValue(ddlType2) '.SelectedValue.Trim
        Dim myFuncName_like As String = qFuncName.Text
        Dim myType3 As String = TIMS.GetListValue(ddlType3) '.SelectedValue.Trim
        'Dim myAcc_like As String=s_UQACC_lk

        'DbAccess.Open(objconn) '.OpenDbConn(objconn)
        Dim sql As String = ""
        sql &= " SELECT so.UserID, so.UserName, so.TRANSTYPE, so.FUNCNAME" & vbCrLf
        sql &= " ,so.TransTime, so.Conditions, so.BeforeValues, so.AfterValues" & vbCrLf
        sql &= " ,so.TargetTable" & vbCrLf
        'sql &= " ,ROW_NUMBER() Over (Partition By so.UserID, so.UserName, so.TRANSTYPE, so.FUNCNAME, a.TransTime Order By so.TransTime Desc) xSort" & vbCrLf
        'sql &= " ,ROW_NUMBER() Over (Order By so.TransTime Desc) xSort" & vbCrLf
        sql &= " FROM ( SELECT a.UserID,c.NAME UserName" & vbCrLf
        sql &= " ,SUBSTRING(a.TransTime, 1, 19) TransTime" & vbCrLf
        sql &= " ,a.Conditions, a.BeforeValues, a.AfterValues,a.TargetTable" & vbCrLf
        sql &= " ,CASE a.TransType WHEN 'Insert' THEN '新增' WHEN 'Update' THEN '修改' WHEN 'Delete' THEN '刪除' ELSE '' END TRANSTYPE" & vbCrLf
        'sql &= " ,dbo.FN_GET_FUNCPATH(a.FUNCPATH) FUNCPATH" & vbCrLf
        sql &= " ,b.FUNCNAME,b.SPAGE" & vbCrLf
        If qKeyWord1.Text <> "" Then
            sql &= " ,(case when a.Conditions like '%'+@KeyWord1+'%' OR a.BeforeValues like '%'+@KeyWord1+'%' OR a.AfterValues like '%'+@KeyWord1+'%' then 1 end) KeyWord1"
        Else
            sql &= " ,0 KeyWord1"
        End If
        sql &= " FROM dbo.SYS_TRANS_LOG a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN dbo.V_FUNCTION3 b WITH(NOLOCK) ON b.SPAGE=dbo.FN_GET_FUNCPATH(a.FUNCPATH)" & vbCrLf
        sql &= " JOIN dbo.AUTH_ACCOUNT c WITH(NOLOCK) ON c.ACCOUNT=a.UserID" & vbCrLf
        sql &= " WHERE a.TargetTable<>'SYS_TRANS_LOG' AND a.FUNCPATH<>'/SelectPlan'" & vbCrLf
        If qtablename.Text <> "" Then sql &= " AND upper(a.TargetTable)=@TargetTable" & vbCrLf
        If myqDate1 <> "" Then sql &= " AND a.TransTime>=@qDate1" & vbCrLf
        If myqDate2 <> "" Then sql &= " AND a.TransTime<=@qDate2" & vbCrLf
        'V_FUNCTION3
        If myType1 <> "" OrElse myType2 <> "" Then
            If myType1 <> "" Then sql &= " AND b.KIND=@type1" & vbCrLf
            If myType2 <> "" Then sql &= " AND b.PARENT=@type2" & vbCrLf
        End If
        'V_FUNCTION2/V_FUNCTION3
        If myFuncName_like <> "" Then sql &= " AND b.FUNCNAME LIKE '%'+@fn1+'%'" & vbCrLf

        If myType3 <> "" Then sql &= " AND a.TransType=@type3" & vbCrLf
        'If myAcc_like <> "" Then sql &= " AND a.UserID LIKE '%'+ @acc1 +'%'" & vbCrLf
        If s_UQACC_lk <> "" Then sql &= " AND UPPER(c.ACCOUNT) LIKE '%'+@UQACC+'%'" & vbCrLf
        sql &= " ) so" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If qKeyWord1.Text <> "" Then sql &= " and so.KeyWord1=1" & vbCrLf
        'sql &= " AND so.xSort='1'" & vbCrLf
        sql &= " ORDER BY so.TransTime DESC" & vbCrLf

        Dim sCmd As New SqlCommand(sql, objconn)
        sCmd.CommandTimeout = 600
        Dim dt As New DataTable '= Nothing
        With sCmd
            .Parameters.Clear()
            If qKeyWord1.Text <> "" Then .Parameters.Add("KeyWord1", SqlDbType.VarChar).Value = qKeyWord1.Text
            If qtablename.Text <> "" Then .Parameters.Add("TargetTable", SqlDbType.VarChar).Value = qtablename.Text
            If myqDate1 <> "" Then .Parameters.Add("qDate1", SqlDbType.VarChar).Value = (myqDate1 & " 00:00:00")
            If myqDate2 <> "" Then .Parameters.Add("qDate2", SqlDbType.VarChar).Value = (myqDate2 & " 23:59:59")
            If myType1 <> "" Then .Parameters.Add("type1", SqlDbType.VarChar).Value = myType1
            If myType2 <> "" Then .Parameters.Add("type2", SqlDbType.VarChar).Value = myType2
            If myFuncName_like <> "" Then .Parameters.Add("fn1", SqlDbType.VarChar).Value = myFuncName_like
            If myType3 <> "" Then .Parameters.Add("type3", SqlDbType.VarChar).Value = myType3
            If s_UQACC_lk <> "" Then .Parameters.Add("UQACC", SqlDbType.VarChar).Value = s_UQACC_lk
            'dt.Load(.ExecuteReader())
        End With

        If TIMS.sUtl_ChkTest() Then
            Dim hPARMS As Hashtable = TIMS.CONVERTPAR2HASHTB(sCmd.Parameters)
            TIMS.WriteLog(Me, $"--{vbCrLf}{TIMS.GetMyValue5(hPARMS)}{vbCrLf}--#SYS_06_010,SSQL:{vbCrLf}{sql}")
        End If

        'Dim parms As Hashtable=New Hashtable()
        'If myqDate1 <> "" Then parms.Add("qDate1", myqDate1 + " 00:00:00")
        'If myqDate2 <> "" Then parms.Add("qDate2", myqDate2 + " 23:59:59")
        'If myType1 <> "" Then parms.Add("type1", myType1)
        'If myType2 <> "" Then parms.Add("type2", myType2)
        'If myFuncName_like <> "" Then parms.Add("fn1", myFuncName_like)
        'If myType3 <> "" Then parms.Add("type3", myType3)
        'If s_UQACC_lk <> "" Then parms.Add("UQACC", s_UQACC_lk)
        'Dim dt As DataTable=Nothing
        'TIMS.writeLog_1(Me, "SYS_06_010", sql, parms)

        Dim flag_error As Boolean = True '預設為錯誤 ! 查詢正確時為false 
        Try
            'dt=DbAccess.GetDataTable(sql, objconn, parms)
            dt.Load(sCmd.ExecuteReader())
            flag_error = False
        Catch ex As Exception
            Dim cst_fun_page_name As String = "##SYS_06_010.aspx, "
            Dim slogMsg1 As String = ""
            slogMsg1 &= cst_fun_page_name & "sql: " & sql & vbCrLf
            slogMsg1 &= cst_fun_page_name & "parms: " & TIMS.GetMyValue3(sCmd.Parameters) & vbCrLf
            'Call TIMS.SendMailTest(slogMsg1)
            Dim strErrmsg As String = ""
            strErrmsg &= "ex.Message:" & vbCrLf & ex.Message & vbCrLf
            strErrmsg &= "ex.ToString:" & vbCrLf & ex.ToString & vbCrLf
            strErrmsg &= "slogMsg1:" & vbCrLf & slogMsg1 & vbCrLf
            'strErrmsg=Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.SendMailTest(strErrmsg)
        End Try
        msg.Text = "查無資料"
        tb_Sch.Visible = False
        If flag_error Then
            Common.MessageBox(Me, TIMS.cst_ErrorMsg13)
            Exit Sub
        End If
        If dt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        If dt Is Nothing Then Return
        If dt.Rows.Count = 0 Then Return

        msg.Text = ""
        tb_Sch.Visible = True
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    '調整某欄位所顯示的內容
    Protected Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const cst_col_紀錄時間 As Integer = 4
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim originLogTime As String = Convert.ToString(drv("TransTime"))
                Dim tmpLogDate As String = If(Len(originLogTime) > 0, originLogTime.Substring(0, 10), "")
                '#Region "依照Web.config的REPLACE2ROC_YEARS(西元年/民國年置換)參數,調整[紀錄時間]顯示格式內容，by:20181019"
                Dim newLogDate As String = If(flag_ROC, TIMS.Cdate17(tmpLogDate), tmpLogDate)
                Dim tmpLogTime As String = If(Len(originLogTime) > 0, originLogTime.Substring(11, 8), "")
                e.Item.Cells(cst_col_紀錄時間).Text = newLogDate + " " + tmpLogTime

                Dim s_TargetTable As String = Convert.ToString(drv("TargetTable"))
                'Dim myMode As String=Convert.ToString(drv("TRANSTYPE")) ' e.Item.Cells(2).Text
                'Dim hidCol1 As Label=e.Item.FindControl("hidCol1") 'Convert.ToString(drv("Conditions"))
                'Dim hidCol2 As Label=e.Item.FindControl("hidCol2") 'Convert.ToString(drv("BeforeValues"))
                'Dim hidCol3 As Label=e.Item.FindControl("hidCol3") 'Convert.ToString(drv("AfterValues"))
                Dim lblInfo As Label = e.Item.FindControl("lblInfo")

                Dim myStr As String = ""
                Select Case Convert.ToString(drv("TRANSTYPE"))'myMode
                    Case cst_myMode_新增
                        Dim iArray() As String = Convert.ToString(drv("BeforeValues")).Split(",") 'hidCol2.Text.Split(",")
                        Dim t_Str As String = ""
                        'Dim i As Integer=0
                        For i As Integer = 0 To iArray.Length - 1
                            If iArray(i).Trim.ToUpper.Contains(cst_PASSWORD_word) Then Continue For
                            If t_Str = "" Then t_Str += iArray(i) Else t_Str += (", " + iArray(i))
                        Next
                        'myStr="<font color='#0066FF'><b>" + "新增內容：" + "</b></font>" + "<br/>" + t_Str
                        myStr = String.Format("<font color='#0066FF'><b>新增內容：</b></font> {1} <br/>{0} ", t_Str, s_TargetTable)
                        lblInfo.Text = myStr

                    Case cst_myMode_修改
                        Dim uArray1_o() As String = Convert.ToString(drv("BeforeValues")).Split(",") 'hidCol2.Text.Split(",")
                        Dim uArray2_o() As String = Convert.ToString(drv("AfterValues")).Split(",") 'hidCol3.Text.Split(",")
                        Dim uArray1_n() As String = New String(uArray1_o.Length) {}
                        Dim uArray2_n() As String = New String(uArray2_o.Length) {}

                        '"排除關鍵字為『PASSWORD』內容"
                        For i As Integer = 0 To uArray1_o.Length - 1
                            If uArray1_o(i).ToUpper.Contains(cst_PASSWORD_word) Then uArray1_n(i) = "" Else uArray1_n(i) = TIMS.ClearSQM(uArray1_o(i))
                        Next
                        For j As Integer = 0 To uArray2_o.Length - 1
                            If uArray2_o(j).ToUpper.Contains(cst_PASSWORD_word) Then uArray2_n(j) = "" Else uArray2_n(j) = TIMS.ClearSQM(uArray2_o(j))
                        Next

                        '比對[修改前]與[修改後]的內容是否相同"
                        Dim before_V As String = ""
                        Dim after_V As String = ""
                        If uArray1_n.Length = uArray2_n.Length Then
                            'Dim k As Integer=0
                            For k As Integer = 0 To uArray1_n.Length - 1
                                If Convert.ToString(uArray1_n(k)) <> "" And Convert.ToString(uArray2_n(k)) <> "" Then
                                    If Not uArray1_n(k).Equals(uArray2_n(k)) Then
                                        If before_V = "" Then before_V += uArray1_n(k) Else before_V += (", " + uArray1_n(k))
                                        If after_V = "" Then after_V += uArray2_n(k) Else after_V += (", " + uArray2_n(k))
                                    End If
                                End If
                            Next
                        Else
                            before_V = "－"
                            after_V = "－"
                        End If

                        myStr = ""
                        'Conditions-條件
                        myStr += String.Format("<font color='#227700'><b>修改以下條件之資料：</b></font> {1} <br/>{0}<br/>", TIMS.ClearSQM(drv("Conditions")), s_TargetTable)
                        myStr += String.Format("<span style='background-color:#FFFF77;'>" + "異動欄位修改前：" + "</span>" + "<br/>{0}<br/>", before_V)
                        myStr += String.Format("<span style='background-color:#FFFF77;'>" + "異動欄位修改後：" + "</span>" + "<br/>{0}", after_V)
                        lblInfo.Text = myStr

                    Case Else
                        myStr = String.Format("<font color='#FF3333'><b>刪除以下條件之資料：</b></font>  {1} <br/>{0}", TIMS.ClearSQM(drv("Conditions")), s_TargetTable)
                        lblInfo.Text = myStr

                End Select

        End Select

    End Sub

    Sub Utl_EXP1()

        'Const cst_功能欄 As Integer=4
        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        sSearch1()

        'DataGrid1.Columns(4).Visible=False
        'Call sUtl_ViewStateValue()  '設定 ViewState Value
        'Call search1()
        'Dim txtTitle As String="產業關鍵字設定"
        ''Dim txtTitle As String=GetAssemblyTitle()

        'txtTitle="EXP_" & TIMS.GetDateNo2 & ".xls" 'txtTitle & ".xls" '"離退訓人數統計表.xls"
        Dim sFileName As String = $"EXP_{TIMS.GetDateNo2()}.xls"
        sFileName = HttpUtility.UrlEncode(sFileName, System.Text.Encoding.UTF8)
        Response.Clear()
        Response.Buffer = True
        Response.Charset = "UTF-8" '設定字集
        Response.AppendHeader("Content-Disposition", "attachment;filename=" & sFileName)
        Response.ContentEncoding = System.Text.Encoding.GetEncoding("UTF-8")
        'Response.ContentType="application/vnd.ms-excel " '內容型態設為Excel
        '文件內容指定為Excel
        Response.ContentType = "application/ms-excel;charset=utf-8"
        Common.RespWrite(Me, "<meta http-equiv=Content-Type content=text/html;charset=UTF-8>")
        ''套CSS值
        'Common.RespWrite(Me, "<style>")
        'Common.RespWrite(Me, "td{mso-number-format:""\@"";}")
        'Common.RespWrite(Me, "</style>")

        DataGrid1.AllowPaging = False '關閉分頁功能
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        'DataGrid1.Visible=False
    End Sub

    Protected Sub bt_export1_Click(sender As Object, e As EventArgs) Handles bt_export1.Click
        Utl_EXP1()
    End Sub
End Class
