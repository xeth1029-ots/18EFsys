Public Class TR_04_023_R
    Inherits AuthBasePage

    'Dim PrintNumC As Integer
    'Dim PrintNum1 As Integer
    'Dim PrintNum2 As Integer
    'Dim PrintNum3 As Integer
    'Dim PrintNum4 As Integer
    'Dim PrintNum5 As Integer
    'Dim PrintNum6 As Integer
    'Dim PrintNum7 As Integer
    'Dim PrintNum8 As Integer
    'Dim PrintNum9 As Integer
    'Dim PrintNum10 As Integer
    'Dim PrintNum11 As Integer
    'Dim PrintNum12 As Integer
    'Dim PrintNum13 As Integer

    'VIEW2
    'V_GETJOBBYDAYS
    'V_STUDENTCOUNT
    'STUD_GETJOBBYDAYS

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

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        If Not Page.IsPostBack Then
            PageControler1.Visible = False

            yearlist = TIMS.GetSyear(yearlist)
            Common.SetListItem(yearlist, sm.UserInfo.Years)
            yearlist.Items.Remove(yearlist.Items.FindByValue(""))
            DistID = TIMS.Get_DistID(DistID)
            If DistID.Items.FindByValue("") Is Nothing Then
                DistID.Items.Insert(0, New ListItem("全部", ""))
            End If
            Tcitycode = TIMS.Get_CityName(Tcitycode, TIMS.dtNothing)

            Call TIMS.Get_TPlan2(chkTPlanID0, chkTPlanID1, chkTPlanIDX, objconn)

            'center.Text = sm.UserInfo.OrgName
            'RIDValue.Value = sm.UserInfo.RID
            'PlanID.Value = sm.UserInfo.PlanID
            'OCID.Style("display") = "none"
            'Print.Visible = False
            btnExport1.Visible = False
            'PageControler1.Visible = False
            'msg.Text = cst_NODATAMsg11
            'Button3_Click(sender, e)

            If sm.UserInfo.LID = "2" Then   '2010/05/24 改成若是委訓單位登入下列欄位就不顯示
                Year_TR.Style("display") = "none"
                DistID_TR.Style("display") = "none"
                'PlanID_TR.Style("display") = "none"
                TPlanID0_TR.Style("display") = "none"
                TPlanID1_TR.Style("display") = "none"
                TPlanIDX_TR.Style("display") = "none"

                'Check_TR.Style("display") = "none"
                'Button2.Style("display") = "none"
            Else
                'LID: 0.1.
                Year_TR.Style("display") = "inline"
                DistID_TR.Style("display") = "inline"
                'PlanID_TR.Style("display") = "inline"
                TPlanID0_TR.Style("display") = "inline"
                TPlanID1_TR.Style("display") = "inline"
                TPlanIDX_TR.Style("display") = "inline"

                'Check_TR.Style("display") = "inline"
                'Button2.Style("display") = "inline"
            End If

        End If

        DistID.Attributes("onclick") = "ClearData();"
        'TPlanID.Attributes("onclick") = "ClearData();"
        Me.chkTPlanID0.Attributes("onclick") = "ClearData();"
        Me.chkTPlanID1.Attributes("onclick") = "ClearData();"
        Me.chkTPlanIDX.Attributes("onclick") = "ClearData();"

        Query.Attributes("OnClick") = "javascript:return chk()"

        '選擇全部轄區
        DistID.Attributes("onclick") = "SelectAll('DistID','DistHidden');"
        Tcitycode.Attributes("onclick") = "SelectAll('Tcitycode','TcityHidden');"
        '選擇全部訓練計畫
        'TPlanID.Attributes("onclick") = "SelectAll('TPlanID','TPlanHidden');"
        chkTPlanID0.Attributes("onclick") = "SelectAll('chkTPlanID0','TPlanID0HID');"
        chkTPlanID1.Attributes("onclick") = "SelectAll('chkTPlanID1','TPlanID1HID');"
        chkTPlanIDX.Attributes("onclick") = "SelectAll('chkTPlanIDX','TPlanIDXHID');"


        'If CheckData.Checked = True Then   '增加是否為單一機構查詢
        '    center.Text = sm.UserInfo.OrgName
        '    RIDValue.Value = sm.UserInfo.RID
        '    PlanID.Value = sm.UserInfo.PlanID

        '    Button2.Disabled = True
        '    OCID.Items.Clear()
        '    Class_TR.Style("display") = "none"
        '    'Table4.Style("display") = "none"
        '    If sm.UserInfo.RID = "A" Then
        '        Org_TR.Style("display") = "none"
        '    Else
        '        Org_TR.Style("display") = "inline"
        '    End If
        'Else
        '    Class_TR.Style("display") = "inline"
        '    Org_TR.Style("display") = "inline"
        '    Button2.Disabled = False
        '    'Table4.Style("display") = "none"
        'End If

        'CheckData.Attributes("OnClick") = "Enabled_OCID('" & sm.UserInfo.OrgName & "','" & sm.UserInfo.RID & "','" & sm.UserInfo.PlanID & "');"

        If sm.UserInfo.DistID = "000" Then
            DistID.Enabled = True
        Else
            DistID.SelectedValue = sm.UserInfo.DistID
            DistID.Enabled = False
        End If

        'DataGrid1_Detail_1.Visible = False
        'DataGrid1_Detail_2.Visible = False
        'DataGrid1_Detail_3.Visible = False
        'DataGrid1_Detail_4.Visible = False
        'DataGrid1_Detail_5.Visible = False
        'DataGrid1_Detail_6.Visible = False
        'Button3.Style("display") = "none"

        'Me.ViewState("SVID") = ""
        'If TIMS.Server_Path() = "DEMO" Then
        '    If sm.UserInfo.Years >= "2009" Then '測試機mark
        '        Me.ViewState("SVID") = TIMS.GetSVID(sm.UserInfo.TPlanID)
        '    End If
        'End If

    End Sub

#Region "NO USE"
    'Sub x1()
    '    Dim x1 As Integer = 0
    '    Dim x2 As Integer = 0
    '    PrintNumC = 30 * TIMS.Rnd1X() + 1 + 10
    '    x1 = PrintNumC
    '    x1 = x1 - 3
    '    If x1 > 0 Then PrintNum1 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum1
    '    If x1 > 0 Then PrintNum2 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum2
    '    If x1 > 0 Then PrintNum3 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum3
    '    If x1 > 0 Then PrintNum4 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum4
    '    If x1 > 0 Then PrintNum5 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum5
    '    If x1 > 0 Then PrintNum6 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum6
    '    If x1 > 0 Then PrintNum7 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum7
    '    If x1 > 0 Then PrintNum8 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum8
    '    If x1 > 0 Then PrintNum9 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum9
    '    If x1 > 0 Then PrintNum10 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum10
    '    If x1 > 0 Then PrintNum11 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum11
    '    If x1 > 0 Then PrintNum12 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum12
    '    If x1 > 0 Then PrintNum13 = x1 * TIMS.Rnd1X() : x1 = x1 - PrintNum13
    '    'PrintNum3 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum4 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum5 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum6 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum7 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum8 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum9 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum10 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum11 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum12 = 10 * TIMS.Rnd1X() + 1
    '    'PrintNum13 = 10 * TIMS.Rnd1X() + 1

    'End Sub
#End Region


    '查詢 SQL
    Sub Search1(ByVal sType As Integer)
        'sType 1:統計 2:明細

        '辦訓地縣市
        Dim TCityCode2 As String = ""
        TCityCode2 = ""
        For i As Integer = 1 To Tcitycode.Items.Count - 1
            If Tcitycode.Items.Item(i).Selected = True Then
                'If Tcitycode.Items.Item(i).Text <> "全部" Then
                'End If
                If TCityCode2 <> "" Then TCityCode2 += ","
                TCityCode2 += Tcitycode.Items.Item(i).Value
            End If
        Next

        '選擇轄區
        Dim DistIDValue As String = ""
        DistIDValue = ""
        For Each objitem As ListItem In Me.DistID.Items
            If objitem.Selected AndAlso objitem.Value <> "" Then
                If DistIDValue <> "" Then DistIDValue &= ","
                DistIDValue &= "'" & objitem.Value & "'"
            End If
        Next

        Dim sql As String = ""
        Select Case sType  'sType 1:統計 2:明細
            Case 1 'sType 1:統計 2:明細
                sql = "" & vbCrLf
                sql += " SELECT ROWNUM 序號" & vbCrLf
                sql += " ,cc.CTName 縣市別" & vbCrLf
                sql += " ,cc.planname 訓練計畫" & vbCrLf
                sql += " ,cc.orgname 訓練單位" & vbCrLf
                sql += " ,cc.classcname2 班級名稱" & vbCrLf
                sql += " ,CONVERT(varchar, cc.stdate, 111) 開訓日期" & vbCrLf
                sql += " ,CONVERT(varchar, cc.ftdate, 111) 結訓日期" & vbCrLf
                sql += " ,s.closecount 結訓人數" & vbCrLf
                sql += " ,dbo.NVL(v3.m1,0) 就業1個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m2,0) 就業2個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m3,0) 就業3個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m4,0) 就業4個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m5,0) 就業5個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m6,0) 就業6個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m7,0) 就業7個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m8,0) 就業8個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m9,0) 就業9個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m10,0) 就業10個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m11,0) 就業11個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m12,0) 就業12個月" & vbCrLf
                sql += " ,dbo.NVL(v3.m13,0) 就業12個月以上" & vbCrLf
                sql += " from VIEW2 cc " & vbCrLf
                sql += " join V_STUDENTCOUNT s on s.ocid =cc.ocid " & vbCrLf
                sql += " left join V_GETJOBBYDAYS v3 on v3.ocid=cc.ocid " & vbCrLf
                sql += " where 1=1" & vbCrLf
            Case 2 'sType 1:統計 2:明細
                sql = "" & vbCrLf
                sql += " SELECT ROWNUM 序號" & vbCrLf
                sql += " ,cc.CTName 縣市別" & vbCrLf
                sql += " ,cc.planname 訓練計畫" & vbCrLf
                sql += " ,cc.orgname 訓練單位" & vbCrLf
                sql += " ,cc.classcname2 班級名稱" & vbCrLf
                sql += " ,CONVERT(varchar, cc.stdate, 111) 開訓日期" & vbCrLf
                sql += " ,CONVERT(varchar, cc.ftdate, 111) 結訓日期" & vbCrLf
                sql += " ,s.closecount 結訓人數" & vbCrLf
                sql += " ,v3.name 姓名" & vbCrLf
                sql += " ,v3.DAYS 就業天數" & vbCrLf
                sql += " from VIEW2 cc " & vbCrLf
                sql += " join V_STUDENTCOUNT s on s.ocid =cc.ocid " & vbCrLf '依班。
                sql += " left join STUD_GETJOBBYDAYS v3 on v3.OCID=cc.OCID " & vbCrLf '依學員
                sql += " where 1=1" & vbCrLf
        End Select

        '年度
        If yearlist.SelectedValue <> "" Then
            sql += " and cc.Years='" & yearlist.SelectedValue & "'" & vbCrLf
        End If
        '轄區選擇
        If DistIDValue <> "" Then
            sql += " and cc.DistID IN (" & DistIDValue & ")" & vbCrLf
        End If
        '辦訓地縣市 
        If TCityCode2 <> "" Then
            sql += " and cc.CTID in (" & TCityCode2 & ")" & vbCrLf
        End If
        '大計畫
        Dim TPlanValue As String = ""
        TPlanValue = TIMS.Get_TPlan2Val(chkTPlanID0, chkTPlanID1, chkTPlanIDX)
        If TPlanValue <> "" Then
            sql += " and cc.TPlanID IN (" & TPlanValue & ")" & vbCrLf
        End If
        '開訓區間
        If STDate1.Text <> "" Then
            sql += " and cc.STDate>=" & TIMS.To_date(STDate1.Text) & vbCrLf '& "','yyyy/mm/dd')" & vbCrLf
        End If
        If STDate2.Text <> "" Then
            sql += " and cc.STDate<=" & TIMS.To_date(STDate2.Text) & vbCrLf 'convert(datetime, '" & STDate2.Text & "', 111)" & vbCrLf
        End If
        '結訓區間
        If FTDate1.Text <> "" Then
            sql += " and cc.FTDate>=" & TIMS.To_date(FTDate1.Text) & vbCrLf 'convert(datetime, '" & FTDate1.Text & "', 111)" & vbCrLf
        End If
        If FTDate2.Text <> "" Then
            sql += " and cc.FTDate<=" & TIMS.To_date(FTDate2.Text) & vbCrLf 'convert(datetime, '" & FTDate2.Text & "', 111)" & vbCrLf
        End If
        sql &= " ORDER BY cc.orgname,cc.classcname2 "

        Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            '.Parameters.Add("xxx", SqlDbType.VarChar).Value = ""
            dt.Load(.ExecuteReader())
        End With


        Table4.Visible = True
        Table4.Style("display") = "inline"
        DataGrid1.Visible = False
        'Print.Visible = False
        btnExport1.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count > 0 Then
            Table4.Visible = True
            Table4.Style("display") = "inline"
            DataGrid1.Visible = True
            'Print.Visible = True
            btnExport1.Visible = True
            PageControler1.Visible = True

            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
            'PageControler1.SqlPrimaryKeyDataCreate(sqlstr_class, "OCID", "DistID,PlanID,OCID,CyclType")
        Else
            Table4.Style("display") = "none"
            Common.MessageBox(Me, "查無資料")
            Exit Sub
        End If

    End Sub

    Protected Sub Query_Click(sender As Object, e As EventArgs) Handles Query.Click
        Call Search1(1)
    End Sub

    '匯出
    Protected Sub btnExport1_Click(sender As Object, e As EventArgs) Handles btnExport1.Click
        'Const Cst_功能欄位 As Integer = 14

        DataGrid1.AllowPaging = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1(1)

        Dim sFileName As String = ""
        sFileName = HttpUtility.UrlEncode("訓練成效統計表.xls", System.Text.Encoding.UTF8)
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
        'DataGrid1.Columns(Cst_功能欄位).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid1.AllowPaging = True
        'DataGrid1.Columns(Cst_功能欄位).Visible = True
        'Call TIMS.CloseDbConn(objconn)
    End Sub

    '匯出明細
    Protected Sub btnExport2_Click(sender As Object, e As EventArgs) Handles btnExport2.Click

        DataGrid1.AllowPaging = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Call Search1(2)

        Dim sFileName As String = ""
        sFileName = HttpUtility.UrlEncode("訓練成效明細表.xls", System.Text.Encoding.UTF8)
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
        'DataGrid1.Columns(Cst_功能欄位).Visible = False
        DataGrid1.EnableViewState = False  '把ViewState給關了

        Dim objStringWriter As New System.IO.StringWriter
        Dim objHtmlTextWriter As New System.Web.UI.HtmlTextWriter(objStringWriter)
        Div1.RenderControl(objHtmlTextWriter)
        Common.RespWrite(Me, Convert.ToString(objStringWriter))
        Response.End()

        DataGrid1.AllowPaging = True
        'DataGrid1.Columns(Cst_功能欄位).Visible = True

        'Call TIMS.CloseDbConn(objconn)
    End Sub
End Class