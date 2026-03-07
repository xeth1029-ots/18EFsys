Partial Class SD_14_007
    Inherits AuthBasePage

    'Private currentPage As Integer
    'Private totalPages As Integer
    'ReportQuery 'SQControl.aspx
    'SD_14_007_b'已轉班'SD_14_007_1_b'未轉班'SD_14_007_2_b'變更審核
    'SD_14_007'已轉班 // 'SD_14_007_1'未轉班 // 'SD_14_007_2'變更審核

    '師資/助教名冊 'SD_14_007*.jrxml/0:未轉班/1:已轉班/2:變更待審
    Const cst_reportFN0 As String = "SD_14_007"
    Const cst_reportFN1 As String = "SD_14_007_1"
    Const cst_reportFN2 As String = "SD_14_007_2"

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
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        'Dim Kind, sql, OrgIdKind As String
        '20090505 andy  edit
        '-------------------
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = DataGrid2
        '-------------------

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Me.Radio1.SelectedIndex = 0
            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        ROC_Years.Value = sm.UserInfo.Years - 1911
        'Years2.Value = sm.UserInfo.Years

        'SMpath.Value = TIMS.Server_Path
        'btnPrint1.Attributes("onclick") = "return CheckPrint('" & ReportQuery.GetSmartQueryPath & "');"
        btnPrint1.Attributes("onclick") = "javascript:return CheckPrint();"
        Button1.Attributes("onclick") = "return CheckSearch();"
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call ClearValue1()
        Call bindData() '頁籤
        '28:產業人才投資方案
        Dim sTitle As String = TIMS.GetTPlanName(sm.UserInfo.TPlanID, objconn)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Dim PNAME As String = PlanPoint.SelectedItem.Text
            Select Case PlanPoint.SelectedValue
                Case "1", "2"
                    sTitle &= "（" & PNAME & "）"
            End Select
        End If
        KindValue.Value = sTitle
    End Sub

    Private Sub bindData()

        Select Case Radio1.SelectedValue
            Case "0" '未轉班
                TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid2) '顯示列數不正確
                If Me.TxtPageSize.Text <> Me.DataGrid2.PageSize Then Me.DataGrid2.PageSize = Me.TxtPageSize.Text
                CreatePlan(1)
            Case "1" '已轉班
                TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確
                If Me.TxtPageSize.Text <> Me.DataGrid1.PageSize Then Me.DataGrid1.PageSize = Me.TxtPageSize.Text
                CreateClass()
            Case "2" '變更待審
                TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid2) '顯示列數不正確
                If Me.TxtPageSize.Text <> Me.DataGrid2.PageSize Then Me.DataGrid2.PageSize = Me.TxtPageSize.Text
                CreatePlan(2)
            Case Else
                TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1) '顯示列數不正確
                Common.MessageBox(Me, "未選擇，班級狀態!")
                Exit Sub
        End Select
    End Sub

    '已轉班
    Private Sub CreateClass()
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then Exit Sub
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        SelectValue.Value = ""

        Dim sql As String = ""
        sql &= " SELECT a.OCID" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,a.STDate ,a.FTDate" & vbCrLf
        sql &= " ,b.OrgName ,b.ORGKINDGW" & vbCrLf
        sql &= " FROM Class_ClassInfo a" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME b ON a.RID=b.RID AND b.RelShip LIKE '" & RelShip & "%'" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql += " AND a.PlanID IN (SELECT PlanID FROM ID_Plan WHERE TPlanID='" & sm.UserInfo.TPlanID & "' AND Years='" & sm.UserInfo.Years & "')" & vbCrLf
        Else
            sql += " AND a.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        If STDate1.Text <> "" Then sql += " AND a.STDate >= " & TIMS.To_date(STDate1.Text) & vbCrLf
        If STDate2.Text <> "" Then sql += " AND a.STDate <= " & TIMS.To_date(STDate2.Text) & vbCrLf
        If FTDate1.Text <> "" Then sql += " AND a.FTDate >= " & TIMS.To_date(FTDate1.Text) & vbCrLf
        If FTDate2.Text <> "" Then sql += " AND a.FTDate <= " & TIMS.To_date(FTDate2.Text) & vbCrLf

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                sql &= " AND b.OrgKind <> 10" & vbCrLf
            Else
                sql &= " AND b.OrgKind=10" & vbCrLf
            End If
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        If TIMS.dtNODATA(dt) Then Return
        'If dt.Rows.Count > 0 Then End If

        DataGridTable.Visible = True
        msg.Text = ""
        DataGrid1.Visible = True
        PageControler1.Visible = True
        DataGrid2.Visible = False
        PageControler2.Visible = False
        'PageControler1.SqlString = sql
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
            Case ListItemType.Item, ListItemType.AlternatingItem
                If e.Item.ItemType = ListItemType.Item Then
                    e.Item.CssClass = ""
                End If
                Dim drv As DataRowView = e.Item.DataItem
                Dim OCID As HtmlInputCheckBox = e.Item.FindControl("OCID")

                OCID.Value = drv("OCID")
                OCID.Attributes("onclick") = "SelectItem(this.checked,this.value);"

                Dim OCIDArray As Array = Split(SelectValue.Value, ",")
                For i As Integer = 0 To OCIDArray.Length - 1
                    If drv("OCID").ToString = OCIDArray(i) Then
                        OCID.Checked = True
                    End If
                Next
        End Select
    End Sub

    Private Sub CreatePlan(ByVal statusKind As Int16)   'statusKind (1=未轉班, 2=變更待審)
        Me.PLANIDValue.Value = ""
        Me.SeqNoValue.Value = ""
        Me.ComIDNOValue.Value = ""

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = ""
        sql &= " WITH WAR1 AS ( SELECT * FROM PLAN_REVISE WHERE AltDataID=11 AND ReviseStatus IS NULL" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND PlanID IN (SELECT ip.PLANID FROM ID_PLAN ip" & vbCrLf
            sql &= " WHERE ip.TPlanID='" & sm.UserInfo.TPlanID & "' AND ip.Years='" & sm.UserInfo.Years & "')" & vbCrLf
        Else
            sql &= " AND PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If
        sql &= " )" & vbCrLf

        sql &= " SELECT a.PlanID ,a.ComIDNO ,a.SeqNo" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,a.STDate,a.FDDate FTDate" & vbCrLf
        sql &= " ,c.OrgName,b.ORGKINDGW" & vbCrLf
        Select Case statusKind
            Case "1" 'statusKind (1=未轉班, 2=變更待審)
            Case "2" 'statusKind (1=未轉班, 2=變更待審)
                'PPIPK : yyyy-MM-dd 'CONVERT(VARCHAR(10),f.CDate, 120)
                sql &= " ,CONVERT(VARCHAR(10),f.CDate, 120) CDate ,f.SubSeqNO ,f.AltDataID ,f.ModifyDate" & vbCrLf
        End Select
        sql &= " FROM dbo.PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN dbo.ID_PLAN ip ON ip.planid=a.planid" & vbCrLf
        sql &= " JOIN dbo.VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO c ON c.comidno=a.comidno" & vbCrLf
        '**by Milor 20080422，加入IsApprPaper='Y'的條件，濾除掉草稿
        sql &= " AND a.IsApprPaper='Y'" & vbCrLf
        Select Case statusKind
            Case "1" 'statusKind (1=未轉班, 2=變更待審)
                sql &= " AND a.TransFlag='N'" & vbCrLf
            Case "2" 'statusKind (1=未轉班, 2=變更待審)
                sql &= " AND a.TransFlag='Y'" & vbCrLf
        End Select
        Select Case statusKind
            Case "1" 'statusKind (1=未轉班, 2=變更待審)
            Case "2" 'statusKind (1=未轉班, 2=變更待審)
                sql &= " JOIN WAR1 f ON a.PlanID=f.PlanID AND a.ComIDNO=f.ComIDNO AND a.SeqNO=f.SeqNO" & vbCrLf
        End Select
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "'" & vbCrLf
            sql &= " AND ip.DistID='" & sm.UserInfo.DistID & "'" & vbCrLf
        End If
        If RelShip <> "" Then
            sql &= " AND b.RelShip LIKE '" & RelShip & "%'" & vbCrLf
        End If

        '28:產業人才投資方案
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                sql &= " AND b.OrgKind<>10" & vbCrLf
            Else
                sql &= " AND b.OrgKind=10" & vbCrLf
            End If
        End If
        sql &= " ORDER BY a.STDate,a.ClassName" & vbCrLf

        'Select Case statusKind'    Case "1" 'statusKind (1=未轉班, 2=變更待審)'    Case "2" 'statusKind (1=未轉班, 2=變更待審)'End Select

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        DataGridTable.Visible = False
        msg.Text = "查無資料"

        If TIMS.dtNODATA(dt) Then Return
        'If dt.Rows.Count > 0 Then End If

        DataGridTable.Visible = True
        msg.Text = ""
        DataGrid1.Visible = False
        PageControler1.Visible = False
        DataGrid2.Visible = True
        PageControler2.Visible = True
        PageControler2.PageDataTable = dt
        PageControler2.ControlerLoad()
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Const cst_申請變更時間 As Integer = 5
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = ""
                'e.Item.Cells(8).Style.Add("display", "none")
                'e.Item.Cells(9).Style.Add("display", "none")
                'e.Item.Cells(10).Style.Add("display", "none")
                e.Item.Cells(cst_申請變更時間).Visible = False
                If Radio1.SelectedValue = "2" Then e.Item.Cells(cst_申請變更時間).Visible = True '0.未轉班 1.已轉班 2.變更待審
            Case ListItemType.Item, ListItemType.AlternatingItem
                'e.Item.Cells(8).Style.Add("display", "none")
                'e.Item.Cells(9).Style.Add("display", "none")
                'e.Item.Cells(10).Style.Add("display", "none")
                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = "SD_TD2"
                Dim drv As DataRowView = e.Item.DataItem
                Dim SeqNo As HtmlInputCheckBox = e.Item.FindControl("SeqNo")
                Dim PlanID As HtmlInputHidden = e.Item.FindControl("PlanID")
                Dim ComIDNO As HtmlInputHidden = e.Item.FindControl("ComIDNO")
                Dim SubSeqNo As HtmlInputHidden = e.Item.FindControl("SubSeqNo")
                Dim CDateValue As HtmlInputHidden = e.Item.FindControl("CDateValue")
                Dim TechID As HtmlInputHidden = e.Item.FindControl("TechID")
                Dim hPTDRID As HtmlInputHidden = e.Item.FindControl("hPTDRID")

                e.Item.Cells(cst_申請變更時間).Visible = False
                If Radio1.SelectedValue = "2" Then '0.未轉班 1.已轉班 2.變更待審
                    e.Item.Cells(cst_申請變更時間).Visible = True
                    e.Item.Cells(cst_申請變更時間).Text = Convert.ToString(drv("ModifyDate"))
                End If
                SeqNo.Value = drv("SeqNo")
                PlanID.Value = drv("PlanID")
                ComIDNO.Value = drv("ComIDNO")
                'PPIPK
                SubSeqNo.Value = ""
                CDateValue.Value = ""
                TechID.Value = ""
                If Radio1.SelectedValue = "2" Then
                    SubSeqNo.Value = drv("SubSeqNo")
                    CDateValue.Value = Convert.ToString(drv("CDate")) '.ToString("yyyy-MM-dd")
                    TechID.Value = TIMS.GetTechID(objconn, PlanID.Value, ComIDNO.Value, SeqNo.Value, SubSeqNo.Value, CDateValue.Value)
                    hPTDRID.Value = TIMS.GetPTDRID(objconn, PlanID.Value, ComIDNO.Value, SeqNo.Value, SubSeqNo.Value, CDateValue.Value)
                End If
                'If Radio1.SelectedValue <> 2 Then SeqNo.Attributes("onclick") = "SelectSeqNo(this.checked,this.value,'" & PlanID.Value & "','" & ComIDNO.Value & "');"
                'SeqNo.Attributes("onclick") = "SelectSeqNo(this.checked,this.value,'" & PlanID.Value & "','" & ComIDNO.Value & "');"
                Dim SeqNoArray As Array = Split(SeqNoValue.Value, ",")
                For i As Integer = 0 To SeqNoArray.Length - 1
                    If drv("SeqNo").ToString = SeqNoArray(i) Then SeqNo.Checked = True
                Next
        End Select
    End Sub

    'Private Function GetTechID(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNO As String, ByVal SubSeqNO As String, ByVal CDateVal As String) As String
    '    Dim rst As String = ""
    '    If PlanID = "" OrElse ComIDNO = "" OrElse SeqNO = "" OrElse SubSeqNO = "" OrElse CDateVal = "" Then Return rst
    '    Using dt As New DataTable
    '        Dim sql As String = ""
    '        sql &= " SELECT NewData11_1 FROM dbo.PLAN_REVISE WITH(NOLOCK)" & vbCrLf
    '        sql &= " WHERE altDataID=11 AND PlanID='" & PlanID & "' AND ComIDNO='" & ComIDNO & "' AND SeqNO='" & SeqNO & "'" & vbCrLf
    '        sql &= " AND SubSeqNO='" & SubSeqNO & "' AND CDate=" & TIMS.To_date(CDateVal)
    '        Dim oCmd As New SqlCommand(sql, objconn)
    '        With oCmd
    '            .Parameters.Clear()
    '            dt.Load(.ExecuteReader())
    '        End With
    '        If TIMS.dtHaveDATA(dt) Then rst = Convert.ToString(dt.Rows(0)("NewData11_1"))
    '    End Using
    '    Return rst
    'End Function

    'Private Function GetPTDRID(ByVal PlanID As String, ByVal ComIDNO As String, ByVal SeqNO As String, ByVal SubSeqNO As String, ByVal CDateVal As String) As String
    '    Dim rst As String = ""
    '    If PlanID = "" OrElse ComIDNO = "" OrElse SeqNO = "" OrElse SubSeqNO = "" OrElse CDateVal = "" Then Return rst
    '    Using dt As New DataTable
    '        Dim sql As String = ""
    '        sql &= " SELECT MAX(PTDRID) PTDRID FROM dbo.PLAN_TRAINDESC_REVISE WITH(NOLOCK)" & vbCrLf
    '        sql &= " WHERE PlanID='" & PlanID & "' AND ComIDNO='" & ComIDNO & "' AND SeqNO='" & SeqNO & "' AND SubSeqNO='" & SubSeqNO & "' AND CDate=" & TIMS.To_date(CDateVal)
    '        Dim oCmd As New SqlCommand(sql, objconn)
    '        With oCmd
    '            .Parameters.Clear()
    '            dt.Load(.ExecuteReader())
    '        End With
    '        If TIMS.dtHaveDATA(dt) Then rst = Convert.ToString(dt.Rows(0)("PTDRID"))
    '    End Using
    '    Return rst
    'End Function

    Sub ClearValue1()
        SelectValue.Value = ""
        SeqNoValue.Value = ""
        PLANIDValue.Value = ""
        ComIDNOValue.Value = ""
        SubSeqNoValue.Value = ""
        CDateItem.Value = ""
        selsqlstr.Value = ""
        TechIDvalue.Value = ""
    End Sub

    Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        'Button1_Click(sender, e)
        Call ClearValue1()
        DataGridTable.Visible = False
    End Sub

    Private Sub PlanPoint_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles PlanPoint.SelectedIndexChanged
        'Button1_Click(sender, e)
        DataGridTable.Visible = False
    End Sub

    '列印
    Protected Sub btnPrint1_Click(sender As Object, e As EventArgs) Handles btnPrint1.Click
        Dim myvalue As String = ""

        ROC_Years.Value = TIMS.ClearSQM(ROC_Years.Value)
        SelectValue.Value = TIMS.ClearSQM(SelectValue.Value)
        TechIDvalue.Value = TIMS.ClearSQM(TechIDvalue.Value)
        TIMS.SetMyValue(myvalue, "Years", ROC_Years.Value) 'sm.UserInfo.Years - 1911 'TIMS.SetMyValue(myvalue, "Years2", Years2.Value) 'sm.UserInfo.Years
        TIMS.SetMyValue(myvalue, "OCID", SelectValue.Value)
        TIMS.SetMyValue(myvalue, "TechID", TechIDvalue.Value)
        TIMS.SetMyValue(myvalue, "selsqlstr", selsqlstr.Value)
        TIMS.SetMyValue(myvalue, "PLANID", PLANIDValue.Value)
        TIMS.SetMyValue(myvalue, "ComIDNO", ComIDNOValue.Value)
        TIMS.SetMyValue(myvalue, "SEQNO", SeqNoValue.Value)
        TIMS.SetMyValue(myvalue, "Title", KindValue.Value)
        '<asp@ListItem Value="0">未轉班</asp@ListItem>
        '<asp@ListItem Value="1">已轉班</asp@ListItem>
        '<asp@ListItem Value="2">變更待審</asp@ListItem>
        '//vRadio1 0:未轉班(PPIPK) x1:已轉班 2:變更待審(PPIPK)'selsqlstr
        Dim fileN As String = ""
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1)
        Select Case v_Radio1'Radio1.SelectedValue
            Case "0" '未轉班
                If selsqlstr.Value = "" Then
                    Common.MessageBox(Me, "未勾選列印資料!!")
                    Exit Sub
                End If
                fileN = cst_reportFN1
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN1, myvalue)
            Case "1" '已轉班
                If SelectValue.Value = "" Then
                    Common.MessageBox(Me, "未勾選列印資料!!")
                    Exit Sub
                End If
                fileN = cst_reportFN0
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN0, myvalue)
            Case "2" '變更待審
                Dim selval As String = ""
                Dim techidval As String = ""
                Dim PTDRIDValue As String = ""
                For Each eItem As DataGridItem In DataGrid2.Items
                    'Dim drv As DataRowView = e.Item.DataItem
                    Dim SeqNo As HtmlInputCheckBox = eItem.FindControl("SeqNo")
                    Dim PlanID As HtmlInputHidden = eItem.FindControl("PlanID")
                    Dim ComIDNO As HtmlInputHidden = eItem.FindControl("ComIDNO")
                    Dim SubSeqNo As HtmlInputHidden = eItem.FindControl("SubSeqNo")
                    Dim CDateValue As HtmlInputHidden = eItem.FindControl("CDateValue")
                    Dim TechID As HtmlInputHidden = eItem.FindControl("TechID")
                    Dim hPTDRID As HtmlInputHidden = eItem.FindControl("hPTDRID")
                    PlanID.Value = TIMS.ClearSQM(PlanID.Value)
                    ComIDNO.Value = TIMS.ClearSQM(ComIDNO.Value)
                    SeqNo.Value = TIMS.ClearSQM(SeqNo.Value)
                    SubSeqNo.Value = TIMS.ClearSQM(SubSeqNo.Value)
                    CDateValue.Value = TIMS.ClearSQM(CDateValue.Value)
                    TechID.Value = TIMS.ClearSQM(TechID.Value)
                    hPTDRID.Value = TIMS.ClearSQM(hPTDRID.Value)
                    If (SeqNo.Checked) AndAlso PlanID.Value <> "" AndAlso ComIDNO.Value <> "" AndAlso SeqNo.Value <> "" AndAlso SubSeqNo.Value <> "" AndAlso CDateValue.Value <> "" Then
                        Dim selval2 As String = String.Format("\'{0}-{1}-{2}-{3}-{4}\'", PlanID.Value, ComIDNO.Value, SeqNo.Value, SubSeqNo.Value, CDateValue.Value)
                        selval &= String.Concat(If(selval <> "", ",", ""), selval2)
                        If TechID.Value <> "" Then techidval &= String.Concat(If(techidval <> "", ",", ""), TechID.Value)
                        If hPTDRID.Value <> "" Then PTDRIDValue &= String.Concat(If(PTDRIDValue <> "", ",", ""), hPTDRID.Value)
                    End If
                Next
                selsqlstr.Value = selval
                TechIDvalue.Value = techidval
                If selsqlstr.Value = "" Then
                    Common.MessageBox(Me, "勾選列印資料有誤!")
                    Exit Sub
                End If
                If TechIDvalue.Value = "" Then
                    Common.MessageBox(Me, "勾選列印資料有誤!!")
                    Exit Sub
                End If
                If PTDRIDValue = "" Then
                    Common.MessageBox(Me, "勾選列印資料有誤.")
                    Exit Sub
                End If

                myvalue = ""
                TIMS.SetMyValue(myvalue, "Years", ROC_Years.Value) 'sm.UserInfo.Years - 1911
                'TIMS.SetMyValue(myvalue, "Years2", Years2.Value) 'sm.UserInfo.Years
                TIMS.SetMyValue(myvalue, "OCID", SelectValue.Value)
                TIMS.SetMyValue(myvalue, "TechID", TechIDvalue.Value)
                TIMS.SetMyValue(myvalue, "selsqlstr", selsqlstr.Value)
                TIMS.SetMyValue(myvalue, "PTDRID", PTDRIDValue)
                TIMS.SetMyValue(myvalue, "PLANID", PLANIDValue.Value)
                'TIMS.SetMyValue(myvalue, "ComIDNO", ComIDNOValue.Value)
                'TIMS.SetMyValue(myvalue, "SEQNO", SeqNoValue.Value)
                TIMS.SetMyValue(myvalue, "Title", KindValue.Value)

                fileN = cst_reportFN2
                'TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_reportFN2, myvalue)
            Case Else
                Common.MessageBox(Me, "未選擇，班級狀態!")
                Exit Sub
        End Select
        If fileN = "" Then
            Common.MessageBox(Me, "未選擇，班級狀態!")
            Exit Sub
        End If
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, fileN, myvalue)
    End Sub
End Class