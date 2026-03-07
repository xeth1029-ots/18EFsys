Partial Class SD_14_006
    Inherits AuthBasePage

    'SD_14_006_1 (未轉班、已轉班)
    'SD_14_006_2 (變更待審)
    'SD_14_006_*.jrxml
    Const cst_printFN1 As String = "SD_14_006_1" 'SD_14_006_1.jrxml (未轉班、已轉班)
    Const cst_printFN2 As String = "SD_14_006_2" 'SD_14_006_2.jrxml (變更待審)

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Me.Radio1.SelectedIndex = 0
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        ROC_Years.Value = sm.UserInfo.Years - 1911
        Years2.Value = sm.UserInfo.Years

        Button1.Attributes("onclick") = "return CheckSearch();"
        'print.Attributes("onclick")="CheckPrint('" & ReportQuery.GetSmartQueryPath & "');"
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Call SearchPlan1()
    End Sub

    Private Sub SearchPlan1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim RelShip As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim SearchStr As String = ""
        If sm.UserInfo.LID <> 0 Then SearchStr &= $" AND ip.PlanID={sm.UserInfo.PlanID}"
        SearchStr &= " AND a.RID IN (SELECT RID FROM Auth_Relship WHERE RelShip LIKE '" & RelShip & "%') "
        If STDate1.Text <> "" Then
            SearchStr &= " AND a.STDate >= " & TIMS.To_date(STDate1.Text) '" & STDate1.Text & "','yyyy/MM/dd')"
        End If
        If STDate2.Text <> "" Then
            SearchStr &= " AND a.STDate <= " & TIMS.To_date(STDate2.Text) '" & STDate2.Text & "','yyyy/MM/dd')"
        End If
        If FTDate1.Text <> "" Then
            SearchStr &= " AND a.FDDate >= " & TIMS.To_date(FTDate1.Text) '" & FTDate1.Text & "','yyyy/MM/dd')"
        End If
        If FTDate2.Text <> "" Then
            SearchStr &= " AND a.FDDate <= " & TIMS.To_date(FTDate2.Text) '" & FTDate2.Text & "','yyyy/MM/dd')"
        End If
        'If Radio1.SelectedIndex=0 Then SearchStr &= " and TransFlag ='N' " Else SearchStr &= " and TransFlag ='Y' "
        Select Case Radio1.SelectedIndex
            Case 0     '未轉班
                SearchStr &= " AND a.TransFlag='N'"
            Case 1, 2    '已轉班 or 待審核 
                SearchStr &= " AND a.TransFlag='Y'"
        End Select

        Dim sql As String = "" & vbCrLf
        sql &= " SELECT a.PlanID ,a.ComIDNO ,a.SeqNo" & vbCrLf
        sql &= " ,CONCAT(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),dbo.FN_GET_RESULTBUTTON_YR(a.RESULTBUTTON)) ClassCName" & vbCrLf
        sql &= " ,a.STDate ,a.FDDate FTDate ,b.OrgName" & vbCrLf
        Select Case Radio1.SelectedValue
            Case "2"
                '090417 andy  edit '待審核 
                sql &= " ,CONVERT(VARCHAR, c.ModifyDate, 120) ModifyDate ,c.Subseqno ,CONVERT(VARCHAR, c.CDate, 111) CDate" & vbCrLf
                sql &= " FROM PLAN_PLANINFO a" & vbCrLf
                sql &= " JOIN VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
                sql &= " JOIN ID_PLAN ip ON ip.planid=a.planid" & vbCrLf
                sql &= " LEFT JOIN PLAN_REVISE c ON c.PlanID=a.PlanID AND c.ComIDNO=a.ComIDNO AND c.SeqNO=a.SeqNO" & vbCrLf
            Case Else
                sql &= " FROM PLAN_PLANINFO a" & vbCrLf
                sql &= " JOIN VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
                sql &= " JOIN id_Plan ip ON ip.planid=a.planid" & vbCrLf
        End Select

        sql &= " WHERE a.IsApprPaper='Y'"
        sql &= $" AND ip.TPlanID='{sm.UserInfo.TPlanID}' AND ip.Years='{sm.UserInfo.Years}'"
        sql &= SearchStr
        Select Case Radio1.SelectedValue
            Case "2"
                '090417 andy  edit  '待審核 
                sql &= " AND c.ReviseStatus IS NULL AND c.AltDataID=14 "
                sql &= " ORDER BY a.STDate, a.ClassName, OrgName, c.ModifyDate "
            Case Else
                sql &= " ORDER BY a.STDate, a.ClassName, OrgName "
        End Select
        Select Case Radio1.SelectedValue
            Case "2"
                DataGrid1.Columns(5).Visible = True
            Case Else
                DataGrid1.Columns(5).Visible = False
        End Select

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        'DataGrid1.Visible=True
        'PageControler1.Visible=True
        'PageControler1.SqlString=sql
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        DataGridTable.Visible = False
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'e.Item.Cells(6).Style.Add("display", "none")
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "head_navy"
                'e.Item.Cells(8).Style.Add("display", "none")  'SubSeqNO
                'e.Item.Cells(9).Style.Add("display", "none")  'CDateValue
                'If Radio1.SelectedValue <> 2 Then e.Item.Cells(5).Visible=False
            Case ListItemType.Item, ListItemType.AlternatingItem
                'e.Item.Cells(8).Style.Add("display", "none")
                'e.Item.Cells(9).Style.Add("display", "none")
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                Dim chkSeqNo As HtmlInputCheckBox = e.Item.FindControl("chkSeqNo") 'chkSeqNo
                Dim HidRadio1 As HtmlInputHidden = e.Item.FindControl("HidRadio1")
                Dim hidPlanID As HtmlInputHidden = e.Item.FindControl("hidPlanID")
                Dim hidComIDNO As HtmlInputHidden = e.Item.FindControl("hidComIDNO")
                Dim hidSeqNo As HtmlInputHidden = e.Item.FindControl("hidSeqNo")
                Dim hidSubSeqNo As HtmlInputHidden = e.Item.FindControl("hidSubSeqNo")
                Dim hidCDateValue As HtmlInputHidden = e.Item.FindControl("hidCDateValue")
                Dim labModifydate As Label = e.Item.FindControl("labModifydate")
                HidRadio1.Value = Radio1.SelectedValue
                hidPlanID.Value = drv("PlanID")
                hidComIDNO.Value = drv("ComIDNO")
                hidSeqNo.Value = drv("SeqNo")
                Select Case Radio1.SelectedValue
                    Case "2"
                        hidSubSeqNo.Value = Convert.ToString(drv("SubSeqNo"))
                        hidCDateValue.Value = TIMS.Cdate3(drv("CDate"), "yyyyMMdd") '.ToString("yyyyMMdd")
                        labModifydate.Text = Convert.ToString(drv("ModifyDate"))
                    Case Else
                        hidSubSeqNo.Value = "" 'drv("ComIDNO")
                        hidCDateValue.Value = "" 'drv("SeqNo")
                        labModifydate.Text = ""
                        'DataGrid1.Items(5).Visible=False
                End Select
        End Select
    End Sub

    Protected Sub BtnPrint_Click(sender As Object, e As EventArgs) Handles BtnPrint.Click
        Dim ii As Integer = 0
        Dim PCSValue As String = ""
        Dim PlanIDValue As String = Convert.ToString(sm.UserInfo.PlanID)
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim chkSeqNo As HtmlInputCheckBox = eItem.FindControl("chkSeqNo")
            Dim HidRadio1 As HtmlInputHidden = eItem.FindControl("HidRadio1")
            Dim hidPlanID As HtmlInputHidden = eItem.FindControl("hidPlanID")
            Dim hidComIDNO As HtmlInputHidden = eItem.FindControl("hidComIDNO")
            Dim hidSeqNo As HtmlInputHidden = eItem.FindControl("hidSeqNo")
            Dim hidSubSeqNo As HtmlInputHidden = eItem.FindControl("hidSubSeqNo")
            Dim hidCDateValue As HtmlInputHidden = eItem.FindControl("hidCDateValue")
            Dim labModifydate As Label = eItem.FindControl("labModifydate")
            Dim xTmp1 As String = ""
            If HidRadio1.Value = Radio1.SelectedValue AndAlso chkSeqNo.Checked Then
                ii += 1
                Select Case Radio1.SelectedValue
                    Case "2" '待審核 (變更待審)
                        xTmp1 = $"{hidPlanID.Value}x{hidComIDNO.Value}x{hidSeqNo.Value}x{hidSubSeqNo.Value}x{hidCDateValue.Value}"
                        PCSValue &= $"{If(PCSValue <> "", ",", "")}\'{xTmp1}\'"
                    Case Else
                        xTmp1 = $"{hidPlanID.Value}x{hidComIDNO.Value}x{hidSeqNo.Value}"
                        PCSValue &= $"{If(PCSValue <> "", ",", "")}\'{xTmp1}\'"
                End Select
                If PlanIDValue <> hidPlanID.Value Then PlanIDValue = hidPlanID.Value
            End If
        Next
        If ii = 0 OrElse PCSValue = "" Then
            Common.MessageBox(Me, "請選擇要列印的班級!!")
            Exit Sub
        End If
        '//變更待審
        'var sUrl='../../SQControl.aspx?SQ_AutoLogout=true&sys=BussinessTrain';
        'var filename='&filename=SD_14_006_1&statusTyp=1&path=' + SMpath;
        'var value1='&PlanID=' + PlanIDValue.value + '&ComIDNO=' + ComIDNOValue.value + '&SEQNO=' + SeqNoValue.value + '&Years=' + Years.value;
        'if (document.form1.Radio1_2.checked) {
        'filename='&filename=SD_14_006_2&statusTyp=2&path=' + SMpath;
        'value1='&PlanID=' + PlanIDValue.value + '&ComIDNO=' + ComIDNOValue.value + '&SEQNO=' + SeqNoValue.value + '&Years=' + Years2.value;
        '}
        'openPrint(sUrl + filename + value1);
        '}
        'var value1='&PlanID=' + PlanIDValue.value + '&ComIDNO=' + ComIDNOValue.value + '&SEQNO=' + SeqNoValue.value + '&Years=' + Years.value;
        Dim sFilename1 As String = ""
        Dim sValue1 As String = ""
        Select Case Radio1.SelectedValue
            Case "2" '待審核 (變更待審)
                sFilename1 = cst_printFN2
                sValue1 = ""
                sValue1 &= "&Years=" & Years2.Value
                sValue1 &= "&PCSValue=" & PCSValue
                Select Case CStr(sm.UserInfo.LID)
                    Case "2" '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                        sValue1 &= "&PlanID=" & sm.UserInfo.PlanID
                    Case Else
                        sValue1 &= "&PlanID=" & PlanIDValue
                End Select
                'sValue1 &= "&RID=" & sm.UserInfo.RID
            Case Else
                sFilename1 = cst_printFN1
                sValue1 = ""
                sValue1 &= "&Years=" & ROC_Years.Value
                sValue1 &= "&PCSValue=" & PCSValue
                Select Case CStr(sm.UserInfo.LID)
                    Case "2" '階層代碼【0:署(局) 1:分署(中心) 2:委訓】
                        sValue1 &= "&PlanID=" & sm.UserInfo.PlanID
                    Case Else
                        sValue1 &= "&PlanID=" & PlanIDValue
                End Select
                'sValue1 &= "&RID=" & sm.UserInfo.RID
        End Select
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sFilename1, sValue1)
    End Sub
End Class