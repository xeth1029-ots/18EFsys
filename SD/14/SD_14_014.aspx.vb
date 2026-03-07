Partial Class SD_14_014
    Inherits AuthBasePage

    'SD_14_014_1,'SD_14_014,'SD_14_014*.jrxml,
    '教學環境資料表 '0:未轉班'1:已轉班'2:變更待審
    Const cst_printFN1 As String = "SD_14_014" '0:未轉班' 1:已轉班 '教學環境資料表
    Const cst_printFN2 As String = "SD_14_014_1" '2:變更待審
    'OJT-24031203：<系統> 產投 -教學環境資料表：新增列印【混成課程_教學環境資料表】功能
    Const cst_printFN1R As String = "SD_14_014R" '教學環境資料表(遠距課程)

    Dim V_Radio1 As String = ""
    Const cst_eCommandName_PRINT2 As String = "PRINT2" '教學環境資料表
    Const cst_eCommandName_PRINT2R As String = "PRINT2R" '教學環境資料表(遠距課程)
    '申請變更時間
    Const Cst_申請變更時間 As Integer = 4 '5
    'Const Cst_變數功能 As Integer=6

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        PageControler1.PageDataGrid = DataGrid1

        'CYears.Value=sm.UserInfo.Years - 1911 '民國年
        'Years9.Value=sm.UserInfo.Years '西元年
        'BtnPrint1.DisabledCssClass

        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Me.Radio1.SelectedIndex = 0

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
            Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

            TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
            If HistoryRID.Rows.Count <> 0 Then
                center.Attributes("onclick") = "showObj('HistoryList2');"
                center.Style("CURSOR") = "hand"
            End If

            Button1.Attributes("onclick") = "return CheckSearch();"
            'BtnPrint1.Attributes("onclick")="CheckPrint();"
        End If
    End Sub

    Private Sub DataGrid1_ItemCommand(source As Object, e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        If e.CommandArgument = "" OrElse e.CommandName = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim vRadio1 As String = TIMS.GetMyValue(sCmdArg, "Radio1")
        Dim vPCS2 As String = TIMS.GetMyValue(sCmdArg, "PCS2")
        Dim vPCS As String = TIMS.GetMyValue(sCmdArg, "PCS")
        If vRadio1 = "" Then Exit Sub
        If vPCS2 = "" AndAlso vPCS = "" Then Exit Sub '2個至少應該有1值
        Dim eCommandName As String = TIMS.ClearSQM(e.CommandName)
        Call Utl_PrintX(sCmdArg, eCommandName)
    End Sub

    Private Sub Datagrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        'V_Radio1=TIMS.GetListValue(Radio1)
        '0:未轉班'1:已轉班'2:變更待審 'SelectValue1.Value=Radio1.SelectedValue
        V_Radio1 = Radio1.SelectedValue
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.Cells(Cst_變數功能).Style.Add("display", "none")
                e.Item.CssClass = "head_navy"
                If V_Radio1 <> "2" Then e.Item.Cells(Cst_申請變更時間).Visible = False
            Case ListItemType.Item, ListItemType.AlternatingItem
                'e.Item.Cells(Cst_變數功能).Style.Add("display", "none")
                If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                Dim drv As DataRowView = e.Item.DataItem
                'Dim chkSeqNo As HtmlInputCheckBox=e.Item.FindControl("chkSeqNo")
                Dim hid_PCS As HtmlInputHidden = e.Item.FindControl("hid_PCS")
                Dim hid_PCS2 As HtmlInputHidden = e.Item.FindControl("hid_PCS2")
                Dim hid_Radio1 As HtmlInputHidden = e.Item.FindControl("hid_Radio1")
                Dim LabModifyDate As Label = e.Item.FindControl("LabModifyDate")
                Dim BtnPrint2 As Button = e.Item.FindControl("BtnPrint2")
                Dim BtnPrint2R As Button = e.Item.FindControl("BtnPrint2R")

                BtnPrint2R.Visible = If(Convert.ToString(drv("DISTANCE_Y")) = "Y", True, False)

                e.Item.Cells(Cst_申請變更時間).Visible = False
                If V_Radio1 = "2" Then
                    e.Item.Cells(Cst_申請變更時間).Visible = True
                    LabModifyDate.Text = Convert.ToString(drv("ModifyDate"))
                End If
                'chkSeqNo.Value=drv("SeqNo")
                hid_Radio1.Value = Convert.ToString(drv("Radio1"))
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "Radio1", hid_Radio1.Value)

                Select Case Convert.ToString(drv("Radio1"))
                    Case "2"
                        '0:未轉班'1:已轉班'2:變更待審
                        TIMS.SetMyValue(sCmdArg, "PCS2", Convert.ToString(drv("PCS2")))
                        hid_PCS2.Value = Convert.ToString(drv("PCS2"))
                    Case Else
                        TIMS.SetMyValue(sCmdArg, "PCS", Convert.ToString(drv("PCS")))
                        hid_PCS.Value = Convert.ToString(drv("PCS"))
                End Select

                BtnPrint2.CommandArgument = sCmdArg
                BtnPrint2R.CommandArgument = sCmdArg
                BtnPrint2.CommandName = cst_eCommandName_PRINT2
                BtnPrint2R.CommandName = cst_eCommandName_PRINT2R

        End Select
    End Sub

    Private Sub Radio1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Radio1.SelectedIndexChanged
        DataGridTable.Visible = False
    End Sub

    ''' <summary>'0:未轉班'1:已轉班 </summary>
    Sub gSearch01()
        '0:未轉班'1:已轉班'2:變更待審 'SelectValue1.Value=Radio1.SelectedValue
        V_Radio1 = TIMS.GetListValue(Radio1) 'V_Radio1=Radio1.SelectedValue
        Select Case V_Radio1 'Radio1.SelectedValue
            Case "0", "1"
            Case "2"
        End Select

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        '變數防止 SQL INJECTION 攻擊
        ViewState("RIDValue") = TIMS.ClearSQM(RIDValue.Value) '.Trim.Replace("'", "''")
        ViewState("STDate1") = TIMS.ClearSQM(STDate1.Text) '.Trim.Replace("'", "''")
        ViewState("STDate2") = TIMS.ClearSQM(STDate2.Text) '.Trim.Replace("'", "''")
        ViewState("FTDate1") = TIMS.ClearSQM(FTDate1.Text) '.Trim.Replace("'", "''")
        ViewState("FTDate2") = TIMS.ClearSQM(FTDate2.Text) '.Trim.Replace("'", "''")

        Dim RelShip As String = TIMS.GET_RelshipforRID($"{ViewState("RIDValue")}", objconn)

        Dim sql As String = ""
        sql &= " SELECT a.PlanID,a.ComIDNO,a.SeqNo" & vbCrLf
        '0:未轉班'1:已轉班'2:變更待審
        Select Case V_Radio1'Radio1.SelectedValue
            Case "0"
                sql &= " ,'0' Radio1 " & vbCrLf
            Case "1"
                sql &= " ,'1' Radio1 " & vbCrLf
        End Select
        sql &= " ,CONCAT(a.PLANID,'-',a.COMIDNO,'-',a.SEQNO) PCS" & vbCrLf
        sql &= " ,CONCAT(dbo.FN_GET_CLASSCNAME(a.ClassName,a.CyclType),dbo.FN_GET_RESULTBUTTON_YR(a.RESULTBUTTON)) CLASSCNAME" & vbCrLf
        sql &= " ,CONVERT(varchar, a.STDate, 111) STDATE" & vbCrLf
        sql &= " ,CONVERT(varchar, a.FDDate, 111) FTDATE" & vbCrLf
        'sql &= " ,b.ORGNAME,a.RMTID,CASE WHEN a.RMTID IS NOT NULL THEN 'Y' END RMTID_Y" & vbCrLf
        sql &= " ,b.ORGNAME,a.DISTANCE,CASE WHEN a.DISTANCE='2' AND a.RMTID IS NOT NULL THEN 'Y' END DISTANCE_Y" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a" & vbCrLf
        sql &= " JOIN VIEW_RIDNAME b ON a.RID=b.RID" & vbCrLf
        sql &= " JOIN ID_PLAN ip ON ip.PLANID=b.PLANID" & vbCrLf
        sql &= " WHERE a.IsApprPaper='Y' AND b.RelShip LIKE '" & RelShip & "%' " & vbCrLf
        Select Case Radio1.SelectedIndex
            Case 0  '未轉班
                sql &= " AND a.TransFlag='N' " & vbCrLf
            Case 1, 2 '已轉班、變更待審 
                sql &= " AND a.TransFlag='Y' " & vbCrLf
        End Select
        If STDate1.Text <> "" Then
            sql &= " AND a.STDate >= " & TIMS.To_date(ViewState("STDate1")) & vbCrLf '
        End If
        If STDate2.Text <> "" Then
            sql &= " AND a.STDate <= " & TIMS.To_date(ViewState("STDate2")) & vbCrLf 'convert(datetime, '" & ViewState("STDate2") & "', 111)"
        End If
        If FTDate1.Text <> "" Then
            sql &= " AND a.FDDate >= " & TIMS.To_date(ViewState("FTDate1")) & vbCrLf 'convert(datetime, '" & ViewState("FTDate1") & "', 111)"
        End If
        If FTDate2.Text <> "" Then
            sql &= " AND a.FDDate <= " & TIMS.To_date(ViewState("FTDate2")) & vbCrLf 'convert(datetime, '" & ViewState("FTDate2") & "', 111)"
        End If
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>
    ''' '2:變更待審
    ''' </summary>
    Sub gSearch2()
        '0:未轉班'1:已轉班 '2:變更待審 'Radio1.SelectedValue '2:變更待審
        V_Radio1 = TIMS.GetListValue(Radio1) 'V_Radio1=Radio1.SelectedValue
        Select Case V_Radio1'Radio1.SelectedValue
            Case "0", "1"
            Case "2"
        End Select

        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        '變數防止 SQL INJECTION 攻擊
        ViewState("RIDValue") = TIMS.ClearSQM(RIDValue.Value) '.Trim.Replace("'", "''")
        ViewState("STDate1") = TIMS.ClearSQM(STDate1.Text) '.Trim.Replace("'", "''")
        ViewState("STDate2") = TIMS.ClearSQM(STDate2.Text) '.Trim.Replace("'", "''")
        ViewState("FTDate1") = TIMS.ClearSQM(FTDate1.Text) '.Trim.Replace("'", "''")
        ViewState("FTDate2") = TIMS.ClearSQM(FTDate2.Text) '.Trim.Replace("'", "''")

        Dim RelShip As String = TIMS.GET_RelshipforRID($"{ViewState("RIDValue")}", objconn)

        Dim sql As String = ""
        sql &= " SELECT a.PlanID ,a.ComIDNO ,a.SeqNo ,'2' Radio1" & vbCrLf
        sql &= " ,CONCAT(a.PLANID,'-',a.COMIDNO,'-',a.SEQNO) PCS" & vbCrLf
        sql &= " ,CONCAT(a.PLANID,'-',a.COMIDNO,'-',a.SEQNO,'-',c.SUBSEQNO,'-',FORMAT(c.CDATE,'yyyy-MM-dd')) PCS2" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.STDate, 111) STDATE" & vbCrLf
        sql &= " ,CONVERT(VARCHAR, a.FDDate, 111) FTDATE" & vbCrLf
        'sql &= " ,b.ORGNAME,a.RMTID,CASE WHEN a.RMTID IS NOT NULL THEN 'Y' END RMTID_Y" & vbCrLf
        sql &= " ,b.ORGNAME,a.DISTANCE,CASE WHEN a.DISTANCE='2' AND a.RMTID IS NOT NULL THEN 'Y' END DISTANCE_Y" & vbCrLf
        'Radio1.SelectedValue'2:變更待審
        sql &= " ,CONVERT(VARCHAR, c.modifydate, 120) ModifyDate ,c.Subseqno ,CONVERT(VARCHAR, c.CDate, 111) CDATE" & vbCrLf
        sql &= " FROM PLAN_PLANINFO a " & vbCrLf
        sql &= " JOIN VIEW_RIDNAME b ON a.RID=b.RID " & vbCrLf
        sql &= " JOIN ID_Plan ip ON ip.PlanID=b.PlanID " & vbCrLf
        'Radio1.SelectedValue'2:變更待審
        sql &= " LEFT JOIN PLAN_REVISE c ON c.PlanID=a.PlanID AND c.ComIDNO=a.ComIDNO AND c.SeqNO=a.SeqNO " & vbCrLf
        sql &= " WHERE c.ReviseStatus IS NULL  AND c.AltDataID=14 " & vbCrLf
        sql &= " AND b.RelShip LIKE '" & RelShip & "%' " & vbCrLf
        'Radio1.SelectedValue'2:變更待審
        sql &= " AND a.TransFlag='Y' AND a.IsApprPaper='Y' " & vbCrLf
        If STDate1.Text <> "" Then sql &= " AND a.STDate >= " & TIMS.To_date(ViewState("STDate1")) & vbCrLf
        If STDate2.Text <> "" Then sql &= " AND a.STDate <= " & TIMS.To_date(ViewState("STDate2")) & vbCrLf
        If FTDate1.Text <> "" Then sql &= " AND a.FDDate >= " & TIMS.To_date(ViewState("FTDate1")) & vbCrLf
        If FTDate2.Text <> "" Then sql &= " AND a.FDDate <= " & TIMS.To_date(ViewState("FTDate2")) & vbCrLf
        If sm.UserInfo.RID = "A" Then
            sql &= " AND ip.TPlanID='" & sm.UserInfo.TPlanID & "' " & vbCrLf
            sql &= " AND ip.Years='" & sm.UserInfo.Years & "' " & vbCrLf
        Else
            sql &= " AND ip.PlanID='" & sm.UserInfo.PlanID & "' " & vbCrLf
        End If
        sql &= " ORDER BY a.STDate, a.ClassName, b.OrgName, c.ModifyDate " & vbCrLf

        DataGridTable.Visible = False
        msg.Text = "查無資料"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
        If TIMS.dtNODATA(dt) Then Return

        DataGridTable.Visible = True
        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Hid_PCSALL.Value = ""
        Hid_PCS2ALL.Value = ""

        '0:未轉班'1:已轉班'2:變更待審
        V_Radio1 = TIMS.GetListValue(Radio1) 'V_Radio1=Radio1.SelectedValue
        SelectValue1.Value = V_Radio1 'Radio1.SelectedValue
        '0:未轉班'1:已轉班'2:變更待審
        Select Case Radio1.SelectedValue
            Case "0", "1"
                Call gSearch01() '0:未轉班'1:已轉班
            Case "2"
                Call gSearch2() '2:變更待審
            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                Exit Sub
        End Select

    End Sub

    '單筆列印
    Sub Utl_PrintX(ByVal sCmdArg As String, ByVal eCommandName As String)
        Dim sPrint_Test As String = TIMS.Utl_GetConfigSet("printtest")
        Dim TSTPRINT As String = If(sPrint_Test = "Y", "2", "1") '正式區1／'測試區2

        'Hid_PCSALL.Value=""'Hid_PCS2ALL.Value=""
        '0:未轉班'1:已轉班'2:變更待審 'SelectValue1.Value=Radio1.SelectedValue
        Dim vRadio1 As String = TIMS.GetMyValue(sCmdArg, "Radio1")
        Dim vPCS2 As String = TIMS.GetMyValue(sCmdArg, "PCS2")
        Dim vPCS As String = TIMS.GetMyValue(sCmdArg, "PCS")
        If vRadio1 = "" Then Exit Sub
        If vPCS2 = "" AndAlso vPCS = "" Then Exit Sub '2個至少應該有1值
        Dim sfilename1 As String = "" 'cst_printFN1
        Dim sMyValue As String = ""

        ROC_Years.Value = (sm.UserInfo.Years - 1911)
        Select Case SelectValue1.Value
            Case "0", "1" '0:未轉班'1:已轉班
                Select Case eCommandName
                    Case cst_eCommandName_PRINT2
                        sfilename1 = cst_printFN1

                    Case cst_eCommandName_PRINT2R
                        sfilename1 = cst_printFN1R

                End Select
                sMyValue &= "&Years=" & ROC_Years.Value
                sMyValue &= "&selsqlstr=" & vPCS

            Case "2" '2:變更待審
                sfilename1 = cst_printFN2
                sMyValue &= "&Years=" & sm.UserInfo.Years
                sMyValue &= "&selsqlstr=" & vPCS2

            Case Else
                Common.MessageBox(Me, TIMS.cst_NODATAMsg2)
                Exit Sub
        End Select

        sMyValue &= "&TPlanID=" & sm.UserInfo.TPlanID
        sMyValue &= "&SYears=" & sm.UserInfo.Years
        sMyValue &= "&TSTPRINT=" & TSTPRINT '正式區1 '測試區2
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, sfilename1, sMyValue)
    End Sub

End Class
