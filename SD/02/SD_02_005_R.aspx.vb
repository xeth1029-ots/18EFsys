Partial Class SD_02_005_R
    Inherits AuthBasePage

    'aspx 列印
    'Maintest_GradeList
    'Maintest_GradeList2
    'http://vm-tims:8080/ReportServer2/report.do?RptID=Maintest_GradeList&OCID1=95948&UserID=oudou
    Const cst_printFN1 As String = "Maintest_GradeList2" '"Maintest_GradeList"
    Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        '分頁設定 Start
        PageControler1.PageDataGrid = DataGrid1
        '分頁設定 End

        'Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
        blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        If blnP0 Then Trwork2013a.Visible = True
#Region "(No Use)"

        ''就服單位協助報名
        'Trwork2013a.Visible = False
        'If sm.UserInfo.Years >= 2013 AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then Trwork2013a.Visible = True
        'End If

#End Region

        If Not IsPostBack Then
            PageControler1.Visible = False
            Button1.Visible = False
            msg.Visible = False
            'CTPanel.Visible = True
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            query.Attributes("onclick") = "return CheckData();"
            Me.radiobtn1.Attributes("onclick") = "showPanel();"
            'Button1.Attributes("onclick") = "CheckPrint();return false;"
            Button1.Attributes("onclick") = "return CheckPrint();"
            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button3_Click(sender, e)
            End If
        End If

#Region "(No Use)"

        'sql = "SELECT a.*,b.OrgName,c.RID as RIDValue,d.TrainID,d.TrainName FROM "
        'sql += "(SELECT * FROM Plan_PlanInfo WHERE PlanID='" & Request("PlanID") & "' and ComIDNO='" & Request("ComIDNO") & "' and SeqNO='" & Request("SeqNO") & "') a "
        'sql += "JOIN Org_OrgInfo b ON a.ComIDNO=b.ComIDNO "
        'sql += "JOIN Auth_Relship c ON c.OrgID=b.OrgID "
        'sql += "LEFT JOIN Key_TrainType d ON a.TMID=d.TMID"
        'dr = DbAccess.GetOneRow(sql)

#End Region

        If Request("PlanID") <> "" AndAlso Request("ComIDNO") <> "" AndAlso Request("SeqNO") <> "" Then
            Dim pcs As String = ""
            Call TIMS.SetMyValue(pcs, "PlanID", Request("PlanID"))
            Call TIMS.SetMyValue(pcs, "ComIDNO", Request("ComIDNO"))
            Call TIMS.SetMyValue(pcs, "SeqNO", Request("SeqNO"))
            Dim dr As DataRow = TIMS.Get_PlanInfo(objconn, pcs)
            If dr IsNot Nothing Then
                trainValue.Value = dr("TMID").ToString
                If dr("TMID").ToString <> "" Then TB_career_id.Text = "[" & dr("TrainID").ToString & "]" & dr("TrainName").ToString
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, historyrid, "HistoryList2", "RIDValue", "center")
        If historyrid.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        Dim strScript As String
        strScript = "<script>showPanel();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        'Button1.Attributes("onclick") = "CheckPrint();return false;"
    End Sub

#Region "(No Use)"

    'Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
    '    'Dim cGuid As String =   ReportQuery.GetGuid(Page)
    '    'Dim Url As String =   ReportQuery.GetUrl(Page)
    '    'Dim strScript As String
    '    'strScript = "<script language=""javascript"">" + vbCrLf
    '    'strScript += "window.open('" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=Maintest_GradeList&path=TIMS&OCID1=" & Me.OCIDValue1.Value & "');" + vbCrLf
    '    'strScript += "</script>"
    '    'Page.RegisterStartupScript("window_onload", strScript)
    '    'Response.Redirect("" & Url & "GUID=" + cGuid + "&SQ_AutoLogout=true&sys=list&filename=Maintest_GradeList&path=TIMS&OCID1=" & Me.OCIDValue1.Value & "")
    'End Sub

#End Region

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Dim dr_Class As DataRowView = e.Item.DataItem
        Select Case e.Item.ItemType
            Case ListItemType.Header
                Dim CheckboxAll As HtmlInputCheckBox = e.Item.FindControl("CheckboxAll")
                CheckboxAll.Attributes("onclick") = "ChangeAll(this);"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Checkbox1.Value = dr_Class("OCID")
                Checkbox1.Attributes("onclick") = "InsertValue(this.checked,this.value)"
                'If PrintValue.Value.IndexOf(Checkbox1.Value) <> -1 Then Checkbox1.Checked = True
                If PrintValue.Value.IndexOf(dr_Class("OCID")) <> -1 Then Checkbox1.Checked = True
        End Select
    End Sub


    Private Sub Query_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles query.Click
        Dim sql As String = ""
        Dim parms As Hashtable = New Hashtable()
        sql = ""
        sql += " SELECT a.OCID ,a.STDate ,a.FTDate ,a.CyclType "
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME" & vbCrLf
        sql += " FROM Class_ClassInfo a "
        sql += " WHERE 1=1 "

#Region "(No Use)"

        '當有指定班別時，系統將會忽略訓練職類及開訓期間
        'If OCIDValue1.Value <> "" Then
        '    sql += "and a.OCID='" & OCIDValue1.Value & "' "
        'Else
        '    sql += "and a.TMID='" & trainValue.Value & "' "
        '    If stdate1.Text <> "" Then sql += "and a.STDate >= '" & stdate1.Text & "' "
        '    If stdate2.Text <> "" Then sql += "and a.STDate <= '" & stdate2.Text & "' "
        'End If

#End Region

        Select Case Me.radiobtn1.SelectedValue
            Case 1
                sql += " AND a.OCID = @OCID "
                parms.Add("OCID", OCIDValue1.Value)
            Case Else
                sql += " AND a.TMID = @TMID "
                parms.Add("TMID", trainValue.Value)
                If STDate1.Text <> "" Then
                    'sql += " AND a.STDate >= " & TIMS.to_date(STDate1.Text) '& "','YYYY/MM/DD') " '★
                    sql += " AND a.STDate >= @STDate1 "
                    parms.Add("STDate1", STDate1.Text)
                End If
                If STDate2.Text <> "" Then
                    'sql += " AND a.STDate <= " & TIMS.to_date(STDate2.Text) 'convert(datetime, '" & stdate2.Text & "', 111) " '★
                    sql += " AND a.STDate <= @STDate2 "
                    parms.Add("STDate2", STDate2.Text)
                End If
        End Select
        sql += " ORDER BY a.OCID ,a.CyclType "
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        msg.Text = "查無資料!!"
        table4.Visible = False
        PageControler1.Visible = False
        Button1.Visible = False
        msg.Visible = True

        If dt.Rows.Count > 0 Then
            table4.Visible = True
            PageControler1.Visible = True
            Button1.Visible = True
            msg.Visible = False
            'pagecontroler1.SqlString = sql
            PageControler1.PageDataTable = dt
            PageControler1.PrimaryKey = "OCID"
            PageControler1.Sort = "CyclType"
            PageControler1.ControlerLoad()
        End If
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)  '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        table4.Visible = False
        msg.Visible = False
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        table4.Visible = False
    End Sub

    Private Sub radiobtn1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles radiobtn1.SelectedIndexChanged
        TB_career_id.Text = ""
        trainValue.Value = ""
        STDate1.Text = ""
        STDate2.Text = ""
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        table4.Visible = False
        msg.Visible = False
    End Sub

    'Button1 GO TO aspx row 135
    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        PrintValue.Value = TIMS.ClearSQM(PrintValue.Value)
        If PrintValue.Value = "" Then
            Common.MessageBox(Me, "請勾選本頁要列印的學員班級!")
            Exit Sub
        End If
        'Maintest_GradeList
        'Dim sFileName As String = "Maintest_GradeList"
        Dim xMyValue As String = ""
        xMyValue &= "&OCID1=" & PrintValue.Value
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, xMyValue)
    End Sub
End Class