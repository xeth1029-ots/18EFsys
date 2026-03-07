Partial Class TechID2
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me, 9)
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then Call Create()
        Close.Attributes("onclick") = "OpenProMenu(0);"
        Open.Attributes("onclick") = "OpenProMenu(1);"
        Close.Style("CURSOR") = "hand"
        Open.Style("CURSOR") = "hand"
        Me.DataGrid1.Attributes("name") = "DataGrid1"
        If State.Value = "0" Then
            Page.RegisterStartupScript("loading", "<script>OpenProMenu(0);</script>")
        Else
            Page.RegisterStartupScript("loading", "<script>OpenProMenu(1);</script>")
        End If
        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Button2.Visible = True
        Else
            Button2.Visible = False
        End If
    End Sub

    Sub Create()
        Dim rqProcessType As String = TIMS.ClearSQM(Request("ProcessType"))
        Dim rqPlanID As String = TIMS.ClearSQM(Request("PlanID"))
        Dim rqComIDNO As String = TIMS.ClearSQM(Request("ComIDNO"))
        Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqRID As String = TIMS.ClearSQM(Request("RID"))
        Dim rqCTName As String = TIMS.ClearSQM(Request("CTName"))
        HidCTName.Value = rqCTName
        TeachID.Value = ""
        TeachName.Value = ""

        Dim sRID As String = ""
        Dim sKindEngage1 As String = ""
        Dim sSqlWhere1 As String = ""

        Select Case rqProcessType
            Case "planid"
                sSqlWhere1 = ""
                sSqlWhere1 += " AND a.RID IN (SELECT RID "
                sSqlWhere1 += " FROM Plan_PlanInfo "
                sSqlWhere1 += " WHERE PlanID='" & rqPlanID & "'"
                sSqlWhere1 += " AND ComIDNO='" & rqComIDNO & "'"
                sSqlWhere1 += " AND SeqNo='" & rqSeqNo & "')"
                sKindEngage1 = KindEngage1.SelectedValue
                sKindEngage1 = TIMS.ClearSQM(sKindEngage1)
                If sKindEngage1 = "%" Then sKindEngage1 = ""
            Case Else
                sRID = rqRID
                'If TeachCName.Text = "%" Then TeachCName.Text = ""
                'If TeachCName.Text <> "" Then TeachCName.Text = Trim(TeachCName.Text)
                'sTeachCName = TeachCName.Text
                'If TeacherID.Text <> "" Then TeacherID.Text = Trim(TeacherID.Text)
                'sTeacherID = TeacherID.Text
                sKindEngage1 = KindEngage1.SelectedValue
                sKindEngage1 = TIMS.ClearSQM(sKindEngage1)
                If sKindEngage1 = "%" Then sKindEngage1 = ""
        End Select

        Call sUtl_Search1(sRID, "", "", sKindEngage1, sSqlWhere1)
    End Sub

    Private Sub KindEngage_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles KindEngage.SelectedIndexChanged
        Call Create()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                e.Item.CssClass = "TR_04002_TD"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim Radio1 As HtmlInputRadioButton = e.Item.FindControl("Radio1")
                Dim Checkbox1 As HtmlInputCheckBox = e.Item.FindControl("Checkbox1")
                Dim drv As DataRowView = e.Item.DataItem
                Radio1.Value = TIMS.reValue(drv("TechID")) '.ToString
                Checkbox1.Value = TIMS.reValue(drv("TechID")) '.ToString
                Radio1.Visible = True
                Checkbox1.Visible = False
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    Radio1.Visible = False
                    Checkbox1.Visible = True
                End If
                If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                    If TIMS.CHK_INDEX(Convert.ToString(drv("TechID")), HidCTName.Value) Then
                        Checkbox1.Checked = True
                        Dim sTmp As String = ""
                        sTmp = TIMS.reValue(drv("TechID"))
                        If TeachID.Value <> "" Then TeachID.Value &= ","
                        TeachID.Value &= sTmp
                        sTmp = TIMS.reValue(drv("TeachCName"))
                        If TeachName.Value <> "" Then TeachName.Value &= ","
                        TeachName.Value &= sTmp
                        sTmp = TIMS.reValue(drv("DegreeID"))
                        If DegreeID.Value <> "" Then DegreeID.Value &= ","
                        DegreeID.Value &= sTmp
                        sTmp = TIMS.reValue(drv("DegreeName"))
                        If DegreeName.Value <> "" Then DegreeName.Value &= ","
                        DegreeName.Value &= sTmp
                        sTmp = TIMS.reValue(drv("Major"))
                        If Major.Value <> "" Then Major.Value &= ","
                        DegreeName.Value &= sTmp
                    End If
                End If
                'Radio1.Attributes("onclick") = "ReturnTechID('" & drv("TechID") & "','" & drv("TeachCName") & "','" & drv("DegreeID") & "','" & drv("DegreeName") & "','" & drv("Major") & "')"
                Dim drv_Major As String = TIMS.reValue(drv("Major"))
                drv_Major = drv_Major.Replace(vbCrLf, "\n")
                Checkbox1.Attributes("onclick") = "SelectTechID(this.checked,'" & TIMS.reValue(drv("TechID")) & "','" & TIMS.reValue(drv("TeachCName")) & "','" & TIMS.reValue(drv("DegreeID")) & "','" & TIMS.reValue(drv("DegreeName")) & "','" & drv_Major & "');"
                e.Item.Cells(1).Text = "外聘"
                If Convert.ToString(drv("KindEngage")) = "1" Then
                    e.Item.Cells(1).Text = "內聘"
                End If
            Case ListItemType.Footer
                If DataGrid1.Items.Count = 0 Then
                    DataGrid1.ShowFooter = True
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(New TableCell)
                    e.Item.Cells(0).ColumnSpan = DataGrid1.Columns.Count
                    e.Item.Cells(0).Text = "查無資料!"
                    e.Item.Cells(0).HorizontalAlign = HorizontalAlign.Center
                Else
                    DataGrid1.ShowFooter = False
                End If
        End Select
    End Sub

    Sub sUtl_Search1(ByVal sRID As String, ByVal sTeachCName As String, ByVal sTeacherID As String, ByVal sKindEngage1 As String, ByVal sSqlWhere1 As String)
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.TechID, a.TeachCName, a.DegreeID, c.Name DegreeName " & vbCrLf
        sql += "       ,replace(ISNULL(a.Specialty1,' '),',',' ') + replace(ISNULL(a.Specialty2,' '),',',' ') " & vbCrLf
        sql += "        + replace(ISNULL(a.Specialty3,' '),',',' ') + replace(ISNULL(a.Specialty4,' '),',',' ') " & vbCrLf
        sql += "        + replace(ISNULL(a.Specialty5,' '),',',' ') AS Major " & vbCrLf
        'sql += "      ,CASE ISNULL(CONVERT(varchar, b.TechID),' ') WHEN ' ' THEN 'N' ELSE 'Y' END ptchk " & vbCrLf
        sql += "       ,a.KindEngage ,a.TeacherID " & vbCrLf
        sql += " FROM Teach_TeacherInfo a " & vbCrLf
        sql += " LEFT JOIN key_Degree c ON a.DegreeID = c.DegreeID " & vbCrLf
        sql += " WHERE 1=1 AND a.WorkStatus = '1' " & vbCrLf
        ' from Teach_TeacherInfo a
        If sSqlWhere1 <> "" Then sql += sSqlWhere1
        If sRID <> "" Then sql += " AND a.RID = '" & sRID & "' " & vbCrLf
        'If TeachCName.Text = "%" Then TeachCName.Text = ""
        If sTeachCName <> "" Then
            sTeachCName = UCase(sTeachCName)
            sql += " AND (1!=1 "
            'sql += " OR regexp_like (a.TeachCName ,N'" & sTeachCName & "','i') " & vbCrLf
            'sql += " OR regexp_like (a.TeachEName ,N'" & sTeachCName & "','i') " & vbCrLf
            sql += " OR UPPER(a.TeachCName) LIKE '%" & sTeachCName & "%' " & vbCrLf
            sql += " OR UPPER(a.TeachEName) LIKE '%" & sTeachCName & "%' " & vbCrLf
            sql += " ) " & vbCrLf
        End If
        'If TeacherID.Text = "%" Then TeacherID.Text = ""
        If sTeacherID <> "" Then
            sTeacherID = UCase(sTeacherID)
            'sql += " AND regexp_like (a.TeacherID ,'" & sTeacherID & "','i') " & vbCrLf
            sql += " AND UPPER(a.TeacherID) LIKE '%" & sTeacherID & "%' " & vbCrLf
        End If
        Select Case sKindEngage1
            Case "1", "2"
                sql += " AND a.KindEngage = '" & sKindEngage1 & "' " & vbCrLf
        End Select
        sql += " ORDER BY a.KindEngage ,a.TeachCName ,a.TeacherID " & vbCrLf
        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '清空暫存值
        TeachID.Value = ""
        TeachName.Value = ""
        DegreeID.Value = ""
        DegreeName.Value = ""
        Dim sRID As String = ""
        'Dim sTeachCName As String = ""
        'Dim sTeacherID As String = ""
        'Dim sKindEngage1 As String = ""
        'Dim sSqlWhere1 As String = ""
        sRID = TIMS.ClearSQM(Request("RID"))
        If TeachCName.Text = "%" Then TeachCName.Text = ""
        TeachCName.Text = TIMS.ClearSQM(TeachCName.Text)
        TeacherID.Text = TIMS.ClearSQM(TeacherID.Text)
        Dim v_KindEngage1 As String = TIMS.GetListValue(KindEngage1)
        If v_KindEngage1 = "%" Then v_KindEngage1 = ""
        'If TeachCName.Text <> "" Then TeachCName.Text = Trim(TeachCName.Text)
        'sTeachCName = TIMS.ClearSQM(TeachCName.Text)
        'If TeacherID.Text <> "" Then TeacherID.Text = Trim(TeacherID.Text)
        'sTeacherID = TIMS.ClearSQM(TeacherID.Text)
        'sKindEngage1 = KindEngage1.SelectedValue
        Call sUtl_Search1(sRID, TeachCName.Text, TeacherID.Text, v_KindEngage1, "")
#Region "(No Use)"

        'sql = ""
        'sql &= " SELECT * FROM "
        'sql += "       (SELECT a.*, "
        'sql += "               case when IsNull(Specialty1,'') ='' then '' "
        'sql += "               else replace(Specialty1,',','') end+ "
        'sql += "               case when IsNull(Specialty2,'') ='' then '' "
        'sql += "               else replace(Specialty2,',','') end+ "
        'sql += "               case when IsNull(Specialty3,'') ='' then '' "
        'sql += "               else replace(Specialty3,',','') end+ "
        'sql += "               case when IsNull(Specialty4,'') ='' then '' "
        'sql += "               else replace(Specialty4,',','') end+ "
        'sql += "               case when IsNull(Specialty5,'') ='' then '' "
        'sql += "               else replace(Specialty5,',','') end as major, "
        'sql += "               k.name DegreeName "
        'sql += "          FROM Teach_TeacherInfo a "
        'sql += "          left join key_degree k on a.DegreeID=k.DegreeID) Teach_TeacherInfo "
        'sql += "         WHERE WorkStatus = '1' and RID='" & Request("RID") & "' "
        'sql += "           and (TeachCName like '%" & Replace(TeachCName.Text, " ", "%") & "%' or TeachEName like '%" & Replace(TeachCName.Text, " ", "%") & "%') "
        'sql += "           and TeacherID like '%" & Replace(TeacherID.Text, " ", "%") & "%' "
        'sql += "           and KindEngage like '" & KindEngage1.SelectedValue & "' "

#End Region
    End Sub
End Class