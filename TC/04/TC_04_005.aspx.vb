Partial Class TC_04_005
    Inherits AuthBasePage

    '未檢送資料註記
    Const cst_tc_msg1 As String = "勾選後，請記得按儲存，才會有儲存資料"
    Const cst_tc_msg2 As String = "(勾選後，請記得按儲存，才會有儲存資料)"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then LabTMID.Text = "訓練業別"

        Pagecontroler1.PageDataGrid = dgPlan

        Labmsg1.Text = cst_tc_msg1
        Labmsg2.Text = cst_tc_msg2

        If Not IsPostBack Then
            cCreate1()
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Org.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

    End Sub

    Sub cCreate1()
        '有不區分
        OrgKind2 = TIMS.Get_RblSearchPlan(Me, OrgKind2)
        Common.SetListItem(OrgKind2, "A")

        '依申請階段 '表示 (1：上半年、2：下半年、3：政策性產業)
        If tr_AppStage_TP28.Visible Then
            AppStage2 = TIMS.Get_AppStage2_NotCase(AppStage2)
            Common.SetListItem(AppStage2, "")
        End If

        '取得訓練計畫
        TPlanid.Value = sm.UserInfo.TPlanID 'DbAccess.ExecuteScalar(Sqlstr, objconn)
        '(加強操作便利性)2005/4/1-melody
        RIDValue.Value = sm.UserInfo.RID
        'Sqlstr = "select orgname from Auth_Relship a join Org_orginfo b on  a.orgid=b.orgid where a.RID='" & sm.UserInfo.RID & "'"
        center.Text = sm.UserInfo.OrgName 'DbAccess.ExecuteScalar(Sqlstr, objconn)

        DataGridTable1.Visible = False
        '勾選框僅提供分署、署勾選，訓練單位成反灰，並隱藏「儲存」按鈕。
        If sm.UserInfo.LID = 2 Then
            btnQuery.Enabled = False
            BtnSaveData1.Enabled = False
            TIMS.Tooltip(btnQuery, TIMS.cst_ErrorMsg14, True)
            TIMS.Tooltip(BtnSaveData1, TIMS.cst_ErrorMsg14, True)
        End If
    End Sub

    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, dgPlan)

        UNIT_SDATE.Text = TIMS.ClearSQM(UNIT_SDATE.Text)
        UNIT_EDATE.Text = TIMS.ClearSQM(UNIT_EDATE.Text)
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        ClassName.Text = TIMS.ClearSQM(ClassName.Text)
        CyclType.Text = TIMS.ClearSQM(CyclType.Text)

        UNIT_SDATE.Text = TIMS.Cdate3(UNIT_SDATE.Text)
        UNIT_EDATE.Text = TIMS.Cdate3(UNIT_EDATE.Text)
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        'Dim vRelship As String = TIMS.GET_RelshipforRID(RIDValue.Value, objconn)

        Dim sql As String = ""
        sql &= " SELECT P1.PLANID,P1.COMIDNO,P1.SEQNO" & vbCrLf
        sql &= " ,dbo.FN_GET_ROC_YEAR(P2.YEARS) PlanYear_ROC" & vbCrLf
        sql &= " ,format(P1.APPLIEDDATE,'yyyy/MM/dd') APPLIEDDATE" & vbCrLf
        sql &= " ,format(P1.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(P1.FDDATE,'yyyy/MM/dd') FDDATE" & vbCrLf
        sql &= " ,P2.DISTNAME" & vbCrLf
        sql &= " ,A1.RID" & vbCrLf
        sql &= " ,O1.ORGNAME" & vbCrLf
        sql &= " ,C1.OCID" & vbCrLf
        sql &= " ,P1.CLASSNAME" & vbCrLf
        sql &= " ,P1.CYCLTYPE" & vbCrLf
        sql &= " ,p1.DataNotSent" & vbCrLf
        sql &= " FROM dbo.PLAN_PLANINFO P1" & vbCrLf
        sql &= " JOIN dbo.AUTH_RELSHIP A1 ON P1.RID=A1.RID" & vbCrLf
        sql &= " JOIN dbo.ORG_ORGINFO O1 ON A1.OrgID=O1.OrgID" & vbCrLf
        sql &= " JOIN dbo.VIEW_PLAN P2 on P2.PLANID=P1.PLANID" & vbCrLf
        sql &= " LEFT JOIN dbo.PLAN_VERREPORT pvr ON P1.PlanID = pvr.PlanID AND  P1.ComIDNO = pvr.ComIDNO AND P1.SeqNO = pvr.SeqNo" & vbCrLf
        sql &= " LEFT JOIN dbo.CLASS_CLASSINFO C1 on C1.PLANID=P1.PLANID AND C1.COMIDNO=P1.COMIDNO AND C1.SEQNO=P1.SEQNO" & vbCrLf
        'sql &= " WHERE vp.TPLANID='28'" & vbCrLf'sql &= " AND vp.YEARS='2020'" & vbCrLf
        sql &= " WHERE pvr.IsApprPaper='Y'" & vbCrLf '正式
        sql &= " AND P2.PlanKind=2" & vbCrLf '計畫種類:1.自辦／2.委外
        '依登入年度
        sql &= " AND P2.Years='" & sm.UserInfo.Years & "'" & vbCrLf
        '依登入計畫
        sql &= " AND P2.TPlanID ='" & sm.UserInfo.TPlanID & "'" & vbCrLf
        'DistType: 搜尋型態: 0:依轄區 1:依訓練機構
        'Dim v_DistType As String = TIMS.GetListValue(DistType)
        '依訓練機構
        Select Case sm.UserInfo.LID
            Case 0 '署-非轄區
                If RIDValue.Value.Length = 1 Then
                    Dim sDISTID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)
                    If sDISTID <> "000" Then sql &= " AND P2.DistID ='" & sDISTID & "'"
                    If sDISTID = "" Then sql &= " AND 1!=1"
                ElseIf RIDValue.Value.Length > 1 Then
                    sql &= " AND P1.RID ='" & RIDValue.Value & "'"
                End If
            Case 1 '轄區
                sql &= " AND P2.PLANID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                sql &= " AND P2.DISTID='" & sm.UserInfo.DistID & "'" & vbCrLf
                If RIDValue.Value.Length > 0 AndAlso RIDValue.Value <> sm.UserInfo.RID Then
                    sql &= " AND P1.RID ='" & RIDValue.Value & "'"
                End If
            Case 2 '委訓限定
                sql &= " AND P2.PLANID='" & sm.UserInfo.PlanID & "'" & vbCrLf
                sql &= " AND P2.DISTID='" & sm.UserInfo.DistID & "'" & vbCrLf
                sql &= " AND P1.RID ='" & sm.UserInfo.RID & "'"
        End Select

        Dim v_OrgKind2 As String = TIMS.GetListValue(OrgKind2)
        Dim v_PlanMode As String = TIMS.GetListValue(PlanMode)
        '依申請階段
        Dim v_AppStage2 As String = "" 'TIMS.GetListValue(AppStage2)
        If tr_AppStage_TP28.Visible Then v_AppStage2 = TIMS.GetListValue(AppStage2)

        Select Case v_PlanMode'PlanMode.SelectedValue
            Case "S" '審核中的
                sql &= " AND pvr.SecResult IS NULL" & vbCrLf
                '產投不判斷 P1.AppliedResult 依 pvr.SecResult 為準
                'sql += " AND P1.AppliedResult IS NULL" & vbCrLf 'sql += " AND P1.ResultButton IS NULL" & vbCrLf 'NULL:已送出不可修改 Y:還原可修改
            Case "Y" '已通過
                sql &= " AND pvr.SecResult='Y'" & vbCrLf
            Case "R" '退件修正
                sql &= " AND pvr.SecResult in ('R','N')" & vbCrLf
        End Select

        If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            'Me.LabTMID.Text = "訓練業別"
            If Me.jobValue.Value <> "" Then
                sql &= " AND (P1.TMID = " & jobValue.Value & vbCrLf
                sql &= " OR P1.TMID IN (" & vbCrLf
                sql &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '職類別
                sql &= " select TMID from Key_TrainType where parent IN (" & vbCrLf '業別
                sql &= " select TMID from Key_TrainType where busid ='G')" & vbCrLf '產業人才投資方案類
                sql &= " AND TMID =" & jobValue.Value & " )))" & vbCrLf
            End If
        Else
            If Me.trainValue.Value <> "" Then
                sql &= " AND P1.TMID = " & Me.trainValue.Value & vbCrLf
            End If
        End If

        '通俗職類
        If txtCJOB_NAME.Text <> "" Then sql &= " AND P1.CJOB_UNKEY = " & cjobValue.Value & "" & vbCrLf

        If Me.UNIT_SDATE.Text <> "" Then sql &= " AND P1.AppliedDate >= " & TIMS.To_date(Me.UNIT_SDATE.Text) & vbCrLf

        If Me.UNIT_EDATE.Text <> "" Then sql &= " AND P1.AppliedDate <= " & TIMS.To_date(Me.UNIT_EDATE.Text) & vbCrLf

        If Me.start_date.Text <> "" Then sql &= " AND P1.STDate >= " & TIMS.To_date(Me.start_date.Text) & vbCrLf

        If Me.end_date.Text <> "" Then sql &= " AND P1.STDate <= " & TIMS.To_date(Me.end_date.Text) & vbCrLf

        If ClassName.Text <> "" Then sql &= " AND P1.ClassName like '%" & ClassName.Text & "%'"

        If CyclType.Text <> "" Then
            If CyclType.Text.Length < 2 Then CyclType.Text = "0" & CInt(Val(CyclType.Text))
            sql &= " AND P1.CyclType='" & CyclType.Text & "'"
        End If

        Select Case v_OrgKind2'OrgKind2.SelectedValue
            Case "G", "W"
                'sql &= " AND O1.OrgKind2='" & OrgKind2.SelectedValue & "'"
                sql &= " AND O1.OrgKind2='" & v_OrgKind2 & "'"
        End Select
        '依申請階段
        If v_AppStage2 <> "" Then sql &= " AND P1.AppStage= '" & v_AppStage2 & "'" & vbCrLf
        '檢送資料-未檢送 未檢送資料
        If CB_DataNotSent_SCH.Checked Then sql &= " AND P1.DataNotSent='Y'" & vbCrLf
        '排序
        sql &= " ORDER BY O1.OrgName,P1.STDate,P1.ClassName"

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        DataGridTable1.Visible = False
        msg1.Text = TIMS.cst_NODATAMsg1
        If dt.Rows.Count = 0 Then Return

        DataGridTable1.Visible = True
        msg1.Text = ""
        Pagecontroler1.PageDataTable = dt
        Pagecontroler1.ControlerLoad()
    End Sub

    Protected Sub btnQuery_Click(sender As Object, e As EventArgs) Handles btnQuery.Click
        Call sSearch1()
    End Sub

    Private Sub dgPlan_ItemDataBound(sender As Object, e As DataGridItemEventArgs) Handles dgPlan.ItemDataBound
        Const Cst_Index As Integer = 0

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem '帶入一行資料
                e.Item.Cells(Cst_Index).Text = e.Item.ItemIndex + 1 + dgPlan.CurrentPageIndex * dgPlan.PageSize
                Dim CB_DataNotSent As CheckBox = e.Item.FindControl("CB_DataNotSent")
                Dim Hid_PCS As HiddenField = e.Item.FindControl("Hid_PCS")
                Dim Hid_OCID As HiddenField = e.Item.FindControl("Hid_OCID")
                Dim HID_DataNotSent As HiddenField = e.Item.FindControl("HID_DataNotSent") '未檢送資料 未檢送資料註記

                Hid_PCS.Value = String.Format("{0}x{1}x{2}", drv("PlanID"), drv("COMIDNO"), drv("SEQNO"))
                Hid_OCID.Value = String.Format("{0}", drv("OCID"))
                HID_DataNotSent.Value = Convert.ToString(drv("DataNotSent")) '未檢送資料註記
                CB_DataNotSent.Checked = If(Convert.ToString(drv("DataNotSent")).Equals("Y"), True, False)

            Case ListItemType.Footer
                If dgPlan.Items.Count = 0 Then
                    dgPlan.ShowFooter = True
                    Dim mycell As New TableCell
                    mycell.ColumnSpan = e.Item.Cells.Count
                    mycell.Text = "目前沒有任何資料!"
                    e.Item.Cells.Clear()
                    e.Item.Cells.Add(mycell)
                    e.Item.HorizontalAlign = HorizontalAlign.Center
                Else
                    dgPlan.ShowFooter = False
                End If
        End Select
    End Sub

    Sub SaveData1()
        '未檢送資料
        Dim iRowAdd As Integer = 0
        Dim iRowDel As Integer = 0
        For Each eItem As DataGridItem In dgPlan.Items
            Dim CB_DataNotSent As CheckBox = eItem.FindControl("CB_DataNotSent")
            Dim Hid_PCS As HiddenField = eItem.FindControl("Hid_PCS")
            Dim HID_DataNotSent As HiddenField = eItem.FindControl("HID_DataNotSent")
            Dim v_DataNotSent As String = If(CB_DataNotSent.Checked, "Y", "N")

            If HID_DataNotSent.Value = "" AndAlso v_DataNotSent.Equals("Y") Then
                iRowAdd += 1
                Exit For
            ElseIf HID_DataNotSent.Value <> "" AndAlso v_DataNotSent.Equals("N") Then
                iRowDel += 1
                Exit For
            End If
        Next
        If iRowAdd = 0 AndAlso iRowDel = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg13)
            Return
        End If

        For Each eItem As DataGridItem In dgPlan.Items
            Dim CB_DataNotSent As CheckBox = eItem.FindControl("CB_DataNotSent")
            Dim Hid_PCS As HiddenField = eItem.FindControl("Hid_PCS")
            Dim HID_DataNotSent As HiddenField = eItem.FindControl("HID_DataNotSent")
            Dim v_DataNotSent As String = If(CB_DataNotSent IsNot Nothing, If(CB_DataNotSent.Checked, "Y", "N"), "")

            If HID_DataNotSent.Value = "" AndAlso v_DataNotSent.Equals("Y") Then
                '未檢送資料
                Dim parms As New Hashtable
                parms.Clear()
                parms.Add("DataNotSent", v_DataNotSent)
                parms.Add("MODIFYACCT", sm.UserInfo.UserID)
                parms.Add("PCS", Hid_PCS.Value)
                Dim sql As String = ""
                sql = ""
                sql &= " UPDATE PLAN_PLANINFO"
                sql &= " SET DataNotSent=@DataNotSent" 'Y
                sql &= " ,MODIFYACCT=@MODIFYACCT"
                sql &= " ,MODIFYDATE=GETDATE()"
                sql &= " WHERE CONCAT(PLANID,'x',COMIDNO,'x',SEQNO)=@PCS"
                DbAccess.ExecuteNonQuery(sql, objconn, parms)
            ElseIf HID_DataNotSent.Value <> "" AndAlso v_DataNotSent.Equals("N") Then
                '未檢送資料
                Dim parms_2 As New Hashtable
                parms_2.Clear()
                parms_2.Add("MODIFYACCT", sm.UserInfo.UserID)
                parms_2.Add("PCS", Hid_PCS.Value)
                Dim sql_2 As String = ""
                sql_2 = ""
                sql_2 &= " UPDATE PLAN_PLANINFO"
                sql_2 &= " SET DataNotSent=NULL" 'v_DataNotSent.Equals("N") NULL
                sql_2 &= " ,MODIFYACCT=@MODIFYACCT"
                sql_2 &= " ,MODIFYDATE=GETDATE()"
                sql_2 &= " WHERE CONCAT(PLANID,'x',COMIDNO,'x',SEQNO)=@PCS"
                DbAccess.ExecuteNonQuery(sql_2, objconn, parms_2)
            End If
        Next

        Common.MessageBox(Me, "儲存成功")

        Call sSearch1()
    End Sub

    Protected Sub BtnSaveData1_Click(sender As Object, e As EventArgs) Handles BtnSaveData1.Click
        SaveData1()
    End Sub

End Class