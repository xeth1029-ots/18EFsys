Partial Class SD_14_018
    Inherits AuthBasePage

    '432 訓練計畫專職/工作人員名冊
    '訓練計畫專職人員名冊	
    'SD_14_018 (空白表單)
    'SD_14_018_2 (含有資料。)
    'Dim dt As DataTable = Nothing
    'Dim oCmd As SqlCommand = Nothing
    'ORG_MEMBER
    'SD_14_018_2*.jrxml

    Const cst_printFN1 As String = "SD_14_018_2"

    Const cst_cmdname_edt As String = "edt"
    Const cst_cmdname_edt2 As String = "edt2"

    Const cst_func_lock1 As String = "lock1"
    Const cst_func_lock2 As String = "lock2"

    'Const cst_FSQ1_上半年 As String = "01"
    'Const cst_FSQ1_下半年 As String = "02"
    'Const cst_FSQ1_政策性產業 As String = "03"
    'Const cst_FSQ1_進階政策性產業 As String = "04"
    'Const cst_FSQ1_上半年N As String = "上半年"
    'Const cst_FSQ1_下半年N As String = "下半年"
    'Const cst_FSQ1_政策性產業N As String = "政策性產業"
    'Const cst_FSQ1_進階政策性產業N As String = " 進階政策性產業"

    'Dim au As New cAUTH
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
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End

        If Not IsPostBack Then
            cCreate1()
        End If

        'TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center", "hbtnSearch")
        'If HistoryRID.Rows.Count <> 0 Then
        '    center.Attributes("onclick") = "showObj('HistoryList2');"
        '    center.Style("CURSOR") = "hand"
        'End If
        'Years.Value = sm.UserInfo.Years - 1911

    End Sub

    Sub cCreate1()
        btnAdd.Visible = False
        btnUPD.Visible = False
        btnCancel.Visible = False
        BtnPrint1.Visible = False

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        orgid_value.Value = sm.UserInfo.OrgID

        rblFSQ1_S = TIMS.Get_APPSTAGE_FSQ1(rblFSQ1_S)
        Common.SetListItem(rblFSQ1_S, "00")

        ddlFSQ1_A = TIMS.Get_APPSTAGE_FSQ1(ddlFSQ1_A)
        'Common.SetListItem(rblFSQ1_S, "00")

        ddl_year = TIMS.Get_Years(ddl_year, objconn)
        Common.SetListItem(ddl_year, sm.UserInfo.Years)
        'Call GetListYears()
        'Call Search1()

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        If Convert.ToString(sm.UserInfo.RID).Length > 1 Then
            '委訓單位?
            Button2.Disabled = True
            center.Enabled = False
        Else
            Button2.Disabled = False
            center.Enabled = True

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx?btnName=btnSearch');"
            Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))
        End If

    End Sub

    Private Sub DataGrid1_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        hidsOMID.Value = ""
        If e.CommandArgument = "" Then Exit Sub
        Dim sCmdArg As String = e.CommandArgument
        Dim OMID As String = TIMS.GetMyValue(sCmdArg, "OMID")
        If OMID = "" Then Exit Sub
        Dim ROWNUM1 As String = TIMS.GetMyValue(sCmdArg, "ROWNUM1")
        'If ROWNUM1 = "" Then Exit Sub

        Select Case e.CommandName
            Case cst_cmdname_edt '修改
                Call Utl_EDITU1(OMID, ROWNUM1)
            Case cst_cmdname_edt2 '修改-僅可修改電話
                Call Utl_EDITU2(OMID, ROWNUM1) '修改-僅可修改電話
            Case "del" '刪除
                Call Utl_DELETE1(OMID)
                Call Search1()
            Case "lock1" '解鎖
                Call Utl_UNLOCK(OMID)
                Call Search1()
        End Select
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim hidOMID As HtmlInputHidden = e.Item.FindControl("hidOMID")
                Dim HidROWNUM1 As HtmlInputHidden = e.Item.FindControl("HidROWNUM1")

                Dim LabFSQ1_DG As Label = e.Item.FindControl("LabFSQ1_DG")
                Dim btEdit As Button = e.Item.FindControl("btEdit")
                Dim btDel As Button = e.Item.FindControl("btDel")
                Dim btLOCK As Button = e.Item.FindControl("btLOCK")

                Dim txt_FSQ1 As String = TIMS.GET_APPSTAGE_FSQ1_N(Convert.ToString(drv("FSQ1")))
                LabFSQ1_DG.Text = txt_FSQ1 'cst_FSQ1_上半年N/cst_FSQ1_下半年N

                HidROWNUM1.Value = Convert.ToString(drv("ROWNUM1"))
                hidOMID.Value = Convert.ToString(drv("OMID"))
                Dim sCmdArg As String = ""
                TIMS.SetMyValue(sCmdArg, "OMID", Convert.ToString(drv("OMID")))
                TIMS.SetMyValue(sCmdArg, "ROWNUM1", Convert.ToString(drv("ROWNUM1")))

                btEdit.CommandArgument = sCmdArg 'Convert.ToString(drv("OMID"))
                btDel.CommandArgument = sCmdArg 'Convert.ToString(drv("OMID"))
                btDel.Attributes("onclick") = TIMS.cst_confirm_delmsg1

                btLOCK.CommandArgument = sCmdArg 'Convert.ToString(drv("OMID"))
                btLOCK.Attributes("onclick") = "return confirm('您確定要解鎖這一筆資料?');"

                'https://jira.turbotech.com.tw/browse/TIMSC-277
                '2018年度適用，訓練計畫專職人員名冊功能修改，將列印空白表單功能拿掉，
                '訓練單位新增專職人員時，可選擇年度別 (上半年/下半年)，
                '人員新增後訓練單位無法自行進行修改，修改鈕是灰階狀態，
                '訓練單位使用時，不用顯示刪除和解鎖鈕，
                '需向分署申請開啟修改權限後，訓練單位使用時的修改鈕才會是正常可按的，
                '此部分僅針對個別的人員修改，分署使用此功能時，人員清單右方有修改、刪除和解鎖鈕，
                '刪除及解鎖鈕只有分署以上層級使用時，會看到也可以用，按下解鎖鈕，
                '訓練單位的該人員修改鈕才可以按，並可進行修改。

                '(預設)委訓單位
                btEdit.Visible = False
                btDel.Visible = False
                btLOCK.Visible = False
                btLOCK.Enabled = True '有效
                Select Case Convert.ToString(sm.UserInfo.LID)
                    Case "0", "1" '分署以上層級使用
                        btEdit.Visible = True
                        btDel.Visible = True
                        btLOCK.Visible = True
                        If Convert.ToString(drv("LOCK1")) = "N" Then
                            btLOCK.Enabled = False '失效
                            TIMS.Tooltip(btEdit, "資料已解鎖!!")
                        End If
                    Case "2"
                        '委訓單位
                        '(已解鎖)
                        If Convert.ToString(drv("LOCK1")) = "N" Then
                            btEdit.CommandName = cst_cmdname_edt
                            btEdit.Visible = True
                            TIMS.Tooltip(btEdit, "資料已解鎖，可修改")
                        Else
                            btEdit.CommandName = cst_cmdname_edt2
                            btEdit.Visible = True
                            TIMS.Tooltip(btEdit, "資料未解鎖，僅可修改部份資訊")
                        End If
                End Select
        End Select
    End Sub

#Region "(No Use)"

    'Private Sub GetListYears()
    '    Dim sql As String = ""
    '    sql = " select distinct Years from id_plan where tplanid=@tplanid "
    '    dt = New DataTable
    '    oCmd = New SqlCommand(sql, objconn)
    '    With oCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("tplanid", SqlDbType.VarChar).Value = sm.UserInfo.TPlanID
    '        dt.Load(.ExecuteReader())
    '    End With
    '    If dt.Rows.Count > 0 Then
    '        ddl_year.DataSource = dt
    '        ddl_year.DataTextField = "Years"
    '        ddl_year.DataValueField = "Years"
    '        ddl_year.DataBind()
    '        Common.SetListItem(ddl_year, sm.UserInfo.Years)
    '        'ddl_year.SelectedValue = sm.UserInfo.Years
    '    End If
    'End Sub

    '列印空白表單
    'Private Sub button5_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button5.ServerClick
    '    Dim errMsg As String = ""
    '    If RIDValue.Value = "" Then
    '        Errmsg &= "訓練機構尚未選取！\n"
    '    End If
    '    If ddl_year.SelectedValue = "" Then
    '        Errmsg &= "計畫年度尚未選取！\n"
    '    End If
    '    If errMsg <> "" Then
    '        '錯誤
    '        Page.RegisterStartupScript("errMsg", "<script>alert('" & errMsg & "');</script>")
    '        Exit Sub
    '    Else
    '        Dim MyValue As String = ""
    '        MyValue = "Years=" & ddl_year.SelectedValue & "&RID=" & RIDValue.Value & "&PlanID=" & sm.UserInfo.PlanID & ""
    '        If sm.UserInfo.RID = "A" AndAlso RIDValue.Value <> sm.UserInfo.RID Then
    '            '職訓局 選擇其他機構列印
    '            Dim sql As String = "SELECT rwPlanID FROM VIEW_RWPLANRID WHERE TPLANID='" & sm.UserInfo.TPlanID & "' AND RID='" & RIDValue.Value & "' AND YEARS='" & ddl_year.SelectedValue & "'"
    '            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
    '            If Not dr Is Nothing Then
    '                MyValue = "Years=" & ddl_year.SelectedValue & "&RID=" & RIDValue.Value & "&PlanID=" & dr("rwPlanID").ToString & ""
    '            End If
    '        End If
    '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "Report", "SD_14_018", MyValue)
    '    End If
    'End Sub

#End Region

    '清除重設
    Sub ClearList1()
        tROWNUM1.Text = ""
        tTitle.Text = ""
        tName.Text = ""
        tPhone.Text = ""
        btnAdd.Visible = True
        btnUPD.Visible = False
        btnCancel.Visible = False
        BtnPrint1.Visible = False

        hidsOMID.Value = ""
        Common.SetListItem(ddlFSQ1_A, "01") 'cst_FSQ1_上半年
        Utl_Lockdata2("")
    End Sub

    Sub Utl_Lockdata2(ByVal func1 As String)
        tTitle.Enabled = True
        tName.Enabled = True
        ddlFSQ1_A.Enabled = True
        Select Case func1
            Case cst_func_lock1
                ddlFSQ1_A.Enabled = False
                TIMS.Tooltip(ddlFSQ1_A, "不提供此項目的修改!", True)
            Case cst_func_lock2
                tTitle.Enabled = False
                tName.Enabled = False
                ddlFSQ1_A.Enabled = False
                TIMS.Tooltip(ddlFSQ1_A, "不提供此項目的修改!", True)
        End Select
    End Sub

    '查詢
    Sub Search1()
        Call ClearList1() '重設

        Dim v_rblFSQ1_S As String = TIMS.GetListValue(rblFSQ1_S)
        Dim v_rblSORT_TYPE1 As String = TIMS.GetListValue(rblSORT_TYPE1)

        'ALTER  TABLE [dbo].[ORG_MEMBER] ADD [SORTNO1] [float] null
        Dim sql As String = ""
        sql &= " SELECT a.OMID" & vbCrLf '/*PK*/
        sql &= " ,ISNULL(a.SORTNO1,(ROW_NUMBER() OVER(ORDER BY a.MODIFYDATE))) AS ROWNUM1"
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,a.PLANID" & vbCrLf
        sql &= " ,a.YEARS" & vbCrLf
        sql &= " ,a.TITLE1" & vbCrLf
        sql &= " ,a.CNAME" & vbCrLf
        sql &= " ,a.PHONE1" & vbCrLf
        sql &= " ,a.MODIFYACCT" & vbCrLf
        sql &= " ,a.MODIFYDATE" & vbCrLf
        sql &= " ,a.FSQ1" & vbCrLf
        sql &= " ,a.LOCK1" & vbCrLf
        sql &= " ,a.SORTNO1" & vbCrLf
        sql &= " FROM ORG_MEMBER a" & vbCrLf
        sql &= " WHERE a.RID=@RID" & vbCrLf
        If RIDValue.Value.Length = 1 Then sql &= " AND a.PlanID = @PlanID" & vbCrLf
        sql &= " AND a.Years = @Years" & vbCrLf
        If v_rblFSQ1_S <> "00" AndAlso v_rblFSQ1_S <> "" Then sql &= " AND a.FSQ1 = @FSQ1" & vbCrLf
        Select Case v_rblSORT_TYPE1
            Case "2"
                sql &= " ORDER BY ROWNUM1" & vbCrLf
            Case Else '"1"
                sql &= " ORDER BY a.MODIFYDATE" & vbCrLf
        End Select

        Dim flag_chktest As Boolean = TIMS.sUtl_ChkTest()
        'TIMS.writeLog(Me, "##SD_14_018 sql:" & sql)
        Dim v_ddl_year As String = TIMS.GetListValue(ddl_year)
        Dim sCmd As New SqlCommand(sql, objconn)
        'Call TIMS.OpenDbConn(objconn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = RIDValue.Value
            If RIDValue.Value.Length = 1 Then
                .Parameters.Add("PlanID", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
            End If
            .Parameters.Add("Years", SqlDbType.VarChar).Value = v_ddl_year 'ddl_year.SelectedValue
            If v_rblFSQ1_S <> "00" AndAlso v_rblFSQ1_S <> "" Then
                .Parameters.Add("FSQ1", SqlDbType.VarChar).Value = v_rblFSQ1_S 'rblFSQ1_S.SelectedValue
            End If

            If (flag_chktest) Then
                Dim s_parms As String = TIMS.GetMyValue3(sCmd.Parameters)
                TIMS.WriteLog(Me, "##SD_14_018 sql:" & sql)
                TIMS.WriteLog(Me, "##SD_14_018 s_parms :" & s_parms)
            End If
            dt.Load(.ExecuteReader())
        End With

        BtnPrint1.Visible = False
        lMsg.Text = "查無資料!"
        DataGrid1.Visible = False
        If dt.Rows.Count > 0 Then
            BtnPrint1.Visible = True
            Select Case sm.UserInfo.LID
                Case 2
                    Update_SORTNO1(dt)
            End Select

            lMsg.Text = ""
            DataGrid1.Visible = True
            DataGrid1.DataSource = dt
            DataGrid1.DataBind()
        End If
    End Sub

    '更新sort資訊 (若為空，新增進去) 依 ROWNUM1
    Sub Update_SORTNO1(ByRef dt1 As DataTable)
        Dim s_sql As String = "SELECT 1 FROM ORG_MEMBER WHERE OMID=@OMID AND SORTNO1 IS NULL"
        Dim sCmd As New SqlCommand(s_sql, objconn)
        Dim u_sql As String = "UPDATE ORG_MEMBER SET SORTNO1=@SORTNO1 WHERE OMID=@OMID AND SORTNO1 IS NULL"
        Dim uCmd As New SqlCommand(u_sql, objconn)
        For Each dr1 As DataRow In dt1.Rows
            Dim dtP As New DataTable
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("OMID", SqlDbType.Int).Value = Val(dr1("OMID"))
                dtP.Load(.ExecuteReader())
            End With
            If dtP.Rows.Count > 0 Then
                '有資料執行update
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("SORTNO1", SqlDbType.Float).Value = CDbl(dr1("ROWNUM1"))
                    .Parameters.Add("OMID", SqlDbType.Int).Value = Val(dr1("OMID"))
                    .ExecuteNonQuery()
                End With
            End If
        Next
    End Sub

    '列印
    Sub pPrint1()
        Dim errMsg As String = ""

        Dim v_ddl_year As String = TIMS.GetListValue(ddl_year)
        Dim v_rblSORT_TYPE1 As String = TIMS.GetListValue(rblSORT_TYPE1)
        'Dim v_ddlFSQ1_A As String = TIMS.GetListValue(ddlFSQ1_A)
        '年度區間/申請階段  00/01/02/03  不區分/上半年/下半年/政策性產業 
        Dim v_rblFSQ1_S As String = TIMS.GetListValue(rblFSQ1_S)

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        If RIDValue.Value = "" Then errMsg &= "訓練機構尚未選取！\n"

        If v_ddl_year = "" Then errMsg &= "計畫年度尚未選取！\n"

        If v_rblFSQ1_S = "" Then errMsg &= "申請階段尚未選取！\n"
        If v_rblFSQ1_S = "00" Then errMsg &= "申請階段,不可選不區分！\n"

        If errMsg <> "" Then '錯誤
            Common.MessageBox(Me, errMsg)
            Exit Sub
        End If

        Dim vPlanID As String = Convert.ToString(sm.UserInfo.PlanID)
        If sm.UserInfo.RID = "A" AndAlso RIDValue.Value <> sm.UserInfo.RID Then
            '職訓局 選擇其他機構列印
            Dim parms As New Hashtable
            parms.Add("TPLANID", sm.UserInfo.TPlanID)
            parms.Add("RID", RIDValue.Value)
            parms.Add("YEARS", v_ddl_year)
            Dim sql As String = ""
            sql &= " SELECT RWPLANID FROM VIEW_RWPLANRID"
            sql &= " WHERE TPLANID=@TPLANID AND RID=@RID AND YEARS=@YEARS"
            Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
            If dr IsNot Nothing Then vPlanID = Convert.ToString(dr("rwPlanID"))
        End If

        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "Years=" & v_ddl_year 'ddl_year.SelectedValue
        MyValue &= "&RID=" & RIDValue.Value
        'MyValue &= "&PlanID=" & vPlanID  'Convert.ToString(sm.UserInfo.PlanID)
        MyValue &= "&FSQ1=" & v_rblFSQ1_S 'rblFSQ1_S.SelectedValue  'edit，by:20181114 (顯示)
        'MyValue &= "&FSQ2=" & v_rblFSQ1_S 'rblFSQ1_S.SelectedValue  'edit，by:20181114 (條件)
        Select Case v_rblSORT_TYPE1
            Case "2"
                MyValue &= "&SORT2=Y"
                'ORDER BY ROWNUM1" & vbCrLf
            Case Else '"1"
                MyValue &= "&SORT1=Y"
                'ORDER BY a.MODIFYDATE" & vbCrLf
        End Select
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub

    '新增/修改
    Sub sSaveData1()

        Dim iSql As String = ""
        iSql &= " INSERT INTO ORG_MEMBER(OMID ,RID ,PlanID ,Years ,SORTNO1" & vbCrLf
        iSql &= " ,Title1 ,CName ,Phone1 ,FSQ1 ,ModifyAcct ,ModifyDate )" & vbCrLf
        iSql &= " VALUES (@OMID ,@RID ,@PlanID ,@Years ,@SORTNO1" & vbCrLf
        iSql &= " ,@Title1 ,@CName ,@Phone1 ,@FSQ1 ,@ModifyAcct ,GETDATE() )" & vbCrLf

        Dim uSql As String = ""
        uSql &= " UPDATE ORG_MEMBER" & vbCrLf
        uSql &= " SET SORTNO1=@SORTNO1" & vbCrLf
        uSql &= " ,Title1 = @Title1" & vbCrLf
        uSql &= " ,CName = @CName" & vbCrLf
        uSql &= " ,Phone1 = @Phone1" & vbCrLf
        uSql &= " ,FSQ1 = @FSQ1" & vbCrLf
        uSql &= " ,LOCK1 = NULL" & vbCrLf
        uSql &= " ,ModifyAcct = @ModifyAcct" & vbCrLf
        uSql &= " ,ModifyDate = GETDATE()" & vbCrLf
        uSql &= " WHERE OMID = @OMID" & vbCrLf

        tROWNUM1.Text = TIMS.CDBL1(tROWNUM1.Text)
        Dim v_ddl_year As String = TIMS.GetListValue(ddl_year)
        Dim v_ddlFSQ1_A As String = TIMS.GetListValue(ddlFSQ1_A)
        If hidsOMID.Value = "" Then
            '新增
            Dim iOMID As Integer = DbAccess.GetNewId(objconn, "ORG_MEMBER_OMID_SEQ,ORG_MEMBER,OMID")
            Dim myParam As Hashtable = New Hashtable
            myParam.Add("OMID", iOMID)
            myParam.Add("RID", RIDValue.Value)
            myParam.Add("PlanID", sm.UserInfo.PlanID)
            myParam.Add("Years", v_ddl_year) 'ddl_year.SelectedValue)
            myParam.Add("SORTNO1", Val(tROWNUM1.Text))
            myParam.Add("Title1", tTitle.Text)
            myParam.Add("CName", tName.Text)
            myParam.Add("Phone1", tPhone.Text)
            myParam.Add("FSQ1", v_ddlFSQ1_A) 'ddlFSQ1_A.SelectedValue)

            myParam.Add("ModifyAcct", sm.UserInfo.UserID)
            DbAccess.ExecuteNonQuery(iSql, objconn, myParam)

        Else
            '修改
            Dim myParam As Hashtable = New Hashtable
            myParam.Add("SORTNO1", Val(tROWNUM1.Text))
            myParam.Add("Title1", tTitle.Text)
            myParam.Add("CName", tName.Text)
            myParam.Add("Phone1", tPhone.Text)
            myParam.Add("FSQ1", v_ddlFSQ1_A) 'ddlFSQ1_A.SelectedValue)

            myParam.Add("ModifyAcct", sm.UserInfo.UserID)
            myParam.Add("OMID", Val(hidsOMID.Value))
            DbAccess.ExecuteNonQuery(uSql, objconn, myParam)

        End If

        Call Search1()
    End Sub

    Function GET_DATAROW1(ByVal vOMID As String) As DataRow
        Dim dr1 As DataRow = Nothing
        'hidsOMID.Value = ""
        Dim sql As String = ""
        sql = " SELECT * FROM ORG_MEMBER WHERE OMID = @OMID" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)
        'Call TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OMID", SqlDbType.VarChar).Value = vOMID
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count > 0 Then dr1 = dt1.Rows(0)
        Return dr1
    End Function

    '修改
    Sub Utl_EDITU1(ByVal vOMID As String, ByVal vROWNUM1 As String)
        hidsOMID.Value = ""
        Dim dr1 As DataRow = GET_DATAROW1(vOMID)
        If dr1 Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        hidsOMID.Value = vOMID 'Convert.ToString(e.CommandArgument)
        tROWNUM1.Text = vROWNUM1
        tTitle.Text = Convert.ToString(dr1("Title1"))
        tName.Text = Convert.ToString(dr1("CName"))
        tPhone.Text = Convert.ToString(dr1("Phone1"))
        Common.SetListItem(ddlFSQ1_A, Convert.ToString(dr1("FSQ1")))

        btnAdd.Visible = False
        btnUPD.Visible = True
        btnCancel.Visible = True
        Utl_Lockdata2(cst_func_lock1)
        'btnAdd.Text = "儲存"
    End Sub

    'Utl_EDITU2 '修改-僅可修改電話
    Sub Utl_EDITU2(ByVal vOMID As String, ByVal vROWNUM1 As String)
        hidsOMID.Value = ""
        Dim dr1 As DataRow = GET_DATAROW1(vOMID)
        If dr1 Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        hidsOMID.Value = vOMID 'Convert.ToString(e.CommandArgument)
        tROWNUM1.Text = vROWNUM1
        tTitle.Text = Convert.ToString(dr1("Title1"))
        tName.Text = Convert.ToString(dr1("CName"))
        tPhone.Text = Convert.ToString(dr1("Phone1"))
        Common.SetListItem(ddlFSQ1_A, Convert.ToString(dr1("FSQ1")))

        btnAdd.Visible = False
        btnUPD.Visible = True
        btnCancel.Visible = True
        Utl_Lockdata2(cst_func_lock2)
        'btnAdd.Text = "儲存"
    End Sub

    '刪除
    Sub Utl_DELETE1(ByVal vOMID As String)
        Dim sql As String = ""
        sql = " DELETE ORG_MEMBER WHERE OMID = @OMID" & vbCrLf
        Dim myParam As Hashtable = New Hashtable
        myParam.Add("OMID", vOMID)
        DbAccess.ExecuteNonQuery(sql, objconn, myParam)
    End Sub

    '解鎖
    Sub Utl_UNLOCK(ByVal vOMID As String)
        Dim dr1 As DataRow = GET_DATAROW1(vOMID)
        If dr1 Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        'myParam.Clear()
        Dim myParam As New Hashtable From {
            {"MODIFYACCT", sm.UserInfo.UserID},
            {"OMID", Val(vOMID)}
        } 'myParam = New Hashtable
        Dim u_sql As String = ""
        u_sql &= " UPDATE ORG_MEMBER" & vbCrLf
        u_sql &= " SET LOCK1 = 'N'" & vbCrLf
        u_sql &= " ,MODIFYACCT = @MODIFYACCT" & vbCrLf
        u_sql &= " ,MODIFYDATE = GETDATE()" & vbCrLf
        u_sql &= " WHERE OMID = @OMID" & vbCrLf
        DbAccess.ExecuteNonQuery(u_sql, objconn, myParam)

    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        tROWNUM1.Text = TIMS.ClearSQM(tROWNUM1.Text)
        tTitle.Text = TIMS.ClearSQM(tTitle.Text)
        tName.Text = TIMS.ClearSQM(tName.Text)
        tPhone.Text = TIMS.ClearSQM(tPhone.Text)

        Dim v_ddl_year As String = TIMS.GetListValue(ddl_year)
        Dim v_ddlFSQ1_A As String = TIMS.GetListValue(ddlFSQ1_A)

        If RIDValue.Value = "" Then Errmsg &= "訓練機構尚未選取！" & vbCrLf

        If v_ddl_year = "" Then Errmsg &= "計畫年度尚未選取！" & vbCrLf

        If tROWNUM1.Text = "" Then Errmsg &= "請填寫排序" & vbCrLf

        If tTitle.Text = "" Then Errmsg &= "請填寫職稱" & vbCrLf
        If tName.Text = "" Then Errmsg &= "請填寫姓名" & vbCrLf
        If tPhone.Text = "" Then Errmsg &= "請填寫聯絡電話" & vbCrLf
        'If v_ddlFSQ1_A = "" Then Errmsg &= "年度別尚未選取！" & vbCrLf
        If v_ddlFSQ1_A = "" Then Errmsg &= "請選擇申請階段" & vbCrLf
        If v_ddlFSQ1_A = "00" Then Errmsg &= "申請階段,參數有誤" & vbCrLf

        If Errmsg <> "" Then Return False

        Dim drRID As DataRow = TIMS.Get_RID_DR(RIDValue.Value, objconn)
        If drRID IsNot Nothing Then
            If Convert.ToString(drRID("years")) <> "" AndAlso Convert.ToString(drRID("years")) <> v_ddl_year Then
                Errmsg &= "訓練機構 業務年度 選擇有誤!" & vbCrLf
            End If
            Select Case sm.UserInfo.LID
                Case 0
                Case 1
                    If Convert.ToString(drRID("planid")) <> "0" AndAlso Convert.ToString(drRID("planid")) <> sm.UserInfo.PlanID Then
                        Errmsg &= "業務計畫 與登入計畫不同 選擇有誤!(分署)" & vbCrLf
                    End If
                Case 2
                    If Convert.ToString(drRID("planid")) <> sm.UserInfo.PlanID Then
                        Errmsg &= "訓練機構 業務計畫 與登入計畫不同 選擇有誤!!(訓練機構)" & vbCrLf
                    End If
            End Select
        End If

        If Not TIMS.IsNumeric1(tROWNUM1.Text) Then
            Errmsg &= "排序，請填寫數字!" & vbCrLf
        End If
        If Errmsg <> "" Then Return False

        'If Not tTitle.Enabled Then Return True '修改狀態下不檢核資料
        Dim dr1 As DataRow = Nothing
        hidsOMID.Value = TIMS.ClearSQM(hidsOMID.Value)
        If hidsOMID.Value <> "" Then dr1 = GET_DATAROW1(hidsOMID.Value)

        Dim sql As String = ""
        sql &= " SELECT 'x' x" & vbCrLf
        sql &= " FROM ORG_MEMBER a" & vbCrLf
        sql &= " WHERE a.RID = @RID" & vbCrLf
        'sql &= " AND PlanID = @PlanID" & vbCrLf
        sql &= " AND a.Years = @Years" & vbCrLf
        sql &= " AND a.Title1 = @Title1" & vbCrLf
        sql &= " AND a.CName = @CName" & vbCrLf
        If v_ddlFSQ1_A <> "00" AndAlso v_ddlFSQ1_A <> "" Then
            sql &= " AND a.FSQ1 = @FSQ1 "
            '.Parameters.Add("@FSQ1", SqlDbType.VarChar).Value = v_ddlFSQ1_A
        Else
            If dr1 IsNot Nothing Then
                sql &= " AND a.FSQ1 = @FSQ1 "
                '.Parameters.Add("@FSQ1", SqlDbType.VarChar).Value = Convert.ToString(dr1("FSQ1"))
            End If
        End If
        If hidsOMID.Value <> "" Then
            sql &= " AND a.OMID != @OMID" & vbCrLf
            'da.SelectCommand.Parameters.Add("@OMID", SqlDbType.VarChar).Value = hidsOMID.Value
        End If
        'Select Case v_ddlFSQ1_A'ddlFSQ1_A.SelectedValue
        '    Case "00"
        '    Case cst_FSQ1_上半年
        '        sql &= " AND FSQ1 = '01' "
        '    Case cst_FSQ1_下半年
        '        sql &= " AND FSQ1 = '02' "
        'End Select

        Dim sCmd As New SqlCommand(sql, objconn)
        '依 RIDValue.Value 重新取得 orgid
        'Call TIMS.OpenDbConn(objconn)
        Dim dt1 As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("@RID", SqlDbType.VarChar).Value = RIDValue.Value
            '.Parameters.Add("@PlanID", SqlDbType.VarChar).Value = sm.UserInfo.PlanID
            .Parameters.Add("@Years", SqlDbType.VarChar).Value = v_ddl_year ' ddl_year.SelectedValue
            .Parameters.Add("@Title1", SqlDbType.NVarChar).Value = tTitle.Text
            .Parameters.Add("@CName", SqlDbType.NVarChar).Value = tName.Text
            If hidsOMID.Value <> "" Then
                'sql += " and OMID!=@OMID" & vbCrLf
                .Parameters.Add("@OMID", SqlDbType.VarChar).Value = hidsOMID.Value
            End If
            If v_ddlFSQ1_A <> "00" AndAlso v_ddlFSQ1_A <> "" Then
                'sql &= " AND a.FSQ1 = @FSQ1 "
                .Parameters.Add("@FSQ1", SqlDbType.VarChar).Value = v_ddlFSQ1_A
            Else
                If dr1 IsNot Nothing Then
                    'sql &= " AND a.FSQ1 = @FSQ1 "
                    .Parameters.Add("@FSQ1", SqlDbType.VarChar).Value = Convert.ToString(dr1("FSQ1"))
                End If
            End If
            dt1.Load(.ExecuteReader())
        End With
        If dt1.Rows.Count > 0 Then
            Errmsg &= "該單位、計畫年度已經有同職稱、姓名資料" & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢
    Protected Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        btnAdd.Visible = True
        btnUPD.Visible = False
        btnCancel.Visible = False
        BtnPrint1.Visible = False

        Call Search1()
    End Sub

    '新增 / 儲存
    Private Sub btnAdd_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAdd.Click, btnUPD.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        Call sSaveData1()
    End Sub

    '列印
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        Call pPrint1()
    End Sub

    '取消
    Protected Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        ClearList1()
    End Sub

End Class