Partial Class SD_05_002
    Inherits AuthBasePage

    Const cst_msgNG1 As String = "請先完成訓期已滿1/2學員離退訓作業!"
    Dim CPdt As DataTable
    'Dim FunDr As DataRow
    Dim Days1 As Integer
    Dim Days2 As Integer

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False
    Dim flag_File1_csv As Boolean = False

    'Stud_CardRecord
    'xup KEY_LEAVE set nouse =null where leaveid in ('06','09');//update
    'SELECT LEAVEID,NAME,LEAVESORT,MINUSPOINT FROM KEY_LEAVE WHERE NOUSE IS NULL ORDER BY LEAVESORT
    'SD_05_002_edit.aspx
    'SD_05_002_add.aspx
    'SD_05_002_Wrong.aspx
    'Dim flagYear2017 As Boolean = False
    Dim dtLEAVE As DataTable
    'Dim oDataGridX As DataGrid = Nothing
    Const cst_DG1_班別 As Integer = 2
    Const cst_DG1_學號 As Integer = 3

    'Dim au As New cAUTH
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)

        'oDataGridX = DataGrid1
        dtLEAVE = TIMS.Get_dtLEAVE(objconn)
        'Call CREATE_DG1(DataGrid1, dtLEAVE)
        PageControler1.PageDataGrid = DataGrid1

        '取出設定天數檔 Start
        Call TIMS.Get_SysDays(Days1, Days2, objconn)
        '取出設定天數檔 End

        'Dim flagYear2017 As Boolean = False
        'flagYear2017 = TIMS.Get_UseLEAVE_2017(Me)

        '分頁設定 Start
        'DataGrid1B.Visible = False
        'oDataGridX = DataGrid1
        'If flagYear2017 Then oDataGridX = DataGrid1B
        'oDataGridX.Visible = True
        '分頁設定 End

        If Not IsPostBack Then
            Call cCreate1()
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');ShowFrame();"
            HistoryTable.Attributes("onclick") = "ShowFrame();"
            OCID1.Style("CURSOR") = "hand"
        End If
        Button6.Attributes("onclick") = String.Concat(TIMS.Get_javascript_openOrg_js(sm), "SetOneOCID();")

        '檢查帳號的功能權限-----------------------------------Start
        'Button1.Enabled = False '查詢
        'If au.blnCanSech Then Button1.Enabled = True
        'Button2.Enabled = False '新增
        'If au.blnCanAdds Then Button2.Enabled = True
        '檢查帳號的功能權限-----------------------------------End
    End Sub

    Sub cCreate1()
        'DataGrid1.Visible = False
        'oDataGridX.Visible = False
        Dim dtLeaveF As DataTable = TIMS.GET_LEAVEdt(TIMS.cst_sex_F, objconn)
        Dim dtLeaveM As DataTable = TIMS.GET_LEAVEdt(TIMS.cst_sex_M, objconn)
        msg.Text = ""
        Table4.Visible = False
        LeaveID = TIMS.GET_LEAVE(LeaveID, dtLeaveF, dtLeaveM, TIMS.cst_sex_F, 1)
        'LeaveID.Items.Insert(LeaveID.Items.Count, New ListItem("全勤", "99"))
        RIDValue.Value = sm.UserInfo.RID
        center.Text = sm.UserInfo.OrgName

        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)

        Call UseKeepSearchStr()

        Button1.Attributes("onclick") = "javascript:return search()"
        Button2.Attributes("onclick") = "javascript:return schocid()"
    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        If start_date.Text.Trim <> "" Then
            Try
                If IsDate(start_date.Text) Then
                    start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
                Else
                    Errmsg += "請假期間開始日期有誤，應為日期格式" & start_date.Text & vbCrLf
                End If
            Catch ex As Exception
                Errmsg += "請假期間開始日期有誤，應為日期格式" & start_date.Text & vbCrLf
            End Try
        End If
        If end_date.Text.Trim <> "" Then
            Try
                If IsDate(start_date.Text) Then
                    end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
                Else
                    Errmsg += "請假期間結束日期有誤，應為日期格式" & end_date.Text & vbCrLf
                End If
            Catch ex As Exception
                Errmsg += "請假期間結束日期有誤，應為日期格式" & end_date.Text & vbCrLf
            End Try
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Sub search1()
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Datagrid2.Visible = False
        Button7.Visible = False

        StdName.Text = TIMS.ClearSQM(StdName.Text)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        cjobValue.Value = TIMS.ClearSQM(cjobValue.Value)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WS1 AS (" & vbCrLf

        sql &= " SELECT cs.SOCID" & vbCrLf
        sql &= " ,cs.STUDSTATUS" & vbCrLf
        sql &= " ,cs.STUDENTID" & vbCrLf
        sql &= " ,cs.STUDID2" & vbCrLf
        sql &= " ,cs.NAME" & vbCrLf
        sql &= " ,cs.OCID" & vbCrLf
        sql &= " ,cs.YEARS" & vbCrLf
        sql &= " ,cs.CLASSCNAME" & vbCrLf
        sql &= " ,cs.CYCLTYPE" & vbCrLf
        sql &= " ,cs.STDATE" & vbCrLf
        sql &= " ,cs.FTDATE" & vbCrLf
        sql &= " ,cs.RID,cs.TPLANID,cs.PLANID" & vbCrLf
        sql &= " ,cs.CLASSID,cs.CJOB_UNKEY" & vbCrLf
        sql &= " FROM V_STUDENTINFO cs" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then
            sql &= " AND cs.TPLANID='" & sm.UserInfo.TPlanID & "'" & vbCrLf
            sql &= " AND cs.YEARS='" & sm.UserInfo.Years & "'" & vbCrLf
        Else
            sql &= " and cs.PLANID='" & sm.UserInfo.PlanID & "'" & vbCrLf
        End If

        If StdName.Text <> "" Then
            'Sql &= " and ss.Name like '%'+'" & StdName.Text.Replace("'", "''") & "'+'%'" & vbCrLf
            sql &= " and cs.Name like '%" & StdName.Text & "%'" & vbCrLf 'fix ORA-01722: invalid number
        End If
        If OCIDValue1.Value <> "" Then
            sql &= " and cs.OCID='" & OCIDValue1.Value & "'" & vbCrLf
        End If
        If RIDValue.Value <> "" Then
            sql &= " and cs.RID='" & RIDValue.Value & "'" & vbCrLf
        Else
            sql &= " and cs.RID='" & sm.UserInfo.RID & "'" & vbCrLf
        End If
        If cjobValue.Value <> "" Then
            sql &= " and cs.CJOB_UNKEY='" & cjobValue.Value & "'" & vbCrLf
        End If
        sql &= " )" & vbCrLf

        sql &= " ,WS2 AS (" & vbCrLf
        sql &= " SELECT a.SOCID, a.DaSource" & vbCrLf
        sql &= " ,MAX(a.LeaveDate) LastDate" & vbCrLf
        Dim s_OTHER_LEAVE As String = ""
        For Each dr1 As DataRow In dtLEAVE.Rows
            If s_OTHER_LEAVE <> "" Then s_OTHER_LEAVE &= ","
            s_OTHER_LEAVE &= Convert.ToString(dr1("LEAVEID"))
            sql &= String.Format(" ,SUM(CASE a.LeaveID WHEN '{0}' THEN a.Hours ELSE 0 END) {1}", dr1("LEAVEID"), dr1("ENGNAME")) & vbCrLf '病假
        Next
        Dim s_OTHER_LEAVEin As String = TIMS.CombiSQLIN(s_OTHER_LEAVE)

        'sql &= " ,SUM(CASE WHEN a.LeaveID = '01' THEN a.Hours ELSE 0 END) Sick" & vbCrLf '病假
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '02' THEN a.Hours ELSE 0 END) Reason" & vbCrLf '事假
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '03' THEN a.Hours ELSE 0 END) Publics" & vbCrLf '公假
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '04' THEN a.Hours ELSE 0 END) Skips" & vbCrLf '曠課
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '05' THEN a.Hours ELSE 0 END) Dead" & vbCrLf '喪假
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '06' THEN a.Hours ELSE 0 END) Late" & vbCrLf '遲到
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '07' THEN a.Hours ELSE 0 END) Marry" & vbCrLf '婚假
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '08' THEN a.Hours ELSE 0 END) Birth" & vbCrLf '陪產假
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '09' THEN a.Hours ELSE 0 END) notPunch" & vbCrLf '未打卡
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '10' THEN a.Hours ELSE 0 END) Weekly" & vbCrLf '集會(週會)
        'sql &= " ,SUM(CASE WHEN a.LeaveID = '11' THEN a.Hours ELSE 0 END) Health" & vbCrLf '生理假
        'If flagYear2017 Then
        '    '03','05','11','01','02','04'
        '    'Publics,Dead,Health,Sick,Reason,Skips
        '    sql &= " ,SUM(CASE WHEN a.LeaveID NOT IN ('03','05','11','01','02','04') THEN a.Hours ELSE 0 END) other" & vbCrLf
        'Else
        '    sql &= " ,SUM(CASE WHEN a.LeaveID NOT IN ('01','02','03','04','05','06','07','08') THEN a.Hours ELSE 0 END) other" & vbCrLf
        'End If
        sql &= String.Format(" ,SUM(CASE WHEN a.LeaveID NOT IN ({0}) THEN a.Hours ELSE 0 END) other", s_OTHER_LEAVEin) & vbCrLf
        'sql &= " ,SUM(CASE WHEN a.LeaveID NOT IN ('01','02','03','04','05','06','07','08') THEN a.Hours ELSE 0 END) other" & vbCrLf
        sql &= " FROM STUD_TURNOUT a" & vbCrLf
        sql &= " JOIN WS1 s on s.socid =a.socid" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        If start_date.Text <> "" Then
            sql &= " and a.LeaveDate >= " & TIMS.To_date(start_date.Text) & vbCrLf
        End If
        If end_date.Text <> "" Then
            sql &= " and a.LeaveDate <= " & TIMS.To_date(end_date.Text) & vbCrLf
        End If
        If LeaveID.SelectedValue <> "" Then
            sql &= " and a.LeaveID = '" & LeaveID.SelectedValue & "'" & vbCrLf
        End If
        sql &= " Group By a.SOCID, a.DaSource" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " SELECT a.SOCID" & vbCrLf
        sql &= " ,a.StudStatus" & vbCrLf
        sql &= " ,a.StudentID,a.STUDID2" & vbCrLf
        sql &= " ,a.Name" & vbCrLf
        sql &= " ,a.OCID" & vbCrLf
        sql &= " ,a.Years" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(a.CLASSCNAME,a.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,a.CyclType" & vbCrLf
        'sql &= " ,a.LevelType" & vbCrLf
        sql &= " ,a.FTDate" & vbCrLf
        'sql &= " ,a.IsClosed" & vbCrLf
        sql &= " ,a.RID" & vbCrLf
        sql &= " ,a.ClassID" & vbCrLf
        '資料來源 null:1:人員填寫 2:線上請假 3:刷卡紀錄。
        sql &= " ,case ISNULL(g.DaSource,'1') when '1' then '人員填寫'" & vbCrLf
        sql &= "  when '2' then '線上請假'" & vbCrLf
        sql &= "  when '3' then '刷卡紀錄' end DaSource" & vbCrLf
        sql &= " ,FORMAT(g.LastDate,'yyyy/MM/dd') LastDate" & vbCrLf
        For Each dr1 As DataRow In dtLEAVE.Rows
            sql &= String.Format(" ,ISNULL(g.{0},0) {0}", dr1("ENGNAME")) & vbCrLf
        Next
        'sql &= " ,ISNULL(g.Reason,0) Reason" & vbCrLf
        'sql &= " ,ISNULL(g.Publics,0) Publics" & vbCrLf
        'sql &= " ,ISNULL(g.Skips,0) Skips" & vbCrLf
        'sql &= " ,ISNULL(g.Dead,0) Dead" & vbCrLf
        'sql &= " ,ISNULL(g.Late,0) Late" & vbCrLf
        'sql &= " ,ISNULL(g.Marry,0) Marry" & vbCrLf
        'sql &= " ,ISNULL(g.Birth,0) Birth" & vbCrLf
        'sql &= " ,ISNULL(g.notPunch,0) notPunch" & vbCrLf
        'sql &= " ,ISNULL(g.Health,0) Health" & vbCrLf

        sql &= " ,ISNULL(g.other,0) other" & vbCrLf
        sql &= " FROM WS1 a" & vbCrLf
        sql &= " JOIN WS2 g on g.socid =a.socid" & vbCrLf
        sql &= " WHERE 1=1"
        Dim flag_LeaveID_Use As Boolean = False
        Dim v_LeaveID As String = TIMS.GetListValue(LeaveID)
        For Each dr1 As DataRow In dtLEAVE.Rows
            If v_LeaveID = Convert.ToString(dr1("LEAVEID")) Then
                sql &= String.Format("and ISNULL(g.{0},0) > 0", dr1("ENGNAME")) & vbCrLf
                flag_LeaveID_Use = True
            End If
            If flag_LeaveID_Use Then Exit For
        Next
        If Not flag_LeaveID_Use AndAlso v_LeaveID <> "" Then
            '(有值才查詢其他)
            sql &= " and ISNULL(g.other,0) > 0 " & vbCrLf
        End If

        'Select Case LeaveID.SelectedValue
        '    Case "01"
        '        sql &= " and ISNULL(g.Sick,0) > 0 " & vbCrLf
        '    Case "02"
        '        sql &= " and ISNULL(g.Reason,0) > 0 " & vbCrLf
        '    Case "03"
        '        sql &= " and ISNULL(g.Publics,0) > 0 " & vbCrLf
        '    Case "04"
        '        sql &= " and ISNULL(g.Skips,0) > 0 " & vbCrLf
        '    Case "05"
        '        sql &= " and ISNULL(g.Dead,0) > 0 " & vbCrLf
        '    Case "06"
        '        sql &= " and ISNULL(g.Late,0) > 0 " & vbCrLf
        '    Case "07"
        '        sql &= " and ISNULL(g.Marry,0) > 0 " & vbCrLf
        '    Case "08"
        '        sql &= " and ISNULL(g.Birth,0) > 0 " & vbCrLf
        '    Case "11"
        '        sql &= " and ISNULL(g.Health,0) > 0 " & vbCrLf
        '    Case Else
        '        If LeaveID.SelectedValue <> "" Then '(有值才查詢其他)
        '            sql &= " and ISNULL(g.other,0) > 0 " & vbCrLf
        '        End If
        'End Select
        sql &= " order by a.ClassID,a.CyclType,a.StudentID " & vbCrLf

        Dim dt As DataTable
        dt = DbAccess.GetDataTable(sql, objconn)
        CPdt = dt.Copy()

        Table4.Visible = False
        PageControler1.Visible = False
        msg.Text = "查無資料!"
        If dt.Rows.Count = 0 Then Return 'EXIT

        Table4.Visible = True
        PageControler1.Visible = True

        msg.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.Sort = "ClassID,CyclType,StudentID"
        PageControler1.ControlerLoad()

        DataGrid1.Columns(cst_DG1_班別).Visible = If(OCID1.Text = "", True, False)
        'If OCID1.Text = "" Then oDataGridX.Columns(cst_班別).Visible = True
    End Sub

    '查詢學員鈕
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call search1()
    End Sub

    Private Sub DataGridX_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        ', DataGrid1B.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                Dim drv As DataRowView = e.Item.DataItem
                '序號
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
                'e.Item.Cells(cst_學號).Text = Right(e.Item.Cells(cst_學號).Text, 2)

                Dim v_LeaveID As String = TIMS.GetListValue(LeaveID)
                Dim sCmdArg As String = ""
                sCmdArg = ""
                TIMS.SetMyValue(sCmdArg, "SOCID", Convert.ToString(drv("SOCID")))
                TIMS.SetMyValue(sCmdArg, "OCID", Convert.ToString(drv("OCID")))
                TIMS.SetMyValue(sCmdArg, "LeaveID", v_LeaveID)

                Dim Button3 As Button = e.Item.FindControl("Button3") '修改
                Dim Button5 As Button = e.Item.FindControl("Button5") '查看
                Button3.CommandName = "edit"
                Button3.CommandArgument = sCmdArg & "&Proecess=edit"

                'Button3.Enabled = True
                'If Not au.blnCanMod Then
                '    Button3.Enabled = False
                '    TIMS.Tooltip(Button3, "無此權限。")
                'End If

                Button5.CommandName = "view"
                Button5.CommandArgument = sCmdArg & "&Proecess=view"

                Select Case drv("StudStatus").ToString
                    Case "1", "4" '在、續訓
                        If sm.UserInfo.RoleID <= 1 AndAlso sm.UserInfo.LID <= 1 Then
                            Button3.Visible = True '修改
                            Button5.Visible = True '查看
                        Else
                            Button3.Visible = True '修改
                            TIMS.Tooltip(Button3, "學員尚未結訓!!可修改。")
                            Button5.Visible = False '查看
                        End If

                    Case "2", "3" '離、退
                        If sm.UserInfo.RoleID <= 1 AndAlso sm.UserInfo.LID <= 1 Then
                            Button3.Visible = True '修改
                            Button5.Visible = True '查看
                        Else
                            Button3.Visible = False '修改
                            Button5.Visible = True '查看
                            TIMS.Tooltip(Button5, "學員已經離退訓!!僅供查看。")
                        End If

                    Case "5" '結訓
                        Select Case Val(sm.UserInfo.RoleID)
                            Case 0, 1
                                'Const cst_i75 As Integer = 75
                                '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成75天
                                '暫時先改這樣,以後還會再改
                                'If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                '    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) >= cst_i75 Then
                                '        Button3.Visible = False
                                '        Button5.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫超過 " & cst_i75 & "天!!僅供查看。")
                                '    Else
                                '        Button3.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫 " & cst_i75 & "天內!!可修改。")
                                '        Button5.Visible = False
                                '    End If
                                'Else
                                '    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) >= Days2 Then
                                '        Button3.Visible = False
                                '        Button5.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫超過 " & Days2 & "天!!僅供查看。")
                                '    Else
                                '        Button3.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫 " & Days2 & "天內!!可修改。")
                                '        Button5.Visible = False
                                '    End If
                                'End If
                                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) >= Days2 Then
                                    Button3.Visible = False
                                    Button5.Visible = True
                                    TIMS.Tooltip(Button5, "結訓後,該計畫超過 " & Days2 & "天!!僅供查看。")
                                Else
                                    Button3.Visible = True
                                    TIMS.Tooltip(Button5, "結訓後,該計畫 " & Days2 & "天內!!可修改。")
                                    Button5.Visible = False
                                End If

                            Case Else
                                'Const cst_i60 As Integer = 60
                                '判斷計畫是否為補助辦理保母職業訓練(46)與辦理照顧服務員職業訓練(47)時,限製天數改成60天
                                '暫時先改這樣,以後還會再改
                                'If sm.UserInfo.TPlanID = 46 Or sm.UserInfo.TPlanID = 47 Then
                                '    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) >= cst_i60 Then
                                '        Button3.Visible = False
                                '        Button5.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫超過 " & cst_i60 & "天!!僅供查看。")
                                '    Else
                                '        Button3.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫 " & cst_i60 & "天內!!可修改。")
                                '        Button5.Visible = False
                                '    End If
                                'Else
                                '    If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) >= Days1 Then
                                '        Button3.Visible = False
                                '        Button5.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫超過 " & Days1 & "天!!僅供查看。")
                                '    Else
                                '        Button3.Visible = True
                                '        TIMS.Tooltip(Button5, "結訓後,該計畫 " & Days1 & "天內!!可修改。")
                                '        Button5.Visible = False
                                '    End If
                                'End If

                                If DateDiff(DateInterval.Day, drv("FTDate"), Now.Date) >= Days1 Then
                                    Button3.Visible = False
                                    Button5.Visible = True
                                    TIMS.Tooltip(Button5, "結訓後,該計畫超過 " & Days1 & "天!!僅供查看。")
                                Else
                                    Button3.Visible = True
                                    TIMS.Tooltip(Button5, "結訓後,該計畫 " & Days1 & "天內!!可修改。")
                                    Button5.Visible = False
                                End If

                        End Select
                End Select
        End Select

    End Sub

    Private Sub DataGridX_ItemCommand(ByVal source As Object, ByVal e As System.Web.UI.WebControls.DataGridCommandEventArgs) Handles DataGrid1.ItemCommand
        ', DataGrid1B.ItemCommand
        msg.Text = ""
        If e.CommandName = "" Then Exit Sub
        If e.CommandArgument = "" Then Exit Sub

        Select Case e.CommandName
            Case "edit"
                'https://jira.turbotech.com.tw/browse/TIMSC-246
                Dim OCID_1 As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
                If TIMS.Chk_SELRESULTBLIDET(OCID_1, objconn) Then
                    msg.Text = cst_msgNG1
                    Common.MessageBox(Me, cst_msgNG1)
                    Exit Sub
                End If

                KeepSearchStr()
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "SD_05_002_edit.aspx?ID=" & Request("ID") & "&" & e.CommandArgument)

            Case "view"
                'https://jira.turbotech.com.tw/browse/TIMSC-246
                Dim OCID_1 As String = TIMS.GetMyValue(e.CommandArgument, "OCID")
                If TIMS.Chk_SELRESULTBLIDET(OCID_1, objconn) Then
                    msg.Text = cst_msgNG1
                    Common.MessageBox(Me, cst_msgNG1)
                    Exit Sub
                End If

                KeepSearchStr()
                Call TIMS.CloseDbConn(objconn)
                TIMS.Utl_Redirect1(Me, "SD_05_002_edit.aspx?ID=" & Request("ID") & "&" & e.CommandArgument)
        End Select

    End Sub

    '新增 鈕
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.RespWrite(Me, "<script>alert('" & "請選取班級資料!!!" & "');</script>")
            Exit Sub
        End If
        'https://jira.turbotech.com.tw/browse/TIMSC-246
        If TIMS.Chk_SELRESULTBLIDET(OCIDValue1.Value, objconn) Then
            Common.MessageBox(Me, cst_msgNG1)
            Exit Sub
        End If
        Call KeepSearchStr()
        'Response.Redirect("SD_05_002_add.aspx?ID=" & Request("ID") & "&Proecess=add")

        Dim url1 As String = "SD_05_002_add.aspx?ID=" & TIMS.ClearSQM(Request("ID"))
        url1 &= "&Proecess=add"
        url1 &= "&OCID1=" & OCIDValue1.Value
        Call TIMS.Utl_Redirect(Me, objconn, url1)

        'End If
    End Sub

    '設定 Session("SearchStr") 
    Sub KeepSearchStr()
        Dim s_SearchStr As String = ""
        s_SearchStr = "prg=SD_05_002"
        s_SearchStr += "&center=" & center.Text & "&RIDValue=" & RIDValue.Value
        s_SearchStr += "&TMID1=" & TMID1.Text & "&TMIDValue1=" & TMIDValue1.Value
        s_SearchStr += "&OCID1=" & OCID1.Text & "&OCIDValue1=" & OCIDValue1.Value
        s_SearchStr += "&start_date=" & start_date.Text
        s_SearchStr += "&end_date=" & end_date.Text
        s_SearchStr += "&LeaveID=" & LeaveID.SelectedValue
        s_SearchStr += "&StdName=" & StdName.Text
        s_SearchStr += "&PageIndex=" & DataGrid1.CurrentPageIndex + 1
        s_SearchStr += "&ShowTable=" & If(Table4.Visible, "true", "false")
        Session("SearchStr") = s_SearchStr
    End Sub

    Sub UseKeepSearchStr()
        'Session("SearchStr") = "prg=SD_05_002"
        'Session("SearchStr") += "&center=" & center.Text & "&RIDValue=" & RIDValue.Value
        'Session("SearchStr") += "&TMID1=" & TMID1.Text & "&TMIDValue1=" & TMIDValue1.Value
        'Session("SearchStr") += "&OCID1=" & OCID1.Text & "&OCIDValue1=" & OCIDValue1.Value
        'Session("SearchStr") += "&start_date=" & start_date.Text
        'Session("SearchStr") += "&end_date=" & end_date.Text
        'Session("SearchStr") += "&LeaveID=" & LeaveID.SelectedValue
        'Session("SearchStr") += "&StdName=" & StdName.Text
        'Session("SearchStr") += "&PageIndex=" & oDataGridX.CurrentPageIndex + 1
        'Session("SearchStr") += "&ShowTable=" & Table4.Visible

        If Session("SearchStr") Is Nothing Then Return
        Dim s_SearchStr As String = Convert.ToString(Session("SearchStr"))
        Session("SearchStr") = Nothing
        If TIMS.GetMyValue(s_SearchStr, "prg") <> "SD_05_002" Then Return
        center.Text = TIMS.GetMyValue(s_SearchStr, "center")
        RIDValue.Value = TIMS.GetMyValue(s_SearchStr, "RIDValue")
        TMID1.Text = TIMS.GetMyValue(s_SearchStr, "TMID1")
        TMIDValue1.Value = TIMS.GetMyValue(s_SearchStr, "TMIDValue1")
        OCID1.Text = TIMS.GetMyValue(s_SearchStr, "OCID1")
        OCIDValue1.Value = TIMS.GetMyValue(s_SearchStr, "OCIDValue1")
        start_date.Text = TIMS.GetMyValue(s_SearchStr, "start_date")
        end_date.Text = TIMS.GetMyValue(s_SearchStr, "end_date")
        Common.SetListItem(LeaveID, TIMS.GetMyValue(s_SearchStr, "LeaveID"))
        StdName.Text = TIMS.GetMyValue(s_SearchStr, "StdName")
        If (LCase(TIMS.GetMyValue(s_SearchStr, "ShowTable")) <> "true") Then Return
        Dim iPageIndex As Integer = Val(TIMS.GetMyValue(s_SearchStr, "PageIndex"))
        Call search1()
        If iPageIndex < 1 Then Return '>= 0 
        '有資料SHOW出 跳頁
        PageControler1.PageIndex = iPageIndex ' Me.ViewState("PageIndex")
        PageControler1.DataTableCreate(CPdt, PageControler1.Sort, PageControler1.PageIndex)

    End Sub


    '檢查輸入資料是否正確
    Function CheckData4(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        If start_date.Text <> "" AndAlso Not TIMS.IsDate1(start_date.Text) Then
            Errmsg += "請假期間開始日期有誤，應為日期格式" & start_date.Text & vbCrLf
        End If
        If end_date.Text <> "" AndAlso Not TIMS.IsDate1(end_date.Text) Then
            Errmsg += "請假期間結束日期有誤，應為日期格式" & start_date.Text & vbCrLf
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    '查詢班級紀錄 鈕
    Private Sub Button4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button4.Click
        Dim Errmsg As String = ""
        Call CheckData4(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim dt As DataTable
        Dim dt1 As DataTable
        Dim dr As DataRow
        Dim dr1 As DataRow
        Dim Cst_TestDate As Date = CDate("2000/01/01")
        Dim TempDate As Date = CDate("2000/01/01")

        Dim TempMon As String
        Dim TempDay1 As String
        Dim TempDay2 As String
        Dim TempCnt As Integer
        'Dim DataRound As String
        Dim DefSDate As String
        Dim DefFDate As String

        If OCIDValue1.Value = "" Then
            Common.RespWrite(Me, "<script>alert('" & "請選擇班級" & "');</script>")
            Exit Sub
        End If

        Table4.Visible = False
        TempCnt = 0
        TempDay1 = ""
        TempDay2 = ""

        '建立DataGird用的DataTable格式 Start
        dt = New DataTable
        dt.Columns.Add(New DataColumn("YearMon"))                       '月份
        dt.Columns.Add(New DataColumn("Recorded"))                      '已紀錄日期
        dt.Columns.Add(New DataColumn("NoRecord"))                      '未紀錄日期
        '建立DataGird用的DataTable格式 End

        DefSDate = "" & sm.UserInfo.Years & "/01/01"
        DefFDate = "" & sm.UserInfo.Years & "/12/31"
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)

        Dim parms As New Hashtable
        parms.Add("OCID", OCIDValue1.Value)
        parms.Add("SCHOOLDATE1", If(start_date.Text <> "", start_date.Text, DefSDate))
        parms.Add("SCHOOLDATE2", If(start_date.Text <> "", end_date.Text, DefFDate))

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " WITH WS1 AS (" & vbCrLf
        sql &= " 	SELECT s1.SCHOOLDATE FROM CLASS_SCHEDULE s1" & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        sql &= " 	and s1.OCID =@OCID" & vbCrLf
        sql &= "    and s1.SCHOOLDATE >= @SCHOOLDATE1" & vbCrLf
        sql &= "    and s1.SCHOOLDATE <= @SCHOOLDATE2" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " ,WS2 AS (" & vbCrLf
        sql &= " 	SELECT DISTINCT b1.LEAVEDATE" & vbCrLf
        sql &= " 	FROM CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " 	JOIN STUD_TURNOUT b1 on b1.socid =cs.socid" & vbCrLf
        sql &= " 	WHERE 1=1" & vbCrLf
        sql &= " 	and cs.OCID =@OCID" & vbCrLf
        sql &= " )" & vbCrLf
        sql &= " SELECT s1.SCHOOLDATE ,ISNULL(s2.LEAVEDATE," & TIMS.To_date(Cst_TestDate) & ") LEAVEDATE" & vbCrLf
        sql &= " FROM WS1 S1" & vbCrLf
        sql &= " LEFT JOIN WS2 S2 ON S2.LEAVEDATE=S1.SCHOOLDATE" & vbCrLf
        sql &= " ORDER by s1.SCHOOLDATE" & vbCrLf
        Try
            dt1 = DbAccess.GetDataTable(sql, objconn, parms)
        Catch ex As Exception
            TIMS.LOG.Error(ex.Message, ex)
            Common.MessageBox(Me, "資料庫效能異常，請重新查詢!!")
            Exit Sub
        End Try

        If dt1.Rows.Count > 0 Then
            For Each dr1 In dt1.Rows
                TempCnt = TempCnt + 1
                If dt1.Rows.Count = 1 Then
                    TempMon = Format(DatePart("yyyy", dr1("SchoolDate")) - 1911, "00") & "年" & Format(DatePart("m", dr1("SchoolDate")), "00") & "月"
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("YearMon") = TempMon
                    If dr1("SchoolDate") = dr1("LeaveDate") Then
                        dr("Recorded") = Format(DatePart("d", dr1("SchoolDate")), "00")
                        dr("NoRecord") = ""
                    Else
                        dr("Recorded") = ""
                        dr("NoRecord") = Format(DatePart("d", dr1("SchoolDate")), "00")
                    End If
                ElseIf CDate(TempDate).ToString("yyyy/MM/dd") = CDate(Cst_TestDate).ToString("yyyy/MM/dd") Then
                    If dr1("SchoolDate") = dr1("LeaveDate") Then
                        TempDay1 = Format(DatePart("d", dr1("SchoolDate")), "00") & ","
                    Else
                        TempDay2 = Format(DatePart("d", dr1("SchoolDate")), "00") & ","
                    End If
                ElseIf DateDiff(DateInterval.Month, TempDate, dr1("SchoolDate")) > 0 Then
                    TempMon = Format(DatePart("yyyy", TempDate) - 1911, "00") & "年" & Format(DatePart("m", TempDate), "00") & "月"
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("YearMon") = TempMon

                    If Len(TempDay1) = 0 Then
                        dr("Recorded") = ""
                    ElseIf Len(TempDay1) > 45 Then
                        dr("Recorded") = Left(TempDay1, 44) & vbCrLf & Mid(TempDay1, 46, Len(TempDay1) - 46)
                    Else
                        dr("Recorded") = Left(TempDay1, Len(TempDay1) - 1)
                    End If

                    If Len(TempDay2) = 0 Then
                        dr("NoRecord") = ""
                    ElseIf Len(TempDay2) > 45 Then
                        dr("NoRecord") = Left(TempDay2, 44) & vbCrLf & Mid(TempDay2, 46, Len(TempDay2) - 46)
                    Else
                        dr("NoRecord") = Left(TempDay2, Len(TempDay2) - 1)
                    End If
                    TempDay1 = ""
                    TempDay2 = ""
                    If dr1("SchoolDate") = dr1("LeaveDate") Then
                        TempDay1 = Format(DatePart("d", dr1("SchoolDate")), "00") & ","
                    Else
                        TempDay2 = Format(DatePart("d", dr1("SchoolDate")), "00") & ","
                    End If
                Else
                    If dr1("SchoolDate") = dr1("LeaveDate") Then
                        TempDay1 = TempDay1 & Format(DatePart("d", dr1("SchoolDate")), "00") & ","
                    Else
                        TempDay2 = TempDay2 & Format(DatePart("d", dr1("SchoolDate")), "00") & ","
                    End If
                End If
                TempDate = dr1("SchoolDate")
                If TempCnt > 1 And TempCnt = dt1.Rows.Count Then
                    TempMon = Format(DatePart("yyyy", dr1("SchoolDate")) - 1911, "00") & "年" & Format(DatePart("m", dr1("SchoolDate")), "00") & "月"
                    dr = dt.NewRow
                    dt.Rows.Add(dr)
                    dr("YearMon") = TempMon

                    If Len(TempDay1) = 0 Then
                        dr("Recorded") = ""
                    ElseIf Len(TempDay1) > 45 Then
                        dr("Recorded") = Left(TempDay1, 44) & vbCrLf & Mid(TempDay1, 46, Len(TempDay1) - 46)
                    Else
                        dr("Recorded") = Left(TempDay1, Len(TempDay1) - 1)
                    End If

                    If Len(TempDay2) = 0 Then
                        dr("NoRecord") = ""
                    ElseIf Len(TempDay2) > 45 Then
                        dr("NoRecord") = Left(TempDay2, 44) & vbCrLf & Mid(TempDay2, 46, Len(TempDay2) - 46)
                    Else
                        dr("NoRecord") = Left(TempDay2, Len(TempDay2) - 1)
                    End If
                End If
            Next
        End If

        Datagrid2.Visible = False
        Button7.Visible = False
        msg.Text = "查無資料!"
        If dt.Rows.Count = 0 Then Return

        Datagrid2.Visible = True
        Button7.Visible = True
        msg.Text = ""

        Datagrid2.DataSource = dt
        Datagrid2.DataBind()

        dt.Dispose()
        dt = Nothing
    End Sub

    '回上一頁 鈕
    Private Sub Button7_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button7.Click
        Datagrid2.Visible = False
        Button7.Visible = False
    End Sub

#Region "NO USE"
    'Private Sub DataGrid1_PreRender(ByVal sender As Object, ByVal e As System.EventArgs) Handles DataGrid1.PreRender
    '    ''若不列入缺曠課為勾選則「總計時數」顯示為0，不將顯示結果存至資料庫 
    '    'For Each item As DataGridItem In DataGrid1.Items
    '    '    Dim TurnoutIgnore As HtmlInputCheckBox = item.FindControl("TurnoutIgnore")
    '    '    If TurnoutIgnore.Checked = True Then
    '    '        item.Cells(15).Text = "0"
    '    '    End If
    '    'Next
    'End Sub

#End Region

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        Table4.Visible = False
    End Sub

    Private Sub Button9_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button9.Click
        '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

#Region "ImportData filed"
    Const cst_Defdatetime1 As String = "1999/01/01" '初始日期。
    Dim SOCdt As DataTable
    Dim gYMDvalue As String '日期yyyy/MM/dd
    Dim gSOCID As String = "" '學號
    Dim gCardNum As String = "" '卡號
    Dim gSesNum As String '節數 (節次 1~12)
    Dim gSwitch2 As String '上下課 (1:上課,2:下課)
    Dim gRecTime As DateTime '時間
    Dim gSCRID As String = "" '流水號
#End Region

    'gYMDvalue'V_CARDINFOLOG STUD_CARDRECORD STUD_CARDINFOLOG
    '1.取得班級學員基本資料
    Function sUtl_GetStuOfClass(ByVal OCIDValue As String) As DataTable
        OCIDValue = TIMS.ClearSQM(OCIDValue)
        Dim Rsdt As New DataTable
        If OCIDValue = "" Then Return Rsdt
        'select * from Stud_CardInfoLog where socid in ( SELECT socid FROM CLASS_STUDENTSOFCLASS where ocid = @ocid )
        Dim sSql As String = ""
        sSql = "" & vbCrLf
        sSql &= " SELECT s.SOCID, c.CardNum" & vbCrLf '卡號
        sSql &= " FROM V_STUDENTINFO s" & vbCrLf
        sSql &= " LEFT JOIN V_CARDINFOLOG c on c.socid =s.socid " & vbCrLf
        sSql &= " WHERE 1=1" & vbCrLf
        sSql &= " AND s.OCID= @OCID" & vbCrLf
        'sSql &= " AND c.CardNum = @CardNum" & vbCrLf
        Dim Cmd As New SqlCommand(sSql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With Cmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue
            Rsdt.Load(.ExecuteReader())
        End With
        Return Rsdt
    End Function

    '2.依卡號班級，取得班級學號SOCID  
    Function sUtl_GetSOCID(ByRef dtStuOfClass As DataTable, ByVal sCardNum As String) As String
        Dim rst As String = ""
        Try
            Dim f As String = "CardNum='" & sCardNum & "'"
            If dtStuOfClass.Select(f).Length > 0 Then
                rst = dtStuOfClass.Select(f)(0)("SOCID")
            End If
        Catch ex As Exception
        End Try
        Return rst
    End Function

    '3.依學號SOCID  卡號，檢查有無卡號
    Function sUtl_ChkCardNum(ByRef dtStuOfClass As DataTable, ByVal sSOCID As String, ByVal sCardNum As String) As Boolean
        Dim rst As Boolean = False '查無資料
        Try
            Dim f As String = ""
            f = ""
            f &= " CardNum='" & sCardNum & "'" '卡號
            f &= " AND SOCID='" & sSOCID & "'"
            If dtStuOfClass.Select(f).Length = 1 Then rst = True
        Catch ex As Exception
        End Try
        Return rst
    End Function

    '    Dim colArray As Array 
    '檢查後若有錯誤顯示錯誤訊息。無錯誤為空。
    Function CheckImportData(ByRef colArray As Array) As String
        Dim Reason As String = ""
        Const cst_Len As Integer = 25 '應該有多少欄位資料。
        'Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        'Dim sql As String
        'Dim dr As DataRow
        If colArray.Length <> cst_Len Then
            Reason += "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Return Reason
        End If

        '去除可能空白
        For i As Integer = 0 To colArray.Length - 1
            If colArray(i) <> "" Then colArray(i) = TIMS.ClearSQM(colArray(i))
            Select Case i
                Case 0
                    gCardNum = ""
                    gSOCID = ""
                    '1.卡號切換為大寫
                    gCardNum = TIMS.ChangeIDNO(colArray(i).ToString()) '卡號
                    '2.依卡號班級，取得 學號SOCID  
                    gSOCID = sUtl_GetSOCID(SOCdt, gCardNum) '班級 學號SOCID  
                    '3.依學號SOCID  卡號，檢查有無卡號
                    If gSOCID <> "" AndAlso gCardNum <> "" Then
                        If Not sUtl_ChkCardNum(SOCdt, gSOCID, gCardNum) Then
                            Reason += "卡號格式有誤<BR>"
                            Exit For
                        End If
                    Else
                        Reason += "卡號格式有誤<BR>"
                        Exit For
                    End If
                Case Else '其他1~24
                    '1.若有資料，檢查時間格式
                    If colArray(i).ToString() <> "" Then
                        Dim strValue As String = colArray(i).ToString()
                        If Not TIMS.sUtl_chkTIMEFormat(gYMDvalue, strValue) Then
                            Reason += "時間格式有誤:(" & i & "). " & strValue & "<BR>"
                            Exit For
                        Else
                            '時間格式無誤。
                            gRecTime = Convert.ToDateTime(gYMDvalue & " " & strValue)
                        End If
                    End If
            End Select
        Next

        Return Reason
    End Function

    '取得節次與上下課。
    Sub sUtl_GetiRow(iRow As Integer, ByRef rSesNum As Integer, ByRef rSwitch2 As Integer)
        Select Case iRow
            Case 1, 2
                rSesNum = 1
            Case 3, 4
                rSesNum = 2
            Case 5, 6
                rSesNum = 3
            Case 7, 8
                rSesNum = 4
            Case 9, 10
                rSesNum = 5
            Case 11, 12
                rSesNum = 6
            Case 13, 14
                rSesNum = 7
            Case 15, 16
                rSesNum = 8
            Case 17, 18
                rSesNum = 9
            Case 19, 20
                rSesNum = 10
            Case 21, 22
                rSesNum = 11
            Case 23, 24
                rSesNum = 12
        End Select
        Select Case iRow
            Case 1, 3, 5, 7, 9, 11, 13, 15, 17, 19, 21, 23
                rSwitch2 = 1
            Case 2, 4, 6, 7, 10, 12, 14, 16, 18, 20, 22, 24
                rSwitch2 = 2
        End Select
    End Sub

    '依資料位置檢查是否有DATA 有的話寫入sSCRID
    Function sUtl_ChkValue1(iRow As Integer, sSOCID As String, sCardNum As String, ByRef sSCRID As String) As Boolean
        Dim rst As Boolean = False
        Dim rSesNum As Integer = 0
        Dim rSwitch2 As Integer = 0
        Call sUtl_GetiRow(iRow, rSesNum, rSwitch2)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT SCRID " & vbCrLf
        sql &= " ,RecTime" & vbCrLf
        sql &= " FROM Stud_CardRecord" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " and SOCID = @SOCID" & vbCrLf
        sql &= " and CardNum = @CardNum" & vbCrLf
        sql &= " and SesNum = @SesNum" & vbCrLf
        sql &= " and Switch2 = @Switch2" & vbCrLf
        Dim sCmd As New SqlCommand(sql, objconn)

        Call TIMS.OpenDbConn(objconn)
        Dim odt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = sSOCID
            .Parameters.Add("CardNum", SqlDbType.VarChar).Value = sCardNum
            .Parameters.Add("SesNum", SqlDbType.VarChar).Value = rSesNum '1~12
            .Parameters.Add("Switch2", SqlDbType.VarChar).Value = rSwitch2 '1或2
            odt.Load(.ExecuteReader())
        End With
        sSCRID = ""
        If odt.Rows.Count > 0 Then
            sSCRID = odt.Rows(0)("SCRID")
            rst = True
        End If
        Return rst
    End Function

    '修改時間就好
    Sub sUtl_UpdataValue1(iRow As Integer, sSCRID As String, sRecTime As DateTime)
        Dim rSesNum As Integer = 0
        Dim rSwitch2 As Integer = 0
        Call sUtl_GetiRow(iRow, rSesNum, rSwitch2)

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " UPDATE Stud_CardRecord " & vbCrLf
        sql &= " SET RecTime = @RecTime" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        sql &= " AND SCRID = @SCRID" & vbCrLf
        Dim cmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        With cmd
            .Parameters.Clear()
            .Parameters.Add("RecTime", SqlDbType.DateTime).Value = sRecTime
            .Parameters.Add("SCRID", SqlDbType.VarChar).Value = sSCRID
            .ExecuteNonQuery()
        End With
    End Sub

    '新增時間 資料。
    Sub sUtl_AddValue1(iRow As Integer, sSOCID As String, sCardNum As String, sRecTime As DateTime)
        Dim rSesNum As Integer = 0
        Dim rSwitch2 As Integer = 0
        Call sUtl_GetiRow(iRow, rSesNum, rSwitch2)

        Dim sql As String = ""
        'Sql &= "  /* IDENTITY(1,1): SCRID */ " & vbCrLf
        sql = "" & vbCrLf
        sql &= " INSERT INTO Stud_CardRecord(" & vbCrLf
        sql &= " SCRID,SOCID,CardNum,SesNum,Switch2,RecTime" & vbCrLf
        sql &= " ) VALUES (" & vbCrLf
        sql &= " @SCRID,@SOCID,@CardNum,@SesNum,@Switch2,@RecTime" & vbCrLf
        sql &= " )" & vbCrLf
        Dim s_cmd As New SqlCommand(sql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim iSCRID As Integer = DbAccess.GetNewId(objconn, "STUD_CARDRECORD_SCRID_SEQ,STUD_CARDRECORD,SCRID")
        With s_cmd
            .Parameters.Clear()
            .Parameters.Add("SCRID", SqlDbType.Int).Value = iSCRID
            .Parameters.Add("SOCID", SqlDbType.VarChar).Value = sSOCID
            .Parameters.Add("CardNum", SqlDbType.VarChar).Value = sCardNum
            .Parameters.Add("SesNum", SqlDbType.VarChar).Value = rSesNum
            .Parameters.Add("Switch2", SqlDbType.VarChar).Value = rSwitch2
            .Parameters.Add("RecTime", SqlDbType.DateTime).Value = sRecTime
            .ExecuteNonQuery()
        End With
    End Sub

    '執行匯入動作。(檔名(日期 yyyyMMdd))
    Sub SUtl_ImprotX1(ByRef FullFileName1 As String)
        '上傳檔案
        File1.PostedFile.SaveAs(FullFileName1)
        'Common.MessageBox(Me, Request.BinaryRead(File1.PostedFile.ContentLength).ToString)

        Dim dt_xls As DataTable = Nothing
        Dim Reason As String = "" '儲存錯誤的原因
        '取得內容
        If (flag_File1_xls) Then
            Const cst_FirstCol1 As String = "卡號"
            dt_xls = TIMS.GetDataTable_XlsFile(FullFileName1, "", Reason, cst_FirstCol1)
            If Reason <> "" Then
                Common.MessageBox(Me, "無法匯入!!" & Reason)
                Exit Sub
            End If
        End If
        If (flag_File1_ods) Then dt_xls = TIMS.GetDataTable_ODSFile(FullFileName1)
        If (flag_File1_csv) Then dt_xls = TIMS.GetDataTable_CSVFile(FullFileName1)
        '刪除檔案 'IO.File.Delete(FullFileName1)
        TIMS.MyFileDelete(FullFileName1)

        Reason = TIMS.Chk_DTXLS1(dt_xls, flag_File1_xls, flag_File1_ods, flag_File1_csv)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        '將檔案讀出放入記憶體
        'Dim sr As System.IO.Stream
        'Dim srr As System.IO.StreamReader
        'sr = System.IO.File.OpenRead(FullFileName1)
        'srr = New System.IO.StreamReader(sr, System.Text.Encoding.Default)

        'Dim OneRow As String            'srr.ReadLine 一行一行的資料
        'Dim RowIndex As Integer = 0     '讀取行累計數
        'Dim Reason As String = ""       '儲存錯誤的原因
        'Dim strArray As String()

        Dim dtWrong As New DataTable    '儲存錯誤資料的DataTable
        Dim drWrong As DataRow
        '建立錯誤資料格式Table----------------Start
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("CardNum"))
        dtWrong.Columns.Add(New DataColumn("Reason"))

        Dim iRowIndex As Integer = 1
        For Each dr1 As DataRow In dt_xls.Rows

            dr1(0) = TIMS.ChangeIDNO(Convert.ToString(dr1(0))) '卡號

            Reason = CheckImportData(dr1.ItemArray) '檢查資料正確性

            If Reason = "" Then
                For iRow As Integer = 1 To 24 '跑24次資料塞入
                    gRecTime = Convert.ToDateTime(cst_Defdatetime1)
                    If dr1.ItemArray(iRow) <> "" Then
                        '時間格式無誤。
                        gRecTime = Convert.ToDateTime(gYMDvalue & " " & dr1.ItemArray(iRow))

                        '查詢資料庫有無資料。
                        If sUtl_ChkValue1(iRow, gSOCID, gCardNum, gSCRID) Then
                            '有
                            Call sUtl_UpdataValue1(iRow, gSCRID, gRecTime)
                        Else
                            '沒有
                            Call sUtl_AddValue1(iRow, gSOCID, gCardNum, gRecTime)
                        End If
                    End If
                Next

            Else
                '有錯誤。 '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)
                drWrong("Index") = iRowIndex
                drWrong("CardNum") = gCardNum '卡號
                drWrong("Reason") = Reason '問題。
            End If
            iRowIndex += 1 '讀取行累計數
        Next

        If dtWrong.Rows.Count = 0 Then
            Common.MessageBox(Me, "資料匯入完成。")
            Return
        End If

        Session("MyWrongTable") = dtWrong
        Page.RegisterStartupScript("", "<script>if(confirm('資料匯入成功，但有錯誤的資料無法匯入，是否要檢視原因?')){window.open('SD_05_002_Wrong.aspx','','width=700,height=600,location=0,status=0,menubar=0,scrollbars=1,resizable=0');}</script>")
    End Sub

    '匯入鈕。
    Protected Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "輸入班級有誤(尚無學員資料)!!")
            Exit Sub
        End If

        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value)
        If drCC Is Nothing Then
            Common.MessageBox(Me, "輸入班級有誤!!")
            Exit Sub
        End If

        SOCdt = sUtl_GetStuOfClass(OCIDValue1.Value)
        If SOCdt.Rows.Count = 0 Then
            Common.MessageBox(Me, "輸入班級有誤(尚無學員資料)!!")
            Exit Sub
        End If

        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods, flag_File1_csv)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Return
        End If
        '檢核檔案狀況 異常為False (有錯誤訊息視窗)
        Dim MyPostedFile As HttpPostedFile = Nothing
        If flag_File1_xls Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "xls") Then Return
        ElseIf flag_File1_ods Then
            If Not TIMS.HttpCHKFile(Me, File1, MyPostedFile, "ods") Then Return
        End If

        gYMDvalue = ""
        Dim sFileName1 As String = sMyFileName.Split(".")(0)
        Dim flag_errFileName As Boolean = False
        If Not flag_errFileName AndAlso Len(sFileName1) <> 8 Then flag_errFileName = True
        If Not flag_errFileName Then
            '字串yyyyMMdd取得日期yyyy/MM/dd。
            gYMDvalue = TIMS.sUtl_YMDValue(sFileName1)
        End If
        If Not flag_errFileName AndAlso gYMDvalue = "" Then flag_errFileName = True
        If flag_errFileName Then
            Common.MessageBox(Me, "檔案名稱有誤! 檔案名稱格式應為yyyyMMdd")
            Exit Sub
        End If

        Const Cst_FileSavePath As String = "~/SD/05/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call SUtl_ImprotX1(FullFileName1)
    End Sub

    'Sub CREATE_DG1(ByRef oDG1 As DataGrid, ByRef dtLEAVE As DataTable)
    '    'Const cst_DG1_序號 As Integer = 0
    '    'Const cst_DG1_最近登錄日期 As Integer = 1
    '    'Const cst_DG1_班別 As Integer = 2
    '    'Const cst_DG1_學號 As Integer = 3
    '    'Const cst_DG1_學員姓名 As Integer = 4
    '    Dim iCol As Integer = 5
    '    For Each dr1 As DataRow In dtLEAVE.Rows
    '        Dim s_HeaderText As String = Convert.ToString(dr1("NAME"))
    '        Dim s_DataField As String = Convert.ToString(dr1("ENGNAME"))
    '        oDG1.Columns.AddAt(iCol, CreateBoundColumn(s_DataField, s_HeaderText))
    '        iCol += 1
    '    Next
    'End Sub

    'Function CreateBoundColumn(DataFieldValue As String, HeaderTextValue As String) As BoundColumn
    '    '// This version of the CreateBoundColumn method sets only the
    '    '// DataField And HeaderText properties.
    '    '// Create a BoundColumn.
    '    Dim column As New BoundColumn()
    '    '// Set the properties of the BoundColumn.
    '    column.DataField = DataFieldValue
    '    column.HeaderText = HeaderTextValue
    '    column.HeaderStyle.HorizontalAlign = HorizontalAlign.Center
    '    column.ItemStyle.HorizontalAlign = HorizontalAlign.Center
    '    Return column
    'End Function

    'Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    'End Sub
End Class

