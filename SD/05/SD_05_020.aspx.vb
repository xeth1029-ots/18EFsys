Partial Class SD_05_020
    Inherits AuthBasePage

    'SD_05_020,'SD_05_020_3,'SD_05_020_1,'SD_05_020_2,
    Const cst_printFN0 As String = "SD_05_020"
    Const cst_printFN1 As String = "SD_05_020_1"
    Const cst_printFN2 As String = "SD_05_020_2"

    Dim flag_File1_xls As Boolean = False
    Dim flag_File1_ods As Boolean = False

    Dim sMemo As String = "" '(查詢原因)
    Dim aIDNO As String '身分證號碼 --varchar(15)
    Dim Idnew As String '新的IDNO

    Dim dt As DataTable
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁,'檢查Session是否存在 Start,'TIMS.CheckSession(Me),TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1
        PageControler2.PageDataGrid = DataGrid2
        PageControler4.PageDataGrid = Datagrid4
        trButton10.Visible = False '不顯示

        If Not Page.IsPostBack Then
            DataGridTable.Visible = False
            DataGridTable2.Visible = False
            DataGridTable4.Visible = False

            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            Print.Enabled = False
            eMeng.Visible = False

            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            If sm.UserInfo.LID <> 2 Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button7_Click(sender, e)
            End If
        End If

        org_tr.Visible = False '不顯示
        class_tr.Visible = False '不顯示
        idno_tr.Visible = True '顯示
        date_tr.Visible = True '顯示
        'file_tr.Visible = True '顯示
        If searchMode.SelectedIndex = 0 Then
            org_tr.Visible = True '顯示
            class_tr.Visible = True '顯示
            idno_tr.Visible = False '不顯示
            date_tr.Visible = False '不顯示
            'file_tr.Visible = False '不顯示
        End If

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= "1" Then
            Button8.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button8.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "Button1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'Print.Attributes("onclick") = "javascript:return CheckPrint();"
        Button1.Attributes("onclick") = "javascript:return CheckPrint();"
        Button10.Attributes("onclick") = "javascript:return EXLINPUT_Chenk();"

    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

        If searchMode.SelectedIndex = 0 Then
            TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        Else
            TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid2)
        End If

        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        '呼叫查詢
        Call LoDATE()

    End Sub

    '呼叫查詢
    Sub LoDATE()
        TIMS.SUtl_TxtPageSize(Me, TxtPageSize, Datagrid4) '顯示列數不正確

        Select Case searchMode.SelectedValue '1:依班別查詢 2:依個別學員查詢
            Case "1" '假如是以班級查詢
                If OCIDValue1.Value = "" Then
                    Common.MessageBox(Me, "以班級查詢，請選擇班級名稱!!")
                    Exit Sub
                End If
                Call searchMode1_1()

            Case "2" '2:依個別學員查詢
                If IDNO.Text <> "" OrElse Idnew <> "" Then
                    '以身分證字號及重複日期查詢,找出這個身證字號在所指定的時間,重複參訓的班級
                    Call LoDateIDNO()
                Else
                    Common.MessageBox(Me, "資料有誤，請輸入有效身分證號，謝謝!!")
                    Exit Sub
                End If
            Case Else
                If Idnew <> "" Then
                    '以身分證字號及重複日期查詢,找出這個身證字號在所指定的時間,重複參訓的班級
                    Call LoDateIDNO()
                Else
                    Common.MessageBox(Me, "資料有誤，請重新選擇要查詢的項目，謝謝!!")
                    Exit Sub
                End If
        End Select
    End Sub

    '查詢原因
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        Dim V_searchMode As String = TIMS.GetListValue(searchMode)
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        center.Text = TIMS.ClearSQM(center.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)

        If V_searchMode <> "" Then RstMemo &= String.Concat("&查詢類型=", V_searchMode)
        If start_date.Text <> "" Then RstMemo &= String.Concat("&重複參訓日期起日=", start_date.Text)
        If end_date.Text <> "" Then RstMemo &= String.Concat("&重複參訓日期迄日=", end_date.Text)
        If center.Text <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Text)
        If IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", IDNO.Text)
        Return RstMemo
    End Function
    'SQL '假如是以班級查詢 是一般計畫
    Sub searchMode1_1()
        Dim STDate As String = ""
        Dim FTDate As String = ""
        Dim sql As String = ""
        sql &= " select format(STDate,'yyyy/MM/dd') STDate ,format(FTDate,'yyyy/MM/dd') FTDate" & vbCrLf
        sql &= " FROM CLASS_CLASSINFO WHERE OCID=@OCID" & vbCrLf

        If OCIDValue1.Value <> "" Then
            Dim sCmd As New SqlCommand(sql, objconn)
            dt = New DataTable
            Call TIMS.OpenDbConn(objconn)
            With sCmd
                .Parameters.Clear()
                .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
                dt.Load(.ExecuteReader())
            End With
            If dt.Rows.Count > 0 Then
                Dim dr As DataRow = dt.Rows(0)
                STDate = dr("STDate")
                FTDate = dr("FTDate")
            End If
        End If
        STDate = TIMS.Cdate3(STDate)
        FTDate = TIMS.Cdate3(FTDate)

        sql = "" & vbCrLf
        sql &= " WITH WAS1 AS ( select ss.Name" & vbCrLf
        sql &= " ,ss.IDNO ,ss.DISTNAME ,ss.PLANNAME ,ss.ORGNAME ,ss.CLASSCNAME2 classname" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_ONCLASS(ss.PlanID,ss.ComIDNO,ss.SeqNo,'WEEKTIME') WEEKS" & vbCrLf
        sql &= " ,concat(format(ss.STDate,'yyyy/MM/dd'),' ~ ',format(ss.FTDate,'yyyy/MM/dd')) Sfdate,ss.STUDSTATUS,ss.STUDSTATUS2" & vbCrLf
        sql &= " FROM dbo.V_STUDENTINFO ss" & vbCrLf
        sql &= " WHERE ss.OCID =@OCID" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " ,WAS2 AS ( select ss.Name" & vbCrLf
        sql &= " ,ss.IDNO,ss.DISTNAME,ss.planname,ss.OrgName" & vbCrLf
        sql &= " ,ss.CLASSCNAME2 classname" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_ONCLASS(ss.PlanID,ss.ComIDNO,ss.SeqNo,'WEEKTIME') WEEKS" & vbCrLf
        sql &= " ,concat(format(ss.STDate,'yyyy/MM/dd'),' ~ ',format(ss.FTDate,'yyyy/MM/dd')) Sfdate,ss.STUDSTATUS,ss.STUDSTATUS2" & vbCrLf
        sql &= " FROM WAS1 a" & vbCrLf
        sql &= " JOIN dbo.V_STUDENTINFO ss ON ss.IDNO=a.IDNO and ss.OCID <> @OCID" & vbCrLf
        sql &= " CROSS JOIN (SELECT STDate,FTDate FROM class_classinfo WHERE OCID =@OCID) CJ1" & vbCrLf
        sql &= " WHERE ( (CJ1.STDate<=ss.STDate and ss.STDate<=CJ1.FTDate) or (CJ1.STDate<=ss.FTDate and ss.FTDate<= CJ1.FTDate)" & vbCrLf
        sql &= " or (ss.STDate<= CJ1.STDate and ss.FTDate>= CJ1.FTDate) )" & vbCrLf
        sql &= " )" & vbCrLf

        sql &= " select t1.Name" & vbCrLf
        sql &= " ,t1.IDNO" & vbCrLf
        sql &= " ,t1.DISTNAME" & vbCrLf
        sql &= " ,t1.planname" & vbCrLf
        sql &= " ,t1.OrgName" & vbCrLf
        sql &= " ,t1.classname" & vbCrLf
        sql &= " ,t1.Sfdate,t1.STUDSTATUS2 STUDSTATUS_N1" & vbCrLf
        sql &= " ,t1.WEEKS" & vbCrLf
        sql &= " ,t2.DISTNAME DISTNAME2" & vbCrLf
        sql &= " ,t2.planname planname2" & vbCrLf
        sql &= " ,t2.OrgName OrgName2" & vbCrLf
        sql &= " ,t2.classname classname2" & vbCrLf
        sql &= " ,t2.Sfdate Sfdate2" & vbCrLf
        sql &= " ,t2.WEEKS WEEKS2" & vbCrLf
        sql &= " from WAS1 t1" & vbCrLf
        sql &= " join WAS2 t2 on t1.idno=t2.idno" & vbCrLf
        sql &= " ORDER BY t1.IDNO" & vbCrLf

        Dim sCmd2 As New SqlCommand(sql, objconn)
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCIDValue1.Value
            dt = New DataTable
            dt.Load(.ExecuteReader())
        End With

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "IDNO,NAME,DISTNAME,PLANNAME,ORGNAME,CLASSNAME,SFDATE,DISTNAME2,PLANNAME2,ORGNAME2,CLASSNAME2,SFDATE2,WEEKS2,STUDSTATUS_N1")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        eMeng.Visible = False
        DataGridTable.Visible = False
        Print.Enabled = False
        msg.Visible = True
        msg.Text = "查無資料!!"
        If TIMS.dtNODATA(dt) Then Return

        'If dt.Rows.Count > 0 Then 'End If
        'eMeng.Visible = False
        DataGridTable.Visible = True
        Print.Enabled = True
        msg.Visible = False
        msg.Text = ""

        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

    End Sub

    '以身分證字號及重複日期查詢,找出這個身證字號在所指定的時間,重複參訓的班級
    Sub LoDateIDNO()
        start_date.Text = TIMS.Cdate3(start_date.Text)
        end_date.Text = TIMS.Cdate3(end_date.Text)
        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))

        Dim sql As String = ""
        sql &= " select ss.Name Name3" & vbCrLf
        sql &= " ,ss.IDNO IDNO3" & vbCrLf
        sql &= " ,ss.DISTNAME DISTNAME3" & vbCrLf
        sql &= " ,ss.planname PLANNAME3" & vbCrLf
        sql &= " ,ss.OrgName ORGNAME3" & vbCrLf
        sql &= " ,ss.CLASSCNAME2 CLASSNAME3" & vbCrLf
        sql &= " ,concat(format(ss.STDate,'yyyy/MM/dd'),' ~ ',format(ss.FTDate,'yyyy/MM/dd')) Sfdate3" & vbCrLf
        sql &= " ,h.ExamName" & vbCrLf
        sql &= " ,ss.STUDSTATUS" & vbCrLf
        sql &= " ,ss.STUDSTATUS2" & vbCrLf
        sql &= " ,dbo.FN_GET_PLAN_ONCLASS(ss.PlanID,ss.ComIDNO,ss.SeqNo,'WEEKTIME') WEEKS" & vbCrLf
        sql &= " FROM dbo.V_STUDENTINFO ss" & vbCrLf
        sql &= " LEFT JOIN Stud_TechExam h ON ss.SOCID=h.SOCID" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and exists( select s2.SOCID" & vbCrLf
        sql &= "   from Class_ClassInfo s1" & vbCrLf
        sql &= "   join Class_StudentsOfClass s2 on s2.OCID=s1.OCID" & vbCrLf
        sql &= "   join Stud_StudentInfo s3 on s3.SID=s2.SID" & vbCrLf
        sql &= "   where s3.IDNO=ss.IDNO" & vbCrLf
        sql &= "   and s1.notopen ='N'" & vbCrLf
        sql &= "   and s2.SOCID != ss.SOCID" & vbCrLf
        sql &= "   and ( (ss.STDate<=s1.STDate and s1.STDate<=ss.FTDate)" & vbCrLf
        sql &= "   	or ( ss.STDate <=s1.FTDate and s1.FTDate< ss.FTDate)" & vbCrLf
        sql &= "   	or (s1.STDate <= ss.STDate and s1.FTDate >= ss.FTDate ) ) )" & vbCrLf
        sql &= " and ss.IDNO in ( select ss.IDNO FROM dbo.V_STUDENTINFO ss where exists (" & vbCrLf
        sql &= "   select s2.SOCID" & vbCrLf
        sql &= "   from Class_ClassInfo s1" & vbCrLf
        sql &= "   join Class_StudentsOfClass s2 on s2.OCID=s1.OCID" & vbCrLf
        sql &= "   join Stud_StudentInfo s3 on s3.SID=s2.SID" & vbCrLf
        sql &= "   where s3.IDNO=ss.IDNO" & vbCrLf
        sql &= "   and s2.SOCID<>ss.SOCID" & vbCrLf
        sql &= "   and s1.notopen ='N'" & vbCrLf
        sql &= "   and ( (ss.STDate<=s1.STDate and s1.STDate<=ss.FTDate) or ( ss.STDate <=s1.FTDate and s1.FTDate< ss.FTDate)" & vbCrLf
        sql &= "   	or (s1.STDate <= ss.STDate and s1.FTDate >= ss.FTDate ) ) )" & vbCrLf

        If IDNO.Text <> "" Then
            sql &= " and ss.IDNO = @IDNO" & vbCrLf
        Else
            '有一外部參數呼叫。
            If Idnew IsNot Nothing Then Idnew = UCase(Idnew)
            If Idnew <> "" Then sql &= " and ss.IDNO IN (" & Idnew & ")" & vbCrLf
        End If
        sql &= " 	)" & vbCrLf
        If start_date.Text <> "" AndAlso end_date.Text <> "" Then
            sql &= " and (1!=1" & vbCrLf
            sql &= " 	or (@STDate<=ss.STDate and ss.STDate<=@FTDate)" & vbCrLf
            sql &= " 	or (@STDate<=ss.FTDate and ss.FTDate<= @FTDate)" & vbCrLf
            sql &= " )" & vbCrLf
        Else
            If start_date.Text <> "" Then
                '如果只輸入一個日期的話,就找出這個學員在這個日期所參加的課程
                'sql &= " and  ( @STDate between cc.STDate and cc.FTDate)" & vbCrLf
                sql &= " and  (ss.STDate <=@STDate and @STDate<= ss.FTDate)" & vbCrLf
            End If
            If end_date.Text <> "" Then
                '如果只輸入一個日期的話,就找出這個學員在這個日期所參加的課程
                sql &= " and  (cc.STDate <=@FTDate and @FTDate<= cc.FTDate)" & vbCrLf
            End If
        End If

        Call TIMS.OpenDbConn(objconn)
        'Dim reader As OracleDataReader
        Dim cmd As New SqlCommand(sql, objconn)
        cmd.Parameters.Clear()
        If IDNO.Text <> "" Then
            cmd.Parameters.Add("IDNO", SqlDbType.VarChar).Value = IDNO.Text
        End If
        If start_date.Text <> "" AndAlso end_date.Text <> "" Then
            cmd.Parameters.Add("STDate", SqlDbType.DateTime).Value = CDate(start_date.Text)
            cmd.Parameters.Add("FTDate", SqlDbType.DateTime).Value = CDate(end_date.Text)
        Else
            If start_date.Text <> "" Then
                '如果只輸入一個日期的話,就找出這個學員在這個日期所參加的課程
                cmd.Parameters.Add("STDate", SqlDbType.DateTime).Value = CDate(start_date.Text)
            End If
            If end_date.Text <> "" Then
                '如果只輸入一個日期的話,就找出這個學員在這個日期所參加的課程
                cmd.Parameters.Add("FTDate", SqlDbType.DateTime).Value = CDate(end_date.Text)
            End If
        End If
        dt = New DataTable
        dt.Load(cmd.ExecuteReader())

        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        sMemo = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "IDNO3,DISTNAME3,PLANNAME3,ORGNAME3,SFDATE3,EXAMNAME,WEEKS,STUDSTATUS2")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        eMeng.Visible = False
        DataGridTable2.Visible = False
        Print.Enabled = False
        msg.Visible = True
        msg.Text = "查無資料!!"
        If TIMS.dtNODATA(dt) Then Return

        'If dt.Rows.Count > 0 Then 'End If
        'eMeng.Visible = False
        DataGridTable2.Visible = True
        Print.Enabled = True
        msg.Visible = False
        msg.Text = ""

        PageControler2.PageDataTable = dt
        PageControler2.ControlerLoad()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 
        End Select
        'If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
        '    e.Item.Cells(0).Text = e.Item.ItemIndex + 1 + DataGrid1.PageSize * DataGrid1.CurrentPageIndex
        'End If
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As System.Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號  
        End Select
    End Sub

    Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
        If searchMode.SelectedIndex = 0 Then
            '依班級
            If OCIDValue1.Value = "" Then
                Common.MessageBox(Me, "請選擇班級名稱!!")
                Exit Sub
            End If

            Dim STDate As String = ""
            Dim FTDate As String = ""
            Dim dr As DataRow
            dr = TIMS.GetOCIDDate(OCIDValue1.Value)
            If Not dr Is Nothing Then
                STDate = Common.FormatDate(dr("STDate"))
                FTDate = Common.FormatDate(dr("FTDate"))
            End If
            'STDate = DbAccess.ExecuteScalar("SELECT to_date(STDate,'yyyy/MM/dd') STDate FROM class_classinfo WHERE OCID='" & OCIDValue1.Value & "'")
            'FTDate = DbAccess.ExecuteScalar("SELECT to_date(FTDate,'yyyy/MM/dd') FTDate FROM class_classinfo WHERE OCID='" & OCIDValue1.Value & "'")
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN0, "OCID=" & OCIDValue1.Value & "&STDate=" & STDate & "&FTDate=" & FTDate)
            'If Session("TPlanID") <> "15" Then
            '    SmartQuery.PrintReport(Me, "Report", "SD_05_020", "OCID=" & OCIDValue1.Value & "&STDate=" & STDate & "&FTDate=" & FTDate)
            'Else
            '    SmartQuery.PrintReport(Me, "Report", "SD_05_020_3", "OCID=" & OCIDValue1.Value & "&STDate=" & STDate & "&FTDate=" & FTDate)
            'End If

        Else

            If start_date.Text <> "" OrElse end_date.Text <> "" Then
            Else
                Common.MessageBox(Me, "請填寫重複參訓日期!!")
                Exit Sub
            End If

            Dim newIDNO As String = ""
            If IDNO.Text <> "" Then
                newIDNO = "\'" & IDNO.Text & "\'"
            Else
                newIDNO = Idnew2.Value '多組IDNO
            End If

            If start_date.Text <> "" AndAlso end_date.Text <> "" Then
                TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, "&IDNO=" & newIDNO & "&STDate=" & start_date.Text.ToString & "&FTDate=" & end_date.Text.ToString)
            Else
                If start_date.Text <> "" Then
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, "&IDNO=" & newIDNO & "&STDate=" & start_date.Text.ToString)
                End If
                If end_date.Text <> "" Then
                    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, "&IDNO=" & newIDNO & "&FTDate=" & end_date.Text.ToString)
                End If
            End If

        End If
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        DataGridTable2.Visible = False
        DataGridTable4.Visible = False
        eMeng.Visible = False
        msg.Visible = False
        '判斷機構是否只有一個班級'如果只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
        DataGridTable2.Visible = False
        DataGridTable4.Visible = False
        eMeng.Visible = False
        msg.Visible = False
    End Sub

    Sub SImportFile1(ByRef FullFileName1 As String)
        Dim htSS As New Hashtable
        TIMS.SetMyValue2(htSS, "FullFileName1", FullFileName1)
        TIMS.SetMyValue2(htSS, "FirstCol", "身分證號碼") '任1欄位名稱(必填)
        Dim Reason As String = ""
        '上傳檔案/取得內容
        Dim dt_xls As DataTable = TIMS.Get_File1data(File1, Reason, htSS, flag_File1_xls, flag_File1_ods)
        If Reason <> "" Then
            Common.MessageBox(Me, Reason)
            Exit Sub
        End If

        Idnew = ""
        Idnew2.Value = ""

        Dim iRowIndex As Integer = 0 '讀取行累計數
        '建立錯誤資料格式Table----------------Start
        'Dim Reason As String = "" '儲存錯誤的原因
        Dim drWrong As DataRow = Nothing
        Dim dtWrong As New DataTable '儲存錯誤資料的DataTable
        dtWrong.Columns.Add(New DataColumn("Index"))
        dtWrong.Columns.Add(New DataColumn("IDNO"))
        dtWrong.Columns.Add(New DataColumn("Reason"))
        '建立錯誤資料格式Table----------------End
        For i As Integer = 0 To dt_xls.Rows.Count - 1
            Reason = ""
            Dim colArray As Array = dt_xls.Rows(i).ItemArray 'Split(OneRow, flag)
            Reason += CheckImportData(colArray) '檢查資料正確性

            If Reason <> "" Then
                '錯誤資料，填入錯誤資料表
                drWrong = dtWrong.NewRow
                dtWrong.Rows.Add(drWrong)

                drWrong("Index") = iRowIndex
                'If colArray.Length > 5 Then
                drWrong("IDNO") = TIMS.ChangeIDNO(aIDNO)
                drWrong("Reason") = Reason
                'End If
            Else
                If Idnew <> "" Then Idnew += ","
                Idnew += "'" & TIMS.ChangeIDNO(aIDNO) & "'" '取得 匯入的idno
                If Idnew2.Value <> "" Then Idnew2.Value += ","
                Idnew2.Value += "\'" & TIMS.ChangeIDNO(aIDNO) & "\'" '取得 匯入的idno

            End If

            iRowIndex += 1 '讀取行累計數
        Next

        If dtWrong.Rows.Count <> 0 Then
            eMeng.Visible = True
            Datagrid3.Style.Item("display") = "inline"

            Datagrid3.Visible = True
            Datagrid3.DataSource = dtWrong
            Datagrid3.DataBind()
            Common.MessageBox(Me, "身分證字號有誤,請重新修正後,再匯入,謝謝!!!")

        Else
            '呼叫查詢
            Call LoDATE()
        End If

    End Sub

    '匯入名冊
    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        Dim sMyFileName As String = ""
        Dim sErrMsg As String = TIMS.ChkFile1(File1, sMyFileName, flag_File1_xls, flag_File1_ods)
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

        Const Cst_FileSavePath As String = "~/SD/05/Temp/"
        Call TIMS.MyCreateDir(Me, Cst_FileSavePath)

        '3. 使用 Path.GetExtension 取得副檔名 (包含點，例如 .jpg)
        Dim fileNM_Ext As String = System.IO.Path.GetExtension(File1.PostedFile.FileName).ToLower()
        sMyFileName = $"f{TIMS.GetDateNo()}_{TIMS.GetRandomS4()}{fileNM_Ext}"
        Dim FullFileName1 As String = Server.MapPath(Cst_FileSavePath & sMyFileName)
        Call SImportFile1(FullFileName1)
    End Sub

    Function CheckImportData(ByVal colArray As Array) As String
        Dim Reason As String = ""
        'Dim SearchEngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ- "
        'Dim sql As String
        'Dim dr As DataRow
        Const cst_Len As Integer = 1
        If colArray.Length <> cst_Len Then
            'Reason += "欄位數量不正確(應該為" & cst_Len & "個欄位)<BR>"
            Reason += "欄位對應有誤<BR>"
            Reason += "請注意欄位中是否有半形逗點<BR>"
            aIDNO = colArray(0).ToString
        Else
            aIDNO = TIMS.ChangeIDNO(colArray(0).ToString) '身分證號碼
            If aIDNO = "" Then
                Reason += "必須填寫身分證號碼<BR>"
            Else
                Dim IDNOFlag As Boolean = True
                Dim EngStr As String = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
                If EngStr.IndexOf(aIDNO.ToUpper.Chars(0)) <> -1 And EngStr.IndexOf(aIDNO.ToUpper.Chars(1)) <> -1 Then '判斷身分證字號的前2碼若是英文字母就不檢查身分證字號
                    IDNOFlag = True
                Else   '檢查身分證字號
                    If aIDNO.Length <> 10 Then
                        IDNOFlag = False
                    ElseIf aIDNO.Chars(1) <> "1" And aIDNO.Chars(1) <> "2" Then
                        IDNOFlag = False
                    ElseIf EngStr.IndexOf(aIDNO.ToUpper.Chars(0)) = -1 Then
                        IDNOFlag = False
                    ElseIf aIDNO = "A123456789" Then
                        IDNOFlag = False
                    End If
                    If Not IDNOFlag Then
                        Reason += "身分證號碼錯誤!<BR>"
                    End If
                End If
            End If
        End If
        Return Reason
    End Function

    Private Sub Datagrid4_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles Datagrid4.ItemDataBound
        If e.Item.ItemType <> ListItemType.Header And e.Item.ItemType <> ListItemType.Footer Then
            Dim drv As DataRowView = e.Item.DataItem
            e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號 
            If e.Item.Cells(4).Text = "學習券" Then
                e.Item.Cells(7).Text = drv("Sfdate5")
            Else
                e.Item.Cells(7).Text = drv("Sfdate4")
            End If
        End If
    End Sub

    Protected Sub searchMode_SelectedIndexChanged(sender As Object, e As EventArgs) Handles searchMode.SelectedIndexChanged
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        IDNO.Text = ""
        start_date.Text = ""
        end_date.Text = ""
        DataGridTable.Visible = False
        DataGridTable2.Visible = False
        DataGridTable4.Visible = False
        msg.Visible = False
    End Sub

End Class
