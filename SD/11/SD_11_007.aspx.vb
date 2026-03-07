Partial Class SD_11_007
    Inherits AuthBasePage

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        PageControler1.PageDataGrid = DataGrid1 '設定分頁

        If Not Page.IsPostBack Then
            CCreate1()
        End If

        If searchMode.SelectedIndex = 0 Then
            org_tr.Visible = True
            class_tr.Visible = True
            idno_tr.Visible = False
            birthday_tr.Visible = False
            date_tr.Visible = False
        ElseIf searchMode.SelectedIndex = 1 Then
            org_tr.Visible = False
            class_tr.Visible = False
            idno_tr.Visible = True
            birthday_tr.Visible = True
            date_tr.Visible = False
        Else
            org_tr.Visible = False
            class_tr.Visible = False
            idno_tr.Visible = False
            birthday_tr.Visible = False
            date_tr.Visible = True
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        'Print.Attributes("onclick") = "javascript:return CheckPrint();"   '(由於ReportServer目前並無相對應的報表ID,經公司內容討論後,先將[列印]功能拿掉，By:20180914)
        Button2.Attributes("onclick") = "javascript:return CheckPrint();"
    End Sub

    Sub CCreate1()
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        Dim V_INQUIRY As String = Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me)))
        If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

        If sm.UserInfo.LID <> "2" Then
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        Else
            Call SCH_OnlyOne_OCID()
        End If
    End Sub

#Region "(由於ReportServer目前並無相對應的報表ID,經公司內容討論後,先將[列印]功能拿掉，By:20180914)"

    'Private Sub Print_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Print.Click
    '    Dim MyValue As String = ""
    '    MyValue = ""
    '    MyValue &= "OCID=" & OCIDValue1.Value
    '    MyValue &= "&IDNO=" & IDNO.Text
    '    MyValue &= "&birthday=" & birthday.Text
    '    MyValue &= "&Stdate1=" & start_date.Text
    '    MyValue &= "&Stdate2=" & end_date.Text
    '    TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, "SD_11_007", MyValue)
    'End Sub

#End Region

    Sub SCH_OnlyOne_OCID()
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        Table4.Visible = False
        'msg.Visible = False
        msg.Text = ""
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        Table4.Visible = False
        'msg.Visible = False
        msg.Text = ""
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        SCH_OnlyOne_OCID()
    End Sub

    Private Sub searchMode_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles searchMode.SelectedIndexChanged
        center.Text = ""
        RIDValue.Value = ""
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        IDNO.Text = ""
        birthday.Text = ""

        start_date.Text = ""
        end_date.Text = ""
        Table4.Visible = False
        'msg.Visible = False
        msg.Text = ""
    End Sub
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Text = TIMS.ClearSQM(center.Text)
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        birthday.Text = TIMS.ClearSQM(birthday.Text)
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)

        If center.Text <> "" Then RstMemo &= String.Concat("&center=", center.Text)
        If RIDValue.Value <> "" Then RstMemo &= String.Concat("&RID=", RIDValue.Value)
        If IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", IDNO.Text)
        If birthday.Text <> "" Then RstMemo &= String.Concat("&birthday=", birthday.Text)
        If start_date.Text <> "" Then RstMemo &= String.Concat("&start_date=", start_date.Text)
        If end_date.Text <> "" Then RstMemo &= String.Concat("&end_date=", end_date.Text)
        Return RstMemo
    End Function

    Sub sSearch1()
        'Dim dt As DataTable

        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        start_date.Text = TIMS.ClearSQM(start_date.Text)
        end_date.Text = TIMS.ClearSQM(end_date.Text)
        If start_date.Text <> "" Then start_date.Text = TIMS.Cdate3(start_date.Text)
        If end_date.Text <> "" Then end_date.Text = TIMS.Cdate3(end_date.Text)

        birthday.Text = TIMS.ClearSQM(birthday.Text)
        If birthday.Text <> "" Then birthday.Text = TIMS.Cdate3(birthday.Text)

        IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(IDNO.Text))

        '班級資訊
        Dim DateStr1 As String = ""
        'convert(datetime, '" & end_date.Text & "', 111) " & vbCrLf
        '學員資訊
        Dim DateStr2 As String = ""
        Dim vSchMode As String = TIMS.ClearSQM(searchMode.SelectedValue)
        Select Case vSchMode
            Case "1" '班級
                Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
                If drCC Is Nothing Then
                    Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
                    Exit Sub
                End If

                If RIDValue.Value.Length > 1 Then DateStr1 &= " AND cc.RID = '" & RIDValue.Value & "' " & vbCrLf
                If OCIDValue1.Value <> "" Then DateStr1 &= " AND cc.OCID = '" & OCIDValue1.Value & "' " & vbCrLf
            Case "2" '學員
                If IDNO.Text = "" Then
                    Common.MessageBox(Me, "請輸入至少一個查詢條件!(身分證號碼)")
                    Exit Sub
                End If

                If IDNO.Text <> "" Then DateStr2 &= " AND ss.IDNO = '" & IDNO.Text & "' " & vbCrLf
                If birthday.Text <> "" Then DateStr2 &= " AND ss.birthday = " & TIMS.To_date(birthday.Text) & vbCrLf 'convert(datetime, '" & birthday.Text & "', 111) " & vbCrLf

            Case "3" '使用開訓日期
                If start_date.Text = "" OrElse end_date.Text = "" Then
                    Common.MessageBox(Me, "日期區間查詢條件不可為空!")
                    Exit Sub
                End If
                '日期置換
                If DateDiff(DateInterval.Day, CDate(start_date.Text), CDate(end_date.Text)) < 0 Then
                    Dim tmp As String = start_date.Text
                    start_date.Text = end_date.Text
                    end_date.Text = tmp
                End If
                If start_date.Text <> "" Then DateStr1 &= " AND cc.Stdate >= " & TIMS.To_date(start_date.Text) & vbCrLf
                If end_date.Text <> "" Then DateStr1 &= " AND cc.Stdate <= " & TIMS.To_date(end_date.Text) & vbCrLf
            Case Else
                Common.MessageBox(Me, "查詢類型有誤!")
                Exit Sub
        End Select

        If DateStr1 = "" AndAlso DateStr2 = "" Then
            Common.MessageBox(Me, "請輸入至少一個查詢條件!")
            Exit Sub
        End If

        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim sql As String = ""
        Select Case vSchMode
            Case "1" '班級
                sql = "" & vbCrLf
                sql &= " WITH WC1 AS (SELECT SOCID,SID,OCID,SETID,etenterdate,sernum FROM class_studentsofclass WHERE OCID='" & OCIDValue1.Value & "')" & vbCrLf
                sql &= " ,WC2 AS (SELECT  max(b.SB2ID) SB2ID,b.IDNO,b.SOCID FROM STUD_BLIGATEDATA28 b where 1=1 and CHANGEMODE<>2 AND  b.socid in (select socid from wc1)  GROUP BY b.IDNO,b.SOCID )" & vbCrLf
            Case "2" '學員
                sql = "" & vbCrLf
                sql &= " WITH WC1S AS (SELECT SID FROM stud_studentinfo WHERE 1=1 and idno='" & IDNO.Text & "')" & vbCrLf
                sql &= " ,WC1 AS (SELECT cs.SOCID,cs.SID,cs.OCID,cs.SETID,cs.etenterdate,cs.sernum FROM WC1S a join class_studentsofclass cs on cs.sid =a.sid )" & vbCrLf
                sql &= " ,WC2 AS (SELECT  max(b.SB2ID) SB2ID,b.IDNO,b.SOCID FROM STUD_BLIGATEDATA28 b where 1=1 and CHANGEMODE<>2 AND  b.socid in (select socid from wc1)  GROUP BY b.IDNO,b.SOCID )" & vbCrLf
            Case "3" '使用開訓日期
                sql = "" & vbCrLf
                sql &= " WITH WC1C AS (SELECT ocid FROM CLASS_CLASSINFO WHERE 1=1 and STDATE>=" & TIMS.To_date(start_date.Text) & " and STDATE<=" & TIMS.To_date(end_date.Text) & ")" & vbCrLf
                sql &= " ,WC1 AS (SELECT cs.SOCID,cs.SID,cs.OCID,cs.SETID,cs.etenterdate,cs.sernum FROM WC1C a join class_studentsofclass cs on cs.ocid =a.ocid )" & vbCrLf
                sql &= " ,WC2 AS (SELECT  max(b.SB2ID) SB2ID,b.IDNO,b.SOCID FROM STUD_BLIGATEDATA28 b where 1=1 and CHANGEMODE<>2 AND  b.socid in (select socid from wc1)  GROUP BY b.IDNO,b.SOCID )" & vbCrLf
            Case Else
                Common.MessageBox(Me, "查詢類型有誤!")
                Exit Sub
        End Select

        sql &= " SELECT ssd.NAME" & vbCrLf
        sql &= " ,ss.IDNO" & vbCrLf
        sql &= " ,format(ss.BIRTHDAY,'yyyy/MM/dd') BIRTHDAY" & vbCrLf
        'sql &= " ,se2.ActNo" & vbCrLf 'sql &= " ,se2.Actname" & vbCrLf
        sql &= " ,b3.ACTNO" & vbCrLf
        sql &= " ,b3.COMNAME" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(cc.CLASSCNAME,cc.CYCLTYPE) CLASSCNAME" & vbCrLf
        'sql &= " --,d.MdateADD" & vbCrLf 'sql &= " --,b.MdateL" & vbCrLf
        sql &= " ,format(b3.MDATE,'yyyy/MM/dd') MDATE" & vbCrLf
        sql &= " ,format(cc.STDATE,'yyyy/MM/dd') STDATE" & vbCrLf
        sql &= " ,format(b3.MODIFYDATE,'yyyy/MM/dd') MODIFYDATE" & vbCrLf
        'sql &= " ,b3.MDATE" & vbCrLf 'sql &= " ,cc.Stdate" & vbCrLf
        sql &= " FROM WC1 cs" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO cc on cc.ocid =cs.ocid" & vbCrLf
        sql &= " JOIN STUD_STUDENTINFO ss ON cs.sid = ss.sid" & vbCrLf
        sql &= " JOIN STUD_SUBDATA ssd ON ss.sid = ssd.sid" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE se ON cs.setid = se.setid AND cs.etenterdate = se.enterdate AND cs.sernum = se.sernum" & vbCrLf
        sql &= " JOIN STUD_ENTERTRAIN2 se2 ON se.esernum = se2.esernum AND se.seid = se2.seid" & vbCrLf
        sql &= " LEFT JOIN WC2 b2 on b2.socid=cs.socid and b2.idno=ss.idno" & vbCrLf
        sql &= " LEFT JOIN STUD_BLIGATEDATA28 b3 on b3.SB2ID=b2.SB2ID" & vbCrLf
        sql &= " WHERE 1=1 " & vbCrLf
        '班級資訊-條件
        If DateStr1 <> "" Then
            '班級資訊-條件
            sql &= DateStr1
        ElseIf DateStr2 <> "" Then
            '學員資訊-條件
            sql &= DateStr2
        Else
            Common.MessageBox(Me, "查詢類型有誤!")
            Exit Sub
        End If

        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        Dim sMemo As String = GET_SEARCH_MEMO()
        '--查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "NAME,IDNO,BIRTHDAY,ACTNO,COMNAME")
        Call TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        Table4.Visible = False
        'msg.Visible = True
        msg.Text = "查無資料!!"
        If dt.Rows.Count > 0 Then
            Table4.Visible = True
            'msg.Visible = True
            msg.Text = ""
            PageControler1.PageDataTable = dt
            PageControler1.ControlerLoad()
        End If

    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Call sSearch1()

    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.AlternatingItem, ListItemType.Item
                e.Item.Cells(0).Text = TIMS.Get_DGSeqNo(sender, e) '序號
        End Select
    End Sub
End Class