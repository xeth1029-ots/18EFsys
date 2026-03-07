Partial Class CP_01_007_add9
    Inherits AuthBasePage

    'CLASS_UNEXPECTTEL
    'CLASS_UNEXPECTTELAPPLY
    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        'InitializeComponent()
        Dim MyValue As Integer = 1000
        Item10_Note.MaxLength = MyValue
        TIMS.Tooltip(Item10_Note, "欄位長度" & CStr(MyValue))
        Item10_Other.MaxLength = MyValue
        TIMS.Tooltip(Item10_Other, "欄位長度" & CStr(MyValue))
    End Sub

    'Dim blnCanAdds As Boolean = False '新增
    'Dim blnCanMod As Boolean = False '修改
    'Dim blnCanDel As Boolean = False '刪除
    'Dim blnCanSech As Boolean = False '查詢
    'Dim blnCanPrnt As Boolean = False '列印
    'UPDATE TABLE : CLASS_UNEXPECTTEL .CLASS_UNEXPECTTELAPPLY 

    Dim dtCOLUMNS As DataTable
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在 End

        dtCOLUMNS = TIMS.Get_USERTABCOLUMNS("CLASS_UNEXPECTTEL,CLASS_UNEXPECTTELAPPLY", objconn)

        If Not IsPostBack Then
            Call cCREATE1()
            Call cCREATE2() '必須放在'create(Request("OCID"), Request("SeqNo")) 之前
            Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
            Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
            If rqOCID <> "" AndAlso rqSeqNo <> "" Then Call SHOW_DATA1(rqOCID, rqSeqNo)
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
        Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

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
        Button1.Attributes("onclick") = "javascript:return chkdata()"

        'If Not IsPostBack Then
        'End If

        Dim rqType As String = TIMS.ClearSQM(Request("Type"))
        Dim rqState As String = TIMS.ClearSQM(Request("State"))

        If rqState = "View" Then Button1.Visible = False

        'Button1.Enabled = False
        'If blnCanAdds Then Button1.Enabled = True

#Region "(No Use)"

        'Dim FunDr As DataRow
        ''檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.RoleID <> 0 Then
        '    If sm.UserInfo.FunDt Is Nothing Then
        '        Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '        Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        '    Else
        '        Dim FunDt As DataTable = sm.UserInfo.FunDt
        '        Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '        If FunDrArray.Length = 0 Then
        '            Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '            Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '        Else
        '            FunDr = FunDrArray(0)
        '            If FunDr("Adds") = 1 Then
        '                Button1.Enabled = True
        '            Else
        '                Button1.Enabled = False
        '            End If
        '        End If
        '    End If
        'End If

#End Region
        '訪視計畫表用
        Button4.Visible = True '回查詢頁面-4
        Button5.Visible = False '回查詢頁面-5
        If rqType = "CT" AndAlso rqState = "View" Then
            Button4.Visible = False '回查詢頁面-4
            Button5.Visible = True '回查詢頁面-5
        ElseIf rqType = "CT" Then
            Button4.Visible = False '回查詢頁面-4
            Button5.Visible = True '回查詢頁面-5
        End If
    End Sub

    Sub cCREATE1()
        If Session("SearchStr") IsNot Nothing Then Me.ViewState("SearchStr") = Session("SearchStr")
        If Session("_SearchStr") IsNot Nothing Then Me.ViewState("_SearchStr") = Session("_SearchStr")
        'Session("SearchStr") = Nothing
        'Session("_SearchStr") = Nothing
        Dim rqRIDValue As String = TIMS.ClearSQM(Request("RIDValue"))
        'Dim rqTMID1 As String = TIMS.ClearSQM(Request("TMID1"))
        'Dim rqOCID1 As String = TIMS.ClearSQM(Request("OCID1"))
        Dim rqTMIDValue1 As String = TIMS.ClearSQM(Request("TMIDValue1"))
        Dim rqOCIDValue1 As String = TIMS.ClearSQM(Request("OCIDValue1"))

        If rqRIDValue <> "" Then RIDValue.Value = rqRIDValue

        'If Request("center") <> "" Then center.Text = Request("center")
        Dim v_pj As String = TIMS.GetMyValue(ViewState("SearchStr"), "pj")
        'fix 機構名稱顯示不出來的問題
        If ViewState("SearchStr") IsNot Nothing AndAlso v_pj = "CP01007" Then
            center.Text = TIMS.UrlDecode1(TIMS.GetMyValue(ViewState("SearchStr"), "center")) 'Replace(MyValue, "%26", "&")
            TMID1.Text = TIMS.UrlDecode1(TIMS.GetMyValue(ViewState("SearchStr"), "TMID1"))
            OCID1.Text = TIMS.UrlDecode1(TIMS.GetMyValue(ViewState("SearchStr"), "OCID1"))
        End If

        'If rqTMID1 <> "" Then TMID1.Text = rqTMID1 'Request("TMID1")
        'If rqOCID1 <> "" Then OCID1.Text = rqOCID1 'Request("OCID1")
        If rqTMIDValue1 <> "" Then TMIDValue1.Value = rqTMIDValue1 'Request("TMIDValue1")
        If rqOCIDValue1 <> "" Then OCIDValue1.Value = rqOCIDValue1 'Request("OCIDValue1")
        VisitorOrgNAME.Text = sm.UserInfo.OrgName
    End Sub

    Sub CreateDataTable1(ByRef dt As DataTable)
        dt = New DataTable
        dt.TableName = "CLASS_UNEXPECTTELQuestion"
        dt.Columns.Add(New DataColumn("ShowItem"))
        dt.Columns.Add(New DataColumn("Item"))
        dt.Columns.Add(New DataColumn("Question"))
        dt.Columns.Add(New DataColumn("Answer"))
        dt.Columns.Add(New DataColumn("checkbox1"))
        dt.Columns.Add(New DataColumn("ckcolumn"))
    End Sub
    Sub CreateDataRow(ByRef dt As DataTable, v_ShowItem As String, v_DataItem As String, v_Question As String, v_Answer As String, v_checkbox1 As String, v_ckcolumn As String)
        Dim dr As DataRow = dt.NewRow
        dt.Rows.Add(dr)
        dr("ShowItem") = v_ShowItem '"7"
        dr("Item") = v_DataItem '"7"
        dr("Question") = v_Question '"上課講師是否有遲到早退？"
        dr("Answer") = v_Answer '"符合,不符合"
        dr("checkbox1") = v_checkbox1 '"尚未上實體課程"
        dr("ckcolumn") = v_ckcolumn '"尚未上實體課程"
    End Sub
    Sub CreateDataRow(ByRef dt As DataTable, v_ShowItem As String, v_DataItem As String, v_Question As String, v_Answer As String)
        CreateDataRow(dt, v_ShowItem, v_DataItem, v_Question, v_Answer, Nothing, Nothing)
    End Sub

    '建立DataGrid表格
    Sub cCREATE2()
        PageControler1.PageDataGrid = DataGrid1
        '建立資料格式Table -- Start
        Dim dt As DataTable = Nothing
        CreateDataTable1(dt)

#Region "(No Use)"
        'dt = New DataTable
        'dt.TableName = "CLASS_UNEXPECTTELQuestion"
        'dt.Columns.Add(New DataColumn("ShowItem"))
        'dt.Columns.Add(New DataColumn("Item"))
        'dt.Columns.Add(New DataColumn("Question"))
        'dt.Columns.Add(New DataColumn("Answer"))
        'dt.Columns.Add(New DataColumn("checkbox1"))

        'insert into CLASS_UNEXPECTTELQuestion select '1','是否有在該單位的班級上課？課程名稱及上課地點？','相符合,不符合'
        'insert into CLASS_UNEXPECTTELQuestion select '2','開訓日期及結訓日期？每周上課時間？總時數？','相符合,不符合'
        'insert into CLASS_UNEXPECTTELQuestion select '3','師資教學品質是否良好？','良好,尚可'
        'insert into CLASS_UNEXPECTTELQuestion select '4','場地及設備品質是否良好？','良好,尚可'
        'insert into CLASS_UNEXPECTTELQuestion select '5','教材是否足夠？','良好,尚可'
        'insert into CLASS_UNEXPECTTELQuestion select '6','最後一次上課時間？上次上課講師姓名或姓式？','相符合,不符合'
        'insert into CLASS_UNEXPECTTELQuestion select '7','上課時間是否有超過一半？','相符合,不符合'
        'insert into CLASS_UNEXPECTTELQuestion select '8','是否有繳交費用？','相符合,不符合'
        'insert into CLASS_UNEXPECTTELQuestion select '9','參訓後對您是否有幫助？','良好,尚可'
#End Region

        CreateDataRow(dt, "A", "A", "聯絡時間", "DTime")
        CreateDataRow(dt, "B", "B", "撥話線路(電話抽訪)", "UsePhone")
        CreateDataRow(dt, "C", "C", "受訪學員姓名", "Name")
        CreateDataRow(dt, "D", "D", "聯絡電話", "Phone")
        CreateDataRow(dt, "E", "E", "身份", "一般,特殊")

        CreateDataRow(dt, "1", "1", "是否有在該單位的班級上課？課程名稱及上課地點？", "符合,不符合")
        CreateDataRow(dt, "2", "2", "開訓日期及結訓日期？每周上課時間？總時數？", "符合,不符合")
        CreateDataRow(dt, "3", "3", "師資教學品質是否良好？", "良好,尚可,待改善")
        CreateDataRow(dt, "4", "4", "實體場地及設備品質是否良好？", "良好,尚可,待改善", "(尚未上實體課程)", "NotENTITY")
        CreateDataRow(dt, "5", "45", "遠距課程流暢度及品質是否良好？", "良好,尚可,待改善", "(尚未上遠距課程)", "NotREMOTELY")
        CreateDataRow(dt, "6", "5", "教材是否足夠？", "良好,尚可,待改善")
        CreateDataRow(dt, "7", "6", "最近一次上課時間？上課講師姓名或姓氏？", "符合,不符合")
        CreateDataRow(dt, "8", "7", "上課講師是否有遲到早退？", "符合,不符合")
        '2010 年以後的資料顯示
        CreateDataRow(dt, "9", "8", "繳交費用額度？", "Money")
        CreateDataRow(dt, "10", "9", "參訓後對您是否有幫助？", "良好,尚可,無幫助")

        '建立資料格式Table----------------End
        PageControler1.DataTableCreate(dt)
        PageControler1.Visible = False
    End Sub

    Sub SHOW_DATA1(ByVal OCID As String, ByVal SeqNo As String)
        Dim parms As New Hashtable From {{"OCID", OCID}}
        Dim sql As String = ""
        sql &= " SELECT c1.OCID ,c1.SEQNO ,c1.APPLYDATE" & vbCrLf
        sql &= " ,c1.ITEM1_1 ,c1.ITEM1_2 ,c1.ITEM1_3 ,c1.ITEM1_4 ,c1.ITEM1_5 ,c1.ITEM1_NOTE" & vbCrLf
        sql &= " ,c1.ITEM2_1 ,c1.ITEM2_2 ,c1.ITEM2_3 ,c1.ITEM2_4 ,c1.ITEM2_5 ,c1.ITEM2_NOTE " & vbCrLf
        sql &= " ,c1.ITEM3_1 ,c1.ITEM3_2 ,c1.ITEM3_3 ,c1.ITEM3_4 ,c1.ITEM3_5 ,c1.ITEM3_NOTE " & vbCrLf
        sql &= " ,c1.ITEM4_1 ,c1.ITEM4_2 ,c1.ITEM4_3 ,c1.ITEM4_4 ,c1.ITEM4_5 ,c1.ITEM4_NOTE " & vbCrLf
        sql &= " ,c1.ITEM45_1 ,c1.ITEM45_2 ,c1.ITEM45_3 ,c1.ITEM45_4 ,c1.ITEM45_5 ,c1.ITEM45_NOTE " & vbCrLf
        sql &= " ,c1.NotENTITY ,c1.NotREMOTELY" & vbCrLf
        sql &= " ,c1.ITEM5_1 ,c1.ITEM5_2 ,c1.ITEM5_3 ,c1.ITEM5_4 ,c1.ITEM5_5 ,c1.ITEM5_NOTE " & vbCrLf
        sql &= " ,c1.ITEM6_1 ,c1.ITEM6_2 ,c1.ITEM6_3 ,c1.ITEM6_4 ,c1.ITEM6_5 ,c1.ITEM6_NOTE " & vbCrLf
        sql &= " ,c1.ITEM7_1 ,c1.ITEM7_2 ,c1.ITEM7_3 ,c1.ITEM7_4 ,c1.ITEM7_5 ,c1.ITEM7_NOTE " & vbCrLf
        sql &= " ,c1.ITEM8_1 ,c1.ITEM8_2 ,c1.ITEM8_3 ,c1.ITEM8_4 ,c1.ITEM8_5 ,c1.ITEM8_NOTE " & vbCrLf
        sql &= " ,c1.ITEM9_1 ,c1.ITEM9_2 ,c1.ITEM9_3 ,c1.ITEM9_4 ,c1.ITEM9_5 ,c1.ITEM9_NOTE" & vbCrLf
        sql &= " ,c1.ITEM10 ,c1.ITEM10_1 ,c1.ITEM10_NOTE ,c1.ITEM10_OTHER ,c1.VISITORNAME " & vbCrLf
        sql &= " ,c1.RID ,c1.ISCLEAR " & vbCrLf
        sql &= " ,c1.ITEM8_1_99 ,c1.ITEM8_2_99 ,c1.ITEM8_3_99 ,c1.ITEM8_4_99 ,c1.ITEM8_5_99" & vbCrLf
        sql &= " ,c1.ORGID" & vbCrLf
        sql &= " ,c2.ITEMA_1 ,c2.ITEMA_2 ,c2.ITEMA_3 ,c2.ITEMA_4 ,c2.ITEMA_5 ,c2.ITEMA_NOTE" & vbCrLf
        sql &= " ,c2.ITEMB_1 ,c2.ITEMB_2 ,c2.ITEMB_3 ,c2.ITEMB_4 ,c2.ITEMB_5 ,c2.ITEMB_NOTE " & vbCrLf
        sql &= " ,c2.ITEMC_1 ,c2.ITEMC_2 ,c2.ITEMC_3 ,c2.ITEMC_4 ,c2.ITEMC_5 ,c2.ITEMC_NOTE" & vbCrLf
        sql &= " ,c2.ITEMD_1 ,c2.ITEMD_2 ,c2.ITEMD_3 ,c2.ITEMD_4 ,c2.ITEMD_5 ,c2.ITEMD_NOTE " & vbCrLf
        sql &= " ,c2.ITEME_1 ,c2.ITEME_2 ,c2.ITEME_3 ,c2.ITEME_4 ,c2.ITEME_5 ,c2.ITEME_NOTE" & vbCrLf
        sql &= " ,c1.TELVISITREASON" & vbCrLf
        sql &= " ,c1.REALVISITDATE" & vbCrLf
        sql &= " FROM CLASS_UNEXPECTTEL c1 " & vbCrLf
        sql &= " LEFT JOIN CLASS_UNEXPECTTELAPPLY c2 ON c1.OCID = c2.OCID AND c1.SeqNO = c2.SeqNO " & vbCrLf
        sql &= " WHERE c1.OCID = @OCID " & vbCrLf
        If SeqNo <> "" Then
            sql &= " AND c1.SEQNO = @SEQNO " & vbCrLf
            parms.Add("SEQNO", SeqNo)
        End If

        'Dim i As Integer = 0
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

        If dr IsNot Nothing Then

            ApplyDate.Text = If(flag_ROC, TIMS.Cdate17(dr("ApplyDate")), TIMS.Cdate3(dr("ApplyDate"))) 'edit，by:20181018

            Common.SetListItem(rbl_TELVISITREASON, Convert.ToString(dr("TELVISITREASON")))

            REALVISITDATE.Text = If(flag_ROC, TIMS.Cdate17(dr("REALVISITDATE")), TIMS.Cdate3(dr("REALVISITDATE"))) 'edit，by:20181018

            For Each item As DataGridItem In DataGrid1.Items
                Dim cb1_show As CheckBox = item.FindControl("cb1_show")
                Dim hid_dataitem As HiddenField = item.FindControl("hid_dataitem")
                Dim hid_ckcolumn As HiddenField = item.FindControl("hid_ckcolumn")
                Dim rdoAnswer1 As RadioButtonList = item.FindControl("rdoAnswer1")
                Dim rdoAnswer2 As RadioButtonList = item.FindControl("rdoAnswer2")
                Dim rdoAnswer3 As RadioButtonList = item.FindControl("rdoAnswer3")
                Dim rdoAnswer4 As RadioButtonList = item.FindControl("rdoAnswer4")
                Dim rdoAnswer5 As RadioButtonList = item.FindControl("rdoAnswer5")
                Dim txtNote As TextBox = item.FindControl("txtNote")
                Dim txtAnswer1 As TextBox = item.FindControl("txtAnswer1")
                Dim txtAnswer2 As TextBox = item.FindControl("txtAnswer2")
                Dim txtAnswer3 As TextBox = item.FindControl("txtAnswer3")
                Dim txtAnswer4 As TextBox = item.FindControl("txtAnswer4")
                Dim txtAnswer5 As TextBox = item.FindControl("txtAnswer5")
                hid_ckcolumn.Value = TIMS.ClearSQM(hid_ckcolumn.Value)
                If hid_ckcolumn.Value <> "" Then cb1_show.Checked = (Convert.ToString(dr(hid_ckcolumn.Value)) = "Y")
                hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value)
                Dim N As String = hid_dataitem.Value '項次
                If N <> "" Then
                    Select Case N
                        Case "A", "B", "C", "D"
                            txtAnswer1.Text = dr("Item" & N & "_1").ToString
                            txtAnswer2.Text = dr("Item" & N & "_2").ToString
                            txtAnswer3.Text = dr("Item" & N & "_3").ToString
                            txtAnswer4.Text = dr("Item" & N & "_4").ToString
                            txtAnswer5.Text = dr("Item" & N & "_5").ToString
                        Case "8"
                            If CInt(sm.UserInfo.Years) >= 2010 Then  '2010年第8題顯示
                                If IsDBNull(dr("Item" & N & "_1_99")) = False Then txtAnswer1.Text = dr("Item" & N & "_1_99")
                                If IsDBNull(dr("Item" & N & "_2_99")) = False Then txtAnswer2.Text = dr("Item" & N & "_2_99")
                                If IsDBNull(dr("Item" & N & "_3_99")) = False Then txtAnswer3.Text = dr("Item" & N & "_3_99")
                                If IsDBNull(dr("Item" & N & "_4_99")) = False Then txtAnswer4.Text = dr("Item" & N & "_4_99")
                                If IsDBNull(dr("Item" & N & "_5_99")) = False Then txtAnswer5.Text = dr("Item" & N & "_5_99")
                            Else
                                If dr("Item" & N & "_1").ToString <> "" Then Common.SetListItem(rdoAnswer1, dr("Item" & N & "_1").ToString)
                                If dr("Item" & N & "_2").ToString <> "" Then Common.SetListItem(rdoAnswer2, dr("Item" & N & "_2").ToString)
                                If dr("Item" & N & "_3").ToString <> "" Then Common.SetListItem(rdoAnswer3, dr("Item" & N & "_3").ToString)
                                If dr("Item" & N & "_4").ToString <> "" Then Common.SetListItem(rdoAnswer4, dr("Item" & N & "_4").ToString)
                                If dr("Item" & N & "_5").ToString <> "" Then Common.SetListItem(rdoAnswer5, dr("Item" & N & "_5").ToString)
                            End If
                        Case Else
                            If dr("Item" & N & "_1").ToString <> "" Then Common.SetListItem(rdoAnswer1, dr("Item" & N & "_1").ToString)
                            If dr("Item" & N & "_2").ToString <> "" Then Common.SetListItem(rdoAnswer2, dr("Item" & N & "_2").ToString)
                            If dr("Item" & N & "_3").ToString <> "" Then Common.SetListItem(rdoAnswer3, dr("Item" & N & "_3").ToString)
                            If dr("Item" & N & "_4").ToString <> "" Then Common.SetListItem(rdoAnswer4, dr("Item" & N & "_4").ToString)
                            If dr("Item" & N & "_5").ToString <> "" Then Common.SetListItem(rdoAnswer5, dr("Item" & N & "_5").ToString)
                    End Select
                    '===================
                    Dim NoteColName As String = String.Concat("Item", N, "_Note")
                    txtNote.Text = dr(NoteColName).ToString

                    'Dim MyValue As Integer = 0
                    'For Each drCol As DataRow In dtCOLUMNS.Rows
                    '    MyValue = drCol("character_maxiMum_length")
                    '    If Convert.ToString(drCol("column_name")) = NoteColName Then
                    '        txtNote.MaxLength = MyValue
                    '        TIMS.Tooltip(txtNote, "欄位長度" & CStr(MyValue))
                    '        Exit For
                    '    End If
                    'Next
                End If
            Next

            'Item10 結論 正常 不正常， 須加以查核  其他附加說明
            'If dr("Item10").ToString <> "" Then Item10.SelectedValue = dr("Item10").ToString
            Common.SetListItem(Item10, Convert.ToString(dr("Item10"))) '1/2
            Item10_1.Checked = If(Convert.ToString(dr("Item10_1")) = "1", True, False)
            Item10_Note.Text = dr("Item10_Note").ToString
            Item10_Other.Text = dr("Item10_Other").ToString
            VisitorName.Text = Convert.ToString(dr("VisitorName"))

            If dr("OrgID").ToString <> "" Then VisitorOrgNAME.Text = TIMS.GET_ORGNAME(dr("OrgID"), objconn)
        End If

        Dim drC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If drC IsNot Nothing Then
            center.Text = drC("ORGNAME")
            RIDValue.Value = drC("RID").ToString
            TMID1.Text = Convert.ToString(drC("TRAINNAME2"))
            TMIDValue1.Value = drC("TMID")
            OCID1.Text = drC("CLASSCNAME2")
            OCIDValue1.Value = drC("OCID")
            center.Enabled = False
            TMID1.Enabled = False
            OCID1.Enabled = False
            Button2.Disabled = True
            Button3.Disabled = True
        End If
    End Sub

    ''' <summary>先取出最大SeqNo</summary>
    ''' <param name="parms"></param>
    ''' <param name="oConn"></param>
    ''' <returns></returns>
    Function GET_UNEXPECTTEL_MAXSEQNO(ByRef parms As Hashtable, ByRef oConn As SqlConnection) As Integer
        Dim iSeqNo As Integer = 0
        'Dim v_rqState As String = TIMS.GetMyValue2(parms, "State") 'Request("State") = "Add"
        Dim v_OCIDValue1 As String = TIMS.GetMyValue2(parms, "OCIDValue1")

        '先取出最大SeqNo
        Dim pms_s1 As New Hashtable From {{"OCID", v_OCIDValue1}}
        Dim sql_s1 As String = " SELECT MAX(SEQNO) NUM FROM CLASS_UNEXPECTTEL WITH(NOLOCK) WHERE OCID=@OCID"
        Dim dr As DataRow = DbAccess.GetOneRow(sql_s1, oConn, pms_s1)
        If dr Is Nothing Then Return iSeqNo '使用MAX 都會回傳資料 沒有資料為異常 回傳為0

        If IsDBNull(dr("NUM")) Then
            iSeqNo = 1
            Return iSeqNo
        End If

        iSeqNo = CInt(dr("NUM")) + 1
        Return iSeqNo
    End Function

    Function rdoDisplay(ByRef rdo As RadioButtonList, ByVal Answer As String, ByVal flag As String) As RadioButtonList
        Dim ary As System.Array
        ary = Split(Answer, flag)
        rdo.Items.Clear()
        For i As Integer = 0 To UBound(ary)
            rdo.Items.Add(ary(i))
            rdo.Items(i).Value = CStr(i + 1)
        Next
        Return rdo
    End Function

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Const CST_Answer1 As Integer = 2
        Const CST_Answer2 As Integer = 3
        Const CST_Answer3 As Integer = 4
        Const CST_Answer4 As Integer = 5
        Const CST_Answer5 As Integer = 6
        Const cst_flag1 As String = "," '逗號 有分隔

        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim cb1_show As CheckBox = e.Item.FindControl("cb1_show")
                Dim hid_dataitem As HiddenField = e.Item.FindControl("hid_dataitem")
                Dim hid_ckcolumn As HiddenField = e.Item.FindControl("hid_ckcolumn")
                Dim labShowItem As Label = e.Item.FindControl("labShowItem")
                Dim labquestion As Label = e.Item.FindControl("labquestion")

                Dim rdoAnswer1 As RadioButtonList = e.Item.Cells(CST_Answer1).FindControl("rdoAnswer1")
                Dim rdoAnswer2 As RadioButtonList = e.Item.Cells(CST_Answer2).FindControl("rdoAnswer2")
                Dim rdoAnswer3 As RadioButtonList = e.Item.Cells(CST_Answer3).FindControl("rdoAnswer3")
                Dim rdoAnswer4 As RadioButtonList = e.Item.Cells(CST_Answer4).FindControl("rdoAnswer4")
                Dim rdoAnswer5 As RadioButtonList = e.Item.Cells(CST_Answer5).FindControl("rdoAnswer5")
                Dim txtAnswer1 As TextBox = e.Item.Cells(CST_Answer1).FindControl("txtAnswer1")
                Dim txtAnswer2 As TextBox = e.Item.Cells(CST_Answer2).FindControl("txtAnswer2")
                Dim txtAnswer3 As TextBox = e.Item.Cells(CST_Answer3).FindControl("txtAnswer3")
                Dim txtAnswer4 As TextBox = e.Item.Cells(CST_Answer4).FindControl("txtAnswer4")
                Dim txtAnswer5 As TextBox = e.Item.Cells(CST_Answer5).FindControl("txtAnswer5")

                hid_dataitem.Value = Convert.ToString(drv("Item"))
                labShowItem.Text = Convert.ToString(drv("ShowItem"))
                labquestion.Text = Convert.ToString(drv("question"))

                If (Convert.ToString(drv("checkbox1")) <> "") AndAlso (Convert.ToString(drv("ckcolumn")) <> "") Then
                    cb1_show.Visible = True
                    cb1_show.Text = Convert.ToString(drv("checkbox1"))
                    hid_ckcolumn.Value = Convert.ToString(drv("ckcolumn"))
                    Dim js_cbl1A As String = String.Concat("clear_rdoAnswer('#", cb1_show.ClientID, "', '#", rdoAnswer1.ClientID, "');")
                    'Dim js_rdoA1 As String = String.Concat("clear_cb1_show('#", cb1_show.ClientID, "', '#", rdoAnswer1.ClientID, "');")
                    cb1_show.Attributes("onclick") = js_cbl1A
                    'rdoAnswer1.Attributes("onclick") = js_rdoA1
                End If

                If drv("Answer").ToString.IndexOf(cst_flag1) <> -1 Then
                    rdoAnswer1 = rdoDisplay(rdoAnswer1, drv("Answer").ToString, cst_flag1)
                    rdoAnswer2 = rdoDisplay(rdoAnswer2, drv("Answer").ToString, cst_flag1)
                    rdoAnswer3 = rdoDisplay(rdoAnswer3, drv("Answer").ToString, cst_flag1)
                    rdoAnswer4 = rdoDisplay(rdoAnswer4, drv("Answer").ToString, cst_flag1)
                    rdoAnswer5 = rdoDisplay(rdoAnswer5, drv("Answer").ToString, cst_flag1)
                    txtAnswer1.Visible = False
                    txtAnswer2.Visible = False
                    txtAnswer3.Visible = False
                    txtAnswer4.Visible = False
                    txtAnswer5.Visible = False
                Else
                    Const cst_MoneyFormat As String = "請填寫數字格式"
                    Select Case Convert.ToString(drv("Answer"))
                        Case "Money"
                            TIMS.Tooltip(txtAnswer1, cst_MoneyFormat)
                            TIMS.Tooltip(txtAnswer2, cst_MoneyFormat)
                            TIMS.Tooltip(txtAnswer3, cst_MoneyFormat)
                            TIMS.Tooltip(txtAnswer4, cst_MoneyFormat)
                            TIMS.Tooltip(txtAnswer5, cst_MoneyFormat)
                    End Select
                End If

                'Dim ShowItem As String = Convert.ToString(drv("ShowItem"))
                Dim N As String = Convert.ToString(drv("Item"))
                Dim txtNote As TextBox = e.Item.FindControl("txtNote")
                Dim s_NoteColName As String = String.Concat("Item", N, "_Note")
                Dim MyValue As Integer = 0
                For Each drCol As DataRow In dtCOLUMNS.Rows
                    MyValue = drCol("CHARACTER_MAXIMUM_LENGTH")
                    If UCase(drCol("COLUMN_NAME")) = UCase(s_NoteColName) Then
                        txtNote.MaxLength = MyValue
                        TIMS.Tooltip(txtNote, "欄位長度" & CStr(MyValue))
                        Exit For
                    End If
                Next
        End Select
    End Sub

    Sub CHECK_CLASS_UNEXPECTTEL(ByRef ErrMsg As String)
        ErrMsg = ""
        Dim fg_cb4 As Boolean = False
        Dim fg_cb5 As Boolean = False
        '依欄位名稱，修改資料
        For Each item As DataGridItem In DataGrid1.Items
            Dim labShowItem As Label = item.FindControl("labShowItem")
            Dim cb1_show As CheckBox = item.FindControl("cb1_show")
            Dim hid_dataitem As HiddenField = item.FindControl("hid_dataitem")
            Dim hid_ckcolumn As HiddenField = item.FindControl("hid_ckcolumn")
            Dim rdoAnswer1 As RadioButtonList = item.FindControl("rdoAnswer1")
            Dim rdoAnswer2 As RadioButtonList = item.FindControl("rdoAnswer2")
            Dim rdoAnswer3 As RadioButtonList = item.FindControl("rdoAnswer3")
            Dim rdoAnswer4 As RadioButtonList = item.FindControl("rdoAnswer4")
            Dim rdoAnswer5 As RadioButtonList = item.FindControl("rdoAnswer5")
            Dim txtNote As TextBox = item.FindControl("txtNote")
            Dim txtAnswer1 As TextBox = item.FindControl("txtAnswer1")
            Dim txtAnswer2 As TextBox = item.FindControl("txtAnswer2")
            Dim txtAnswer3 As TextBox = item.FindControl("txtAnswer3")
            Dim txtAnswer4 As TextBox = item.FindControl("txtAnswer4")
            Dim txtAnswer5 As TextBox = item.FindControl("txtAnswer5")
            txtAnswer1.Text = TIMS.ClearSQM(txtAnswer1.Text)
            txtAnswer2.Text = TIMS.ClearSQM(txtAnswer2.Text)
            txtAnswer3.Text = TIMS.ClearSQM(txtAnswer3.Text)
            txtAnswer4.Text = TIMS.ClearSQM(txtAnswer4.Text)
            txtAnswer5.Text = TIMS.ClearSQM(txtAnswer5.Text)
            'Dim N As String = TIMS.ClearSQM(item.Cells(0).Text) '項次
            hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value) '項次
            labShowItem.Text = TIMS.ClearSQM(labShowItem.Text) '項次
            Dim showN As String = labShowItem.Text 'hid_dataitem.Value '項次
            If showN <> "" AndAlso IsNumeric(showN) Then '該迴圈只處理數字
                If showN = "9" Then '2024:第9題 And CInt(sm.UserInfo.Years) >= 2010  2010年的第8題
                    If Convert.ToString(txtAnswer1.Text) <> "" Then
                        If Not IsNumeric(txtAnswer1.Text) Then
                            ErrMsg += "項次" & showN & "(繳交費用額度-訪問一)必須為數字\n"
                        End If
                    Else
                        ErrMsg += "項次" & showN & "(繳交費用額度-訪問一)不可為空\n"
                    End If
                    If Convert.ToString(txtAnswer2.Text) <> "" Then
                        If Not IsNumeric(txtAnswer2.Text) Then
                            ErrMsg += "項次" & showN & "(繳交費用額度-訪問二)必須為數字\n"
                        End If
                    End If
                    If Convert.ToString(txtAnswer3.Text) <> "" Then
                        If Not IsNumeric(txtAnswer3.Text) Then
                            ErrMsg += "項次" & showN & "(繳交費用額度-訪問三)必須為數字\n"
                        End If
                    End If
                    If Convert.ToString(txtAnswer4.Text) <> "" Then
                        If Not IsNumeric(txtAnswer4.Text) Then
                            ErrMsg += "項次" & showN & "(繳交費用額度-訪問四)必須為數字\n"
                        End If
                    End If
                    If Convert.ToString(txtAnswer5.Text) <> "" Then
                        If Not IsNumeric(txtAnswer5.Text) Then
                            ErrMsg += "項次" & showN & "(繳交費用額度-訪問五)必須為數字\n"
                        End If
                    End If
                ElseIf showN = "4" Then ' 第4題
                    If cb1_show.Visible AndAlso cb1_show.Text <> "" AndAlso hid_ckcolumn.Value <> "" Then
                        fg_cb4 = cb1_show.Checked
                        If rdoAnswer1 IsNot Nothing AndAlso rdoAnswer1.Visible AndAlso TIMS.GetListValue(rdoAnswer1) = "" Then
                            If Not cb1_show.Checked Then ErrMsg += "項次" & showN & String.Format("的訪問一不可為空(或請勾選 {0})\n", cb1_show.Text)
                        End If
                    End If
                ElseIf showN = "5" Then ' 第5題
                    If cb1_show.Visible AndAlso cb1_show.Text <> "" AndAlso hid_ckcolumn.Value <> "" Then
                        fg_cb5 = cb1_show.Checked
                        If rdoAnswer1 IsNot Nothing AndAlso rdoAnswer1.Visible AndAlso TIMS.GetListValue(rdoAnswer1) = "" Then
                            If Not cb1_show.Checked Then ErrMsg += "項次" & showN & String.Format("的訪問一不可為空(或請勾選 {0})\n", cb1_show.Text)
                        End If
                    End If

                Else
                    If rdoAnswer1.SelectedIndex = -1 OrElse rdoAnswer1.SelectedValue = "" Then
                        ErrMsg += "項次" & showN & "的訪問一不可為空\n"
                    End If
                End If
                txtNote.Text = TIMS.ClearSQM(txtNote.Text)
                If txtNote.Text.Trim.Length > 100 Then ErrMsg += "項次" & showN & "的備註／說明事項 超過系統長度100\n"
                'If ErrMsg = "" Then dr("Item" & N & "_Note") = txtNote.Text
            End If
        Next
        'Item10 結論 正常 不正常， 須加以查核  其他附加說明
        Dim v_Item10 As String = TIMS.GetListValue(Item10)
        If v_Item10 = "" Then ErrMsg += "結論 (正常 不正常) 請擇一勾選\n"

        If fg_cb4 AndAlso fg_cb5 Then
            ErrMsg += "項次4, 項次5 不可同時勾選\n"
        End If
    End Sub

    Function CHECK_CLASS_UNEXPECTTELAPPLY() As String
        Dim rst As String = ""
        'If OCID = "" Then Return rst
        'If SeqNo = "" Then Return rst
        Dim ErrMsg As String = ""
        For Each item As DataGridItem In DataGrid1.Items
            Dim labShowItem As Label = item.FindControl("labShowItem")
            Dim cb1_show As CheckBox = item.FindControl("cb1_show")
            Dim hid_dataitem As HiddenField = item.FindControl("hid_dataitem")
            Dim hid_ckcolumn As HiddenField = item.FindControl("hid_ckcolumn")
            Dim rdoAnswer1 As RadioButtonList = item.FindControl("rdoAnswer1")
            Dim rdoAnswer2 As RadioButtonList = item.FindControl("rdoAnswer2")
            Dim rdoAnswer3 As RadioButtonList = item.FindControl("rdoAnswer3")
            Dim rdoAnswer4 As RadioButtonList = item.FindControl("rdoAnswer4")
            Dim rdoAnswer5 As RadioButtonList = item.FindControl("rdoAnswer5")
            Dim txtNote As TextBox = item.FindControl("txtNote")
            Dim txtAnswer1 As TextBox = item.FindControl("txtAnswer1")
            Dim txtAnswer2 As TextBox = item.FindControl("txtAnswer2")
            Dim txtAnswer3 As TextBox = item.FindControl("txtAnswer3")
            Dim txtAnswer4 As TextBox = item.FindControl("txtAnswer4")
            Dim txtAnswer5 As TextBox = item.FindControl("txtAnswer5")
            'Dim N As String = TIMS.ClearSQM(item.Cells(0).Text) '項次
            hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value) '項次
            labShowItem.Text = TIMS.ClearSQM(labShowItem.Text) '項次
            Dim showN As String = labShowItem.Text 'hid_dataitem.Value '項次
            If showN = "" Then Exit For '不可為空
            If IsNumeric(showN) Then Continue For '該迴圈只處理非數字
            Select Case showN
                Case "E"
                    If rdoAnswer1.SelectedIndex = -1 OrElse rdoAnswer1.SelectedValue = "" Then
                        ErrMsg += "項次" & showN & "的訪問一不可為空\n"
                    End If
                Case Else
                    txtAnswer1.Text = TIMS.ClearSQM(txtAnswer1.Text)
                    txtAnswer2.Text = TIMS.ClearSQM(txtAnswer2.Text)
                    txtAnswer3.Text = TIMS.ClearSQM(txtAnswer3.Text)
                    txtAnswer4.Text = TIMS.ClearSQM(txtAnswer4.Text)
                    txtAnswer5.Text = TIMS.ClearSQM(txtAnswer5.Text)
                    If txtAnswer1.Text = "" Then
                        ErrMsg += "項次" & showN & "的訪問一不可為空\n"
                    End If
            End Select
            txtNote.Text = TIMS.ClearSQM(txtNote.Text)
            If txtNote.Text.Trim.Length > 100 Then ErrMsg &= "項次" & showN & "的備註／說明事項 超過系統長度100\n"
        Next
        If ErrMsg <> "" Then
            rst = ErrMsg
            Return rst '異常離開
        End If
        Return rst
    End Function

    Sub SAVE_CLASS_UNEXPECTTEL(ByRef MyPage As Page, ByRef parms As Hashtable, ByRef oConn As SqlConnection, ByRef iSeqNo As Integer)
        Dim v_rqState As String = TIMS.GetMyValue2(parms, "State") 'Request("State") = "Add"
        Dim v_OCIDValue1 As String = TIMS.GetMyValue2(parms, "OCIDValue1") 'OCIDValue1.value/Request("OCID") 
        Dim v_rqSEQNO As String = TIMS.GetMyValue2(parms, "SeqNo") 'Request("SeqNo")

        Dim sql As String = ""
        Dim ErrMsg As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        'Dim iSeqNo As Integer = 0
        'Dim iSeqNo As Integer = GET_UNEXPECTTEL_MAXSEQNO(parms, oConn)
        Dim v_TELVISITREASON As String = TIMS.GetListValue(rbl_TELVISITREASON)
        ErrMsg = ""
        If v_rqState = "Add" Then '表示新增狀態
            '先取出最大SeqNo'OCIDValue1
            iSeqNo = GET_UNEXPECTTEL_MAXSEQNO(parms, oConn)

            sql = " SELECT * FROM CLASS_UNEXPECTTEL WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, da, oConn)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = v_OCIDValue1
            dr("SeqNo") = iSeqNo
        Else
            '修改
            sql = "SELECT * FROM CLASS_UNEXPECTTEL WHERE OCID='" & v_OCIDValue1 & "' and SeqNo='" & v_rqSEQNO & "'"
            dt = DbAccess.GetDataTable(sql, da, oConn)
            If dt.Rows.Count <> 1 Then
                Common.MessageBox(MyPage, "資料異常，請重新查詢!")
                Exit Sub
            End If
            dr = dt.Rows(0)
            iSeqNo = Val(v_rqSEQNO) 'Request("SeqNo")
        End If

        ApplyDate.Text = TIMS.ClearSQM(ApplyDate.Text)
        REALVISITDATE.Text = TIMS.ClearSQM(REALVISITDATE.Text)
        dr("ApplyDate") = If(ApplyDate.Text <> "", If(flag_ROC, TIMS.Cdate18(ApplyDate.Text), TIMS.Cdate2(ApplyDate.Text)), TIMS.Cdate2(TIMS.GetSysDate(oConn)))  'edit，by:20181018

        dr("REALVISITDATE") = If(REALVISITDATE.Text <> "", If(flag_ROC, TIMS.Cdate18(REALVISITDATE.Text), TIMS.Cdate2(REALVISITDATE.Text)), Convert.DBNull)
        'Dim v_TELVISITREASON As String = TIMS.ClearSQM(rbl_TELVISITREASON.SelectedValue)
        dr("TELVISITREASON") = If(v_TELVISITREASON <> "", v_TELVISITREASON, Convert.DBNull)

        '依欄位名稱，修改資料
        For Each item As DataGridItem In DataGrid1.Items
            Dim cb1_show As CheckBox = item.FindControl("cb1_show")
            Dim hid_dataitem As HiddenField = item.FindControl("hid_dataitem")
            Dim hid_ckcolumn As HiddenField = item.FindControl("hid_ckcolumn")
            Dim rdoAnswer1 As RadioButtonList = item.FindControl("rdoAnswer1")
            Dim rdoAnswer2 As RadioButtonList = item.FindControl("rdoAnswer2")
            Dim rdoAnswer3 As RadioButtonList = item.FindControl("rdoAnswer3")
            Dim rdoAnswer4 As RadioButtonList = item.FindControl("rdoAnswer4")
            Dim rdoAnswer5 As RadioButtonList = item.FindControl("rdoAnswer5")
            Dim txtNote As TextBox = item.FindControl("txtNote")
            Dim txtAnswer1 As TextBox = item.FindControl("txtAnswer1")
            Dim txtAnswer2 As TextBox = item.FindControl("txtAnswer2")
            Dim txtAnswer3 As TextBox = item.FindControl("txtAnswer3")
            Dim txtAnswer4 As TextBox = item.FindControl("txtAnswer4")
            Dim txtAnswer5 As TextBox = item.FindControl("txtAnswer5")
            'Dim N As String = TIMS.ClearSQM(item.Cells(0).Text) '項次
            hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value) '項次
            Dim N As String = hid_dataitem.Value '項次
            '該迴圈只處理數字
            If N <> "" AndAlso IsNumeric(N) Then '該迴圈只處理數字
                If N = "8" AndAlso CInt(sm.UserInfo.Years) >= 2010 Then '2010年的第8題-8:繳交費用額度？
                    dr("Item" & N & "_1_99") = If(txtAnswer1.Text <> "", CInt(txtAnswer1.Text), Convert.DBNull)
                    dr("Item" & N & "_2_99") = If(txtAnswer2.Text <> "", CInt(txtAnswer2.Text), Convert.DBNull)
                    dr("Item" & N & "_3_99") = If(txtAnswer3.Text <> "", CInt(txtAnswer3.Text), Convert.DBNull)
                    dr("Item" & N & "_4_99") = If(txtAnswer4.Text <> "", CInt(txtAnswer4.Text), Convert.DBNull)
                    dr("Item" & N & "_5_99") = If(txtAnswer5.Text <> "", CInt(txtAnswer5.Text), Convert.DBNull)
                Else
                    dr("Item" & N & "_1") = If(rdoAnswer1.SelectedValue <> "", rdoAnswer1.SelectedValue, Convert.DBNull)
                    dr("Item" & N & "_2") = If(rdoAnswer2.SelectedValue <> "", rdoAnswer2.SelectedValue, Convert.DBNull)
                    dr("Item" & N & "_3") = If(rdoAnswer3.SelectedValue <> "", rdoAnswer3.SelectedValue, Convert.DBNull)
                    dr("Item" & N & "_4") = If(rdoAnswer4.SelectedValue <> "", rdoAnswer4.SelectedValue, Convert.DBNull)
                    dr("Item" & N & "_5") = If(rdoAnswer5.SelectedValue <> "", rdoAnswer5.SelectedValue, Convert.DBNull)
                End If
                txtNote.Text = TIMS.ClearSQM(txtNote.Text)
                dr("Item" & N & "_Note") = txtNote.Text

                If cb1_show.Visible AndAlso cb1_show.Text <> "" AndAlso hid_ckcolumn.Value <> "" Then
                    dr(hid_ckcolumn.Value) = If(cb1_show.Checked, "Y", Convert.DBNull)
                End If
            End If
        Next

        'Item10 結論 正常 不正常， 須加以查核  其他附加說明
        dr("Item10") = If(Item10.SelectedValue <> "", Item10.SelectedValue, Convert.DBNull)
        '其他附加說明 1 have 2:no have
        dr("Item10_1") = If(Item10_1.Checked, "1", "2")
        dr("Item10_Note") = If(Item10_Note.Text = "", Convert.DBNull, Item10_Note.Text)
        dr("Item10_Other") = If(Item10_Other.Text = "", Convert.DBNull, Item10_Other.Text)
        dr("OrgID") = sm.UserInfo.OrgID
        dr("VisitorName") = VisitorName.Text
        dr("RID") = sm.UserInfo.RID
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    Function SAVE_CLASS_UNEXPECTTELAPPLY(ByVal OCID As String, ByVal SeqNo As String) As String
        Dim rst As String = ""

        If OCID = "" OrElse SeqNo = "" Then Return rst

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        'Dim ErrMsg As String = ""

        Dim da As New SqlDataAdapter
        sql = " SELECT * FROM CLASS_UNEXPECTTELAPPLY WHERE OCID='" & OCID & "' AND SeqNo='" & SeqNo & "'"
        dt = DbAccess.GetDataTable(sql, da, objconn)
        If dt.Rows.Count > 0 Then
            dr = dt.Rows(0)
        Else
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = OCID
            dr("SeqNo") = SeqNo
        End If

        For Each item As DataGridItem In DataGrid1.Items
            Dim cb1_show As CheckBox = item.FindControl("cb1_show")
            Dim hid_dataitem As HiddenField = item.FindControl("hid_dataitem")
            Dim hid_ckcolumn As HiddenField = item.FindControl("hid_ckcolumn")
            Dim rdoAnswer1 As RadioButtonList = item.FindControl("rdoAnswer1")
            Dim rdoAnswer2 As RadioButtonList = item.FindControl("rdoAnswer2")
            Dim rdoAnswer3 As RadioButtonList = item.FindControl("rdoAnswer3")
            Dim rdoAnswer4 As RadioButtonList = item.FindControl("rdoAnswer4")
            Dim rdoAnswer5 As RadioButtonList = item.FindControl("rdoAnswer5")
            Dim txtNote As TextBox = item.FindControl("txtNote")
            Dim txtAnswer1 As TextBox = item.FindControl("txtAnswer1")
            Dim txtAnswer2 As TextBox = item.FindControl("txtAnswer2")
            Dim txtAnswer3 As TextBox = item.FindControl("txtAnswer3")
            Dim txtAnswer4 As TextBox = item.FindControl("txtAnswer4")
            Dim txtAnswer5 As TextBox = item.FindControl("txtAnswer5")
            'Dim s_N As String = TIMS.ClearSQM(item.Cells(0).Text) '項次
            hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value) '項次
            Dim s_N As String = hid_dataitem.Value '項次
            If s_N = "" Then Exit For '不可為空
            If IsNumeric(s_N) Then Continue For '該迴圈只處理非數字
            Select Case s_N
                Case "E" 'E	身份
                    dr("Item" & s_N & "_1") = If(rdoAnswer1.SelectedValue <> "", rdoAnswer1.SelectedValue, Convert.DBNull)
                    dr("Item" & s_N & "_2") = If(rdoAnswer2.SelectedValue <> "", rdoAnswer2.SelectedValue, Convert.DBNull)
                    dr("Item" & s_N & "_3") = If(rdoAnswer3.SelectedValue <> "", rdoAnswer3.SelectedValue, Convert.DBNull)
                    dr("Item" & s_N & "_4") = If(rdoAnswer4.SelectedValue <> "", rdoAnswer4.SelectedValue, Convert.DBNull)
                    dr("Item" & s_N & "_5") = If(rdoAnswer5.SelectedValue <> "", rdoAnswer5.SelectedValue, Convert.DBNull)
                Case Else
                    dr("Item" & s_N & "_1") = If(txtAnswer1.Text <> "", txtAnswer1.Text, Convert.DBNull)
                    dr("Item" & s_N & "_2") = If(txtAnswer2.Text <> "", txtAnswer2.Text, Convert.DBNull)
                    dr("Item" & s_N & "_3") = If(txtAnswer3.Text <> "", txtAnswer3.Text, Convert.DBNull)
                    dr("Item" & s_N & "_4") = If(txtAnswer4.Text <> "", txtAnswer4.Text, Convert.DBNull)
                    dr("Item" & s_N & "_5") = If(txtAnswer5.Text <> "", txtAnswer5.Text, Convert.DBNull)
            End Select
            txtNote.Text = TIMS.ClearSQM(txtNote.Text)
            dr("Item" & s_N & "_Note") = If(txtNote.Text <> "", txtNote.Text, Convert.DBNull)
        Next
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        'If ErrMsg <> "" Then'    rst = ErrMsg'    Return rst '異常離開'End If
        DbAccess.UpdateDataTable(dt, da)
        Return rst
    End Function

    ''' <summary>'儲存</summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim s_ERRMSG As String = ""
        Call CHECK_CLASS_UNEXPECTTEL(s_ERRMSG)
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, Replace(s_ERRMSG, "\n", vbCrLf))
            Exit Sub
        End If
        'Dim s_ERRMSG As String = ""
        s_ERRMSG += CHECK_CLASS_UNEXPECTTELAPPLY()
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, Replace(s_ERRMSG, "\n", vbCrLf))
            Exit Sub
        End If
        Dim rqState As String = TIMS.ClearSQM(Request("State"))
        Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
        Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))

        Dim iSeqNo As Integer = 0
        '表示新增狀態
        Dim v_OCIDValue1 As String = If(rqState = "Add", OCIDValue1.Value, rqOCID)

        Dim parms As New Hashtable From {{"State", rqState}, {"OCIDValue1", v_OCIDValue1}, {"SeqNo", rqSeqNo}}
        Dim errFlag As Boolean = False
        Try
            Call SAVE_CLASS_UNEXPECTTEL(Me, parms, objconn, iSeqNo)
            Call SAVE_CLASS_UNEXPECTTELAPPLY(v_OCIDValue1, CStr(iSeqNo))
        Catch ex As Exception
            Dim strErrmsg As String = ""
            strErrmsg &= String.Format("rqState-v_OCIDValue1-rqSeqNo:{0}-{1}-{2}", rqState, v_OCIDValue1, rqSeqNo) & vbCrLf
            strErrmsg &= String.Format("sm.UserID:{0}", sm.UserInfo.UserID) & vbCrLf
            strErrmsg &= TIMS.GetErrorMsg(Me, ex) '取得錯誤資訊寫入
            strErrmsg &= "ex.ToString : " & ex.ToString & vbCrLf
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg, ex)
            errFlag = True
        End Try
        If errFlag Then
            '有錯誤發生!
            Common.MessageBox(Me, "儲存失敗，請重新操作!!(發生異常)")
            Exit Sub
        End If

        If Session("SearchStr") Is Nothing Then Session("SearchStr") = Me.ViewState("SearchStr")
        If Session("_SearchStr") Is Nothing Then Session("_SearchStr") = Me.ViewState("_SearchStr")
        Common.MessageBox(Me, "儲存成功")

        Dim rqType As String = TIMS.ClearSQM(Request("Type"))
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim s_Url1 As String = "CP_01_007.aspx?ID=" & RqID
        If rqType = "CT" Then s_Url1 = "CP_01_008.aspx?ID=" & RqID
        TIMS.Utl_Redirect1(Me, s_Url1)
    End Sub

    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        Dim RqID As String = TIMS.Get_MRqID(Me)
        If Session("SearchStr") Is Nothing Then Session("SearchStr") = Me.ViewState("SearchStr")
        If Session("_SearchStr") Is Nothing Then Session("_SearchStr") = Me.ViewState("_SearchStr")
        Dim s_Url1 As String = "CP_01_007.aspx?ID=" & RqID
        TIMS.Utl_Redirect1(Me, s_Url1)
    End Sub

    Private Sub Button5_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.ServerClick
        Dim RqID As String = TIMS.Get_MRqID(Me)
        '訪視計畫表用
        If Session("SearchStr") Is Nothing Then Session("SearchStr") = Me.ViewState("SearchStr")
        If Session("_SearchStr") Is Nothing Then Session("_SearchStr") = Me.ViewState("_SearchStr")
        Dim s_Url1 As String = "CP_01_008.aspx?ID=" & RqID
        TIMS.Utl_Redirect1(Me, s_Url1)
    End Sub

End Class
