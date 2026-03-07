Partial Class CP_01_006_add9t
    Inherits AuthBasePage

    'CLASS_UNEXPECTVISITOR
    'CLASS_UNEXPECTVISITORAPPLY
    '不預告實地訪查紀錄表-抽訪學員紀錄

    'UPDATE TABLE : CLASS_UNEXPECTTEL .CLASS_UNEXPECTTELAPPLY 
    Const cst_printFN1 As String = "CP_01_006_D2" '2024 '"CP_01_006_D1" 'PRINT 有資料

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        'InitializeComponent()

        Dim MyValue As Integer = 1000
        Item10_Note.MaxLength = MyValue
        TIMS.Tooltip(Item10_Note, "欄位長度" & CStr(MyValue))
        'Item10_Other.MaxLength = MyValue
        'TIMS.Tooltip(Item10_Other, "欄位長度" & CStr(MyValue))
    End Sub

    Const cst_TABLE_NAME_APPLY_1 As String = "CLASS_UNEXPECTVISITORAPPLY"
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

        dtCOLUMNS = TIMS.Get_USERTABCOLUMNS("CLASS_UNEXPECTVISITORAPPLY", objconn)

        If Session("SearchStr") IsNot Nothing Then ViewState("SearchStr") = Session("SearchStr")
        If Not IsPostBack Then
            Call cCREATE1()
            Call cCREATE2() '必須放在'create(Request("OCID"), Request("SeqNo")) 之前
            Dim rqOCID As String = TIMS.ClearSQM(Request("OCID"))
            Dim rqSeqNo As String = TIMS.ClearSQM(Request("SeqNo"))
            Call SHOW_DATA1(rqOCID, rqSeqNo)
        End If

        BtnSave1.Attributes("onclick") = "javascript:return save_chkdata()"

    End Sub

    ''' <summary>SearchStr</summary>
    Sub cCREATE1()
        'Dim rqOCID As String = TIMS.ClearSQM(Request("OCID")) 'OCID
        'Dim rqSEQNO As String = TIMS.ClearSQM(Request("SeqNo"))
        ''Dim rqState As String = TIMS.ClearSQM(Request("State"))
        'If rqOCID = "" OrElse rqSEQNO = "" Then Exit Sub

        'fix 機構名稱顯示不出來的問題
        If ViewState("SearchStr") IsNot Nothing Then
            Dim str_SearchStr As String = Convert.ToString(ViewState("SearchStr"))
            Dim MyValue As String = ""
            MyValue = TIMS.GetMyValue(str_SearchStr, "center")
            center.Text = TIMS.ClearSQM(Replace(MyValue, "%26", "&"))
            MyValue = TIMS.GetMyValue(str_SearchStr, "TMID1")
            TMID1.Text = TIMS.ClearSQM(Replace(MyValue, "%26", "&"))
            MyValue = TIMS.GetMyValue(str_SearchStr, "OCID1")
            OCID1.Text = TIMS.ClearSQM(Replace(MyValue, "%26", "&"))
            TMIDValue1.Value = TIMS.GetMyValue(str_SearchStr, "TMIDValue1")
            OCIDValue1.Value = TIMS.GetMyValue(str_SearchStr, "OCIDValue1")
            labSFDATE_TW.Text = TIMS.GetMyValue(str_SearchStr, "SFDATE_TW")

        End If

    End Sub

    Sub CreateDataTable1(ByRef dt As DataTable)
        dt = New DataTable
        'Const cst_TABLE_NAME_APPLY_1 As String = "CLASS_UNEXPECTVISITORAPPLY"
        dt.TableName = cst_TABLE_NAME_APPLY_1 'x"CLASS_UNEXPECTTELQuestion"
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

    ''' <summary> 建立DataGrid表格 </summary>
    Sub cCREATE2()
        'Dim rqOCID As String = TIMS.ClearSQM(Request("OCID")) 'OCID
        'Dim rqSEQNO As String = TIMS.ClearSQM(Request("SeqNo"))
        'Dim rqState As String = TIMS.ClearSQM(Request("State"))
        'If rqOCID = "" OrElse rqSEQNO = "" Then Exit Sub

        'CLASS_UNEXPECTVISITOR 'PageControler1.PageDataGrid = DataGrid1
        '建立資料格式Table -- Start
        Dim dt As DataTable = Nothing
        CreateDataTable1(dt)

        CreateDataRow(dt, "A", "A", "受訪學員姓名", "lockName")
        CreateDataRow(dt, "B", "B", "身份", "一般,特殊")
        CreateDataRow(dt, "1", "1", "是否有在該單位的班級上課？課程名稱及上課地點？", "符合,不符合")
        CreateDataRow(dt, "2", "2", "開訓日期及結訓日期？每周上課時間？總時數？", "符合,不符合")
        CreateDataRow(dt, "3", "3", "師資教學品質是否良好？", "良好,尚可,待改善")
        CreateDataRow(dt, "4", "4", "實體場地及設備品質是否良好？", "良好,尚可,待改善", "(尚未上實體課程)", "NotENTITY")
        CreateDataRow(dt, "5", "45", "遠距課程流暢度及品質是否良好？", "良好,尚可,待改善", "(尚未上遠距課程)", "NotREMOTELY")
        CreateDataRow(dt, "6", "5", "教材是否足夠？", "良好,尚可,待改善")
        CreateDataRow(dt, "7", "6", "最近一次上課時間？上課講師姓名或姓氏？", "符合,不符合")
        CreateDataRow(dt, "8", "7", "上課講師是否有遲到早退？", "符合,不符合")
        CreateDataRow(dt, "9", "8", "繳交費用額度？", "Money")
        CreateDataRow(dt, "10", "9", "參訓後對您是否有幫助？", "良好,尚可,無幫助")

        '建立資料格式Table----------------End
        'PageControler1.DataTableCreate(dt)
        'PageControler1.Visible = False
        DataGrid1.DataSource = dt
        DataGrid1.DataBind()
    End Sub

    Sub SHOW_DATA1(ByVal OCID As String, ByVal SeqNo As String)
        Dim parms As New Hashtable From {{"OCID", OCID}, {"SEQNO", SeqNo}}
        Dim sql As String = ""
        sql &= " SELECT a.OCID ,a.SEQNO ,a.APPLYDATE" & vbCrLf
        sql &= " ,c1.ITEM1_1 ,c1.ITEM1_2 ,c1.ITEM1_3 ,c1.ITEM1_4 ,c1.ITEM1_5 ,c1.ITEM1_NOTE " & vbCrLf
        sql &= " ,c1.ITEM2_1 ,c1.ITEM2_2 ,c1.ITEM2_3 ,c1.ITEM2_4 ,c1.ITEM2_5 ,c1.ITEM2_NOTE " & vbCrLf
        sql &= " ,c1.ITEM3_1 ,c1.ITEM3_2 ,c1.ITEM3_3 ,c1.ITEM3_4 ,c1.ITEM3_5 ,c1.ITEM3_NOTE " & vbCrLf
        sql &= " ,c1.ITEM4_1 ,c1.ITEM4_2 ,c1.ITEM4_3 ,c1.ITEM4_4 ,c1.ITEM4_5 ,c1.ITEM4_NOTE " & vbCrLf
        sql &= " ,c1.ITEM45_1 ,c1.ITEM45_2 ,c1.ITEM45_3 ,c1.ITEM45_4 ,c1.ITEM45_5 ,c1.ITEM45_NOTE " & vbCrLf
        sql &= " ,c1.NotENTITY ,c1.NotREMOTELY" & vbCrLf
        sql &= " ,c1.ITEM5_1 ,c1.ITEM5_2 ,c1.ITEM5_3 ,c1.ITEM5_4 ,c1.ITEM5_5 ,c1.ITEM5_NOTE " & vbCrLf
        sql &= " ,c1.ITEM6_1 ,c1.ITEM6_2 ,c1.ITEM6_3 ,c1.ITEM6_4 ,c1.ITEM6_5 ,c1.ITEM6_NOTE " & vbCrLf
        sql &= " ,c1.ITEM7_1 ,c1.ITEM7_2 ,c1.ITEM7_3 ,c1.ITEM7_4 ,c1.ITEM7_5 ,c1.ITEM7_NOTE " & vbCrLf
        sql &= " ,c1.ITEM8_1_99 ,c1.ITEM8_2_99 ,c1.ITEM8_3_99 ,c1.ITEM8_4_99 ,c1.ITEM8_5_99 ,c1.ITEM8_NOTE " & vbCrLf
        sql &= " ,c1.ITEM9_1 ,c1.ITEM9_2 ,c1.ITEM9_3 ,c1.ITEM9_4 ,c1.ITEM9_5 ,c1.ITEM9_NOTE" & vbCrLf
        sql &= " ,c1.ITEM10_NOTE " & vbCrLf
        'VISITWAY: 1:實地訪查  '2:視訊訪查
        sql &= " ,a.VISITORNAME ,a.VISITWAY" & vbCrLf
        sql &= " ,a.ORGID ,a.RID" & vbCrLf ',c1.ISCLEAR 
        sql &= " ,a.STUD_NAME ITEMA_1,a.STUD_NAME2 ITEMA_2 ,c1.ITEMA_3 ,c1.ITEMA_4 ,c1.ITEMA_5 ,c1.ITEMA_NOTE" & vbCrLf
        sql &= " ,c1.ITEMB_1 ,c1.ITEMB_2 ,c1.ITEMB_3 ,c1.ITEMB_4 ,c1.ITEMB_5 ,c1.ITEMB_NOTE " & vbCrLf
        'sql &= " ,c1.TELVISITREASON ,c1.REALVISITDATE" & vbCrLf
        sql &= " FROM dbo.CLASS_UNEXPECTVISITOR a WITH(NOLOCK)" & vbCrLf
        sql &= " LEFT JOIN dbo.CLASS_UNEXPECTVISITORAPPLY c1 WITH(NOLOCK) ON c1.OCID = A.OCID AND c1.SeqNO = A.SeqNO " & vbCrLf
        sql &= " WHERE a.OCID=@OCID AND a.SEQNO=@SEQNO" & vbCrLf
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)
        If dr IsNot Nothing Then
            ApplyDate.Text = If(flag_ROC, TIMS.Cdate17(dr("APPLYDATE")), TIMS.Cdate3(dr("APPLYDATE")))

            For Each item As DataGridItem In DataGrid1.Items
                'Dim labShowItem As Label = item.FindControl("labShowItem")
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
                'Dim N As String = CStr(item.Cells(0).Text) '項次
                hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value) '項次
                'labShowItem.Text = TIMS.ClearSQM(labShowItem.Text) '項次
                Dim N As String = hid_dataitem.Value '項次
                If N <> "" Then
                    Select Case N
                        Case "8" '數字儲存
                            txtAnswer1.Text = dr("Item" & N & "_1_99").ToString
                            txtAnswer2.Text = dr("Item" & N & "_2_99").ToString
                            txtAnswer3.Text = dr("Item" & N & "_3_99").ToString
                            txtAnswer4.Text = dr("Item" & N & "_4_99").ToString
                            txtAnswer5.Text = dr("Item" & N & "_5_99").ToString
                        Case "A"
                            txtAnswer1.Text = dr("Item" & N & "_1").ToString
                            txtAnswer2.Text = dr("Item" & N & "_2").ToString
                            txtAnswer3.Text = dr("Item" & N & "_3").ToString
                            txtAnswer4.Text = dr("Item" & N & "_4").ToString
                            txtAnswer5.Text = dr("Item" & N & "_5").ToString
                        Case Else '(B) (1-9 (排除8))
                            If dr("Item" & N & "_1").ToString <> "" Then Common.SetListItem(rdoAnswer1, dr("Item" & N & "_1").ToString)
                            If dr("Item" & N & "_2").ToString <> "" Then Common.SetListItem(rdoAnswer2, dr("Item" & N & "_2").ToString)
                            If dr("Item" & N & "_3").ToString <> "" Then Common.SetListItem(rdoAnswer3, dr("Item" & N & "_3").ToString)
                            If dr("Item" & N & "_4").ToString <> "" Then Common.SetListItem(rdoAnswer4, dr("Item" & N & "_4").ToString)
                            If dr("Item" & N & "_5").ToString <> "" Then Common.SetListItem(rdoAnswer5, dr("Item" & N & "_5").ToString)
                    End Select
                    Dim NoteColName As String = "Item" & N & "_Note"
                    txtNote.Text = dr(NoteColName).ToString
                End If
            Next

            'If dr("Item10").ToString <> "" Then Item10.SelectedValue = dr("Item10").ToString
            'Common.SetListItem(Item10, Convert.ToString(dr("Item10"))) '1/2
            'Item10_1.Checked = If(Convert.ToString(dr("Item10_1")) = "1", True, False)
            Item10_Note.Text = dr("Item10_Note").ToString
            'Item10_Other.Text = dr("Item10_Other").ToString
            VisitorName.Text = Convert.ToString(dr("VisitorName"))
            'VISITWAY: 1:實地訪查  '2:視訊訪查
            Dim v_rblVISITWAY As String = If(Convert.ToString(dr("VISITWAY")) <> "", dr("VISITWAY"), "1")
            Common.SetListItem(rblVISITWAY, v_rblVISITWAY)
            rblVISITWAY.Enabled = False
            VisitorOrgNAME.Text = TIMS.GET_ORGNAME(dr("OrgID"), objconn)
        End If
        VisitorName.Enabled = False
        VisitorOrgNAME.Enabled = False
        ApplyDate.Enabled = False

        Dim drC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If drC IsNot Nothing Then
            center.Text = Convert.ToString(drC("OrgName"))
            RIDValue.Value = Convert.ToString(drC("RID"))
            TMID1.Text = Convert.ToString(drC("TrainName2"))
            TMIDValue1.Value = Convert.ToString(drC("TMID"))
            OCID1.Text = Convert.ToString(drC("CLASSCNAME2"))
            OCIDValue1.Value = Convert.ToString(drC("OCID"))
            center.Enabled = False
            TMID1.Enabled = False
            OCID1.Enabled = False
        End If
    End Sub

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
        Const flag As String = "," '逗號 有分隔

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
                    Dim s_CBA1 As String = cb1_show.ClientID
                    Dim s_RDOB1 As String = rdoAnswer1.UniqueID
                    Dim s_RDOB2 As String = rdoAnswer2.UniqueID
                    Dim s_RDOB3 As String = rdoAnswer3.UniqueID
                    Dim s_RDOB4 As String = rdoAnswer4.UniqueID
                    Dim s_RDOB5 As String = rdoAnswer5.UniqueID
                    cb1_show.Attributes("onclick") = String.Format("CHANGE_CB1('{0}','{1}','{2}','{3}','{4}','{5}');", s_CBA1, s_RDOB1, s_RDOB2, s_RDOB3, s_RDOB4, s_RDOB5)
                End If

                If drv("Answer").ToString.IndexOf(flag) <> -1 Then
                    'rdoDisplay
                    rdoAnswer1 = rdoDisplay(rdoAnswer1, drv("Answer").ToString, flag)
                    rdoAnswer2 = rdoDisplay(rdoAnswer2, drv("Answer").ToString, flag)
                    rdoAnswer3 = rdoDisplay(rdoAnswer3, drv("Answer").ToString, flag)
                    rdoAnswer4 = rdoDisplay(rdoAnswer4, drv("Answer").ToString, flag)
                    rdoAnswer5 = rdoDisplay(rdoAnswer5, drv("Answer").ToString, flag)

                    txtAnswer1.Visible = False
                    txtAnswer2.Visible = False
                    txtAnswer3.Visible = False
                    txtAnswer4.Visible = False
                    txtAnswer5.Visible = False
                Else
                    Const cst_MoneyFormat_Msg As String = "請填寫數字格式"
                    Select Case Convert.ToString(drv("Answer"))
                        Case "Money"
                            TIMS.Tooltip(txtAnswer1, cst_MoneyFormat_Msg)
                            TIMS.Tooltip(txtAnswer2, cst_MoneyFormat_Msg)
                            TIMS.Tooltip(txtAnswer3, cst_MoneyFormat_Msg)
                            TIMS.Tooltip(txtAnswer4, cst_MoneyFormat_Msg)
                            TIMS.Tooltip(txtAnswer5, cst_MoneyFormat_Msg)
                        Case "lockName"
                            txtAnswer1.Enabled = False
                            txtAnswer2.Enabled = False
                            txtAnswer3.Enabled = False
                            txtAnswer4.Enabled = False
                            txtAnswer5.Enabled = False
                    End Select
                End If

                Dim N As String = Convert.ToString(drv("Item"))
                Dim txtNote As TextBox = e.Item.FindControl("txtNote")
                Dim NoteColName As String = "Item" & N & "_Note"
                Dim MyValue As Integer = 0
                For Each drCol As DataRow In dtCOLUMNS.Rows
                    MyValue = drCol("CHARACTER_MAXIMUM_LENGTH")
                    If UCase(drCol("COLUMN_NAME")) = UCase(NoteColName) Then
                        txtNote.MaxLength = MyValue
                        TIMS.Tooltip(txtNote, "欄位長度" & CStr(MyValue))
                        Exit For
                    End If
                Next
        End Select
    End Sub

    ''' <summary>檢核-該迴圈只處理數字</summary>
    ''' <param name="ErrMsg"></param>
    Sub CHECK_CLASS_UNEXPECTVISITOR_NUM(ByRef ErrMsg As String)
        ErrMsg = ""
        '依欄位名稱，修改資料
        Dim fg_cb4 As Boolean = False
        Dim fg_cb5 As Boolean = False
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
            'Dim N As String = CStr(item.Cells(0).Text) '項次
            hid_dataitem.Value = TIMS.ClearSQM(hid_dataitem.Value) '項次
            labShowItem.Text = TIMS.ClearSQM(labShowItem.Text) '項次
            Dim showN As String = labShowItem.Text 'hid_dataitem.Value '項次
            'Dim N As String = TIMS.ClearSQM(item.Cells(0).Text) '項次
            If showN <> "" AndAlso IsNumeric(showN) Then '該迴圈只處理數字
                If showN = "9" Then ' 第9題
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
                    If rdoAnswer1 IsNot Nothing AndAlso rdoAnswer1.Visible AndAlso TIMS.GetListValue(rdoAnswer1) = "" Then
                        ErrMsg += "項次" & showN & "的訪問一不可為空\n"
                    End If
                    'If txtAnswer1 IsNot Nothing Then
                    '    If txtAnswer1.Visible AndAlso txtAnswer1.Text = "" Then
                    '        ErrMsg += "項次" & N & "的訪問一不可為空\n"
                    '    End If
                    'End If
                End If
                txtNote.Text = TIMS.ClearSQM(txtNote.Text)
                If txtNote.Text.Trim.Length > 100 Then ErrMsg += "項次" & showN & "的備註／說明事項 超過系統長度100\n"
                'If ErrMsg = "" Then dr("Item" & N & "_Note") = txtNote.Text
            End If
        Next

        If fg_cb4 AndAlso fg_cb5 Then
            ErrMsg += "項次4, 項次5 不可同時勾選\n"
        End If
    End Sub

    ''' <summary>
    ''' 該迴圈只處理非數字
    ''' </summary>
    ''' <param name="ErrMsg"></param>
    Sub CHECK_CLASS_UNEXPECTVISITOR_CHAR(ByRef ErrMsg As String)
        Dim rst As String = ""
        'If OCID = "" Then Return rst
        'If SeqNo = "" Then Return rst
        'Dim ErrMsg As String = ""
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
                Case "B"
                    Dim v_rdoAnswer1 As String = TIMS.GetListValue(rdoAnswer1)
                    If v_rdoAnswer1 = "" Then ErrMsg += "項次" & showN & "的訪問一不可為空\n"
                Case Else 'Case "A"
                    txtAnswer1.Text = TIMS.ClearSQM(txtAnswer1.Text)
                    'txtAnswer2.Text = TIMS.ClearSQM(txtAnswer2.Text)
                    'txtAnswer3.Text = TIMS.ClearSQM(txtAnswer3.Text)
                    'txtAnswer4.Text = TIMS.ClearSQM(txtAnswer4.Text)
                    'txtAnswer5.Text = TIMS.ClearSQM(txtAnswer5.Text)
                    If txtAnswer1.Text = "" Then ErrMsg += "項次" & showN & "的訪問一不可為空\n"
            End Select
            txtNote.Text = TIMS.ClearSQM(txtNote.Text)
            If txtNote.Text.Length > 100 Then ErrMsg &= "項次" & showN & "的備註／說明事項 超過系統長度100\n"
        Next
        'If ErrMsg <> "" Then Return ErrMsg '異常離開
        'Return rst '正常為空
    End Sub


    '非數字項次儲存 (A/B) 數字項次儲存 (1-9)
    Sub SAVE_CLASS_UNEXPECTVISITORAPPLY(ByRef oConn As SqlConnection, ByRef parms As Hashtable)
        Dim v_OCIDValue1 As String = TIMS.GetMyValue2(parms, "OCIDValue1") 'OCIDValue1.value/Request("OCID") 
        Dim v_rqSEQNO As String = TIMS.GetMyValue2(parms, "rqSEQNO") 'Request("SeqNo")

        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim da As New SqlDataAdapter

        sql = " SELECT * FROM CLASS_UNEXPECTVISITORAPPLY WHERE OCID=@OCIDValue1 AND SEQNO=@rqSEQNO"
        dt = DbAccess.GetDataTable(sql, oConn, parms)
        If dt.Rows.Count = 0 Then
            sql = " SELECT * FROM CLASS_UNEXPECTVISITORAPPLY WHERE 1<>1 "
            dt = DbAccess.GetDataTable(sql, da, oConn)
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("OCID") = v_OCIDValue1
            dr("SeqNo") = v_rqSEQNO
            'ApplyDate.Text = TIMS.ClearSQM(ApplyDate.Text)
            'dr("ApplyDate") = If(flag_ROC, TIMS.cdate18(ApplyDate.Text), TIMS.cdate2(ApplyDate.Text))  'edit，by:20181018
        Else
            '修改
            sql = "SELECT * FROM CLASS_UNEXPECTVISITORAPPLY WHERE OCID=" & v_OCIDValue1 & " AND SEQNO=" & v_rqSEQNO & ""
            dt = DbAccess.GetDataTable(sql, da, oConn)
            If dt.Rows.Count <> 1 Then
                Common.MessageBox(Me, "資料異常，請重新查詢!")
                Exit Sub
            End If
            dr = dt.Rows(0)
        End If

        For Each item As DataGridItem In DataGrid1.Items
            'Dim labShowItem As Label = item.FindControl("labShowItem")
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
            'labShowItem.Text = TIMS.ClearSQM(labShowItem.Text) '項次
            Dim N As String = hid_dataitem.Value '項次
            If N = "" Then Exit For '不可為空
            'If IsNumeric(N) Then Continue For '該迴圈只處理非數字
            If IsNumeric(N) Then
                '只處理數字
                Select Case N
                    Case "8" '數字儲存
                        txtAnswer1.Text = TIMS.ClearSQM(txtAnswer1.Text)
                        txtAnswer2.Text = TIMS.ClearSQM(txtAnswer2.Text)
                        txtAnswer3.Text = TIMS.ClearSQM(txtAnswer3.Text)
                        txtAnswer4.Text = TIMS.ClearSQM(txtAnswer4.Text)
                        txtAnswer5.Text = TIMS.ClearSQM(txtAnswer5.Text)
                        dr("Item" & N & "_1_99") = If(txtAnswer1.Text <> "", CInt(txtAnswer1.Text), Convert.DBNull)
                        dr("Item" & N & "_2_99") = If(txtAnswer2.Text <> "", CInt(txtAnswer2.Text), Convert.DBNull)
                        dr("Item" & N & "_3_99") = If(txtAnswer3.Text <> "", CInt(txtAnswer3.Text), Convert.DBNull)
                        dr("Item" & N & "_4_99") = If(txtAnswer4.Text <> "", CInt(txtAnswer4.Text), Convert.DBNull)
                        dr("Item" & N & "_5_99") = If(txtAnswer5.Text <> "", CInt(txtAnswer5.Text), Convert.DBNull)
                    Case Else
                        dr("Item" & N & "_1") = If(rdoAnswer1.SelectedValue <> "", rdoAnswer1.SelectedValue, Convert.DBNull)
                        dr("Item" & N & "_2") = If(rdoAnswer2.SelectedValue <> "", rdoAnswer2.SelectedValue, Convert.DBNull)
                        dr("Item" & N & "_3") = If(rdoAnswer3.SelectedValue <> "", rdoAnswer3.SelectedValue, Convert.DBNull)
                        dr("Item" & N & "_4") = If(rdoAnswer4.SelectedValue <> "", rdoAnswer4.SelectedValue, Convert.DBNull)
                        dr("Item" & N & "_5") = If(rdoAnswer5.SelectedValue <> "", rdoAnswer5.SelectedValue, Convert.DBNull)
                End Select
            Else
                '只處理非數字
                Select Case N
                    Case "B" '身份
                        Dim v_rdoAnswer1 As String = TIMS.GetListValue(rdoAnswer1)
                        Dim v_rdoAnswer2 As String = TIMS.GetListValue(rdoAnswer2)
                        Dim v_rdoAnswer3 As String = TIMS.GetListValue(rdoAnswer3)
                        Dim v_rdoAnswer4 As String = TIMS.GetListValue(rdoAnswer4)
                        Dim v_rdoAnswer5 As String = TIMS.GetListValue(rdoAnswer5)
                        dr("Item" & N & "_1") = If(v_rdoAnswer1 <> "", v_rdoAnswer1, Convert.DBNull)
                        dr("Item" & N & "_2") = If(v_rdoAnswer2 <> "", v_rdoAnswer2, Convert.DBNull)
                        dr("Item" & N & "_3") = If(v_rdoAnswer3 <> "", v_rdoAnswer3, Convert.DBNull)
                        dr("Item" & N & "_4") = If(v_rdoAnswer4 <> "", v_rdoAnswer4, Convert.DBNull)
                        dr("Item" & N & "_5") = If(v_rdoAnswer5 <> "", v_rdoAnswer5, Convert.DBNull)
                    Case Else 'A
                        dr("Item" & N & "_1") = If(txtAnswer1.Text <> "", txtAnswer1.Text, Convert.DBNull)
                        dr("Item" & N & "_2") = If(txtAnswer2.Text <> "", txtAnswer2.Text, Convert.DBNull)
                        dr("Item" & N & "_3") = If(txtAnswer3.Text <> "", txtAnswer3.Text, Convert.DBNull)
                        dr("Item" & N & "_4") = If(txtAnswer4.Text <> "", txtAnswer4.Text, Convert.DBNull)
                        dr("Item" & N & "_5") = If(txtAnswer5.Text <> "", txtAnswer5.Text, Convert.DBNull)
                End Select
            End If
            txtNote.Text = TIMS.ClearSQM(txtNote.Text)
            dr("Item" & N & "_Note") = If(txtNote.Text <> "", txtNote.Text, Convert.DBNull)

            If cb1_show.Visible AndAlso cb1_show.Text <> "" AndAlso hid_ckcolumn.Value <> "" Then
                dr(hid_ckcolumn.Value) = If(cb1_show.Checked, "Y", Convert.DBNull)
            End If
        Next
        dr("ITEM10_NOTE") = If(Item10_Note.Text = "", Convert.DBNull, Item10_Note.Text)
        dr("ModifyAcct") = sm.UserInfo.UserID
        dr("ModifyDate") = Now
        DbAccess.UpdateDataTable(dt, da)
    End Sub

    '儲存
    Private Sub BtnSave1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles BtnSave1.Click
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        Dim s_ERRMSG As String = ""
        Call CHECK_CLASS_UNEXPECTVISITOR_NUM(s_ERRMSG)
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, Replace(s_ERRMSG, "\n", vbCrLf))
            Exit Sub
        End If

        s_ERRMSG = ""
        Call CHECK_CLASS_UNEXPECTVISITOR_CHAR(s_ERRMSG)
        If s_ERRMSG <> "" Then
            Common.MessageBox(Me, Replace(s_ERRMSG, "\n", vbCrLf))
            Exit Sub
        End If

        Dim v_OCIDValue1 As String = TIMS.ClearSQM(Request("OCID"))
        Dim v_rqSEQNO As String = TIMS.ClearSQM(Request("SEQNO"))
        Dim parms As New Hashtable From {{"OCIDValue1", v_OCIDValue1}, {"rqSEQNO", v_rqSEQNO}}
        Try
            Call SAVE_CLASS_UNEXPECTVISITORAPPLY(objconn, parms)
        Catch ex As Exception
            Call TIMS.WriteTraceLog(Me, ex)
            Common.MessageBox(Me, "儲存失敗!")
            Exit Sub
        End Try
        Common.MessageBox(Me, "儲存成功")

        If Session("SearchStr") Is Nothing Then Session("SearchStr") = ViewState("SearchStr")
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim s_Url1 As String = "CP_01_006.aspx?ID=" & RqID
        TIMS.Utl_Redirect1(Me, s_Url1)
    End Sub

    Protected Sub BtnBack1_Click(sender As Object, e As EventArgs) Handles BtnBack1.Click
        If Session("SearchStr") Is Nothing Then Session("SearchStr") = ViewState("SearchStr")
        Dim RqID As String = TIMS.Get_MRqID(Me)
        Dim s_Url1 As String = "CP_01_006.aspx?ID=" & RqID
        TIMS.Utl_Redirect1(Me, s_Url1)
    End Sub

    '列印
    Protected Sub BtnPrint1_Click(sender As Object, e As EventArgs) Handles BtnPrint1.Click
        Dim v_rqOCID As String = TIMS.ClearSQM(Request("OCID"))
        Dim v_rqSEQNO As String = TIMS.ClearSQM(Request("SEQNO"))
        Dim s_RID As String = sm.UserInfo.RID
        If sm.UserInfo.RID = "A" Then
            Dim drC As DataRow = TIMS.GetOCIDDate(v_rqOCID, objconn)
            If drC Is Nothing Then Return
            s_RID = drC("RID")
        End If

        Dim myValue As String = ""
        myValue &= "OCID=" & v_rqOCID
        myValue &= "&SeqNo=" & v_rqSEQNO
        myValue &= "&RID=" & s_RID 'sm.UserInfo.RID
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myValue)
    End Sub

    Protected Sub DataGrid1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DataGrid1.SelectedIndexChanged

    End Sub
End Class
