Partial Class CP_01_007_add
    Inherits AuthBasePage

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: 此為 Web Form 設計工具所需的方法呼叫
        '請勿使用程式碼編輯器進行修改。
        'InitializeComponent()

        Dim MyValue As Integer = 0
        MyValue = 1000
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
    'UPDATE TABLE : Class_UnExpectTel .Class_UnexpectTelApply 
    Dim dtCOLUMNS As DataTable
    Dim objconn As SqlConnection
    Dim flag_ROC As Boolean = TIMS.CHK_REPLACE2ROC_YEARS()

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁

        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, blnCanAdds, blnCanMod, blnCanDel, blnCanSech, blnCanPrnt) '2011 取得功能按鈕權限值
        TIMS.OpenDbConn(objconn)
        '檢查Session是否存在--------------------------End

        Dim strSql As String = ""
#Region "(No Use)"

        'Dim dtCOLUMNS As DataTable 
        'strSql = "" & vbCrLf
        'strSql &= " SELECT TABLE_NAME,COLUMN_NAME,DATA_TYPE,DATA_LENGTH CHARACTER_MAXIMUM_LENGTH" & vbCrLf
        'strSql &= " FROM USER_TAB_COLUMNS" & vbCrLf
        'strSql &= " WHERE UPPER(TABLE_NAME) IN ('CLASS_UNEXPECTTEL','CLASS_UNEXPECTTELAPPLY') " & vbCrLf
        'strSql &= " AND UPPER(DATA_TYPE) IN ('NVARCHAR2','VARCHAR2','CHAR')" & vbCrLf

#End Region

        'sql server
        strSql = ""
        strSql &= " SELECT TABLE_NAME, COLUMN_NAME, DATA_TYPE ,CHARACTER_MAXIMUM_LENGTH ,CHARACTER_MAXIMUM_LENGTH CHAR_LENGTH " & vbCrLf
        strSql &= " FROM INFORMATION_SCHEMA.COLUMNS "
        strSql &= " WHERE TABLE_NAME IN ('CLASS_UNEXPECTTEL','CLASS_UNEXPECTTELAPPLY') AND DATA_TYPE IN ('NVARCHAR','VARCHAR','CHAR','NCHAR') " & vbCrLf
        dtCOLUMNS = DbAccess.GetDataTable(strSql, objconn)

        If Not IsPostBack Then
            Me.ViewState("SearchStr") = Session("SearchStr")
            Session("SearchStr") = Nothing
            Me.ViewState("_SearchStr") = Session("_SearchStr")
            Session("_SearchStr") = Nothing
            If Request("RIDValue") <> "" Then RIDValue.Value = Request("RIDValue")

            'If Request("center") <> "" Then center.Text = Request("center")
            'fix 機構名稱顯示不出來的問題
            If Not ViewState("SearchStr") Is Nothing Then
                Dim MyValue As String = ""

                MyValue = TIMS.GetMyValue(ViewState("SearchStr"), "center")
                center.Text = Replace(MyValue, "%26", "&")
            End If

            If Request("TMID1") <> "" Then TMID1.Text = Request("TMID1")
            If Request("OCID1") <> "" Then OCID1.Text = Request("OCID1")
            If Request("TMIDValue1") <> "" Then TMIDValue1.Value = Request("TMIDValue1")
            If Request("OCIDValue1") <> "" Then OCIDValue1.Value = Request("OCIDValue1")
            VisitorOrgNAME.Text = sm.UserInfo.OrgName
        End If

        If sm.UserInfo.RID = "A" Or sm.UserInfo.RoleID <= 1 Then
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg.aspx');"
        Else
            Button2.Attributes("onclick") = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        End If

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

        If Not IsPostBack Then
            create2() '必須放在'create(Request("OCID"), Request("SeqNo")) 之前
            If Request("OCID") <> "" Then Call create(Request("OCID"), Request("SeqNo"))
        End If

        If Request("State") = "View" Then Button1.Visible = False

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
        If Request("Type") = "CT" And Request("State") = "View" Then
            Button4.Visible = False
            Button5.Visible = True
        ElseIf Request("Type") = "CT" Then
            Button4.Visible = False
            Button5.Visible = True
        Else
            Button4.Visible = True
            Button5.Visible = False
        End If
    End Sub

    '建立DataGrid表格
    Sub create2()
        PageControler1.PageDataGrid = DataGrid1
        '建立資料格式Table----------------Start
        Dim dt As New DataTable
        dt.TableName = "Class_UnExpectTelQuestion"
        dt.Columns.Add(New DataColumn("Item"))
        dt.Columns.Add(New DataColumn("Question"))
        dt.Columns.Add(New DataColumn("Answer"))

        Dim dr As DataRow
#Region "(No Use)"

        'insert into Class_UnExpectTelQuestion select '1','是否有在該單位的班級上課？課程名稱及上課地點？','相符合,不符合'
        'insert into Class_UnExpectTelQuestion select '2','開訓日期及結訓日期？每周上課時間？總時數？','相符合,不符合'
        'insert into Class_UnExpectTelQuestion select '3','師資教學品質是否良好？','良好,尚可'
        'insert into Class_UnExpectTelQuestion select '4','場地及設備品質是否良好？','良好,尚可'
        'insert into Class_UnExpectTelQuestion select '5','教材是否足夠？','良好,尚可'
        'insert into Class_UnExpectTelQuestion select '6','最後一次上課時間？上次上課講師姓名或姓式？','相符合,不符合'
        'insert into Class_UnExpectTelQuestion select '7','上課時間是否有超過一半？','相符合,不符合'
        'insert into Class_UnExpectTelQuestion select '8','是否有繳交費用？','相符合,不符合'
        'insert into Class_UnExpectTelQuestion select '9','參訓後對您是否有幫助？','良好,尚可'

#End Region

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "A"
        dr("Question") = "聯絡時間"
        dr("Answer") = "DTime"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "B"
        dr("Question") = "撥話線路(電話抽訪)"
        dr("Answer") = "UsePhone"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "C"
        dr("Question") = "受訪學員姓名"
        dr("Answer") = "Name"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "D"
        dr("Question") = "聯絡電話"
        dr("Answer") = "Phone"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "E"
        dr("Question") = "身份"
        dr("Answer") = "一般,特殊"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "1"
        dr("Question") = "是否有在該單位的班級上課？課程名稱及上課地點？"
        dr("Answer") = "相符合,不符合"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "2"
        dr("Question") = "開訓日期及結訓日期？每周上課時間？總時數？"
        dr("Answer") = "相符合,不符合"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "3"
        dr("Question") = "師資教學品質是否良好？"
        dr("Answer") = "良好,尚可,待改善"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "4"
        dr("Question") = "場地及設備品質是否良好？"
        dr("Answer") = "良好,尚可,待改善"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "5"
        dr("Question") = "教材是否足夠？"
        dr("Answer") = "良好,尚可,待改善"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "6"
        dr("Question") = "最近一次上課時間？上課講師姓名或姓氏？"
        dr("Answer") = "相符合,不符合"

        If CInt(sm.UserInfo.Years) < 2010 Then
            dr = dt.NewRow   '2010 年以前的資料顯示
            dt.Rows.Add(dr)
            dr("Item") = "7"
            dr("Question") = "上課時間是否有超過一半？"
            dr("Answer") = "相符合,不符合"
        Else                  '2010 年以後的資料顯示
            dr = dt.NewRow
            dt.Rows.Add(dr)
            dr("Item") = "7"
            dr("Question") = "上課講師是否有遲到早退？"
            dr("Answer") = "相符合,不符合"
        End If

#Region "(No Use)"

        ''2010 年以前的資料顯示
        'dr = dt.NewRow
        'dt.Rows.Add(dr)
        'dr("Item") = "8"
        'dr("Question") = "是否有繳交費用？"
        'dr("Answer") = "相符合,不符合"

#End Region

        '2010 年以後的資料顯示
        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "8"
        dr("Question") = "繳交費用額度？"
        dr("Answer") = "Money"

        dr = dt.NewRow
        dt.Rows.Add(dr)
        dr("Item") = "9"
        dr("Question") = "參訓後對您是否有幫助？"
        dr("Answer") = "良好,尚可"

        '建立資料格式Table----------------End
        PageControler1.DataTableCreate(dt)
        PageControler1.Visible = False
    End Sub

    Sub create(ByVal OCID As String, ByVal SeqNo As String)
        Dim parms As Hashtable = New Hashtable()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT c1.OCID ,c1.SEQNO ,c1.APPLYDATE ,c1.ITEM1_1 ,c1.ITEM1_2 ,c1.ITEM1_3 ,c1.ITEM1_4 ,c1.ITEM1_5 " & vbCrLf
        sql &= "        ,c1.ITEM1_NOTE ,c1.ITEM2_1 ,c1.ITEM2_2 ,c1.ITEM2_3 ,c1.ITEM2_4 ,c1.ITEM2_5 ,c1.ITEM2_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM3_1 ,c1.ITEM3_2 ,c1.ITEM3_3 ,c1.ITEM3_4 ,c1.ITEM3_5 ,c1.ITEM3_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM4_1 ,c1.ITEM4_2 ,c1.ITEM4_3 ,c1.ITEM4_4 ,c1.ITEM4_5 ,c1.ITEM4_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM5_1 ,c1.ITEM5_2 ,c1.ITEM5_3 ,c1.ITEM5_4 ,c1.ITEM5_5 ,c1.ITEM5_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM6_1 ,c1.ITEM6_2 ,c1.ITEM6_3 ,c1.ITEM6_4 ,c1.ITEM6_5 ,c1.ITEM6_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM7_1 ,c1.ITEM7_2 ,c1.ITEM7_3 ,c1.ITEM7_4 ,c1.ITEM7_5 ,c1.ITEM7_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM8_1 ,c1.ITEM8_2 ,c1.ITEM8_3 ,c1.ITEM8_4 ,c1.ITEM8_5 ,c1.ITEM8_NOTE " & vbCrLf
        sql &= "        ,c1.ITEM9_1 ,c1.ITEM9_2 ,c1.ITEM9_3 ,c1.ITEM9_4 ,c1.ITEM9_5 ,c1.ITEM9_NOTE" & vbCrLf
        sql &= "        ,c1.ITEM10 ,c1.ITEM10_1 ,c1.ITEM10_NOTE ,c1.ITEM10_OTHER ,c1.VISITORNAME " & vbCrLf
        sql &= "        ,c1.RID ,c1.ISCLEAR ,c1.ITEM8_1_99 ,c1.ITEM8_2_99 ,c1.ITEM8_3_99 ,c1.ITEM8_4_99 ,c1.ITEM8_5_99 ,c1.ORGID" & vbCrLf
        sql &= "        ,c2.ITEMA_1 ,c2.ITEMA_2 ,c2.ITEMA_3 ,c2.ITEMA_4 ,c2.ITEMA_5 ,c2.ITEMA_NOTE ,c2.ITEMB_1 ,c2.ITEMB_2 ,c2.ITEMB_3 ,c2.ITEMB_4 ,c2.ITEMB_5 ,c2.ITEMB_NOTE " & vbCrLf
        sql &= "        ,c2.ITEMC_1 ,c2.ITEMC_2 ,c2.ITEMC_3 ,c2.ITEMC_4 ,c2.ITEMC_5 ,c2.ITEMC_NOTE ,c2.ITEMD_1 ,c2.ITEMD_2 ,c2.ITEMD_3 ,c2.ITEMD_4 ,c2.ITEMD_5 ,c2.ITEMD_NOTE " & vbCrLf
        sql &= "        ,c2.ITEME_1 ,c2.ITEME_2 ,c2.ITEME_3 ,c2.ITEME_4 ,c2.ITEME_5 ,c2.ITEME_NOTE" & vbCrLf
        sql &= " FROM CLASS_UNEXPECTTEL c1 " & vbCrLf
        sql &= " LEFT JOIN CLASS_UNEXPECTTELAPPLY c2 ON c1.OCID = c2.OCID AND c1.SeqNO = c2.SeqNO " & vbCrLf
        sql &= " WHERE 1=1 AND c1.OCID = @OCID " & vbCrLf

        parms.Add("OCID", OCID)
        If SeqNo <> "" Then
            sql &= " AND c1.SEQNO = @SEQNO " & vbCrLf
            parms.Add("SEQNO", SeqNo)
        End If

        Dim i As Integer = 0
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn, parms)

        If Not dr Is Nothing Then
            If flag_ROC Then
                ApplyDate.Text = TIMS.cdate17(dr("ApplyDate"))  'edit，by:20181018
            Else
                ApplyDate.Text = dr("ApplyDate")  'edit，by:20181018
            End If

            For Each item As DataGridItem In DataGrid1.Items
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
                Dim N As String = CStr(item.Cells(0).Text) '項次
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
                    Dim NoteColName As String = ""
                    NoteColName = "Item" & N & "_Note"
                    txtNote.Text = dr(NoteColName).ToString
#Region "(No Use)"

                    'Dim MyValue As Integer = 0
                    'For Each drCol As DataRow In dtCOLUMNS.Rows
                    '    MyValue = drCol("character_maxiMum_length")
                    '    If Convert.ToString(drCol("column_name")) = NoteColName Then
                    '        txtNote.MaxLength = MyValue
                    '        TIMS.Tooltip(txtNote, "欄位長度" & CStr(MyValue))
                    '        Exit For
                    '    End If
                    'Next

#End Region
                End If
            Next

            If dr("Item10").ToString <> "" Then Item10.SelectedValue = dr("Item10").ToString
            If dr("Item10_1").ToString <> "" Then
                If dr("Item10_1").ToString = "1" Then Item10_1.Checked = True Else Item10_1.Checked = False
            End If
            Item10_Note.Text = dr("Item10_Note").ToString
            Item10_Other.Text = dr("Item10_Other").ToString
            VisitorName.Text = Convert.ToString(dr("VisitorName"))
            If dr("OrgID").ToString <> "" Then VisitorOrgNAME.Text = TIMS.GET_OrgName(dr("OrgID"), objconn)
        End If

#Region "(No Use)"

        'Dim strSql As String = ""
        'strSql = "" & vbCrLf
        'strSql &= " SELECT b.TMID" & vbCrLf
        'strSql &= " ,b.JobID" & vbCrLf
        'strSql &= " , b.JobName" & vbCrLf
        'strSql &= " , a.OCID" & vbCrLf
        'strSql &= " , a.ClassCName" & vbCrLf
        'strSql &= " , a.CyclType" & vbCrLf
        'strSql &= " , a.LevelType" & vbCrLf
        'strSql &= " , a.RID" & vbCrLf
        'strSql &= " , c.OrgName" & vbCrLf
        'strSql &= " FROM Class_ClassInfo a" & vbCrLf
        'strSql &= " join Key_TrainType b on a.TMID=b.TMID " & vbCrLf
        'strSql &= " join Auth_Relship d on d.RID=a.RID" & vbCrLf
        'strSql &= " join Org_OrgInfo c on d.OrgID=c.OrgID" & vbCrLf
        'strSql &= " WHERE 1=1" & vbCrLf
        'strSql &= " and a.OCID='" & OCID & "'" & vbCrLf
        'dr = DbAccess.GetOneRow(strSql, objconn)

#End Region

        Dim drC As DataRow = TIMS.GetOCIDDate(OCID, objconn)
        If Not drC Is Nothing Then
            center.Text = drC("OrgName")
            RIDValue.Value = drC("RID").ToString
            TMID1.Text = Convert.ToString(drC("TrainName2"))
            TMIDValue1.Value = drC("TMID")
            OCID1.Text = drC("ClassCName2")
            OCIDValue1.Value = drC("OCID")
            center.Enabled = False
            TMID1.Enabled = False
            OCID1.Enabled = False
            Button2.Disabled = True
            Button3.Disabled = True
        End If
    End Sub

    Function UPDATE_Class_UnexpectTelApply(ByVal OCID As String, ByVal SeqNo As String) As String
        Dim rst As String = ""
        Dim sql As String = ""
        Dim dt As DataTable = Nothing
        Dim dr As DataRow = Nothing
        Dim ErrMsg As String = ""

        If OCID <> "" And SeqNo <> "" Then
            Dim da As New SqlDataAdapter
            sql = " SELECT * FROM CLASS_UNEXPECTTELAPPLY WHERE OCID = '" & OCID & "' AND SeqNo = '" & SeqNo & "'"
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
                Dim N As String = CStr(item.Cells(0).Text) '項次
                If N <> "" Then
                    If Not IsNumeric(N) Then
                        Select Case N
                            Case "E"
                                If rdoAnswer1.SelectedIndex <> -1 Then
                                    dr("Item" & N & "_1") = rdoAnswer1.SelectedValue
                                Else
                                    ErrMsg += "項次" & N & "的訪問一不可為空\n"
                                End If
                                If rdoAnswer2.SelectedIndex <> -1 Then dr("Item" & N & "_2") = rdoAnswer2.SelectedValue
                                If rdoAnswer3.SelectedIndex <> -1 Then dr("Item" & N & "_3") = rdoAnswer3.SelectedValue
                                If rdoAnswer4.SelectedIndex <> -1 Then dr("Item" & N & "_4") = rdoAnswer4.SelectedValue
                                If rdoAnswer5.SelectedIndex <> -1 Then dr("Item" & N & "_5") = rdoAnswer5.SelectedValue
                            Case Else
                                If txtAnswer1.Text <> "" Then
                                    dr("Item" & N & "_1") = txtAnswer1.Text
                                Else
                                    ErrMsg += "項次" & N & "的訪問一不可為空\n"
                                    dr("Item" & N & "_1") = Convert.DBNull
                                End If
                                If txtAnswer2.Text <> "" Then
                                    dr("Item" & N & "_2") = txtAnswer2.Text
                                Else
                                    dr("Item" & N & "_2") = Convert.DBNull
                                End If
                                If txtAnswer3.Text <> "" Then
                                    dr("Item" & N & "_3") = txtAnswer3.Text
                                Else
                                    dr("Item" & N & "_3") = Convert.DBNull
                                End If
                                If txtAnswer4.Text <> "" Then
                                    dr("Item" & N & "_4") = txtAnswer4.Text
                                Else
                                    dr("Item" & N & "_4") = Convert.DBNull
                                End If
                                If txtAnswer5.Text <> "" Then
                                    dr("Item" & N & "_5") = txtAnswer5.Text
                                Else
                                    dr("Item" & N & "_5") = Convert.DBNull
                                End If
                        End Select
                        txtNote.Text = txtNote.Text.Trim
                        If txtNote.Text.Trim.Length > 100 Then ErrMsg &= "項次" & N & "的備註／說明事項 超過系統長度100\n"
                        If ErrMsg = "" Then
                            If txtNote.Text <> "" Then
                                dr("Item" & N & "_Note") = txtNote.Text
                            Else
                                dr("Item" & N & "_Note") = Convert.DBNull
                            End If
                        End If
                    End If
                End If
            Next
            Try
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                If ErrMsg = "" Then
                    DbAccess.UpdateDataTable(dt, da)
                Else
                    rst = ErrMsg
                End If
            Catch ex As Exception
                rst = ex.Message
            End Try
        End If
        Return rst
    End Function

    '儲存
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim sql As String = ""
        Dim ErrMsg As String = ""
        Dim da As SqlDataAdapter = Nothing
        Dim dt As DataTable
        Dim dr As DataRow
        Dim SeqNo As Integer = 0

        Try
            ErrMsg = ""
            If Request("State") = "Add" Then '表示新增狀態
                '先取出最大SeqNo
                sql = " SELECT MAX(SeqNO) num FROM Class_UnexpectTel WHERE OCID = '" & OCIDValue1.Value & "' "
                dr = DbAccess.GetOneRow(sql, objconn)
                If Not dr Is Nothing Then
                    If IsDBNull(dr("num")) Then
                        SeqNo = 1
                    Else
                        SeqNo = CInt(dr("num")) + 1
                    End If
                End If
                sql = " SELECT * FROM Class_UnexpectTel WHERE 1<>1 "
                dt = DbAccess.GetDataTable(sql, da, objconn)
                dr = dt.NewRow
                dt.Rows.Add(dr)
                dr("OCID") = OCIDValue1.Value
                dr("SeqNo") = SeqNo
            Else
                '修改
                sql = "SELECT * FROM Class_UnexpectTel WHERE OCID='" & Request("OCID") & "' and SeqNo='" & Request("SeqNo") & "'"
                dt = DbAccess.GetDataTable(sql, da, objconn)
                If dt.Rows.Count <> 1 Then
                    Common.MessageBox(Me, "資料異常，請重新查詢!")
                    Exit Sub
                End If
                dr = dt.Rows(0)
                SeqNo = Request("SeqNo")
            End If

            If flag_ROC Then
                dr("ApplyDate") = TIMS.cdate18(ApplyDate.Text)  'edit，by:20181018
            Else
                dr("ApplyDate") = ApplyDate.Text  'edit，by:20181018
            End If

            '依欄位名稱，修改資料
            For Each item As DataGridItem In DataGrid1.Items
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
                Dim N As String = CStr(item.Cells(0).Text) '項次
                If N <> "" Then
                    If IsNumeric(N) Then
                        If N = "8" And CInt(sm.UserInfo.Years) >= 2010 Then '2010年的第8題
                            If Convert.ToString(txtAnswer1.Text) <> "" Then
                                If IsNumeric(txtAnswer1.Text) Then
                                    dr("Item" & N & "_1_99") = CInt(txtAnswer1.Text)
                                Else
                                    ErrMsg += "項次" & N & "必須為數字\n"
                                End If
                            Else
                                ErrMsg += "項次" & N & "的訪問一不可為空\n"
                            End If
                            If Convert.ToString(txtAnswer2.Text) <> "" Then
                                If IsNumeric(txtAnswer2.Text) Then
                                    dr("Item" & N & "_2_99") = CInt(txtAnswer2.Text)
                                Else
                                    ErrMsg += "項次" & N & "必須為數字\n"
                                End If
                            End If
                            If Convert.ToString(txtAnswer3.Text) <> "" Then
                                If IsNumeric(txtAnswer3.Text) Then
                                    dr("Item" & N & "_3_99") = CInt(txtAnswer3.Text)
                                Else
                                    ErrMsg += "項次" & N & "必須為數字\n"
                                End If
                            End If
                            If Convert.ToString(txtAnswer4.Text) <> "" Then
                                If IsNumeric(txtAnswer4.Text) Then
                                    dr("Item" & N & "_4_99") = CInt(txtAnswer4.Text)
                                Else
                                    ErrMsg += "項次" & N & "必須為數字\n"
                                End If
                            End If
                            If Convert.ToString(txtAnswer5.Text) <> "" Then
                                If IsNumeric(txtAnswer5.Text) Then
                                    dr("Item" & N & "_5_99") = CInt(txtAnswer5.Text)
                                Else
                                    ErrMsg += "項次" & N & "必須為數字\n"
                                End If
                            End If
                        Else
                            If rdoAnswer1.SelectedIndex <> -1 Then
                                dr("Item" & N & "_1") = rdoAnswer1.SelectedValue
                            Else
                                ErrMsg += "項次" & N & "的訪問一不可為空\n"
                            End If
                            If rdoAnswer2.SelectedIndex <> -1 Then dr("Item" & N & "_2") = rdoAnswer2.SelectedValue
                            If rdoAnswer3.SelectedIndex <> -1 Then dr("Item" & N & "_3") = rdoAnswer3.SelectedValue
                            If rdoAnswer4.SelectedIndex <> -1 Then dr("Item" & N & "_4") = rdoAnswer4.SelectedValue
                            If rdoAnswer5.SelectedIndex <> -1 Then dr("Item" & N & "_5") = rdoAnswer5.SelectedValue
                        End If
                        txtNote.Text = txtNote.Text.Trim
                        If txtNote.Text.Trim.Length > 100 Then ErrMsg += "項次" & N & "的備註／說明事項 超過系統長度100\n"
                        If ErrMsg = "" Then dr("Item" & N & "_Note") = txtNote.Text
                    End If
                End If
            Next
            If Item10.SelectedIndex <> -1 Then dr("Item10") = Item10.SelectedValue
            If Item10_1.Checked = True Then
                dr("Item10_1") = "1"
            Else
                dr("Item10_1") = "2"
            End If

            dr("Item10_Note") = IIf(Item10_Note.Text = "", Convert.DBNull, Item10_Note.Text)
            dr("Item10_Other") = IIf(Item10_Other.Text = "", Convert.DBNull, Item10_Other.Text)
            dr("OrgID") = sm.UserInfo.OrgID
            dr("VisitorName") = VisitorName.Text
            dr("RID") = sm.UserInfo.RID
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            ErrMsg += UPDATE_Class_UnexpectTelApply(OCIDValue1.Value, SeqNo.ToString)
            If ErrMsg <> "" Then
                Common.MessageBox(Me, ErrMsg.Replace("\n", "<br>"))
                Exit Sub
            End If
            DbAccess.UpdateDataTable(dt, da)

            Session("SearchStr") = Me.ViewState("SearchStr")
            Session("_SearchStr") = Me.ViewState("_SearchStr")
            Common.MessageBox(Me, "儲存成功")

            If Request("Type") = "CT" Then
                TIMS.Utl_Redirect1(Me, "CP_01_008.aspx?ID=" & Request("ID"))
            Else
                TIMS.Utl_Redirect1(Me, "CP_01_007.aspx?ID=" & Request("ID"))
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        Session("SearchStr") = Me.ViewState("SearchStr")
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        TIMS.Utl_Redirect1(Me, "CP_01_007.aspx?ID=" & Request("ID"))
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

                If drv("Answer").ToString.IndexOf(flag) <> -1 Then
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

                Dim N As String = Convert.ToString(drv("Item"))
                Dim txtNote As TextBox = e.Item.FindControl("txtNote")
                Dim NoteColName As String = ""
                NoteColName = "Item" & N & "_Note"

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

    Private Sub Button5_ServerClick(ByVal sender As Object, ByVal e As System.EventArgs) Handles Button5.ServerClick
        '訪視計畫表用
        Session("SearchStr") = Me.ViewState("SearchStr")
        Session("_SearchStr") = Me.ViewState("_SearchStr")
        TIMS.Utl_Redirect1(Me, "CP_01_008.aspx?ID=" & Request("ID"))
    End Sub
End Class