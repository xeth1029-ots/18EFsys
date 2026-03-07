Partial Class SD_14_013
    Inherits AuthBasePage

#Region "NO USE"
    '2016年後版
    'SD_14_013_2016 @BussinessTrain '產投
    'SD_14_013_2_2016 @BussinessTrain '提升勞工自主學習計畫
    'SD_14_013_*.jrxml

    '補助經費申請書(學員補助申請書)
    'xx: SD_14_013_2009 'SD_14_013_2_2009
    '2010年後版
    'SD_14_013_2010 @BussinessTrain '產投
    'SD_14_013_2_2010 @BussinessTrain '提升勞工自主學習計畫
    '2013年後版
    'SD_14_013_2013 @BussinessTrain '產投
    'SD_14_013_2_2013 @BussinessTrain '提升勞工自主學習計畫
    'SD_14_013_*.jrxml

    'Dim gTestflag As Boolean = False '測試
    '依年度判斷使用各報表
    'Const cst_2010 As String = "2010"
    'Const cst_2013 As String = "2013"
    'Const cst_2016 As String = "2016"
    'Const cst_2017 As String = "2017"

#End Region

    'Const cst_printFNoth As String = "SD_14_013_2016" '非產投計畫使用-充飛
    'Const cst_printFN17G As String = "SD_14_013_2017"
    'Const cst_printFN17W As String = "SD_14_013_2_2017"

    'SD_14_013_2018*.JRXML
    Const cst_printFNoth2 As String = "SD_14_013_2019" '非產投計畫使用

    Const cst_printFN18G As String = "SD_14_013_2018G"
    Const cst_printFN18W As String = "SD_14_013_2018W"
    'Const cst_printFN1O As String = "OJTSD140138G"(報名網空白)
    'Const cst_printFN2O As String = "OJTSD140138W"(報名網空白)
    'Const cst_printFN1S As String = "OJTSD140138GS"(報名網簽名)
    'Const cst_printFN2S As String = "OJTSD140138WS"(報名網簽名)

    'Const cst_flgCIShow As String = "flgCIShow"
    'Dim flgCIShow As Boolean = False '是否可正常顯示個資。false:不可以 true:可以。
    Const cst_DG1_身分證號碼 As Integer = 3
    'Const cst_DG1_出生日期 As Integer = 4

    Const cst_ilimitedmax As Integer = 999

    Dim iPYNum As Integer = 1 'iPYNum = TIMS.sUtl_GetPYNum(Me) '1:2017前 2:2017 3:2018
    'Dim au As New cAUTH
    Dim sMemo As String = "" '(查詢原因)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        'Call TIMS.GetAuth(Me, au.blnCanAdds, au.blnCanMod, au.blnCanDel, au.blnCanSech, au.blnCanPrnt) '2011 取得功能按鈕權限值
        Call TIMS.OpenDbConn(objconn) '開啟連線
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        'Dim gTestflag As Boolean = False '測試
        'If TIMS.Utl_GetConfigSet("TestTPID54x2012") = "Y" Then gTestflag = True '測試
        'hidgTestflag.Value = TIMS.Utl_GetConfigSet("TestTPID54x2012")
        If Not IsPostBack Then
            msg.Text = ""
            DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            '取出鍵詞-查詢原因-INQUIRY
            Dim V_INQUIRY As String = Session($"{TIMS.cst_GSE_V_INQUIRY}{TIMS.Get_MRqID(Me)}")
            If TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES) Then Call TIMS.GET_INQUIRY(ddl_INQUIRY_Sch, objconn, V_INQUIRY)

            Radio1.SelectedIndex = 0
            HidYears.Value = sm.UserInfo.Years
            Years.Value = sm.UserInfo.Years - 1911

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');"
            Button2.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

            Button1.Attributes("onclick") = "return CheckSearch();"
            Button4.Attributes("onclick") = "ClearData();"
            'Button3.Attributes("onclick") = "PrintReport('" & ReportQuery.GetSmartQueryPath & "');"

            PlanPoint = TIMS.Get_RblPlanPoint(Me, PlanPoint, objconn)
            Common.SetListItem(PlanPoint, "1")

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button5_Click(sender, e)
            End If
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

    End Sub

    '檢查輸入資料是否正確
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        '訓練機構為必填／若無只能查詢自身
        RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Name.Text = TIMS.ClearSQM(Name.Text)
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        If RIDValue.Value = "" Then
            Errmsg += "機構選擇有誤，請重新選擇!" & vbCrLf
        End If

        If STDate1.Text <> "" Then
            'STDate1.Text = Trim(STDate1.Text)
            If Not TIMS.IsDate1(STDate1.Text) Then
                Errmsg += "開訓區間 起始日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate1.Text = CDate(STDate1.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If STDate2.Text <> "" Then
            'STDate2.Text = Trim(STDate2.Text)
            If Not TIMS.IsDate1(STDate2.Text) Then
                Errmsg += "開訓區間 迄止日期格式有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                STDate2.Text = CDate(STDate2.Text).ToString("yyyy/MM/dd")
            End If
        End If

        If Errmsg = "" Then
            If STDate1.Text <> "" AndAlso STDate2.Text <> "" Then
                If CDate(STDate1.Text) > CDate(STDate2.Text) Then
                    Errmsg += "【開訓區間】的起日不得大於【開訓區間】的迄日!!" & vbCrLf
                End If
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function
    '查詢原因
    Private Function GET_SEARCH_MEMO() As String
        Dim RstMemo As String = ""
        center.Text = TIMS.ClearSQM(center.Text)
        STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        STDate2.Text = TIMS.ClearSQM(STDate2.Text)
        IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        Name.Text = TIMS.ClearSQM(Name.Text)
        Dim v_Radio1 As String = TIMS.GetListValue(Radio1) 'AppStage.SelectedValue
        If center.Text <> "" Then RstMemo &= String.Concat("&訓練機構=", center.Text)
        If STDate1.Text <> "" Then RstMemo &= String.Concat("&開訓期間1=", STDate1.Text)
        If STDate2.Text <> "" Then RstMemo &= String.Concat("&開訓期間2=", STDate2.Text)
        If IDNO.Text <> "" Then RstMemo &= String.Concat("&身分證號碼=", IDNO.Text)
        If Name.Text <> "" Then RstMemo &= String.Concat("&學員姓名=", Name.Text)
        If v_Radio1 <> "" Then RstMemo &= String.Concat("&學員狀態=", v_Radio1)
        Return RstMemo
    End Function
    '查詢
    Sub SSearch1()
        'trlimitedmax999.Visible = False
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)
        'flgCIShow = (TIMS.GetListValue(rblWorkMode) = "2") '2:True '可正常顯示個資。1:FALSE:模糊顯示

        '訓練機構為必填／若無只能查詢自身
        'RIDValue.Value = TIMS.ClearSQM(RIDValue.Value)
        'If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'IDNO.Text = TIMS.ClearSQM(IDNO.Text)
        'Name.Text = TIMS.ClearSQM(Name.Text)
        'STDate1.Text = TIMS.ClearSQM(STDate1.Text)
        'STDate2.Text = TIMS.ClearSQM(STDate2.Text)

        Dim s_TOPSEL As String = If(sm.UserInfo.LID = 0, String.Concat("TOP ", cst_ilimitedmax), "")

        Dim sql As String = ""
        sql &= String.Concat(" SELECT ", s_TOPSEL, " a.SID") & vbCrLf
        sql &= " ,a.STUDID2 ,dbo.FN_GET_MASK1(a.IDNO) IDNO" & vbCrLf
        sql &= " ,a.NAME" & vbCrLf
        'sql &= " ,format(a.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " ,a.SOCID" & vbCrLf
        sql &= " ,a.ClassCName" & vbCrLf
        sql &= " ,a.CyclType" & vbCrLf
        sql &= " ,a.CLASSCNAME2" & vbCrLf
        sql &= " ,CONVERT(varchar, a.STDate, 111) STDate" & vbCrLf
        sql &= " ,CONVERT(varchar, a.FTDate, 111) FTDate" & vbCrLf
        sql &= " FROM dbo.V_STUDENTINFO a" & vbCrLf
        'sql &= " JOIN dbo.AUTH_RELSHIP b ON a.RID=b.RID and a.PlanID=b.PlanID" & vbCrLf
        'sql &= " JOIN dbo.ORG_ORGINFO c ON b.OrgID=c.OrgID" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SUBSIDYCOST sc ON sc.SOCID=a.SOCID" & vbCrLf
        sql &= " WHERE a.STUDSTATUS NOT IN (2,3)" & vbCrLf '排除離退訓
        sql &= " and a.TPlanID='" & sm.UserInfo.TPlanID & "'"

        If Not sm.UserInfo.RID = "A" Then
            sql &= " and a.RID='" & RIDValue.Value & "'"
            sql &= " and a.PlanID='" & sm.UserInfo.PlanID & "'"
        ElseIf Len(RIDValue.Value) <> 1 Then
            sql &= " and a.RID='" & RIDValue.Value & "'"
        End If

        If OCIDValue1.Value <> "" Then sql &= " and a.OCID='" & OCIDValue1.Value & "'"

        If IDNO.Text <> "" Then sql &= " and a.IDNO='" & IDNO.Text & "'"

        If Name.Text <> "" Then sql &= " and a.Name like '%" & Name.Text & "%'"

        If STDate1.Text <> "" Then sql &= " and a.STDate>= " & TIMS.To_date(STDate1.Text) & vbCrLf

        If STDate2.Text <> "" Then sql &= " and a.STDate<= " & TIMS.To_date(STDate2.Text) & vbCrLf

        '資料都沒有填寫，限定年度
        If OCIDValue1.Value = "" AndAlso IDNO.Text = "" AndAlso Name.Text = "" AndAlso STDate1.Text = "" AndAlso STDate2.Text = "" Then
            sql &= String.Concat(" and a.YEARS='", sm.UserInfo.Years, "'")
        End If

        '已結訓且已申請經費
        If Radio1.SelectedIndex = 1 Then
            sql &= " AND sc.SOCID IS NOT NULL"
            'SearchStr += " and a.SOCID IN (SELECT SOCID FROM Stud_SubsidyCost)"
        End If
        '28:產業人才投資方案
        KindValue.Value = "" 'TIMS.GetTPlanName(sm.UserInfo.TPlanID)
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            If PlanPoint.SelectedValue = "1" Then
                sql &= " and a.OrgKind2='G' "
                KindValue.Value = "G"
            Else
                sql &= " and a.OrgKind2='W' "
                KindValue.Value = "W"
            End If
        End If
        sql &= " ORDER BY a.STUDID2"

        Dim dt As DataTable = Nothing
        Try
            dt = DbAccess.GetDataTable(sql, objconn)
        Catch ex As Exception
            Common.MessageBox(Me, TIMS.cst_ErrorMsg9) '"資料庫效能異常，請重新查詢")
            'Common.MessageBox(Me, ex.ToString)
            Dim strErrmsg As String = ""
            strErrmsg += "/*  sql: */" & vbCrLf
            strErrmsg += sql & vbCrLf
            strErrmsg += "/*  ex.ToString: */" & vbCrLf
            strErrmsg += ex.ToString & vbCrLf
            strErrmsg += TIMS.GetErrorMsg(Me) '取得錯誤資訊寫入
            'strErrmsg = Replace(strErrmsg, vbCrLf, "<br>" & vbCrLf)
            Call TIMS.WriteTraceLog(strErrmsg)
            Exit Sub
        End Try

        'SOCID,STUDID2,NAME,IDNO,BIRTHDAY,CLASSCNAME2,STDATE,FTDATE
        '查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        Dim MRqID As String = TIMS.Get_MRqID(Me)
        sMemo = GET_SEARCH_MEMO()
        '查詢原因:INQNO--查詢結果筆數:RESCNT--查詢清單結果:RESDESC 
        Dim vRESDESC As String = TIMS.GET_RESDESCdt(dt, "SOCID,STUDID2,NAME,IDNO")
        Call TIMS.SubInsAccountLog1(Me, MRqID, TIMS.cst_wm查詢, TIMS.cst_wmdip2, OCIDValue1.Value, sMemo, objconn, v_INQUIRY, dt.Rows.Count, vRESDESC)

        DataGridTable.Visible = False
        msg.Visible = True
        msg.Text = "查無資料"
        SOCIDValue.Value = "" '(DataGrid1_ItemDataBound)

        trlimitedmax999.Visible = False
        If dt.Rows.Count = 0 Then Return

        sMemo = $"&OCID={OCIDValue1.Value}&IDNO={IDNO.Text}&sql={sql}"
        Dim s_FUNID As String = TIMS.Get_MRqID(Me)
        TIMS.SubInsAccountLog1(Me, s_FUNID, TIMS.cst_wm查詢, 2, OCIDValue1.Value, sMemo)

        DataGridTable.Visible = True
        msg.Visible = False
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()

        trlimitedmax999.Visible = (s_TOPSEL <> "")
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If
        '取出鍵詞-查詢原因
        Dim v_INQUIRY As String = TIMS.GetListValue(ddl_INQUIRY_Sch)
        If (TIMS.Utl_GetConfigSet(TIMS.cst_appkey_INQUIRY).Equals(TIMS.cst_YES)) Then
            If (v_INQUIRY = "") Then Common.MessageBox(Me, "請選擇「查詢原因」") : Return
            Session(String.Concat(TIMS.cst_GSE_V_INQUIRY, TIMS.Get_MRqID(Me))) = v_INQUIRY
        End If

        Call SSearch1()
    End Sub

    Private Sub DataGrid1_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid1.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Header
                'e.Item.CssClass = "head_navy"
                Dim checkbox3 As HtmlInputCheckBox = e.Item.FindControl("checkbox3")
                checkbox3.Attributes("onclick") = "ChangeAll(this);"
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                'checkbox3'checkbox2'hf_SOCID'hf_SID
                Dim checkbox2 As HtmlInputCheckBox = e.Item.FindControl("checkbox2")
                Dim hf_SOCID As HiddenField = e.Item.FindControl("hf_SOCID")
                Dim hf_SID As HiddenField = e.Item.FindControl("hf_SID")
                checkbox2.Value = Convert.ToString(drv("SOCID"))
                hf_SOCID.Value = Convert.ToString(drv("SOCID"))
                hf_SID.Value = Convert.ToString(drv("SID"))

                '不可顯示個資。
                'If Not flgCIShow Then
                'e.Item.Cells(cst_DG1_身分證號碼).Text = TIMS.strMask(e.Item.Cells(cst_DG1_身分證號碼).Text, 1)
                'e.Item.Cells(cst_DG1_出生日期).Text = TIMS.strMask(e.Item.Cells(cst_DG1_出生日期).Text, 2)
                'End If

                'If e.Item.ItemType = ListItemType.Item Then e.Item.CssClass = ""
                'Dim DTR As HtmlTableRow = e.Item.FindControl("DTR")
                'Dim IMG1 As HtmlImage = e.Item.FindControl("IMG1")
                'IMG1.Attributes("onclick") = "ShowDetail('" & DTR.ClientID & "','" & IMG1.ClientID & "')"

                'Dim DetailTable As Table = e.Item.FindControl("DetailTable")
                'Dim MyRow As TableRow = Nothing
                'Dim MyCell As TableCell = Nothing
                'Dim AllCheckBox As New HtmlInputCheckBox
                'TIMS.CreateRow(DetailTable, MyRow)
                'TIMS.CreateCell(MyRow, MyCell)
                'AllCheckBox.Attributes("onclick") = "SelectAll(this.checked," & e.Item.ItemIndex + 1 & ")"
                'MyCell.Controls.Add(AllCheckBox)
                'TIMS.CreateCell(MyRow, MyCell)
                'MyCell.Text = "全選"
                'Call TIMS.CreateRow(DetailTable, MyRow)
                'Call TIMS.CreateCell(MyRow, MyCell)

                'Dim CheckBox1 As New HtmlInputCheckBox
                'CheckBox1.Value = drv("SOCID")
                'CheckBox1.Attributes("onclick") = "SelectItem(this.checked,this.value);"
                'If OCIDValue1.Value <> "" AndAlso Convert.ToString(drv("SOCID")) <> "" Then
                '    CheckBox1.Checked = True
                '    If SOCIDValue.Value.IndexOf(Convert.ToString(drv("SOCID"))) = -1 Then
                '        If SOCIDValue.Value <> "" Then SOCIDValue.Value &= ","
                '        SOCIDValue.Value &= Convert.ToString(drv("SOCID"))
                '    End If
                'End If
                'MyCell.Controls.Add(CheckBox1)
                'MyCell.Width = Unit.Pixel(23)
                'Call TIMS.CreateCell(MyRow, MyCell)
                'MyCell.Text = Convert.ToString(drv("CLASSCNAME2")) '.ToString & "第" & CStr(drv("CyclType")) & "期"
                'MyCell.Text &= "(" & drv("STDate") & "~" & drv("FTDate") & ")"
                ''If ShowTR Then
                ''    DTR.Style("display") = "inline"
                ''    IMG1.Src = "../../images/n02.gif"
                ''Else
                ''    DTR.Style("display") = "none"
                ''    IMG1.Src = "../../images/n01.gif"
                ''End If
                'IMG1.Src = "../../images/n02.gif"
                'DTR.Style("display") = ""
        End Select
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGridTable.Visible = False
        msg.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGridTable.Visible = False
        msg.Visible = False
    End Sub

    '列印
    Protected Sub BtnPrint3_Click(sender As Object, e As EventArgs) Handles BtnPrint3.Click
        Dim s_SOCIDVals As String = ""
        For Each eItem As DataGridItem In DataGrid1.Items
            Dim checkbox2 As HtmlInputCheckBox = eItem.FindControl("checkbox2")
            Dim hf_SOCID As HiddenField = eItem.FindControl("hf_SOCID")
            hf_SOCID.Value = TIMS.ClearSQM(hf_SOCID.Value)
            If checkbox2.Checked AndAlso hf_SOCID.Value <> "" Then
                s_SOCIDVals &= String.Concat(If(s_SOCIDVals <> "", ",", ""), hf_SOCID.Value)
            End If
        Next
        SOCIDValue.Value = s_SOCIDVals
        SOCIDValue.Value = TIMS.ClearSQM(SOCIDValue.Value)

        If SOCIDValue.Value = "" Then '(DataGrid1_ItemDataBound)
            Common.MessageBox(Me, "請先勾選要列印的學員")
            Exit Sub
        End If
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case KindValue.Value
                Case "G", "W" 'G:產業人才投資計畫'W:提升勞工自主學習計畫
                Case Else
                    Common.MessageBox(Me, "請先勾選要列印的計畫")
                    Exit Sub
            End Select
        End If
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        '學員資料審核鈕
        If Convert.ToString(drCC("AppliedResultR")) <> "Y" Then
            Common.MessageBox(Me, "學員資料尚未審核!!")
            Exit Sub
        End If

        'OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'SOCIDValue.Value = TIMS.ClearSQM(SOCIDValue.Value)
        Dim v_OCIDValue As String = OCIDValue1.Value
        If v_OCIDValue = "" Then v_OCIDValue = TIMS.Get_OCIDforSOCID(SOCIDValue.Value, objconn)

        iPYNum = TIMS.sUtl_GetPYNum(Me)
        Dim filename As String = "" 'cst_printFN1 '"SD_14_013_2016" '其他非產投計畫使用報表列印
        Dim myValue As String = ""
        TIMS.SetMyValue(myValue, "RID", RIDValue.Value)
        TIMS.SetMyValue(myValue, "OCID", v_OCIDValue)
        TIMS.SetMyValue(myValue, "SOCID", SOCIDValue.Value)
        TIMS.SetMyValue(myValue, "Years", Years.Value)

        filename = cst_printFNoth2 '非產投計畫使用
        If TIMS.Cst_TPlanID28.IndexOf(sm.UserInfo.TPlanID) > -1 Then
            Select Case KindValue.Value
                Case "G"
                    filename = cst_printFN18G
                Case "W"
                    filename = cst_printFN18W
            End Select
        End If

        sMemo = $"{myValue}&prt={filename}"
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm列印, 2, OCIDValue1.Value, sMemo)
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, filename, myValue)
    End Sub

End Class
