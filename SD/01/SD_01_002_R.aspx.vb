Partial Class SD_01_002_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "member_list"

    '報名管道
    Const cst_rw2不區分 As String = "A"
    Const cst_rw2一般推介單 As String = "CH4"
    Const cst_rw2免試推介單 As String = "EPW"
    Const cst_rw2專案核定報名 As String = "EP2P"
    '<asp@ListItem Value="A" Selected="True">不區分</asp@ListItem>
    '<asp@ListItem Value="CH4">一般推介單</asp@ListItem>
    '<asp@ListItem Value="EPW">免試推介單</asp@ListItem>
    '<asp@ListItem Value="EP2P">專案核定報名</asp@ListItem>

    'Select concat('/',IDENTITYID,'.',name) idnm
    'From dbo.KEY_IDENTITY WITH(NOLOCK)
    'Where IDENTITYID Not IN (Select IDENTITYID FROM dbo.PLAN_IDENTITY With(NOLOCK) WHERE isEnabled = 'N' and TPLANID='06')

    'iReport: member_list
    Const cst_EnterPathW As String = "W" '就服站代碼
    Const cst_EnterPathNameW As String = "<br />(就服單位協助報名)" '說明
    Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)

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
        '檢查Session是否存在 End

        blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        If blnP0 Then Trwork2013a.Visible = True

        ''就服單位協助報名
        'Trwork2013a.Visible = False
        'If sm.UserInfo.Years >= 2013 _
        '    AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then
        '        Trwork2013a.Visible = True
        '    End If
        'End If

        'PageControler1 = Me.FindControl("PageControler1")
        PageControler1.PageDataGrid = DataGrid2
        If Not IsPostBack Then
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            PageControler1.Visible = False

            '直接列印
            Button1.Attributes.Add("OnClick", "return chk();")
            'Button1.Attributes("onclick") = "javascript:return chk();"

            '查詢
            Button3.Attributes.Add("OnClick", "return chk();")
            'Button3.Attributes("onclick") = "javascript:return chk();"
            Button2.Attributes.Add("OnClick", "return chk();") '匯出

            '群組列印
            Button4.Attributes.Add("OnClick", "return CheckPrint();") '群組列印(up)
            '群組列印
            Button5.Attributes.Add("OnClick", "return CheckPrint();") '群組列印(down)

            Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg.aspx?name=Stu_Maintain');"
            Const cst_javascript_openOrg_FMT2 As String = "javascript:openOrg('../../Common/LevOrg1.aspx');"
            Button8.Attributes("onclick") = If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, cst_javascript_openOrg_FMT1, cst_javascript_openOrg_FMT2)

            'PageControler1.PageDataGrid = DataGrid2

            ''090406 andy edit
            ''-----------
            'If RIDValue.Value <> "" Then sm.UserInfo.RID = RIDValue.Value
            ''------------
            'If RIDValue.Value <> "" Then
            '    Me.RID.Value = RIDValue.Value
            'Else
            '    Me.RID.Value = sm.UserInfo.RID
            'End If
            'Me.DistID.Value = sm.UserInfo.DistID
            'Me.TPlanID.Value = sm.UserInfo.TPlanID

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button13_Click(sender, e)
            End If
        End If

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        'TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", , , "TMIDValue1", "TMID1")
        'If HistoryTable.Rows.Count <> 0 Then
        '    OCID1.Attributes("onclick") = "showObj('HistoryList');"
        '    OCID1.Style("CURSOR") = "hand"
        'End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

    End Sub

    '直接列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Call sUtl_Print1()
    End Sub

    '匯出(Excel)
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim dt As DataTable
        Dim sql As String = ""
        Dim parms As Hashtable = New Hashtable()
        sql = sUtl_Search1_SQL(2, parms) '1:查詢sql 2:匯出sql
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        Call ExportX1(dt)

    End Sub

    Sub ExportX1(ByRef dt As DataTable)
        Dim sFileName1 As String = "student" & TIMS.GetDateNo2()
        'Dim strSTYLE As String = ""
        'strSTYLE &= ("<style>")

        Dim strHTML As String = ""
        strHTML &= ("<div>")
        strHTML &= ("<table>")

        'Response.AddHeader("content-disposition", "attachment; filename=" & HttpUtility.UrlEncode("student", System.Text.Encoding.UTF8) & ".xls")
        'Response.ContentType = "Application/octet-stream"
        'Response.ContentEncoding = System.Text.Encoding.GetEncoding("Big5")

        Dim ExportStr As String = ""
        '建立輸出文字
        'ExportStr = "報名職類" & vbTab & "報名編號" & vbTab & "姓名" & vbTab & "報名日期" & vbTab & "地址" & vbTab & "聯絡電話" & vbTab
        'ExportStr = "報名職類" & vbTab & "報名編號" & vbTab & "姓名" & vbTab & "報名日期" & vbTab & "性別" & vbTab & "生日" & vbTab & "身分證號碼" & vbTab & "學校" & vbTab & "科系" & vbTab & "地址" & vbTab & "聯絡電話" & vbTab & "備註" & vbTab
        '欄名「 報名職類」改為「班別名稱」，「報名編號」改為「報名流水編號」
        'ExportStr = "班別名稱" & vbTab & "報名流水編號" & vbTab & "姓名" & vbTab & "報名日期" & vbTab & "身分別" & vbTab & "性別" & vbTab & "生日" & vbTab & "身分證號碼" & vbTab & "學校" & vbTab & "科系" & vbTab & "地址" & vbTab & "聯絡電話" & vbTab & "備註" & vbTab
        'ExportStr = "班別名稱" & vbTab & "報名流水編號" & vbTab & "姓名" & vbTab & "報名日期" & vbTab & "身分別" & vbTab & "性別" & vbTab & "生日" & vbTab & "身分證號碼" & vbTab & "學校" & vbTab & "科系" & vbTab & "地址" & vbTab & "聯絡電話" & vbTab & "E-mail" & vbTab & "備註" & vbTab  '090406 andy edit 加入email欄位
        '090424 andy edit 加入行動電話欄位

        '班別名稱	報名流水編號	准考證號碼	姓名	
        '報名日期	身分別	性別	學校	科系	E-mail	畢業狀況	最高學歷	兵役狀況	報名管道	備註

        ExportStr = ""
        ExportStr &= "班別名稱" & vbTab
        ExportStr &= "報名流水編號" & vbTab
        ExportStr &= "准考證號碼" & vbTab
        ExportStr &= "姓名" & vbTab

        ExportStr &= "報名日期" & vbTab
        ExportStr &= "身分別" & vbTab
        ExportStr &= "性別" & vbTab
        'ExportStr &= "生日" & vbTab
        'ExportStr &= "身分證號碼" & vbTab
        ExportStr &= "學校" & vbTab
        ExportStr &= "科系" & vbTab
        'ExportStr &= "地址" & vbTab
        'ExportStr &= "聯絡電話" & vbTab
        ExportStr &= "E-mail" & vbTab
        'ExportStr &= "行動電話" & vbTab
        ExportStr &= "畢業狀況" & vbTab
        'ExportStr &= "婚姻狀況" & vbTab
        ExportStr &= "最高學歷" & vbTab
        ExportStr &= "兵役狀況" & vbTab
        ExportStr &= "報名管道" & vbTab
        'ExportStr &= "是否同意" & vbTab
        ExportStr &= "備註"
        strHTML &= TIMS.Get_TABLETR(Replace(ExportStr, vbTab, ","))

        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        '建立資料面
        Dim i_ROW As Integer = 0
        For Each dr As DataRow In dt.Rows
            i_ROW += 1
            ExportStr = ""
            ExportStr &= Convert.ToString(dr("ClassCName")) & vbTab '班別名稱
            ExportStr &= Convert.ToString(i_ROW) & vbTab '報名流水編號
            ExportStr &= Convert.ToString(dr("EXAMNO")) & vbTab '准考證號碼
            ExportStr &= Convert.ToString(dr("Name")) & vbTab '姓名

            'ExportStr &= Convert.ToString( FormatDateTime(dr("RelEnterDate"), DateFormat.ShortDate) & vbTab '報名日期
            ExportStr &= Convert.ToString(TIMS.Cdate3(dr("RelEnterDate"))) & vbTab '報名日期
            '單一或多重身分別
            Dim IdentityName As String = If(dr("IdentityID").ToString <> "", TIMS.Get_IdentityName(dr("IdentityID").ToString, objconn), "")
            ExportStr &= Convert.ToString(IdentityName) & vbTab
            ExportStr &= Convert.ToString(dr("Sex")) & vbTab '性別

            'ExportStr &= TIMS.cdate3(dr("Birthday")) & vbTab '生日
            'ExportStr &= Convert.ToString(dr("IDNO")) & vbTab '身分證號碼
            ExportStr &= Convert.ToString(dr("School")) & vbTab '學校
            ExportStr &= Convert.ToString(dr("Department")) & vbTab '科系'
            'START修改因ZipCode錯誤join後資料無法顯示 20090717 by waiming
            'ExportStr &= Convert.ToString( dr("Address") & vbTab '地址
            'ExportStr &= Convert.ToString( dr("ZipID1") & dr("ZipID2") & dr("CityName") & dr("ZipName") & dr("Address") & vbTab '地址
            'ExportStr &= Convert.ToString(dr("ZipID1") & dr("ZipID2") & dr("ZipName") & dr("Address")) & vbTab '地址 'END修改因ZipCode錯誤join後資料無法顯示
            'ExportStr &= Convert.ToString(dr("Phone1")) & vbTab   '連絡電話
            ExportStr &= Convert.ToString(dr("EMAIL")) & vbTab ' email   '090406 andy edit 加入email欄位
            'ExportStr &= Convert.ToString(dr("CellPhone")) & vbTab        '090424 andy edit 加入行動電話欄位
            ExportStr &= Convert.ToString(dr("GradName")) & vbTab '畢業狀況
            'ExportStr &= Convert.ToString(dr("Marital")) & vbTab '婚姻狀況
            ExportStr &= Convert.ToString(dr("DegreeName")) & vbTab '最高學歷
            ExportStr &= Convert.ToString(dr("Military")) & vbTab '兵役狀況
            ExportStr &= Convert.ToString(dr("EnterChannel")) & vbTab '報名管道
            'ExportStr &= Convert.ToString(dr("IsAgree")) & vbTab '是否同意
            ExportStr &= Convert.ToString(dr("Notes"))  '備註
            'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))

            strHTML &= TIMS.Get_TABLETR(Replace(ExportStr, vbTab, ","))
        Next
        strHTML &= ("</table>")
        strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        'parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "")
        'TIMS.CloseDbConn(objconn)
        'Response.End()
    End Sub

    ''' <summary> 檢核 </summary>
    ''' <param name="Errmsg"></param>
    ''' <returns></returns>
    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        'Select Case sm.UserInfo.LID
        '    Case 2
        '    Case Else
        '        If OCIDValue1.Value = "" Then
        '            Errmsg += "非委訓單位，請選擇1班級資訊" & vbCrLf
        '        End If
        'End Select

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If OCIDValue1.Value <> "" AndAlso drCC Is Nothing Then
            Errmsg += TIMS.cst_NODATAMsg2 & vbCrLf
        End If

        If start_date.Text <> "" Then
            start_date.Text = Trim(start_date.Text)
            If TIMS.IsDate1(start_date.Text) Then
                start_date.Text = CDate(start_date.Text).ToString("yyyy/MM/dd")
            Else
                Errmsg += "開訓日期 起始日期格式有誤" & vbCrLf
            End If
        End If
        If end_date.Text <> "" Then
            end_date.Text = Trim(end_date.Text)
            If TIMS.IsDate1(end_date.Text) Then
                end_date.Text = CDate(end_date.Text).ToString("yyyy/MM/dd")
            Else
                Errmsg += "開訓日期 迄止日期格式有誤" & vbCrLf
            End If
        End If
        If EnterDate_start.Text <> "" Then
            EnterDate_start.Text = Trim(EnterDate_start.Text)
            If TIMS.IsDate1(EnterDate_start.Text) Then
                EnterDate_start.Text = CDate(EnterDate_start.Text).ToString("yyyy/MM/dd")
            Else
                Errmsg += "報名日期 起始日期格式有誤" & vbCrLf
            End If
        End If
        If EnterDate_end.Text <> "" Then
            EnterDate_end.Text = Trim(EnterDate_end.Text)
            If TIMS.IsDate1(EnterDate_end.Text) Then
                EnterDate_end.Text = CDate(EnterDate_end.Text).ToString("yyyy/MM/dd")
            Else
                Errmsg += "報名日期 迄止日期格式有誤" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    ''' <summary>
    ''' '1:查詢sql 2:匯出sql
    ''' </summary>
    ''' <param name="sType">1:查詢sql 2:匯出sql</param>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function sUtl_Search1_SQL(ByVal sType As Integer, ByRef parms As Hashtable) As String
        Dim rst As String = ""

        '1:查詢sql 2:匯出sql
        ''sql += " AND i.OCID='57359'" & vbCrLf
        Dim sql As String = ""
        sql &= " SELECT b.EXAMNO" & vbCrLf
        sql &= " ,b.SETID" & vbCrLf
        sql &= " ,a.IDNO" & vbCrLf
        sql &= " ,a.Birthday" & vbCrLf
        sql &= " ,i.ClassCName" & vbCrLf
        sql &= " ,i.YEARS" & vbCrLf
        sql &= " ,i.PLANNAME" & vbCrLf
        sql &= " ,i.OrgName" & vbCrLf
        sql &= " ,i.TrainName" & vbCrLf
        sql &= " ,a.Name" & vbCrLf
        sql &= " ,b.RelEnterDate" & vbCrLf
        sql &= " ,ISNULL(a.Phone1,'') Phone1" & vbCrLf
        sql &= " ,ISNULL(a.Phone2,'') Phone2" & vbCrLf
        sql &= " ,ISNULL(a.CellPhone,'') CellPhone" & vbCrLf
        sql &= " ,ISNULL(dbo.FN_GZIP3(a.ZipCode),'') ZipID1" & vbCrLf
        sql &= " ,ISNULL(dbo.FN_GZIP2(a.ZipCODE6W),'') ZipID2" & vbCrLf
        sql &= " ,g.CTName CityName" & vbCrLf
        sql &= " ,g.ZipName" & vbCrLf
        sql &= " ,a.Address" & vbCrLf
        sql &= " ,ISNULL(kg.Name,'') GradName" & vbCrLf
        sql &= " ,dbo.DECODE10(a.Sex,'M','男','F','女','1','男','2','女','女') Sex" & vbCrLf
        sql &= " ,a.School" & vbCrLf
        sql &= " ,a.Department" & vbCrLf
        sql &= " ,a.email" & vbCrLf
        sql &= " ,b.Notes" & vbCrLf
        sql &= " ,b.IdentityID" & vbCrLf
        sql &= " ,a.MaritalStatus" & vbCrLf
        sql &= " ,dbo.DECODE6(a.MaritalStatus,1,'已婚',2,'未婚','未填寫') Marital" & vbCrLf
        sql &= " ,ISNULL(kd.Name,'未填寫') DegreeName" & vbCrLf
        sql &= " ,ISNULL(km.Name,'未填寫') Military" & vbCrLf
        sql &= " ,dbo.DECODE10(b.EnterChannel,1,'網路',2,'現場',3,'通訊',4,'推介','網路') EnterChannel" & vbCrLf
        sql &= " ,dbo.DECODE6(a.IsAgree,'Y','Y','N','N','需確認') IsAgree" & vbCrLf
        '大於0 以「*」星號標註在訓中之學員，惟不可以本項標註作為甄試資格不符或不予錄訓之依據。
        sql &= " ,CASE WHEN dbo.FN_CHKSTUDCOUNT(a.IDNO,b.OCID1) > 0 THEN 1 ELSE 0 END StudCount" & vbCrLf
        sql &= " FROM dbo.STUD_ENTERTEMP a" & vbCrLf
        sql &= " JOIN dbo.STUD_ENTERTYPE b ON a.SETID=b.SETID" & vbCrLf
        sql &= " JOIN dbo.VIEW2 i on i.OCID=b.OCID1" & vbCrLf
        sql &= " LEFT JOIN dbo.VIEW_ZIPNAME g ON g.ZipCode = a.ZipCode" & vbCrLf
        sql &= " LEFT JOIN dbo.Key_GradState kg ON kg.GradID=a.GradID" & vbCrLf
        sql &= " LEFT JOIN dbo.Key_Degree kd ON kd.DegreeID = a.DegreeID" & vbCrLf
        sql &= " LEFT JOIN dbo.Key_Military km ON km.MilitaryID = a.MilitaryID" & vbCrLf
        'sql &= " AND i.TPLANID='06' AND i.YEARS='2019'" & vbCrLf
        sql &= " WHERE i.TPlanID = @TPlanID " & vbCrLf
        parms.Add("TPlanID", sm.UserInfo.TPlanID)

        Select Case sm.UserInfo.LID
            Case 2
                sql &= " AND i.DistID = @DistID " & vbCrLf
                parms.Add("DistID", sm.UserInfo.DistID)
        End Select

        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        sql &= " AND i.RID = @RID " & vbCrLf
        parms.Add("RID", RIDValue.Value)

        '條件值
        If Me.OCIDValue1.Value <> "" Then
            'i.OCID
            sql &= " AND b.OCID1 = @OCID1 " & vbCrLf
            sql &= " AND i.OCID = @OCID " & vbCrLf
            parms.Add("OCID1", Me.OCIDValue1.Value)
            parms.Add("OCID", Me.OCIDValue1.Value)
        End If
        If Me.cjobValue.Value <> "" Then
            sql &= " AND i.CJOB_UNKEY = @CJOB_UNKEY " & vbCrLf
            parms.Add("CJOB_UNKEY", Me.cjobValue.Value)
        End If
        If Me.start_date.Text <> "" Then
            sql &= " AND i.STDate >= @STDate1 " & vbCrLf
            parms.Add("STDate1", Me.start_date.Text)
        End If
        If Me.end_date.Text <> "" Then
            sql &= " AND i.STDate <= @STDate2 " & vbCrLf
            parms.Add("STDate2", Me.end_date.Text)
        End If

        If Me.EnterDate_start.Text <> "" Then
            sql &= " AND b.RelEnterDate >= @RelEnterDate1 " & vbCrLf
            parms.Add("RelEnterDate1", Me.EnterDate_start.Text)
        End If
        If Me.EnterDate_end.Text <> "" Then
            sql &= " AND b.RelEnterDate <= @RelEnterDate2 " & vbCrLf
            parms.Add("RelEnterDate2", Me.EnterDate_end.Text)
        End If

        'Select Case rblEnterPathW.SelectedValue
        '    Case "Y" '是 就服單位協助報名
        '        sql &= " AND ISNULL(b.EnterPath,' ') = '" & cst_EnterPathW & "' " & vbCrLf
        '    Case "N" '不是 就服單位協助報名
        '        sql &= " AND ISNULL(b.EnterPath,' ') != '" & cst_EnterPathW & "' " & vbCrLf
        'End Select
        Select Case rblEnterPathW2.SelectedValue
            Case cst_rw2不區分
            Case cst_rw2一般推介單
                sql &= " AND ISNULL(b.ENTERCHANNEL,0) = 4 " & vbCrLf
                sql &= " AND ISNULL(b.ENTERPATH,' ') != 'W' " & vbCrLf '排除免試
                sql &= " AND ISNULL(b.ENTERPATH2,' ') != 'P' " '專案核定
            Case cst_rw2免試推介單
                sql &= " AND ISNULL(b.ENTERPATH,' ') = 'W' " & vbCrLf
            Case cst_rw2專案核定報名
                sql &= " AND ISNULL(b.ENTERPATH2,' ') = 'P' " & vbCrLf
        End Select

        '1:查詢sql 2:匯出sql
        Select Case sType
            Case 1
                '2009/06/11 改成依准考證號排序
                sql &= " ORDER BY i.OCID, b.EXAMNO, b.RelEnterDate" & vbCrLf
            Case Else
                sql &= " ORDER BY i.OCID, b.RelEnterDate" & vbCrLf
        End Select
        rst = sql
        Return rst
    End Function

    Sub sSearch1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        Dim dt As DataTable
        Dim sql As String = ""
        Dim parms As Hashtable = New Hashtable()

        '1:查詢sql 2:匯出sql
        sql = sUtl_Search1_SQL(1, parms)
        dt = DbAccess.GetDataTable(sql, objconn, parms)

        Table5.Visible = False
        msg.Text = "查無資料!!"
        msg.Visible = True
        Me.Button4.Visible = False
        Me.Button5.Visible = False
        DataGrid2.Visible = False
        PageControler1.Visible = False

        If dt.Rows.Count = 0 Then Exit Sub

        Table5.Visible = True
        msg.Text = ""
        msg.Visible = False
        Me.Button4.Visible = True
        Me.Button5.Visible = True
        DataGrid2.Visible = True
        PageControler1.Visible = True
        'PageControler1.SqlString = sql
        PageControler1.PageDataTable = dt '.SqlString = sql
        PageControler1.ControlerLoad()
    End Sub

    '查詢
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call sSearch1()
    End Sub

    Private Sub DataGrid2_ItemDataBound(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.DataGridItemEventArgs) Handles DataGrid2.ItemDataBound
        Select Case e.Item.ItemType
            Case ListItemType.Item, ListItemType.AlternatingItem
                Dim drv As DataRowView = e.Item.DataItem
                Dim EXAMNO As HtmlInputHidden = e.Item.FindControl("EXAMNO")
                Dim HIDNO As HtmlInputHidden = e.Item.FindControl("HIDNO")
                Dim labStar1 As Label = e.Item.FindControl("labStar1")
                Dim labName As Label = e.Item.FindControl("labName")
                Dim labRelEnterDate As Label = e.Item.FindControl("labRelEnterDate")

                labStar1.Visible = If(Convert.ToString(drv("StudCount")) = "1", True, False)

                EXAMNO.Value = "'" & drv("EXAMNO") & "'"
                HIDNO.Value = "'" & drv("IDNO") & "'"
                'e.Item.Cells(5).Text = FormatDateTime(drv("RelEnterDate"), 2)
                'e.Item.Cells(5).Text = Common.FormatDate(drv("RelEnterDate"))
                labName.Text = Convert.ToString(drv("Name"))
                labRelEnterDate.Text = If(Convert.ToString(drv("RelEnterDate")) <> "", Common.FormatDate(drv("RelEnterDate")), "")
        End Select
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn) '判斷機構是否只有一個班級
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        DataGrid2.Visible = False
        Me.Button4.Visible = False
        Me.Button5.Visible = False
        msg.Visible = False
        PageControler1.Visible = False
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        '如果只有一個班級
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
        DataGrid2.Visible = False
        Me.Button4.Visible = False
        Me.Button5.Visible = False
        msg.Visible = False
        PageControler1.Visible = False
    End Sub

    '直接列印
    Sub sUtl_Print1()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=member_list&ExamID=' + ExamID + 
        '&DistID=' + document.getElementById('DistID').value + 
        '&OCID1=' + document.getElementById('OCIDValue1').value + 
        '&CJOB_UNKEY=' + document.getElementById('cjobValue').value + '&RID=' + document.getElementById('RID').value +
        '&TPlanID=' + document.getElementById('TPlanID').value + 
        '&STDate1=' + document.getElementById('start_date').value + '&STDate2=
        ' + document.getElementById('end_date').value + '&RelEnterDate1=' + document.getElementById('EnterDate_start').value +
        '&RelEnterDate2=' + document.getElementById('EnterDate_end').value);
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If OCIDValue1.Value <> "" AndAlso drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        Dim myvalue As String = ""
        myvalue = ""
        'myvalue &= "&ExamID=" & Replace(hidExamID.Value, "'", "\'")
        myvalue &= "&TPlanID=" & sm.UserInfo.TPlanID
        myvalue &= "&OCID1=" & OCIDValue1.Value
        myvalue &= "&CJOB_UNKEY=" & cjobValue.Value
        Select Case sm.UserInfo.LID
            Case 0
                myvalue &= "&DistID=" & s_DistID 'sm.UserInfo.DistID
                myvalue &= "&RID=" & RIDValue.Value
            Case Else
                myvalue &= "&DistID=" & sm.UserInfo.DistID
                myvalue &= "&RID=" & sm.UserInfo.RID
        End Select
        myvalue &= "&STDate1=" & start_date.Text
        myvalue &= "&STDate2=" & end_date.Text
        myvalue &= "&RelEnterDate1=" & EnterDate_start.Text
        myvalue &= "&RelEnterDate2=" & EnterDate_end.Text

        '就服單位協助報名
        'Select Case rblEnterPathW.SelectedValue
        '    Case "A"
        '    Case "Y", "N"
        '        myvalue &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
        'End Select
        Select Case rblEnterPathW2.SelectedValue
            Case cst_rw2不區分
            Case cst_rw2一般推介單
                myvalue &= "&ENTERPATH1=Y"
            Case cst_rw2免試推介單
                myvalue &= "&ENTERPATH2=Y"
            Case cst_rw2專案核定報名
                myvalue &= "&ENTERPATH3=Y"
        End Select
        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myvalue)
    End Sub

    '群組列印
    Sub sUtl_Print2()
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        'openPrint('../../SQControl.aspx?SQ_AutoLogout=true&sys=member&filename=member_list&ExamID=' + ExamID + 
        '&DistID=' + document.getElementById('DistID').value + 
        '&OCID1=' + document.getElementById('OCIDValue1').value + 
        '&CJOB_UNKEY=' + document.getElementById('cjobValue').value + '&RID=' + document.getElementById('RID').value +
        '&TPlanID=' + document.getElementById('TPlanID').value + 
        '&STDate1=' + document.getElementById('start_date').value + '&STDate2=
        ' + document.getElementById('end_date').value + '&RelEnterDate1=' + document.getElementById('EnterDate_start').value +
        '&RelEnterDate2=' + document.getElementById('EnterDate_end').value);

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If OCIDValue1.Value <> "" AndAlso drCC Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If RIDValue.Value = "" Then RIDValue.Value = sm.UserInfo.RID
        Dim s_DistID As String = TIMS.Get_DistID_RID(RIDValue.Value, objconn)

        hidExamID.Value = ""
        Dim IDNOArray As New ArrayList
        For Each eItem As DataGridItem In DataGrid2.Items
            Dim Checkbox1 As HtmlInputCheckBox = eItem.FindControl("Checkbox1")
            Dim EXAMNO As HtmlInputHidden = eItem.FindControl("EXAMNO")
            If Checkbox1.Checked AndAlso EXAMNO.Value <> "" Then
                If (hidExamID.Value <> "") Then hidExamID.Value &= ","
                hidExamID.Value &= EXAMNO.Value
            End If
        Next

        Dim myvalue As String = ""
        myvalue = ""
        myvalue &= "&ExamID=" & Replace(hidExamID.Value, "'", "\'")
        myvalue &= "&TPlanID=" & sm.UserInfo.TPlanID
        Select Case sm.UserInfo.LID
            Case 0
                myvalue &= "&DistID=" & s_DistID 'sm.UserInfo.DistID
                myvalue &= "&RID=" & RIDValue.Value
            Case Else
                myvalue &= "&DistID=" & sm.UserInfo.DistID
                myvalue &= "&RID=" & sm.UserInfo.RID
        End Select
        myvalue &= "&OCID1=" & OCIDValue1.Value
        myvalue &= "&CJOB_UNKEY=" & cjobValue.Value
        myvalue &= "&STDate1=" & start_date.Text
        myvalue &= "&STDate2=" & end_date.Text
        myvalue &= "&RelEnterDate1=" & EnterDate_start.Text
        myvalue &= "&RelEnterDate2=" & EnterDate_end.Text

        '就服單位協助報名
        'Select Case rblEnterPathW.SelectedValue
        '    Case "A"
        '    Case "Y", "N"
        '        myvalue &= "&EnterPath" & rblEnterPathW.SelectedValue & "=W"
        'End Select

        Select Case rblEnterPathW2.SelectedValue
            Case cst_rw2不區分
            Case cst_rw2一般推介單
                myvalue &= "&ENTERPATH1=Y"
            Case cst_rw2免試推介單
                myvalue &= "&ENTERPATH2=Y"
            Case cst_rw2專案核定報名
                myvalue &= "&ENTERPATH3=Y"
        End Select

        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, myvalue)
    End Sub

    '群組列印
    Protected Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Call sUtl_Print2()
    End Sub

    '群組列印
    Protected Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Call sUtl_Print2()
    End Sub

End Class
