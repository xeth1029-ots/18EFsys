Partial Class SD_02_008
    Inherits AuthBasePage

    Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), titlelab1, titlelab2)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        'Dim blnP0 As Boolean = False '報名管道(職前計畫顯示)
        blnP0 = TIMS.Get_TPlanID_P0(Me, objconn)
        Trwork2013a.Visible = False '報名管道(職前計畫顯示)
        If blnP0 Then Trwork2013a.Visible = True

#Region "(No Use)"

        ''就服單位協助報名
        'Trwork2013a.Visible = False
        'If sm.UserInfo.Years >= 2013 _
        '    AndAlso TIMS.Cst_TPlanID0237AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
        '    If TIMS.Utl_GetConfigSet("work2013") = "Y" Then Trwork2013a.Visible = True
        'End If

#End Region

        If Not IsPostBack Then
            cCreate1()
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

        button1.Attributes("onclick") = "javascript:return search();"
    End Sub

    Sub cCreate1()
        '婚姻狀況,學歷,學校名稱,科系名稱,畢業狀況,兵役 ,MaritalStatus,DegreeID,School,Department,GradID,Military
        Dim sortItem1 As String = "班別名稱,准考證號碼,報名日期,姓名,身分證號碼,出生日期,性別,畢業狀況,郵遞區號,通訊地址,聯絡電話(日),聯絡電話(夜),行動電話,E_MAIL,參訓身分別,報名管道,筆試成績,口試成績,總成績,總成績名次"
        Dim sortItem2 As String = "ClassCName,ExamNo,RelEnterDate,Name,IDNO_MK,BIRTHDAY_MK,Sex,GradID,ZipCode6W,Address,Phone1,Phone2,CellPhone,Email,IdentityName,EnterChannel,WriteResult,OralResult,TotalResult,RSort"
        Dim sortItem1SP As String() = sortItem1.Split(",")
        Dim sortItem2SP As String() = sortItem2.Split(",")
        sort.Items.Clear()
        For i As Integer = 0 To sortItem1SP.Length - 1
            sort.Items.Add(New ListItem(sortItem1SP(i), sortItem2SP(i)))
        Next

        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID
        If sm.UserInfo.LID <> "2" Then
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        Else
            TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
        End If
    End Sub

    '匯出
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles button1.Click
        Call creattable()
    End Sub

    '匯出(設計)
    Sub creattable()
        'Dim date_str As String = ""
        'date_str = Common.FormatDate(Now().Date.ToString).Substring(0, 4) & Common.FormatDate(Now().Date.ToString).Substring(5, 2) & Common.FormatDate(Now().Date.ToString).Substring(8, 2)

        Dim v_rblEnterPathW As String = TIMS.GetListValue(rblEnterPathW)

        Const cst_免試正取 As String = "免試正取"

        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " SELECT s1.SETID" & vbCrLf
        sql &= " ,C1.ClassCName" & vbCrLf
        sql &= " ,s1.ExamNo" & vbCrLf
        sql &= " ,s1.RelEnterDate" & vbCrLf
        sql &= " ,S2.Name" & vbCrLf
        sql &= " ,S2.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(S2.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,format(S2.Birthday,'yyyy/MM/dd') Birthday" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK2(S2.Birthday) BIRTHDAY_MK" & vbCrLf
        sql &= " ,case S2.Sex when 'M' then '男' when 'F' then '女' end Sex" & vbCrLf
        sql &= " ,S2.DegreeID" & vbCrLf
        sql &= " ,S2.School" & vbCrLf
        sql &= " ,S2.Department" & vbCrLf
        sql &= " ,case S2.MaritalStatus when 1 then '已婚' when 2 then '未婚' end MaritalStatus" & vbCrLf
        sql &= " ,case S2.GradID when '01' then '畢業' when '02' then '肄業' when '03' then '在學中' end GradID" & vbCrLf
        sql &= " ,case S2.MilitaryID when '01' then '已役' when '02' then '未役' when '03' then '免役' when '04' then '在役中' end Military" & vbCrLf
        sql &= " ,concat(I1.ZipName, S2.Address) Address" & vbCrLf
        sql &= " ,S2.Phone1" & vbCrLf
        sql &= " ,S2.Phone2" & vbCrLf
        sql &= " ,S2.CellPhone" & vbCrLf
        sql &= " ,S2.Email" & vbCrLf
        sql &= " ,s1.IdentityID" & vbCrLf
        sql &= " ,dbo.FN_GET_IDENTNAME(s1.IdentityID) IdentityName" & vbCrLf

        sql &= " ,S2.ZipCode" & vbCrLf
        sql &= " ,dbo.FN_GET_ZIPCODE(S2.ZipCode,S2.ZipCode6W) ZipCode6W" & vbCrLf
        'EnterChannel_N
        sql &= " ,case s1.EnterChannel when '1' then '網路報名' when '2' then '現場報名' when '3' then '通訊報名' when '4' then '推介報名' end EnterChannel" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(s1.EnterPath,' ') !='W' THEN CONVERT(VARCHAR, s1.WriteResult) ELSE '免試正取' END WriteResult" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(s1.EnterPath,' ') !='W' THEN CONVERT(VARCHAR, s1.OralResult) ELSE '免試正取' END OralResult" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(s1.EnterPath,' ') !='W' THEN CONVERT(VARCHAR, s1.TotalResult) ELSE '免試正取' END TotalResult" & vbCrLf
        sql &= " ,K1.NAME DegreeID" & vbCrLf
        sql &= " ,s1.EnterPath" & vbCrLf
        sql &= " ,CASE WHEN ISNULL(s1.EnterPath,' ') !='W' THEN '0' ELSE '免試正取' END RSort" & vbCrLf
        'sql += " ,0 as RSort" & vbCrLf
        'sql += " --SELECT s1.OCID1" & vbCrLf
        sql &= " FROM STUD_ENTERTYPE s1" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP S2 ON s1.SETID = S2.SETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO C1 ON C1.OCID = s1.OCID1" & vbCrLf
        sql &= " JOIN ID_PLAN ip on ip.PLANID=C1.PLANID" & vbCrLf
        sql &= " JOIN VIEW_ZIPNAME I1 ON I1.ZipCode = S2.ZipCode" & vbCrLf
        sql &= " JOIN KEY_DEGREE K1 ON K1.DegreeID = S2.DegreeID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf

#Region "(No Use)"
        'sql += " Where 1=1" & vbCrLf
        'sql += " --and  CONVERT(varchar, s1.TotalResult) !='0'" & vbCrLf
        'sql += " --AND s1.EnterPath='W'" & vbCrLf
        'sql += " --AND ROWNUM <=10" & vbCrLf
        'sql += " AND s1.OCID1='57498'" & vbCrLf
#End Region

        sql &= " AND s1.OCID1 = @OCID1 " & vbCrLf
        sql &= " AND S1.CCLID IS NULL" & vbCrLf

        Select Case v_rblEnterPathW 'rblEnterPathW.SelectedValue
            Case "Y" '是 就服單位協助報名
                sql &= " AND ISNULL(s1.EnterPath,' ') = '" & TIMS.cst_EnterPathW & "' " & vbCrLf
            Case "N" '不是 就服單位協助報名
                sql &= " AND ISNULL(s1.EnterPath,' ') != '" & TIMS.cst_EnterPathW & "' " & vbCrLf
        End Select
        sql &= " ORDER BY s1.ExamNo " & vbCrLf
        Call TIMS.OpenDbConn(objconn)
        Dim dt As DataTable
        Dim oCmd As New SqlCommand(sql, objconn)
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("OCID1", SqlDbType.VarChar).Value = OCIDValue1.Value
            'Table.Load(.ExecuteReader())
            dt = DbAccess.GetDataTable(oCmd.CommandText, objconn, oCmd.Parameters)
        End With

        '總成績名次
        If dt.Rows.Count > 0 Then
            '總成績名次
            dt.DefaultView.Sort = "TotalResult desc"
            Dim tmpDt As DataTable = TIMS.dv2dt(dt.DefaultView) '總成績名次
            Dim tdr As DataRow
            '總成績名次
            For ix As Integer = 0 To tmpDt.Rows.Count - 1
                tdr = tmpDt.Rows(ix)
                Dim drow As DataRow = dt.Select("SETID='" & Convert.ToString(tdr("SETID")) & "'")(0)
                Select Case Convert.ToString(drow("RSort"))
                    Case cst_免試正取
                    Case Else
                        drow("RSort") = ix + 1
                End Select
            Next
            tmpDt.Dispose()
            dt.AcceptChanges()
        End If
        'For Each dr As DataRow In dv.Table.Rows
        '    dr("SETID")
        'Next
        ExportX1(dt)
    End Sub

    Sub ExportX1(ByRef dt As DataTable)
        If dt Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim sFileName1 As String = "學員報名甄試" & TIMS.GetDateNo2()
        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= "&ACT=" & sFileName1 & vbCrLf
        'sMemo &= String.Format("&parms=({0})", TIMS.GET_PARMSVAL(parms)) & vbCrLf
        sMemo &= "&OCIDValue1=" & OCIDValue1.Value & vbCrLf
        sMemo &= "&COUNT=" & dt.Rows.Count & vbCrLf
        TIMS.SubInsAccountLog1(Me, TIMS.Get_MRqID(Me), TIMS.cst_wm匯出, TIMS.cst_wmdip1, OCIDValue1.Value, sMemo)

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        'strHTML &= ("<div>")
        'strHTML &= ("<table>")
        strHTML &= ("<table cellspacing=""1"" cellpadding=""1"" border=""1"">")

        '建立輸出文字
        Dim ExportStr As String = ""
        Dim v_sort As String = TIMS.GetCblValue(sort) 'sort.SelectedValue

        If v_sort = "" Then
            For i_sortC As Integer = 0 To sort.Items.Count - 1
                Dim TXT_CSort As String = sort.Items.Item(i_sortC).Text
                ExportStr &= String.Concat(If(ExportStr <> "", ",", ""), TXT_CSort)
            Next
        Else
            For i_sortC As Integer = 0 To sort.Items.Count - 1
                If sort.Items.Item(i_sortC).Selected Then
                    Dim TXT_CSort As String = sort.Items.Item(i_sortC).Text
                    ExportStr &= String.Concat(If(ExportStr <> "", ",", ""), TXT_CSort)
                End If
            Next
        End If
        strHTML &= TIMS.Get_TABLETR(ExportStr)

        'ExportStr += vbCrLf
        'Common.RespWrite(Me, TIMS.sUtl_AntiXss(ExportStr))
        '建立資料面
        For Each dr As DataRow In dt.Rows
            ExportStr = ""
            If v_sort = "" Then
                For i_sortC As Integer = 0 To sort.Items.Count - 1
                    Dim TXT_CSort As String = TIMS.ClearSQM(dr(sort.Items.Item(i_sortC).Value))
                    ExportStr &= String.Concat(If(ExportStr <> "", TIMS.cst_SplitB1, ""), TXT_CSort)
                Next
            Else
                For i_sortC As Integer = 0 To Me.sort.Items.Count - 1
                    If Me.sort.Items(i_sortC).Selected Then
                        Dim TXT_CSort As String = TIMS.ClearSQM(dr(sort.Items.Item(i_sortC).Value))
                        ExportStr &= String.Concat(If(ExportStr <> "", TIMS.cst_SplitB1, ""), TXT_CSort)
                    End If
                Next
            End If
            strHTML &= TIMS.Get_TABLETR(ExportStr, True, TIMS.cst_SplitB1)
        Next
        strHTML &= ("</table>")
        'strHTML &= ("</div>")

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") ' Response.End()
    End Sub

    '參訓身分別
    'Function GET_IdentityName(ByVal drIdentityID As String) As String
    '    Dim rst As String = ""
    '    '參訓身分別
    '    'Dim IdentityName As String = ""
    '    'Dim IdentityID As String() 
    '    'Dim StrSql As String = ""
    '    'Dim drow As DataRow 
    '    'IdentityName = ""
    '    If drIdentityID <> "" Then
    '        Dim IdentityID As String() = Split(drIdentityID, ",")
    '        For i As Integer = 0 To IdentityID.Length - 1
    '            If IdentityID(i) <> "" Then
    '                Dim StrSql As String = ""
    '                StrSql = " SELECT Name From Key_Identity WHERE IdentityID = '" & IdentityID(i) & "' "
    '                Dim drow As DataRow = DbAccess.GetOneRow(StrSql, objconn)
    '                If Not drow Is Nothing Then
    '                    If rst <> "" Then rst &= "/"
    '                    rst &= drow("Name")
    '                End If
    '            End If
    '        Next
    '    End If
    '    If rst = "" Then rst = "未填寫"
    '    Return rst
    'End Function

    '判斷機構是否只有一個班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '判斷機構是否只有一個班級
        Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
    End Sub
End Class