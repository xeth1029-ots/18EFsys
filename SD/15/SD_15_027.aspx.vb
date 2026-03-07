Public Class SD_15_027
    Inherits AuthBasePage

    't:(勞保勾稽)產投／充飛 STUD_BLIGATEDATA28 SOCID
    't:(勞退)產投／充飛／自辦在職 STUD_BLIPERSON28 SOCID
    't:(勞保勾稽)自辦在職：STUD_SELRESULTBLI OCID IDNO
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objconn)
        TIMS.Get_TitleLab(objconn, Request("ID"), TitleLab1, TitleLab2)
        '檢查Session是否存在 End
        PageControler1.PageDataGrid = DataGrid1

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, Historytable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1")
        If Historytable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        If Not IsPostBack Then
            Call sCreate1()
        End If
    End Sub

    Sub sCreate1()
        BtnSch1.Attributes("onclick") = "return CheckSearch();"
        BtnClearOCIDValue.Attributes("onclick") = "return ClearData();"

        Dim js_BtnLevOrg1_1 As String = "javascript:openOrg('../../Common/LevOrg1.aspx');"
        If sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1 Then js_BtnLevOrg1_1 = "javascript:openOrg('../../Common/LevOrg.aspx');"
        BtnLevOrg1.Attributes("onclick") = js_BtnLevOrg1_1

        'Table4.Visible = False
        PageControler1.Visible = False
        lab_msg_1.Text = ""
        center.Text = sm.UserInfo.OrgName
        RIDValue.Value = sm.UserInfo.RID

        '若只有管理一個班級，自動協助帶出班級--by andy 2009-02-25
        TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
    End Sub

    ''' <summary>
    ''' 查詢／匯出 必要檢核
    ''' </summary>
    ''' <param name="errMsg"></param>
    ''' <returns></returns>
    Function checkData1(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True 'False

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'If OCIDValue1.Value = "" Then Exit Sub
        s_IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(s_IDNO.Text))
        s_NAME.Text = TIMS.ClearSQM(s_NAME.Text)

        Select Case sm.UserInfo.TPlanID
            Case TIMS.Cst_TPlanID70
            Case TIMS.Cst_TPlanID06
            Case TIMS.Cst_TPlanID28
            Case Else
                'Common.MessageBox(Me, "該計畫不提供此功能!!") 'Exit Sub
                errMsg &= String.Concat(TIMS.cst_ErrorMsg17, vbCrLf) '"該計畫不提供此功能!!" & vbCrLf
                Return False
        End Select

        '若角色無權限，顯示：「 該功能，無權限使用!」
        '只有署可以查， 分署、單位均無權限
        '區域產業據點計畫 增加 【勞保勞退資料查詢】功能
        '階層代碼 0:署 1:中心 2:委訓 
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                errMsg &= "該功能，無權限使用!" & vbCrLf
                Return False
        End Select

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'If OCIDValue1.Value = "" Then
        '    errMsg &= "請選擇職類班別!!" & vbCrLf
        '    Return False
        'End If
        '請選擇職類班別 或輸入身分證號碼!
        If OCIDValue1.Value = "" AndAlso s_IDNO.Text = "" Then
            errMsg &= "請選擇職類班別 或輸入身分證號碼!" & vbCrLf
            Return False
        End If

        If errMsg <> "" Then rst = False
        Return rst
    End Function

    ''' <summary>查詢</summary>
    Sub sSearch1()
        Call TIMS.SUtl_TxtPageSize(Me, TxtPageSize, DataGrid1)

        Dim sql As String = ""
        Dim parms As New Hashtable
        sql = get_sch1_sql(parms)
        If sql = "" Then Exit Sub
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        tb_Sch_DG1.Visible = False
        lab_msg_1.Text = TIMS.cst_NODATAMsg1 '"查無資料"
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= "&ACT=勞保勞退資料查詢" & vbCrLf
        'If sFileName1 <> "" Then sMemo &= String.Format("&sFileName1=({0})", sFileName1) & vbCrLf
        If center.Text <> "" Then sMemo &= String.Format("&center=({0})", center.Text) & vbCrLf
        If RIDValue.Value <> "" Then sMemo &= String.Format("&RIDValue=({0})", RIDValue.Value) & vbCrLf
        If OCID1.Text <> "" Then sMemo &= String.Format("&OCID1=({0})", OCID1.Text) & vbCrLf
        If OCIDValue1.Value <> "" Then sMemo &= String.Format("&OCIDValue1=({0})", OCIDValue1.Value) & vbCrLf
        If s_IDNO.Text <> "" Then sMemo &= String.Format("&IDNO=({0})", s_IDNO.Text) & vbCrLf
        If s_NAME.Text <> "" Then sMemo &= String.Format("&NAME=({0})", s_NAME.Text) & vbCrLf
        sMemo &= String.Format("&parms=({0})", TIMS.GetMyValue3(parms)) & vbCrLf
        'sMemo &= "&sql=" & sql & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm查詢, TIMS.cst_wmdip1, OCIDValue1.Value, sMemo)

        tb_Sch_DG1.Visible = True
        lab_msg_1.Text = ""
        PageControler1.PageDataTable = dt
        PageControler1.ControlerLoad()
    End Sub

    ''' <summary>匯出</summary>
    Sub Export1()
        Dim sql As String = ""
        Dim parms As New Hashtable
        sql = get_sch1_sql(parms)
        If sql = "" Then Exit Sub
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)

        tb_Sch_DG1.Visible = False
        lab_msg_1.Text = TIMS.cst_NODATAMsg1 '"查無資料"
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        Const cst_TitleS1 As String = "勞保勞退資料"
        Dim sFileName1 As String = cst_TitleS1 & TIMS.GetToday(objconn)

        Dim sMemo As String = ""
        sMemo = ""
        sMemo &= "&ACT=勞保勞退資料查詢" & vbCrLf
        If sFileName1 <> "" Then sMemo &= String.Format("&sFileName1=({0})", sFileName1) & vbCrLf
        If center.Text <> "" Then sMemo &= String.Format("&center=({0})", center.Text) & vbCrLf
        If RIDValue.Value <> "" Then sMemo &= String.Format("&RIDValue=({0})", RIDValue.Value) & vbCrLf
        If OCID1.Text <> "" Then sMemo &= String.Format("&OCID1=({0})", OCID1.Text) & vbCrLf
        If OCIDValue1.Value <> "" Then sMemo &= String.Format("&OCIDValue1=({0})", OCIDValue1.Value) & vbCrLf
        If s_IDNO.Text <> "" Then sMemo &= String.Format("&IDNO=({0})", s_IDNO.Text) & vbCrLf
        If s_NAME.Text <> "" Then sMemo &= String.Format("&NAME=({0})", s_NAME.Text) & vbCrLf
        sMemo &= String.Format("&parms=({0})", TIMS.GetMyValue3(parms)) & vbCrLf
        'sMemo &= "&sql=" & sql & vbCrLf
        TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm匯出, TIMS.cst_wmdip1, OCIDValue1.Value, sMemo)

        Dim sTitle1 As String = String.Concat((sm.UserInfo.Years - 1911), "年度－勞保勞退資料表")

        Dim sPattern As String = ""
        Dim sColumn As String = ""
        sPattern = "序號,轄區分署,年度,訓練計畫,姓名,身分證號,訓練機構,班別名稱,勞保薪資級距,勞退月提繳級距"
        sColumn = "SEQNUM,DISTNAME,YEARS,PLANNAME,SNAME,IDNO_MK,ORGNAME,CLASSCNAME2,SALARY,WAGE"

        Dim sPatternA() As String = Split(sPattern, ",")
        Dim sColumnA() As String = Split(sColumn, ",")
        Dim iColSpanCount As Integer = sColumnA.Length

        '套CSS值
        Dim strSTYLE As String = ""
        strSTYLE &= "<style>"
        strSTYLE &= "td{mso-number-format:""\@"";}"
        strSTYLE &= ".noDecFormat{mso-number-format:""0"";}"
        strSTYLE &= "</style>"

        Dim strHTML As String = ""
        strHTML &= "<div>"
        strHTML &= "<table cellspacing=""1"" cellpadding=""1"" width=""100%"" border=""1"">"
        'Common.RespWrite(Me, "<tr>")

        '標題抬頭
        Dim ExportStr As String = "" '建立輸出文字
        ExportStr = "<tr>"
        ExportStr &= "<td colspan='" & iColSpanCount & "' align='center'>" & sTitle1 & "</td>" '& vbTab
        ExportStr &= "</tr>" & vbCrLf
        ExportStr &= "<tr>"
        For i As Integer = 0 To sPatternA.Length - 1
            ExportStr &= "<td>" & sPatternA(i) & "</td>" '& vbTab
        Next
        ExportStr &= "</tr>" & vbCrLf
        strHTML &= ExportStr

        '建立資料面
        Dim iNum As Integer = 0
        For Each dr As DataRow In dt.DefaultView.Table.Rows
            iNum += 1
            ExportStr = "<tr>"
            For i As Integer = 0 To sColumnA.Length - 1
                ExportStr &= "<td>" & Convert.ToString(dr(sColumnA(i))) & "</td>" '& vbTab
            Next
            ExportStr &= "</tr>" & vbCrLf
            strHTML &= ExportStr
            'Common.RespWrite(Me, ExportStr)
        Next
        strHTML &= "</table>"
        strHTML &= "</div>"

        Dim parmsExp As New Hashtable
        parmsExp.Add("ExpType", TIMS.GetListValue(RBListExpType))
        parmsExp.Add("FileName", sFileName1)
        parmsExp.Add("strSTYLE", strSTYLE)
        parmsExp.Add("strHTML", strHTML)
        parmsExp.Add("ResponseNoEnd", "Y")
        TIMS.Utl_ExportRp1(Me, parmsExp)
        TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
    End Sub

    ''' <summary>
    ''' sql ／參數組合資訊
    ''' </summary>
    ''' <param name="parms"></param>
    ''' <returns></returns>
    Function get_sch1_sql(ByRef parms As Hashtable) As String
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        'If OCIDValue1.Value = "" Then Exit Sub
        s_IDNO.Text = TIMS.ChangeIDNO(TIMS.ClearSQM(s_IDNO.Text))
        s_NAME.Text = TIMS.ClearSQM(s_NAME.Text)

        'Dim parms As New Hashtable
        parms.Clear()
        If OCIDValue1.Value <> "" Then parms.Add("OCID", OCIDValue1.Value)
        If s_IDNO.Text <> "" Then parms.Add("IDNO", s_IDNO.Text)
        If s_NAME.Text <> "" Then parms.Add("NAMElk", s_NAME.Text)
        If OCIDValue1.Value = "" AndAlso s_IDNO.Text = "" Then Return "" '至少要有某資訊

        '「轄區分署」、「年度」、「訓練計畫」、「姓名」、「身分證號碼」、「訓練機構」、「班別名稱」、「勞保薪資級距」、「勞退月提繳級距」。
        Dim sql As String = ""
        sql = "" & vbCrLf
        'sql &= " SELECT TOP 100 s.OCID" & vbCrLf
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY S.SOCID) SEQNUM" & vbCrLf
        sql &= " ,s.RID,s.OCID,s.DISTID,s.DISTNAME" & vbCrLf
        sql &= " ,s.TPLANID ,s.PLANNAME" & vbCrLf
        sql &= " ,s.YEARS ,S.ORGNAME" & vbCrLf
        sql &= " ,S.SOCID,S.NAME SNAME" & vbCrLf
        sql &= " ,S.IDNO" & vbCrLf
        sql &= " ,dbo.FN_GET_MASK1(S.IDNO) IDNO_MK" & vbCrLf
        sql &= " ,dbo.FN_GET_CLASSCNAME(S.CLASSCNAME,S.CYCLTYPE) CLASSCNAME2" & vbCrLf
        sql &= " ,ISNULL(b2.SALARY,b4.SALARY) SALARY" & vbCrLf '勞保薪資級距
        sql &= " ,b3.WAGE" & vbCrLf '勞退月提繳級距
        sql &= " FROM dbo.V_STUDENTINFO s" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_BLIGATEDATA28 b2 on b2.socid=s.socid AND b2.idno=s.idno" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_BLIPERSON28 b3 on b3.socid=s.socid AND b3.idno=s.idno" & vbCrLf
        sql &= " LEFT JOIN dbo.STUD_SELRESULTBLI b4 on b4.IDNO=s.IDNO AND b4.OCID=s.OCID" & vbCrLf
        sql &= " WHERE 1=1" & vbCrLf
        sql &= " AND S.TPLANID=@TPLANID" & vbCrLf
        sql &= " AND s.YEARS=@YEARS" & vbCrLf
        parms.Add("TPLANID", sm.UserInfo.TPlanID)
        parms.Add("YEARS", sm.UserInfo.Years)
        Select Case sm.UserInfo.LID
            Case 1
                parms.Add("PLANID", sm.UserInfo.PlanID)
                sql &= " AND S.PLANID=@PLANID" & vbCrLf
            Case 2
                parms.Add("RID", sm.UserInfo.RID)
                sql &= " AND s.RID=@RID" & vbCrLf
        End Select
        'sql &= " and s.OCID =122549" & vbCrLf
        If OCIDValue1.Value <> "" Then sql &= " AND s.OCID =@OCID " & vbCrLf
        If s_IDNO.Text <> "" Then sql &= " AND s.IDNO =@IDNO " & vbCrLf
        If s_NAME.Text <> "" Then sql &= " AND s.NAME like '%'+@NAMElk+'%' " & vbCrLf
        sql &= " ORDER BY S.SOCID" & vbCrLf
        Return sql
    End Function


    Private Sub BtnGETvalue2_Click(sender As Object, e As System.EventArgs) Handles BtnGETvalue2.Click
        'BtnGETvalue2
        '判斷機構是否只有一個班級
        Dim dr As DataRow
        dr = TIMS.GET_OnlyOne_OCID(Me, RIDValue.Value, objconn)
        '不只一個班級
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse Convert.ToString(dr("total")) <> "1" Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub

    ''' <summary>
    ''' 清理異常的勾稽資料1  b.idno!=ss.idno
    ''' </summary>
    Sub ClearNGIDNO1()
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql &= " select cs.OCID,cs.SOCID" & vbCrLf
        sql &= " ,ss.name,ss.idno" & vbCrLf
        sql &= " ,b.name bname,b.idno bidno" & vbCrLf
        sql &= " ,format(b.MODIFYDATE,'yyyy/MM/dd HH:mm:ss') bMODIFYDATE" & vbCrLf
        sql &= " from CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " join STUD_STUDENTINFO ss on ss.sid=cs.sid" & vbCrLf
        sql &= " join STUD_BLIGATEDATA28 b on b.socid =cs.socid and b.idno!=ss.idno" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and b.MODIFYDATE>=getdate()-400" & vbCrLf
        Dim dt1 As DataTable = DbAccess.GetDataTable(sql, objconn)
        If dt1.Rows.Count = 0 Then Return

        Dim s_logmsg1 As String = ""
        s_logmsg1 &= String.Format("##SD_15_027 Sub ClearNGIDNO1: 清理異常的勾稽資料1： {0} 筆", dt1.Rows.Count) & vbCrLf
        For Each dr1 As DataRow In dt1.Rows
            s_logmsg1 &= String.Format("OCID：{0},SOCID：{1},idno：{2},name：{3},bidno：{4},bname：{5},bMODIFYDATE：{6}", dr1("OCID"), dr1("SOCID"), dr1("idno"), dr1("name"), dr1("bidno"), dr1("bname"), dr1("bMODIFYDATE")) & vbCrLf
        Next
        TIMS.LOG.Warn(s_logmsg1)

        sql = "" & vbCrLf
        sql &= " DELETE STUD_BLIGATEDATA28" & vbCrLf
        sql &= " from CLASS_STUDENTSOFCLASS cs" & vbCrLf
        sql &= " join STUD_STUDENTINFO ss on ss.sid=cs.sid" & vbCrLf
        sql &= " join STUD_BLIGATEDATA28 b on b.socid =cs.socid and b.idno!=ss.idno" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        sql &= " and b.MODIFYDATE>=getdate()-400" & vbCrLf
        DbAccess.ExecuteNonQuery(sql, objconn)
    End Sub

    '查詢
    Protected Sub BtnSch1_Click(sender As Object, e As EventArgs) Handles BtnSch1.Click
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                Common.MessageBox(Me, TIMS.cst_ErrorMsg5s)
                Exit Sub
        End Select
        Dim sErrMsg As String = ""
        Call checkData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        '清理異常的勾稽資料1
        Call ClearNGIDNO1()

        Call sSearch1()
    End Sub

    '匯出
    Protected Sub BtnExp1_Click(sender As Object, e As EventArgs) Handles BtnExp1.Click
        Select Case sm.UserInfo.LID
            Case 0
            Case Else
                Common.MessageBox(Me, TIMS.cst_ErrorMsg5s)
                Exit Sub
        End Select
        Dim sErrMsg As String = ""
        Call checkData1(sErrMsg)
        If sErrMsg <> "" Then
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        '清理異常的勾稽資料1
        Call ClearNGIDNO1()
        '匯出
        Call Export1()
    End Sub
End Class
