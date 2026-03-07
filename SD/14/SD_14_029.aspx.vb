Partial Class SD_14_029
    Inherits AuthBasePage

    Const cst_printFN1 As String = "SD_14_029" '2011

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

        If Not IsPostBack Then
            labmsg.Text = "" '清空
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID
            '判斷機構是否只有一個班級3
            Call TIMS.GET_OnlyOne_OCID3(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn)
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

    End Sub

    ''' <summary>檢核錄取人數</summary>
    ''' <param name="v_OCID"></param>
    ''' <returns></returns>
    Function CHK_STUD_ADMISSION_Y(ByRef v_OCID As String) As Integer
        Dim rst As Integer = 0
        If v_OCID = "" Then Return rst

        Dim parms As New Hashtable From {{"OCID", v_OCID}}
        Dim sql As String = ""
        sql &= " SELECT ROW_NUMBER() OVER(ORDER BY d.SETID) AS ROWSETID" & vbCrLf
        sql &= " ,d.IDNO" & vbCrLf
        'sql &= " ,concat(d.NAME,CASE cs.StudStatus WHEN '2' THEN N'(離訓)' WHEN '3' THEN N'(退訓)' ELSE '' END) STDNAME" & vbCrLf
        sql &= " ,d.SETID ,cs.SOCID" & vbCrLf
        sql &= " FROM STUD_SELRESULT b" & vbCrLf
        sql &= " JOIN STUD_ENTERTYPE a ON a.SETID=b.SETID AND a.EnterDate=b.EnterDate AND a.SerNum=b.SerNum AND a.OCID1=b.OCID" & vbCrLf
        sql &= " JOIN STUD_ENTERTEMP d ON d.SETID=a.SETID" & vbCrLf
        sql &= " JOIN CLASS_CLASSINFO e ON e.OCID=b.OCID" & vbCrLf
        sql &= " LEFT JOIN STUD_STUDENTINFO ss on ss.IDNO=d.IDNO" & vbCrLf
        sql &= " LEFT JOIN CLASS_STUDENTSOFCLASS cs on cs.ocid=a.ocid1 AND cs.SID=ss.SID" & vbCrLf
        sql &= " WHERE b.OCID=@OCID" '#{OCID}" & vbCrLf
        sql &= " and b.ADMISSION='Y'" & vbCrLf ' /* 140288 是否錄取:錄取 (除了正取才是錄取) N 不錄取 */" & vbCrLf
        sql &= " and b.SelResultID='01'" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return rst

        rst = dt.Rows.Count
        Return rst
    End Function

    Function checkData2(ByRef errMsg As String) As Boolean
        Dim rst As Boolean = True 'False

        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim iStd As Integer = CHK_STUD_ADMISSION_Y(OCIDValue1.Value)
        If iStd = 0 Then errMsg &= "查無班級錄取人數資料!" & vbCrLf
        rst = If(errMsg <> "", False, True)
        If errMsg <> "" Then Return rst

        Return rst
    End Function

    Public Shared Function CHK_STUD_ADP_STUDATTEND(oConn As SqlConnection, v_OCID As String) As Integer
        Dim rst As Integer = 0
        If v_OCID = "" Then Return rst

        Dim parms As New Hashtable From {{"OCID", Val(v_OCID)}}
        Dim sql As String = " SELECT 1 CROW FROM ADP_STUDATTEND a WHERE a.OCID=@OCID" '#{OCID}" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, oConn, parms)
        If dt Is Nothing OrElse dt.Rows.Count = 0 Then Return rst

        rst = dt.Rows.Count
        Return rst
    End Function

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim errMsg As String = ""
        labmsg.Text = ""

        'Years.Value = TIMS.ClearSQM(Years.Value)
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        Dim drCC As DataRow = TIMS.GetOCIDDate(OCIDValue1.Value, objconn)
        If OCIDValue1.Value = "" Then
            errMsg &= "請選擇 職類/班別" & vbCrLf
            'Common.MessageBox(Me, errMsg)
            'Return '  Exit Sub
        ElseIf drCC Is Nothing Then
            errMsg &= "職類/班別 選擇有誤!" & vbCrLf
        End If
        If errMsg <> "" Then
            labmsg.Text = errMsg
            Common.MessageBox(Me, errMsg)
            Return
        End If

        Dim sErrMsg As String = ""
        Call checkData2(sErrMsg)
        If sErrMsg <> "" Then
            labmsg.Text = sErrMsg
            Common.MessageBox(Me, sErrMsg)
            Exit Sub
        End If

        Dim rPMS As New Hashtable From {{"HP", 2}, {"OCID", OCIDValue1.Value}}

        Call TIMS.UPDATE_ADP_STUDATTEND(sm, objconn, OCIDValue1.Value, rPMS)

        Dim sErrMsg3 As String = ""
        Dim iStd As Integer = CHK_STUD_ADP_STUDATTEND(objconn, OCIDValue1.Value)
        If iStd = 0 Then sErrMsg3 &= "查無 學員線上簽到(退)明細 資料!" & vbCrLf
        If sErrMsg3 <> "" Then
            labmsg.Text = sErrMsg3
            Common.MessageBox(Me, sErrMsg3)
            Exit Sub
        End If

        Dim s_MSD As String = Convert.ToString(drCC("MSD"))
        Dim MyValue As String = ""
        TIMS.SetMyValue(MyValue, "TPlanID", sm.UserInfo.TPlanID)
        TIMS.SetMyValue(MyValue, "MSD", s_MSD)
        TIMS.SetMyValue(MyValue, "OCID", OCIDValue1.Value)
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
    End Sub


    '判斷機構是否只有一個班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn) '判斷機構是否只有一個班級
    End Sub

End Class