Partial Class SD_14_009
    Inherits AuthBasePage

    'ReportQuery
    'SQControl.aspx
    'SD_14_009_16 '(2016)新的 SD_14_009_16*.jrxml
    'SD_14_009 '(2014)
    'SD_14_009_00 '民國100年舊(2011)

    'SD_14_009_16*.jrxml
    Const cst_printFN1 As String = "SD_14_009_00" '2011
    Const cst_printFN2 As String = "SD_14_009" '2014
    Const cst_printFN3 As String = "SD_14_009_16" '2016

    '參考程式：SD_05_014_add
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

        Years.Value = sm.UserInfo.Years - 1911

        If Not IsPostBack Then
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

#Region "(No Use)"

    ''檢查是否已經審核  
    'Function ChkAppliedResultM(ByVal OCID As String, ByVal tConn As SqlConnection) As Boolean
    '    If OCID = "" Then Return False
    '    Dim Rst As Boolean = False
    '    'Dim dr As DataRow
    '    Dim sql As String = ""
    '    sql = "SELECT AppliedResultM FROM Class_ClassInfo WHERE OCID=@OCID "
    '    Dim sCmd As New SqlCommand(sql, tConn)
    '    Call TIMS.OpenDbConn(tConn)
    '    Dim dt As New DataTable
    '    With sCmd
    '        .Parameters.Clear()
    '        .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
    '        dt.Load(.ExecuteReader())
    '    End With
    '    If dt.Rows.Count > 0 Then
    '        Select Case Convert.ToString(dt.Rows(0)("AppliedResultM"))
    '            Case "Y", "N"
    '                Rst = True
    '        End Select
    '    End If
    '    Return Rst
    'End Function

#End Region

    '檢查是否已提出審核申請
    Function CheckStud_SubsidyCost(ByVal OCID As String, ByVal tConn As SqlConnection) As Boolean
        If OCID = "" Then Return False
        Dim Rst As Boolean = False
        Dim sql As String = ""
        sql = ""
        sql += " SELECT 'x' x "
        sql += " FROM Class_ClassInfo a"
        sql += " JOIN Class_StudentsOfClass b ON a.OCID = b.OCID "
        sql += " JOIN Stud_SubsidyCost c ON b.SOCID = c.SOCID "
        sql += " WHERE 0 = 0 "
        sql += "    AND a.OCID = @OCID "
        sql += "    AND c.SOCID IS NOT NULL "
        Dim sCmd As New SqlCommand(sql, tConn)
        Call TIMS.OpenDbConn(tConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then Rst = True '已提出補助申請
        Return Rst
    End Function

    '檢查該班是否已結訓
    Function CheckIsClosed(ByVal OCID As String, ByVal tConn As SqlConnection) As Boolean
        'IsClosed	是否結訓
        Dim Rst As Boolean = False '未結訓
        Dim sql As String = ""
        sql = ""
        sql += " SELECT a.IsClosed "
        sql += " FROM Class_ClassInfo a "
        sql += " WHERE 0 = 0 "
        sql += "    AND a.OCID = @OCID "
        Dim sCmd As New SqlCommand(sql, tConn)
        Call TIMS.OpenDbConn(tConn)
        Dim dt As New DataTable
        With sCmd
            .Parameters.Clear()
            .Parameters.Add("OCID", SqlDbType.VarChar).Value = OCID
            dt.Load(.ExecuteReader())
        End With
        If dt.Rows.Count > 0 Then
            If Convert.ToString(dt.Rows(0)("IsClosed")) = "Y" Then Rst = True '已結訓
        End If
        Return Rst
    End Function

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Me.labmsg.Text = "查無資料!"
        Dim tmpMsg1 As String = "查無資料!"
        'Dim OCID As String = OCIDValue1.Value
        OCIDValue1.Value = TIMS.ClearSQM(OCIDValue1.Value)
        If OCIDValue1.Value = "" Then
            Common.MessageBox(Me, "請選擇班級")
            Exit Sub
        End If

        Dim flag1 As Boolean = CheckStud_SubsidyCost(OCIDValue1.Value, objconn)
        Dim flag2 As Boolean = CheckIsClosed(OCIDValue1.Value, objconn)
        If Not flag1 Then tmpMsg1 = "班級尚未送交補助金申請!!"
        If Not flag2 Then tmpMsg1 = "班級尚未結訓!!"
        Me.labmsg.Text = tmpMsg1 '"查無資料!"

        'If flag1 AndAlso flag2 Then  'edit，by:20181029
        If True Then  '依照承辦人需求,將目前限制條件都先拿掉，by:20181029
            Me.labmsg.Text = ""
            Dim vsFileName1 As String = cst_printFN1 '"SD_14_009_00"
            If sm.UserInfo.Years >= "2012" Then vsFileName1 = cst_printFN2 '"SD_14_009"
            If sm.UserInfo.Years >= "2016" Then vsFileName1 = cst_printFN3 '"SD_14_009_16"
            Dim MyValue As String = ""
            MyValue = "Years=" & Years.Value
            MyValue &= "&OCID=" & OCIDValue1.Value
            TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, vsFileName1, MyValue)
#Region "(No Use)"

            'Dim strScript As String
            'strScript = "<script language=""javascript"">" + vbCrLf
            'strScript += "     CheckPrint('" & ReportQuery.GetSmartQueryPath & "') ;" + vbCrLf
            'strScript += "</script>"
            'Me.RegisterStartupScript("", strScript)

#End Region
        End If
    End Sub

    '判斷機構是否只有一個班級
    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        Call TIMS.GET_OnlyOne_OCID2(Me, RIDValue.Value, TMID1, OCID1, TMIDValue1, OCIDValue1, objconn) '判斷機構是否只有一個班級
    End Sub
End Class