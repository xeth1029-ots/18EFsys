Public Class SD_07_002_R
    Inherits AuthBasePage

    Const cst_printFN1 As String = "skill_list"
    'Const cst_printFN2 As String = "SD_07_002_R2"
    'Const cst_printFN3 As String = "SD_07_002_R3"
    'Const cst_printFN4 As String = "SD_07_002_R4"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        If Not IsPostBack Then
            'DataGridTable.Visible = False
            center.Text = sm.UserInfo.OrgName
            RIDValue.Value = sm.UserInfo.RID

            If sm.UserInfo.LID <> "2" Then
                TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
            Else
                Button13_Click(sender, e)
            End If
        End If

        Const cst_javascript_openOrg_FMT1 As String = "javascript:openOrg('../../Common/LevOrg{0}.aspx');SetOneOCID();"
        Button8.Attributes("onclick") = String.Format(cst_javascript_openOrg_FMT1, If(sm.UserInfo.RID = "A" OrElse sm.UserInfo.RoleID <= 1, "", "1"))

        TIMS.ShowHistoryRID(Me, HistoryRID, "HistoryList2", "RIDValue", "center")
        If HistoryRID.Rows.Count <> 0 Then
            center.Attributes("onclick") = "showObj('HistoryList2');"
            center.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "RIDValue", "center", "TMIDValue1", "TMID1", True, "btnSchExamTime")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If

        '列印
        Button1.Attributes("onclick") = "javascript:return ReportPrint();"

    End Sub

    '列印
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        '測試班級是否有資料
        Dim dt As DataTable
        Dim sql As String = ""
        sql = ""
        sql &= " select a.ExamTime,a.ExamName "
        sql &= " FROM V_STUDTECHEXAM a" '請直接改view 因為改變技能檢定的基礎設定
        sql &= " where 1=1"
        If OCIDValue1.Value <> "" Then
            sql &= " and a.ocid ='" & OCIDValue1.Value & "'"
        End If
        If ddlKindTime.SelectedValue <> "" Then
            sql &= " and a.ExamTime ='" & ddlKindTime.SelectedValue & "'"
        End If
        sql &= " order by a.ExamTime"
        dt = DbAccess.GetDataTable(sql, objconn)
        Button1.Enabled = True
        If dt.Rows.Count = 0 Then
            Button1.Enabled = False
            Common.MessageBox(Me, "查無資料，無法列印，請重新查詢!!")
            Exit Sub
        End If

        '列印
        Dim MyValue As String = ""
        MyValue = ""
        MyValue &= "&OCID=" & OCIDValue1.Value
        MyValue &= "&TMID=" & TMIDValue1.Value
        MyValue &= "&KindTime=" & ddlKindTime.SelectedValue
        'MyValue &= "&KindTime=" & ddlKindTime.SelectedValue
        Call TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)

        'Select Case ddlKindTime.SelectedValue
        '    Case "1"
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN1, MyValue)
        '    Case "2"
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN2, MyValue)
        '    Case "3"
        '        TIMS.CloseDbConn(objconn) : ReportQuery.PrintReport(Me, cst_printFN3, MyValue)
        'End Select
    End Sub

    '查詢技能檢定職類。
    Private Sub btnSchExamTime_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSchExamTime.Click
        ddlKindTime.Items.Clear()

        Dim dt As DataTable
        Dim sql As String = ""
        sql = ""
        sql &= " select a.ExamTime,a.ExamName "
        sql &= " FROM V_STUDTECHEXAM a" '請直接改view 因為改變技能檢定的基礎設定
        sql &= " where 1=1"
        If OCIDValue1.Value <> "" Then
            sql &= " and a.ocid ='" & OCIDValue1.Value & "'"
        End If
        sql &= " order by a.ExamTime"
        dt = DbAccess.GetDataTable(sql, objconn)
        Button1.Enabled = True
        If dt.Rows.Count = 0 Then
            Button1.Enabled = False
            Common.MessageBox(Me, "查無資料，無法列印，請重新查詢!!")
            Exit Sub
        End If

        With ddlKindTime
            .DataSource = dt
            .DataTextField = "ExamName"
            .DataValueField = "ExamTime"
            .DataBind()
        End With
    End Sub

#Region "last function search tmid ocid"
    Private Sub btnSetOneOCID_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnSetOneOCID.Click
        If sm.UserInfo.LID <> "2" Then
            '20090220 andy edit 當登入帳號為 一般使用者、承辦人 該帳號有賦于班級時(只有一個時)帶出該班級
            TIMS.GET_OnlyOne_OCID(Me, TMID1, TMIDValue1, OCID1, OCIDValue1)
        End If
    End Sub

    '查詢是否只有一個班
    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '判斷機構是否只有一個班級
        Dim dr As DataRow = TIMS.GET_OnlyOne_OCID(RIDValue.Value, objconn)
        ddlKindTime.Items.Clear()
        TMID1.Text = ""
        OCID1.Text = ""
        TMIDValue1.Value = ""
        OCIDValue1.Value = ""
        '如果只有一個班級
        If dr Is Nothing OrElse $"{dr("TOTAL")}" = "" OrElse TIMS.CINT1(dr("TOTAL")) <> 1 Then Return
        TMID1.Text = dr("trainname")
        OCID1.Text = dr("classname")
        TMIDValue1.Value = dr("trainid")
        OCIDValue1.Value = dr("ocid")
    End Sub
#End Region

End Class

