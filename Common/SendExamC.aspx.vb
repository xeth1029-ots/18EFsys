Partial Class SendExamC
    Inherits AuthBasePage

    '檢定職類與考試級別
    '/* 檢定職類 、 職類群*/
    'KEY_EXAM3/KEY_EXAMGROUP
    'fieldnameGP.Value = TIMS.ClearSQM(Request("NMGP")) '不定(欄位)
    '  fieldnameXM.Value = TIMS.ClearSQM(Request("NMXM")) '不定(欄位)
    '  fieldnameLV.Value = TIMS.ClearSQM(Request("NMLV")) '不定(欄位)
    '  fieldvalueXM.Value = TIMS.ClearSQM(Request("VLXM")) '不定(欄位值)
    '  fieldvalueLV.Value = TIMS.ClearSQM(Request("VLLV")) '不定(欄位值)
    Const cst_ffvv As String = "NMGP,NMXM,NMLV,VLXM,VLLV,BTN1"

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        If TIMS.ChkSession(sm) Then
            Common.RespWrite(Me, "<script>alert('您的登入資訊已經遺失，請重新登入');window.close();</script>")
            TIMS.Utl_RespWriteEnd(Me, objconn, "") 'Response.End()
        End If
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        For Each sRqVal As String In cst_ffvv.Split(",")
            If TIMS.CheckInput(Request(sRqVal)) Then
                Common.MessageBox(Me, TIMS.cst_ErrorMsg2)
                Exit Sub
            End If
        Next
        fieldnameGP.Value = TIMS.ClearSQM(Request("NMGP")) '不定(欄位)
        fieldnameXM.Value = TIMS.ClearSQM(Request("NMXM")) '不定(欄位)
        fieldnameLV.Value = TIMS.ClearSQM(Request("NMLV")) '不定(欄位)
        fieldvalueXM.Value = TIMS.ClearSQM(Request("VLXM")) '不定(欄位值)
        fieldvalueLV.Value = TIMS.ClearSQM(Request("VLLV")) '不定(欄位值)
        fieldbtnN1.Value = TIMS.ClearSQM(Request("BTN1")) '不定(按鈕)

        If Not IsPostBack Then Call cCreate1()
    End Sub

    Sub cCreate1()
        'ddlEXAMGROUP
        '職類群
        'ddlEXAM3
        '檢定職類
        'ddlEXLEVEL
        '級別
        'tbCJOB2.Visible = True
        ddlEXAMGROUP.Items.Clear()
        ddlEXAM3.Items.Clear()
        ddlEXLEVEL.Items.Clear()
        ddlEXAMGROUP = GetKeyExam3(ddlEXAMGROUP, "", 1)
    End Sub

    '設計 DropDownList  
    Function GetKeyExam3(ByRef obj As ListControl, ByVal s_value As String, ByRef iType As Integer) As ListControl
        Dim parms As New Hashtable
        Dim sql As String = ""
        Select Case iType
            Case 1
                sql = " SELECT EGID,JGNAME FROM KEY_EXAMGROUP ORDER BY EGID "
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)
                With obj
                    .DataSource = dt
                    .DataTextField = "JGNAME"
                    .DataValueField = "EGID"
                    .DataBind()
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End With
            Case 2
                parms.Clear()
                parms.Add("EGID", s_value)
                sql = " SELECT EXAMID,CONCAT('(',EXAMID,')',EXNAME) EXNAME_N FROM KEY_EXAM3 WHERE EGID=@EGID ORDER BY EXAMID "
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
                With obj
                    .DataSource = dt
                    .DataTextField = "EXNAME_N"
                    .DataValueField = "EXAMID"
                    .DataBind()
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End With
            Case 3
                'Const cst_xExamLevelv1 As String = "1,2,3,4,5"
                'Const cst_xExamLeveln1 As String = "甲級,乙級,丙級,單一級,不分級"
                parms.Clear()
                parms.Add("EXAMID", s_value)
                sql = "" & vbCrLf
                sql &= " SELECT '1' LEVELID,'甲' LEVELNAME FROM KEY_EXAM3 WHERE EXAMID=@EXAMID AND EXLEVEL LIKE '%甲%'" & vbCrLf
                sql &= " UNION SELECT '2' LEVELID,'乙' LEVELNAME FROM KEY_EXAM3 WHERE EXAMID=@EXAMID AND EXLEVEL LIKE '%乙%'" & vbCrLf
                sql &= " UNION SELECT '3' LEVELID,'丙' LEVELNAME FROM KEY_EXAM3 WHERE EXAMID=@EXAMID AND EXLEVEL LIKE '%丙%'" & vbCrLf
                sql &= " UNION SELECT '4' LEVELID,'單一' LEVELNAME FROM KEY_EXAM3 WHERE EXAMID=@EXAMID AND EXLEVEL LIKE '%單一%'" & vbCrLf
                'sql &= " UNION SELECT '5' LEVELID,'不分' LEVELNAME FROM KEY_EXAM3 WHERE EXAMID=@EXAMID AND EXLEVEL LIKE '%不分%'" & vbCrLf
                Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn, parms)
                With obj
                    .DataSource = dt
                    .DataTextField = "LEVELNAME"
                    .DataValueField = "LEVELID"
                    .DataBind()
                    .Items.Insert(0, New ListItem(TIMS.cst_ddl_PleaseChoose3, ""))
                End With
        End Select
        Return obj
    End Function

    Protected Sub ddlEXAMGROUP_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEXAMGROUP.SelectedIndexChanged
        Dim v_ddlEXAMGROUP As String = TIMS.GetListValue(ddlEXAMGROUP)
        ddlEXAM3.Items.Clear()
        ddlEXLEVEL.Items.Clear()
        If v_ddlEXAMGROUP = "" Then Return
        ddlEXAM3 = GetKeyExam3(ddlEXAM3, v_ddlEXAMGROUP, 2)
    End Sub

    Protected Sub ddlEXAM3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEXAM3.SelectedIndexChanged
        Dim v_ddlEXAM3 As String = TIMS.GetListValue(ddlEXAM3)
        ddlEXLEVEL.Items.Clear()
        If v_ddlEXAM3 = "" Then Return
        ddlEXLEVEL = GetKeyExam3(ddlEXLEVEL, v_ddlEXAM3, 3)
    End Sub
End Class