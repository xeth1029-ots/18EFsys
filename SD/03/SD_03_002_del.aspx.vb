Partial Class SD_03_002_del
    Inherits AuthBasePage

    Dim rqSOCID As String = ""
    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '檢查Session是否存在 Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在 End

        rqSOCID = Request("SOCID")
        rqSOCID = TIMS.ClearSQM(rqSOCID)
        If rqSOCID = "" Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If

        If Not IsPostBack Then Call create()
        Button1.Attributes("onclick") = "return CheckData();"
    End Sub

    Sub create()
        Dim dr As DataRow
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT a.StudentID ,b.Name " & vbCrLf
        sql += " FROM Class_StudentsOfClass a " & vbCrLf
        sql += " JOIN Stud_StudentInfo b ON a.SID = b.SID " & vbCrLf
        sql += " WHERE 1=1 " & vbCrLf
        sql += "    AND a.SOCID = '" & rqSOCID & "' " & vbCrLf
        dr = DbAccess.GetOneRow(sql, objconn)
        If dr Is Nothing Then
            Common.MessageBox(Me, TIMS.cst_NODATAMsg1)
            Exit Sub
        End If
        StudentID.Text = Right(dr("StudentID"), 2)
        Name.Text = dr("Name")
    End Sub

    Sub sUtl_Delete1()
        'ByVal MyPage As Page, ByVal ss3 As String, ByVal oConn As SqlConnection
        'Dim rqSOCID As String = TIMS.GetMyValue(ss3, "rqSOCID")
        'Dim rDelResaon As String = TIMS.GetMyValue(ss3, "rDelResaon") 'DelResaon.SelectedValue
        'Dim flagCanStudSelResultUpdate As Boolean = False
        '2006/03/28 add conn by matt
        'TIMS.TestDbConn(Me, conn, True)
        'sql = "SELECT * FROM Class_StudentsOfClass WHERE SOCID='" & rqSOCID & "'"

        Dim sql As String = ""
        sql = ""
        sql &= " SELECT f.OrgID "
        sql &= "        ,f.OrgName "
        sql &= "        ,d.PlanName "
        sql &= "        ,a.StudentID "
        sql &= "        ,e.Name "
        sql &= "        ,e.IDNO "
        sql &= "        ,a.StudStatus "
        sql &= "        ,b.PlanID "
        sql &= "        ,b.ComIDNO "
        sql &= "        ,b.SeqNo "
        sql &= "        ,b.ClassCName "
        sql &= "        ,b.RID "
        sql &= "        ,a.OCID "
        sql &= " FROM Class_StudentsOfClass a "
        sql += " JOIN Class_ClassInfo b ON a.OCID = b.OCID "
        sql += " JOIN ID_Plan c ON b.PlanID = c.PlanID "
        sql += " JOIN Key_Plan d ON c.TPlanID = d.TPlanID "
        sql += " JOIN Stud_StudentInfo e ON a.SID = e.SID "
        sql += " JOIN Org_OrgInfo f ON b.ComIDNO = f.ComIDNO "
        sql += " WHERE a.SOCID = '" & rqSOCID & "' "
        Dim drS As DataRow = DbAccess.GetOneRow(sql, objconn)
        If drS Is Nothing Then
            Common.MessageBox(Me, "學員資料遺失，請重新查詢資料!!")
            Exit Sub
        End If
        Dim xIDNO As String = Convert.ToString(drS("IDNO"))
        Dim xOCID As String = Convert.ToString(drS("OCID"))

        Dim flagCanStudSelResultUpdate As Boolean = False
        sql = "" & vbCrLf
        sql += " SELECT r.SETID " & vbCrLf
        sql += "        ,CONVERT(VARCHAR, r.ENTERDATE, 111) ETEnterDate " & vbCrLf
        sql += "        ,r.SERNUM " & vbCrLf
        sql += "        ,r.OCID " & vbCrLf
        sql += " FROM Stud_SelResult r " & vbCrLf
        sql += " JOIN Stud_entertype b ON b.setid = r.setid AND b.enterdate = r.enterdate AND b.sernum = r.sernum AND b.ocid1 = r.ocid " & vbCrLf
        sql += " JOIN Stud_entertemp a ON a.setid = b.setid " & vbCrLf
        sql += " WHERE 1=1" & vbCrLf
        sql += "    AND a.IDNO = @IDNO " & vbCrLf
        sql += "    AND b.OCID1 = @OCID1 " & vbCrLf
        Dim sCmd2 As New SqlCommand(sql, objconn)
        Dim dt2 As New DataTable
        Dim dr2 As DataRow = Nothing
        With sCmd2
            .Parameters.Clear()
            .Parameters.Add("IDNO", SqlDbType.VarChar).Value = xIDNO
            .Parameters.Add("OCID1", SqlDbType.Int).Value = xOCID
            dt2.Load(.ExecuteReader())
        End With
        If dt2.Rows.Count > 0 Then
            dr2 = dt2.Rows(0)
            flagCanStudSelResultUpdate = True
        End If
        Select Case DelResaon.SelectedValue
            Case "1", "2", "3"
            Case Else
                Common.MessageBox(Me, "學員資料遺失，請重新查詢資料!!")
                Exit Sub
        End Select

        'Dim StudStatus As String = ""
        Dim DelNote As String = ""
        sql = ""
        sql += " SELECT f.OrgID ,f.OrgName ,d.PlanName ,a.StudentID ,e.Name ,a.StudStatus ,b.PlanID ,b.ComIDNO ,b.SeqNo ,b.ClassCName ,b.RID ,a.OCID "
        sql += " FROM Class_StudentsOfClass a "
        sql += " JOIN Class_ClassInfo b ON a.OCID = b.OCID "
        sql += " JOIN ID_Plan c ON b.PlanID = c.PlanID "
        sql += " JOIN Key_Plan d ON c.TPlanID = d.TPlanID "
        sql += " JOIN Stud_StudentInfo e ON a.SID = e.SID "
        sql += " JOIN Org_OrgInfo f ON b.ComIDNO = f.ComIDNO "
        sql += " WHERE 1=1 "
        sql += " AND a.SOCID = '" & rqSOCID & "' "
        Dim dr As DataRow = DbAccess.GetOneRow(sql, objconn)
        Dim STUDSTATUS_N As String = TIMS.GET_STUDSTATUS_N(dr("StudStatus"))

        DelNote = "刪除[" & dr("PlanName") & "]-[" & dr("OrgName") & "]-[" & dr("ClassCName") & "]-[(" & dr("StudentID") & ")" & dr("Name") & "]-[" & STUDSTATUS_N & "]-[" & DelResaon.SelectedItem.Text & IIf(DelResaon.SelectedIndex = DelResaon.Items.Count - 1, ":" & DelReasonOther.Text, "") & "]"
        TIMS.InsertDelLog(sm.UserInfo.UserID, Request("ID"), sm.UserInfo.DistID, DelNote, dr("OrgID"), dr("RID"), dr("PlanID"), dr("ComIDNO"), dr("SeqNO"), dr("OCID"), rqSOCID, DelResaon.SelectedValue, DelReasonOther.Text)

        Call TIMS.SubInsAccountLog1(Me, Request("ID"), TIMS.cst_wm刪除, "2", Convert.ToString(dr("OCID")), DelNote)

        '刪除學員資料
        Dim iRst As Integer = TIMS.sUtl_DelSTUDENTSOFCLASS(Me, rqSOCID, objconn)
        If objconn Is Nothing Then
            Common.MessageBox(Me, "學員資料遺失，請重新查詢資料!!")
            Exit Sub
        End If
        If iRst = 0 Then
            Common.MessageBox(Me, "學員資料遺失，請重新查詢資料!!")
            Exit Sub
        End If

        '0:中間有異常產生 1:完整結束
        Dim iRst2 As Integer = 0
        iRst2 = sUtl_Delete2(dr2, flagCanStudSelResultUpdate, objconn)

        Dim sScript As String = ""

        '0:中間有異常產生 1:完整結束
        If iRst2 = 0 Then
            'Common.MessageBox(Me, "學員資料遺失，請重新查詢資料!!")
            sScript = ""
            sScript += "<script>"
            sScript += "alert('刪除失敗，請重新執行!!!\n\n若持續出現此問題，請連絡系統管理人員!!!謝謝');"
            sScript += "window.close();"
            sScript += "</script>"
            Common.RespWrite(Me, TIMS.sUtl_AntiXss(sScript))
            Exit Sub
            'Common.MessageBox(Me, "刪除失敗，請重新執行!!!\n\n若持續出現此問題，請連絡系統管理人員!!!謝謝")
        End If

        sScript = ""
        sScript += "<script language=javascript>"
        sScript += "function GetValue() {"
        sScript += "if(window.opener.document.form1.Button1){"
        sScript += "window.opener.document.form1.Button1.click();"
        sScript += "}"
        sScript += "}"

        sScript += "GetValue();" '執行 'SD_03_002.aspx（Button1）查詢按鈕。
        sScript += "</script>"
        sScript += "<script>"
        sScript += "alert('刪除成功!!');"
        sScript += "window.close();"
        sScript += "</script>"
        Common.RespWrite(Me, TIMS.sUtl_AntiXss(sScript))

        'Common.RespWrite(Me, "<script>alert('刪除成功');opener.document.getElementById('Button1').click();window.close();</script>")
        'SD_03_002.aspx（Button1）查詢按鈕。
    End Sub

    '0:中間有異常產生 1:完整結束
    Function sUtl_Delete2(ByVal dr2 As DataRow, ByVal flagCanStudSelResultUpdate As Boolean, ByVal conn As SqlConnection) As Integer
        Dim iRst2 As Integer = 0 '0:中間有異常產生 1:完整結束

        Dim dr As DataRow = Nothing
        Dim sqlAdp As New SqlDataAdapter
        Dim sql As String = ""
        Dim sqlstr As String = ""

        Dim Trans As SqlTransaction = DbAccess.BeginTrans(conn)
        Try
            'update 甄試結果試算檔Stud_SelResult
            If flagCanStudSelResultUpdate Then
                sqlstr = ""
                sqlstr &= " UPDATE Stud_SelResult "
                sqlstr &= " SET AppliedStatus = 'N' "
                sqlstr &= " WHERE 1=1 "
                sqlstr &= "    AND SETID = @SETID "
                sqlstr &= "    AND EnterDate = @EnterDate "
                sqlstr &= "    AND SerNum = @SerNum "
                sqlstr &= "    AND OCID = @OCID "
                Dim uCmd As New SqlCommand(sqlstr, conn, Trans)
                With uCmd
                    .Parameters.Clear()
                    .Parameters.Add("SETID", SqlDbType.Int).Value = dr2("SETID")
                    .Parameters.Add("EnterDate", SqlDbType.DateTime).Value = CDate(dr2("ETEnterDate"))
                    .Parameters.Add("SerNum", SqlDbType.Int).Value = dr2("SerNum")
                    .Parameters.Add("OCID", SqlDbType.Int).Value = dr2("OCID")
                    '.ExecuteNonQuery()  'edit，by:20181017
                    DbAccess.ExecuteNonQuery(uCmd.CommandText, conn, uCmd.Parameters)  'edit，by:20181017
                End With
            End If

#Region "(No Use)"

            '刪除報名資料，造成歷史資料追查問題，故排除刪報名資料 by AMU(Milor) --2009-04-27
            'If dr("ETEnterDate").ToString <> "" Then
            '    sql = " DELETE Stud_EnterType WHERE SETID = '" & dr("SETID").ToString & "' AND EnterDate = '" & FormatDateTime(dr("ETEnterDate"), 2) & "' AND SerNum = '" & dr("SerNum").ToString & "' "
            '    DbAccess.ExecuteNonQuery(sql, Trans)
            'End If
            'sql = " DELETE Class_StudentsOfClass WHERE SOCID='" & rqSOCID & "'"
            'DbAccess.ExecuteNonQuery(sql, Trans)

#End Region

            '刪除結訓成績 (分數大於0)
            sql = " DELETE Stud_TrainingResults WHERE SOCID = '" & rqSOCID & "' "
            DbAccess.ExecuteNonQuery(sql, Trans)

            Dim da As SqlDataAdapter = Nothing
            Dim dt As DataTable = Nothing
            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                '變更Modify人時
                'Dim sqlStr As String = ""
                sqlstr = " UPDATE Stud_ServicePlace SET ModifyAcct = @ModifyAcct ,ModifyDate = GETDATE() WHERE SOCID = @SOCID "
                'Dim sqlAdp As New SqlDataAdapter
                With sqlAdp
                    .UpdateCommand = New SqlCommand(sqlstr, conn, Trans)
                    .UpdateCommand.Parameters.Clear()
                    .UpdateCommand.Parameters.Add("ModifyAcct", SqlDbType.VarChar).Value = Convert.ToString(sm.UserInfo.UserID)
                    .UpdateCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rqSOCID
                    '.UpdateCommand.ExecuteNonQuery()  'edit，by:20181017
                    DbAccess.ExecuteNonQuery(sqlstr, conn, sqlAdp.UpdateCommand.Parameters)  'edit，by:20181017
                End With
                '搬移資料
                sqlstr = " INSERT INTO Stud_ServicePlaceDelData SELECT * FROM Stud_ServicePlace WHERE SOCID = @SOCID "
                With sqlAdp
                    .InsertCommand = New SqlCommand(sqlstr, conn, Trans)
                    .InsertCommand.Parameters.Clear()
                    .InsertCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rqSOCID
                    '.InsertCommand.ExecuteNonQuery()  'edit，by:20181017
                    DbAccess.ExecuteNonQuery(sqlstr, conn, sqlAdp.InsertCommand.Parameters)  'edit，by:20181017
                End With
                '刪除Class_StudentsOfClass
                sqlstr = " DELETE Stud_ServicePlace WHERE SOCID = @SOCID "
                With sqlAdp
                    .DeleteCommand = New SqlCommand(sqlstr, conn, Trans)
                    .DeleteCommand.Parameters.Clear()
                    .DeleteCommand.Parameters.Add("SOCID", SqlDbType.Int).Value = rqSOCID
                    '.DeleteCommand.ExecuteNonQuery()  'edit，by:20181017
                    DbAccess.ExecuteNonQuery(sqlstr, conn, sqlAdp.DeleteCommand.Parameters)  'edit，by:20181017
                End With

                '學員參訓背景(產學訓)
                sql = " SELECT * FROM Stud_TrainBG WHERE SOCID = '" & rqSOCID & "' "
                dr = DbAccess.GetOneRow(sql, Trans)

                '學員參訓背景(產學訓) (刪除檔)
                sql = " SELECT * FROM Stud_DelTrainBG WHERE 1<>1 "
                dt = DbAccess.GetDataTable(sql, da, Trans)
                If Not dr Is Nothing Then
                    Dim dr1 As DataRow = dt.NewRow
                    dt.Rows.Add(dr1)
                    dr1("SOCID") = dr("SOCID")
                    dr1("Q1") = dr("Q1")
                    dr1("Q3") = dr("Q3")
                    dr1("Q3_Other") = dr("Q3_Other")
                    dr1("Q4") = dr("Q4")
                    dr1("Q5") = dr("Q5")
                    dr1("Q61") = dr("Q61")
                    dr1("Q62") = dr("Q62")
                    dr1("Q63") = dr("Q63")
                    dr1("Q64") = dr("Q64")
                    dr1("ModifyAcct") = sm.UserInfo.UserID
                    dr1("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    sql = " DELETE Stud_TrainBG WHERE SOCID = '" & rqSOCID & "' "
                    DbAccess.ExecuteNonQuery(sql, Trans)
                End If

                '學員參訓背景(產學訓) (刪除檔)
                sql = " INSERT INTO Stud_DelTrainBGQ2 (SOCID,Q2) SELECT SOCID ,Q2 FROM Stud_TrainBGQ2 WHERE SOCID = '" & rqSOCID & "' "
                DbAccess.ExecuteNonQuery(sql, Trans)

                '學員參訓背景(產學訓) 
                sql = " DELETE Stud_TrainBGQ2 WHERE SOCID = '" & rqSOCID & "' "
                DbAccess.ExecuteNonQuery(sql, Trans)
            End If

            '消除三合一資料
            '學習券
            sql = " SELECT * FROM Adp_DGTRNData WHERE SOCID = '" & rqSOCID & "' "
            dt = DbAccess.GetDataTable(sql, da, Trans)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                dr("SOCID") = Convert.DBNull
                dr("ARVL_STATE") = 0
                dr("ARVL_DATE") = Convert.DBNull
                dr("ARVL_UNIT_NAME") = Convert.DBNull
                dr("ARVL_ORG_NAME") = Convert.DBNull
                dr("ARVL_ORG_DOCNO") = Convert.DBNull
                dr("ARVL_SDATE") = Convert.DBNull
                dr("ARVL_EDATE") = Convert.DBNull
                dr("ARVL_UNIT_PROMOTER") = Convert.DBNull
                dr("ARVL_UNIT_TEL") = Convert.DBNull
                dr("ACT_END_DATE") = Convert.DBNull
                dr("ARVL_CLASS_NAME") = Convert.DBNull
                dr("ARVL_CLASS_NO") = Convert.DBNull
                dr("TransToTIMS") = "N"
                dr("TIMSModifyDate") = Now
                DbAccess.UpdateDataTable(dt, da, Trans)
            End If
            '職訓券
            sql = " SELECT * FROM Adp_TRNData WHERE SOCID = '" & rqSOCID & "' "
            dt = DbAccess.GetDataTable(sql, da, Trans)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                dr("SOCID") = Convert.DBNull
                dr("ARVL_STATE") = 0
                dr("ARVL_DATE") = Convert.DBNull
                dr("ARVL_SDATE") = Convert.DBNull
                dr("ARVL_EDATE") = Convert.DBNull
                dr("ARVL_UNIT_NAME") = Convert.DBNull
                dr("ARVL_HOURS") = Convert.DBNull
                dr("ARVL_CLASS_NAME") = Convert.DBNull
                dr("ARVL_UNIT_ZIP") = Convert.DBNull
                dr("ARVL_UNIT_ADDR") = Convert.DBNull
                dr("ARVL_UNIT_USER") = Convert.DBNull
                dr("ARVL_UNIT_TEL") = Convert.DBNull
                dr("ARVL_FSH_DATE") = Convert.DBNull
                dr("TransToTIMS") = "N"
                dr("TIMSModifyDate") = Now
                DbAccess.UpdateDataTable(dt, da, Trans)
            End If

            '推介券
            sql = " SELECT * FROM Adp_GOVTRNData WHERE SOCID = '" & rqSOCID & "' "
            dt = DbAccess.GetDataTable(sql, da, Trans)
            If dt.Rows.Count <> 0 Then
                dr = dt.Rows(0)
                dr("SOCID") = Convert.DBNull
                dr("ARVL_STATE") = 0
                dr("ARVL_DATE") = Convert.DBNull
                dr("ARVL_EDATE") = Convert.DBNull
                dr("ARVL_SDATE") = Convert.DBNull
                dr("ARVL_UNIT_NAME") = Convert.DBNull
                dr("ARVL_HOURS") = Convert.DBNull
                dr("ARVL_CLASS_NAME") = Convert.DBNull
                dr("ARVL_UNIT_ZIP") = Convert.DBNull
                dr("ARVL_UNIT_ADDR") = Convert.DBNull
                dr("ARVL_UNIT_USER") = Convert.DBNull
                dr("ARVL_UNIT_TEL") = Convert.DBNull
                dr("ARVL_FSH_DATE") = Convert.DBNull
                dr("TransToTIMS") = "N"
                dr("TIMSModifyDate") = Now
                DbAccess.UpdateDataTable(dt, da, Trans)
            End If
            DbAccess.CommitTrans(Trans)
            iRst2 = 1
        Catch ex As Exception
            DbAccess.RollbackTrans(Trans)
            iRst2 = 0
        End Try

        Return iRst2
    End Function

    '刪除鈕 (?)
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session) If TIMS.ChkSession(Me) Then Exit Sub
        If rqSOCID = "" Then rqSOCID = TIMS.ClearSQM(Request("SOCID"))
        rqSOCID = TIMS.ClearSQM(rqSOCID)
        If rqSOCID = "" Then
            Common.MessageBox(Me, "學員資料遺失，請重新查詢資料!!")
            Exit Sub
        End If
        Call sUtl_Delete1()
    End Sub
End Class