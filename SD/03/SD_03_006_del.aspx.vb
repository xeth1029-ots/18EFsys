Partial Class SD_03_006_del
    Inherits System.Web.UI.Page

    Dim objconn As OracleConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在---------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        '檢查Session是否存在---------------------------End

        If Not IsPostBack Then
            create()
        End If

        Button1.Attributes("onclick") = "return CheckData();"
    End Sub

    Sub create()
        Dim sql As String
        Dim dr As DataRow

        sql = "SELECT a.StudentID,b.Name FROM "
        sql += "(SELECT * FROM Class_StudentsOfClass WHERE SOCID='" & Request("SOCID") & "') a "
        sql += "JOIN Stud_StudentInfo b ON a.SID=b.SID "
        dr = DbAccess.GetOneRow(sql, objconn)

        StudentID.Text = Right(dr("StudentID"), 2)
        Name.Text = dr("Name")
    End Sub

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim DelNote As String = ""
        Dim StudStatus As String = ""
        Dim sql As String = ""
        Dim dr As DataRow
        Dim dt As DataTable
        Dim da As OracleDataAdapter = Nothing
        Dim Trans As OracleTransaction
        'Dim conn As OracleConnection
        '2006/03/ add conn by matt
        'conn = DbAccess.GetConnection
        Trans = DbAccess.BeginTrans(objconn)

        Try
            '2006/03/ add conn by matt
            sql = "" & vbCrLf
            sql += " SELECT f.OrgID,f.OrgName,d.PlanName,a.StudentID,e.Name,a.StudStatus,b.PlanID,b.ComIDNO,b.SeqNo,b.ClassCName,b.RID,a.OCID " & vbCrLf
            sql += " FROM Class_StudentsOfClass a" & vbCrLf
            sql += " JOIN Class_ClassInfo b ON a.OCID=b.OCID" & vbCrLf
            sql += " JOIN ID_Plan c ON b.PlanID=c.PlanID" & vbCrLf
            sql += " JOIN Key_Plan d ON c.TPlanID=d.TPlanID" & vbCrLf
            sql += " JOIN Stud_StudentInfo e ON a.SID=e.SID" & vbCrLf
            sql += " JOIN Org_OrgInfo f ON b.ComIDNO=f.ComIDNO" & vbCrLf
            sql += " where 1=1" & vbCrLf
            sql += " and a.SOCID='" & Request("SOCID") & "'" & vbCrLf

            dr = DbAccess.GetOneRow(sql, Trans)
            Select Case Convert.ToString(dr("StudStatus"))
                Case "1"
                    StudStatus = "在訓"
                Case "2"
                    StudStatus = "離訓"
                Case "3"
                    StudStatus = "退訓"
                Case "4"
                    StudStatus = "續訓"
                Case "5"
                    StudStatus = "結訓"
            End Select
            DelNote = "刪除[" & dr("PlanName") & "]-[" & dr("OrgName") & "]-[" & dr("ClassCName") & "]-[(" & dr("StudentID") & ")" & dr("Name") & "]-[" & StudStatus & "]-[" & DelResaon.SelectedItem.Text & IIf(DelResaon.SelectedIndex = DelResaon.Items.Count - 1, ":" & DelReasonOther.Text, "") & "]"
            TIMS.InsertDelLog(sm.UserInfo.UserID, Request("ID"), sm.UserInfo.DistID, DelNote, dr("OrgID"), dr("RID"), dr("PlanID"), dr("ComIDNO"), dr("SeqNO"), dr("OCID"), Request("SOCID"), DelResaon.SelectedValue, DelReasonOther.Text)

            sql = "SELECT * FROM Class_StudentsOfClass WHERE SOCID='" & Request("SOCID") & "'"
            dr = DbAccess.GetOneRow(sql, Trans)
            sql = "SELECT * FROM Class_DelStudentsOfClass WHERE 1<>1"
            dt = DbAccess.GetDataTable(sql, da, Trans)
            Dim dr1 As DataRow
            dr1 = dt.NewRow
            dt.Rows.Add(dr1)
            dr1("SOCID") = dr("SOCID")
            dr1("OCID") = dr("OCID")
            dr1("SID") = dr("SID")
            dr1("StudentID") = dr("StudentID")
            dr1("LevelNo") = dr("LevelNo")
            dr1("EnterDate") = dr("EnterDate")
            dr1("OpenDate") = dr("OpenDate")
            dr1("CloseDate") = dr("CloseDate")
            dr1("RejectTDate1") = dr("RejectTDate1")
            dr1("RejectTDate2") = dr("RejectTDate2")
            dr1("RTReasonID") = dr("RTReasonID")
            dr1("RTReasoOther") = dr("RTReasoOther")
            dr1("StudStatus") = dr("StudStatus")
            dr1("TotalResult") = dr("TotalResult")
            dr1("BehaviorResult") = dr("BehaviorResult")
            dr1("Rank") = dr("Rank")
            dr1("IsOnJob") = dr("IsOnJob")
            dr1("TRNDMode") = dr("TRNDMode")
            dr1("TRNDType") = dr("TRNDType")
            dr1("EnterChannel") = dr("EnterChannel")
            dr1("BudgetID") = dr("BudgetID")
            dr1("GetCertificate") = dr("GetCertificate")
            dr1("GetSubsidy") = dr("GetSubsidy")
            dr1("IdentityID") = dr("IdentityID")
            dr1("MIdentityID") = dr("MIdentityID")
            dr1("SubsidyID") = dr("SubsidyID")
            dr1("TrainHours") = dr("TrainHours")
            dr1("ActNo") = dr("ActNo")
            dr1("SETID") = dr("SETID")
            dr1("ETEnterDate") = dr("ETEnterDate")
            dr1("SerNum") = dr("SerNum")
            dr1("PMode") = dr("PMode")
            dr1("Unit1Hour") = dr("Unit1Hour")
            dr1("Unit2Hour") = dr("Unit2Hour")
            dr1("Unit3Hour") = dr("Unit3Hour")
            dr1("Unit4Hour") = dr("Unit4Hour")
            dr1("ModifyAcct") = sm.UserInfo.UserID
            dr1("ModifyDate") = Now
            DbAccess.UpdateDataTable(dt, da, Trans)

            If dr("ETEnterDate").ToString <> "" Then
                sql = "DELETE Stud_EnterType WHERE SETID='" & dr("SETID").ToString & "' and EnterDate='" & FormatDateTime(dr("ETEnterDate"), 2) & "' and SerNum='" & dr("SerNum").ToString & "'"
                DbAccess.ExecuteNonQuery(sql, Trans)
            End If
            sql = "DELETE Class_StudentsOfClass WHERE SOCID='" & Request("SOCID") & "'"
            DbAccess.ExecuteNonQuery(sql, Trans)

            If TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then
                sql = "SELECT * FROM Stud_ServicePlace WHERE SOCID='" & Request("SOCID") & "'"
                dr = DbAccess.GetOneRow(sql, Trans)
                sql = "SELECT * FROM Stud_DelServicePlace WHERE 1<>1"
                dt = DbAccess.GetDataTable(sql, da, Trans)

                If Not dr Is Nothing Then
                    dr1 = dt.NewRow
                    dt.Rows.Add(dr1)
                    dr1("SOCID") = dr("SOCID")
                    dr1("AcctMode") = dr("AcctMode")
                    dr1("PostNo") = dr("PostNo")
                    dr1("AcctHeadNo") = dr("AcctHeadNo")
                    dr1("AcctExNo") = dr("AcctExNo")
                    dr1("AcctNo") = dr("AcctNo")
                    dr1("BankName") = dr("BankName")
                    dr1("ExBankName") = dr("ExBankName")
                    dr1("FirDate") = dr("FirDate")
                    dr1("Uname") = dr("Uname")
                    dr1("Intaxno") = dr("Intaxno")
                    dr1("ServDept") = dr("ServDept")
                    dr1("JobTitle") = dr("JobTitle")
                    dr1("Zip") = dr("Zip")
                    dr1("Addr") = dr("Addr")
                    dr1("Tel") = dr("Tel")
                    dr1("Fax") = dr("Fax")
                    dr1("SDate") = dr("SDate")
                    dr1("SJDate") = dr("SJDate")
                    dr1("SPDate") = dr("SPDate")
                    dr1("ModifyAcct") = sm.UserInfo.UserID
                    dr1("ModifyDate") = Now
                    DbAccess.UpdateDataTable(dt, da, Trans)

                    sql = "DELETE Stud_ServicePlace WHERE SOCID='" & Request("SOCID") & "'"
                    DbAccess.ExecuteNonQuery(sql, Trans)
                End If

                sql = "SELECT * FROM Stud_TrainBG WHERE SOCID='" & Request("SOCID") & "'"
                dr = DbAccess.GetOneRow(sql, Trans)
                sql = "SELECT * FROM Stud_DelTrainBG WHERE 1<>1"
                dt = DbAccess.GetDataTable(sql, da, Trans)

                If Not dr Is Nothing Then
                    dr1 = dt.NewRow
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

                    sql = "DELETE Stud_TrainBG WHERE SOCID='" & Request("SOCID") & "'"
                    DbAccess.ExecuteNonQuery(sql, Trans)
                End If

                sql = "INSERT INTO Stud_DelTrainBGQ2 (SOCID,Q2) SELECT SOCID,Q2 FROM Stud_TrainBGQ2 WHERE SOCID='" & Request("SOCID") & "'"
                DbAccess.ExecuteNonQuery(sql, Trans)
            End If

            '消除三合一資料
            '學習券
            sql = "SELECT * FROM Adp_DGTRNData WHERE SOCID='" & Request("SOCID") & "'"
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
            sql = "SELECT * FROM Adp_TRNData WHERE SOCID='" & Request("SOCID") & "'"
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
            sql = "SELECT * FROM Adp_GOVTRNData WHERE SOCID='" & Request("SOCID") & "'"
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
            Common.RespWrite(Me, "<script>alert('刪除成功');opener.document.getElementById('Button1').click();window.close();</script>")
        Catch ex As Exception
            DbAccess.RollbackTrans(Trans)
            Throw ex
        End Try
    End Sub
End Class
