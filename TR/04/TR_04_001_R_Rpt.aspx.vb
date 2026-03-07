Partial Class TR_04_001_R_Rpt
    Inherits AuthBasePage

    'Dim arrItem(113, 113) As String
    'Dim arrSpanRow(114) As String
    'Dim sourceDT As DataTable = Nothing
    'Dim PName As String = ""
    'Dim plankind As String = ""

    Dim DistName As String = "DistName"
    Dim mytitle As String = "title"

    Dim tDt_main As New DataTable
    Dim tDt_q1 As New DataTable
    Dim tDt_q2 As New DataTable
    Dim tDt_q3 As New DataTable
    Dim tDt_q4 As New DataTable
    Dim tDt_q5 As New DataTable
    Dim tDt_q6 As New DataTable
    Dim tDt_q7 As New DataTable
    Dim tDt_q8 As New DataTable
    Dim tDt_q9 As New DataTable
    Dim tDt_q10 As New DataTable
    Dim tDt_q11 As New DataTable

    Dim objconn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objconn)
    End Sub

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        objconn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload

        Dim STDate As String = "2013/1/1" '"2005/1/1"
        Dim STDate2 As String = "2013/4/30" '"2005/4/30"
        Dim DistID As String = "%"

        Dim tDt As New DataTable

        'STDate = Request("STDate")
        'STDate2 = Request("STDate2")
        'DistID = Request("DistID")
        'DistName = Request("DistName")
        'title = Request("title")

        tDt_main = db_main(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q1 = db_q1(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q2 = db_q2(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q3 = db_q3(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q4 = db_q4(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q5 = db_q5(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q6 = db_q6(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q7 = db_q7(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q8 = db_q8(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q9 = db_q9(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q10 = db_q10(STDate, STDate2, DistID, DistName, mytitle)
        tDt_q11 = db_q11(STDate, STDate2, DistID, DistName, mytitle)
        'Dim i As Integer
        'i = tDt.Rows.Count

        PrintDiv(tDt_main, "AAAA", "150", "150", 40, "10", "V")

    End Sub


    Private Function db_main(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += "SELECT 1 as ShowTable,'本月人數' as TableName,CONVERT(varchar, Share_ID) as Share_ID,CONVERT(varchar, Share_Name) as Share_Name FROM Adp_ShareSource WHERE Share_Type='301'"
        sql += " UNION"
        sql += " SELECT 1 as ShowTable,'本月人數' as TableName,'99' as Share_ID,'小計' as Share_Name FROM Adp_ShareSource WHERE rownum=1"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,'累積人數' as TableName,CONVERT(varchar, Share_ID) as Share_ID,CONVERT(varchar, Share_Name) as Share_Name FROM Adp_ShareSource WHERE Share_Type='301'"
        sql += " UNION"
        sql += " SELECT 2 as ShowTable,'累積人數' as TableName,'99' as Share_ID,'小計' as Share_Name FROM Adp_ShareSource WHERE rownum=1"
        sql += " Order By ShowTable,Share_ID"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q1(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += "SELECT 1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 總人數 FROM"
        sql += " Adp_DGTRNData"
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo))"
        sql += " Group By OBJECT_TYPE"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,case when COUNT(1)=0 then 0 else  (COUNT(1))/2 end as  總人數 FROM"
        sql += " Adp_DGTRNData"
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 總人數 FROM"
        sql += " Adp_DGTRNData"
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo))"
        sql += " Group By OBJECT_TYPE"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,case when COUNT(1)=0 then 0 else (COUNT(1))/2 end as 總人數 FROM"
        sql += " Adp_DGTRNData"
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo))		"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q2(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT 1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 男 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN"
        sql += "  (SELECT SID FROM Stud_StudentInfo WHERE Sex='M'))"
        sql += " Group By OBJECT_TYPE"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,case when COUNT(1)=0 then 0 else (COUNT(1))/2 end as 男"
        sql += "  FROM Adp_DGTRNData a  WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo"
        sql += "  WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID"
        sql += "  IN (SELECT SID FROM Stud_StudentInfo WHERE Sex='M')) Union ALL SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 男 FROM "
        sql += " Adp_DGTRNData a  WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM "
        sql += " Stud_StudentInfo WHERE Sex='M')) Group By a.OBJECT_TYPE Union"
        sql += "  SELECT  2 as ShowTable,'99' as Share_ID,case when COUNT(1)=0 then 0 else (COUNT(1))/2 end as 男 FROM Adp_DGTRNData"
        sql += "   WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=" & TIMS.To_date(STDate)
        sql += "  and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo"
        sql += "  WHERE Sex='M'))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q3(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 女 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE Sex='F'))"
        sql += " Group By OBJECT_TYPE"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 女 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE Sex='F'))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 女 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE Sex='F'))"
        sql += " Group By OBJECT_TYPE"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 女 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE Sex='F'))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q4(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 年齡45以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=45))"
        sql += " Group By OBJECT_TYPE"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 年齡45以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=45))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 年齡45以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=45))"
        sql += " Group By OBJECT_TYPE"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 年齡45以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=45))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q5(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as age3644 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=36 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=44))"
        sql += " Group By OBJECT_TYPE"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as age3644 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=36 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=44))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as age3644 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=36 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=44))"
        sql += " Group By OBJECT_TYPE"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as  age3644 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=36 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=44))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q6(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as age2635 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and"
        sql += "  SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=26 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=35))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as age2635 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=26 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=35))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as age2635 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=" & TIMS.To_date(STDate)
        sql += " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID "
        sql += " FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=26 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=35))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as age2635 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)>=26 and datepart(year,getdate())-DATEPART(YEAR, Birthday)<=35))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q7(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as age25以下 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)<=25))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as age25以下 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)<=25))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as  age25以下 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)<=25))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as age25以下 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE"
        sql += "  STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and"
        sql += "  SID IN (SELECT SID FROM Stud_StudentInfo WHERE datepart(year,getdate())-DATEPART(YEAR, Birthday)<=25))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q8(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += "  SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 大專以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and"
        sql += "  SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID>=3))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2  as 大專以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID>=3))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 大專以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=" & TIMS.To_date(STDate)
        sql += " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and SID IN (SELECT SID "
        sql += " FROM Stud_StudentInfo WHERE DegreeID>=3))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 大專以上 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE"
        sql += "  STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) "
        sql += " and SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID>=3))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q9(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 高中職 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=2))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 高中職 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=2))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 高中職 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo "
        sql += " WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=2))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2  as 高中職 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo "
        sql += " WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "'))"
        sql += "  and SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=2))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q10(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += "  SELECT  1 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 國中小 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo "
        sql += " WHERE STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=1))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " UNION"
        sql += " SELECT  1 as ShowTable,'99' as Share_ID,(COUNT(1))/2  as 國中小 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=DateAdd('M',-1," & TIMS.To_date(STDate2) & ") and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) "
        sql += " and SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=1))"
        sql += " Union ALL"
        sql += " SELECT 2 as ShowTable,CONVERT(varchar, OBJECT_TYPE) as Share_ID,COUNT(1) as 國中小 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE"
        sql += "  STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) "
        sql += " and SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=1))"
        sql += " Group By CONVERT(varchar, OBJECT_TYPE)"
        sql += " Union"
        sql += " SELECT  2 as ShowTable,'99' as Share_ID,(COUNT(1))/2 as 國中小 FROM"
        sql += " Adp_DGTRNData "
        sql += " WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE "
        sql += " STDate>=" & TIMS.To_date(STDate) & " and STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "')) and "
        sql += " SID IN (SELECT SID FROM Stud_StudentInfo WHERE DegreeID=1))"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function

    Private Function db_q11(ByVal STDate As String, ByVal STDate2 As String, ByVal DistID As String, ByVal DistName As String, ByVal title As String) As DataTable
        Dim resDt As DataTable = Nothing
        Dim sql As String = ""

        sql += " SELECT 查訪班次,已開班數量,參加第三單元人數,參加第二單元人數,參加第一單元人數,完成研習人數,a.Share_ID"
        sql += " ,case when dbo.NVL(f.已開班數量,0)=0 then '0' else Round(g.查訪班次/f.已開班數量,2)*100+'" + DistID + "' end as 查訪率  FROM"
        sql += " (SELECT Share_ID FROM Adp_ShareSource WHERE Share_Type='301') a"
        sql += " JOIN (SELECT COUNT(1) as 完成研習人數 FROM Adp_DGTRNData WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE StudStatus='5' and OCID"
        sql += "  IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like"
        sql += "  '" + DistID + "')) "
        sql += " and SID IN (SELECT SID FROM Stud_StudentInfo))) b ON 1=1"
        sql += " JOIN (SELECT COUNT(1) as 參加第一單元人數 FROM Adp_DGTRNData WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE StudStatus='5' and "
        sql += " OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID "
        sql += " like '" + DistID + "') and Class_Unit like '1__') and SID IN (SELECT SID FROM Stud_StudentInfo))) c ON 1=1"
        sql += " JOIN (SELECT COUNT(1) as 參加第二單元人數 FROM Adp_DGTRNData WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE StudStatus='5' and "
        sql += " OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID"
        sql += "  like '" + DistID + "') and Class_Unit like '_1_') and SID IN (SELECT SID FROM Stud_StudentInfo))) d ON 1=1"
        sql += " JOIN (SELECT COUNT(1) as 參加第三單元人數 FROM Adp_DGTRNData WHERE SOCID IN (SELECT SOCID FROM Class_StudentsOfClass WHERE StudStatus='5' and"
        sql += "  OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID"
        sql += "  like '" + DistID + "') and Class_Unit like '__1') and SID IN (SELECT SID FROM Stud_StudentInfo))) e ON 1=1"
        sql += " JOIN (SELECT COUNT(1) as 已開班數量 FROM Class_ClassInfo WHERE STDate>= " & TIMS.To_date(STDate) & " and STDate<= " & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan"
        sql += "  WHERE DistID like '" + DistID + "') and Class_Unit like '__1' and STDate <=getdate()) f ON 1=1"
        sql += " JOIN (SELECT COUNT(1) as 查訪班次 FROM Class_Visitor WHERE OCID IN (SELECT OCID FROM Class_ClassInfo WHERE STDate>=" & TIMS.To_date(STDate) & " and "
        sql += " STDate<=" & TIMS.To_date(STDate2) & " and PlanID IN (SELECT PlanID FROM ID_Plan WHERE DistID like '" + DistID + "') and Class_Unit like '__1' and STDate <=getdate())) g ON 1=1"

        resDt = DbAccess.GetDataTable(sql, objconn)
        Return resDt
    End Function


    Private Sub PrintDiv(ByVal dt As DataTable, ByVal selRpt As String, ByVal Field1_width As String, ByVal Field2_width As String, ByVal RCount As Integer, ByVal font_size As String, ByVal portrait As String)
        'dt:要顯示的資料,selRpt:,Field1_width:標題題目的寬度,Field2_width:標題題目的寬度,RCount:每頁筆數,font_size:內容字型大小,portrait:直式/橫式

        Dim tmpDT As New DataTable
        'Dim tmpDR As DataRow
        'Dim tmpObj As Object
        Dim sql As String = ""
        Dim PageCount As Int32 = 0  'Pages
        Dim ReportCount As Integer = RCount '每頁筆數
        Dim ColCount As Integer = 0
        Dim intTmp As Integer = 0
        Dim rsCursor As Integer = 0   '報表內容列印的NO
        Dim intPageRecord As Integer = RCount '每頁列印幾筆

        Dim nt As HtmlTable
        Dim nr As HtmlTableRow
        Dim nc As HtmlTableCell
        Dim nl As HtmlGenericControl
        Dim strStyle As String = "font-size:" + font_size + "pt;font-family:DFKai-SB"
        Dim int_width As Integer
        Dim strWatermarkImg As String
        Dim strWatermarkDiv As String
        Dim intWatermarkTop As Integer

        Dim temp As String
        Dim temp2 As String
        Dim str As String
        Dim c As Integer
        Dim b As Boolean
        Dim t As Integer

        tmpDT = dt
        ColCount = dt.Columns.Count

        intTmp = tmpDT.Rows.Count
        PageCount = 1
        'If (intTmp Mod ReportCount) = 0 Then
        '    PageCount = (intTmp / ReportCount) - 1
        'Else
        '    PageCount = intTmp / ReportCount
        'End If

        '表格寬度的設定
        'If portrait = "H" Then
        '    int_width = Int((550 - Field1_width - Field2_width) / 19)
        '    strWatermarkImg = "TIMS_1.jpg"
        'Else
        int_width = 100 'Int((820 - Field1_width - Field2_width) / 19)
        strWatermarkImg = "TIMS_2.jpg"
        'End If

        If dt.Rows.Count > 0 Then


            For i As Integer = 0 To PageCount - 1
                '加背景圖的div
                If portrait = "H" Then
                    intWatermarkTop = i * 800
                Else
                    intWatermarkTop = i * 550
                End If
                strWatermarkDiv = "<div style='position:absolute;z-index:-1; margin:0;padding:0;left:0px;top: " + intWatermarkTop.ToString + "px;'><img src='../../images/rptpic/temple/" + strWatermarkImg + "' /></div>"
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                nl.InnerHtml = strWatermarkDiv

                '表頭
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "0")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = DistName

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "100%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("style", "font-size:14pt;font-family:DFKai-SB")
                nc.InnerHtml = mytitle

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "列印日期：" + Now().ToShortDateString()

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:13pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "20%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:11pt;font-family:DFKai-SB")
                nc.InnerHtml = "頁數：" + (i + 1).ToString + " / " + PageCount.ToString

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "80%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", "font-size:13pt;font-family:DFKai-SB")
                nc.InnerHtml = ""

                'Column Header
                nt = New HtmlTable
                nt.Attributes.Add("style", "width:100%; BORDER-TOP-STYLE: none;FONT-FAMILY: 標楷體;BORDER-RIGHT-STYLE: none;BORDER-LEFT-STYLE: none;BORDER-COLLAPSE: collapse;BORDER-BOTTOM-STYLE: none")
                nt.Attributes.Add("align", "center")
                nt.Attributes.Add("border", "2")
                div_print.Controls.Add(nt)

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "10%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", "2")
                nc.Attributes.Add("rowspan", "2")
                nc.InnerHtml = "短期電腦研習課程"

                t = 0
                c = 0
                str = ""
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    If (temp = "1") Then
                        str = tDt_main.Rows(j).Item("TableName").ToString
                        c += 1
                    End If
                Next
                t += c

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "45%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", c)
                nc.InnerHtml = str

                c = 0
                str = ""
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    If (temp = "2") Then
                        str = tDt_main.Rows(j).Item("TableName").ToString
                        c += 1
                    End If
                Next
                t += c

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                nc.Attributes.Add("width", "45%")
                nc.Attributes.Add("align", "center")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", c)
                nc.InnerHtml = str


                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    If (temp = "1") Then
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", Field1_width)
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = tDt_main.Rows(j).Item("Share_Name").ToString
                    End If
                Next

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    If (temp = "2") Then
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", Field1_width)
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        nc.InnerHtml = tDt_main.Rows(j).Item("Share_Name").ToString
                    End If
                Next

                'row1
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "總人數"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q1.Rows.Count - 1
                            If temp = tDt_q1.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q1.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q1.Rows(k).Item("總人數").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q1.Rows.Count - 1
                            If temp = tDt_q1.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q1.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q1.Rows(k).Item("總人數").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                'row2
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("rowspan", 2)
                nc.InnerHtml = "性別"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "男"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q2.Rows.Count - 1
                            If temp = tDt_q2.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q2.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q2.Rows(k).Item("男").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q2.Rows.Count - 1
                            If temp = tDt_q2.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q2.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q2.Rows(k).Item("男").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "女"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q3.Rows.Count - 1
                            If temp = tDt_q3.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q3.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q3.Rows(k).Item("女").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q3.Rows.Count - 1
                            If temp = tDt_q3.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q3.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q3.Rows(k).Item("女").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                'row3
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("rowspan", 4)
                nc.InnerHtml = "年齡"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "45以上"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q4.Rows.Count - 1
                            If temp = tDt_q4.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q4.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q4.Rows(k).Item("年齡45以上").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q4.Rows.Count - 1
                            If temp = tDt_q4.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q4.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q4.Rows(k).Item("年齡45以上").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "36-44"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q5.Rows.Count - 1
                            If temp = tDt_q5.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q5.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q5.Rows(k).Item("age3644").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q5.Rows.Count - 1
                            If temp = tDt_q5.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q5.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q5.Rows(k).Item("age3644").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "26-35"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q6.Rows.Count - 1
                            If temp = tDt_q6.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q6.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q6.Rows(k).Item("age2635").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q6.Rows.Count - 1
                            If temp = tDt_q6.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q6.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q6.Rows(k).Item("age2635").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "25以下"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q7.Rows.Count - 1
                            If temp = tDt_q7.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q7.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q7.Rows(k).Item("age25以下").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q7.Rows.Count - 1
                            If temp = tDt_q7.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q7.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q7.Rows(k).Item("age25以下").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                'row4
                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("rowspan", 3)
                nc.InnerHtml = "學歷"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "大專以上"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q8.Rows.Count - 1
                            If temp = tDt_q8.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q8.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q8.Rows(k).Item("大專以上").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q8.Rows.Count - 1
                            If temp = tDt_q8.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q8.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q8.Rows(k).Item("大專以上").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "高中職"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q9.Rows.Count - 1
                            If temp = tDt_q9.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q9.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q9.Rows(k).Item("高中職").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q9.Rows.Count - 1
                            If temp = tDt_q9.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q9.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q9.Rows(k).Item("高中職").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.InnerHtml = "國中小"

                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "1") Then
                        b = False
                        For k As Integer = 0 To tDt_q10.Rows.Count - 1
                            If temp = tDt_q10.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q10.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q10.Rows(k).Item("國中小").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next
                For j As Integer = 0 To tDt_main.Rows.Count - 1
                    temp = tDt_main.Rows(j).Item("ShowTable").ToString
                    temp2 = tDt_main.Rows(j).Item("Share_ID").ToString
                    If (temp = "2") Then
                        b = False
                        For k As Integer = 0 To tDt_q10.Rows.Count - 1
                            If temp = tDt_q10.Rows(k).Item("ShowTable").ToString And temp2 = tDt_q10.Rows(k).Item("Share_ID").ToString Then
                                str = tDt_q10.Rows(k).Item("國中小").ToString
                                b = True
                                Exit For
                            End If
                        Next
                        nc = New HtmlTableCell
                        nr.Controls.Add(nc)
                        'nc.Attributes.Add("width", "40%")
                        nc.Attributes.Add("align", "center")
                        nc.Attributes.Add("style", strStyle)
                        If b Then
                            nc.InnerHtml = str
                        Else
                            nc.InnerHtml = "0"
                        End If
                    End If
                Next

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "完成研習人數"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("完成研習人數").ToString

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "參加第一單元人數"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("參加第一單元人數").ToString

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "參加第二單元人數"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("參加第二單元人數").ToString

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "參加第三單元人數"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("參加第三單元人數").ToString

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "已開班數量"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("已開班數量").ToString

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "查訪班次"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("查訪班次").ToString

                nr = New HtmlTableRow
                nt.Controls.Add(nr)

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", 2)
                nc.InnerHtml = "查訪率%"

                nc = New HtmlTableCell
                nr.Controls.Add(nc)
                'nc.Attributes.Add("width", "40%")
                nc.Attributes.Add("align", "left")
                nc.Attributes.Add("style", strStyle)
                nc.Attributes.Add("colspan", t)
                nc.InnerHtml = tDt_q11.Rows(0).Item("查訪率").ToString


[CONTINUE]:
                '表尾
                If rsCursor + 1 > tmpDT.Rows.Count Then
                    GoTo out
                End If
                '換頁列印
                nl = New HtmlGenericControl
                div_print.Controls.Add(nl)
                nl.InnerHtml = "<p style='line-height:2px;margin:0cm;margin-bottom:0.0001pt;mso-pagination:widow-orphan;'><br clear=all style='mso-special-character:line-break;page-break-before:always'>"

            Next
out:
        End If

    End Sub

    Private Sub exc_Print(ByVal portrait As String)
        Dim strScript As String = ""
        strScript = "<script language=""javascript"">window.print();</script>"
        Page.RegisterStartupScript("window_onload", strScript)
        Return

        'Dim strScript As String = ""
        'strScript = "<script language=""javascript"">" + vbCrLf
        ''strScript = "function print() {"
        'strScript += "if (!factory.object) {"
        ''strScript += "return"
        'strScript += "} else {"
        'strScript += "factory.printing.header = """";"
        'strScript += "factory.printing.footer = """";"
        'strScript += "factory.printing.leftMargin = 5; "
        'strScript += "factory.printing.topMargin = 10; "
        'strScript += "factory.printing.rightMargin = 5; "
        'strScript += "factory.printing.bottomMargin = 10; "
        'strScript += "factory.printing.portrait = " + portrait + ";"
        'strScript += "factory.printing.Print(true);"
        'strScript += "window.close();"
        'strScript += "}"
        ''strScript += "}"
        'strScript += "</script>"
        'Page.RegisterStartupScript("window_onload", strScript)
        'Return
    End Sub

End Class
