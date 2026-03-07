Partial Class SD_05_003_add
    Inherits AuthBasePage

    ''Dim sql As String
    ''Dim da As SqlDataAdapter = nothing
    ''Dim conn As SqlConnection
    'Dim FunDr As DataRow
    Dim objConn As SqlConnection

    Private Sub sUtl_PageUnload(ByVal sender As Object, ByVal e As System.EventArgs)
        Call TIMS.CloseDbConn(objConn) '--關閉連線
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        '在這裡放置使用者程式碼以初始化網頁
        '檢查Session是否存在--------------------------Start
        ' (直接在 AuthBasePage 處理, 不用個別檢查 Session)  TIMS.CheckSession(Me)
        'TIMS.Get_TitleLab(Request("ID"), TitleLab1, TitleLab2)
        objConn = DbAccess.GetConnection()
        AddHandler MyBase.Unload, AddressOf sUtl_PageUnload
        Call TIMS.OpenDbConn(objConn) '--開啟連線
        '檢查Session是否存在--------------------------End

        If Not IsPostBack Then
            SOCID.Items.Clear()
            SOCID.Items.Add(New ListItem("請選擇班別", 0))
        End If

        TIMS.ShowHistoryClass(Me, HistoryTable, "HistoryList", "OCIDValue1", "OCID1", "", "", "TMIDValue1", "TMID1", True, "Button5")
        If HistoryTable.Rows.Count <> 0 Then
            OCID1.Attributes("onclick") = "showObj('HistoryList');"
            OCID1.Style("CURSOR") = "hand"
        End If
        TIMS.ShowHistoryClass(Me, HistoryTable2, "HistoryList2", "OCIDValue2", "OCID2", "", "", "TMIDValue2", "TMID2")
        If HistoryTable2.Rows.Count <> 0 Then
            OCID2.Attributes("onclick") = "showObj('HistoryList2');"
            OCID2.Style("CURSOR") = "hand"
        End If

        '檢查帳號的功能權限-----------------------------------Start
        'If sm.UserInfo.FunDt Is Nothing Then
        '    Common.RespWrite(Me, "<script>alert('Session過期');</script>")
        '    Common.RespWrite(Me, "<script>top.location.href='../../logout.aspx';</script>")
        'Else
        '    Dim FunDt As DataTable = sm.UserInfo.FunDt
        '    Dim FunDrArray() As DataRow = FunDt.Select("FunID='" & Request("ID") & "'")

        '    If FunDrArray.Length = 0 Then
        '        Common.RespWrite(Me, "<script>alert('您無權限使用該功能');</script>")
        '        Common.RespWrite(Me, "<script>location.href='../../main2.aspx';</script>")
        '    Else
        '       'Dim FunDr As DataRow
        '        FunDr = FunDrArray(0)
        '        If FunDr("Adds") = 1 Then
        '            Button1.Enabled = True
        '        Else
        '            Button1.Enabled = False
        '        End If
        '        If FunDr("Sech") = 1 Then
        '            Button2.Disabled = False
        '            Button3.Disabled = False
        '        Else
        '            Button2.Disabled = True
        '            Button3.Disabled = True
        '        End If
        '    End If
        'End If
        '檢查帳號的功能權限-----------------------------------End

        Me.ViewState("_search") = Nothing
        If Not Session("_search") Is Nothing Then Me.ViewState("_search") = Session("_search")

        Button1.Attributes("onclick") = "javascript:return chkdata()"
        Button5.Style("display") = "none"
        Button4.Attributes("onclick") = "location.href='SD_05_003.aspx?ID=" & Request("ID") & "';"
    End Sub

    Function CheckData1(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(ApplyDate.Text) <> "" Then
            ApplyDate.Text = Trim(ApplyDate.Text)
            If Not TIMS.IsDate1(ApplyDate.Text) Then
                Errmsg += "轉班日期輸入有誤" & vbCrLf
            End If
            If Errmsg = "" Then
                ApplyDate.Text = CDate(ApplyDate.Text).ToString("yyyy/MM/dd")
            End If
        Else
            ApplyDate.Text = ""
            Errmsg += "請先選擇轉班日期" & vbCrLf
        End If

        If Trim(OCIDValue1.Value) = "" Then
            Errmsg += "請選擇 有效 原職類/班級" & vbCrLf
        Else
            If Not IsNumeric(OCIDValue1.Value) Then
                Errmsg += "請選擇 有效 原職類/班級" & vbCrLf
            End If
        End If

        If Trim(SOCID.SelectedValue) = "" Then
            Errmsg += "請選擇 有效 學員姓名" & vbCrLf
        Else
            If Not IsNumeric(SOCID.SelectedValue) Then
                Errmsg += "請選擇 有效 學員姓名" & vbCrLf
            End If
        End If

        If Trim(OCIDValue2.Value) = "" Then
            Errmsg += "請選擇/輸入 有效 轉到職類/班級" & vbCrLf
        Else
            If Not IsNumeric(OCIDValue2.Value) Then
                Errmsg += "請選擇/輸入 有效 轉到職類/班級" & vbCrLf
            End If
        End If

        '確認轉班學員，與輸入欄位
        Dim dr As DataRow = Nothing 'Class_ClassInfo:1
        Dim dr2 As DataRow = Nothing 'Class_ClassInfo:2
        'Dim dt3 As DataTable = Nothing 'Class_StudentsOfClass:3
        'Dim Sql As String = ""
        If Errmsg = "" Then
            'Sql = ""
            'Sql += " SELECT NORID,OTHERREASON,ISBUSINESS,APPLIEDRESULTR,APPLIEDRESULTM,PNUM,   " & vbCrLf
            'Sql += " ISCONT,QAYSDATE,QAYFDATE,LASTSTATE,TADDRESSZIP6W,EVTA_NOSHOW,   " & vbCrLf
            'Sql += " ETRAIN_SHOW,ECOMMENT,COMPANYNAME,NOTICE,CJOB_UNKEY,EXAMPERIOD,   " & vbCrLf
            'Sql += " OCID,CLSID,PLANID,YEARS,CYCLTYPE,LEVELTYPE,RID,CLASSCNAME,CLASSENGNAME,   " & vbCrLf
            'Sql += " CONTENT,PURPOSE," & vbCrLf
            'Sql += " TPROPERTYID,TMID,CLID,SENTERDATE,FENTERDATE,CHECKINDATE,EXAMDATE,STDATE,   " & vbCrLf
            'Sql += " FTDATE,TADDRESSZIP,TADDRESS,THOURS,TNUM,TDEADLINE,TPERIOD,NOTOPEN,   " & vbCrLf
            'Sql += " ISAPPLIC,RELSHIP,COMIDNO,SEQNO,ISCALCULATE,ISSUCCESS,CTNAME,MODIFYACCT,   " & vbCrLf
            'Sql += " MODIFYDATE,CLASSNUM,LEVELCOUNT,ISFULLDATE,CLASS_UNIT,ISCLOSED,BGTIME   " & vbCrLf
            'Sql += " FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "'"
            'dr = DbAccess.GetOneRow(Sql, objConn)
            'Dim dr As DataRow = Nothing 'Class_ClassInfo:1
            dr = TIMS.GetOCIDDate(OCIDValue1.Value, objConn)
            If dr Is Nothing Then
                Errmsg += "請選擇 有效 原職類/班級" & vbCrLf
            End If
            'Sql = ""
            'Sql += " SELECT NORID,OTHERREASON,ISBUSINESS,APPLIEDRESULTR,APPLIEDRESULTM,PNUM,   " & vbCrLf
            'Sql += " ISCONT,QAYSDATE,QAYFDATE,LASTSTATE,TADDRESSZIP6W,EVTA_NOSHOW,   " & vbCrLf
            'Sql += " ETRAIN_SHOW,ECOMMENT,COMPANYNAME,NOTICE,CJOB_UNKEY,EXAMPERIOD,   " & vbCrLf
            'Sql += " OCID,CLSID,PLANID,YEARS,CYCLTYPE,LEVELTYPE,RID,CLASSCNAME,CLASSENGNAME,   " & vbCrLf
            'Sql += " CONTENT,PURPOSE," & vbCrLf
            'Sql += " TPROPERTYID,TMID,CLID,SENTERDATE,FENTERDATE,CHECKINDATE,EXAMDATE,STDATE,   " & vbCrLf
            'Sql += " FTDATE,TADDRESSZIP,TADDRESS,THOURS,TNUM,TDEADLINE,TPERIOD,NOTOPEN,   " & vbCrLf
            'Sql += " ISAPPLIC,RELSHIP,COMIDNO,SEQNO,ISCALCULATE,ISSUCCESS,CTNAME,MODIFYACCT,   " & vbCrLf
            'Sql += " MODIFYDATE,CLASSNUM,LEVELCOUNT,ISFULLDATE,CLASS_UNIT,ISCLOSED,BGTIME   " & vbCrLf
            'Sql += " FROM Class_ClassInfo WHERE OCID='" & OCIDValue2.Value & "'"
            'dr2 = DbAccess.GetOneRow(Sql, objConn)
            'Dim dr2 As DataRow = Nothing 'Class_ClassInfo:2
            dr2 = TIMS.GetOCIDDate(OCIDValue2.Value, objConn)
            If dr2 Is Nothing Then
                Errmsg += "請選擇 有效 轉到職類/班級" & vbCrLf
            End If
        End If

        If Errmsg = "" Then
            If OCIDValue1.Value = OCIDValue2.Value Then
                Errmsg += "原職類/班級 不可相同於 轉到職類/班級" & vbCrLf
            End If
            'If DateDiff(DateInterval.Day, CDate(dr("STDate")), CDate(Now)) > 14 Then
            '    Errmsg += "原職類/班級 比對 轉班日期 超過開訓日兩週不能轉班" & vbCrLf
            'End If
            If Convert.ToString(dr("IsClosed")) = "Y" Then
                Errmsg += "原職類/班級 已經結訓 不能轉班" & vbCrLf
            End If
            'If DateDiff(DateInterval.Day, CDate(dr2("STDate")), CDate(Now)) > 14 Then
            '    Errmsg += "轉到職類/班級 比對 轉班日期 超過開訓日兩週不能轉班" & vbCrLf
            'End If
            If Convert.ToString(dr2("IsClosed")) = "Y" Then
                Errmsg += "轉到職類/班級 已經結訓 不能轉班" & vbCrLf
            End If
        End If

        If Errmsg = "" Then
            'Sql = ""
            'Sql &= " SELECT a.SOCID,b.Name "
            'Sql += " FROM (SELECT SID,SOCID,OCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') a "
            'Sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID"
            'dt3 = DbAccess.GetDataTable(Sql, objConn)
            Dim dt3 As DataTable = Nothing 'Class_StudentsOfClass:3
            dt3 = Get_Stddt(OCIDValue1.Value, objConn)
            If dt3.Rows.Count = 0 Then
                Errmsg += "原職類/班級 查無 在訓學生資料 不能轉班" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim Errmsg As String = ""
        Call CheckData1(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        '轉班學號變更 by AMU 
        '=============== 轉班學號變更 START =======================================================
        Dim MaxStudentID As Integer
        Dim StudentID As String = ""
        Dim sql As String = ""
        sql = "" & vbCrLf
        sql += " SELECT cc.OCID" & vbCrLf
        sql += " ,cc.Years" & vbCrLf
        sql += " ,d.ClassID" & vbCrLf
        sql += " ,cc.CyclType " & vbCrLf
        sql += " ,e.DistID" & vbCrLf
        sql += " ,f.OrgKind" & vbCrLf
        sql += " ,pp.IsBusiness" & vbCrLf
        sql += " ,pp.ClassCate" & vbCrLf
        sql += " ,cc.TMID" & vbCrLf
        sql += " FROM Class_ClassInfo cc " & vbCrLf
        sql += " JOIN ID_Class d ON d.CLSID =cc.CLSID" & vbCrLf
        sql += " JOIN Auth_Relship e ON e.RID=cc.RID" & vbCrLf
        sql += " JOIN Org_OrgInfo f ON f.OrgID=e.OrgID" & vbCrLf
        sql += " JOIN Plan_PlanInfo pp on pp.PlanID=cc.PlanID AND pp.ComIDNO=cc.ComIDNO AND pp.SeqNo=cc.SeqNo" & vbCrLf
        sql += " WHERE cc.OCID='" & OCIDValue2.Value & "'" & vbCrLf
        Dim dr1 As DataRow = DbAccess.GetOneRow(sql, objConn)

        If Not TIMS.Cst_TPlanID28AppPlan.IndexOf(sm.UserInfo.TPlanID) > -1 Then

            '學號增加的方式應該是要去除前面的固定長度字串，才做流水號處理，並將流水號由原本的2碼變3碼。
            If dr1("Years").ToString.Length = 4 Then
                Me.ViewState("StudentID") = Microsoft.VisualBasic.Right(dr1("Years").ToString, 2) & "0" & dr1("ClassID") & dr1("CyclType")
            ElseIf dr1("Years").ToString.Length = 2 Then
                Me.ViewState("StudentID") = dr1("Years").ToString & "0" & dr1("ClassID") & dr1("CyclType")
            Else
                Me.ViewState("StudentID") = Microsoft.VisualBasic.Right(Now.Year.ToString, 2) & "0" & dr1("ClassID") & dr1("CyclType")
            End If
            'Me.ViewState("StudentID") = dr1("Years") & "0" & dr1("ClassID") & dr1("CyclType")
            sql = "SELECT dbo.NVL(max(CONVERT(numeric, replace(StudentID,'" & Me.ViewState("StudentID") & "',''))),0)+1 as MaxNum FROM Class_StudentsOfClass WHERE OCID='" & dr1("OCID") & "'"
            MaxStudentID = DbAccess.ExecuteScalar(sql, objConn)
            StudentID = Me.ViewState("StudentID") & Format(MaxStudentID, "00#")

        Else

            'sm.UserInfo.TPlanID = "28"
            '因為有資料庫交易問題所以提前呼叫
            StudentID = TIMS.Get_TPlanID28_StudentID(
                                          dr1("Years").ToString,
                                          dr1("DistID").ToString,
                                          dr1("OrgKind").ToString,
                                          dr1("IsBusiness").ToString,
                                          dr1("ClassID").ToString,
                                          dr1("CyclType").ToString,
                                          dr1("ClassCate").ToString,
                                          dr1("TMID").ToString, objConn)

            Me.ViewState("StudentID") = StudentID

            sql = "" & vbCrLf
            sql += " SELECT CONVERT(numeric, dbo.NVL(MAX( dbo.SUBSTR2(StudentID, -2)),0))+1 MaxNum " & vbCrLf
            sql += " FROM Class_StudentsOfClass" & vbCrLf
            sql += " WHERE OCID ='" & dr1("OCID") & "'" & vbCrLf
            'sql2 = "SELECT convert(int,ISNULL(RIGHT(MAX(StudentID),2),0))+1 as MaxNum FROM Class_StudentsOfClass WHERE OCID='" & dr1("OCID") & "'"
            MaxStudentID = DbAccess.ExecuteScalar(sql, objConn)
            StudentID = Me.ViewState("StudentID") & Format(MaxStudentID, "0#")
        End If
        '=============== 轉班學號變更 END =======================================================

        Dim da As SqlDataAdapter = Nothing
        Dim da1 As SqlDataAdapter = Nothing
        Dim oConn As SqlConnection = DbAccess.GetConnection()
        Dim oTrans As SqlTransaction = DbAccess.BeginTrans(oConn)
        Try
            'Dim dt As DataTable
            'Dim dt1 As DataTable
            'Dim dr As DataRow
            'Dim dr1 As DataRow
            'Dim da1 As SqlDataAdapter
            sql = "SELECT * FROM Stud_TranClassRecord WHERE 1<>1"
            Dim dt As DataTable = DbAccess.GetDataTable(sql, da, oTrans)
            Dim dr As DataRow = dt.NewRow
            dr("SOCID") = SOCID.SelectedValue
            dr("OrigClassID") = OCIDValue1.Value
            dr("NewClassID") = OCIDValue2.Value
            dr("ApplyDate") = ApplyDate.Text
            dr("Reason") = Reason.Text
            dr("ModifyAcct") = sm.UserInfo.UserID
            dr("ModifyDate") = Now
            dt.Rows.Add(dr)
            DbAccess.UpdateDataTable(dt, da, oTrans)

            sql = "SELECT SOCID,OCID,StudentID FROM Class_StudentsOfClass WHERE SOCID='" & SOCID.SelectedValue & "'"
            Dim dt1 As DataTable = DbAccess.GetDataTable(sql, da1, oTrans)
            dr1 = dt1.Rows(0)
            dr1("OCID") = OCIDValue2.Value
            dr1("StudentID") = StudentID
            DbAccess.UpdateDataTable(dt1, da1, oTrans)
            DbAccess.CommitTrans(oTrans)

        Catch ex As Exception
            DbAccess.RollbackTrans(oTrans)
            Common.MessageBox(Me, "儲存失敗!!")
            Exit Sub
        End Try
        Call TIMS.CloseDbConn(oConn)

        Common.RespWrite(Me, "<script>alert('新增成功');</script>")
        Common.RespWrite(Me, "<script>location.href='SD_05_003.aspx?ID=" & Request("ID") & "';</script>")
    End Sub

    Private Sub Button4_ServerClick(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.ServerClick
        If Not Me.ViewState("_search") Is Nothing Then Session("_search") = Me.ViewState("_search")
        TIMS.Utl_Redirect1(Me, "SD_05_003.aspx?ID=" & Request("ID"))
    End Sub

    Function CheckData2(ByRef Errmsg As String) As Boolean
        Dim Rst As Boolean = True
        Errmsg = ""

        If Trim(OCIDValue1.Value) = "" Then
            Errmsg += "請選擇 有效 原職類/班級" & vbCrLf
        Else
            If Not IsNumeric(OCIDValue1.Value) Then
                Errmsg += "請選擇 有效 原職類/班級" & vbCrLf
            End If
        End If

        '確認轉班學員，與輸入欄位
        Dim dr As DataRow = Nothing 'Class_ClassInfo:1
        'Dim dr2 As DataRow = Nothing 'Class_ClassInfo:2
        Dim dt3 As DataTable = Nothing 'Class_StudentsOfClass:3

        Dim Sql As String = ""
        If Errmsg = "" Then
            Sql = ""
            Sql += " SELECT NORID,OTHERREASON,ISBUSINESS,APPLIEDRESULTR,APPLIEDRESULTM,PNUM,   " & vbCrLf
            Sql += " ISCONT,QAYSDATE,QAYFDATE,LASTSTATE,TADDRESSZIP6W,EVTA_NOSHOW,   " & vbCrLf
            Sql += " ETRAIN_SHOW,ECOMMENT,COMPANYNAME,NOTICE,CJOB_UNKEY,EXAMPERIOD,   " & vbCrLf
            Sql += " OCID,CLSID,PLANID,YEARS,CYCLTYPE,LEVELTYPE,RID,CLASSCNAME,CLASSENGNAME,   " & vbCrLf
            Sql += " CONTENT,PURPOSE," & vbCrLf
            Sql += " TPROPERTYID,TMID,CLID,SENTERDATE,FENTERDATE,CHECKINDATE,EXAMDATE,STDATE,   " & vbCrLf
            Sql += " FTDATE,TADDRESSZIP,TADDRESS,THOURS,TNUM,TDEADLINE,TPERIOD,NOTOPEN,   " & vbCrLf
            Sql += " ISAPPLIC,RELSHIP,COMIDNO,SEQNO,ISCALCULATE,ISSUCCESS,CTNAME,MODIFYACCT,   " & vbCrLf
            Sql += " MODIFYDATE,CLASSNUM,LEVELCOUNT,ISFULLDATE,CLASS_UNIT,ISCLOSED,BGTIME   " & vbCrLf
            Sql += " FROM Class_ClassInfo WHERE OCID='" & OCIDValue1.Value & "'"
            dr = DbAccess.GetOneRow(Sql, objConn)
            If dr Is Nothing Then
                Errmsg += "請選擇 有效 原職類/班級" & vbCrLf
            End If
        End If

        If Errmsg = "" Then
            If OCIDValue1.Value = OCIDValue2.Value Then
                Errmsg += "原職類/班級 不可相同於 轉到職類/班級" & vbCrLf
            End If
            'If DateDiff(DateInterval.Day, CDate(dr("STDate")), CDate(Now)) > 14 Then
            '    Errmsg += "原職類/班級 比對 轉班日期 超過開訓日兩週不能轉班" & vbCrLf
            'End If
            If Convert.ToString(dr("IsClosed")) = "Y" Then
                Errmsg += "原職類/班級 已經結訓 不能轉班" & vbCrLf
            End If
        End If

        If Errmsg = "" Then
            'Sql = "SELECT a.SOCID,b.Name FROM "
            'Sql += "(SELECT SID,SOCID,OCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') a "
            'Sql += "JOIN Stud_StudentInfo b ON a.SID=b.SID"
            dt3 = Get_Stddt(OCIDValue1.Value, objConn)
            If dt3.Rows.Count = 0 Then
                Errmsg += "原職類/班級 查無 在訓學生資料 不能轉班" & vbCrLf
            End If
        End If

        If Errmsg <> "" Then Rst = False
        Return Rst
    End Function

    Sub Add_Student()
        SOCID.Items.Clear()

        If OCIDValue1.Value <> "" Then
            'Dim Sql As String = ""
            'Sql = ""
            'Sql &= " SELECT a.SOCID,b.Name "
            'Sql += " FROM (SELECT SID,SOCID,OCID FROM Class_StudentsOfClass WHERE OCID='" & OCIDValue1.Value & "') a "
            'Sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID"

            Dim dtResult As DataTable = Get_Stddt(OCIDValue1.Value, objConn)
            'Dim dtResult As DataTable = DbAccess.GetDataTable(Sql, objConn)
            If dtResult.Rows.Count = 0 Then
                Common.MessageBox(Me, "原職類/班級 查無 在訓學生資料")
                Exit Sub
            Else
                With SOCID
                    .DataSource = dtResult
                    .DataTextField = "Name"
                    .DataValueField = "SOCID"
                    .DataBind()
                    .Items.Insert(0, New ListItem("===請選擇===", ""))
                End With
            End If

        End If
    End Sub

    Function Get_Stddt(ByVal OCID As String, ByRef tConn As SqlConnection) As DataTable
        Dim rDt As DataTable = Nothing
        Dim sql As String = ""
        sql = ""
        sql &= " SELECT a.SOCID,b.Name "
        sql += " FROM (SELECT SID,SOCID,OCID FROM Class_StudentsOfClass WHERE OCID='" & OCID & "') a "
        sql += " JOIN Stud_StudentInfo b ON a.SID=b.SID"
        rDt = DbAccess.GetDataTable(sql, tConn)
        Return rDt
    End Function

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        Dim Errmsg As String = ""
        Call CheckData2(Errmsg)
        If Errmsg <> "" Then
            Common.MessageBox(Page, Errmsg)
            Exit Sub
        End If

        SOCID.Items.Clear()
        Call Add_Student()
    End Sub
End Class
