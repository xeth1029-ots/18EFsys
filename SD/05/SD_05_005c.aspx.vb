Partial Class SD_05_005c
    Inherits AuthBasePage

    'CLASS_SCHEDULE/PLAN_SCHEDULE/COURSE_COURSEINFO/STUD_TRAININGRESULTS
    Dim tmpValue1 As String = "" '組合暫存空間。
    Dim vOCID1 As String = "" '班級代碼暫存。
    Dim odr As DataRow = Nothing '班級資料暫存。

    Const cst_noOCID As String = "班級資料遺失~~"

    Dim objconn As SqlConnection

    'Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
    '    ClientScript.RegisterForEventValidation("ListBox1", "This is Option 1")
    '    ClientScript.RegisterForEventValidation("ListBox1", "This is Option 2")
    '    ClientScript.RegisterForEventValidation("ListBox1", "This is Option 3")
    '    ' Uncomment the line below when you want to specifically register the option for event validation.
    '    ' ClientScript.RegisterForEventValidation("DropDownList1", "Is this option registered for event validation?")
    '    MyBase.Render(writer)
    'End Sub

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
            Button1.Attributes("onclick") = "return check_data();"
            Button2.Attributes("onclick") = "javascript:return result();"
            Button3.Attributes("onclick") = "MoveItem(1);"
            Button4.Attributes("onclick") = "MoveItem(2);"
            Button5.Attributes("onclick") = "MoveItem(3);"
            Button6.Attributes("onclick") = "MoveItem(4);"

            Label1.Visible = False
            Table3.Visible = False

            vOCID1 = TIMS.sUtl_GetRqValue(Me, "OCID")
            odr = Nothing : If vOCID1 <> "" Then odr = TIMS.GetOCIDDate(vOCID1, objconn)
            If odr Is Nothing Then
                Common.MessageBox(Me, cst_noOCID)
                Exit Sub
            End If

            Button7.Visible = False
        End If

        'If UCase(sm.UserInfo.UserID) = "SNOOPY" Then
        '    Button7.Visible = True
        'End If
    End Sub

    '查詢
    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        vOCID1 = TIMS.sUtl_GetRqValue(Me, "OCID")
        odr = Nothing
        If vOCID1 <> "" Then odr = TIMS.GetOCIDDate(vOCID1, objconn)
        If odr Is Nothing Then
            Common.MessageBox(Me, cst_noOCID)
            Exit Sub
        End If

        course.Value = ""
        TextBox1.Text = TIMS.ClearSQM(TextBox1.Text)
        '重建時間配當--完成
        If TextBox1.Text = "@@" Then
            Button7.Visible = True
            TextBox1.Text = "0"
        End If

        '是數字嗎？ture:是
        Dim flag_isnum_1 As Boolean = TIMS.IsNumeric1(TextBox1.Text)
        Dim i_hours As Double = 0
        If flag_isnum_1 Then i_hours = Val(TextBox1.Text)

        '20081215 andy  edit 
        Dim sql As String = ""
        sql &= " SELECT  DISTINCT  CourID  FROM  (" & vbCrLf
        sql &= " SELECT   CourID , OCID, SumCourse  FROM" & vbCrLf
        sql &= " (SELECT  CourID , OCID," & vbCrLf
        sql &= " ISNULL(W1,0)  +ISNULL(W2,0)+ISNULL(W3,0)+ISNULL(W4,0)+ISNULL(W5,0)+ISNULL(W6,0)+ISNULL(W7,0)+ISNULL(W8,0)+ISNULL(W9,0)+ISNULL(W10,0)" & vbCrLf
        sql &= " +ISNULL(W11,0)  +ISNULL(W12,0)+ISNULL(W13,0)+ISNULL(W14,0)+ISNULL(W15,0)+ISNULL(W16,0)+ISNULL(W17,0)+ISNULL(W18,0)+ISNULL(W19,0)+ISNULL(W20,0)" & vbCrLf
        sql &= " +ISNULL(W21,0)  +ISNULL(W22,0)+ISNULL(W23,0)+ISNULL(W24,0)+ISNULL(W25,0)+ISNULL(W26,0)+ISNULL(W27,0)+ISNULL(W28,0)+ISNULL(W29,0)+ISNULL(W30,0)" & vbCrLf
        sql &= " +ISNULL(W31,0)  +ISNULL(W32,0)+ISNULL(W33,0)+ISNULL(W34,0)+ISNULL(W35,0)+ISNULL(W36,0)+ISNULL(W37,0)+ISNULL(W38,0)+ISNULL(W39,0)+ISNULL(W40,0)" & vbCrLf
        sql &= " +ISNULL(W41,0)  +ISNULL(W42,0)+ISNULL(W43,0)+ISNULL(W44,0)+ISNULL(W45,0)+ISNULL(W46,0)+ISNULL(W47,0)+ISNULL(W48,0)+ISNULL(W49,0)+ISNULL(W50,0)" & vbCrLf
        sql &= " +ISNULL(W51,0)  +ISNULL(W52,0)+ISNULL(W53,0)+ISNULL(W54,0)+ISNULL(W55,0)  SumCourse" & vbCrLf
        sql &= " FROM PLAN_SCHEDULE WITH(NOLOCK)" & vbCrLf
        sql &= $" WHERE OCID={vOCID1} and TitleItem=5  ) b" & vbCrLf
        sql &= $" WHERE SumCourse > {i_hours}" & vbCrLf
        sql &= " ) c" & vbCrLf
        Dim dt As DataTable = DbAccess.GetDataTable(sql, objconn)

        Dim s_CourIDstr As String = ""
        If dt.Rows.Count > 0 Then
            For Each dr As DataRow In dt.Rows
                If Convert.ToString(dr("CourID")) <> "" Then
                    tmpValue1 = "'" & Convert.ToString(dr("CourID")) & "'"
                    If tmpValue1 <> "" AndAlso s_CourIDstr.IndexOf(tmpValue1) = -1 Then '重複字過濾
                        If s_CourIDstr <> "" Then s_CourIDstr &= ","
                        s_CourIDstr &= tmpValue1
                    End If
                End If
            Next
        End If
        'If CourIDstr = Nothing Then
        '    CourIDstr = "''"
        'End If

        If s_CourIDstr <> "" Then s_CourIDstr = Trim(s_CourIDstr)
        sql = "" & vbCrLf
        sql &= " SELECT Distinct CourseName ,CourID" & vbCrLf
        sql &= " FROM COURSE_COURSEINFO WITH(NOLOCK)" & vbCrLf
        sql &= " where 1=1" & vbCrLf
        If s_CourIDstr <> "" Then
            sql &= " and CourID in (" & s_CourIDstr & ")" & vbCrLf
        Else
            sql &= " and 1<>1" & vbCrLf
        End If
        sql &= " order by CourseName" & vbCrLf

        '------------------------
        'sql = "SELECT a.CourseName, a.CourID FROM Course_CourseInfo a, (SELECT * FROM Plan_Schedule WHERE OCID='" & Request("OCID") & "' and (COALESCE(W1,0)+COALESCE(W2,0)+COALESCE(W3,0)+COALESCE(W4,0)+COALESCE(W5,0)+COALESCE(W6,0)+COALESCE(W7,0)+COALESCE(W8,0)+COALESCE(W9,0)+COALESCE(W10,0)+COALESCE(W11,0)+COALESCE(W12,0)+COALESCE(W13,0)+COALESCE(W14,0)+COALESCE(W15,0)+COALESCE(W16,0)+COALESCE(W17,0)+COALESCE(W18,0)+COALESCE(W19,0)+COALESCE(W20,0)+COALESCE(W21,0)+COALESCE(W22,0)+COALESCE(W23,0)+COALESCE(W24,0)+COALESCE(W25,0)+COALESCE(W26,0)+COALESCE(W27,0)+COALESCE(W28,0)+COALESCE(W29,0)+COALESCE(W30,0)+COALESCE(W31,0)+COALESCE(W32,0)+COALESCE(W33,0)+COALESCE(W34,0)+COALESCE(W35,0)+COALESCE(W36,0)+COALESCE(W37,0)+COALESCE(W38,0)+COALESCE(W39,0)+COALESCE(W40,0)+COALESCE(W41,0)+COALESCE(W42,0)+COALESCE(W43,0)+COALESCE(W44,0)+COALESCE(W45,0)+COALESCE(W46,0)+COALESCE(W47,0)+COALESCE(W48,0)+COALESCE(W49,0)+COALESCE(W50,0)+COALESCE(W51,0)+COALESCE(W52,0)+COALESCE(W53,0)+COALESCE(W54,0)+COALESCE(W55,0))>'" & hours & "' and TitleItem=5) b WHERE a.CourID=b.CourID"
        '-------------
        'sql = "" & vbCrLf
        'sql += " SELECT Distinct a.CourseName, a.CourID" & vbCrLf
        'sql += " FROM Course_CourseInfo a," & vbCrLf
        'sql += " (SELECT * FROM Plan_Schedule" & vbCrLf
        'sql += " WHERE OCID='" & Request("OCID") & "'" & vbCrLf
        'sql += " and (COALESCE(W1,0)+COALESCE(W2,0)+COALESCE(W3,0)+COALESCE(W4,0)+COALESCE(W5,0)+COALESCE(W6,0)+COALESCE(W7,0)+COALESCE(W8,0)+COALESCE(W9,0)+COALESCE(W10,0)+COALESCE(W11,0)+COALESCE(W12,0)+COALESCE(W13,0)+COALESCE(W14,0)+COALESCE(W15,0)+COALESCE(W16,0)+COALESCE(W17,0)+COALESCE(W18,0)+COALESCE(W19,0)+COALESCE(W20,0)+COALESCE(W21,0)+COALESCE(W22,0)+COALESCE(W23,0)+COALESCE(W24,0)+COALESCE(W25,0)+COALESCE(W26,0)+COALESCE(W27,0)+COALESCE(W28,0)+COALESCE(W29,0)+COALESCE(W30,0)+COALESCE(W31,0)+COALESCE(W32,0)+COALESCE(W33,0)+COALESCE(W34,0)+COALESCE(W35,0)+COALESCE(W36,0)+COALESCE(W37,0)+COALESCE(W38,0)+COALESCE(W39,0)+COALESCE(W40,0)+COALESCE(W41,0)+COALESCE(W42,0)+COALESCE(W43,0)+COALESCE(W44,0)+COALESCE(W45,0)+COALESCE(W46,0)+COALESCE(W47,0)+COALESCE(W48,0)+COALESCE(W49,0)+COALESCE(W50,0)+COALESCE(W51,0)+COALESCE(W52,0)+COALESCE(W53,0)+COALESCE(W54,0)+COALESCE(W55,0))>'" & hours & "' and TitleItem=5) b" & vbCrLf
        'sql += " WHERE a.CourID=b.CourID or a.CourseID ='complextest'" & vbCrLf
        '-------------

        '在資料庫新增一個綜合評量的課程
        'insert INTO Course_CourseInfo(CourseID,CoursENAME,Classification1,Classification2,RID,ModifyAcct,ModifyDATE)
        'VALUES('complextest','綜合評量',1,0,'','sys',getdate())

        ' and SumOfHour>'" & hours & "'
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count = 0 Then
            Common.MessageBox(Me, "查無資料!")
            Table3.Visible = False
        Else
            Label1.Visible = True
            Table3.Visible = True
            Button2.Visible = True
            With ListBox1
                .DataSource = dt
                .DataTextField = "CourseName"
                .DataValueField = "CourID"
                .DataBind()
            End With
        End If

        '搜尋主課程
        '如果有輸入成績，先把那些課程選出
        sql = "" & vbCrLf
        sql &= " select Distinct concat(b.CourseName,'(*)') CourseName,b.CourID" & vbCrLf
        sql &= " FROM STUD_TRAININGRESULTS a WITH(NOLOCK)" & vbCrLf
        sql &= " JOIN CLASS_STUDENTSOFCLASS c WITH(NOLOCK) on c.socid=a.socid" & vbCrLf
        sql &= " JOIN COURSE_COURSEINFO b WITH(NOLOCK) ON a.CourID=b.CourID" & vbCrLf
        sql &= $" WHERE c.OCID={vOCID1}" & vbCrLf
        dt = DbAccess.GetDataTable(sql, objconn)
        If dt.Rows.Count <> 0 Then
            Label1.Visible = True
            Table3.Visible = True
            Button2.Visible = True
            With ListBox2
                .DataSource = dt
                .DataTextField = "CourseName"
                .DataValueField = "CourID"
                .DataBind()
            End With

            For Each dr As DataRow In dt.Rows
                If course.Value <> "" Then course.Value &= ","
                course.Value &= "'" & dr("CourID") & "'"

                If Not ListBox1.Items.FindByValue(dr("CourID")) Is Nothing Then
                    ListBox1.Items.Remove(ListBox1.Items.FindByValue(dr("CourID")))
                End If
            Next
        End If
    End Sub

    '送出
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'Common.RespWrite(Me, "   window.opener.document.form1.DBClass.value='" & Replace(course.Value, "'", "\'") & "';" & vbCrLf)
        'Common.RespWrite(Me, "   window.opener.document.form1.Hidden2.value='" & course.SelectedItem.Text & "'" & vbCrLf)
        Dim courseValue As String = Replace(course.Value, "'", "\'")
        Dim ssScript1 As String = ""
        ssScript1 &= "<script language=javascript>" & vbCrLf
        ssScript1 &= "   var mylabel=window.opener.document.getElementById('Label1');" & vbCrLf
        ssScript1 &= "   mylabel.innerHTML='課程選取完畢!';" & vbCrLf
        ssScript1 &= "   window.opener.document.form1.ChooseClass.value='" & courseValue & "';" & vbCrLf
        ssScript1 &= "   window.close();" & vbCrLf
        ssScript1 &= "</script>" & vbCrLf
        Common.RespWrite(Me, ssScript1)

        'Common.RespWrite(Me, "" & vbCrLf)
        '	function SendDate(){
        '		var mylabel=window.opener.document.getElementById('Label1');
        '      mylabel.innerHTML = getRadioValue(document.form1.course)
        '	}
        '</script> 
    End Sub

    '重建時間配當表
    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        vOCID1 = TIMS.sUtl_GetRqValue(Me, "OCID")
        odr = Nothing : If vOCID1 <> "" Then odr = TIMS.GetOCIDDate(vOCID1, objconn)
        If odr Is Nothing Then
            Common.MessageBox(Me, cst_noOCID)
            Exit Sub
        End If

        Dim OrgID As String = ""
        Dim SUBsql As String = ""

        Dim dt2 As DataTable = Nothing
        Dim da2 As SqlDataAdapter = Nothing
        Dim trans As SqlTransaction = Nothing
        Dim STDate As Date = odr("STDate") '開訓日期
        Dim FTDate As Date = odr("FTDate") '結訓日期
        Dim ccRID1 As String = Convert.ToString(odr("RID")).Substring(0, 1) '取1個RIDS
        If ccRID1 = "" Then
            Common.MessageBox(Me, cst_noOCID)
            Exit Sub
        End If

        'Dim sql As String = ""
        'conn = DbAccess.GetConnection()
        'sql = "SELECT STDate,FTDate FROM Class_ClassInfo WHERE OCID = " & vOCID1
        'STDate = odr("STDate")
        'FTDate = odr("FTDate")

        '增加空白日期存入資料庫--   Start
        Dim dt As DataTable = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim StartDate As Date = STDate
        Dim EndDate As Date = FTDate
        Dim sql As String = ""
        Dim flag As Boolean = False             '要是為True，必須回存日期
        'tConn = DbAccess.GetConnection
        '儲存課表
        Call TIMS.OpenDbConn(objconn)
        'SELECT SCHOOLDATE FROM CLASS_SCHEDULE WHERE ocid =99721 ORDER BY 1
        'SELECT * FROM CLASS_SCHEDULE WHERE ocid =99721 and SchoolDate< dbo.fn_date('20170510') ORDER BY 1
        'delete CLASS_SCHEDULE WHERE  1=1AND CSID in ('5799380','5802574')
        'SELECT * FROM PLAN_SCHEDULE WHERE OCID=99721 ORDER BY 1
        'sql = "SELECT * FROM CLASS_SCHEDULE WHERE OCID='" & OCID.ToString & "' and SchoolDate>=convert(datetime, '" & StartDate & "', 111) and SchoolDate<=convert(datetime, '" & EndDate & "', 111) and Formal='Y'"
        sql = ""
        sql &= $" SELECT * FROM CLASS_SCHEDULE WHERE OCID={vOCID1}"
        sql &= $" and SchoolDate>={TIMS.To_date(StartDate)} and SchoolDate<= {TIMS.To_date(EndDate)}"
        'sql &= " and Formal='Y'"
        dt = DbAccess.GetDataTable(sql, da, objconn)
        While (StartDate <= EndDate)
            If dt.Select($"SchoolDate='{StartDate}'").Length = 0 Then
                Dim iCSID As Integer = DbAccess.GetNewId(objconn, "CLASS_SCHEDULE_CSID_SEQ,CLASS_SCHEDULE,CSID")
                Dim dr As DataRow = dt.NewRow
                dt.Rows.Add(dr)
                dr("CSID") = iCSID
                dr("OCID") = vOCID1
                dr("SchoolDate") = StartDate
                dr("Formal") = "Y"
                dr("Type") = 2
                dr("ModifyAcct") = sm.UserInfo.UserID
                dr("ModifyDate") = Now
                flag = True
            End If
            StartDate = StartDate.AddDays(1)
        End While
        If flag Then
            DbAccess.UpdateDataTable(dt, da)
        Else
            dt = Nothing
        End If
        '增加空白日期存入資料庫--   End

        SUBsql = "SELECT COURID,COURSENAME,MAINCOURID FROM COURSE_COURSEINFO WHERE LEFT(RID,1)=@RID"
        Dim oCmd As New SqlCommand(SUBsql, objconn)
        Call TIMS.OpenDbConn(objconn)
        Dim dtCourseData As New DataTable '= Nothing
        With oCmd
            .Parameters.Clear()
            .Parameters.Add("RID", SqlDbType.VarChar).Value = ccRID1
            dtCourseData.Load(.ExecuteReader())
        End With
        'CourseDataTable = DbAccess.GetDataTable(SUBsql, objconn)
        Try
            trans = DbAccess.BeginTrans(objconn)
            'Plan_Schedule
            '採新增課程可跨年度，因應報表tabel設定此功能 by AMU 20091001
            SD_04_002.AddNew_Plan_Schedule(Me, STDate, FTDate, vOCID1, dt2, da2, trans, dtCourseData)
            DbAccess.CommitTrans(trans)
        Catch ex As Exception
            DbAccess.RollbackTrans(trans)
            Throw ex
        End Try

        Common.MessageBox(Me, "重建時間配當完成")
        Button7.Visible = False
    End Sub
End Class
